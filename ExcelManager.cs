using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace KiCadExcelBridge
{
    public class ExcelManager
    {
        private readonly List<string> _filePaths;
        private readonly List<SheetMapping> _sheetMappings;
        // Key: fullPath + "::" + sheetName
        private readonly Dictionary<string, List<Dictionary<int, string>>> _data = new();
        private readonly Dictionary<string, CategoryInfo> _categories = new(StringComparer.OrdinalIgnoreCase);

        private class CategoryInfo
        {
            public string Id { get; init; } = string.Empty;
            public string Name { get; init; } = string.Empty;
            public SheetMapping Mapping { get; init; }
            public Dictionary<string, string> Filters { get; init; } = new();
        }

        public ExcelManager(List<string> filePaths, List<SheetMapping> sheetMappings)
        {
            _filePaths = filePaths;
            _sheetMappings = sheetMappings;
            EpplusLicenseHelper.EnsureNonCommercialLicense();
            LoadData();
        }

        private void LoadData()
        {
            _data.Clear();

            foreach (var mapping in _sheetMappings)
            {
                var filePath = mapping.SourceFile;
                if (!_filePaths.Contains(filePath))
                {
                    continue;
                }

                var key = GetDataKey(filePath, mapping.SheetName);
                if (_data.ContainsKey(key))
                {
                    // Already loaded for this sheet (multiple mappings may reference it)
                    continue;
                }

                try
                {
                    if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        _data[key] = LoadWorksheetRows(filePath, mapping);
                    }
                    else if (Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        _data[key] = LoadCsvRows(filePath, mapping);
                    }
                }
                catch
                {
                    // ignore file read errors for now
                }
            }

            RefreshCategories();
        }

        private void RefreshCategories()
        {
            _categories.Clear();
            var slugCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var mapping in _sheetMappings)
            {
                var baseLabel = mapping.CategoryLabel?.Trim();
                if (string.IsNullOrWhiteSpace(baseLabel)) continue;

                var splitFields = mapping.FieldMappings
                    .Where(f => f.Split && f.ColumnIndex.HasValue && f.ColumnIndex.Value > 0)
                    .ToList();

                if (splitFields.Count == 0)
                {
                    // Standard category
                    var slug = GenerateUniqueSlug(baseLabel, slugCounts);
                    _categories[slug] = new CategoryInfo
                    {
                        Id = slug,
                        Name = baseLabel,
                        Mapping = mapping
                    };
                }
                else
                {
                    // Split category
                    var key = GetDataKey(mapping.SourceFile, mapping.SheetName);
                    if (!_data.TryGetValue(key, out var rows)) continue;

                    // Find distinct combinations
                    var combinations = new HashSet<string>();
                    var combinationValues = new Dictionary<string, Dictionary<string, string>>();

                    foreach (var row in rows)
                    {
                        var currentValues = new Dictionary<string, string>();
                        var combinationKeyParts = new List<string>();

                        foreach (var field in splitFields)
                        {
                            var val = GetRowValue(row, field.ColumnIndex.Value);
                            currentValues[field.FieldName] = val;
                            combinationKeyParts.Add(val);
                        }

                        var combinationKey = string.Join("|||", combinationKeyParts);
                        if (combinations.Add(combinationKey))
                        {
                            combinationValues[combinationKey] = currentValues;
                        }
                    }

                    foreach (var comboKey in combinations)
                    {
                        var values = combinationValues[comboKey];
                        // Build name: "Category - Val1 - Val2"
                        var nameParts = new List<string> { baseLabel };
                        nameParts.AddRange(values.Values);
                        var fullName = string.Join(" - ", nameParts);

                        var slug = GenerateUniqueSlug(fullName, slugCounts);
                        _categories[slug] = new CategoryInfo
                        {
                            Id = slug,
                            Name = fullName,
                            Mapping = mapping,
                            Filters = values
                        };
                    }
                }
            }
        }

        private static string GenerateUniqueSlug(string name, Dictionary<string, int> counts)
        {
            var baseSlug = Slugify(name);
            if (!counts.ContainsKey(baseSlug))
            {
                counts[baseSlug] = 1;
                return baseSlug;
            }

            var count = counts[baseSlug];
            counts[baseSlug] = count + 1;
            return $"{baseSlug}-{count}";
        }

        private static List<Dictionary<int, string>> LoadWorksheetRows(string filePath, SheetMapping mapping)
        {
            var rows = new List<Dictionary<int, string>>();
            try
            {
                var stream = FileHelpers.OpenFileForReadWithFallback(filePath, maxRetries: 4, out var actualPath, out var usedFallback);
                if (stream == null)
                {
                    return rows;
                }

                using (stream)
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[mapping.SheetName];
                    if (worksheet == null || worksheet.Dimension == null)
                    {
                        // cleanup fallback copy if any
                    }
                    else
                    {
                        var endRow = worksheet.Dimension.End.Row;
                        var endColumn = worksheet.Dimension.End.Column;
                        var startRow = mapping.IgnoreHeader ? 2 : 1;
                        if (startRow <= endRow)
                        {
                            for (int row = startRow; row <= endRow; row++)
                            {
                                var rowData = new Dictionary<int, string>();
                                for (int col = 1; col <= endColumn; col++)
                                {
                                    rowData[col] = worksheet.Cells[row, col].Text?.Trim() ?? string.Empty;
                                }
                                rows.Add(rowData);
                            }
                        }
                    }
                }

                // If we used a fallback temporary file, attempt to delete it
                try
                {
                    // actualPath set by helper; only delete if it looks like a temp copy
                    if (!string.IsNullOrEmpty(actualPath) && usedFallback)
                    {
                        File.Delete(actualPath);
                    }
                }
                catch
                {
                    // ignore deletion errors
                }
            }
            catch
            {
                // ignore
            }

            return rows;
        }

        private static List<Dictionary<int, string>> LoadCsvRows(string filePath, SheetMapping mapping)
        {
            var rows = new List<Dictionary<int, string>>();
            try
            {
                var stream = FileHelpers.OpenFileForReadWithFallback(filePath, maxRetries: 4, out var actualPath, out var usedFallback);
                if (stream == null)
                {
                    return rows;
                }

                using (stream)
                using (var sr = new StreamReader(stream))
                {
                    var lines = new List<string>();
                    while (!sr.EndOfStream)
                    {
                        lines.Add(sr.ReadLine() ?? string.Empty);
                    }

                    var startIndex = mapping.IgnoreHeader ? 1 : 0;
                    for (int i = startIndex; i < lines.Count; i++)
                    {
                        var values = lines[i].Split(',');
                        var rowData = new Dictionary<int, string>();
                        for (int col = 0; col < values.Length; col++)
                        {
                            rowData[col + 1] = values[col].Trim();
                        }
                        rows.Add(rowData);
                    }
                }

                try
                {
                    if (!string.IsNullOrEmpty(actualPath) && usedFallback)
                    {
                        File.Delete(actualPath);
                    }
                }
                catch
                {
                    // ignore
                }
            }
            catch
            {
                // ignore
            }

            return rows;
        }

        private static string GetDataKey(string filePath, string sheetName) => filePath + "::" + sheetName;

        public IEnumerable<object> GetCategories()
        {
            return _categories.Values.Select(c => new { id = c.Id, name = c.Name, description = string.Empty });
        }

        public IEnumerable<object> GetPartsForCategory(string categoryId)
        {
            if (string.IsNullOrWhiteSpace(categoryId) || !_categories.TryGetValue(categoryId, out var category))
            {
                return Enumerable.Empty<object>();
            }

            var parts = new List<object>();
            var sheet = category.Mapping;
            
            var idMapping = FindFieldMapping(sheet, "ID");
            if (idMapping?.ColumnIndex is not int idColumn || idColumn <= 0)
            {
                return parts;
            }

            var key = GetDataKey(sheet.SourceFile, sheet.SheetName);
            if (!_data.TryGetValue(key, out var sheetData))
            {
                return parts;
            }

            var config = ConfigurationManager.Load();

            foreach (var row in sheetData)
            {
                // Check filters
                if (category.Filters != null && category.Filters.Count > 0)
                {
                    var match = true;
                    foreach (var filter in category.Filters)
                    {
                        var field = FindFieldMapping(sheet, filter.Key);
                        if (field?.ColumnIndex is int colIdx && colIdx > 0)
                        {
                            var rowVal = GetRowValue(row, colIdx);
                            if (!string.Equals(rowVal, filter.Value, StringComparison.OrdinalIgnoreCase))
                            {
                                match = false;
                                break;
                            }
                        }
                    }
                    if (!match) continue;
                }

                var rawId = GetRowValue(row, idColumn);
                if (string.IsNullOrWhiteSpace(rawId))
                {
                    continue;
                }

                var partId = ComposePartId(category.Id, rawId);
                var name = GetFieldValue(sheet, row, "PartNumber");
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = rawId;
                }

                var description = GetFieldValue(sheet, row, "Description");
                var symbolIdStr = GetFieldValue(sheet, row, "Symbol");
                
                if (!string.IsNullOrWhiteSpace(symbolIdStr) && !string.IsNullOrWhiteSpace(config.SymbolPrefix))
                {
                    symbolIdStr = $"{config.SymbolPrefix}:{symbolIdStr}";
                }
                
                parts.Add(new { id = partId, name, description, symbolIdStr });
            }

            return parts;
        }

        public object? GetPartDetails(string partId)
        {
            if (string.IsNullOrWhiteSpace(partId))
            {
                return null;
            }
            var (library, rawId) = SplitPartId(partId);
            
            if (!_categories.TryGetValue(library, out var category))
            {
                return null;
            }

            var sheet = category.Mapping;
            var idMapping = FindFieldMapping(sheet, "ID");
            if (idMapping?.ColumnIndex is not int idColumn || idColumn <= 0)
            {
                return null;
            }

            var key = GetDataKey(sheet.SourceFile, sheet.SheetName);
            if (!_data.TryGetValue(key, out var sheetData))
            {
                return null;
            }

            var config = ConfigurationManager.Load();

            foreach (var row in sheetData)
            {
                // Check filters
                if (category.Filters != null && category.Filters.Count > 0)
                {
                    var match = true;
                    foreach (var filter in category.Filters)
                    {
                        var field = FindFieldMapping(sheet, filter.Key);
                        if (field?.ColumnIndex is int colIdx && colIdx > 0)
                        {
                            var rowVal = GetRowValue(row, colIdx);
                            if (!string.Equals(rowVal, filter.Value, StringComparison.OrdinalIgnoreCase))
                            {
                                match = false;
                                break;
                            }
                        }
                    }
                    if (!match) continue;
                }

                var candidateId = GetRowValue(row, idColumn);
                // Compare slugified candidate id to the slug provided in the request
                var candidateSlug = Slugify(candidateId);
                if (!string.Equals(candidateSlug, rawId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var fields = BuildFieldDictionary(sheet, row);

                var partName = GetFieldValue(sheet, row, "PartNumber");
                if (string.IsNullOrWhiteSpace(partName))
                {
                    partName = candidateId;
                }

                var symbolIdStr = GetFieldValue(sheet, row, "Symbol");
                if (!string.IsNullOrWhiteSpace(symbolIdStr) && !string.IsNullOrWhiteSpace(config.SymbolPrefix))
                {
                    symbolIdStr = $"{config.SymbolPrefix}:{symbolIdStr}";
                }

                var exclude = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "PartNumber", "ID", "Symbol", "Footprint" };
                var fieldsPayload = fields
                    .Where(kvp => !exclude.Contains(kvp.Key))
                    .ToDictionary(
                        kvp => kvp.Key,
                        kvp => new { value = ApplyPrefixIfNeeded(kvp.Key, kvp.Value.value, config), visible = kvp.Value.visible ? "True" : "False" },
                        StringComparer.OrdinalIgnoreCase);

                var footprintValue = GetFieldValue(sheet, row, "Footprint");
                if (!string.IsNullOrWhiteSpace(footprintValue) && !string.IsNullOrWhiteSpace(config.FootprintPrefix))
                {
                    footprintValue = $"{config.FootprintPrefix}:{footprintValue}";
                }
                
                if (fields.ContainsKey("Footprint"))
                {
                    var footprintVisible = fields["Footprint"].visible;
                    fieldsPayload["Footprint"] = new { value = footprintValue, visible = footprintVisible ? "True" : "False" };
                }

                return new
                {
                    id = ComposePartId(category.Id, candidateId),
                    name = partName,
                    symbolIdStr = symbolIdStr,
                    exclude_from_bom = "False",
                    exclude_from_board = "False",
                    exclude_from_sim = "False",
                    fields = fieldsPayload
                };
            }

            return null;
        }

        private static FieldMapping? FindFieldMapping(SheetMapping mapping, string fieldName)
        {
            return mapping.FieldMappings?.FirstOrDefault(m => string.Equals(m.FieldName, fieldName, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetFieldValue(SheetMapping mapping, Dictionary<int, string> row, string fieldName)
        {
            var field = FindFieldMapping(mapping, fieldName);
            if (field?.ColumnIndex is int column && column > 0)
            {
                return GetRowValue(row, column);
            }

            return string.Empty;
        }

        private static Dictionary<string, (string value, bool visible)> BuildFieldDictionary(SheetMapping mapping, Dictionary<int, string> row)
        {
            var result = new Dictionary<string, (string value, bool visible)>(StringComparer.OrdinalIgnoreCase);
            foreach (var field in mapping.FieldMappings)
            {
                if (!field.ColumnIndex.HasValue || field.ColumnIndex.Value <= 0)
                {
                    continue;
                }

                var value = GetRowValue(row, field.ColumnIndex.Value);
                result[field.FieldName] = (value, field.Visible);
            }

            return result;
        }

        private static string GetRowValue(Dictionary<int, string> row, int columnIndex)
        {
            return row.TryGetValue(columnIndex, out var value) ? value ?? string.Empty : string.Empty;
        }

        private static string ComposePartId(string library, string rawId)
        {
            var lib = library?.Trim() ?? string.Empty;
            var id = rawId?.Trim() ?? string.Empty;
            // Slugify the part id to ensure URL-safe identifiers (no spaces/unsafe chars)
            var slug = Slugify(id);
            return string.IsNullOrEmpty(lib) ? slug : $"{lib}:{slug}";
        }

        private static (string library, string rawId) SplitPartId(string partId)
        {
            var trimmed = partId.Trim();
            var separatorIndex = trimmed.IndexOf(':');
            if (separatorIndex <= 0 || separatorIndex >= trimmed.Length - 1)
            {
                return (trimmed, trimmed);
            }

            var library = trimmed.Substring(0, separatorIndex);
            var rawId = trimmed.Substring(separatorIndex + 1);
            return (library, rawId);
        }

        private static string Slugify(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return string.Empty;
            var s = input.Trim().ToLowerInvariant();
            // Replace non-letter/digit characters with hyphen
            s = Regex.Replace(s, "[^a-z0-9]+", "-");
            // Collapse multiple hyphens
            s = Regex.Replace(s, "-+", "-");
            // Trim leading/trailing hyphens
            s = s.Trim('-');
            return s;
        }

        private static string ApplyPrefixIfNeeded(string fieldName, string value, AppConfiguration config)
        {
            // This helper is for fields in the payload; Footprint is handled separately
            // but included here for consistency in case it's in the payload
            if (string.IsNullOrWhiteSpace(value)) return value;
            
            if (string.Equals(fieldName, "Footprint", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(config.FootprintPrefix))
            {
                return $"{config.FootprintPrefix}:{value}";
            }
            
            return value;
        }
    }
}

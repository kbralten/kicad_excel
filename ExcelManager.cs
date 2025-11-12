using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace KiCadExcelBridge
{
    public class ExcelManager
    {
        private readonly List<string> _filePaths;
        private readonly List<SheetMapping> _sheetMappings;
        // Key: fullPath + "::" + sheetName
        private readonly Dictionary<string, List<Dictionary<int, string>>> _data = new();

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
        }

        private static List<Dictionary<int, string>> LoadWorksheetRows(string filePath, SheetMapping mapping)
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[mapping.SheetName];
            var rows = new List<Dictionary<int, string>>();
            if (worksheet == null || worksheet.Dimension == null)
            {
                return rows;
            }

            var endRow = worksheet.Dimension.End.Row;
            var endColumn = worksheet.Dimension.End.Column;
            var startRow = mapping.IgnoreHeader ? 2 : 1;
            if (startRow > endRow)
            {
                return rows;
            }

            for (int row = startRow; row <= endRow; row++)
            {
                var rowData = new Dictionary<int, string>();
                for (int col = 1; col <= endColumn; col++)
                {
                    rowData[col] = worksheet.Cells[row, col].Text?.Trim() ?? string.Empty;
                }
                rows.Add(rowData);
            }

            return rows;
        }

        private static List<Dictionary<int, string>> LoadCsvRows(string filePath, SheetMapping mapping)
        {
            var rows = new List<Dictionary<int, string>>();
            var lines = File.ReadAllLines(filePath);
            var startIndex = mapping.IgnoreHeader ? 1 : 0;
            for (int i = startIndex; i < lines.Length; i++)
            {
                var values = lines[i].Split(',');
                var rowData = new Dictionary<int, string>();
                for (int col = 0; col < values.Length; col++)
                {
                    rowData[col + 1] = values[col].Trim();
                }
                rows.Add(rowData);
            }

            return rows;
        }

        private static string GetDataKey(string filePath, string sheetName) => filePath + "::" + sheetName;

        public IEnumerable<object> GetCategories()
        {
            var categories = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var mapping in _sheetMappings)
            {
                var label = mapping.CategoryLabel?.Trim();
                if (string.IsNullOrWhiteSpace(label))
                {
                    continue;
                }

                if (!categories.ContainsKey(label))
                {
                    categories[label] = label;
                }
            }

            return categories.Select(pair => new { id = pair.Key, name = pair.Value, description = string.Empty });
        }

        public IEnumerable<object> GetPartsForCategory(string categoryId)
        {
            if (string.IsNullOrWhiteSpace(categoryId))
            {
                return Enumerable.Empty<object>();
            }

            var parts = new List<object>();
            foreach (var sheet in _sheetMappings)
            {
                var label = sheet.CategoryLabel?.Trim();
                if (string.IsNullOrWhiteSpace(label) || !string.Equals(label, categoryId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var idMapping = FindFieldMapping(sheet, "ID");
                if (idMapping?.ColumnIndex is not int idColumn || idColumn <= 0)
                {
                    continue;
                }

                var key = GetDataKey(sheet.SourceFile, sheet.SheetName);
                if (!_data.TryGetValue(key, out var sheetData))
                {
                    continue;
                }

                foreach (var row in sheetData)
                {
                    var rawId = GetRowValue(row, idColumn);
                    if (string.IsNullOrWhiteSpace(rawId))
                    {
                        continue;
                    }

                    var partId = ComposePartId(label, rawId);
                    var name = GetFieldValue(sheet, row, "PartNumber");
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        name = rawId;
                    }

                    var description = GetFieldValue(sheet, row, "Description");
                    var symbolIdStr = GetFieldValue(sheet, row, "Symbol");
                    parts.Add(new { id = partId, name, description, symbolIdStr });
                }
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
            foreach (var sheet in _sheetMappings)
            {
                var label = sheet.CategoryLabel?.Trim();
                if (string.IsNullOrWhiteSpace(label) || !string.Equals(label, library, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var idMapping = FindFieldMapping(sheet, "ID");
                if (idMapping?.ColumnIndex is not int idColumn || idColumn <= 0)
                {
                    continue;
                }

                var key = GetDataKey(sheet.SourceFile, sheet.SheetName);
                if (!_data.TryGetValue(key, out var sheetData))
                {
                    continue;
                }

                foreach (var row in sheetData)
                {
                    var candidateId = GetRowValue(row, idColumn);
                    if (!string.Equals(candidateId, rawId, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var fields = BuildFieldDictionary(sheet, row);

                    // name should come from PartNumber (fall back to rawId)
                    var partName = GetFieldValue(sheet, row, "PartNumber");
                    if (string.IsNullOrWhiteSpace(partName))
                    {
                        partName = rawId;
                    }

                    // symbolIdStr should be taken from the "Symbol" field
                    var symbolIdStr = GetFieldValue(sheet, row, "Symbol");

                    // Exclude PartNumber, ID and Symbol from the fields payload since they are provided separately
                    var exclude = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "PartNumber", "ID", "Symbol" };
                    var fieldsPayload = fields
                        .Where(kvp => !exclude.Contains(kvp.Key))
                        .ToDictionary(
                            kvp => kvp.Key,
                            kvp => new { value = kvp.Value.value, visible = kvp.Value.visible ? "True" : "False" },
                            StringComparer.OrdinalIgnoreCase);

                    return new
                    {
                        id = ComposePartId(label, candidateId),
                        name = partName,
                        symbolIdStr = symbolIdStr,
                        exclude_from_bom = "False",
                        exclude_from_board = "False",
                        exclude_from_sim = "False",
                        fields = fieldsPayload
                    };
                }
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
            return string.IsNullOrEmpty(lib) ? id : $"{lib}:{id}";
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
    }
}

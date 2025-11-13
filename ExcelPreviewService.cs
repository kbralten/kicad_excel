using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;

namespace KiCadExcelBridge
{
    internal static class ExcelPreviewService
    {
        public static ExcelPreviewResult LoadPreview(string filePath, string sheetName, bool ignoreHeader, int maxRows)
        {
            var columns = new List<ExcelColumnOption>();
            var rows = new List<Dictionary<int, string>>();

            if (!File.Exists(filePath))
            {
                return new ExcelPreviewResult(columns, rows);
            }

            try
            {
                var extension = Path.GetExtension(filePath);
                if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    var stream = FileHelpers.OpenFileForReadWithFallback(filePath, maxRetries: 3, out var actualPath, out var usedFallback);
                    if (stream == null)
                    {
                        return new ExcelPreviewResult(columns, rows);
                    }

                    using (stream)
                    using (var package = new ExcelPackage(stream))
                    {
                        var worksheet = package.Workbook.Worksheets[sheetName];
                        if (worksheet == null || worksheet.Dimension == null)
                        {
                            return new ExcelPreviewResult(columns, rows);
                        }

                        var endColumn = worksheet.Dimension.End.Column;
                        var startRow = ignoreHeader ? 2 : 1;
                        var endRow = Math.Min(worksheet.Dimension.End.Row, startRow + Math.Max(maxRows, 1) - 1);

                        for (int col = 1; col <= endColumn; col++)
                        {
                            var header = ignoreHeader ? worksheet.Cells[1, col].Text?.Trim() : null;
                            header = string.IsNullOrWhiteSpace(header) ? null : header;
                            columns.Add(new ExcelColumnOption
                            {
                                Index = col,
                                Letter = ColumnIndexToLetter(col),
                                Header = header
                            });
                        }

                        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                        {
                            var rowDict = new Dictionary<int, string>();
                            for (int col = 1; col <= endColumn; col++)
                            {
                                rowDict[col] = worksheet.Cells[rowIndex, col].Text;
                            }
                            rows.Add(rowDict);
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
                        // ignore deletion errors
                    }
                }
                else if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    var maxLines = ignoreHeader ? maxRows + 1 : maxRows;
                    var takeCount = maxLines > 0 ? maxLines : int.MaxValue;
                    var lines = new List<string>();
                    using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var sr = new StreamReader(fs))
                    {
                        for (int i = 0; i < takeCount && !sr.EndOfStream; i++)
                        {
                            lines.Add(sr.ReadLine() ?? string.Empty);
                        }
                    }
                    if (lines.Count == 0)
                    {
                        return new ExcelPreviewResult(columns, rows);
                    }

                    var headers = new List<string>();
                    var firstLineValues = SplitCsvLine(lines[0]);
                    if (ignoreHeader)
                    {
                        headers = firstLineValues;
                    }
                    else
                    {
                        headers = Enumerable.Repeat(string.Empty, firstLineValues.Count).ToList();
                    }

                    for (int col = 0; col < headers.Count; col++)
                    {
                        var header = ignoreHeader ? headers[col]?.Trim() : null;
                        header = string.IsNullOrWhiteSpace(header) ? null : header;
                        columns.Add(new ExcelColumnOption
                        {
                            Index = col + 1,
                            Letter = ColumnIndexToLetter(col + 1),
                            Header = header
                        });
                    }

                    var startIndex = ignoreHeader ? 1 : 0;
                    var totalRows = 0;
                    for (int i = startIndex; i < lines.Count && totalRows < maxRows; i++)
                    {
                        var values = SplitCsvLine(lines[i]);
                        var rowDict = new Dictionary<int, string>();
                        var columnCount = Math.Max(columns.Count, values.Count);
                        for (int col = 0; col < columnCount; col++)
                        {
                            if (col >= columns.Count)
                            {
                                columns.Add(new ExcelColumnOption
                                {
                                    Index = col + 1,
                                    Letter = ColumnIndexToLetter(col + 1),
                                    Header = null
                                });
                            }

                            rowDict[col + 1] = col < values.Count ? values[col] : string.Empty;
                        }
                        rows.Add(rowDict);
                        totalRows++;
                    }
                }
            }
            catch
            {
                // Ignore preview errors and return what was collected so far (possibly empty)
            }

            return new ExcelPreviewResult(columns, rows);
        }

        private static string ColumnIndexToLetter(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = new StringBuilder();

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName.Insert(0, Convert.ToChar(65 + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnName.ToString();
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            if (line == null)
            {
                return result;
            }

            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];
                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            result.Add(sb.ToString());
            return result;
        }
    }

    internal sealed class ExcelPreviewResult
    {
        public ExcelPreviewResult(List<ExcelColumnOption> columns, List<Dictionary<int, string>> rows)
        {
            Columns = columns;
            Rows = rows;
        }

        public List<ExcelColumnOption> Columns { get; }
        public List<Dictionary<int, string>> Rows { get; }
    }
}

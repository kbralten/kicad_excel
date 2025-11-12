using System;

namespace KiCadExcelBridge
{
    public class FieldMapping
    {
        public string FieldName { get; set; } = string.Empty;
        public int? ColumnIndex { get; set; }
        public string? ColumnHeader { get; set; }
        public bool Visible { get; set; } = true;
        public string Category { get; set; } = "Common";
        public bool IsRequired { get; set; }

        public bool IsCustom => string.Equals(Category, "Custom", StringComparison.OrdinalIgnoreCase) && !IsRequired;
    }
}

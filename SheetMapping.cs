using System.Collections.Generic;

namespace KiCadExcelBridge
{
    public class SheetMapping
    {
        public string SourceFile { get; set; } = string.Empty;
        public string SheetName { get; set; } = string.Empty;
        public string CategoryLabel { get; set; } = string.Empty;
        public bool IgnoreHeader { get; set; } = true;
        public List<FieldMapping> FieldMappings { get; set; } = new List<FieldMapping>();
    }
}

using System.Collections.Generic;

namespace KiCadExcelBridge
{
    public class AppConfiguration
    {
        public List<string> SourceFiles { get; set; } = new List<string>();
        public List<SheetMapping> SheetMappings { get; set; } = new List<SheetMapping>();
        public string SymbolPrefix { get; set; } = "symbol";
        public string FootprintPrefix { get; set; } = "footprint";
        public int ServerPort { get; set; } = 8088;
    }
}

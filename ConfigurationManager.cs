using System;
using System.IO;
using System.Text.Json;

namespace KiCadExcelBridge
{
    public static class ConfigurationManager
    {
        private static readonly string _configFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "KiCadExcelBridge");
        private static readonly string _configFile = Path.Combine(_configFolder, "config.json");

        public static AppConfiguration Load()
        {
            if (!File.Exists(_configFile))
            {
                return new AppConfiguration();
            }

            try
            {
                var json = File.ReadAllText(_configFile);

                // Parse JSON to detect legacy global IgnoreHeader property so we can migrate
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                bool hasRootIgnore = root.TryGetProperty("IgnoreHeader", out var rootIgnoreElem);
                bool? rootIgnoreValue = hasRootIgnore ? rootIgnoreElem.GetBoolean() : (bool?)null;

                var config = JsonSerializer.Deserialize<AppConfiguration>(json) ?? new AppConfiguration();

                foreach (var sheet in config.SheetMappings)
                {
                    sheet.FieldMappings ??= FieldMappingDefaults.CreateDefaults();
                    FieldMappingDefaults.EnsureDefaults(sheet.FieldMappings);
                }

                // If the legacy root IgnoreHeader exists, but per-sheet IgnoreHeader was not present
                // in the saved SheetMappings entries, apply the root value to those entries.
                if (rootIgnoreValue.HasValue && root.TryGetProperty("SheetMappings", out var sheetMappingsElement) && sheetMappingsElement.ValueKind == JsonValueKind.Array)
                {
                    int i = 0;
                    foreach (var item in sheetMappingsElement.EnumerateArray())
                    {
                        var hasPerSheetIgnore = item.TryGetProperty("IgnoreHeader", out _);
                        if (!hasPerSheetIgnore)
                        {
                            if (i < config.SheetMappings.Count)
                            {
                                config.SheetMappings[i].IgnoreHeader = rootIgnoreValue.Value;
                            }
                        }
                        i++;
                    }
                }

                return config;
            }
            catch
            {
                // Handle deserialization errors, maybe return a default config
                return new AppConfiguration();
            }
        }

        public static void Save(AppConfiguration config)
        {
            try
            {
                if (!Directory.Exists(_configFolder))
                {
                    Directory.CreateDirectory(_configFolder);
                }

                var json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configFile, json);
            }
            catch
            {
                // Handle serialization/IO errors
            }
        }
    }
}

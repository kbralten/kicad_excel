using System;
using System.Collections.Generic;
using System.Linq;

namespace KiCadExcelBridge
{
    internal static class FieldMappingDefaults
    {
        private static readonly List<FieldDefinition> _defaults = new()
        {
            new FieldDefinition("ID", "Required", defaultVisible: false, isRequired: true),
            new FieldDefinition("Symbol", "Required", defaultVisible: true, isRequired: true),
            new FieldDefinition("Footprint", "Required", defaultVisible: true, isRequired: true),
            new FieldDefinition("Value", "Required", defaultVisible: true, isRequired: true),
            new FieldDefinition("PartNumber", "Common", defaultVisible: true, isRequired: false),
            new FieldDefinition("Manufacturer", "Common", defaultVisible: true, isRequired: false),
            new FieldDefinition("Supplier", "Common", defaultVisible: false, isRequired: false),
            new FieldDefinition("Datasheet", "Common", defaultVisible: false, isRequired: false),
            new FieldDefinition("Description", "Common", defaultVisible: true, isRequired: false)
        };

        public static List<FieldMapping> CreateDefaults()
        {
            return _defaults
                .Select(d => new FieldMapping
                {
                    FieldName = d.Name,
                    Category = d.Category,
                    Visible = d.DefaultVisible,
                    IsRequired = d.IsRequired
                })
                .ToList();
        }

        public static void EnsureDefaults(List<FieldMapping> mappings)
        {
            if (mappings == null)
            {
                return;
            }

            foreach (var def in _defaults)
            {
                var existing = mappings.FirstOrDefault(m => string.Equals(m.FieldName, def.Name, StringComparison.OrdinalIgnoreCase));
                if (existing == null)
                {
                    mappings.Add(new FieldMapping
                    {
                        FieldName = def.Name,
                        Category = def.Category,
                        Visible = def.DefaultVisible,
                        IsRequired = def.IsRequired
                    });
                }
                else
                {
                    existing.Category = def.Category;
                    existing.IsRequired = def.IsRequired;
                    if (!existing.ColumnIndex.HasValue)
                    {
                        existing.Visible = def.DefaultVisible;
                    }
                }
            }
        }

        private record FieldDefinition(string Name, string Category, bool defaultVisible, bool isRequired)
        {
            public bool DefaultVisible => defaultVisible;
            public bool IsRequired => isRequired;
        }
    }
}

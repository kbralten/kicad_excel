using System;
using System.Reflection;
using OfficeOpenXml;

namespace KiCadExcelBridge
{
    internal static class EpplusLicenseHelper
    {
        public static void EnsureNonCommercialLicense()
        {
            try
            {
                var excelType = typeof(ExcelPackage);

                // 1) Try new API: ExcelPackage.License may be an enum or an object with methods
                var licenseProp = excelType.GetProperty("License", BindingFlags.Static | BindingFlags.Public);
                if (licenseProp != null)
                {
                    var propType = licenseProp.PropertyType;

                    // If it's an enum, set it directly
                    if (propType.IsEnum)
                    {
                        var enumValue = Enum.Parse(propType, "NonCommercial", ignoreCase: true);
                        licenseProp.SetValue(null, enumValue);
                        return;
                    }

                    // If it's an object, try to find a method that accepts a license enum or string
                    var licenseObj = licenseProp.GetValue(null);
                    if (licenseObj != null)
                    {
                        var licType = licenseObj.GetType();
                        // Look for methods like Set, SetLicense, SetLicenseContext, Register, or SetLicenseKey
                        var candidates = licType.GetMethods(BindingFlags.Instance | BindingFlags.Public)
                            .Where(m => m.Name.IndexOf("Set", StringComparison.OrdinalIgnoreCase) >= 0 || m.Name.IndexOf("Register", StringComparison.OrdinalIgnoreCase) >= 0)
                            .ToArray();

                        foreach (var method in candidates)
                        {
                            var parameters = method.GetParameters();
                            if (parameters.Length == 1)
                            {
                                var pType = parameters[0].ParameterType;
                                try
                                {
                                    if (pType.IsEnum)
                                    {
                                        var enumValue = Enum.Parse(pType, "NonCommercial", ignoreCase: true);
                                        method.Invoke(licenseObj, new object[] { enumValue });
                                        return;
                                    }
                                    else if (pType == typeof(string))
                                    {
                                        method.Invoke(licenseObj, new object[] { "NonCommercial" });
                                        return;
                                    }
                                }
                                catch
                                {
                                    // ignore and try next candidate
                                }
                            }
                        }
                    }
                }

                // 2) Fallback to older API: ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                var oldProp = excelType.GetProperty("LicenseContext", BindingFlags.Static | BindingFlags.Public);
                if (oldProp != null)
                {
                    var oldType = oldProp.PropertyType;
                    if (oldType.IsEnum)
                    {
                        var enumValue = Enum.Parse(oldType, "NonCommercial", ignoreCase: true);
                        oldProp.SetValue(null, enumValue);
                        return;
                    }
                }

                // 3) As a last attempt, try to locate an enum type named LicenseContext in the assembly and assign via License property
                var asm = excelType.Assembly;
                var licenseContextType = asm.GetType("OfficeOpenXml.LicenseContext") ?? asm.GetType("OfficeOpenXml.Licensing.LicenseContext");
                if (licenseContextType != null && licenseContextType.IsEnum && licenseProp != null)
                {
                    var enumValue = Enum.Parse(licenseContextType, "NonCommercial", ignoreCase: true);
                    if (licenseProp.PropertyType.IsAssignableFrom(licenseContextType))
                    {
                        licenseProp.SetValue(null, enumValue);
                        return;
                    }
                }
            }
            catch
            {
                // If anything goes wrong, swallow — the caller will receive EPPlus' own exception with guidance.
            }
        }
    }
}

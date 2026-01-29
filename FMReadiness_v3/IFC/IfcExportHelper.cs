using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Revit.DB;

namespace FMReadiness_v3.IFC
{
    public static class IfcExportHelper
    {
        private const string PsetFileName = "FMReadiness_Psets.txt";
        private const string PsetFolderName = "IFC";

        public static string? EnsureUserDefinedPsetFile()
        {
            try
            {
                var assemblyDir = Path.GetDirectoryName(typeof(IfcExportHelper).Assembly.Location) ?? string.Empty;
                var folderPath = Path.Combine(assemblyDir, PsetFolderName);
                var filePath = Path.Combine(folderPath, PsetFileName);

                var desiredContent = GetDefaultPsetContent();
                if (File.Exists(filePath))
                {
                    var existing = File.ReadAllText(filePath);
                    if (!IsCompatiblePsetContent(existing))
                    {
                        File.WriteAllText(filePath, desiredContent);
                    }

                    return filePath;
                }

                Directory.CreateDirectory(folderPath);
                File.WriteAllText(filePath, desiredContent);
                return filePath;
            }
            catch
            {
                return null;
            }
        }

        public static void ConfigureIfc4Options(IFCExportOptions options, string userPsetPath, bool includeRevitPropertySets)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            // IFC4 Design Transfer View is more permissive for MEP export than IFC4 RV.
            options.FileVersion = IFCVersion.IFC4DTV;
            options.ExportBaseQuantities = includeRevitPropertySets;

            // Prefer strongly-typed properties when available (varies by IFC exporter build).
            TrySetOptionProperty(options, "ExportInternalRevitPropertySets", includeRevitPropertySets);
            TrySetOptionProperty(options, "ExportIFCCommonPropertySets", includeRevitPropertySets);
            TrySetOptionProperty(options, "ExportMaterialPsets", includeRevitPropertySets);
            TrySetOptionProperty(options, "ExportUserDefinedPsets", true);
            TrySetOptionProperty(options, "ExportUserDefinedPsetsFileName", userPsetPath);
            TrySetOptionProperty(options, "ExportUserDefinedParameterMapping", false);
            TrySetOptionProperty(options, "ExportUserDefinedParameterMappingFileName", string.Empty);
            TrySetOptionProperty(options, "ExportSchedulesAsPsets", false);

            // IFC Exporter expects these option keys (same ones used by the UI exporter).
            options.AddOption("ExportUserDefinedPsets", "true");
            options.AddOption("ExportUserDefinedPsetsFileName", userPsetPath);

            // Explicitly disable parameter mapping unless needed later.
            options.AddOption("ExportUserDefinedParameterMapping", "false");
            options.AddOption("ExportUserDefinedParameterMappingFileName", string.Empty);

            options.AddOption("ExportInternalRevitPropertySets", includeRevitPropertySets ? "true" : "false");
            options.AddOption("ExportIFCCommonPropertySets", includeRevitPropertySets ? "true" : "false");
            options.AddOption("ExportMaterialPsets", includeRevitPropertySets ? "true" : "false");
            // Avoid exporter writes that require an open transaction.
            options.AddOption("StoreIFCGUID", "false");

            // Force plain IFC to simplify downstream verification.
            options.AddOption("FileType", "IFC");
            options.AddOption("IFCFileType", "IFC");
        }

        public static string GetOptionPropertyDiagnostics(IFCExportOptions options)
        {
            if (options == null) return "options=null";

            var names = new[]
            {
                "ExportInternalRevitPropertySets",
                "ExportIFCCommonPropertySets",
                "ExportMaterialPsets",
                "ExportBaseQuantities",
                "ExportUserDefinedPsets",
                "ExportUserDefinedPsetsFileName",
                "ExportUserDefinedParameterMapping",
                "ExportUserDefinedParameterMappingFileName",
                "ExportSchedulesAsPsets"
            };

            var parts = new List<string>(names.Length);
            foreach (var name in names)
            {
                var prop = options.GetType().GetProperty(name);
                if (prop == null)
                {
                    parts.Add($"{name}=<missing>");
                    continue;
                }

                object? value = null;
                try
                {
                    value = prop.GetValue(options);
                }
                catch
                {
                    value = "<error>";
                }

                parts.Add($"{name}={(value ?? "<null>")}");
            }

            return string.Join("; ", parts);
        }

        private static bool TrySetOptionProperty(IFCExportOptions options, string propertyName, object? value)
        {
            try
            {
                var prop = options.GetType().GetProperty(propertyName);
                if (prop == null || !prop.CanWrite) return false;

                var targetType = prop.PropertyType;
                if (value == null)
                {
                    if (!targetType.IsValueType || Nullable.GetUnderlyingType(targetType) != null)
                    {
                        prop.SetValue(options, null);
                        return true;
                    }

                    return false;
                }

                if (targetType.IsInstanceOfType(value))
                {
                    prop.SetValue(options, value);
                    return true;
                }

                if (targetType == typeof(bool))
                {
                    prop.SetValue(options, Convert.ToBoolean(value));
                    return true;
                }

                if (targetType == typeof(string))
                {
                    prop.SetValue(options, value.ToString());
                    return true;
                }

                if (targetType.IsEnum)
                {
                    var parsed = Enum.Parse(targetType, value.ToString() ?? string.Empty, true);
                    prop.SetValue(options, parsed);
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public static string GetDefaultPsetContent()
        {
            return string.Join(Environment.NewLine, new[]
            {
                "PropertySet:\tFMReadiness\tI\tIfcDistributionElement",
                "FM_Barcode\tText\tFM_Barcode",
                "FM_UniqueAssetId\tText\tFM_UniqueAssetId",
                "FM_InstallationDate\tText\tFM_InstallationDate",
                "FM_WarrantyStart\tText\tFM_WarrantyStart",
                "FM_WarrantyEnd\tText\tFM_WarrantyEnd",
                "FM_Criticality\tText\tFM_Criticality",
                "FM_Trade\tText\tFM_Trade",
                "FM_PMTemplateId\tText\tFM_PMTemplateId",
                "FM_PMFrequencyDays\tInteger\tFM_PMFrequencyDays",
                "FM_Building\tText\tFM_Building",
                "FM_LocationSpace\tText\tFM_LocationSpace",
                "",
                "PropertySet:\tFMReadinessType\tT\tIfcDistributionElementType",
                "Manufacturer\tText\tManufacturer",
                "Model\tText\tModel",
                "TypeMark\tText\tType Mark"
            });
        }

        private static bool IsCompatiblePsetContent(string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return false;

            return HasRequiredLine(content, "PropertySet:", "FMReadiness\tI\tIfcDistributionElement")
                && HasRequiredLine(content, "FM_Barcode", "FM_Barcode\tText\tFM_Barcode")
                && HasRequiredLine(content, "PropertySet:", "FMReadinessType\tT\tIfcDistributionElementType")
                && HasRequiredLine(content, "TypeMark", "TypeMark\tText\tType Mark");
        }

        private static bool HasRequiredLine(string content, string startsWith, string requiredSubstring)
        {
            var lines = content.Split(new[] { "\r\n", "\n", "\r" }, StringSplitOptions.None);
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmed))
                {
                    continue;
                }

                if (trimmed.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                if (trimmed.StartsWith(startsWith, StringComparison.Ordinal)
                    && trimmed.Contains(requiredSubstring, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }
    }
}

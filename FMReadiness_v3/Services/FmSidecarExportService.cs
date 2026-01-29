using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using Autodesk.Revit.DB;

namespace FMReadiness_v3.Services
{
    /// <summary>
    /// Service to export FM parameters as a sidecar JSON file.
    /// The sidecar file is keyed by IFC GlobalId and can be merged
    /// with the DigitalTwin viewer's metadata.json.
    /// </summary>
    public class FmSidecarExportService
    {
        private readonly Document _doc;

        public FmSidecarExportService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// Export FM parameters for all elements to a sidecar JSON file.
        /// Uses the stored IFC GUID parameter if available (from "Store IFC GUID" export option),
        /// otherwise falls back to computing it from the Revit UniqueId.
        /// </summary>
        /// <param name="outputPath">Path for the output .fm_params.json file</param>
        /// <returns>Export result with statistics</returns>
        public ExportResult Export(string outputPath)
        {
            var result = new ExportResult();
            var sidecarData = new Dictionary<string, FmElementData>();

            // Get all FM-relevant elements
            var collectorService = new CollectorService(_doc);
            var elements = collectorService.GetAllFmElements();
            result.TotalElements = elements.Count;

            foreach (var element in elements)
            {
                try
                {
                    // Try to get stored IFC GUID first (from "Store IFC GUID" export option)
                    var ifcGlobalId = GetStoredIfcGuid(element);
                    bool usedStoredGuid = !string.IsNullOrEmpty(ifcGlobalId);

                    // Fall back to computed GUID if not stored
                    if (!usedStoredGuid)
                    {
                        ifcGlobalId = ConvertToIfcGuid(element.UniqueId);
                    }

                    if (string.IsNullOrEmpty(ifcGlobalId))
                    {
                        result.SkippedNoGuid++;
                        continue;
                    }

                    if (usedStoredGuid)
                        result.StoredGuidCount++;
                    else
                        result.ComputedGuidCount++;

                    // Get element type for type parameters
                    var typeId = element.GetTypeId();
                    ElementType? elementType = null;
                    if (typeId != ElementId.InvalidElementId)
                        elementType = _doc.GetElement(typeId) as ElementType;

                    // Collect FM instance parameters
                    var fmParams = new FmParameters
                    {
                        FM_Barcode = GetParamValue(element, "FM_Barcode"),
                        FM_UniqueAssetId = GetParamValue(element, "FM_UniqueAssetId"),
                        FM_InstallationDate = GetParamValue(element, "FM_InstallationDate"),
                        FM_WarrantyStart = GetParamValue(element, "FM_WarrantyStart"),
                        FM_WarrantyEnd = GetParamValue(element, "FM_WarrantyEnd"),
                        FM_Criticality = GetParamValue(element, "FM_Criticality"),
                        FM_Trade = GetParamValue(element, "FM_Trade"),
                        FM_PMTemplateId = GetParamValue(element, "FM_PMTemplateId"),
                        FM_PMFrequencyDays = GetParamValue(element, "FM_PMFrequencyDays"),
                        FM_Building = GetParamValue(element, "FM_Building"),
                        FM_LocationSpace = GetParamValue(element, "FM_LocationSpace")
                    };

                    // Collect type parameters
                    var typeParams = new FmTypeParameters();
                    if (elementType != null)
                    {
                        typeParams.Manufacturer = GetParamValue(elementType, "Manufacturer");
                        typeParams.Model = GetParamValue(elementType, "Model");
                        typeParams.TypeMark = GetParamValue(elementType, BuiltInParameter.ALL_MODEL_TYPE_MARK);
                    }

                    // Only include elements that have at least one FM parameter set
                    if (HasAnyValue(fmParams) || HasAnyValue(typeParams))
                    {
                        var elementData = new FmElementData
                        {
                            FMReadiness = fmParams,
                            FMReadinessType = typeParams,
                            _meta = new ElementMeta
                            {
                                RevitElementId = GetElementIdValue(element.Id),
                                RevitUniqueId = element.UniqueId,
                                Category = element.Category?.Name,
                                Family = elementType?.FamilyName,
                                TypeName = elementType?.Name
                            }
                        };

                        sidecarData[ifcGlobalId] = elementData;
                        result.ExportedElements++;
                    }
                    else
                    {
                        result.SkippedNoData++;
                    }
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"Element {element.Id}: {ex.Message}");
                }
            }

            // Write JSON file
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                PropertyNamingPolicy = null // Keep property names as-is
            };

            var json = JsonSerializer.Serialize(sidecarData, options);
            File.WriteAllText(outputPath, json);

            result.OutputPath = outputPath;
            result.FileSizeBytes = new FileInfo(outputPath).Length;

            return result;
        }

        /// <summary>
        /// Get the stored IFC GUID from the element's IfcGUID parameter.
        /// This parameter is populated when exporting to IFC with "Store IFC GUID" enabled.
        /// </summary>
        private static string? GetStoredIfcGuid(Element element)
        {
            // Try the built-in IFC_GUID parameter
            var param = element.get_Parameter(BuiltInParameter.IFC_GUID);
            if (param != null && param.HasValue && param.StorageType == StorageType.String)
            {
                var value = param.AsString();
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            // Also try looking up by name (in case it's a shared parameter)
            param = element.LookupParameter("IfcGUID");
            if (param != null && param.HasValue && param.StorageType == StorageType.String)
            {
                var value = param.AsString();
                if (!string.IsNullOrWhiteSpace(value))
                    return value;
            }

            return null;
        }

        /// <summary>
        /// Convert Revit UniqueId to IFC GlobalId (22-character base64 encoding).
        /// This is the standard algorithm used by Revit's IFC exporter.
        /// NOTE: This is a fallback method. Prefer using the stored IFC GUID parameter
        /// when available (from "Store IFC GUID" export option).
        /// </summary>
        private static string? ConvertToIfcGuid(string uniqueId)
        {
            if (string.IsNullOrEmpty(uniqueId) || uniqueId.Length < 36)
                return null;

            try
            {
                // Revit UniqueId format: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX-XXXXXXXX
                // First 36 chars are the EpisodeId GUID, last 8 are element-specific
                var episodeIdStr = uniqueId.Substring(0, 36);
                var elementIdStr = uniqueId.Substring(37, 8);

                var episodeId = new Guid(episodeIdStr);
                var elementId = int.Parse(elementIdStr, System.Globalization.NumberStyles.HexNumber);

                // XOR the element ID into the last 4 bytes of the GUID
                var guidBytes = episodeId.ToByteArray();
                var elementBytes = BitConverter.GetBytes(elementId);

                // XOR element ID with last 4 bytes of GUID (bytes 12-15)
                for (int i = 0; i < 4; i++)
                {
                    guidBytes[12 + i] ^= elementBytes[i];
                }

                // Convert to IFC base64 encoding
                return ConvertToIfcBase64(guidBytes);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Convert 16-byte GUID to IFC 22-character base64 string.
        /// IFC uses a custom base64 alphabet.
        /// </summary>
        private static string ConvertToIfcBase64(byte[] guid)
        {
            // IFC base64 alphabet (different from standard base64)
            const string base64Chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_$";

            // Convert GUID bytes to a 128-bit number, then encode as base64
            // Output is always 22 characters

            var result = new char[22];
            int n = 0;

            // Process 3 bytes at a time (24 bits) -> 4 base64 chars (6 bits each)
            // 16 bytes = 5 groups of 3 + 1 byte remaining

            // First 2 bytes -> 4 chars (with padding)
            uint num = (uint)((guid[0] << 16) | (guid[1] << 8));
            result[n++] = base64Chars[(int)((num >> 18) & 0x3F)];
            result[n++] = base64Chars[(int)((num >> 12) & 0x3F)];
            result[n++] = base64Chars[(int)((num >> 6) & 0x3F)];
            result[n++] = base64Chars[(int)(num & 0x3F)];

            // Remaining 14 bytes in groups of 3
            for (int i = 2; i < 16; i += 3)
            {
                if (i + 2 < 16)
                {
                    num = (uint)((guid[i] << 16) | (guid[i + 1] << 8) | guid[i + 2]);
                    result[n++] = base64Chars[(int)((num >> 18) & 0x3F)];
                    result[n++] = base64Chars[(int)((num >> 12) & 0x3F)];
                    result[n++] = base64Chars[(int)((num >> 6) & 0x3F)];
                    result[n++] = base64Chars[(int)(num & 0x3F)];
                }
                else if (i + 1 < 16)
                {
                    num = (uint)((guid[i] << 16) | (guid[i + 1] << 8));
                    result[n++] = base64Chars[(int)((num >> 18) & 0x3F)];
                    result[n++] = base64Chars[(int)((num >> 12) & 0x3F)];
                    result[n++] = base64Chars[(int)((num >> 6) & 0x3F)];
                }
                else
                {
                    num = (uint)(guid[i] << 16);
                    result[n++] = base64Chars[(int)((num >> 18) & 0x3F)];
                    result[n++] = base64Chars[(int)((num >> 12) & 0x3F)];
                }
            }

            return new string(result, 0, 22);
        }

        private static string? GetParamValue(Element element, string paramName)
        {
            var param = element.LookupParameter(paramName);
            return ExtractParamValue(param);
        }

        private static string? GetParamValue(Element element, BuiltInParameter bip)
        {
            var param = element.get_Parameter(bip);
            return ExtractParamValue(param);
        }

        private static string? ExtractParamValue(Parameter? param)
        {
            if (param == null || !param.HasValue)
                return null;

            return param.StorageType switch
            {
                StorageType.String => param.AsString(),
                StorageType.Integer => param.AsInteger().ToString(),
                StorageType.Double => param.AsDouble().ToString(),
                StorageType.ElementId => param.AsElementId()?.ToString(),
                _ => param.AsValueString()
            };
        }

        private static bool HasAnyValue(object obj)
        {
            if (obj == null) return false;

            foreach (var prop in obj.GetType().GetProperties())
            {
                if (prop.Name.StartsWith("_")) continue; // Skip meta properties
                var value = prop.GetValue(obj);
                if (value != null && !string.IsNullOrWhiteSpace(value.ToString()))
                    return true;
            }
            return false;
        }

        private static int GetElementIdValue(ElementId id)
        {
#if REVIT2024_OR_GREATER
            return (int)id.Value;
#else
            return id.IntegerValue;
#endif
        }
    }

    #region Data Models

    public class ExportResult
    {
        public string OutputPath { get; set; } = string.Empty;
        public int TotalElements { get; set; }
        public int ExportedElements { get; set; }
        public int SkippedNoGuid { get; set; }
        public int SkippedNoData { get; set; }
        public int StoredGuidCount { get; set; }
        public int ComputedGuidCount { get; set; }
        public long FileSizeBytes { get; set; }
        public List<string> Errors { get; } = new List<string>();
    }

    public class FmElementData
    {
        public FmParameters? FMReadiness { get; set; }
        public FmTypeParameters? FMReadinessType { get; set; }
        public ElementMeta? _meta { get; set; }
    }

    public class FmParameters
    {
        public string? FM_Barcode { get; set; }
        public string? FM_UniqueAssetId { get; set; }
        public string? FM_InstallationDate { get; set; }
        public string? FM_WarrantyStart { get; set; }
        public string? FM_WarrantyEnd { get; set; }
        public string? FM_Criticality { get; set; }
        public string? FM_Trade { get; set; }
        public string? FM_PMTemplateId { get; set; }
        public string? FM_PMFrequencyDays { get; set; }
        public string? FM_Building { get; set; }
        public string? FM_LocationSpace { get; set; }
    }

    public class FmTypeParameters
    {
        public string? Manufacturer { get; set; }
        public string? Model { get; set; }
        public string? TypeMark { get; set; }
    }

    public class ElementMeta
    {
        public int RevitElementId { get; set; }
        public string? RevitUniqueId { get; set; }
        public string? Category { get; set; }
        public string? Family { get; set; }
        public string? TypeName { get; set; }
    }

    #endregion
}

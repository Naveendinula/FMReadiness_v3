using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;

namespace FMReadiness_v3.Services
{
    /// <summary>
    /// Service for resolving COBie field values with alias fallback and computed value support.
    /// Implements the mapping layer between Revit parameters and COBie field names.
    /// </summary>
    public class CobieMappingService
    {
        private readonly PresetService _presetService;
        private readonly bool _writeAliases;

        /// <summary>
        /// Policy for reading values - controls alias fallback behavior.
        /// </summary>
        public enum ReadPolicy
        {
            /// <summary>Primary parameter only</summary>
            PrimaryOnly,
            /// <summary>Primary first, then aliases in order</summary>
            PrimaryThenAliases,
            /// <summary>First non-empty value from any source</summary>
            FirstAvailable
        }

        /// <summary>
        /// Policy for writing values - controls which parameters get written.
        /// </summary>
        public enum WritePolicy
        {
            /// <summary>Write to primary COBie parameter only</summary>
            PrimaryOnly,
            /// <summary>Write to primary and all aliases</summary>
            PrimaryAndAliases,
            /// <summary>Write to aliases only (preserve COBie params)</summary>
            AliasesOnly
        }

        public ReadPolicy CurrentReadPolicy { get; set; } = ReadPolicy.PrimaryThenAliases;
        public WritePolicy CurrentWritePolicy { get; set; } = WritePolicy.PrimaryAndAliases;

        public CobieMappingService(PresetService presetService)
        {
            _presetService = presetService;
            _writeAliases = presetService.CurrentPreset?.WriteAliases ?? true;

            // Set default write policy based on preset
            CurrentWritePolicy = _writeAliases ? WritePolicy.PrimaryAndAliases : WritePolicy.PrimaryOnly;
        }

        #region Read Operations

        /// <summary>
        /// Resolves a COBie field value from an element using the current preset configuration.
        /// </summary>
        public (bool success, string? value, string? source) ResolveFieldValue(
            Element element,
            Element? typeElement,
            CobieFieldSpec field,
            Document doc)
        {
            if (field == null)
                return (false, null, null);

            var targetElement = field.Scope == "type" ? typeElement : element;
            if (targetElement == null && field.Scope == "type")
                return (false, null, null);

            // 1. Check for computed value first
            if (field.Computed != null && !string.IsNullOrEmpty(field.Computed.Source))
            {
                var computed = ResolveComputedValue(element, field.Computed.Source, doc);
                if (computed.success)
                    return (true, computed.value, "computed:" + field.Computed.Source);
            }

            // 2. Try primary COBie parameter
            if (!string.IsNullOrEmpty(field.RevitBuiltIn))
            {
                var result = TryGetBuiltinParam(targetElement!, field.RevitBuiltIn);
                if (result.ok && !string.IsNullOrWhiteSpace(result.value))
                    return (true, result.value, "builtin:" + field.RevitBuiltIn);
            }

            if (!string.IsNullOrEmpty(field.RevitParam))
            {
                var result = TryGetNamedParam(targetElement!, field.RevitParam);
                if (result.ok && !string.IsNullOrWhiteSpace(result.value))
                    return (true, result.value, "param:" + field.RevitParam);
            }

            // 3. If read policy allows, try aliases
            if (CurrentReadPolicy != ReadPolicy.PrimaryOnly && field.AliasParams != null)
            {
                foreach (var alias in field.AliasParams)
                {
                    if (string.IsNullOrEmpty(alias)) continue;

                    var result = TryGetNamedParam(targetElement!, alias);
                    if (result.ok && !string.IsNullOrWhiteSpace(result.value))
                        return (true, result.value, "alias:" + alias);
                }
            }

            // 4. Return default value if defined
            if (!string.IsNullOrEmpty(field.DefaultValue))
                return (true, field.DefaultValue, "default");

            return (false, null, null);
        }

        /// <summary>
        /// Gets all field values for an element based on the current preset.
        /// </summary>
        public Dictionary<string, FieldValueResult> GetAllFieldValues(
            Element element,
            Element? typeElement,
            Document doc,
            string tableName = "Component")
        {
            var results = new Dictionary<string, FieldValueResult>();
            var fields = _presetService.GetAllFields(tableName);

            foreach (var field in fields)
            {
                if (string.IsNullOrEmpty(field.CobieKey)) continue;

                var (success, value, source) = ResolveFieldValue(element, typeElement, field, doc);
                results[field.CobieKey] = new FieldValueResult
                {
                    CobieKey = field.CobieKey,
                    Label = field.Label ?? field.CobieKey,
                    Value = value,
                    HasValue = success && !string.IsNullOrWhiteSpace(value),
                    Source = source,
                    Required = field.Required,
                    Group = field.Group ?? "Other",
                    DataType = field.DataType ?? "string"
                };
            }

            return results;
        }

        /// <summary>
        /// Resolves computed values from element properties.
        /// </summary>
        public (bool success, string? value) ResolveComputedValue(Element element, string computedId, Document doc)
        {
            if (string.IsNullOrEmpty(computedId))
                return (false, null);

            switch (computedId)
            {
                case "Element.UniqueId":
                    return (true, element.UniqueId);

                case "Element.TypeName":
                    return GetTypeName(element, doc);

                case "Element.LevelName":
                    return GetLevelName(element, doc);

                case "Element.RoomOrSpace":
                    return GetRoomOrSpace(element, doc);

                case "Space.Name":
                    return GetSpaceName(element);

                case "Space.Number":
                    return GetSpaceNumber(element);

                case "Space.LevelName":
                    return GetSpaceLevelName(element, doc);

                case "Level.Name":
                    return GetLevelNameDirect(element);

                case "Level.Elevation":
                    return GetLevelElevation(element);

                default:
                    return (false, null);
            }
        }

        #endregion

        #region Write Operations

        /// <summary>
        /// Sets a COBie field value on an element, respecting write policy for aliases.
        /// </summary>
        public (bool success, int paramsUpdated) SetFieldValue(
            Document doc,
            Element element,
            Element? typeElement,
            CobieFieldSpec field,
            string value)
        {
            if (field == null || string.IsNullOrEmpty(value))
                return (false, 0);

            var targetElement = field.Scope == "type" ? typeElement : element;
            if (targetElement == null)
                return (false, 0);

            int updated = 0;

            // Write to primary parameter
            if (CurrentWritePolicy != WritePolicy.AliasesOnly)
            {
                if (!string.IsNullOrEmpty(field.RevitBuiltIn))
                {
                    if (TrySetBuiltinParam(targetElement, field.RevitBuiltIn, value))
                        updated++;
                }
                else if (!string.IsNullOrEmpty(field.RevitParam))
                {
                    if (TrySetNamedParam(doc, targetElement, field.RevitParam, value))
                        updated++;
                }
            }

            // Write to aliases if policy allows
            if (CurrentWritePolicy != WritePolicy.PrimaryOnly && _writeAliases && field.AliasParams != null)
            {
                foreach (var alias in field.AliasParams)
                {
                    if (string.IsNullOrEmpty(alias)) continue;
                    if (TrySetNamedParam(doc, targetElement, alias, value))
                        updated++;
                }
            }

            return (updated > 0, updated);
        }

        /// <summary>
        /// Simplified overload for setting a field value on a single element.
        /// </summary>
        public bool SetFieldValue(Element element, CobieFieldSpec field, string value, bool writeAliases = true)
        {
            if (element == null || field == null || string.IsNullOrEmpty(value))
                return false;

            var originalPolicy = CurrentWritePolicy;
            if (!writeAliases)
                CurrentWritePolicy = WritePolicy.PrimaryOnly;

            try
            {
                var doc = element.Document;
                Element? typeElement = null;
                
                if (field.Scope == "type" && element is ElementType)
                {
                    // Element is already a type
                    var (success, _) = SetFieldValue(doc, element, element, field, value);
                    return success;
                }
                else if (field.Scope == "type")
                {
                    var typeId = element.GetTypeId();
                    if (typeId != ElementId.InvalidElementId)
                        typeElement = doc.GetElement(typeId);
                }

                var (result, _) = SetFieldValue(doc, element, typeElement, field, value);
                return result;
            }
            finally
            {
                CurrentWritePolicy = originalPolicy;
            }
        }

        /// <summary>
        /// Sets multiple field values on an element.
        /// </summary>
        public (int fieldsUpdated, int paramsUpdated) SetFieldValues(
            Document doc,
            Element element,
            Element? typeElement,
            Dictionary<string, string> values,
            string tableName = "Component")
        {
            var fields = _presetService.GetAllFields(tableName);
            int fieldsUpdated = 0;
            int totalParamsUpdated = 0;

            foreach (var field in fields)
            {
                if (string.IsNullOrEmpty(field.CobieKey)) continue;
                if (!values.TryGetValue(field.CobieKey, out var value)) continue;
                if (string.IsNullOrEmpty(value)) continue;

                var (success, paramsUpdated) = SetFieldValue(doc, element, typeElement, field, value);
                if (success)
                {
                    fieldsUpdated++;
                    totalParamsUpdated += paramsUpdated;
                }
            }

            return (fieldsUpdated, totalParamsUpdated);
        }

        #endregion

        #region Validation

        /// <summary>
        /// Validates field values against rules defined in the preset.
        /// </summary>
        public List<ValidationError> ValidateFieldValues(
            Dictionary<string, FieldValueResult> fieldValues,
            string tableName = "Component")
        {
            var errors = new List<ValidationError>();
            var fields = _presetService.GetAllFields(tableName);

            foreach (var field in fields)
            {
                if (string.IsNullOrEmpty(field.CobieKey)) continue;

                if (!fieldValues.TryGetValue(field.CobieKey, out var result))
                    continue;

                // Required field check
                if (field.Required && !result.HasValue)
                {
                    errors.Add(new ValidationError
                    {
                        CobieKey = field.CobieKey,
                        Label = field.Label ?? field.CobieKey,
                        Rule = "required",
                        Message = $"Required field '{field.Label}' is missing"
                    });
                }

                // Date format check
                if (field.Rules?.Contains("date") == true && result.HasValue)
                {
                    if (!IsValidDateFormat(result.Value))
                    {
                        errors.Add(new ValidationError
                        {
                            CobieKey = field.CobieKey,
                            Label = field.Label ?? field.CobieKey,
                            Rule = "date",
                            Message = $"Field '{field.Label}' has invalid date format (expected YYYY-MM-DD)"
                        });
                    }
                }
            }

            return errors;
        }

        /// <summary>
        /// Checks for uniqueness violations across a collection of elements.
        /// </summary>
        public Dictionary<string, List<int>> CheckUniqueness(
            IEnumerable<(int elementId, Dictionary<string, FieldValueResult> values)> elementValues,
            string tableName = "Component")
        {
            var violations = new Dictionary<string, List<int>>();
            var fields = _presetService.GetAllFields(tableName);
            var uniqueFields = fields.Where(f => f.Rules?.Contains("unique") == true).ToList();

            foreach (var field in uniqueFields)
            {
                if (string.IsNullOrEmpty(field.CobieKey)) continue;

                var valueToElements = new Dictionary<string, List<int>>();

                foreach (var (elementId, values) in elementValues)
                {
                    if (!values.TryGetValue(field.CobieKey, out var result)) continue;
                    if (!result.HasValue || string.IsNullOrWhiteSpace(result.Value)) continue;

                    if (!valueToElements.ContainsKey(result.Value))
                        valueToElements[result.Value] = new List<int>();

                    valueToElements[result.Value].Add(elementId);
                }

                // Find duplicates
                foreach (var kvp in valueToElements)
                {
                    if (kvp.Value.Count > 1)
                    {
                        if (!violations.ContainsKey(field.CobieKey))
                            violations[field.CobieKey] = new List<int>();

                        violations[field.CobieKey].AddRange(kvp.Value);
                    }
                }
            }

            return violations;
        }

        private bool IsValidDateFormat(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return true;

            // Accept various date formats
            var formats = new[]
            {
                "yyyy-MM-dd",
                "yyyy/MM/dd",
                "MM/dd/yyyy",
                "dd/MM/yyyy",
                "yyyy-MM-ddTHH:mm:ss"
            };

            return DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _);
        }

        #endregion

        #region COBie Readiness Score

        /// <summary>
        /// Calculates COBie readiness score for an element.
        /// </summary>
        public CobieReadinessScore CalculateReadinessScore(
            Dictionary<string, FieldValueResult> fieldValues,
            List<ValidationError> validationErrors,
            string tableName = "Component")
        {
            var fields = _presetService.GetAllFields(tableName);
            var score = new CobieReadinessScore();

            foreach (var field in fields)
            {
                if (string.IsNullOrEmpty(field.CobieKey)) continue;

                if (field.Required)
                    score.TotalRequiredFields++;
                else
                    score.TotalOptionalFields++;

                if (fieldValues.TryGetValue(field.CobieKey, out var result) && result.HasValue)
                {
                    if (field.Required)
                        score.PopulatedRequiredFields++;
                    else
                        score.PopulatedOptionalFields++;
                }
            }

            score.ValidationErrors = validationErrors.Count;
            score.CalculateScores();

            return score;
        }

        #endregion

        #region Private Parameter Helpers

        private (bool ok, string? value) TryGetBuiltinParam(Element element, string? bipName)
        {
            if (element == null || string.IsNullOrEmpty(bipName))
                return (false, null);

            if (!Enum.TryParse<BuiltInParameter>(bipName, out var bip))
                return (false, null);

            var param = element.get_Parameter(bip);
            return ExtractParamValue(param);
        }

        private (bool ok, string? value) TryGetNamedParam(Element element, string? paramName)
        {
            if (element == null || string.IsNullOrEmpty(paramName))
                return (false, null);

            var param = element.LookupParameter(paramName);
            return ExtractParamValue(param);
        }

        private (bool ok, string? value) ExtractParamValue(Parameter? param)
        {
            if (param == null || !param.HasValue)
                return (false, null);

            string? value = null;
            switch (param.StorageType)
            {
                case StorageType.String:
                    value = param.AsString();
                    break;
                case StorageType.Integer:
                    value = param.AsInteger().ToString();
                    break;
                case StorageType.Double:
                    value = param.AsValueString() ?? param.AsDouble().ToString("F2");
                    break;
                case StorageType.ElementId:
                    var id = param.AsElementId();
                    value = id != null ? GetElementIdValue(id).ToString() : null;
                    break;
            }

            if (string.IsNullOrWhiteSpace(value))
                return (false, null);

            return (true, value);
        }

        private bool TrySetBuiltinParam(Element element, string bipName, string value)
        {
            if (!Enum.TryParse<BuiltInParameter>(bipName, out var bip))
                return false;

            var param = element.get_Parameter(bip);
            if (param == null || param.IsReadOnly)
                return false;

            return TrySetParamValue(param, value);
        }

        private bool TrySetNamedParam(Document doc, Element element, string paramName, string value)
        {
            var param = element.LookupParameter(paramName);
            if (param == null)
            {
                // Try to create the parameter if it doesn't exist
                // This would require shared parameter setup - skip for now
                return false;
            }

            if (param.IsReadOnly)
                return false;

            return TrySetParamValue(param, value);
        }

        private bool TrySetParamValue(Parameter param, string? value)
        {
            if (param == null || param.IsReadOnly || value == null)
                return false;

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        return param.Set(value);
                    case StorageType.Integer:
                        if (int.TryParse(value, out var intVal))
                            return param.Set(intVal);
                        break;
                    case StorageType.Double:
                        if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out var dblVal))
                            return param.Set(dblVal);
                        break;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        #endregion

        #region Computed Value Helpers

        private (bool ok, string? value) GetTypeName(Element element, Document doc)
        {
            var typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
                return (false, null);

            var elementType = doc.GetElement(typeId) as ElementType;
            if (elementType == null)
                return (false, null);

            return (true, elementType.Name);
        }

        private (bool ok, string? value) GetLevelName(Element element, Document doc)
        {
            var levelId = element.LevelId;
            if (levelId == null || levelId == ElementId.InvalidElementId)
            {
                var levelParam = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
                               ?? element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
                               ?? element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);

                if (levelParam != null && levelParam.HasValue)
                    levelId = levelParam.AsElementId();
            }

            if (levelId == null || levelId == ElementId.InvalidElementId)
                return (false, null);

            var level = doc.GetElement(levelId) as Level;
            if (level == null)
                return (false, null);

            return (true, level.Name);
        }

        private (bool ok, string? value) GetRoomOrSpace(Element element, Document doc)
        {
            var point = GetElementPoint(element);
            if (point == null)
                return (false, null);

            var room = doc.GetRoomAtPoint(point);
            if (room != null && !string.IsNullOrWhiteSpace(room.Name))
            {
                var roomNumber = room.Number;
                var roomName = room.Name;
                var display = !string.IsNullOrEmpty(roomNumber) ? roomNumber + " - " + roomName : roomName;
                return (true, display);
            }

            var space = doc.GetSpaceAtPoint(point);
            if (space != null && !string.IsNullOrWhiteSpace(space.Name))
            {
                var spaceNumber = space.Number;
                var spaceName = space.Name;
                var display = !string.IsNullOrEmpty(spaceNumber)
                    ? spaceNumber + " - " + spaceName
                    : spaceName;
                return (true, display);
            }

            return (false, null);
        }

        private (bool ok, string? value) GetSpaceName(Element element)
        {
            if (element is Room room)
                return (true, room.Name);
            if (element is Space space)
                return (true, space.Name);
            return (false, null);
        }

        private (bool ok, string? value) GetSpaceNumber(Element element)
        {
            if (element is Room room)
                return (true, room.Number);
            if (element is Space space)
                return (true, space.Number);
            return (false, null);
        }

        private (bool ok, string? value) GetSpaceLevelName(Element element, Document doc)
        {
            Level? level = null;

            if (element is Room room)
                level = room.Level;
            else if (element is Space space)
                level = space.Level;

            if (level == null)
                return (false, null);

            return (true, level.Name);
        }

        private (bool ok, string? value) GetLevelNameDirect(Element element)
        {
            if (element is Level level)
                return (true, level.Name);
            return (false, null);
        }

        private (bool ok, string? value) GetLevelElevation(Element element)
        {
            if (element is Level level)
                return (true, level.Elevation.ToString("F2"));
            return (false, null);
        }

        private XYZ? GetElementPoint(Element element)
        {
            var locationPoint = element.Location as LocationPoint;
            if (locationPoint != null)
                return locationPoint.Point;

            var bb = element.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) * 0.5;

            return null;
        }

        private static int GetElementIdValue(ElementId id)
        {
#if REVIT2026_OR_GREATER
            return unchecked((int)id.Value);
#else
            return id.IntegerValue;
#endif
        }

        #endregion
    }

    #region Result Classes

    public class FieldValueResult
    {
        public string CobieKey { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
        public string? Value { get; set; }
        public bool HasValue { get; set; }
        public string? Source { get; set; }
        public bool Required { get; set; }
        public string Group { get; set; } = "Other";
        public string DataType { get; set; } = "string";
    }

    public class ValidationError
    {
        public string CobieKey { get; set; } = string.Empty;
        public string Label { get; set; } = string.Empty;
        public string Rule { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
    }

    public class CobieReadinessScore
    {
        public int TotalRequiredFields { get; set; }
        public int PopulatedRequiredFields { get; set; }
        public int TotalOptionalFields { get; set; }
        public int PopulatedOptionalFields { get; set; }
        public int ValidationErrors { get; set; }

        public double RequiredFieldsScore { get; private set; }
        public double OptionalFieldsScore { get; private set; }
        public double OverallScore { get; private set; }
        public bool IsCobieReady { get; private set; }

        public void CalculateScores()
        {
            RequiredFieldsScore = TotalRequiredFields > 0
                ? (double)PopulatedRequiredFields / TotalRequiredFields * 100
                : 100;

            OptionalFieldsScore = TotalOptionalFields > 0
                ? (double)PopulatedOptionalFields / TotalOptionalFields * 100
                : 100;

            // Overall score: 70% required fields, 30% optional
            OverallScore = (RequiredFieldsScore * 0.7) + (OptionalFieldsScore * 0.3);

            // COBie ready = all required fields populated and no validation errors
            IsCobieReady = PopulatedRequiredFields == TotalRequiredFields && ValidationErrors == 0;
        }
    }

    #endregion
}

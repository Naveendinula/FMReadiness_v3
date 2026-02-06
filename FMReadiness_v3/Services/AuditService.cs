using System;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;

namespace FMReadiness_v3.Services
{
    public class AuditService
    {
        public class AuditReport
        {
            public int ElementsWithMissingData { get; set; }
            public Dictionary<string, int> MissingParamCounts { get; set; } = new Dictionary<string, int>();

            public int ElementsWithMissingTypeData { get; set; }
            public Dictionary<string, int> MissingTypeParamCounts { get; set; } = new Dictionary<string, int>();

            public int TotalAuditedAssets { get; set; }
            public int FullyReadyAssets { get; set; }
            public double AverageReadinessScore { get; set; }
            public string AuditProfileName { get; set; } = string.Empty;
            public AuditScoreMode ScoreMode { get; set; } = AuditScoreMode.AllEditable;

            public List<ElementAuditResult> ElementResults { get; set; } = new List<ElementAuditResult>();

            public Dictionary<string, double> AverageGroupScores { get; set; } = new Dictionary<string, double>();

            public Dictionary<string, List<int>> UniquenessViolations { get; set; } = new Dictionary<string, List<int>>();
        }

        public AuditReport RunFullAudit(
            Document doc,
            IEnumerable<Element> elements,
            Dictionary<string, CategoryConfig> rules,
            AuditScoreMode scoreMode = AuditScoreMode.AllEditable)
        {
            var report = new AuditReport();
            if (doc == null || elements == null || rules == null || rules.Count == 0)
                return report;
            var elementList = elements.ToList();
            report.ScoreMode = scoreMode;

            // Phase 1: collect all field values for uniqueness checks.
            var fieldValuesMap = new Dictionary<int, Dictionary<string, string>>();
            var uniqueFieldValues = new Dictionary<string, Dictionary<string, List<int>>>();

            foreach (var element in elementList)
            {
                var categoryKey = GetCategoryKey(element);
                if (string.IsNullOrEmpty(categoryKey) || !rules.ContainsKey(categoryKey))
                    continue;

                var config = rules[categoryKey];
                if (config.Groups == null || config.Groups.Count == 0)
                    continue;
                var elementId = GetElementIdValue(element.Id);
                var fieldValues = new Dictionary<string, string>();

                Element typeElement = null;
                var typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElement = doc.GetElement(typeId);

                foreach (var groupEntry in config.Groups)
                {
                    var groupConfig = groupEntry.Value;
                    if (groupConfig?.Fields == null || groupConfig.Fields.Count == 0)
                        continue;

                    foreach (var field in groupConfig.Fields)
                    {
                        if (field == null) continue;
                        var result = TryGetFieldValue(element, typeElement, field, doc);
                        fieldValues[field.Key] = result.ok ? result.value : null;

                        if (HasRule(field, "unique") && result.ok && !string.IsNullOrWhiteSpace(result.value))
                        {
                            if (!uniqueFieldValues.ContainsKey(field.Key))
                                uniqueFieldValues[field.Key] = new Dictionary<string, List<int>>();

                            if (!uniqueFieldValues[field.Key].ContainsKey(result.value))
                                uniqueFieldValues[field.Key][result.value] = new List<int>();

                            uniqueFieldValues[field.Key][result.value].Add(elementId);
                        }
                    }
                }

                fieldValuesMap[elementId] = fieldValues;
            }

            // Phase 2: identify uniqueness violations.
            var duplicateElementFields = new Dictionary<int, HashSet<string>>();
            foreach (var fieldEntry in uniqueFieldValues)
            {
                var fieldKey = fieldEntry.Key;
                var valueMap = fieldEntry.Value;

                foreach (var valueEntry in valueMap)
                {
                    var elementIds = valueEntry.Value;
                    if (elementIds.Count <= 1) continue;

                    if (!report.UniquenessViolations.ContainsKey(fieldKey))
                        report.UniquenessViolations[fieldKey] = new List<int>();

                    report.UniquenessViolations[fieldKey].AddRange(elementIds);

                    foreach (var eid in elementIds)
                    {
                        if (!duplicateElementFields.ContainsKey(eid))
                            duplicateElementFields[eid] = new HashSet<string>();
                        duplicateElementFields[eid].Add(fieldKey);
                    }
                }
            }

            // Phase 3: score each element.
            var scoreSum = 0.0;
            var scoredAssets = 0;
            var groupScoreSums = new Dictionary<string, double>();
            var groupScoreCounts = new Dictionary<string, int>();

            foreach (var element in elementList)
            {
                var categoryKey = GetCategoryKey(element);
                if (string.IsNullOrEmpty(categoryKey) || !rules.ContainsKey(categoryKey))
                    continue;

                var config = rules[categoryKey];
                if (config.Groups == null || config.Groups.Count == 0)
                    continue;
                var elementId = GetElementIdValue(element.Id);

                if (!fieldValuesMap.ContainsKey(elementId))
                    continue;

                var fieldValues = fieldValuesMap[elementId];
                HashSet<string> duplicateFields;
                if (!duplicateElementFields.TryGetValue(elementId, out duplicateFields))
                    duplicateFields = new HashSet<string>();

                Element typeElement = null;
                var typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                    typeElement = doc.GetElement(typeId);

                var missingFields = new List<MissingFieldInfo>();
                var groupScores = new Dictionary<string, double>();
                var totalScored = 0;
                var totalMissing = 0;
                var scoreAllFields = scoreMode == AuditScoreMode.AllEditable;

                foreach (var groupEntry in config.Groups)
                {
                    var groupName = groupEntry.Key;
                    var groupConfig = groupEntry.Value;
                    if (groupConfig?.Fields == null || groupConfig.Fields.Count == 0)
                        continue;

                    var fields = groupConfig.Fields.Where(f => f != null).ToList();
                    if (fields.Count == 0)
                        continue;

                    var groupTotal = 0;
                    var groupMissing = 0;

                    foreach (var field in fields)
                    {
                        var isRequired = IsFieldRequired(field);
                        var hasUniqueRule = HasRule(field, "unique");
                        var isScoredField = scoreAllFields || isRequired || hasUniqueRule;
                        if (!isScoredField)
                            continue;

                        string value;
                        fieldValues.TryGetValue(field.Key, out value);
                        var isMissing = string.IsNullOrWhiteSpace(value);
                        var isDuplicate = duplicateFields.Contains(field.Key);
                        var countsMissing = scoreAllFields || isRequired;
                        var isFailed = isDuplicate || (countsMissing && isMissing);

                        groupTotal++;

                        if (isFailed)
                        {
                            groupMissing++;
                            missingFields.Add(new MissingFieldInfo
                            {
                                FieldKey = field.Key,
                                Group = groupName,
                                Scope = field.Scope ?? string.Empty,
                                Required = isRequired,
                                FieldLabel = field.Label,
                                Reason = isDuplicate ? "duplicate" : null
                            });

                            var paramKey = $"[{groupName}] {field.Label}";
                            if (field.Scope == "type")
                            {
                                int cnt;
                                report.MissingTypeParamCounts.TryGetValue(paramKey, out cnt);
                                report.MissingTypeParamCounts[paramKey] = cnt + 1;
                            }
                            else
                            {
                                int cnt;
                                report.MissingParamCounts.TryGetValue(paramKey, out cnt);
                                report.MissingParamCounts[paramKey] = cnt + 1;
                            }
                        }
                    }

                    var groupScore = groupTotal > 0 ? 1.0 - ((double)groupMissing / groupTotal) : 1.0;
                    groupScore = Clamp01(groupScore);
                    groupScores[groupName] = groupScore;

                    if (!groupScoreSums.ContainsKey(groupName))
                    {
                        groupScoreSums[groupName] = 0;
                        groupScoreCounts[groupName] = 0;
                    }

                    groupScoreSums[groupName] += groupScore;
                    groupScoreCounts[groupName]++;

                    totalScored += groupTotal;
                    totalMissing += groupMissing;
                }

                report.TotalAuditedAssets++;
                if (totalMissing > 0)
                {
                    report.ElementsWithMissingData++;
                    if (missingFields.Any(f => string.Equals(f.Scope, "type", StringComparison.OrdinalIgnoreCase)))
                    {
                        report.ElementsWithMissingTypeData++;
                    }
                }

                var overallScore = totalScored > 0 ? 1.0 - ((double)totalMissing / totalScored) : 1.0;
                overallScore = Clamp01(overallScore);

                scoreSum += overallScore;
                scoredAssets++;

                if (totalMissing == 0)
                    report.FullyReadyAssets++;

                var familyName = string.Empty;
                var typeName = string.Empty;
                var elementType = typeElement as ElementType;
                if (elementType != null)
                {
                    familyName = elementType.FamilyName ?? string.Empty;
                    typeName = elementType.Name ?? string.Empty;
                }

                var missingParamsStr = string.Join(", ", missingFields.Select(f =>
                    f.Reason == "duplicate" ? $"[{f.Group}] {f.FieldLabel} (dup)" : $"[{f.Group}] {f.FieldLabel}"));

                report.ElementResults.Add(new ElementAuditResult
                {
                    ElementId = elementId,
                    Category = element.Category != null ? element.Category.Name : categoryKey,
                    Family = familyName,
                    Type = typeName,
                    MissingCount = totalMissing,
                    ReadinessScore = overallScore,
                    MissingParams = missingParamsStr,
                    GroupScores = groupScores,
                    MissingFields = missingFields
                });
            }

            report.AverageReadinessScore = scoredAssets == 0 ? 1.0 : (scoreSum / scoredAssets);

            foreach (var groupEntry in groupScoreSums)
            {
                var groupName = groupEntry.Key;
                var sum = groupEntry.Value;
                var count = groupScoreCounts[groupName];
                report.AverageGroupScores[groupName] = count > 0 ? sum / count : 1.0;
            }

            return report;
        }

        private static bool IsFieldRequired(FieldSpec field)
        {
            if (field == null)
                return false;

            if (field.Required.HasValue)
                return field.Required.Value;

            if (HasRule(field, "optional"))
                return false;

            if (HasRule(field, "required"))
                return true;

            return true;
        }

        private static bool HasRule(FieldSpec field, string rule)
        {
            if (field?.Rules == null || string.IsNullOrWhiteSpace(rule))
                return false;

            return field.Rules.Any(r => string.Equals(r, rule, StringComparison.OrdinalIgnoreCase));
        }

        private static string GetCategoryKey(Element element)
        {
            var category = element.Category;
            if (category == null) return string.Empty;

            try
            {
                var bic = (BuiltInCategory)GetElementIdValue(category.Id);
                return bic.ToString();
            }
            catch
            {
                return category.Name ?? string.Empty;
            }
        }

        private static int GetElementIdValue(ElementId id)
        {
#if REVIT2026_OR_GREATER
            return unchecked((int)id.Value);
#else
            return id.IntegerValue;
#endif
        }

        private static double Clamp01(double value)
        {
            if (value < 0) return 0;
            if (value > 1) return 1;
            return value;
        }

        private (bool ok, string value) TryGetFieldValue(Element element, Element typeElement, FieldSpec field, Document doc)
        {
            if (field?.Source == null || string.IsNullOrWhiteSpace(field.Source.Type))
                return (false, null);

            var targetElement = field.Scope == "type" ? typeElement : element;
            var sourceType = field.Source.Type;

            switch (sourceType.ToLowerInvariant())
            {
                case "builtin":
                    return TryGetBuiltinParam(targetElement, field.Source.Id);
                case "name":
                    return TryGetNamedParam(targetElement, field.Source.Value);
                case "sharedguid":
                    return TryGetSharedGuidParam(targetElement, field.Source.Id);
                case "computed":
                    return TryGetComputedField(element, field.Source.Id, doc);
                default:
                    return (false, null);
            }
        }

        private (bool ok, string value) TryGetBuiltinParam(Element element, string bipName)
        {
            if (element == null || string.IsNullOrEmpty(bipName))
                return (false, null);

            if (!TryParseEnumValue(bipName, out BuiltInParameter bip))
                return (false, null);

            var param = element.get_Parameter(bip);
            return ExtractParamValue(param);
        }

        private (bool ok, string value) TryGetNamedParam(Element element, string paramName)
        {
            if (element == null || string.IsNullOrEmpty(paramName))
                return (false, null);

            var param = element.LookupParameter(paramName);
            return ExtractParamValue(param);
        }

        private (bool ok, string value) TryGetSharedGuidParam(Element element, string guidStr)
        {
            if (element == null || string.IsNullOrEmpty(guidStr))
                return (false, null);

            Guid guid;
            if (!Guid.TryParse(guidStr, out guid))
                return (false, null);

            var param = element.get_Parameter(guid);
            return ExtractParamValue(param);
        }

        private (bool ok, string value) ExtractParamValue(Parameter param)
        {
            if (param == null || !param.HasValue)
                return (false, null);

            string value = null;
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

        private static bool TryParseEnumValue<TEnum>(string? value, out TEnum result)
            where TEnum : struct, Enum
        {
            result = default;
            if (string.IsNullOrWhiteSpace(value)) return false;

            try
            {
                if (Enum.TryParse(value, out result)) return true;
            }
            catch
            {
                // Fall through to numeric parsing.
            }

            if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numeric))
                return false;

            var underlying = Enum.GetUnderlyingType(typeof(TEnum));
            try
            {
                var converted = Convert.ChangeType(numeric, underlying, CultureInfo.InvariantCulture);
                result = (TEnum)Enum.ToObject(typeof(TEnum), converted);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private (bool ok, string value) TryGetComputedField(Element element, string computedId, Document doc)
        {
            if (string.IsNullOrEmpty(computedId))
                return (false, null);

            if (computedId == "Element.UniqueId")
                return (true, element.UniqueId);

            if (computedId == "Element.LevelName")
                return GetLevelName(element, doc);

            if (computedId == "Element.RoomOrSpace")
                return GetRoomOrSpace(element, doc);

            return (false, null);
        }

        private (bool ok, string value) GetLevelName(Element element, Document doc)
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

        private (bool ok, string value) GetRoomOrSpace(Element element, Document doc)
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
                    ? "Space: " + spaceNumber + " - " + spaceName
                    : "Space: " + spaceName;
                return (true, display);
            }

            return (false, null);
        }

        private XYZ GetElementPoint(Element element)
        {
            var locationPoint = element.Location as LocationPoint;
            if (locationPoint != null)
                return locationPoint.Point;

            var bb = element.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) * 0.5;

            return null;
        }
    }
}

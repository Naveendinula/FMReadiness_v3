using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.UI;
using FMReadiness_v3.Services;
using FMReadiness_v3.UI.Panes;

namespace FMReadiness_v3.UI.ExternalEvents
{
    /// <summary>
    /// Handles parameter editing operations from the WebView2 UI.
    /// </summary>
    public class ParameterEditorExternalEventHandler : IExternalEventHandler
    {
        public enum OperationType
        {
            GetSelectedElements,
            GetCategoryStats,
            SetInstanceParams,
            SetCategoryParams,
            SetTypeParams,
            CopyComputedToParam,
            RefreshAudit  // New operation to refresh audit after changes
        }

        public OperationType CurrentOperation { get; set; }
        public string? JsonPayload { get; set; }

        // Callback to send results back to WebView
        public Action<string>? PostResult { get; set; }

        public void Execute(UIApplication app)
        {
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;
            var doc = uidoc.Document;

            try
            {
                switch (CurrentOperation)
                {
                    case OperationType.GetSelectedElements:
                        HandleGetSelectedElements(uidoc, doc);
                        break;
                    case OperationType.GetCategoryStats:
                        HandleGetCategoryStats(doc);
                        break;
                    case OperationType.SetInstanceParams:
                        HandleSetInstanceParams(doc);
                        break;
                    case OperationType.SetCategoryParams:
                        HandleSetCategoryParams(doc);
                        break;
                    case OperationType.SetTypeParams:
                        HandleSetTypeParams(doc);
                        break;
                    case OperationType.CopyComputedToParam:
                        HandleCopyComputedToParam(doc);
                        break;
                    case OperationType.RefreshAudit:
                        HandleRefreshAudit(doc);
                        break;
                }
            }
            catch (Exception ex)
            {
                SendOperationResult(false, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Re-runs the audit and pushes updated results to the UI.
        /// This is called after parameter changes to keep the audit results in sync.
        /// </summary>
        private void HandleRefreshAudit(Document doc)
        {
            try
            {
                // Load checklist config
                var checklist = new ChecklistService();
                if (!checklist.LoadConfig())
                {
                    SendOperationResult(false, "Could not load checklist configuration");
                    return;
                }

                // Collect elements
                var collector = new CollectorService(doc);
                var elements = collector.GetAllFmElements();

                // Run audit
                var auditService = new AuditService();
                var report = auditService.RunFullAudit(doc, elements, checklist.Rules);

                // Push results to WebView - this will update the audit results tab
                WebViewPaneController.UpdateFromReport(report);

                SendOperationResult(true, "Audit refreshed", false);
            }
            catch (Exception ex)
            {
                SendOperationResult(false, $"Failed to refresh audit: {ex.Message}");
            }
        }

        private void HandleGetSelectedElements(UIDocument uidoc, Document doc)
        {
            var selectedIds = uidoc.Selection.GetElementIds();
            var elements = new List<object>();

            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element == null) continue;

                // Skip non-MEP elements
                var cat = GetBuiltInCategory(element);
                if (cat != BuiltInCategory.OST_MechanicalEquipment &&
                    cat != BuiltInCategory.OST_DuctTerminal &&
                    cat != BuiltInCategory.OST_DuctAccessory &&
                    cat != BuiltInCategory.OST_PipeAccessory)
                    continue;

                var typeId = element.GetTypeId();
                ElementType? elementType = null;
                if (typeId != ElementId.InvalidElementId)
                    elementType = doc.GetElement(typeId) as ElementType;

                // Count instances of this type
                int typeInstanceCount = 0;
                if (typeId != ElementId.InvalidElementId)
                {
                    typeInstanceCount = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .Where(e => e.GetTypeId() == typeId)
                        .Count();
                }

                // Get instance parameters
                var instanceParams = new Dictionary<string, string?>
                {
                    ["FM_Barcode"] = GetParamValue(element, "FM_Barcode"),
                    ["FM_UniqueAssetId"] = GetParamValue(element, "FM_UniqueAssetId"),
                    ["FM_InstallationDate"] = GetParamValue(element, "FM_InstallationDate"),
                    ["FM_WarrantyStart"] = GetParamValue(element, "FM_WarrantyStart"),
                    ["FM_WarrantyEnd"] = GetParamValue(element, "FM_WarrantyEnd"),
                    ["FM_Criticality"] = GetParamValue(element, "FM_Criticality"),
                    ["FM_Trade"] = GetParamValue(element, "FM_Trade"),
                    ["FM_PMTemplateId"] = GetParamValue(element, "FM_PMTemplateId"),
                    ["FM_PMFrequencyDays"] = GetParamValue(element, "FM_PMFrequencyDays"),
                    ["FM_Building"] = GetParamValue(element, "FM_Building"),
                    ["FM_LocationSpace"] = GetParamValue(element, "FM_LocationSpace")
                };

                // Get type parameters
                var typeParams = new Dictionary<string, string?>();
                if (elementType != null)
                {
                    typeParams["Manufacturer"] = GetParamValue(elementType, "Manufacturer");
                    typeParams["Model"] = GetParamValue(elementType, "Model");
                    typeParams["TypeMark"] = GetParamValue(elementType, BuiltInParameter.ALL_MODEL_TYPE_MARK);
                }

                // Get computed values
                var computed = new Dictionary<string, string?>
                {
                    ["UniqueId"] = element.UniqueId,
                    ["Level"] = GetElementLevelName(element, doc),
                    ["RoomSpace"] = GetElementRoomSpace(element, doc)
                };

                elements.Add(new
                {
                    elementId = GetElementIdValue(element.Id),
                    category = element.Category?.Name ?? "Unknown",
                    family = elementType?.FamilyName ?? "Unknown",
                    typeName = elementType?.Name ?? "Unknown",
                    typeId = typeId != ElementId.InvalidElementId ? (int?)GetElementIdValue(typeId) : null,
                    typeInstanceCount = typeInstanceCount,
                    instanceParams = instanceParams,
                    typeParams = typeParams,
                    computed = computed
                });
            }

            var result = new { type = "selectedElementsData", elements = elements };
            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            PostResult?.Invoke(json);
        }

        private void HandleGetCategoryStats(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload)) return;

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var categoryStr = jsonDoc.RootElement.GetProperty("category").GetString();

            if (!Enum.TryParse<BuiltInCategory>(categoryStr, out var category))
            {
                SendCategoryStats(categoryStr ?? "", 0);
                return;
            }

            var count = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .GetElementCount();

            SendCategoryStats(categoryStr ?? "", count);
        }

        private void HandleSetInstanceParams(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload)) return;

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var elementIds = jsonDoc.RootElement.GetProperty("elementIds")
                .EnumerateArray()
                .Select(e => new ElementId(e.GetInt32()))
                .ToList();

            var paramsToSet = jsonDoc.RootElement.GetProperty("params")
                .EnumerateObject()
                .ToDictionary(p => p.Name, p => p.Value.GetString());

            int updated = 0;
            using (var tx = new Transaction(doc, "Set FM Instance Parameters"))
            {
                tx.Start();
                foreach (var id in elementIds)
                {
                    var element = doc.GetElement(id);
                    if (element == null) continue;

                    foreach (var kvp in paramsToSet)
                    {
                        if (SetParamValue(element, kvp.Key, kvp.Value))
                            updated++;
                    }
                }
                tx.Commit();
            }

            SendOperationResult(true, $"Updated {updated} parameter values on {elementIds.Count} elements", true);
        }

        private void HandleSetCategoryParams(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload)) return;

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var categoryStr = jsonDoc.RootElement.GetProperty("category").GetString();
            var onlyBlanks = jsonDoc.RootElement.GetProperty("onlyBlanks").GetBoolean();

            var paramsToSet = jsonDoc.RootElement.GetProperty("params")
                .EnumerateObject()
                .ToDictionary(p => p.Name, p => p.Value.GetString());

            if (!Enum.TryParse<BuiltInCategory>(categoryStr, out var category))
            {
                SendOperationResult(false, "Invalid category");
                return;
            }

            var elements = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToElements();

            int updated = 0;
            int skipped = 0;

            using (var tx = new Transaction(doc, "Bulk Set FM Parameters"))
            {
                tx.Start();
                foreach (var element in elements)
                {
                    foreach (var kvp in paramsToSet)
                    {
                        if (onlyBlanks)
                        {
                            var existingValue = GetParamValue(element, kvp.Key);
                            if (!string.IsNullOrWhiteSpace(existingValue))
                            {
                                skipped++;
                                continue;
                            }
                        }

                        if (SetParamValue(element, kvp.Key, kvp.Value))
                            updated++;
                    }
                }
                tx.Commit();
            }

            var msg = $"Updated {updated} values on {elements.Count} elements";
            if (skipped > 0) msg += $" ({skipped} skipped - not blank)";
            SendOperationResult(true, msg);
        }

        private void HandleSetTypeParams(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload)) return;

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var typeIdInt = jsonDoc.RootElement.GetProperty("typeId").GetInt32();
            var typeId = new ElementId(typeIdInt);

            var paramsToSet = jsonDoc.RootElement.GetProperty("params")
                .EnumerateObject()
                .ToDictionary(p => p.Name, p => p.Value.GetString());

            var elementType = doc.GetElement(typeId) as ElementType;
            if (elementType == null)
            {
                SendOperationResult(false, "Type not found");
                return;
            }

            int updated = 0;
            using (var tx = new Transaction(doc, "Set FM Type Parameters"))
            {
                tx.Start();
                foreach (var kvp in paramsToSet)
                {
                    if (kvp.Key == "TypeMark")
                    {
                        var param = elementType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK);
                        if (param != null && !param.IsReadOnly)
                        {
                            param.Set(kvp.Value ?? "");
                            updated++;
                        }
                    }
                    else if (SetParamValue(elementType, kvp.Key, kvp.Value))
                    {
                        updated++;
                    }
                }
                tx.Commit();
            }

            // Count affected instances
            var instanceCount = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(e => e.GetTypeId() == typeId)
                .Count();

            SendOperationResult(true, $"Updated {updated} type parameters (affects {instanceCount} instances)", true);
        }

        private void HandleCopyComputedToParam(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload)) return;

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var elementIds = jsonDoc.RootElement.GetProperty("elementIds")
                .EnumerateArray()
                .Select(e => new ElementId(e.GetInt32()))
                .ToList();

            var sourceField = jsonDoc.RootElement.GetProperty("sourceField").GetString();
            var targetParam = jsonDoc.RootElement.GetProperty("targetParam").GetString();

            int updated = 0;
            using (var tx = new Transaction(doc, "Copy Computed to Parameter"))
            {
                tx.Start();
                foreach (var id in elementIds)
                {
                    var element = doc.GetElement(id);
                    if (element == null) continue;

                    string? value = sourceField switch
                    {
                        "UniqueId" => element.UniqueId,
                        "Level" => GetElementLevelName(element, doc),
                        "RoomSpace" => GetElementRoomSpace(element, doc),
                        _ => null
                    };

                    if (!string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(targetParam))
                    {
                        if (SetParamValue(element, targetParam, value))
                            updated++;
                    }
                }
                tx.Commit();
            }

            SendOperationResult(true, $"Copied {sourceField} to {targetParam} on {updated} elements", true);
        }

        // Helper methods
        private string? GetParamValue(Element element, string paramName)
        {
            var param = element.LookupParameter(paramName);
            if (param == null || !param.HasValue) return null;

            return param.StorageType switch
            {
                StorageType.String => param.AsString(),
                StorageType.Integer => param.AsInteger().ToString(),
                StorageType.Double => param.AsValueString() ?? param.AsDouble().ToString("F2"),
                _ => null
            };
        }

        private string? GetParamValue(Element element, BuiltInParameter bip)
        {
            var param = element.get_Parameter(bip);
            if (param == null || !param.HasValue) return null;

            return param.StorageType switch
            {
                StorageType.String => param.AsString(),
                StorageType.Integer => param.AsInteger().ToString(),
                StorageType.Double => param.AsValueString() ?? param.AsDouble().ToString("F2"),
                _ => null
            };
        }

        private bool SetParamValue(Element element, string paramName, string? value)
        {
            var param = element.LookupParameter(paramName);
            if (param == null || param.IsReadOnly) return false;

            try
            {
                switch (param.StorageType)
                {
                    case StorageType.String:
                        param.Set(value ?? "");
                        return true;
                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                        {
                            param.Set(intVal);
                            return true;
                        }
                        break;
                    case StorageType.Double:
                        if (double.TryParse(value, out double dblVal))
                        {
                            param.Set(dblVal);
                            return true;
                        }
                        break;
                }
            }
            catch { }
            return false;
        }

        private string? GetElementLevelName(Element element, Document doc)
        {
            // Try direct LevelId
            var levelId = element.LevelId;
            if (levelId != null && levelId != ElementId.InvalidElementId)
            {
                var level = doc.GetElement(levelId) as Level;
                if (level != null) return level.Name;
            }

            // Try level parameters
            var levelParams = new[] {
                BuiltInParameter.FAMILY_LEVEL_PARAM,
                BuiltInParameter.SCHEDULE_LEVEL_PARAM,
                BuiltInParameter.RBS_START_LEVEL_PARAM,
                BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM
            };

            foreach (var bip in levelParams)
            {
                var param = element.get_Parameter(bip);
                if (param != null && param.HasValue)
                {
                    var paramLevelId = param.AsElementId();
                    if (paramLevelId != null && paramLevelId != ElementId.InvalidElementId)
                    {
                        var level = doc.GetElement(paramLevelId) as Level;
                        if (level != null) return level.Name;
                    }
                }
            }

            return null;
        }

        private string? GetElementRoomSpace(Element element, Document doc)
        {
            XYZ? point = null;

            if (element.Location is LocationPoint lp)
                point = lp.Point;
            else if (element.Location is LocationCurve lc)
                point = lc.Curve.Evaluate(0.5, true);
            else
            {
                var bb = element.get_BoundingBox(null);
                if (bb != null)
                    point = (bb.Min + bb.Max) * 0.5;
            }

            if (point == null) return null;

            // Try Room
            var room = doc.GetRoomAtPoint(point);
            if (room != null && !string.IsNullOrWhiteSpace(room.Name))
            {
                var display = !string.IsNullOrEmpty(room.Number) ? $"{room.Number} - {room.Name}" : room.Name;
                return display;
            }

            // Try Space
            var space = doc.GetSpaceAtPoint(point);
            if (space != null && !string.IsNullOrWhiteSpace(space.Name))
            {
                var display = !string.IsNullOrEmpty(space.Number) ? $"{space.Number} - {space.Name}" : space.Name;
                return $"Space: {display}";
            }

            return null;
        }

        private void SendOperationResult(bool success, string message, bool refreshAudit = false)
        {
            var result = new { type = "operationResult", success, message, refreshAudit };
            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            PostResult?.Invoke(json);
        }

        private void SendCategoryStats(string category, int count)
        {
            var result = new { type = "categoryStats", category, count };
            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            PostResult?.Invoke(json);
        }

        private static BuiltInCategory? GetBuiltInCategory(Element element)
        {
            var category = element.Category;
            if (category == null) return null;

            var categoryId = category.Id;
#if REVIT2026_OR_GREATER
            var catValue = unchecked((int)categoryId.Value);
#else
            var catValue = categoryId.IntegerValue;
#endif
            return Enum.IsDefined(typeof(BuiltInCategory), catValue) ? (BuiltInCategory)catValue : null;
        }

        private static int GetElementIdValue(ElementId id)
        {
#if REVIT2026_OR_GREATER
            return unchecked((int)id.Value);
#else
            return id.IntegerValue;
#endif
        }

        public string GetName() => "FMReadiness_v3.ParameterEditorHandler";
    }
}

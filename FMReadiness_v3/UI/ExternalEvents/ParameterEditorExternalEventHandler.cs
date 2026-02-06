using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Globalization;
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
    /// Supports COBie presets and dynamic field mapping.
    /// </summary>
    public class ParameterEditorExternalEventHandler : IExternalEventHandler
    {
        public enum OperationType
        {
            GetSelectedElements,
            FocusFixElement,
            SetSelectionElements,
            ApplySelectionScope,
            GetCategoryStats,
            SetInstanceParams,
            SetCategoryParams,
            SetTypeParams,
            CopyComputedToParam,
            RefreshAudit,
            // COBie preset operations
            GetAvailablePresets,
            LoadPreset,
            GetPresetFields,
            SetCobieFieldValues,
            ValidateCobieReadiness,
            EnsureCobieParameters,
            // COBie editor operations
            SetCobieTypeFieldValues,
            SetCobieCategoryFieldValues,
            CopyComputedToCobieParam
        }

        public OperationType CurrentOperation { get; set; }
        public string? JsonPayload { get; set; }

        // Callback to send results back to WebView
        public Action<string>? PostResult { get; set; }

        // Preset service for COBie mapping
        private readonly PresetService _presetService = new PresetService();
        private CobieMappingService? _mappingService;

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
                    case OperationType.FocusFixElement:
                        HandleFocusFixElement(uidoc, doc);
                        break;
                    case OperationType.SetSelectionElements:
                        HandleSetSelectionElements(uidoc, doc);
                        break;
                    case OperationType.ApplySelectionScope:
                        HandleApplySelectionScope(uidoc, doc);
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
                    // New COBie preset operations
                    case OperationType.GetAvailablePresets:
                        HandleGetAvailablePresets();
                        break;
                    case OperationType.LoadPreset:
                        HandleLoadPreset();
                        break;
                    case OperationType.GetPresetFields:
                        HandleGetPresetFields();
                        break;
                    case OperationType.SetCobieFieldValues:
                        HandleSetCobieFieldValues(doc);
                        break;
                    case OperationType.ValidateCobieReadiness:
                        HandleValidateCobieReadiness(uidoc, doc);
                        break;
                    case OperationType.EnsureCobieParameters:
                        HandleEnsureCobieParameters(app, doc);
                        break;
                    // COBie editor operations
                    case OperationType.SetCobieTypeFieldValues:
                        HandleSetCobieTypeFieldValues(doc);
                        break;
                    case OperationType.SetCobieCategoryFieldValues:
                        HandleSetCobieCategoryFieldValues(doc);
                        break;
                    case OperationType.CopyComputedToCobieParam:
                        HandleCopyComputedToCobieParam(doc);
                        break;
                }
            }
            catch (Exception ex)
            {
                LogException(ex);
                SendOperationResult(false, $"Error in {CurrentOperation}: {ex.Message}");
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
                var resolver = new AuditProfileResolverService();
                if (!resolver.TryResolveRules(out var rules, out var profileName, out var errorMessage))
                {
                    SendOperationResult(false, errorMessage);
                    return;
                }

                var collector = new CollectorService(doc);
                var elements = collector.GetAllFmElements();
                var scoreMode = AuditProfileState.GetScoreMode();

                var auditService = new AuditService();
                var report = auditService.RunFullAudit(doc, elements, rules, scoreMode);
                report.AuditProfileName = profileName;

                WebViewPaneController.UpdateFromReport(report);

                SendOperationResult(true, "Audit refreshed", false);
            }
            catch (Exception ex)
            {
                SendOperationResult(false, $"Failed to refresh audit: {ex.Message}");
            }
        }

        #region COBie Preset Handlers

        /// <summary>
        /// Returns list of available preset configurations.
        /// </summary>
        private void HandleGetAvailablePresets()
        {
            var presets = _presetService.GetAvailablePresets();
            var currentPreset = _presetService.CurrentPresetName;

            var result = new
            {
                type = "availablePresets",
                presets = presets.Select(p => new
                {
                    fileName = p.FileName,
                    name = p.Name,
                    description = p.Description,
                    version = p.Version
                }).ToList(),
                currentPreset = currentPreset
            };

            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            PostResult?.Invoke(json);
        }

        /// <summary>
        /// Loads a preset by filename.
        /// </summary>
        private void HandleLoadPreset()
        {
            if (string.IsNullOrEmpty(JsonPayload))
            {
                SendOperationResult(false, "No preset specified");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var fileName = jsonDoc.RootElement.GetProperty("fileName").GetString();

            if (string.IsNullOrEmpty(fileName))
            {
                SendOperationResult(false, "Invalid preset filename");
                return;
            }

            if (_presetService.LoadPreset(fileName))
            {
                _mappingService = new CobieMappingService(_presetService);
                var preset = _presetService.CurrentPreset;
                AuditProfileState.SetActivePreset(_presetService.CurrentPresetName, preset?.Name);

                var result = new
                {
                    type = "presetLoaded",
                    success = true,
                    preset = new
                    {
                        name = preset?.Name,
                        version = preset?.Version,
                        description = preset?.Description,
                        writeAliases = preset?.WriteAliases ?? true,
                        categories = preset?.Categories ?? new List<string>()
                    }
                };

                var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
                PostResult?.Invoke(json);
            }
            else
            {
                SendOperationResult(false, $"Failed to load preset: {fileName}");
            }
        }

        /// <summary>
        /// Returns field definitions from current preset organized by group.
        /// </summary>
        private void HandleGetPresetFields()
        {
            if (_presetService.CurrentPreset == null)
            {
                _presetService.LoadDefaultPreset();
                _mappingService = new CobieMappingService(_presetService);
            }

            AuditProfileState.SetActivePreset(_presetService.CurrentPresetName, _presetService.CurrentPreset?.Name);

            var fieldsByGroup = _presetService.GetFieldsByGroup("Component");
            var typeFields = _presetService.GetFieldsByGroup("Type");

            var groups = new List<object>();
            foreach (var groupEntry in fieldsByGroup)
            {
                groups.Add(new
                {
                    name = groupEntry.Key,
                    fields = groupEntry.Value.Select(f => new
                    {
                        cobieKey = f.CobieKey,
                        label = f.Label,
                        scope = f.Scope,
                        dataType = f.DataType,
                        required = f.Required,
                        revitParam = f.RevitParam,
                        aliasParams = f.AliasParams ?? new List<string>(),
                        rules = f.Rules ?? new List<string>(),
                        hasComputed = f.Computed != null,
                        computedSource = f.Computed?.Source
                    }).ToList()
                });
            }

            var typeGroups = new List<object>();
            foreach (var groupEntry in typeFields)
            {
                typeGroups.Add(new
                {
                    name = groupEntry.Key,
                    fields = groupEntry.Value.Where(f => f.Scope == "type").Select(f => new
                    {
                        cobieKey = f.CobieKey,
                        label = f.Label,
                        scope = f.Scope,
                        dataType = f.DataType,
                        required = f.Required,
                        revitParam = f.RevitParam,
                        aliasParams = f.AliasParams ?? new List<string>(),
                        rules = f.Rules ?? new List<string>()
                    }).ToList()
                });
            }

            var result = new
            {
                type = "presetFields",
                presetName = _presetService.CurrentPreset?.Name ?? "Default",
                componentGroups = groups,
                typeGroups = typeGroups
            };

            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            PostResult?.Invoke(json);
        }

        /// <summary>
        /// Sets COBie field values using the mapping service with alias support.
        /// </summary>
        private void HandleSetCobieFieldValues(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload))
            {
                SendOperationResult(false, "No payload");
                return;
            }

            EnsureMappingService();

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var root = jsonDoc.RootElement;
            var elementIds = root.GetProperty("elementIds")
                .EnumerateArray()
                .Select(e => new ElementId(e.GetInt32()))
                .ToList();

            var writeAliases = true;
            if (root.TryGetProperty("writeAliases", out var writeAliasesProp))
                writeAliases = writeAliasesProp.GetBoolean();

            var fieldValues = new Dictionary<string, string>();

            if (root.TryGetProperty("fields", out var fieldsProp))
            {
                foreach (var prop in fieldsProp.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.String)
                    {
                        fieldValues[prop.Name] = prop.Value.GetString() ?? string.Empty;
                    }
                    else if (prop.Value.TryGetProperty("value", out var valueProp))
                    {
                        fieldValues[prop.Name] = valueProp.GetString() ?? string.Empty;
                    }
                }
            }
            else if (root.TryGetProperty("fieldValues", out var fieldValuesProp))
            {
                foreach (var prop in fieldValuesProp.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.String)
                    {
                        fieldValues[prop.Name] = prop.Value.GetString() ?? string.Empty;
                    }
                    else if (prop.Value.TryGetProperty("value", out var valueProp))
                    {
                        fieldValues[prop.Name] = valueProp.GetString() ?? string.Empty;
                    }
                }
            }

            if (fieldValues.Count == 0)
            {
                SendOperationResult(false, "No field values provided");
                return;
            }

            int totalFieldsUpdated = 0;
            int totalParamsUpdated = 0;
            var originalPolicy = _mappingService!.CurrentWritePolicy;
            _mappingService.CurrentWritePolicy = writeAliases
                ? CobieMappingService.WritePolicy.PrimaryAndAliases
                : CobieMappingService.WritePolicy.PrimaryOnly;

            try
            {
                if (!RunInTransaction(doc, "Set COBie Field Values", () =>
                {
                    foreach (var id in elementIds)
                    {
                        var element = doc.GetElement(id);
                        if (element == null) continue;

                        var typeId = element.GetTypeId();
                        Element? typeElement = typeId != ElementId.InvalidElementId
                            ? doc.GetElement(typeId)
                            : null;

                        var (fieldsUpdated, paramsUpdated) = _mappingService!.SetFieldValues(
                            doc, element, typeElement, fieldValues);

                        totalFieldsUpdated += fieldsUpdated;
                        totalParamsUpdated += paramsUpdated;
                    }
                }, out var error))
                {
                    SendOperationResult(false, error);
                    return;
                }
            }
            finally
            {
                _mappingService.CurrentWritePolicy = originalPolicy;
            }

            SendOperationResult(true,
                $"Updated {totalFieldsUpdated} fields ({totalParamsUpdated} parameters) on {elementIds.Count} elements",
                true);
        }

        /// <summary>
        /// Validates COBie readiness for selected elements.
        /// </summary>
        private void HandleValidateCobieReadiness(UIDocument uidoc, Document doc)
        {
            EnsureMappingService();

            var selectedIds = uidoc.Selection.GetElementIds();
            if (selectedIds.Count == 0)
            {
                SendOperationResult(false, "No elements selected");
                return;
            }

            var elementResults = new List<object>();
            var allFieldValues = new List<(int elementId, Dictionary<string, FieldValueResult> values)>();
            var cobieReadyCount = 0;
            var scoreTotal = 0.0;

            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element == null || !IsTargetCategory(element)) continue;

                var typeId = element.GetTypeId();
                Element? typeElement = typeId != ElementId.InvalidElementId
                    ? doc.GetElement(typeId)
                    : null;

                var fieldValues = _mappingService!.GetAllFieldValues(element, typeElement, doc);
                var validationErrors = _mappingService.ValidateFieldValues(fieldValues);
                var score = _mappingService.CalculateReadinessScore(fieldValues, validationErrors);
                var roundedOverallScore = Math.Round(score.OverallScore, 1);
                if (score.IsCobieReady)
                    cobieReadyCount++;
                scoreTotal += roundedOverallScore;

                allFieldValues.Add((GetElementIdValue(id), fieldValues));

                elementResults.Add(new
                {
                    elementId = GetElementIdValue(id),
                    category = element.Category?.Name ?? "Unknown",
                    cobieReady = score.IsCobieReady,
                    overallScore = roundedOverallScore,
                    requiredScore = Math.Round(score.RequiredFieldsScore, 1),
                    optionalScore = Math.Round(score.OptionalFieldsScore, 1),
                    requiredFields = new
                    {
                        total = score.TotalRequiredFields,
                        populated = score.PopulatedRequiredFields
                    },
                    optionalFields = new
                    {
                        total = score.TotalOptionalFields,
                        populated = score.PopulatedOptionalFields
                    },
                    validationErrors = validationErrors.Select(e => new
                    {
                        field = e.Label,
                        rule = e.Rule,
                        message = e.Message
                    }).ToList(),
                    missingRequired = fieldValues
                        .Where(fv => fv.Value.Required && !fv.Value.HasValue)
                        .Select(fv => fv.Value.Label)
                        .ToList()
                });
            }

            // Check uniqueness violations
            var uniquenessViolations = _mappingService!.CheckUniqueness(allFieldValues);

            var result = new
            {
                type = "cobieReadinessResult",
                elements = elementResults,
                uniquenessViolations = uniquenessViolations.Select(kv => new
                {
                    field = kv.Key,
                    elementIds = kv.Value
                }).ToList(),
                summary = new
                {
                    totalElements = elementResults.Count,
                    cobieReadyCount = cobieReadyCount,
                    averageScore = elementResults.Count > 0
                        ? Math.Round(scoreTotal / elementResults.Count, 1)
                        : 0
                }
            };

            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            PostResult?.Invoke(json);
        }

        private void EnsureMappingService()
        {
            if (_mappingService == null)
            {
                if (_presetService.CurrentPreset == null)
                    _presetService.LoadDefaultPreset();
                _mappingService = new CobieMappingService(_presetService);
            }
        }

        #endregion

        #region COBie Editor Handlers

        private void HandleEnsureCobieParameters(UIApplication app, Document doc)
        {
            if (_presetService.CurrentPreset == null)
                _presetService.LoadDefaultPreset();

            if (!string.IsNullOrEmpty(JsonPayload))
            {
                using var jsonDoc = JsonDocument.Parse(JsonPayload);
                if (jsonDoc.RootElement.TryGetProperty("presetFile", out var presetFileProp))
                {
                    var presetFile = presetFileProp.GetString();
                    if (!string.IsNullOrWhiteSpace(presetFile))
                    {
                        if (!_presetService.LoadPreset(presetFile))
                        {
                            SendOperationResult(false, $"Failed to load preset: {presetFile}");
                            return;
                        }
                    }
                }
            }

            var preset = _presetService.CurrentPreset;
            if (preset == null)
            {
                SendOperationResult(false, "No preset loaded");
                return;
            }

            AuditProfileState.SetActivePreset(_presetService.CurrentPresetName, preset.Name);

            bool includeAliases = true;
            bool replaceCobie = false;
            bool removeFmAliases = false;
            string mode = "cobie";
            if (!string.IsNullOrEmpty(JsonPayload))
            {
                using var jsonDoc = JsonDocument.Parse(JsonPayload);
                if (jsonDoc.RootElement.TryGetProperty("includeAliases", out var includeAliasesProp))
                    includeAliases = includeAliasesProp.GetBoolean();
                if (jsonDoc.RootElement.TryGetProperty("replaceCobie", out var replaceCobieProp))
                    replaceCobie = replaceCobieProp.GetBoolean();
                if (jsonDoc.RootElement.TryGetProperty("removeFmAliases", out var removeFmAliasesProp))
                    removeFmAliases = removeFmAliasesProp.GetBoolean();
                if (jsonDoc.RootElement.TryGetProperty("mode", out var modeProp))
                    mode = modeProp.GetString() ?? "cobie";
            }

            if (removeFmAliases)
            {
                includeAliases = false;
            }

            var service = new CobieParameterService();
            var result = service.EnsureParameters(app, doc, preset, includeAliases, replaceCobie, removeFmAliases);

            var message = $"Created {result.Created}, updated {result.UpdatedBindings}, skipped {result.Skipped}.";
            if (result.Removed > 0)
            {
                message += $" Removed {result.Removed} COBie parameters.";
            }
            if (result.Warnings.Count > 0)
            {
                message += $" Warnings: {string.Join(" | ", result.Warnings)}";
            }

            SendOperationResult(true, message, refreshAudit: true);

            if (string.Equals(mode, "cobie", StringComparison.OrdinalIgnoreCase))
            {
                var ensuredPayload = new
                {
                    type = "cobieParametersEnsured",
                    success = true,
                    created = result.Created,
                    updated = result.UpdatedBindings,
                    skipped = result.Skipped,
                    warnings = result.Warnings
                };
                var json = JsonSerializer.Serialize(ensuredPayload, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
                PostResult?.Invoke(json);
            }
            else if (string.Equals(mode, "legacy", StringComparison.OrdinalIgnoreCase))
            {
                var ensuredPayload = new
                {
                    type = "legacyParametersEnsured",
                    success = true,
                    created = result.Created,
                    updated = result.UpdatedBindings,
                    skipped = result.Skipped,
                    removed = result.Removed,
                    warnings = result.Warnings
                };
                var json = JsonSerializer.Serialize(ensuredPayload, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
                PostResult?.Invoke(json);
            }
        }

        private void HandleSetCobieTypeFieldValues(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload))
            {
                SendOperationResult(false, "No payload provided");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var root = jsonDoc.RootElement;

            if (!root.TryGetProperty("typeId", out var typeIdProp))
            {
                SendOperationResult(false, "No type ID provided");
                return;
            }

            var typeId = new ElementId(typeIdProp.GetInt32());
            var typeElement = doc.GetElement(typeId) as ElementType;
            if (typeElement == null)
            {
                SendOperationResult(false, "Type element not found");
                return;
            }

            var writeAliases = true;
            if (root.TryGetProperty("writeAliases", out var writeAliasesProp))
                writeAliases = writeAliasesProp.GetBoolean();

            var fieldValues = new Dictionary<string, string>();
            if (root.TryGetProperty("fieldValues", out var fieldValuesProp))
            {
                foreach (var prop in fieldValuesProp.EnumerateObject())
                {
                    if (prop.Value.TryGetProperty("value", out var valueProp))
                    {
                        fieldValues[prop.Name] = valueProp.GetString() ?? "";
                    }
                }
            }

            if (fieldValues.Count == 0)
            {
                SendOperationResult(false, "No field values provided");
                return;
            }

            EnsureMappingService();

            using (var tx = new Transaction(doc, "Set COBie Type Fields"))
            {
                tx.Start();
                int updatedCount = 0;

                foreach (var kvp in fieldValues)
                {
                    var field = FindFieldByKey(kvp.Key);
                    if (field != null)
                    {
                        var success = _mappingService!.SetFieldValue(typeElement, field, kvp.Value, writeAliases);
                        if (success) updatedCount++;
                    }
                    else
                    {
                        // Try setting directly by parameter name
                        var param = typeElement.LookupParameter(kvp.Key);
                        if (param != null && !param.IsReadOnly)
                        {
                            if (TrySetParamValue(param, kvp.Value))
                                updatedCount++;
                        }
                    }
                }

                tx.Commit();
                SendOperationResult(true, $"Updated {updatedCount} type fields", refreshAudit: true);
            }
        }

        private void HandleSetCobieCategoryFieldValues(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload))
            {
                SendOperationResult(false, "No payload provided");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var root = jsonDoc.RootElement;

            if (!root.TryGetProperty("category", out var categoryProp))
            {
                SendOperationResult(false, "No category provided");
                return;
            }

            var categoryStr = categoryProp.GetString();
            if (!Enum.TryParse<BuiltInCategory>(categoryStr, out var category))
            {
                SendOperationResult(false, $"Invalid category: {categoryStr}");
                return;
            }

            var onlyBlanks = true;
            if (root.TryGetProperty("onlyBlanks", out var onlyBlanksProp))
                onlyBlanks = onlyBlanksProp.GetBoolean();

            var writeAliases = true;
            if (root.TryGetProperty("writeAliases", out var writeAliasesProp))
                writeAliases = writeAliasesProp.GetBoolean();

            var fieldValues = new Dictionary<string, string>();
            if (root.TryGetProperty("fieldValues", out var fieldValuesProp))
            {
                foreach (var prop in fieldValuesProp.EnumerateObject())
                {
                    if (prop.Value.TryGetProperty("value", out var valueProp))
                    {
                        fieldValues[prop.Name] = valueProp.GetString() ?? "";
                    }
                }
            }

            if (fieldValues.Count == 0)
            {
                SendOperationResult(false, "No field values provided");
                return;
            }

            EnsureMappingService();

            var elements = new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToElements();

            using (var tx = new Transaction(doc, "Set COBie Category Fields"))
            {
                tx.Start();
                int updatedElements = 0;

                foreach (var element in elements)
                {
                    bool anyUpdated = false;
                    foreach (var kvp in fieldValues)
                    {
                        var field = FindFieldByKey(kvp.Key);
                        if (field != null)
                        {
                            if (onlyBlanks)
                            {
                                var existing = _mappingService!.ResolveFieldValue(element, null, field, doc);
                                if (existing.success && !string.IsNullOrWhiteSpace(existing.value))
                                    continue;
                            }

                            var success = _mappingService!.SetFieldValue(element, field, kvp.Value, writeAliases);
                            if (success) anyUpdated = true;
                        }
                        else
                        {
                            // Try setting directly by parameter name
                            var param = element.LookupParameter(kvp.Key);
                            if (param != null && !param.IsReadOnly)
                            {
                                if (onlyBlanks && !string.IsNullOrWhiteSpace(param.AsString() ?? param.AsValueString()))
                                    continue;

                                if (TrySetParamValue(param, kvp.Value))
                                    anyUpdated = true;
                            }
                        }
                    }

                    if (anyUpdated) updatedElements++;
                }

                tx.Commit();
                SendOperationResult(true, $"Updated {updatedElements} elements", refreshAudit: true);
            }
        }

        private void HandleCopyComputedToCobieParam(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload))
            {
                SendOperationResult(false, "No payload provided");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var root = jsonDoc.RootElement;

            if (!root.TryGetProperty("elementIds", out var elementIdsProp))
            {
                SendOperationResult(false, "No element IDs provided");
                return;
            }

            if (!root.TryGetProperty("sourceField", out var sourceFieldProp))
            {
                SendOperationResult(false, "No source field provided");
                return;
            }

            if (!root.TryGetProperty("targetCobieKey", out var targetKeyProp))
            {
                SendOperationResult(false, "No target COBie key provided");
                return;
            }

            var sourceField = sourceFieldProp.GetString();
            var targetCobieKey = targetKeyProp.GetString();

            var writeAliases = true;
            if (root.TryGetProperty("writeAliases", out var writeAliasesProp))
                writeAliases = writeAliasesProp.GetBoolean();

            EnsureMappingService();

            var targetField = FindFieldByKey(targetCobieKey);
            if (targetField == null)
            {
                SendOperationResult(false, $"COBie field not found: {targetCobieKey}");
                return;
            }

            using (var tx = new Transaction(doc, "Copy Computed to COBie"))
            {
                tx.Start();
                int updatedCount = 0;

                foreach (var idElem in elementIdsProp.EnumerateArray())
                {
                    var elementId = new ElementId(idElem.GetInt32());
                    var element = doc.GetElement(elementId);
                    if (element == null) continue;

                    string? computedValue = null;

                    switch (sourceField)
                    {
                        case "UniqueId":
                            computedValue = element.UniqueId;
                            break;
                        case "RoomSpace":
                            computedValue = GetRoomOrSpaceName(element, doc);
                            break;
                        case "Level":
                            computedValue = GetLevelName(element, doc);
                            break;
                    }

                    if (!string.IsNullOrEmpty(computedValue))
                    {
                        var success = _mappingService!.SetFieldValue(element, targetField, computedValue, writeAliases);
                        if (success) updatedCount++;
                    }
                }

                tx.Commit();
                SendOperationResult(true, $"Updated {updatedCount} elements", refreshAudit: true);
            }
        }

        private CobieFieldSpec? FindFieldByKey(string? cobieKey)
        {
            if (string.IsNullOrEmpty(cobieKey)) return null;

            var preset = _presetService.CurrentPreset;
            if (preset?.Tables == null) return null;

            foreach (var table in preset.Tables.Values)
            {
                if (table.Fields == null) continue;
                var field = table.Fields.FirstOrDefault(f =>
                    string.Equals(f.CobieKey, cobieKey, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(f.RevitParam, cobieKey, StringComparison.OrdinalIgnoreCase));
                if (field != null) return field;
            }

            return null;
        }

        private string? GetRoomOrSpaceName(Element element, Document doc)
        {
            try
            {
                if (element is FamilyInstance fi && fi.Space != null)
                    return fi.Space.Name;

                if (element is FamilyInstance fi2 && fi2.Room != null)
                    return fi2.Room.Name;

                var point = GetElementPoint(element);
                if (point != null)
                {
                    var room = doc.GetRoomAtPoint(point);
                    if (room != null) return room.Name;

                    var spaces = new FilteredElementCollector(doc)
                        .OfClass(typeof(Space))
                        .Cast<Space>()
                        .Where(s => s.IsPointInSpace(point))
                        .ToList();

                    if (spaces.Count > 0) return spaces[0].Name;
                }
            }
            catch { }

            return null;
        }

        private string? GetLevelName(Element element, Document doc)
        {
            try
            {
                var levelId = element.LevelId;
                if (levelId != null && levelId != ElementId.InvalidElementId)
                {
                    var level = doc.GetElement(levelId) as Level;
                    if (level != null) return level.Name;
                }

                var levelParam = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                if (levelParam != null && levelParam.HasValue)
                {
                    var paramLevelId = levelParam.AsElementId();
                    if (paramLevelId != null && paramLevelId != ElementId.InvalidElementId)
                    {
                        var level = doc.GetElement(paramLevelId) as Level;
                        if (level != null) return level.Name;
                    }
                }
            }
            catch { }

            return null;
        }

        private XYZ? GetElementPoint(Element element)
        {
            if (element.Location is LocationPoint lp)
                return lp.Point;

            if (element.Location is LocationCurve lc)
                return lc.Curve.Evaluate(0.5, true);

            var bb = element.get_BoundingBox(null);
            if (bb != null)
                return (bb.Min + bb.Max) * 0.5;

            return null;
        }

        #endregion

        private sealed class SelectedElementsSnapshot
        {
            public IReadOnlyList<int> ElementIds { get; }
            public string Json { get; }

            public SelectedElementsSnapshot(IReadOnlyList<int> elementIds, string json)
            {
                ElementIds = elementIds;
                Json = json;
            }
        }

        internal bool TryBuildSelectedElementsSnapshot(UIDocument uidoc, Document doc, out IReadOnlyList<int> elementIds, out string json)
        {
            elementIds = Array.Empty<int>();
            json = string.Empty;

            try
            {
                var snapshot = BuildSelectedElementsSnapshot(uidoc, doc);
                elementIds = snapshot.ElementIds;
                json = snapshot.Json;
                return true;
            }
            catch (Exception ex)
            {
                LogException(ex);
                return false;
            }
        }

        private void HandleGetSelectedElements(UIDocument uidoc, Document doc)
        {
            if (TryBuildSelectedElementsSnapshot(uidoc, doc, out _, out var json))
            {
                PostResult?.Invoke(json);
            }
        }

        private void HandleFocusFixElement(UIDocument uidoc, Document doc)
        {
            if (string.IsNullOrWhiteSpace(JsonPayload))
            {
                SendOperationResult(false, "No element specified for fix");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            if (!jsonDoc.RootElement.TryGetProperty("elementId", out var elementIdProp)
                || !elementIdProp.TryGetInt32(out var elementIdInt))
            {
                SendOperationResult(false, "Invalid element id for fix");
                return;
            }

            var elementId = new ElementId(elementIdInt);
            var element = doc.GetElement(elementId);
            if (element == null)
            {
                SendOperationResult(false, $"Element {elementIdInt} not found");
                return;
            }

            uidoc.Selection.SetElementIds(new List<ElementId> { elementId });

            try
            {
                uidoc.ShowElements(elementId);
            }
            catch
            {
                // Ignore view-focus failures and still return selected element data.
            }

            var selectionIds = uidoc.Selection.GetElementIds();
            if (selectionIds != null && selectionIds.Contains(elementId))
            {
                var snapshot = BuildSelectedElementsSnapshot(doc, selectionIds);
                PostResult?.Invoke(snapshot.Json);
            }
            else
            {
                var snapshot = BuildSelectedElementsSnapshot(doc, new List<ElementId> { elementId });
                PostResult?.Invoke(snapshot.Json);
            }
        }

        private void HandleSetSelectionElements(UIDocument uidoc, Document doc)
        {
            if (string.IsNullOrWhiteSpace(JsonPayload))
            {
                SendOperationResult(false, "No selection payload");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            if (!jsonDoc.RootElement.TryGetProperty("elementIds", out var idsProp)
                || idsProp.ValueKind != JsonValueKind.Array)
            {
                SendOperationResult(false, "No elementIds provided");
                return;
            }

            var targetIds = new List<ElementId>();
            foreach (var idProp in idsProp.EnumerateArray())
            {
                if (idProp.ValueKind != JsonValueKind.Number || !idProp.TryGetInt32(out var idValue))
                    continue;

                var elementId = new ElementId(idValue);
                var element = doc.GetElement(elementId);
                if (element == null || !IsTargetCategory(element))
                    continue;

                targetIds.Add(elementId);
            }

            uidoc.Selection.SetElementIds(targetIds);

            if (targetIds.Count == 1)
            {
                try
                {
                    uidoc.ShowElements(targetIds[0]);
                }
                catch
                {
                    // Ignore view-focus failures.
                }
            }

            var snapshot = BuildSelectedElementsSnapshot(doc, uidoc.Selection.GetElementIds());
            PostResult?.Invoke(snapshot.Json);
        }

        private SelectedElementsSnapshot BuildSelectedElementsSnapshot(UIDocument uidoc, Document doc)
        {
            var selectedIds = uidoc.Selection.GetElementIds();
            return BuildSelectedElementsSnapshot(doc, selectedIds);
        }

        private SelectedElementsSnapshot BuildSelectedElementsSnapshot(Document doc, ICollection<ElementId> selectedIds)
        {
            var elements = new List<object>();
            var includedIds = new List<int>();
            // Allow COBie values for multi-select (up to 100 elements for performance)
            var includeCobieValues = selectedIds.Count > 0 && selectedIds.Count <= 100;

            if (includeCobieValues)
            {
                EnsureMappingService();
            }

            foreach (var id in selectedIds)
            {
                var element = doc.GetElement(id);
                if (element == null) continue;

                // Skip non-MEP elements
                if (!IsTargetCategory(element))
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

                Dictionary<string, string?>? cobieInstanceParams = null;
                Dictionary<string, string?>? cobieTypeParams = null;

                if (includeCobieValues)
                {
                    var cobieInstanceValues = _mappingService!.GetAllFieldValues(element, elementType, doc, "Component");
                    cobieInstanceParams = cobieInstanceValues.ToDictionary(
                        kvp => kvp.Key,
                        kvp => kvp.Value.Value);

                    var cobieTypeValues = _mappingService!.GetAllFieldValues(element, elementType, doc, "Type");
                    cobieTypeParams = cobieTypeValues.ToDictionary(
                        kvp => kvp.Key,
                        kvp => kvp.Value.Value);
                }

                var elementIdValue = GetElementIdValue(element.Id);
                includedIds.Add(elementIdValue);

                elements.Add(new
                {
                    elementId = elementIdValue,
                    category = element.Category?.Name ?? "Unknown",
                    family = elementType?.FamilyName ?? "Unknown",
                    typeName = elementType?.Name ?? "Unknown",
                    typeId = typeId != ElementId.InvalidElementId ? (int?)GetElementIdValue(typeId) : null,
                    typeInstanceCount = typeInstanceCount,
                    instanceParams = instanceParams,
                    typeParams = typeParams,
                    computed = computed,
                    cobieInstanceParams = cobieInstanceParams,
                    cobieTypeParams = cobieTypeParams
                });
            }

            includedIds.Sort();

            var result = new { type = "selectedElementsData", elements = elements };
            var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
            return new SelectedElementsSnapshot(includedIds, json);
        }

        private void HandleApplySelectionScope(UIDocument uidoc, Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload))
            {
                SendOperationResult(false, "No scope specified");
                return;
            }

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var root = jsonDoc.RootElement;
            if (!root.TryGetProperty("scope", out var scopeProp))
            {
                SendOperationResult(false, "No scope specified");
                return;
            }

            var scope = scopeProp.GetString();
            if (string.IsNullOrWhiteSpace(scope))
            {
                SendOperationResult(false, "Invalid scope");
                return;
            }

            List<ElementId> targetIds;

            switch (scope)
            {
                case "selection":
                    targetIds = uidoc.Selection.GetElementIds().ToList();
                    break;
                case "activeView":
                    var view = doc.ActiveView;
                    if (view == null)
                    {
                        SendOperationResult(false, "No active view");
                        return;
                    }

                    targetIds = new FilteredElementCollector(doc, view.Id)
                        .WhereElementIsNotElementType()
                        .WherePasses(CreateTargetCategoryFilter())
                        .ToElementIds()
                        .ToList();
                    break;
                case "category":
                    if (!root.TryGetProperty("category", out var categoryProp))
                    {
                        SendOperationResult(false, "No category specified");
                        return;
                    }

                    var categoryStr = categoryProp.GetString();
                    if (!TryParseBuiltInCategory(categoryStr, out var category))
                    {
                        SendOperationResult(false, "Invalid category");
                        return;
                    }

                    if (!IsTargetCategory(category))
                    {
                        SendOperationResult(false, "Category not supported in editor");
                        return;
                    }

                    targetIds = new FilteredElementCollector(doc)
                        .OfCategory(category)
                        .WhereElementIsNotElementType()
                        .ToElementIds()
                        .ToList();
                    break;
                case "selectedType":
                    ElementId? typeId = null;
                    if (root.TryGetProperty("typeId", out var typeIdProp)
                        && typeIdProp.ValueKind == JsonValueKind.Number)
                    {
                        typeId = new ElementId(typeIdProp.GetInt32());
                    }

                    if (typeId == null || typeId == ElementId.InvalidElementId)
                    {
                        var selectedTypeIds = uidoc.Selection.GetElementIds()
                            .Select(id => doc.GetElement(id))
                            .Where(e => e != null && IsTargetCategory(e))
                            .Select(e => e!.GetTypeId())
                            .Where(id => id != ElementId.InvalidElementId)
                            .Distinct()
                            .ToList();

                        if (selectedTypeIds.Count != 1)
                        {
                            SendOperationResult(false, "Select elements of a single type");
                            return;
                        }

                        typeId = selectedTypeIds[0];
                    }

                    var resolvedTypeId = typeId;
                    targetIds = new FilteredElementCollector(doc)
                        .WhereElementIsNotElementType()
                        .WherePasses(CreateTargetCategoryFilter())
                        .Where(e => e.GetTypeId() == resolvedTypeId)
                        .Select(e => e.Id)
                        .ToList();
                    break;
                default:
                    SendOperationResult(false, "Unsupported scope");
                    return;
            }

            uidoc.Selection.SetElementIds(targetIds);

            HandleGetSelectedElements(uidoc, doc);

            if (targetIds.Count == 0)
            {
                SendOperationResult(false, "No elements found for scope");
            }
            else
            {
                SendOperationResult(true, $"Selected {targetIds.Count} elements");
            }
        }

        private void HandleGetCategoryStats(Document doc)
        {
            if (string.IsNullOrEmpty(JsonPayload)) return;

            using var jsonDoc = JsonDocument.Parse(JsonPayload);
            var categoryStr = jsonDoc.RootElement.GetProperty("category").GetString();

            if (!TryParseBuiltInCategory(categoryStr, out var category))
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
            if (!RunInTransaction(doc, "Set FM Instance Parameters", () =>
            {
                foreach (var id in elementIds)
                {
                    var element = doc.GetElement(id);
                    if (element == null) continue;

                    foreach (var kvp in paramsToSet)
                    {
                        if (SetInstanceParamValue(doc, element, kvp.Key, kvp.Value))
                            updated++;
                    }
                }
            }, out var error))
            {
                SendOperationResult(false, error);
                return;
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

            if (!TryParseBuiltInCategory(categoryStr, out var category))
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

            if (!RunInTransaction(doc, "Bulk Set FM Parameters", () =>
            {
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

                        if (SetInstanceParamValue(doc, element, kvp.Key, kvp.Value))
                            updated++;
                    }
                }
            }, out var error))
            {
                SendOperationResult(false, error);
                return;
            }

            var msg = $"Updated {updated} values on {elements.Count} elements";
            if (skipped > 0) msg += $" ({skipped} skipped - not blank)";
            SendOperationResult(true, msg, true);
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
            if (!RunInTransaction(doc, "Set FM Type Parameters", () =>
            {
                foreach (var kvp in paramsToSet)
                {
                    if (kvp.Key == "TypeMark")
                    {
                        var param = elementType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK);
                        if (param != null && !param.IsReadOnly)
                        {
                            if (TrySetParamValue(param, kvp.Value))
                                updated++;
                        }
                    }
                    else if (SetParamValue(elementType, kvp.Key, kvp.Value))
                    {
                        updated++;
                    }
                }
            }, out var error))
            {
                SendOperationResult(false, error);
                return;
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
            if (!RunInTransaction(doc, "Copy Computed to Parameter", () =>
            {
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
                        if (SetInstanceParamValue(doc, element, targetParam, value))
                            updated++;
                    }
                }
            }, out var error))
            {
                SendOperationResult(false, error);
                return;
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
                return TrySetParamValue(param, value);
            }
            catch { }
            return false;
        }

        private bool SetInstanceParamValue(Document doc, Element element, string paramName, string? value)
        {
            var param = element.LookupParameter(paramName);
            if (param == null || param.IsReadOnly) return false;

            if (IsTypeParameter(doc, element, paramName, param))
                return false;

            try
            {
                return TrySetParamValue(param, value);
            }
            catch
            {
                return false;
            }
        }

        private static bool TrySetParamValue(Parameter param, string? value)
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

            return false;
        }

        private static bool IsTypeParameter(Document doc, Element element, string paramName, Parameter param)
        {
            var bindingKind = GetBindingKind(doc, paramName);
            if (bindingKind == ParamBindingKind.Type) return true;
            if (bindingKind == ParamBindingKind.Instance) return false;

            var typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId) return false;

            var typeElement = doc.GetElement(typeId);
            if (typeElement == null) return false;

            var typeParam = typeElement.LookupParameter(paramName);
            if (typeParam == null) return false;

            return typeParam.Id == param.Id;
        }

        private enum ParamBindingKind
        {
            Unknown,
            Instance,
            Type
        }

        private static ParamBindingKind GetBindingKind(Document doc, string paramName)
        {
            if (doc == null || string.IsNullOrWhiteSpace(paramName)) return ParamBindingKind.Unknown;

            var map = doc.ParameterBindings;
            var iterator = map.ForwardIterator();
            iterator.Reset();
            while (iterator.MoveNext())
            {
                if (iterator.Key is not Definition definition) continue;
                if (!definition.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase)) continue;

                if (iterator.Current is InstanceBinding) return ParamBindingKind.Instance;
                if (iterator.Current is TypeBinding) return ParamBindingKind.Type;
                return ParamBindingKind.Unknown;
            }

            return ParamBindingKind.Unknown;
        }

        private static bool RunInTransaction(Document doc, string name, Action action, out string error)
        {
            error = string.Empty;

            if (doc.IsReadOnly)
            {
                error = "Document is read-only. Check out the model or make it editable before updating parameters.";
                return false;
            }

            // Prefer a dedicated transaction when possible.
            try
            {
                using var tx = new Transaction(doc, name);
                var txStatus = tx.Start();
                if (txStatus == TransactionStatus.Started)
                {
                    try
                    {
                        action();
                        tx.Commit();
                        return true;
                    }
                    catch (Exception ex)
                    {
                        tx.RollBack();
                        error = ex.Message;
                        return false;
                    }
                }
            }
            catch
            {
                // Ignore and fall through to sub-transaction or direct execution.
            }

            if (doc.IsModifiable)
            {
                try
                {
                    using var sub = new SubTransaction(doc);
                    var status = sub.Start();
                    if (status == TransactionStatus.Started)
                    {
                        try
                        {
                            action();
                            sub.Commit();
                            return true;
                        }
                        catch (Exception ex)
                        {
                            sub.RollBack();
                            error = ex.Message;
                            return false;
                        }
                    }

                    error = $"Could not start sub-transaction ({status}).";
                    return false;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            error = "Document is not modifiable and a transaction could not be started.";
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

        private static readonly BuiltInCategory[] TargetCategories =
        {
            BuiltInCategory.OST_MechanicalEquipment,
            BuiltInCategory.OST_DuctTerminal,
            BuiltInCategory.OST_DuctAccessory,
            BuiltInCategory.OST_PipeAccessory
        };

        private static readonly HashSet<long> TargetCategoryValues = new(
            TargetCategories.Select(GetBuiltInCategoryValue));

        private static ElementMulticategoryFilter CreateTargetCategoryFilter()
        {
            return new ElementMulticategoryFilter(TargetCategories);
        }

        private static bool IsTargetCategory(Element element)
        {
            var category = element.Category;
            if (category == null) return false;

            var rawValue = GetElementIdValueLong(category.Id);
            return TargetCategoryValues.Contains(rawValue);
        }

        private static bool IsTargetCategory(BuiltInCategory category)
        {
            return TargetCategoryValues.Contains(GetBuiltInCategoryValue(category));
        }

        private static long GetBuiltInCategoryValue(BuiltInCategory category)
        {
            return Convert.ToInt64(category);
        }

        private static long GetElementIdValueLong(ElementId id)
        {
#if REVIT2026_OR_GREATER
            return id.Value;
#else
            return id.IntegerValue;
#endif
        }

        private static bool TryParseBuiltInCategory(string? value, out BuiltInCategory category)
        {
            return TryParseEnumValue(value, out category);
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

        private static int GetElementIdValue(ElementId id)
        {
#if REVIT2026_OR_GREATER
            return unchecked((int)id.Value);
#else
            return id.IntegerValue;
#endif
        }

        private static void LogException(Exception ex)
        {
            try
            {
                var logDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "FMReadiness_v3",
                    "logs");
                Directory.CreateDirectory(logDir);
                var logPath = Path.Combine(logDir, "parameter-editor.log");
                var entry = $"[{DateTime.UtcNow:O}] {ex}\n\n";
                File.AppendAllText(logPath, entry);
            }
            catch
            {
                // Ignore logging failures.
            }
        }

        public string GetName() => "FMReadiness_v3.ParameterEditorHandler";
    }
}

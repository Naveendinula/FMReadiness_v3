using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using FMReadiness_v3.Services;
using FMReadiness_v3.UI.ExternalEvents;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.UI.Panes
{
    public static class WebViewPaneController
    {
        private static AuditWebPane? _paneInstance;
        private static SelectZoomExternalEventHandler? _selectZoomHandler;
        private static ExternalEvent? _selectZoomEvent;
        private static Get2dViewsExternalEventHandler? _get2dViewsHandler;
        private static ExternalEvent? _get2dViewsEvent;
        private static Open2dViewExternalEventHandler? _open2dViewHandler;
        private static ExternalEvent? _open2dViewEvent;
        private static ParameterEditorExternalEventHandler? _paramEditorHandler;
        private static ExternalEvent? _paramEditorEvent;
        private static string? _cachedJson;
        private static EventHandler<IdlingEventArgs>? _idlingHandler;
        private static readonly TimeSpan SelectionSyncInterval = TimeSpan.FromMilliseconds(350);
        private static DateTime _lastSelectionSyncUtc = DateTime.MinValue;
        private static IReadOnlyList<int> _lastSelectionIds = Array.Empty<int>();
        private static bool _autoSyncEnabled = true;
        private static bool _selectionLocked;
        private static bool _forceSelectionSync;

        public static void Initialize()
        {
            _selectZoomHandler = new SelectZoomExternalEventHandler();
            _selectZoomEvent = ExternalEvent.Create(_selectZoomHandler);

            _get2dViewsHandler = new Get2dViewsExternalEventHandler();
            _get2dViewsEvent = ExternalEvent.Create(_get2dViewsHandler);

            _open2dViewHandler = new Open2dViewExternalEventHandler();
            _open2dViewEvent = ExternalEvent.Create(_open2dViewHandler);

            _paramEditorHandler = new ParameterEditorExternalEventHandler
            {
                PostResult = PostToWebView
            };
            _paramEditorEvent = ExternalEvent.Create(_paramEditorHandler);
        }

        public static void RegisterPane(AuditWebPane pane)
        {
            _paneInstance = pane;
            _forceSelectionSync = true;

            var cachedJson = _cachedJson;
            if (!string.IsNullOrEmpty(cachedJson))
            {
                _paneInstance.PostAuditResults(cachedJson!);
            }
        }

        public static void StartSelectionAutoSync(UIControlledApplication application)
        {
            if (_idlingHandler != null) return;
            _idlingHandler = OnIdling;
            application.Idling += _idlingHandler;
        }

        public static void StopSelectionAutoSync(UIControlledApplication application)
        {
            if (_idlingHandler == null) return;
            application.Idling -= _idlingHandler;
            _idlingHandler = null;
        }

        public static void SetSelectionSyncState(bool autoSyncEnabled, bool selectionLocked)
        {
            _autoSyncEnabled = autoSyncEnabled;
            _selectionLocked = selectionLocked;

            if (_autoSyncEnabled && !_selectionLocked)
            {
                _forceSelectionSync = true;
            }
        }

        private static void PostToWebView(string json)
        {
            _paneInstance?.PostAuditResults(json);
        }

        public static void RequestSelectZoom(int elementId)
        {
            if (_selectZoomHandler == null || _selectZoomEvent == null) return;

            _selectZoomHandler.PendingElementId = elementId;
            _selectZoomEvent.Raise();
        }

        public static void Request2dViews(int elementId)
        {
            if (_get2dViewsHandler == null || _get2dViewsEvent == null) return;

            _get2dViewsHandler.PendingElementId = elementId;
            _get2dViewsEvent.Raise();
        }

        public static void RequestOpen2dView(int viewId)
        {
            if (_open2dViewHandler == null || _open2dViewEvent == null) return;

            _open2dViewHandler.PendingViewId = viewId;
            _open2dViewEvent.Raise();
        }

        public static void RequestParameterEditorOperation(ParameterEditorExternalEventHandler.OperationType operation, string? jsonPayload = null)
        {
            if (_paramEditorHandler == null || _paramEditorEvent == null) return;

            _paramEditorHandler.CurrentOperation = operation;
            _paramEditorHandler.JsonPayload = jsonPayload;
            _paramEditorEvent.Raise();
        }

        public static void RequestAuditRefresh()
        {
            RequestParameterEditorOperation(ParameterEditorExternalEventHandler.OperationType.RefreshAudit);
        }

        internal static void Post2dViewOptions(int elementId, IReadOnlyList<ViewOption> views)
        {
            if (_paneInstance == null) return;

            var dto = new TwoDViewOptionsDto
            {
                type = "2dViewOptions",
                elementId = elementId,
                views = views.Select(v => new ViewDto
                {
                    viewId = v.ViewId,
                    name = v.Name,
                    viewType = v.ViewType
                }).ToList()
            };

            var json = SerializeToJson(dto);
            _paneInstance.PostAuditResults(json);
        }

        public static void UpdateFromReport(AuditService.AuditReport report)
        {
            var dto = new AuditResultsDto
            {
                type = "auditResults",
                summary = new SummaryDto
                {
                    overallReadinessPercent = Math.Round(report.AverageReadinessScore * 100, 0),
                    fullyReady = report.FullyReadyAssets,
                    totalAudited = report.TotalAuditedAssets,
                    auditProfile = report.AuditProfileName,
                    groupScores = report.AverageGroupScores.ToDictionary(
                        kvp => kvp.Key,
                        kvp => (int)Math.Round(kvp.Value * 100))
                },
                rows = new List<RowDto>()
            };

            foreach (var result in report.ElementResults)
            {
                dto.rows.Add(new RowDto
                {
                    elementId = result.ElementId,
                    category = result.Category,
                    family = result.Family,
                    type = result.Type,
                    missingCount = result.MissingCount,
                    readinessPercent = (int)Math.Round(result.ReadinessScore * 100),
                    missingParams = result.MissingParams,
                    missingFields = (result.MissingFields ?? new List<MissingFieldInfo>())
                        .Select(f => new MissingFieldDto
                        {
                            key = f.FieldKey,
                            label = f.FieldLabel,
                            group = f.Group,
                            scope = f.Scope,
                            required = f.Required,
                            reason = f.Reason ?? string.Empty
                        })
                        .Where(f => !string.IsNullOrWhiteSpace(f.key) || !string.IsNullOrWhiteSpace(f.label))
                        .GroupBy(f => string.IsNullOrWhiteSpace(f.key)
                            ? $"{f.label}|{f.reason}"
                            : $"{f.key}|{f.reason}")
                        .Select(g => g.First())
                        .ToList(),
                    groupScores = result.GroupScores.ToDictionary(
                        kvp => kvp.Key,
                        kvp => (int)Math.Round(kvp.Value * 100))
                });
            }

            var json = SerializeToJson(dto);

            _cachedJson = json;

            _paneInstance?.PostAuditResults(json);
        }

        internal class ViewOption
        {
            public int ViewId { get; }
            public string Name { get; }
            public string ViewType { get; }

            public ViewOption(int viewId, string name, string viewType)
            {
                ViewId = viewId;
                Name = name;
                ViewType = viewType;
            }
        }

        private static string SerializeToJson<T>(T value)
        {
            var settings = new DataContractJsonSerializerSettings
            {
                UseSimpleDictionaryFormat = true
            };
            var serializer = new DataContractJsonSerializer(typeof(T), settings);
            using var stream = new MemoryStream();
            serializer.WriteObject(stream, value);
            return Encoding.UTF8.GetString(stream.ToArray());
        }

        private static void OnIdling(object sender, IdlingEventArgs e)
        {
            if (!_autoSyncEnabled || _selectionLocked) return;
            if (_paneInstance == null || _paramEditorHandler == null) return;

            var now = DateTime.UtcNow;
            if (!_forceSelectionSync && now - _lastSelectionSyncUtc < SelectionSyncInterval) return;
            _lastSelectionSyncUtc = now;

            var uiApp = sender as UIApplication ?? TryGetUiApplication();
            if (uiApp == null) return;
            var uidoc = uiApp.ActiveUIDocument;
            if (uidoc == null) return;

            if (!_paramEditorHandler.TryBuildSelectedElementsSnapshot(uidoc, uidoc.Document, out var elementIds, out var json))
                return;

            if (!_forceSelectionSync && AreSameSelection(elementIds, _lastSelectionIds))
                return;

            _forceSelectionSync = false;
            _lastSelectionIds = elementIds.ToArray();
            _paneInstance.PostAuditResults(json);
        }

        private static bool AreSameSelection(IReadOnlyList<int> current, IReadOnlyList<int> previous)
        {
            if (ReferenceEquals(current, previous)) return true;
            if (current.Count != previous.Count) return false;
            for (int i = 0; i < current.Count; i++)
            {
                if (current[i] != previous[i]) return false;
            }

            return true;
        }

        private static UIApplication? TryGetUiApplication()
        {
            try
            {
                return Context.UiApplication;
            }
            catch
            {
                return null;
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

        [DataContract]
        private class AuditResultsDto
        {
            [DataMember(Name = "type")]
            public string type { get; set; } = "auditResults";
            [DataMember(Name = "summary")]
            public SummaryDto? summary { get; set; }
            [DataMember(Name = "rows")]
            public List<RowDto> rows { get; set; } = new();
        }

        [DataContract]
        private class SummaryDto
        {
            [DataMember(Name = "overallReadinessPercent")]
            public double overallReadinessPercent { get; set; }
            [DataMember(Name = "fullyReady")]
            public int fullyReady { get; set; }
            [DataMember(Name = "totalAudited")]
            public int totalAudited { get; set; }
            [DataMember(Name = "auditProfile")]
            public string auditProfile { get; set; } = string.Empty;
            [DataMember(Name = "groupScores")]
            public Dictionary<string, int> groupScores { get; set; } = new();
        }

        [DataContract]
        private class RowDto
        {
            [DataMember(Name = "elementId")]
            public int elementId { get; set; }
            [DataMember(Name = "category")]
            public string category { get; set; } = string.Empty;
            [DataMember(Name = "family")]
            public string family { get; set; } = string.Empty;
            [DataMember(Name = "type")]
            public string type { get; set; } = string.Empty;
            [DataMember(Name = "missingCount")]
            public int missingCount { get; set; }
            [DataMember(Name = "readinessPercent")]
            public int readinessPercent { get; set; }
            [DataMember(Name = "missingParams")]
            public string missingParams { get; set; } = string.Empty;
            [DataMember(Name = "missingFields")]
            public List<MissingFieldDto> missingFields { get; set; } = new();
            [DataMember(Name = "groupScores")]
            public Dictionary<string, int> groupScores { get; set; } = new();
        }

        [DataContract]
        private class MissingFieldDto
        {
            [DataMember(Name = "key")]
            public string key { get; set; } = string.Empty;
            [DataMember(Name = "label")]
            public string label { get; set; } = string.Empty;
            [DataMember(Name = "group")]
            public string group { get; set; } = string.Empty;
            [DataMember(Name = "scope")]
            public string scope { get; set; } = string.Empty;
            [DataMember(Name = "required")]
            public bool required { get; set; }
            [DataMember(Name = "reason")]
            public string reason { get; set; } = string.Empty;
        }

        [DataContract]
        private class TwoDViewOptionsDto
        {
            [DataMember(Name = "type")]
            public string type { get; set; } = "2dViewOptions";
            [DataMember(Name = "elementId")]
            public int elementId { get; set; }
            [DataMember(Name = "views")]
            public List<ViewDto> views { get; set; } = new();
        }

        [DataContract]
        private class ViewDto
        {
            [DataMember(Name = "viewId")]
            public int viewId { get; set; }
            [DataMember(Name = "name")]
            public string name { get; set; } = string.Empty;
            [DataMember(Name = "viewType")]
            public string viewType { get; set; } = string.Empty;
        }

        private class Get2dViewsExternalEventHandler : IExternalEventHandler
        {
            public int? PendingElementId { get; set; }

            public void Execute(UIApplication app)
            {
                if (PendingElementId == null) return;
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null) return;
                var doc = uidoc.Document;

                var elementId = new ElementId(PendingElementId.Value);
                var element = doc.GetElement(elementId);
                PendingElementId = null;
                if (element == null) return;

                var levelId = GetElementLevel(element, doc);
                if (levelId == null || levelId == ElementId.InvalidElementId)
                {
                    var allPlanViews = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewPlan))
                        .Cast<ViewPlan>()
                        .Where(v => !v.IsTemplate)
                        .Where(v => v.ViewType == ViewType.FloorPlan || v.ViewType == ViewType.CeilingPlan || v.ViewType == ViewType.EngineeringPlan)
                        .OrderBy(v => v.GenLevel?.Elevation ?? 0)
                        .ThenBy(v => v.Name)
                        .Select(v => new ViewOption(GetElementIdValue(v.Id), v.Name, v.ViewType.ToString()))
                        .ToList();

                    Post2dViewOptions(GetElementIdValue(elementId), allPlanViews);
                    return;
                }

                var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewPlan))
                    .Cast<ViewPlan>()
                    .Where(v => !v.IsTemplate)
                    .Where(v => v.GenLevel != null && v.GenLevel.Id == levelId)
                    .Where(v => v.ViewType == ViewType.FloorPlan || v.ViewType == ViewType.CeilingPlan || v.ViewType == ViewType.EngineeringPlan)
                    .Select(v => new ViewOption(GetElementIdValue(v.Id), v.Name, v.ViewType.ToString()))
                    .ToList();

                Post2dViewOptions(GetElementIdValue(elementId), views);
            }

            private ElementId? GetElementLevel(Element element, Document doc)
            {
                var levelId = element.LevelId;
                if (levelId != null && levelId != ElementId.InvalidElementId)
                    return levelId;

                var levelParam = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM);
                if (levelParam != null && levelParam.HasValue)
                {
                    var paramLevelId = levelParam.AsElementId();
                    if (paramLevelId != null && paramLevelId != ElementId.InvalidElementId)
                        return paramLevelId;
                }

                levelParam = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM);
                if (levelParam != null && levelParam.HasValue)
                {
                    var paramLevelId = levelParam.AsElementId();
                    if (paramLevelId != null && paramLevelId != ElementId.InvalidElementId)
                        return paramLevelId;
                }

                levelParam = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM);
                if (levelParam != null && levelParam.HasValue)
                {
                    var paramLevelId = levelParam.AsElementId();
                    if (paramLevelId != null && paramLevelId != ElementId.InvalidElementId)
                        return paramLevelId;
                }

                levelParam = element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM);
                if (levelParam != null && levelParam.HasValue)
                {
                    var paramLevelId = levelParam.AsElementId();
                    if (paramLevelId != null && paramLevelId != ElementId.InvalidElementId)
                        return paramLevelId;
                }

                if (element is FamilyInstance fi && fi.Host != null)
                {
                    var hostLevelId = GetElementLevel(fi.Host, doc);
                    if (hostLevelId != null && hostLevelId != ElementId.InvalidElementId)
                        return hostLevelId;
                }

                var point = GetElementPoint(element);
                if (point != null)
                {
                    var levels = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .OrderBy(l => l.Elevation)
                        .ToList();

                    Level? bestLevel = null;
                    foreach (var level in levels)
                    {
                        if (level.Elevation <= point.Z + 0.1)
                            bestLevel = level;
                        else
                            break;
                    }

                    if (bestLevel != null)
                        return bestLevel.Id;

                    if (levels.Count > 0)
                        return levels[0].Id;
                }

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

            public string GetName() => "FMReadiness_v3.Get2dViewsHandler";
        }

        private class Open2dViewExternalEventHandler : IExternalEventHandler
        {
            public int? PendingViewId { get; set; }

            public void Execute(UIApplication app)
            {
                if (PendingViewId == null) return;
                var uidoc = app.ActiveUIDocument;
                if (uidoc == null) return;

                var viewId = new ElementId(PendingViewId.Value);
                PendingViewId = null;

                var view = uidoc.Document.GetElement(viewId) as View;
                if (view == null) return;
                if (view.IsTemplate) return;

                uidoc.RequestViewChange(view);
            }

            public string GetName() => "FMReadiness_v3.Open2dViewHandler";
        }
    }
}

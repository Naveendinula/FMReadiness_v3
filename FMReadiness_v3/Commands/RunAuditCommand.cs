using System;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using FMReadiness_v3.Services;
using FMReadiness_v3.UI.Panes;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class RunAuditCommand : ExternalCommand
    {
        public override void Execute()
        {
            var doc = Context.ActiveDocument ?? Context.Document;
            var uiApp = Context.UiApplication;

            if (doc == null)
            {
                TaskDialog.Show("FM Readiness", "No active document. Open a model and try again.");
                return;
            }

            if (uiApp == null)
            {
                TaskDialog.Show("FM Readiness", "Revit UI application context is unavailable.");
                return;
            }

            try
            {
                var checklist = new ChecklistService();
                if (!checklist.LoadConfig()) return;

                var collector = new CollectorService(doc);
                var elements = collector.GetAllFmElements();

                var auditService = new AuditService();
                var report = auditService.RunFullAudit(doc, elements, checklist.Rules);

                WebViewPaneController.UpdateFromReport(report);

                var pane = uiApp.GetDockablePane(PaneIds.FMReadinessPaneId);
                if (pane != null && !pane.IsShown())
                {
                    pane.Show();
                }

                var total = report.TotalAuditedAssets;
                var ready = report.FullyReadyAssets;
                var readinessPct = report.AverageReadinessScore * 100.0;

                var topMissingInstance = report.MissingParamCounts
                    .OrderByDescending(x => x.Value)
                    .Take(5)
                    .Select(x => $"- {x.Key} ({x.Value})");

                var topMissingType = report.MissingTypeParamCounts
                    .OrderByDescending(x => x.Value)
                    .Take(5)
                    .Select(x => $"- {x.Key} ({x.Value})");

                var message =
                    "FM Readiness Summary\n" +
                    $"Overall readiness: {Math.Round(readinessPct, 0)}%\n" +
                    $"Fully ready assets: {ready} / {total}\n\n" +
                    "-----------------------------------\n\n" +
                    "Missing INSTANCE params found\n" +
                    $"Elements with missing instance data: {report.ElementsWithMissingData}\n\n" +
                    $"Top missing instance params:\n{string.Join("\n", topMissingInstance)}\n\n" +
                    "-----------------------------------\n\n" +
                    "Missing TYPE params found\n" +
                    $"Elements with missing type data: {report.ElementsWithMissingTypeData}\n\n" +
                    $"Top missing type params:\n{string.Join("\n", topMissingType)}\n\n" +
                    "-----------------------------------\n" +
                    "Results shown in FM Readiness pane.";

                TaskDialog.Show("FM Data Readiness Audit", message);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FM Readiness - Error", ex.ToString());
            }
        }
    }
}

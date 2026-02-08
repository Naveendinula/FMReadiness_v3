using System;
using System.Linq;
using System.Collections.Generic;
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
                var resolver = new AuditProfileResolverService();
                if (!resolver.TryResolveRules(out var rules, out var profileName, out var errorMessage))
                {
                    TaskDialog.Show("FM Readiness", errorMessage);
                    return;
                }

                var collector = new CollectorService(doc);
                var elements = collector.GetAllFmElements();
                var scoreMode = AuditProfileState.GetScoreMode();

                var auditService = new AuditService();
                var report = auditService.RunFullAudit(doc, elements, rules, scoreMode);
                report.AuditProfileName = profileName;

                WebViewPaneController.UpdateFromReport(report);

                var pane = uiApp.GetDockablePane(PaneIds.FMReadinessPaneId);
                if (pane != null && !pane.IsShown())
                {
                    pane.Show();
                }

                var total = report.TotalAuditedAssets;
                var ready = report.FullyReadyAssets;
                var readinessPct = report.AverageReadinessScore * 100.0;

                var topMissingComponent = report.MissingParamCounts
                    .OrderByDescending(x => x.Value)
                    .ThenBy(x => x.Key)
                    .Take(5)
                    .ToList();

                var topMissingType = report.MissingTypeParamCounts
                    .OrderByDescending(x => x.Value)
                    .ThenBy(x => x.Key)
                    .Take(5)
                    .ToList();

                var dialog = new TaskDialog("FM Data Readiness Audit")
                {
                    MainInstruction = $"FM audit complete — {Math.Round(readinessPct, 0)}% readiness",
                    MainContent =
                        $"Audit profile: {profileName}\n" +
                        $"Audit scope: {AuditProfileState.GetScoreModeLabel(scoreMode)}\n" +
                        $"Fully ready assets: {ready} / {total}\n\n" +
                        $"Assets with missing component data: {report.ElementsWithMissingData}\n" +
                        $"Assets with missing type data: {report.ElementsWithMissingTypeData}",
                    ExpandedContent =
                        "Top missing component fields:\n" +
                        FormatTopMissing(topMissingComponent) +
                        "\n\nTop missing type fields:\n" +
                        FormatTopMissing(topMissingType),
                    FooterText = "Detailed results are shown in the FM Readiness pane."
                };

                dialog.Show();
            }
            catch (Exception ex)
            {
                TaskDialog.Show("FM Readiness - Error", ex.ToString());
            }
        }

        private static string FormatTopMissing(IEnumerable<KeyValuePair<string, int>> missingCounts)
        {
            var items = missingCounts
                .Where(x => x.Value > 0)
                .Take(5)
                .ToList();

            if (items.Count == 0)
            {
                return "None";
            }

            return string.Join("\n", items.Select((x, index) => $"{index + 1}. {x.Key} ({x.Value})"));
        }
    }
}

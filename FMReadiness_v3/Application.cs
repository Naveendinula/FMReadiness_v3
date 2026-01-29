using FMReadiness_v3.Commands;
using FMReadiness_v3.UI.Panes;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        private static AuditResultsPaneProvider? _paneProvider;

        public override void OnStartup()
        {
            WebViewPaneController.Initialize();
            RegisterDockablePane();
            CreateRibbon();
        }

        private void RegisterDockablePane()
        {
            _paneProvider = new AuditResultsPaneProvider();
            base.Application.RegisterDockablePane(
                PaneIds.FMReadinessPaneId,
                "FM Readiness",
                _paneProvider);
        }

        private void CreateRibbon()
        {
            var panel = base.Application.CreatePanel("FM Tools", "Digital Twin");

            panel.AddPushButton<RunAuditCommand>("Run FM\nAudit")
                .SetImage("/FMReadiness_v3;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/FMReadiness_v3;component/Resources/Icons/RibbonIcon32.png")
                .SetToolTip("Checks FM data completeness and shows results in the FM Readiness pane.");

            panel.AddPushButton<ShowPaneCommand>("FM Pane")
                .SetToolTip("Show or hide the FM Readiness results pane.");

            panel.AddPushButton<ExportFmSidecarCommand>("Export FM\nSidecar")
                .SetToolTip("Exports FM parameters as a sidecar JSON file for the DigitalTwin viewer.\n\nThe sidecar file contains FM parameters keyed by IFC GlobalId.\nUpload it alongside your IFC file to show FM data in the viewer.");
        }
    }
}

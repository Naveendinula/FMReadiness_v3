using Autodesk.Revit.UI;

namespace FMReadiness_v3.UI.Panes
{
    public class AuditResultsPaneProvider : IDockablePaneProvider
    {
        public AuditWebPane? WebPaneInstance { get; private set; }

        public void SetupDockablePane(DockablePaneProviderData data)
        {
            WebPaneInstance = new AuditWebPane();
            data.FrameworkElement = WebPaneInstance;
            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Right
            };
            data.VisibleByDefault = true;
        }
    }
}


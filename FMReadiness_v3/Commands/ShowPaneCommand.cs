using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using FMReadiness_v3.UI.Panes;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class ShowPaneCommand : ExternalCommand
    {
        public override void Execute()
        {
            var uiApp = Context.UiApplication;
            var pane = uiApp.GetDockablePane(PaneIds.FMReadinessPaneId);

            if (pane == null)
            {
                TaskDialog.Show("FM Readiness", "Could not find FM Readiness pane.");
                return;
            }

            if (pane.IsShown())
            {
                pane.Hide();
            }
            else
            {
                pane.Show();
            }
        }
    }
}


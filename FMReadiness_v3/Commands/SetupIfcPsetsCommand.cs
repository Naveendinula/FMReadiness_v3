using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using FMReadiness_v3.IFC;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.Commands
{
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class SetupIfcPsetsCommand : ExternalCommand
    {
        public override void Execute()
        {
            var uiApp = Context.UiApplication;
            if (uiApp == null)
            {
                TaskDialog.Show("FM Readiness", "Revit UI application context is unavailable.");
                return;
            }

            var psetPath = IfcExportHelper.EnsureUserDefinedPsetFile();
            if (string.IsNullOrWhiteSpace(psetPath))
            {
                TaskDialog.Show("FM Readiness", "Unable to create the IFC property set configuration file.");
                return;
            }

            var dialog = new TaskDialog("FM Readiness - IFC Property Sets")
            {
                MainInstruction = "User-defined FM property sets file is ready.",
                MainContent =
                    "To export these parameters to IFC:\n" +
                    "1. Open IFC Export and click Modify Setup.\n" +
                    "2. Go to Property Sets tab.\n" +
                    "3. Check 'Export user defined property sets'.\n" +
                    "4. Browse to the file below.\n\n" +
                    $"File:\n{psetPath}"
            };

            dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open IFC Export dialog");
            dialog.CommonButtons = TaskDialogCommonButtons.Close;
            var result = dialog.Show();

            if (result == TaskDialogResult.CommandLink1)
            {
                var cmdId = TryGetIfcExportCommandId();
                if (cmdId != null)
                {
                    uiApp.PostCommand(cmdId);
                }
                else
                {
                    TaskDialog.Show(
                        "FM Readiness",
                        "Could not find the IFC Export command in this Revit version. " +
                        "Open IFC Export manually from File > Export > IFC.");
                }
            }
        }

        private static RevitCommandId? TryGetIfcExportCommandId()
        {
            var candidates = new[] { "IFCExport", "ExportIFC", "ExportIfc" };
            foreach (var name in candidates)
            {
                if (Enum.TryParse(name, out PostableCommand command))
                {
                    var cmdId = RevitCommandId.LookupPostableCommandId(command);
                    if (cmdId != null)
                    {
                        return cmdId;
                    }
                }
            }

            return null;
        }

    }
}

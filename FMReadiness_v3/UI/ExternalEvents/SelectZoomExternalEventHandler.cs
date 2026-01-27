using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace FMReadiness_v3.UI.ExternalEvents
{
    public class SelectZoomExternalEventHandler : IExternalEventHandler
    {
        public int? PendingElementId { get; set; }

        public void Execute(UIApplication app)
        {
            if (PendingElementId == null) return;

            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) return;

            var elementId = new ElementId(PendingElementId.Value);
            var element = uidoc.Document.GetElement(elementId);
            if (element == null)
            {
                TaskDialog.Show("FM Readiness", $"Element {PendingElementId} not found in the document.");
                PendingElementId = null;
                return;
            }

            uidoc.Selection.SetElementIds(new List<ElementId> { elementId });
            uidoc.ShowElements(elementId);

            PendingElementId = null;
        }

        public string GetName() => "FMReadiness_v3.SelectZoomHandler";
    }
}


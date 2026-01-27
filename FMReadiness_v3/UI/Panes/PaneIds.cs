using System;
using Autodesk.Revit.UI;

namespace FMReadiness_v3.UI.Panes
{
    public static class PaneIds
    {
        public static readonly Guid FMReadinessPaneGuid = new Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890");
        public static readonly DockablePaneId FMReadinessPaneId = new DockablePaneId(FMReadinessPaneGuid);
    }
}


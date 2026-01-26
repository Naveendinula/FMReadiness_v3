using FMReadiness_v3.Commands;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3
{
    /// <summary>
    ///     Application entry point
    /// </summary>
    [UsedImplicitly]
    public class Application : ExternalApplication
    {
        public override void OnStartup()
        {
            CreateRibbon();
        }

        private void CreateRibbon()
        {
            var panel = Application.CreatePanel("Commands", "FMReadiness_v3");

            panel.AddPushButton<StartupCommand>("Execute")
                .SetImage("/FMReadiness_v3;component/Resources/Icons/RibbonIcon16.png")
                .SetLargeImage("/FMReadiness_v3;component/Resources/Icons/RibbonIcon32.png");
        }
    }
} knkjnjnkjn
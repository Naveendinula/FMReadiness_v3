using Autodesk.Revit.Attributes;
using FMReadiness_v3.ViewModels;
using FMReadiness_v3.Views;
using Nice3point.Revit.Toolkit.External;

namespace FMReadiness_v3.Commands
{
    /// <summary>
    ///     External command entry point.
    /// </summary>
    [UsedImplicitly]
    [Transaction(TransactionMode.Manual)]
    public class StartupCommand : ExternalCommand
    {
        public override void Execute()
        {
            var viewModel = new FMReadiness_v3ViewModel();
            var view = new FMReadiness_v3View(viewModel);
            view.ShowDialog();
        }
    }
}
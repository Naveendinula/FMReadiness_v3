using FMReadiness_v3.ViewModels;

namespace FMReadiness_v3.Views
{
    public sealed partial class FMReadiness_v3View
    {
        public FMReadiness_v3View(FMReadiness_v3ViewModel viewModel)
        {
            DataContext = viewModel;
            InitializeComponent();
        }
    }
}
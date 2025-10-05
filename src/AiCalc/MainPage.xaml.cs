using AiCalc.ViewModels;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;

namespace AiCalc;

public sealed partial class MainPage : Page
{
    public MainPage()
    {
        ViewModel = new WorkbookViewModel();
        InitializeComponent();
        DataContext = ViewModel;
    }

    public WorkbookViewModel ViewModel { get; }
}

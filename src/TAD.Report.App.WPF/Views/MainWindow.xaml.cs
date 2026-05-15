using System.Windows;
using TAD.Report.App.WPF.ViewModels;

namespace TAD.Report.App.WPF.Views;

public partial class MainWindow : Window
{
    public MainWindow(MainViewModel viewModel)
    {
        InitializeComponent();
        DataContext = viewModel;
    }
}

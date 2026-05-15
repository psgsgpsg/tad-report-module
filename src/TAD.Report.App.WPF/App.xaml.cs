using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using TAD.Report.App.WPF.ViewModels;
using TAD.Report.App.WPF.Views;
using TAD.Report.Core.Interfaces;
using TAD.Report.Infrastructure.PowerPoint.Services;

namespace TAD.Report.App.WPF;

public partial class App : Application
{
    private ServiceProvider? _serviceProvider;

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        var services = new ServiceCollection();

        services.AddSingleton<IReportGenerator, PowerPointReportGenerator>();
        services.AddSingleton<MainViewModel>();
        services.AddSingleton<MainWindow>();

        _serviceProvider = services.BuildServiceProvider();

        var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
        MainWindow = mainWindow;
        mainWindow.Show();
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _serviceProvider?.Dispose();
        base.OnExit(e);
    }
}

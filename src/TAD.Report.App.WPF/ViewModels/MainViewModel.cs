using System.IO;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Win32;
using TAD.Report.Core.Interfaces;
using TAD.Report.Core.Models;

namespace TAD.Report.App.WPF.ViewModels;

/// <summary>
/// Main screen VM: binds sample <see cref="ReportData"/> and exports via <see cref="IReportGenerator"/>.
/// </summary>
public sealed partial class MainViewModel : ObservableObject
{
    private readonly IReportGenerator _reportGenerator;
    private readonly IAsyncRelayCommand _exportReportAsyncCommand;

    public MainViewModel(IReportGenerator reportGenerator)
    {
        _reportGenerator = reportGenerator ?? throw new ArgumentNullException(nameof(reportGenerator));
        ReportModel = CreateDesignTimeReportSample();
        _exportReportAsyncCommand = new AsyncRelayCommand(ExportReportAsyncCore);
    }

    /// <summary>Explicit command so binding works even if source generators are skipped in the IDE/build pipeline.</summary>
    public IAsyncRelayCommand ExportReportAsyncCommand => _exportReportAsyncCommand;

    /// <summary>Aggregate bound to the view (title, metrics, test case rows, trends).</summary>
    [ObservableProperty]
    private ReportData _reportModel = new();

    /// <summary>True while export / IO is in progress.</summary>
    [ObservableProperty]
    private bool _isBusy;

    /// <summary>Last export or validation error for safe UI binding (cleared on success).</summary>
    [ObservableProperty]
    private string? _errorMessage;

    private async Task ExportReportAsyncCore()
    {
        ErrorMessage = null;

        var dialog = new SaveFileDialog
        {
            Title = "PPTX 저장 위치 선택",
            Filter = "PowerPoint presentation (*.pptx)|*.pptx",
            DefaultExt = ".pptx",
            AddExtension = true,
            FileName = "TAD_TestReport.pptx",
        };

        var owner = Application.Current?.MainWindow;
        var accepted = owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
        if (accepted != true || string.IsNullOrWhiteSpace(dialog.FileName))
            return;

        try
        {
            IsBusy = true;

            var result = await _reportGenerator.GenerateReportAsync(ReportModel).ConfigureAwait(true);

            if (!result.IsSuccess)
            {
                ErrorMessage = string.IsNullOrWhiteSpace(result.ErrorMessage)
                    ? "Report generation failed."
                    : result.ErrorMessage;
                return;
            }

            if (result.Value is not { Length: > 0 } bytes)
            {
                ErrorMessage = "Report generation returned no data.";
                return;
            }

            await File.WriteAllBytesAsync(dialog.FileName, bytes).ConfigureAwait(true);
            ErrorMessage = null;

            MessageBox.Show(
                $"파일이 저장되었습니다.\n{dialog.FileName}",
                "PPT 보고서 발행",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            ErrorMessage = ex.Message;
            MessageBox.Show(
                ex.Message,
                "PPT 보고서 발행 오류",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        finally
        {
            IsBusy = false;
        }
    }

    private static ReportData CreateDesignTimeReportSample()
    {
        var report = new ReportData
        {
            Title = "TAD Test Report",
            Date = DateTime.Now.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture),
            TestCases =
            [
                new TestCaseResult
                {
                    No = 1,
                    Name = "Login flow",
                    Result = "PASS",
                    Description = string.Empty,
                    Remarks = "OK",
                },
                new TestCaseResult
                {
                    No = 2,
                    Name = "Export report",
                    Result = "FAIL",
                    Description = "Timeout waiting for dialog.",
                    Screenshot = LoadSampleScreenshot(),
                    Remarks = "Retry later",
                },
            ],
            Trends =
            [
                new DailyTrend { Date = "05-12", PassCount = 12, FailCount = 1 },
                new DailyTrend { Date = "05-13", PassCount = 10, FailCount = 2 },
            ],
        };

        return report;
    }

    private static byte[] LoadSampleScreenshot()
    {
        var screenshotPath = Path.Combine(AppContext.BaseDirectory, "assets", "export_report_screenshot.png");
        return File.Exists(screenshotPath) ? File.ReadAllBytes(screenshotPath) : [];
    }
}

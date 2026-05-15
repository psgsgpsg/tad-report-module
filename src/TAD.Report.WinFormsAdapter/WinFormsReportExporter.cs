using System.Windows.Forms;
using TAD.Report.Core.Models;

namespace TAD.Report.WinFormsAdapter;

/// <summary>
/// WinForms-friendly helper that owns the save dialog and user-facing success/error messages.
/// </summary>
public static class WinFormsReportExporter
{
    public static async Task<bool> ExportWithSaveDialogAsync(
        IWin32Window? owner,
        ReportData data,
        ReportExportOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        options ??= new ReportExportOptions();

        using var dialog = new SaveFileDialog
        {
            Title = options.SaveDialogTitle,
            Filter = "PowerPoint presentation (*.pptx)|*.pptx",
            DefaultExt = ".pptx",
            AddExtension = true,
            FileName = options.DefaultFileName,
        };

        var accepted = owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
        if (accepted != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.FileName))
            return false;

        var service = new ReportExportService(options);
        var result = await service.ExportAsync(data, dialog.FileName, cancellationToken).ConfigureAwait(true);
        if (!result.IsSuccess)
        {
            MessageBox.Show(
                owner,
                result.ErrorMessage ?? "Report export failed.",
                "PPT 보고서 발행 오류",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return false;
        }

        MessageBox.Show(
            owner,
            $"파일이 저장되었습니다.{Environment.NewLine}{result.Value}",
            "PPT 보고서 발행",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
        return true;
    }
}

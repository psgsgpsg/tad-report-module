using TAD.Report.Core;
using TAD.Report.Core.Interfaces;
using TAD.Report.Core.Models;
using TAD.Report.Infrastructure.PowerPoint.Services;

namespace TAD.Report.WinFormsAdapter;

/// <summary>
/// UI-independent export facade intended for WinForms projects.
/// </summary>
public sealed class ReportExportService
{
    private readonly IReportGenerator _reportGenerator;

    public ReportExportService(ReportExportOptions options)
        : this(new PowerPointReportGenerator(
            (options ?? throw new ArgumentNullException(nameof(options))).TemplatePath,
            options.CompanyLogoPath))
    {
    }

    public ReportExportService(IReportGenerator reportGenerator)
    {
        _reportGenerator = reportGenerator ?? throw new ArgumentNullException(nameof(reportGenerator));
    }

    /// <summary>
    /// Generates a PPTX report and writes it to <paramref name="outputPath"/>.
    /// </summary>
    public async Task<Result<string>> ExportAsync(
        ReportData data,
        string outputPath,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(outputPath))
            return Result.Failure<string>("Output path is required.");

        var generated = await _reportGenerator.GenerateReportAsync(data, cancellationToken).ConfigureAwait(false);
        if (!generated.IsSuccess)
            return Result.Failure<string>(generated.ErrorMessage ?? "Report generation failed.");

        if (generated.Value is not { Length: > 0 } bytes)
            return Result.Failure<string>("Report generation returned no data.");

        var directory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(directory))
            Directory.CreateDirectory(directory);

        await File.WriteAllBytesAsync(outputPath, bytes, cancellationToken).ConfigureAwait(false);
        return Result.Success(outputPath);
    }
}

using TAD.Report.Core.Models;

namespace TAD.Report.Core.Interfaces;

/// <summary>
/// Builds a report document from <see cref="ReportData"/> (PPTX or other binary payload).
/// </summary>
public interface IReportGenerator
{
    /// <summary>
    /// Generates a report file as a byte array from the supplied aggregate.
    /// </summary>
    /// <param name="data">Source metrics and collections for placeholder binding.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Success with file bytes, or failure with an error message.</returns>
    Task<Result<byte[]>> GenerateReportAsync(ReportData data, CancellationToken cancellationToken = default);
}

namespace TAD.Report.Core.Models;

/// <summary>
/// Aggregated pass/fail counts for a calendar day (date format: MM-dd).
/// </summary>
public sealed class DailyTrend
{
    /// <summary>Day label in MM-dd format.</summary>
    public string Date { get; set; } = string.Empty;

    /// <summary>Total passed tests on this day.</summary>
    public int PassCount { get; set; }

    /// <summary>Total failed tests on this day.</summary>
    public int FailCount { get; set; }
}

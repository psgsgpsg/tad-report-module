namespace TAD.Report.Core.Models;

/// <summary>
/// Root aggregate for PPTX report binding (placeholders: {{TITLE}}, {{DATE}}, {{TOTAL}}, {{PASS}}, {{FAIL}}, {{RATE}}).
/// </summary>
public sealed class ReportData
{
    /// <summary>Report main title (maps to {{TITLE}}).</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Execution date (maps to {{DATE}}).</summary>
    public string Date { get; set; } = string.Empty;

    /// <summary>Top-right company logo image bytes.</summary>
    public byte[] CompanyLogo { get; set; } = [];

    /// <summary>Collection of per-test results.</summary>
    public List<TestCaseResult> TestCases { get; set; } = [];

    /// <summary>Daily trend series for charts.</summary>
    public List<DailyTrend> Trends { get; set; } = [];

    /// <summary>Total number of test cases (maps to {{TOTAL}}).</summary>
    public int Total => TestCases.Count;

    /// <summary>Count of tests with Result == "PASS" (maps to {{PASS}}), ignoring case.</summary>
    public int Pass => TestCases.Count(static t => IsResult(t, "PASS"));

    /// <summary>Count of tests with Result == "FAIL" (maps to {{FAIL}}), ignoring case.</summary>
    public int Fail => TestCases.Count(static t => IsResult(t, "FAIL"));

    /// <summary>Pass rate in percent (maps to {{RATE}}%); zero when there are no tests.</summary>
    public double Rate => Total > 0 ? (double)Pass / Total * 100.0 : 0.0;

    private static bool IsResult(TestCaseResult testCase, string expected) =>
        string.Equals(testCase.Result, expected, StringComparison.OrdinalIgnoreCase);
}

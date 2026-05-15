namespace TAD.Report.Core.Models;

/// <summary>
/// Represents a single test case row for report generation (placeholder mapping: {{NAME}}, {{DESC}}).
/// </summary>
public sealed class TestCaseResult
{
    /// <summary>Sequential number.</summary>
    public int No { get; set; }

    /// <summary>Test name (maps to {{NAME}}).</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Outcome label: "PASS" or "FAIL".</summary>
    public string Result { get; set; } = string.Empty;

    /// <summary>Error logs or details (maps to {{DESC}}).</summary>
    public string Description { get; set; } = string.Empty;

    /// <summary>Failure screenshot payload; may be empty when not applicable.</summary>
    public byte[] Screenshot { get; set; } = [];

    /// <summary>Additional comments.</summary>
    public string Remarks { get; set; } = string.Empty;
}

namespace TAD.Report.WinFormsAdapter;

/// <summary>
/// Runtime paths and save-dialog defaults for WinForms report export integration.
/// </summary>
public sealed class ReportExportOptions
{
    /// <summary>PowerPoint template path. Defaults to assets/tad_report_template.pptx under AppContext.BaseDirectory.</summary>
    public string TemplatePath { get; set; } = Path.Combine(AppContext.BaseDirectory, "assets", "tad_report_template.pptx");

    /// <summary>Fallback company logo path. Defaults to assets/company_logo.png under AppContext.BaseDirectory.</summary>
    public string CompanyLogoPath { get; set; } = Path.Combine(AppContext.BaseDirectory, "assets", "company_logo.png");

    /// <summary>Default file name shown by the WinForms save dialog.</summary>
    public string DefaultFileName { get; set; } = "TAD_TestReport.pptx";

    /// <summary>Save dialog title.</summary>
    public string SaveDialogTitle { get; set; } = "PPTX 저장 위치 선택";
}

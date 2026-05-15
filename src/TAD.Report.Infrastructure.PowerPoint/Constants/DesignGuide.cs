namespace TAD.Report.Infrastructure.PowerPoint.Constants;

/// <summary>
/// Corporate identity and layout constants from <c>docs/design_guide.md</c>.
/// </summary>
public static class DesignGuide
{
    // --- Brand colors (RGB hex without # for Open XML srgbClr/@val) ---
    public const string MainCorporateColorHex = "1E3A8A";
    public const string PassColorHex = "108981";
    public const string FailColorHex = "EF4444";
    public const string NeutralDarkColorHex = "333333";

    /// <summary>Full #RRGGBB form for UI or diagnostics.</summary>
    public const string MainCorporateColor = "#1E3A8A";
    public const string PassColor = "#108981";
    public const string FailColor = "#EF4444";
    public const string NeutralDarkColor = "#333333";

    // --- Typography ---
    public const string PrimaryFont = "맑은 고딕";
    public const string SecondaryFont = "나눔고딕";

    // --- EMU math (1 in = 914_400 EMU) ---
    public const long EmuPerInch = 914_400L;

    public static long EmuFromInch(double inches) => (long)(inches * EmuPerInch);

    /// <summary>Slide 3 failure screenshot frame: top-left per design guide (inches).</summary>
    public static long FailureImageOffsetXEmu => EmuFromInch(6.5);

    public static long FailureImageOffsetYEmu => EmuFromInch(1.8);

    /// <summary>Bounding box width/height inside which the screenshot is letterboxed (contain).</summary>
    public static long FailureImageMaxWidthEmu => EmuFromInch(6.3);

    public static long FailureImageMaxHeightEmu => EmuFromInch(5.0);

    /// <summary>Company logo anchor (inches), all slides except cover.</summary>
    public static long CompanyLogoOffsetXEmu => EmuFromInch(11.5);

    public static long CompanyLogoOffsetYEmu => EmuFromInch(0.1);

    /// <summary>Default logo width on slide (inches); height follows image aspect ratio.</summary>
    public const double CompanyLogoWidthInches = 1.35;
}

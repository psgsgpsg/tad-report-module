# TAD Test Report System Specification

## 1. System Overview & Target Environment
- **Objective**: Create a WPF-based Test Case PPTX Report Generator module.
- **Root Namespace**: Strictly locked to `TAD.Report`

### 1.1 IDE & Toolset Cross-Compatibility
- **Main Solution IDE**: Visual Studio 2022 (v17.x)
- **Module Development IDE**: Visual Studio 2026 (v18.x)
- **Compatibility Rule**: 
  - 이 모듈은 Visual Studio 2026 환경에서 .NET 8.0을 타겟으로 빌드되지만, 향후 VS2022 기반의 본 솔루션에 소스 코드로 통합되어야 하므로 **C# 12 표준 문법 및 .NET 8.0 SDK 표준 라이브러리 규격만을 엄격히 준수**합니다.
  - VS2026 전용 AI 기능이나 타겟 SDK 버전 전용 컴파일러 옵션은 메인 솔루션 병합 시 빌드 크래시를 유발하므로 사용을 절대 금지합니다.

### 1.2 Reference Development PC Specification (Host Hardware)
본 모듈의 빌드 및 검증이 수행된 개발 장비의 하드웨어 기준 명세입니다.
- **Host OS**: Windows 11 Pro (64-bit)
- **Processor (CPU)**: 12th Gen Intel(R) Core(TM) i7-1260P(2.10 GHz)
- **System Memory (RAM)**: 16.0GB
- **Target Framework**: .NET 8.0 (Long-Term Support)
- **UI Architecture**: WPF with MVVM Pattern

---

## 2. Directory & Solution Structure
The solution folder hierarchy is structured as follows. AI must generate and reference source code files exactly within these relative paths.

```text
tad-report-module/
│
├── .gitignore
├── report-specification.md                # This instruction file
│
├── docs/                                  # Documentation and layout images
│   ├── deployment-guide.md                # Runtime asset deployment rules
│   ├── design_guide.md                    # Corporate colors and asset rules
│   ├── export-report-customization-guide.md # Export report customization guide
│   ├── optimization-change-log.md         # Summary of applied changes
│   ├── powerpoint_template_layout.png     # Blueprint layout
│   ├── project-analysis.md                # Architecture and implementation analysis
│   └── winforms-integration-guide.md      # WinForms host integration guide
│
├── assets/                                # Static resources for runtime
│   ├── company_logo.png                   # Embedded logo image
│   ├── export_report_screenshot.png       # Optional sample failure screenshot
│   └── tad_report_template.pptx           # PPTX slide master template
│
└── src/                                   # C# Source Code Root
    ├── TAD.Report.sln                     # Root Solution File
    │
    ├── TAD.Report.Core/                   # Domain & Pure Logic (No external PPT/UI dependencies)
    │   ├── TAD.Report.Core.csproj
    │   ├── Models/
    │   │   ├── TestCaseResult.cs
    │   │   ├── DailyTrend.cs
    │   │   └── ReportData.cs
    │   └── Interfaces/
    │       └── IReportGenerator.cs
    │
    ├── TAD.Report.Infrastructure.PowerPoint/ # PPTX Processing (OpenXML Engine)
    │   ├── TAD.Report.Infrastructure.PowerPoint.csproj
    │   ├── Constants/
    │   │   └── DesignGuide.cs
    │   └── Services/
    │       └── PowerPointReportGenerator.cs
    │
    ├── TAD.Report.WinFormsAdapter/        # Optional WinForms integration facade
    │   ├── TAD.Report.WinFormsAdapter.csproj
    │   ├── ReportExportOptions.cs
    │   ├── ReportExportService.cs
    │   └── WinFormsReportExporter.cs
    │
    └── TAD.Report.App.WPF/                # Desktop Application UI (MVVM Window)
        ├── TAD.Report.App.WPF.csproj
        ├── App.xaml
        ├── App.xaml.cs                    # IoC Container (ServiceCollection Setup)
        ├── Views/
        │   └── MainWindow.xaml
        └── ViewModels/
            └── MainViewModel.cs
```

---

## 3. Domain Models (`TAD.Report.Core`)
Ensure the following models are implemented precisely.

### TestCaseResult
- `int No`: Sequential number
- `string Name`: Test name (maps to `{{NAME}}`)
- `string Result`: "PASS" or "FAIL" (case-insensitive when counted/rendered)
- `string Description`: Error logs or details (maps to `{{DESC}}`)
- `byte[] Screenshot`: Failure screenshot image byte array
- `string Remarks`: Additional comments

### DailyTrend
- `string Date`: Format "MM-dd"
- `int PassCount`: Total passed tests on this day
- `int FailCount`: Total failed tests on this day

### ReportData
- `string Title`: Report main title (maps to `{{TITLE}}`)
- `string Date`: Execution date (maps to `{{DATE}}`)
- `byte[] CompanyLogo`: Top-right company logo image
- `List<TestCaseResult> TestCases`: Collection of results
- `List<DailyTrend> Trends`: Collection of trends
- **Calculated Fields**:
  - `Total`: `TestCases.Count` (maps to `{{TOTAL}}`)
  - `Pass`: `TestCases.Count(t => Result equals "PASS", ignoring case)` (maps to `{{PASS}}`)
  - `Fail`: `TestCases.Count(t => Result equals "FAIL", ignoring case)` (maps to `{{FAIL}}`)
  - `Rate`: `Total > 0 ? (double)Pass / Total * 100 : 0` (maps to `{{RATE}}%`)

---

## 4. Design & Layout Rules (`TAD.Report.Infrastructure.PowerPoint`)
Strictly follow corporate design guidelines defined in `docs/design_guide.md` using code constants.
- **Main Corporate Color**: `#1E3A8A` (Navy Blue)
- **Status - PASS Color**: `#108981` (Green)
- **Status - FAIL Color**: `#EF4444` (Red)
- **Primary Font**: "맑은 고딕" (Malgun Gothic)
- **Company Logo Position**: `Inches(11.5), Inches(0.1)` on all non-cover slides

### Slide Component Rules
1. **Slide 1 & 2 (Text Replacement)**: Scan shape text frames and swap placeholders `{{TITLE}}`, `{{DATE}}`, `{{TOTAL}}`, `{{PASS}}`, `{{FAIL}}`, `{{RATE}}` with `ReportData` metrics.
2. **Slide 3 (Dynamic Scaling)**: Clone the failure detail slide on the fly for entries where `Result == "FAIL"` ignoring case. Replace `{{NAME}}`, `{{RESULT}}`, and `{{DESC}}`, then render the `Screenshot` byte stream into the target image region while preserving aspect ratio.
   - Prefer a PowerPoint picture object named `TAD_FailureScreenshot`.
   - If that picture object is missing and screenshot bytes are valid, insert a new picture at the configured failure image coordinates.
   - If screenshot bytes are empty or invalid, skip only image insertion and keep report generation running.
3. **Slide 4 (Dynamic Table Generation)**: Build dynamic table rows matching `TestCases` array length. Colorize result cell foreground dynamically (Green/Red).
4. **Charts (Slide 2 & 5)**: 
   - Slide 2: Feed `Pass` and `Fail` integers to the Pie Chart wrapper.
   - Slide 5: Feed `List<DailyTrend>` timestamps and counts to the Line Chart wrapper.

---

## 5. UI & Injection Rules (`TAD.Report.App.WPF`)
- Build `MainViewModel` cleanly decoupled from UI using `CommunityToolkit.Mvvm`.
- Use Microsoft Dependency Injection (`Microsoft.Extensions.DependencyInjection`).
- Resolve `IReportGenerator` interface via structural constructor injection. Never instantiate `PowerPointReportGenerator` manually inside ViewModels.
- Run all I/O and report export tasks fully out-of-process asynchronously using C# `async/await`.

---

## 6. Required NuGet Packages & CLI Dependencies
Every project layers must only use the specified packages below. Do not import unlisted third-party UI or PDF engines.

### 1) TAD.Report.Core
- No external NuGet packages needed (Pure .NET 8 Standard Library).

### 2) TAD.Report.Infrastructure.PowerPoint
- `DocumentFormat.OpenXml` (Version 3.1.x)
- `SixLabors.ImageSharp` (Version 3.1.x)

### 3) TAD.Report.App.WPF
- `CommunityToolkit.Mvvm` (Version 8.2.x)
- `Microsoft.Extensions.DependencyInjection` (Version 8.0.x)

### 4) TAD.Report.WinFormsAdapter
- No external NuGet packages needed.
- References `TAD.Report.Core` and `TAD.Report.Infrastructure.PowerPoint`.
- Targets `net8.0-windows` with `UseWindowsForms=true`.

---

## 7. Enterprise Production Guidelines

### 1) Logging Strategy
- Inject and use **`Microsoft.Extensions.Logging.Abstractions.ILogger`** in all Infrastructure services.

### 2) Error Handling & Result Pattern
- Methods in `IReportGenerator` must return a unified wrapper object (e.g., `Result<byte[]>`) indicating `IsSuccess`, `ErrorMessage`, and `Value`.

### 3) Configuration
- Externalize all metadata configurations into standard **`appsettings.json`**.
- Map settings using `IOptions<T>` patterns via Dependency Injection.
- Runtime asset deployment rules are documented in `docs/deployment-guide.md`.

### 4) Testability
- Prepare for an **xUnit** Test Project (`TAD.Report.Tests`) to run pipeline automation via `dotnet test`.
- Current regression coverage verifies:
  - PPTX generation returns a non-empty binary payload.
  - Failure detail placeholders such as `{{RESULT}}` and `{{DESC}}` are replaced.
  - Failure slides are cloned when multiple failures exist.
  - Failure detail slides are removed when no failures exist.
  - Invalid logo or screenshot image bytes do not stop report generation.
  - `PASS`/`FAIL` calculated fields are case-insensitive.

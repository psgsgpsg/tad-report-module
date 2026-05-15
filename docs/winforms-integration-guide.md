# WinForms 통합 가이드

본 문서는 WinForms 기반 본 프로젝트에서 TAD Report PPTX 생성 기능을 가져다 쓰는 방법을 설명합니다.

## 1. 권장 통합 구조

WinForms 본 프로젝트는 WPF 앱을 참조하지 않고, 아래 라이브러리만 참조하는 구성이 가장 안전합니다.

```text
WinForms 본 프로젝트
├─ TAD.Report.Core
├─ TAD.Report.Infrastructure.PowerPoint
└─ TAD.Report.WinFormsAdapter
```

`TAD.Report.App.WPF`는 샘플/데모 앱으로만 유지합니다.

## 2. 추가된 어댑터 프로젝트

새 프로젝트 `TAD.Report.WinFormsAdapter`는 WinForms 화면에서 보고서 생성을 쉽게 호출하기 위한 얇은 어댑터입니다.

| 파일 | 역할 |
| --- | --- |
| `ReportExportOptions.cs` | 템플릿 경로, 로고 경로, 저장 대화상자 기본값 |
| `ReportExportService.cs` | UI와 무관한 PPTX 생성 및 파일 저장 서비스 |
| `WinFormsReportExporter.cs` | WinForms `SaveFileDialog`, `MessageBox` 포함 편의 API |

## 3. 가장 간단한 사용법

WinForms 버튼 클릭 이벤트에서 아래처럼 호출할 수 있습니다.

```csharp
private async void btnExportReport_Click(object sender, EventArgs e)
{
    var data = BuildReportDataFromCurrentScreen();

    var options = new ReportExportOptions
    {
        TemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "tad_report_template.pptx"),
        CompanyLogoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "company_logo.png"),
        DefaultFileName = "TAD_TestReport.pptx",
    };

    await WinFormsReportExporter.ExportWithSaveDialogAsync(this, data, options);
}
```

## 4. UI를 직접 제어하고 싶을 때

본 프로젝트에서 이미 저장 경로를 가지고 있거나, 자체 메시지 처리를 하고 싶다면 `ReportExportService`를 직접 사용합니다.

```csharp
var options = new ReportExportOptions
{
    TemplatePath = templatePath,
    CompanyLogoPath = logoPath,
};

var service = new ReportExportService(options);
var result = await service.ExportAsync(reportData, outputPath);

if (!result.IsSuccess)
{
    MessageBox.Show(result.ErrorMessage, "보고서 생성 오류");
    return;
}
```

## 5. ReportData 매핑 예시

본 프로젝트의 테스트 결과 객체를 `ReportData`로 변환하는 매퍼를 두는 것을 권장합니다.

```csharp
private static ReportData BuildReportDataFromCurrentScreen()
{
    return new ReportData
    {
        Title = "TAD 자동화 테스트 결과 보고서",
        Date = DateTime.Now.ToString("yyyy-MM-dd"),
        TestCases =
        [
            new TestCaseResult
            {
                No = 1,
                Name = "Export report",
                Result = "FAIL",
                Description = "Timeout waiting for dialog.",
                Screenshot = File.Exists(screenshotPath) ? File.ReadAllBytes(screenshotPath) : [],
                Remarks = "확인 필요",
            },
        ],
        Trends =
        [
            new DailyTrend { Date = "05-15", PassCount = 10, FailCount = 1 },
        ],
    };
}
```

## 6. assets 배포

WinForms 실행 폴더 기준으로 아래 파일이 필요합니다.

```text
assets/
├─ tad_report_template.pptx
├─ company_logo.png
└─ export_report_screenshot.png   # 선택
```

복사 규칙과 `.csproj` 예시는 `docs/deployment-guide.md`에서 관리합니다.

## 7. 주의사항

- `PowerPointReportGenerator`는 WPF에 의존하지 않습니다.
- WinForms 프로젝트에서는 `TAD.Report.App.WPF`를 참조하지 않습니다.
- 템플릿과 로고 경로는 `ReportExportOptions`로 명시하는 방식을 권장합니다.
- 스크린샷이 없으면 `Screenshot = []`를 넣으면 됩니다.
- `PASS`/`FAIL` 값은 대소문자를 구분하지 않습니다.

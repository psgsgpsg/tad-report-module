# 배포 가이드

본 문서는 WPF 샘플 앱 또는 WinForms 본 프로젝트에서 PPTX 보고서 생성 기능을 실행할 때 필요한 리소스 배포 규칙을 정리합니다.

## 1. 실행 폴더 구조

보고서 생성기는 기본적으로 실행 파일 폴더 아래의 `assets` 폴더를 사용할 수 있습니다.

```text
실행폴더/
├─ TAD.Report.App.WPF.exe 또는 본 WinForms 실행 파일
└─ assets/
   ├─ tad_report_template.pptx
   ├─ company_logo.png
   └─ export_report_screenshot.png   # 선택
```

| 파일 | 필수 여부 | 용도 |
| --- | --- | --- |
| `tad_report_template.pptx` | 필수 | PPTX 보고서 템플릿 |
| `company_logo.png` | 권장 | 표지 제외 슬라이드 우상단 회사 로고 |
| `export_report_screenshot.png` | 선택 | WPF 샘플 Export report 실패 스크린샷 |

`company_logo.png`는 회사 로고 전용입니다. Export report 스크린샷 대체 이미지로 사용하지 않습니다.

## 2. WPF 샘플 앱 배포

현재 `TAD.Report.App.WPF.csproj`는 루트 `assets`의 템플릿과 회사 로고를 출력 폴더로 복사합니다.

```xml
<None Include="..\..\assets\tad_report_template.pptx"
      Link="assets\tad_report_template.pptx"
      CopyToOutputDirectory="PreserveNewest"
      Condition="Exists('..\..\assets\tad_report_template.pptx')" />
<None Include="..\..\assets\company_logo.png"
      Link="assets\company_logo.png"
      CopyToOutputDirectory="PreserveNewest"
      Condition="Exists('..\..\assets\company_logo.png')" />
```

WPF 샘플에서 `export_report_screenshot.png`를 사용하려면 같은 방식으로 추가합니다.

```xml
<None Include="..\..\assets\export_report_screenshot.png"
      Link="assets\export_report_screenshot.png"
      CopyToOutputDirectory="PreserveNewest"
      Condition="Exists('..\..\assets\export_report_screenshot.png')" />
```

## 3. WinForms 본 프로젝트 배포

WinForms 본 프로젝트에서 `ReportExportOptions`로 경로를 명시하는 방식을 권장합니다.

```csharp
var options = new ReportExportOptions
{
    TemplatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "tad_report_template.pptx"),
    CompanyLogoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "company_logo.png"),
};
```

본 프로젝트 `.csproj`에는 아래처럼 리소스 복사 항목을 둡니다.

```xml
<ItemGroup>
  <None Include="assets\tad_report_template.pptx" CopyToOutputDirectory="PreserveNewest" />
  <None Include="assets\company_logo.png" CopyToOutputDirectory="PreserveNewest" />
  <None Include="assets\export_report_screenshot.png"
        CopyToOutputDirectory="PreserveNewest"
        Condition="Exists('assets\export_report_screenshot.png')" />
</ItemGroup>
```

## 4. 경로 정책

- 라이브러리 기본값은 `AppContext.BaseDirectory/assets`입니다.
- 본 프로젝트에서는 `ReportExportOptions.TemplatePath`, `ReportExportOptions.CompanyLogoPath`로 명시하는 편이 안전합니다.
- 테스트나 독립 실행 환경에서는 출력 폴더에 `assets`가 복사됐는지 먼저 확인합니다.

## 5. 확인 명령

```powershell
cd C:\Workspace\tad-report-module\src
dotnet build TAD.Report.sln
Get-ChildItem .\artifacts\bin\TAD.Report.App.WPF\debug\assets
```

예상 파일:

```text
company_logo.png
tad_report_template.pptx
```

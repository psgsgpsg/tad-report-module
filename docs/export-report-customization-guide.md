# Export Report 커스터마이징 가이드

이 문서는 `Export report` 실패 상세 슬라이드의 텍스트, 결과값, 상세 설명, 스크린샷 이미지, 이미지 영역을 코드에서 수정하는 방법을 정리합니다.

## 1. 관련 파일

| 목적 | 파일 |
| --- | --- |
| 샘플 Export report 데이터 | `src/TAD.Report.App.WPF/ViewModels/MainViewModel.cs` |
| PPTX 생성 핵심 로직 | `src/TAD.Report.Infrastructure.PowerPoint/Services/PowerPointReportGenerator.cs` |
| 이미지/로고 좌표 및 크기 상수 | `src/TAD.Report.Infrastructure.PowerPoint/Constants/DesignGuide.cs` |
| 도메인 모델 | `src/TAD.Report.Core/Models/TestCaseResult.cs` |
| PPT 템플릿 | `assets/tad_report_template.pptx` |

## 2. Export report 데이터 수정 위치

WPF 앱의 샘플 `Export report` 항목은 `MainViewModel.cs`의 `CreateDesignTimeReportSample()` 안에 있습니다.

```csharp
new TestCaseResult
{
    No = 2,
    Name = "Export report",
    Result = "FAIL",
    Description = "Timeout waiting for dialog.",
    Screenshot = LoadSampleScreenshot(),
    Remarks = "Retry later",
},
```

수정 가능한 값은 다음과 같습니다.

| 속성 | 설명 |
| --- | --- |
| `Name` | 실패 상세 슬라이드의 테스트 이름 |
| `Result` | `{{RESULT}}`에 들어갈 값 |
| `Description` | `{{DESC}}` 또는 상세 설명 영역에 들어갈 값 |
| `Screenshot` | 실패 상세 슬라이드 이미지 영역에 들어갈 이미지 바이트 |
| `Remarks` | 테스트 리스트 슬라이드의 비고 |

## 3. 커스텀 이미지 넣기

현재 샘플은 실행 폴더의 `assets/export_report_screenshot.png`가 있을 때만 스크린샷 영역에 이미지를 넣습니다. 파일이 없으면 스크린샷 삽입을 건너뛰며, 회사 로고를 스크린샷 대체 이미지로 사용하지 않습니다.

```csharp
private static byte[] LoadSampleScreenshot()
{
    var screenshotPath = Path.Combine(AppContext.BaseDirectory, "assets", "export_report_screenshot.png");
    return File.Exists(screenshotPath) ? File.ReadAllBytes(screenshotPath) : [];
}
```

원하는 이미지 파일을 넣으려면 아래처럼 변경할 수 있습니다.

```csharp
private static byte[] LoadSampleScreenshot()
{
    var imagePath = Path.Combine(AppContext.BaseDirectory, "assets", "export_report_screenshot.png");
    return File.Exists(imagePath) ? File.ReadAllBytes(imagePath) : [];
}
```

이 경우 `assets/export_report_screenshot.png` 파일을 추가하고, `.csproj`에서 출력 폴더로 복사되도록 등록해야 합니다. 배포/복사 규칙의 전체 기준은 `docs/deployment-guide.md`를 따릅니다.

## 4. 실패 상세 슬라이드 치환 위치

`PowerPointReportGenerator.cs`의 `ApplyFailureTextReplacements()`에서 실패 상세 슬라이드 텍스트를 치환합니다.

```csharp
var map = new Dictionary<string, string>(StringComparer.Ordinal)
{
    ["{{NAME}}"] = SafeText(fails[i].Name),
    ["{{RESULT}}"] = SafeText(fails[i].Result),
    ["{{DESC}}"] = description,
    ["상세 내용 / 오류 메시지 /\n로그\n(선택 영역)"] = description,
    ["상세 내용 / 오류 메시지 / 로그 (선택 영역)"] = description,
};
```

템플릿의 플레이스홀더를 추가하려면 이 맵에 항목을 추가합니다.

예시:

```csharp
["{{REMARKS}}"] = SafeText(fails[i].Remarks),
```

그리고 PowerPoint 템플릿에도 `{{REMARKS}}` 텍스트를 넣어야 합니다.

## 5. 상세 내용 영역에 원하는 내용 넣기

PowerPoint 실패 상세 슬라이드의 `"상세 내용 / 오류 메시지 / 로그 (선택 영역)"` 영역은 현재 `TestCaseResult.Description` 값으로 채워집니다.

데이터를 넣는 위치는 `MainViewModel.cs`의 `CreateDesignTimeReportSample()`입니다.

```csharp
new TestCaseResult
{
    No = 2,
    Name = "Export report",
    Result = "FAIL",
    Description = "Timeout waiting for dialog.",
    Screenshot = LoadSampleScreenshot(),
    Remarks = "Retry later",
},
```

원하는 상세 내용을 넣으려면 `Description` 값을 수정합니다.

예시:

```csharp
Description =
    "PPT 보고서 발행 중 저장 대화상자 응답 지연이 발생했습니다.\n" +
    "원인: SaveFileDialog 호출 후 타임아웃\n" +
    "조치: UI 스레드 점유 여부 및 파일 경로 권한 확인 필요",
```

실제 운영 데이터에서는 `ReportData.TestCases`를 만드는 쪽에서 `Description`에 오류 메시지, 로그, 스택 트레이스, 재현 절차 등을 넣으면 됩니다.

```csharp
reportData.TestCases.Add(new TestCaseResult
{
    No = 2,
    Name = "Export report",
    Result = "FAIL",
    Description = errorLogText,
    Screenshot = screenshotBytes,
    Remarks = "확인 필요",
});
```

문구가 PPT에 들어가는 실제 치환 코드는 `PowerPointReportGenerator.cs`의 `ApplyFailureTextReplacements()`입니다.

```csharp
var description = SafeText(fails[i].Description);
var map = new Dictionary<string, string>(StringComparer.Ordinal)
{
    ["{{NAME}}"] = SafeText(fails[i].Name),
    ["{{RESULT}}"] = SafeText(fails[i].Result),
    ["{{DESC}}"] = description,
    ["상세 내용 / 오류 메시지 /\n로그\n(선택 영역)"] = description,
    ["상세 내용 / 오류 메시지 / 로그 (선택 영역)"] = description,
};
```

즉, 수정이 필요한 지점은 보통 두 곳입니다.

| 수정 목적 | 수정 위치 |
| --- | --- |
| 상세 내용에 들어갈 실제 문구 변경 | `TestCaseResult.Description` 값을 만드는 코드 |
| 템플릿의 다른 문구나 토큰을 상세 내용으로 치환 | `ApplyFailureTextReplacements()`의 `map` |

권장 방식은 PowerPoint 템플릿의 상세 내용 영역에 안내 문구 대신 `{{DESC}}` 플레이스홀더를 넣는 것입니다. 그러면 코드에서는 아래 항목만 유지해도 됩니다.

```csharp
["{{DESC}}"] = description,
```

템플릿에 고정 안내 문구를 그대로 두면, 문구의 줄바꿈이나 띄어쓰기가 바뀔 때 치환이 실패할 수 있습니다.

## 6. 이미지 삽입 로직 위치

실패 상세 슬라이드에 이미지를 실제로 넣는 코드는 `BindFailureScreenshot()`입니다.

```csharp
private static void BindFailureScreenshot(SlidePart slidePart, byte[] screenshot)
```

현재 동작 방식은 다음과 같습니다.

1. 실패 상세 슬라이드에서 `TAD_FailureScreenshot` 이름을 가진 `Picture`를 우선 찾습니다.
2. 없으면 회사 로고가 아닌 기존 `Picture`를 fallback으로 사용합니다.
3. 교체할 `Picture`가 없으면 지정 좌표에 새 `Picture`를 생성합니다.
4. `TestCaseResult.Screenshot` 바이트를 이미지로 로드합니다.
5. 지정된 이미지 영역 안에 비율 유지 방식으로 맞춥니다.
6. 기존 이미지 파트를 교체하거나 새 이미지 파트를 추가합니다.

중요 코드:

```csharp
var picture = FindFailureScreenshotPicture(slidePart);
```

템플릿에 이미지가 여러 개 있을 수 있으므로, 가능하면 PowerPoint에서 스크린샷 placeholder 개체 이름을 `TAD_FailureScreenshot`으로 지정하는 것을 권장합니다.

## 7. 이미지 영역 위치/크기 수정

스크린샷 영역의 위치와 크기는 `DesignGuide.cs`에서 조정합니다.

```csharp
public static long FailureImageOffsetXEmu => EmuFromInch(6.5);
public static long FailureImageOffsetYEmu => EmuFromInch(1.8);

public static long FailureImageMaxWidthEmu => EmuFromInch(6.3);
public static long FailureImageMaxHeightEmu => EmuFromInch(5.0);
```

| 상수 | 의미 |
| --- | --- |
| `FailureImageOffsetXEmu` | 이미지 영역의 왼쪽 X 위치 |
| `FailureImageOffsetYEmu` | 이미지 영역의 위쪽 Y 위치 |
| `FailureImageMaxWidthEmu` | 이미지 영역 최대 너비 |
| `FailureImageMaxHeightEmu` | 이미지 영역 최대 높이 |

예시: 이미지를 더 왼쪽 위로 옮기고 크게 표시하기

```csharp
public static long FailureImageOffsetXEmu => EmuFromInch(5.8);
public static long FailureImageOffsetYEmu => EmuFromInch(1.5);

public static long FailureImageMaxWidthEmu => EmuFromInch(6.8);
public static long FailureImageMaxHeightEmu => EmuFromInch(5.3);
```

## 8. 이미지 맞춤 방식

현재 이미지는 `contain` 방식입니다. 즉 이미지 전체가 잘리지 않도록 영역 안에 맞추고, 남는 여백은 가운데 정렬합니다.

관련 코드:

```csharp
var (cx, cy) = MeasureContainedExtentEmu(image.Width, image.Height, maxCx, maxCy);
var offsetX = originX + (maxCx - cx) / 2;
var offsetY = originY + (maxCy - cy) / 2;
```

이미지를 영역에 꽉 채우고 일부 잘려도 되는 `cover` 방식이 필요하면 `MeasureContainedExtentEmu()` 대신 별도 계산 로직을 추가해야 합니다.

## 9. 스크린샷 placeholder 이름

현재 코드는 `TAD_FailureScreenshot` 이름을 우선 사용하고, 없으면 fallback으로 회사 로고가 아닌 기존 이미지를 사용합니다. 템플릿을 더 안정적으로 만들려면 PowerPoint에서 스크린샷 영역 이미지의 개체 이름을 `TAD_FailureScreenshot`으로 지정하세요.

```csharp
var picture = slidePart.Slide.Descendants<P.Picture>()
    .FirstOrDefault(pic =>
        pic.NonVisualPictureProperties?
            .NonVisualDrawingProperties?
            .Name?
            .Value == "TAD_FailureScreenshot");
```

해당 이름의 이미지가 없고 스크린샷 바이트가 유효하면, 코드는 `DesignGuide.cs`의 이미지 좌표에 새 그림 개체를 삽입합니다.

## 10. 수정 후 검증 명령

수정 후에는 아래 명령으로 빌드와 테스트를 확인합니다.

```powershell
cd C:\Workspace\tad-report-module\src
dotnet build TAD.Report.sln
dotnet test TAD.Report.sln
```

실행 파일 출력 위치에 `assets`가 복사됐는지도 확인합니다.

```powershell
Get-ChildItem C:\Workspace\tad-report-module\src\artifacts\bin\TAD.Report.App.WPF\debug\assets
```

자세한 assets 배포 규칙은 `docs/deployment-guide.md`를 참고합니다.

## 11. 빠른 수정 체크리스트

1. Export report 문구 변경
   - `MainViewModel.cs`의 `Name`, `Description`, `Result` 수정

2. Export report 이미지 변경
   - `Screenshot`에 원하는 이미지 파일의 `byte[]` 지정
   - 이미지 파일을 `assets`에 추가
   - `.csproj`에 `CopyToOutputDirectory` 등록

3. 이미지 위치/크기 변경
   - `DesignGuide.cs`의 `FailureImageOffsetXEmu`, `FailureImageOffsetYEmu`, `FailureImageMaxWidthEmu`, `FailureImageMaxHeightEmu` 수정

4. 템플릿 플레이스홀더 추가
   - `assets/tad_report_template.pptx`에 `{{TOKEN}}` 추가
   - `PowerPointReportGenerator.cs`의 치환 맵에 같은 토큰 추가

5. 템플릿 이미지 영역 안정화
   - PPT에서 스크린샷 이미지 개체 이름을 `TAD_FailureScreenshot`으로 지정
   - 코드가 해당 이름을 우선 사용하므로 fallback 동작을 줄일 수 있음

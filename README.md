# TAD Report Module

[![.NET 8.0](https://img.shields.io/badge/.NET-8.0-512BD4.svg)](https://dotnet.microsoft.com/download/dotnet/8.0)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**TAD Report Module**은 테스트 실행 결과 데이터를 분석하여 시각화된 PowerPoint(`.pptx`) 보고서를 자동으로 생성하는 .NET 8 기반의 엔터프라이즈급 리포팅 엔진입니다.

이 모듈은 사전에 정의된 PowerPoint 슬라이드 마스터 템플릿을 기반으로 텍스트 플레이스홀더 치환, 동적 슬라이드 복제, 차트 데이터 업데이트 및 테스트 결과 테이블 자동 생성을 수행합니다.

## 🚀 주요 기능

- **자동 보고서 생성**: `ReportData` 모델을 입력받아 단 몇 초 만에 완성된 PPTX 파일 생성.
- **동적 슬라이드 복제**: 실패(FAIL)한 테스트 케이스 수에 맞춰 상세 분석 슬라이드를 자동으로 복제 및 제거.
- **이미지 자동 바인딩**: 실패 스크린샷 및 기업 로고를 지정된 영역에 비율을 유지하며 자동 삽입.
- **데이터 시각화**: 종합 요약 파이 차트 및 일자별 Pass/Fail 추이 라인 차트 자동 갱신.
- **엔터프라이즈 디자인 준수**: 지정된 기업 컬러(#1E3A8A) 및 레이아웃 가이드라인 엄격 준수.
- **유연한 연동**: WPF 메인 앱 제공 및 기존 WinForms 프로젝트와의 통합을 위한 어댑터 레이어 포함.

## 🛠 기술 스택

- **Runtime**: .NET 8.0 (LTS)
- **Language**: C# 12
- **UI Framework**: WPF (Windows Presentation Foundation)
- **Engines**: 
  - [Open XML SDK](https://github.com/dotnet/Open-XML-SDK) (PowerPoint 조작)
  - [ImageSharp](https://github.com/SixLabors/ImageSharp) (이미지 처리)
- **Architecture**: MVVM Pattern with CommunityToolkit.Mvvm, Microsoft DI

## 📂 프로젝트 구조

```text
src/
├── TAD.Report.Core/                   # 도메인 모델 및 인터페이스 (Pure Logic)
├── TAD.Report.Infrastructure.PowerPoint/ # Open XML 기반 PPTX 생성 엔진
├── TAD.Report.WinFormsAdapter/        # WinForms 프로젝트 통합용 어댑터
├── TAD.Report.App.WPF/                # 데스크톱 리포팅 도구 (WPF 앱)
└── TAD.Report.Tests/                  # xUnit 기반 단위 및 통합 테스트
```

## 🏁 시작하기

### 사전 요구 사항
- [.NET 8.0 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) 이상
- Visual Studio 2022 v17.x 또는 Visual Studio 2026 v18.x (권장)

### 설치 및 빌드
```powershell
# 저장소 복제
git clone https://github.com/psgsgpsg/tad-report-module.git
cd tad-report-module/src

# 솔루션 빌드
dotnet build TAD.Report.sln

# 테스트 실행
dotnet test TAD.Report.sln
```

### 실행
```powershell
dotnet run --project TAD.Report.App.WPF/TAD.Report.App.WPF.csproj
```

## 📄 보고서 템플릿 규약

보고서 생성기는 `assets/tad_report_template.pptx` 파일을 기반으로 동작하며, 다음 플레이스홀더를 지원합니다:

| 플레이스홀더 | 설명 |
| :--- | :--- |
| `{{TITLE}}` | 보고서 제목 |
| `{{DATE}}` | 실행 일자 |
| `{{TOTAL}}` | 전체 테스트 케이스 수 |
| `{{PASS}}` | 통과(PASS) 수 |
| `{{FAIL}}` | 실패(FAIL) 수 |
| `{{RATE}}` | 통과율 (%) |

상세한 디자인 가이드 및 좌표 설정은 `docs/design_guide.md`와 `src/TAD.Report.Infrastructure.PowerPoint/Constants/DesignGuide.cs`를 참조하세요.

## 🤝 기여하기

1. 이 저장소를 Fork 합니다.
2. 새로운 기능 브랜치를 만듭니다 (`git checkout -b feature/NewFeature`).
3. 변경 사항을 Commit 합니다 (`git commit -m 'Add some NewFeature'`).
4. 브랜치에 Push 합니다 (`git push origin feature/NewFeature`).
5. Pull Request를 생성합니다.

## 📜 라이선스

이 프로젝트는 MIT 라이선스에 따라 라이선스가 부여됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

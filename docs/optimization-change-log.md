# 최적화 변경 내역

본 문서는 코드와 문서 구조에 반영된 주요 변경사항만 요약합니다. 상세 사용법은 각 전용 가이드를 참고합니다.

## 변경 요약

| 구분 | 변경 내용 | 상세 문서 |
| --- | --- | --- |
| 템플릿 안정성 | PPTX 필수 슬라이드, placeholder, 차트, 표 검사 추가 | `docs/project-analysis.md` |
| 실패 스크린샷 | `TAD_FailureScreenshot` 이름 우선, 없으면 fallback 또는 새 이미지 삽입 | `docs/export-report-customization-guide.md` |
| 이미지 방어 | 깨진 로고/스크린샷 바이트가 있어도 보고서 생성 지속 | `docs/project-analysis.md` |
| 샘플 이미지 | `company_logo.png`를 스크린샷 대체 이미지로 사용하지 않도록 분리 | `docs/deployment-guide.md` |
| WinForms 통합 | `TAD.Report.WinFormsAdapter` 프로젝트 추가 | `docs/winforms-integration-guide.md` |
| 로고 위치 | 회사 로고 앵커를 `Inches(11.5), Inches(0.1)`로 조정 | `docs/design_guide.md` |
| 결과 판정 | `PASS`/`FAIL` 계산을 대소문자 무시 방식으로 통일 | `docs/project-analysis.md` |
| 테스트 | 슬라이드 복제/제거, placeholder, 이미지 방어, 대소문자 테스트 보강 | `docs/project-analysis.md` |

## 현재 검증 상태

```text
dotnet build TAD.Report.sln
결과: 성공, 경고 0개, 오류 0개

dotnet test TAD.Report.sln
결과: 테스트 7개 통과, 실패 0개
```

## 남은 개선 후보

- `ILogger<PowerPointReportGenerator>` 주입으로 이미지 건너뜀, 템플릿 검사 실패, placeholder 잔존 여부 기록
- 템플릿 경로, 로고 경로, 이미지 좌표를 옵션 또는 설정 파일로 외부화
- PowerPoint가 placeholder를 여러 Run으로 나누는 경우까지 방어하는 문단 단위 치환 로직 추가
- 테스트 전용 PNG 생성 헬퍼를 추가해 테스트가 회사 로고 파일에 덜 의존하도록 개선

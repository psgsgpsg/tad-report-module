# TAD Test Report System - Design Guide

본 문서는 `TAD` 테스트 결과 보고서 PPTX 자동 생성 모듈이 준수해야 하는 시각적 정체성(Corporate Identity)과 레이아웃 배치 규칙을 정의합니다. 모든 코드 상수와 스타일 적용 로직은 본 가이드라인을 엄격히 따릅니다.

---

## 1. 전역 디자인 시스템 (Global Design System)

### 1.1 지정 색상 (Brand Colors)
보고서의 일관성을 위해 아래의 고정 Hex RGB 색상 코드만을 사용합니다.
- **Main Corporate Color**: `#1E3A8A` (딥 네이비) - 표지, 제목 바, 슬라이드 헤더 테두리
- **Status - PASS Color**: `#108981` (청록 계열 그린) - 패스 텍스트, 성공 그래프 지표
- **Status - FAIL Color**: `#EF4444` (경고 레드) - 실패 텍스트, 에러 하이라이트, 실패 그래프 지표
- **Neutral Dark Color**: `#333333` (다크 그레이) - 본문 일반 텍스트 및 기본 표 글자

### 1.2 타이포그래피 (Typography)
C# 오픈소스 라이브러리로 텍스트 렌더링 시 폰트 깨짐 방지를 위해 시스템 기본 폰트를 지정합니다.
- **주 폰트 (Primary)**: `맑은 고딕 (Malgun Gothic)` - 제목, 데이터 수치, 태그 치환 텍스트
- **부 폰트 (Secondary)**: `나눔고딕 (NanumGothic)` 또는 `Arial` - 표 내부 상세 내용 및 본문

---

## 2. 슬라이드별 컴포넌트 배치 및 데이터 바인딩 규칙

### Slide 1: 표지 (Cover)
- **비주얼 구성**: 중앙 정렬 레이아웃, 하단 여백에 실행 정보 배치
- **데이터 매핑**:
  - `{{TITLE}}`: 폰트 크기 `44pt`, Bold, 색상 `#1E3A8A`
  - `Date: {{DATE}}`: 폰트 크기 `14pt`, Regular, 색상 `#333333`

### Slide 2: 종합 요약 (Test Summary)
- **비주얼 구성**: 좌측 요약 메트릭스 리스트, 우측 요약 파이 차트 (50:50 분할)
- **데이터 매핑**:
  - `• Total Tests: {{TOTAL}}`
  - `• Passed: {{PASS}}`
  - `• Failed: {{FAIL}}`
  - `• Success Rate: {{RATE}}%`
- **차트 규칙 (Pie Chart)**:
  - 데이터 영역에 `Pass`와 `Fail` 카운트 반영
  - 계열(Series) 색상 매핑: PASS 조각 ➡️ `#108981`, FAIL 조각 ➡️ `#EF4444`

### Slide 3: 실패 세부 정보 (Failure Detail)
- **비주얼 구성**: 좌측 에러 로그 영역, 우측 실패 시점 스크린샷 뷰어
- **동적 생성 조건**: 전체 결과 중 `Result == "FAIL"` 인 아이템 개수만큼 이 슬라이드를 복사(Clone)하여 런타임에 동적 추가
- **결과 판정 기준**: `Result` 값은 대소문자를 구분하지 않고 `"FAIL"`로 판정함 (`"fail"`, `"Fail"` 허용)
- **데이터 매핑**:
  - 제목 바: `Failure Detail: {{NAME}}` (폰트 크기 `24pt`)
  - 결과 텍스트: `{{RESULT}}`
  - 본문 텍스트 박스: `{{DESC}}` (스택 트레이스 및 에러 메시지 출력, 줄바꿈 유지)
- **이미지 바인딩**:
  - 우측 고정 사각형 프레임 영역(`Inches(6.5), Inches(1.8)`) 좌표 내부에 `Screenshot` 바이너리 데이터를 비율 유지 방식으로 삽입
  - PowerPoint 이미지 개체 이름이 `TAD_FailureScreenshot`이면 해당 개체를 우선 교체
  - 해당 이미지 개체가 없고 스크린샷 바이트가 유효하면 코드가 동일 좌표에 새 이미지 개체를 생성
  - `Screenshot`이 비어 있거나 이미지로 읽을 수 없으면 보고서 생성을 중단하지 않고 이미지 삽입만 건너뜀

### Slide 4: 테스트 리스트 (Test Case List)
- **비주얼 구성**: 상단 요약 타이틀, 하단 전체 목록 테이블 (Grid 형태)
- **데이터 매핑**:
  - 테이블 헤더: `No | Test Name | Result | Execution Time | Remarks` 고정
  - `TestCases` 리스트를 순회하며 행(Row)을 하단으로 동적 추가
- **테이블 스타일링 규칙**:
  - 헤더 배경색: `#1E3A8A`, 글자색: `#FFFFFF`
  - 데이터 행 내부 `Result` 셀 텍스트 조건부 서식:
    - 값이 `"PASS"` 일 때 ➡️ 텍스트 색상 `#108981` (Bold)
    - 값이 `"FAIL"` 일 때 ➡️ 텍스트 색상 `#EF4444` (Bold)
    - 대소문자는 구분하지 않음 (`"pass"`, `"fail"` 허용)

### Slide 5: 일자별 추이 (Daily Trend Analysis)
- **비주얼 구성**: 전체 화면을 활용하는 대형 꺾은선형 차트
- **데이터 매핑**:
  - X축 (Categories): 최근 테스트 수행 일자 목록 (`List<DailyTrend>.Date` 목록)
  - Y축 (Values): 해당 일자의 Pass 수와 Fail 수
- **차트 규칙 (Line Chart)**:
  - 계열 1 (PASS 꺾은선): 선 색상 `#108981`, 표식(Marker) 포함
  - 계열 2 (FAIL 꺾은선): 선 색상 `#EF4444`, 표식(Marker) 포함

---

## 3. 개발 프레임워크 제약 사항 (Technical Constraints)

- **정적 리소스**: 우상단 회사 로고(`company_logo.png`)는 모든 슬라이드(표지 제외)의 고정 좌표(`Inches(11.5), Inches(0.1)`)에 동일한 크기로 자동 삽입되어야 함.
- **샘플 스크린샷 리소스**: Export report 샘플 스크린샷은 선택 리소스 `export_report_screenshot.png`를 사용하며, `company_logo.png`를 스크린샷 대체 이미지로 사용하지 않음.
- **예외 처리**: 데이터 바인딩 시 `{{TAG}}` 대상 값이 `null` 이거나 빈 문자열일 경우, 에러를 내지 않고 빈 칸(`""`)으로 치환하여 템플릿 문구 노출을 방지함. 이미지 바이트가 손상된 경우에도 해당 이미지 삽입만 건너뜀.

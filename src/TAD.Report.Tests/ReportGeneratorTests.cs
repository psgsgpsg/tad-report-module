using Xunit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using TAD.Report.Core.Models;
using TAD.Report.Core;
using TAD.Report.Infrastructure.PowerPoint.Services;

namespace TAD.Report.Tests
{
    public class ReportGeneratorTests
    {
        private static readonly byte[] SamplePng =
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "assets", "company_logo.png"));

        // 1. 검증에 사용할 모의(Mock) 데이터 생성 헬퍼 메서드
        private ReportData CreateMockReportData()
        {
            var reportData = new ReportData
            {
                Title = "TAD 자동화 테스트 결과 보고서",
                Date = DateTime.Now.ToString("yyyy-MM-dd"),
                CompanyLogo = SamplePng
            };

            // PASS 데이터 3개 추가
            for (int i = 1; i <= 3; i++)
            {
                reportData.TestCases.Add(new TestCaseResult
                {
                    No = i,
                    Name = $"TAD_Auth_Module_Test_0{i}",
                    Result = "PASS",
                    Description = "정상 작동 완료",
                    Remarks = "N/A"
                });
            }

            // FAIL 데이터 2개 추가 (스크린샷 바이너리 포함)
            for (int i = 4; i <= 5; i++)
            {
                reportData.TestCases.Add(new TestCaseResult
                {
                    No = i,
                    Name = $"TAD_Payment_API_Error_0{i}",
                    Result = "FAIL",
                    Description = "HTTP 500 Internal Server Error 발생\nat TAD.Payment.Service.Process()",
                    Screenshot = SamplePng,
                    Remarks = "긴급 수정 필요"
                });
            }

            // 일자별 추이 더미 데이터 추가
            reportData.Trends.Add(new DailyTrend { Date = "05-12", PassCount = 10, FailCount = 2 });
            reportData.Trends.Add(new DailyTrend { Date = "05-13", PassCount = 15, FailCount = 1 });
            reportData.Trends.Add(new DailyTrend { Date = "05-14", PassCount = 3, FailCount = 2 });

            return reportData;
        }

        // 2. PPTX 파일 생성이 성공하고 바이너리가 리턴되는지 검증
        [Fact]
        public async Task GenerateReport_WithValidData_ReturnsSuccessResult()
        {
            // Arrange (준비)
            var mockData = CreateMockReportData();
            var generator = new PowerPointReportGenerator(); 

            // Act (실행)
            var result = await generator.GenerateReportAsync(mockData);

            // Assert (단언/검증 - xUnit2002 에러 유발 코드 제거 완료)
            Assert.True(result.IsSuccess, result.ErrorMessage);
            Assert.NotNull(result.Value);
            Assert.NotEmpty(result.Value);
        }

        [Fact]
        public async Task GenerateReport_ReplacesFailureDetailPlaceholders()
        {
            // Arrange
            var mockData = CreateMockReportData();
            var generator = new PowerPointReportGenerator();

            // Act
            var result = await generator.GenerateReportAsync(mockData);

            // Assert
            Assert.True(result.IsSuccess, result.ErrorMessage);
            Assert.NotNull(result.Value);

            using var stream = new MemoryStream(result.Value);
            using var document = PresentationDocument.Open(stream, false);
            var text = string.Join(
                Environment.NewLine,
                document.PresentationPart!.SlideParts
                    .SelectMany(slidePart => slidePart.Slide.Descendants<A.Text>())
                    .Select(t => t.Text));

            Assert.DoesNotContain("{{RESULT}}", text);
            Assert.DoesNotContain("{{DESC}}", text);
            Assert.Contains("FAIL", text);
            Assert.Contains("HTTP 500 Internal Server Error", text);
        }

        [Fact]
        public async Task GenerateReport_WithMultipleFailures_ClonesFailureSlides()
        {
            // Arrange
            var mockData = CreateMockReportData();
            var generator = new PowerPointReportGenerator();

            // Act
            var result = await generator.GenerateReportAsync(mockData);

            // Assert
            Assert.True(result.IsSuccess, result.ErrorMessage);
            Assert.NotNull(result.Value);

            using var stream = new MemoryStream(result.Value);
            using var document = PresentationDocument.Open(stream, false);
            var slideCount = document.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().Count();

            Assert.Equal(6, slideCount);
        }

        [Fact]
        public async Task GenerateReport_WithNoFailures_RemovesFailureSlide()
        {
            // Arrange
            var mockData = CreateMockReportData();
            foreach (var testCase in mockData.TestCases)
            {
                testCase.Result = "PASS";
                testCase.Description = string.Empty;
                testCase.Screenshot = [];
            }

            var generator = new PowerPointReportGenerator();

            // Act
            var result = await generator.GenerateReportAsync(mockData);

            // Assert
            Assert.True(result.IsSuccess, result.ErrorMessage);
            Assert.NotNull(result.Value);

            using var stream = new MemoryStream(result.Value);
            using var document = PresentationDocument.Open(stream, false);
            var slideCount = document.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().Count();

            Assert.Equal(4, slideCount);
        }

        [Fact]
        public async Task GenerateReport_WithInvalidImageBytes_SucceedsWithoutImageReplacement()
        {
            // Arrange
            var mockData = CreateMockReportData();
            mockData.CompanyLogo = [1, 2, 3, 4];
            foreach (var failure in mockData.TestCases.Where(t => t.Result == "FAIL"))
                failure.Screenshot = [5, 6, 7, 8];

            var generator = new PowerPointReportGenerator();

            // Act
            var result = await generator.GenerateReportAsync(mockData);

            // Assert
            Assert.True(result.IsSuccess, result.ErrorMessage);
            Assert.NotNull(result.Value);
            Assert.NotEmpty(result.Value);
        }


        // 3. 명세서에 정의된 자동 연산 프로퍼티(Total, Pass, Fail, Rate) 정밀 검증
        [Fact]
        public void ReportData_CalculatedFields_AreCorrect()
        {
            // Arrange & Act
            var data = CreateMockReportData();

            // Assert (명세서 3번 항목 수식 검증)
            Assert.Equal(5, data.Total);       // 3 (PASS) + 2 (FAIL) = 5
            Assert.Equal(3, data.Pass);        // PASS 개수 = 3
            Assert.Equal(2, data.Fail);        // FAIL 개수 = 2
            Assert.Equal(60.0, data.Rate);     // (3 / 5) * 100 = 60.0%
        }

        [Fact]
        public void ReportData_CalculatedFields_AreCaseInsensitive()
        {
            // Arrange
            var data = new ReportData
            {
                TestCases =
                [
                    new TestCaseResult { Result = "pass" },
                    new TestCaseResult { Result = "Pass" },
                    new TestCaseResult { Result = "FAIL" },
                    new TestCaseResult { Result = "fail" },
                ],
            };

            // Assert
            Assert.Equal(4, data.Total);
            Assert.Equal(2, data.Pass);
            Assert.Equal(2, data.Fail);
            Assert.Equal(50.0, data.Rate);
        }
    }
}

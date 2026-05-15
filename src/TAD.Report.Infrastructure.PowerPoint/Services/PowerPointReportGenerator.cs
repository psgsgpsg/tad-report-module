using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;
using TAD.Report.Core;
using TAD.Report.Core.Interfaces;
using TAD.Report.Core.Models;
using TAD.Report.Infrastructure.PowerPoint.Constants;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace TAD.Report.Infrastructure.PowerPoint.Services;

/// <summary>
/// Open XML–based PPTX generator following <c>report-specification.md</c> and <c>design_guide.md</c>.
/// </summary>
public sealed class PowerPointReportGenerator : IReportGenerator
{
    private const int CoverSlideIndex = 0;
    private const int SummarySlideIndex = 1;
    private const int FailureTemplateSlideIndex = 2;
    private const int OriginalTestListSlideIndex = 3;
    private const int OriginalTrendSlideIndex = 4;
    private const int MinimumTemplateSlideCount = 5;
    private const string FailureScreenshotPictureName = "TAD_FailureScreenshot";

    private readonly string _templatePath;
    private readonly string _fallbackLogoPath;

    public PowerPointReportGenerator()
        : this(
            global::System.IO.Path.Combine(AppContext.BaseDirectory, "assets", "tad_report_template.pptx"),
            global::System.IO.Path.Combine(AppContext.BaseDirectory, "assets", "company_logo.png"))
    {
    }

    public PowerPointReportGenerator(string templatePath, string? fallbackLogoPath = null)
    {
        _templatePath = templatePath ?? throw new ArgumentNullException(nameof(templatePath));
        _fallbackLogoPath = fallbackLogoPath ?? global::System.IO.Path.Combine(AppContext.BaseDirectory, "assets", "company_logo.png");
    }

    /// <inheritdoc />
    public async Task<Result<byte[]>> GenerateReportAsync(ReportData data, CancellationToken cancellationToken = default)
    {
        try
        {
            ArgumentNullException.ThrowIfNull(data);

            if (!File.Exists(_templatePath))
            {
                throw new FileNotFoundException($"Report template not found at '{_templatePath}'.");
            }

            var templateBytes = await File.ReadAllBytesAsync(_templatePath, cancellationToken).ConfigureAwait(false);
            var output = await Task.Run(() => BuildFromTemplate(templateBytes, data), cancellationToken).ConfigureAwait(false);
            return Result.Success(output);
        }
        catch (Exception ex)
        {
            return Result.Failure<byte[]>(ex.Message);
        }
    }

    private byte[] BuildFromTemplate(byte[] templatePptx, ReportData data)
    {
        var buffer = new MemoryStream(templatePptx.Length);
        buffer.Write(templatePptx);
        buffer.Position = 0;

        using var presentation = PresentationDocument.Open(buffer, true);
        var presentationPart = presentation.PresentationPart
            ?? throw new InvalidOperationException("Presentation is missing a presentation part.");

        ValidateTemplate(presentationPart);

        var fails = data.TestCases
            .Where(t => string.Equals(t.Result, "FAIL", StringComparison.OrdinalIgnoreCase))
            .ToList();

        AdjustFailureSlides(presentationPart, templatePptx, fails);

        var failureCount = fails.Count;
        var tableSlideIndex = failureCount == 0 ? OriginalTestListSlideIndex - 1 : failureCount + FailureTemplateSlideIndex;
        var trendSlideIndex = tableSlideIndex + 1;

        ApplyMetricTextReplacements(presentationPart, data, CoverSlideIndex, SummarySlideIndex);
        ApplyFailureTextReplacements(presentationPart, fails, FailureTemplateSlideIndex);

        UpdatePieChartOnSlide(presentationPart, SummarySlideIndex, data.Pass, data.Fail);
        UpdateLineChartOnSlide(presentationPart, trendSlideIndex, data.Trends);
        BuildTestCaseTable(presentationPart, tableSlideIndex, data.TestCases);

        var logoBytes = ResolveLogoBytes(data);
        if (logoBytes is { Length: > 0 })
            InjectCompanyLogoOnAllButCover(presentationPart, logoBytes);

        presentation.Save();
        buffer.Position = 0;
        return buffer.ToArray();
    }

    private static void ValidateTemplate(PresentationPart presentationPart)
    {
        var slideIds = presentationPart.Presentation?.SlideIdList?.Elements<SlideId>().ToList()
            ?? throw new InvalidOperationException("Presentation has no SlideIdList.");

        if (slideIds.Count < MinimumTemplateSlideCount)
        {
            throw new InvalidOperationException(
                $"Report template must contain at least {MinimumTemplateSlideCount} slides. Actual: {slideIds.Count}.");
        }

        var summarySlide = GetSlidePart(presentationPart, slideIds, SummarySlideIndex);
        if (!summarySlide.GetPartsOfType<ChartPart>().Any(part => part.ChartSpace.Descendants<PieChart>().Any()))
            throw new InvalidOperationException("Report template summary slide is missing a pie chart.");

        var failureSlide = GetSlidePart(presentationPart, slideIds, FailureTemplateSlideIndex);
        var failureTokens = new[] { "{{NAME}}", "{{RESULT}}", "{{DESC}}" };
        var failureText = string.Concat(failureSlide.Slide.Descendants<A.Text>().Select(t => t.Text));
        foreach (var token in failureTokens)
        {
            if (!failureText.Contains(token, StringComparison.Ordinal))
                throw new InvalidOperationException($"Report template failure slide is missing placeholder '{token}'.");
        }

        var tableSlide = GetSlidePart(presentationPart, slideIds, OriginalTestListSlideIndex);
        if (!tableSlide.Slide.Descendants<A.Table>().Any())
            throw new InvalidOperationException("Report template test list slide is missing a table.");

        var trendSlide = GetSlidePart(presentationPart, slideIds, OriginalTrendSlideIndex);
        if (!trendSlide.GetPartsOfType<ChartPart>().Any(part => part.ChartSpace.Descendants<LineChart>().Any()))
            throw new InvalidOperationException("Report template trend slide is missing a line chart.");
    }

    private static SlidePart GetSlidePart(PresentationPart presentationPart, IReadOnlyList<SlideId> slideIds, int slideIndex)
    {
        if (slideIndex < 0 || slideIndex >= slideIds.Count)
            throw new InvalidOperationException($"Report template slide index {slideIndex} is out of range.");

        return (SlidePart)presentationPart.GetPartById(slideIds[slideIndex].RelationshipId!);
    }

    private static void AdjustFailureSlides(PresentationPart presentationPart, byte[] pristineTemplate, IReadOnlyList<TestCaseResult> fails)
    {
        var slideIdList = presentationPart.Presentation!.SlideIdList
            ?? throw new InvalidOperationException("Presentation has no SlideIdList.");

        var orderedSlideIds = slideIdList.Elements<SlideId>().ToList();
        var failureTemplateSlideId = orderedSlideIds[FailureTemplateSlideIndex];
        var failureTemplatePart = (SlidePart)presentationPart.GetPartById(failureTemplateSlideId.RelationshipId!);

        if (fails.Count == 0)
        {
            failureTemplateSlideId.Remove();
            presentationPart.DeletePart(failureTemplatePart);
            return;
        }

        var anchor = failureTemplateSlideId;
        for (var i = 1; i < fails.Count; i++)
        {
            using var ms = new MemoryStream(pristineTemplate, writable: false);
            using var sourceDoc = PresentationDocument.Open(ms, false);
            var sourcePresentationPart = sourceDoc.PresentationPart!;
            var sourceFailureSlideId = sourcePresentationPart.Presentation!.SlideIdList!
                .Elements<SlideId>().ElementAt(FailureTemplateSlideIndex);
            var sourceFailurePart = (SlidePart)sourcePresentationPart.GetPartById(sourceFailureSlideId.RelationshipId!);

            var imported = (SlidePart)presentationPart.AddPart(sourceFailurePart);
            var newSlideId = new SlideId
            {
                Id = NextSlideId(slideIdList),
                RelationshipId = presentationPart.GetIdOfPart(imported)
            };
            slideIdList.InsertAfter(newSlideId, anchor);
            anchor = newSlideId;
        }
    }

    private static uint NextSlideId(SlideIdList slideIdList)
    {
        var max = slideIdList.Elements<SlideId>().Select(s => s.Id?.Value ?? 0u).DefaultIfEmpty(255u).Max();
        return max + 1;
    }

    private byte[]? ResolveLogoBytes(ReportData data)
    {
        if (data.CompanyLogo is { Length: > 0 })
            return data.CompanyLogo;

        return File.Exists(_fallbackLogoPath) ? File.ReadAllBytes(_fallbackLogoPath) : null;
    }

    private static void ApplyMetricTextReplacements(PresentationPart presentationPart, ReportData data, params int[] slideIndexes)
    {
        var map = BuildMetricReplacementMap(data);
        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
        foreach (var index in slideIndexes)
        {
            if (index < 0 || index >= slideIds.Count)
                continue;

            var part = (SlidePart)presentationPart.GetPartById(slideIds[index].RelationshipId!);
            ReplaceTextInElement(part.Slide, map);
        }
    }

    private static void ApplyFailureTextReplacements(PresentationPart presentationPart, IReadOnlyList<TestCaseResult> fails, int firstFailureSlideIndex)
    {
        if (fails.Count == 0)
            return;

        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
        for (var i = 0; i < fails.Count; i++)
        {
            var slideIndex = firstFailureSlideIndex + i;
            if (slideIndex >= slideIds.Count)
                break;

            var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[slideIndex].RelationshipId!);
            var description = SafeText(fails[i].Description);
            var map = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["{{NAME}}"] = SafeText(fails[i].Name),
                ["{{RESULT}}"] = SafeText(fails[i].Result),
                ["{{DESC}}"] = description,
                ["상세 내용 / 오류 메시지 /\n로그\n(선택 영역)"] = description,
                ["상세 내용 / 오류 메시지 / 로그 (선택 영역)"] = description,
            };
            ReplaceTextInElement(slidePart.Slide, map);
            BindFailureScreenshot(slidePart, fails[i].Screenshot);
        }
    }

    private static string SafeText(string? value) => value ?? string.Empty;

    private static Dictionary<string, string> BuildMetricReplacementMap(ReportData data)
    {
        var rateText = data.Rate.ToString("0.##", CultureInfo.InvariantCulture);
        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["{{TITLE}}"] = SafeText(data.Title),
            ["{{DATE}}"] = SafeText(data.Date),
            ["{{TOTAL}}"] = data.Total.ToString(CultureInfo.InvariantCulture),
            ["{{PASS}}"] = data.Pass.ToString(CultureInfo.InvariantCulture),
            ["{{FAIL}}"] = data.Fail.ToString(CultureInfo.InvariantCulture),
            ["{{RATE}}"] = rateText,
        };
    }

    private static void ReplaceTextInElement(OpenXmlElement root, IReadOnlyDictionary<string, string> replacements)
    {
        foreach (var text in root.Descendants<A.Text>())
        {
            var original = text.Text;
            if (string.IsNullOrEmpty(original))
                continue;

            var updated = original;
            foreach (var (token, value) in replacements)
                updated = updated.Replace(token, value, StringComparison.Ordinal);

            if (!string.Equals(updated, original, StringComparison.Ordinal))
                text.Text = updated;
        }
    }

    private static void BindFailureScreenshot(SlidePart slidePart, byte[] screenshot)
    {
        var maxCx = DesignGuide.FailureImageMaxWidthEmu;
        var maxCy = DesignGuide.FailureImageMaxHeightEmu;
        var originX = DesignGuide.FailureImageOffsetXEmu;
        var originY = DesignGuide.FailureImageOffsetYEmu;

        if (screenshot is not { Length: > 0 })
            return;

        Image<Rgba32> image;
        try
        {
            image = Image.Load<Rgba32>(screenshot);
        }
        catch (Exception)
        {
            return;
        }

        using (image)
        {
            var (cx, cy) = MeasureContainedExtentEmu(image.Width, image.Height, maxCx, maxCy);
            var offsetX = originX + (maxCx - cx) / 2;
            var offsetY = originY + (maxCy - cy) / 2;

            var targetW = Math.Max(1, (int)Math.Round(EmuToPixel(cx)));
            var targetH = Math.Max(1, (int)Math.Round(EmuToPixel(cy)));
            image.Mutate(ctx => ctx.Resize(targetW, targetH));

            using var pngStream = new MemoryStream();
            image.SaveAsPng(pngStream);
            pngStream.Position = 0;

            var picture = FindFailureScreenshotPicture(slidePart);
            var blip = picture?.Descendants<A.Blip>().FirstOrDefault();
            if (picture is not null && blip?.Embed?.Value is string embedId)
            {
                var imagePart = (ImagePart)slidePart.GetPartById(embedId);
                imagePart.FeedData(pngStream);
                ApplyPictureTransform(picture, offsetX, offsetY, cx, cy);
                return;
            }

            var newImagePart = slidePart.AddImagePart(ImagePartType.Png);
            newImagePart.FeedData(pngStream);

            var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;
            var newPicture = BuildPictureShape(
                tree,
                slidePart.GetIdOfPart(newImagePart),
                FailureScreenshotPictureName,
                offsetX,
                offsetY,
                cx,
                cy);
            tree.Append(newPicture);
        }
    }

    private static P.Picture? FindFailureScreenshotPicture(SlidePart slidePart)
    {
        var pictures = slidePart.Slide.Descendants<P.Picture>().ToList();
        return pictures.FirstOrDefault(pic =>
                string.Equals(
                    pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value,
                    FailureScreenshotPictureName,
                    StringComparison.Ordinal))
            ?? pictures.FirstOrDefault(pic =>
                !string.Equals(
                    pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value,
                    "TAD_CompanyLogo",
                    StringComparison.Ordinal));
    }

    private static void ApplyPictureTransform(P.Picture picture, long offsetX, long offsetY, long cx, long cy)
    {
        var spPr = picture.GetFirstChild<P.ShapeProperties>()
            ?? picture.AppendChild(new P.ShapeProperties());

        var xfrm = spPr.GetFirstChild<A.Transform2D>();
        xfrm?.Remove();

        xfrm = new A.Transform2D(
            new A.Offset { X = offsetX, Y = offsetY },
            new A.Extents { Cx = cx, Cy = cy });
        spPr.InsertAt(xfrm, 0);
    }

    private static (long cx, long cy) MeasureContainedExtentEmu(int pixelWidth, int pixelHeight, long maxCx, long maxCy)
    {
        var imgCx = PixelToEmu(pixelWidth);
        var imgCy = PixelToEmu(pixelHeight);
        var scale = Math.Min(maxCx / (double)imgCx, maxCy / (double)imgCy);
        return ((long)(imgCx * scale), (long)(imgCy * scale));
    }

    private static long PixelToEmu(int px) => (long)Math.Round(px * DesignGuide.EmuPerInch / 96.0);

    private static double EmuToPixel(long emu) => emu * 96.0 / DesignGuide.EmuPerInch;

    private static void InjectCompanyLogoOnAllButCover(PresentationPart presentationPart, byte[] logoPngBytes)
    {
        Image<Rgba32> image;
        try
        {
            image = Image.Load<Rgba32>(logoPngBytes);
        }
        catch (Exception)
        {
            return;
        }

        using var imageScope = image;
        var maxCx = DesignGuide.EmuFromInch(DesignGuide.CompanyLogoWidthInches);
        var (cx, cy) = MeasureContainedExtentEmu(imageScope.Width, imageScope.Height, maxCx, long.MaxValue);
        var targetW = Math.Max(1, (int)Math.Round(EmuToPixel(cx)));
        var targetH = Math.Max(1, (int)Math.Round(EmuToPixel(cy)));
        imageScope.Mutate(ctx => ctx.Resize(targetW, targetH));

        using var png = new MemoryStream();
        imageScope.SaveAsPng(png);
        var logoBytes = png.ToArray();

        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
        for (var i = 1; i < slideIds.Count; i++)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[i].RelationshipId!);
            InjectCompanyLogo(slidePart, logoBytes, cx, cy);
        }
    }

    private static void InjectCompanyLogo(SlidePart slidePart, ReadOnlySpan<byte> pngBytes, long cx, long cy)
    {
        var tree = slidePart.Slide.CommonSlideData!.ShapeTree!;
        var existing = tree.Descendants<P.Picture>()
            .FirstOrDefault(pic => pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value == "TAD_CompanyLogo");

        if (existing is not null)
        {
            var blip = existing.Descendants<A.Blip>().FirstOrDefault();
            if (blip?.Embed?.Value is string embedId)
            {
                using var ms = new MemoryStream();
                ms.Write(pngBytes);
                ms.Position = 0;
                var imagePart = (ImagePart)slidePart.GetPartById(embedId);
                imagePart.FeedData(ms);
                ApplyPictureTransform(existing, DesignGuide.CompanyLogoOffsetXEmu, DesignGuide.CompanyLogoOffsetYEmu, cx, cy);
            }

            return;
        }

        var imagePartNew = slidePart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream())
        {
            ms.Write(pngBytes);
            ms.Position = 0;
            imagePartNew.FeedData(ms);
        }

        var relId = slidePart.GetIdOfPart(imagePartNew);
        var picture = BuildPictureShape(tree, relId, "TAD_CompanyLogo", DesignGuide.CompanyLogoOffsetXEmu, DesignGuide.CompanyLogoOffsetYEmu, cx, cy);
        tree.Append(picture);
    }

    private static P.Picture BuildPictureShape(OpenXmlElement shapeTree, string embedRelationshipId, string shapeName, long x, long y, long cx, long cy)
    {
        return new P.Picture(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = NextShapeTreeId(shapeTree), Name = shapeName },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = embedRelationshipId, CompressionState = A.BlipCompressionValues.Print },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
    }

    private static uint NextShapeTreeId(OpenXmlElement shapeTree)
    {
        var max = shapeTree.Descendants<P.NonVisualDrawingProperties>()
            .Select(n => n.Id?.Value ?? 0u)
            .DefaultIfEmpty(0u)
            .Max();
        return max + 1;
    }

    private static void BuildTestCaseTable(PresentationPart presentationPart, int tableSlideIndex, IReadOnlyList<TestCaseResult> testCases)
    {
        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
        if (tableSlideIndex < 0 || tableSlideIndex >= slideIds.Count)
            return;

        var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[tableSlideIndex].RelationshipId!);
        var table = slidePart.Slide.Descendants<A.Table>().FirstOrDefault();
        if (table is null)
            return;

        StyleTableHeader(table);

        var rows = table.Elements<A.TableRow>().ToList();
        if (rows.Count == 0)
            return;

        A.TableRow rowTemplate;
        if (rows.Count > 1)
        {
            rowTemplate = (A.TableRow)rows[1].CloneNode(true);
            foreach (var extra in rows.Skip(1).ToList())
                extra.Remove();
        }
        else
        {
            rowTemplate = CreateSyntheticDataRow();
        }

        foreach (var tc in testCases)
        {
            var newRow = (A.TableRow)rowTemplate.CloneNode(true);
            SetRowTexts(newRow, tc);
            StyleResultCell(newRow, tc.Result);
            table.AppendChild(newRow);
        }
    }

    private static A.TableRow CreateSyntheticDataRow()
    {
        var row = new A.TableRow();
        for (var c = 0; c < 5; c++)
        {
            row.Append(new A.TableCell(
                new A.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(string.Empty))))));
        }

        return row;
    }

    private static void StyleTableHeader(A.Table table)
    {
        var headerRow = table.Elements<A.TableRow>().FirstOrDefault();
        if (headerRow is null)
            return;

        foreach (var cell in headerRow.Elements<A.TableCell>())
        {
            cell.TableCellProperties ??= new A.TableCellProperties();
            cell.TableCellProperties.RemoveAllChildren<A.SolidFill>();
            cell.TableCellProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = DesignGuide.MainCorporateColorHex }));

            foreach (var run in cell.Descendants<A.Run>())
            {
                var rPr = run.RunProperties ?? run.InsertAt(new A.RunProperties(), 0);
                rPr.RemoveAllChildren<A.SolidFill>();
                rPr.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFFFF" }));
                rPr.Bold = true;
                rPr.Append(new A.LatinFont { Typeface = DesignGuide.PrimaryFont });
                rPr.Append(new A.EastAsianFont { Typeface = DesignGuide.PrimaryFont });
            }
        }
    }

    private static void SetRowTexts(A.TableRow row, TestCaseResult testCase)
    {
        var cells = row.Elements<A.TableCell>().ToList();
        var values = new[]
        {
            testCase.No.ToString(CultureInfo.InvariantCulture),
            SafeText(testCase.Name),
            SafeText(testCase.Result),
            string.Empty, // Execution time not present on domain model
            SafeText(testCase.Remarks),
        };

        for (var i = 0; i < cells.Count && i < values.Length; i++)
            ReplaceCellText(cells[i], values[i]);
    }

    private static void ReplaceCellText(A.TableCell cell, string text)
    {
        var body = cell.TextBody ?? cell.AppendChild(new A.TextBody());
        body.RemoveAllChildren();

        var paragraph = new A.Paragraph(new A.Run(new A.Text(text)));
        body.Append(paragraph);
    }

    private static void StyleResultCell(A.TableRow row, string result)
    {
        var cells = row.Elements<A.TableCell>().ToList();
        if (cells.Count < 3)
            return;

        var resultCell = cells[2];
        var isPass = string.Equals(result, "PASS", StringComparison.OrdinalIgnoreCase);
        var isFail = string.Equals(result, "FAIL", StringComparison.OrdinalIgnoreCase);
        if (!isPass && !isFail)
            return;

        var color = isPass ? DesignGuide.PassColorHex : DesignGuide.FailColorHex;
        foreach (var run in resultCell.Descendants<A.Run>())
        {
            var rPr = run.RunProperties ?? run.InsertAt(new A.RunProperties(), 0);
            rPr.Bold = true;
            rPr.RemoveAllChildren<A.SolidFill>();
            rPr.Append(new A.SolidFill(new A.RgbColorModelHex { Val = color }));
            rPr.Append(new A.LatinFont { Typeface = DesignGuide.PrimaryFont });
            rPr.Append(new A.EastAsianFont { Typeface = DesignGuide.PrimaryFont });
        }
    }

    private static void UpdatePieChartOnSlide(PresentationPart presentationPart, int slideIndex, int pass, int fail)
    {
        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
        if (slideIndex < 0 || slideIndex >= slideIds.Count)
            return;

        var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[slideIndex].RelationshipId!);
        foreach (var chartPart in slidePart.GetPartsOfType<ChartPart>())
        {
            var pie = chartPart.ChartSpace.Descendants<PieChart>().FirstOrDefault();
            if (pie is null)
                continue;

            var series = pie.Descendants<PieChartSeries>().FirstOrDefault();
            if (series is null)
                continue;

            var values = series.GetFirstChild<Values>() ?? series.AppendChild(new Values());
            values.RemoveAllChildren();

            var numLit = new NumberLiteral(
                new FormatCode("General"),
                new PointCount { Val = 2 },
                new NumericPoint(new NumericValue(pass.ToString(CultureInfo.InvariantCulture))) { Index = 0 },
                new NumericPoint(new NumericValue(fail.ToString(CultureInfo.InvariantCulture))) { Index = 1 });
            values.AppendChild(numLit);

            ApplySeriesRgb(series, 0, DesignGuide.PassColorHex);
            ApplySeriesRgb(series, 1, DesignGuide.FailColorHex);
            break;
        }
    }

    private static void ApplySeriesRgb(OpenXmlCompositeElement series, int pointIndex, string rgbHex)
    {
        var dPt = series.Descendants<DataPoint>().FirstOrDefault(dp =>
            dp.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Index>()?.Val?.Value == (uint)pointIndex);
        if (dPt is null)
        {
            dPt = new DataPoint(new ChartShapeProperties(new A.SolidFill(new A.RgbColorModelHex { Val = rgbHex })))
            {
                Index = new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = (uint)pointIndex }
            };
            series.AppendChild(dPt);
            return;
        }

        var spPr = dPt.GetFirstChild<ChartShapeProperties>() ?? dPt.InsertAt(new ChartShapeProperties(), 0);
        spPr.RemoveAllChildren<A.SolidFill>();
        spPr.AppendChild(new A.SolidFill(new A.RgbColorModelHex { Val = rgbHex }));
    }

    private static void UpdateLineChartOnSlide(PresentationPart presentationPart, int slideIndex, IReadOnlyList<DailyTrend> trends)
    {
        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<SlideId>().ToList();
        if (slideIndex < 0 || slideIndex >= slideIds.Count)
            return;

        var slidePart = (SlidePart)presentationPart.GetPartById(slideIds[slideIndex].RelationshipId!);
        foreach (var chartPart in slidePart.GetPartsOfType<ChartPart>())
        {
            var line = chartPart.ChartSpace.Descendants<LineChart>().FirstOrDefault();
            if (line is null)
                continue;

            var seriesList = line.Elements<LineChartSeries>().ToList();
            if (seriesList.Count < 2)
                continue;

            WriteStringLiteral(seriesList[0].GetFirstChild<CategoryAxisData>() ?? seriesList[0].AppendChild(new CategoryAxisData()),
                trends.Select(t => SafeText(t.Date)).ToList());

            WriteNumberLiteral(seriesList[0].GetFirstChild<Values>() ?? seriesList[0].AppendChild(new Values()),
                trends.Select(t => t.PassCount).ToList());

            WriteStringLiteral(seriesList[1].GetFirstChild<CategoryAxisData>() ?? seriesList[1].AppendChild(new CategoryAxisData()),
                trends.Select(t => SafeText(t.Date)).ToList());

            WriteNumberLiteral(seriesList[1].GetFirstChild<Values>() ?? seriesList[1].AppendChild(new Values()),
                trends.Select(t => t.FailCount).ToList());

            ApplyLineSeriesColor(seriesList[0], DesignGuide.PassColorHex);
            ApplyLineSeriesColor(seriesList[1], DesignGuide.FailColorHex);
            break;
        }
    }

    private static void WriteStringLiteral(OpenXmlCompositeElement parent, IReadOnlyList<string> points)
    {
        parent.RemoveAllChildren();
        var literal = new StringLiteral(new PointCount { Val = (uint)points.Count });
        for (var i = 0; i < points.Count; i++)
            literal.AppendChild(new StringPoint(new NumericValue(points[i])) { Index = (uint)i });

        parent.AppendChild(literal);
    }

    private static void WriteNumberLiteral(Values values, IReadOnlyList<int> points)
    {
        values.RemoveAllChildren();
        var literal = new NumberLiteral(
            new FormatCode("General"),
            new PointCount { Val = (uint)points.Count });
        for (var i = 0; i < points.Count; i++)
            literal.AppendChild(new NumericPoint(new NumericValue(points[i].ToString(CultureInfo.InvariantCulture))) { Index = (uint)i });

        values.AppendChild(literal);
    }

    private static void ApplyLineSeriesColor(LineChartSeries series, string rgbHex)
    {
        var spPr = series.GetFirstChild<ChartShapeProperties>() ?? series.InsertAt(new ChartShapeProperties(), 0);
        spPr.RemoveAllChildren<A.SolidFill>();
        spPr.AppendChild(new A.SolidFill(new A.RgbColorModelHex { Val = rgbHex }));
    }
}

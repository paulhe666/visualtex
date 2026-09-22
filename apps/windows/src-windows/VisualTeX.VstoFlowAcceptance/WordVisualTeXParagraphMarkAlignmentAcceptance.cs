using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class ScreenRectMetric
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
    }

    private sealed class InkMetric
    {
        public int Count { get; set; }
        public int MinX { get; set; }
        public int MaxX { get; set; }
        public int MinY { get; set; }
        public int MaxY { get; set; }
        public double CenterY { get; set; }
        public double MeanY { get; set; }
        public int Height { get; set; }
    }

    private static void RunWordVisualTeXParagraphMarkAlignmentAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The paragraph-mark alignment acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"paragraph-mark-alignment-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "display-formula.svg");
        var pngPath = Path.Combine(assetRoot, "display-formula.png");
        var fixtureSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"320\" height=\"100\" viewBox=\"0 0 320 100\"><text x=\"118\" y=\"27\" font-family=\"Cambria Math\" font-size=\"28\">a+b</text><line x1=\"96\" y1=\"50\" x2=\"224\" y2=\"50\" stroke=\"#111111\" stroke-width=\"4\"/><text x=\"118\" y=\"91\" font-family=\"Cambria Math\" font-size=\"28\">c+d</text></svg>";
        var fixtureLatex =
            @"\int_0^\infty e^{-x^2}\,dx=\frac{\sqrt{\pi}}{2}";
        var fixtureMathMl = string.Empty;
        var exportWidthPx = 320d;
        var exportHeightPx = 100d;
        var exportBaselinePx = double.TryParse(
            Environment.GetEnvironmentVariable("VISUALTEX_DISPLAY_MARK_BASELINE_PX"),
            out var requestedBaselinePx)
            ? requestedBaselinePx
            : 72d;
        var productionFixturePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_DISPLAY_MARK_EXPORT_FIXTURE");
        if (!string.IsNullOrWhiteSpace(productionFixturePath))
        {
            var resolvedFixturePath = Path.GetFullPath(
                Environment.ExpandEnvironmentVariables(
                    productionFixturePath!.Trim().Trim('"')));
            using var fixtureDocument = JsonDocument.Parse(
                File.ReadAllText(resolvedFixturePath));
            var root = fixtureDocument.RootElement;
            fixtureSvg = root.GetProperty("svg").GetString()
                ?? throw new InvalidDataException(
                    "The production display fixture has no SVG.");
            exportWidthPx = root.GetProperty("width").GetDouble();
            exportHeightPx = root.GetProperty("height").GetDouble();
            exportBaselinePx = root.GetProperty("baseline").GetDouble();
            if (root.TryGetProperty("mathMl", out var mathMlElement))
                fixtureMathMl = mathMlElement.GetString() ?? string.Empty;
            fixtureLatex = @"\oiint_{\Sigma}F\,\mathrm{d}S11";
        }
        File.WriteAllText(
            svgPath,
            fixtureSvg,
            new UTF8Encoding(false));
        var emfWidth = Math.Max(1, (int)Math.Ceiling(exportWidthPx));
        var emfHeight = Math.Max(1, (int)Math.Ceiling(exportHeightPx));
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
            svgPath,
            emfWidth,
            emfHeight);
        var pngDataUrl = CreatePngDataUrl(
            "paragraph-mark-alignment",
            emfWidth,
            emfHeight);
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(
                pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Window? window = null;
        Word.View? view = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? paragraphMark = null;
        Word.Font? paragraphMarkFont = null;
        Word.Range? objectResult = null;
        Word.Font? objectResultFont = null;
        System.Drawing.Bitmap? screenshot = null;
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Add(Visible: true);
            var documentPath = Path.Combine(
                artifactRoot,
                "VisualTeX-Unnumbered-Paragraph-Mark-Alignment.docx");
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Content.Text =
                "VisualTeX paragraph-mark alignment acceptance\r"
                + "Above formula\r\r"
                + "Below formula\r";
            document.Save();
            document.Activate();

            window = application.ActiveWindow;
            var wordWindowHandle = new IntPtr(window.Hwnd);
            _ = SetWindowPos(
                wordWindowHandle,
                IntPtr.Zero,
                80,
                60,
                1320,
                900,
                0x0040);
            _ = SetForegroundWindow(wordWindowHandle);
            view = window.View;
            view.ShowAll = true;

            var service = new WordFormulaService(application);
            var insertion = document.Range(
                "VisualTeX paragraph-mark alignment acceptance\rAbove formula\r".Length,
                "VisualTeX paragraph-mark alignment acceptance\rAbove formula\r".Length);
            string formulaId;
            try
            {
                application.Selection.SetRange(
                    insertion.Start,
                    insertion.End);
                formulaId = Guid.NewGuid().ToString("D");
                var session = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: fixtureLatex);
                // Exercise the user-visible path that can differ from a fresh
                // unnumbered insert: create the managed numbered display host,
                // then remove numbering through the real ReplaceOle path.
                session.Numbered = true;
                session.DisplayMode = "block";
                session.ExportResult = new OfficeExportDocument
                {
                    Width = (float)exportWidthPx,
                    Height = (float)exportHeightPx,
                    Baseline = (float)exportBaselinePx,
                    MathMl = fixtureMathMl,
                };
                service.InsertOle(session, pngPath, emfPath);
            }
            finally { Release(insertion); }

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(
                document,
                formulaId);
            var numberedMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException(
                    "Paragraph-mark alignment numbered fixture lost VisualTeX metadata.");
            AssertTrue(numberedMetadata.Numbered,
                "Paragraph-mark alignment setup did not create a numbered display formula.");
            shapeRange = shape.Range.Duplicate;
            var unnumberSession = CreateNumberedPerformanceSession(
                "edit",
                formulaId,
                document.FullName,
                WordRangeReference(shapeRange.Start, shapeRange.End),
                numberedMetadata,
                latex: fixtureLatex);
            unnumberSession.Numbered = false;
            unnumberSession.DisplayMode = "block";
            unnumberSession.ExportResult = new OfficeExportDocument
            {
                Width = (float)exportWidthPx,
                Height = (float)exportHeightPx,
                Baseline = (float)exportBaselinePx,
                MathMl = fixtureMathMl,
            };
            service.ReplaceOle(unnumberSession, pngPath, emfPath);
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(
                document,
                formulaId);
            var metadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException(
                    "Paragraph-mark alignment unnumbered result lost VisualTeX metadata.");
            AssertTrue(!metadata.Numbered,
                "Numbered-to-unnumbered edit did not persist Numbered=false.");
            AssertEqual(
                "block",
                metadata.DisplayMode,
                "Paragraph-mark alignment fixture is not display math.");

            // Measure the persisted Word layout rather than the transient OLE
            // selection frame left by ReplaceOle. Reopen the isolated fixture so
            // GetPoint and the screenshot see the same settled document state a
            // user sees after normal editing/save.
            document.Save();
            Release(shape); shape = null;
            Release(view); view = null;
            Release(window); window = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                Visible: true);
            document.Activate();
            window = application.ActiveWindow;
            wordWindowHandle = new IntPtr(window.Hwnd);
            _ = SetWindowPos(
                wordWindowHandle,
                IntPtr.Zero,
                80,
                60,
                1320,
                900,
                0x0040);
            _ = SetForegroundWindow(wordWindowHandle);
            view = window.View;
            view.ShowAll = true;
            shape = FindVisualTeXOleByFormulaIdForNumberToggle(
                document,
                formulaId);
            metadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException(
                    "Saved/reopened paragraph-mark alignment result lost VisualTeX metadata.");

            shapeRange = shape.Range.Duplicate;
            paragraphs = shapeRange.Paragraphs;
            AssertEqual(
                1,
                paragraphs.Count,
                "Paragraph-mark alignment fixture does not own exactly one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            paragraphMark = document.Range(
                paragraphRange.End - 1,
                paragraphRange.End);
            AssertEqual(
                "\r",
                paragraphMark.Text,
                "Paragraph-mark alignment fixture does not end in an ordinary paragraph mark.");
            paragraphMarkFont = paragraphMark.Font;

            for (var position = shapeRange.Start;
                 position < shapeRange.End;
                 position++)
            {
                Release(objectResultFont); objectResultFont = null;
                Release(objectResult); objectResult = null;
                objectResult = document.Range(position, position + 1);
                if (!string.Equals(
                        objectResult.Text,
                        "\u0001",
                        StringComparison.Ordinal))
                    continue;
                objectResultFont = objectResult.Font;
                break;
            }
            if (objectResultFont is null)
                throw new InvalidDataException(
                    "Paragraph-mark alignment fixture has no U+0001 VisualTeX object-result character.");

            document.Repaginate();
            window.ScrollIntoView(paragraphRange, true);
            // ReplaceOle can rebuild Word's view state; force formatting marks
            // on immediately before measurement instead of relying on the value
            // set before the edit transaction.
            view.ShowAll = true;
            // ReplaceOle can leave the embedded object selected. Its selection
            // frame is dark and spans the complete OLE rectangle, which would
            // contaminate the pixel measurement. Move the caret to the next
            // paragraph without changing document content before capturing.
            application.Selection.SetRange(
                paragraphRange.End,
                paragraphRange.End);
            try { application.ScreenRefresh(); } catch { }
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(180);

            var shapeScreen = ReadAcceptanceScreenRect(window, objectResult);
            var paragraphMarkScreen = ReadAcceptanceScreenRect(
                window,
                paragraphMark);
            var captureLeft = Math.Max(
                0,
                Math.Min(shapeScreen.X, paragraphMarkScreen.X) - 36);
            var captureTop = Math.Max(
                0,
                Math.Min(shapeScreen.Y, paragraphMarkScreen.Y) - 40);
            var captureRight = Math.Max(
                shapeScreen.X + Math.Max(1, shapeScreen.Width),
                paragraphMarkScreen.X + Math.Max(40, paragraphMarkScreen.Width + 28))
                + 36;
            var captureBottom = Math.Max(
                shapeScreen.Y + Math.Max(1, shapeScreen.Height),
                paragraphMarkScreen.Y + Math.Max(40, paragraphMarkScreen.Height + 28))
                + 40;
            var captureWidth = Math.Max(120, captureRight - captureLeft);
            var captureHeight = Math.Max(120, captureBottom - captureTop);
            screenshot = new System.Drawing.Bitmap(
                captureWidth,
                captureHeight,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (var graphics = System.Drawing.Graphics.FromImage(screenshot))
            {
                graphics.CopyFromScreen(
                    captureLeft,
                    captureTop,
                    0,
                    0,
                    new System.Drawing.Size(
                        captureWidth,
                        captureHeight),
                    System.Drawing.CopyPixelOperation.SourceCopy);
            }
            var screenshotPath = Path.Combine(
                artifactRoot,
                "visualtex-unnumbered-paragraph-mark.png");
            screenshot.Save(
                screenshotPath,
                System.Drawing.Imaging.ImageFormat.Png);

            var formulaInkRect = new ScreenRectMetric
            {
                X = Math.Max(0, shapeScreen.X - captureLeft),
                Y = Math.Max(0, shapeScreen.Y - captureTop),
                Width = Math.Max(1, shapeScreen.Width),
                Height = Math.Max(1, shapeScreen.Height),
            };
            var markInkRect = new ScreenRectMetric
            {
                // Measure only Word's own paragraph-mark rectangle. Expanding
                // left/right or above/below can mix neighbouring formula ink into
                // this ROI and turn a visible displacement into a false pass.
                X = Math.Max(
                    0,
                    paragraphMarkScreen.X - captureLeft),
                Y = Math.Max(
                    0,
                    paragraphMarkScreen.Y - captureTop),
                Width = Math.Max(
                    1,
                    paragraphMarkScreen.Width),
                Height = Math.Max(
                    1,
                    paragraphMarkScreen.Height),
            };
            if (markInkRect.X + markInkRect.Width > captureWidth)
                markInkRect.Width = captureWidth - markInkRect.X;
            if (markInkRect.Y + markInkRect.Height > captureHeight)
                markInkRect.Height = captureHeight - markInkRect.Y;

            var formulaInk = MeasureAcceptanceInk(
                screenshot,
                formulaInkRect);
            var paragraphMarkInk = MeasureAcceptanceInk(
                screenshot,
                markInkRect);
            AssertTrue(
                formulaInk.Count > 0,
                "The screenshot contains no measurable formula ink.");
            AssertTrue(
                paragraphMarkInk.Count > 0,
                "The screenshot contains no measurable paragraph-mark ink. ShowAll may not be enabled.");
            var verticalCenterDeltaPx =
                paragraphMarkInk.CenterY - formulaInk.CenterY;

            var semanticFontSize =
                FormulaFontSize.ResolveSemanticFontSize(metadata);
            AssertTrue(
                metadata.WordDisplayPreviewInkHeightRatio.HasValue
                && metadata.WordDisplayPreviewBottomWhitespaceRatio.HasValue,
                "Paragraph-mark alignment fixture lost its display preview ink metrics.");
            var expectedObjectPosition =
                WordInlineAlignment.CalculateDisplayInkCenterPosition(
                    shape.Height,
                    (float)metadata.WordDisplayPreviewInkHeightRatio!.Value,
                    (float)metadata.WordDisplayPreviewBottomWhitespaceRatio!.Value,
                    paragraphMarkFont.Size);
            var measurement = new
            {
                formulaId,
                screenshot = screenshotPath,
                capture = new
                {
                    left = captureLeft,
                    top = captureTop,
                    width = captureWidth,
                    height = captureHeight,
                },
                shapeRange = new
                {
                    start = shapeRange.Start,
                    end = shapeRange.End,
                    objectResultStart = objectResult.Start,
                    objectResultEnd = objectResult.End,
                    shapeWidthPt = shape.Width,
                    shapeHeightPt = shape.Height,
                    objectResultFontPositionPt =
                        objectResultFont.Position,
                    expectedObjectPositionPt =
                        expectedObjectPosition,
                    screen = shapeScreen,
                },
                paragraphMark = new
                {
                    start = paragraphMark.Start,
                    end = paragraphMark.End,
                    text = paragraphMark.Text,
                    fontPositionPt = paragraphMarkFont.Position,
                    screen = paragraphMarkScreen,
                },
                export = new
                {
                    renderWidthPx = metadata.RenderWidthPx,
                    renderHeightPx = metadata.RenderHeightPx,
                    baselinePx = metadata.Baseline,
                    semanticFontSizePt = semanticFontSize,
                },
                formulaInk,
                paragraphMarkInk,
                verticalCenterDeltaPx,
            };
            File.WriteAllText(
                Path.Combine(
                    artifactRoot,
                    "visualtex-unnumbered-paragraph-mark-measurement.json"),
                JsonSerializer.Serialize(
                    measurement,
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,
                    }),
                new UTF8Encoding(false));
            document.Save();

            // This in-process capture remains diagnostic because Word can report
            // an inflated EMBED-field rectangle while the OLE server is still
            // attached to this automation process. The release acceptance uses
            // capture-word-formula-baseline.ps1 after reopen for the authoritative
            // pixel-center assertion. Here require the persisted structural rule.
            AssertNear(
                expectedObjectPosition,
                objectResultFont.Position,
                0.1f,
                "The unnumbered VisualTeX display OLE did not retain the preview-ink-center position.");
            AssertNear(
                0f,
                paragraphMarkFont.Position,
                0.1f,
                "The paragraph mark acquired a manual vertical Font.Position offset.");

            Console.WriteLine(
                $"[VISUALTEX PARAGRAPH MARK] formulaCenterY={formulaInk.CenterY:F2}px paragraphMarkCenterY={paragraphMarkInk.CenterY:F2}px delta={verticalCenterDeltaPx:F2}px objectPosition={objectResultFont.Position:F2}pt paragraphMarkPosition={paragraphMarkFont.Position:F2}pt.");
        }
        finally
        {
            screenshot?.Dispose();
            Release(objectResultFont);
            Release(objectResult);
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
            Release(view);
            Release(window);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); }
            catch { }
            Release(application);
            try { Directory.Delete(assetRoot, recursive: true); }
            catch { }
            ForceComCleanup();
        }
    }

    private static ScreenRectMetric ReadAcceptanceScreenRect(
        Word.Window window,
        Word.Range range)
    {
        window.GetPoint(
            out var left,
            out var top,
            out var width,
            out var height,
            range);
        return new ScreenRectMetric
        {
            X = left,
            Y = top,
            Width = width,
            Height = height,
        };
    }

    private static InkMetric MeasureAcceptanceInk(
        System.Drawing.Bitmap bitmap,
        ScreenRectMetric rectangle)
    {
        var minX = int.MaxValue;
        var minY = int.MaxValue;
        var maxX = -1;
        var maxY = -1;
        var count = 0;
        double sumY = 0;
        var x0 = Math.Max(0, rectangle.X);
        var y0 = Math.Max(0, rectangle.Y);
        var x1 = Math.Min(
            bitmap.Width,
            x0 + Math.Max(0, rectangle.Width));
        var y1 = Math.Min(
            bitmap.Height,
            y0 + Math.Max(0, rectangle.Height));
        for (var y = y0; y < y1; y++)
        {
            for (var x = x0; x < x1; x++)
            {
                var color = bitmap.GetPixel(x, y);
                if (color.R >= 150
                    || color.G >= 150
                    || color.B >= 150)
                    continue;
                count++;
                sumY += y;
                minX = Math.Min(minX, x);
                maxX = Math.Max(maxX, x);
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            }
        }
        if (count == 0)
            return new InkMetric();
        return new InkMetric
        {
            Count = count,
            MinX = minX,
            MaxX = maxX,
            MinY = minY,
            MaxY = maxY,
            CenterY = (minY + maxY) / 2.0,
            MeanY = sumY / count,
            Height = maxY - minY + 1,
        };
    }
}

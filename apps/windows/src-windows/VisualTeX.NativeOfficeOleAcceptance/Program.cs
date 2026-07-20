using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.PowerPointVsto;
using VisualTeX.WordVsto;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.NativeOfficeOleAcceptance;

internal static class Program
{
    private const string FormulaClsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}";
    private const string TypeLibraryId = "{DF66EC66-3B3A-4675-A7BE-30456A04EB96}";

    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length >= 1
            && string.Equals(args[0], "--installed-real-visual", StringComparison.OrdinalIgnoreCase))
        {
            if (args.Length < 3)
            {
                Console.Error.WriteLine(
                    "Usage: VisualTeX.NativeOfficeOleAcceptance.exe --installed-real-visual <SVG-fixture-directory> <artifact-directory>");
                return 2;
            }
            try
            {
                return RunInstalledRealVisualComparison(
                    Path.GetFullPath(args[1]),
                    Path.GetFullPath(args[2]));
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length < 1)
        {
            Console.Error.WriteLine(
                "Usage: VisualTeX.NativeOfficeOleAcceptance.exe <FormulaOleServer.exe> [artifact-directory]");
            return 2;
        }

        var serverPath = Path.GetFullPath(args[0]);
        if (!File.Exists(serverPath))
        {
            Console.Error.WriteLine($"Formula OLE LocalServer does not exist: {serverPath}");
            return 2;
        }
        if (HasExistingRegistration())
        {
            Console.Error.WriteLine(
                "VisualTeX Formula OLE is already registered. Acceptance refuses to overwrite an existing installation.");
            return 3;
        }

        var artifactRoot = args.Length >= 2
            ? Path.GetFullPath(args[1])
            : Path.Combine(Path.GetTempPath(), $"VisualTeX-Native-Office-OLE-{Guid.NewGuid():N}");
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(artifactRoot);
        Directory.CreateDirectory(previewRoot);
        var tracePath = Path.Combine(artifactRoot, "ole-server-trace.log");
        Environment.SetEnvironmentVariable("VISUALTEX_OLE_TRACE_PATH", tracePath);

        var wordPath = Path.Combine(artifactRoot, "VisualTeX-Native-OLE-Word.docx");
        var wordOmmlPath = Path.Combine(artifactRoot, "VisualTeX-OMML-OLE-RoundTrip.docx");
        var powerPointPath = Path.Combine(artifactRoot, "VisualTeX-Native-OLE-PowerPoint.pptx");
        var powerPointConversionPath = Path.Combine(
            artifactRoot,
            "VisualTeX-PowerPoint-Picture-To-OLE.pptx");
        var formulaId = Guid.NewGuid().ToString();
        var ommlFormulaId = Guid.NewGuid().ToString();
        var powerPointConversionFormulaId = Guid.NewGuid().ToString();
        var repeatedConversionFormulaId = Guid.NewGuid().ToString();
        var initial = CreatePreviewSet(previewRoot, formulaId, "initial", 420, 130);
        var updated = CreatePreviewSet(previewRoot, formulaId, "updated", 520, 150);
        var wide = CreatePreviewSet(
            previewRoot,
            powerPointConversionFormulaId,
            "wide-conversion",
            840,
            130);
        var display = CreatePreviewSet(
            previewRoot,
            repeatedConversionFormulaId,
            "word-display",
            360,
            56);
        CopyPreviewSet(initial, artifactRoot, "initial");
        CopyPreviewSet(updated, artifactRoot, "updated");
        CopyPreviewSet(wide, artifactRoot, "wide-conversion");
        CopyPreviewSet(display, artifactRoot, "word-display");
        var registered = false;

        try
        {
            Console.WriteLine("[1/9] Registering the per-user ATL LocalServer...");
            RunRegistration(serverPath, "/RegServerPerUser");
            registered = true;
            AssertRegistrationPresent(serverPath);

            Console.WriteLine("[2/9] Verifying Word OMML/OLE editing, mixed numbering, and cross-reference updates...");
            VerifyWordOmmlOleRoundTrip(wordOmmlPath, ommlFormulaId, initial, updated);
            VerifyWordMixedNumberingScenarios(initial, updated);
            VerifyRepeatedWordConversionsGeometryAndNativeSync(
                repeatedConversionFormulaId,
                display);

            Console.WriteLine("[3/9] Creating a real Word OLE object and saving DOCX...");
            CreateWordDocument(wordPath, formulaId, initial);

            Console.WriteLine("[4/9] Creating a real PowerPoint OLE object and saving PPTX...");
            CreatePowerPointDocument(powerPointPath, formulaId, initial);

            Console.WriteLine("[5/9] Converting a realistically resized PowerPoint picture formula to OLE...");
            VerifyPowerPointPictureToOleConversion(
                powerPointConversionPath,
                powerPointConversionFormulaId,
                wide,
                wide);

            ForceComCleanup();
            Console.WriteLine("[6/9] Unregistering the server and reopening cached previews offline...");
            RunRegistration(serverPath, "/UnregServerPerUser");
            registered = false;
            AssertRegistrationAbsent();
            VerifyWordCachedPreviewOffline(wordPath, formulaId);
            VerifyPowerPointCachedPreviewOffline(powerPointPath, formulaId);
            VerifyPowerPointCachedPreviewOffline(
                powerPointConversionPath,
                powerPointConversionFormulaId);

            Console.WriteLine("[7/9] Re-registering and updating the persisted Word OLE object...");
            RunRegistration(serverPath, "/RegServerPerUser");
            registered = true;
            AssertRegistrationPresent(serverPath);
            UpdateAndVerifyWord(wordPath, formulaId, updated);

            Console.WriteLine("[8/9] Updating the persisted PowerPoint OLE object...");
            UpdateAndVerifyPowerPoint(powerPointPath, formulaId, updated);

            Console.WriteLine("[9/9] Final offline reopen and cleanup verification...");
            ForceComCleanup();
            RunRegistration(serverPath, "/UnregServerPerUser");
            registered = false;
            AssertRegistrationAbsent();
            VerifyWordCachedPreviewOffline(wordPath, formulaId);
            VerifyPowerPointCachedPreviewOffline(powerPointPath, formulaId);
            VerifyPowerPointCachedPreviewOffline(
                powerPointConversionPath,
                powerPointConversionFormulaId);

            Console.WriteLine("VisualTeX real Word OMML/OLE and PowerPoint native OLE acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        catch (Exception error)
        {
            Console.Error.WriteLine(error);
            Console.Error.WriteLine($"Acceptance artifacts retained for diagnosis: {artifactRoot}");
            return 1;
        }
        finally
        {
            ForceComCleanup();
            if (registered || HasExistingRegistration())
            {
                try { RunRegistration(serverPath, "/UnregServerPerUser"); } catch { }
            }
            TryDelete(initial.SvgPath);
            TryDelete(initial.EmfPath);
            TryDelete(initial.PngPath);
            TryDelete(updated.SvgPath);
            TryDelete(updated.EmfPath);
            TryDelete(updated.PngPath);
            TryDelete(wide.SvgPath);
            TryDelete(wide.EmfPath);
            TryDelete(wide.PngPath);
            TryDelete(display.SvgPath);
            TryDelete(display.EmfPath);
            TryDelete(display.PngPath);
        }
    }

    private static int RunInstalledRealVisualComparison(
        string fixtureRoot,
        string artifactRoot)
    {
        if (!Directory.Exists(fixtureRoot))
            throw new DirectoryNotFoundException($"SVG fixture directory does not exist: {fixtureRoot}");
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "The installed-real-visual mode requires the formally installed VisualTeX OLE registration.");

        Directory.CreateDirectory(artifactRoot);
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(previewRoot);
        var formulas = new[]
        {
            @"\int ac",
            @"\int x\,dy",
            @"\alpha\beta d f d f aaaaabbbbb",
            @"\frac{a+b}{c+d}+\sqrt{x^2+y^2}",
            @"\int_0^1 x^2\,dx+\sum_{n=1}^{\infty}\frac{1}{n^2}",
        };
        var previews = new List<PreviewSet>();
        var comparisonRows = new List<ComparisonRow>();
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        try
        {
            for (var index = 0; index < formulas.Length; index++)
            {
                var fixture = Path.Combine(fixtureRoot, $"formula-{index + 1}.svg");
                previews.Add(CreatePreviewSetFromSvgFixture(
                    previewRoot,
                    fixture,
                    $"real-{index + 1}"));
            }

            application = new PowerPoint.Application
            {
                Visible = Office.MsoTriState.msoTrue,
            };
            presentation = application.Presentations.Add(Office.MsoTriState.msoTrue);
            presentation.PageSetup.SlideWidth = 960f;
            presentation.PageSetup.SlideHeight = 540f;
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            var service = new PowerPointFormulaService(application);

            AddComparisonLabel(slide, "同一 SVG/EMF 来源：转换前图片", 28f, 12f, 360f);
            AddComparisonLabel(slide, "转换后 VisualTeX OLE", 500f, 12f, 360f);

            for (var index = 0; index < previews.Count; index++)
            {
                var preview = previews[index];
                var formulaId = Guid.NewGuid().ToString();
                var ratio = preview.Width / (float)Math.Max(1, preview.Height);
                var height = Math.Max(38f, Math.Min(66f, preview.Height * 1.05f));
                var width = height * ratio;
                if (width > 380f)
                {
                    var scale = 380f / width;
                    width *= scale;
                    height *= scale;
                }
                var top = 50f + index * 94f + (72f - height) / 2f;
                var sourceLeft = 35f;
                var candidateLeft = 505f;
                PowerPoint.Shape? source = null;
                PowerPoint.Shape? candidate = null;
                PowerPoint.Tags? tags = null;
                PowerPoint.Shape? converted = null;
                try
                {
                    source = slide.Shapes.AddPicture(
                        preview.PngPath,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        sourceLeft,
                        top,
                        width,
                        height);
                    source.Name = $"VisualTeXSource_{index + 1}";

                    candidate = slide.Shapes.AddPicture(
                        preview.PngPath,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        candidateLeft,
                        top,
                        width,
                        height);
                    var metadata = CreateMetadata(
                        formulaId,
                        preview,
                        formulas[index],
                        $"real-visual-{index + 1}");
                    var encoded = FormulaMetadataCodec.Encode(metadata);
                    candidate.Name = $"VisualTeX_{formulaId}";
                    candidate.AlternativeText = encoded;
                    tags = candidate.Tags;
                    tags.Add("VisualTeXFormulaId", formulaId);
                    tags.Add("VisualTeXMetadata", encoded);

                    var session = new OfficeSessionDocument
                    {
                        Id = Guid.NewGuid().ToString(),
                        Mode = "edit",
                        Host = "powerpoint",
                        FormulaId = formulaId,
                        Title = $"Real formula visual comparison {index + 1}",
                        Lines = new List<FormulaLine>
                        {
                            new() { Id = Guid.NewGuid().ToString(), Latex = formulas[index] },
                        },
                        CodeFormat = "raw",
                        DisplayMode = "block",
                        ObjectMode = FormulaOleContract.NativeOleMode,
                        Numbered = false,
                        Dirty = false,
                        SourceObjectId = candidate.Name,
                        OriginalMetadata = metadata,
                        ExportResult = new OfficeExportDocument
                        {
                            Width = preview.Width,
                            Height = preview.Height,
                            Baseline = preview.Height * 0.62f,
                        },
                    };
                    var result = service.ReplaceOle(
                        session,
                        preview.PngPath,
                        preview.EmfPath);
                    Release(candidate);
                    candidate = null;
                    converted = slide.Shapes[result.ObjectId];

                    var exportWidth = 2400;
                    var exportHeight = Math.Max(1, (int)Math.Round(exportWidth * height / width));
                    var sourcePath = Path.Combine(artifactRoot, $"formula-{index + 1}-picture.png");
                    var olePath = Path.Combine(artifactRoot, $"formula-{index + 1}-ole.png");
                    source.Export(
                        sourcePath,
                        PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                        exportWidth,
                        exportHeight,
                        PowerPoint.PpExportMode.ppScaleXY);
                    converted.Export(
                        olePath,
                        PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                        exportWidth,
                        exportHeight,
                        PowerPoint.PpExportMode.ppScaleXY);
                    ReportIndependentShapeExports(
                        sourcePath,
                        olePath,
                        $"real formula {index + 1}");
                    comparisonRows.Add(new ComparisonRow(
                        index + 1,
                        sourceLeft,
                        candidateLeft,
                        top,
                        width,
                        height));
                    Console.WriteLine(
                        $"  real formula {index + 1} shape: source={source.Width:0.###}x{source.Height:0.###} pt; "
                        + $"OLE={converted.Width:0.###}x{converted.Height:0.###} pt; "
                        + $"EMF natural={preview.Width}x{preview.Height} CSS px.");
                }
                finally
                {
                    Release(converted);
                    Release(tags);
                    Release(candidate);
                    Release(source);
                }
            }

            var presentationPath = Path.Combine(
                artifactRoot,
                "VisualTeX-Real-Formula-Picture-vs-OLE.pptx");
            var slidePath = Path.Combine(
                artifactRoot,
                "VisualTeX-Real-Formula-Picture-vs-OLE.png");
            presentation.SaveAs(
                presentationPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                Office.MsoTriState.msoFalse);
            slide.Export(slidePath, "PNG", 1920, 1080);
            AssertSameSlidePairBounds(
                slidePath,
                comparisonRows,
                presentation.PageSetup.SlideWidth,
                presentation.PageSetup.SlideHeight,
                "same-slide before reopen");
            Console.WriteLine($"Real formula same-slide visual artifact: {slidePath}");

            presentation.Close();
            Release(slide);
            slide = null;
            Release(presentation);
            presentation = null;

            presentation = application.Presentations.Open(
                presentationPath,
                Office.MsoTriState.msoTrue,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue);
            slide = presentation.Slides[1];
            var reopenedPath = Path.Combine(
                artifactRoot,
                "VisualTeX-Real-Formula-Picture-vs-OLE-Reopened.png");
            slide.Export(reopenedPath, "PNG", 1920, 1080);
            AssertSameSlidePairBounds(
                reopenedPath,
                comparisonRows,
                presentation.PageSetup.SlideWidth,
                presentation.PageSetup.SlideHeight,
                "same-slide after reopen");
            Console.WriteLine($"Saved/reopened same-slide visual artifact: {reopenedPath}");
            return 0;
        }
        finally
        {
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
            foreach (var preview in previews)
            {
                TryDelete(preview.SvgPath);
                TryDelete(preview.EmfPath);
                TryDelete(preview.PngPath);
            }
        }
    }

    private static PreviewSet CreatePreviewSetFromSvgFixture(
        string previewRoot,
        string fixturePath,
        string suffix)
    {
        if (!File.Exists(fixturePath))
            throw new FileNotFoundException("Real formula SVG fixture does not exist.", fixturePath);
        var document = XDocument.Load(fixturePath, LoadOptions.None);
        var root = document.Root
            ?? throw new InvalidDataException("Real formula SVG fixture has no root element.");
        if (!double.TryParse(
                root.Attribute("width")?.Value,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var width)
            || !double.TryParse(
                root.Attribute("height")?.Value,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var height)
            || width <= 0
            || height <= 0)
            throw new InvalidDataException("Real formula SVG fixture dimensions are invalid.");

        var token = $"{suffix}-{Guid.NewGuid():N}";
        var svgPath = Path.Combine(previewRoot, token + ".svg");
        var pngPath = Path.Combine(previewRoot, token + ".png");
        File.Copy(fixturePath, svgPath, true);
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
            svgPath,
            (float)width,
            (float)height);
        OfficeOlePreview.ValidateVectorEmf(emfPath);
        var browserPngPath = Path.Combine(
            Path.GetDirectoryName(fixturePath)
                ?? throw new InvalidOperationException("Real formula fixture has no parent directory."),
            Path.GetFileNameWithoutExtension(fixturePath) + "-browser.png");
        if (File.Exists(browserPngPath))
        {
            File.Copy(browserPngPath, pngPath, true);
        }
        else
        {
            CreatePngFromEmf(
                emfPath,
                pngPath,
                Math.Max(1, (int)Math.Ceiling(width * 4)),
                Math.Max(1, (int)Math.Ceiling(height * 4)));
        }

        var bytes = File.ReadAllBytes(emfPath);
        if (bytes.Length < 40)
            throw new InvalidDataException("Real formula EMF header is truncated.");
        var frameLeft = BitConverter.ToInt32(bytes, 24);
        var frameTop = BitConverter.ToInt32(bytes, 28);
        var frameRight = BitConverter.ToInt32(bytes, 32);
        var frameBottom = BitConverter.ToInt32(bytes, 36);
        Console.WriteLine(
            $"  {suffix} SVG={width:0.####}x{height:0.####} px; "
            + $"EMF rclFrame={frameLeft},{frameTop},{frameRight},{frameBottom} "
            + $"({frameRight - frameLeft}x{frameBottom - frameTop} HIMETRIC).");
        return new PreviewSet(
            svgPath,
            emfPath,
            pngPath,
            Math.Max(1, (int)Math.Round(width)),
            Math.Max(1, (int)Math.Round(height)));
    }

    private static void AddComparisonLabel(
        PowerPoint.Slide slide,
        string text,
        float left,
        float top,
        float width)
    {
        PowerPoint.Shape? label = null;
        PowerPoint.TextFrame? frame = null;
        PowerPoint.TextRange? range = null;
        try
        {
            label = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                left,
                top,
                width,
                28f);
            frame = label.TextFrame;
            range = frame.TextRange;
            range.Text = text;
            range.Font.Size = 16f;
            range.Font.Bold = Office.MsoTriState.msoTrue;
        }
        finally
        {
            Release(range);
            Release(frame);
            Release(label);
        }
    }

    private static void CopyPreviewSet(
        PreviewSet preview,
        string artifactRoot,
        string prefix)
    {
        File.Copy(preview.SvgPath, Path.Combine(artifactRoot, prefix + ".svg"), true);
        File.Copy(preview.EmfPath, Path.Combine(artifactRoot, prefix + ".emf"), true);
        File.Copy(preview.PngPath, Path.Combine(artifactRoot, prefix + ".png"), true);
    }

    private static PreviewSet CreatePreviewSet(
        string previewRoot,
        string formulaId,
        string suffix,
        int width,
        int height)
    {
        var token = $"{formulaId}-{suffix}-{Guid.NewGuid():N}";
        var svgPath = Path.Combine(previewRoot, token + ".svg");
        var pngPath = Path.Combine(previewRoot, token + ".png");
        var strokeWidth = Math.Max(2f, height / 28f);
        var svg = $"""
            <svg xmlns="http://www.w3.org/2000/svg"
                 viewBox="0 0 {width} {height}">
              <rect x="0" y="0" width="{width}" height="{height}" fill="transparent" opacity="0.001" stroke="none" />
              <path d="M {width * 0.08f:0.###} {height * 0.72f:0.###}
                       C {width * 0.30f:0.###} {height * 0.12f:0.###},
                         {width * 0.52f:0.###} {height * 0.92f:0.###},
                         {width * 0.76f:0.###} {height * 0.28f:0.###}"
                    fill="none" stroke="#111111" stroke-width="{strokeWidth:0.###}" />
              <line x1="{width * 0.55f:0.###}" y1="{height * 0.48f:0.###}"
                    x2="{width * 0.92f:0.###}" y2="{height * 0.48f:0.###}"
                    stroke="#111111" stroke-width="{strokeWidth:0.###}" />
            </svg>
            """;
        File.WriteAllText(svgPath, svg, new UTF8Encoding(false));
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, width, height);
        OfficeOlePreview.ValidateVectorEmf(emfPath);
        CreatePngFromEmf(emfPath, pngPath, width * 2, height * 2);
        return new PreviewSet(svgPath, emfPath, pngPath, width, height);
    }

    private static void CreatePngFromEmf(
        string emfPath,
        string pngPath,
        int width,
        int height)
    {
        using var metafile = new Metafile(emfPath);
        using var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.Clear(Color.Transparent);
        graphics.CompositingMode = CompositingMode.SourceOver;
        graphics.CompositingQuality = CompositingQuality.HighQuality;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.SmoothingMode = SmoothingMode.HighQuality;
        graphics.DrawImage(metafile, new Rectangle(0, 0, width, height));
        bitmap.Save(pngPath, ImageFormat.Png);
    }

    private static FormulaMetadata CreateMetadata(
        string formulaId,
        PreviewSet preview,
        string latex,
        string suffix)
    {
        var now = DateTimeOffset.UtcNow.ToString("O");
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = $"Native Office OLE {suffix}",
            Latex = latex,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            Numbered = false,
            RenderWidthPx = preview.Width,
            RenderHeightPx = preview.Height,
            Baseline = preview.Height * 0.72,
            CreatedWithVersion = "1.1.0",
            UpdatedWithVersion = "1.1.0",
            CreatedAt = now,
            UpdatedAt = now,
        };
        metadata.Validate();
        return metadata;
    }

    private static void VerifyWordOmmlOleRoundTrip(
        string path,
        string formulaId,
        PreviewSet initial,
        PreviewSet updated)
    {
        const string initialLatex = @"\frac{a}{b}+x^2";
        const string initialMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo>"
            + "<msup><mi>x</mi><mn>2</mn></msup></math>";
        const string updatedLatex = @"\sqrt{x^2+y^2}";
        const string updatedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
            + "<msup><mi>y</mi><mn>2</mn></msup></mrow></msqrt></math>";
        const string finalLatex = @"\int_0^1 t^2\,dt";
        const string finalMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup>"
            + "<msup><mi>t</mi><mn>2</mn></msup><mi>d</mi><mi>t</mi></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Range? range = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.InlineShape? shape = null;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            selection = application.Selection;
            selection.SetRange(0, 0);
            var service = new WordFormulaService(application);

            var createSession = CreateWordSession(
                formulaId,
                "create",
                FormulaOleContract.WordOmmlMode,
                initialLatex,
                initialMathMl,
                initial,
                numbered: true,
                originalMetadata: null);
            service.InsertOmml(createSession, initialMathMl);
            Assert(document.OMaths.Count >= 1, "Word OMML insertion did not create an OMath.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("OMML bookmark was not created.");
            var stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException("OMML Custom XML metadata was not readable.");
            AssertEqual(initialLatex, stored.Latex, "OMML metadata lost the original LaTeX.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var initialOmmlSelection = ReadSelectedFormula(application, service, equationRange);
            AssertEqual(
                FormulaOleContract.WordOmmlMode,
                initialOmmlSelection.ObjectMode ?? string.Empty,
                "OMML selection was not recognized as wordOmml.");
            Assert(
                !WordDoubleClickRouting.ShouldOpenVisualTeX(initialOmmlSelection),
                "Word OMML double-click would incorrectly leave the native Word editor.");
            Release(equationRange);
            equationRange = null;

            AssertEqual(
                "1",
                WordEquationNumbering.Reconcile(document).ToString(),
                "OMML formula did not participate in Word equation numbering.");
            Assert(
                CountFieldCodes(document, "SEQ ") >= 1,
                "OMML numbering did not create a native Word SEQ field.");
            AssertWordEquationNumberOutsideOmml(document, formulaId);
            var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            Assert(targets.Count == 1, "OMML formula was not exposed as a Word cross-reference target.");
            Release(selection);
            selection = application.Selection;
            selection.SetRange(document.Content.End - 1, document.Content.End - 1);
            selection.TypeParagraph();
            WordEquationNumbering.InsertEquationReference(
                document,
                selection,
                targets[0],
                EquationReferenceStyle.Parenthesized);
            Assert(
                CountFieldCodes(document, "REF ") >= 2,
                "OMML visible numbering and inserted cross-reference did not use native REF fields.");

            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            maths = equationRange.OMaths;
            math = maths[1];
            math.Linearize();
            math.BuildUp();
            Release(equationRange);
            equationRange = math.Range;
            Assert(
                equationRange.WordOpenXML.Contains("oMath")
                    && document.OMaths.Count >= 1
                    && WordOmmlFormulaStore.FindByFormulaId(document, formulaId) is not null,
                "Word-native OMML Linearize/BuildUp editing cycle did not preserve the formula.");
            Release(math);
            math = null;
            Release(maths);
            maths = null;
            Release(equationRange);
            equationRange = null;
            Release(bookmark);
            bookmark = null;

            var updatedSession = CreateWordSession(
                formulaId,
                "edit",
                FormulaOleContract.WordOmmlMode,
                updatedLatex,
                updatedMathMl,
                updated,
                numbered: true,
                originalMetadata: stored);
            service.ReplaceOmml(updatedSession, updatedMathMl);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("VisualTeX OMML edit removed the bookmark.");
            stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException("Updated OMML metadata was not readable.");
            AssertEqual(updatedLatex, stored.Latex, "VisualTeX OMML edit did not update LaTeX metadata.");
            Release(bookmark);
            bookmark = null;

            var oleSession = CreateWordSession(
                formulaId,
                "edit",
                FormulaOleContract.NativeOleMode,
                updatedLatex,
                updatedMathMl,
                updated,
                numbered: true,
                originalMetadata: stored);
            service.ReplaceOle(oleSession, updated.PngPath, updated.EmfPath);
            shape = FindWordFormula(document, formulaId)
                ?? throw new InvalidOperationException("OMML to OLE conversion did not create an InlineShape.");
            range = shape.Range;
            var convertedOleSelection = ReadSelectedFormula(application, service, range);
            AssertEqual(
                FormulaOleContract.NativeOleMode,
                convertedOleSelection.ObjectMode ?? string.Empty,
                "Converted OLE formula was not recognized as nativeOle.");
            Assert(
                WordDoubleClickRouting.ShouldOpenVisualTeX(convertedOleSelection),
                "Converted OLE formula would not route double-click to the VisualTeX formula editor.");
            Assert(
                WordOmmlFormulaStore.FindByFormulaId(document, formulaId) is null,
                "OMML bookmark remained after conversion to OLE.");
            Release(range);
            range = null;
            Release(shape);
            shape = null;

            var finalSession = CreateWordSession(
                formulaId,
                "edit",
                FormulaOleContract.WordOmmlMode,
                finalLatex,
                finalMathMl,
                initial,
                numbered: true,
                originalMetadata: oleSession.ToMetadata());
            service.ReplaceOmml(finalSession, finalMathMl);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("OLE to OMML conversion did not restore the bookmark.");
            stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException("Round-trip OMML metadata was not readable.");
            AssertEqual(finalLatex, stored.Latex, "OLE to OMML conversion lost final LaTeX metadata.");
            Assert(document.OMaths.Count >= 1, "OLE to OMML conversion did not restore an OMath.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var roundTripOmmlSelection = ReadSelectedFormula(application, service, equationRange);
            Assert(
                !WordDoubleClickRouting.ShouldOpenVisualTeX(roundTripOmmlSelection),
                "OLE to OMML conversion did not restore Word-native double-click editing.");
            Release(equationRange);
            equationRange = null;
            WordEquationNumbering.Reconcile(document);
            AssertWordEquationNumberOutsideOmml(document, formulaId);
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);

            Release(bookmark);
            bookmark = null;
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = application.Documents.Open(
                FileName: path,
                ConfirmConversions: false,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);
            Assert(document.OMaths.Count >= 1, "Saved DOCX did not preserve native OMML.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("Saved DOCX did not preserve the OMML bookmark.");
            stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException("Saved DOCX did not preserve OMML Custom XML metadata.");
            AssertEqual(finalLatex, stored.Latex, "Saved DOCX changed OMML LaTeX metadata.");
            Assert(
                CountFieldCodes(document, "SEQ ") >= 1 && CountFieldCodes(document, "REF ") >= 2,
                "Saved DOCX did not preserve native numbering and cross-reference fields.");
            AssertWordEquationNumberOutsideOmml(document, formulaId);
        }
        finally
        {
            Release(shape);
            Release(math);
            Release(maths);
            Release(equationRange);
            Release(bookmark);
            Release(range);
            Release(selection);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void VerifyRepeatedWordConversionsGeometryAndNativeSync(
        string formulaId,
        PreviewSet preview)
    {
        const string initialLatex = @"\frac{a+b}{c+d}+\sqrt{x^2+y^2}";
        const string initialMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow>"
            + "<mrow><mi>c</mi><mo>+</mo><mi>d</mi></mrow></mfrac><mo>+</mo>"
            + "<msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
            + "<msup><mi>y</mi><mn>2</mn></msup></mrow></msqrt></math>";
        const string editedLatex = @"\frac{a+b}{c+d}+\sqrt{x^2+y^2}+z^3";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow>"
            + "<mrow><mi>c</mi><mo>+</mo><mi>d</mi></mrow></mfrac><mo>+</mo>"
            + "<msqrt><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
            + "<msup><mi>y</mi><mn>2</mn></msup></mrow></msqrt><mo>+</mo>"
            + "<msup><mi>z</mi><mn>3</mn></msup></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.InlineShape? shape = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        try
        {
            application = new Word.Application
            {
                Visible = true,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            document.Activate();
            application.ActiveWindow.View.Type = Word.WdViewType.wdPrintView;
            selection = application.Selection;
            selection.SetRange(0, 0);
            var service = new WordFormulaService(application);

            service.InsertOmml(
                CreateWordSession(
                    formulaId,
                    "create",
                    FormulaOleContract.WordOmmlMode,
                    initialLatex,
                    initialMathMl,
                    preview,
                    numbered: true,
                    originalMetadata: null),
                initialMathMl);
            WordEquationNumbering.Reconcile(document);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("Repeated-conversion OMML bookmark was not created.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            paragraph = equationRange.Paragraphs[1];
            paragraphRange = paragraph.Range;
            var stableParagraphStart = paragraphRange.Start;
            AssertNumberedFormulaGeometry(
                application,
                document,
                formulaId,
                equationRange,
                null,
                expectedOleHeight: null,
                context: "initial OMML");

            var stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException("Initial OMML metadata is missing.");
            Assert(
                !string.IsNullOrWhiteSpace(stored.NativeOmmlFingerprint),
                "New OMML formula did not store its content fingerprint.");

            AppendUsingWordNativeEquationEditor(document, bookmark, "+z^3");
            Release(equationRange);
            equationRange = null;
            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("Word-native editing removed the VisualTeX OMML anchor.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var refreshedSelection = ReadSelectedFormula(application, service, equationRange);
            AssertEqual(
                FormulaOleContract.WordOmmlMode,
                refreshedSelection.ObjectMode ?? string.Empty,
                "Word-native edited OMML was no longer recognized by VisualTeX.");
            var refreshed = refreshedSelection.Metadata
                ?? throw new InvalidOperationException("Word-native edited OMML returned no VisualTeX metadata.");
            Assert(
                refreshed.Latex.IndexOf("z", StringComparison.Ordinal) >= 0
                    && refreshed.Latex.IndexOf("^{3}", StringComparison.Ordinal) >= 0,
                $"VisualTeX did not import the Word-native equation addition. Imported LaTeX: {refreshed.Latex}");
            Assert(
                !string.Equals(
                    stored.NativeOmmlFingerprint,
                    refreshed.NativeOmmlFingerprint,
                    StringComparison.OrdinalIgnoreCase),
                "Word-native edit did not change the OMML fingerprint seen by VisualTeX.");

            // Commit a VisualTeX edit based on the source reconstructed from the
            // current Word OMML. This proves the native addition is not only
            // visible in the editor payload, but can survive the next update.
            var normalizedSession = CreateWordSession(
                formulaId,
                "edit",
                FormulaOleContract.WordOmmlMode,
                editedLatex,
                editedMathMl,
                preview,
                numbered: true,
                originalMetadata: refreshed);
            service.ReplaceOmml(normalizedSession, editedMathMl);

            for (var cycle = 1; cycle <= 5; cycle++)
            {
                Release(bookmark);
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException($"Cycle {cycle}: OMML bookmark is missing.");
                var cycleMetadata = WordOmmlFormulaStore.TryRead(document, bookmark)
                    ?? throw new InvalidOperationException($"Cycle {cycle}: OMML metadata is missing.");
                var oleSession = CreateWordSession(
                    formulaId,
                    "edit",
                    FormulaOleContract.NativeOleMode,
                    editedLatex,
                    editedMathMl,
                    preview,
                    numbered: true,
                    originalMetadata: cycleMetadata);
                service.ReplaceOle(oleSession, preview.PngPath, preview.EmfPath);
                shape = FindWordFormula(document, formulaId)
                    ?? throw new InvalidOperationException($"Cycle {cycle}: OMML to OLE produced no object.");
                equationRange = shape.Range;
                paragraph = equationRange.Paragraphs[1];
                paragraphRange = paragraph.Range;
                AssertEqual(
                    stableParagraphStart.ToString(),
                    paragraphRange.Start.ToString(),
                    $"Cycle {cycle}: OMML to OLE moved the formula to another paragraph.");
                AssertClose(
                    preview.Height * 0.75f,
                    shape.Height,
                    1.5f,
                    $"Cycle {cycle}: OLE formula height changed.");
                AssertClose(
                    preview.Width / (float)preview.Height,
                    shape.Width / shape.Height,
                    0.05f,
                    $"Cycle {cycle}: OLE formula aspect ratio changed.");
                AssertNumberedFormulaGeometry(
                    application,
                    document,
                    formulaId,
                    equationRange,
                    shape,
                    preview.Height * 0.75f,
                    $"OLE cycle {cycle}");
                var oleSelection = ReadSelectedFormula(application, service, equationRange);
                Assert(
                    WordDoubleClickRouting.ShouldOpenVisualTeX(oleSelection),
                    $"Cycle {cycle}: converted OLE no longer routes to VisualTeX editing.");
                Assert(
                    WordOmmlFormulaStore.FindByFormulaId(document, formulaId) is null,
                    $"Cycle {cycle}: stale OMML selection anchor remained after conversion to OLE.");
                Assert(
                    document.OMaths.Count == 0,
                    $"Cycle {cycle}: an empty native OMath container still surrounds the converted OLE object.");
                Word.OMaths? paragraphMaths = null;
                try
                {
                    paragraphMaths = paragraphRange.OMaths;
                    Assert(
                        paragraphMaths.Count == 0,
                        $"Cycle {cycle}: converted OLE remains inside the native equation paragraph container.");
                }
                finally { Release(paragraphMaths); }

                var ommlSession = CreateWordSession(
                    formulaId,
                    "edit",
                    FormulaOleContract.WordOmmlMode,
                    editedLatex,
                    editedMathMl,
                    preview,
                    numbered: true,
                    originalMetadata: oleSession.ToMetadata());
                service.ReplaceOmml(ommlSession, editedMathMl);
                Release(shape);
                shape = null;
                Release(equationRange);
                equationRange = null;
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;

                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidOperationException($"Cycle {cycle}: OLE to OMML did not restore the anchor.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                paragraph = equationRange.Paragraphs[1];
                paragraphRange = paragraph.Range;
                AssertEqual(
                    stableParagraphStart.ToString(),
                    paragraphRange.Start.ToString(),
                    $"Cycle {cycle}: OLE to OMML moved the formula to another paragraph.");
                var ommlSelection = ReadSelectedFormula(application, service, equationRange);
                Assert(
                    !WordDoubleClickRouting.ShouldOpenVisualTeX(ommlSelection),
                    $"Cycle {cycle}: OMML no longer preserves Word-native double-click editing.");
                AssertNumberedFormulaGeometry(
                    application,
                    document,
                    formulaId,
                    equationRange,
                    null,
                    expectedOleHeight: null,
                    context: $"OMML cycle {cycle}");
                AssertEqual(
                    "1",
                    WordEquationNumbering.Reconcile(document).ToString(),
                    $"Cycle {cycle}: numbered formula count changed.");
                var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
                AssertTargetNumbers(targets, (formulaId, "1"));
            }

            Console.WriteLine(
                "  Word repeated conversion: 5 OMML↔OLE cycles, geometry, numbering, editing routes, and native OMML source sync passed.");
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(shape);
            Release(equationRange);
            Release(bookmark);
            Release(selection);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AppendUsingWordNativeEquationEditor(
        Word.Document document,
        Word.Bookmark bookmark,
        string suffix)
    {
        Word.Range? equationRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? insertion = null;
        try
        {
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            maths = equationRange.OMaths;
            Assert(maths.Count == 1, "Word-native OMML edit could not locate exactly one OMath.");
            math = maths[1];
            math.Linearize();
            insertion = math.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertBefore(suffix);
            math.BuildUp();
            var rebuiltText = math.Range.Text ?? string.Empty;
            Assert(
                rebuiltText.IndexOf("z", StringComparison.OrdinalIgnoreCase) >= 0,
                $"Word-native equation editor did not retain the inserted suffix. Text: {rebuiltText}");
        }
        finally
        {
            Release(insertion);
            Release(math);
            Release(maths);
            Release(equationRange);
        }
    }

    private static void AssertNumberedFormulaGeometry(
        Word.Application application,
        Word.Document document,
        string formulaId,
        Word.Range formulaRange,
        Word.InlineShape? oleShape,
        float? expectedOleHeight,
        string context)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabStops = null;
        Word.ListFormat? listFormat = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        Word.Range? formulaStart = null;
        Word.Range? formulaEnd = null;
        Word.Range? numberEnd = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Microsoft.Office.Interop.Word.Font? formulaFont = null;
        try
        {
            document.Repaginate();
            Thread.Sleep(100);
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraphRange.ParagraphFormat;
            tabStops = format.TabStops;
            Assert(tabStops.Count >= 2, $"{context}: numbered paragraph lost center/right tab stops.");
            listFormat = paragraphRange.ListFormat;
            Assert(
                listFormat.ListType == Word.WdListType.wdListNoNumbering,
                $"{context}: numbered formula paragraph inherited a list marker/black square.");

            bookmarks = document.Bookmarks;
            var visibleName = WordEquationNumbering.EquationBookmarkName(formulaId);
            Assert(bookmarks.Exists(visibleName), $"{context}: visible equation number bookmark is missing.");
            numberBookmark = bookmarks[visibleName];
            numberRange = numberBookmark.Range;
            Assert(
                (numberRange.Text ?? string.Empty).StartsWith("\t(", StringComparison.Ordinal)
                    && (numberRange.Text ?? string.Empty).EndsWith(")", StringComparison.Ordinal),
                $"{context}: equation number text is malformed: '{numberRange.Text}'.");

            formulaStart = formulaRange.Duplicate;
            formulaStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            formulaEnd = formulaRange.Duplicate;
            formulaEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            numberEnd = numberRange.Duplicate;
            numberEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

            var formulaStartX = Convert.ToSingle(
                formulaStart.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage]);
            var formulaEndX = oleShape is null
                ? Convert.ToSingle(formulaEnd.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage])
                : formulaStartX + oleShape.Width;
            var formulaY = Convert.ToSingle(
                formulaStart.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            var numberEndX = Convert.ToSingle(
                numberEnd.Information[Word.WdInformation.wdHorizontalPositionRelativeToPage]);
            var numberY = Convert.ToSingle(
                numberEnd.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);

            sections = formulaRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            var textLeft = pageSetup.LeftMargin;
            var textRight = pageSetup.PageWidth - pageSetup.RightMargin;
            var textCenter = (textLeft + textRight) / 2f;
            var formulaCenter = (formulaStartX + formulaEndX) / 2f;

            Assert(formulaStartX >= textLeft - 8f, $"{context}: formula begins outside the text area.");
            Assert(formulaEndX < numberEndX - 4f, $"{context}: formula overlaps or follows its number.");
            AssertClose(textCenter, formulaCenter, 18f, $"{context}: formula is not centered in the text area.");
            AssertClose(textRight, numberEndX, 24f, $"{context}: equation number is not right-aligned.");
            AssertClose(formulaY, numberY, 28f, $"{context}: equation number is not vertically aligned with the formula.");

            if (TryGetWordRangeScreenRect(
                    application,
                    formulaRange,
                    out var formulaLeftPx,
                    out var formulaTopPx,
                    out var formulaWidthPx,
                    out var formulaHeightPx)
                && TryGetWordRangeScreenRect(
                    application,
                    numberRange,
                    out var numberLeftPx,
                    out var numberTopPx,
                    out var numberWidthPx,
                    out var numberHeightPx))
            {
                var formulaCenterYPx = formulaTopPx + formulaHeightPx / 2f;
                var numberCenterYPx = numberTopPx + numberHeightPx / 2f;
                var centerTolerancePx = Math.Max(8f, formulaHeightPx * 0.22f);
                AssertClose(
                    formulaCenterYPx,
                    numberCenterYPx,
                    centerTolerancePx,
                    $"{context}: visible equation number is not centered beside the actual rendered formula glyphs.");
                Assert(
                    formulaWidthPx > 4 && formulaHeightPx > 4,
                    $"{context}: Word returned an invalid visible formula rectangle.");
                if (oleShape is null)
                    Assert(
                        formulaHeightPx >= 18,
                        $"{context}: display OMML is visually the same size as an inline formula ({formulaHeightPx}px high).");
                Console.WriteLine(
                    $"  {context} visible rect: formula={formulaLeftPx},{formulaTopPx} "
                    + $"{formulaWidthPx}x{formulaHeightPx}px; number={numberLeftPx},{numberTopPx} "
                    + $"{numberWidthPx}x{numberHeightPx}px.");
            }

            if (oleShape is null)
            {
                formulaFont = formulaRange.Font;
                var fontSize = formulaFont.Size;
                Assert(
                    !float.IsNaN(fontSize) && fontSize >= 14f && fontSize <= 24f,
                    $"{context}: display OMML font size is not larger than inline text: {fontSize} pt.");
            }
            else if (expectedOleHeight.HasValue)
            {
                AssertClose(
                    expectedOleHeight.Value,
                    oleShape.Height,
                    1.5f,
                    $"{context}: OLE formula physical height is abnormal.");
            }

            Console.WriteLine(
                $"  {context} geometry: formula={formulaStartX:0.0}-{formulaEndX:0.0} pt, "
                + $"center={formulaCenter:0.0}/{textCenter:0.0}, numberEnd={numberEndX:0.0}/{textRight:0.0}, "
                + $"y={formulaY:0.0}/{numberY:0.0}.");
        }
        finally
        {
            Release(formulaFont);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(numberEnd);
            Release(formulaEnd);
            Release(formulaStart);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(listFormat);
            Release(tabStops);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryGetWordRangeScreenRect(
        Word.Application application,
        Word.Range range,
        out int left,
        out int top,
        out int width,
        out int height)
    {
        left = top = width = height = 0;
        Word.Window? window = null;
        try
        {
            window = application.ActiveWindow;
            if (window is null) return false;
            try { window.ScrollIntoView(range, true); } catch { }
            Thread.Sleep(40);
            window.GetPoint(out left, out top, out width, out height, range);
            return width > 0 && height > 0;
        }
        catch
        {
            left = top = width = height = 0;
            return false;
        }
        finally { Release(window); }
    }

    private static void VerifyWordMixedNumberingScenarios(
        PreviewSet initial,
        PreviewSet updated)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><mn>1</mn></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>=</mo><mi>c</mi></math>";
        const string insertedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msqrt><mrow><mi>p</mi><mo>+</mo><mi>q</mi></mrow></msqrt></math>";

        var firstId = Guid.NewGuid().ToString();
        var secondId = Guid.NewGuid().ToString();
        var insertedId = Guid.NewGuid().ToString();
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            var service = new WordFormulaService(application);

            ResetWordSelection(application, ref selection, 0, 0);
            service.InsertOmml(
                CreateWordSession(
                    firstId,
                    "create",
                    FormulaOleContract.WordOmmlMode,
                    @"x^2+1",
                    firstMathMl,
                    initial,
                    numbered: true,
                    originalMetadata: null),
                firstMathMl);

            ResetWordSelection(
                application,
                ref selection,
                document.Content.End - 1,
                document.Content.End - 1);
            service.InsertOle(
                CreateWordSession(
                    secondId,
                    "create",
                    FormulaOleContract.NativeOleMode,
                    @"\frac{a}{b}=c",
                    secondMathMl,
                    updated,
                    numbered: true,
                    originalMetadata: null),
                updated.PngPath,
                updated.EmfPath);

            AssertEqual(
                "2",
                WordEquationNumbering.Reconcile(document).ToString(),
                "Mixed OMML/OLE document did not expose two numbered formulas.");
            var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            AssertTargetNumbers(
                targets,
                (firstId, "1"),
                (secondId, "2"));

            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, firstId)
                ?? throw new InvalidOperationException("Mixed numbering OMML formula bookmark is missing.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertCleanFormulaParagraph(range, "OMML numbered formula");
            Release(range);
            range = null;
            Release(bookmark);
            bookmark = null;

            shape = FindWordFormula(document, secondId)
                ?? throw new InvalidOperationException("Mixed numbering OLE formula is missing.");
            range = shape.Range;
            AssertCleanFormulaParagraph(range, "OLE numbered formula");
            var oleSelection = ReadSelectedFormula(application, service, range);
            Assert(
                WordDoubleClickRouting.ShouldOpenVisualTeX(oleSelection),
                "Mixed numbering OLE formula double-click would not open the formula editor.");
            Release(range);
            range = null;
            Release(shape);
            shape = null;

            ResetWordSelection(
                application,
                ref selection,
                document.Content.End - 1,
                document.Content.End - 1);
            selection!.TypeParagraph();
            var firstTarget = targets.Single(target => target.FormulaId == firstId);
            var secondTarget = targets.Single(target => target.FormulaId == secondId);
            WordEquationNumbering.InsertEquationReference(
                document,
                selection!,
                firstTarget,
                EquationReferenceStyle.Parenthesized);
            selection!.TypeText(" ");
            WordEquationNumbering.InsertEquationReference(
                document,
                selection!,
                secondTarget,
                EquationReferenceStyle.Parenthesized);
            AssertReferenceResult(document, firstId, "1");
            AssertReferenceResult(document, secondId, "2");

            // A real user often inserts a new equation above existing content.
            // Existing REF fields must follow the original formula identities.
            ResetWordSelection(application, ref selection, 0, 0);
            service.InsertOmml(
                CreateWordSession(
                    insertedId,
                    "create",
                    FormulaOleContract.WordOmmlMode,
                    @"\sqrt{p+q}",
                    insertedMathMl,
                    initial,
                    numbered: true,
                    originalMetadata: null),
                insertedMathMl);
            WordEquationNumbering.Reconcile(document);
            targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            AssertTargetNumbers(
                targets,
                (insertedId, "1"),
                (firstId, "2"),
                (secondId, "3"));
            AssertReferenceResult(document, firstId, "2");
            AssertReferenceResult(document, secondId, "3");

            // Reproduce the real failure mode: the user selects only the native
            // OMath and presses Delete. Word removes the equation but leaves the
            // collapsed VisualTeX bookmark/custom XML part behind. Reconcile
            // must detect that stale anchor and renumber every surviving field.
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, insertedId)
                ?? throw new InvalidOperationException("Inserted OMML formula bookmark is missing before delete.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            range.Delete();
            Release(range);
            range = null;
            Release(bookmark);
            bookmark = null;
            WordEquationNumbering.Reconcile(document);
            targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            AssertTargetNumbers(
                targets,
                (firstId, "1"),
                (secondId, "2"));
            AssertReferenceResult(document, firstId, "1");
            AssertReferenceResult(document, secondId, "2");
            Assert(
                WordOmmlFormulaStore.FindByFormulaId(document, insertedId) is null,
                "Deleting a numbered OMML formula left its metadata anchor behind.");
        }
        finally
        {
            Release(shape);
            Release(range);
            Release(bookmark);
            Release(selection);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertTargetNumbers(
        IReadOnlyList<EquationReferenceTarget> targets,
        params (string FormulaId, string Number)[] expected)
    {
        Assert(
            targets.Count == expected.Length,
            $"Expected {expected.Length} Word equation targets, actual {targets.Count}.");
        foreach (var item in expected)
        {
            var target = targets.SingleOrDefault(candidate => candidate.FormulaId == item.FormulaId)
                ?? throw new InvalidOperationException(
                    $"ACCEPTANCE FAILURE: Formula {item.FormulaId} is missing from Word cross-reference targets.");
            AssertEqual(
                item.Number,
                target.NumberText,
                $"Formula {item.FormulaId} has the wrong Word equation number.");
        }
    }

    private static void AssertReferenceResult(
        Word.Document document,
        string formulaId,
        string expectedNumber)
    {
        var bookmarkName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
        Word.Fields? fields = null;
        var matching = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    field = fields[index];
                    if (field.Type != Word.WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if ((code.Text ?? string.Empty).IndexOf(
                            bookmarkName,
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    field.Update();
                    result = field.Result;
                    AssertEqual(
                        expectedNumber,
                        (result.Text ?? string.Empty).Trim(),
                        $"REF field for formula {formulaId} did not update.");
                    matching++;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        Assert(matching >= 2,
            $"Expected both visible numbering and body reference fields for formula {formulaId}.");
    }

    private static void AssertCleanFormulaParagraph(Word.Range formulaRange, string context)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? format = null;
        Word.ListFormat? listFormat = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;
            listFormat = paragraphRange.ListFormat;
            Assert(
                listFormat.ListType == Word.WdListType.wdListNoNumbering,
                $"{context} retained bullet or numbering formatting.");
            Assert(format.PageBreakBefore == 0, $"{context} retained PageBreakBefore formatting.");
            Assert(format.KeepTogether == 0, $"{context} retained KeepTogether formatting.");
            Assert(format.KeepWithNext == 0, $"{context} retained KeepWithNext formatting.");
        }
        finally
        {
            Release(listFormat);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static OfficeSessionDocument CreateWordSession(
        string formulaId,
        string mode,
        string objectMode,
        string latex,
        string mathMl,
        PreviewSet preview,
        bool numbered,
        FormulaMetadata? originalMetadata)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString(),
            Mode = mode,
            Host = "word",
            FormulaId = formulaId,
            Title = "Word OMML/OLE acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = objectMode,
            Numbered = numbered,
            Dirty = true,
            OriginalMetadata = originalMetadata,
            ExportResult = new OfficeExportDocument
            {
                MathMl = mathMl,
                Width = preview.Width,
                Height = preview.Height,
                Baseline = preview.Height * 0.72f,
            },
        };
    }

    private static void ResetWordSelection(
        Word.Application application,
        ref Word.Selection? selection,
        int start,
        int end)
    {
        Release(selection);
        selection = application.Selection;
        selection.SetRange(start, end);
    }

    private static OfficeSelection ReadSelectedFormula(
        Word.Application application,
        WordFormulaService service,
        Word.Range selectedRange)
    {
        Word.Selection? selection = null;
        try
        {
            selection = application.Selection;
            selection.SetRange(selectedRange.Start, selectedRange.End);
            return service.ReadSelection(selection);
        }
        finally { Release(selection); }
    }

    private static Word.InlineShape? FindWordFormula(
        Word.Document document,
        string formulaId)
    {
        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is null
                        || !string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var result = shape;
                    shape = null;
                    return result;
                }
                finally { Release(shape); }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private static void AssertWordEquationNumberOutsideOmml(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Range? equationRange = null;
        Word.OMaths? equationMaths = null;
        Word.OMath? equation = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = WordEquationNumbering.EquationBookmarkName(formulaId);
            Assert(bookmarks.Exists(numberName), "Visible Word equation number bookmark is missing.");
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            var documentXml = XDocument.Parse(document.Content.WordOpenXML);
            var numberBookmarkNode = documentXml
                .Descendants()
                .FirstOrDefault(element =>
                    element.Name.LocalName == "bookmarkStart"
                    && element.Attributes().Any(attribute =>
                        attribute.Name.LocalName == "name"
                        && string.Equals(
                            attribute.Value,
                            numberName,
                            StringComparison.Ordinal)));
            Assert(numberBookmarkNode is not null, "Visible Word equation number bookmark is absent from OpenXML.");
            var bookmarkInsideMath = numberBookmarkNode!
                .Ancestors()
                .Any(element => element.Name.LocalName is "oMath" or "oMathPara");
            Assert(
                !bookmarkInsideMath,
                "Visible Word equation number bookmark is structurally inside the OMML object.");
            var numberText = numberRange.Text ?? string.Empty;
            Assert(
                numberText.StartsWith("\t(", StringComparison.Ordinal)
                    && numberText.EndsWith(")", StringComparison.Ordinal),
                $"Visible Word equation number has invalid text/layout: '{numberText}'.");

            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException("OMML formula bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            equationMaths = equationRange.OMaths;
            Assert(equationMaths.Count == 1, "Numbered OMML formula does not contain exactly one OMath.");
            equation = equationMaths[1];
            Assert(
                equation.Type == Word.WdOMathType.wdOMathInline,
                "Numbered OMML formula must use inline OMath inside the display paragraph.");
            Assert(
                numberRange.Start >= equationRange.End,
                "Visible Word equation number overlaps the OMML formula range.");
        }
        finally
        {
            Release(equation);
            Release(equationMaths);
            Release(equationRange);
            Release(formulaBookmark);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static int CountFieldCodes(Word.Document document, string marker)
    {
        Word.Fields? fields = null;
        var count = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if ((code.Text ?? string.Empty).IndexOf(
                            marker,
                            StringComparison.OrdinalIgnoreCase) >= 0)
                        count++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        return count;
    }

    private static void CreateWordDocument(
        string path,
        string formulaId,
        PreviewSet preview)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? range = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            Console.WriteLine("  Word: starting Application");
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            Console.WriteLine("  Word: creating document");
            document = application.Documents.Add();
            range = document.Content;
            range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Console.WriteLine("  Word: calling InlineShapes.AddOLEObject");
            shape = AddWordOleObject(document, range);
            Console.WriteLine("  Word: OLE object inserted");
            var metadata = CreateMetadata(formulaId, preview, @"x^2+\frac{1}{y}", "initial");
            format = shape.OLEFormat;
            AssertEqual(FormulaOleContract.ProgId, format.ProgID, "Word inserted the wrong OLE class.");
            Console.WriteLine("  Word: acquiring custom interface");
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            var native = (IVisualTeXFormulaObject)oleObject;
            Console.WriteLine("  Word: initializing JSON/EMF/PNG");
            FormulaOleInterop.Initialize(native, metadata, preview.EmfPath, preview.PngPath);
            AssertNativeMetadata(native, formulaId);
            Console.WriteLine("  Word: initialization completed");
            shape.Width = preview.Width * 0.75f;
            Console.WriteLine("  Word: saving DOCX");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine("  Word: DOCX saved");
        }
        finally
        {
            Release(oleObject);
            Release(format);
            Release(shape);
            Release(range);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static Word.InlineShape AddWordOleObject(
        Word.Document document,
        Word.Range range) =>
        document.InlineShapes.AddOLEObject(
            ClassType: FormulaOleContract.ProgId,
            LinkToFile: false,
            DisplayAsIcon: false,
            Range: range);

    private static void CreatePowerPointDocument(
        string path,
        string formulaId,
        PreviewSet preview)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            application = new PowerPoint.Application();
            presentation = application.Presentations.Add(Office.MsoTriState.msoFalse);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            shape = slide.Shapes.AddOLEObject(
                90f,
                80f,
                preview.Width * 0.75f,
                preview.Height * 0.75f,
                FormulaOleContract.ProgId,
                string.Empty,
                Office.MsoTriState.msoFalse,
                string.Empty,
                0,
                string.Empty,
                Office.MsoTriState.msoFalse);
            var metadata = CreateMetadata(formulaId, preview, @"\int_0^1 f(x)\,dx", "initial");
            format = shape.OLEFormat;
            AssertEqual(
                FormulaOleContract.ProgId,
                format.ProgID,
                "PowerPoint inserted the wrong OLE class.");
            oleObject = format.Object;
            var native = (IVisualTeXFormulaObject)oleObject;
            FormulaOleInterop.Initialize(native, metadata, preview.EmfPath, preview.PngPath);
            AssertNativeMetadata(native, formulaId);
            shape.Name = $"VisualTeX_{formulaId}";
            presentation.SaveAs(
                path,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                Office.MsoTriState.msoFalse);
        }
        finally
        {
            Release(oleObject);
            Release(format);
            Release(shape);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void VerifyPowerPointPictureToOleConversion(
        string path,
        string formulaId,
        PreviewSet originalPreview,
        PreviewSet convertedPreview)
    {
        const float originalLeft = 60f;
        const float originalTop = 80f;
        const float originalWidth = 600f;
        const string formulaLatex = @"\alpha\beta d f d f d f d f d f d f aaaaabbbbb";
        var originalHeight = originalWidth
            * originalPreview.Height
            / (float)originalPreview.Width;
        // Export at high resolution so one-pixel antialiasing differences do
        // not masquerade as a visible horizontal/vertical squeeze.
        var exportWidth = 4800;
        var exportHeight = Math.Max(
            1,
            (int)Math.Round(exportWidth * originalHeight / originalWidth));
        var artifactDirectory = Path.GetDirectoryName(path)
            ?? throw new InvalidOperationException("PowerPoint acceptance path has no parent directory.");
        var pictureRenderPath = Path.Combine(
            artifactDirectory,
            "PowerPoint-Picture-Before-OLE.png");
        var oleRenderPath = Path.Combine(
            artifactDirectory,
            "PowerPoint-OLE-After-Conversion.png");
        var reopenedRenderPath = Path.Combine(
            artifactDirectory,
            "PowerPoint-OLE-After-Reopen.png");

        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? picture = null;
        PowerPoint.Tags? tags = null;
        PowerPoint.Shape? converted = null;
        PowerPoint.OLEFormat? format = null;
        PowerPoint.Presentation? reopened = null;
        PowerPoint.Slide? reopenedSlide = null;
        PowerPoint.Shape? reopenedShape = null;
        try
        {
            application = new PowerPoint.Application
            {
                Visible = Office.MsoTriState.msoTrue,
            };
            presentation = application.Presentations.Add(Office.MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            picture = slide.Shapes.AddPicture(
                originalPreview.PngPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                originalLeft,
                originalTop,
                originalWidth,
                originalHeight);

            var originalMetadata = CreateMetadata(
                formulaId,
                originalPreview,
                formulaLatex,
                "picture-source");
            var encoded = FormulaMetadataCodec.Encode(originalMetadata);
            picture.Name = $"VisualTeX_{formulaId}";
            picture.AlternativeText = encoded;
            tags = picture.Tags;
            tags.Add("VisualTeXFormulaId", formulaId);
            tags.Add("VisualTeXMetadata", encoded);
            picture.Export(
                pictureRenderPath,
                PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                exportWidth,
                exportHeight,
                PowerPoint.PpExportMode.ppScaleXY);

            var expected = (Width: originalWidth, Height: originalHeight);
            var originalCenterX = originalLeft + originalWidth / 2f;
            var originalCenterY = originalTop + originalHeight / 2f;

            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString(),
                Mode = "edit",
                Host = "powerpoint",
                FormulaId = formulaId,
                Title = "PowerPoint picture to OLE acceptance",
                Lines = new List<FormulaLine>
                {
                    new()
                    {
                        Id = Guid.NewGuid().ToString(),
                        Latex = formulaLatex,
                    },
                },
                CodeFormat = "latex",
                DisplayMode = "block",
                ObjectMode = FormulaOleContract.NativeOleMode,
                Numbered = false,
                Dirty = true,
                SourceObjectId = picture.Name,
                OriginalMetadata = originalMetadata,
                ExportResult = new OfficeExportDocument
                {
                    Width = convertedPreview.Width,
                    Height = convertedPreview.Height,
                    Baseline = convertedPreview.Height * 0.72f,
                },
            };

            var service = new PowerPointFormulaService(application);
            service.ReplaceOle(
                session,
                convertedPreview.PngPath,
                convertedPreview.EmfPath);

            Assert(slide.Shapes.Count == 1, "PowerPoint picture to OLE conversion left a duplicate shape.");
            converted = slide.Shapes[1];
            Assert(
                converted.Type == Office.MsoShapeType.msoEmbeddedOLEObject,
                "PowerPoint picture to OLE conversion did not create an embedded OLE object.");
            format = converted.OLEFormat;
            AssertEqual(
                FormulaOleContract.ProgId,
                format.ProgID,
                "PowerPoint picture conversion created the wrong OLE class.");

            Console.WriteLine(
                $"  PowerPoint conversion geometry: expected={expected.Width:0.###}x{expected.Height:0.###}, "
                + $"actual={converted.Width:0.###}x{converted.Height:0.###}, "
                + $"left/top={converted.Left:0.###}/{converted.Top:0.###}, "
                + $"center={converted.Left + converted.Width / 2f:0.###}/{converted.Top + converted.Height / 2f:0.###}");
            AssertClose(expected.Width, converted.Width, 0.75f,
                "PowerPoint converted OLE width does not match the formula's natural aspect ratio.");
            AssertClose(expected.Height, converted.Height, 0.75f,
                "PowerPoint converted OLE height changed the apparent formula font size.");
            AssertClose(
                convertedPreview.Width / (float)convertedPreview.Height,
                converted.Width / converted.Height,
                0.03f,
                "PowerPoint converted OLE formula is visually flattened or stretched.");
            AssertClose(
                originalCenterX,
                converted.Left + converted.Width / 2f,
                0.75f,
                "PowerPoint picture to OLE conversion moved the formula horizontally.");
            AssertClose(
                originalCenterY,
                converted.Top + converted.Height / 2f,
                0.75f,
                "PowerPoint picture to OLE conversion moved the formula vertically.");
            Assert(
                converted.Height >= originalHeight * 0.95f,
                "PowerPoint converted OLE formula became noticeably flatter than the source picture.");
            converted.Export(
                oleRenderPath,
                PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                exportWidth,
                exportHeight,
                PowerPoint.PpExportMode.ppScaleXY);
            AssertRenderedFormulaBoundsEquivalent(
                pictureRenderPath,
                oleRenderPath,
                "PowerPoint picture→OLE conversion");

            presentation.SaveAs(
                path,
                PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                Office.MsoTriState.msoFalse);
            presentation.Close();
            Release(format);
            format = null;
            Release(converted);
            converted = null;
            Release(tags);
            tags = null;
            Release(picture);
            picture = null;
            Release(slide);
            slide = null;
            Release(presentation);
            presentation = null;

            reopened = application.Presentations.Open(
                path,
                Office.MsoTriState.msoTrue,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse);
            reopenedSlide = reopened.Slides[1];
            reopenedShape = reopenedSlide.Shapes[1];
            Assert(
                reopenedShape.Type == Office.MsoShapeType.msoEmbeddedOLEObject,
                "Saved PowerPoint conversion did not remain an OLE object.");
            AssertClose(expected.Width, reopenedShape.Width, 0.75f,
                "PowerPoint save/reopen changed the converted OLE width.");
            AssertClose(expected.Height, reopenedShape.Height, 0.75f,
                "PowerPoint save/reopen changed the converted OLE height.");
            reopenedShape.Export(
                reopenedRenderPath,
                PowerPoint.PpShapeFormat.ppShapeFormatPNG,
                exportWidth,
                exportHeight,
                PowerPoint.PpExportMode.ppScaleXY);
            AssertRenderedFormulaBoundsEquivalent(
                pictureRenderPath,
                reopenedRenderPath,
                "PowerPoint saved/reopened OLE cache");
        }
        finally
        {
            Release(reopenedShape);
            Release(reopenedSlide);
            if (reopened is not null)
            {
                try { reopened.Close(); } catch { }
            }
            Release(reopened);
            Release(format);
            Release(converted);
            Release(tags);
            Release(picture);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void ReportIndependentShapeExports(
        string sourcePath,
        string convertedPath,
        string context)
    {
        using var source = new Bitmap(sourcePath);
        using var converted = new Bitmap(convertedPath);
        var sourceBounds = FindDarkPixelBounds(source);
        var convertedBounds = FindDarkPixelBounds(converted);
        Console.WriteLine(
            $"  {context} independent exports: source bitmap={source.Width}x{source.Height}, "
            + $"dark={sourceBounds.Width}x{sourceBounds.Height}; "
            + $"OLE bitmap={converted.Width}x{converted.Height}, "
            + $"dark={convertedBounds.Width}x{convertedBounds.Height}.");
    }

    private static void AssertSameSlidePairBounds(
        string screenshotPath,
        IReadOnlyList<ComparisonRow> rows,
        float slideWidth,
        float slideHeight,
        string context)
    {
        using var screenshot = new Bitmap(screenshotPath);
        var scaleX = screenshot.Width / slideWidth;
        var scaleY = screenshot.Height / slideHeight;
        foreach (var row in rows)
        {
            var sourceRegion = Rectangle.Round(new RectangleF(
                row.SourceLeft * scaleX,
                row.Top * scaleY,
                row.Width * scaleX,
                row.Height * scaleY));
            var oleRegion = Rectangle.Round(new RectangleF(
                row.OleLeft * scaleX,
                row.Top * scaleY,
                row.Width * scaleX,
                row.Height * scaleY));
            sourceRegion.Intersect(new Rectangle(0, 0, screenshot.Width, screenshot.Height));
            oleRegion.Intersect(new Rectangle(0, 0, screenshot.Width, screenshot.Height));
            var sourceBounds = FindDarkPixelBounds(screenshot, sourceRegion);
            var oleBounds = FindDarkPixelBounds(screenshot, oleRegion);
            Assert(
                sourceBounds.Width > 0 && sourceBounds.Height > 0,
                $"{context}, formula {row.Index}: source crop contains no visible formula pixels.");
            Assert(
                oleBounds.Width > 0 && oleBounds.Height > 0,
                $"{context}, formula {row.Index}: OLE crop contains no visible formula pixels.");

            var widthScale = oleBounds.Width / (float)sourceBounds.Width;
            var heightScale = oleBounds.Height / (float)sourceBounds.Height;
            var aspectScale =
                (oleBounds.Width / (float)oleBounds.Height)
                / (sourceBounds.Width / (float)sourceBounds.Height);
            Console.WriteLine(
                $"  {context}, formula {row.Index}: source={sourceBounds.Width}x{sourceBounds.Height}; "
                + $"OLE={oleBounds.Width}x{oleBounds.Height}; widthScale={widthScale:0.###}; "
                + $"heightScale={heightScale:0.###}; aspectScale={aspectScale:0.###}.");
            Assert(
                widthScale >= 0.94f && widthScale <= 1.06f,
                $"{context}, formula {row.Index}: OLE content width changed ({widthScale:0.###}×)." );
            Assert(
                heightScale >= 0.92f && heightScale <= 1.08f,
                $"{context}, formula {row.Index}: OLE content height changed ({heightScale:0.###}×)." );
            Assert(
                aspectScale >= 0.94f && aspectScale <= 1.06f,
                $"{context}, formula {row.Index}: internal glyph aspect changed ({aspectScale:0.###}×)." );
        }
    }

    private static Rectangle FindDarkPixelBounds(Bitmap bitmap, Rectangle region)
    {
        var left = region.Right;
        var top = region.Bottom;
        var right = region.Left - 1;
        var bottom = region.Top - 1;
        for (var y = region.Top; y < region.Bottom; y++)
        {
            for (var x = region.Left; x < region.Right; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.A < 16) continue;
                if (pixel.R + pixel.G + pixel.B >= 660) continue;
                left = Math.Min(left, x);
                top = Math.Min(top, y);
                right = Math.Max(right, x);
                bottom = Math.Max(bottom, y);
            }
        }
        return right < left || bottom < top
            ? Rectangle.Empty
            : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
    }

    private static void AssertRenderedFormulaBoundsEquivalent(
        string sourcePath,
        string convertedPath,
        string context)
    {
        using var source = new Bitmap(sourcePath);
        using var converted = new Bitmap(convertedPath);
        var sourceBounds = FindDarkPixelBounds(source);
        var convertedBounds = FindDarkPixelBounds(converted);
        Assert(
            sourceBounds.Width > 0 && sourceBounds.Height > 0,
            $"{context}: source picture export contains no visible formula pixels.");
        Assert(
            convertedBounds.Width > 0 && convertedBounds.Height > 0,
            $"{context}: OLE export contains no visible formula pixels.");

        var widthScale = convertedBounds.Width / (float)sourceBounds.Width;
        var heightScale = convertedBounds.Height / (float)sourceBounds.Height;
        var sourceAspect = sourceBounds.Width / (float)sourceBounds.Height;
        var convertedAspect = convertedBounds.Width / (float)convertedBounds.Height;
        var aspectScale = convertedAspect / sourceAspect;
        Console.WriteLine(
            $"  {context} visible pixels: source={sourceBounds.X},{sourceBounds.Y} "
            + $"{sourceBounds.Width}x{sourceBounds.Height}; converted={convertedBounds.X},{convertedBounds.Y} "
            + $"{convertedBounds.Width}x{convertedBounds.Height}; "
            + $"widthScale={widthScale:0.###}, heightScale={heightScale:0.###}, aspectScale={aspectScale:0.###}.");

        Assert(
            widthScale >= 0.94f && widthScale <= 1.06f,
            $"{context}: formula content became horizontally squeezed or widened ({widthScale:0.###}×)." );
        Assert(
            heightScale >= 0.92f && heightScale <= 1.08f,
            $"{context}: formula content became vertically flattened or enlarged ({heightScale:0.###}×)." );
        Assert(
            aspectScale >= 0.94f && aspectScale <= 1.06f,
            $"{context}: internal formula glyph aspect ratio changed ({aspectScale:0.###}×)." );
    }

    private static Rectangle FindDarkPixelBounds(Bitmap bitmap)
    {
        var left = bitmap.Width;
        var top = bitmap.Height;
        var right = -1;
        var bottom = -1;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = 0; x < bitmap.Width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.A < 16) continue;
                if (pixel.R + pixel.G + pixel.B >= 660) continue;
                left = Math.Min(left, x);
                top = Math.Min(top, y);
                right = Math.Max(right, x);
                bottom = Math.Max(bottom, y);
            }
        }
        return right < left || bottom < top
            ? Rectangle.Empty
            : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
    }

    private static void VerifyWordCachedPreviewOffline(string path, string formulaId)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Open(path, ReadOnly: true, Visible: false);
            shapes = document.InlineShapes;
            Assert(shapes.Count == 1, "Word offline reopen lost the embedded OLE object.");
            shape = shapes[1];
            Assert(shape.Width > 1 && shape.Height > 1, "Word offline cached preview has invalid size.");
            Assert(
                shape.Type == Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject,
                "Word offline reopen changed the embedded OLE object into another shape type.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void VerifyPowerPointCachedPreviewOffline(string path, string formulaId)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shapes? shapes = null;
        PowerPoint.Shape? shape = null;
        try
        {
            application = new PowerPoint.Application();
            presentation = application.Presentations.Open(
                path,
                Office.MsoTriState.msoTrue,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse);
            slide = presentation.Slides[1];
            shapes = slide.Shapes;
            Assert(shapes.Count == 1, "PowerPoint offline reopen lost the embedded OLE object.");
            shape = shapes[1];
            Assert(shape.Width > 1 && shape.Height > 1, "PowerPoint offline cached preview has invalid size.");
            AssertEqual($"VisualTeX_{formulaId}", shape.Name, "PowerPoint formula identity changed offline.");
            Assert(
                shape.Type == Office.MsoShapeType.msoEmbeddedOLEObject,
                "PowerPoint offline reopen changed the embedded OLE object into another shape type.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void UpdateAndVerifyWord(
        string path,
        string formulaId,
        PreviewSet preview)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.OLEFormat? format = null;
        object? oleObject = null;
        string? extractedPng = null;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Open(path, ReadOnly: false, Visible: false);
            shapes = document.InlineShapes;
            shape = shapes[1];
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            var native = (IVisualTeXFormulaObject)oleObject;
            var metadata = CreateMetadata(formulaId, preview, @"e^{i\pi}+1=0", "updated");
            FormulaOleInterop.Update(native, metadata, preview.EmfPath, preview.PngPath);
            AssertNativeMetadata(native, formulaId);
            document.Save();
            extractedPng = OlePngPreviewExtractor.MaterializePng(oleObject, formulaId);
            AssertPng(extractedPng, "Word updated OLE PNG cache");
        }
        finally
        {
            TryDelete(extractedPng);
            Release(oleObject);
            Release(format);
            Release(shape);
            Release(shapes);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdSaveChanges); } catch { }
            }
            Release(document);
            if (application is not null)
            {
                try { application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void UpdateAndVerifyPowerPoint(
        string path,
        string formulaId,
        PreviewSet preview)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.OLEFormat? format = null;
        object? oleObject = null;
        string? extractedPng = null;
        try
        {
            application = new PowerPoint.Application();
            presentation = application.Presentations.Open(
                path,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse);
            slide = presentation.Slides[1];
            shape = slide.Shapes[1];
            format = shape.OLEFormat;
            oleObject = format.Object;
            var native = (IVisualTeXFormulaObject)oleObject;
            var metadata = CreateMetadata(
                formulaId,
                preview,
                @"\sum_{n=1}^{\infty}\frac{1}{n^2}",
                "updated");
            FormulaOleInterop.Update(native, metadata, preview.EmfPath, preview.PngPath);
            AssertNativeMetadata(native, formulaId);
            presentation.Save();
            extractedPng = OlePngPreviewExtractor.MaterializePng(oleObject, formulaId);
            AssertPng(extractedPng, "PowerPoint updated OLE PNG cache");
        }
        finally
        {
            TryDelete(extractedPng);
            Release(oleObject);
            Release(format);
            Release(shape);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Close(); } catch { }
            }
            Release(presentation);
            if (application is not null)
            {
                try { application.Quit(); } catch { }
            }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertNativeMetadata(
        IVisualTeXFormulaObject formula,
        string formulaId)
    {
        var metadata = FormulaOleInterop.ReadMetadata(formula);
        AssertEqual(formulaId, metadata.FormulaId, "Persisted formula UUID changed.");
    }

    private static void AssertPng(string path, string context)
    {
        var bytes = File.ReadAllBytes(path);
        Assert(
            bytes.Length >= 8
            && bytes[0] == 137
            && bytes[1] == 80
            && bytes[2] == 78
            && bytes[3] == 71,
            context + " is invalid.");
    }

    private static void RunRegistration(string serverPath, string argument)
    {
        using var process = Process.Start(new ProcessStartInfo(serverPath, argument)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
        }) ?? throw new InvalidOperationException("Unable to start the Formula OLE LocalServer registration command.");
        if (!process.WaitForExit(15000))
        {
            try { process.Kill(); } catch { }
            throw new TimeoutException($"Formula OLE registration command timed out: {argument}");
        }
        if (process.ExitCode != 0)
            throw new InvalidOperationException(
                $"Formula OLE registration command {argument} failed with exit code {process.ExitCode}.");
    }

    private static bool HasExistingRegistration()
    {
        var view = Environment.Is64BitProcess
            ? RegistryView.Registry64
            : RegistryView.Registry32;
        using var currentUser = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
        using var classes = currentUser.OpenSubKey("Software\\Classes");
        using var progId = classes?.OpenSubKey(FormulaOleContract.ProgId);
        using var clsid = classes?.OpenSubKey($"CLSID\\{FormulaClsid}");
        using var typeLib = classes?.OpenSubKey($"TypeLib\\{TypeLibraryId}");
        var found = progId is not null || clsid is not null || typeLib is not null;
        if (found)
        {
            Console.Error.WriteLine(
                $"Registration precheck ({view}): ProgID={progId is not null}, CLSID={clsid is not null}, TypeLib={typeLib is not null}.");
        }
        return found;
    }

    private static void AssertRegistrationPresent(string serverPath)
    {
        using var key = Registry.CurrentUser.OpenSubKey(
            $"Software\\Classes\\CLSID\\{FormulaClsid}\\LocalServer32")
            ?? throw new InvalidOperationException("Formula OLE LocalServer32 registration is missing.");
        var registered = Convert.ToString(key.GetValue(null))?.Trim('"');
        AssertEqual(
            Path.GetFullPath(serverPath),
            registered is null ? string.Empty : Path.GetFullPath(registered),
            "Formula OLE LocalServer32 points to the wrong executable.");
    }

    private static void AssertRegistrationAbsent()
    {
        Assert(!HasExistingRegistration(), "Formula OLE registration remained after per-user unregistration.");
    }

    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException("ACCEPTANCE FAILURE: " + message);
    }

    private static void AssertEqual(string expected, string actual, string message)
    {
        if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"ACCEPTANCE FAILURE: {message} Expected '{expected}', actual '{actual}'.");
    }

    private static void AssertClose(
        float expected,
        float actual,
        float tolerance,
        string message)
    {
        if (float.IsNaN(actual)
            || float.IsInfinity(actual)
            || Math.Abs(expected - actual) > tolerance)
            throw new InvalidOperationException(
                $"ACCEPTANCE FAILURE: {message} Expected {expected:0.###}, actual {actual:0.###}, tolerance {tolerance:0.###}.");
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }

    private static void ForceComCleanup()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        Thread.Sleep(150);
    }

    private static void TryDelete(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        try { File.Delete(path); } catch { }
    }

    private sealed class ComparisonRow
    {
        public ComparisonRow(
            int index,
            float sourceLeft,
            float oleLeft,
            float top,
            float width,
            float height)
        {
            Index = index;
            SourceLeft = sourceLeft;
            OleLeft = oleLeft;
            Top = top;
            Width = width;
            Height = height;
        }

        public int Index { get; }
        public float SourceLeft { get; }
        public float OleLeft { get; }
        public float Top { get; }
        public float Width { get; }
        public float Height { get; }
    }

    private sealed class PreviewSet
    {
        public PreviewSet(
            string svgPath,
            string emfPath,
            string pngPath,
            int width,
            int height)
        {
            SvgPath = svgPath;
            EmfPath = emfPath;
            PngPath = pngPath;
            Width = width;
            Height = height;
        }

        public string SvgPath { get; }
        public string EmfPath { get; }
        public string PngPath { get; }
        public int Width { get; }
        public int Height { get; }
    }
}

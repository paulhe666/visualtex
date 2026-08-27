using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO.Compression;
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
    private static readonly (string Command, string Symbol, string CodePoint)[]
        ExtendedIntegralOperators =
        {
            ("oiint", "∯", "222F"),
            ("oiiint", "∰", "2230"),
            ("intclockwise", "∱", "2231"),
            ("varointclockwise", "∲", "2232"),
            ("ointctrclockwise", "∳", "2233"),
            ("sumint", "⨋", "2A0B"),
            ("iiiint", "⨌", "2A0C"),
            ("intbar", "⨍", "2A0D"),
            ("intBar", "⨎", "2A0E"),
            ("fint", "⨏", "2A0F"),
            ("cirfnint", "⨐", "2A10"),
            ("awint", "⨑", "2A11"),
            ("intctrclockwise", "⨑", "2A11"),
            ("rppolint", "⨒", "2A12"),
            ("scpolint", "⨓", "2A13"),
            ("npolint", "⨔", "2A14"),
            ("pointint", "⨕", "2A15"),
            ("quatint", "⨖", "2A16"),
            ("intlarhk", "⨗", "2A17"),
            ("intx", "⨘", "2A18"),
            ("intcap", "⨙", "2A19"),
            ("intcup", "⨚", "2A1A"),
            ("upint", "⨛", "2A1B"),
            ("lowint", "⨜", "2A1C"),
        };
    private static readonly (string Command, string Accent)[] NativeAccentOperators =
        {
            ("hat", "\u0302"),
            ("widehat", "\u0302"),
            ("tilde", "\u0303"),
            ("widetilde", "\u0303"),
            ("vec", "\u20D7"),
            ("overrightarrow", "\u20D7"),
            ("overleftarrow", "\u20D6"),
            ("dot", "\u0307"),
            ("ddot", "\u0308"),
            ("check", "\u030C"),
            ("breve", "\u0306"),
            ("acute", "\u0301"),
            ("grave", "\u0300"),
            ("mathring", "\u030A"),
        };

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr window, out NativeRect rectangle);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr window);

    [STAThread]
    private static string ResolveArtifactRoot(string? argument)
    {
        var baseRoot = Path.Combine(
            Path.GetTempPath(),
            "VisualTeX",
            "acceptance",
            "native-office-ole");
        if (string.IsNullOrWhiteSpace(argument))
            return Path.Combine(baseRoot, Guid.NewGuid().ToString("N"));

        var expanded = Environment.ExpandEnvironmentVariables(
            argument!.Trim().Trim('"'));
        if (Path.IsPathRooted(expanded))
            return Path.GetFullPath(expanded);
        var segments = expanded
            .Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar)
            .Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries)
            .Where(segment => !string.Equals(segment, ".", StringComparison.Ordinal))
            .ToArray();
        if (segments.Any(segment => string.Equals(segment, "..", StringComparison.Ordinal)))
            throw new InvalidDataException(
                "A relative native Office acceptance path cannot contain '..'. Use an absolute path for an explicit external destination.");
        return Path.Combine(
            baseRoot,
            segments.Length == 0 ? Guid.NewGuid().ToString("N") : Path.Combine(segments));
    }

    private static int Main(string[] args)
    {
        if (args.Length >= 1
            && string.Equals(
                args[0],
                "--installed-word-ole-numbering-promotion",
                StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var promotionArtifactRoot = args.Length >= 2
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"VisualTeX-Word-OLE-Numbering-Promotion-{Guid.NewGuid():N}");
                Directory.CreateDirectory(promotionArtifactRoot);
                var promotionPreviewRoot = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "VisualTeX",
                    "office",
                    "temp");
                Directory.CreateDirectory(promotionPreviewRoot);
                var promotionPreviewFormulaId = Guid.NewGuid().ToString();
                var promotionInitial = CreatePreviewSet(
                    promotionPreviewRoot,
                    promotionPreviewFormulaId,
                    "unnumbered",
                    420,
                    130);
                var promotionUpdated = CreatePreviewSet(
                    promotionPreviewRoot,
                    promotionPreviewFormulaId,
                    "numbered",
                    520,
                    150);
                VerifyWordOleNumberingPromotion(
                    promotionArtifactRoot,
                    promotionInitial,
                    promotionUpdated);
                Console.WriteLine(
                    "Installed Word OLE numbering promotion acceptance passed: "
                    + "an unnumbered display OLE was edited to Numbered=true without "
                    + "reusing a deleted InlineShape COM object.");
                Console.WriteLine($"Artifacts: {promotionArtifactRoot}");
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(args[0], "--native-word-number-size-probe", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var probeArtifactRoot = args.Length >= 2
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"VisualTeX-Native-Word-Number-Size-{Guid.NewGuid():N}");
                return RunNativeWordNumberSizeProbe(probeArtifactRoot);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(args[0], "--installed-word-bulk-import", StringComparison.OrdinalIgnoreCase))
        {
            if (args.Length < 4)
            {
                Console.Error.WriteLine(
                    "Usage: VisualTeX.NativeOfficeOleAcceptance.exe --installed-word-bulk-import <source.md|source.tex> <omml|ole> <artifact-directory>");
                return 2;
            }
            try
            {
                var bulkSourcePath = Path.GetFullPath(args[1]);
                var bulkObjectMode = args[2];
                var bulkArtifactRoot = Path.GetFullPath(args[3]);
                // Set acceptance inputs before entering any method that references
                // Word COM types. Office can bootstrap WINWORD while the acceptance
                // method is being JIT-compiled; that process must inherit these
                // variables or the add-in will open the interactive import dialog.
                Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
                Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", bulkSourcePath);
                Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "auto");
                Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", bulkObjectMode);
                Directory.CreateDirectory(bulkArtifactRoot);
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG",
                    Path.Combine(bulkArtifactRoot, "addin-bulk-import.log"));
                return RunInstalledWordBulkImportAcceptance(
                    bulkSourcePath,
                    bulkObjectMode,
                    bulkArtifactRoot);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                try
                {
                    var bulkArtifactRoot = Path.GetFullPath(args[3]);
                    Directory.CreateDirectory(bulkArtifactRoot);
                    File.WriteAllText(
                        Path.Combine(bulkArtifactRoot, "acceptance-error.txt"),
                        error.ToString(),
                        Encoding.UTF8);
                }
                catch { }
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(
                args[0],
                "--installed-word-inline-roundtrip-anchor",
                StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var roundTripArtifactRoot = args.Length >= 2
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"VisualTeX-Word-Inline-RoundTrip-{Guid.NewGuid():N}");
                return RunInstalledWordInlineRoundTripAnchorAcceptance(
                    roundTripArtifactRoot);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(args[0], "--installed-font-size", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var fontSizeArtifactRoot = args.Length >= 2
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"VisualTeX-Formula-Font-Size-{Guid.NewGuid():N}");
                return RunInstalledFormulaFontSizeAcceptance(fontSizeArtifactRoot);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(args[0], "--installed-mixed-text-vector", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var mixedTextArtifactRoot = args.Length >= 2
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"VisualTeX-Mixed-Text-Vector-{Guid.NewGuid():N}");
                return RunInstalledMixedTextVectorAcceptance(mixedTextArtifactRoot);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(args[0], "--installed-tall-matrix-vector", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var tallMatrixArtifactRoot = args.Length >= 2
                    ? Path.GetFullPath(args[1])
                    : Path.Combine(
                        Path.GetTempPath(),
                        $"VisualTeX-Tall-Matrix-Vector-{Guid.NewGuid():N}");
                return RunInstalledTallMatrixVectorAcceptance(tallMatrixArtifactRoot);
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

        if (args.Length >= 1
            && string.Equals(args[0], "--installed-word-baseline-visual", StringComparison.OrdinalIgnoreCase))
        {
            if (args.Length < 3)
            {
                Console.Error.WriteLine(
                    "Usage: VisualTeX.NativeOfficeOleAcceptance.exe --installed-word-baseline-visual <SVG-fixture> <artifact-directory>");
                return 2;
            }
            try
            {
                return RunInstalledWordBaselineVisualComparison(
                    Path.GetFullPath(args[1]),
                    Path.GetFullPath(args[2]));
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }

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

        var artifactRoot = ResolveArtifactRoot(args.Length >= 2 ? args[1] : null);
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
            VerifyWordOleNumberingPromotion(artifactRoot, initial, updated);
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

    private static int RunInstalledWordBaselineVisualComparison(
        string fixturePath,
        string artifactRoot)
    {
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "The installed Word baseline visual mode requires the formally installed VisualTeX OLE registration.");
        if (!File.Exists(fixturePath))
            throw new FileNotFoundException("Word baseline SVG fixture does not exist.", fixturePath);

        Directory.CreateDirectory(artifactRoot);
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(previewRoot);
        var preview = CreatePreviewSetFromSvgFixture(
            previewRoot,
            fixturePath,
            "word-baseline-visual");
        var exportedBaseline = ReadSvgFixtureBaseline(fixturePath);
        var formulaText = fixturePath.IndexOf(
                "no-descender",
                StringComparison.OrdinalIgnoreCase) >= 0
            ? "dfdfdfdf"
            : "dfgfdgfdgfgfgf";
        const float renderFontSizePt = 14f;
        const float targetFontSizePt = 42f;
        var variants = new[]
        {
            (Label: "N", HasBaselineMetadata: true),
            (Label: "L", HasBaselineMetadata: false),
        };

        var documentPath = Path.Combine(artifactRoot, "VisualTeX-Word-Baseline-Visual.docx");
        var pdfPath = Path.Combine(artifactRoot, "VisualTeX-Word-Baseline-Visual.pdf");
        var screenshotPath = Path.Combine(artifactRoot, "VisualTeX-Word-Baseline-Visual.png");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Window? window = null;
        try
        {
            application = new Word.Application
            {
                Visible = true,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            window = application.ActiveWindow;
            window.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            window.View.Type = Word.WdViewType.wdPrintView;
            window.View.Zoom.Percentage = 100;
            document.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            document.PageSetup.LeftMargin = 36f;
            document.PageSetup.RightMargin = 36f;
            document.PageSetup.TopMargin = 30f;
            document.PageSetup.BottomMargin = 30f;
            selection = application.Selection;
            var service = new WordFormulaService(application);

            selection.Font.Size = 15f;
            selection.Font.Bold = 1;
            selection.TypeText("VisualTeX Word 行内公式基线实机对照");
            selection.TypeParagraph();
            selection.Font.Bold = 0;
            selection.Font.Size = 11f;
            selection.TypeText(
                $"R=Word 原生 OMML；N=完整元数据 OLE；L=旧版缺少基线元数据 OLE；"
                + $"正文与目标公式均为 {targetFontSizePt:0.#} pt。");
            selection.TypeParagraph();

            var nativeMathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>" + formulaText + "</mi></math>";
            selection.Font.Size = targetFontSizePt;
            selection.Font.Position = 0;
            selection.TypeText("R中：");
            var nativeFormulaId = Guid.NewGuid().ToString();
            var nativeSession = CreateWordSession(
                nativeFormulaId,
                "create",
                FormulaOleContract.WordOmmlMode,
                formulaText,
                nativeMathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: targetFontSizePt);
            nativeSession.DisplayMode = "inline";
            service.InsertOmml(nativeSession, nativeMathMl);
            selection = application.Selection;
            selection.Font.Position = 0;
            selection.TypeParagraph();

            selection.Font.Size = renderFontSizePt;
            selection.Font.Position = 0;
            selection.TypeText("O中：");
            var originalFormulaId = Guid.NewGuid().ToString();
            var originalSession = CreateWordSession(
                originalFormulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                formulaText,
                nativeMathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: renderFontSizePt);
            originalSession.DisplayMode = "inline";
            originalSession.ExportResult!.Baseline = exportedBaseline;
            service.InsertOle(originalSession, preview.PngPath, preview.EmfPath);
            var originalShape = FindWordFormula(document, originalFormulaId)
                ?? throw new InvalidOperationException("Word baseline visual probe did not create its original-size OLE formula.");
            try
            {
                Console.WriteLine(
                    $"  O: width={originalShape.Width:0.###}pt; height={originalShape.Height:0.###}pt; "
                    + $"position={originalShape.Range.Font.Position:0.###}pt.");
            }
            finally { Release(originalShape); }
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.Font.Position = 0;
            selection.TypeParagraph();

            foreach (var variant in variants)
            {
                selection.Font.Bold = 0;
                selection.Font.Size = targetFontSizePt;
                selection.Font.Position = 0;
                selection.TypeText(variant.Label + "中：");

                var formulaId = Guid.NewGuid().ToString();
                var session = CreateWordSession(
                    formulaId,
                    "create",
                    FormulaOleContract.NativeOleMode,
                    formulaText,
                    nativeMathMl,
                    preview,
                    numbered: false,
                    originalMetadata: null,
                    fontSizePt: renderFontSizePt);
                session.DisplayMode = "inline";
                session.ExportResult!.Baseline = variant.HasBaselineMetadata
                    ? exportedBaseline
                    : null;
                service.InsertOle(session, preview.PngPath, preview.EmfPath);
                var shape = FindWordFormula(document, formulaId)
                    ?? throw new InvalidOperationException("Word baseline visual probe did not create its OLE formula.");
                try
                {
                    var initialPosition = shape.Range.Font.Position;
                    shape.Range.Select();
                    service.SetSelectedFormulaFontSize(targetFontSizePt);
                    Console.WriteLine(
                        $"  {variant.Label}: width={shape.Width:0.###}pt; height={shape.Height:0.###}pt; "
                        + $"initialPosition={initialPosition:0.###}pt; "
                        + $"finalPosition={shape.Range.Font.Position:0.###}pt.");
                }
                finally { Release(shape); }

                selection = application.Selection;
                selection.EndKey(Word.WdUnits.wdStory);
                selection.Font.Position = 0;
                selection.TypeParagraph();
            }

            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.ExportAsFixedFormat(
                pdfPath,
                Word.WdExportFormat.wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                Range: Word.WdExportRange.wdExportAllDocument,
                Item: Word.WdExportItem.wdExportDocumentContent,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);
            selection.EndKey(Word.WdUnits.wdStory);
            selection.Font.Position = 0;
            document.Activate();
            application.ScreenRefresh();
            Thread.Sleep(1200);
            CaptureWindow(window.Hwnd, screenshotPath);
            Console.WriteLine($"Word baseline visual DOCX: {documentPath}");
            Console.WriteLine($"Word baseline visual PDF: {pdfPath}");
            Console.WriteLine($"Word baseline visual screenshot: {screenshotPath}");
            return 0;
        }
        finally
        {
            Release(window);
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
            TryDelete(preview.SvgPath);
            TryDelete(preview.EmfPath);
            TryDelete(preview.PngPath);
        }
    }

    private static float ReadSvgFixtureBaseline(string fixturePath)
    {
        var document = XDocument.Load(fixturePath, LoadOptions.None);
        var root = document.Root
            ?? throw new InvalidDataException("Word baseline SVG fixture has no root element.");
        if (!double.TryParse(
                root.Attribute("height")?.Value,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var height)
            || height <= 0)
            throw new InvalidDataException("Word baseline SVG fixture height is invalid.");
        var viewBox = (root.Attribute("viewBox")?.Value ?? string.Empty)
            .Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
        if (viewBox.Length != 4
            || !double.TryParse(viewBox[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var y)
            || !double.TryParse(viewBox[3], NumberStyles.Float, CultureInfo.InvariantCulture, out var viewHeight)
            || viewHeight <= 0)
            throw new InvalidDataException("Word baseline SVG fixture viewBox is invalid.");
        return (float)Math.Max(0, Math.Min(height, -y / viewHeight * height));
    }

    private static NativeRect CaptureWindow(int windowHandle, string outputPath)
    {
        var handle = new IntPtr(windowHandle);
        SetForegroundWindow(handle);
        Thread.Sleep(300);
        if (!GetWindowRect(handle, out var rectangle))
            throw new InvalidOperationException("Unable to read the Word window bounds for visual capture.");
        var width = Math.Max(1, rectangle.Right - rectangle.Left);
        var height = Math.Max(1, rectangle.Bottom - rectangle.Top);
        using var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
        using var graphics = Graphics.FromImage(bitmap);
        graphics.CopyFromScreen(
            rectangle.Left,
            rectangle.Top,
            0,
            0,
            new Size(width, height),
            CopyPixelOperation.SourceCopy);
        bitmap.Save(outputPath, ImageFormat.Png);
        return rectangle;
    }

    private static void AnalyzeWordBaselineScreenshot(
        string screenshotPath,
        NativeRect capturedWindow,
        Word.Window window,
        IReadOnlyList<WordBaselineVisualRow> rows)
    {
        using var screenshot = new Bitmap(screenshotPath);
        foreach (var row in rows)
        {
            var textBounds = GetWordRangeScreenshotBounds(
                window,
                row.BodyRange,
                capturedWindow,
                screenshot.Size);
            var latinBounds = GetWordRangeScreenshotBounds(
                window,
                row.LatinRange,
                capturedWindow,
                screenshot.Size);
            var formulaBounds = GetWordRangeScreenshotBounds(
                window,
                row.FormulaRange,
                capturedWindow,
                screenshot.Size);
            if (textBounds.Width > 28) textBounds.Width -= 28;
            if (formulaBounds.Width > 28)
            {
                formulaBounds.X += 18;
                formulaBounds.Width -= 18;
            }
            var textPixels = FindWordInkPixelBounds(screenshot, textBounds);
            var latinPixels = FindWordInkPixelBounds(screenshot, latinBounds);
            var formulaPixels = FindWordInkPixelBounds(screenshot, formulaBounds);
            if (textPixels.IsEmpty || latinPixels.IsEmpty || formulaPixels.IsEmpty)
            {
                Console.WriteLine(
                    $"  {row.Label}: visual pixel bounds unavailable; text={textBounds}, latin={latinBounds}, formula={formulaBounds}.");
                continue;
            }
            var bottomDelta = formulaPixels.Bottom - textPixels.Bottom;
            var latinBottomDelta = formulaPixels.Bottom - latinPixels.Bottom;
            var centerDelta =
                (formulaPixels.Top + formulaPixels.Bottom) / 2f
                - (textPixels.Top + textPixels.Bottom) / 2f;
            Console.WriteLine(
                $"  {row.Label}: textPixels={textPixels}; latinPixels={latinPixels}; formulaPixels={formulaPixels}; "
                + $"formulaBottom-textBottom={bottomDelta}px; formulaBottom-latinBottom={latinBottomDelta}px; "
                + $"centerDelta={centerDelta:0.###}px.");
        }
    }

    private static Rectangle FindWordInkPixelBounds(Bitmap bitmap, Rectangle region)
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
                if (pixel.A < 32) continue;
                if (pixel.R + pixel.G + pixel.B >= 480) continue;
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

    private static Rectangle GetWordRangeScreenshotBounds(
        Word.Window window,
        Word.Range range,
        NativeRect capturedWindow,
        Size screenshotSize)
    {
        int left;
        int top;
        int width;
        int height;
        window.GetPoint(out left, out top, out width, out height, range);
        var rectangle = new Rectangle(
            left - capturedWindow.Left,
            top - capturedWindow.Top,
            Math.Max(1, width),
            Math.Max(1, height));
        rectangle.Inflate(4, 4);
        rectangle.Intersect(new Rectangle(Point.Empty, screenshotSize));
        return rectangle;
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

    private static int RunNativeWordNumberSizeProbe(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "Word-Native-Equation-Number-Font-Size-Probe.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertionRange = null;
        Word.Table? table = null;
        Word.Cell? formulaCell = null;
        Word.Cell? numberCell = null;
        Word.Range? formulaInput = null;
        Word.OMaths? maths = null;
        Word.Range? addedMathRange = null;
        Word.OMath? math = null;
        Word.Range? formulaRange = null;
        Word.Font? formulaFont = null;
        Word.Range? numberCellRange = null;
        Word.Range? fieldRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? numberRange = null;
        Word.Font? numberFont = null;
        try
        {
            application = new Word.Application
            {
                Visible = false,
                DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
            };
            document = application.Documents.Add();
            insertionRange = document.Range(0, 0);
            table = document.Tables.Add(insertionRange, 1, 3);
            table.AllowAutoFit = false;
            table.Borders.Enable = 0;
            formulaCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            formulaCell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            numberCell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            formulaInput = formulaCell.Range.Duplicate;
            formulaInput.End = Math.Max(formulaInput.Start, formulaInput.End - 1);
            formulaInput.Text = "x^2+y^2=z^2";
            formulaInput.Font.Size = 14f;
            maths = formulaInput.OMaths;
            addedMathRange = maths.Add(formulaInput);
            Release(maths);
            maths = addedMathRange.OMaths;
            math = maths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            formulaRange = math.Range;
            formulaFont = formulaRange.Font;
            formulaFont.Size = 14f;

            numberCellRange = numberCell.Range.Duplicate;
            numberCellRange.End = Math.Max(numberCellRange.Start, numberCellRange.End - 1);
            numberCellRange.Text = "()";
            fieldRange = document.Range(numberCellRange.Start + 1, numberCellRange.Start + 1);
            fields = document.Fields;
            object fieldType = -1;
            object fieldCode = "SEQ Equation \\* ARABIC";
            object preserveFormatting = true;
            field = fields.Add(
                fieldRange,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            field.Update();
            numberRange = numberCell.Range.Duplicate;
            numberRange.End = Math.Max(numberRange.Start, numberRange.End - 1);
            numberFont = numberRange.Font;
            numberFont.Size = 11f;

            var formulaSizeBefore = formulaRange.Font.Size;
            var numberSizeBefore = numberRange.Font.Size;
            formulaRange.Font.Size = 24f;
            math.BuildUp();
            field.Update();
            var formulaSizeAfter = formulaRange.Font.Size;
            var numberSizeAfter = numberRange.Font.Size;

            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"Native Word numbered equation probe: formula {formulaSizeBefore:0.###}pt -> {formulaSizeAfter:0.###}pt; "
                + $"SEQ number {numberSizeBefore:0.###}pt -> {numberSizeAfter:0.###}pt.");
            Console.WriteLine($"Native Word probe artifact: {documentPath}");
            AssertClose(24f, formulaSizeAfter, 0.1f,
                "Native Word OMML formula did not accept the requested 24 pt size.");
            AssertClose(numberSizeBefore, numberSizeAfter, 0.1f,
                "Native Word unexpectedly changed the independent SEQ equation number when only the formula size changed.");
            return 0;
        }
        finally
        {
            Release(numberFont);
            Release(numberRange);
            Release(field);
            Release(fields);
            Release(fieldRange);
            Release(numberCellRange);
            Release(formulaFont);
            Release(formulaRange);
            Release(math);
            Release(addedMathRange);
            Release(maths);
            Release(formulaInput);
            Release(numberCell);
            Release(formulaCell);
            Release(table);
            Release(insertionRange);
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

    private static int RunInstalledWordBulkImportAcceptance(
        string sourcePath,
        string objectMode,
        string artifactRoot)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("Bulk import source file was not found.", sourcePath);
        if (!string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(objectMode, "ole", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentOutOfRangeException(nameof(objectMode), objectMode, "Expected omml or ole.");
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "VisualTeX Office integration must be installed before bulk import acceptance.");
        var runningWord = Process.GetProcessesByName("WINWORD");
        var preexistingWordProcessIds = runningWord
            .Select(process => process.Id)
            .ToHashSet();
        int? bootstrapWordProcessId = null;
        var useIsolatedWordInstance = false;
        try
        {
            using var currentProcess = Process.GetCurrentProcess();
            var currentStartedAt = currentProcess.StartTime;
            var bootstrapCandidates = runningWord
                .Where(process =>
                {
                    try
                    {
                        return process.StartTime >= currentStartedAt.AddSeconds(-1);
                    }
                    catch
                    {
                        return false;
                    }
                })
                .ToArray();
            useIsolatedWordInstance = runningWord.Length > 0 && bootstrapCandidates.Length == 0;
            if (useIsolatedWordInstance)
            {
                Console.Error.WriteLine(
                    "Existing Word windows detected; starting a separate hidden Word instance for acceptance.");
            }
            else if (bootstrapCandidates.Length == 1)
            {
                bootstrapWordProcessId = bootstrapCandidates[0].Id;
                Console.Error.WriteLine(
                    $"Reusing acceptance-owned Word bootstrap process {bootstrapWordProcessId.Value}.");
            }
        }
        finally
        {
            foreach (var process in runningWord) process.Dispose();
        }

        Directory.CreateDirectory(artifactRoot);
        var source = File.ReadAllText(sourcePath, Encoding.UTF8);
        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Auto,
            string.Equals(objectMode, "ole", StringComparison.OrdinalIgnoreCase)
                ? WordBulkFormulaObjectMode.Ole
                : WordBulkFormulaObjectMode.Omml);
        var expectedFormulaCount = parsed.FormulaCount;
        var expectedListCount = parsed.Blocks.Count(block =>
            block.Kind is WordBulkBlockKind.Bullet or WordBulkBlockKind.Numbered);
        var outputPath = Path.Combine(
            artifactRoot,
            $"VisualTeX-Word-Bulk-{objectMode.ToLowerInvariant()}.docx");
        TryDelete(outputPath);

        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "auto");
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", objectMode);

        Process? wordProcess = null;
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        VisualTeXSessionClient? client = null;
        var converterSessionIds = new List<string>();
        var temporaryPreviewPaths = new List<string>();
        try
        {
            if (useIsolatedWordInstance)
            {
                application = new Word.Application();
                var isolatedDeadline = DateTime.UtcNow.AddSeconds(10);
                while (wordProcess is null && DateTime.UtcNow < isolatedDeadline)
                {
                    var candidates = Process.GetProcessesByName("WINWORD");
                    try
                    {
                        wordProcess = candidates
                            .Where(process => !preexistingWordProcessIds.Contains(process.Id))
                            .OrderByDescending(process =>
                            {
                                try { return process.StartTime; }
                                catch { return DateTime.MinValue; }
                            })
                            .FirstOrDefault();
                        foreach (var candidate in candidates)
                        {
                            if (!ReferenceEquals(candidate, wordProcess)) candidate.Dispose();
                        }
                    }
                    catch
                    {
                        foreach (var candidate in candidates) candidate.Dispose();
                        throw;
                    }
                    if (wordProcess is null) Thread.Sleep(100);
                }
                Assert(
                    wordProcess is not null,
                    "Word did not create a separate process for isolated acceptance; existing documents were left untouched.");
                RetryRejectedOfficeCall(() => application.Visible = false);
            }
            else
            {
                if (bootstrapWordProcessId.HasValue)
                {
                    wordProcess = Process.GetProcessById(bootstrapWordProcessId.Value);
                }
                else
                {
                    var wordPath = ResolveOfficeApplicationPath("WINWORD.EXE");
                    wordProcess = Process.Start(new ProcessStartInfo(wordPath)
                    {
                        UseShellExecute = true,
                    }) ?? throw new InvalidOperationException("Word desktop process did not start.");
                }
                application = WaitForWordApplication(TimeSpan.FromSeconds(30));
                RetryRejectedOfficeCall(() => application.Visible = true);
            }
            RetryRejectedOfficeCall(() => application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone);
            while (RetryRejectedOfficeCall(() => application.Documents.Count) > 0)
            {
                Word.Document? existing = null;
                try
                {
                    existing = application.Documents[1];
                    existing.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                finally { Release(existing); }
            }
            document = RetryRejectedOfficeCall(() => application.Documents.Add());
            var service = new WordFormulaService(application);
            var sourceSelection = service.ReadSelection();
            client = new VisualTeXSessionClient();
            client.EnsureHealthyAsync(CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            var prepared = PrepareBulkFormulasForAcceptance(
                client,
                parsed,
                objectMode,
                sourceSelection.DocumentId,
                service.ReadCurrentTypingFontSize(),
                converterSessionIds,
                temporaryPreviewPaths);
            var insertResult = service.InsertBulkDocument(
                parsed,
                prepared,
                sourceSelection.DocumentId,
                sourceSelection.ObjectId);
            Assert(
                insertResult.FormulaCount == expectedFormulaCount,
                $"Direct bulk insertion returned {insertResult.FormulaCount} formulas; expected {expectedFormulaCount}.");
            foreach (var sessionId in converterSessionIds)
            {
                client.CompleteAsync(sessionId, CancellationToken.None)
                    .GetAwaiter()
                    .GetResult();
            }

            IReadOnlyList<FormulaMetadata> metadata = ReadBulkFormulaMetadata(document, objectMode);

            var formulaIds = metadata.Select(item => item.FormulaId).ToArray();
            Assert(
                metadata.Count == expectedFormulaCount,
                $"Expected {expectedFormulaCount} bulk formulas, found {metadata.Count}. "
                + $"FormulaIds=[{string.Join(", ", formulaIds)}].");
            Assert(
                formulaIds.Distinct(StringComparer.OrdinalIgnoreCase).Count()
                    == expectedFormulaCount,
                "Bulk imported formulas do not have independent formula IDs. "
                + $"FormulaIds=[{string.Join(", ", formulaIds)}].");
            Assert(
                metadata.Count(item => item.DisplayMode == "inline") == parsed.InlineFormulaCount,
                "Bulk inline formula count does not match the parsed source.");
            Assert(
                metadata.Count(item => item.DisplayMode == "block") == parsed.DisplayFormulaCount,
                "Bulk display formula count does not match the parsed source.");

            AssertBulkNativeText(document, parsed);
            var listParagraphCount = CountListParagraphs(document);
            Assert(
                listParagraphCount >= expectedListCount,
                $"Expected at least {expectedListCount} native list paragraphs, found {listParagraphCount}.");
            if (parsed.Blocks[0].Kind == WordBulkBlockKind.Heading)
            {
                Assert(
                    IsHeadingParagraph(document.Paragraphs[1]),
                    "The first imported block is not a native Word heading.");
            }
            AssertNativeInlineStyles(document, parsed);
            AssertBulkInlineTextRemainsNative(document, parsed, objectMode);
            AssertBulkDisplaySpacing(document, parsed, objectMode);
            AssertBulkEditRouting(application, document, metadata, objectMode);
            AssertBulkExtendedIntegralOperators(document, parsed, metadata, objectMode);
            AssertBulkNativeAccents(document, parsed, metadata, objectMode);

            var before = metadata.ToDictionary(
                item => item.FormulaId,
                item => item.FontSizePt ?? 0d,
                StringComparer.OrdinalIgnoreCase);
            SelectFormula(document, metadata[0], objectMode);
            service.SetSelectedFormulaFontSize(before[metadata[0].FormulaId] + 1d);
            Thread.Sleep(250);
            IReadOnlyList<FormulaMetadata> after = ReadBulkFormulaMetadata(document, objectMode);
            var afterMap = after.ToDictionary(
                item => item.FormulaId,
                item => item.FontSizePt ?? 0d,
                StringComparer.OrdinalIgnoreCase);
            Assert(
                afterMap[metadata[0].FormulaId] > before[metadata[0].FormulaId],
                "The selected bulk formula did not increase its font size.");
            foreach (var item in metadata.Skip(1))
            {
                Assert(
                    Math.Abs(afterMap[item.FormulaId] - before[item.FormulaId]) <= 0.1,
                    $"Changing one bulk formula also changed formula {item.FormulaId}.");
            }
            AssertBulkLayout(document, after, objectMode);

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;
            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: true,
                AddToRecentFiles: false);
            var reopenedMetadata = ReadBulkFormulaMetadata(reopened, objectMode);
            Assert(
                reopenedMetadata.Count == expectedFormulaCount,
                "Bulk formula count changed after save/reopen.");
            Assert(
                reopenedMetadata.Select(item => item.FormulaId).ToHashSet(StringComparer.OrdinalIgnoreCase)
                    .SetEquals(afterMap.Keys),
                "Bulk formula IDs changed after save/reopen.");
            var reopenedSelected = reopenedMetadata.First(item =>
                item.FormulaId.Equals(metadata[0].FormulaId, StringComparison.OrdinalIgnoreCase));
            Assert(
                Math.Abs((reopenedSelected.FontSizePt ?? 0d) - afterMap[metadata[0].FormulaId]) <= 0.1,
                "Selected formula size was not preserved after save/reopen.");
            AssertBulkLayout(reopened, reopenedMetadata, objectMode);
            AssertBulkDisplaySpacing(reopened, parsed, objectMode);
            AssertBulkExtendedIntegralOperators(reopened, parsed, reopenedMetadata, objectMode);
            AssertReopenedExtendedIntegralEditSource(
                application,
                reopened,
                reopenedMetadata,
                objectMode);
            AssertBulkNativeAccents(reopened, parsed, reopenedMetadata, objectMode);

            Console.WriteLine(
                $"Bulk {objectMode.ToUpperInvariant()} acceptance passed: "
                + $"blocks={parsed.Blocks.Count}, formulas={expectedFormulaCount}, "
                + $"inline={parsed.InlineFormulaCount}, display={parsed.DisplayFormulaCount}, "
                + $"lists={listParagraphCount}, selectedSize="
                + $"{before[metadata[0].FormulaId]:0.##}->{afterMap[metadata[0].FormulaId]:0.##} pt.");
            Console.WriteLine($"Artifact: {outputPath}");
            return 0;
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", null);
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { application?.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            foreach (var path in temporaryPreviewPaths) TryDelete(path);
            client?.Dispose();
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
            if (wordProcess is not null)
            {
                try
                {
                    if (!wordProcess.WaitForExit(5000)) wordProcess.Kill();
                }
                catch { }
                wordProcess.Dispose();
            }
        }
    }

    private static IReadOnlyDictionary<string, PreparedWordBulkFormula> PrepareBulkFormulasForAcceptance(
        VisualTeXSessionClient client,
        WordBulkImportDocument document,
        string objectMode,
        string? sourceDocumentId,
        double fontSizePt,
        ICollection<string> converterSessionIds,
        ICollection<string> temporaryPreviewPaths)
    {
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        var sessionObjectMode = string.Equals(objectMode, "ole", StringComparison.OrdinalIgnoreCase)
            ? FormulaOleContract.NativeOleMode
            : FormulaOleContract.WordOmmlMode;
        var normalizedFontSize = FormulaFontSize.Normalize(fontSizePt);
        foreach (var run in document.Blocks
                     .SelectMany(block => block.Runs)
                     .Where(candidate => candidate.IsFormula))
        {
            var line = new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = run.Latex,
            };
            var session = client.CreateSessionAsync(
                    new CreateVstoSessionRequest
                    {
                        Mode = "create",
                        Host = "word",
                        SourceDocumentId = sourceDocumentId,
                        Title = "Bulk import acceptance formula",
                        Lines = new List<FormulaLine> { line },
                        ActiveLineId = line.Id,
                        CodeFormat = "latex",
                        DisplayMode = run.DisplayMode,
                        ObjectMode = sessionObjectMode,
                        Numbered = false,
                        FontSizePt = normalizedFontSize,
                        AutoCommitOnClose = false,
                    },
                    CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            converterSessionIds.Add(session.Id);
            client.OpenConverterAsync(session.Id, CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            session = client.WaitForCommitAsync(
                    session.Id,
                    TimeSpan.FromMinutes(3),
                    CancellationToken.None)
                .GetAwaiter()
                .GetResult();
            if (session.Status == "failed")
                throw new InvalidOperationException(
                    session.Error ?? $"Bulk formula render failed: {run.Latex}");
            if (session.Status == "cancelled")
                throw new OperationCanceledException(
                    $"Bulk formula render was cancelled: {run.Latex}");
            var export = session.ExportResult
                ?? throw new InvalidOperationException(
                    $"Bulk formula has no export result: {run.Latex}");
            if (string.IsNullOrWhiteSpace(export.MathMl))
                throw new InvalidDataException(
                    $"Bulk formula has no MathML result: {run.Latex}");
            AssertNoUnknownCommandMathMl(run.Latex, export.MathMl!);
            AssertLatexCompatibilityMathMl(run.Latex, export.MathMl!);
            AssertExtendedIntegralMathMl(run.Latex, export.MathMl!);

            string? pngPath = null;
            string? svgPath = null;
            string? emfPath = null;
            if (string.Equals(
                    sessionObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                pngPath = client.MaterializePng(session);
                temporaryPreviewPaths.Add(pngPath);
                svgPath = client.MaterializeSvg(session);
                temporaryPreviewPaths.Add(svgPath);
                AssertNoUnknownCommandSvg(run.Latex, svgPath);
                AssertExtendedIntegralSvg(run.Latex, svgPath);
                emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                    svgPath,
                    export.Width,
                    export.Height);
                temporaryPreviewPaths.Add(emfPath);
            }

            var independentSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "create",
                Host = "word",
                FormulaId = Guid.NewGuid().ToString("D"),
                SourceDocumentId = sourceDocumentId,
                Title = "Bulk import acceptance formula",
                Lines = new List<FormulaLine>
                {
                    new()
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        Latex = run.Latex,
                    },
                },
                CodeFormat = "latex",
                DisplayMode = run.DisplayMode,
                ObjectMode = sessionObjectMode,
                Numbered = false,
                FontSizePt = normalizedFontSize,
                Status = "committing",
                Dirty = true,
                ExportResult = export,
            };
            prepared.Add(run.Id, new PreparedWordBulkFormula
            {
                Run = run,
                Session = independentSession,
                MathMl = export.MathMl,
                PngPath = pngPath,
                EmfPath = emfPath,
            });
        }
        return prepared;
    }

    private static IEnumerable<(string Command, string Symbol, string CodePoint)>
        FindExtendedIntegralOperators(string latex)
    {
        foreach (var item in ExtendedIntegralOperators)
        {
            if (latex.IndexOf("\\" + item.Command, StringComparison.Ordinal) >= 0)
                yield return item;
        }
    }

    private static void AssertNoUnknownCommandMathMl(string latex, string mathMl)
    {
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var errors = document
            .Descendants(presentationMath + "mtext")
            .Where(element =>
                element.Value.IndexOf('\\') >= 0
                || string.Equals(
                    element.Attribute("mathcolor")?.Value,
                    "red",
                    StringComparison.OrdinalIgnoreCase))
            .Select(element => element.Value.Trim())
            .ToArray();
        Assert(
            errors.Length == 0,
            "MathJax emitted unknown-command error text during bulk formula conversion. "
            + $"Latex='{latex}', Errors=[{string.Join(", ", errors)}].");
    }

    private static void AssertLatexCompatibilityMathMl(string latex, string mathMl)
    {
        if (latex.IndexOf("\\bm", StringComparison.Ordinal) >= 0
            || latex.IndexOf("\\boldsymbol", StringComparison.Ordinal) >= 0)
        {
            Assert(
                mathMl.IndexOf("mathvariant=\"bold-italic\"", StringComparison.OrdinalIgnoreCase) >= 0,
                "Bold mathematical symbols did not retain bold-italic semantics in MathML. "
                + $"Latex='{latex}', MathML='{mathMl}'.");
            Assert(
                mathMl.IndexOf("\\bm", StringComparison.Ordinal) < 0
                && mathMl.IndexOf("\\boldsymbol", StringComparison.Ordinal) < 0,
                "Bold-symbol commands leaked into MathML as literal text.");
        }
        if (latex.IndexOf("\\begin{vmatrix}", StringComparison.Ordinal) >= 0
            || latex.IndexOf("\\begin{Vmatrix}", StringComparison.Ordinal) >= 0)
        {
            var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
            XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
            var table = document.Descendants(presentationMath + "mtable").SingleOrDefault();
            Assert(table is not null, "vmatrix MathML does not contain a real mtable.");
            var rows = table!.Elements(presentationMath + "mtr").ToArray();
            Assert(rows.Length > 0, "vmatrix MathML contains no matrix rows.");
            Assert(
                rows.All(row => row.Elements(presentationMath + "mtd").Any()),
                "vmatrix MathML contains an empty matrix row.");
        }
    }

    private static void AssertNoUnknownCommandSvg(string latex, string svgPath)
    {
        var svg = File.ReadAllText(svgPath, Encoding.UTF8);
        Assert(
            svg.IndexOf("data-mml-node=\"mtext\" fill=\"red\"", StringComparison.OrdinalIgnoreCase) < 0
            && svg.IndexOf("data-mml-node=\"merror\"", StringComparison.OrdinalIgnoreCase) < 0,
            "MathJax emitted a red command/error glyph in the OLE vector preview. "
            + $"Latex='{latex}', Svg='{svgPath}'.");
    }

    private static void AssertExtendedIntegralMathMl(string latex, string mathMl)
    {
        var operators = FindExtendedIntegralOperators(latex).ToArray();
        if (operators.Length == 0) return;

        Assert(
            mathMl.IndexOf("mathcolor=\"red\"", StringComparison.OrdinalIgnoreCase) < 0,
            $"Extended integral MathML contains an unknown-command error: {mathMl}");
        foreach (var item in operators)
        {
            var entity = $"&#x{item.CodePoint};";
            Assert(
                mathMl.IndexOf(item.Symbol, StringComparison.Ordinal) >= 0
                || mathMl.IndexOf(entity, StringComparison.OrdinalIgnoreCase) >= 0,
                $"\\{item.Command} did not become Unicode {item.Symbol} in MathML: {mathMl}");
            Assert(
                mathMl.IndexOf("\\" + item.Command, StringComparison.Ordinal) < 0,
                $"\\{item.Command} leaked into MathML as literal text: {mathMl}");
        }
    }

    private static void AssertExtendedIntegralSvg(string latex, string svgPath)
    {
        var operators = FindExtendedIntegralOperators(latex).ToArray();
        if (operators.Length == 0) return;

        var svg = File.ReadAllText(svgPath, Encoding.UTF8);
        Assert(svg.IndexOf("<svg", StringComparison.OrdinalIgnoreCase) >= 0,
            $"Extended integral export is not SVG: {svgPath}");
        Assert(
            svg.IndexOf("fill=\"red\"", StringComparison.OrdinalIgnoreCase) < 0
            && svg.IndexOf("#ff0000", StringComparison.OrdinalIgnoreCase) < 0,
            $"Extended integral SVG contains an error glyph: {svgPath}");
        Assert(
            svg.IndexOf("<path", StringComparison.OrdinalIgnoreCase) >= 0
            || svg.IndexOf("<use", StringComparison.OrdinalIgnoreCase) >= 0,
            $"Extended integral SVG contains no rendered vector glyph: {svgPath}");
        foreach (var item in operators)
        {
            Assert(
                svg.IndexOf("\\" + item.Command, StringComparison.Ordinal) < 0,
                $"\\{item.Command} leaked into SVG as literal text: {svgPath}");
            var renderedCommand = string.Equals(
                item.Command,
                "intctrclockwise",
                StringComparison.Ordinal)
                ? "awint"
                : item.Command;
            Assert(
                svg.IndexOf(
                    $"data-visualtex-integral=\"{renderedCommand}\"",
                    StringComparison.Ordinal) >= 0,
                $"\\{item.Command} did not use the VisualTeX large vector glyph in OLE SVG: {svgPath}");
            Assert(
                svg.IndexOf($">{item.Symbol}</text>", StringComparison.Ordinal) < 0,
                $"\\{item.Command} fell back to a small system-font character in OLE SVG: {svgPath}");
        }
    }

    private static void AssertBulkExtendedIntegralOperators(
        Word.Document document,
        WordBulkImportDocument parsed,
        IReadOnlyList<FormulaMetadata> metadata,
        string objectMode)
    {
        var sourceLatex = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();
        var operators = sourceLatex
            .SelectMany(FindExtendedIntegralOperators)
            .GroupBy(item => item.Command, StringComparer.Ordinal)
            .Select(group => group.First())
            .ToArray();
        if (operators.Length == 0) return;

        var persistedLatex = metadata.Select(item => item.Latex ?? string.Empty).ToArray();
        foreach (var item in operators)
        {
            Assert(
                persistedLatex.Any(latex =>
                    latex.IndexOf("\\" + item.Command, StringComparison.Ordinal) >= 0),
                $"Bulk import lost the editable source command \\{item.Command}.");
        }

        if (!string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
            return;

        var package = XDocument.Parse(document.Content.WordOpenXML, LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var allMathText = string.Concat(package.Descendants(math + "t").Select(text => text.Value));
        foreach (var item in operators)
        {
            Assert(
                allMathText.IndexOf("\\" + item.Command, StringComparison.Ordinal) < 0,
                $"Bulk OMML contains literal command text \\{item.Command}.");
            var narySymbols = package
                .Descendants(math + "nary")
                .Select(nary => nary.Element(math + "naryPr")
                    ?.Element(math + "chr")
                    ?.Attribute(math + "val")
                    ?.Value)
                .Where(value => !string.IsNullOrEmpty(value))
                .Cast<string>()
                .ToArray();
            Assert(
                narySymbols.Contains(item.Symbol, StringComparer.Ordinal),
                $"Bulk OMML did not preserve \\{item.Command} as native n-ary operator {item.Symbol}. "
                + $"Actual n-ary symbols=[{string.Join(",", narySymbols.Select(symbol => $"{symbol}:U+{char.ConvertToUtf32(symbol, 0):X4}"))}].");
        }
    }

    private static void AssertReopenedExtendedIntegralEditSource(
        Word.Application application,
        Word.Document document,
        IReadOnlyList<FormulaMetadata> metadata,
        string objectMode)
    {
        if (!string.Equals(objectMode, "ole", StringComparison.OrdinalIgnoreCase))
            return;
        var contour = metadata.FirstOrDefault(item =>
            (item.Latex ?? string.Empty).IndexOf("\\oiint", StringComparison.Ordinal) >= 0);
        if (contour is null) return;

        SelectFormula(document, contour, objectMode);
        var reopened = new WordFormulaService(application).ReadSelection();
        Assert(
            string.Equals(
                reopened.ObjectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal),
            "Save/reopen changed the \\oiint OLE object mode.");
        Assert(
            string.Equals(reopened.FormulaId, contour.FormulaId, StringComparison.OrdinalIgnoreCase),
            "Save/reopen routed the \\oiint OLE object to a different formulaId.");
        var reopenedLatex = reopened.Metadata?.Latex ?? string.Empty;
        Assert(
            reopenedLatex.IndexOf("\\oiint", StringComparison.Ordinal) >= 0,
            "Save/reopen lost the editable \\oiint command from OLE metadata. "
            + $"ActualLatex='{reopenedLatex}'.");
        Assert(
            reopenedLatex.IndexOf('∯') < 0,
            "Save/reopen serialized \\oiint as a raw Unicode glyph instead of canonical LaTeX.");
    }

    private static void AssertBulkNativeAccents(
        Word.Document document,
        WordBulkImportDocument parsed,
        IReadOnlyList<FormulaMetadata> metadata,
        string objectMode)
    {
        var sourceLatex = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();
        var accents = NativeAccentOperators
            .Where(item => sourceLatex.Any(latex =>
                latex.IndexOf("\\" + item.Command, StringComparison.Ordinal) >= 0))
            .ToArray();
        if (accents.Length == 0) return;

        var persistedLatex = metadata.Select(item => item.Latex ?? string.Empty).ToArray();
        foreach (var item in accents)
        {
            Assert(
                persistedLatex.Any(latex =>
                    latex.IndexOf("\\" + item.Command, StringComparison.Ordinal) >= 0),
                $"Bulk import lost the editable accent command \\{item.Command}.");
        }

        if (!string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
            return;

        var package = XDocument.Parse(document.Content.WordOpenXML, LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var accentCharacters = package
            .Descendants(math + "acc")
            .Select(accent => accent.Element(math + "accPr")
                ?.Element(math + "chr")
                ?.Attribute(math + "val")
                ?.Value ?? "\u0302")
            .ToArray();
        foreach (var item in accents)
        {
            Assert(
                accentCharacters.Contains(item.Accent, StringComparer.Ordinal),
                $"Bulk OMML did not preserve \\{item.Command} as native m:acc character "
                + $"U+{char.ConvertToUtf32(item.Accent, 0):X4}. Actual accents=[{string.Join(",", accentCharacters.Select(value => $"U+{char.ConvertToUtf32(value, 0):X4}"))}].");
        }

        var allMathText = string.Concat(package.Descendants(math + "t").Select(text => text.Value));
        Assert(
            allMathText.IndexOf('\uFFFD') < 0 && allMathText.IndexOf('□') < 0,
            "Bulk OMML accent formula contains a replacement or placeholder box glyph.");
    }

    private static string ResolveOfficeApplicationPath(string fileName)
    {
        using var key = Registry.LocalMachine.OpenSubKey(
            $@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{fileName}");
        var registered = Convert.ToString(key?.GetValue(null), CultureInfo.InvariantCulture);
        if (!string.IsNullOrWhiteSpace(registered) && File.Exists(registered))
            return registered;
        foreach (var root in new[]
                 {
                     Environment.GetEnvironmentVariable("ProgramW6432"),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                     Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                 }.Where(value => !string.IsNullOrWhiteSpace(value)))
        {
            var candidate = Path.Combine(root!, "Microsoft Office", "Root", "Office16", fileName);
            if (File.Exists(candidate)) return candidate;
        }
        throw new FileNotFoundException($"Unable to locate {fileName}.");
    }

    private static void RetryRejectedOfficeCall(Action action)
    {
        _ = RetryRejectedOfficeCall(() =>
        {
            action();
            return true;
        });
    }

    private static T RetryRejectedOfficeCall<T>(Func<T> action)
    {
        var deadline = DateTime.UtcNow.AddSeconds(20);
        Exception? lastError = null;
        do
        {
            try { return action(); }
            catch (COMException error) when (
                error.HResult == unchecked((int)0x80010001)
                || error.HResult == unchecked((int)0x8001010A))
            {
                lastError = error;
                Thread.Sleep(250);
            }
        } while (DateTime.UtcNow < deadline);
        throw new TimeoutException("Office continued rejecting COM calls during startup.", lastError);
    }

    private static Word.Application WaitForWordApplication(TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        Exception? lastError = null;
        do
        {
            Thread.Sleep(250);
            try
            {
                return (Word.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception error) when (error is COMException or InvalidCastException)
            {
                lastError = error;
            }
        } while (DateTime.UtcNow < deadline);
        throw new TimeoutException("Word did not register in the Running Object Table.", lastError);
    }

    private static IReadOnlyList<FormulaMetadata> ReadBulkFormulaMetadata(
        Word.Document document,
        string objectMode)
    {
        var result = new List<FormulaMetadata>();
        if (string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
        {
            foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata is not null) result.Add(metadata);
            }
            return result.OrderBy(item => FormulaStart(document, item, objectMode)).ToArray();
        }
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
                    if (metadata is not null)
                    {
                        metadata.FontSizePt = FormulaFontSize.InferOleFontSize(
                            shape.Width,
                            shape.Height,
                            metadata);
                        result.Add(metadata);
                    }
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }
        return result.OrderBy(item => FormulaStart(document, item, objectMode)).ToArray();
    }

    private static int FormulaStart(
        Word.Document document,
        FormulaMetadata metadata,
        string objectMode)
    {
        if (string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
        {
            Word.Bookmark? bookmark = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, metadata.FormulaId);
                return bookmark?.Range.Start ?? int.MaxValue;
            }
            finally { Release(bookmark); }
        }
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
                    var candidate = WordFormulaMetadataReader.TryRead(shape);
                    if (candidate?.FormulaId == metadata.FormulaId) return shape.Range.Start;
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }
        return int.MaxValue;
    }

    private static int CountListParagraphs(Word.Document document)
    {
        Word.Paragraphs? paragraphs = null;
        var count = 0;
        try
        {
            paragraphs = document.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Word.Paragraph? paragraph = null;
                Word.Range? range = null;
                Word.ListFormat? list = null;
                try
                {
                    paragraph = paragraphs[index];
                    range = paragraph.Range;
                    list = range.ListFormat;
                    if (list.ListType != Word.WdListType.wdListNoNumbering) count++;
                }
                finally
                {
                    Release(list);
                    Release(range);
                    Release(paragraph);
                }
            }
        }
        finally { Release(paragraphs); }
        return count;
    }

    private static bool IsHeadingParagraph(Word.Paragraph paragraph)
    {
        Word.Range? range = null;
        object? styleObject = null;
        try
        {
            range = paragraph.Range;
            styleObject = range.get_Style();
            var style = Convert.ToString(((dynamic)styleObject).NameLocal, CultureInfo.InvariantCulture)
                        ?? Convert.ToString(styleObject, CultureInfo.InvariantCulture)
                        ?? string.Empty;
            return style.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0
                   || style.IndexOf("标题", StringComparison.Ordinal) >= 0;
        }
        catch { return false; }
        finally
        {
            Release(styleObject);
            Release(range);
        }
    }

    private static void AssertBulkNativeText(
        Word.Document document,
        WordBulkImportDocument parsed)
    {
        var samples = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => !run.IsFormula && !string.IsNullOrWhiteSpace(run.Text))
            .Select(run => run.Text.Trim())
            .Where(value => value.Length >= 2)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        foreach (var sample in samples)
        {
            Word.Range? found = null;
            try
            {
                found = FindNativeTextRange(document, sample);
                Assert(found is not null, "Native Word text is missing after bulk import: " + sample);
            }
            finally { Release(found); }
        }
    }

    private static void AssertNativeInlineStyles(
        Word.Document document,
        WordBulkImportDocument parsed)
    {
        var styledRuns = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => !run.IsFormula
                          && !string.IsNullOrWhiteSpace(run.Text)
                          && (run.Bold || run.Italic || run.Code))
            .ToArray();
        foreach (var sample in styledRuns)
        {
            Word.Range? range = null;
            Word.Font? font = null;
            try
            {
                range = FindNativeTextRange(document, sample.Text);
                Assert(range is not null, $"Native styled text is missing: {sample.Text}");
                font = range!.Font;
                if (sample.Bold)
                    Assert(font.Bold != 0, "Bold text was not preserved as native Word bold.");
                if (sample.Italic)
                    Assert(font.Italic != 0, "Italic text was not preserved as native Word italic.");
                if (sample.Code)
                {
                    var names = new List<string>();
                    try { names.Add(font.Name ?? string.Empty); } catch { }
                    try { names.Add(font.NameAscii ?? string.Empty); } catch { }
                    Assert(
                        names.Any(name =>
                            name.IndexOf("Consolas", StringComparison.OrdinalIgnoreCase) >= 0),
                        "Code text was not preserved with the native Word code font. "
                        + $"Sample='{sample.Text}', Fonts=[{string.Join(", ", names)}], "
                        + $"Range={range.Start}:{range.End}.");
                }
            }
            finally
            {
                Release(font);
                Release(range);
            }
        }
    }

    private static Word.Range? FindNativeTextRange(
        Word.Document document,
        string value)
    {
        Word.Range? searchRange = null;
        Word.Find? find = null;
        Word.Range? result = null;
        try
        {
            searchRange = document.Content.Duplicate;
            find = searchRange.Find;
            find.ClearFormatting();
            var found = find.Execute(
                FindText: value,
                MatchCase: true,
                MatchWholeWord: false,
                MatchWildcards: false,
                Forward: true,
                Wrap: Word.WdFindWrap.wdFindStop,
                Format: false);
            if (!found) return null;
            result = searchRange.Duplicate;
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(find);
            Release(searchRange);
        }
    }

    private static void AssertBulkDisplaySpacing(
        Word.Document document,
        WordBulkImportDocument parsed,
        string objectMode)
    {
        for (var index = 0; index + 1 < parsed.Blocks.Count; index++)
        {
            var block = parsed.Blocks[index];
            if (block.Kind != WordBulkBlockKind.Paragraph
                || parsed.Blocks[index + 1].Kind != WordBulkBlockKind.DisplayFormula)
                continue;

            // An inline formula splits the surrounding prose into multiple
            // native Word runs with an object character between them. Joining
            // those runs creates a string that never exists contiguously in
            // Word, so locate the paragraph by its longest native text span.
            var sample = block.Runs
                .Where(run => !run.IsFormula)
                .Select(run => run.Text.Trim())
                .Where(text => text.Length > 0)
                .OrderByDescending(text => text.Length)
                .FirstOrDefault() ?? string.Empty;
            if (sample.Length == 0) continue;

            Word.Range? found = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.ParagraphFormat? format = null;
            try
            {
                found = FindNativeTextRange(document, sample);
                Assert(found is not null,
                    "Could not locate prose before a bulk display formula: " + sample);
                paragraphs = found!.Paragraphs;
                Assert(paragraphs.Count > 0,
                    "Prose before a bulk display formula has no Word paragraph: " + sample);
                paragraph = paragraphs[1];
                format = paragraph.Format;
                if (string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
                {
                    AssertClose(
                        8f,
                        format.SpaceAfter,
                        0.2f,
                        "Bulk prose before native OMML did not retain Word's native paragraph spacing. "
                        + $"Text='{sample}'.");
                }
                else
                {
                    Assert(
                        format.SpaceAfter <= 0.1f,
                        "Bulk prose before OLE was not compacted for the preview's internal padding. "
                        + $"Text='{sample}', SpaceAfter={format.SpaceAfter:0.##} pt.");
                }
            }
            finally
            {
                Release(format);
                Release(paragraph);
                Release(paragraphs);
                Release(found);
            }
        }
    }

    private static void AssertBulkInlineTextRemainsNative(
        Word.Document document,
        WordBulkImportDocument parsed,
        string objectMode)
    {
        if (!string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
            return;

        var trailingSamples = new List<string>();
        foreach (var block in parsed.Blocks)
        {
            for (var index = 0; index < block.Runs.Count; index++)
            {
                var run = block.Runs[index];
                if (!run.IsFormula || run.DisplayMode != "inline") continue;
                var trailing = block.Runs
                    .Skip(index + 1)
                    .Where(candidate => !candidate.IsFormula)
                    .Select(candidate => candidate.Text)
                    .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
                var sample = trailing?.Trim();
                if (!string.IsNullOrWhiteSpace(sample) && sample!.Length >= 4)
                    trailingSamples.Add(sample);
            }
        }
        if (trailingSamples.Count == 0) return;

        var openXml = document.Content.WordOpenXML;
        var package = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var mathText = string.Concat(package
            .Descendants(math + "oMath")
            .SelectMany(equation => equation.Descendants(math + "t"))
            .Select(text => text.Value));
        foreach (var sample in trailingSamples.Distinct(StringComparer.Ordinal))
        {
            Assert(
                mathText.IndexOf(sample, StringComparison.Ordinal) < 0,
                "Bulk inline OMML absorbed following native Word text into the math zone: "
                + sample);
        }
    }

    private static void AssertBulkEditRouting(
        Word.Application application,
        Word.Document document,
        IReadOnlyList<FormulaMetadata> metadata,
        string objectMode)
    {
        var service = new WordFormulaService(application);
        var editedInlineOle = false;
        foreach (var item in metadata)
        {
            SelectFormula(document, item, objectMode);
            var source = service.ReadSelection();
            Assert(source.Metadata is not null, $"Formula {item.FormulaId} cannot be routed to the editor.");
            Assert(
                string.Equals(source.FormulaId, item.FormulaId, StringComparison.OrdinalIgnoreCase),
                $"Formula {item.FormulaId} routed to a different formula ID "
                + $"{source.FormulaId ?? "<null>"}. "
                + $"ExpectedLatex='{item.Latex}', ActualLatex='{source.Metadata?.Latex}'.");
            AssertClose(
                FormulaFontSize.ResolveSemanticFontSize(item),
                FormulaFontSize.ResolveSemanticFontSize(source.Metadata!),
                0.01f,
                "Opening a bulk-imported formula for editing changed its semantic font size. "
                + $"FormulaId={item.FormulaId}.");

            if (editedInlineOle
                || !string.Equals(objectMode, "ole", StringComparison.OrdinalIgnoreCase)
                || !string.Equals(item.DisplayMode, "inline", StringComparison.Ordinal))
                continue;

            Word.InlineShape? beforeShape = null;
            Word.InlineShape? afterShape = null;
            PreviewSet? editedPreview = null;
            try
            {
                beforeShape = FindWordFormula(document, item.FormulaId)
                    ?? throw new InvalidOperationException(
                        "Bulk inline OLE disappeared before the edit acceptance.");
                var beforeWidth = beforeShape.Width;
                var beforeHeight = beforeShape.Height;
                var baselineSize = FormulaFontSize.OleSizeAt(
                    item,
                    FormulaFontSize.ResolveSemanticFontSize(item));
                Console.WriteLine(
                    "  [OLE inline 11 pt baseline] "
                    + $"latex='{item.Latex}'; actual={beforeWidth:0.###}x{beforeHeight:0.###} pt; "
                    + $"semantic-svg={baselineSize.Width:0.###}x{baselineSize.Height:0.###} pt; "
                    + $"error={Math.Abs(beforeWidth - baselineSize.Width) / Math.Max(1f, baselineSize.Width):P1}/"
                    + $"{Math.Abs(beforeHeight - baselineSize.Height) / Math.Max(1f, baselineSize.Height):P1}; "
                    + $"render={item.RenderWidthPx:0.###}x{item.RenderHeightPx:0.###} px.");
                var renderWidth = Math.Max(2, (int)Math.Ceiling((item.RenderWidthPx ?? 24d) * 1.9d));
                var renderHeight = Math.Max(2, (int)Math.Ceiling(item.RenderHeightPx ?? 11d));
                var previewRoot = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "VisualTeX",
                    "office",
                    "temp");
                Directory.CreateDirectory(previewRoot);
                editedPreview = CreatePreviewSet(
                    previewRoot,
                    item.FormulaId,
                    "bulk-inline-double-click-edit",
                    renderWidth,
                    renderHeight);
                var editedSession = CreateWordSession(
                    item.FormulaId,
                    "edit",
                    FormulaOleContract.NativeOleMode,
                    item.Latex + "+1",
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></math>",
                    editedPreview,
                    numbered: false,
                    originalMetadata: source.Metadata,
                    fontSizePt: FormulaFontSize.ResolveSemanticFontSize(source.Metadata!));
                editedSession.DisplayMode = "inline";
                service.ReplaceOle(
                    editedSession,
                    editedPreview.PngPath,
                    editedPreview.EmfPath);

                afterShape = FindWordFormula(document, item.FormulaId)
                    ?? throw new InvalidOperationException(
                        "Bulk inline OLE disappeared after the edit acceptance.");
                Assert(
                    afterShape.Width >= beforeWidth * 1.5f,
                    "Editing a bulk-imported inline OLE kept the old width and shrank the new content. "
                    + $"FormulaId={item.FormulaId}; Before={beforeWidth:0.##} pt; "
                    + $"After={afterShape.Width:0.##} pt.");
                AssertClose(
                    beforeHeight,
                    afterShape.Height,
                    1f,
                    "Editing a bulk-imported inline OLE changed its glyph height.");
                afterShape.Range.Select();
                var editedSource = service.ReadSelection();
                AssertClose(
                    FormulaFontSize.ResolveSemanticFontSize(item),
                    FormulaFontSize.ResolveSemanticFontSize(editedSource.Metadata),
                    0.01f,
                    "Edited bulk inline OLE no longer reports its original semantic font size.");
                editedInlineOle = true;
            }
            finally
            {
                Release(afterShape);
                Release(beforeShape);
                if (editedPreview is not null)
                {
                    TryDelete(editedPreview.SvgPath);
                    TryDelete(editedPreview.EmfPath);
                    TryDelete(editedPreview.PngPath);
                }
            }
        }
    }

    private static void SelectFormula(
        Word.Document document,
        FormulaMetadata metadata,
        string objectMode)
    {
        if (string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? range = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, metadata.FormulaId)
                    ?? throw new InvalidOperationException($"OMML bookmark is missing: {metadata.FormulaId}");
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                range.Select();
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
            return;
        }
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
                    var candidate = WordFormulaMetadataReader.TryRead(shape);
                    if (candidate?.FormulaId != metadata.FormulaId) continue;
                    shape.Range.Select();
                    return;
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }
        throw new InvalidOperationException($"OLE formula is missing: {metadata.FormulaId}");
    }

    private static void AssertBulkLayout(
        Word.Document document,
        IReadOnlyList<FormulaMetadata> metadata,
        string objectMode)
    {
        foreach (var item in metadata)
        {
            if (string.Equals(objectMode, "omml", StringComparison.OrdinalIgnoreCase))
            {
                Word.Bookmark? bookmark = null;
                Word.Range? equation = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, item.FormulaId);
                    Assert(bookmark is not null, $"OMML bookmark disappeared: {item.FormulaId}");
                    equation = WordOmmlFormulaStore.GetEquationRange(bookmark!);
                    var equationOpenXml = equation.WordOpenXML;
                    AssertNoLiteralLatexCommandsInOmml(
                        equationOpenXml,
                        item.FormulaId);
                    AssertNoEmptyOmmlScriptSlots(equationOpenXml, item.FormulaId);
                    if (item.Latex.IndexOf("\\begin{vmatrix}", StringComparison.Ordinal) >= 0
                        || item.Latex.IndexOf("\\begin{Vmatrix}", StringComparison.Ordinal) >= 0)
                        AssertNativeOmmlVmatrix(equationOpenXml, item.FormulaId);
                    if (item.Latex.IndexOf("\\begin{align", StringComparison.Ordinal) >= 0
                        || item.Latex.IndexOf("\\begin{aligned", StringComparison.Ordinal) >= 0)
                    {
                        AssertOmmlMatrixColumnAlignment(
                            equation,
                            new[] { "right", "left" },
                            expectedRows: 0,
                            "Bulk aligned OMML did not preserve the ampersand alignment columns.");
                    }
                    if (item.DisplayMode == "block")
                        Assert(
                            equation.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter,
                            "Bulk display OMML is not centered like a normal display formula.");
                }
                finally
                {
                    Release(equation);
                    Release(bookmark);
                }
                continue;
            }
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
                        var candidate = WordFormulaMetadataReader.TryRead(shape);
                        if (candidate?.FormulaId != item.FormulaId) continue;
                        var expectedSize = FormulaFontSize.OleSizeAt(
                            item,
                            FormulaFontSize.ResolveSemanticFontSize(item));
                        var widthError = Math.Abs(shape.Width - expectedSize.Width)
                            / Math.Max(1f, expectedSize.Width);
                        var heightError = Math.Abs(shape.Height - expectedSize.Height)
                            / Math.Max(1f, expectedSize.Height);
                        Assert(
                            widthError <= 0.10f && heightError <= 0.10f,
                            "Bulk OLE physical size does not match the SVG's 96 dpi to Word 72 dpi semantic size. "
                            + $"FormulaId={item.FormulaId}; Actual={shape.Width:0.###}x{shape.Height:0.###} pt; "
                            + $"Expected={expectedSize.Width:0.###}x{expectedSize.Height:0.###} pt; "
                            + $"Error={widthError:P1}/{heightError:P1}.");
                        if (string.Equals(item.Latex.Trim(), "x", StringComparison.Ordinal))
                        {
                            Console.WriteLine(
                                "  [OLE 11 pt x] "
                                + $"actual={shape.Width:0.###}x{shape.Height:0.###} pt; "
                                + $"semantic-svg={expectedSize.Width:0.###}x{expectedSize.Height:0.###} pt; "
                                + $"error={widthError:P1}/{heightError:P1}; "
                                + $"render={item.RenderWidthPx:0.###}x{item.RenderHeightPx:0.###} px.");
                        }
                        if (item.DisplayMode == "inline")
                        {
                            var objectModelPosition = shape.Range.Font.Position;
                            var persistedPosition = ReadPersistedRunFontPosition(shape);
                            Assert(
                                persistedPosition < 0,
                                "Bulk inline OLE did not receive the normal baseline offset. "
                                + $"FormulaId={item.FormulaId}; "
                                + $"ObjectModelPosition={objectModelPosition}; "
                                + $"PersistedPosition={persistedPosition}; "
                                + $"Shape={shape.Width:0.##}x{shape.Height:0.##} pt; "
                                + $"RenderHeight={item.RenderHeightPx:0.##}; "
                                + $"Baseline={item.Baseline:0.##}; FontSize={item.FontSizePt:0.##}.");
                            AssertClose(
                                0f,
                                ReadParagraphMarkFontPosition(shape.Range),
                                0.1f,
                                "Bulk inline OLE contaminated the paragraph typing baseline.");
                        }
                        else
                            Assert(
                                shape.Range.ParagraphFormat.Alignment
                                == Word.WdParagraphAlignment.wdAlignParagraphCenter,
                                "Bulk display OLE is not centered like a normal display formula.");
                        break;
                    }
                    finally { Release(shape); }
                }
            }
            finally { Release(shapes); }
        }
    }

    private static void AssertNoLiteralLatexCommandsInOmml(
        string wordOpenXml,
        string formulaId)
    {
        var document = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var literals = document
            .Descendants(math + "t")
            .Select(text => text.Value)
            .Where(value => value.IndexOf('\\') >= 0)
            .ToArray();
        Assert(
            literals.Length == 0,
            "OMML contains literal LaTeX command text. "
            + $"FormulaId={formulaId}; Values=[{string.Join(", ", literals)}].");
    }

    private static void AssertNativeOmmlVmatrix(
        string wordOpenXml,
        string formulaId)
    {
        var document = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var delimiter = document.Descendants(math + "d").SingleOrDefault(element =>
            element.Element(math + "dPr")?
                .Element(math + "begChr")?
                .Attribute(math + "val")?
                .Value is "|" or "‖");
        Assert(
            delimiter is not null,
            "vmatrix did not become one native OMML delimiter. "
            + $"FormulaId={formulaId}.");
        var matrix = delimiter!.Descendants(math + "m").SingleOrDefault();
        Assert(matrix is not null,
            "vmatrix native delimiter does not contain an OMML matrix. "
            + $"FormulaId={formulaId}.");
        var rows = matrix!.Elements(math + "mr").ToArray();
        Assert(rows.Length > 0,
            "vmatrix OMML matrix contains no rows. "
            + $"FormulaId={formulaId}.");
        foreach (var row in rows)
        {
            var cells = row.Elements(math + "e").ToArray();
            Assert(cells.Length > 0,
                "vmatrix OMML matrix contains an empty row. "
                + $"FormulaId={formulaId}.");
            foreach (var cell in cells)
            {
                Assert(
                    cell.Descendants(math + "t").Any(text =>
                        !string.IsNullOrWhiteSpace(text.Value)),
                    "vmatrix OMML matrix contains a dotted empty slot. "
                    + $"FormulaId={formulaId}; Cell={cell.ToString(SaveOptions.DisableFormatting)}.");
            }
        }
    }

    private static void AssertNoEmptyOmmlScriptSlots(
        string wordOpenXml,
        string formulaId)
    {
        var document = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        foreach (var script in document.Descendants().Where(element =>
                     element.Name == math + "sSub"
                     || element.Name == math + "sSup"
                     || element.Name == math + "sSubSup"))
        {
            AssertOmmlSlotHasVisibleContent(script, math + "e", formulaId);
            if (script.Name == math + "sSub" || script.Name == math + "sSubSup")
                AssertOmmlSlotHasVisibleContent(script, math + "sub", formulaId);
            if (script.Name == math + "sSup" || script.Name == math + "sSubSup")
                AssertOmmlSlotHasVisibleContent(script, math + "sup", formulaId);
        }
    }

    private static void AssertOmmlSlotHasVisibleContent(
        XElement script,
        XName slotName,
        string formulaId)
    {
        var slot = script.Element(slotName);
        var hasText = slot?
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "t")
            .Any(element => !string.IsNullOrWhiteSpace(element.Value)) == true;
        Assert(
            hasText,
            "OMML script contains an empty slot that Word would render as a dotted placeholder. "
            + $"FormulaId={formulaId}; Script={script.Name.LocalName}; Slot={slotName.LocalName}; "
            + $"Xml={script.ToString(SaveOptions.DisableFormatting)}");
    }

    private static int RunInstalledWordInlineRoundTripAnchorAcceptance(
        string artifactRoot)
    {
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "VisualTeX Formula OLE must be installed before running the Word inline round-trip acceptance.");

        Directory.CreateDirectory(artifactRoot);
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(previewRoot);
        var formulaId = Guid.NewGuid().ToString();
        var preview = CreatePreviewSet(
            previewRoot,
            formulaId,
            "inline-roundtrip-anchor",
            196,
            44);
        try
        {
            VerifyWordInlineOleOmmlRoundTripAndAnchorRecovery(
                Path.Combine(
                    artifactRoot,
                    "VisualTeX-Word-Inline-OLE-OMML-OLE.docx"),
                formulaId,
                preview);
            Console.WriteLine(
                "VisualTeX Word inline OLE/OMML size and trailing-anchor recovery acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        finally
        {
            ForceComCleanup();
            TryDelete(preview.SvgPath);
            TryDelete(preview.EmfPath);
            TryDelete(preview.PngPath);
        }
    }

    private static void VerifyWordInlineOleOmmlRoundTripAndAnchorRecovery(
        string path,
        string formulaId,
        PreviewSet preview)
    {
        const string latex = @"v_{GS}+111111";
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow>"
            + "<msub><mi>v</mi><mi>GS</mi></msub><mo>+</mo><mn>111111</mn>"
            + "</mrow></math>";
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.InlineShape? shape = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
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
            selection.TypeText("inline: ");
            var service = new WordFormulaService(application);
            var createSession = CreateWordSession(
                formulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 11);
            createSession.DisplayMode = "inline";
            service.InsertOle(createSession, preview.PngPath, preview.EmfPath);
            shape = FindWordFormula(document, formulaId)
                ?? throw new InvalidOperationException(
                    "Word inline OLE round-trip fixture was not created.");

            // Reproduce the real failure: a user-sized inline OLE is smaller than
            // the current renderer's natural extent. The old OMML -> OLE path
            // discarded this physical size and recreated the larger natural one.
            shape.LockAspectRatio = Office.MsoTriState.msoFalse;
            shape.Width *= 0.68f;
            shape.Height *= 0.68f;
            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
            var expectedWidth = shape.Width;
            var expectedHeight = shape.Height;
            var oleMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidOperationException(
                    "Word inline OLE round-trip metadata is missing.");

            var toOmmlSession = CreateWordSession(
                formulaId,
                "edit",
                FormulaOleContract.WordOmmlMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: oleMetadata,
                fontSizePt: 11);
            toOmmlSession.DisplayMode = "inline";
            service.ReplaceOmml(toOmmlSession, mathMl);
            Release(shape);
            shape = null;
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException(
                    "Inline OLE to OMML conversion did not create its anchor.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var storedMetadata = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException(
                    "Inline OLE physical dimensions were not persisted in OMML metadata.");
            AssertClose(
                expectedWidth,
                (float)(storedMetadata.WordInlineOleWidthPt ?? 0),
                0.05f,
                "Inline OLE to OMML conversion lost the physical width.");
            AssertClose(
                expectedHeight,
                (float)(storedMetadata.WordInlineOleHeightPt ?? 0),
                0.05f,
                "Inline OLE to OMML conversion lost the physical height.");

            // Simulate Word moving the collapsed bookmark from the equation's
            // leading edge to its trailing edge after native OMath rebuilding.
            var bookmarkName = WordOmmlFormulaStore.BookmarkName(formulaId);
            var trailingPosition = equationRange.End;
            bookmark.Delete();
            Release(bookmark);
            bookmark = null;
            object trailingStart = trailingPosition;
            object trailingEnd = trailingPosition;
            Word.Range? trailingRange = null;
            Word.Bookmarks? bookmarks = null;
            try
            {
                trailingRange = document.Range(ref trailingStart, ref trailingEnd);
                bookmarks = document.Bookmarks;
                bookmark = bookmarks.Add(bookmarkName, trailingRange);
            }
            finally
            {
                Release(bookmarks);
                Release(trailingRange);
            }

            Release(equationRange);
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var reopened = ReadSelectedFormula(application, service, equationRange);
            AssertEqual(
                FormulaOleContract.WordOmmlMode,
                reopened.ObjectMode ?? string.Empty,
                "VisualTeX could not reopen an OMML formula whose anchor moved behind it.");
            var reopenedMetadata = reopened.Metadata
                ?? throw new InvalidOperationException(
                    "Recovered trailing-anchor OMML returned no metadata.");

            var backToOleSession = CreateWordSession(
                formulaId,
                "edit",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: reopenedMetadata,
                fontSizePt: 11);
            backToOleSession.DisplayMode = "inline";
            service.ReplaceOle(
                backToOleSession,
                preview.PngPath,
                preview.EmfPath);
            shape = FindWordFormula(document, formulaId)
                ?? throw new InvalidOperationException(
                    "Inline OMML to OLE round trip did not recreate the object.");
            AssertClose(
                expectedWidth,
                shape.Width,
                0.75f,
                "Inline OLE became wider after OLE to OMML to OLE conversion.");
            AssertClose(
                expectedHeight,
                shape.Height,
                0.75f,
                "Inline OLE became taller after OLE to OMML to OLE conversion.");
            AssertClose(
                0f,
                ReadParagraphMarkFontPosition(shape.Range),
                0.1f,
                "Inline OLE round trip contaminated the paragraph baseline.");
            service.NormalizeInlineOleParagraphBaselinesBeforeSave(document);
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            Release(shape);
            shape = FindWordFormula(document, formulaId)
                ?? throw new InvalidOperationException(
                    "Saved inline OLE/OMML round-trip formula is missing.");
            AssertClose(
                expectedWidth,
                shape.Width,
                0.75f,
                "Saved inline OLE became wider after OLE to OMML to OLE conversion.");
            AssertClose(
                expectedHeight,
                shape.Height,
                0.75f,
                "Saved inline OLE became taller after OLE to OMML to OLE conversion.");
            AssertClose(
                0f,
                ReadParagraphMarkFontPosition(shape.Range),
                0.1f,
                "Saved inline OLE round trip contaminated the paragraph baseline.");
            Console.WriteLine(
                $"  inline round trip preserved {shape.Width:0.###}x{shape.Height:0.###} pt after save/reopen and recovered a trailing OMML anchor.");
        }
        finally
        {
            Release(equationRange);
            Release(bookmark);
            Release(shape);
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
        }
    }

    private static int RunInstalledFormulaFontSizeAcceptance(string artifactRoot)
    {
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "VisualTeX Formula OLE must be installed before running font-size acceptance.");

        Directory.CreateDirectory(artifactRoot);
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(previewRoot);
        var formulaId = Guid.NewGuid().ToString();
        var preview = CreatePreviewSet(previewRoot, formulaId, "font-size", 400, 100);
        var smallInlinePreview = CreatePreviewSet(
            previewRoot,
            Guid.NewGuid().ToString(),
            "inline-edit-small",
            25,
            11);
        var widerInlinePreview = CreatePreviewSet(
            previewRoot,
            Guid.NewGuid().ToString(),
            "inline-edit-wider",
            49,
            11);
        try
        {
            Console.WriteLine("[1/2] Verifying Word OMML and VisualTeX OLE font sizing...");
            VerifyWordFormulaFontSizing(
                Path.Combine(artifactRoot, "VisualTeX-Word-Formula-Font-Size.docx"),
                formulaId,
                preview,
                smallInlinePreview,
                widerInlinePreview);

            Console.WriteLine("[2/2] Verifying PowerPoint VisualTeX OLE font sizing...");
            VerifyPowerPointFormulaFontSizing(
                Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Formula-Font-Size.pptx"),
                Guid.NewGuid().ToString(),
                preview);

            Console.WriteLine("VisualTeX real Office formula font-size acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        finally
        {
            ForceComCleanup();
            TryDelete(preview.SvgPath);
            TryDelete(preview.EmfPath);
            TryDelete(preview.PngPath);
            TryDelete(smallInlinePreview.SvgPath);
            TryDelete(smallInlinePreview.EmfPath);
            TryDelete(smallInlinePreview.PngPath);
            TryDelete(widerInlinePreview.SvgPath);
            TryDelete(widerInlinePreview.EmfPath);
            TryDelete(widerInlinePreview.PngPath);
        }
    }

    private static int RunInstalledTallMatrixVectorAcceptance(string artifactRoot)
    {
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "VisualTeX Formula OLE must be installed before running tall-matrix vector acceptance.");

        Directory.CreateDirectory(artifactRoot);
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(previewRoot);
        var fixturePath = Path.Combine(previewRoot, $"tall-matrix-{Guid.NewGuid():N}.svg");
        File.WriteAllText(fixturePath, BuildTallMatrixMathJaxSvg(), new UTF8Encoding(false));
        var preview = CreatePreviewSetFromSvgFixture(
            previewRoot,
            fixturePath,
            $"tall-matrix-{Guid.NewGuid():N}");
        try
        {
            OfficeOlePreview.ValidateVectorEmf(preview.EmfPath);
            AssertTallMatrixBracketsContinuous(preview.PngPath);
            File.Copy(preview.SvgPath, Path.Combine(artifactRoot, "tall-matrix-mathjax.svg"), true);
            File.Copy(preview.EmfPath, Path.Combine(artifactRoot, "tall-matrix-vector.emf"), true);
            File.Copy(preview.PngPath, Path.Combine(artifactRoot, "tall-matrix-vector-replay.png"), true);

            var documentPath = Path.Combine(artifactRoot, "VisualTeX-Word-Tall-Matrix-Vector.docx");
            VerifyWordTallMatrixOle(documentPath, preview);
            Console.WriteLine("VisualTeX tall-matrix vector Office acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        finally
        {
            ForceComCleanup();
            TryDelete(fixturePath);
            TryDelete(preview.SvgPath);
            TryDelete(preview.EmfPath);
            TryDelete(preview.PngPath);
        }
    }

    private static string BuildTallMatrixMathJaxSvg()
    {
        const string glyphPath =
            "M60 120Q80 420 250 420Q355 420 390 330Q425 420 505 420Q555 420 555 370Q555 345 525 330Q505 325 490 345Q485 360 495 380Q470 390 450 370Q430 345 405 250L365 90Q355 45 380 35Q395 30 420 45L435 60L465 25Q415 -20 365 -20Q310 -20 280 25Q220 -20 155 -20Q75 -20 35 45Q5 95 60 120ZM285 335Q250 380 205 380Q145 380 120 300Q95 220 105 120Q115 35 170 20Q225 5 280 60Q305 90 315 135L350 275Q350 300 335 320Q320 335 285 335Z";
        var builder = new StringBuilder(24000);
        builder.Append("<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:xlink=\"http://www.w3.org/1999/xlink\" width=\"307.648\" height=\"269.8666666666666\" role=\"img\" focusable=\"false\" viewBox=\"-428.5714285714286 -7478.571428571428 16481.14285714286 14457.142857142857\">");
        builder.Append("<rect x=\"-428.5714285714286\" y=\"-7478.571428571428\" width=\"16481.14285714286\" height=\"14457.142857142857\" fill=\"#000000\" fill-opacity=\"0.001\"/>");
        builder.Append("<defs>");
        builder.Append("<path id=\"left-top\" d=\"M319 -645V1154H666V1070H403V-645H319Z\"/>");
        builder.Append("<path id=\"left-middle\" d=\"M319 0V602H403V0H319Z\"/>");
        builder.Append("<path id=\"left-bottom\" d=\"M319 -644V1155H403V-560H666V-644H319Z\"/>");
        builder.Append("<path id=\"right-top\" d=\"M0 1070V1154H347V-645H263V1070H0Z\"/>");
        builder.Append("<path id=\"right-middle\" d=\"M263 0V602H347V0H263Z\"/>");
        builder.Append("<path id=\"right-bottom\" d=\"M263 -560V1155H347V-644H0V-560H263Z\"/>");
        builder.Append("<path id=\"matrix-a\" d=\"").Append(glyphPath).Append("\"/>");
        builder.Append("</defs><g stroke=\"#111111\" fill=\"#111111\" stroke-width=\"0\" transform=\"scale(1,-1)\"><g>");
        builder.Append("<g><use href=\"#left-top\" transform=\"translate(0,5896)\"/><use href=\"#left-bottom\" transform=\"translate(0,-5906)\"/><svg width=\"667\" height=\"10202\" y=\"-4851\" x=\"0\" viewBox=\"0 2550.5 667 10202\"><use href=\"#left-middle\" transform=\"scale(1,25.42)\"/></svg></g>");
        builder.Append("<g transform=\"translate(667,0)\">");
        for (var row = 0; row < 10; row++)
        {
            var y = 6300 - row * 1400;
            builder.Append("<g transform=\"translate(0,").Append(y).Append(")\">");
            for (var column = 0; column < 10; column++)
            {
                var x = column * 1529;
                builder.Append("<use href=\"#matrix-a\" transform=\"translate(")
                    .Append(x)
                    .Append(",0)\"/>");
            }
            builder.Append("</g>");
        }
        builder.Append("</g>");
        builder.Append("<g transform=\"translate(14957,0)\"><use href=\"#right-top\" transform=\"translate(0,5896)\"/><use href=\"#right-bottom\" transform=\"translate(0,-5906)\"/><svg width=\"667\" height=\"10202\" y=\"-4851\" x=\"0\" viewBox=\"0 2550.5 667 10202\"><use href=\"#right-middle\" transform=\"scale(1,25.42)\"/></svg></g>");
        builder.Append("</g></g></svg>");
        return builder.ToString();
    }

    private static void AssertTallMatrixBracketsContinuous(string pngPath)
    {
        using var bitmap = new Bitmap(pngPath);
        var top = (int)Math.Round(bitmap.Height * 0.15);
        var bottom = (int)Math.Round(bitmap.Height * 0.85);
        var leftStart = (int)Math.Round(bitmap.Width * 0.02);
        var leftEnd = (int)Math.Round(bitmap.Width * 0.12);
        var rightStart = (int)Math.Round(bitmap.Width * 0.88);
        var rightEnd = (int)Math.Round(bitmap.Width * 0.98);
        var totalRows = bottom - top + 1;
        var leftRows = 0;
        var rightRows = 0;
        for (var y = top; y <= bottom; y++)
        {
            if (ContainsDarkPixel(bitmap, y, leftStart, leftEnd)) leftRows++;
            if (ContainsDarkPixel(bitmap, y, rightStart, rightEnd)) rightRows++;
        }
        var required = (int)Math.Floor(totalRows * 0.92);
        Assert(leftRows >= required,
            $"Tall matrix left bracket is broken: {leftRows}/{totalRows} continuous rows.");
        Assert(rightRows >= required,
            $"Tall matrix right bracket is broken: {rightRows}/{totalRows} continuous rows.");
        Console.WriteLine(
            $"  tall-matrix bracket coverage left={leftRows}/{totalRows}; right={rightRows}/{totalRows}.");
    }

    private static bool ContainsDarkPixel(Bitmap bitmap, int y, int left, int right)
    {
        left = Math.Max(0, left);
        right = Math.Min(bitmap.Width - 1, right);
        for (var x = left; x <= right; x++)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.A > 24 && pixel.R + pixel.G + pixel.B < 660) return true;
        }
        return false;
    }

    private static void VerifyWordTallMatrixOle(string path, PreviewSet preview)
    {
        const string latex = @"\begin{bmatrix}a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\\a&a&a&a&a&a&a&a&a&a\end{bmatrix}";
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mtext>10 by 10 matrix</mtext></math>";
        var formulaId = Guid.NewGuid().ToString();
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
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
            var session = CreateWordSession(
                formulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 14);
            session.DisplayMode = "block";
            var service = new WordFormulaService(application);
            service.InsertOle(session, preview.PngPath, preview.EmfPath);
            shape = FindWordFormula(document, formulaId)
                ?? throw new InvalidOperationException("Word tall-matrix OLE object was not created.");
            var expectedWidth = shape.Width;
            var expectedHeight = shape.Height;
            Assert(expectedHeight > 150f,
                $"Word tall-matrix OLE object is unexpectedly short: {expectedHeight:0.###} pt.");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            Release(shape);
            shape = FindWordFormula(document, formulaId)
                ?? throw new InvalidOperationException("Saved Word tall-matrix OLE object is missing.");
            AssertClose(expectedWidth, shape.Width, 0.75f,
                "Saved Word tall-matrix OLE object changed width.");
            AssertClose(expectedHeight, shape.Height, 0.75f,
                "Saved Word tall-matrix OLE object changed height.");
            Console.WriteLine(
                $"  Word tall-matrix OLE saved/reopened at {shape.Width:0.###}x{shape.Height:0.###} pt.");
        }
        finally
        {
            Release(shape);
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
        }
    }

    private static int RunInstalledMixedTextVectorAcceptance(string artifactRoot)
    {
        if (!HasExistingRegistration())
            throw new InvalidOperationException(
                "VisualTeX Formula OLE must be installed before running mixed-text vector acceptance.");

        Directory.CreateDirectory(artifactRoot);
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp");
        Directory.CreateDirectory(previewRoot);
        var fixturePath = Path.Combine(previewRoot, $"mixed-text-{Guid.NewGuid():N}.svg");
        File.WriteAllText(fixturePath, MixedTextMathJaxSvg, new UTF8Encoding(false));
        var preview = CreatePreviewSetFromSvgFixture(
            previewRoot,
            fixturePath,
            $"mixed-text-{Guid.NewGuid():N}");
        try
        {
            OfficeOlePreview.ValidateVectorEmf(preview.EmfPath);
            AssertMixedTextRegionRendered(preview.PngPath);
            File.Copy(preview.SvgPath, Path.Combine(artifactRoot, "mixed-text-mathjax.svg"), true);
            File.Copy(preview.EmfPath, Path.Combine(artifactRoot, "mixed-text-vector.emf"), true);
            File.Copy(preview.PngPath, Path.Combine(artifactRoot, "mixed-text-vector-replay.png"), true);

            Console.WriteLine("[1/2] Verifying Word mixed-text vector OLE save/reopen...");
            VerifyWordFormulaFontSizing(
                Path.Combine(artifactRoot, "VisualTeX-Word-Mixed-Text-Vector.docx"),
                Guid.NewGuid().ToString(),
                preview,
                preview,
                preview);

            Console.WriteLine("[2/2] Verifying PowerPoint mixed-text OLE/SVG save/reopen...");
            VerifyPowerPointFormulaFontSizing(
                Path.Combine(artifactRoot, "VisualTeX-PowerPoint-Mixed-Text-Vector.pptx"),
                Guid.NewGuid().ToString(),
                preview);

            Console.WriteLine("VisualTeX mixed-text vector Office acceptance passed.");
            Console.WriteLine($"Artifacts: {artifactRoot}");
            return 0;
        }
        finally
        {
            ForceComCleanup();
            TryDelete(fixturePath);
            TryDelete(preview.SvgPath);
            TryDelete(preview.EmfPath);
            TryDelete(preview.PngPath);
        }
    }

    private static void AssertMixedTextRegionRendered(string pngPath)
    {
        using var bitmap = new Bitmap(pngPath);
        var darkSamples = 0;
        var left = bitmap.Width / 3;
        for (var y = 0; y < bitmap.Height; y += 2)
        for (var x = left; x < bitmap.Width; x += 2)
        {
            var pixel = bitmap.GetPixel(x, y);
            if (pixel.A > 24 && pixel.R + pixel.G + pixel.B < 660) darkSamples++;
        }
        Assert(
            darkSamples > 150,
            $"Mixed-text EMF replay lost the Chinese glyph outlines: {darkSamples} dark samples in the text region.");
        Console.WriteLine($"  mixed-text replay contains {darkSamples} dark text-region samples.");
    }

    private const string MixedTextMathJaxSvg = """
        <svg xmlns="http://www.w3.org/2000/svg"
             xmlns:xlink="http://www.w3.org/1999/xlink"
             width="280" height="64"
             viewBox="0 -1100 7000 1600">
          <rect x="0" y="-1100" width="7000" height="1600"
                fill="#000000" fill-opacity="0.001"/>
          <defs>
            <path id="math-s" d="M80 300 Q180 520 420 430 Q570 370 450 250 Q360 170 210 190 Q90 205 75 100 Q65 25 185 10 Q390 -10 525 145"/>
            <path id="math-f" d="M120 -250 Q245 -250 285 -40 L390 500 Q430 700 625 700 Q770 700 825 595 Q850 545 815 490 Q780 445 725 465 Q665 485 675 555 Q680 610 735 635 Q690 670 625 655 Q520 630 490 490 L445 260 H650 L625 170 H425 L340 -275 Q315 -405 210 -500 Q135 -565 55 -545 Q-5 -530 5 -465 Q15 -410 70 -395 Q125 -380 150 -430 Q165 -455 150 -490 Q220 -455 240 -350 L345 170 H155 L180 260 H365 L410 500 Q445 700 625 745"/>
          </defs>
          <g stroke="#111111" fill="#111111" stroke-width="0" transform="scale(1,-1)">
            <g data-mml-node="math">
              <use xlink:href="#math-s"/>
              <use xlink:href="#math-f" transform="translate(650,0)"/>
              <use xlink:href="#math-s" transform="translate(1600,0)"/>
              <use xlink:href="#math-f" transform="translate(2250,0)"/>
              <g data-mml-node="TeXAtom" transform="translate(3200,0)">
                <text data-variant="normal" transform="scale(1,-1)"
                      font-size="1000px" font-family="serif">的地方地方</text>
              </g>
            </g>
          </g>
        </svg>
        """;

    private static void VerifyWordFormulaFontSizing(
        string path,
        string ommlFormulaId,
        PreviewSet preview,
        PreviewSet smallInlinePreview,
        PreviewSet widerInlinePreview)
    {
        const string latex = @"\frac{a}{b}+x^2";
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo>"
            + "<msup><mi>x</mi><mn>2</mn></msup></math>";
        const string alignedLatex =
            @"\begin{aligned} R_o' &= \frac{v_t}{i_d} \\ A_v &= -g_mR_d \\ R_i &= R_{g1}\parallel R_{g2} \end{aligned}";
        const string alignedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mtable displaystyle=\"true\" columnalign=\"right left\" columnspacing=\"0em\">"
            + "<mtr><mtd><msubsup><mi>R</mi><mi>o</mi><mo>&#x2032;</mo></msubsup></mtd>"
            + "<mtd><mi></mi><mo>=</mo><mfrac><msub><mi>v</mi><mi>t</mi></msub><msub><mi>i</mi><mi>d</mi></msub></mfrac></mtd></mtr>"
            + "<mtr><mtd><msub><mi>A</mi><mi>v</mi></msub></mtd>"
            + "<mtd><mi></mi><mo>=</mo><mo>&#x2212;</mo><msub><mi>g</mi><mi>m</mi></msub><msub><mi>R</mi><mi>d</mi></msub></mtd></mtr>"
            + "<mtr><mtd><msub><mi>R</mi><mi>i</mi></msub></mtd>"
            + "<mtd><mi></mi><mo>=</mo><msub><mi>R</mi><mrow><mi>g</mi><mn>1</mn></mrow></msub><mo>&#x2225;</mo><msub><mi>R</mi><mrow><mi>g</mi><mn>2</mn></mrow></msub></mtd></mtr>"
            + "</mtable></math>";
        const string compatibilityLatex =
            @"\nabla\times\bm F=\begin{vmatrix}\bm e_x&\bm e_y&\bm e_z\\\partial_x&\partial_y&\partial_z\\F_x&F_y&F_z\end{vmatrix}.";
        const string compatibilityMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi mathvariant=\"normal\">&#x2207;</mi><mo>&#xD7;</mo>"
            + "<mi mathvariant=\"bold-italic\">F</mi><mo>=</mo>"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">|</mo>"
            + "<mtable><mtr>"
            + "<mtd><msub><mi mathvariant=\"bold-italic\">e</mi><mi>x</mi></msub></mtd>"
            + "<mtd><msub><mi mathvariant=\"bold-italic\">e</mi><mi>y</mi></msub></mtd>"
            + "<mtd><msub><mi mathvariant=\"bold-italic\">e</mi><mi>z</mi></msub></mtd></mtr>"
            + "<mtr><mtd><msub><mi>&#x2202;</mi><mi>x</mi></msub></mtd>"
            + "<mtd><msub><mi>&#x2202;</mi><mi>y</mi></msub></mtd>"
            + "<mtd><msub><mi>&#x2202;</mi><mi>z</mi></msub></mtd></mtr>"
            + "<mtr><mtd><msub><mi>F</mi><mi>x</mi></msub></mtd>"
            + "<mtd><msub><mi>F</mi><mi>y</mi></msub></mtd>"
            + "<mtd><msub><mi>F</mi><mi>z</mi></msub></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\">|</mo></mrow><mo>.</mo></math>";
        const string compatibilityEditedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi mathvariant=\"normal\">&#x2207;</mi><mo>&#xD7;</mo>"
            + "<mi mathvariant=\"bold-italic\">F</mi><mo>=</mo>"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">|</mo>"
            + "<mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
            + "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\">|</mo></mrow><mo>+</mo><mn>1</mn></math>";
        var oleFormulaId = Guid.NewGuid().ToString();
        var legacyOleFormulaId = Guid.NewGuid().ToString();
        var inlineEditOleFormulaId = Guid.NewGuid().ToString();
        var alignedOmmlFormulaId = Guid.NewGuid().ToString();
        var compatibilityOmmlFormulaId = Guid.NewGuid().ToString();
        var blockOmmlFormulaId = Guid.NewGuid().ToString();
        var blockOleFormulaId = Guid.NewGuid().ToString();
        var verifyInlineEditExpansion =
            widerInlinePreview.Width > smallInlinePreview.Width * 1.2f;
        var expectedLegacyInlinePosition = 0f;
        var editedInlineWidth = 0f;
        var editedInlineHeight = 0f;
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.Font? font = null;
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
            selection.Font.Size = 24f;
            selection.TypeText("试一下字号：");
            var service = new WordFormulaService(application);

            var ommlSession = CreateWordSession(
                ommlFormulaId,
                "create",
                FormulaOleContract.WordOmmlMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 14);
            ommlSession.DisplayMode = "inline";
            service.InsertOmml(ommlSession, mathMl);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, ommlFormulaId)
                ?? throw new InvalidOperationException("Word OMML font-size bookmark was not created.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            selection.SetRange(range.Start, range.End);
            AssertClose(18f, service.SetSelectedFormulaFontSize(18), 0.01f,
                "Word OMML font-size command returned the wrong point size.");
            Release(font);
            font = range.Font;
            AssertClose(18f, font.Size, 0.01f,
                "Word OMML range did not use the requested 18 pt base size.");
            var stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidOperationException("Word OMML font-size metadata was not persisted.");
            AssertClose(18f, FormulaFontSize.ResolveSemanticFontSize(stored), 0.01f,
                "Word OMML metadata did not preserve the requested size.");
            AssertClose(0f, application.Selection.Font.Position, 0.1f,
                "Word OMML font sizing left the typing caret vertically shifted.");

            Release(range);
            range = null;
            Release(bookmark);
            bookmark = null;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText("这里呢：");
            var oleSession = CreateWordSession(
                oleFormulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 14);
            oleSession.DisplayMode = "inline";
            service.InsertOle(oleSession, preview.PngPath, preview.EmfPath);
            shape = FindWordFormula(document, oleFormulaId)
                ?? throw new InvalidOperationException("Word OLE font-size object was not created.");
            range = shape.Range;
            selection.SetRange(range.Start, range.End);
            var metadata = oleSession.ToMetadata();
            var expectedOle = FormulaFontSize.OleSizeAt(metadata, 21);
            service.SetSelectedFormulaFontSize(21);
            AssertClose(expectedOle.Width, shape.Width, 0.75f,
                "Word OLE width does not match the requested semantic font size.");
            AssertClose(expectedOle.Height, shape.Height, 0.75f,
                "Word OLE height does not match the requested semantic font size.");
            var expectedInlinePosition = WordInlineAlignment.CalculateFontPosition(
                shape.Height,
                (float)(metadata.RenderHeightPx ?? 0),
                metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null);
            AssertClose(expectedInlinePosition, shape.Range.Font.Position, 0.1f,
                "Word OLE inline baseline was lost after changing font size.");
            AssertClose(0f, ReadParagraphMarkFontPosition(shape.Range), 0.1f,
                "Word OLE font sizing contaminated the paragraph default baseline.");
            AssertClose(0f, application.Selection.Font.Position, 0.1f,
                "Word OLE font sizing left the typing caret vertically shifted.");

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText("旧公式：");
            var legacySession = CreateWordSession(
                legacyOleFormulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 14);
            legacySession.DisplayMode = "inline";
            legacySession.ExportResult!.Baseline = null;
            service.InsertOle(legacySession, preview.PngPath, preview.EmfPath);
            Release(shape);
            shape = FindWordFormula(document, legacyOleFormulaId)
                ?? throw new InvalidOperationException("Word legacy OLE font-size object was not created.");
            var legacyInitialPosition = shape.Range.Font.Position;
            Assert(
                legacyInitialPosition < 0,
                "Word legacy OLE insertion did not receive a fallback inline baseline.");
            shape.Range.Select();
            service.SetSelectedFormulaFontSize(42);
            expectedLegacyInlinePosition = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                shape.Height,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: legacyInitialPosition,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42);
            AssertClose(
                expectedLegacyInlinePosition,
                shape.Range.Font.Position,
                0.1f,
                "Word legacy OLE formula did not scale its existing baseline when metadata was missing.");
            AssertClose(0f, ReadParagraphMarkFontPosition(shape.Range), 0.1f,
                "Word legacy OLE font sizing contaminated the paragraph default baseline.");

            if (verifyInlineEditExpansion)
            {
                selection.EndKey(Word.WdUnits.wdStory);
                selection.TypeParagraph();
                selection.TypeText("批量行内 OLE 双击编辑：");
                var inlineEditCreateSession = CreateWordSession(
                    inlineEditOleFormulaId,
                    "create",
                    FormulaOleContract.NativeOleMode,
                    "v_{GS}",
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mi>v</mi><mi>GS</mi></msub></math>",
                    smallInlinePreview,
                    numbered: false,
                    originalMetadata: null,
                    fontSizePt: 11);
                inlineEditCreateSession.DisplayMode = "inline";
                service.InsertOle(
                    inlineEditCreateSession,
                    smallInlinePreview.PngPath,
                    smallInlinePreview.EmfPath);
                Release(shape);
                shape = FindWordFormula(document, inlineEditOleFormulaId)
                    ?? throw new InvalidOperationException("Word small inline OLE edit fixture was not created.");
                var initialInlineWidth = shape.Width;
                var initialInlineHeight = shape.Height;

                // Follow the actual double-click path. The add-in selects the
                // existing object and calls ReadSelection before creating the
                // edit Session. The previous acceptance bypassed this step, so
                // it never exercised small-formula font-size inference.
                shape.Range.Select();
                var capturedSelection = service.ReadSelection();
                var capturedMetadata = capturedSelection.Metadata
                    ?? throw new InvalidOperationException("Word inline OLE double-click metadata is missing.");
                AssertClose(
                    11f,
                    FormulaFontSize.ResolveSemanticFontSize(capturedMetadata),
                    0.01f,
                    "Selecting a small inline OLE formula changed 11 pt into a smaller edit-session size.");

                var inlineEditSession = CreateWordSession(
                    inlineEditOleFormulaId,
                    "edit",
                    FormulaOleContract.NativeOleMode,
                    "v_{GS}111111",
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><msub><mi>v</mi><mi>GS</mi></msub><mn>111111</mn></mrow></math>",
                    widerInlinePreview,
                    numbered: false,
                    originalMetadata: capturedMetadata,
                    fontSizePt: FormulaFontSize.ResolveSemanticFontSize(capturedMetadata));
                inlineEditSession.DisplayMode = "inline";
                service.ReplaceOle(
                    inlineEditSession,
                    widerInlinePreview.PngPath,
                    widerInlinePreview.EmfPath);
                Release(shape);
                shape = FindWordFormula(document, inlineEditOleFormulaId)
                    ?? throw new InvalidOperationException("Edited Word inline OLE formula is missing.");
                editedInlineWidth = shape.Width;
                editedInlineHeight = shape.Height;
                Assert(
                    editedInlineWidth >= initialInlineWidth * 1.5f,
                    "Editing a wider small inline OLE formula kept the old width and shrank its content. "
                    + $"Before={initialInlineWidth:0.##} pt; After={editedInlineWidth:0.##} pt.");
                AssertClose(
                    initialInlineHeight,
                    editedInlineHeight,
                    0.75f,
                    "Editing a wider small inline OLE formula changed its 11 pt glyph height.");
                var editedMetadata = WordFormulaMetadataReader.TryRead(shape)
                    ?? throw new InvalidOperationException("Edited Word inline OLE metadata is missing.");
                AssertClose(
                    11f,
                    FormulaFontSize.ResolveSemanticFontSize(editedMetadata),
                    0.01f,
                    "Edited small inline OLE formula did not preserve its semantic 11 pt size.");
                AssertClose(0f, ReadParagraphMarkFontPosition(shape.Range), 0.1f,
                    "Editing a wider inline OLE formula contaminated the paragraph baseline.");
            }

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText("OMML aligned 对齐：");
            var alignedSession = CreateWordSession(
                alignedOmmlFormulaId,
                "create",
                FormulaOleContract.WordOmmlMode,
                alignedLatex,
                alignedMathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 11);
            alignedSession.DisplayMode = "block";
            service.InsertOmml(alignedSession, alignedMathMl);
            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, alignedOmmlFormulaId)
                ?? throw new InvalidOperationException("Word aligned OMML bookmark was not created.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertOmmlMatrixColumnAlignment(
                range,
                new[] { "right", "left" },
                expectedRows: 3,
                "Inserted Word aligned OMML did not preserve the ampersand alignment columns.");

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText("直接 OMML 兼容链：");
            var compatibilitySession = CreateWordSession(
                compatibilityOmmlFormulaId,
                "create",
                FormulaOleContract.WordOmmlMode,
                compatibilityLatex,
                compatibilityMathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 11);
            compatibilitySession.DisplayMode = "block";
            service.InsertOmml(compatibilitySession, compatibilityMathMl);
            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, compatibilityOmmlFormulaId)
                ?? throw new InvalidOperationException("Direct compatibility OMML bookmark was not created.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertNoLiteralLatexCommandsInOmml(range.WordOpenXML, compatibilityOmmlFormulaId);
            AssertNativeOmmlVmatrix(range.WordOpenXML, compatibilityOmmlFormulaId);
            var compatibilityInitialPath = Path.Combine(
                Path.GetDirectoryName(path)!,
                "VisualTeX-Word-Compatibility-OMML-Initial.docx");
            service.NormalizeInlineOleParagraphBaselinesBeforeSave(document);
            document.SaveAs2(
                compatibilityInitialPath,
                Word.WdSaveFormat.wdFormatXMLDocument);

            // Follow the same selection/read/edit/write-back sequence used by a
            // real double-click edit, rather than only testing initial insert.
            range.Select();
            var compatibilitySource = service.ReadSelection();
            var compatibilityMetadata = compatibilitySource.Metadata
                ?? throw new InvalidOperationException("Direct compatibility OMML edit metadata is missing.");
            var compatibilityEditSession = CreateWordSession(
                compatibilityOmmlFormulaId,
                "edit",
                FormulaOleContract.WordOmmlMode,
                compatibilityLatex + "+1",
                compatibilityEditedMathMl,
                preview,
                numbered: false,
                originalMetadata: compatibilityMetadata,
                fontSizePt: FormulaFontSize.ResolveSemanticFontSize(compatibilityMetadata));
            compatibilityEditSession.DisplayMode = "block";
            service.ReplaceOmml(compatibilityEditSession, compatibilityEditedMathMl);
            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, compatibilityOmmlFormulaId)
                ?? throw new InvalidOperationException("Direct compatibility OMML disappeared after edit/write-back.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertNoLiteralLatexCommandsInOmml(range.WordOpenXML, compatibilityOmmlFormulaId);
            AssertNativeOmmlVmatrix(range.WordOpenXML, compatibilityOmmlFormulaId);

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var blockOmmlLeadStart = selection.Start;
            selection.TypeText("编号 OMML 行间公式前文");
            selection.ParagraphFormat.SpaceAfter = 8f;
            var blockOmmlSession = CreateWordSession(
                blockOmmlFormulaId,
                "create",
                FormulaOleContract.WordOmmlMode,
                latex,
                mathMl,
                preview,
                numbered: true,
                originalMetadata: null,
                fontSizePt: 14);
            service.InsertOmml(blockOmmlSession, mathMl);
            AssertParagraphSpaceAfterAt(
                document,
                blockOmmlLeadStart,
                8f,
                "Word numbered display OMML did not retain native Word prose spacing.");
            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, blockOmmlFormulaId)
                ?? throw new InvalidOperationException("Word display OMML font-size bookmark was not created.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            range.Select();
            service.SetSelectedFormulaFontSize(24);
            AssertClose(0f, range.Font.Position, 0.1f,
                "Word display OMML formula inherited an inline baseline offset.");
            Assert(
                range.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter,
                "Word display OMML formula paragraph is not centered after font sizing.");
            AssertEquationNumberFontSize(document, blockOmmlFormulaId, 24f,
                "Word display OMML equation number did not follow the formula font size.");

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            var blockOleLeadStart = selection.Start;
            selection.TypeText("编号 OLE 行间公式前文");
            selection.ParagraphFormat.SpaceAfter = 8f;
            var blockOleSession = CreateWordSession(
                blockOleFormulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: true,
                originalMetadata: null,
                fontSizePt: 14);
            service.InsertOle(blockOleSession, preview.PngPath, preview.EmfPath);
            AssertParagraphSpaceAfterAt(
                document,
                blockOleLeadStart,
                0f,
                "Word numbered display OLE did not compact the preceding prose spacing.");
            Release(shape);
            shape = FindWordFormula(document, blockOleFormulaId)
                ?? throw new InvalidOperationException("Word display OLE font-size object was not created.");
            shape.Range.Select();
            service.SetSelectedFormulaFontSize(24);
            AssertClose(0f, shape.Range.Font.Position, 0.1f,
                "Word display OLE formula inherited an inline baseline offset.");
            Assert(
                shape.Range.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter,
                "Word display OLE formula paragraph is not centered after font sizing.");
            AssertEquationNumberFontSize(document, blockOleFormulaId, 24f,
                "Word display OLE equation number did not follow the formula font size.");

            service.NormalizeInlineOleParagraphBaselinesBeforeSave(document);
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            Release(selection);
            selection = application.Selection;
            var reopenedService = new WordFormulaService(application);
            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, ommlFormulaId)
                ?? throw new InvalidOperationException("Saved Word OMML font-size bookmark is missing.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            Release(font);
            font = range.Font;
            AssertClose(18f, font.Size, 0.01f,
                "Saved Word OMML formula did not preserve 18 pt.");

            Release(shape);
            shape = FindWordFormula(document, oleFormulaId)
                ?? throw new InvalidOperationException("Saved Word OLE font-size object is missing.");
            AssertClose(expectedOle.Height, shape.Height, 0.75f,
                "Saved Word OLE formula did not preserve its physical font-size scale.");
            AssertClose(expectedInlinePosition, ReadPersistedRunFontPosition(shape), 0.1f,
                "Saved Word OLE formula did not preserve its inline baseline.");
            AssertClose(0f, ReadParagraphMarkFontPosition(shape.Range), 0.1f,
                "Saved Word OLE formula contaminated the paragraph default baseline.");
            Release(range);
            range = shape.Range;
            document.Activate();
            range.Select();
            AssertClose(21f, reopenedService.GetSelectedFormulaFontSize() ?? 0f, 0.5f,
                "Saved Word OLE semantic font size could not be inferred from geometry.");

            Release(shape);
            shape = FindWordFormula(document, legacyOleFormulaId)
                ?? throw new InvalidOperationException("Saved legacy Word OLE font-size object is missing.");
            AssertClose(
                expectedLegacyInlinePosition,
                ReadPersistedRunFontPosition(shape),
                0.1f,
                "Saved legacy Word OLE formula did not preserve its scaled inline baseline.");
            AssertClose(0f, ReadParagraphMarkFontPosition(shape.Range), 0.1f,
                "Saved legacy Word OLE formula contaminated the paragraph default baseline.");

            if (verifyInlineEditExpansion)
            {
                Release(shape);
                shape = FindWordFormula(document, inlineEditOleFormulaId)
                    ?? throw new InvalidOperationException("Saved edited inline Word OLE formula is missing.");
                AssertClose(
                    editedInlineWidth,
                    shape.Width,
                    0.75f,
                    "Saved edited inline Word OLE formula did not preserve its expanded width.");
                AssertClose(
                    editedInlineHeight,
                    shape.Height,
                    0.75f,
                    "Saved edited inline Word OLE formula did not preserve its glyph height.");
                var savedInlineEditMetadata = WordFormulaMetadataReader.TryRead(shape)
                    ?? throw new InvalidOperationException("Saved edited inline Word OLE metadata is missing.");
                Assert(
                    Math.Abs((savedInlineEditMetadata.RenderWidthPx ?? 0) - widerInlinePreview.Width) <= 1,
                    "Saved edited inline Word OLE formula retained the old render width metadata.");
            }

            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, alignedOmmlFormulaId)
                ?? throw new InvalidOperationException("Saved Word aligned OMML bookmark is missing.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertOmmlMatrixColumnAlignment(
                range,
                new[] { "right", "left" },
                expectedRows: 3,
                "Saved Word aligned OMML lost the ampersand alignment columns.");

            Release(bookmark);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, compatibilityOmmlFormulaId)
                ?? throw new InvalidOperationException("Saved direct compatibility OMML bookmark is missing.");
            Release(range);
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            AssertNoLiteralLatexCommandsInOmml(range.WordOpenXML, compatibilityOmmlFormulaId);
            AssertNativeOmmlVmatrix(range.WordOpenXML, compatibilityOmmlFormulaId);
            range.Select();
            var reopenedCompatibilitySource = reopenedService.ReadSelection();
            Assert(
                reopenedCompatibilitySource.Metadata is not null
                && reopenedCompatibilitySource.Metadata.Latex.EndsWith("+1", StringComparison.Ordinal),
                "Saved direct compatibility OMML did not preserve edited LaTeX metadata.");

            AssertEquationNumberFontSize(document, blockOmmlFormulaId, 24f,
                "Saved Word display OMML equation number did not preserve its formula font size.");
            AssertEquationNumberFontSize(document, blockOleFormulaId, 24f,
                "Saved Word display OLE equation number did not preserve its formula font size.");
            AssertTextParagraphSpaceAfter(
                document,
                "编号 OMML 行间公式前文",
                8f,
                "Saved Word numbered display OMML did not preserve native Word spacing.");
            AssertTextParagraphSpaceAfter(
                document,
                "编号 OLE 行间公式前文",
                0f,
                "Saved Word numbered display OLE restored loose preceding spacing.");
        }
        finally
        {
            Release(font);
            Release(range);
            Release(bookmark);
            Release(shape);
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
        }
    }

    private static void VerifyPowerPointFormulaFontSizing(
        string path,
        string formulaId,
        PreviewSet preview)
    {
        const string latex = @"\frac{a}{b}+x^2";
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo>"
            + "<msup><mi>x</mi><mn>2</mn></msup></math>";
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPoint.Shape? svgShape = null;
        var svgFormulaId = Guid.NewGuid().ToString();
        try
        {
            application = new PowerPoint.Application
            {
                Visible = Office.MsoTriState.msoTrue,
            };
            presentation = application.Presentations.Add(Office.MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            var service = new PowerPointFormulaService(application);
            var session = CreateWordSession(
                formulaId,
                "create",
                FormulaOleContract.NativeOleMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 14);
            session.Host = "powerpoint";
            var result = service.InsertOle(session, preview.PngPath, preview.EmfPath);
            shape = slide.Shapes[result.ObjectId];
            var centerX = shape.Left + shape.Width / 2f;
            var centerY = shape.Top + shape.Height / 2f;
            shape.Select();
            var metadata = session.ToMetadata();
            var expected = FormulaFontSize.OleSizeAt(metadata, 24, 600f, 400f);
            service.SetSelectedFormulaFontSize(24);
            AssertClose(expected.Width, shape.Width, 0.75f,
                "PowerPoint OLE width does not match the requested semantic font size.");
            AssertClose(expected.Height, shape.Height, 0.75f,
                "PowerPoint OLE height does not match the requested semantic font size.");
            AssertClose(centerX, shape.Left + shape.Width / 2f, 0.75f,
                "PowerPoint OLE horizontal center moved during font sizing.");
            AssertClose(centerY, shape.Top + shape.Height / 2f, 0.75f,
                "PowerPoint OLE vertical center moved during font sizing.");

            var svgSession = CreateWordSession(
                svgFormulaId,
                "create",
                FormulaOleContract.CrossPlatformPictureMode,
                latex,
                mathMl,
                preview,
                numbered: false,
                originalMetadata: null,
                fontSizePt: 18);
            svgSession.Host = "powerpoint";
            var svgResult = service.Insert(svgSession, preview.SvgPath);
            svgShape = slide.Shapes[svgResult.ObjectId];
            Assert(svgShape.Type != Office.MsoShapeType.msoEmbeddedOLEObject,
                "PowerPoint SVG formula was inserted as an OLE object.");
            Assert(svgShape.Type != Office.MsoShapeType.msoLinkedPicture,
                "PowerPoint SVG formula was linked instead of embedded.");

            presentation.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            presentation.Close();
            Release(presentation);
            presentation = application.Presentations.Open(
                path,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue);
            Release(slide);
            slide = presentation.Slides[1];
            Release(shape);
            shape = slide.Shapes[$"VisualTeX_{formulaId}"];
            AssertClose(expected.Width, shape.Width, 0.75f,
                "Saved PowerPoint OLE formula did not preserve width.");
            AssertClose(expected.Height, shape.Height, 0.75f,
                "Saved PowerPoint OLE formula did not preserve height.");
            Release(svgShape);
            svgShape = slide.Shapes[$"VisualTeX_{svgFormulaId}"];
            Assert(svgShape.Type != Office.MsoShapeType.msoEmbeddedOLEObject,
                "Saved PowerPoint SVG formula reopened as an OLE object.");

            presentation.Close();
            Release(presentation);
            presentation = null;
            AssertPowerPointPackageContainsEmbeddedSvg(path);
        }
        finally
        {
            Release(svgShape);
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
        }
    }

    private static void AssertPowerPointPackageContainsEmbeddedSvg(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        var svgEntries = archive.Entries
            .Where(entry => entry.FullName.StartsWith("ppt/media/", StringComparison.OrdinalIgnoreCase)
                && entry.FullName.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
            .ToList();
        Assert(svgEntries.Count > 0,
            "PowerPoint package does not contain an embedded SVG under ppt/media.");
        var contentTypes = archive.GetEntry("[Content_Types].xml")
            ?? throw new InvalidOperationException(
                "ACCEPTANCE FAILURE: PowerPoint package has no [Content_Types].xml.");
        using var reader = new StreamReader(contentTypes.Open(), Encoding.UTF8, true);
        var xml = reader.ReadToEnd();
        Assert(xml.IndexOf("image/svg+xml", StringComparison.OrdinalIgnoreCase) >= 0,
            "PowerPoint package does not declare the image/svg+xml content type.");
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
            FontSizePt = FormulaFontSize.DefaultPt,
            RenderFontSizePt = FormulaFontSize.DefaultPt,
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

    private static void VerifyWordOleNumberingPromotion(
        string artifactRoot,
        PreviewSet initial,
        PreviewSet updated)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>x</mi><mo>=</mo><mn>1</mn></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>y</mi><mo>=</mo><mn>2</mn></math>";
        const string oleMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>u</mi><mo>+</mo><mi>v</mi><mo>=</mo><mi>w</mi></math>";

        var firstId = Guid.NewGuid().ToString();
        var secondId = Guid.NewGuid().ToString();
        var promotedOleId = Guid.NewGuid().ToString();
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
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
                    "x=1",
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
            service.InsertOmml(
                CreateWordSession(
                    secondId,
                    "create",
                    FormulaOleContract.WordOmmlMode,
                    "y=2",
                    secondMathMl,
                    initial,
                    numbered: true,
                    originalMetadata: null),
                secondMathMl);

            ResetWordSelection(
                application,
                ref selection,
                document.Content.End - 1,
                document.Content.End - 1);
            service.InsertOle(
                CreateWordSession(
                    promotedOleId,
                    "create",
                    FormulaOleContract.NativeOleMode,
                    "u+v=w",
                    oleMathMl,
                    initial,
                    numbered: false,
                    originalMetadata: null),
                initial.PngPath,
                initial.EmfPath);
            AssertEqual(
                "2",
                WordEquationNumbering.Reconcile(document).ToString(),
                "The unnumbered OLE fixture unexpectedly participated in numbering.");

            shape = FindWordFormula(document, promotedOleId)
                ?? throw new InvalidOperationException(
                    "Unnumbered OLE promotion fixture is missing before edit.");
            range = shape.Range;
            var selected = ReadSelectedFormula(application, service, range);
            var originalMetadata = selected.Metadata
                ?? throw new InvalidOperationException(
                    "Unnumbered OLE promotion fixture returned no metadata.");
            Assert(
                !originalMetadata.Numbered,
                "Unnumbered OLE promotion fixture unexpectedly started numbered.");
            Release(range);
            range = null;
            Release(shape);
            shape = null;

            service.ReplaceOle(
                CreateWordSession(
                    promotedOleId,
                    "edit",
                    FormulaOleContract.NativeOleMode,
                    "u+v=w+1",
                    oleMathMl,
                    updated,
                    numbered: true,
                    originalMetadata),
                updated.PngPath,
                updated.EmfPath);

            Assert(
                document.InlineShapes.Count == 1,
                "Editing an unnumbered OLE to numbered duplicated or deleted the formula.");
            shape = FindWordFormula(document, promotedOleId)
                ?? throw new InvalidOperationException(
                    "Editing an unnumbered OLE to numbered deleted the formula.");
            range = shape.Range;
            Assert(
                (bool)range.get_Information(Word.WdInformation.wdWithInTable),
                "Editing an unnumbered OLE to numbered left it outside the numbered table.");
            var refreshed = ReadSelectedFormula(application, service, range);
            Assert(
                refreshed.Metadata?.Numbered == true,
                "Editing an unnumbered OLE to numbered did not persist Numbered=true.");
            AssertEqual(
                "3",
                WordEquationNumbering.Reconcile(document).ToString(),
                "OLE numbering promotion did not expose three numbered formulas.");
            var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            AssertTargetNumbers(
                targets,
                (firstId, "1"),
                (secondId, "2"),
                (promotedOleId, "3"));

            Directory.CreateDirectory(artifactRoot);
            var path = Path.Combine(
                artifactRoot,
                "word-office2019-unnumbered-ole-edit-to-numbered.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "  Word OLE numbering promotion: unnumbered display OLE edited to "
                + "Numbered=true, one OLE retained, table migration and 1/2/3 sequence passed.");
        }
        finally
        {
            Release(range);
            Release(shape);
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
        FormulaMetadata? originalMetadata,
        double fontSizePt = FormulaFontSize.DefaultPt)
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
            FontSizePt = FormulaFontSize.Normalize(fontSizePt),
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
        var found = false;
        foreach (var hive in new[] { RegistryHive.CurrentUser, RegistryHive.LocalMachine })
        {
            using var baseKey = RegistryKey.OpenBaseKey(hive, view);
            using var classes = baseKey.OpenSubKey("Software\\Classes");
            using var progId = classes?.OpenSubKey(FormulaOleContract.ProgId);
            using var clsid = classes?.OpenSubKey($"CLSID\\{FormulaClsid}");
            using var typeLib = classes?.OpenSubKey($"TypeLib\\{TypeLibraryId}");
            var hiveFound = progId is not null || clsid is not null || typeLib is not null;
            if (!hiveFound) continue;
            found = true;
            Console.Error.WriteLine(
                $"Registration precheck ({hive}, {view}): ProgID={progId is not null}, CLSID={clsid is not null}, TypeLib={typeLib is not null}.");
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

    private static float ReadPersistedRunFontPosition(Word.InlineShape shape)
    {
        var document = XDocument.Parse(shape.Range.WordOpenXML);
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var position = document
            .Descendants(word + "position")
            .Select(element => element.Attribute(word + "val")?.Value)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
        if (!int.TryParse(position, NumberStyles.Integer, CultureInfo.InvariantCulture, out var halfPoints))
            return 0f;
        return halfPoints / 2f;
    }

    private static float ReadParagraphMarkFontPosition(Word.Range formulaRange)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? mark = null;
        Word.Font? font = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            mark = paragraphRange.Duplicate;
            mark.SetRange(paragraphRange.End - 1, paragraphRange.End);
            font = mark.Font;
            return font.Position == (int)Word.WdConstants.wdUndefined ? 0f : font.Position;
        }
        finally
        {
            Release(font);
            Release(mark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void AssertOmmlMatrixColumnAlignment(
        Word.Range formulaRange,
        IReadOnlyList<string> expectedAlignments,
        int expectedRows,
        string message)
    {
        var package = XDocument.Parse(
            formulaRange.WordOpenXML,
            LoadOptions.PreserveWhitespace);
        XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var matrix = package
            .Descendants(math + "m")
            .FirstOrDefault(candidate =>
            {
                var rows = candidate.Elements(math + "mr").ToArray();
                return (expectedRows <= 0 || rows.Length == expectedRows)
                    && rows.Length > 0
                    && rows.All(row =>
                        row.Elements(math + "e").Count() == expectedAlignments.Count);
            });
        Assert(matrix is not null, message + " No matching native OMML matrix was found.");

        var actual = new List<string>();
        var columns = matrix!
            .Element(math + "mPr")?
            .Element(math + "mcs")?
            .Elements(math + "mc")
            .ToArray() ?? Array.Empty<XElement>();
        foreach (var column in columns)
        {
            var properties = column.Element(math + "mcPr");
            var countText = properties?
                .Element(math + "count")?
                .Attribute(math + "val")?
                .Value;
            var count = int.TryParse(countText, out var parsedCount)
                ? Math.Max(1, parsedCount)
                : 1;
            var alignment = properties?
                .Element(math + "mcJc")?
                .Attribute(math + "val")?
                .Value ?? "center";
            for (var index = 0; index < count; index++) actual.Add(alignment);
        }

        Assert(
            actual.SequenceEqual(expectedAlignments, StringComparer.OrdinalIgnoreCase),
            message
            + $" Expected=[{string.Join(",", expectedAlignments)}]; "
            + $"Actual=[{string.Join(",", actual)}]; "
            + $"OMML={matrix.ToString(SaveOptions.DisableFormatting)}");
    }

    private static void AssertEquationNumberFontSize(
        Word.Document document,
        string formulaId,
        float expectedFontSize,
        string message)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.Font? font = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = WordEquationNumbering.EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name))
                throw new InvalidOperationException($"ACCEPTANCE FAILURE: {message} Number bookmark '{name}' is missing.");
            bookmark = bookmarks[name];
            range = bookmark.Range;
            font = range.Font;
            AssertClose(expectedFontSize, font.Size, 0.1f, message);
        }
        finally
        {
            Release(font);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AssertTextParagraphSpaceAfter(
        Word.Document document,
        string text,
        float expected,
        string message)
    {
        Word.Range? range = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        try
        {
            range = FindNativeTextRange(document, text);
            Assert(range is not null, message + " Text is missing: " + text);
            paragraphs = range!.Paragraphs;
            Assert(paragraphs.Count > 0, message + " Paragraph is missing.");
            paragraph = paragraphs[1];
            format = paragraph.Format;
            AssertClose(expected, format.SpaceAfter, 0.1f, message);
        }
        finally
        {
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }

    private static void AssertParagraphSpaceAfterAt(
        Word.Document document,
        int position,
        float expected,
        string message)
    {
        Word.Range? anchor = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        try
        {
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            Assert(paragraphs.Count > 0, message + " Paragraph is missing.");
            paragraph = paragraphs[1];
            format = paragraph.Format;
            AssertClose(expected, format.SpaceAfter, 0.1f, message);
        }
        finally
        {
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
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

    private sealed class WordBaselineVisualRow : IDisposable
    {
        public WordBaselineVisualRow(
            string label,
            Word.Range bodyRange,
            Word.Range latinRange,
            Word.Range formulaRange)
        {
            Label = label;
            BodyRange = bodyRange;
            LatinRange = latinRange;
            FormulaRange = formulaRange;
        }

        public string Label { get; }
        public Word.Range BodyRange { get; }
        public Word.Range LatinRange { get; }
        public Word.Range FormulaRange { get; }

        public void Dispose()
        {
            Release(FormulaRange);
            Release(LatinRange);
            Release(BodyRange);
        }
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

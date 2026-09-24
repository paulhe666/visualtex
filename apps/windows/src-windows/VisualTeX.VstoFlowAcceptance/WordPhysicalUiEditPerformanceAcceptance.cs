using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Windows.Automation;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const int PhysicalUiPerfInlineCount = 100;
    private const int PhysicalUiPerfDisplayCount = 50;
    private const int PhysicalUiPerfTotalCount =
        PhysicalUiPerfInlineCount + PhysicalUiPerfDisplayCount;
    private const double PhysicalUiPerfUnnumberedLimitMs = 700d;
    private const double PhysicalUiPerfNumberedLimitMs = 1000d;

    private sealed class PhysicalUiEditMetric
    {
        public int Index { get; set; }
        public string Format { get; set; } = string.Empty;
        public string Scenario { get; set; } = string.Empty;
        public string DisplayMode { get; set; } = string.Empty;
        public bool Numbered { get; set; }
        public double OpenVisibleMilliseconds { get; set; }
        public double OpenReadyMilliseconds { get; set; }
        public double ApplyHiddenMilliseconds { get; set; }
    }

    private static void RunWordPhysicalUiEditPerformanceAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var format = (
                Environment.GetEnvironmentVariable("VISUALTEX_UI_EDIT_PERF_FORMAT")
                ?? "visualtex")
            .Trim()
            .ToLowerInvariant();
        var scenario = (
                Environment.GetEnvironmentVariable("VISUALTEX_UI_EDIT_PERF_SCENARIO")
                ?? "prepare")
            .Trim()
            .ToLowerInvariant();

        if (format is not ("visualtex" or "mathtype" or "omml"))
            throw new ArgumentOutOfRangeException(
                nameof(format),
                format,
                "VISUALTEX_UI_EDIT_PERF_FORMAT must be visualtex, mathtype, or omml.");

        if (scenario == "prepare")
        {
            PreparePhysicalUiEditPerformanceFixtures(format, artifactRoot);
            return;
        }

        if (scenario is "create-inline" or "create-display" or "create-numbered")
        {
            MeasurePhysicalUiCreatePerformance(client, format, scenario, artifactRoot);
            return;
        }

        if (scenario is not ("inline" or "display" or "numbered"))
            throw new ArgumentOutOfRangeException(
                nameof(scenario),
                scenario,
                "VISUALTEX_UI_EDIT_PERF_SCENARIO must be prepare, inline, display, numbered, create-inline, create-display, or create-numbered.");

        MeasurePhysicalUiEditPerformance(client, format, scenario, artifactRoot);
    }

    private static void PreparePhysicalUiEditPerformanceFixtures(
        string format,
        string artifactRoot)
    {
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var previousNumberPosition = WordEquationNumbering.GetDefaultMathTypeNumberPosition();
        Word.Application? application = null;
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"physical-ui-edit-perf-{format}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);

        var genericLatex = @"x_{1}=1";
        var genericSvg = PerformanceSvg(1);
        var assets = CreatePerformanceOleAssets(
            assetRoot,
            1,
            genericLatex,
            genericSvg);
        var wmfPath = Path.Combine(assetRoot, "formula-001.wmf");
        File.WriteAllBytes(
            wmfPath,
            MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                assets.EmfPath,
                widthPt: 120,
                heightPt: 36));

        try
        {
            var unnumberedPath = PhysicalUiPerfFixturePath(
                artifactRoot,
                format,
                numberedDisplays: false);
            if (File.Exists(unnumberedPath))
            {
                Console.WriteLine(
                    $"[PHYSICAL UI PERF PREP REUSE] format={format}; fixture={unnumberedPath}");
            }
            else if (format == "mathtype")
            {
                CreateStagedMathTypeUnnumberedPerformanceFixture(
                    unnumberedPath,
                    genericSvg,
                    assets.PngPath,
                    assets.EmfPath,
                    wmfPath);
            }
            else
            {
                application = CreateWordApplication(visible: false);
                try { application.ScreenUpdating = false; } catch { }
                CreatePhysicalUiEditPerformanceFixture(
                    application,
                    format,
                    numberedDisplays: false,
                    unnumberedPath,
                    genericSvg,
                    assets.PngPath,
                    assets.EmfPath,
                    wmfPath);
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(application);
                application = null;
                ForceComCleanup();
            }

            // Keep each 150-formula fixture in a fresh Word process. Reusing the
            // same automation host for a second large OLE corpus can retain COM/
            // presentation resources from the first fixture and terminate Word
            // near the 100-object boundary, which is unrelated to edit latency.
            var numberedPath = PhysicalUiPerfFixturePath(
                artifactRoot,
                format,
                numberedDisplays: true);
            if (File.Exists(numberedPath))
            {
                Console.WriteLine(
                    $"[PHYSICAL UI PERF PREP REUSE] format={format}; fixture={numberedPath}");
            }
            else
            {
                application = CreateWordApplication(visible: false);
                try { application.ScreenUpdating = false; } catch { }
                CreatePhysicalUiEditPerformanceFixture(
                    application,
                    format,
                    numberedDisplays: true,
                    numberedPath,
                    genericSvg,
                    assets.PngPath,
                    assets.EmfPath,
                    wmfPath);
            }

            Console.WriteLine(
                $"[PHYSICAL UI PERF PREPARE PASS] format={format}; "
                + $"unnumbered={unnumberedPath} "
                + $"({PhysicalUiPerfInlineCount} inline + {PhysicalUiPerfDisplayCount} display); "
                + $"numbered={numberedPath} ({PhysicalUiPerfDisplayCount} numbered display formulas).");
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition(previousNumberPosition);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void CreatePhysicalUiEditPerformanceFixture(
        Word.Application application,
        string format,
        bool numberedDisplays,
        string path,
        string genericSvg,
        string pngPath,
        string emfPath,
        string wmfPath)
    {
        Word.Document? document = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormat(document, "continuous");
            WordEquationNumbering.SetDefaultCreateObjectMode(
                PhysicalUiPerfObjectMode(format));
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(numberedDisplays);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");

            var formulaCount = numberedDisplays
                ? PhysicalUiPerfDisplayCount
                : PhysicalUiPerfTotalCount;
            InsertPhysicalUiEditPerformanceFormulaRange(
                application,
                document,
                format,
                numberedDisplays,
                1,
                formulaCount,
                genericSvg,
                pngPath,
                emfPath,
                wmfPath);

            AssertPhysicalUiPerfInventory(
                document,
                format,
                numberedDisplays,
                $"fixture before save {format}");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertPhysicalUiPerfInventory(
                document,
                format,
                numberedDisplays,
                $"fixture after reopen {format}");
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
        }
    }

    private static void CreateStagedMathTypeUnnumberedPerformanceFixture(
        string path,
        string genericSvg,
        string pngPath,
        string emfPath,
        string wmfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            try { application.ScreenUpdating = false; } catch { }
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormat(document, "continuous");
            WordEquationNumbering.SetDefaultCreateObjectMode(
                FormulaOleContract.MathTypeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");

            InsertPhysicalUiEditPerformanceFormulaRange(
                application,
                document,
                "mathtype",
                numberedDisplays: false,
                firstIndex: 1,
                lastIndex: PhysicalUiPerfInlineCount,
                genericSvg,
                pngPath,
                emfPath,
                wmfPath);

            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            application = null;
            ForceComCleanup();
        }

        try
        {
            application = CreateWordApplication(visible: false);
            try { application.ScreenUpdating = false; } catch { }
            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormat(document, "continuous");
            WordEquationNumbering.SetDefaultCreateObjectMode(
                FormulaOleContract.MathTypeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");

            AssertEqual(
                PhysicalUiPerfInlineCount,
                document.InlineShapes.Count,
                "staged MathType fixture lost inline formulas before display append.");

            InsertPhysicalUiEditPerformanceFormulaRange(
                application,
                document,
                "mathtype",
                numberedDisplays: false,
                firstIndex: PhysicalUiPerfInlineCount + 1,
                lastIndex: PhysicalUiPerfTotalCount,
                genericSvg,
                pngPath,
                emfPath,
                wmfPath);

            AssertPhysicalUiPerfInventory(
                document,
                "mathtype",
                numberedDisplays: false,
                "staged MathType fixture before final save");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertPhysicalUiPerfInventory(
                document,
                "mathtype",
                numberedDisplays: false,
                "staged MathType fixture after final reopen");
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void InsertPhysicalUiEditPerformanceFormulaRange(
        Word.Application application,
        Word.Document document,
        string format,
        bool numberedDisplays,
        int firstIndex,
        int lastIndex,
        string genericSvg,
        string pngPath,
        string emfPath,
        string wmfPath)
    {
        var service = new WordFormulaService(application);
        for (var index = firstIndex; index <= lastIndex; index++)
        {
            document.Activate();
            SelectDocumentEnd(document);
            if (index > 1)
                application.Selection.TypeParagraph();

            var display = numberedDisplays || index > PhysicalUiPerfInlineCount;
            if (!display)
                application.Selection.TypeText($"Inline {index:000}: ");

            var session = CreatePhysicalUiPerfSession(
                index,
                format,
                display ? "block" : "inline",
                numberedDisplays && display,
                genericSvg,
                pngPath);

            switch (format)
            {
                case "visualtex":
                    service.InsertOle(session, pngPath, emfPath);
                    break;
                case "mathtype":
                    service.InsertMathTypeOle(
                        session,
                        session.ExportResult?.MathMl
                            ?? throw new InvalidDataException(
                                $"MathType fixture #{index} has no MathML."),
                        emfPath,
                        isolatedNativePreviewWmfPath: wmfPath,
                        isolatedNativePreviewWidthPt: 120,
                        isolatedNativePreviewHeightPt: 36,
                        isolatedNativePreviewWordPosition:
                            session.Numbered || session.DisplayMode == "inline"
                                ? -3
                                : 0,
                        isolatedNativePreviewAttempted: true);
                    break;
                case "omml":
                    service.InsertOmml(
                        session,
                        session.ExportResult?.MathMl
                            ?? throw new InvalidDataException(
                                $"OMML fixture #{index} has no MathML."));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(format), format, null);
            }

            if (!display)
                application.Selection.TypeText(" tail");

            if (index % 25 == 0 || index == lastIndex)
                Console.WriteLine(
                    $"[PHYSICAL UI PERF PREP {format}] "
                    + $"numberedDisplays={numberedDisplays} "
                    + $"{index}/{lastIndex}");
        }
    }

    private static OfficeSessionDocument CreatePhysicalUiPerfSession(
        int index,
        string format,
        string displayMode,
        bool numbered,
        string genericSvg,
        string pngPath)
    {
        var latex = $@"x_{{{index}}}={index}";
        var display = displayMode == "block" ? "block" : "inline";
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\""
            + display
            + "\"><msub><mi>x</mi><mn>"
            + index.ToString(CultureInfo.InvariantCulture)
            + "</mn></msub><mo>=</mo><mn>"
            + index.ToString(CultureInfo.InvariantCulture)
            + "</mn></math>";
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Host = "word",
            Mode = "create",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = $"Physical UI performance formula {index}",
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = latex,
                },
            },
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = PhysicalUiPerfObjectMode(format),
            Numbered = numbered,
            MathTypeNumberPosition = "right",
            FontSizePt = 12,
            ExportResult = new OfficeExportDocument
            {
                Svg = genericSvg,
                SvgBase64 = "data:image/svg+xml;base64,"
                    + Convert.ToBase64String(Encoding.UTF8.GetBytes(genericSvg)),
                PngBase64 = "data:image/png;base64,"
                    + Convert.ToBase64String(File.ReadAllBytes(pngPath)),
                MathMl = mathMl,
                Width = ExportWidth,
                Height = ExportHeight,
                Baseline = ExportBaseline,
            },
        };
    }

    private static string PhysicalUiPerfObjectMode(string format) => format switch
    {
        "visualtex" => FormulaOleContract.NativeOleMode,
        "mathtype" => FormulaOleContract.MathTypeOleMode,
        "omml" => FormulaOleContract.WordOmmlMode,
        _ => throw new ArgumentOutOfRangeException(nameof(format), format, null),
    };

    private static string PhysicalUiPerfFixturePath(
        string artifactRoot,
        string format,
        bool numberedDisplays) =>
        Path.Combine(
            artifactRoot,
            $"Physical-UI-Edit-{format}-"
            + (numberedDisplays ? "numbered" : "unnumbered")
            + ".docx");

    private static void AssertPhysicalUiPerfInventory(
        Word.Document document,
        string format,
        bool numberedDisplays,
        string context)
    {
        var expectedFormulaCount = numberedDisplays
            ? PhysicalUiPerfDisplayCount
            : PhysicalUiPerfTotalCount;
        if (format == "omml")
            AssertEqual(
                expectedFormulaCount,
                document.OMaths.Count,
                context + ": OMML formula count mismatch.");
        else
            AssertEqual(
                expectedFormulaCount,
                document.InlineShapes.Count,
                context + ": OLE formula count mismatch.");

        if (format == "mathtype")
            AssertEqual(
                numberedDisplays ? PhysicalUiPerfDisplayCount : 0,
                CountStressMathTypePlaceRefFields(document),
                context + ": MathType numbered-host count mismatch.");
        else if (format == "visualtex")
            AssertEqual(
                numberedDisplays ? PhysicalUiPerfDisplayCount : 0,
                CountVisualTeXPlaceRefFields(document),
                context + ": VisualTeX numbered-host count mismatch.");
        else
            AssertEqual(
                numberedDisplays ? PhysicalUiPerfDisplayCount : 0,
                CountPhysicalUiPerfNumberedOmmlHosts(document),
                context + ": OMML numbered-host count mismatch.");
    }

    private static int CountPhysicalUiPerfNumberedOmmlHosts(
        Word.Document document)
    {
        var count = 0;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Release(range);
                range = null;
                Release(math);
                math = maths[index];
                range = math.Range.Duplicate;
                var openXml = range.WordOpenXML ?? string.Empty;
                if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(openXml))
                    count++;
            }
            return count;
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
        }
    }

    private static void MeasurePhysicalUiEditPerformance(
        VisualTeXSessionClient client,
        string format,
        string scenario,
        string artifactRoot)
    {
        var numbered = scenario == "numbered";
        var fixturePath = PhysicalUiPerfFixturePath(
            artifactRoot,
            format,
            numberedDisplays: numbered);
        if (!File.Exists(fixturePath))
            throw new FileNotFoundException(
                "Physical UI edit performance fixture is missing. Run scenario=prepare first.",
                fixturePath);

        var start = scenario switch
        {
            "inline" => 1,
            "display" => PhysicalUiPerfInlineCount + 1,
            _ => 1,
        };
        var end = scenario switch
        {
            "inline" => PhysicalUiPerfInlineCount,
            "display" => PhysicalUiPerfTotalCount,
            _ => PhysicalUiPerfDisplayCount,
        };
        if (int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_UI_EDIT_PERF_START"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var configuredStart))
            start = Math.Max(start, configuredStart);
        if (int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_UI_EDIT_PERF_END"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var configuredEnd))
            end = Math.Min(end, configuredEnd);
        if (start > end)
            throw new InvalidOperationException(
                $"Physical UI performance sample range is empty: {start}>{end}.");

        var threshold = numbered
            ? PhysicalUiPerfNumberedLimitMs
            : PhysicalUiPerfUnnumberedLimitMs;
        var previousAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousTracePath =
            Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var tracePath = Path.Combine(
            artifactRoot,
            $"physical-ui-edit-{format}-{scenario}-vsto.trace.log");
        try { if (File.Exists(tracePath)) File.Delete(tracePath); } catch { }
        Word.Application? application = null;
        Word.Document? document = null;
        string? workingFixturePath = null;
        using var codeBaseOverride = AttachActiveWord
            ? null
            : CreateUiE2eWordVstoCodeBaseOverride();

        try
        {
            if (FindVisiblePhysicalUiPerfEditorWindow() != IntPtr.Zero)
                throw new InvalidOperationException(
                    "A VisualTeX Office editor for the performance target process is already visible. "
                    + "The physical UI performance run will not close or reuse an unrelated editor.");

            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                tracePath);
            workingFixturePath = Path.Combine(
                artifactRoot,
                $"Physical-UI-Edit-{format}-{scenario}-run-{Guid.NewGuid():N}.docx");
            File.Copy(fixturePath, workingFixturePath, overwrite: false);
            application = CreateWordApplication(visible: true);
            document = application.Documents.Open(
                workingFixturePath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);
            document.Activate();
            try
            {
                application.ActiveWindow.WindowState =
                    Word.WdWindowState.wdWindowStateMaximize;
            }
            catch { }

            AssertPhysicalUiPerfInventory(
                document,
                format,
                numberedDisplays: numbered,
                $"physical UI measurement preflight {format}/{scenario}");

            var wordWindow = new IntPtr(application.ActiveWindow.Hwnd);
            var wordRoot = AutomationElement.FromHandle(wordWindow);
            EnsurePhysicalUiPerfRibbonTab(wordWindow, wordRoot);

            var metrics = new List<PhysicalUiEditMetric>(end - start + 1);
            // Word startup asks the companion to create the reusable editor shell
            // while it is still hidden. Capture that exact HWND before any
            // stopwatch starts so sample #1 uses the same direct visibility probe
            // as every later sample instead of repeatedly discovering top-level
            // windows after the click.
            var reusableEditorWindow = WaitForPhysicalUiPerfEditorHandle(
                TimeSpan.FromSeconds(5));
            for (var index = start; index <= end; index++)
            {
                SelectPhysicalUiPerfFormula(document, format, index);
                document.Activate();
                SetForegroundWindow(wordWindow);
                WaitForPhysicalUiPerfSelectionStable(
                    document,
                    format,
                    index,
                    TimeSpan.FromSeconds(2));

                wordRoot = AutomationElement.FromHandle(wordWindow);
                var editButton = WaitForUiElement(
                    wordRoot,
                    "编辑所选公式",
                    ControlType.Button,
                    TimeSpan.FromSeconds(8),
                    requireEnabled: true);
                PreparePhysicalUiPerfClick(
                    editButton,
                    $"Word Ribbon 编辑所选公式 #{index}");

                var openWatch = StartPhysicalUiPerfClick(
                    editButton,
                    $"Word Ribbon 编辑所选公式 #{index}");
                var editorWindow = WaitForVisibleOfficeEditorWindowFast(
                    TimeSpan.FromSeconds(8),
                    reusableEditorWindow);
                reusableEditorWindow = editorWindow;
                var openVisibleMs = openWatch.Elapsed.TotalMilliseconds;

                var editorRoot = AutomationElement.FromHandle(editorWindow);
                var replacement = WaitForUiElement(
                    editorRoot,
                    "欧拉恒等式",
                    ControlType.Button,
                    TimeSpan.FromSeconds(8),
                    requireEnabled: true);
                var openReadyMs = openWatch.Elapsed.TotalMilliseconds;
                openWatch.Stop();

                SetForegroundWindow(editorWindow);
                ClickUiElementByMouse(
                    replacement,
                    $"欧拉恒等式 edit replacement #{index}");
                var update = WaitForUiElement(
                    editorRoot,
                    "更新公式",
                    ControlType.Button,
                    TimeSpan.FromSeconds(8),
                    requireEnabled: true);
                PreparePhysicalUiPerfClick(
                    update,
                    $"更新公式 #{index}");

                var traceOffsetBeforeApply = PhysicalUiPerfTraceLength(tracePath);
                var applyWatch = StartPhysicalUiPerfClick(
                    update,
                    $"更新公式 #{index}");
                WaitForSpecificOfficeEditorHiddenFast(
                    editorWindow,
                    TimeSpan.FromSeconds(8));
                applyWatch.Stop();

                var metric = new PhysicalUiEditMetric
                {
                    Index = index,
                    Format = format,
                    Scenario = scenario,
                    DisplayMode = scenario == "inline" ? "inline" : "block",
                    Numbered = numbered,
                    OpenVisibleMilliseconds = openVisibleMs,
                    OpenReadyMilliseconds = openReadyMs,
                    ApplyHiddenMilliseconds = applyWatch.Elapsed.TotalMilliseconds,
                };
                metrics.Add(metric);
                Console.WriteLine(
                    $"[PHYSICAL UI PERF {format}/{scenario} "
                    + $"{index - start + 1:D3}/{end - start + 1:D3}] "
                    + $"index={index} openVisible={metric.OpenVisibleMilliseconds:0.0}ms "
                    + $"openReady={metric.OpenReadyMilliseconds:0.0}ms "
                    + $"applyHidden={metric.ApplyHiddenMilliseconds:0.0}ms "
                    + $"limit={threshold:0}ms");

                // The editor hides before the VSTO write-back/finalization path
                // necessarily releases its operation gate. Wait for the existing
                // production trace marker emitted immediately after Release() so
                // sample N tail work can never inflate sample N+1 open latency.
                WaitForPhysicalUiPerfOperationRelease(
                    tracePath,
                    traceOffsetBeforeApply,
                    TimeSpan.FromSeconds(5));
                WaitForPhysicalUiPerfWordQuiescence(document, wordWindow);
            }

            AssertPhysicalUiPerfInventory(
                document,
                format,
                numberedDisplays: numbered,
                $"physical UI measurement after edits {format}/{scenario}");
            WritePhysicalUiEditPerformanceReport(
                metrics,
                threshold,
                artifactRoot,
                format,
                scenario,
                fixturePath);
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            if (!string.IsNullOrWhiteSpace(workingFixturePath))
            {
                try { File.Delete(workingFixturePath); } catch { }
            }
        }
    }


    private static void MeasurePhysicalUiCreatePerformance(
        VisualTeXSessionClient client,
        string format,
        string scenario,
        string artifactRoot)
    {
        _ = client;
        var numbered = scenario == "create-numbered";
        var inline = scenario == "create-inline";
        var sampleCount = inline
            ? PhysicalUiPerfInlineCount
            : PhysicalUiPerfDisplayCount;
        var start = 1;
        var end = sampleCount;
        if (int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_UI_EDIT_PERF_START"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var configuredStart))
            start = Math.Max(1, configuredStart);
        if (int.TryParse(
                Environment.GetEnvironmentVariable("VISUALTEX_UI_EDIT_PERF_END"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var configuredEnd))
            end = Math.Min(sampleCount, configuredEnd);
        if (start > end)
            throw new InvalidOperationException(
                $"Physical UI create performance sample range is empty: {start}>{end}.");
        var measuredCount = end - start + 1;
        var threshold = numbered
            ? PhysicalUiPerfNumberedLimitMs
            : PhysicalUiPerfUnnumberedLimitMs;
        var previousCreateMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        var previousNumberPosition = WordEquationNumbering.GetDefaultMathTypeNumberPosition();
        var previousAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousTracePath =
            Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var tracePath = Path.Combine(
            artifactRoot,
            $"physical-ui-create-{format}-{(inline ? "inline" : numbered ? "numbered" : "display")}-vsto.trace.log");
        try { if (File.Exists(tracePath)) File.Delete(tracePath); } catch { }

        Word.Application? application = null;
        Word.Document? document = null;
        using var codeBaseOverride = AttachActiveWord
            ? null
            : CreateUiE2eWordVstoCodeBaseOverride();

        try
        {
            if (FindVisiblePhysicalUiPerfEditorWindow() != IntPtr.Zero)
                throw new InvalidOperationException(
                    "A VisualTeX Office editor for the performance target process is already visible. "
                    + "The physical UI create performance run will not close or reuse an unrelated editor.");

            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                tracePath);

            if (format != "omml")
                WordEquationNumbering.SetDefaultCreateObjectMode(
                    PhysicalUiPerfObjectMode(format));
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(numbered);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition("right");

            application = CreateWordApplication(visible: true);
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormat(document, "continuous");
            try
            {
                application.ActiveWindow.WindowState =
                    Word.WdWindowState.wdWindowStateMaximize;
            }
            catch { }

            var wordWindow = new IntPtr(application.ActiveWindow.Hwnd);
            var wordRoot = AutomationElement.FromHandle(wordWindow);
            EnsurePhysicalUiPerfRibbonTab(wordWindow, wordRoot);
            var reusableEditorWindow = WaitForPhysicalUiPerfEditorHandle(
                TimeSpan.FromSeconds(5));

            var metrics = new List<PhysicalUiEditMetric>(measuredCount);
            for (var index = start; index <= end; index++)
            {
                document.Activate();
                SelectDocumentEnd(document);
                if (index > start)
                    application.Selection.TypeParagraph();
                if (inline)
                    application.Selection.TypeText($"Inline {index:000}: ");

                SetForegroundWindow(wordWindow);
                PumpUi(TimeSpan.FromMilliseconds(30));
                wordRoot = AutomationElement.FromHandle(wordWindow);
                var createButtonName = format == "omml"
                    ? inline ? "OMML 行内公式" : "OMML 行间公式"
                    : inline ? "OLE 行内公式" : "OLE 行间公式";
                var createButton = WaitForUiElement(
                    wordRoot,
                    createButtonName,
                    ControlType.Button,
                    TimeSpan.FromSeconds(8),
                    requireEnabled: true);
                PreparePhysicalUiPerfClick(
                    createButton,
                    $"Word Ribbon {createButtonName} #{index}");

                var openWatch = StartPhysicalUiPerfClick(
                    createButton,
                    $"Word Ribbon {createButtonName} #{index}");
                var editorWindow = WaitForVisibleOfficeEditorWindowFast(
                    TimeSpan.FromSeconds(8),
                    reusableEditorWindow);
                reusableEditorWindow = editorWindow;
                var openVisibleMs = openWatch.Elapsed.TotalMilliseconds;

                var editorRoot = AutomationElement.FromHandle(editorWindow);
                var replacement = WaitForUiElement(
                    editorRoot,
                    "欧拉恒等式",
                    ControlType.Button,
                    TimeSpan.FromSeconds(8),
                    requireEnabled: true);
                var openReadyMs = openWatch.Elapsed.TotalMilliseconds;
                openWatch.Stop();

                SetForegroundWindow(editorWindow);
                ClickUiElementByMouse(
                    replacement,
                    $"欧拉恒等式 create replacement #{index}");
                var finish = WaitForUiElement(
                    editorRoot,
                    "完成并插入",
                    ControlType.Button,
                    TimeSpan.FromSeconds(8),
                    requireEnabled: true);
                PreparePhysicalUiPerfClick(
                    finish,
                    $"完成并插入 #{index}");

                var traceOffsetBeforeApply = PhysicalUiPerfTraceLength(tracePath);
                var applyWatch = StartPhysicalUiPerfClick(
                    finish,
                    $"完成并插入 #{index}");
                WaitForSpecificOfficeEditorHiddenFast(
                    editorWindow,
                    TimeSpan.FromSeconds(8));
                applyWatch.Stop();

                var metric = new PhysicalUiEditMetric
                {
                    Index = index,
                    Format = format,
                    Scenario = scenario,
                    DisplayMode = inline ? "inline" : "block",
                    Numbered = numbered,
                    OpenVisibleMilliseconds = openVisibleMs,
                    OpenReadyMilliseconds = openReadyMs,
                    ApplyHiddenMilliseconds = applyWatch.Elapsed.TotalMilliseconds,
                };
                metrics.Add(metric);
                Console.WriteLine(
                    $"[PHYSICAL UI CREATE {format}/{scenario} "
                    + $"{index - start + 1:D3}/{measuredCount:D3}] "
                    + $"openVisible={metric.OpenVisibleMilliseconds:0.0}ms "
                    + $"openReady={metric.OpenReadyMilliseconds:0.0}ms "
                    + $"insertHidden={metric.ApplyHiddenMilliseconds:0.0}ms "
                    + $"limit={threshold:0}ms");

                WaitForPhysicalUiPerfOperationRelease(
                    tracePath,
                    traceOffsetBeforeApply,
                    TimeSpan.FromSeconds(8));
                WaitForPhysicalUiPerfWordQuiescence(document, wordWindow);
            }

            var actualFormulaCount = format == "omml"
                ? document.OMaths.Count
                : document.InlineShapes.Count;
            AssertEqual(
                measuredCount,
                actualFormulaCount,
                $"physical UI create {format}/{scenario}: formula count mismatch.");

            var actualNumberedCount = format switch
            {
                "visualtex" => CountVisualTeXPlaceRefFields(document),
                "mathtype" => CountStressMathTypePlaceRefFields(document),
                "omml" => CountPhysicalUiPerfNumberedOmmlHosts(document),
                _ => 0,
            };
            AssertEqual(
                numbered ? measuredCount : 0,
                actualNumberedCount,
                $"physical UI create {format}/{scenario}: numbered host count mismatch.");

            WritePhysicalUiCreatePerformanceReport(
                metrics,
                threshold,
                artifactRoot,
                format,
                scenario);
        }
        finally
        {
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousNumbered);
            WordEquationNumbering.SetDefaultMathTypeNumberPosition(previousNumberPosition);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void WritePhysicalUiCreatePerformanceReport(
        IReadOnlyList<PhysicalUiEditMetric> metrics,
        double threshold,
        string artifactRoot,
        string format,
        string scenario)
    {
        var suffix = scenario.Replace("create-", string.Empty);
        var prefix = $"physical-ui-create-{format}-{suffix}";
        var csvPath = Path.Combine(artifactRoot, prefix + ".csv");
        var jsonPath = Path.Combine(artifactRoot, prefix + ".json");
        var csv = new List<string>
        {
            "index,format,scenario,display_mode,numbered,open_visible_ms,open_ready_ms,insert_hidden_ms,limit_ms",
        };
        csv.AddRange(metrics.Select(item => string.Join(
            ",",
            item.Index.ToString(CultureInfo.InvariantCulture),
            item.Format,
            item.Scenario,
            item.DisplayMode,
            item.Numbered ? "true" : "false",
            item.OpenVisibleMilliseconds.ToString("0.000", CultureInfo.InvariantCulture),
            item.OpenReadyMilliseconds.ToString("0.000", CultureInfo.InvariantCulture),
            item.ApplyHiddenMilliseconds.ToString("0.000", CultureInfo.InvariantCulture),
            threshold.ToString("0.000", CultureInfo.InvariantCulture))));
        File.WriteAllLines(csvPath, csv, new UTF8Encoding(false));

        var openValues = metrics
            .Select(item => item.OpenVisibleMilliseconds)
            .OrderBy(value => value)
            .ToArray();
        var readyValues = metrics
            .Select(item => item.OpenReadyMilliseconds)
            .OrderBy(value => value)
            .ToArray();
        var insertValues = metrics
            .Select(item => item.ApplyHiddenMilliseconds)
            .OrderBy(value => value)
            .ToArray();
        var violations = metrics
            .Where(item =>
                item.OpenVisibleMilliseconds > threshold
                || item.ApplyHiddenMilliseconds > threshold)
            .OrderByDescending(item =>
                Math.Max(
                    item.OpenVisibleMilliseconds,
                    item.ApplyHiddenMilliseconds))
            .ToArray();

        var report = new
        {
            GeneratedAt = DateTimeOffset.Now,
            Format = format,
            Scenario = scenario,
            SampleCount = metrics.Count,
            ThresholdMilliseconds = threshold,
            TimingDefinition = new
            {
                Open =
                    "physical left-button click on the Word Ribbon create-formula button -> Office editor top-level window becomes visible",
                OpenReady =
                    "same physical click -> enabled 欧拉恒等式 editor control is discoverable; diagnostic only",
                Insert =
                    "physical left-button click on 完成并插入 -> the exact Office editor window is no longer visible",
            },
            OpenVisible = PhysicalUiPerfStats(openValues),
            OpenReady = PhysicalUiPerfStats(readyValues),
            InsertHidden = PhysicalUiPerfStats(insertValues),
            ViolationCount = violations.Length,
            Violations = violations.Take(30).ToArray(),
            Samples = metrics,
        };
        File.WriteAllText(
            jsonPath,
            JsonSerializer.Serialize(
                report,
                new JsonSerializerOptions { WriteIndented = true }),
            new UTF8Encoding(false));

        Console.WriteLine(
            $"[PHYSICAL UI CREATE SUMMARY] format={format} scenario={scenario} "
            + $"samples={metrics.Count} limit={threshold:0}ms; "
            + $"open p50={PhysicalUiPerfPercentile(openValues, 0.50):0.0}ms "
            + $"p95={PhysicalUiPerfPercentile(openValues, 0.95):0.0}ms "
            + $"max={(openValues.Length == 0 ? 0 : openValues[openValues.Length - 1]):0.0}ms; "
            + $"insert p50={PhysicalUiPerfPercentile(insertValues, 0.50):0.0}ms "
            + $"p95={PhysicalUiPerfPercentile(insertValues, 0.95):0.0}ms "
            + $"max={(insertValues.Length == 0 ? 0 : insertValues[insertValues.Length - 1]):0.0}ms; "
            + $"violations={violations.Length}; csv={csvPath}; json={jsonPath}");

        if (violations.Length > 0)
        {
            var worst = string.Join(
                Environment.NewLine,
                violations.Take(15).Select(item =>
                    $"#{item.Index} open={item.OpenVisibleMilliseconds:0.0}ms "
                    + $"insert={item.ApplyHiddenMilliseconds:0.0}ms"));
            throw new InvalidDataException(
                $"Physical UI create performance missed the strict {threshold:0}ms per-operation gate "
                + $"for {format}/{scenario}: {violations.Length}/{metrics.Count} samples failed."
                + Environment.NewLine
                + worst
                + Environment.NewLine
                + $"Full report: {jsonPath}");
        }
    }

    private static void EnsurePhysicalUiPerfRibbonTab(
        IntPtr wordWindow,
        AutomationElement wordRoot)
    {
        SetForegroundWindow(wordWindow);
        var visualTeXTab = WaitForUiElement(
            wordRoot,
            "VisualTeX",
            ControlType.TabItem,
            TimeSpan.FromSeconds(12),
            requireEnabled: true);
        if (!visualTeXTab.TryGetCurrentPattern(
                SelectionItemPattern.Pattern,
                out var pattern)
            || pattern is not SelectionItemPattern selection
            || !selection.Current.IsSelected)
            ClickUiElementByMouse(
                visualTeXTab,
                "Word VisualTeX Ribbon tab for physical UI performance");
        PumpUi(TimeSpan.FromMilliseconds(200));
    }

    private static void SelectPhysicalUiPerfFormula(
        Word.Document document,
        string format,
        int index)
    {
        Word.InlineShape? shape = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            if (format == "omml")
            {
                math = document.OMaths[index];
                range = math.Range.Duplicate;
            }
            else
            {
                shape = document.InlineShapes[index];
                range = shape.Range.Duplicate;
            }
            document.Activate();
            document.Application.ActiveWindow.ScrollIntoView(range, true);
            range.Select();
        }
        finally
        {
            Release(range);
            Release(math);
            Release(shape);
        }
    }

    private static void WaitForPhysicalUiPerfSelectionStable(
        Word.Document document,
        string format,
        int index,
        TimeSpan timeout)
    {
        var deadline = Stopwatch.StartNew();
        var stablePasses = 0;
        var lastStart = int.MinValue;
        var lastEnd = int.MinValue;
        while (deadline.Elapsed < timeout)
        {
            Word.InlineShape? shape = null;
            Word.OMath? math = null;
            Word.Range? targetRange = null;
            Word.Selection? selection = null;
            Word.Range? selectionRange = null;
            try
            {
                if (format == "omml")
                {
                    math = document.OMaths[index];
                    targetRange = math.Range.Duplicate;
                }
                else
                {
                    shape = document.InlineShapes[index];
                    targetRange = shape.Range.Duplicate;
                }

                selection = document.Application.Selection;
                selectionRange = selection.Range.Duplicate;
                var start = selectionRange.Start;
                var end = selectionRange.End;
                var containsTarget =
                    start <= targetRange.Start
                    && end >= targetRange.End;
                if (containsTarget
                    && start == lastStart
                    && end == lastEnd)
                {
                    stablePasses++;
                    if (stablePasses >= 2) return;
                }
                else
                {
                    stablePasses = containsTarget ? 1 : 0;
                    lastStart = start;
                    lastEnd = end;
                }
            }
            catch (COMException)
            {
                stablePasses = 0;
            }
            finally
            {
                Release(selectionRange);
                Release(selection);
                Release(targetRange);
                Release(math);
                Release(shape);
            }
            PumpUi(TimeSpan.FromMilliseconds(20));
        }
        throw new TimeoutException(
            $"Word selection did not stabilize on physical UI performance formula #{index}.");
    }

    private static void WaitForPhysicalUiPerfWordQuiescence(
        Word.Document document,
        IntPtr wordWindow)
    {
        // The editor is hidden as soon as the user's Apply action is accepted,
        // while the VSTO dispatcher can still be finishing the Word write. This
        // settling wait is deliberately outside both measured intervals so the
        // next sample never inherits work from the preceding formula.
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < TimeSpan.FromSeconds(3))
        {
            try
            {
                document.Activate();
                if (document.Windows.Count > 0)
                {
                    SetForegroundWindow(wordWindow);
                    PumpUi(TimeSpan.FromMilliseconds(600));
                    return;
                }
            }
            catch (COMException)
            {
                // Word can transiently report no active document window while a
                // structural display-formula write is completing. Retry outside
                // the performance stopwatch rather than polluting the next sample.
            }
            PumpUi(TimeSpan.FromMilliseconds(40));
        }
        throw new TimeoutException(
            "Word did not become quiescent after the previous physical UI formula edit.");
    }

    private static long PhysicalUiPerfTraceLength(string tracePath)
    {
        try
        {
            return File.Exists(tracePath)
                ? new FileInfo(tracePath).Length
                : 0L;
        }
        catch
        {
            return 0L;
        }
    }

    private static void WaitForPhysicalUiPerfOperationRelease(
        string tracePath,
        long startOffset,
        TimeSpan timeout)
    {
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < timeout)
        {
            System.Windows.Forms.Application.DoEvents();
            try
            {
                if (File.Exists(tracePath))
                {
                    using var stream = new FileStream(
                        tracePath,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite | FileShare.Delete);
                    var offset = Math.Max(0L, Math.Min(startOffset, stream.Length));
                    stream.Position = offset;
                    using var reader = new StreamReader(
                        stream,
                        Encoding.UTF8,
                        detectEncodingFromByteOrderMarks: true,
                        bufferSize: 4096,
                        leaveOpen: false);
                    var tail = reader.ReadToEnd();
                    if (tail.IndexOf(
                            "ribbon-session-operation-released",
                            StringComparison.Ordinal) >= 0)
                        return;
                }
            }
            catch (IOException)
            {
                // The VSTO trace writer may be flushing the same file. Retry.
            }
            Thread.Sleep(5);
        }

        throw new TimeoutException(
            "The VisualTeX Word add-in did not release its formula-operation gate after the editor became hidden.");
    }

    private static void PreparePhysicalUiPerfClick(
        AutomationElement element,
        string description)
    {
        var bounds = element.Current.BoundingRectangle;
        if (bounds.Width <= 1 || bounds.Height <= 1)
            throw new InvalidOperationException(
                $"UI element '{description}' has invalid bounds {bounds}.");
        try { element.SetFocus(); } catch { }
        var x = (int)Math.Round(bounds.Left + bounds.Width / 2d);
        var y = (int)Math.Round(bounds.Top + bounds.Height / 2d);
        if (!SetCursorPos(x, y))
            throw new InvalidOperationException(
                $"Unable to position the mouse on '{description}'.");
        // Cursor positioning/focus settling is intentionally outside the stopwatch.
        Thread.Sleep(25);
    }

    private static Stopwatch StartPhysicalUiPerfClick(
        AutomationElement element,
        string description)
    {
        // Word can finish a ScrollIntoView/Ribbon layout pass after the earlier
        // focus-settling step and move the pointer target. Re-read the live UIA
        // bounds and position the cursor immediately before LEFTDOWN. Cursor
        // placement is outside the measured interval; the stopwatch begins
        // directly before the physical mouse press, matching the user-visible
        // timing contract.
        var bounds = element.Current.BoundingRectangle;
        if (bounds.Width <= 1 || bounds.Height <= 1)
            throw new InvalidOperationException(
                $"UI element '{description}' has invalid final bounds {bounds}.");
        var x = (int)Math.Round(bounds.Left + bounds.Width / 2d);
        var y = (int)Math.Round(bounds.Top + bounds.Height / 2d);
        if (!SetCursorPos(x, y))
            throw new InvalidOperationException(
                $"Unable to position the mouse on '{description}' immediately before click.");
        var watch = Stopwatch.StartNew();
        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
        return watch;
    }

    private static string ResolvePhysicalUiPerfVisualTeXProcessPath()
    {
        var configured = (
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_UI_EDIT_PERF_VISUALTEX_PROCESS_PATH")
                ?? string.Empty)
            .Trim()
            .Trim('"');
        if (!string.IsNullOrWhiteSpace(configured))
        {
            try { return Path.GetFullPath(configured); }
            catch { return configured; }
        }

        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\VisualTeX\OfficeIntegration");
            var registered = (key?.GetValue("ExecutablePath") as string ?? string.Empty)
                .Trim()
                .Trim('"');
            if (!string.IsNullOrWhiteSpace(registered))
            {
                try { return Path.GetFullPath(registered); }
                catch { return registered; }
            }
        }
        catch { }
        return string.Empty;
    }

    private static IntPtr FindPhysicalUiPerfEditorWindow(bool visibleOnly)
    {
        var requiredPath = ResolvePhysicalUiPerfVisualTeXProcessPath();
        if (string.IsNullOrWhiteSpace(requiredPath) && visibleOnly)
            return FindVisibleOfficeEditorWindow();

        var result = IntPtr.Zero;
        EnumWindows((windowHandle, _) =>
        {
            if (visibleOnly && !IsWindowVisible(windowHandle)) return true;
            GetWindowThreadProcessId(windowHandle, out var processId);
            if (processId == 0) return true;
            try
            {
                using var process = Process.GetProcessById((int)processId);
                if (!string.Equals(
                        process.ProcessName,
                        "visualtex",
                        StringComparison.OrdinalIgnoreCase))
                    return true;
                var processPath = process.MainModule?.FileName ?? string.Empty;
                try { processPath = Path.GetFullPath(processPath); }
                catch { }
                if (!string.Equals(
                        processPath,
                        requiredPath,
                        StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch
            {
                return true;
            }

            var length = GetWindowTextLength(windowHandle);
            if (length <= 0) return true;
            var title = new StringBuilder(length + 1);
            GetWindowText(windowHandle, title, title.Capacity);
            if (title.ToString().IndexOf(
                    "Office",
                    StringComparison.OrdinalIgnoreCase) < 0)
                return true;
            result = windowHandle;
            return false;
        }, IntPtr.Zero);
        return result;
    }

    private static IntPtr FindVisiblePhysicalUiPerfEditorWindow() =>
        FindPhysicalUiPerfEditorWindow(visibleOnly: true);

    private static IntPtr WaitForPhysicalUiPerfEditorHandle(TimeSpan timeout)
    {
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < timeout)
        {
            System.Windows.Forms.Application.DoEvents();
            var handle = FindPhysicalUiPerfEditorWindow(visibleOnly: false);
            if (handle != IntPtr.Zero) return handle;
            Thread.Sleep(10);
        }
        throw new TimeoutException(
            "The prewarmed VisualTeX Office editor HWND was not created before the physical UI performance measurement.");
    }

    private static IntPtr WaitForVisibleOfficeEditorWindowFast(
        TimeSpan timeout,
        IntPtr reusableEditorWindow)
    {
        var deadline = Stopwatch.StartNew();
        if (reusableEditorWindow != IntPtr.Zero)
        {
            while (deadline.Elapsed < timeout)
            {
                System.Windows.Forms.Application.DoEvents();
                if (IsWindowVisible(reusableEditorWindow))
                    return reusableEditorWindow;
                _ = GetWindowThreadProcessId(reusableEditorWindow, out var processId);
                if (processId == 0)
                    break;
                Thread.Sleep(1);
            }
        }

        while (deadline.Elapsed < timeout)
        {
            System.Windows.Forms.Application.DoEvents();
            var handle = FindVisiblePhysicalUiPerfEditorWindow();
            if (handle != IntPtr.Zero) return handle;
            Thread.Sleep(5);
        }
        throw new TimeoutException(
            "The VisualTeX Office editor did not become visible during the physical UI performance measurement.");
    }

    private static void WaitForSpecificOfficeEditorHiddenFast(
        IntPtr editorWindow,
        TimeSpan timeout)
    {
        var deadline = Stopwatch.StartNew();
        while (deadline.Elapsed < timeout)
        {
            System.Windows.Forms.Application.DoEvents();
            if (!IsWindowVisible(editorWindow)) return;
            Thread.Sleep(5);
        }
        throw new TimeoutException(
            "The VisualTeX Office editor did not become hidden during the physical UI performance measurement.");
    }

    private static void WritePhysicalUiEditPerformanceReport(
        IReadOnlyList<PhysicalUiEditMetric> metrics,
        double threshold,
        string artifactRoot,
        string format,
        string scenario,
        string fixturePath)
    {
        var prefix = $"physical-ui-edit-{format}-{scenario}";
        var csvPath = Path.Combine(artifactRoot, prefix + ".csv");
        var jsonPath = Path.Combine(artifactRoot, prefix + ".json");
        var csv = new List<string>
        {
            "index,format,scenario,display_mode,numbered,open_visible_ms,open_ready_ms,apply_hidden_ms,limit_ms",
        };
        csv.AddRange(metrics.Select(item => string.Join(
            ",",
            item.Index.ToString(CultureInfo.InvariantCulture),
            item.Format,
            item.Scenario,
            item.DisplayMode,
            item.Numbered ? "true" : "false",
            item.OpenVisibleMilliseconds.ToString("0.000", CultureInfo.InvariantCulture),
            item.OpenReadyMilliseconds.ToString("0.000", CultureInfo.InvariantCulture),
            item.ApplyHiddenMilliseconds.ToString("0.000", CultureInfo.InvariantCulture),
            threshold.ToString("0.000", CultureInfo.InvariantCulture))));
        File.WriteAllLines(csvPath, csv, new UTF8Encoding(false));

        var openValues = metrics
            .Select(item => item.OpenVisibleMilliseconds)
            .OrderBy(value => value)
            .ToArray();
        var readyValues = metrics
            .Select(item => item.OpenReadyMilliseconds)
            .OrderBy(value => value)
            .ToArray();
        var applyValues = metrics
            .Select(item => item.ApplyHiddenMilliseconds)
            .OrderBy(value => value)
            .ToArray();
        var violations = metrics
            .Where(item =>
                item.OpenVisibleMilliseconds > threshold
                || item.ApplyHiddenMilliseconds > threshold)
            .OrderByDescending(item =>
                Math.Max(
                    item.OpenVisibleMilliseconds,
                    item.ApplyHiddenMilliseconds))
            .ToArray();

        var report = new
        {
            GeneratedAt = DateTimeOffset.Now,
            Format = format,
            Scenario = scenario,
            Fixture = fixturePath,
            SampleCount = metrics.Count,
            ThresholdMilliseconds = threshold,
            TimingDefinition = new
            {
                Open =
                    "physical left-button click on Word Ribbon 编辑所选公式 -> Office editor top-level window becomes visible",
                OpenReady =
                    "same physical click -> enabled 欧拉恒等式 editor control is discoverable; diagnostic only",
                Apply =
                    "physical left-button click on 更新公式 -> the exact Office editor window is no longer visible",
            },
            OpenVisible = PhysicalUiPerfStats(openValues),
            OpenReady = PhysicalUiPerfStats(readyValues),
            ApplyHidden = PhysicalUiPerfStats(applyValues),
            ViolationCount = violations.Length,
            Violations = violations.Take(30).ToArray(),
            Samples = metrics,
        };
        File.WriteAllText(
            jsonPath,
            JsonSerializer.Serialize(
                report,
                new JsonSerializerOptions { WriteIndented = true }),
            new UTF8Encoding(false));

        Console.WriteLine(
            $"[PHYSICAL UI PERF SUMMARY] format={format} scenario={scenario} "
            + $"samples={metrics.Count} limit={threshold:0}ms; "
            + $"open p50={PhysicalUiPerfPercentile(openValues, 0.50):0.0}ms "
            + $"p95={PhysicalUiPerfPercentile(openValues, 0.95):0.0}ms "
            + $"max={(openValues.Length == 0 ? 0 : openValues[openValues.Length - 1]):0.0}ms; "
            + $"apply p50={PhysicalUiPerfPercentile(applyValues, 0.50):0.0}ms "
            + $"p95={PhysicalUiPerfPercentile(applyValues, 0.95):0.0}ms "
            + $"max={(applyValues.Length == 0 ? 0 : applyValues[applyValues.Length - 1]):0.0}ms; "
            + $"violations={violations.Length}; csv={csvPath}; json={jsonPath}");

        if (violations.Length > 0)
        {
            var worst = string.Join(
                Environment.NewLine,
                violations.Take(15).Select(item =>
                    $"#{item.Index} open={item.OpenVisibleMilliseconds:0.0}ms "
                    + $"apply={item.ApplyHiddenMilliseconds:0.0}ms"));
            throw new InvalidDataException(
                $"Physical UI edit performance missed the strict {threshold:0}ms per-operation gate "
                + $"for {format}/{scenario}: {violations.Length}/{metrics.Count} samples failed."
                + Environment.NewLine
                + worst
                + Environment.NewLine
                + $"Full report: {jsonPath}");
        }
    }

    private static object PhysicalUiPerfStats(IReadOnlyList<double> sorted) =>
        new
        {
            P50Milliseconds = PhysicalUiPerfPercentile(sorted, 0.50),
            P95Milliseconds = PhysicalUiPerfPercentile(sorted, 0.95),
            MaximumMilliseconds = sorted.Count == 0 ? 0 : sorted[sorted.Count - 1],
            AverageMilliseconds = sorted.Count == 0 ? 0 : sorted.Average(),
        };

    private static double PhysicalUiPerfPercentile(
        IReadOnlyList<double> sorted,
        double percentile)
    {
        if (sorted.Count == 0) return 0;
        var bounded = Math.Max(0, Math.Min(1, percentile));
        var position = bounded * (sorted.Count - 1);
        var lower = (int)Math.Floor(position);
        var upper = (int)Math.Ceiling(position);
        if (lower == upper) return sorted[lower];
        var weight = position - lower;
        return sorted[lower] * (1d - weight) + sorted[upper] * weight;
    }
}

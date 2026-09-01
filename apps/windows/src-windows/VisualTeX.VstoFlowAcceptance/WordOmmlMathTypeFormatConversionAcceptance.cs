using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledOmmlMathTypeFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixturePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_OMML_MATHTYPE_INSTALLED_FIXTURE");
        if (string.IsNullOrWhiteSpace(fixturePath) || !File.Exists(fixturePath))
            throw new InvalidOperationException(
                "Installed OMML↔MathType acceptance requires VISUALTEX_OMML_MATHTYPE_INSTALLED_FIXTURE pointing to the saved OMML core fixture.");

        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var tracePath = Path.Combine(
            artifactRoot,
            "installed-omml-mathtype-format-conversion.trace.log");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                "1");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                tracePath);

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException(
                    "Installed OMML↔MathType acceptance requires MathType.exe process count to be zero before Word starts.");

            var writablePath = Path.Combine(
                artifactRoot,
                "installed-omml-mathtype-input.docx");
            File.Copy(Path.GetFullPath(fixturePath), writablePath, overwrite: true);

            application = CreateWordApplication(visible: true);
            document = application.Documents.Open(
                writablePath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(7, document.OMaths.Count,
                "Installed OMML↔MathType fixture must start with the seven saved OMML equations from core acceptance.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed OMML↔MathType fixture unexpectedly contains a MathType object before the installed add-in test.");

            // Add one equation through Word itself after the fixture opens. This
            // deliberately has no VisualTeX custom XML/bookmark metadata and proves
            // the installed Ribbon path also handles arbitrary native OMath content.
            const string nativeToken = "VT_INSTALLED_PURE_NATIVE_OMML";
            AppendAcceptanceText(
                document,
                $" installed-before-native {nativeToken} installed-after-native\r");
            Word.Range? nativeRange = null;
            try
            {
                nativeRange = InsertPureNativeOmml(document, nativeToken, "r+2");
            }
            finally { Release(nativeRange); }
            AssertEqual(8, document.OMaths.Count,
                "Installed acceptance did not create the additional pure Word-native OMath source.");

            addIns = application.COMAddIns;
            object addInKey = "VisualTeX.WordVsto";
            installedAddIn = addIns.Item(ref addInKey);
            if (!installedAddIn.Connect)
                installedAddIn.Connect = true;
            for (var index = 0; index < 80 && installedAddIn.Object is null; index++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
            callbacksObject = installedAddIn.Object
                ?? throw new InvalidOperationException(
                    "Installed VisualTeX.WordVsto automation object was unavailable. The acceptance refuses to fall back to a locally constructed ThisAddIn.");
            dynamic callbacks = callbacksObject;

            SelectFirstOmmlEquations(document, 2);
            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertOmmlToMathTypeSelection(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline);
            AssertEqual(6, document.OMaths.Count,
                "Installed OMML→MathType selection callback changed the wrong OMML count.");
            AssertEqual(2, CountMathTypeOleShapes(document),
                "Installed OMML→MathType selection callback did not create two MathType objects.");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertOmmlToMathTypeDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline);
            AssertEqual(0, document.OMaths.Count,
                "Installed OMML→MathType document callback left OMML sources behind in the mixed document.");
            AssertEqual(8, CountMathTypeOleShapes(document),
                "Installed OMML→MathType document callback did not preserve the two already-converted MathType objects while converting the remaining six OMML equations.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Installed OMML→MathType conversion did not preserve the single numbered equation as MTPlaceRef.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after installed OMML→MathType callbacks");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after installed OMML→MathType callbacks");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "installed OMML→MathType callbacks");

            var mathTypePath = Path.Combine(
                artifactRoot,
                "Installed-OMML-To-MathType-Acceptance.docx");
            document.SaveAs2(mathTypePath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                mathTypePath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(8, CountMathTypeOleShapes(document),
                "Installed OMML→MathType save/reopen lost a MathType object.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Installed OMML→MathType save/reopen changed the numbered equation.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after installed OMML→MathType save/reopen");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after installed OMML→MathType save/reopen");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "installed OMML→MathType save/reopen");

            SelectFirstMathTypeEquations(document, 2);
            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertMathTypeToOmmlSelection(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=MathType target=OMML",
                mathTypeBaseline);
            AssertEqual(2, document.OMaths.Count,
                "Installed MathType→OMML selection callback did not create two OMath equations.");
            AssertEqual(6, CountMathTypeOleShapes(document),
                "Installed MathType→OMML selection callback changed the wrong MathType count.");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertMathTypeToOmmlDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=MathType target=OMML",
                mathTypeBaseline);
            AssertEqual(8, document.OMaths.Count,
                "Installed MathType→OMML document callback did not restore all eight equations in the mixed document.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed MathType→OMML document callback left MathType sources behind.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "Installed MathType→OMML callbacks did not restore the single numbered OMML equation.");
            AssertOmmlSemanticCoverage(document);
            AssertInstalledOmmlMathTypeProseSurvived(document);
            AssertNoNewMathTypeProcess(mathTypeBaseline, "installed MathType→OMML callbacks");

            var finalPath = Path.Combine(
                artifactRoot,
                "Installed-MathType-To-OMML-Acceptance.docx");
            document.SaveAs2(finalPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(8, document.OMaths.Count,
                "Installed MathType→OMML save/reopen lost a native equation.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed MathType→OMML save/reopen restored a MathType object.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "Installed MathType→OMML save/reopen changed numbering state.");
            AssertOmmlSemanticCoverage(document);
            AssertInstalledOmmlMathTypeProseSurvived(document);
            AssertNoNewMathTypeProcess(mathTypeBaseline, "full installed OMML↔MathType acceptance");

            Console.WriteLine(
                "[OMML↔MATHTYPE INSTALLED] Installed VisualTeX.WordVsto COM automation object executed all four new Ribbon callbacks. Selection/document conversion in both directions passed with a mixed document, one runtime-created pure Word OMath, numbering, complex semantics, Equation.DSMT4/Equation Native integrity, live metafile ink, no OlePres, save/reopen persistence, and MathTypeProcessCount=0.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void ResetInstalledFormatConversionTrace(string tracePath)
    {
        try { File.Delete(tracePath); } catch { }
    }

    private static int WaitForInstalledOmmlMathTypeConversion(
        string tracePath,
        string directionMarker,
        IReadOnlyCollection<int> mathTypeBaseline,
        bool allowTransientMathTypeProcess = false)
    {
        var deadline = DateTime.UtcNow.AddSeconds(120);
        var completionMarker = "format-conversion-complete " + directionMarker;
        var stoppedMarker = "format-conversion-stopped " + directionMarker;
        var failedMarker = "format-conversion-failed";
        var peakAdditionalMathTypeProcesses = 0;
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(200);
            var startedMathType = SnapshotMathTypeProcessIds()
                .Except(mathTypeBaseline)
                .ToArray();
            peakAdditionalMathTypeProcesses = Math.Max(
                peakAdditionalMathTypeProcesses,
                startedMathType.Length);
            var controlledRpcHelpers = startedMathType
                .Where(MathTypeNativePreviewRenderer.IsControlledMathTypeRpcHelperProcess)
                .ToArray();
            if (controlledRpcHelpers.Length > 1)
                throw new InvalidOperationException(
                    "Installed OMML↔MathType conversion started more than one controlled -mtrpc helper: "
                    + string.Join(", ", controlledRpcHelpers));
            var unexpectedMathType = startedMathType
                .Except(controlledRpcHelpers)
                .ToArray();
            if (!allowTransientMathTypeProcess && unexpectedMathType.Length > 0)
                throw new InvalidOperationException(
                    "Installed OMML↔MathType conversion started unexpected MathType.exe: "
                    + string.Join(", ", unexpectedMathType));
            if (!File.Exists(tracePath)) continue;
            string trace;
            try { trace = File.ReadAllText(tracePath); }
            catch { continue; }
            var stoppedIndex = trace.IndexOf(stoppedMarker, StringComparison.Ordinal);
            if (stoppedIndex >= 0)
                throw new InvalidOperationException(
                    "Installed Ribbon conversion reported a stopped transaction: "
                    + trace.Substring(stoppedIndex).Trim());
            var failedIndex = trace.IndexOf(failedMarker, StringComparison.Ordinal);
            if (failedIndex >= 0)
                throw new InvalidOperationException(
                    "Installed Ribbon conversion reported a failed transaction: "
                    + trace.Substring(failedIndex).Trim());
            if (trace.IndexOf(completionMarker, StringComparison.Ordinal) < 0)
                continue;
            for (var settle = 0; settle < 50; settle++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
                var remainingMathType = SnapshotMathTypeProcessIds()
                    .Except(mathTypeBaseline)
                    .ToArray();
                peakAdditionalMathTypeProcesses = Math.Max(
                    peakAdditionalMathTypeProcesses,
                    remainingMathType.Length);
                if (remainingMathType.Length == 0)
                    return peakAdditionalMathTypeProcesses;
                var remainingControlledRpcHelpers = remainingMathType
                    .Where(MathTypeNativePreviewRenderer.IsControlledMathTypeRpcHelperProcess)
                    .ToArray();
                if (remainingControlledRpcHelpers.Length > 1)
                    throw new InvalidOperationException(
                        "Installed OMML↔MathType conversion retained more than one controlled -mtrpc helper: "
                        + string.Join(", ", remainingControlledRpcHelpers));
                var remainingUnexpected = remainingMathType
                    .Except(remainingControlledRpcHelpers)
                    .ToArray();
                if (remainingUnexpected.Length == 0)
                    return peakAdditionalMathTypeProcesses;
                if (!allowTransientMathTypeProcess)
                    throw new InvalidOperationException(
                        "Installed OMML↔MathType conversion left unexpected MathType.exe running: "
                        + string.Join(", ", remainingUnexpected));
            }
            throw new InvalidOperationException(
                "Installed Ribbon conversion completed but its transient MathType native-preview helper did not exit.");
        }
        throw new TimeoutException(
            $"Installed Ribbon conversion did not report '{completionMarker}' within the acceptance deadline.");
    }

    private static void AssertInstalledOmmlMathTypeProseSurvived(Word.Document document)
    {
        var requiredInOrder = new[]
        {
            "before-inline-hbar-greek",
            "after-inline-hbar-greek",
            "before-display-fraction-integral",
            "after-display-fraction-integral",
            "before-display-matrix",
            "after-display-matrix",
            "before-inline-accents-vector",
            "after-inline-accents-vector",
            "before-pure-native-omml",
            "after-pure-native-omml",
            "installed-before-native",
            "installed-after-native",
        };
        var text = document.Content.Text ?? string.Empty;
        var previousIndex = -1;
        foreach (var marker in requiredInOrder)
        {
            var index = text.IndexOf(marker, StringComparison.Ordinal);
            AssertTrue(index > previousIndex,
                $"Installed OMML↔MathType conversion removed or reordered adjacent prose marker '{marker}'.");
            previousIndex = index;
        }
    }

    private static void RunWordMathTypeOmmlSelectionViewAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixture = ResolveMathTypeNativeEditorFixture();
        if (!File.Exists(fixture))
            throw new FileNotFoundException(
                "A genuine MathType-generated Equation.DSMT4 fixture is required for the selection-view acceptance.",
                fixture);
        var path = Path.Combine(
            artifactRoot,
            "MathType-To-OMML-Selection-View.docx");
        File.Copy(fixture, path, overwrite: true);

        using var host = new WordPerformanceHost(path);
        Word.Range? prefixRange = null;
        Word.Range? suffixRange = null;
        Word.Range? content = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Selection? selection = null;
        Word.Window? window = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        try
        {
            var prefix = new System.Text.StringBuilder();
            for (var index = 1; index <= 180; index++)
                prefix.Append($"VIEW-PREFIX-{index:D3} 这是用于验证 MathType 转 OMML 后页面位置保持的长文档正文。\r");
            prefixRange = host.Document.Range(0, 0);
            prefixRange.Text = prefix.ToString();

            content = host.Document.Content;
            var suffixStart = Math.Max(content.Start, content.End - 1);
            suffixRange = host.Document.Range(suffixStart, suffixStart);
            var suffix = new System.Text.StringBuilder();
            for (var index = 1; index <= 60; index++)
                suffix.Append($"\rVIEW-SUFFIX-{index:D3} 页面位置回归测试尾部正文。");
            suffixRange.Text = suffix.ToString();

            host.Application.Visible = true;
            host.Document.Activate();
            host.Application.ActiveWindow.Activate();
            System.Windows.Forms.Application.DoEvents();

            shapes = host.Document.InlineShapes;
            if (shapes.Count != 1)
                throw new InvalidDataException(
                    $"Selection-view fixture contains {shapes.Count} inline shapes instead of one MathType equation.");
            shape = shapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidDataException(
                    "Selection-view fixture is no longer a MathType Equation.DSMT4 object.");
            shapeRange = shape.Range.Duplicate;
            shapeRange.Select();
            window = host.Application.ActiveWindow;
            window.ScrollIntoView(shapeRange, true);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);
            selection = host.Application.Selection;
            var selectionStartBefore = selection.Start;
            var selectionEndBefore = selection.End;
            var verticalBefore = window.VerticalPercentScrolled;
            var horizontalBefore = window.HorizontalPercentScrolled;
            if (verticalBefore < 20)
                throw new InvalidDataException(
                    $"Selection-view setup did not place the MathType equation deep enough in the document; vertical={verticalBefore}%.");

            host.AddIn.OnConvertMathTypeToOmmlSelection(new object());
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(15));
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(250);

            AssertEqual(0, CountMathTypeOleShapes(host.Document),
                "MathType→OMML selection-view acceptance left the MathType source behind.");
            maths = host.Document.OMaths;
            AssertEqual(1, maths.Count,
                "MathType→OMML selection-view acceptance did not create exactly one native OMath.");
            math = maths[1];
            mathRange = math.Range.Duplicate;

            Release(selection);
            selection = host.Application.Selection;
            var verticalAfter = window.VerticalPercentScrolled;
            var horizontalAfter = window.HorizontalPercentScrolled;
            var selectionStartAfter = selection.Start;
            var selectionEndAfter = selection.End;
            if (verticalAfter <= 5)
                throw new InvalidDataException(
                    $"MathType→OMML selected conversion jumped to the beginning of the document: before={verticalBefore}%, after={verticalAfter}%.");
            if (Math.Abs(verticalAfter - verticalBefore) > 3)
                throw new InvalidDataException(
                    $"MathType→OMML selected conversion changed the vertical viewport too much: before={verticalBefore}%, after={verticalAfter}%.");
            if (Math.Abs(horizontalAfter - horizontalBefore) > 3)
                throw new InvalidDataException(
                    $"MathType→OMML selected conversion changed horizontal viewport: before={horizontalBefore}%, after={horizontalAfter}%.");
            if (Math.Abs(selectionStartAfter - mathRange.Start) > 64
                || Math.Abs(selectionStartAfter - selectionStartBefore) > 64)
                throw new InvalidDataException(
                    "MathType→OMML selected conversion restored the viewport but moved Selection away from the converted equation. "
                    + $"before={selectionStartBefore}:{selectionEndBefore}; after={selectionStartAfter}:{selectionEndAfter}; omml={mathRange.Start}:{mathRange.End}.");

            host.Save(path);
            Console.WriteLine(
                "[MathType→OMML selection view] Preserved long-document viewport and selection vicinity: "
                + $"vertical {verticalBefore}%→{verticalAfter}%, selection {selectionStartBefore}:{selectionEndBefore}→{selectionStartAfter}:{selectionEndAfter}.");
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(window);
            Release(selection);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(content);
            Release(suffixRange);
            Release(prefixRange);
        }
    }

    private static void RunWordOmmlMathTypeSinglePerformanceAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var tracePath = Path.Combine(
            artifactRoot,
            "single-omml-to-mathtype.trace.log");
        try { File.Delete(tracePath); } catch { }
        var previousFormatAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousInjectedFailure = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                "1");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                tracePath);

            using var host = new WordPerformanceHost(documentPath: null);
            var prewarmWatch = System.Diagnostics.Stopwatch.StartNew();
            AssertTrue(
                MathTypeNativePreviewRenderer.WaitForSharedSessionPrewarm(
                    TimeSpan.FromSeconds(12)),
                "MathType shared native-preview prewarm did not finish in time.");
            prewarmWatch.Stop();
            Console.WriteLine(
                $"[SINGLE OMML→MT PREWARM] elapsedMs={prewarmWatch.ElapsedMilliseconds}");
            const string token = "VT_SINGLE_OMML_TO_MT_PERF";
            Word.Range? nativeRange = null;
            try
            {
                host.Document.Content.Text = $"before {token} after\r";
                nativeRange = InsertPureNativeOmml(
                    host.Document,
                    token,
                    "(x+1)/(y-2)");
                nativeRange.Select();

                var sessionsBefore = SnapshotSessionIds();
                var watch = System.Diagnostics.Stopwatch.StartNew();
                host.AddIn.OnConvertOmmlToMathTypeSelection(new object());
                var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(15);
                while (DateTime.UtcNow < deadline)
                {
                    System.Windows.Forms.Application.DoEvents();
                    if (host.Document.OMaths.Count == 0
                        && CountMathTypeOleShapes(host.Document) == 1)
                        break;
                    Thread.Sleep(15);
                }
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(5));
                watch.Stop();

                AssertEqual(0, host.Document.OMaths.Count,
                    "Single OMML→MathType Ribbon conversion left the source OMath behind.");
                AssertEqual(1, CountMathTypeOleShapes(host.Document),
                    "Single OMML→MathType Ribbon conversion did not create exactly one MathType OLE.");
                AssertEveryMathTypeProgId(host.Document);
                AssertSingleMathTypeTypingBoundaryMatchesProse(host);

                var sessionsAfter = SnapshotSessionIds();
                var createdSessions = sessionsAfter.Except(
                    sessionsBefore,
                    StringComparer.OrdinalIgnoreCase).ToArray();
                AssertEqual(0, createdSessions.Length,
                    "Direct-source OMML→MathType unexpectedly created Companion converter sessions: "
                    + string.Join(",", createdSessions));

                var trace = File.Exists(tracePath)
                    ? File.ReadAllText(tracePath)
                    : string.Empty;
                AssertTrue(
                    trace.IndexOf(
                        "format-conversion-render-bypass sourceMode=wordOmml targetMode=mathTypeOle targets=1 reason=source-mathml-ready",
                        StringComparison.Ordinal) >= 0,
                    "Single OMML→MathType trace did not prove the direct-MathML renderer bypass.");
                AssertTrue(
                    trace.IndexOf(
                        "format-conversion-render-start",
                        StringComparison.Ordinal) < 0,
                    "Single OMML→MathType still entered the Companion converter-render path.");

                var outputPath = Path.Combine(
                    artifactRoot,
                    "Single-OMML-To-MathType-Performance.docx");
                host.Save(outputPath);
                Console.WriteLine(
                    $"[SINGLE OMML→MT PERF] elapsedMs={watch.ElapsedMilliseconds}; "
                    + "companionSessions=0; sourceMathMlBypass=true; MathTypeObjects=1; "
                    + $"output={outputPath}");

                AssertSinglePureNativeDisplayOmmlFastPath(host);
                AssertSingleManagedOmmlAvoidsDirectDelete(host, tracePath);
                AssertSharedMathTypePreviewBatchConversion(host, tracePath);
                AssertSingleNumberedManagedOmmlFastPath(host, tracePath);
            }
            finally { Release(nativeRange); }

            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                "1");
            using (var rollbackHost = new WordPerformanceHost(documentPath: null))
            {
                Word.Range? rollbackRange = null;
                try
                {
                    const string rollbackToken = "VT_SINGLE_OMML_FAST_ROLLBACK";
                    rollbackHost.Document.Content.Text =
                        $"rollback-before {rollbackToken} rollback-after\r";
                    rollbackRange = InsertPureNativeOmml(
                        rollbackHost.Document,
                        rollbackToken,
                        "(p+q)/(r+s)");
                    rollbackRange.Select();
                    rollbackHost.AddIn.OnConvertOmmlToMathTypeSelection(new object());
                    WaitForAddInIdle(
                        rollbackHost.AddIn,
                        TimeSpan.FromSeconds(5));
                    AssertEqual(
                        1,
                        rollbackHost.Document.OMaths.Count,
                        "Injected failure did not restore the direct-delete OMML source.");
                    AssertEqual(
                        0,
                        CountMathTypeOleShapes(rollbackHost.Document),
                        "Injected failure left a MathType OLE after direct-delete rollback.");
                    var rollbackText = rollbackHost.Document.Content.Text ?? string.Empty;
                    AssertTrue(
                        rollbackText.IndexOf("rollback-before", StringComparison.Ordinal) >= 0
                        && rollbackText.IndexOf("rollback-after", StringComparison.Ordinal) >= 0,
                        "Direct-delete rollback damaged adjacent prose.");
                    Console.WriteLine(
                        "[SINGLE OMML→MT ROLLBACK] Injected post-delete failure restored the native OMath and adjacent prose.");
                }
                finally { Release(rollbackRange); }
            }
            using (var managedRollbackHost = new WordPerformanceHost(documentPath: null))
            {
                Word.Bookmark? managedBookmark = null;
                Word.Range? managedRange = null;
                try
                {
                    const string managedMathMl =
                        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
                        + "<mrow><munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">lim</mo>"
                        + "<mrow><mi>n</mi><mo>→</mo><mo>∞</mo></mrow></munder>"
                        + "<msup><mfenced><mrow><mn>1</mn><mo>+</mo><mfrac><mn>1</mn><mi>n</mi></mfrac>"
                        + "</mrow></mfenced><mi>n</mi></msup><mo>=</mo>"
                        + "<mi mathvariant=\"normal\">e</mi></mrow></math>";
                    managedRollbackHost.Document.Content.Text =
                        "managed-rollback-before\rmanaged-rollback-after\r";
                    managedRollbackHost.Document.Range(
                        managedRollbackHost.Document.Paragraphs[1].Range.End - 1,
                        managedRollbackHost.Document.Paragraphs[1].Range.End - 1).Select();
                    var managedService = new WordFormulaService(managedRollbackHost.Application);
                    var managedSession = CreateOmmlMathTypeAcceptanceSession(
                        managedMathMl,
                        "block",
                        false,
                        FormulaOleContract.WordOmmlMode);
                    managedService.InsertOmml(managedSession, managedMathMl);
                    managedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                        managedRollbackHost.Document,
                        managedSession.FormulaId)
                        ?? throw new InvalidDataException(
                            "Managed rollback setup lost its VTOMML bookmark.");
                    managedRange = WordOmmlFormulaStore.GetEquationRange(managedBookmark);
                    managedRange.Select();
                    managedRollbackHost.AddIn.OnConvertOmmlToMathTypeSelection(new object());
                    WaitForAddInIdle(managedRollbackHost.AddIn, TimeSpan.FromSeconds(8));
                    AssertEqual(1, managedRollbackHost.Document.OMaths.Count,
                        "Injected managed OMML failure did not restore the source OMath.");
                    AssertEqual(0, CountMathTypeOleShapes(managedRollbackHost.Document),
                        "Injected managed OMML failure left a MathType OLE behind.");
                    Release(managedBookmark);
                    managedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                        managedRollbackHost.Document,
                        managedSession.FormulaId);
                    AssertTrue(managedBookmark is not null,
                        "Injected managed OMML failure lost its VTOMML identity.");
                    var managedText = managedRollbackHost.Document.Content.Text ?? string.Empty;
                    AssertTrue(
                        managedText.IndexOf("managed-rollback-before", StringComparison.Ordinal) >= 0
                        && managedText.IndexOf("managed-rollback-after", StringComparison.Ordinal) >= 0,
                        "Injected managed OMML rollback damaged adjacent prose.");
                    Console.WriteLine(
                        "[MANAGED OMML→MT ROLLBACK] Injected post-delete failure restored the managed block OMath, VTOMML identity and adjacent prose.");
                }
                finally
                {
                    Release(managedRange);
                    Release(managedBookmark);
                }
            }
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                previousInjectedFailure);
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                previousInjectedFailure);
        }
    }

    private static void AssertSinglePureNativeDisplayOmmlFastPath(
        WordPerformanceHost host)
    {
        Word.Range? nativeRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? displayRange = null;
        try
        {
            const string token = "VT_SINGLE_DISPLAY_OMML_FAST";
            AppendAcceptanceText(
                host.Document,
                $"\rdisplay-before\r{token}\rdisplay-after\r");
            nativeRange = InsertPureNativeOmml(
                host.Document,
                token,
                "(u+v)/(w+1)");
            maths = nativeRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "Display fast-path setup did not retain exactly one native OMath.");
            math = maths[1];
            math.Type = Word.WdOMathType.wdOMathDisplay;
            displayRange = math.Range.Duplicate;
            displayRange.Select();

            var paragraphCountBefore = host.Document.Paragraphs.Count;
            var beforeCount = CountMathTypeOleShapes(host.Document);
            host.AddIn.OnConvertOmmlToMathTypeSelection(new object());
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(5));
            AssertEqual(
                0,
                host.Document.OMaths.Count,
                "Single display OMML fast path left the source OMath behind.");
            AssertEqual(
                beforeCount + 1,
                CountMathTypeOleShapes(host.Document),
                "Single display OMML fast path did not create one MathType OLE.");
            AssertEqual(
                paragraphCountBefore,
                host.Document.Paragraphs.Count,
                "Single display OMML→MathType inserted an extra paragraph below the converted equation.");
            AssertMathTypeDisplayFollowedImmediatelyByText(
                host.Document,
                "display-after",
                "Single display OMML→MathType");
            var text = host.Document.Content.Text ?? string.Empty;
            AssertTrue(
                text.IndexOf("display-before", StringComparison.Ordinal) >= 0
                && text.IndexOf("display-after", StringComparison.Ordinal) >= 0,
                "Single display OMML fast path damaged adjacent paragraphs.");
            Console.WriteLine(
                "[SINGLE DISPLAY OMML→MT] Pure native display OMath used the direct replacement path without damaging adjacent paragraphs.");
        }
        finally
        {
            Release(displayRange);
            Release(math);
            Release(maths);
            Release(nativeRange);
        }
    }

    private static void AssertMathTypeDisplayFollowedImmediatelyByText(
        Word.Document document,
        string expectedFollowingText,
        string context)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? followingParagraph = null;
        Word.Range? followingRange = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaRange = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        try
        {
            paragraphs = document.Paragraphs;
            var followingIndex = -1;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(followingRange);
                followingRange = null;
                Release(followingParagraph);
                followingParagraph = paragraphs[index];
                followingRange = followingParagraph.Range;
                var text = (followingRange.Text ?? string.Empty)
                    .Trim('\r', '\a', '\v', '\f', '\t', ' ');
                if (!string.Equals(text, expectedFollowingText, StringComparison.Ordinal))
                    continue;
                followingIndex = index;
                break;
            }
            AssertTrue(followingIndex > 1,
                $"{context} could not resolve the paragraph following the converted display formula.");

            formulaParagraph = paragraphs[followingIndex - 1];
            formulaRange = formulaParagraph.Range;
            shapes = formulaRange.InlineShapes;
            var mathTypeCount = 0;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) mathTypeCount++;
            }
            AssertEqual(1, mathTypeCount,
                $"{context} left an empty paragraph between the MathType display equation and '{expectedFollowingText}'.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(formulaRange);
            Release(formulaParagraph);
            Release(followingRange);
            Release(followingParagraph);
            Release(paragraphs);
        }
    }

    private static void AssertSingleManagedOmmlAvoidsDirectDelete(
        WordPerformanceHost host,
        string tracePath)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><msup><mi>t</mi><mn>2</mn></msup><mo>+</mo><mn>1</mn></math>";
        var service = new WordFormulaService(host.Application);
        var session = CreateOmmlMathTypeAcceptanceSession(
            mathMl,
            "inline",
            false,
            FormulaOleContract.WordOmmlMode);
        Word.Bookmark? bookmark = null;
        Word.Bookmark? staleBookmark = null;
        Word.Range? equationRange = null;
        try
        {
            SelectDocumentEnd(host.Document);
            service.InsertOmml(session, mathMl);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(
                host.Document,
                session.FormulaId)
                ?? throw new InvalidDataException(
                    "Managed OMML safety setup lost its VTOMML bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            equationRange.Select();
            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Managed OMML safety check did not capture exactly one equation.");
            AssertTrue(plan.Targets[0].SourceIsManagedOmml,
                "A VisualTeX-managed OMML equation was misclassified as pure native OMML.");

            var traceBefore = File.Exists(tracePath)
                ? File.ReadAllText(tracePath)
                : string.Empty;
            var traceLengthBefore = traceBefore.Length;
            equationRange.Select();
            var beforeCount = CountMathTypeOleShapes(host.Document);
            host.AddIn.OnConvertOmmlToMathTypeSelection(new object());
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(5));
            AssertEqual(
                beforeCount + 1,
                CountMathTypeOleShapes(host.Document),
                "Managed OMML did not convert to MathType.");
            staleBookmark = WordOmmlFormulaStore.FindByFormulaId(
                host.Document,
                session.FormulaId);
            AssertTrue(
                staleBookmark is null,
                "Managed OMML conversion left its VTOMML bookmark behind.");
            var trace = File.Exists(tracePath)
                ? File.ReadAllText(tracePath)
                : string.Empty;
            var traceSuffix = traceLengthBefore >= 0 && traceLengthBefore < trace.Length
                ? trace.Substring(traceLengthBefore)
                : trace;
            AssertTrue(
                traceSuffix.IndexOf(
                    "format-conversion-direct-omml-delete",
                    StringComparison.Ordinal) < 0,
                "Managed OMML incorrectly entered the pure-native direct-delete fast path.");
            Console.WriteLine(
                "[MANAGED OMML SAFETY] VisualTeX-managed OMML stayed on the metadata-aware replacement path and removed its VTOMML anchor.");
        }
        finally
        {
            Release(staleBookmark);
            Release(equationRange);
            Release(bookmark);
        }
    }

    private static void AssertSingleNumberedManagedOmmlFastPath(
        WordPerformanceHost host,
        string tracePath)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mrow><mi>y</mi><mo>−</mo><mn>2</mn></mrow></mfrac></math>";
        var service = new WordFormulaService(host.Application);
        var session = CreateOmmlMathTypeAcceptanceSession(
            mathMl,
            "block",
            true,
            FormulaOleContract.WordOmmlMode);
        Word.Bookmark? bookmark = null;
        Word.Bookmark? staleBookmark = null;
        Word.Range? equationRange = null;
        try
        {
            SelectDocumentEnd(host.Document);
            service.InsertOmml(session, mathMl);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(
                host.Document,
                session.FormulaId)
                ?? throw new InvalidDataException(
                    "Numbered managed OMML performance setup lost its VTOMML bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            equationRange.Select();
            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Numbered managed OMML performance capture did not isolate one equation.");
            AssertTrue(
                plan.Targets[0].SourceIsManagedOmml && plan.Targets[0].Numbered,
                "Numbered managed OMML performance source lost its managed/numbered state.");
            AssertEqual(0, CountMathTypePlaceRefFields(host.Document),
                "Numbered managed OMML performance setup unexpectedly has an older MathType number field.");

            var traceBefore = File.Exists(tracePath)
                ? File.ReadAllText(tracePath)
                : string.Empty;
            var traceLengthBefore = traceBefore.Length;
            var beforeCount = CountMathTypeOleShapes(host.Document);
            equationRange.Select();
            var watch = System.Diagnostics.Stopwatch.StartNew();
            host.AddIn.OnConvertOmmlToMathTypeSelection(new object());
            var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(10);
            while (DateTime.UtcNow < deadline)
            {
                System.Windows.Forms.Application.DoEvents();
                if (WordOmmlFormulaStore.FindByFormulaId(host.Document, session.FormulaId) is null
                    && CountMathTypeOleShapes(host.Document) == beforeCount + 1)
                    break;
                Thread.Sleep(15);
            }
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(5));
            watch.Stop();

            AssertEqual(beforeCount + 1, CountMathTypeOleShapes(host.Document),
                "Single numbered managed OMML did not convert to one MathType OLE.");
            AssertEqual(1, CountMathTypePlaceRefFields(host.Document),
                "Single numbered managed OMML did not create exactly one MTPlaceRef field.");
            staleBookmark = WordOmmlFormulaStore.FindByFormulaId(
                host.Document,
                session.FormulaId);
            AssertTrue(staleBookmark is null,
                "Single numbered managed OMML conversion left its VTOMML identity behind.");

            var trace = File.Exists(tracePath)
                ? File.ReadAllText(tracePath)
                : string.Empty;
            var traceSuffix = traceLengthBefore < trace.Length
                ? trace.Substring(traceLengthBefore)
                : trace;
            AssertTrue(
                traceSuffix.IndexOf(
                    "format-conversion-direct-numbered-omml-delete",
                    StringComparison.Ordinal) >= 0,
                "Single numbered managed OMML did not use the direct numbered source-removal path.");
            AssertTrue(
                traceSuffix.IndexOf(
                    "format-conversion-numbering-local-mathtype-single finalized=1",
                    StringComparison.Ordinal) >= 0,
                "Single numbered managed OMML still used the document-wide MathType numbering refresh.");
            AssertTrue(
                watch.ElapsedMilliseconds < 2500,
                $"Single numbered managed OMML→MathType remains too slow after prewarm: {watch.ElapsedMilliseconds}ms.");
            Console.WriteLine(
                $"[SINGLE NUMBERED OMML→MT PERF] elapsedMs={watch.ElapsedMilliseconds}; directNumberedDelete=true; localNumberFinalize=true.");
        }
        finally
        {
            Release(staleBookmark);
            Release(equationRange);
            Release(bookmark);
        }
    }

    private static void AssertSharedMathTypePreviewBatchConversion(
        WordPerformanceHost host,
        string tracePath)
    {
        var beforeCount = CountMathTypeOleShapes(host.Document);
        var traceBefore = File.Exists(tracePath)
            ? File.ReadAllText(tracePath)
            : string.Empty;
        for (var index = 1; index <= 6; index++)
        {
            var token = $"VT_SHARED_BATCH_OMML_{index}";
            AppendAcceptanceText(
                host.Document,
                $" batch-before-{index} {token} batch-after-{index}\r");
            Word.Range? range = null;
            try
            {
                range = InsertPureNativeOmml(
                    host.Document,
                    token,
                    $"x_{index}+{index}");
            }
            finally { Release(range); }
        }
        AssertEqual(6, host.Document.OMaths.Count,
            "Shared-preview batch setup did not create six native OMath equations.");

        host.AddIn.OnConvertOmmlToMathTypeDocument(new object());
        WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
        AssertEqual(0, host.Document.OMaths.Count,
            "Shared-preview batch conversion left native OMath equations behind.");
        AssertEqual(
            beforeCount + 6,
            CountMathTypeOleShapes(host.Document),
            "Shared-preview batch conversion did not create six MathType OLE objects.");
        AssertEveryMathTypeProgId(host.Document);
        var trace = File.Exists(tracePath)
            ? File.ReadAllText(tracePath)
            : string.Empty;
        var suffix = trace.Length > traceBefore.Length
            ? trace.Substring(traceBefore.Length)
            : trace;
        AssertTrue(
            suffix.IndexOf(
                "format-conversion-native-preview-batch-complete formulas=6",
                StringComparison.Ordinal) >= 0,
            "Shared-preview batch trace did not render all six MathType previews in one batch.");
        for (var index = 1; index <= 6; index++)
        {
            AssertTrue(
                (host.Document.Content.Text ?? string.Empty).IndexOf(
                    $"batch-before-{index}",
                    StringComparison.Ordinal) >= 0
                && (host.Document.Content.Text ?? string.Empty).IndexOf(
                    $"batch-after-{index}",
                    StringComparison.Ordinal) >= 0,
                $"Shared-preview batch conversion damaged prose around formula {index}.");
        }
        Console.WriteLine(
            "[SHARED MATHTYPE BATCH] Six native OMML equations used one native-preview batch and preserved adjacent prose.");
    }

    private static void AssertSingleMathTypeTypingBoundaryMatchesProse(
        WordPerformanceHost host)
    {
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Range? preceding = null;
        Word.Range? typed = null;
        Microsoft.Office.Interop.Word.Font? precedingFont = null;
        Microsoft.Office.Interop.Word.Font? typedFont = null;
        try
        {
            shapes = host.Document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    shape = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (shape is null)
                throw new InvalidDataException(
                    "Single OMML→MathType typing-boundary check found no MathType OLE.");

            shapeRange = shape.Range;
            preceding = host.Document.Range(
                Math.Max(host.Document.Content.Start, shapeRange.Start - 1),
                shapeRange.Start);
            precedingFont = preceding.Font;
            var expectedPosition = precedingFont.Position;
            if (expectedPosition == (int)Word.WdConstants.wdUndefined)
                expectedPosition = 0;

            var selection = host.Application.Selection;
            const string probe = "VTMT_TYPING_PROBE";
            selection.SetRange(shapeRange.End, shapeRange.End);
            var probeStart = selection.Start;
            selection.TypeText(probe);
            typed = host.Document.Range(probeStart, probeStart + probe.Length);
            typedFont = typed.Font;
            AssertEqual(
                expectedPosition,
                typedFont.Position,
                "Text typed immediately after the fast-path MathType OLE inherited the wrong baseline.");
            typed.Delete();
        }
        finally
        {
            Release(typedFont);
            Release(precedingFont);
            Release(typed);
            Release(preceding);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
        }
    }

    private static void RunWordMathTypeOmmlConsecutiveNumberedAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-omml-consecutive-numbered.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><text x=\"8\" y=\"78\" font-family=\"Cambria Math\" font-size=\"46\">det(A-λI)=0</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 112);
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>q</mi><mo>=</mo><mn>7</mn></mrow></math>";

        var mathTypeBaseline = SnapshotMathTypeProcessIds();
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            // This structural round-trip does not require a foreground Word window.
            // A second visible Word instance can block indefinitely behind the
            // user's active Word/OLE window before COM activation returns, which
            // tests foreground arbitration rather than formula conversion.
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Content.Text = "VT-CONSECUTIVE-BEFORE\r";
            var service = new WordFormulaService(application);

            for (var index = 0; index < 9; index++)
            {
                SelectDocumentEnd(document);
                service.InsertMathTypeOle(
                    CreateOmmlMathTypeAcceptanceSession(
                        mathMl,
                        "block",
                        numbered: true,
                        FormulaOleContract.MathTypeOleMode),
                    mathMl,
                    emfPath);
                if (index < 8)
                    AppendAcceptanceText(document, "\r");
            }
            AppendAcceptanceText(document, "\rVT-CONSECUTIVE-AFTER\r");

            AssertEqual(9, CountMathTypeOleShapes(document),
                "Consecutive numbered MathType→OMML setup did not create nine MathType equations.");
            AssertEqual(9, CountMathTypePlaceRefFields(document),
                "Consecutive numbered MathType→OMML setup did not create nine MTPlaceRef numbers.");

            var mathTypeReferenceTargets = MathTypeEquationReferences.GetTargets(document);
            AssertEqual(9, mathTypeReferenceTargets.Count,
                "Consecutive numbered MathType→OMML setup did not expose nine MathType reference targets.");
            SelectDocumentEnd(document);
            var referenceSelection = application.Selection;
            MathTypeEquationReferences.InsertReference(
                document,
                referenceSelection,
                mathTypeReferenceTargets[1],
                Word.WdColor.wdColorAutomatic);
            Release(referenceSelection);
            var referenceAlias = FindMathTypeReferenceAliasForAcceptance(document)
                ?? throw new InvalidDataException(
                    "Consecutive numbered MathType→OMML setup did not create a ZEqnNum reference alias.");
            AssertTrue(document.Bookmarks.Exists(referenceAlias),
                "Consecutive numbered MathType→OMML setup created a reference field without its ZEqnNum bookmark.");
            var referenceTextBefore = ReadMathTypeReferenceResultForAcceptance(document, referenceAlias);
            AssertTrue(!string.IsNullOrWhiteSpace(referenceTextBefore),
                "Consecutive numbered MathType→OMML setup created an empty MathType reference result.");
            var referenceBoldBefore = ReadMathTypeReferenceBoldForAcceptance(document, referenceAlias);
            AssertEqual(0, referenceBoldBefore,
                "Consecutive numbered MathType→OMML setup unexpectedly created a bold MathType reference.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(9, plan.Targets.Count,
                "Consecutive numbered MathType→OMML capture did not find all nine equations.");
            AssertTrue(
                plan.Targets.All(target => target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal)),
                "Consecutive numbered MathType→OMML capture lost numbered display ownership.");

            var secondPlanTarget = plan.Targets
                .OrderBy(target => target.SourceStart)
                .ElementAt(1);
            var preparedTargets = PrepareOmmlMathTypeTargets(plan, emfPath);
            var result = service.ApplyFormulaFormatConversionPlan(
                plan,
                preparedTargets);
            AssertEqual(9, result.FormulaCount,
                "Consecutive numbered MathType→OMML conversion did not convert all nine equations: "
                + string.Join(" | ", result.Failures));
            AssertEqual(0, result.FailedFormulaCount,
                "Consecutive numbered MathType→OMML conversion reported a failure: "
                + string.Join(" | ", result.Failures));
            AssertConsecutiveNumberedOmmlStructure(application, document, "live conversion");
            var secondOmmlFormulaId = preparedTargets[secondPlanTarget.Id].Session.FormulaId;
            AssertConvertedMathTypeReferenceAlias(
                application,
                document,
                referenceAlias,
                referenceTextBefore,
                secondOmmlFormulaId,
                "live conversion");

            var outputPath = Path.Combine(
                artifactRoot,
                "MathType-To-OMML-Consecutive-Numbered-9.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertConsecutiveNumberedOmmlStructure(application, document, "save/reopen");
            AssertConvertedMathTypeReferenceAlias(
                application,
                document,
                referenceAlias,
                referenceTextBefore,
                secondOmmlFormulaId,
                "save/reopen");

            var visualTeXReferenceTargets = WordEquationNumbering.GetEquationReferenceTargets(document);
            var visualTeXReferenceTarget = visualTeXReferenceTargets.Single(target =>
                string.Equals(target.FormulaId, secondOmmlFormulaId, StringComparison.Ordinal));
            SelectDocumentEnd(document);
            var visualTeXReferenceSelection = application.Selection;
            WordEquationNumbering.InsertEquationReference(
                document,
                visualTeXReferenceSelection,
                visualTeXReferenceTarget,
                EquationReferenceStyle.NumberOnly);
            Release(visualTeXReferenceSelection);
            var visualTeXReferenceAlias = WordEquationNumbering.NativeNumberBookmarkName(secondOmmlFormulaId);
            var visualTeXReferenceTextBefore = ReadMathTypeReferenceResultForAcceptance(
                document,
                visualTeXReferenceAlias);
            AssertEqual(
                visualTeXReferenceTarget.NumberText.Trim(),
                visualTeXReferenceTextBefore.Trim(),
                "OMML reference setup did not create a dynamic VTEqNum reference to the second formula.");
            AssertVisualTeXReferenceAlias(
                application,
                document,
                visualTeXReferenceAlias,
                visualTeXReferenceTextBefore,
                secondOmmlFormulaId,
                "OMML reference setup");
            var preReverseAliases = MathTypeEquationReferences.CaptureFormatConversionAliasesFromVisualTeX(
                document,
                secondOmmlFormulaId);
            AssertTrue(preReverseAliases.Any(alias => string.Equals(
                    alias.Name,
                    referenceAlias,
                    StringComparison.OrdinalIgnoreCase)),
                "OMML→MathType setup did not capture the inherited ZEqnNum compatibility alias.");
            AssertTrue(preReverseAliases.Any(alias => string.Equals(
                    alias.Name,
                    visualTeXReferenceAlias,
                    StringComparison.OrdinalIgnoreCase)),
                "OMML→MathType setup did not capture the dynamic VTEqNum compatibility alias.");

            service = new WordFormulaService(application);
            var ommlToMathTypePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(9, ommlToMathTypePlan.Targets.Count,
                "Reference round-trip OMML→MathType capture did not find all nine formulas.");
            var ommlToMathTypePrepared = PrepareOmmlMathTypeTargets(ommlToMathTypePlan, emfPath);
            var ommlToMathTypeResult = service.ApplyFormulaFormatConversionPlan(
                ommlToMathTypePlan,
                ommlToMathTypePrepared);
            AssertEqual(9, ommlToMathTypeResult.FormulaCount,
                "Reference round-trip OMML→MathType did not convert all nine formulas: "
                + string.Join(" | ", ommlToMathTypeResult.Failures));
            AssertEqual(0, ommlToMathTypeResult.FailedFormulaCount,
                "Reference round-trip OMML→MathType failed: "
                + string.Join(" | ", ommlToMathTypeResult.Failures));
            AssertEqual(0, document.OMaths.Count,
                "Reference round-trip OMML→MathType left OMML formulas behind.");
            AssertEqual(9, CountMathTypeOleShapes(document),
                "Reference round-trip OMML→MathType did not create nine MathType equations.");
            AssertEqual(9, CountMathTypePlaceRefFields(document),
                "Reference round-trip OMML→MathType did not recreate nine MTPlaceRef fields.");
            AssertReferenceAliasesOnMathType(
                application,
                document,
                referenceAlias,
                referenceTextBefore,
                referenceBoldBefore,
                visualTeXReferenceAlias,
                visualTeXReferenceTextBefore,
                "OMML→MathType live conversion");

            var mathTypeRoundTripPath = Path.Combine(
                artifactRoot,
                "OMML-To-MathType-Reference-RoundTrip.docx");
            document.SaveAs2(mathTypeRoundTripPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                mathTypeRoundTripPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertReferenceAliasesOnMathType(
                application,
                document,
                referenceAlias,
                referenceTextBefore,
                referenceBoldBefore,
                visualTeXReferenceAlias,
                visualTeXReferenceTextBefore,
                "OMML→MathType save/reopen");

            service = new WordFormulaService(application);
            var secondMathTypeToOmmlPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(9, secondMathTypeToOmmlPlan.Targets.Count,
                "Second MathType→OMML reference round-trip capture did not find all nine formulas.");
            var secondMathTypeTarget = secondMathTypeToOmmlPlan.Targets
                .OrderBy(target => target.SourceStart)
                .ElementAt(1);
            var secondMathTypeToOmmlPrepared = PrepareOmmlMathTypeTargets(
                secondMathTypeToOmmlPlan,
                emfPath);
            var secondMathTypeToOmmlResult = service.ApplyFormulaFormatConversionPlan(
                secondMathTypeToOmmlPlan,
                secondMathTypeToOmmlPrepared);
            AssertEqual(9, secondMathTypeToOmmlResult.FormulaCount,
                "Second MathType→OMML reference round-trip did not convert all nine formulas: "
                + string.Join(" | ", secondMathTypeToOmmlResult.Failures));
            AssertEqual(0, secondMathTypeToOmmlResult.FailedFormulaCount,
                "Second MathType→OMML reference round-trip failed: "
                + string.Join(" | ", secondMathTypeToOmmlResult.Failures));
            var finalSecondOmmlFormulaId = secondMathTypeToOmmlPrepared[secondMathTypeTarget.Id]
                .Session.FormulaId;
            AssertConsecutiveNumberedOmmlStructure(application, document, "second live conversion");
            AssertConvertedMathTypeReferenceAlias(
                application,
                document,
                referenceAlias,
                referenceTextBefore,
                finalSecondOmmlFormulaId,
                "second MathType→OMML live conversion");
            AssertVisualTeXReferenceAlias(
                application,
                document,
                visualTeXReferenceAlias,
                visualTeXReferenceTextBefore,
                finalSecondOmmlFormulaId,
                "second MathType→OMML live conversion");

            var finalRoundTripPath = Path.Combine(
                artifactRoot,
                "MathType-OMML-Reference-Repeated-RoundTrip.docx");
            document.SaveAs2(finalRoundTripPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalRoundTripPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertConsecutiveNumberedOmmlStructure(application, document, "second save/reopen");
            AssertConvertedMathTypeReferenceAlias(
                application,
                document,
                referenceAlias,
                referenceTextBefore,
                finalSecondOmmlFormulaId,
                "second MathType→OMML save/reopen");
            AssertVisualTeXReferenceAlias(
                application,
                document,
                visualTeXReferenceAlias,
                visualTeXReferenceTextBefore,
                finalSecondOmmlFormulaId,
                "second MathType→OMML save/reopen");

            AssertMathTypeDisplayToOmmlPreservesParagraphStructure(
                application,
                document,
                emfPath);
            AssertNoNewMathTypeProcess(
                mathTypeBaseline,
                "consecutive numbered MathType→OMML acceptance");

            Console.WriteLine(
                "[CONSECUTIVE NUMBERED MATHTYPE→OMML] Nine adjacent numbered MathType display equations converted to nine independent direct-SEQ 1x3 hosts with one genuine center-cell wdOMathDisplay each, zero Shape/TextBox artifacts, preserved cross-reference aliases, and repeated save/reopen persistence.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertConsecutiveNumberedOmmlStructure(
        Word.Application application,
        Word.Document document,
        string stage)
    {
        AssertEqual(9, document.OMaths.Count,
            $"Consecutive numbered MathType→OMML {stage} did not retain nine OMath equations.");
        AssertEqual(0, CountMathTypeOleShapes(document),
            $"Consecutive numbered MathType→OMML {stage} left a MathType source behind.");
        AssertEqual(9, document.Tables.Count,
            $"Consecutive numbered MathType→OMML {stage} did not retain one direct-SEQ 1x3 table per formula.");
        AssertEqual(9, CountManagedNumberedOmml(document),
            $"Consecutive numbered MathType→OMML {stage} lost managed numbering ownership.");
        AssertEqual(0, document.Shapes.Count,
            $"Consecutive numbered MathType→OMML {stage} recreated a retired Shape/TextBox number.");
        AssertManagedNativeOmmlInterTableSeparatorsCompact(
            document,
            WordOmmlFormulaStore.FormulaIds(document).ToArray(),
            $"consecutive numbered MathType→OMML {stage}");

        var ordered = new List<(string FormulaId, int Start)>();
        foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? equationRange = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (metadata?.Numbered != true
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal))
                    continue;
                equationRange = WordOmmlFormulaStore
                    .GetEquationRangeVerifiedForStructuralEdit(
                        document,
                        formulaId,
                        metadata);
                ordered.Add((formulaId, equationRange.Start));
            }
            finally
            {
                Release(equationRange);
                Release(bookmark);
            }
        }
        ordered.Sort((left, right) => left.Start.CompareTo(right.Start));
        AssertEqual(9, ordered.Count,
            $"Consecutive numbered MathType→OMML {stage} did not resolve nine managed formula identities.");
        var orderedIdSet = ordered
            .Select(item => item.FormulaId)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var targetByFormulaId = WordEquationNumbering
            .GetEquationReferenceTargets(document)
            .Where(target => orderedIdSet.Contains(target.FormulaId))
            .ToDictionary(
                target => target.FormulaId,
                target => target.NumberText.Trim(),
                StringComparer.OrdinalIgnoreCase);
        AssertEqual(9, targetByFormulaId.Count,
            $"Consecutive numbered MathType→OMML {stage} did not expose nine dynamic reference targets.");

        for (var index = 0; index < ordered.Count; index++)
        {
            var formulaId = ordered[index].FormulaId;
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                $"consecutive MathType→OMML {stage} formula {index + 1}");

            Word.Range? visibleRange = null;
            Word.Fields? fields = null;
            Word.Field? reference = null;
            Word.Range? code = null;
            Word.Range? result = null;
            Word.Range? prefixRange = null;
            try
            {
                visibleRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    formulaId)
                    ?? throw new InvalidDataException(
                        $"Consecutive numbered MathType→OMML {stage} formula {index + 1} lost its visible direct-SEQ label.");
                AssertEqual(Word.WdStoryType.wdMainTextStory, visibleRange.StoryType,
                    $"Consecutive numbered MathType→OMML {stage} formula {index + 1} moved its number outside the main document story.");
                AssertTrue((bool)visibleRange.get_Information(Word.WdInformation.wdWithInTable),
                    $"Consecutive numbered MathType→OMML {stage} formula {index + 1} moved its number outside the 1x3 table.");
                fields = visibleRange.Fields;
                AssertEqual(1, fields.Count,
                    $"Consecutive numbered MathType→OMML {stage} formula {index + 1} has an invalid direct SEQ count.");
                reference = fields[1];
                reference.Update();
                // Heading-aware direct-table numbering stores the heading prefix
                // as ordinary Word text before the SEQ field, while the field result
                // itself is only the local ordinal. Reconstruct the visible number
                // from the literal prefix + Field.Result rather than Range.Text;
                // the latter exposes the complete SEQ instruction when Alt+F9 /
                // ShowFieldCodes is enabled.
                code = reference.Code;
                result = reference.Result;
                var prefixStart = Math.Min(visibleRange.End, visibleRange.Start + 1);
                var fieldBegin = Math.Max(prefixStart, code.Start - 1);
                prefixRange = document.Range(prefixStart, fieldBegin);
                var actualNumber = NormalizeNumberedOmmlLabel(
                        (prefixRange.Text ?? string.Empty)
                        + (result.Text ?? string.Empty))
                    .Trim('(', ')');
                AssertEqual(
                    targetByFormulaId[formulaId],
                    actualNumber,
                    $"Consecutive numbered MathType→OMML {stage} formula {index + 1} does not match its dynamic reference target.");
                var trailingOrdinal = System.Text.RegularExpressions.Regex.Match(
                    actualNumber,
                    @"(?<ordinal>\d+)\s*$",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant);
                AssertTrue(trailingOrdinal.Success,
                    $"Consecutive numbered MathType→OMML {stage} formula {index + 1} has no trailing ordinal: '{actualNumber}'.");
                AssertEqual(
                    (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    trailingOrdinal.Groups["ordinal"].Value,
                    $"Consecutive numbered MathType→OMML {stage} numbering is not continuous at formula {index + 1}.");
            }
            finally
            {
                Release(prefixRange);
                Release(result);
                Release(code);
                Release(reference);
                Release(fields);
                Release(visibleRange);
            }
        }
    }

    private static void AssertConsecutiveNumberedOmmlFormulaStructure(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Range? equationRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Fields? equationFields = null;
        Word.Paragraphs? formulaParagraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaParagraphRange = null;
        Word.Range? visibleRange = null;
        Word.Fields? visibleFields = null;
        Word.Field? reference = null;
        Word.Range? fieldCode = null;
        Word.Shape? shape = null;
        Word.Range? shapeAnchor = null;
        Word.Paragraphs? anchorParagraphs = null;
        Word.Paragraph? anchorParagraph = null;
        Word.Range? anchorParagraphRange = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                formulaId)
                ?? throw new InvalidDataException(
                    context + ": the managed OMML identity bookmark is missing.");
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaBookmark)
                ?? throw new InvalidDataException(
                    context + ": OMML metadata is missing.");
            AssertTrue(metadata.Numbered
                    && string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal),
                context + ": metadata no longer describes a numbered block formula.");
            equationRange = WordOmmlFormulaStore
                .GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
            AssertTrue(!(bool)equationRange.get_Information(
                    Word.WdInformation.wdWithInTable),
                context + ": the formula is still inside a legacy table.");
            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": the formula does not contain exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": the formula is not genuine Word Display math.");
            equationFields = equationRange.Fields;
            AssertEqual(0, equationFields.Count,
                context + ": a field leaked inside OMath.");

            formulaParagraphs = equationRange.Paragraphs;
            AssertEqual(1, formulaParagraphs.Count,
                context + ": the formula spans more than one Word paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            var ownerXml = formulaParagraphRange.WordOpenXML ?? string.Empty;
            AssertTrue(ownerXml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": m:oMathPara is missing.");
            AssertTrue(ownerXml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) < 0,
                context + ": the obsolete m:eqArr/#(...) wrapper returned.");
            AssertTrue(ownerXml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) < 0,
                context + ": field controls leaked into the formula paragraph.");
            AssertTrue(ownerXml.IndexOf(" REF ", StringComparison.OrdinalIgnoreCase) < 0,
                context + ": the dynamic REF leaked into the formula paragraph.");
            AssertTrue(ownerXml.IndexOf("<w:drawing", StringComparison.OrdinalIgnoreCase) < 0,
                context + ": the external number Shape leaked into the formula paragraph.");

            visibleRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                document,
                formulaId)
                ?? throw new InvalidDataException(
                    context + ": the external equation-number range is missing.");
            AssertEqual(Word.WdStoryType.wdTextFrameStory, visibleRange.StoryType,
                context + ": the visible equation number is not external Word text.");
            visibleFields = visibleRange.Fields;
            AssertEqual(1, visibleFields.Count,
                context + ": the external label does not contain exactly one REF field.");
            reference = visibleFields[1];
            fieldCode = reference.Code;
            AssertTrue((fieldCode.Text ?? string.Empty).IndexOf(
                    "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the external field targets the wrong equation-number bookmark.");
            AssertEqual(0, fieldCode.OMaths.Count,
                context + ": the external REF code is inside OMath.");

            shape = FindNumberedOmmlShape(document, formulaId, context);
            shapeAnchor = shape.Anchor;
            anchorParagraphs = shapeAnchor.Paragraphs;
            AssertEqual(1, anchorParagraphs.Count,
                context + ": the Shape anchor spans more than one paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            AssertEqual(anchorParagraphRange.End, formulaParagraphRange.Start,
                context + ": the dedicated Shape anchor is not immediately before the display formula.");
            AssertEqual(0, anchorParagraphRange.OMaths.Count,
                context + ": mathematical content leaked into the Shape anchor paragraph.");
            AssertEqual(0, anchorParagraphRange.Fields.Count,
                context + ": fields leaked into the Shape anchor paragraph.");
        }
        finally
        {
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(shapeAnchor);
            Release(shape);
            Release(fieldCode);
            Release(reference);
            Release(visibleFields);
            Release(visibleRange);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(equationFields);
            Release(math);
            Release(maths);
            Release(equationRange);
            Release(formulaBookmark);
        }
    }

    private static string? FindMathTypeReferenceAliasForAcceptance(Word.Document document)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ');
                var marker = "REF ZEqnNum";
                var markerIndex = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (markerIndex < 0) continue;
                var start = markerIndex + "REF ".Length;
                var end = start;
                while (end < text.Length
                       && !char.IsWhiteSpace(text[end])
                       && text[end] != '\\'
                       && text[end] != '\u0014'
                       && text[end] != '\u0015')
                    end++;
                var alias = text.Substring(start, end - start).Trim();
                if (alias.StartsWith("ZEqnNum", StringComparison.OrdinalIgnoreCase))
                    return alias;
            }
            return null;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static string ReadMathTypeReferenceResultForAcceptance(
        Word.Document document,
        string alias)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty).TrimStart();
                if (!text.StartsWith("REF " + alias, StringComparison.OrdinalIgnoreCase))
                    continue;
                result = field.Result;
                return (result.Text ?? string.Empty).Trim();
            }
            return string.Empty;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int ReadMathTypeReferenceBoldForAcceptance(
        Word.Document document,
        string alias)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty).TrimStart();
                if (!text.StartsWith("REF " + alias, StringComparison.OrdinalIgnoreCase))
                    continue;
                result = field.Result;
                font = result.Font;
                return font.Bold;
            }
            throw new InvalidDataException($"Could not find REF {alias} for bold-format acceptance.");
        }
        finally
        {
            Release(font);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void AssertConvertedMathTypeReferenceAlias(
        Word.Application application,
        Word.Document document,
        string alias,
        string expectedReferenceText,
        string formulaId,
        string stage)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? aliasBookmark = null;
        Word.Bookmark? nativeBookmark = null;
        Word.Range? aliasRange = null;
        Word.Range? nativeRange = null;
        Word.Range? visibleNumberRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Field? outerGoTo = null;
        Word.Field? nestedRef = null;
        Word.Range? referenceResult = null;
        Word.Selection? selection = null;
        try
        {
            bookmarks = document.Bookmarks;
            var nativeName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertTrue(bookmarks.Exists(alias),
                $"MathType→OMML {stage} lost reference compatibility bookmark {alias}.");
            AssertTrue(bookmarks.Exists(nativeName),
                $"MathType→OMML {stage} lost native OMML number bookmark {nativeName}.");
            aliasBookmark = bookmarks[alias];
            nativeBookmark = bookmarks[nativeName];
            aliasRange = aliasBookmark.Range;
            nativeRange = nativeBookmark.Range;
            visibleNumberRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    $"MathType→OMML {stage} cannot resolve the visible number range for {formulaId}.");
            AssertEqual(visibleNumberRange.Start, aliasRange.Start,
                $"MathType→OMML {stage} restored {alias} at the wrong visible-number start position.");
            AssertEqual(visibleNumberRange.End, aliasRange.End,
                $"MathType→OMML {stage} restored {alias} at the wrong visible-number end position.");
            AssertTrue(nativeRange.Start >= aliasRange.Start
                       && nativeRange.End <= aliasRange.End,
                $"MathType→OMML {stage} restored {alias} outside the durable VTEqNum number identity.");
            AssertEqual(
                expectedReferenceText.Trim(),
                ReadVisibleEquationNumber(document, formulaId),
                $"MathType→OMML {stage} restored {alias} over the wrong visible number text.");

            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (text.StartsWith("REF " + alias, StringComparison.OrdinalIgnoreCase))
                {
                    Release(nestedRef);
                    nestedRef = field;
                    field = null;
                    continue;
                }
                if (text.StartsWith("GOTOBUTTON " + alias, StringComparison.OrdinalIgnoreCase))
                {
                    Release(outerGoTo);
                    outerGoTo = field;
                    field = null;
                }
            }
            AssertTrue(nestedRef is not null,
                $"MathType→OMML {stage} lost nested REF {alias}.");
            AssertTrue(outerGoTo is not null,
                $"MathType→OMML {stage} lost GOTOBUTTON {alias}.");
            nestedRef!.Update();
            referenceResult = nestedRef.Result;
            AssertEqual(
                expectedReferenceText.Trim(),
                (referenceResult.Text ?? string.Empty).Trim(),
                $"MathType→OMML {stage} changed the visible MathType reference text.");

            outerGoTo!.DoClick();
            selection = application.Selection;
            AssertTrue(
                selection.Start >= aliasRange.Start && selection.Start <= aliasRange.End,
                $"MathType→OMML {stage} GOTOBUTTON did not navigate to restored bookmark {alias}.");
        }
        finally
        {
            Release(selection);
            Release(referenceResult);
            Release(nestedRef);
            Release(outerGoTo);
            Release(code);
            Release(field);
            Release(fields);
            Release(visibleNumberRange);
            Release(nativeRange);
            Release(aliasRange);
            Release(nativeBookmark);
            Release(aliasBookmark);
            Release(bookmarks);
        }
    }

    private static string ReadRangeTextWithFieldResults(
        Word.Document document,
        Word.Range range)
    {
        Word.View? view = null;
        var restoreFieldCodes = false;
        try
        {
            view = document.ActiveWindow.View;
            restoreFieldCodes = view.ShowFieldCodes;
            if (restoreFieldCodes)
            {
                view.ShowFieldCodes = false;
                System.Windows.Forms.Application.DoEvents();
            }
            return (range.Text ?? string.Empty).Trim();
        }
        finally
        {
            if (view is not null && restoreFieldCodes)
            {
                try { view.ShowFieldCodes = true; } catch { }
            }
            Release(view);
        }
    }

    private static void AssertReferenceAliasesOnMathType(
        Word.Application application,
        Word.Document document,
        string mathTypeAlias,
        string expectedMathTypeText,
        int expectedMathTypeBold,
        string visualTeXAlias,
        string expectedVisualTeXText,
        string stage)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? mathTypeBookmark = null;
        Word.Bookmark? visualTeXBookmark = null;
        Word.Range? mathTypeRange = null;
        Word.Range? visualTeXRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Field? goTo = null;
        Word.Field? visualTeXGoTo = null;
        Word.Range? result = null;
        Word.Selection? selection = null;
        Word.View? view = null;
        var restoreFieldCodes = false;
        try
        {
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(mathTypeAlias),
                $"{stage} lost MathType compatibility bookmark {mathTypeAlias}.");
            AssertTrue(bookmarks.Exists(visualTeXAlias),
                $"{stage} lost VisualTeX compatibility bookmark {visualTeXAlias}.");
            mathTypeBookmark = bookmarks[mathTypeAlias];
            visualTeXBookmark = bookmarks[visualTeXAlias];
            mathTypeRange = mathTypeBookmark.Range;
            visualTeXRange = visualTeXBookmark.Range;
            AssertEqual(
                expectedMathTypeText.Trim(),
                ReadRangeTextWithFieldResults(document, mathTypeRange),
                $"{stage} restored the MathType alias over the wrong visible-number span.");
            AssertEqual(
                expectedVisualTeXText.Trim(),
                ReadRangeTextWithFieldResults(document, visualTeXRange),
                $"{stage} restored the VisualTeX alias over the wrong number-only span.");

            fields = document.Fields;
            var mathTypeRefFound = false;
            var visualTeXRefFound = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (text.StartsWith("REF " + mathTypeAlias, StringComparison.OrdinalIgnoreCase))
                {
                    Release(result);
                    result = field.Result;
                    AssertEqual(expectedMathTypeText.Trim(), (result.Text ?? string.Empty).Trim(),
                        $"{stage} changed the MathType reference result.");
                    AssertEqual(expectedMathTypeBold, result.Font.Bold,
                        $"{stage} changed the MathType reference bold formatting.");
                    mathTypeRefFound = true;
                }
                if (text.StartsWith("REF " + visualTeXAlias, StringComparison.OrdinalIgnoreCase))
                {
                    field.Update();
                    Release(result);
                    result = field.Result;
                    AssertEqual(expectedVisualTeXText.Trim(), (result.Text ?? string.Empty).Trim(),
                        $"{stage} changed the VisualTeX reference result.");
                    visualTeXRefFound = true;
                }
                if (text.StartsWith("GOTOBUTTON " + mathTypeAlias, StringComparison.OrdinalIgnoreCase))
                {
                    Release(goTo);
                    goTo = field;
                    field = null;
                    continue;
                }
                if (text.StartsWith("GOTOBUTTON " + visualTeXAlias, StringComparison.OrdinalIgnoreCase))
                {
                    Release(visualTeXGoTo);
                    visualTeXGoTo = field;
                    field = null;
                }
            }
            AssertTrue(mathTypeRefFound,
                $"{stage} lost the dynamic REF field for {mathTypeAlias}.");
            AssertTrue(visualTeXRefFound,
                $"{stage} lost the dynamic REF field for {visualTeXAlias}.");
            AssertTrue(goTo is not null,
                $"{stage} lost GOTOBUTTON {mathTypeAlias}.");
            AssertTrue(visualTeXGoTo is not null,
                $"{stage} lost GOTOBUTTON {visualTeXAlias}.");

            // Updating nested REF fields can cause Word to rematerialize the outer
            // GOTOBUTTON field. Resolve it again immediately before DoClick so the
            // test exercises the same live field object a real user double-clicks.
            Release(goTo);
            goTo = null;
            Release(visualTeXGoTo);
            visualTeXGoTo = null;
            Release(code);
            code = null;
            Release(field);
            field = null;
            Release(fields);
            fields = document.Fields;
            var expectedGoToSelectionStart = -1;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (MathTypeEquationReferences.IsMathTypePlaceRefCode(text)
                    && mathTypeRange.Start >= code.Start
                    && mathTypeRange.End <= code.End)
                {
                    expectedGoToSelectionStart = Math.Max(document.Content.Start, code.Start - 1);
                }
                if (text.StartsWith("GOTOBUTTON " + mathTypeAlias, StringComparison.OrdinalIgnoreCase))
                {
                    Release(goTo);
                    goTo = field;
                    field = null;
                    continue;
                }
                if (!text.StartsWith("GOTOBUTTON " + visualTeXAlias, StringComparison.OrdinalIgnoreCase))
                    continue;
                Release(visualTeXGoTo);
                visualTeXGoTo = field;
                field = null;
            }
            AssertTrue(goTo is not null,
                $"{stage} lost live GOTOBUTTON {mathTypeAlias} after REF refresh.");
            AssertTrue(visualTeXGoTo is not null,
                $"{stage} lost live GOTOBUTTON {visualTeXAlias} after REF refresh.");
            AssertTrue(expectedGoToSelectionStart >= 0,
                $"{stage} could not resolve the MTPlaceRef owner of {mathTypeAlias}.");
            view = document.ActiveWindow.View;
            restoreFieldCodes = view.ShowFieldCodes;
            if (restoreFieldCodes)
            {
                view.ShowFieldCodes = false;
                System.Windows.Forms.Application.DoEvents();
            }
            goTo!.DoClick();
            selection = application.Selection;
            AssertEqual(expectedGoToSelectionStart, selection.Start,
                $"{stage} GOTOBUTTON did not navigate to the MTPlaceRef owner of {mathTypeAlias}; bookmark={mathTypeRange.Start}:{mathTypeRange.End}, selection={selection.Start}:{selection.End}.");
            Release(selection);
            selection = null;
            visualTeXGoTo!.DoClick();
            selection = application.Selection;
            AssertEqual(expectedGoToSelectionStart, selection.Start,
                $"{stage} GOTOBUTTON did not navigate to the MTPlaceRef owner of {visualTeXAlias}; bookmark={visualTeXRange.Start}:{visualTeXRange.End}, selection={selection.Start}:{selection.End}.");
        }
        finally
        {
            if (view is not null && restoreFieldCodes)
            {
                try { view.ShowFieldCodes = true; } catch { }
            }
            Release(view);
            Release(selection);
            Release(result);
            Release(visualTeXGoTo);
            Release(goTo);
            Release(code);
            Release(field);
            Release(fields);
            Release(visualTeXRange);
            Release(mathTypeRange);
            Release(visualTeXBookmark);
            Release(mathTypeBookmark);
            Release(bookmarks);
        }
    }

    private static void AssertVisualTeXReferenceAlias(
        Word.Application application,
        Word.Document document,
        string alias,
        string expectedReferenceText,
        string formulaId,
        string stage)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? aliasBookmark = null;
        Word.Bookmark? nativeBookmark = null;
        Word.Range? aliasRange = null;
        Word.Range? nativeRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        Word.Field? goTo = null;
        Word.Selection? selection = null;
        Word.View? view = null;
        var restoreFieldCodes = false;
        var referenceFound = false;
        try
        {
            bookmarks = document.Bookmarks;
            var nativeName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertTrue(bookmarks.Exists(alias),
                $"{stage} lost VisualTeX reference compatibility bookmark {alias}.");
            AssertTrue(bookmarks.Exists(nativeName),
                $"{stage} lost native OMML number bookmark {nativeName}.");
            aliasBookmark = bookmarks[alias];
            nativeBookmark = bookmarks[nativeName];
            aliasRange = aliasBookmark.Range;
            nativeRange = nativeBookmark.Range;
            AssertEqual(nativeRange.Start, aliasRange.Start,
                $"{stage} restored {alias} at the wrong number-only start position.");
            AssertEqual(nativeRange.End, aliasRange.End,
                $"{stage} restored {alias} at the wrong number-only end position.");
            AssertEqual(
                expectedReferenceText.Trim(),
                ReadVisibleEquationNumber(document, formulaId).Trim().Trim('(', ')'),
                $"{stage} restored {alias} over the wrong number text.");

            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty).TrimStart();
                if (text.StartsWith("REF " + alias, StringComparison.OrdinalIgnoreCase))
                {
                    field.Update();
                    Release(result);
                    result = field.Result;
                    AssertEqual(expectedReferenceText.Trim(), (result.Text ?? string.Empty).Trim(),
                        $"{stage} changed the dynamic VisualTeX reference result.");
                    referenceFound = true;
                    continue;
                }
                if (!text.StartsWith("GOTOBUTTON " + alias, StringComparison.OrdinalIgnoreCase))
                    continue;
                Release(goTo);
                goTo = field;
                field = null;
            }
            AssertTrue(referenceFound, $"{stage} lost dynamic REF {alias}.");
            AssertTrue(goTo is not null,
                $"{stage} lost the navigable GOTOBUTTON field for {alias}.");
            view = document.ActiveWindow.View;
            restoreFieldCodes = view.ShowFieldCodes;
            if (restoreFieldCodes)
            {
                view.ShowFieldCodes = false;
                System.Windows.Forms.Application.DoEvents();
            }
            goTo!.DoClick();
            selection = application.Selection;
            AssertTrue(
                selection.Start >= Math.Max(document.Content.Start, aliasRange.Start - 1)
                && selection.Start <= Math.Min(document.Content.End, aliasRange.End + 1),
                $"{stage} GOTOBUTTON did not navigate to {alias}; bookmark={aliasRange.Start}:{aliasRange.End}, selection={selection.Start}:{selection.End}.");
        }
        finally
        {
            if (view is not null && restoreFieldCodes)
            {
                try { view.ShowFieldCodes = true; } catch { }
            }
            Release(view);
            Release(selection);
            Release(goTo);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(nativeRange);
            Release(aliasRange);
            Release(nativeBookmark);
            Release(aliasBookmark);
            Release(bookmarks);
        }
    }

    private static void RunWordOmmlMathTypeFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var pngPath = Path.Combine(artifactRoot, "omml-mathtype-format-conversion.png");
        var svgPath = Path.Combine(artifactRoot, "omml-mathtype-format-conversion.svg");
        WriteAcceptancePng(pngPath, "OMML↔MT", 360, 112);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><text x=\"8\" y=\"78\" font-family=\"Cambria Math\" font-size=\"46\">α + ℏ + ∫ + A</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 112);

        var sources = new[]
        {
            new OmmlMathTypeAcceptanceFormula(
                "inline-hbar-greek",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>ℏ</mi><mo>+</mo><mi>α</mi><mo>=</mo><mi>β</mi></mrow></math>",
                "inline",
                false),
            new OmmlMathTypeAcceptanceFormula(
                "display-fraction-integral",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mfrac><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msub><mi>y</mi><mn>1</mn></msub></mrow><mi>z</mi></mfrac><mo>+</mo><msubsup><mo>∫</mo><mn>0</mn><mi>∞</mi></msubsup><mi>f</mi><mo>⁡</mo><mfenced><mi>x</mi></mfenced></mrow></math>",
                "block",
                true),
            new OmmlMathTypeAcceptanceFormula(
                "display-matrix",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>A</mi><mo>=</mo><mfenced open=\"[\" close=\"]\"><mtable><mtr><mtd><mi>α</mi></mtd><mtd><mi>β</mi></mtd></mtr><mtr><mtd><mi>γ</mi></mtd><mtd><mi>δ</mi></mtd></mtr></mtable></mfenced></mrow></math>",
                "block",
                false),
            new OmmlMathTypeAcceptanceFormula(
                "inline-accents-vector",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mover accent=\"true\"><mi>v</mi><mo>→</mo></mover><mo>+</mo><mover accent=\"true\"><mi>x</mi><mo>¯</mo></mover><mo>+</mo><mover accent=\"true\"><mi>y</mi><mo>^</mo></mover></mrow></math>",
                "inline",
                false),
        };
        const string existingMathTypeMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>q</mi><mo>=</mo><mn>7</mn></math>";
        const string ommlSentinelMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>s</mi><mo>=</mo><mn>9</mn></math>";

        var mathTypeProcessesBefore = SnapshotMathTypeProcessIds();
        if (mathTypeProcessesBefore.Count > 0)
            Console.WriteLine(
                $"[OMML↔MATHTYPE CORE] Preserving {mathTypeProcessesBefore.Count} pre-existing MathType process(es); every stage still verifies that conversion starts no additional process.");

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            // Keep this structural conversion acceptance hidden. The test still
            // creates real Equation.DSMT4 objects and verifies their metafile
            // presentations, but a visible automation window can be captured by
            // Office's modal activation wizard before the first OMath is inserted.
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "OMML MathType format-conversion acceptance\r";
            var service = new WordFormulaService(application);

            for (var index = 0; index < sources.Length; index++)
                InsertManagedOmmlBetweenProse(document, service, sources[index]);

            SelectDocumentEnd(document);
            service.InsertMathTypeOle(
                CreateOmmlMathTypeAcceptanceSession(
                    existingMathTypeMathMl,
                    "inline",
                    false,
                    FormulaOleContract.MathTypeOleMode),
                existingMathTypeMathMl,
                emfPath);
            AppendAcceptanceText(document, " existing-mathtype\r");

            AssertEqual(4, document.OMaths.Count,
                "OMML→MathType setup did not create exactly four OMML source equations.");
            AssertEqual(1, CountMathTypeOleShapes(document),
                "OMML→MathType setup did not retain exactly one pre-existing MathType object.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "OMML→MathType setup lost the numbered OMML source state.");

            SelectFirstOmmlEquations(document, 2);
            var selectionPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(2, selectionPlan.Targets.Count,
                "OMML→MathType selection capture did not isolate exactly two OMML equations.");
            AssertTrue(selectionPlan.Targets.All(target => !string.IsNullOrWhiteSpace(target.SourceMathMl)),
                "OMML→MathType selection capture did not preserve canonical source MathML.");
            var selectionResult = service.ApplyFormulaFormatConversionPlan(
                selectionPlan,
                PrepareOmmlMathTypeTargets(selectionPlan, emfPath));
            AssertEqual(2, selectionResult.FormulaCount,
                "OMML→MathType selection conversion did not replace two equations. Failures: "
                + string.Join(" | ", selectionResult.Failures));
            AssertEqual(0, selectionResult.FailedFormulaCount,
                $"OMML→MathType selection conversion failed: {string.Join(" | ", selectionResult.Failures)}");
            AssertEqual(2, document.OMaths.Count,
                "OMML→MathType selection conversion changed the wrong OMML count.");
            AssertEqual(3, CountMathTypeOleShapes(document),
                "OMML→MathType selection conversion did not preserve the pre-existing MathType object.");

            var documentPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(2, documentPlan.Targets.Count,
                "OMML→MathType document capture did not find only the remaining two OMML equations.");
            var documentResult = service.ApplyFormulaFormatConversionPlan(
                documentPlan,
                PrepareOmmlMathTypeTargets(documentPlan, emfPath));
            AssertEqual(2, documentResult.FormulaCount,
                "OMML→MathType document conversion did not replace the remaining two equations. Failures: "
                + string.Join(" | ", documentResult.Failures));
            AssertEqual(0, documentResult.FailedFormulaCount,
                $"OMML→MathType document conversion failed: {string.Join(" | ", documentResult.Failures)}");
            AssertEqual(0, document.OMaths.Count,
                "OMML→MathType document conversion left OMML equations behind.");
            AssertEqual(5, CountMathTypeOleShapes(document),
                "OMML→MathType document conversion produced the wrong MathType count.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "OMML→MathType conversion did not recreate the numbered source as one MTPlaceRef field.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after OMML→MathType core conversion");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after OMML→MathType core conversion");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "OMML→MathType core conversion");

            var mathTypePath = Path.Combine(
                artifactRoot,
                "OMML-To-MathType-Core-Acceptance.docx");
            document.SaveAs2(mathTypePath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                mathTypePath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(5, CountMathTypeOleShapes(document),
                "Saved/reopened OMML→MathType document lost a MathType object.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Saved/reopened OMML→MathType document changed numbering.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after OMML→MathType save/reopen");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after OMML→MathType save/reopen");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "OMML→MathType save/reopen");

            service = new WordFormulaService(application);
            SelectDocumentEnd(document);
            service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    ommlSentinelMathMl,
                    "inline",
                    false,
                    FormulaOleContract.WordOmmlMode),
                ommlSentinelMathMl);
            AppendAcceptanceText(document, " existing-omml-sentinel\r");
            AssertEqual(1, document.OMaths.Count,
                "MathType→OMML mixed-document setup did not create the OMML sentinel.");

            SelectFirstMathTypeEquations(document, 2);
            var reverseSelectionPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(2, reverseSelectionPlan.Targets.Count,
                "MathType→OMML selection capture did not isolate exactly two MathType equations.");
            AssertTrue(reverseSelectionPlan.Targets.All(target => !string.IsNullOrWhiteSpace(target.SourceMathMl)),
                "MathType→OMML selection capture did not preserve Equation Native MathML.");
            var reverseSelectionResult = service.ApplyFormulaFormatConversionPlan(
                reverseSelectionPlan,
                PrepareOmmlMathTypeTargets(reverseSelectionPlan, emfPath));
            AssertEqual(2, reverseSelectionResult.FormulaCount,
                "MathType→OMML selection conversion did not replace two equations. Failures: "
                + string.Join(" | ", reverseSelectionResult.Failures));
            AssertEqual(0, reverseSelectionResult.FailedFormulaCount,
                $"MathType→OMML selection conversion failed: {string.Join(" | ", reverseSelectionResult.Failures)}");
            AssertEqual(3, CountMathTypeOleShapes(document),
                "MathType→OMML selection conversion removed the wrong MathType objects.");
            AssertEqual(3, document.OMaths.Count,
                "MathType→OMML selection conversion did not retain the pre-existing OMML sentinel.");

            var reverseDocumentPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(3, reverseDocumentPlan.Targets.Count,
                "MathType→OMML document capture did not find only the remaining MathType objects.");
            var reverseDocumentResult = service.ApplyFormulaFormatConversionPlan(
                reverseDocumentPlan,
                PrepareOmmlMathTypeTargets(reverseDocumentPlan, emfPath));
            AssertEqual(3, reverseDocumentResult.FormulaCount,
                "MathType→OMML document conversion did not replace the remaining MathType equations. Failures: "
                + string.Join(" | ", reverseDocumentResult.Failures));
            AssertEqual(0, reverseDocumentResult.FailedFormulaCount,
                $"MathType→OMML document conversion failed: {string.Join(" | ", reverseDocumentResult.Failures)}");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "MathType→OMML document conversion left MathType OLE objects behind.");
            AssertEqual(6, document.OMaths.Count,
                "MathType→OMML document conversion produced the wrong OMML count or removed the sentinel.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "MathType→OMML conversion did not restore the numbered formula state.");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "MathType→OMML core conversion");

            var finalPath = Path.Combine(
                artifactRoot,
                "MathType-To-OMML-Core-Acceptance.docx");
            document.SaveAs2(finalPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Saved/reopened MathType→OMML document restored a MathType OLE object.");
            AssertEqual(6, document.OMaths.Count,
                "Saved/reopened MathType→OMML document lost a native OMML equation.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "Saved/reopened MathType→OMML document changed the numbered OMML state.");
            AssertOmmlSemanticCoverage(document);
            AssertOmmlConversionProseSurvived(document, sources);

            service = new WordFormulaService(application);
            const string nativeToken = "VT_PURE_NATIVE_OMML_SOURCE";
            AppendAcceptanceText(
                document,
                $" before-pure-native-omml {nativeToken} after-pure-native-omml\r");
            Word.Range? nativeRange = null;
            try
            {
                nativeRange = InsertPureNativeOmml(document, nativeToken, "n+1");
                nativeRange.Select();
                var nativePlan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
                AssertEqual(1, nativePlan.Targets.Count,
                    "Pure Word-native OMath selection was not recognized as one OMML conversion source.");
                AssertTrue(!nativePlan.Targets[0].SourceIsManagedOmml,
                    "Pure Word-native OMath was incorrectly treated as VisualTeX-managed OMML.");
                var nativeResult = service.ApplyFormulaFormatConversionPlan(
                    nativePlan,
                    PrepareOmmlMathTypeTargets(nativePlan, emfPath));
                AssertEqual(1, nativeResult.FormulaCount,
                    "Pure Word-native OMath did not convert to MathType. Failures: "
                    + string.Join(" | ", nativeResult.Failures));
            }
            finally { Release(nativeRange); }
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Pure Word-native OMML conversion did not create exactly one Equation.DSMT4 object.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after pure native OMML→MathType");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after pure native OMML→MathType");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "pure native OMML→MathType");

            SelectFirstMathTypeEquations(document, 1);
            var nativeReversePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(1, nativeReversePlan.Targets.Count,
                "Pure native round-trip MathType source was not captured for OMML restoration.");
            var nativeReverseResult = service.ApplyFormulaFormatConversionPlan(
                nativeReversePlan,
                PrepareOmmlMathTypeTargets(nativeReversePlan, emfPath));
            AssertEqual(1, nativeReverseResult.FormulaCount,
                "Pure native MathType→OMML round-trip failed: "
                + string.Join(" | ", nativeReverseResult.Failures));
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Pure native OMML round-trip left a MathType object behind.");
            AssertEqual(7, document.OMaths.Count,
                "Pure native OMML round-trip produced the wrong final OMath count.");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(7, document.OMaths.Count,
                "Pure Word-native OMML round-trip did not survive save/reopen.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Pure Word-native OMML round-trip restored a MathType object after save/reopen.");
            AssertTrue((document.Content.Text ?? string.Empty).IndexOf(
                           "before-pure-native-omml",
                           StringComparison.Ordinal) >= 0
                       && (document.Content.Text ?? string.Empty).IndexOf(
                           "after-pure-native-omml",
                           StringComparison.Ordinal) >= 0,
                "Pure Word-native OMML round-trip damaged adjacent prose.");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "full OMML↔MathType acceptance");

            AssertExactLimitOmmlToMathTypeConversion(
                application,
                emfPath,
                mathTypeProcessesBefore);

            AssertMathTypeDisplayToOmmlPreservesParagraphStructure(
                application,
                document,
                emfPath);
            AssertMathTypeToOmmlPartialBatchFinalization(
                application,
                document,
                emfPath,
                artifactRoot);
            AssertVisualTeXOleToOmmlPreservesNumberFormat(
                application,
                document,
                artifactRoot);

            Console.WriteLine(
                "[OMML↔MATHTYPE CORE] Selection + document conversion in both directions passed with VisualTeX-managed and pure Word-native OMath sources, mixed source types, inline/display formulas, numbering, heading/section number-format preservation, display-paragraph preservation, hbar/Greek/fraction/integral/subscript/superscript/matrix/accent/vector semantics, adjacent prose preservation, non-empty MathType live previews, no OlePres, save/reopen persistence, and MathTypeProcessCount=0.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void AssertExactLimitOmmlToMathTypeConversion(
        Word.Application application,
        string emfPath,
        IReadOnlyCollection<int> mathTypeProcessesBefore)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mrow><munder><mo data-mjx-texclass=\"OP\" movablelimits=\"true\">lim</mo>"
            + "<mrow><mi>n</mi><mo>→</mo><mo>∞</mo></mrow></munder>"
            + "<msup><mfenced><mrow><mn>1</mn><mo>+</mo><mfrac><mn>1</mn><mi>n</mi></mfrac>"
            + "</mrow></mfenced><mi>n</mi></msup><mo>=</mo>"
            + "<mi mathvariant=\"normal\">e</mi></mrow></math>";
        Word.Document? document = null;
        Word.Range? sourceRange = null;
        try
        {
            document = application.Documents.Add();
            document.Content.Text = "before-exact-limit\r";
            var service = new WordFormulaService(application);
            var source = new OmmlMathTypeAcceptanceFormula(
                "exact-user-limit",
                mathMl,
                "block",
                false);
            InsertManagedOmmlBetweenProse(document, service, source);
            AssertEqual(1, document.OMaths.Count,
                "Exact user limit setup did not create one block OMML formula.");
            sourceRange = document.OMaths[1].Range.Duplicate;
            sourceRange.Select();

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Exact user limit OMML was not captured as one MathType conversion target.");
            AssertTrue(plan.Targets[0].Latex.IndexOf("\\lim", StringComparison.Ordinal) >= 0,
                "Exact user limit OMML capture lost its \\lim semantics: " + plan.Targets[0].Latex);
            var result = service.ApplyFormulaFormatConversionPlan(
                plan,
                PrepareOmmlMathTypeTargets(plan, emfPath));
            AssertEqual(0, result.FailedFormulaCount,
                "Exact user limit OMML→MathType conversion failed: "
                + string.Join(" | ", result.Failures));
            AssertEqual(1, result.FormulaCount,
                "Exact user limit OMML→MathType conversion did not convert one formula.");
            AssertEqual(0, document.OMaths.Count,
                "Exact user limit OMML→MathType conversion left the source OMath behind.");
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Exact user limit OMML→MathType conversion did not create one Equation.DSMT4 object.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "exact user limit OMML→MathType");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "exact user limit OMML→MathType");
            AssertNoNewMathTypeProcess(
                mathTypeProcessesBefore,
                "exact user limit OMML→MathType");
            Console.WriteLine(
                "[EXACT USER LIMIT OMML→MATHTYPE] \\lim_{n→∞}((1+1/n)^n)=e converted transactionally with no rollback.");
        }
        finally
        {
            Release(sourceRange);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void AssertMathTypeDisplayToOmmlPreservesParagraphStructure(
        Word.Application application,
        Word.Document returnDocument,
        string emfPath)
    {
        AssertMathTypeDisplayToOmmlPreservesParagraphStructure(
            application,
            returnDocument,
            emfPath,
            numbered: false);
        AssertMathTypeDisplayToOmmlPreservesParagraphStructure(
            application,
            returnDocument,
            emfPath,
            numbered: true);
    }

    private static void AssertMathTypeDisplayToOmmlPreservesParagraphStructure(
        Word.Application application,
        Word.Document returnDocument,
        string emfPath,
        bool numbered)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><mi>f</mi><mfenced><mi>x</mi></mfenced></math>";
        Word.Document? probe = null;
        Word.Paragraph? blankParagraph = null;
        Word.Range? insertion = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.OMath? convertedMath = null;
        Word.Range? convertedRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Paragraph? precedingParagraph = null;
        Word.Table? numberedTable = null;
        Word.Range? numberedTableRange = null;
        try
        {
            probe = application.Documents.Add();
            probe.Content.Text = "VT-DISPLAY-BEFORE\r\rVT-DISPLAY-AFTER\r";
            blankParagraph = probe.Paragraphs[2];
            insertion = blankParagraph.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertion.Select();

            var service = new WordFormulaService(application);
            service.InsertMathTypeOle(
                CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    "block",
                    numbered,
                    FormulaOleContract.MathTypeOleMode),
                mathMl,
                emfPath);
            AssertEqual(1, CountMathTypeOleShapes(probe),
                "Display-paragraph regression setup did not create one MathType equation.");
            shape = probe.InlineShapes[1];
            shapeRange = shape.Range.Duplicate;
            if (numbered)
                AssertTrue(
                    MathTypeOleInterop.TryReadDisplayNumberPosition(shape, out _),
                    "Numbered MathType→OMML regression setup did not create an MTPlaceRef number.");
            var paragraphCountBefore = probe.Paragraphs.Count;
            shapeRange.Select();

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(1, plan.Targets.Count,
                "Display-paragraph regression did not capture the MathType display equation.");
            var preparedTargets = PrepareOmmlMathTypeTargets(plan, emfPath);
            var result = service.ApplyFormulaFormatConversionPlan(
                plan,
                preparedTargets);
            AssertEqual(1, result.FormulaCount,
                "Display MathType→OMML regression failed: " + string.Join(" | ", result.Failures));
            AssertEqual(0, result.FailedFormulaCount,
                "Display MathType→OMML regression reported a failed formula.");
            AssertEqual(1, probe.OMaths.Count,
                "MathType→OMML display conversion did not retain exactly one OMath.");

            convertedMath = probe.OMaths[1];
            convertedRange = convertedMath.Range.Duplicate;
            paragraphs = probe.Paragraphs;

            if (numbered)
            {
                AssertEqual(1, probe.Tables.Count,
                    "Numbered MathType→OMML conversion did not create exactly one managed 1x3 numbering table.");
                AssertEqual(0, probe.Shapes.Count,
                    "Numbered MathType→OMML conversion created a floating number Shape.");
                AssertEqual(0, probe.Frames.Count,
                    "Numbered MathType→OMML conversion left a hidden caption Frame outside the 1x3 host.");
                var convertedFormulaId = preparedTargets[plan.Targets[0].Id]
                    .Session.FormulaId;
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    probe,
                    convertedFormulaId,
                    "numbered MathType→OMML direct-SEQ 1x3 host");

                numberedTable = probe.Tables[1];
                numberedTableRange = numberedTable.Range.Duplicate;
                Word.Range? precedingProbe = null;
                Word.Paragraphs? precedingParagraphs = null;
                Word.Range? precedingRange = null;
                try
                {
                    var precedingStart = Math.Max(
                        probe.Content.Start,
                        numberedTableRange.Start - 1);
                    precedingProbe = probe.Range(precedingStart, numberedTableRange.Start);
                    precedingParagraphs = precedingProbe.Paragraphs;
                    AssertTrue(precedingParagraphs.Count > 0,
                        "Converted numbered display OMML has no preceding body paragraph to validate.");
                    precedingParagraph = precedingParagraphs[1];
                    precedingRange = precedingParagraph.Range;
                    var precedingText = (precedingRange.Text ?? string.Empty)
                        .Trim('\r', '\a', '\v', '\f', '\t', ' ');
                    AssertEqual("VT-DISPLAY-BEFORE", precedingText,
                        "Numbered MathType→OMML inserted a visible blank paragraph above the managed 1x3 formula.");
                    Console.WriteLine(
                        $"[NUMBERED MATHTYPE DISPLAY→OMML 1X3] paragraphs={paragraphCountBefore}->{probe.Paragraphs.Count}; formula={convertedRange.Start}:{convertedRange.End}; preceding='{precedingText}'; tables={probe.Tables.Count}; frames={probe.Frames.Count}; shapes={probe.Shapes.Count}.");
                }
                finally
                {
                    Release(precedingRange);
                    Release(precedingParagraphs);
                    Release(precedingProbe);
                }
            }
            else
            {
                AssertEqual(0, probe.Tables.Count,
                    "Unnumbered MathType→OMML unexpectedly created a numbering table.");
                AssertEqual(paragraphCountBefore, probe.Paragraphs.Count,
                    "Unnumbered MathType→OMML display conversion changed the body paragraph count.");

                var formulaParagraphIndex = -1;
                for (var index = 1; index <= paragraphs.Count; index++)
                {
                    Release(formulaParagraph);
                    formulaParagraph = paragraphs[index];
                    Word.Range? range = null;
                    try
                    {
                        range = formulaParagraph.Range;
                        if (convertedRange.Start >= range.Start
                            && convertedRange.End <= range.End)
                        {
                            formulaParagraphIndex = index;
                            break;
                        }
                    }
                    finally { Release(range); }
                }
                AssertTrue(formulaParagraphIndex > 1,
                    "Converted unnumbered display OMML has no preceding body paragraph to validate.");
                precedingParagraph = paragraphs[formulaParagraphIndex - 1];
                Word.Range? precedingRange = null;
                try
                {
                    precedingRange = precedingParagraph.Range;
                    var precedingText = (precedingRange.Text ?? string.Empty)
                        .Trim('\r', '\a', '\v', '\f', '\t', ' ');
                    AssertEqual("VT-DISPLAY-BEFORE", precedingText,
                        "Unnumbered MathType→OMML inserted a blank paragraph above the display equation.");
                    Console.WriteLine(
                        $"[UNNUMBERED MATHTYPE DISPLAY→OMML PARAGRAPH] paragraphs={paragraphCountBefore}->{probe.Paragraphs.Count}; preceding='{precedingText}'.");
                }
                finally { Release(precedingRange); }
            }
        }
        finally
        {
            Release(numberedTableRange);
            Release(numberedTable);
            Release(precedingParagraph);
            Release(formulaParagraph);
            Release(paragraphs);
            Release(convertedRange);
            Release(convertedMath);
            Release(shapeRange);
            Release(shape);
            Release(insertion);
            Release(blankParagraph);
            if (probe is not null)
            {
                try { probe.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(probe);
            try { returnDocument.Activate(); } catch { }
        }
    }

    private static void AssertMathTypeToOmmlPartialBatchFinalization(
        Word.Application application,
        Word.Document returnDocument,
        string emfPath,
        string artifactRoot)
    {
        const string firstMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";
        const string thirdMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>";

        var previousAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_ACCEPTANCE");
        var previousFailure = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
        Word.Document? probe = null;
        Word.Bookmark? convertedBookmark = null;
        Word.Range? convertedRange = null;
        try
        {
            probe = application.Documents.Add();
            probe.Content.Text =
                "partial-before\r\rpartial-between-one-two\r\rpartial-between-two-three\r\rpartial-after\r";
            var service = new WordFormulaService(application);

            void InsertAtBlankParagraph(
                int paragraphIndex,
                string mathMl,
                bool numbered)
            {
                Word.Paragraph? paragraph = null;
                Word.Range? insertion = null;
                try
                {
                    paragraph = probe.Paragraphs[paragraphIndex];
                    insertion = paragraph.Range.Duplicate;
                    insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    insertion.Select();
                    service.InsertMathTypeOle(
                        CreateOmmlMathTypeAcceptanceSession(
                            mathMl,
                            "block",
                            numbered,
                            FormulaOleContract.MathTypeOleMode),
                        mathMl,
                        emfPath,
                        preserveExistingDisplayParagraphBoundary: true);
                }
                finally
                {
                    Release(insertion);
                    Release(paragraph);
                }
            }

            // Insert from the end so earlier blank-paragraph indexes remain stable.
            InsertAtBlankParagraph(6, thirdMathMl, numbered: false);
            InsertAtBlankParagraph(4, secondMathMl, numbered: false);
            InsertAtBlankParagraph(2, firstMathMl, numbered: true);
            AssertEqual(3, CountMathTypeOleShapes(probe),
                "Partial MathType→OMML setup did not create exactly three MathType sources.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            var ordered = plan.Targets.OrderBy(target => target.SourceStart).ToArray();
            AssertEqual(3, ordered.Length,
                "Partial MathType→OMML setup did not capture three sources.");
            AssertTrue(ordered[0].Numbered,
                "Partial MathType→OMML setup did not preserve the first source as numbered display MathType.");
            var prepared = PrepareOmmlMathTypeTargets(plan, emfPath);

            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                ordered[1].SourceFormulaId);
            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);

            AssertEqual(1, result.FormulaCount,
                "Partial MathType→OMML conversion did not retain exactly the first committed target. Failures: "
                + string.Join(" | ", result.Failures));
            AssertEqual(1, result.FailedFormulaCount,
                "Partial MathType→OMML conversion should report only the injected second-item failure. Failures: "
                + string.Join(" | ", result.Failures));
            AssertEqual(2, CountMathTypeOleShapes(probe),
                "Partial MathType→OMML conversion did not preserve the failed and unprocessed MathType sources.");
            AssertEqual(1, probe.OMaths.Count,
                "Partial MathType→OMML conversion did not retain exactly one committed OMath.");
            var failureText = string.Join(" | ", result.Failures);
            AssertTrue(
                failureText.IndexOf(
                    "Injected format-conversion failure after deleting the source host.",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                "Partial MathType→OMML conversion lost the primary injected failure: " + failureText);
            AssertTrue(
                failureText.IndexOf("bookmark drifted", StringComparison.OrdinalIgnoreCase) < 0
                && failureText.IndexOf("could not be recovered uniquely", StringComparison.OrdinalIgnoreCase) < 0,
                "Partial MathType→OMML finalization replaced the primary failure with a stale OMML identity error: "
                + failureText);

            var firstConvertedFormulaId = prepared[ordered[0].Id].Session.FormulaId;
            var convertedMetadata = WordOmmlFormulaStore.TryRead(
                    probe,
                    firstConvertedFormulaId)
                ?? throw new InvalidDataException(
                    "Partial MathType→OMML conversion lost the committed target metadata.");
            convertedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    probe,
                    firstConvertedFormulaId)
                ?? throw new InvalidDataException(
                    "Partial MathType→OMML conversion lost the committed VTOMML bookmark.");
            convertedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                probe,
                firstConvertedFormulaId,
                convertedMetadata);
            AssertTrue(
                WordEquationNumbering.HasReusableNumberedNativeOmmlDirectTableHost(
                    probe,
                    convertedRange,
                    firstConvertedFormulaId),
                "The successfully committed numbered target in a partial batch was not finalized as the required 1x3 direct-SEQ host.");
            var liveFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                convertedRange.WordOpenXML);
            AssertEqual(
                liveFingerprint,
                convertedMetadata.NativeOmmlFingerprint ?? string.Empty,
                "The successfully committed OMML target retained a provisional/stale fingerprint after a later item failed.");

            var remainingPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(2, remainingPlan.Targets.Count,
                "Partial MathType→OMML conversion did not leave exactly the failed and unprocessed MathType sources.");
            var remainingSignatures = remainingPlan.Targets
                .Select(target => MathTypeMtefCodec.SemanticSignature(
                    target.SourceMathMl
                    ?? throw new InvalidDataException("A remaining MathType source lost its MathML.")))
                .OrderBy(signature => signature, StringComparer.Ordinal)
                .ToArray();
            var expectedRemainingSignatures = new[]
                {
                    MathTypeMtefCodec.SemanticSignature(secondMathMl),
                    MathTypeMtefCodec.SemanticSignature(thirdMathMl),
                }
                .OrderBy(signature => signature, StringComparer.Ordinal)
                .ToArray();
            AssertEqual(
                string.Join("\n", expectedRemainingSignatures),
                string.Join("\n", remainingSignatures),
                "Partial MathType→OMML rollback changed the failed or unprocessed MathType source semantics.");

            var outputPath = Path.Combine(
                artifactRoot,
                "MathType-To-OMML-Partial-Batch-Finalization.docx");
            probe.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "[PARTIAL MT→OMML] First numbered target finalized as 1x3 direct-SEQ with a live fingerprint; "
                + "the injected second target rolled back, the third remained untouched, and the primary failure was preserved.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                previousFailure);
            Release(convertedRange);
            Release(convertedBookmark);
            if (probe is not null)
            {
                try { probe.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(probe);
            try { returnDocument.Activate(); } catch { }
        }
    }

    private static void AssertVisualTeXOleToOmmlPreservesNumberFormat(
        Word.Application application,
        Word.Document returnDocument,
        string artifactRoot)
    {
        // Use a previously generated, saved VisualTeX OLE document instead of
        // creating a fresh native OLE inside this isolated acceptance process.
        // The latter intentionally suppresses the installed OLE server and can be
        // denied by COM before format conversion starts; opening the saved fixture
        // exercises the real existing-document path reported by users.
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup><mo>=</mo><msup><mi>z</mi><mn>2</mn></msup></math>";
        var fixturePath = Path.GetFullPath(Path.Combine(
            "artifacts",
            "issue11-probe",
            "issue11-heading1-ole.docx"));
        if (!File.Exists(fixturePath))
            throw new FileNotFoundException(
                "The saved numbered VisualTeX OLE fixture is required for number-format conversion acceptance.",
                fixturePath);
        var previousDefaultFormat = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        var workingPath = Path.Combine(
            artifactRoot,
            "VisualTeX-OLE-To-OMML-Heading2-Number-Format.docx");
        File.Copy(fixturePath, workingPath, overwrite: true);

        Word.Document? probe = null;
        Word.Document? reopened = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            probe = application.Documents.Open(
                workingPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            var service = new WordFormulaService(application);
            service.SetEquationNumberFormat(EquationNumberFormat.Heading2DotId);
            AssertEqual(
                EquationNumberFormat.Heading2DotId,
                service.GetEquationNumberFormatId(),
                "Number-format preservation fixture did not enter heading2-dot mode.");

            for (var index = 1; index <= probe.InlineShapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = probe.InlineShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(candidate);
                    if (metadata?.Numbered != true
                        || !string.Equals(
                            metadata.DisplayMode,
                            "block",
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    shape = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            AssertTrue(shape is not null,
                "Saved fixture has no numbered VisualTeX OLE display equation to convert.");
            var visualTeXOleCountBefore = CountVisualTeXNativeOleShapes(probe);
            shapeRange = shape!.Range.Duplicate;
            shapeRange.Select();
            var selected = service.ReadSelection();
            var sourceMetadata = selected.Metadata
                ?? throw new InvalidDataException(
                    "The saved VisualTeX OLE fixture lost its metadata before OMML conversion.");
            AssertTrue(sourceMetadata.Numbered,
                "The saved VisualTeX OLE fixture lost its numbered state before conversion.");
            AssertEqual("block", sourceMetadata.DisplayMode,
                "The number-format fixture formula is no longer a display equation.");
            var convertedFormulaId = selected.FormulaId ?? sourceMetadata.FormulaId;
            var replacementLatex = MathMlToLatexConverter.Convert(mathMl).Trim();
            var lineId = sourceMetadata.Lines.FirstOrDefault()?.Id;
            if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString("D");
            var targetSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "edit",
                Host = "word",
                FormulaId = convertedFormulaId,
                SourceDocumentId = selected.DocumentId,
                SourceObjectId = selected.ObjectId,
                Title = "VisualTeX OLE to OMML number-format acceptance",
                CodeFormat = "latex",
                DisplayMode = sourceMetadata.DisplayMode,
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = sourceMetadata.Numbered,
                FontSizePt = sourceMetadata.FontSizePt ?? 12,
                OriginalMetadata = sourceMetadata,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId!, Latex = replacementLatex },
                },
                ExportResult = new OfficeExportDocument
                {
                    MathMl = mathMl,
                    Width = 260,
                    Height = 96,
                    Baseline = 72,
                },
            };
            service.ReplaceOmml(targetSession, mathMl);
            AssertEqual(
                visualTeXOleCountBefore - 1,
                CountVisualTeXNativeOleShapes(probe),
                "VisualTeX OLE→OMML number-format regression did not remove exactly the selected VisualTeX OLE.");
            Word.Bookmark? convertedBookmark = null;
            try
            {
                convertedBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    probe,
                    convertedFormulaId);
                AssertTrue(convertedBookmark is not null,
                    "VisualTeX OLE→OMML number-format regression did not retain the converted formula identity as OMML.");
            }
            finally { Release(convertedBookmark); }
            AssertEqual(
                EquationNumberFormat.Heading2DotId,
                service.GetEquationNumberFormatId(),
                "VisualTeX OLE→OMML fell back from section numbering to continuous numbering.");
            var visibleTarget = WordEquationNumbering.GetEquationReferenceTargets(probe)
                .Single(item => string.Equals(
                    item.FormulaId,
                    convertedFormulaId,
                    StringComparison.OrdinalIgnoreCase));
            AssertEqual(2, visibleTarget.NumberText.Count(character => character == '.'),
                "Converted OMML visible number no longer has the section-numbering shape a.b.c.");
            var numberBeforeReopen = visibleTarget.NumberText;

            probe.Save();
            probe.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(probe);
            probe = null;
            reopened = application.Documents.Open(
                workingPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            var reopenedService = new WordFormulaService(application);
            AssertEqual(
                EquationNumberFormat.Heading2DotId,
                reopenedService.GetEquationNumberFormatId(),
                "Saved/reopened converted OMML lost the section-numbering format.");
            var reopenedTarget = WordEquationNumbering.GetEquationReferenceTargets(reopened)
                .Single(item => string.Equals(
                    item.FormulaId,
                    convertedFormulaId,
                    StringComparison.OrdinalIgnoreCase));
            AssertEqual(numberBeforeReopen, reopenedTarget.NumberText,
                "Saved/reopened converted OMML changed its section-style equation number.");
            Console.WriteLine(
                $"[VISUALTEX OLE→OMML NUMBER FORMAT] format={EquationNumberFormat.Heading2DotId}; number={reopenedTarget.NumberText}; save/reopen preserved.");
        }
        finally
        {
            if (reopened is not null)
            {
                try { reopened.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(reopened);
            Release(shapeRange);
            Release(shape);
            if (probe is not null)
            {
                try { probe.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(probe);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previousDefaultFormat);
            try { returnDocument.Activate(); } catch { }
        }
    }

    private sealed class OmmlMathTypeAcceptanceFormula
    {
        internal OmmlMathTypeAcceptanceFormula(
            string name,
            string mathMl,
            string displayMode,
            bool numbered)
        {
            Name = name;
            MathMl = mathMl;
            DisplayMode = displayMode;
            Numbered = numbered;
        }

        internal string Name { get; }
        internal string MathMl { get; }
        internal string DisplayMode { get; }
        internal bool Numbered { get; }
    }

    private static OfficeSessionDocument CreateOmmlMathTypeAcceptanceSession(
        string mathMl,
        string displayMode,
        bool numbered,
        string objectMode)
    {
        var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "OMML MathType conversion acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = objectMode,
            Numbered = numbered,
            MathTypeNumberPosition = "right",
            FontSizePt = 12,
            ExportResult = new OfficeExportDocument
            {
                MathMl = mathMl,
                Width = 360,
                Height = 112,
                Baseline = 82,
            },
        };
    }

    private static IReadOnlyDictionary<string, PreparedWordBulkFormula> PrepareOmmlMathTypeTargets(
        WordFormulaFormatConversionPlan plan,
        string emfPath)
    {
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        foreach (var target in plan.Targets)
        {
            var mathMl = target.SourceMathMl
                ?? throw new InvalidDataException(
                    $"Format-conversion target '{target.Latex}' has no canonical source MathML.");
            prepared[target.Id] = new PreparedWordBulkFormula
            {
                Run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                },
                Session = CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    target.DisplayMode,
                    target.Numbered,
                    plan.TargetMode),
                MathMl = mathMl,
                EmfPath = string.Equals(
                    plan.TargetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                    ? emfPath
                    : null,
            };
        }
        return prepared;
    }

    private static void SelectDocumentEnd(Word.Document document)
    {
        Word.Range? range = null;
        for (var attempt = 0; attempt < 60; attempt++)
        {
            try
            {
                Release(range);
                range = null;
                var end = document.Content.End - 1;
                range = document.Range(end, end);
                range.Select();
                Release(range);
                return;
            }
            catch (System.Runtime.InteropServices.COMException error)
                when (error.HResult == unchecked((int)0x80010001) && attempt < 59)
            {
                Release(range);
                range = null;
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(150);
            }
            finally
            {
                if (attempt == 59)
                    Release(range);
            }
        }
    }

    private static void AppendAcceptanceText(Word.Document document, string text)
    {
        Word.Range? range = null;
        try
        {
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = text;
        }
        finally { Release(range); }
    }

    private static void SelectFirstOmmlEquations(Word.Document document, int count)
    {
        Word.OMaths? maths = null;
        Word.OMath? first = null;
        Word.OMath? last = null;
        Word.Range? firstRange = null;
        Word.Range? lastRange = null;
        Word.Range? selection = null;
        try
        {
            maths = document.OMaths;
            if (maths.Count < count)
                throw new InvalidDataException(
                    $"Expected at least {count} OMML equations, actual {maths.Count}.");
            first = maths[1];
            last = maths[count];
            firstRange = first.Range;
            lastRange = last.Range;
            selection = document.Range(firstRange.Start, lastRange.End);
            selection.Select();
        }
        finally
        {
            Release(selection);
            Release(lastRange);
            Release(firstRange);
            Release(last);
            Release(first);
            Release(maths);
        }
    }

    private static void SelectFirstMathTypeEquations(Word.Document document, int count)
    {
        var starts = new List<(int Start, int End)>();
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                range = shape.Range;
                starts.Add((range.Start, range.End));
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }
        starts.Sort((left, right) => left.Start.CompareTo(right.Start));
        if (starts.Count < count)
            throw new InvalidDataException(
                $"Expected at least {count} MathType equations, actual {starts.Count}.");
        Word.Range? selection = null;
        try
        {
            selection = document.Range(starts[0].Start, starts[count - 1].End);
            selection.Select();
        }
        finally { Release(selection); }
    }

    private static int CountManagedNumberedOmml(Word.Document document)
    {
        var count = 0;
        foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
        {
            Word.Bookmark? bookmark = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (metadata?.Numbered == true) count++;
            }
            finally { Release(bookmark); }
        }
        return count;
    }

    private static void AssertEveryMathTypeProgId(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                count++;
                AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                    $"MathType conversion output #{count} lost ProgID=Equation.DSMT4.");
                var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                AssertTrue(!string.IsNullOrWhiteSpace(mathMl),
                    $"MathType conversion output #{count} has no readable Equation Native MathML.");
            }
            finally { Release(shape); }
        }
        AssertTrue(count > 0, "No MathType conversion outputs were available for ProgID validation.");
    }

    private static void AssertNoNewMathTypeProcess(
        IReadOnlyCollection<int> baseline,
        string stage)
    {
        var started = SnapshotMathTypeProcessIds().Except(baseline).ToArray();
        AssertEqual(0, started.Length,
            $"MathType.exe started during {stage}: {string.Join(", ", started)}");
    }

    private static void InsertManagedOmmlBetweenProse(
        Word.Document document,
        WordFormulaService service,
        OmmlMathTypeAcceptanceFormula source)
    {
        var token = "VT_OMML_SOURCE_" + Guid.NewGuid().ToString("N");
        var surroundingText = string.Equals(
                source.DisplayMode,
                "block",
                StringComparison.Ordinal)
            ? $"before-{source.Name}\r{token}\r after-{source.Name}\r"
            : $"before-{source.Name} {token} after-{source.Name}\r";
        AppendAcceptanceText(document, surroundingText);
        Word.Range? insertion = null;
        try
        {
            insertion = FindAcceptanceTextRange(document, token);
            insertion.Text = string.Empty;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertion.Select();
            service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    source.MathMl,
                    source.DisplayMode,
                    source.Numbered,
                    FormulaOleContract.WordOmmlMode),
                source.MathMl);
        }
        finally { Release(insertion); }
    }

    private static Word.Range FindAcceptanceTextRange(
        Word.Document document,
        string text)
    {
        Word.Range? range = null;
        Word.Find? find = null;
        try
        {
            range = document.Content.Duplicate;
            find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            if (!find.Execute())
                throw new InvalidDataException(
                    $"Acceptance placeholder '{text}' was not found in the Word document.");
            var result = range.Duplicate;
            Release(range);
            range = null;
            return result;
        }
        finally
        {
            Release(find);
            Release(range);
        }
    }

    private static void AssertVisibleMathTypePreviewsWithClipboardRetry(
        Word.Document document,
        string stage)
    {
        Exception? last = null;
        for (var attempt = 1; attempt <= 8; attempt++)
        {
            try
            {
                AssertVisibleMathTypeEmfPreviews(document, stage);
                return;
            }
            catch (InvalidDataException error)
            {
                last = error;
                if (error.ToString().IndexOf(
                        "CLIPBRD_E_CANT_OPEN",
                        StringComparison.OrdinalIgnoreCase) < 0
                    && error.ToString().IndexOf(
                        "clipboard stayed busy",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    throw;
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(400 * attempt);
            }
        }
        throw new InvalidDataException(
            $"Word clipboard remained busy after repeated live-preview validation attempts {stage}.",
            last);
    }

    private static Word.Range InsertPureNativeOmml(
        Word.Document document,
        string placeholder,
        string linearText)
    {
        Word.Range? source = null;
        Word.Range? added = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? result = null;
        try
        {
            source = FindAcceptanceTextRange(document, placeholder);
            var start = source.Start;
            source.Text = linearText;
            source.SetRange(start, start + linearText.Length);
            added = document.OMaths.Add(source);
            maths = added.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "Word did not create exactly one native OMath from the linear source range.");
            math = maths[1];
            math.BuildUp();
            result = math.Range.Duplicate;
            var duplicate = result;
            result = null;
            return duplicate;
        }
        finally
        {
            Release(result);
            Release(math);
            Release(maths);
            Release(added);
            Release(source);
        }
    }

    private static void AssertOmmlConversionProseSurvived(
        Word.Document document,
        IReadOnlyList<OmmlMathTypeAcceptanceFormula> sources)
    {
        var text = document.Content.Text ?? string.Empty;
        var previous = -1;
        foreach (var source in sources)
        {
            var before = "before-" + source.Name;
            var after = "after-" + source.Name;
            var beforeIndex = text.IndexOf(before, StringComparison.Ordinal);
            var afterIndex = text.IndexOf(after, StringComparison.Ordinal);
            AssertTrue(beforeIndex >= 0 && afterIndex > beforeIndex,
                $"OMML↔MathType conversion damaged adjacent prose around '{source.Name}'.");
            AssertTrue(beforeIndex > previous,
                $"OMML↔MathType conversion changed the document order near '{source.Name}'.");
            previous = afterIndex;
        }
    }

    private static void AssertOmmlSemanticCoverage(Word.Document document)
    {
        var latex = new List<string>();
        Word.OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range.Duplicate;
                    var metadata = WordOmmlNativeSource.CreateForNative(document, range);
                    latex.Add((metadata.Latex ?? string.Empty).Replace(" ", string.Empty));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
        }
        finally { Release(maths); }

        var joined = string.Join(" | ", latex);
        AssertTrue(joined.IndexOf("hbar", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("ℏ", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost hbar: {joined}");
        AssertTrue(joined.IndexOf("alpha", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("α", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost Greek alpha: {joined}");
        AssertTrue(joined.IndexOf("frac", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("/", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost fraction structure: {joined}");
        AssertTrue(joined.IndexOf("int", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("∫", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost integral structure: {joined}");
        AssertTrue(joined.IndexOf("matrix", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("begin", StringComparison.OrdinalIgnoreCase) >= 0,
            $"Final OMML semantics lost matrix structure: {joined}");
        AssertTrue(joined.IndexOf("vec", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("over", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("→", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost vector/accent structure: {joined}");
    }
}

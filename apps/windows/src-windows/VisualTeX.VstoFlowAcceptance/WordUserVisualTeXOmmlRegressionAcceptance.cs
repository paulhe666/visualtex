using System.Diagnostics;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunUserVisualTeXOmmlRegressionAcceptance(string artifactRoot)
    {
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER_VT_OMML_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "VISUALTEX_USER_VT_OMML_SOURCE must point to the user's VisualTeX Word document.",
                sourcePath);

        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(artifactRoot, "user-visualtex-to-omml-repro.docx");
        var tracePath = Path.Combine(artifactRoot, "user-visualtex-to-omml-repro.trace.log");
        File.Copy(Path.GetFullPath(sourcePath), outputPath, overwrite: true);
        try { File.Delete(tracePath); } catch { }

        var previousFormatAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousTracePath =
            Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousTracePerf =
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        Exception? conversionError = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF", "1");

            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            WriteUserVisualTeXOmmlSnapshot(
                document,
                Path.Combine(artifactRoot, "before"));
            var visualTeXBefore = CountInstalledVisualTeXOleShapes(document);
            var ommlBefore = document.OMaths.Count;
            Console.WriteLine(
                $"[USER VT→OMML BEFORE] visualTeX={visualTeXBefore}; omml={ommlBefore}; "
                + $"paragraphs={document.Paragraphs.Count}; tables={document.Tables.Count}; "
                + $"bookmarks={document.Bookmarks.Count}; fields={document.Fields.Count}");

            var service = new WordFormulaService(application);
            var planMetrics = new WordOperationMetrics();
            WordFormulaFormatConversionPlan plan;
            var planWatch = Stopwatch.StartNew();
            using (planMetrics)
            {
                plan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.NativeOleMode,
                    FormulaOleContract.WordOmmlMode);
            }
            planWatch.Stop();
            File.WriteAllLines(
                Path.Combine(artifactRoot, "before-plan-targets.txt"),
                plan.Targets
                    .OrderBy(target => target.SourceStart)
                    .Select((target, index) =>
                        $"{index + 1}|formulaId={target.SourceFormulaId}|range={target.SourceObjectId}"
                        + $"|start={target.SourceStart}|display={target.DisplayMode}|numbered={target.Numbered}"
                        + $"|table={target.SourceWithinTable}|font={target.FontSizePt:0.###}"
                        + $"|latex={target.Latex.Replace("\r", "<CR>").Replace("\n", "<LF>")}"));
            File.WriteAllLines(
                Path.Combine(artifactRoot, "before-plan-metrics.txt"),
                new[]
                {
                    $"captureMs={planWatch.Elapsed.TotalMilliseconds:0.###}",
                    $"targets={plan.Targets.Count}",
                }.Concat(
                    planMetrics.Entries
                        .OrderByDescending(pair => pair.Value.Milliseconds)
                        .Select(pair =>
                            $"{pair.Key}|calls={pair.Value.Calls}|ms={pair.Value.Milliseconds:0.###}")));
            Console.WriteLine(
                $"[USER VT→OMML PLAN] targets={plan.Targets.Count}; captureMs={planWatch.Elapsed.TotalMilliseconds:0.###}");

            addIn = new ThisAddIn();
            addIn.OnConnection(
                application,
                Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            ResetInstalledFormatConversionTrace(tracePath);
            var conversionWatch = Stopwatch.StartNew();
            try
            {
                addIn.OnConvertVisualTeXToOmmlDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    "source=VisualTeX target=OMML",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(180));
            }
            catch (Exception error)
            {
                conversionError = error;
            }
            conversionWatch.Stop();

            WriteUserVisualTeXOmmlSnapshot(
                document,
                Path.Combine(artifactRoot, "after"));
            Console.WriteLine(
                $"[USER VT→OMML AFTER] visualTeX={CountInstalledVisualTeXOleShapes(document)}; "
                + $"omml={document.OMaths.Count}; paragraphs={document.Paragraphs.Count}; "
                + $"tables={document.Tables.Count}; bookmarks={document.Bookmarks.Count}; "
                + $"fields={document.Fields.Count}; conversionMs={conversionWatch.Elapsed.TotalMilliseconds:0.###}");

            if (File.Exists(tracePath))
            {
                var trace = File.ReadAllText(tracePath);
                var interesting = trace
                    .Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(line =>
                        line.IndexOf("format-conversion-", StringComparison.OrdinalIgnoreCase) >= 0
                        || line.IndexOf("omml-", StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToArray();
                File.WriteAllLines(
                    Path.Combine(artifactRoot, "conversion-trace-focused.txt"),
                    interesting);
            }

            if (conversionError is not null)
                throw new InvalidOperationException(
                    "The r66 VisualTeX→OMML production-path conversion reproduced a failure. "
                    + $"See artifacts at '{artifactRoot}'.",
                    conversionError);

            AssertEqual(0, CountInstalledVisualTeXOleShapes(document),
                "r66 VisualTeX→OMML left VisualTeX OLE sources behind.");
            AssertEqual(ommlBefore + visualTeXBefore, document.OMaths.Count,
                "r66 VisualTeX→OMML did not create one OMML equation per VisualTeX source.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            WriteUserVisualTeXOmmlSnapshot(
                reopened,
                Path.Combine(artifactRoot, "reopened"));
            AssertEqual(0, CountInstalledVisualTeXOleShapes(reopened),
                "Saved/reopened r66 conversion restored or retained VisualTeX OLE sources.");
            AssertEqual(ommlBefore + visualTeXBefore, reopened.OMaths.Count,
                "Saved/reopened r66 conversion changed the OMML equation count.");
            Console.WriteLine(
                $"[USER VT→OMML PASS] converted={visualTeXBefore}; "
                + $"omml={reopened.OMaths.Count}; output={outputPath}");
        }
        finally
        {
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();

            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_TRACE_FORMAT_PERF",
                previousTracePerf);
        }
    }

    private static void RunUserOmmlVisualTeXRegressionAcceptance(string artifactRoot)
    {
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER_OMML_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "VISUALTEX_USER_OMML_SOURCE must point to the user's OMML Word document.",
                sourcePath);

        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(artifactRoot, "user-omml-to-visualtex-repro.docx");
        var tracePath = Path.Combine(artifactRoot, "user-omml-to-visualtex-repro.trace.log");
        var fixtureCountSetting = Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_OMML_FIXTURE_COUNT");
        var fixtureCount = string.IsNullOrWhiteSpace(fixtureCountSetting) ? 0
            : int.TryParse(fixtureCountSetting, out var requestedCount) && requestedCount >= 0 && requestedCount <= 100
                ? requestedCount : throw new InvalidDataException("OMML fixture count must be between 0 and 100.");
        CopyNumberedOmmlPerformanceFixture(Path.GetFullPath(sourcePath), outputPath, fixtureCount);
        try { File.Delete(tracePath); } catch { }

        var previousFormatAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousAcceptance =
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousTracePath =
            Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var sourceOmmlCount = document.OMaths.Count;
            var sourceVisualTeXCount = CountInstalledVisualTeXOleShapes(document);
            var service = new WordFormulaService(application);
            var metrics = new WordOperationMetrics();
            WordFormulaFormatConversionPlan plan;
            var scanWatch = Stopwatch.StartNew();
            using (metrics)
            {
                plan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
            }
            scanWatch.Stop();
            File.WriteAllLines(
                Path.Combine(artifactRoot, "scan-metrics.txt"),
                new[]
                {
                    $"scanMs={scanWatch.Elapsed.TotalMilliseconds:0.###}",
                    $"targets={plan.Targets.Count}",
                }.Concat(
                    metrics.Entries
                        .OrderByDescending(pair => pair.Value.Milliseconds)
                        .ThenBy(pair => pair.Key, StringComparer.Ordinal)
                        .Select(pair =>
                            $"{pair.Key}|calls={pair.Value.Calls}|ms={pair.Value.Milliseconds:0.###}")));

            Console.WriteLine(
                $"[USER OMML→VT BEFORE] omml={sourceOmmlCount}; visualTeX={sourceVisualTeXCount}; "
                + $"targets={plan.Targets.Count}; scanMs={scanWatch.Elapsed.TotalMilliseconds:0.###}");
            if (sourceOmmlCount <= 0 || plan.Targets.Count != sourceOmmlCount)
                throw new InvalidDataException(
                    $"OMML→VisualTeX fixture inventory mismatch: omml={sourceOmmlCount}, targets={plan.Targets.Count}.");

            addIn = new ThisAddIn();
            addIn.OnConnection(
                application,
                Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            ResetInstalledFormatConversionTrace(tracePath);
            var conversionWatch = Stopwatch.StartNew();
            addIn.OnConvertOmmlToVisualTeXDocument(new object());
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=VisualTeX",
                mathTypeBaseline);
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(180));
            conversionWatch.Stop();

            var finalOmmlCount = document.OMaths.Count;
            var finalVisualTeXCount = CountInstalledVisualTeXOleShapes(document);
            Console.WriteLine(
                $"[USER OMML→VT AFTER] omml={finalOmmlCount}; visualTeX={finalVisualTeXCount}; "
                + $"conversionMs={conversionWatch.Elapsed.TotalMilliseconds:0.###}");
            AssertEqual(0, finalOmmlCount,
                "100 OMML→VisualTeX conversion left OMML sources behind.");
            AssertEqual(sourceVisualTeXCount + sourceOmmlCount, finalVisualTeXCount,
                "100 OMML→VisualTeX conversion did not create one VisualTeX formula per OMML source.");

            var expectedNumbered = plan.Targets.Count(target => target.Numbered);
            AssertEqual(expectedNumbered, CountVisualTeXNumberingBookmarkTriples(document),
                "OMML→VisualTeX conversion lost numbered identities.");
            var persistenceWatch = Stopwatch.StartNew();
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            persistenceWatch.Stop();
            Console.WriteLine($"[OMML→VT PERSISTENCE] saveReopenMs={persistenceWatch.Elapsed.TotalMilliseconds:0.###}");
            AssertEqual(expectedNumbered, CountVisualTeXNumberingBookmarkTriples(reopened),
                "Saved/reopened OMML→VisualTeX result lost numbered identities.");
            AssertEqual(0, reopened.OMaths.Count,
                "Saved/reopened 100 OMML→VisualTeX result regained OMML sources.");
            AssertEqual(sourceVisualTeXCount + sourceOmmlCount,
                CountInstalledVisualTeXOleShapes(reopened),
                "Saved/reopened 100 OMML→VisualTeX result changed the VisualTeX formula count.");

            if (Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_OMML_FIXTURE_ROUNDTRIP") == "1")
            {
                reopened.Activate();
                var returnWatch = Stopwatch.StartNew();
                addIn.OnConvertVisualTeXToOmmlDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(tracePath, "source=VisualTeX target=OMML", mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(180));
                returnWatch.Stop();
                AssertEqual(sourceOmmlCount, reopened.OMaths.Count, "Numbered fixture return leg lost an OMath.");
                AssertEqual(sourceVisualTeXCount, CountInstalledVisualTeXOleShapes(reopened), "Numbered fixture return leg left converted OLEs.");
                AssertEqual(expectedNumbered, CountVisualTeXNumberingBookmarkTriples(reopened), "Numbered fixture return leg lost numbering.");
                reopened.Save();
                reopened.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(reopened);
                reopened = application.Documents.Open(outputPath, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                AssertEqual(sourceOmmlCount, reopened.OMaths.Count, "Saved/reopened return leg lost an OMath.");
                AssertEqual(expectedNumbered, CountVisualTeXNumberingBookmarkTriples(reopened), "Saved/reopened return leg lost numbering.");
                Console.WriteLine($"[OMML FIXTURE ROUNDTRIP PASS] formulas={sourceOmmlCount}; returnMs={returnWatch.Elapsed.TotalMilliseconds:0.###}; saveReopen=PASS");
            }

            if (File.Exists(tracePath))
            {
                var lines = File.ReadAllLines(tracePath)
                    .Where(line => line.IndexOf(
                        "format-conversion-",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    .ToArray();
                File.WriteAllLines(
                    Path.Combine(artifactRoot, "conversion-trace-focused.txt"),
                    lines);
            }
            Console.WriteLine(
                $"[USER OMML→VT PASS] converted={sourceOmmlCount}; "
                + $"scanMs={scanWatch.Elapsed.TotalMilliseconds:0.###}; "
                + $"conversionMs={conversionWatch.Elapsed.TotalMilliseconds:0.###}; output={outputPath}");
        }
        finally
        {
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();

            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
        }
    }

    private static void WriteUserVisualTeXOmmlSnapshot(
        Word.Document document,
        string prefix)
    {
        var lines = new List<string>
        {
            $"document={document.FullName}",
            $"content={document.Content.Start}:{document.Content.End}",
            $"paragraphs={document.Paragraphs.Count}",
            $"inlineShapes={document.InlineShapes.Count}",
            $"omaths={document.OMaths.Count}",
            $"tables={document.Tables.Count}",
            $"bookmarks={document.Bookmarks.Count}",
            $"fields={document.Fields.Count}",
        };

        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            Word.OLEFormat? ole = null;
            try
            {
                shape = document.InlineShapes[index];
                range = shape.Range;
                paragraphs = range.Paragraphs;
                if (paragraphs.Count > 0)
                {
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                }
                string progId;
                try
                {
                    ole = shape.OLEFormat;
                    progId = ole.ProgID ?? string.Empty;
                }
                catch { progId = string.Empty; }

                FormulaMetadata? metadata = null;
                if (WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    try { metadata = WordFormulaMetadataReader.TryRead(shape); }
                    catch { }
                }
                lines.Add(
                    $"SHAPE|{index}|progId={progId}|range={range.Start}:{range.End}"
                    + $"|story={range.StoryType}|paragraph={paragraphRange?.Start}:{paragraphRange?.End}"
                    + $"|withinTable={range.get_Information(Word.WdInformation.wdWithInTable)}"
                    + $"|tables={range.Tables.Count}|omaths={range.OMaths.Count}|fields={range.Fields.Count}"
                    + $"|formulaId={metadata?.FormulaId}|display={metadata?.DisplayMode}"
                    + $"|numbered={metadata?.Numbered}|latex={(metadata?.Latex ?? string.Empty).Replace("\r", "<CR>").Replace("\n", "<LF>")}");
            }
            finally
            {
                Release(ole);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(range);
                Release(shape);
            }
        }

        for (var index = 1; index <= document.OMaths.Count; index++)
        {
            Word.OMath? math = null;
            Word.Range? range = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            try
            {
                math = document.OMaths[index];
                range = math.Range;
                paragraphs = range.Paragraphs;
                if (paragraphs.Count > 0)
                {
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                }
                lines.Add(
                    $"OMATH|{index}|range={range.Start}:{range.End}|story={range.StoryType}"
                    + $"|paragraph={paragraphRange?.Start}:{paragraphRange?.End}"
                    + $"|withinTable={range.get_Information(Word.WdInformation.wdWithInTable)}"
                    + $"|text={(range.Text ?? string.Empty).Replace("\r", "<CR>").Replace("\n", "<LF>")}");
            }
            finally
            {
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(range);
                Release(math);
            }
        }

        for (var index = 1; index <= document.Bookmarks.Count; index++)
        {
            Word.Bookmark? bookmark = null;
            Word.Range? range = null;
            try
            {
                bookmark = document.Bookmarks[index];
                range = bookmark.Range;
                lines.Add(
                    $"BOOKMARK|{index}|name={bookmark.Name}|range={range.Start}:{range.End}"
                    + $"|story={range.StoryType}");
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }

        for (var index = 1; index <= document.Fields.Count; index++)
        {
            Word.Field? field = null;
            Word.Range? code = null;
            Word.Range? result = null;
            try
            {
                field = document.Fields[index];
                code = field.Code;
                result = field.Result;
                lines.Add(
                    $"FIELD|{index}|codeRange={code.Start}:{code.End}"
                    + $"|resultRange={result.Start}:{result.End}"
                    + $"|code={(code.Text ?? string.Empty).Replace("\r", "<CR>").Replace("\n", "<LF>")}"
                    + $"|result={(result.Text ?? string.Empty).Replace("\r", "<CR>").Replace("\n", "<LF>")}");
            }
            finally
            {
                Release(result);
                Release(code);
                Release(field);
            }
        }

        File.WriteAllLines(prefix + ".structure.txt", lines);
        File.WriteAllText(prefix + ".wordopenxml.xml", WordDocumentXml.Read(document));
    }
}

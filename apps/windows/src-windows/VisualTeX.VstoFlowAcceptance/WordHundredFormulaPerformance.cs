using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Xml.Linq;
using Extensibility;
using Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class WordPerformanceHost : IDisposable
    {
        private Array _custom = Array.Empty<object>();
        private COMAddIns? _installedAddIns;
        private COMAddIn? _installedAddIn;
        private bool _disposed;

        internal WordPerformanceHost(string? documentPath)
        {
            Application = CreateWordApplication(visible: false);
            _installedAddIns = Application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                _installedAddIn = _installedAddIns.Item(ref addInIndex);
                if (_installedAddIn.Connect) _installedAddIn.Connect = false;
            }
            catch
            {
                Release(_installedAddIn);
                _installedAddIn = null;
            }

            Document = string.IsNullOrWhiteSpace(documentPath)
                ? Application.Documents.Add()
                : Application.Documents.Open(
                    documentPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
            try { Document.Activate(); }
            catch (COMException error) when (
                error.HResult == unchecked((int)0x80010001)
                || error.HResult == unchecked((int)0x8001010A))
            {
                // Office 2021 can reject explicit activation in a hidden
                // automation instance even though the newly created/opened
                // document is already that instance's ActiveDocument.
            }
            AddIn = new ThisAddIn();
            AddIn.OnConnection(
                Application,
                ext_ConnectMode.ext_cm_AfterStartup,
                AddIn,
                ref _custom);
        }

        internal Word.Application Application { get; }
        internal Word.Document Document { get; }
        internal ThisAddIn AddIn { get; }

        internal void Save(string path)
        {
            Document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try { AddIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref _custom); }
            catch { }
            if (_installedAddIn is not null)
            {
                try { _installedAddIn.Connect = true; } catch { }
            }
            try { Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(Application); } catch { }
            Release(_installedAddIn);
            Release(_installedAddIns);
            Release(Document);
            Release(Application);
            ForceComCleanup();
        }
    }

    private static void RunWordHundredFormulaPerformance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Console.WriteLine("[Word 100 performance] Creating 50 OLE + 50 OMML formulas...");
        var documentPath = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-100-Formula-Performance.docx");
        // The native OLE server intentionally accepts preview files only below
        // LocalAppData\VisualTeX\office\temp. Use the same trusted staging root
        // as production Session commits so this acceptance exercises the real
        // security boundary instead of weakening it for tests.
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"word-100-acceptance-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var timings = new List<PerformanceTimingEntry>();
        List<PerformanceFormulaEntry> formulas;

        using (var host = new WordPerformanceHost(documentPath: null))
        {
            formulas = CreatePerformanceCorpus(host, assetRoot);
            VerifyPerformanceInventory(host.Document, formulas, "after corpus creation");
            host.Save(documentPath);
            Console.WriteLine("[Word 100 performance] Measuring edits before Word restart...");
            RunPerformanceEditPhase(
                client,
                host,
                formulas,
                phase: "before-word-restart",
                round: 1,
                timings);
            VerifyPerformanceInventory(host.Document, formulas, "after first edit pass");
            host.Save(documentPath);
        }

        // Force a genuine Word process/application restart, not merely a document reopen.
        Thread.Sleep(300);
        ForceComCleanup();

        using (var host = new WordPerformanceHost(documentPath))
        {
            VerifyPerformanceInventory(host.Document, formulas, "after Word restart");
            Console.WriteLine("[Word 100 performance] Measuring edits after Word restart...");
            RunPerformanceEditPhase(
                client,
                host,
                formulas,
                phase: "after-word-restart",
                round: 2,
                timings);
            VerifyPerformanceInventory(host.Document, formulas, "after second edit pass");
            host.Save(documentPath);
        }

        var summaries = timings
            .GroupBy(item => item.Phase, StringComparer.Ordinal)
            .Select(group => SummarizePerformance(group.Key, group.ToArray()))
            .ToArray();
        var reportPath = Path.Combine(artifactRoot, "word-100-performance.json");
        var report = new
        {
            GeneratedAt = DateTimeOffset.Now,
            P95ThresholdMilliseconds = WordPerformanceLimitMilliseconds,
            MaximumOutlierThresholdMilliseconds = WordPerformanceOutlierLimitMilliseconds,
            Document = documentPath,
            FormulaCount = formulas.Count,
            OleCount = formulas.Count(item => item.ObjectMode == FormulaOleContract.NativeOleMode),
            OmmlCount = formulas.Count(item => item.ObjectMode == FormulaOleContract.WordOmmlMode),
            InlineCount = formulas.Count(item => item.DisplayMode == "inline"),
            DisplayCount = formulas.Count(item => item.DisplayMode == "block"),
            Summaries = summaries,
            Timings = timings,
        };
        File.WriteAllText(
            reportPath,
            JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }),
            new UTF8Encoding(false));

        foreach (var summary in summaries)
        {
            Console.WriteLine(
                $"  {summary.Phase}: open p50={summary.OpenP50Milliseconds:F1}ms, "
                + $"p95={summary.OpenP95Milliseconds:F1}ms, max={summary.OpenMaximumMilliseconds:F1}ms; "
                + $"apply p50={summary.ApplyP50Milliseconds:F1}ms, "
                + $"p95={summary.ApplyP95Milliseconds:F1}ms, max={summary.ApplyMaximumMilliseconds:F1}ms");
        }

        var violations = timings
            .Where(item =>
                item.OpenMilliseconds > WordPerformanceOutlierLimitMilliseconds
                || item.ApplyMilliseconds > WordPerformanceOutlierLimitMilliseconds)
            .ToArray();
        var percentileViolations = summaries
            .Where(item =>
                item.OpenP95Milliseconds > WordPerformanceLimitMilliseconds
                || item.ApplyP95Milliseconds > WordPerformanceLimitMilliseconds)
            .ToArray();
        if (violations.Length > 0 || percentileViolations.Length > 0)
        {
            var worst = violations
                .OrderByDescending(item => Math.Max(item.OpenMilliseconds, item.ApplyMilliseconds))
                .Take(12)
                .Select(item =>
                    $"{item.Phase} #{item.Index} {item.ObjectMode}/{item.DisplayMode} "
                    + $"open={item.OpenMilliseconds:F1}ms apply={item.ApplyMilliseconds:F1}ms");
            throw new InvalidDataException(
                $"Word formula performance missed its p95 "
                + $"{WordPerformanceLimitMilliseconds:F0}ms / maximum "
                + $"{WordPerformanceOutlierLimitMilliseconds:F0}ms gate "
                + $"(phase violations={percentileViolations.Length}, outliers={violations.Length}). "
                + $"Report: {reportPath}\n"
                + string.Join("\n", worst));
        }

        Console.WriteLine(
            $"[Word 100 performance] All {timings.Count} measurements passed the "
            + $"p95 {WordPerformanceLimitMilliseconds:F0}ms / maximum "
            + $"{WordPerformanceOutlierLimitMilliseconds:F0}ms gate. Report: {reportPath}");
    }

    private static List<PerformanceFormulaEntry> CreatePerformanceCorpus(
        WordPerformanceHost host,
        string assetRoot)
    {
        var service = new WordFormulaService(host.Application);
        var documentId = service.ReadActiveDocumentId();
        var createLimit = int.TryParse(
            Environment.GetEnvironmentVariable("VISUALTEX_PERF_CREATE_LIMIT"),
            out var parsedCreateLimit)
            ? Math.Min(100, Math.Max(1, parsedCreateLimit))
            : 100;
        var formulas = new List<PerformanceFormulaEntry>(createLimit);
        for (var index = 1; index <= createLimit; index++)
        {
            var objectMode = index <= 50
                ? FormulaOleContract.NativeOleMode
                : FormulaOleContract.WordOmmlMode;
            var displayMode = index % 2 == 0 ? "inline" : "block";
            var formulaId = Guid.NewGuid().ToString("D");
            var latex = PerformanceLatex(index, round: 0);
            var line = new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = latex,
            };
            var svg = PerformanceSvg(index);
            var pngDataUrl = CreatePngDataUrl(latex, ExportWidth, ExportHeight);
            var mathMl = PerformanceMathMl(index, round: 0, displayMode);
            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "create",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = documentId,
                SourceObjectId = null,
                Title = $"Performance formula {index}",
                Lines = new List<FormulaLine> { line },
                CodeFormat = "latex",
                DisplayMode = displayMode,
                ObjectMode = objectMode,
                Numbered = false,
                FontSizePt = 11,
                Status = "committing",
                Dirty = true,
                ExportResult = new OfficeExportDocument
                {
                    Svg = svg,
                    SvgBase64 = "data:image/svg+xml;base64,"
                        + Convert.ToBase64String(Encoding.UTF8.GetBytes(svg)),
                    PngBase64 = pngDataUrl,
                    MathMl = mathMl,
                    Width = ExportWidth,
                    Height = ExportHeight,
                    Baseline = ExportBaseline,
                },
            };

            host.Application.Selection.EndKey(Word.WdUnits.wdStory);
            if (index > 1) host.Application.Selection.TypeParagraph();
            if (displayMode == "inline")
                host.Application.Selection.TypeText($"Inline {index}: ");
            else
            {
                host.Application.Selection.TypeText($"Display {index}");
                host.Application.Selection.TypeParagraph();
            }

            if (objectMode == FormulaOleContract.NativeOleMode)
            {
                var (pngPath, emfPath) = CreatePerformanceOleAssets(
                    assetRoot,
                    index,
                    latex,
                    svg);
                try
                {
                    service.InsertOle(session, pngPath, emfPath);
                }
                catch (Exception error)
                {
                    throw new InvalidOperationException(
                        $"Failed to create performance VisualTeX OLE formula #{index}: {error.Message}",
                        error);
                }
            }
            else
            {
                service.InsertOmml(session, mathMl);
            }

            if (displayMode == "inline")
            {
                host.Application.Selection.TypeText(" end");
                if (objectMode == FormulaOleContract.WordOmmlMode)
                    AssertInlineOmmlSuffixOutsideMath(
                        host.Document,
                        host.Application.Selection.Start,
                        index,
                        formulaId);
            }
            formulas.Add(new PerformanceFormulaEntry
            {
                Index = index,
                FormulaId = formulaId,
                ObjectMode = objectMode,
                DisplayMode = displayMode,
                Latex = latex,
            });
            if (index % 10 == 0)
                Console.WriteLine($"  created {index}/100 formulas");
        }
        return formulas;
    }

    private static void AssertInlineOmmlSuffixOutsideMath(
        Word.Document document,
        int caretPosition,
        int formulaIndex,
        string formulaId)
    {
        Word.Range? content = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Bookmark? baselineBookmark = null;
        Word.Range? formulaBookmarkRange = null;
        Word.Range? baselineRange = null;
        Word.Range? resolvedRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? nativeRange = null;
        try
        {
            content = document.Content;
            var wordDocument = XDocument.Parse(content.WordOpenXML);
            var equation = wordDocument
                .Descendants()
                .LastOrDefault(element => element.Name.LocalName == "oMath")
                ?? throw new InvalidOperationException(
                    $"Formula #{formulaIndex} did not create a native Word equation.");
            var equationText = string.Concat(
                equation.Descendants()
                    .Where(element => element.Name.LocalName == "t")
                    .Select(element => element.Value));
            if (equationText.IndexOf("end", StringComparison.Ordinal) >= 0
                || equationText.IndexOf("\u00A0", StringComparison.Ordinal) >= 0)
            {
                formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (formulaBookmark is not null)
                {
                    formulaBookmarkRange = formulaBookmark.Range;
                    resolvedRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
                }
                var baselineName = "VTBL_" + Guid.Parse(formulaId).ToString("N");
                if (document.Bookmarks.Exists(baselineName))
                {
                    baselineBookmark = document.Bookmarks[baselineName];
                    baselineRange = baselineBookmark.Range;
                }
                maths = document.OMaths;
                if (maths.Count > 0)
                {
                    math = maths[maths.Count];
                    nativeRange = math.Range;
                }
                throw new InvalidOperationException(
                    $"Formula #{formulaIndex} physically absorbed trailing prose into its OMML XML: "
                    + equationText
                    + $"; formulaBookmark={formulaBookmarkRange?.Start}:{formulaBookmarkRange?.End}"
                    + $"; baseline={baselineRange?.Start}:{baselineRange?.End}"
                    + $"; resolved={resolvedRange?.Start}:{resolvedRange?.End} '{resolvedRange?.Text}'"
                    + $"; native={nativeRange?.Start}:{nativeRange?.End} '{nativeRange?.Text}'"
                    + $"; caret={caretPosition}");
            }
        }
        finally
        {
            Release(nativeRange);
            Release(math);
            Release(maths);
            Release(resolvedRange);
            Release(baselineRange);
            Release(formulaBookmarkRange);
            Release(baselineBookmark);
            Release(formulaBookmark);
            Release(content);
        }
    }

    private static void RunPerformanceEditPhase(
        VisualTeXSessionClient client,
        WordPerformanceHost host,
        IReadOnlyList<PerformanceFormulaEntry> formulas,
        string phase,
        int round,
        ICollection<PerformanceTimingEntry> timings)
    {
        var service = new WordFormulaService(host.Application);
        var startIndex = int.TryParse(
            Environment.GetEnvironmentVariable("VISUALTEX_PERF_START_INDEX"),
            out var parsedStart)
            ? Math.Max(1, parsedStart)
            : 1;
        var endIndex = int.TryParse(
            Environment.GetEnvironmentVariable("VISUALTEX_PERF_END_INDEX"),
            out var parsedEnd)
            ? Math.Min(formulas.Count, parsedEnd)
            : formulas.Count;
        foreach (var formula in formulas.Where(item =>
                     item.Index >= startIndex && item.Index <= endIndex))
        {
            SelectPerformanceFormula(host.Document, formula);
            var selectedBeforeOpen = service.ReadSelection();
            if (!string.Equals(
                    selectedBeforeOpen.FormulaId,
                    formula.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    $"Formula #{formula.Index} selection resolved to "
                    + $"{selectedBeforeOpen.FormulaId ?? "no formula"} before opening.");
            var existing = SnapshotSessionIds();
            var openWatch = Stopwatch.StartNew();
            host.AddIn.OnEditSelected(new object());
            var sessionId = WaitForNewSessionFast(
                existing,
                "word",
                TimeSpan.FromSeconds(3));
            var session = WaitForEditorReadyFast(
                client,
                sessionId,
                TimeSpan.FromSeconds(3));
            openWatch.Stop();

            AssertEqual("edit", session.Mode, $"Formula #{formula.Index} did not open in edit mode.");
            AssertEqual(formula.FormulaId, session.FormulaId,
                $"Formula #{formula.Index} opened the wrong formulaId.");
            AssertEqual(formula.ObjectMode, session.ObjectMode,
                $"Formula #{formula.Index} opened with the wrong object mode.");
            AssertEqual(formula.DisplayMode, session.DisplayMode,
                $"Formula #{formula.Index} opened with the wrong display mode.");
            AssertEqual(formula.Latex, string.Join("\n", session.Lines.Select(line => line.Latex)),
                $"Formula #{formula.Index} opened another formula's source.");
            if (session.Lines.Any(line =>
                line.Latex.IndexOf("\\tag", StringComparison.Ordinal) >= 0))
                throw new InvalidDataException(
                    $"Formula #{formula.Index} exposed an equation tag inside editable LaTeX.");

            var updatedLatex = PerformanceLatex(formula.Index, round);
            var applyWatch = Stopwatch.StartNew();
            Commit(
                client,
                session,
                formula.DisplayMode,
                formula.ObjectMode,
                updatedLatex,
                mathMl: formula.ObjectMode == FormulaOleContract.WordOmmlMode
                    ? PerformanceMathMl(formula.Index, round, formula.DisplayMode)
                    : null);
            var final = WaitForTerminalFast(
                client,
                sessionId,
                TimeSpan.FromSeconds(3));
            WaitForAddInIdleFast(host.AddIn, TimeSpan.FromSeconds(3));
            applyWatch.Stop();

            AssertEqual("completed", final.Status,
                final.Error ?? $"Formula #{formula.Index} update did not complete.");
            AssertEqual(formula.FormulaId, final.FormulaId,
                $"Formula #{formula.Index} changed identity after update.");
            AssertEqual(updatedLatex, string.Join("\n", final.Lines.Select(line => line.Latex)),
                $"Formula #{formula.Index} committed the wrong source.");
            client.CloseEditorAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();

            host.Document.Activate();
            SelectPerformanceFormula(host.Document, formula);
            var selected = service.ReadSelection();
            AssertEqual(formula.FormulaId, selected.FormulaId,
                $"Formula #{formula.Index} was not reselected after applying the edit.");
            AssertEqual(updatedLatex, selected.Metadata?.Latex,
                $"Formula #{formula.Index} stored incorrect metadata after applying the edit.");
            formula.Latex = updatedLatex;

            timings.Add(new PerformanceTimingEntry
            {
                Phase = phase,
                Index = formula.Index,
                FormulaId = formula.FormulaId,
                ObjectMode = formula.ObjectMode,
                DisplayMode = formula.DisplayMode,
                OpenMilliseconds = openWatch.Elapsed.TotalMilliseconds,
                ApplyMilliseconds = applyWatch.Elapsed.TotalMilliseconds,
            });
            Console.WriteLine(
                $"  {phase} #{formula.Index:000} {formula.ObjectMode}/{formula.DisplayMode}: "
                + $"open={openWatch.Elapsed.TotalMilliseconds:F1}ms, "
                + $"apply={applyWatch.Elapsed.TotalMilliseconds:F1}ms");

            if (formula.Index % 10 == 0)
                VerifyPerformanceInventory(
                    host.Document,
                    formulas,
                    $"{phase} after formula {formula.Index}");
        }
    }

    private static string WaitForNewSessionFast(
        HashSet<string> existing,
        string expectedHost,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            foreach (var directory in Directory.EnumerateDirectories(SessionRoot)
                         .OrderByDescending(Directory.GetLastWriteTimeUtc))
            {
                var id = Path.GetFileName(directory);
                if (string.IsNullOrWhiteSpace(id) || existing.Contains(id)) continue;
                var sessionPath = Path.Combine(directory, "session.json");
                if (!File.Exists(sessionPath)) continue;
                string json;
                try { json = File.ReadAllText(sessionPath); }
                catch { continue; }
                if (json.IndexOf(
                        $"\"host\": \"{expectedHost}\"",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    || json.IndexOf(
                        $"\"host\":\"{expectedHost}\"",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return id;
            }
            Thread.Sleep(4);
        }
        throw new TimeoutException($"No new {expectedHost} Office Session appeared within {timeout.TotalMilliseconds:F0}ms.");
    }

    private static OfficeSessionDocument WaitForEditorReadyFast(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            try
            {
                session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }
            catch
            {
                Thread.Sleep(4);
                continue;
            }
            if (session.Dirty)
                throw new InvalidOperationException(
                    $"Session {sessionId} became dirty before the performance edit began.");
            if (session.Status == "editing") return session;
            if (session.Status is "completed" or "failed" or "cancelled")
                throw new InvalidOperationException(
                    $"Session {sessionId} reached {session.Status} before editor readiness: {session.Error}");
            Thread.Sleep(4);
        }
        throw new TimeoutException(
            $"Session {sessionId} did not signal editor readiness; last status was {session?.Status ?? "unknown"}.");
    }

    private static OfficeSessionDocument WaitForTerminalFast(
        VisualTeXSessionClient client,
        string sessionId,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        OfficeSessionDocument? session = null;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            var sessionPath = Path.Combine(SessionRoot, sessionId, "session.json");
            if (File.Exists(sessionPath))
            {
                try
                {
                    var json = File.ReadAllText(sessionPath);
                    session = JsonSerializer.Deserialize<OfficeSessionDocument>(json);
                }
                catch
                {
                    // Session writes are atomic in production, but antivirus or
                    // filesystem notification timing can briefly expose a file
                    // between replace operations. Retry inside the same deadline.
                }
            }
            if (session is null)
            {
                try
                {
                    session = client.GetSessionAsync(sessionId, CancellationToken.None)
                        .GetAwaiter().GetResult();
                }
                catch (TaskCanceledException)
                {
                    Thread.Sleep(4);
                    continue;
                }
            }
            if (session.Status is "completed" or "failed" or "cancelled") return session;
            session = null;
            Thread.Sleep(4);
        }
        throw new TimeoutException(
            $"Session {sessionId} did not finish; last status was {session?.Status ?? "unknown"}.");
    }

    private static void WaitForAddInIdleFast(object addIn, TimeSpan timeout)
    {
        var field = addIn.GetType().GetField(
            "_operationGate",
            System.Reflection.BindingFlags.Instance
            | System.Reflection.BindingFlags.NonPublic)
            ?? throw new MissingFieldException("Office add-in operation gate is missing.");
        var gate = field.GetValue(addIn) as SemaphoreSlim
            ?? throw new InvalidOperationException("Office add-in operation gate is unavailable.");
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            if (gate.CurrentCount == 1) return;
            Thread.Sleep(2);
        }
        throw new TimeoutException("Office add-in did not return to idle during the performance measurement.");
    }

    private static void SelectPerformanceFormula(
        Word.Document document,
        PerformanceFormulaEntry formula)
    {
        if (formula.ObjectMode == FormulaOleContract.NativeOleMode)
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
                        if (!string.Equals(
                                metadata?.FormulaId,
                                formula.FormulaId,
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        shape.Range.Select();
                        return;
                    }
                    finally { Release(shape); }
                }
            }
            finally { Release(shapes); }
            throw new InvalidDataException($"OLE formula #{formula.Index} no longer exists.");
        }

        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formula.FormulaId)
                ?? throw new InvalidDataException($"OMML formula #{formula.Index} no longer exists.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            equationRange.Select();
        }
        finally
        {
            Release(equationRange);
            Release(bookmark);
        }
    }

    private static void VerifyPerformanceInventory(
        Word.Document document,
        IReadOnlyList<PerformanceFormulaEntry> expected,
        string stage)
    {
        var observed = new Dictionary<string, FormulaMetadata>(StringComparer.OrdinalIgnoreCase);
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
                    if (metadata is not null) observed[metadata.FormulaId] = metadata;
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }

        foreach (var formula in expected.Where(item => item.ObjectMode == FormulaOleContract.WordOmmlMode))
        {
            Word.Bookmark? bookmark = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formula.FormulaId)
                    ?? throw new InvalidDataException(
                        $"{stage}: OMML formula #{formula.Index} is missing.");
                var metadata = WordOmmlFormulaStore.TryRead(document, formula.FormulaId)
                    ?? throw new InvalidDataException(
                        $"{stage}: OMML metadata #{formula.Index} is missing.");
                observed[metadata.FormulaId] = metadata;
            }
            finally { Release(bookmark); }
        }

        AssertEqual(expected.Count, observed.Count,
            $"{stage}: formula inventory count changed.");
        var expectedOleCount = expected.Count(item =>
            item.ObjectMode == FormulaOleContract.NativeOleMode);
        var expectedOmmlCount = expected.Count(item =>
            item.ObjectMode == FormulaOleContract.WordOmmlMode);
        AssertEqual(expectedOleCount, expected.Count(item =>
            item.ObjectMode == FormulaOleContract.NativeOleMode
            && observed.ContainsKey(item.FormulaId)),
            $"{stage}: OLE formula count changed.");
        AssertEqual(expectedOmmlCount, expected.Count(item =>
            item.ObjectMode == FormulaOleContract.WordOmmlMode
            && observed.ContainsKey(item.FormulaId)),
            $"{stage}: OMML formula count changed.");
        foreach (var formula in expected)
        {
            if (!observed.TryGetValue(formula.FormulaId, out var metadata))
                throw new InvalidDataException($"{stage}: formula #{formula.Index} is missing.");
            AssertEqual(formula.Latex, metadata.Latex,
                $"{stage}: formula #{formula.Index} source was mixed with another formula.");
            AssertEqual(formula.DisplayMode, metadata.DisplayMode,
                $"{stage}: formula #{formula.Index} display mode changed.");
            if (metadata.Latex.IndexOf("\\tag", StringComparison.Ordinal) >= 0)
                throw new InvalidDataException(
                    $"{stage}: formula #{formula.Index} persisted an editable equation tag.");
        }
    }

    private static PerformanceSummary SummarizePerformance(
        string phase,
        IReadOnlyList<PerformanceTimingEntry> entries)
    {
        var opens = entries.Select(item => item.OpenMilliseconds).OrderBy(value => value).ToArray();
        var applies = entries.Select(item => item.ApplyMilliseconds).OrderBy(value => value).ToArray();
        return new PerformanceSummary
        {
            Phase = phase,
            Count = entries.Count,
            OpenP50Milliseconds = Percentile(opens, 0.50),
            OpenP95Milliseconds = Percentile(opens, 0.95),
            OpenMaximumMilliseconds = opens.Length > 0 ? opens[opens.Length - 1] : 0,
            ApplyP50Milliseconds = Percentile(applies, 0.50),
            ApplyP95Milliseconds = Percentile(applies, 0.95),
            ApplyMaximumMilliseconds = applies.Length > 0 ? applies[applies.Length - 1] : 0,
        };
    }

    private static double Percentile(double[] sorted, double percentile)
    {
        if (sorted.Length == 0) return 0;
        var index = Math.Max(0, Math.Min(
            sorted.Length - 1,
            (int)Math.Ceiling(sorted.Length * percentile) - 1));
        return sorted[index];
    }

    private static string PerformanceLatex(int index, int round)
    {
        var suffix = round > 0 ? $"+r_{{{round}}}" : string.Empty;
        return $"\\mathrm{{e}}^{{\\mathrm{{i}}\\pi}}+x_{{{index}}}={index}{suffix}";
    }

    private static string PerformanceMathMl(int index, int round, string displayMode)
    {
        var suffix = round > 0
            ? $"<mo>+</mo><msub><mi>r</mi><mn>{round}</mn></msub>"
            : string.Empty;
        var display = displayMode == "block" ? "block" : "inline";
        return "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\""
            + display
            + "\"><msup><mi mathvariant=\"normal\">e</mi><mrow>"
            + "<mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
            + $"<mo>+</mo><msub><mi>x</mi><mn>{index}</mn></msub>"
            + $"<mo>=</mo><mn>{index}</mn>{suffix}</math>";
    }

    private static string PerformanceSvg(int index) =>
        "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 160 32\">"
        + "<rect x=\"0\" y=\"0\" width=\"160\" height=\"32\" fill=\"none\"/>"
        + $"<text x=\"5\" y=\"23\" font-family=\"Cambria Math\" font-size=\"20\" fill=\"#111\">F{index:000}</text>"
        + "</svg>";

    private static (string PngPath, string EmfPath) CreatePerformanceOleAssets(
        string assetRoot,
        int index,
        string latex,
        string svg)
    {
        var svgPath = Path.Combine(assetRoot, $"formula-{index:000}.svg");
        var pngPath = Path.Combine(assetRoot, $"formula-{index:000}.png");
        File.WriteAllText(svgPath, svg, new UTF8Encoding(false));
        using (var bitmap = new Bitmap(320, 64, PixelFormat.Format32bppArgb))
        using (var graphics = Graphics.FromImage(bitmap))
        using (var font = new Font("Cambria Math", 22f, FontStyle.Regular, GraphicsUnit.Pixel))
        using (var brush = new SolidBrush(Color.Black))
        {
            graphics.Clear(Color.Transparent);
            graphics.DrawString($"F{index:000}: {latex}", font, brush, new PointF(2, 8));
            bitmap.Save(pngPath, ImageFormat.Png);
        }
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
            svgPath,
            ExportWidth,
            ExportHeight);
        return (pngPath, emfPath);
    }
}

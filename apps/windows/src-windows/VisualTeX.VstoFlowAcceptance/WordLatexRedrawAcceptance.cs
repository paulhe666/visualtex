using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const int LatexRedrawPerformanceLimitMilliseconds = 250;
    private const uint WordObjIdNativeOm = 0xFFFFFFF0;

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr windowHandle,
        uint objectId,
        ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out object nativeObject);

    private static T RetryRejectedOfficeCall<T>(Func<T> action)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(30);
        while (true)
        {
            try { return action(); }
            catch (COMException error) when (
                (error.HResult == unchecked((int)0x80010001)
                    || error.HResult == unchecked((int)0x8001010A))
                && DateTime.UtcNow < deadline)
            {
                WinForms.Application.DoEvents();
                Thread.Sleep(100);
            }
        }
    }

    private static void RetryRejectedOfficeCall(Action action) =>
        RetryRejectedOfficeCall(() =>
        {
            action();
            return true;
        });

    private sealed class WordSourceContextStressHost : IDisposable
    {
        private readonly bool _ownsApplication;
        private Word.Document? _previousDocument;
        private bool _disposed;

        internal WordSourceContextStressHost()
        {
            if (AttachActiveWord)
            {
                Application = Marshal.GetActiveObject("Word.Application") as Word.Application
                    ?? throw new InvalidOperationException("No active Word instance is available.");
                try
                {
                    _previousDocument = RetryRejectedOfficeCall(
                        () => Application.ActiveDocument);
                }
                catch
                {
                    _previousDocument = null;
                }
            }
            else
            {
                Application = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone,
                };
                _ownsApplication = true;
            }

            Document = RetryRejectedOfficeCall(
                () => Application.Documents.Add(Visible: false));
            if (AttachActiveWord)
            {
                RetryRejectedOfficeCall(() => Document.Activate());
                Word.Windows? windows = null;
                Word.Window? window = null;
                try
                {
                    windows = RetryRejectedOfficeCall(() => Document.Windows);
                    if (RetryRejectedOfficeCall(() => windows.Count) > 0)
                    {
                        window = RetryRejectedOfficeCall(() => windows[1]);
                        RetryRejectedOfficeCall(() => window.Visible = false);
                    }
                }
                finally
                {
                    Release(window);
                    Release(windows);
                }
            }
        }

        internal Word.Application Application { get; }
        internal Word.Document Document { get; }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            try
            {
                RetryRejectedOfficeCall(
                    () => Document.Close(Word.WdSaveOptions.wdDoNotSaveChanges));
            }
            catch { }
            if (_previousDocument is not null)
            {
                try { RetryRejectedOfficeCall(() => _previousDocument.Activate()); }
                catch { }
            }
            if (_ownsApplication)
            {
                try { Application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(_previousDocument);
            Release(Document);
            Release(Application);
            ForceComCleanup();
        }
    }

    private static void RunWordLatexRedraw(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        RunWordLatexRedrawSourceContextStress();
        Console.WriteLine("[Word LaTeX redraw] Prewarming reusable hidden converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();

        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.WordOmmlMode,
            wholeDocument: false,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-OMML.docx");
        const string oleFileName = "VisualTeX-Word-Latex-Redraw-OLE.docx";
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.NativeOleMode,
            wholeDocument: true,
            expectedFileName: oleFileName);
        AssertSavedInlineOleTypingAnchor(
            Path.Combine(artifactRoot, oleFileName));
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.MathTypeOleMode,
            wholeDocument: false,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-MathType-Selection.docx",
            verifyResizeBaseline: false);
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.MathTypeOleMode,
            wholeDocument: true,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-MathType-Document.docx",
            verifyResizeBaseline: false);
    }

    private static void RunWordLatexRedrawOmmlOnly(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Console.WriteLine("[Word LaTeX redraw OMML] Prewarming reusable hidden converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.WordOmmlMode,
            wholeDocument: false,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-OMML-Only.docx");
        RunWordLatexRedrawNumberedOmmlSpacingScenario(artifactRoot);
    }

    private static void RunWordLatexRedrawNumberedOmmlSpacingScenario(
        string artifactRoot)
    {
        var documentPath = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-Latex-Redraw-Numbered-OMML-Spacing.docx");
        var logPath = Path.Combine(
            artifactRoot,
            "word-latex-redraw-numbered-omml-spacing.log");
        TryDeleteAcceptanceFile(documentPath);
        TryDeleteAcceptanceFile(logPath);
        var previousLog = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG");
        var previousNumbering = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS");
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
            logPath);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS",
            "1");
        try
        {
            using var host = new WordPerformanceHost(documentPath: null);
            var selection = host.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory);
            selection.Font.Size = 10.5f;
            selection.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;
            selection.TypeText("BEFORE_NUMBERED_REDRAW");
            selection.TypeParagraph();
            selection.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphCenter;
            selection.TypeText("$$a_1=b_1$$");
            selection.TypeParagraph();
            selection.TypeText("$$a_2=b_2$$");
            selection.TypeParagraph();
            selection.TypeText("$$a_3=b_3$$");
            selection.TypeParagraph();
            selection.ParagraphFormat.Alignment =
                Word.WdParagraphAlignment.wdAlignParagraphLeft;
            selection.TypeText("AFTER_NUMBERED_REDRAW");

            host.Application.Selection.SetRange(
                host.Document.Content.Start,
                host.Document.Content.Start);
            host.AddIn.OnRedrawDocumentToOmml(new object());
            _ = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(4));
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            host.Save(documentPath);

            AssertEqual(3, host.Document.OMaths.Count,
                "First centered numbered OMML redraw created the wrong equation count.");
            AssertEqual(3, host.Document.Tables.Count,
                "First centered numbered OMML redraw did not create one direct-SEQ 1x3 table per display formula.");
            AssertNumberedOmmlRedrawTableSpacing(host.Document);

            // Reproduce the real multi-redraw workflow: flatten the numbered OMML
            // tables back to visible LaTeX, keep those display-source paragraphs
            // centered, then redraw the whole document to numbered OMML again.
            // The second pass used to preserve the centered source paragraph pPr
            // and leave a full-height paragraph between neighboring direct-SEQ
            // tables even though no stale SEQ field remained.
            TryDeleteAcceptanceFile(logPath);
            host.Application.Selection.SetRange(
                host.Document.Content.Start,
                host.Document.Content.Start);
            host.AddIn.OnRedrawDocumentOmmlToLatex(new object());
            _ = WaitForFormulaToLatex(logPath, expectedCompletions: 1);
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            AssertEqual(0, host.Document.OMaths.Count,
                "OMML-to-LaTeX round trip left native equations before the second redraw.");
            AssertEqual(0, host.Document.Tables.Count,
                "OMML-to-LaTeX round trip left numbered direct-SEQ tables before the second redraw.");
            foreach (var latexSource in new[]
                     {
                         "$$a_1=b_1$$",
                         "$$a_2=b_2$$",
                         "$$a_3=b_3$$",
                     })
                SetLatexSourceParagraphAlignment(
                    host.Document,
                    latexSource,
                    Word.WdParagraphAlignment.wdAlignParagraphCenter);

            TryDeleteAcceptanceFile(logPath);
            host.Application.Selection.SetRange(
                host.Document.Content.Start,
                host.Document.Content.Start);
            host.AddIn.OnRedrawDocumentToOmml(new object());
            _ = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(4));
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            host.Save(documentPath);

            AssertEqual(3, host.Document.OMaths.Count,
                "Second centered numbered OMML redraw created the wrong equation count.");
            AssertEqual(3, host.Document.Tables.Count,
                "Second centered numbered OMML redraw did not create one direct-SEQ 1x3 table per display formula.");
            AssertNumberedOmmlRedrawTableSpacing(host.Document);
            Console.WriteLine(
                $"[Word LaTeX redraw] Centered multi-redraw numbered OMML spacing regression passed: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS",
                previousNumbering);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
                previousLog);
        }
    }

    private static void SetLatexSourceParagraphAlignment(
        Word.Document document,
        string latexSource,
        Word.WdParagraphAlignment alignment)
    {
        Word.Range? search = null;
        Word.Find? find = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? format = null;
        try
        {
            search = document.Content.Duplicate;
            find = search.Find;
            find.ClearFormatting();
            find.Text = latexSource;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.MatchWildcards = false;
            AssertTrue(find.Execute(),
                $"Centered multi-redraw source '{latexSource}' was not found after OMML-to-LaTeX conversion.");
            paragraphs = search.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                $"Centered multi-redraw source '{latexSource}' spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraphRange.ParagraphFormat;
            format.Alignment = alignment;
            AssertEqual(alignment, format.Alignment,
                $"Centered multi-redraw source '{latexSource}' did not retain the requested paragraph alignment.");
        }
        finally
        {
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(find);
            Release(search);
        }
    }

    private static void AssertNumberedOmmlRedrawTableSpacing(
        Word.Document document)
    {
        Word.Tables? tables = null;
        var tableRanges = new List<(int Start, int End)>();
        try
        {
            tables = document.Tables;
            for (var index = 1; index <= tables.Count; index++)
            {
                Word.Table? table = null;
                Word.Range? tableRange = null;
                try
                {
                    table = tables[index];
                    AssertEqual(1, table.Rows.Count,
                        $"Numbered OMML redraw table {index} has the wrong row count.");
                    AssertEqual(3, table.Columns.Count,
                        $"Numbered OMML redraw table {index} has the wrong column count.");
                    tableRange = table.Range;
                    tableRanges.Add((tableRange.Start, tableRange.End));
                }
                finally
                {
                    Release(tableRange);
                    Release(table);
                }
            }

            tableRanges.Sort((left, right) => left.Start.CompareTo(right.Start));
            for (var index = 0; index + 1 < tableRanges.Count; index++)
            {
                var left = tableRanges[index];
                var right = tableRanges[index + 1];
                Word.Range? gap = null;
                Word.Font? font = null;
                Word.ParagraphFormat? format = null;
                try
                {
                    gap = document.Range(left.End, right.Start);
                    var gapText = gap.Text ?? string.Empty;
                    if (gapText.Length > 1
                        || gapText.Any(character => character != '\r'))
                        throw new InvalidDataException(
                            $"Numbered OMML redraw left extra body paragraphs between tables {index + 1} and {index + 2}: length={gapText.Length}.");
                    if (gapText.Length == 0) continue;

                    font = gap.Font;
                    format = gap.ParagraphFormat;
                    if (font.Size > 1.25f
                        || format.LineSpacing > 1.25f)
                        throw new InvalidDataException(
                            $"Numbered OMML redraw left a full-height separator between tables {index + 1} and {index + 2}: font={font.Size:F2}pt line={format.LineSpacing:F2}pt.");
                }
                finally
                {
                    Release(format);
                    Release(font);
                    Release(gap);
                }
            }
        }
        finally
        {
            Release(tables);
        }
    }

    private static void RunWordLatexRedrawMathTypeOnly(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Console.WriteLine("[Word LaTeX redraw MathType] Prewarming reusable hidden converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.MathTypeOleMode,
            wholeDocument: false,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-MathType-Selection.docx",
            verifyResizeBaseline: false);
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.MathTypeOleMode,
            wholeDocument: true,
            expectedFileName: "VisualTeX-Word-Latex-Redraw-MathType-Document.docx",
            verifyResizeBaseline: false);
        RunWordMathTypeToLatexRibbonAcceptance(artifactRoot);
    }

    private static void RunWordMathTypeToLatexRibbonAcceptance(string artifactRoot)
    {
        var sourcePath = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-Latex-Redraw-MathType-Selection.docx");
        var outputPath = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-MathType-To-Latex-Ribbon.docx");
        var logPath = Path.Combine(
            artifactRoot,
            "word-mathtype-to-latex-ribbon.log");
        TryDeleteAcceptanceFile(outputPath);
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
            logPath);
        try
        {
            List<(string Latex, string DisplayMode)> expected;
            using (var host = new WordPerformanceHost(sourcePath))
            {
                expected = CaptureMathTypeLatexExpectations(host.Document);
                AssertEqual(4, expected.Count,
                    "MathType-to-LaTeX Ribbon fixture did not contain four MathType formulas.");

                Word.InlineShapes? shapes = null;
                Word.InlineShape? first = null;
                Word.Range? firstRange = null;
                try
                {
                    shapes = host.Document.InlineShapes;
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        Release(first);
                        first = shapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(first)) continue;
                        firstRange = first.Range;
                        host.Application.Selection.SetRange(firstRange.Start, firstRange.End);
                        break;
                    }
                    if (firstRange is null)
                        throw new InvalidDataException(
                            "MathType-to-LaTeX Ribbon fixture has no selectable MathType formula.");
                }
                finally
                {
                    Release(firstRange);
                    Release(first);
                    Release(shapes);
                }

                host.AddIn.OnRedrawSelectionMathTypeToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 1);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertEqual(3, CountMathTypeOleFormulas(host.Document),
                    "Selection MathType-to-LaTeX did not convert exactly one formula.");
                AssertDocumentContains(
                    host.Document,
                    BuildExpectedMathTypeLatexSource(expected[0]));

                CollapseSelectionAtDocumentStart(host);
                host.AddIn.OnRedrawDocumentMathTypeToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 2);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
                AssertEqual(0, CountMathTypeOleFormulas(host.Document),
                    "Document MathType-to-LaTeX left MathType formulas behind.");
                foreach (var item in expected)
                    AssertDocumentContains(
                        host.Document,
                        BuildExpectedMathTypeLatexSource(item));
                host.Save(outputPath);
            }

            using (var reopened = new WordPerformanceHost(outputPath))
            {
                AssertEqual(0, CountMathTypeOleFormulas(reopened.Document),
                    "Reopened MathType-to-LaTeX document restored MathType objects unexpectedly.");
                foreach (var item in expected)
                    AssertDocumentContains(
                        reopened.Document,
                        BuildExpectedMathTypeLatexSource(item));
            }
            Console.WriteLine(
                "[Word MathType→LaTeX] Selection and document Ribbon conversions passed with save/reopen persistence.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
                null);
        }
    }

    private static List<(string Latex, string DisplayMode)> CaptureMathTypeLatexExpectations(
        Word.Document document)
    {
        var result = new List<(string Latex, string DisplayMode)>();
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
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    var metadata = MathTypeOleInterop.ReadMetadata(document.Application, shape);
                    result.Add((metadata.Latex, metadata.DisplayMode));
                }
                finally { Release(shape); }
            }
            return result;
        }
        finally { Release(shapes); }
    }

    private static int CountMathTypeOleFormulas(Word.Document document)
    {
        var count = 0;
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
                    if (MathTypeOleInterop.IsMathTypeOle(shape)) count++;
                }
                finally { Release(shape); }
            }
            return count;
        }
        finally { Release(shapes); }
    }

    private static string BuildExpectedMathTypeLatexSource(
        (string Latex, string DisplayMode) formula) =>
        string.Equals(formula.DisplayMode, "block", StringComparison.OrdinalIgnoreCase)
            ? "$$" + formula.Latex + "$$"
            : "$" + formula.Latex + "$";

    private static void RunWordInlineOleTypingBaseline(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Console.WriteLine("[Word inline OLE baseline] Prewarming reusable hidden converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();

        const string oleFileName = "VisualTeX-Word-Inline-OLE-Typing-Baseline.docx";
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.NativeOleMode,
            wholeDocument: true,
            expectedFileName: oleFileName);
        AssertSavedInlineOleTypingAnchor(Path.Combine(artifactRoot, oleFileName));
    }

    private static void RunWordInlineOleInitialTypingBaseline(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Console.WriteLine("[Word inline OLE initial baseline] Prewarming reusable hidden converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();

        const string oleFileName = "VisualTeX-Word-Inline-OLE-Initial-Typing-Baseline.docx";
        RunWordLatexRedrawScenario(
            artifactRoot,
            objectMode: FormulaOleContract.NativeOleMode,
            wholeDocument: true,
            expectedFileName: oleFileName,
            verifyResizeBaseline: false);
        AssertSavedInlineOleTypingAnchor(Path.Combine(artifactRoot, oleFileName));
    }

    private static void RunExistingWordInlineOleFontStyle(string artifactRoot)
    {
        var buildLogsRoot = Path.Combine(
            Environment.CurrentDirectory,
            "build-logs");
        var fixture = Directory
            .EnumerateFiles(
                buildLogsRoot,
                "VisualTeX-Word-Latex-Redraw-OLE.docx",
                SearchOption.AllDirectories)
            .Where(path => !path.StartsWith(
                artifactRoot,
                StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(File.GetLastWriteTimeUtc)
            .FirstOrDefault()
            ?? throw new FileNotFoundException(
                "No previously verified Word OLE redraw fixture is available.");

        Console.WriteLine(
            "[Word inline OLE font] Opening existing verified fixture: "
            + fixture);
        using var host = new WordPerformanceHost(fixture);
        AssertInlineOleResizeKeepsTrailingProseBaseline(host);
        Console.WriteLine(
            "[Word inline OLE font] Existing fixture passed: "
            + "anchor bookmark preserved, body font inherited, "
            + "typing baseline preserved for trailing prose and paragraph end.");
    }

    private static void RunWordOmmlAnchorRecovery(string artifactRoot)
    {
        using var host = new WordSourceContextStressHost();
        var service = new WordFormulaService(host.Application);
        var formulaId = Guid.NewGuid().ToString("D");
        var lineId = Guid.NewGuid().ToString("D");
        Word.Selection? selection = null;
        Word.Bookmark? bookmark = null;
        Word.Bookmark? movedBookmark = null;
        Word.Bookmark? repairedBookmark = null;
        Word.Bookmarks? bookmarks = null;
        Word.Range? equationRange = null;
        Word.Range? driftAnchor = null;
        Word.Range? repairedAnchor = null;
        Word.Range? content = null;
        try
        {
            host.Application.Visible = true;
            RetryRejectedOfficeCall(() => host.Document.Activate());
            selection = RetryRejectedOfficeCall(() => host.Application.Selection);
            RetryRejectedOfficeCall(() => selection.SetRange(0, 0));

            var insertSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Host = "word",
                Mode = "create",
                FormulaId = formulaId,
                Title = "Word Formula",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                CodeFormat = "latex",
                FontSizePt = 11,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId, Latex = "x^2+y^2=z^2" },
                },
            };
            const string mathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
                + "<msup><mi>y</mi><mn>2</mn></msup><mo>=</mo>"
                + "<msup><mi>z</mi><mn>2</mn></msup></math>";
            service.InsertOmml(insertSession, mathMl);

            bookmark = WordOmmlFormulaStore.FindByFormulaId(host.Document, formulaId)
                ?? throw new InvalidDataException("Anchor recovery acceptance lost the initial VTOMML bookmark.");
            var stored = WordOmmlFormulaStore.TryRead(host.Document, bookmark)
                ?? throw new InvalidDataException("Anchor recovery acceptance lost the initial OMML metadata.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var expectedStart = equationRange.Start;
            var expectedFingerprint = stored.NativeOmmlFingerprint
                ?? throw new InvalidDataException("Anchor recovery acceptance has no native OMML fingerprint.");
            var liveFingerprintBeforeDrift = WordOmmlConverter.ComputeOmmlFingerprint(
                equationRange.WordOpenXML);
            var initiallySynchronized = string.Equals(
                expectedFingerprint,
                liveFingerprintBeforeDrift,
                StringComparison.OrdinalIgnoreCase);
            Console.WriteLine(
                "[Word OMML anchor] stored/live fingerprint equal before drift="
                + initiallySynchronized);
            if (!initiallySynchronized)
                throw new InvalidDataException(
                    "A newly inserted OMML formula persisted a provisional fingerprint instead of the final Word OMath fingerprint.");

            // Reproduce a document created by an older build: its CustomXML
            // metadata contains the converter-side fingerprint even though the
            // VTOMML bookmark is still healthy. Opening that formula must migrate
            // the metadata to the live Word fingerprint immediately.
            var legacyMetadata = FormulaMetadataCodec.Decode(
                FormulaMetadataCodec.Encode(stored))
                ?? throw new InvalidDataException("Could not clone legacy OMML metadata.");
            legacyMetadata.NativeOmmlFingerprint = new string('0', 64);
            WordOmmlFormulaStore.Save(host.Document, legacyMetadata);
            var sessionSnapshot = WordOmmlNativeSource.RefreshForVisualTeX(
                host.Document,
                bookmark,
                legacyMetadata);
            var migratedMetadata = WordOmmlFormulaStore.TryRead(host.Document, bookmark)
                ?? throw new InvalidDataException("Legacy OMML fingerprint migration lost metadata.");
            if (!string.Equals(
                    migratedMetadata.NativeOmmlFingerprint,
                    liveFingerprintBeforeDrift,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "Opening a legacy OMML formula did not persist its live Word fingerprint.");

            // Capture the same immutable range/fingerprint that the real editor
            // carries in its Session. Then deliberately stale the CustomXML again
            // and move VTOMML far away before commit. This reproduces the reported
            // case where the editor opened successfully but Apply later failed.
            var sourceObjectId =
                $"visualtex-word-vsto-range:{equationRange.Start}:{equationRange.End}";
            var staleAfterOpen = FormulaMetadataCodec.Decode(
                FormulaMetadataCodec.Encode(migratedMetadata))
                ?? throw new InvalidDataException("Could not clone stale commit metadata.");
            staleAfterOpen.NativeOmmlFingerprint = new string('f', 64);
            WordOmmlFormulaStore.Save(host.Document, staleAfterOpen);

            RetryRejectedOfficeCall(() => selection.SetRange(host.Document.Content.End - 1, host.Document.Content.End - 1));
            RetryRejectedOfficeCall(() => selection.TypeText(new string('A', 1400)));

            var bookmarkName = WordOmmlFormulaStore.BookmarkName(formulaId);
            RetryRejectedOfficeCall(() => bookmark.Delete());
            Release(bookmark);
            bookmark = null;
            content = RetryRejectedOfficeCall(() => host.Document.Content);
            var driftPosition = Math.Max(content.Start, content.End - 2);
            driftAnchor = RetryRejectedOfficeCall(() => host.Document.Range(driftPosition, driftPosition));
            bookmarks = RetryRejectedOfficeCall(() => host.Document.Bookmarks);
            movedBookmark = RetryRejectedOfficeCall(() => bookmarks.Add(bookmarkName, driftAnchor));
            if (Math.Abs(driftAnchor.Start - expectedStart) <= 512)
                throw new InvalidDataException("Anchor recovery acceptance did not move the VTOMML bookmark far enough.");

            var staleLookupFailed = false;
            try
            {
                Release(equationRange);
                equationRange = WordOmmlFormulaStore.GetEquationRange(movedBookmark);
            }
            catch (InvalidDataException error) when (
                error.Message.IndexOf("OMML anchor", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                staleLookupFailed = true;
            }
            if (!staleLookupFailed)
                throw new InvalidDataException(
                    "The legacy stale-fingerprint setup did not reproduce the OMML-anchor lookup failure.");

            const string editedMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<msup><mi>x</mi><mn>2</mn></msup><mo>+</mo>"
                + "<msup><mi>y</mi><mn>2</mn></msup><mo>=</mo>"
                + "<msup><mi>z</mi><mn>2</mn></msup><mo>+</mo><mn>1</mn></math>";
            var editSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Host = "word",
                Mode = "edit",
                FormulaId = formulaId,
                SourceObjectId = sourceObjectId,
                Title = "Word Formula",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                CodeFormat = "latex",
                FontSizePt = 11,
                OriginalMetadata = sessionSnapshot,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId, Latex = "x^2+y^2=z^2+1" },
                },
            };
            service.ReplaceOmml(editSession, editedMathMl);

            Release(movedBookmark);
            movedBookmark = null;
            repairedBookmark = WordOmmlFormulaStore.FindByFormulaId(host.Document, formulaId)
                ?? throw new InvalidDataException("Commit-time OMML recovery lost the formula bookmark.");
            Release(equationRange);
            equationRange = WordOmmlFormulaStore.GetEquationRange(repairedBookmark);
            var committedMetadata = WordOmmlFormulaStore.TryRead(host.Document, repairedBookmark)
                ?? throw new InvalidDataException("Commit-time OMML recovery lost metadata.");
            var committedFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                equationRange.WordOpenXML);
            if (!string.Equals(
                    committedMetadata.NativeOmmlFingerprint,
                    committedFingerprint,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "OMML edit commit did not persist the final live Word fingerprint.");
            repairedAnchor = repairedBookmark.Range;
            if (Math.Abs(repairedAnchor.Start - equationRange.Start) > 1)
                throw new InvalidDataException(
                    "OMML edit commit did not rebind VTOMML to the final equation start.");

            // Finally drift the now-correct anchor again beyond the local 512
            // character recovery window. With a valid final fingerprint the
            // generic lookup must recover the unique OMath across the document
            // and permanently repair VTOMML without restarting Word.
            RetryRejectedOfficeCall(() => repairedBookmark.Delete());
            Release(repairedBookmark);
            repairedBookmark = null;
            Release(repairedAnchor);
            repairedAnchor = null;
            Release(driftAnchor);
            driftAnchor = RetryRejectedOfficeCall(() => host.Document.Range(driftPosition, driftPosition));
            movedBookmark = RetryRejectedOfficeCall(() => bookmarks.Add(bookmarkName, driftAnchor));
            repairedBookmark = WordOmmlFormulaStore.FindAtRange(host.Document, equationRange)
                ?? throw new InvalidDataException(
                    "A drifted VTOMML formula could not be rediscovered from its visible OMath range.");
            Release(movedBookmark);
            movedBookmark = null;
            Release(equationRange);
            equationRange = WordOmmlFormulaStore.GetEquationRange(repairedBookmark);
            var recoveredFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(equationRange.WordOpenXML);
            if (!string.Equals(
                    recoveredFingerprint,
                    committedFingerprint,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("Full-document anchor recovery resolved the wrong native equation.");

            repairedAnchor = repairedBookmark.Range;
            if (Math.Abs(repairedAnchor.Start - equationRange.Start) > 1)
                throw new InvalidDataException(
                    $"Full-document recovery returned the equation but left VTOMML drifted by {Math.Abs(repairedAnchor.Start - equationRange.Start)} characters.");

            var artifactPath = Path.Combine(artifactRoot, "VisualTeX-Word-OMML-Anchor-Recovery.docx");
            RetryRejectedOfficeCall(() => host.Document.SaveAs2(
                artifactPath,
                Word.WdSaveFormat.wdFormatXMLDocument));
            Console.WriteLine(
                "[Word OMML anchor] Drifted VTOMML bookmark recovered by native fingerprint and rebound to the equation.");
            Console.WriteLine("[Word OMML anchor] Artifact: " + artifactPath);
        }
        finally
        {
            Release(repairedAnchor);
            Release(content);
            Release(driftAnchor);
            Release(equationRange);
            Release(repairedBookmark);
            Release(movedBookmark);
            Release(bookmark);
            Release(bookmarks);
            Release(selection);
        }
    }

    private static void RunWordOmmlBoundaryDigitDirect(string artifactRoot)
    {
        using var host = new WordSourceContextStressHost();
        var formulaId = Guid.NewGuid().ToString("D");
        var lineId = Guid.NewGuid().ToString("D");
        var service = new WordFormulaService(host.Application);
        Word.Selection? selection = null;
        Word.Range? bodyRange = null;
        Microsoft.Office.Interop.Word.Font? bodyFont = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.Range? trailingFormulaCharacter = null;
        Word.Range? documentContent = null;
        Word.Range? trailingTextRange = null;
        Microsoft.Office.Interop.Word.Font? trailingFormulaFont = null;
        Microsoft.Office.Interop.Word.Font? trailingTextFont = null;
        Word.OMaths? maths = null;
        try
        {
            bodyRange = RetryRejectedOfficeCall(() => host.Document.Range(0, 0));
            RetryRejectedOfficeCall(() => bodyRange.Text = "正文前 ");
            bodyFont = RetryRejectedOfficeCall(() => bodyRange.Font);
            RetryRejectedOfficeCall(() => bodyFont.Name = "宋体");
            RetryRejectedOfficeCall(() => bodyFont.Size = 11f);
            host.Application.Visible = true;
            WinForms.Application.DoEvents();
            Thread.Sleep(700);
            RetryRejectedOfficeCall(() => host.Document.Activate());
            selection = RetryRejectedOfficeCall(() => host.Application.Selection);
            RetryRejectedOfficeCall(() => selection.SetRange(bodyRange.End, bodyRange.End));

            var insertSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Host = "word",
                Mode = "create",
                FormulaId = formulaId,
                Title = "Word Formula",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                CodeFormat = "latex",
                FontSizePt = 11,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId, Latex = "x+y" },
                },
            };
            const string initialMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>+</mo><mi>y</mi></math>";
            service.InsertOmml(insertSession, initialMathMl);
            RetryRejectedOfficeCall(() => selection.TypeText(" 后方正文"));

            bookmark = WordOmmlFormulaStore.FindByFormulaId(host.Document, formulaId)
                ?? throw new InvalidDataException(
                    "Direct OMML boundary acceptance could not find the inserted formula bookmark.");
            var stored = WordOmmlFormulaStore.TryRead(host.Document, bookmark)
                ?? throw new InvalidDataException(
                    "Direct OMML boundary acceptance could not read inserted metadata.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var insertedText = equationRange.Text ?? string.Empty;
            if (insertedText.IndexOf('\u200B') >= 0
                || insertedText.IndexOf('\u200C') >= 0
                || insertedText.IndexOf('\u2060') >= 0)
                throw new InvalidDataException(
                    "Initial inline OMML retained a temporary VisualTeX boundary character.");
            Word.Bookmarks? allBookmarks = null;
            try
            {
                allBookmarks = host.Document.Bookmarks;
                var boundaryName = "VTBL_" + Guid.Parse(formulaId).ToString("N");
                if (allBookmarks.Exists(boundaryName))
                    throw new InvalidDataException(
                        "Initial inline OMML retained its temporary VTBL bookmark.");
            }
            finally { Release(allBookmarks); }

            Release(equationRange);
            equationRange = null;
            Release(bookmark);
            bookmark = null;

            var editSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Host = "word",
                Mode = "edit",
                FormulaId = formulaId,
                Title = "Word Formula",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                CodeFormat = "latex",
                FontSizePt = 11,
                OriginalMetadata = stored,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId, Latex = "x+y\u200C1" },
                },
            };
            const string dirtyMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>+</mo><mi>y</mi>"
                + "<mrow data-mjx-texclass=\"ORD\"><mo>&#x200C;</mo></mrow>"
                + "<mn>1</mn></math>";
            service.ReplaceOmml(editSession, dirtyMathMl);

            maths = host.Document.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    $"Direct OMML boundary update created {maths.Count} equations instead of one.");
            Release(maths);
            maths = null;

            bookmark = WordOmmlFormulaStore.FindByFormulaId(host.Document, formulaId)
                ?? throw new InvalidDataException(
                    "Direct OMML boundary update lost the formula bookmark.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var equationText = equationRange.Text ?? string.Empty;
            if (equationText.IndexOf('1') < 0)
                throw new InvalidDataException(
                    "Direct OMML boundary update did not write the appended digit 1.");
            if (equationText.IndexOf('\u200B') >= 0
                || equationText.IndexOf('\u200C') >= 0
                || equationText.IndexOf('\u2060') >= 0)
                throw new InvalidDataException(
                    "Direct OMML boundary update left a VisualTeX boundary character inside OMath.");
            if (equationText.EndsWith(" ", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Direct OMML boundary update left an ordinary trailing space inside OMath.");
            if (equationRange.End > equationRange.Start)
            {
                trailingFormulaCharacter = host.Document.Range(
                    equationRange.End - 1,
                    equationRange.End);
                trailingFormulaFont = trailingFormulaCharacter.Font;
                if (string.Equals(
                        trailingFormulaCharacter.Text,
                        " ",
                        StringComparison.Ordinal)
                    && trailingFormulaFont.Hidden != 0)
                    throw new InvalidDataException(
                        "Direct OMML boundary update retained a hidden ASCII guard at OMath.End.");
            }

            documentContent = host.Document.Content;
            var documentText = documentContent.Text ?? string.Empty;
            var trailingTextOffset = documentText.IndexOf(
                "后方正文",
                StringComparison.Ordinal);
            if (trailingTextOffset < 0)
                throw new InvalidDataException(
                    "Direct OMML boundary update deleted the trailing body text.");
            trailingTextRange = host.Document.Range(
                documentContent.Start + trailingTextOffset,
                documentContent.Start + trailingTextOffset + "后方正文".Length);
            trailingTextFont = trailingTextRange.Font;
            if (trailingTextFont.Hidden != 0)
                throw new InvalidDataException(
                    "Direct OMML boundary update propagated hidden guard formatting to trailing body text.");

            var openXml = equationRange.WordOpenXML;
            if (openXml.IndexOf("200C", StringComparison.OrdinalIgnoreCase) >= 0
                || openXml.IndexOf("8204", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidDataException(
                    "Direct OMML boundary update left U+200C in Word Open XML.");

            var storedAfter = WordOmmlFormulaStore.TryRead(host.Document, bookmark)
                ?? throw new InvalidDataException(
                    "Direct OMML boundary update lost stored metadata.");
            if (!string.Equals(storedAfter.Latex, "x+y1", StringComparison.Ordinal)
                || storedAfter.Latex.IndexOf('\u200C') >= 0)
                throw new InvalidDataException(
                    "Direct OMML boundary update stored dirty LaTeX: " + storedAfter.Latex);
            var reopened = WordOmmlNativeSource.RefreshForVisualTeX(
                host.Document,
                bookmark,
                storedAfter);
            var reopenedLatex = string.Join(
                "\n",
                reopened.Lines.Select(line => line.Latex));
            if (!string.Equals(reopenedLatex, "x+y1", StringComparison.Ordinal)
                || reopenedLatex.IndexOf('\u200C') >= 0)
                throw new InvalidDataException(
                    "Reopening updated OMML returned dirty or duplicated LaTeX: "
                    + reopenedLatex);

            var staleBookmarks = new List<Word.Bookmark>();
            var staleFormulaIds = new List<string> { formulaId };
            try
            {
                for (var staleIndex = 0; staleIndex < 2; staleIndex++)
                {
                    var staleMetadata = FormulaMetadataCodec.Decode(
                        FormulaMetadataCodec.Encode(storedAfter))
                        ?? throw new InvalidDataException(
                            "Could not clone OMML metadata for stale-anchor acceptance.");
                    staleMetadata.FormulaId = Guid.NewGuid().ToString("D");
                    staleMetadata.Lines = staleMetadata.Lines
                        .Select(line => new FormulaLine
                        {
                            Id = Guid.NewGuid().ToString("D"),
                            Latex = line.Latex,
                        })
                        .ToList();
                    staleMetadata.Validate();
                    staleFormulaIds.Add(staleMetadata.FormulaId);
                    var staleBookmark = WordOmmlFormulaStore.Wrap(
                        host.Document,
                        equationRange,
                        staleMetadata);
                    WordOmmlFormulaStore.Save(host.Document, staleMetadata);
                    staleBookmarks.Add(staleBookmark);
                }

                foreach (var staleFormulaId in staleFormulaIds)
                    DeleteOmmlMetadataPartExternally(
                        host.Document,
                        staleFormulaId);

                if (WordOmmlFormulaStore.TryRead(host.Document, bookmark) is not null)
                    throw new InvalidDataException(
                        "OMML metadata cache revived an externally deleted CustomXMLPart.");
                foreach (var staleBookmark in staleBookmarks)
                {
                    if (WordOmmlFormulaStore.TryRead(
                            host.Document,
                            staleBookmark) is not null)
                        throw new InvalidDataException(
                            "A stale duplicate VTOMML anchor retained cached metadata.");
                }

                var detectedCount = service.CountFormulaObjectsForLatex(
                    wholeDocument: true,
                    objectMode: FormulaOleContract.WordOmmlMode);
                if (detectedCount != 1)
                    throw new InvalidDataException(
                        $"One visible OMath with stale duplicate VTOMML anchors was counted as {detectedCount} formulas.");

                var converted = service.ConvertFormulaObjectsToLatex(
                    wholeDocument: true,
                    objectMode: FormulaOleContract.WordOmmlMode);
                if (converted.FormulaCount != 1)
                    throw new InvalidDataException(
                        $"One visible OMath was converted to {converted.FormulaCount} LaTeX sources.");
                Release(maths);
                maths = host.Document.OMaths;
                if (maths.Count != 0)
                    throw new InvalidDataException(
                        "OMML-to-LaTeX conversion left a visible native equation behind.");

                Release(documentContent);
                documentContent = host.Document.Content;
                var convertedDocumentText = documentContent.Text ?? string.Empty;
                var sourceOccurrenceCount = 0;
                for (var searchIndex = 0;
                     (searchIndex = convertedDocumentText.IndexOf(
                         "x+y1",
                         searchIndex,
                         StringComparison.Ordinal)) >= 0;
                     searchIndex += "x+y1".Length)
                    sourceOccurrenceCount++;
                if (sourceOccurrenceCount != 1)
                    throw new InvalidDataException(
                        $"OMML-to-LaTeX conversion emitted x+y1 {sourceOccurrenceCount} times instead of once.");
            }
            finally
            {
                foreach (var staleBookmark in staleBookmarks)
                    Release(staleBookmark);
            }

            var artifactPath = Path.Combine(
                artifactRoot,
                "VisualTeX-Word-OMML-Boundary-Digit-Direct.docx");
            RetryRejectedOfficeCall(() => host.Document.SaveAs2(
                artifactPath,
                Word.WdSaveFormat.wdFormatXMLDocument));
            Console.WriteLine(
                "[Word OMML boundary] Real Word InsertOmml/ReplaceOmml passed: "
                + "digit 1 preserved, U+200C and trailing hidden space excluded, "
                + "trailing prose stayed visible, stale duplicate anchors counted/converted once.");
            Console.WriteLine("[Word OMML boundary] Artifact: " + artifactPath);
        }
        finally
        {
            Release(maths);
            Release(trailingTextFont);
            Release(trailingFormulaFont);
            Release(trailingTextRange);
            Release(documentContent);
            Release(trailingFormulaCharacter);
            Release(equationRange);
            Release(bookmark);
            Release(bodyFont);
            Release(bodyRange);
            Release(selection);
        }
    }

    private static void DeleteOmmlMetadataPartExternally(
        Word.Document document,
        string formulaId)
    {
        Microsoft.Office.Core.CustomXMLParts? parts = null;
        Microsoft.Office.Core.CustomXMLParts? selected = null;
        Microsoft.Office.Core.CustomXMLPart? matched = null;
        try
        {
            parts = document.CustomXMLParts;
            selected = parts.SelectByNamespace(
                WordOmmlFormulaStore.NamespaceUri);
            for (var index = 1; index <= selected.Count; index++)
            {
                Microsoft.Office.Core.CustomXMLPart? part = null;
                try
                {
                    part = selected[index];
                    if (!WordOmmlFormulaStore.TryDecodePartXml(
                            part.XML,
                            out var metadata)
                        || !string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    matched = part;
                    part = null;
                    break;
                }
                finally { Release(part); }
            }
            if (matched is null)
                throw new InvalidDataException(
                    "Stale-anchor acceptance could not find the OMML metadata part to delete.");
            matched.Delete();
        }
        finally
        {
            Release(matched);
            Release(selected);
            Release(parts);
        }
    }

    private readonly struct PixelBounds
    {
        internal PixelBounds(int minX, int minY, int maxX, int maxY, int count)
        {
            MinX = minX;
            MinY = minY;
            MaxX = maxX;
            MaxY = maxY;
            Count = count;
        }

        internal int MinX { get; }
        internal int MinY { get; }
        internal int MaxX { get; }
        internal int MaxY { get; }
        internal int Count { get; }
        internal int Width => MaxX - MinX + 1;
        internal int Height => MaxY - MinY + 1;
    }

    private static void SelectExistingWordInlineOle()
    {
        var handleText = Environment.GetEnvironmentVariable(
            "VISUALTEX_TEST_WORD_DOCUMENT_HWND");
        if (!long.TryParse(handleText, out var handleValue) || handleValue == 0)
            throw new InvalidOperationException(
                "VISUALTEX_TEST_WORD_DOCUMENT_HWND must identify the visible Word document window.");

        object? nativeObject = null;
        Word.Window? window = null;
        Word.Selection? selection = null;
        Word.Range? content = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            var dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
            var result = AccessibleObjectFromWindow(
                new IntPtr(handleValue),
                WordObjIdNativeOm,
                ref dispatchId,
                out nativeObject);
            Marshal.ThrowExceptionForHR(result);
            window = (Word.Window)nativeObject;
            RetryRejectedOfficeCall(window.Activate);
            _ = SetForegroundWindow(new IntPtr(
                RetryRejectedOfficeCall(() => window.Hwnd)));
            selection = RetryRejectedOfficeCall(() => window.Selection);
            RetryRejectedOfficeCall(selection.WholeStory);
            content = RetryRejectedOfficeCall(() => selection.Range);
            shapes = RetryRejectedOfficeCall(() => content.InlineShapes);
            var count = RetryRejectedOfficeCall(() => shapes.Count);
            if (count != 1)
                throw new InvalidDataException(
                    $"Expected one existing inline OLE, found {count}.");
            shape = RetryRejectedOfficeCall(() => shapes[1]);
            shapeRange = RetryRejectedOfficeCall(() => shape.Range);
            var start = RetryRejectedOfficeCall(() => shapeRange.Start);
            var end = RetryRejectedOfficeCall(() => shapeRange.End);
            RetryRejectedOfficeCall(() => selection.SetRange(start, end));
            Console.WriteLine(
                $"Selected existing inline OLE at range {start}-{end} in the visible Word window.");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(content);
            Release(selection);
            Release(window);
            Release(nativeObject);
        }
    }

    private static void RunExistingWordInlineOleVisualBaseline(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var handleText = Environment.GetEnvironmentVariable(
            "VISUALTEX_TEST_WORD_DOCUMENT_HWND");
        if (!long.TryParse(handleText, out var handleValue) || handleValue == 0)
            throw new InvalidOperationException(
                "VISUALTEX_TEST_WORD_DOCUMENT_HWND must identify the visible Word document window.");

        object? nativeObject = null;
        Word.Window? window = null;
        Word.Application? wordApplication = null;
        Word.InlineShapes? inlineShapes = null;
        Word.InlineShape? shape = null;
        Word.Range? formulaRange = null;
        Microsoft.Office.Interop.Word.Font? formulaFont = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Selection? selection = null;
        Word.Range? content = null;
        Word.Range? preceding = null;
        Word.Range? trailing = null;
        Microsoft.Office.Interop.Word.Font? precedingFont = null;
        Microsoft.Office.Interop.Word.Font? trailingFont = null;
        try
        {
            var dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
            var result = AccessibleObjectFromWindow(
                new IntPtr(handleValue),
                WordObjIdNativeOm,
                ref dispatchId,
                out nativeObject);
            Marshal.ThrowExceptionForHR(result);
            window = (Word.Window)nativeObject;
            RetryRejectedOfficeCall(window.Activate);
            var windowHandle = RetryRejectedOfficeCall(() => window.Hwnd);
            _ = SetForegroundWindow(new IntPtr(windowHandle));
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            selection = RetryRejectedOfficeCall(() => window.Selection);
            RetryRejectedOfficeCall(selection.WholeStory);
            content = RetryRejectedOfficeCall(() => selection.Range);
            inlineShapes = RetryRejectedOfficeCall(() => content.InlineShapes);
            var inlineShapeCount = RetryRejectedOfficeCall(() => inlineShapes.Count);
            if (inlineShapeCount != 1)
                throw new InvalidDataException(
                    $"Existing visible baseline fixture expected one inline OLE, found {inlineShapeCount}.");

            var contentText = RetryRejectedOfficeCall(() => content.Text) ?? string.Empty;
            if (!contentText.Contains('X'))
                throw new InvalidDataException(
                    "Existing visible baseline fixture does not contain adjacent text.");
            if (!contentText.TrimEnd('\r', '\a').EndsWith("X", StringComparison.Ordinal))
            {
                var insertionPosition = RetryRejectedOfficeCall(() =>
                    Math.Max(content.Start, content.End - 1));
                RetryRejectedOfficeCall(() =>
                    selection.SetRange(insertionPosition, insertionPosition));
                RetryRejectedOfficeCall(() => selection.Font.Name = "Times New Roman");
                RetryRejectedOfficeCall(() => selection.Font.Size = 42);
                RetryRejectedOfficeCall(() => selection.Font.Position = 0);
                RetryRejectedOfficeCall(() => selection.Font.Superscript = 0);
                RetryRejectedOfficeCall(() => selection.Font.Subscript = 0);
                RetryRejectedOfficeCall(() =>
                    selection.Font.Color = Word.WdColor.wdColorRed);
                RetryRejectedOfficeCall(() => selection.TypeText(" X"));
            }

            shape = RetryRejectedOfficeCall(() => inlineShapes[1]);
            formulaRange = RetryRejectedOfficeCall(() => shape.Range);
            var resizeText = Environment.GetEnvironmentVariable(
                "VISUALTEX_TEST_RESIZE_FONT_SIZE");
            if (double.TryParse(
                    resizeText,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out var resizeFontSize)
                && resizeFontSize > 0)
            {
                var formulaStart = RetryRejectedOfficeCall(() => formulaRange.Start);
                var formulaEnd = RetryRejectedOfficeCall(() => formulaRange.End);
                preceding = RetryRejectedOfficeCall(() => content.Duplicate);
                trailing = RetryRejectedOfficeCall(() => content.Duplicate);
                RetryRejectedOfficeCall(() => preceding.End = formulaStart);
                var contentEnd = RetryRejectedOfficeCall(() => content.End);
                RetryRejectedOfficeCall(() => trailing.Start = formulaEnd);
                var trailingStart = RetryRejectedOfficeCall(() => trailing.Start);
                RetryRejectedOfficeCall(() =>
                    trailing.End = Math.Max(trailingStart, contentEnd - 1));
                precedingFont = RetryRejectedOfficeCall(() => preceding.Font);
                trailingFont = RetryRejectedOfficeCall(() => trailing.Font);
                foreach (var font in new[] { precedingFont, trailingFont })
                {
                    RetryRejectedOfficeCall(() => font.Name = "Times New Roman");
                    RetryRejectedOfficeCall(() => font.Size = (float)resizeFontSize);
                    RetryRejectedOfficeCall(() => font.Position = 0);
                    RetryRejectedOfficeCall(() => font.Superscript = 0);
                    RetryRejectedOfficeCall(() => font.Subscript = 0);
                    RetryRejectedOfficeCall(() =>
                        font.Color = Word.WdColor.wdColorRed);
                }
                RetryRejectedOfficeCall(() =>
                    selection.SetRange(formulaStart, formulaEnd));
                wordApplication = RetryRejectedOfficeCall(() => window.Application);
                var service = new WordFormulaService(wordApplication);
                service.SetSelectedFormulaFontSize(resizeFontSize);
                Release(formulaRange);
                formulaRange = null;
                Release(shape);
                shape = null;
                Release(inlineShapes);
                inlineShapes = null;
                RetryRejectedOfficeCall(selection.WholeStory);
                Release(content);
                content = RetryRejectedOfficeCall(() => selection.Range);
                inlineShapes = RetryRejectedOfficeCall(() => content.InlineShapes);
                shape = RetryRejectedOfficeCall(() => inlineShapes[1]);
                formulaRange = RetryRejectedOfficeCall(() => shape.Range);
            }
            formulaFont = RetryRejectedOfficeCall(() => formulaRange.Font);
            var metadata = RetryRejectedOfficeCall(() =>
                WordFormulaMetadataReader.TryRead(shape));
            paragraphs = RetryRejectedOfficeCall(() => formulaRange.Paragraphs);
            paragraph = RetryRejectedOfficeCall(() => paragraphs[1]);
            var sourceParagraphRange = RetryRejectedOfficeCall(() => paragraph.Range);
            try
            {
                paragraphRange = RetryRejectedOfficeCall(() =>
                    sourceParagraphRange.Duplicate);
            }
            finally { Release(sourceParagraphRange); }
            var paragraphStart = RetryRejectedOfficeCall(() => paragraphRange.Start);
            var paragraphEnd = RetryRejectedOfficeCall(() => paragraphRange.End);
            RetryRejectedOfficeCall(() =>
                paragraphRange.End = Math.Max(paragraphStart, paragraphEnd - 1));

            var finalParagraphStart = RetryRejectedOfficeCall(() => paragraphRange.Start);
            var finalParagraphEnd = RetryRejectedOfficeCall(() => paragraphRange.End);
            RetryRejectedOfficeCall(() =>
                selection.SetRange(finalParagraphStart, finalParagraphEnd));
            RetryRejectedOfficeCall(window.Activate);
            _ = SetForegroundWindow(new IntPtr(windowHandle));
            WinForms.Application.DoEvents();
            Thread.Sleep(300);
            WinForms.Clipboard.Clear();
            RetryRejectedOfficeCall(selection.CopyAsPicture);
            WinForms.Application.DoEvents();
            Thread.Sleep(500);

            var imagePath = Path.Combine(
                artifactRoot,
                "word-inline-ole-visual-baseline-existing-42pt.png");
            ExportClipboardPictureThroughPowerPoint(imagePath);
            using var bitmap = new Bitmap(imagePath);
            var red = FindPixelBounds(bitmap, static color =>
                color.A > 0
                && color.R >= 120
                && color.R >= color.G + 45
                && color.R >= color.B + 45);
            var black = FindPixelBounds(bitmap, static color =>
                color.A > 0
                && color.R <= 95
                && color.G <= 95
                && color.B <= 95);
            if (red.Count < 20 || black.Count < 10)
                throw new InvalidDataException(
                    $"Existing 42 pt pixel classification was insufficient: red={red.Count}, black={black.Count}.");
            var bottomDelta = black.MaxY - red.MaxY;
            var centerDelta = ((black.MinY + black.MaxY) / 2.0)
                - ((red.MinY + red.MaxY) / 2.0);
            var shapeWidth = RetryRejectedOfficeCall(() => shape.Width);
            var shapeHeight = RetryRejectedOfficeCall(() => shape.Height);
            var formulaPosition = RetryRejectedOfficeCall(() => formulaFont.Position);
            Console.WriteLine(
                $"[Word existing inline OLE visual baseline] 42.0 pt: "
                + $"shape={shapeWidth:F2}x{shapeHeight:F2}pt, "
                + $"fontPosition={formulaPosition}, "
                + $"metadataFont={metadata?.FontSizePt?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}, "
                + $"render={metadata?.RenderWidthPx?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}x"
                + $"{metadata?.RenderHeightPx?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}px, "
                + $"baseline={metadata?.Baseline?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}px, "
                + $"red=({red.MinX},{red.MinY})-({red.MaxX},{red.MaxY}), "
                + $"black=({black.MinX},{black.MinY})-({black.MaxX},{black.MaxY}), "
                + $"bottomDelta={bottomDelta}px, centerDelta={centerDelta:F1}px, image={imagePath}");
            if (Math.Abs(bottomDelta) > 1)
                throw new InvalidDataException(
                    $"Office 2021 visible 42 pt OLE baseline still differs by {bottomDelta}px. Screenshot: {imagePath}");
        }
        finally
        {
            Release(trailingFont);
            Release(precedingFont);
            Release(trailing);
            Release(preceding);
            Release(content);
            Release(selection);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaFont);
            Release(formulaRange);
            Release(shape);
            Release(inlineShapes);
            Release(wordApplication);
            Release(window);
            Release(nativeObject);
        }
    }

    private static void RunWordInlineOleVisualBaseline(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Console.WriteLine("[Word inline OLE visual baseline] Prewarming converter...");
        client.PrewarmConverterAsync(CancellationToken.None).GetAwaiter().GetResult();
        Thread.Sleep(700);
        WinForms.Application.DoEvents();

        var logPath = Path.Combine(artifactRoot, "word-inline-ole-visual-baseline.log");
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", logPath);
        try
        {
            using var host = new WordPerformanceHost(documentPath: null);
            var selection = host.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory);
            selection.Font.Name = "Times New Roman";
            selection.Font.Size = 12;
            selection.Font.Color = Word.WdColor.wdColorRed;
            selection.TypeText("X ");
            selection.Font.Color = Word.WdColor.wdColorAutomatic;
            selection.TypeText("$X$");
            selection.Font.Color = Word.WdColor.wdColorRed;
            selection.TypeText(" X");

            var content = host.Document.Content;
            try
            {
                selection.SetRange(content.Start, content.End - 1);
                host.AddIn.OnRedrawSelectionToOle(new object());
            }
            finally { Release(content); }

            _ = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(2));
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            if (host.Document.InlineShapes.Count != 1)
                throw new InvalidDataException(
                    $"Visual baseline fixture expected one OLE formula, found {host.Document.InlineShapes.Count}.");

            foreach (var fontSize in new[] { 12d, 18d, 42d })
                MeasureInlineOleVisualBaseline(host, fontSize, artifactRoot);
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", null);
        }
    }

    private static void MeasureInlineOleVisualBaseline(
        WordPerformanceHost host,
        double fontSize,
        string artifactRoot)
    {
        Word.InlineShape? shape = null;
        Word.Range? formulaRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? preceding = null;
        Word.Range? trailing = null;
        Word.Selection? selection = null;
        Microsoft.Office.Interop.Word.Font? precedingFont = null;
        Microsoft.Office.Interop.Word.Font? trailingFont = null;
        Microsoft.Office.Interop.Word.Font? formulaFont = null;
        try
        {
            shape = host.Document.InlineShapes[1];
            formulaRange = shape.Range;
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            paragraphRange.End = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            preceding = host.Document.Range(paragraphRange.Start, formulaRange.Start);
            trailing = host.Document.Range(formulaRange.End, paragraphRange.End);
            precedingFont = preceding.Font;
            trailingFont = trailing.Font;
            foreach (var font in new[] { precedingFont, trailingFont })
            {
                font.Name = "Times New Roman";
                font.Size = (float)fontSize;
                font.Position = 0;
                font.Superscript = 0;
                font.Subscript = 0;
                font.Color = Word.WdColor.wdColorRed;
            }

            selection = host.Application.Selection;
            selection.SetRange(formulaRange.Start, formulaRange.End);
            var service = new WordFormulaService(host.Application);
            service.SetSelectedFormulaFontSize(fontSize);

            Release(shape);
            shape = host.Document.InlineShapes[1];
            Release(formulaRange);
            formulaRange = shape.Range;
            formulaFont = formulaRange.Font;
            selection.SetRange(formulaRange.Start, formulaRange.End);
            var selectedFormula = service.ReadSelection();
            Console.WriteLine(
                $"[Word inline OLE visual baseline structure] {fontSize:F1} pt: "
                + $"shape={shape.Width:F2}x{shape.Height:F2}pt, "
                + $"fontPosition={formulaFont.Position}, "
                + $"metadataFont={selectedFormula.Metadata?.FontSizePt?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}, "
                + $"render={selectedFormula.Metadata?.RenderWidthPx?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}x"
                + $"{selectedFormula.Metadata?.RenderHeightPx?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}px, "
                + $"baseline={selectedFormula.Metadata?.Baseline?.ToString("F2", CultureInfo.InvariantCulture) ?? "null"}px.");
            Release(paragraphs);
            paragraphs = formulaRange.Paragraphs;
            Release(paragraph);
            paragraph = paragraphs[1];
            Release(paragraphRange);
            paragraphRange = paragraph.Range.Duplicate;
            paragraphRange.End = Math.Max(paragraphRange.Start, paragraphRange.End - 1);

            host.Application.Visible = true;
            host.Document.Activate();
            host.Application.ActiveWindow.Activate();
            _ = SetForegroundWindow(new IntPtr(host.Application.ActiveWindow.Hwnd));
            WinForms.Application.DoEvents();
            Thread.Sleep(300);

            selection.SetRange(paragraphRange.Start, paragraphRange.End);
            WinForms.Clipboard.Clear();
            selection.CopyAsPicture();
            WinForms.Application.DoEvents();
            Thread.Sleep(500);

            var imagePath = Path.Combine(
                artifactRoot,
                $"word-inline-ole-visual-baseline-{fontSize.ToString("0.#", CultureInfo.InvariantCulture).Replace('.', '_')}pt.png");
            ExportClipboardPictureThroughPowerPoint(imagePath);
            using var bitmap = new Bitmap(imagePath);

            var red = FindPixelBounds(bitmap, static color =>
                color.A > 0
                && color.R >= 120
                && color.R >= color.G + 45
                && color.R >= color.B + 45);
            var black = FindPixelBounds(bitmap, static color =>
                color.A > 0
                && color.R <= 95
                && color.G <= 95
                && color.B <= 95);
            if (red.Count < 20 || black.Count < 10)
                throw new InvalidDataException(
                    $"Visual baseline pixel classification was insufficient at {fontSize:F1} pt: "
                    + $"red={red.Count}, black={black.Count}, image={bitmap.Width}x{bitmap.Height}.");

            var bottomDelta = black.MaxY - red.MaxY;
            var centerDelta = ((black.MinY + black.MaxY) / 2.0)
                - ((red.MinY + red.MaxY) / 2.0);
            Console.WriteLine(
                $"[Word inline OLE visual baseline] {fontSize:F1} pt: "
                + $"red=({red.MinX},{red.MinY})-({red.MaxX},{red.MaxY}), "
                + $"black=({black.MinX},{black.MinY})-({black.MaxX},{black.MaxY}), "
                + $"bottomDelta={bottomDelta}px, centerDelta={centerDelta:F1}px, image={imagePath}");
            if (Math.Abs(bottomDelta) > 3)
                throw new InvalidDataException(
                    $"Office 2021 visible inline OLE baseline differs from adjacent text at {fontSize:F1} pt: "
                    + $"formula bottom minus text bottom = {bottomDelta}px. Screenshot: {imagePath}");
        }
        finally
        {
            Release(formulaFont);
            Release(trailingFont);
            Release(precedingFont);
            Release(selection);
            Release(trailing);
            Release(preceding);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaRange);
            Release(shape);
        }
    }

    private static void ExportClipboardPictureThroughPowerPoint(string imagePath)
    {
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.ShapeRange? pasted = null;
        PowerPoint.Shape? shape = null;
        try
        {
            application = new PowerPoint.Application
            {
                Visible = Microsoft.Office.Core.MsoTriState.msoTrue,
            };
            presentation = application.Presentations.Add(
                Microsoft.Office.Core.MsoTriState.msoFalse);
            slide = presentation.Slides.Add(
                1,
                PowerPoint.PpSlideLayout.ppLayoutBlank);
            var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(8);
            Exception? lastError = null;
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    pasted = slide.Shapes.PasteSpecial(
                        PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile);
                    if (pasted.Count > 0) break;
                }
                catch (Exception error)
                {
                    lastError = error;
                    Release(pasted);
                    pasted = null;
                }
                WinForms.Application.DoEvents();
                Thread.Sleep(150);
            }
            if (pasted is null || pasted.Count == 0)
                throw new InvalidDataException(
                    "PowerPoint could not paste Word CopyAsPicture as an enhanced metafile.",
                    lastError);
            shape = pasted[1];
            shape.Export(imagePath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            if (!File.Exists(imagePath) || new FileInfo(imagePath).Length == 0)
                throw new InvalidDataException(
                    $"PowerPoint did not export the visual baseline image: {imagePath}");
        }
        finally
        {
            Release(shape);
            Release(pasted);
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

    private static PixelBounds FindPixelBounds(
        Bitmap bitmap,
        Func<Color, bool> predicate,
        int minX = 0,
        int? maxXExclusive = null)
    {
        var maxExclusive = Math.Min(bitmap.Width, maxXExclusive ?? bitmap.Width);
        var left = bitmap.Width;
        var top = bitmap.Height;
        var right = -1;
        var bottom = -1;
        var count = 0;
        for (var y = 0; y < bitmap.Height; y++)
        {
            for (var x = Math.Max(0, minX); x < maxExclusive; x++)
            {
                var color = bitmap.GetPixel(x, y);
                if (!predicate(color)) continue;
                left = Math.Min(left, x);
                top = Math.Min(top, y);
                right = Math.Max(right, x);
                bottom = Math.Max(bottom, y);
                count++;
            }
        }
        return count == 0
            ? new PixelBounds(0, 0, -1, -1, 0)
            : new PixelBounds(left, top, right, bottom, count);
    }

    private static void RunWordLatexRedrawSourceContextStress()
    {
        using var host = new WordSourceContextStressHost();
        var builder = new StringBuilder(120_000);
        var sourceFormats = new List<(int Start, int Length, float Size)>(1000);
        for (var index = 1; index <= 500; index++)
        {
            builder.Append($"第{index}段中文正文用于检查行内公式：");
            var inlineStart = builder.Length;
            var inlineSource = $"$a_{{{index}}}+b_{{{index}}}=c_{{{index}}}$";
            builder.Append(inlineSource);
            sourceFormats.Add((inlineStart, inlineSource.Length, 9.5f));
            builder.Append("。\r");

            var displayStart = builder.Length;
            var displaySource = $@"\[E_{{{index}}}=m_{{{index}}}c^2\]";
            builder.Append(displaySource);
            sourceFormats.Add((displayStart, displaySource.Length, 8.5f));
            builder.Append('\r');
        }

        Word.Range? content = null;
        Microsoft.Office.Interop.Word.Font? contentFont = null;
        try
        {
            content = RetryRejectedOfficeCall(() => host.Document.Content);
            var documentStart = RetryRejectedOfficeCall(() => content.Start);
            RetryRejectedOfficeCall(() => content.Text = builder.ToString());
            contentFont = RetryRejectedOfficeCall(() => content.Font);
            RetryRejectedOfficeCall(() => contentFont.Name = "宋体");
            RetryRejectedOfficeCall(() => contentFont.Size = 10.5f);
            foreach (var sourceFormat in sourceFormats)
            {
                Word.Range? sourceRange = null;
                Microsoft.Office.Interop.Word.Font? sourceFont = null;
                try
                {
                    sourceRange = RetryRejectedOfficeCall(() => host.Document.Range(
                        documentStart + sourceFormat.Start,
                        documentStart + sourceFormat.Start + sourceFormat.Length));
                    sourceFont = RetryRejectedOfficeCall(() => sourceRange.Font);
                    RetryRejectedOfficeCall(() => sourceFont.Size = sourceFormat.Size);
                }
                finally
                {
                    Release(sourceFont);
                    Release(sourceRange);
                }
            }
        }
        finally
        {
            Release(contentFont);
            Release(content);
        }

        RetryRejectedOfficeCall(() => host.Document.Activate());
        var service = new WordFormulaService(host.Application);
        var stopwatch = Stopwatch.StartNew();
        var plan = service.CaptureLatexRedrawPlan(wholeDocument: true);
        stopwatch.Stop();
        if (plan.Targets.Count != 1000)
            throw new InvalidDataException(
                $"1000-formula redraw scan found {plan.Targets.Count} formulas.");
        var inline = plan.Targets
            .Where(target => string.Equals(target.DisplayMode, "inline", StringComparison.Ordinal))
            .ToArray();
        var display = plan.Targets
            .Where(target => string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            .ToArray();
        if (inline.Length != 500 || display.Length != 500)
            throw new InvalidDataException(
                $"1000-formula redraw scan produced inline={inline.Length}, display={display.Length}.");
        var wrongInline = inline.Count(target => Math.Abs(target.FontSizePt - 10.5) > 0.1);
        var wrongDisplay = display.Count(target => Math.Abs(target.FontSizePt - 10.5) > 0.1);
        var unpreservedDisplay = display.Count(target => !target.PreserveDisplayParagraphBoundary);
        if (wrongInline != 0 || wrongDisplay != 0 || unpreservedDisplay != 0)
            throw new InvalidDataException(
                "1000-formula redraw context inheritance failed: "
                + $"wrongInline={wrongInline}, wrongDisplay={wrongDisplay}, "
                + $"unpreservedDisplay={unpreservedDisplay}.");
        if (stopwatch.Elapsed > TimeSpan.FromSeconds(30))
            throw new InvalidDataException(
                $"1000-formula redraw source scan took {stopwatch.Elapsed.TotalSeconds:F2}s.");
        Console.WriteLine(
            "[Word LaTeX redraw] 1000-formula source context scan passed: "
            + $"{stopwatch.ElapsedMilliseconds} ms, inline=500, display=500.");
    }

    private static void RunWordLatexRedrawDistinctFormulas(string artifactRoot)
    {
        RunWordLatexRedrawDistinctFormulaScenario(
            artifactRoot,
            FormulaOleContract.WordOmmlMode,
            "VisualTeX-Word-Latex-Redraw-Distinct-OMML.docx");
        RunWordLatexRedrawDistinctFormulaScenario(
            artifactRoot,
            FormulaOleContract.NativeOleMode,
            "VisualTeX-Word-Latex-Redraw-Distinct-OLE.docx");
    }

    private static void RunWordLatexRedrawDistinctFormulaScenario(
        string artifactRoot,
        string objectMode,
        string expectedFileName)
    {
        var modeName = objectMode == FormulaOleContract.NativeOleMode ? "OLE" : "OMML";
        var logPath = Path.Combine(
            artifactRoot,
            $"word-latex-redraw-distinct-{modeName.ToLowerInvariant()}.log");
        var documentPath = Path.Combine(artifactRoot, expectedFileName);
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", logPath);
        try
        {
            using var host = new WordPerformanceHost(documentPath: null);
            PopulateDistinctLatexRedrawDocument(host);
            Word.Range? content = null;
            try
            {
                content = host.Document.Content;
                host.Application.Selection.SetRange(content.Start, content.Start);
                if (objectMode == FormulaOleContract.NativeOleMode)
                    host.AddIn.OnRedrawDocumentToOle(new object());
                else
                    host.AddIn.OnRedrawDocumentToOmml(new object());
            }
            finally { Release(content); }

            _ = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(4));
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            AssertDistinctLatexRedrawDocument(host.Document, objectMode);
            host.Save(documentPath);
            Console.WriteLine(
                $"[Word LaTeX redraw] Distinct {modeName} redraw passed: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", null);
        }
    }

    private static void PopulateDistinctLatexRedrawDocument(WordPerformanceHost host)
    {
        var selection = host.Application.Selection;
        selection.HomeKey(Word.WdUnits.wdStory);
        selection.Font.Name = "宋体";
        selection.Font.Size = 10.5f;
        selection.TypeText("Inline A $E=mc^2$ and inline B $a^2+b^2=c^2$. ");
        selection.TypeText(@"Inline C $\sin^2 x+\cos^2 x=1$.");
        selection.TypeParagraph();
        selection.TypeText(@"$$\frac{d}{dx}\left(\sin x\right)=\cos x$$");
        selection.TypeParagraph();
        selection.TypeText(@"$$\int_{-\infty}^{\infty} e^{-x^2}\,dx=\sqrt{\pi}$$");
        selection.TypeParagraph();
        selection.TypeText(@"$$\sum_{n=1}^{\infty}\frac{1}{n^2}=\frac{\pi^2}{6}$$");
        selection.TypeParagraph();
    }

    private static void AssertDistinctLatexRedrawDocument(
        Word.Document document,
        string objectMode)
    {
        const int expectedFormulaCount = 6;
        if (objectMode == FormulaOleContract.WordOmmlMode)
        {
            Word.OMaths? maths = null;
            try
            {
                maths = document.OMaths;
                AssertEqual(
                    expectedFormulaCount,
                    maths.Count,
                    "Distinct-formula OMML redraw created the wrong equation count.");
                var nativeLatex = new List<string>(expectedFormulaCount);
                for (var index = 1; index <= maths.Count; index++)
                {
                    Word.OMath? math = null;
                    Word.Range? range = null;
                    try
                    {
                        math = maths[index];
                        range = math.Range;
                        var metadata = WordOmmlNativeSource.CreateForNative(document, range);
                        nativeLatex.Add(metadata.Latex);
                    }
                    finally
                    {
                        Release(range);
                        Release(math);
                    }
                }
                Console.WriteLine(
                    "[Word LaTeX redraw] Native OMML sources: "
                    + string.Join(" || ", nativeLatex));
                var distinct = nativeLatex
                    .Select(value => Regex.Replace(value, @"\s+", string.Empty))
                    .Distinct(StringComparer.Ordinal)
                    .Count();
                AssertEqual(
                    expectedFormulaCount,
                    distinct,
                    "Distinct LaTeX formulas collapsed to duplicate native OMML equations.");
            }
            finally { Release(maths); }
            return;
        }

        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            var latex = new List<string>();
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is not null && WordFormulaMetadataReader.IsNativeOle(shape))
                        latex.Add(metadata.Latex);
                }
                finally { Release(shape); }
            }
            Console.WriteLine(
                "[Word LaTeX redraw] OLE metadata sources: " + string.Join(" || ", latex));
            AssertEqual(
                expectedFormulaCount,
                latex.Count,
                "Distinct-formula OLE redraw created the wrong object count.");
            AssertEqual(
                expectedFormulaCount,
                latex.Distinct(StringComparer.Ordinal).Count(),
                "Distinct LaTeX formulas collapsed to duplicate OLE metadata.");
        }
        finally { Release(shapes); }
    }

    private static void RunWordLatexRedrawScenario(
        string artifactRoot,
        string objectMode,
        bool wholeDocument,
        string expectedFileName,
        bool verifyResizeBaseline = true)
    {
        var modeName = objectMode == FormulaOleContract.NativeOleMode
            ? "OLE"
            : objectMode == FormulaOleContract.MathTypeOleMode
                ? "MathType"
                : "OMML";
        var logPath = Path.Combine(
            artifactRoot,
            $"word-latex-redraw-{modeName.ToLowerInvariant()}.log");
        var documentPath = Path.Combine(artifactRoot, expectedFileName);
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", logPath);
        try
        {
            using var host = new WordPerformanceHost(documentPath: null);
            PopulateLatexRedrawDocument(host);
            Word.Range? content = null;
            try
            {
                content = host.Document.Content;
                if (wholeDocument)
                {
                    host.Application.Selection.SetRange(content.Start, content.Start);
                    if (objectMode == FormulaOleContract.NativeOleMode)
                        host.AddIn.OnRedrawDocumentToOle(new object());
                    else if (objectMode == FormulaOleContract.MathTypeOleMode)
                        host.AddIn.OnRedrawDocumentToMathType(new object());
                    else
                        host.AddIn.OnRedrawDocumentToOmml(new object());
                }
                else
                {
                    host.Application.Selection.SetRange(content.Start, content.End - 1);
                    if (objectMode == FormulaOleContract.NativeOleMode)
                        host.AddIn.OnRedrawSelectionToOle(new object());
                    else if (objectMode == FormulaOleContract.MathTypeOleMode)
                        host.AddIn.OnRedrawSelectionToMathType(new object());
                    else
                        host.AddIn.OnRedrawSelectionToOmml(new object());
                }
            }
            finally { Release(content); }

            var redrawLog = WaitForLatexRedraw(logPath, TimeSpan.FromMinutes(4));
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(30));
            AssertLatexRedrawDocument(host.Document, objectMode);
            if (objectMode == FormulaOleContract.NativeOleMode)
            {
                AssertInitialInlineOleTypingBoundaryNavigation(host);
                if (verifyResizeBaseline)
                    AssertInlineOleResizeKeepsTrailingProseBaseline(host);
            }
            host.Save(documentPath);
            AssertLatexRedrawPerformance(redrawLog, modeName);
            Console.WriteLine(
                $"[Word LaTeX redraw] {modeName} {(wholeDocument ? "document" : "selection")} redraw passed: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", null);
        }
    }

    private static void PopulateLatexRedrawDocument(WordPerformanceHost host)
    {
        var selection = host.Application.Selection;
        selection.HomeKey(Word.WdUnits.wdStory);

        selection.Font.Name = "宋体";
        selection.Font.Size = 10.5f;
        // Reproduce the user's real document where ordinary body text itself has
        // a -1 pt character position. A correct inline-OLE boundary must preserve
        // that nonzero prose baseline rather than silently normalizing it to 0.
        selection.Font.Position = -1;
        // A supplementary Unicode character before later formulas reproduces
        // Word's story-coordinate/UTF-16 offset mismatch from real documents.
        // The raw LaTeX run is intentionally smaller than the prose, matching
        // word_10000_chinese_500_inline_500_display.docx.
        selection.TypeText("普通正文🙂前 ");
        selection.Font.Size = 9.5f;
        selection.TypeText("$UVI>2$");
        selection.Font.Size = 10.5f;
        // Match the real inline-OLE document structure: following prose starts
        // immediately after the formula. VisualTeX may keep only its U+200C
        // zero-width typing anchor between the OLE and this first body character.
        selection.TypeText("普通正文后。");
        selection.TypeParagraph();
        selection.Font.Position = 0;

        selection.Font.Size = 8.5f;
        selection.TypeText(@"\[E=mc^2\]");
        selection.TypeParagraph();

        selection.Font.Size = 16;
        selection.TypeText("AFTER_DISPLAY_BODY 大字号正文前 ");
        selection.Font.Size = 9.5f;
        selection.TypeText(@"\(f_x:V\to\mathbb{R},\ f_x(y):=\ip{x}{y}\)");
        selection.Font.Size = 16;
        selection.TypeText(" 大字号正文后。");
        selection.TypeParagraph();

        selection.Font.Size = 12;
        selection.TypeText("无线信号前 ");
        selection.Font.Size = 9.5f;
        selection.TypeText(@"$\left( \text{约}1.4\times 10^{-5}eV \right)$");
        selection.Font.Size = 12;
        selection.TypeText(" 无线信号后。");
        selection.TypeParagraph();

        selection.Font.Size = 10.5f;
        selection.TypeText("INVALID_LATEX_BEFORE ");
        selection.Font.Size = 9.5f;
        selection.TypeText(@"$\VisualTeXDefinitelyUnknown{z}$");
        selection.Font.Size = 10.5f;
        selection.TypeText(" INVALID_LATEX_AFTER");
        selection.TypeParagraph();
    }

    private static string WaitForLatexRedraw(string logPath, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var last = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(25);
            try
            {
                if (!File.Exists(logPath)) continue;
                last = File.ReadAllText(logPath, Encoding.UTF8);
                if (last.IndexOf("redraw-failed", StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException(
                        "Word LaTeX redraw failed.\n" + last);
                if (last.IndexOf("redraw-complete", StringComparison.Ordinal) >= 0)
                    return last;
            }
            catch (IOException)
            {
                // The add-in may be appending the current timing line.
            }
        }
        throw new TimeoutException(
            $"Word LaTeX redraw did not complete within {timeout.TotalSeconds:F0}s. Last log:\n{last}");
    }

    private static void AssertLatexRedrawDocument(
        Word.Document document,
        string objectMode)
    {
        Word.Range? content = null;
        Word.OMaths? maths = null;
        Word.InlineShapes? shapes = null;
        try
        {
            content = document.Content;
            var text = content.Text ?? string.Empty;
            foreach (var required in new[]
                     {
                         "普通正文🙂前",
                         "普通正文后",
                         "大字号正文前",
                         "大字号正文后",
                         "无线信号前",
                         "无线信号后",
                         "INVALID_LATEX_BEFORE",
                         @"$\VisualTeXDefinitelyUnknown{z}$",
                         "INVALID_LATEX_AFTER",
                     })
            {
                if (text.IndexOf(required, StringComparison.Ordinal) < 0)
                    throw new InvalidDataException(
                        $"Word LaTeX redraw lost surrounding prose: {required}");
            }
            foreach (var forbidden in new[] { "$UVI>2$", @"\(", @"\)", @"\[", @"\]", @"\ip" })
            {
                if (text.IndexOf(forbidden, StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException(
                        $"Word LaTeX redraw left source LaTeX text in the document: {forbidden}");
            }

            if (objectMode == FormulaOleContract.WordOmmlMode)
            {
                maths = document.OMaths;
                if (maths.Count != 4)
                    throw new InvalidDataException(
                        $"OMML redraw created {maths.Count} equations instead of 4.");
                AssertOmmlFontSize(maths[1], 10.5, 1);
                AssertOmmlFontSize(maths[2], 10.5, 2);
                AssertOmmlFontSize(maths[3], 16, 3);
                AssertOmmlFontSize(maths[4], 12, 4);
                Word.OMath? displayMath = null;
                Word.Range? displayRange = null;
                try
                {
                    displayMath = maths[2];
                    displayRange = displayMath.Range;
                    AssertDisplayFormulaFollowedImmediatelyByText(
                        document,
                        displayRange,
                        "AFTER_DISPLAY_BODY");
                }
                finally
                {
                    Release(displayRange);
                    Release(displayMath);
                }
            }
            else if (objectMode == FormulaOleContract.MathTypeOleMode)
            {
                shapes = document.InlineShapes;
                var mathType = new List<(Word.InlineShape Shape, FormulaMetadata Metadata)>();
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    try
                    {
                        shape = shapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                        var metadata = MathTypeOleInterop.ReadMetadata(document.Application, shape);
                        mathType.Add((shape, metadata));
                        shape = null;
                    }
                    finally { Release(shape); }
                }
                try
                {
                    if (mathType.Count != 4)
                        throw new InvalidDataException(
                            $"MathType redraw created {mathType.Count} Equation.DSMT4 objects instead of 4.");
                    AssertEqual(0, document.OMaths.Count,
                        "MathType redraw unexpectedly left Word OMML equations behind.");
                    var normalizedLatex = mathType
                        .Select(item => (item.Metadata.Latex ?? string.Empty)
                            .Replace(" ", string.Empty)
                            .Replace("{", string.Empty)
                            .Replace("}", string.Empty))
                        .ToArray();
                    AssertTrue(normalizedLatex.Any(value =>
                            value.IndexOf("UVI>2", StringComparison.Ordinal) >= 0),
                        "MathType redraw did not preserve the UVI>2 inline formula.");
                    AssertTrue(normalizedLatex.Any(value =>
                            value.IndexOf("E=mc^2", StringComparison.Ordinal) >= 0),
                        "MathType redraw did not preserve the E=mc^2 display formula.");
                    var display = mathType.FirstOrDefault(item =>
                        string.Equals(item.Metadata.DisplayMode, "block", StringComparison.Ordinal)
                        && (item.Metadata.Latex ?? string.Empty)
                            .Replace(" ", string.Empty)
                            .Replace("{", string.Empty)
                            .Replace("}", string.Empty)
                            .IndexOf("E=mc^2", StringComparison.Ordinal) >= 0);
                    if (display.Shape is null)
                        throw new InvalidDataException(
                            "MathType redraw could not resolve the E=mc^2 display equation.");
                    Word.Range? displayRange = null;
                    try
                    {
                        displayRange = display.Shape.Range;
                        AssertDisplayFormulaFollowedImmediatelyByText(
                            document,
                            displayRange,
                            "AFTER_DISPLAY_BODY");
                    }
                    finally { Release(displayRange); }
                }
                finally
                {
                    foreach (var item in mathType) Release(item.Shape);
                }
            }
            else
            {
                shapes = document.InlineShapes;
                var native = new List<(Word.InlineShape Shape, FormulaMetadata Metadata)>();
                for (var index = 1; index <= shapes.Count; index++)
                {
                    var shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is not null && WordFormulaMetadataReader.IsNativeOle(shape))
                        native.Add((shape, metadata));
                    else
                        Release(shape);
                }
                try
                {
                    if (native.Count != 4)
                        throw new InvalidDataException(
                            $"OLE redraw created {native.Count} VisualTeX objects instead of 4.");
                    AssertOleFontSize(native, "UVI>2", 10.5, "OLE UVI formula font size");
                    AssertOleFontSize(
                        native,
                        @"f_x:V\to\mathbb{R},\ f_x(y):=\ip{x}{y}",
                        16,
                        "OLE inner-product formula font size");
                    AssertOleFontSize(native, "E=mc^2", 10.5, "OLE display formula font size");
                    AssertOleFontSize(
                        native,
                        @"\left( \text{约}1.4\times 10^{-5}eV \right)",
                        12,
                        "OLE wireless formula font size");
                    var display = native.Single(item =>
                        string.Equals(item.Metadata.Latex, "E=mc^2", StringComparison.Ordinal));
                    Word.Range? displayRange = null;
                    try
                    {
                        displayRange = display.Shape.Range;
                        AssertDisplayFormulaFollowedImmediatelyByText(
                            document,
                            displayRange,
                            "AFTER_DISPLAY_BODY");
                    }
                    finally { Release(displayRange); }
                }
                finally
                {
                    foreach (var item in native) Release(item.Shape);
                }
            }
        }
        finally
        {
            Release(shapes);
            Release(maths);
            Release(content);
        }
    }

    private static void AssertInitialInlineOleTypingBoundaryNavigation(
        WordPerformanceHost host)
    {
        Word.InlineShapes? shapes = null;
        Word.InlineShape? target = null;
        Word.Range? formulaRange = null;
        Word.Range? preceding = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? anchorRange = null;
        Word.Selection? selection = null;
        Microsoft.Office.Interop.Word.Font? precedingFont = null;
        try
        {
            shapes = host.Document.InlineShapes;
            FormulaMetadata? metadata = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    var candidateMetadata = WordFormulaMetadataReader.TryRead(candidate);
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)
                        || !string.Equals(candidateMetadata?.Latex, "UVI>2", StringComparison.Ordinal))
                        continue;
                    target = candidate;
                    metadata = candidateMetadata;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (target is null || metadata is null)
                throw new InvalidDataException(
                    "Initial typing-baseline acceptance could not find the UVI inline OLE.");

            formulaRange = target.Range;
            preceding = host.Document.Range(
                Math.Max(0, formulaRange.Start - 1),
                formulaRange.Start);
            precedingFont = preceding.Font;
            var expectedPosition = precedingFont.Position;
            if (expectedPosition == (int)Word.WdConstants.wdUndefined)
                expectedPosition = 0;

            AssertInlineOleTypingAnchor(
                host.Document,
                target,
                metadata,
                expectedPosition,
                precedingFont);

            var bookmarkName = "VTBL_" + Guid.Parse(metadata.FormulaId).ToString("N");
            bookmarks = host.Document.Bookmarks;
            bookmark = bookmarks[bookmarkName];
            anchorRange = bookmark.Range;

            host.Application.Visible = true;
            host.Document.Activate();
            host.Application.ActiveWindow.Activate();
            _ = SetForegroundWindow(new IntPtr(host.Application.ActiveWindow.Hwnd));
            WinForms.Application.DoEvents();
            Thread.Sleep(200);
            selection = host.Application.Selection;

            void AssertAnchorStillBodyFormatted(string stage)
            {
                Release(anchorRange);
                anchorRange = bookmark.Range;
                Microsoft.Office.Interop.Word.Font? anchorFont = null;
                try
                {
                    anchorFont = anchorRange.Font;
                    if (anchorFont.Position != expectedPosition)
                        throw new InvalidDataException(
                            $"{stage}: inline OLE typing anchor baseline changed. "
                            + $"Expected {expectedPosition}, actual {anchorFont.Position}.");
                    AssertBodyCharacterFormattingMatches(
                        precedingFont,
                        anchorFont,
                        stage + " anchor");
                }
                finally { Release(anchorFont); }
            }

            void TypeProbe(string text, string stage)
            {
                Release(anchorRange);
                anchorRange = bookmark.Range;
                selection.SetRange(anchorRange.End, anchorRange.End);
                WinForms.Application.DoEvents();
                Thread.Sleep(100);
                var typedStart = selection.Start;
                Word.Range? typedRange = null;
                Microsoft.Office.Interop.Word.Font? typedFont = null;
                try
                {
                    selection.TypeText(text);
                    WinForms.Application.DoEvents();
                    Thread.Sleep(100);
                    typedRange = host.Document.Range(
                        typedStart,
                        typedStart + text.Length);
                    if (!string.Equals(typedRange.Text, text, StringComparison.Ordinal))
                        throw new InvalidDataException(
                            $"{stage}: Word inserted the typing probe at an unexpected range.");
                    typedFont = typedRange.Font;
                    if (typedFont.Position != expectedPosition)
                        throw new InvalidDataException(
                            $"{stage}: text typed after the inline OLE changed baseline. "
                            + $"Expected {expectedPosition}, actual {typedFont.Position}.");
                    AssertBodyCharacterFormattingMatches(
                        precedingFont,
                        typedFont,
                        stage);
                    typedRange.Delete();
                }
                finally
                {
                    Release(typedFont);
                    Release(typedRange);
                }
                AssertAnchorStillBodyFormatted(stage + " after delete");
            }

            TypeProbe("directprobe", "Initial direct typing");

            Release(anchorRange);
            anchorRange = bookmark.Range;
            if (!string.Equals(anchorRange.Text, "\u200C", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The persistent inline-OLE typing boundary must be U+200C, not a visible space.");
            Word.Range? followingCharacter = null;
            try
            {
                if (anchorRange.End < host.Document.Content.End - 1)
                {
                    followingCharacter = host.Document.Range(anchorRange.End, anchorRange.End + 1);
                    if (string.Equals(followingCharacter.Text, " ", StringComparison.Ordinal))
                        throw new InvalidDataException(
                            "VisualTeX left an ordinary ASCII space after the inline OLE typing anchor.");
                }
            }
            finally { Release(followingCharacter); }

            selection.SetRange(anchorRange.End, anchorRange.End);
            for (var cycle = 1; cycle <= 4; cycle++)
            {
                // Use physical keyboard navigation so Word executes the same
                // caret-affinity path as a real user. There is intentionally no
                // visible post-OLE space: only the U+200C typing anchor exists.
                WinForms.SendKeys.SendWait("{LEFT}");
                WinForms.Application.DoEvents();
                Thread.Sleep(75);
                AssertAnchorStillBodyFormatted($"Physical arrow cycle {cycle} left");
                WinForms.SendKeys.SendWait("{RIGHT}");
                WinForms.Application.DoEvents();
                Thread.Sleep(75);
                AssertAnchorStillBodyFormatted($"Physical arrow cycle {cycle} right");
            }

            TypeProbe("arrowprobe", "Typing after physical left/right navigation");

            Release(anchorRange);
            anchorRange = bookmark.Range;
            var prosePosition = Math.Min(
                Math.Max(anchorRange.End + 1, anchorRange.End),
                Math.Max(anchorRange.End, host.Document.Content.End - 1));
            selection.SetRange(prosePosition, prosePosition);
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            selection.SetRange(anchorRange.End, anchorRange.End);
            WinForms.Application.DoEvents();
            Thread.Sleep(150);
            AssertAnchorStillBodyFormatted("Repositioning from following prose");
            TypeProbe("repositionprobe", "Typing after prose re-selection");

            Console.WriteLine(
                "[Word inline OLE initial baseline] Passed direct typing, four left/right cycles, and prose re-selection without baseline drift.");
        }
        finally
        {
            Release(precedingFont);
            Release(selection);
            Release(anchorRange);
            Release(bookmark);
            Release(bookmarks);
            Release(preceding);
            Release(formulaRange);
            Release(target);
            Release(shapes);
        }
    }

    private static void AssertSavedInlineOleTypingAnchor(string documentPath)
    {
        using var host = new WordPerformanceHost(documentPath);
        Word.InlineShapes? shapes = null;
        Word.InlineShape? target = null;
        Word.Range? formulaRange = null;
        Word.Range? preceding = null;
        Word.Selection? selection = null;
        Word.Range? typedRange = null;
        Microsoft.Office.Interop.Word.Font? precedingFont = null;
        Microsoft.Office.Interop.Word.Font? typedFont = null;
        try
        {
            shapes = host.Document.InlineShapes;
            FormulaMetadata? metadata = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    var candidateMetadata = WordFormulaMetadataReader.TryRead(candidate);
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)
                        || !string.Equals(
                            candidateMetadata?.Latex,
                            "UVI>2",
                            StringComparison.Ordinal))
                        continue;
                    target = candidate;
                    metadata = candidateMetadata;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (target is null || metadata is null)
                throw new InvalidDataException(
                    "Save/reopen typing-anchor acceptance could not find the resized inline OLE.");

            formulaRange = target.Range;
            preceding = host.Document.Range(
                Math.Max(0, formulaRange.Start - 1),
                formulaRange.Start);
            precedingFont = preceding.Font;
            var expectedPosition = precedingFont.Position;
            if (expectedPosition == (int)Word.WdConstants.wdUndefined)
                expectedPosition = 0;
            AssertInlineOleTypingAnchor(
                host.Document,
                target,
                metadata,
                expectedPosition,
                precedingFont);

            host.Application.Visible = true;
            host.Application.ActiveWindow.Activate();
            _ = SetForegroundWindow(new IntPtr(host.Application.ActiveWindow.Hwnd));
            WinForms.Application.DoEvents();
            Thread.Sleep(300);
            selection = host.Application.Selection;
            selection.SetRange(formulaRange.End, formulaRange.End);
            var service = new WordFormulaService(host.Application);
            service.NormalizeTypingCaretAfterInlineFormula(selection);
            WinForms.Application.DoEvents();
            Thread.Sleep(300);
            var typedStart = selection.Start;
            const string typedText = "typedaftersavereopen";
            WinForms.SendKeys.SendWait(typedText);
            WinForms.Application.DoEvents();
            Thread.Sleep(300);
            typedRange = host.Document.Range(
                typedStart,
                typedStart + typedText.Length);
            typedFont = typedRange.Font;
            if (typedFont.Position != expectedPosition)
                throw new InvalidDataException(
                    "Typing after save/reopen inherited the resized OLE baseline. "
                    + $"Expected {expectedPosition}, actual {typedFont.Position}.");
            AssertBodyCharacterFormattingMatches(
                precedingFont,
                typedFont,
                "Typing after save/reopen inline OLE");
            Console.WriteLine(
                "[Word LaTeX redraw] Saved/reopened inline OLE zero-width typing anchor passed real keyboard input.");
        }
        finally
        {
            Release(typedFont);
            Release(typedRange);
            Release(selection);
            Release(precedingFont);
            Release(preceding);
            Release(formulaRange);
            Release(target);
            Release(shapes);
        }
    }

    private static void AssertInlineOleResizeKeepsTrailingProseBaseline(
        WordPerformanceHost host)
    {
        Word.InlineShapes? shapes = null;
        Word.InlineShape? target = null;
        FormulaMetadata? targetMetadata = null;
        Word.Range? formulaRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? preceding = null;
        Word.Range? trailing = null;
        Word.Range? paragraphTail = null;
        Word.Selection? typingSelection = null;
        Word.Window? inputWindow = null;
        Word.Range? typedRange = null;
        Microsoft.Office.Interop.Word.Font? precedingFont = null;
        Microsoft.Office.Interop.Word.Font? trailingFont = null;
        Microsoft.Office.Interop.Word.Font? formulaFont = null;
        Microsoft.Office.Interop.Word.Font? typedFont = null;
        try
        {
            Console.WriteLine("[Word inline OLE font] Stage 1: enumerate inline shapes.");
            shapes = host.Document.InlineShapes;
            Console.WriteLine($"[Word inline OLE font] Inline shape count: {shapes.Count}.");
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? candidate = null;
                try
                {
                    Console.WriteLine($"[Word inline OLE font] Inspect shape {index}/{shapes.Count}.");
                    candidate = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(candidate);
                    Console.WriteLine(
                        $"[Word inline OLE font] Shape {index} latex={metadata?.Latex ?? "<null>"}.");
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)
                        || !string.Equals(metadata?.Latex, "UVI>2", StringComparison.Ordinal))
                        continue;
                    target = candidate;
                    targetMetadata = metadata;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (target is null)
                throw new InvalidDataException(
                    "Inline OLE baseline acceptance could not find the UVI formula.");

            Console.WriteLine("[Word inline OLE font] Stage 2: target UVI formula located.");
            formulaRange = target.Range;
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            preceding = host.Document.Range(
                Math.Max(paragraphRange.Start, formulaRange.Start - 1),
                formulaRange.Start);
            trailing = host.Document.Range(
                formulaRange.End,
                Math.Max(formulaRange.End, paragraphRange.End - 1));
            precedingFont = preceding.Font;
            trailingFont = trailing.Font;
            var expectedPosition = precedingFont.Position;
            if (expectedPosition == (int)Word.WdConstants.wdUndefined)
                expectedPosition = 0;

            // Reproduce the visible corruption reported by users: the ordinary
            // text run after a large inline OLE inherits the object's negative
            // position. Resizing through VisualTeX must restore that run.
            trailingFont.Position = -9;
            host.Application.Selection.SetRange(formulaRange.Start, formulaRange.End);
            var service = new WordFormulaService(host.Application);
            Console.WriteLine("[Word inline OLE font] Stage 3: resize target to 42 pt.");
            service.SetSelectedFormulaFontSize(42);
            Console.WriteLine("[Word inline OLE font] Stage 3 complete.");
            Release(formulaRange);
            formulaRange = target.Range;
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            Release(trailing);
            trailing = host.Document.Range(
                formulaRange.End,
                Math.Max(formulaRange.End, paragraphRange.End - 1));
            Console.WriteLine("[Word inline OLE font] Stage 4: verify typing anchor formatting.");
            AssertInlineOleTypingAnchor(
                host.Document,
                target,
                targetMetadata
                    ?? throw new InvalidDataException(
                        "Inline OLE baseline acceptance lost formula metadata."),
                expectedPosition,
                precedingFont);

            Release(trailingFont);
            trailingFont = trailing.Font;
            if (trailingFont.Position != expectedPosition)
                throw new InvalidDataException(
                    "Resizing an inline OLE left the following prose on a different baseline. "
                    + $"Expected {expectedPosition}, actual {trailingFont.Position}.");
            formulaFont = formulaRange.Font;
            if (formulaFont.Position == trailingFont.Position)
                throw new InvalidDataException(
                    "Inline OLE baseline repair incorrectly removed the formula's own alignment offset.");

            // A collapsed Word insertion point can report Position=0 while the
            // next character still inherits the OLE object's negative baseline.
            // Validate the actual typed run rather than trusting caret metadata.
            Console.WriteLine("[Word inline OLE font] Stage 5: activate test window and type probe.");
            host.Application.Visible = true;
            Console.WriteLine("[Word inline OLE font] Stage 5a: application visible.");
            host.Document.Activate();
            Console.WriteLine("[Word inline OLE font] Stage 5b: document activated.");
            var documentWindows = host.Document.Windows;
            try
            {
                inputWindow = documentWindows[1];
                inputWindow.Visible = true;
                Console.WriteLine("[Word inline OLE font] Stage 5c: document window visible.");
                inputWindow.Activate();
                Console.WriteLine("[Word inline OLE font] Stage 5d: document window activated.");
                _ = SetForegroundWindow(new IntPtr(inputWindow.Hwnd));
                Console.WriteLine("[Word inline OLE font] Stage 5e: foreground requested.");
            }
            finally { Release(documentWindows); }
            WinForms.Application.DoEvents();
            Thread.Sleep(300);
            typingSelection = host.Application.Selection;
            Console.WriteLine("[Word inline OLE font] Stage 5f: selection acquired.");
            typingSelection.SetRange(formulaRange.End, formulaRange.End);
            Console.WriteLine("[Word inline OLE font] Stage 5g: selection moved to formula end.");
            Release(typingSelection);
            typingSelection = null;
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            typingSelection = host.Application.Selection;
            Console.WriteLine("[Word inline OLE font] Stage 5h: event-driven caret normalization completed.");
            var typedStart = typingSelection.Start;
            const string typedText = "typedafterlargeole";
            Console.WriteLine($"[Word inline OLE font] Stage 5i: type at {typedStart}.");
            typingSelection.TypeText(typedText);
            Console.WriteLine("[Word inline OLE font] Stage 5j: TypeText returned.");
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            typedRange = host.Document.Range(
                typedStart,
                typedStart + typedText.Length);
            if (!string.Equals(typedRange.Text, typedText, StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not insert the resize typing probe at the normalized OLE boundary. "
                    + $"Expected '{typedText}', actual '{typedRange.Text}', "
                    + $"formulaEnd={formulaRange.End}, typedStart={typedStart}.");
            typedFont = typedRange.Font;
            if (typedFont.Position != expectedPosition)
                throw new InvalidDataException(
                    "Typing after a resized inline OLE inherited the formula baseline. "
                    + $"Expected {expectedPosition}, actual {typedFont.Position}; "
                    + $"formulaEnd={formulaRange.End}, typedStart={typedStart}; "
                    + DescribeCharacterFormatting(
                        host.Document,
                        Math.Max(0, formulaRange.End - 1),
                        typedStart + typedText.Length + 1));
            AssertBodyCharacterFormattingMatches(
                precedingFont,
                typedFont,
                "Typing after resized inline OLE");

            // Repeat with the OLE as the final content in its paragraph. Without
            // an existing prose run after the object, Word can display a normal
            // caret but apply the object's negative Position to keyboard input.
            Release(typedFont);
            typedFont = null;
            Release(typedRange);
            typedRange = null;
            Release(typingSelection);
            typingSelection = null;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            paragraphTail = host.Document.Range(
                formulaRange.End,
                Math.Max(formulaRange.End, paragraphRange.End - 1));
            paragraphTail.Delete();
            Release(formulaRange);
            formulaRange = target.Range;
            host.Application.Selection.SetRange(formulaRange.Start, formulaRange.End);
            service.SetSelectedFormulaFontSize(44);
            Release(formulaRange);
            formulaRange = target.Range;
            AssertInlineOleTypingAnchor(
                host.Document,
                target,
                targetMetadata!,
                expectedPosition,
                precedingFont);

            typingSelection = host.Application.Selection;
            typingSelection.SetRange(formulaRange.End, formulaRange.End);
            Release(typingSelection);
            typingSelection = null;
            WinForms.Application.DoEvents();
            Thread.Sleep(500);
            typingSelection = host.Application.Selection;
            var paragraphEndTypedStart = typingSelection.Start;
            const string paragraphEndTypedText = "typedatparagraphend";
            typingSelection.TypeText(paragraphEndTypedText);
            WinForms.Application.DoEvents();
            Thread.Sleep(100);
            typedRange = host.Document.Range(
                paragraphEndTypedStart,
                paragraphEndTypedStart + paragraphEndTypedText.Length);
            if (!string.Equals(
                    typedRange.Text,
                    paragraphEndTypedText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not insert the paragraph-end typing probe at the normalized OLE boundary. "
                    + $"Expected '{paragraphEndTypedText}', actual '{typedRange.Text}', "
                    + $"formulaEnd={formulaRange.End}, typedStart={paragraphEndTypedStart}.");
            typedFont = typedRange.Font;
            if (typedFont.Position != expectedPosition)
                throw new InvalidDataException(
                    "Typing after a paragraph-final resized inline OLE inherited the formula baseline. "
                    + $"Expected {expectedPosition}, actual {typedFont.Position}.");
            AssertBodyCharacterFormattingMatches(
                precedingFont,
                typedFont,
                "Typing after paragraph-final inline OLE");
        }
        finally
        {
            Release(typedFont);
            Release(typedRange);
            Release(inputWindow);
            Release(typingSelection);
            Release(formulaFont);
            Release(trailingFont);
            Release(precedingFont);
            Release(paragraphTail);
            Release(trailing);
            Release(preceding);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaRange);
            Release(target);
            Release(shapes);
        }
    }

    private static void AssertInlineOleTypingAnchor(
        Word.Document document,
        Word.InlineShape shape,
        FormulaMetadata metadata,
        int expectedPosition,
        Microsoft.Office.Interop.Word.Font expectedBodyFont)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.Range? anchorRange = null;
        Microsoft.Office.Interop.Word.Font? anchorFont = null;
        try
        {
            var bookmarkName = "VTBL_" + Guid.Parse(metadata.FormulaId).ToString("N");
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
                throw new InvalidDataException(
                    "Inline OLE typing anchor bookmark was not created.");
            bookmark = bookmarks[bookmarkName];
            anchorRange = bookmark.Range;
            formulaRange = shape.Range;
            if (anchorRange.Start != formulaRange.End
                || anchorRange.End != formulaRange.End + 1)
                throw new InvalidDataException(
                    "Inline OLE typing anchor is not immediately after the formula. "
                    + $"Formula={formulaRange.Start}:{formulaRange.End}, "
                    + $"anchor={anchorRange.Start}:{anchorRange.End}.");
            if (!string.Equals(anchorRange.Text, "\u200C", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Inline OLE typing anchor is not the expected zero-width non-joiner.");
            anchorFont = anchorRange.Font;
            if (anchorFont.Position != expectedPosition
                || anchorFont.Hidden != 0
                || anchorFont.Subscript != 0
                || anchorFont.Superscript != 0)
                throw new InvalidDataException(
                    "Inline OLE typing anchor does not carry ordinary body-text formatting. "
                    + $"Position={anchorFont.Position}, Hidden={anchorFont.Hidden}, "
                    + $"Subscript={anchorFont.Subscript}, Superscript={anchorFont.Superscript}.");
            AssertBodyCharacterFormattingMatches(
                expectedBodyFont,
                anchorFont,
                "Inline OLE typing anchor");
        }
        finally
        {
            Release(anchorFont);
            Release(anchorRange);
            Release(formulaRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static string DescribeCharacterFormatting(
        Word.Document document,
        int start,
        int end)
    {
        var builder = new StringBuilder();
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var safeStart = Math.Max(content.Start, start);
            var safeEnd = Math.Min(content.End, Math.Max(safeStart, end));
            for (var position = safeStart; position < safeEnd; position++)
            {
                Word.Range? characterRange = null;
                Microsoft.Office.Interop.Word.Font? characterFont = null;
                try
                {
                    characterRange = document.Range(position, position + 1);
                    characterFont = characterRange.Font;
                    var text = characterRange.Text ?? string.Empty;
                    var codePoint = text.Length == 0
                        ? "empty"
                        : $"U+{(int)text[0]:X4}";
                    if (builder.Length > 0) builder.Append(" | ");
                    builder.Append($"{position}:{codePoint}")
                        .Append($",pos={characterFont.Position}")
                        .Append($",name={characterFont.Name}")
                        .Append($",ascii={characterFont.NameAscii}")
                        .Append($",east={characterFont.NameFarEast}")
                        .Append($",size={characterFont.Size}")
                        .Append($",bold={characterFont.Bold}")
                        .Append($",italic={characterFont.Italic}");
                }
                finally
                {
                    Release(characterFont);
                    Release(characterRange);
                }
            }
        }
        finally { Release(content); }
        return builder.ToString();
    }

    private static void AssertBodyCharacterFormattingMatches(
        Microsoft.Office.Interop.Word.Font expected,
        Microsoft.Office.Interop.Word.Font actual,
        string stage)
    {
        static bool SameName(string? left, string? right) =>
            string.Equals(left ?? string.Empty, right ?? string.Empty,
                StringComparison.OrdinalIgnoreCase);

        var sameSize = Math.Abs(expected.Size - actual.Size) <= 0.1f;
        if (!SameName(expected.Name, actual.Name)
            || !SameName(expected.NameAscii, actual.NameAscii)
            || !SameName(expected.NameFarEast, actual.NameFarEast)
            || !sameSize
            || expected.Bold != actual.Bold
            || expected.Italic != actual.Italic)
        {
            throw new InvalidDataException(
                $"{stage} inherited a different body-text font. "
                + $"Expected Name={expected.Name}, ASCII={expected.NameAscii}, "
                + $"FarEast={expected.NameFarEast}, Size={expected.Size}, "
                + $"Bold={expected.Bold}, Italic={expected.Italic}; "
                + $"actual Name={actual.Name}, ASCII={actual.NameAscii}, "
                + $"FarEast={actual.NameFarEast}, Size={actual.Size}, "
                + $"Bold={actual.Bold}, Italic={actual.Italic}.");
        }
    }

    private static void AssertDisplayFormulaFollowedImmediatelyByText(
        Word.Document document,
        Word.Range formulaRange,
        string expectedText)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? nextAnchor = null;
        Word.Paragraphs? nextParagraphs = null;
        Word.Paragraph? nextParagraph = null;
        Word.Range? nextRange = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Display formula paragraph was not found.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            nextAnchor = document.Range(paragraphRange.End, paragraphRange.End);
            nextParagraphs = nextAnchor.Paragraphs;
            if (nextParagraphs.Count == 0)
                throw new InvalidDataException("Paragraph after the display formula was not found.");
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range;
            var nextText = nextRange.Text ?? string.Empty;
            if (nextText.IndexOf(expectedText, StringComparison.Ordinal) < 0)
                throw new InvalidDataException(
                    "Display formula introduced an empty paragraph before the following prose. "
                    + $"Expected next paragraph to contain '{expectedText}', actual='{nextText.Replace("\r", "<CR>").Replace("\v", "<BR>")}'.");
        }
        finally
        {
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextAnchor);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void AssertOleFontSize(
        IReadOnlyList<(Word.InlineShape Shape, FormulaMetadata Metadata)> formulas,
        string latex,
        double expected,
        string label)
    {
        var matches = formulas
            .Where(item => string.Equals(item.Metadata.Latex, latex, StringComparison.Ordinal))
            .ToArray();
        if (matches.Length != 1)
            throw new InvalidDataException(
                $"{label} matched {matches.Length} OLE formulas for LaTeX: {latex}");
        AssertNear(matches[0].Metadata.FontSizePt ?? 0, expected, label);
    }

    private static void AssertOmmlFontSize(Word.OMath math, double expected, int index)
    {
        Word.Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = math.Range;
            font = range.Font;
            AssertNear(font.Size, expected, $"OMML formula {index} font size");
        }
        finally
        {
            Release(font);
            Release(range);
            Release(math);
        }
    }

    private static void AssertNear(double actual, double expected, string label)
    {
        if (double.IsNaN(actual)
            || double.IsInfinity(actual)
            || Math.Abs(actual - expected) > 0.6)
            throw new InvalidDataException(
                $"{label} was {actual.ToString("0.##", CultureInfo.InvariantCulture)} pt; expected {expected:0.##} pt.");
    }

    private static void AssertLatexRedrawPerformance(string log, string modeName)
    {
        var timings = Regex.Matches(log, @"\brender index=\d+ elapsedMs=(?<ms>\d+)")
            .Cast<Match>()
            .Select(match => long.Parse(match.Groups["ms"].Value, CultureInfo.InvariantCulture))
            .ToArray();
        if (timings.Length != 4)
            throw new InvalidDataException(
                $"{modeName} redraw logged {timings.Length} render timings instead of 4.\n{log}");
        if (log.IndexOf("render-skipped", StringComparison.Ordinal) < 0
            || log.IndexOf("skipped=1", StringComparison.Ordinal) < 0)
            throw new InvalidDataException(
                $"{modeName} redraw did not report exactly one preserved invalid formula.\n{log}");
        var maximum = timings.Max();
        Console.WriteLine(
            $"[Word LaTeX redraw] {modeName} render timings: {string.Join(", ", timings)} ms; max={maximum} ms");
        if (maximum > LatexRedrawPerformanceLimitMilliseconds)
            throw new InvalidDataException(
                $"{modeName} formula redraw exceeded the {LatexRedrawPerformanceLimitMilliseconds} ms target: max={maximum} ms.\n{log}");
    }

    private static void TryDeleteAcceptanceFile(string path)
    {
        try { File.Delete(path); } catch { }
    }
}

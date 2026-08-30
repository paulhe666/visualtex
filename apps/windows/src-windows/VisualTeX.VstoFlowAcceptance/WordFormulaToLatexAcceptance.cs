using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using WinForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class FormulaToLatexCase
    {
        internal string FormulaId { get; set; } = string.Empty;
        internal string ObjectMode { get; set; } = string.Empty;
        internal string DisplayMode { get; set; } = string.Empty;
        internal string Latex { get; set; } = string.Empty;
        internal bool Numbered { get; set; }
    }

    private static void RunWordFormulaToLatex(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "VisualTeX-Word-Formula-To-Latex.docx");
        var logPath = Path.Combine(
            artifactRoot,
            "word-formula-to-latex.log");
        TryDeleteAcceptanceFile(documentPath);
        TryDeleteAcceptanceFile(logPath);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
            logPath);
        const string referenceBookmark = "VTTestEquationNumberReference";
        try
        {
            RunNativeWordOmmlToOleAcceptance(client);
            RunOmmlSelectionReadOnlyAcceptance();
            RunMixedOmmlThenOleToLatexAcceptance(client);
            using (var host = new WordPerformanceHost(documentPath: null))
            {
                WordEquationNumbering.SetEquationNumberFormatPreference(
                    host.Document,
                    EquationNumberFormat.ContinuousId);
                var formulas = PopulateFormulaToLatexDocument(client, host);
                var oleInline = formulas.Single(item =>
                    item.ObjectMode == FormulaOleContract.NativeOleMode
                    && item.DisplayMode == "inline");
                var ommlInline = formulas.Single(item =>
                    item.ObjectMode == FormulaOleContract.WordOmmlMode
                    && item.DisplayMode == "inline");
                var oleDisplay = formulas.Single(item =>
                    item.ObjectMode == FormulaOleContract.NativeOleMode
                    && item.DisplayMode == "block");
                AssertFormulaToLatexRollbackAfterInjectedFailure(
                    host,
                    oleInline,
                    expectedOle: 2,
                    expectedOmml: 2);
                AssertFormulaToLatexEmptyMetadataPreflight(
                    host,
                    oleInline,
                    expectedOle: 2,
                    expectedOmml: 2);

                var insertedReference = InsertNumberingReference(
                    host.Application,
                    host.Document,
                    oleDisplay.FormulaId);
                AssertEqual(referenceBookmark, insertedReference,
                    "The equation reference acceptance bookmark changed unexpectedly.");
                AppendFormulaToLatexDocumentEnd(host.Application);
                AssertReferenceText(host.Document, referenceBookmark, "(1)");
                AssertFormulaObjectCounts(host.Document, expectedOle: 2, expectedOmml: 2);

                MaterializeInlineOleTypingAnchor(host, oleInline);
                SelectFormula(host, oleInline);
                host.AddIn.OnRedrawSelectionOleToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 1);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFormulaObjectCounts(host.Document, expectedOle: 1, expectedOmml: 2);
                AssertDocumentContains(host.Document, "$a=1$");
                AssertDocumentDoesNotContain(host.Document, "$b=2$");

                SelectFormula(host, ommlInline);
                host.AddIn.OnRedrawSelectionOmmlToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 2);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFormulaObjectCounts(host.Document, expectedOle: 1, expectedOmml: 1);
                AssertDocumentContains(host.Document, "$b=2$");

                CollapseSelectionAtDocumentStart(host);
                host.AddIn.OnRedrawDocumentOleToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 3);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFormulaObjectCounts(host.Document, expectedOle: 0, expectedOmml: 1);
                AssertDocumentContains(host.Document, "$$c=3$$");
                AssertReferenceText(host.Document, referenceBookmark, "(1)");
                AssertNoBrokenReferenceText(host.Document);

                CollapseSelectionAtDocumentStart(host);
                host.AddIn.OnRedrawDocumentOmmlToLatex(new object());
                WaitForFormulaToLatex(logPath, expectedCompletions: 4);
                WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));
                AssertFinalFormulaToLatexDocument(host.Document, referenceBookmark);
                host.Save(documentPath);
            }

            using (var reopened = new WordPerformanceHost(documentPath))
            {
                AssertFinalFormulaToLatexDocument(
                    reopened.Document,
                    referenceBookmark);
            }

            Console.WriteLine(
                "Word formula-to-LaTeX acceptance passed: selection and whole-document Ribbon commands restored OLE/OMML formulas to source code independently, numbered tables were flattened, references were preserved as plain text, prose order survived, and the result persisted after save/reopen.");
            Console.WriteLine($"Artifact: {documentPath}");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG",
                null);
        }
    }

    private static void RunOmmlSelectionReadOnlyAcceptance()
    {
        using var host = new WordPerformanceHost(documentPath: null);
        Word.Range? equationRange = null;
        Word.Range? content = null;
        Word.Bookmarks? bookmarks = null;
        try
        {
            equationRange = InsertNativeWordOmml(
                host,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi>r</mi><mo>=</mo><mn>7</mn></math>",
                display: false);
            var beforeText = ReadFormulaToLatexDocumentText(host.Document);
            bookmarks = host.Document.Bookmarks;
            var beforeBookmarkCount = bookmarks.Count;
            var beforeCustomXmlCount = host.Document.CustomXMLParts.Count;
            AssertEqual(
                0,
                WordOmmlFormulaStore.FormulaIds(host.Document).Count,
                "The read-only OMML selection fixture unexpectedly started with VisualTeX metadata.");

            var caret = Math.Min(equationRange.End - 1, equationRange.Start + 1);
            host.Application.Selection.SetRange(caret, caret);
            var service = new WordFormulaService(host.Application);
            var size = service.GetSelectedFormulaFontSize();
            if (!size.HasValue)
                throw new InvalidDataException(
                    "Read-only OMML font-size probing did not recognize the native Word equation.");
            _ = service.GetSelectedFormulaFontSize();

            AssertEqual(
                beforeText,
                ReadFormulaToLatexDocumentText(host.Document),
                "Reading the selected OMML font size changed Word document text.");
            AssertEqual(
                beforeBookmarkCount,
                host.Document.Bookmarks.Count,
                "Reading the selected OMML font size created a Word bookmark.");
            AssertEqual(
                beforeCustomXmlCount,
                host.Document.CustomXMLParts.Count,
                "Reading the selected OMML font size created a CustomXML metadata part.");
            AssertEqual(
                0,
                WordOmmlFormulaStore.FormulaIds(host.Document).Count,
                "Reading the selected OMML font size adopted a native Word equation into VisualTeX metadata.");
            Console.WriteLine(
                "Native OMML selection/font-size probing remained strictly read-only.");
        }
        finally
        {
            Release(bookmarks);
            Release(content);
            Release(equationRange);
        }
    }

    private static void RunMixedOmmlThenOleToLatexAcceptance(
        VisualTeXSessionClient client)
    {
        using var host = new WordPerformanceHost(documentPath: null);
        Word.Selection? selection = null;
        Word.Tables? tables = null;
        Word.Frames? frames = null;
        try
        {
            selection = host.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory);
            var oleDisplay = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.NativeOleMode,
                displayMode: "block",
                numbered: true,
                latex: "u=v",
                mathVariable: "u",
                mathNumber: "1");
            selection.EndKey(Word.WdUnits.wdStory);
            var ommlDisplay = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.WordOmmlMode,
                displayMode: "block",
                numbered: true,
                latex: "p=2",
                mathVariable: "p",
                mathNumber: "2");
            selection.EndKey(Word.WdUnits.wdStory);
            var ommlInline = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.WordOmmlMode,
                displayMode: "inline",
                numbered: false,
                latex: "q=3",
                mathVariable: "q",
                mathNumber: "3");

            AssertFormulaObjectCounts(host.Document, expectedOle: 1, expectedOmml: 2);
            var service = new WordFormulaService(host.Application);
            var ommlResult = service.ConvertFormulaObjectsToLatex(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(2, ommlResult.FormulaCount,
                "The mixed-order fixture did not convert both OMML formulas first.");
            AssertFormulaObjectCounts(host.Document, expectedOle: 1, expectedOmml: 0);
            AssertDocumentContains(host.Document, "$$p=2$$");
            AssertLatexParagraphEscapedCompactStructuralSpacing(
                host.Document,
                "$$p=2$$");
            AssertDocumentContains(host.Document, "$q=3$");
            AssertDocumentDoesNotContain(host.Document, "$$u=v$$");

            var oleResult = service.ConvertFormulaObjectsToLatex(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode);
            AssertEqual(1, oleResult.FormulaCount,
                "The mixed-order fixture did not convert the remaining OLE formula.");
            AssertFormulaObjectCounts(host.Document, expectedOle: 0, expectedOmml: 0);
            frames = host.Document.Frames;
            AssertEqual(0, frames.Count,
                "Mixed OMML-then-OLE conversion left a hidden Word Frame around restored LaTeX source.");
            Release(frames);
            frames = null;
            var text = ReadFormulaToLatexDocumentText(host.Document);
            foreach (var expected in new[] { "$$u=v$$", "$$p=2$$", "$q=3$" })
            {
                var count = text.Split(
                    new[] { expected },
                    StringSplitOptions.None).Length - 1;
                AssertEqual(1, count,
                    $"Mixed OMML-then-OLE conversion did not preserve exactly one '{expected}' source.");
            }
            if (text.IndexOf("(1)", StringComparison.Ordinal) >= 0
                || text.IndexOf("(2)", StringComparison.Ordinal) >= 0
                || text.IndexOf("\r1\r", StringComparison.Ordinal) >= 0
                || text.IndexOf("\r2\r", StringComparison.Ordinal) >= 0)
            {
                throw new InvalidDataException(
                    "Mixed OMML-then-OLE conversion left equation-number residue after flattening numbered tables.\n"
                    + text.Replace("\r", "<CR>").Replace("\a", "<CELL>"));
            }
            tables = host.Document.Tables;
            AssertEqual(0, tables.Count,
                "Mixed OMML-then-OLE conversion left a numbered formula table behind.");
            Console.WriteLine(
                "Mixed numbered OLE + numbered OMML + inline OMML conversion passed in OMML-first/OLE-second order.");
        }
        finally
        {
            Release(frames);
            Release(tables);
            Release(selection);
        }
    }

    private static void RunNativeWordOmmlToOleAcceptance(
        VisualTeXSessionClient client)
    {
        using var host = new WordPerformanceHost(documentPath: null);
        Word.Selection? selection = null;
        Word.Range? nativeRange = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        try
        {
            selection = host.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory);
            selection.Font.Name = "宋体";
            selection.Font.Size = 10.5f;
            selection.TypeText("NATIVE_OLE_BEFORE ");
            nativeRange = InsertNativeWordOmml(
                host,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi>p</mi><mo>=</mo><mn>6</mn></math>",
                display: false);
            selection.SetRange(nativeRange.End, nativeRange.End);
            selection.MoveRight(
                Word.WdUnits.wdCharacter,
                1,
                Word.WdMovementType.wdMove);
            selection.TypeText(" NATIVE_OLE_AFTER");

            AssertEqual(
                0,
                WordOmmlFormulaStore.FormulaIds(host.Document).Count,
                "The native Word OMML fixture unexpectedly had VisualTeX metadata.");
            host.Application.Selection.SetRange(nativeRange.Start, nativeRange.End);
            var existing = SnapshotSessionIds();
            var final = WaitForDirectConversion(
                client,
                existing,
                "word",
                FormulaOleContract.NativeOleMode,
                () => host.AddIn.OnConvertSelected(new object()),
                TimeSpan.FromSeconds(45),
                out _);
            AssertEqual(
                "completed",
                final.Status,
                final.Error ?? "Native Word OMML-to-OLE conversion did not complete.");
            WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(20));

            shapes = host.Document.InlineShapes;
            AssertEqual(
                1,
                shapes.Count,
                "Native Word OMML-to-OLE conversion did not create exactly one OLE object.");
            shape = shapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(shape))
                throw new InvalidDataException(
                    "Native Word OMML conversion created the wrong OLE class.");
            var metadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException(
                    "Native Word OMML-to-OLE conversion did not attach VisualTeX metadata.");
            AssertEqual(
                "p=6",
                (metadata.Latex ?? string.Empty).Replace(" ", string.Empty),
                "Native Word OMML-to-OLE conversion recovered the wrong LaTeX source.");
            AssertDocumentContains(host.Document, "NATIVE_OLE_BEFORE");
            AssertDocumentContains(host.Document, "NATIVE_OLE_AFTER");
            Console.WriteLine(
                "Native Word OMML direct-to-OLE acceptance passed without pre-existing VisualTeX metadata.");
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(nativeRange);
            Release(selection);
        }
    }

    private static Word.Range InsertNativeWordOmml(
        WordPerformanceHost host,
        string mathMl,
        bool display)
    {
        Word.Selection? selection = null;
        Word.Range? insertion = null;
        Word.Range? equationRange = null;
        try
        {
            selection = host.Application.Selection;
            insertion = selection.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            equationRange = WordOmmlConverter.Insert(
                host.Application,
                host.Document,
                insertion,
                mathMl,
                display,
                sourceFingerprint: out _,
                includeLeadingTab: false);
            var result = equationRange.Duplicate;
            return result;
        }
        finally
        {
            Release(equationRange);
            Release(insertion);
            Release(selection);
        }
    }

    private static List<FormulaToLatexCase> PopulateFormulaToLatexDocument(
        VisualTeXSessionClient client,
        WordPerformanceHost host)
    {
        Word.Selection? selection = null;
        try
        {
            selection = host.Application.Selection;
            selection.HomeKey(Word.WdUnits.wdStory);
            selection.Font.Name = "宋体";
            selection.Font.Size = 10.5f;
            selection.TypeText("BEGIN ");
            var oleInline = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.NativeOleMode,
                displayMode: "inline",
                numbered: false,
                latex: "a=1",
                mathVariable: "a",
                mathNumber: "1");
            selection.TypeText(" BETWEEN ");
            var ommlInline = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.WordOmmlMode,
                displayMode: "inline",
                numbered: false,
                latex: "b=2",
                mathVariable: "b",
                mathNumber: "2");
            selection.TypeText(" INLINE_END");
            selection.TypeParagraph();

            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("BEFORE_OLE_DISPLAY");
            selection.TypeParagraph();
            var oleDisplay = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.NativeOleMode,
                displayMode: "block",
                numbered: true,
                latex: "c=3",
                mathVariable: "c",
                mathNumber: "3");
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("AFTER_OLE_DISPLAY");
            selection.TypeParagraph();

            var ommlDisplay = InsertFormulaToLatexCase(
                client,
                host,
                FormulaOleContract.WordOmmlMode,
                displayMode: "block",
                numbered: true,
                latex: "d=4",
                mathVariable: "d",
                mathNumber: "4");
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("AFTER_OMML_DISPLAY");
            selection.TypeParagraph();

            selection.TypeText("BEFORE_NATIVE_OMML ");
            InsertNativeWordOmml(
                host,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi>n</mi><mo>=</mo><mn>5</mn></math>",
                display: false);
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText(" AFTER_NATIVE_OMML");
            selection.TypeParagraph();

            return new List<FormulaToLatexCase>
            {
                oleInline,
                ommlInline,
                oleDisplay,
                ommlDisplay,
            };
        }
        finally { Release(selection); }
    }

    private static FormulaToLatexCase InsertFormulaToLatexCase(
        VisualTeXSessionClient client,
        WordPerformanceHost host,
        string objectMode,
        string displayMode,
        bool numbered,
        string latex,
        string mathVariable,
        string mathNumber)
    {
        var existing = SnapshotSessionIds();
        var display = string.Equals(displayMode, "block", StringComparison.Ordinal);
        if (string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
        {
            if (display) host.AddIn.OnInsertDisplayOmml(new object());
            else host.AddIn.OnInsertInlineOmml(new object());
        }
        else
        {
            if (display) host.AddIn.OnInsertDisplay(new object());
            else host.AddIn.OnInsertInline(new object());
        }

        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        var mathMl = string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal)
            ? $"<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"{(display ? "block" : "inline")}\"><mi>{mathVariable}</mi><mo>=</mo><mn>{mathNumber}</mn></math>"
            : null;
        Commit(
            client,
            session,
            displayMode,
            objectMode,
            latex,
            numbered: numbered,
            mathMl: mathMl);
        var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
        AssertEqual("completed", final.Status,
            final.Error ?? "The formula-to-LaTeX source formula did not complete.");
        client.CloseEditorAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(15));
        return new FormulaToLatexCase
        {
            FormulaId = final.FormulaId
                ?? throw new InvalidDataException(
                    "The formula-to-LaTeX source formula has no formulaId."),
            ObjectMode = objectMode,
            DisplayMode = displayMode,
            Latex = latex,
            Numbered = numbered,
        };
    }

    private static void AppendToFormulaOmmlAndSelect(
        Word.Document document,
        FormulaToLatexCase formula,
        string suffix)
    {
        Word.Range? range = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? insertion = null;
        Word.Range? selectedRange = null;
        try
        {
            range = ResolveFormulaToLatexRange(document, formula);
            maths = range.OMaths;
            if (maths.Count == 0)
                throw new InvalidDataException(
                    $"OMML formula {formula.FormulaId} could not be opened for native editing.");
            math = maths[1];
            math.Linearize();
            insertion = math.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertBefore(suffix);
            math.BuildUp();
            selectedRange = math.Range;
            selectedRange.Select();
        }
        finally
        {
            Release(selectedRange);
            Release(insertion);
            Release(math);
            Release(maths);
            Release(range);
        }
    }

    private static void AppendFormulaToLatexDocumentEnd(Word.Application application)
    {
        Word.Selection? selection = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("DOCUMENT_END");
        }
        finally { Release(selection); }
    }

    private static void AssertFormulaToLatexRollbackAfterInjectedFailure(
        WordPerformanceHost host,
        FormulaToLatexCase formula,
        int expectedOle,
        int expectedOmml)
    {
        var beforeText = ReadFormulaToLatexDocumentText(host.Document);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMULA_TO_LATEX_FAIL_AFTER_DELETE",
            formula.FormulaId);
        try
        {
            SelectFormula(host, formula);
            var service = new WordFormulaService(host.Application);
            try
            {
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: false,
                    FormulaOleContract.NativeOleMode);
                throw new InvalidDataException(
                    "Injected formula-to-LaTeX failure did not interrupt conversion.");
            }
            catch (InvalidOperationException error)
                when (error.Message.IndexOf(
                    "Injected formula-to-LaTeX failure",
                    StringComparison.Ordinal) >= 0)
            {
                // Expected. The service must restore the deleted formula before
                // surfacing the conversion error to its caller.
            }
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMULA_TO_LATEX_FAIL_AFTER_DELETE",
                null);
        }

        AssertEqual(
            beforeText,
            ReadFormulaToLatexDocumentText(host.Document),
            "Formula-to-LaTeX rollback changed document text after an injected post-delete failure.");
        AssertFormulaObjectCounts(host.Document, expectedOle, expectedOmml);
        Word.Range? restored = null;
        try
        {
            restored = ResolveFormulaToLatexRange(host.Document, formula);
            if (restored.Start >= restored.End)
                throw new InvalidDataException(
                    "Formula-to-LaTeX rollback restored a collapsed formula range.");
        }
        finally { Release(restored); }
        Console.WriteLine(
            "Formula-to-LaTeX atomic rollback acceptance passed after an injected post-delete failure.");
    }

    private static void AssertFormulaToLatexEmptyMetadataPreflight(
        WordPerformanceHost host,
        FormulaToLatexCase formula,
        int expectedOle,
        int expectedOmml)
    {
        var beforeText = ReadFormulaToLatexDocumentText(host.Document);
        Environment.SetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMULA_TO_LATEX_EMPTY_SOURCE",
            formula.FormulaId);
        try
        {
            SelectFormula(host, formula);
            var service = new WordFormulaService(host.Application);
            try
            {
                service.ConvertFormulaObjectsToLatex(
                    wholeDocument: false,
                    FormulaOleContract.NativeOleMode);
                throw new InvalidDataException(
                    "Formula-to-LaTeX accepted empty source metadata and deleted a visible formula.");
            }
            catch (InvalidDataException error)
                when (error.Message.IndexOf(
                    "LaTeX 元数据为空",
                    StringComparison.Ordinal) >= 0)
            {
                // Expected: preflight must fail before the destructive undo record.
            }
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMULA_TO_LATEX_EMPTY_SOURCE",
                null);
        }

        AssertEqual(
            beforeText,
            ReadFormulaToLatexDocumentText(host.Document),
            "Empty formula metadata changed document text before conversion was rejected.");
        AssertFormulaObjectCounts(host.Document, expectedOle, expectedOmml);
        Word.Range? restored = null;
        try
        {
            restored = ResolveFormulaToLatexRange(host.Document, formula);
            if (restored.Start >= restored.End)
                throw new InvalidDataException(
                    "Empty-source preflight removed or collapsed the visible OLE formula.");
        }
        finally { Release(restored); }
        Console.WriteLine(
            "Formula-to-LaTeX empty-source preflight acceptance passed without deleting the visible formula.");
    }

    private static void MaterializeInlineOleTypingAnchor(
        WordPerformanceHost host,
        FormulaToLatexCase formula)
    {
        Word.Range? range = null;
        Word.Selection? selection = null;
        Word.Range? content = null;
        try
        {
            range = ResolveFormulaToLatexRange(host.Document, formula);
            selection = host.Application.Selection;
            selection.SetRange(range.End, range.End);
            var service = new WordFormulaService(host.Application);
            service.NormalizeTypingCaretAfterInlineFormula(selection);
            content = host.Document.Content;
            var text = content.Text ?? string.Empty;
            var anchorCount = text.Count(character => character == '\u200C');
            AssertEqual(
                1,
                anchorCount,
                "Clicking the inline OLE boundary did not create exactly one zero-width typing anchor.");
        }
        finally
        {
            Release(content);
            Release(selection);
            Release(range);
        }
    }

    private static void SelectFormula(
        WordPerformanceHost host,
        FormulaToLatexCase formula)
    {
        Word.Range? range = null;
        try
        {
            range = ResolveFormulaToLatexRange(host.Document, formula);
            host.Application.Selection.SetRange(range.Start, range.End);
        }
        finally { Release(range); }
    }

    private static Word.Range ResolveFormulaToLatexRange(
        Word.Document document,
        FormulaToLatexCase formula)
    {
        if (string.Equals(
                formula.ObjectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? range = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    formula.FormulaId)
                    ?? throw new InvalidDataException(
                        $"OMML formula {formula.FormulaId} is missing.");
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                return range.Duplicate;
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }

        Word.InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                try
                {
                    shape = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formula.FormulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    range = shape.Range;
                    return range.Duplicate;
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
        }
        finally { Release(shapes); }
        throw new InvalidDataException(
            $"OLE formula {formula.FormulaId} is missing.");
    }

    private static void CollapseSelectionAtDocumentStart(WordPerformanceHost host)
    {
        Word.Range? content = null;
        try
        {
            content = host.Document.Content;
            host.Application.Selection.SetRange(content.Start, content.Start);
        }
        finally { Release(content); }
    }

    private static string WaitForFormulaToLatex(
        string logPath,
        int expectedCompletions)
    {
        var deadline = DateTime.UtcNow + TimeSpan.FromMinutes(2);
        var last = string.Empty;
        while (DateTime.UtcNow < deadline)
        {
            WinForms.Application.DoEvents();
            Thread.Sleep(25);
            try
            {
                if (!File.Exists(logPath)) continue;
                last = File.ReadAllText(logPath, Encoding.UTF8);
                if (last.IndexOf(
                        "formula-to-latex-failed",
                        StringComparison.Ordinal) >= 0)
                    throw new InvalidDataException(
                        "Word formula-to-LaTeX command failed.\n" + last);
                var completed = last.Split(new[] { "formula-to-latex-complete" },
                    StringSplitOptions.None).Length - 1;
                if (completed >= expectedCompletions) return last;
            }
            catch (IOException)
            {
                // The add-in may be appending the current log line.
            }
        }
        throw new TimeoutException(
            $"Word formula-to-LaTeX command did not complete. Expected {expectedCompletions} completion records. Last log:\n{last}");
    }

    private static void AssertLatexParagraphEscapedCompactStructuralSpacing(
        Word.Document document,
        string latexSource)
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
            AssertTrue(find.Execute(),
                $"Formula-to-LaTeX paragraph '{latexSource}' was not found.");
            paragraphs = search.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                $"Formula-to-LaTeX source '{latexSource}' spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraphRange.ParagraphFormat;
            AssertTrue(
                format.LineSpacingRule != Word.WdLineSpacing.wdLineSpaceExactly
                || format.LineSpacing > 2.01f,
                $"Visible LaTeX source '{latexSource}' inherited VisualTeX's compact 1pt structural line box.");
            Console.WriteLine(
                $"  formula-to-LaTeX line-height '{latexSource}': rule={format.LineSpacingRule} line={format.LineSpacing:0.###}pt font={paragraphRange.Font.Size:0.###}pt.");
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

    private static void AssertFormulaObjectCounts(
        Word.Document document,
        int expectedOle,
        int expectedOmml)
    {
        Word.InlineShapes? shapes = null;
        var oleCount = 0;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (WordFormulaMetadataReader.IsNativeOle(shape)
                        && WordFormulaMetadataReader.TryRead(shape) is not null)
                        oleCount++;
                }
                finally { Release(shape); }
            }
        }
        finally { Release(shapes); }
        var ommlCount = WordOmmlFormulaStore.FormulaIds(document).Count;
        AssertEqual(expectedOle, oleCount,
            "The formula-to-LaTeX command converted the wrong number of OLE formulas.");
        AssertEqual(expectedOmml, ommlCount,
            "The formula-to-LaTeX command converted the wrong number of OMML formulas.");
    }

    private static void AssertFinalFormulaToLatexDocument(
        Word.Document document,
        string referenceBookmark)
    {
        AssertFormulaObjectCounts(document, expectedOle: 0, expectedOmml: 0);
        foreach (var required in new[]
                 {
                     "BEGIN",
                     "$a=1$",
                     "BETWEEN",
                     "$b=2$",
                     "INLINE_END",
                     "BEFORE_OLE_DISPLAY",
                     "$$c=3$$",
                     "AFTER_OLE_DISPLAY",
                     "$$d=4$$",
                     "AFTER_OMML_DISPLAY",
                     "BEFORE_NATIVE_OMML",
                     "$n=5$",
                     "AFTER_NATIVE_OMML",
                     "Reference:",
                     "DOCUMENT_END",
                 })
            AssertDocumentContains(document, required);

        var text = ReadFormulaToLatexDocumentText(document);
        if (text.IndexOf('\u200C') >= 0)
        {
            var positions = Enumerable.Range(0, text.Length)
                .Where(index => text[index] == '\u200C')
                .ToArray();
            var escaped = text
                .Replace("\u200C", "<ZWNJ>")
                .Replace("\r", "<CR>")
                .Replace("\a", "<CELL>");
            throw new InvalidDataException(
                "Formula-to-LaTeX conversion left inline OLE typing anchors at "
                + $"positions [{string.Join(",", positions)}]. Document={escaped}");
        }
        var ordered = new[]
        {
            "BEGIN",
            "$a=1$",
            "BETWEEN",
            "$b=2$",
            "INLINE_END",
            "BEFORE_OLE_DISPLAY",
            "$$c=3$$",
            "AFTER_OLE_DISPLAY",
            "$$d=4$$",
            "AFTER_OMML_DISPLAY",
            "BEFORE_NATIVE_OMML",
            "$n=5$",
            "AFTER_NATIVE_OMML",
            "Reference:",
            "DOCUMENT_END",
        };
        var previous = -1;
        foreach (var marker in ordered)
        {
            var position = text.IndexOf(marker, StringComparison.Ordinal);
            if (position <= previous)
                throw new InvalidDataException(
                    $"Formula-to-LaTeX document order is wrong at '{marker}'.\n{text.Replace("\r", "<CR>").Replace("\a", "<CELL>")}");
            previous = position;
        }

        Word.Tables? tables = null;
        try
        {
            tables = document.Tables;
            AssertEqual(0, tables.Count,
                "Numbered formula tables remained after all formulas were restored to LaTeX.");
        }
        finally { Release(tables); }
        AssertReferenceText(document, referenceBookmark, "(1)");
        AssertNoBrokenReferenceText(document);
    }

    private static void AssertDocumentContains(
        Word.Document document,
        string expected)
    {
        var text = ReadFormulaToLatexDocumentText(document);
        if (text.IndexOf(expected, StringComparison.Ordinal) < 0)
            throw new InvalidDataException(
                $"Formula-to-LaTeX document is missing '{expected}'.\n{text.Replace("\r", "<CR>").Replace("\a", "<CELL>")}");
    }

    private static void AssertDocumentDoesNotContain(
        Word.Document document,
        string forbidden)
    {
        var text = ReadFormulaToLatexDocumentText(document);
        if (text.IndexOf(forbidden, StringComparison.Ordinal) >= 0)
            throw new InvalidDataException(
                $"Formula-to-LaTeX document unexpectedly contains '{forbidden}'.");
    }

    private static void AssertNoBrokenReferenceText(Word.Document document)
    {
        var text = ReadFormulaToLatexDocumentText(document);
        foreach (var forbidden in new[]
                 {
                     "Error! Reference source not found.",
                     "错误! 未找到引用源。",
                     "错误！未找到引用源。",
                 })
        {
            if (text.IndexOf(forbidden, StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidDataException(
                    $"Formula-to-LaTeX conversion left a broken equation reference: {forbidden}");
        }
    }

    private static string ReadFormulaToLatexDocumentText(Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            return content.Text ?? string.Empty;
        }
        finally { Release(content); }
    }
}

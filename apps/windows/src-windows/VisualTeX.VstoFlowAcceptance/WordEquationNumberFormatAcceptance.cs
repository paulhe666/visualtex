using Extensibility;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class EquationNumberRibbonControl : Office.IRibbonControl
    {
        internal EquationNumberRibbonControl(string tag)
        {
            Id = "VisualTeX.WordVsto.NumberFormatAcceptance";
            Tag = tag;
        }

        public string Id { get; }
        public object Context => null!;
        public string Tag { get; }
    }

    private sealed class NumberedFormulaCase
    {
        internal string FormulaId { get; set; } = string.Empty;
        internal string ObjectMode { get; set; } = string.Empty;
    }

    private static void RunWordEquationNumberFormat(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(artifactRoot, "word-equation-number-format.docx");
        DeleteBulkPerformanceArtifact(outputPath);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        Word.Document? freshDocument = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        var previousDefaultFormat = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        var previousDefaultNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        try
        {
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(
                EquationNumberFormat.ContinuousId);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            InsertNumberingHeading(application, document, level: 1, "Chapter A");
            var formulas = new List<NumberedFormulaCase>
            {
                InsertNumberedFormula(client, application, addIn, FormulaOleContract.NativeOleMode, "a_1=1"),
                InsertNumberedFormula(client, application, addIn, FormulaOleContract.WordOmmlMode, "a_2=2"),
            };
            InsertNumberingHeading(application, document, level: 2, "Section A.1");
            formulas.Add(InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.NativeOleMode,
                "a_3=3"));

            InsertNumberingHeading(application, document, level: 1, "Chapter B");
            formulas.Add(InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "b_1=4"));
            InsertNumberingHeading(application, document, level: 2, "Section B.1");
            formulas.Add(InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.NativeOleMode,
                "b_2=5"));

            AssertEquationNumberFormat(
                document,
                addIn,
                EquationNumberFormat.ContinuousId,
                formulas,
                new[] { "1", "2", "3", "4", "5" });

            var referenceBookmark = InsertNumberingReference(
                application,
                document,
                formulas[0].FormulaId);
            var legacyReferenceBookmark = InsertPlainNumberingReference(
                application,
                document,
                formulas[0].FormulaId);
            AssertReferenceText(document, referenceBookmark, "(1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1)");

            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading1DotId,
                formulas,
                new[] { "1.1", "1.2", "1.3", "2.1", "2.2" });
            AssertReferenceText(document, referenceBookmark, "(1.1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1.1)");

            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading1DashId,
                formulas,
                new[] { "1-1", "1-2", "1-3", "2-1", "2-2" });
            AssertReferenceText(document, referenceBookmark, "(1-1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1-1)");

            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading2DotId,
                formulas,
                new[] { "1.0.1", "1.0.2", "1.1.1", "2.0.1", "2.1.1" });
            AssertReferenceText(document, referenceBookmark, "(1.0.1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1.0.1)");

            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading2DashId,
                formulas,
                new[] { "1.0-1", "1.0-2", "1.1-1", "2.0-1", "2.1-1" });
            AssertReferenceText(document, referenceBookmark, "(1.0-1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1.0-1)");

            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading1DashId,
                formulas,
                new[] { "1-1", "1-2", "1-3", "2-1", "2-2" });
            var futureFormula = InsertNumberedFormula(
                client,
                application,
                addIn,
                FormulaOleContract.WordOmmlMode,
                "b_3=6");
            formulas.Add(futureFormula);
            AssertEquationNumberFormat(
                document,
                addIn,
                EquationNumberFormat.Heading1DashId,
                formulas,
                new[] { "1-1", "1-2", "1-3", "2-1", "2-2", "2-3" });
            AssertReferenceText(document, referenceBookmark, "(1-1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1-1)");

            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading2DotId,
                formulas,
                new[] { "1.0.1", "1.0.2", "1.1.1", "2.0.1", "2.1.1", "2.1.2" });
            ApplyEquationNumberFormat(
                addIn,
                document,
                EquationNumberFormat.Heading1DashId,
                formulas,
                new[] { "1-1", "1-2", "1-3", "2-1", "2-2", "2-3" });
            AssertReferenceText(document, referenceBookmark, "(1-1)");
            AssertReferenceText(document, legacyReferenceBookmark, "(1-1)");

            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEquationNumberFormat(
                reopened,
                addIn,
                EquationNumberFormat.Heading1DashId,
                formulas,
                new[] { "1-1", "1-2", "1-3", "2-1", "2-2", "2-3" });
            AssertReferenceText(reopened, referenceBookmark, "(1-1)");
            AssertReferenceText(reopened, legacyReferenceBookmark, "(1-1)");
            AssertEqual(
                EquationNumberFormat.Heading1DashId,
                WordEquationNumbering.GetDefaultEquationNumberFormatId(),
                "The user-level equation-number format did not remember the last Ribbon selection.");

            freshDocument = application.Documents.Add();
            AssertTrue(
                addIn.GetEquationNumberFormatPressed(
                    new EquationNumberRibbonControl(EquationNumberFormat.Heading1DashId)),
                "A fresh Word document did not inherit the remembered user-level number format.");
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(true);
            AssertTrue(
                WordEquationNumbering.GetDefaultDisplayEquationNumbered(),
                "The remembered display-equation numbering checkbox did not persist true.");
            AssertNewDisplaySessionNumberedPreference(
                client,
                addIn,
                expectedNumbered: true);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);
            AssertTrue(
                !WordEquationNumbering.GetDefaultDisplayEquationNumbered(),
                "The remembered display-equation numbering checkbox did not persist false.");
            AssertNewDisplaySessionNumberedPreference(
                client,
                addIn,
                expectedNumbered: false);

            Console.WriteLine(
                "Word equation-number format acceptance passed: Ribbon selection changed all existing OLE/OMML numbers immediately, native cross-references followed the selected format, the document setting survived save/reopen, a fresh document inherited the user-level default, and the numbered-display preference round-tripped through persistent user storage.");
            Console.WriteLine($"Artifact: {outputPath}");
        }
        finally
        {
            if (addIn is not null)
            {
                try
                {
                    addIn.OnDisconnection(
                        ext_DisconnectMode.ext_dm_UserClosed,
                        ref custom);
                }
                catch { }
            }
            try { freshDocument?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(freshDocument);
            Release(reopened);
            Release(document);
            Release(application);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previousDefaultFormat);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousDefaultNumbered);
            ForceComCleanup();
        }
    }

    private static void InsertNumberingHeading(
        Word.Application application,
        Word.Document document,
        int level,
        string text)
    {
        Word.Selection? selection = null;
        Word.Range? headingRange = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            var start = selection.Start;
            selection.TypeText(text);
            selection.TypeParagraph();
            headingRange = document.Range(start, selection.Start);
            object headingStyle = level == 1
                ? Word.WdBuiltinStyle.wdStyleHeading1
                : Word.WdBuiltinStyle.wdStyleHeading2;
            headingRange.set_Style(ref headingStyle);
            object normalStyle = Word.WdBuiltinStyle.wdStyleNormal;
            selection.set_Style(ref normalStyle);
        }
        finally
        {
            Release(headingRange);
            Release(selection);
        }
    }

    private static void AssertNewDisplaySessionNumberedPreference(
        VisualTeXSessionClient client,
        VisualTeX.WordVsto.ThisAddIn addIn,
        bool expectedNumbered)
    {
        var existing = SnapshotSessionIds();
        addIn.OnInsertDisplay(new object());
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        try
        {
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual("block", session.DisplayMode,
                "The display-equation preference probe did not create a block session.");
            AssertEqual(expectedNumbered, session.Numbered,
                $"A new display editor session did not inherit numbered={expectedNumbered}.");
            client.PatchAsync(
                    sessionId,
                    new { status = "cancelled", explicitCancel = true },
                    CancellationToken.None)
                .GetAwaiter().GetResult();
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(30));
            AssertEqual("cancelled", final.Status,
                final.Error ?? "The display-equation preference probe did not cancel cleanly.");
        }
        finally
        {
            try
            {
                client.CloseEditorAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
            }
            catch { }
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
        }
    }

    private static NumberedFormulaCase InsertNumberedFormula(
        VisualTeXSessionClient client,
        Word.Application application,
        VisualTeX.WordVsto.ThisAddIn addIn,
        string objectMode,
        string latex,
        bool numbered = true)
    {
        const string mathMlPrefix =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">";
        const string mathMlSuffix = "</math>";
        var existing = SnapshotSessionIds();
        if (string.Equals(objectMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal))
            addIn.OnInsertDisplayOmml(new object());
        else
            addIn.OnInsertDisplay(new object());
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        var mathMl = string.Equals(objectMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal)
            ? mathMlPrefix + "<mi>x</mi><mo>=</mo><mn>1</mn>" + mathMlSuffix
            : null;
        Commit(
            client,
            session,
            "block",
            objectMode,
            latex,
            numbered: numbered,
            mathMl: mathMl);
        var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
        AssertEqual("completed", final.Status,
            final.Error ?? "The numbered formula did not complete.");
        client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
        WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
        return new NumberedFormulaCase
        {
            FormulaId = final.FormulaId
                ?? throw new InvalidDataException("The numbered formula has no formulaId."),
            ObjectMode = objectMode,
        };
    }

    private static string InsertNumberingReference(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Selection? selection = null;
        Word.Range? referenceRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("Reference: ");
            var start = selection.Start;
            var target = WordEquationNumbering.GetEquationReferenceTargets(document)
                .Single(item => string.Equals(
                    item.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase));
            WordEquationNumbering.InsertEquationReference(
                document,
                selection,
                target,
                EquationReferenceStyle.Parenthesized);
            var end = selection.Start;
            selection.TypeParagraph();
            referenceRange = document.Range(start, end);
            bookmarks = document.Bookmarks;
            const string bookmarkName = "VTTestEquationNumberReference";
            bookmark = bookmarks.Add(bookmarkName, referenceRange);
            return bookmarkName;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(referenceRange);
            Release(selection);
        }
    }

    private static string InsertPlainNumberingReference(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Selection? selection = null;
        Word.Range? fieldInsertion = null;
        Word.Field? reference = null;
        Word.Range? referenceResult = null;
        Word.Range? referenceRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("Plain REF reference: ");
            var start = selection.Start;
            selection.TypeText("(");
            fieldInsertion = document.Range(selection.Start, selection.Start);
            reference = document.Fields.Add(
                fieldInsertion,
                Word.WdFieldType.wdFieldEmpty,
                $"REF {WordEquationNumbering.NativeNumberBookmarkName(formulaId)} \\h",
                PreserveFormatting: true);
            reference.Update();
            referenceResult = reference.Result;
            selection.SetRange(referenceResult.End, referenceResult.End);
            selection.TypeText(")");
            var end = selection.Start;
            selection.TypeParagraph();
            referenceRange = document.Range(start, end);
            bookmarks = document.Bookmarks;
            const string bookmarkName = "VTTestPlainEquationNumberReference";
            bookmark = bookmarks.Add(bookmarkName, referenceRange);
            return bookmarkName;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(referenceRange);
            Release(referenceResult);
            Release(reference);
            Release(fieldInsertion);
            Release(selection);
        }
    }

    private static void ApplyEquationNumberFormat(
        VisualTeX.WordVsto.ThisAddIn addIn,
        Word.Document document,
        string formatId,
        IReadOnlyList<NumberedFormulaCase> formulas,
        IReadOnlyList<string> expectedNumbers)
    {
        var control = new EquationNumberRibbonControl(formatId);
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        addIn.OnEquationNumberFormatChanged(control, pressed: true);
        var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(30);
        Exception? lastError = null;
        while (DateTime.UtcNow < deadline)
        {
            try
            {
                AssertEquationNumberFormat(
                    document,
                    addIn,
                    formatId,
                    formulas,
                    expectedNumbers);
                stopwatch.Stop();
                Console.WriteLine(
                    $"  Equation-number format {formatId}: {formulas.Count} formulas in {stopwatch.ElapsedMilliseconds} ms.");
                return;
            }
            catch (Exception error)
            {
                lastError = error;
                Thread.Sleep(100);
            }
        }
        throw new InvalidOperationException(
            $"The Ribbon did not apply equation-number format {formatId} in time.",
            lastError);
    }

    private static void AssertEquationNumberFormat(
        Word.Document document,
        VisualTeX.WordVsto.ThisAddIn addIn,
        string formatId,
        IReadOnlyList<NumberedFormulaCase> formulas,
        IReadOnlyList<string> expectedNumbers)
    {
        AssertEqual(formulas.Count, expectedNumbers.Count,
            "The equation-number acceptance data is inconsistent.");
        var selectedControl = new EquationNumberRibbonControl(formatId);
        AssertTrue(addIn.GetEquationNumberFormatPressed(selectedControl),
            $"Ribbon format {formatId} is not selected.");
        foreach (var alternate in new[]
                 {
                     EquationNumberFormat.ContinuousId,
                     EquationNumberFormat.Heading1DotId,
                     EquationNumberFormat.Heading1DashId,
                     EquationNumberFormat.Heading2DotId,
                     EquationNumberFormat.Heading2DashId,
                 })
        {
            var pressed = addIn.GetEquationNumberFormatPressed(
                new EquationNumberRibbonControl(alternate));
            AssertEqual(string.Equals(alternate, formatId, StringComparison.Ordinal), pressed,
                $"Ribbon pressed state is wrong for {alternate}.");
        }

        for (var index = 0; index < formulas.Count; index++)
        {
            var visible = ReadEquationNumberBookmarkText(
                document,
                WordEquationNumbering.EquationBookmarkName(formulas[index].FormulaId));
            var native = ReadEquationNumberBookmarkText(
                document,
                WordEquationNumbering.NativeNumberBookmarkName(formulas[index].FormulaId));
            if (!string.Equals(expectedNumbers[index], visible, StringComparison.Ordinal)
                || !string.Equals(expectedNumbers[index], native, StringComparison.Ordinal))
            {
                DumpEquationSequenceFieldInventory(
                    document,
                    $"format={formatId} formulaIndex={index + 1} expected={expectedNumbers[index]} visible={visible} native={native}");
            }
            AssertEqual(expectedNumbers[index], visible,
                $"Visible equation number {index + 1} is wrong for format {formatId}.");
            AssertEqual(expectedNumbers[index], native,
                $"Native caption number {index + 1} is wrong for format {formatId}.");
        }
    }

    private static void DumpEquationSequenceFieldInventory(
        Word.Document document,
        string context)
    {
        Word.Fields? fields = null;
        Word.Bookmarks? bookmarks = null;
        try
        {
            fields = document.Fields;
            bookmarks = document.Bookmarks;
            Console.WriteLine(
                $"  [SEQ inventory] {context}; documentFields={fields.Count}, bookmarks={bookmarks.Count}.");
            for (var fieldIndex = 1; fieldIndex <= fields.Count; fieldIndex++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    field = fields[fieldIndex];
                    code = field.Code;
                    var codeText = code.Text ?? string.Empty;
                    if (codeText.IndexOf(
                            "SEQ VisualTeXEquation",
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    result = field.Result;
                    var aliases = new List<string>();
                    for (var bookmarkIndex = 1; bookmarkIndex <= bookmarks.Count; bookmarkIndex++)
                    {
                        Word.Bookmark? bookmark = null;
                        Word.Range? bookmarkRange = null;
                        try
                        {
                            bookmark = bookmarks[bookmarkIndex];
                            bookmarkRange = bookmark.Range;
                            if (bookmarkRange.StoryType == result.StoryType
                                && bookmarkRange.Start <= result.Start
                                && bookmarkRange.End >= result.End)
                                aliases.Add(bookmark.Name ?? string.Empty);
                        }
                        catch { }
                        finally
                        {
                            Release(bookmarkRange);
                            Release(bookmark);
                        }
                    }
                    Console.WriteLine(
                        $"    SEQ#{fieldIndex}: story={result.StoryType}, code={code.Start}:{code.End} '{codeText.Trim()}', result={result.Start}:{result.End} '{NormalizeEquationNumberText(result.Text)}', aliases=[{string.Join(",", aliases)}]");
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally
        {
            Release(bookmarks);
            Release(fields);
        }
    }

    private static string ReadEquationNumberBookmarkText(
        Word.Document document,
        string bookmarkName)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(bookmarkName),
                $"Equation-number bookmark {bookmarkName} is missing.");
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            return (range.Text ?? string.Empty)
                .Replace("\t", string.Empty)
                .Trim()
                .TrimStart('(', '（')
                .TrimEnd(')', '）');
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AssertReferenceText(
        Word.Document document,
        string bookmarkName,
        string expected)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(bookmarkName),
                "The equation reference bookmark is missing.");
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            var actual = ReadRenderedEquationReferenceText(document, range);
            if (!string.Equals(expected, actual, StringComparison.Ordinal)
                && expected.StartsWith("(", StringComparison.Ordinal)
                && actual.EndsWith(")", StringComparison.Ordinal)
                && range.Start > document.Content.Start)
            {
                Word.Range? expanded = null;
                try
                {
                    // Word can shrink a test-only bookmark around a field so the
                    // literal opening parenthesis sits immediately outside its
                    // start. Validate the rendered field result rather than the
                    // bookmark-boundary movement.
                    expanded = document.Range(range.Start - 1, range.End);
                    var expandedText = ReadRenderedEquationReferenceText(
                        document,
                        expanded);
                    if (expandedText.StartsWith("(", StringComparison.Ordinal))
                        actual = expandedText;
                }
                finally { Release(expanded); }
            }
            AssertEqual(expected, actual,
                "The native equation cross-reference did not follow the selected format.");
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static string ReadRenderedEquationReferenceText(
        Word.Document document,
        Word.Range range)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Fields? nestedFields = null;
        Word.Field? nested = null;
        Word.Range? nestedResult = null;
        try
        {
            var raw = range.Text ?? string.Empty;
            var visibleReference = string.Empty;
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (code.End < range.Start || code.Start > range.End) continue;
                var instruction = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (!instruction.StartsWith(
                        "GOTOBUTTON ",
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                Release(nestedFields);
                nestedFields = code.Fields;
                if (nestedFields.Count != 1)
                    throw new InvalidDataException(
                        "The navigable equation reference does not contain exactly one nested REF field.");
                Release(nested);
                nested = nestedFields[1];
                Release(nestedResult);
                nestedResult = nested.Result;
                visibleReference = (nestedResult.Text ?? string.Empty).Trim();
                break;
            }

            if (visibleReference.Length == 0)
                return raw.Trim();

            var builder = new System.Text.StringBuilder(raw.Length + visibleReference.Length);
            var inserted = false;
            foreach (var character in raw)
            {
                if (character is '\u0013' or '\u0014' or '\u0015')
                {
                    if (!inserted)
                    {
                        builder.Append(visibleReference);
                        inserted = true;
                    }
                    continue;
                }
                builder.Append(character);
            }
            if (!inserted) builder.Append(visibleReference);
            return builder.ToString().Trim();
        }
        finally
        {
            Release(nestedResult);
            Release(nested);
            Release(nestedFields);
            Release(code);
            Release(field);
            Release(fields);
        }
    }
}

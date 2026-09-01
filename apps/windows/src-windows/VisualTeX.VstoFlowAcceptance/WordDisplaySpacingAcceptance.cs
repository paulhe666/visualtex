using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class DisplaySpacingCase
    {
        internal string MarkerBookmark { get; set; } = string.Empty;
        internal string FormulaId { get; set; } = string.Empty;
        internal string ObjectMode { get; set; } = string.Empty;
        internal bool Numbered { get; set; }
    }

    private static void RunWordDisplaySpacing(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(artifactRoot, "word-display-spacing.docx");
        DeleteBulkPerformanceArtifact(outputPath);

        var previousCreateObjectMode = WordEquationNumbering.GetDefaultCreateObjectMode();
        var previousDisplayNumbered = WordEquationNumbering.GetDefaultDisplayEquationNumbered();
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            // This acceptance validates explicit VisualTeX OLE and OMML spacing.
            // It must not inherit or overwrite the user's persisted create mode.
            WordEquationNumbering.SetDefaultCreateObjectMode(
                FormulaOleContract.NativeOleMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(false);
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            var inlineOmmlFormulaId = InsertAndAssertInlineOmmlCaret(
                client,
                application,
                document,
                addIn);

            var cases = new List<DisplaySpacingCase>
            {
                InsertDisplaySpacingCase(
                    client,
                    application,
                    document,
                    addIn,
                    markerIndex: 1,
                    FormulaOleContract.NativeOleMode,
                    numbered: false),
                InsertDisplaySpacingCase(
                    client,
                    application,
                    document,
                    addIn,
                    markerIndex: 2,
                    FormulaOleContract.NativeOleMode,
                    numbered: true),
                InsertDisplaySpacingCase(
                    client,
                    application,
                    document,
                    addIn,
                    markerIndex: 3,
                    FormulaOleContract.WordOmmlMode,
                    numbered: false),
                InsertDisplaySpacingCase(
                    client,
                    application,
                    document,
                    addIn,
                    markerIndex: 4,
                    FormulaOleContract.WordOmmlMode,
                    numbered: true),
            };

            AssertDisplaySpacingCases(document, cases);
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: true,
                AddToRecentFiles: false);
            AssertDisplaySpacingCases(reopened, cases);
            AssertInlineOmmlCaretResult(reopened, inlineOmmlFormulaId);

            Console.WriteLine(
                "Word formula spacing acceptance passed: inline OMML left no VTBL placeholder and "
                + "typed prose stayed outside OMath; unnumbered/numbered OLE and OMML started "
                + "immediately after the preceding paragraph with no inserted blank paragraph.");
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
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            WordEquationNumbering.SetDefaultCreateObjectMode(previousCreateObjectMode);
            WordEquationNumbering.SetDefaultDisplayEquationNumbered(previousDisplayNumbered);
            ForceComCleanup();
        }
    }

    private static string InsertAndAssertInlineOmmlCaret(
        VisualTeXSessionClient client,
        Word.Application application,
        Word.Document document,
        VisualTeX.WordVsto.ThisAddIn addIn)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\">"
            + "<mi>x</mi><mo>=</mo><mn>1</mn></math>";
        Word.Selection? selection = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeText("前文");
            var existing = SnapshotSessionIds();
            addIn.OnInsertInlineOmml(new object());
            var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            AssertEqual(FormulaOleContract.WordOmmlMode, session.ObjectMode,
                "The inline OMML command created the wrong Word object mode.");
            Commit(
                client,
                session,
                "inline",
                FormulaOleContract.WordOmmlMode,
                "x=1",
                mathMl: mathMl);
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            AssertEqual("completed", final.Status,
                final.Error ?? "The inline OMML formula did not complete.");
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));

            // Do not adjust the caret in the acceptance test. This must exercise
            // the insertion point left by VisualTeX itself. Before typing, the
            // equation must already touch the paragraph mark with no hidden guard.
            var formulaId = final.FormulaId
                ?? throw new InvalidDataException("The inline OMML formula has no formulaId.");
            AssertInlineOmmlHasNoBoundaryCharacter(document, formulaId);
            selection.TypeText("正文");
            AssertInlineOmmlCaretResult(document, formulaId);
            selection.TypeParagraph();
            return formulaId;
        }
        finally { Release(selection); }
    }

    private static void AssertInlineOmmlHasNoBoundaryCharacter(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Range? equationRange = null;
        Word.Range? paragraphRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? nextCharacter = null;
        Word.Font? nextFont = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("The inline OMML formula bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            paragraphs = equationRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            AssertEqual(paragraphRange.End - 1, equationRange.End,
                "Inline OMML still has a character between OMath.End and the paragraph mark.");
            nextCharacter = document.Range(equationRange.End, equationRange.End + 1);
            AssertEqual("\r", nextCharacter.Text ?? string.Empty,
                "Inline OMML is followed by a placeholder instead of the paragraph mark.");
            nextFont = nextCharacter.Font;
            AssertEqual(0, nextFont.Hidden,
                "The paragraph mark after inline OMML unexpectedly inherited hidden formatting.");
        }
        finally
        {
            Release(nextFont);
            Release(nextCharacter);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(equationRange);
            Release(formulaBookmark);
        }
    }

    private static void AssertInlineOmmlCaretResult(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Range? equationRange = null;
        Word.Range? following = null;
        Word.Bookmarks? bookmarks = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("The inline OMML formula bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            AssertTrue((equationRange.Text ?? string.Empty).IndexOf("正文", StringComparison.Ordinal) < 0,
                "Text typed after inline OMML was absorbed into the native equation.");

            var contentEnd = document.Content.End;
            object followingStart = equationRange.End;
            object followingEnd = Math.Min(contentEnd, equationRange.End + 2);
            following = document.Range(ref followingStart, ref followingEnd);
            AssertTrue((following.Text ?? string.Empty).StartsWith("正文", StringComparison.Ordinal),
                "The inline OMML caret did not remain immediately before following prose.");
            Release(following);
            object afterProseStart = equationRange.End + 2;
            object afterProseEnd = Math.Min(contentEnd, equationRange.End + 3);
            following = document.Range(ref afterProseStart, ref afterProseEnd);
            AssertEqual("\r", following.Text ?? string.Empty,
                "A hidden placeholder remained after prose typed following inline OMML.");

            bookmarks = document.Bookmarks;
            var boundaryName = "VTBL_" + Guid.Parse(formulaId).ToString("N");
            AssertTrue(!bookmarks.Exists(boundaryName),
                "The inline OMML formula still exposes a VTBL placeholder bookmark.");
        }
        finally
        {
            Release(bookmarks);
            Release(following);
            Release(equationRange);
            Release(formulaBookmark);
        }
    }

    private static DisplaySpacingCase InsertDisplaySpacingCase(
        VisualTeXSessionClient client,
        Word.Application application,
        Word.Document document,
        VisualTeX.WordVsto.ThisAddIn addIn,
        int markerIndex,
        string objectMode,
        bool numbered)
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>x</mi><mo>=</mo><mn>1</mn></math>";
        var markerBookmark = InsertDisplaySpacingMarker(
            application,
            document,
            markerIndex,
            $"{(numbered ? "numbered " : string.Empty)}{objectMode}");
        Word.Font? contaminatedTypingFont = null;
        try
        {
            contaminatedTypingFont = application.Selection.Font;
            contaminatedTypingFont.Bold = 1;
            contaminatedTypingFont.Italic = 1;
            contaminatedTypingFont.StrikeThrough = 1;
            contaminatedTypingFont.Underline = Word.WdUnderline.wdUnderlineSingle;
        }
        finally { Release(contaminatedTypingFont); }
        var existing = SnapshotSessionIds();
        if (string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
            addIn.OnInsertDisplayOmml(new object());
        else
            addIn.OnInsertDisplay(new object());

        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        AssertEqual(objectMode, session.ObjectMode,
            "The display spacing command created the wrong Word object mode.");
        Commit(
            client,
            session,
            "block",
            objectMode,
            "x=1",
            numbered: numbered,
            mathMl: string.Equals(
                    objectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal)
                ? mathMl
                : null);
        var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
        AssertEqual("completed", final.Status,
            final.Error ?? "The display spacing formula did not complete.");
        client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
        WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
        AssertDisplayFormulaFollowingTypingIsUpright(
            application,
            document,
            $"{objectMode}-{(numbered ? "numbered" : "unnumbered")}-{markerIndex}");

        var result = new DisplaySpacingCase
        {
            MarkerBookmark = markerBookmark,
            FormulaId = final.FormulaId
                ?? throw new InvalidDataException("The display spacing formula has no formulaId."),
            ObjectMode = objectMode,
            Numbered = numbered,
        };
        AssertDisplaySpacingCase(document, result);
        return result;
    }

    private static void AssertDisplayFormulaFollowingTypingIsUpright(
        Word.Application application,
        Word.Document document,
        string label)
    {
        const string probe = "VT_DISPLAY_UPRIGHT_PROBE";
        Word.Selection? selection = null;
        Word.Range? range = null;
        Word.Font? font = null;
        try
        {
            selection = application.Selection;
            selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var start = selection.Start;
            selection.TypeText(probe);
            range = document.Range(start, start + probe.Length);
            font = range.Font;
            AssertEqual(0, font.Bold,
                $"{label}: text typed after the display formula inherited bold formatting.");
            AssertEqual(0, font.Italic,
                $"{label}: text typed after the display formula inherited italic formatting.");
            AssertEqual(0, font.StrikeThrough,
                $"{label}: text typed after the display formula inherited strike formatting.");
            AssertEqual(Word.WdUnderline.wdUnderlineNone, font.Underline,
                $"{label}: text typed after the display formula inherited underline formatting.");
            range.Delete();
            selection.SetRange(start, start);
        }
        finally
        {
            Release(font);
            Release(range);
            Release(selection);
        }
    }

    private static string InsertDisplaySpacingMarker(
        Word.Application application,
        Word.Document document,
        int markerIndex,
        string label)
    {
        Word.Selection? selection = null;
        Word.Range? markerRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            var markerText = $"Before {label} {markerIndex}";
            var markerStart = selection.Start;
            selection.TypeText(markerText);
            var markerEnd = selection.End;
            selection.TypeParagraph();
            markerRange = document.Range(markerStart, markerEnd);
            bookmarks = document.Bookmarks;
            var bookmarkName = $"VTTestDisplay{markerIndex}";
            bookmark = bookmarks.Add(bookmarkName, markerRange);
            return bookmarkName;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(markerRange);
            Release(selection);
        }
    }

    private static void AssertDisplaySpacingCases(
        Word.Document document,
        IEnumerable<DisplaySpacingCase> cases)
    {
        foreach (var item in cases)
            AssertDisplaySpacingCase(document, item);
    }

    private static void AssertDisplaySpacingCase(
        Word.Document document,
        DisplaySpacingCase item)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? markerBookmark = null;
        Word.Range? markerRange = null;
        Word.Range? markerParagraphRange = null;
        Word.Range? formulaRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(item.MarkerBookmark),
                $"The marker bookmark {item.MarkerBookmark} is missing.");
            markerBookmark = bookmarks[item.MarkerBookmark];
            markerRange = markerBookmark.Range;
            markerParagraphRange = markerRange.Paragraphs[1].Range;
            formulaRange = string.Equals(
                    item.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal)
                ? FindDisplaySpacingOmmlRange(document, item.FormulaId)
                : FindDisplaySpacingOleRange(document, item.FormulaId);
            var containerStart = DisplayFormulaContainerStart(formulaRange);
            AssertEqual(
                markerParagraphRange.End,
                containerStart,
                $"A blank paragraph was inserted before {(item.Numbered ? "numbered " : string.Empty)}{item.ObjectMode} formula {item.FormulaId}.");
        }
        finally
        {
            Release(formulaRange);
            Release(markerParagraphRange);
            Release(markerRange);
            Release(markerBookmark);
            Release(bookmarks);
        }
    }

    private static Word.Range FindDisplaySpacingOleRange(
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
                Word.Range? range = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formulaId,
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
            throw new InvalidDataException($"OLE formula {formulaId} was not found.");
        }
        finally { Release(shapes); }
    }

    private static Word.Range FindDisplaySpacingOmmlRange(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException($"OMML formula {formulaId} was not found.");
            return WordOmmlFormulaStore.GetEquationRange(bookmark);
        }
        finally { Release(bookmark); }
    }

    private static int DisplayFormulaContainerStart(Word.Range formulaRange)
    {
        Word.Table? table = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        try
        {
            if ((bool)formulaRange.get_Information(Word.WdInformation.wdWithInTable)
                && formulaRange.Tables.Count > 0)
            {
                table = formulaRange.Tables[1];
                return table.Range.Start;
            }
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.Start;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(table);
        }
    }
}

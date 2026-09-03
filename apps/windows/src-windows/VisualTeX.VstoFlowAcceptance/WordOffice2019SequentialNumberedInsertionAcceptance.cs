using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string WordRangeReferencePrefix = "visualtex-word-vsto-range:";

    private static void RunWordOffice2019SequentialNumberedInsertion(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            RunLegacyCompactTypingTailHeadingIsolationScenario(application);
            RunSequentialNumberedInsertionScenario(
                application,
                artifactRoot,
                "drifted-live-selection",
                captureInsideNumberedTable: false);
            RunSequentialNumberedInsertionScenario(
                application,
                artifactRoot,
                "captured-inside-numbered-table",
                captureInsideNumberedTable: true);
            RunMiddleInsertionFromEarlierNumberedFormulaScenario(
                application,
                artifactRoot);
            Console.WriteLine(
                "Office 2019 sequential numbered insertion acceptance passed: "
                + "a captured create anchor survived live-selection drift, a caret "
                + "captured inside an existing numbered table was redirected after "
                + "that formula's native SEQ paragraph, and a middle insertion no "
                + "longer jumps to the document tail.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunLegacyCompactTypingTailHeadingIsolationScenario(
        Word.Application application)
    {
        Word.Document? document = null;
        Word.Range? tail = null;
        Word.Range? following = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.ListTemplate? listTemplate = null;
        Word.ListLevel? listLevel = null;
        Word.ListFormat? listFormat = null;
        Word.Font? font = null;
        Word.ParagraphFormat? format = null;
        var stage = "create-document";
        try
        {
            document = application.Documents.Add();
            document.Activate();
            stage = "create-two-paragraph-fixture";
            document.Content.Text = "\r";
            tail = document.Paragraphs[1].Range.Duplicate;
            following = document.Paragraphs[2].Range.Duplicate;

            // Reproduce the historical generated tail exactly: a one-point empty
            // paragraph can retain Heading 1 plus a list value from the insertion
            // paragraph. Value 10 is the user-reported failure that turned the next
            // equation number from 0.0.x into 10.0.x.
            format = tail.ParagraphFormat;
            format.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
            format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = 1f;
            font = tail.Font;
            font.Size = 1f;
            stage = "create-list-template";
            listTemplate = document.ListTemplates.Add(
                OutlineNumbered: false,
                Name: "VisualTeXAcceptanceLegacyTail");
            listLevel = listTemplate.ListLevels[1];
            listLevel.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevel.NumberFormat = "%1.";
            listLevel.StartAt = 10;
            stage = "apply-list-template";
            listFormat = tail.ListFormat;
            listFormat.ApplyListTemplateWithLevel(
                listTemplate,
                ContinuePreviousList: false,
                ApplyTo: Word.WdListApplyTo.wdListApplyToWholeList,
                DefaultListBehavior: Word.WdDefaultListBehavior.wdWord10ListBehavior,
                ApplyLevel: 1);
            stage = "bookmark-legacy-tail";
            bookmarks = document.Bookmarks;
            bookmark = bookmarks.Add("VTNumberedTypingTail", tail);

            stage = "resolve-heading-scope";
            var scope = WordEquationNumbering.ResolveHeadingScopeAtPosition(
                document,
                following.Start,
                EquationNumberFormat.Heading2DotId);
            AssertEqual(
                "0.0",
                scope.NumberText,
                "A VisualTeX-owned 1 pt typing tail was treated as chapter 10.");

            stage = "expand-legacy-tail";
            application.Selection.SetRange(tail.Start, tail.Start);
            AssertTrue(
                WordEquationNumbering.ExpandCompactTrailingTypingParagraph(
                    application.Selection),
                "The legacy 1 pt typing tail was not upgraded to an ordinary paragraph.");

            stage = "assert-upgraded-tail";
            Release(format);
            format = tail.ParagraphFormat;
            Release(font);
            font = tail.Font;
            Release(listFormat);
            listFormat = tail.ListFormat;
            AssertTrue(font.Size > 1.1f,
                "The upgraded typing tail still exposes a 1 pt caret.");
            AssertEqual(
                Word.WdLineSpacing.wdLineSpaceSingle,
                format.LineSpacingRule,
                "The upgraded typing tail still uses exact 1 pt line spacing.");
            AssertEqual(
                Word.WdOutlineLevel.wdOutlineLevelBodyText,
                format.OutlineLevel,
                "The upgraded typing tail still participates in heading numbering.");
            AssertEqual(
                Word.WdListType.wdListNoNumbering,
                listFormat.ListType,
                "The upgraded typing tail retained the source list number.");
        }
        catch (Exception error)
        {
            throw new InvalidOperationException(
                $"Legacy compact typing-tail acceptance failed at {stage}: "
                + $"{error.GetType().FullName} (0x{error.HResult:X8}) {error.Message}");
        }
        finally
        {
            Release(format);
            Release(font);
            Release(listFormat);
            Release(listLevel);
            Release(listTemplate);
            Release(bookmark);
            Release(bookmarks);
            Release(following);
            Release(tail);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            ForceComCleanup();
        }
    }

    private static void RunSequentialNumberedInsertionScenario(
        Word.Application application,
        string artifactRoot,
        string scenarioName,
        bool captureInsideNumberedTable)
    {
        Word.Document? document = null;
        Word.Table? firstTable = null;
        Word.Range? firstTableRange = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var firstSession = CreateNumberedOmmlSession(
                document,
                WordRangeReference(0, 0),
                "x=1");
            service.InsertOmml(
                firstSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>=</mo><mn>1</mn></math>");

            firstTable = document.Tables[1];
            firstTableRange = firstTable.Range;
            string capturedAnchor;
            if (captureInsideNumberedTable)
            {
                capturedAnchor = WordRangeReference(
                    firstTableRange.Start,
                    firstTableRange.Start);
            }
            else
            {
                capturedAnchor = service.ReadSelection().ObjectId
                    ?? throw new InvalidOperationException(
                        "Word did not expose the post-formula create anchor.");
            }

            // Simulate Office 2019 moving the live Selection back into the old
            // numbered table while the external VisualTeX editor owns focus.
            application.Selection.SetRange(
                firstTableRange.Start,
                firstTableRange.Start);

            var secondSession = CreateNumberedOmmlSession(
                document,
                capturedAnchor,
                "y=2");
            service.InsertOmml(
                secondSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>y</mi><mo>=</mo><mn>2</mn></math>");

            AssertSequentialNumberedInsertion(
                application,
                document,
                firstSession.FormulaId,
                secondSession.FormulaId,
                scenarioName);
            AssertNumberedSpacingCleanup(document, scenarioName);
            if (!captureInsideNumberedTable)
            {
                AssertTerminalTypingParagraphIsOrdinary(
                    document,
                    scenarioName);
            }
            else
            {
                Word.Row? extraRow = null;
                var legacyRowPath = Path.Combine(
                    artifactRoot,
                    $"word-office2019-{scenarioName}-legacy-row.docx");
                try
                {
                    extraRow = firstTable.Rows.Add();
                    AssertEqual(
                        2,
                        firstTable.Rows.Count,
                        $"{scenarioName}: acceptance could not create the legacy empty numbered-table row.");

                    // A freshly added empty Word row can exist only in the COM model
                    // and be omitted from document.Content.WordOpenXML. The real bug
                    // reported by users is a persisted/reopened legacy document, where
                    // the extra row is serialized. Save and reopen so this acceptance
                    // exercises the same repair path without forcing every healthy
                    // 100-formula update to scan Rows.Count through COM.
                    document.SaveAs2(
                        legacyRowPath,
                        Word.WdSaveFormat.wdFormatXMLDocument);
                    Release(extraRow);
                    extraRow = null;
                    Release(firstTableRange);
                    firstTableRange = null;
                    Release(firstTable);
                    firstTable = null;
                    document.Close(Word.WdSaveOptions.wdSaveChanges);
                    Release(document);
                    document = application.Documents.Open(
                        legacyRowPath,
                        ReadOnly: false,
                        AddToRecentFiles: false,
                        Visible: false);
                    document.Activate();
                    firstTable = document.Tables[1];
                    firstTableRange = firstTable.Range;
                    AssertEqual(
                        2,
                        firstTable.Rows.Count,
                        $"{scenarioName}: persisted legacy row did not survive reopen.");

                    WordEquationNumbering.UpdateEquationNumbers(document);
                    AssertEqual(
                        1,
                        firstTable.Rows.Count,
                        $"{scenarioName}: update numbering did not remove the persisted legacy empty numbered-table row.");
                    AssertNumberedSpacingCleanup(document, scenarioName + "-repair");
                }
                finally { Release(extraRow); }
            }
            var artifactPath = Path.Combine(
                artifactRoot,
                $"word-office2019-{scenarioName}.docx");
            document.SaveAs2(
                artifactPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
        }
        finally
        {
            Release(firstTableRange);
            Release(firstTable);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            ForceComCleanup();
        }
    }

    private static void RunMiddleInsertionFromEarlierNumberedFormulaScenario(
        Word.Application application,
        string artifactRoot)
    {
        Word.Document? document = null;
        Word.Table? firstTable = null;
        Word.Range? firstTableRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? firstBookmark = null;
        Word.Bookmark? middleBookmark = null;
        Word.Bookmark? secondBookmark = null;
        Word.Range? firstBookmarkRange = null;
        Word.Range? middleBookmarkRange = null;
        Word.Range? secondBookmarkRange = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);

            var firstSession = CreateNumberedOmmlSession(
                document,
                WordRangeReference(0, 0),
                "x=1");
            service.InsertOmml(
                firstSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>=</mo><mn>1</mn></math>");

            var afterFirstAnchor = service.ReadSelection().ObjectId
                ?? throw new InvalidOperationException(
                    "Middle-insert acceptance could not capture the anchor after the first formula.");
            var secondSession = CreateNumberedOmmlSession(
                document,
                afterFirstAnchor,
                "y=2");
            service.InsertOmml(
                secondSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>y</mi><mo>=</mo><mn>2</mn></math>");

            firstTable = document.Tables[1];
            firstTableRange = firstTable.Range.Duplicate;
            var earlierFormulaAnchor = WordRangeReference(
                firstTableRange.Start,
                firstTableRange.Start);

            // Reproduce the problematic state seen after creating an empty line
            // between two numbered formulas: Word may report the create anchor
            // from the earlier formula's table while the live Selection later
            // drifts elsewhere as the external editor takes focus. The insertion
            // must stay between formula #1 and #2, never at the document tail.
            application.Selection.EndKey(Word.WdUnits.wdStory);
            var middleSession = CreateNumberedOmmlSession(
                document,
                earlierFormulaAnchor,
                "z=3");
            service.InsertOmml(
                middleSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>z</mi><mo>=</mo><mn>3</mn></math>");

            var artifactPath = Path.Combine(
                artifactRoot,
                "word-office2019-middle-insertion.docx");
            Console.WriteLine(
                $"    [diag] middle insertion: tables={document.Tables.Count}; "
                + $"omaths={document.OMaths.Count}; paragraphs={document.Paragraphs.Count}");
            for (var tableIndex = 1; tableIndex <= document.Tables.Count; tableIndex++)
            {
                Word.Table? diagnosticTable = null;
                Word.Range? diagnosticRange = null;
                try
                {
                    diagnosticTable = document.Tables[tableIndex];
                    diagnosticRange = diagnosticTable.Range;
                    Console.WriteLine(
                        $"    [diag] table#{tableIndex}: "
                        + $"rows={diagnosticTable.Rows.Count}; cols={diagnosticTable.Columns.Count}; "
                        + $"range={diagnosticRange.Start}:{diagnosticRange.End}");
                }
                finally
                {
                    Release(diagnosticRange);
                    Release(diagnosticTable);
                }
            }

            AssertEqual(
                3,
                document.Tables.Count,
                "Middle-insert acceptance did not create three independent numbered tables.");
            AssertEqual(
                3,
                document.OMaths.Count,
                "Middle-insert acceptance lost a native OMML formula.");

            bookmarks = document.Bookmarks;
            firstBookmark = bookmarks[WordEquationNumbering.EquationBookmarkName(
                firstSession.FormulaId)];
            middleBookmark = bookmarks[WordEquationNumbering.EquationBookmarkName(
                middleSession.FormulaId)];
            secondBookmark = bookmarks[WordEquationNumbering.EquationBookmarkName(
                secondSession.FormulaId)];
            firstBookmarkRange = firstBookmark.Range;
            middleBookmarkRange = middleBookmark.Range;
            secondBookmarkRange = secondBookmark.Range;
            AssertTrue(
                firstBookmarkRange.Start < middleBookmarkRange.Start,
                "Middle-insert acceptance placed the new formula before the first formula.");
            AssertTrue(
                middleBookmarkRange.Start < secondBookmarkRange.Start,
                "Middle-insert acceptance still moved the new formula to the document tail.");

            // Inserting into the ordinary paragraph after formula #1 must leave a
            // genuine typing paragraph after the new middle formula, followed by
            // the mandatory compact paragraph before the old formula #2. This is
            // the same user-facing contract as Enter at the end of a right-cell
            // number: the ordinary line is deletable, while deleting it must not
            // merge the two independent 1x3 tables.
            AssertMiddleInsertionTypingParagraphIsDeletable(
                document,
                middleSession.FormulaId,
                secondSession.FormulaId);
            AssertNumberedSpacingCleanup(
                document,
                "middle-insertion-after-typing-line-delete");
            document.SaveAs2(
                artifactPath,
                Word.WdSaveFormat.wdFormatXMLDocument);

        }
        finally
        {
            Release(secondBookmarkRange);
            Release(middleBookmarkRange);
            Release(firstBookmarkRange);
            Release(secondBookmark);
            Release(middleBookmark);
            Release(firstBookmark);
            Release(bookmarks);
            Release(firstTableRange);
            Release(firstTable);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateNumberedOmmlSession(
        Word.Document document,
        string sourceObjectId,
        string latex)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            SourceDocumentId = document.FullName,
            SourceObjectId = sourceObjectId,
            Title = "Office 2019 sequential numbered insertion acceptance",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = true,
            FontSizePt = 11,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
        };
    }

    private static string WordRangeReference(int start, int end) =>
        $"{WordRangeReferencePrefix}{start}:{end}";

    private static void AssertMiddleInsertionTypingParagraphIsDeletable(
        Word.Document document,
        string middleFormulaId,
        string followingFormulaId)
    {
        Word.Table? middleTable = null;
        Word.Table? followingTable = null;
        Word.Range? middleTableRange = null;
        Word.Range? followingTableRange = null;
        Word.Range? typingRange = null;
        try
        {
            middleTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    middleFormulaId)
                ?? throw new InvalidDataException(
                    "Middle-insert acceptance lost the newly inserted 1x3 table.");
            followingTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    followingFormulaId)
                ?? throw new InvalidDataException(
                    "Middle-insert acceptance lost the following 1x3 table.");
            middleTableRange = middleTable.Range;
            followingTableRange = followingTable.Range;
            AssertEqual(
                middleTableRange.End + 2,
                followingTableRange.Start,
                "Middle insertion did not leave exactly one ordinary typing paragraph plus one compact table separator.");
            AssertNormalTypingBodyParagraph(
                document,
                middleTableRange.End,
                middleTableRange.End + 1,
                "middle-insertion: ordinary typing paragraph");
            AssertCompactBodyParagraph(
                document,
                middleTableRange.End + 1,
                followingTableRange.Start,
                "middle-insertion: mandatory table separator");

            var tableCount = document.Tables.Count;
            var mathCount = document.OMaths.Count;
            typingRange = document.Range(
                middleTableRange.End,
                middleTableRange.End + 1);
            typingRange.Delete();
            AssertEqual(tableCount, document.Tables.Count,
                "Deleting the middle insertion typing paragraph merged two 1x3 tables.");
            AssertEqual(mathCount, document.OMaths.Count,
                "Deleting the middle insertion typing paragraph removed a formula.");

            Release(middleTableRange);
            middleTableRange = middleTable.Range;
            Release(followingTableRange);
            followingTableRange = followingTable.Range;
            AssertCompactBodyParagraph(
                document,
                middleTableRange.End,
                followingTableRange.Start,
                "middle-insertion: retained separator after deleting typing line");
        }
        finally
        {
            Release(typingRange);
            Release(followingTableRange);
            Release(middleTableRange);
            Release(followingTable);
            Release(middleTable);
        }
    }

    private static void AssertNumberedSpacingCleanup(
        Word.Document document,
        string scenarioName)
    {
        var tableRanges = new List<(int Start, int End)>();
        for (var index = 1; index <= document.Tables.Count; index++)
        {
            Word.Table? table = null;
            Word.Range? tableRange = null;
            try
            {
                table = document.Tables[index];
                AssertEqual(
                    1,
                    table.Rows.Count,
                    $"{scenarioName}: numbered table {index} retained an empty structural row.");
                AssertEqual(
                    3,
                    table.Columns.Count,
                    $"{scenarioName}: numbered table {index} is not the current 1x3 host.");
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
        AssertTrue(tableRanges.Count > 0,
            $"{scenarioName}: no numbered 1x3 table was retained.");
        AssertEqual(document.Content.Start, tableRanges[0].Start,
            $"{scenarioName}: an unexpected body paragraph remains before the first numbered table.");

        for (var index = 1; index < tableRanges.Count; index++)
        {
            AssertCompactBodyParagraph(
                document,
                tableRanges[index - 1].End,
                tableRanges[index].Start,
                $"{scenarioName}: inter-table separator {index}");
        }
        AssertNormalTypingBodyParagraph(
            document,
            tableRanges[tableRanges.Count - 1].End,
            document.Content.End,
            $"{scenarioName}: terminal typing paragraph");

        var ordinaryParagraphCount = 0;
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Frames? frames = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                if ((bool)range.get_Information(Word.WdInformation.wdWithInTable))
                    continue;
                frames = range.Frames;
                if (frames.Count == 0) ordinaryParagraphCount++;
            }
            finally
            {
                Release(frames);
                Release(range);
                Release(paragraph);
            }
        }
        AssertEqual(
            tableRanges.Count,
            ordinaryParagraphCount,
            $"{scenarioName}: a visible or duplicate body paragraph remains around the numbered tables.");
    }

    private static void AssertCompactBodyParagraph(
        Word.Document document,
        int start,
        int end,
        string context)
    {
        Word.Range? range = null;
        Word.Font? font = null;
        Word.ParagraphFormat? format = null;
        try
        {
            range = document.Range(start, end);
            AssertEqual("\r", range.Text,
                context + ": expected exactly one structural paragraph mark.");
            AssertTrue(!(bool)range.get_Information(
                           Word.WdInformation.wdWithInTable)
                       && range.Tables.Count == 0
                       && range.OMaths.Count == 0
                       && range.Fields.Count == 0
                       && range.Bookmarks.Count == 0
                       && range.InlineShapes.Count == 0
                       && range.Frames.Count == 0,
                context + ": structural paragraph owns table, formula, field, bookmark, OLE, or Frame content.");
            font = range.Font;
            format = range.ParagraphFormat;
            AssertTrue(font.Size <= 1.1f,
                context + $": paragraph remains visible at {font.Size:0.###}pt.");
            AssertEqual(
                Word.WdLineSpacing.wdLineSpaceExactly,
                format.LineSpacingRule,
                context + ": paragraph is not exact-height.");
            AssertTrue(format.LineSpacing <= 1.1f,
                context + $": paragraph remains {format.LineSpacing:0.###}pt high.");
        }
        finally
        {
            Release(format);
            Release(font);
            Release(range);
        }
    }

    private static void AssertNormalTypingBodyParagraph(
        Word.Document document,
        int start,
        int end,
        string context)
    {
        Word.Range? range = null;
        Word.Font? font = null;
        Word.ParagraphFormat? format = null;
        try
        {
            range = document.Range(start, end);
            AssertEqual("\r", range.Text,
                context + ": expected exactly one empty paragraph mark.");
            AssertTrue(!(bool)range.get_Information(
                           Word.WdInformation.wdWithInTable)
                       && range.Tables.Count == 0
                       && range.OMaths.Count == 0
                       && range.Fields.Count == 0
                       && range.Bookmarks.Count == 0
                       && range.InlineShapes.Count == 0
                       && range.Frames.Count == 0,
                context + ": typing paragraph owns table, formula, field, bookmark, OLE, or Frame content.");
            font = range.Font;
            format = range.ParagraphFormat;
            AssertTrue(font.Size > 1.1f,
                context + ": typing paragraph inherited the internal 1pt separator font.");
            AssertEqual(
                Word.WdLineSpacing.wdLineSpaceSingle,
                format.LineSpacingRule,
                context + ": typing paragraph is not a normal single-line paragraph.");
        }
        finally
        {
            Release(format);
            Release(font);
            Release(range);
        }
    }

    private static void AssertTerminalTypingParagraphIsOrdinary(
        Word.Document document,
        string scenarioName)
    {
        Word.Paragraph? lastParagraph = null;
        Word.Range? lastRange = null;
        try
        {
            lastParagraph = document.Paragraphs[document.Paragraphs.Count];
            lastRange = lastParagraph.Range.Duplicate;
            AssertEqual(document.Content.End, lastRange.End,
                $"{scenarioName}: the final typing paragraph does not end at the document boundary.");
            AssertNormalTypingBodyParagraph(
                document,
                lastRange.Start,
                lastRange.End,
                $"{scenarioName}: final typing paragraph");
        }
        finally
        {
            Release(lastRange);
            Release(lastParagraph);
        }
    }

    private static void AssertSequentialNumberedInsertion(
        Word.Application application,
        Word.Document document,
        string firstFormulaId,
        string secondFormulaId,
        string scenarioName)
    {
        AssertEqual(
            2,
            document.Tables.Count,
            $"{scenarioName}: the second numbered formula did not create an independent table.");
        AssertEqual(
            2,
            document.OMaths.Count,
            $"{scenarioName}: one native formula was swallowed during sequential insertion.");
        AssertEqual(0, document.Shapes.Count,
            $"{scenarioName}: sequential insertion recreated a retired Shape/TextBox number.");
        AssertEqual(0, document.Frames.Count,
            $"{scenarioName}: sequential insertion recreated a retired caption Frame.");

        AssertOmmlTableNumberLifecyclePhase(
            application,
            document,
            firstFormulaId,
            scenarioName + " first 1x3 host");
        AssertOmmlTableNumberLifecyclePhase(
            application,
            document,
            secondFormulaId,
            scenarioName + " second 1x3 host");
        AssertManagedNativeOmmlInterTableSeparatorsCompact(
            document,
            new[] { firstFormulaId, secondFormulaId },
            scenarioName + " sequential insertion");

        Word.Table? firstTable = null;
        Word.Table? secondTable = null;
        Word.Range? firstTableRange = null;
        Word.Range? secondTableRange = null;
        Word.Range? firstNumberRange = null;
        Word.Range? secondNumberRange = null;
        try
        {
            firstTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    firstFormulaId)
                ?? throw new InvalidDataException(
                    $"{scenarioName}: the first FormulaId lost its 1x3 host.");
            secondTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    secondFormulaId)
                ?? throw new InvalidDataException(
                    $"{scenarioName}: the second FormulaId lost its 1x3 host.");
            firstTableRange = firstTable.Range;
            secondTableRange = secondTable.Range;
            AssertTrue(firstTableRange.Start < secondTableRange.Start,
                $"{scenarioName}: the second formula was inserted before the captured first formula.");
            AssertTrue(firstTableRange.End < secondTableRange.Start,
                $"{scenarioName}: the two 1x3 hosts overlap or were merged.");

            firstNumberRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    firstFormulaId)
                ?? throw new InvalidDataException(
                    $"{scenarioName}: the first right-cell number is missing.");
            secondNumberRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    secondFormulaId)
                ?? throw new InvalidDataException(
                    $"{scenarioName}: the second right-cell number is missing.");
            AssertEqual(
                "1",
                (firstNumberRange.Text ?? string.Empty).Trim().Trim('(', ')'),
                $"{scenarioName}: the first formula was renumbered out of document order.");
            AssertEqual(
                "2",
                (secondNumberRange.Text ?? string.Empty).Trim().Trim('(', ')'),
                $"{scenarioName}: the second formula did not receive sequence number 2.");
        }
        finally
        {
            Release(secondNumberRange);
            Release(firstNumberRange);
            Release(secondTableRange);
            Release(firstTableRange);
            Release(secondTable);
            Release(firstTable);
        }
    }
}

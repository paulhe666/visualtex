using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlTableNumberStressAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The OMML 1x3 stress acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        RunOmmlTableThreeFormulaMutationStress(artifactRoot);
        RunOmmlTableTwentyFormulaPaginationStress(artifactRoot);
        Console.WriteLine(
            "Word OMML 1x3 stress acceptance passed: edit, format changes, number toggle/re-enable and 20-formula pagination/save-reopen retained the direct-SEQ 1x3 host without hidden caption paragraphs.");
    }

    private static void RunOmmlTableThreeFormulaMutationStress(string artifactRoot)
    {
        var path = Path.Combine(artifactRoot, "word-omml-1x3-three-formula-stress.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.Content.Text = "1 Introduction\r";
            Word.Paragraph? heading = null;
            try
            {
                heading = document.Paragraphs[1];
                heading.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
            }
            finally { Release(heading); }
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            for (var index = 0; index < 3; index++)
            {
                var formulaId = Guid.NewGuid().ToString("D");
                formulaIds.Add(formulaId);
                InsertNumberedOmmlAtDocumentEnd(
                    application,
                    document,
                    service,
                    formulaId,
                    @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    QuadraticFormulaMathMl());
            }
            AssertEqual(3, document.Tables.Count,
                "Three numbered OMML formulas did not create exactly three 1x3 tables.");
            AssertVisibleNumbers(document, formulaIds, "(1)", "(2)", "(3)");
            for (var index = 0; index < formulaIds.Count; index++)
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    formulaIds[index],
                    $"three-formula initial {index + 1}");

            // Edit the middle formula in place. Its table and direct SEQ identity
            // must survive; only the center OMath body changes.
            var middleId = formulaIds[1];
            var middleMetadata = WordOmmlFormulaStore.TryRead(document, middleId)
                ?? throw new InvalidDataException("Middle OMML metadata is missing before edit.");
            Word.Range? middleRange = null;
            try
            {
                middleRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    middleId,
                    middleMetadata);
                var editSession = CreateNumberedOmmlTabSession(
                    middleId,
                    document.FullName,
                    middleRange.Start,
                    middleRange.End,
                    latex: @"y=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    originalMetadata: middleMetadata);
                service.ReplaceOmml(
                    editSession,
                    QuadraticFormulaMathMl().Replace("<mi>x</mi>", "<mi>y</mi>"));
            }
            finally { Release(middleRange); }
            AssertEqual(3, document.Tables.Count,
                "Editing the middle numbered OMML changed the table count.");
            AssertVisibleNumbers(document, formulaIds, "(1)", "(2)", "(3)");
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                middleId,
                "middle formula after edit");

            // Format switch: no table or bookmark reconstruction should be needed;
            // only prefix/SEQ results in the right cells change.
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            AssertTrue(WordEquationNumbering.UpdateEquationNumbers(document) >= 3,
                "Heading-number format update did not process the three formulas.");
            AssertVisibleNumbers(document, formulaIds, "(1.1)", "(1.2)", "(1.3)");
            AssertEqual(3, document.Tables.Count,
                "Heading-number format update changed the 1x3 table count.");
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            AssertTrue(WordEquationNumbering.UpdateEquationNumbers(document) >= 3,
                "Continuous-number format restore did not process the three formulas.");
            AssertVisibleNumbers(document, formulaIds, "(1)", "(2)", "(3)");

            // Remove numbering from the middle formula. The managed table must be
            // dismantled back to one standalone true Display OMath, not left as an
            // empty 1x3 shell. Then restore numbering using the same FormulaId.
            middleMetadata = WordOmmlFormulaStore.TryRead(document, middleId)
                ?? throw new InvalidDataException("Middle metadata is missing before number toggle.");
            Word.Range? toggleRange = null;
            try
            {
                toggleRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    middleId,
                    middleMetadata);
                var unnumberSession = CreateNumberedOmmlTabSession(
                    middleId,
                    document.FullName,
                    toggleRange.Start,
                    toggleRange.End,
                    latex: @"y=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    originalMetadata: middleMetadata);
                unnumberSession.Numbered = false;
                service.ReplaceOmml(
                    unnumberSession,
                    QuadraticFormulaMathMl().Replace("<mi>x</mi>", "<mi>y</mi>"));
            }
            finally { Release(toggleRange); }
            AssertEqual(2, document.Tables.Count,
                "Numbered→unnumbered did not dismantle exactly the middle 1x3 table.");
            var unnumberedMetadata = WordOmmlFormulaStore.TryRead(document, middleId)
                ?? throw new InvalidDataException("Middle metadata is missing after unnumbering.");
            AssertTrue(!unnumberedMetadata.Numbered,
                "Numbered→unnumbered did not persist Numbered=false.");
            Word.Bookmark? middleBookmark = null;
            Word.Range? unnumberedRange = null;
            Word.Paragraphs? unnumberedParagraphs = null;
            Word.Paragraph? unnumberedParagraph = null;
            Word.ParagraphFormat? unnumberedParagraphFormat = null;
            try
            {
                middleBookmark = WordOmmlFormulaStore.FindByFormulaId(document, middleId)
                    ?? throw new InvalidDataException("Unnumbered middle formula lost VTOMML identity.");
                unnumberedRange = WordOmmlFormulaStore.GetEquationRange(middleBookmark);
                AssertTrue(!(bool)unnumberedRange.get_Information(Word.WdInformation.wdWithInTable),
                    "Unnumbered middle formula remained in a table.");
                AssertEqual(Word.WdOMathType.wdOMathDisplay, unnumberedRange.OMaths[1].Type,
                    "Unnumbered middle formula degraded from true Display OMath.");
                unnumberedParagraphs = unnumberedRange.Paragraphs;
                AssertEqual(1, unnumberedParagraphs.Count,
                    "Unnumbered middle formula no longer occupies one standalone paragraph.");
                unnumberedParagraph = unnumberedParagraphs[1];
                unnumberedParagraphFormat = unnumberedParagraph.Format;
                AssertTrue(
                    !(unnumberedParagraphFormat.LineSpacingRule == Word.WdLineSpacing.wdLineSpaceExactly
                      && unnumberedParagraphFormat.LineSpacing <= 2.01f),
                    $"Unnumbered middle formula inherited the compact 1pt table separator: rule={unnumberedParagraphFormat.LineSpacingRule}, line={unnumberedParagraphFormat.LineSpacing:0.##}pt.");
                AssertEqual(Word.WdLineSpacing.wdLineSpaceSingle,
                    unnumberedParagraphFormat.LineSpacingRule,
                    "Unnumbered middle formula did not restore ordinary single line spacing after dismantling its 1x3 host.");

                // Simulate a document already damaged by an older build. The normal
                // document-open refresh must self-heal this exact managed-OMML 1pt
                // signature without requiring the user to edit the formula again.
                unnumberedParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                unnumberedParagraphFormat.LineSpacing = 1f;
                AssertTrue(WordEquationNumbering.RefreshNumberedOmmlTabLayouts(document) >= 1,
                    "Document-open OMML refresh did not detect the legacy 1pt standalone formula.");
                Release(unnumberedParagraphFormat); unnumberedParagraphFormat = null;
                Release(unnumberedParagraph); unnumberedParagraph = null;
                Release(unnumberedParagraphs); unnumberedParagraphs = null;
                unnumberedParagraphs = unnumberedRange.Paragraphs;
                unnumberedParagraph = unnumberedParagraphs[1];
                unnumberedParagraphFormat = unnumberedParagraph.Format;
                AssertEqual(Word.WdLineSpacing.wdLineSpaceSingle,
                    unnumberedParagraphFormat.LineSpacingRule,
                    "Document-open OMML refresh did not repair the legacy 1pt standalone formula line box.");
                Console.WriteLine(
                    $"  unnumbered OMML line-box repair: rule={unnumberedParagraphFormat.LineSpacingRule}, line={unnumberedParagraphFormat.LineSpacing:0.##}pt.");

                var renumberSession = CreateNumberedOmmlTabSession(
                    middleId,
                    document.FullName,
                    unnumberedRange.Start,
                    unnumberedRange.End,
                    latex: @"y=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    originalMetadata: unnumberedMetadata);
                service.ReplaceOmml(
                    renumberSession,
                    QuadraticFormulaMathMl().Replace("<mi>x</mi>", "<mi>y</mi>"));
            }
            finally
            {
                Release(unnumberedParagraphFormat);
                Release(unnumberedParagraph);
                Release(unnumberedParagraphs);
                Release(unnumberedRange);
                Release(middleBookmark);
            }
            WordEquationNumbering.UpdateEquationNumbers(document);
            AssertEqual(3, document.Tables.Count,
                "Unnumbered→numbered did not restore exactly one 1x3 table.");
            AssertVisibleNumbers(document, formulaIds, "(1)", "(2)", "(3)");
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                middleId,
                "middle formula after renumber");

            document.Fields.Update();
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                path,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertEqual(3, document.Tables.Count,
                "Three-formula stress save/reopen changed the table count.");
            AssertVisibleNumbers(document, formulaIds, "(1)", "(2)", "(3)");
            foreach (var formulaId in formulaIds)
                AssertDirectTableArtifactsStayInsideRightCell(
                    document,
                    formulaId,
                    "three-formula save/reopen");
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

    private static void RunOmmlTableTwentyFormulaPaginationStress(string artifactRoot)
    {
        var path = Path.Combine(artifactRoot, "word-omml-1x3-twenty-formula-stress.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var ids = new List<string>();
            for (var index = 0; index < 20; index++)
            {
                var id = Guid.NewGuid().ToString("D");
                ids.Add(id);
                InsertNumberedOmmlAtDocumentEnd(
                    application,
                    document,
                    service,
                    id,
                    @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    QuadraticFormulaMathMl());
            }
            AssertEqual(20, document.Tables.Count,
                "Twenty numbered OMML formulas did not create exactly twenty tables.");
            WordEquationNumbering.UpdateEquationNumbers(document);
            for (var index = 0; index < ids.Count; index++)
            {
                AssertEqual(
                    $"({index + 1})",
                    ReadVisibleEquationNumber(document, ids[index]),
                    $"Twenty-formula stress number {index + 1} is incorrect.");
                AssertDirectTableArtifactsStayInsideRightCell(
                    document,
                    ids[index],
                    $"twenty-formula initial {index + 1}");
            }
            // Full visual/geometry checks on representative rows across the flow.
            foreach (var index in new[] { 0, 9, 19 })
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    ids[index],
                    $"twenty-formula representative {index + 1}");

            document.Repaginate();
            var pageCount = document.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
            AssertTrue(pageCount >= 2,
                "Twenty-formula stress did not actually cross a page boundary.");
            document.Fields.Update();
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                path,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertEqual(20, document.Tables.Count,
                "Twenty-formula save/reopen changed the table count.");
            for (var index = 0; index < ids.Count; index++)
            {
                AssertEqual(
                    $"({index + 1})",
                    ReadVisibleEquationNumber(document, ids[index]),
                    $"Reopened twenty-formula number {index + 1} is incorrect.");
                AssertDirectTableArtifactsStayInsideRightCell(
                    document,
                    ids[index],
                    $"twenty-formula save/reopen {index + 1}");
            }
            Console.WriteLine(
                $"  twenty-formula pagination: pages={pageCount}, tables={document.Tables.Count}, first={ReadVisibleEquationNumber(document, ids[0])}, last={ReadVisibleEquationNumber(document, ids[ids.Count - 1])}.");
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

    private static void InsertNumberedOmmlAtDocumentEnd(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string latex,
        string mathMl)
    {
        var insertion = document.Content.End - 1;
        application.Selection.SetRange(insertion, insertion);
        var session = CreateNumberedOmmlTabSession(
            formulaId,
            document.FullName,
            insertion,
            insertion,
            latex,
            originalMetadata: null);
        service.InsertOmml(session, mathMl);
    }

    private static void AssertVisibleNumbers(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        params string[] expected)
    {
        AssertEqual(expected.Length, formulaIds.Count,
            "Visible-number expectation count does not match formula count.");
        for (var index = 0; index < formulaIds.Count; index++)
            AssertEqual(
                expected[index],
                ReadVisibleEquationNumber(document, formulaIds[index]),
                $"Visible equation number {index + 1} is incorrect.");
    }

    private static void AssertDirectTableArtifactsStayInsideRightCell(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Range? formulaRange = null;
        Word.Table? table = null;
        Word.Cell? numberCell = null;
        Word.Range? cellRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? owned = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(context + ": metadata missing.");
            formulaRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            AssertTrue((bool)formulaRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": formula is not in its 1x3 host.");
            table = formulaRange.Tables[1];
            AssertEqual(1, table.Rows.Count, context + ": table gained extra rows.");
            AssertEqual(3, table.Columns.Count, context + ": table lost 1x3 geometry.");
            numberCell = table.Cell(1, 3);
            cellRange = numberCell.Range;
            AssertEqual(1, cellRange.Fields.Count,
                context + ": right cell does not contain exactly one direct SEQ.");
            AssertEqual(0, cellRange.OMaths.Count,
                context + ": right cell contains OMML.");
            bookmarks = document.Bookmarks;
            foreach (var name in new[]
                     {
                         WordEquationNumbering.EquationBookmarkName(formulaId),
                         WordEquationNumbering.NativeCaptionBookmarkName(formulaId),
                         WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                     })
            {
                AssertTrue(bookmarks.Exists(name), context + $": bookmark {name} missing.");
                Release(bookmark); bookmark = bookmarks[name];
                Release(owned); owned = bookmark.Range;
                AssertTrue((bool)owned.get_Information(Word.WdInformation.wdWithInTable),
                    context + $": bookmark {name} escaped to a body paragraph.");
                AssertTrue(owned.Start >= cellRange.Start && owned.End <= cellRange.End,
                    context + $": bookmark {name} escaped the right cell.");
            }
        }
        finally
        {
            Release(owned);
            Release(bookmark);
            Release(bookmarks);
            Release(cellRange);
            Release(numberCell);
            Release(table);
            Release(formulaRange);
        }
    }
}

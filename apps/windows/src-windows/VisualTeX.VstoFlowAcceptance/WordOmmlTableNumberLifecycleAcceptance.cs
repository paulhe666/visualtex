using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlTableNumberLifecycleAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The OMML 1x3 lifecycle acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "word-omml-1x3-number-lifecycle.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Field? externalReference = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Content.End - 1;
            application.Selection.SetRange(insertion, insertion);
            var service = new WordFormulaService(application);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion,
                insertion,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());

            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "01-after-insert");

            externalReference = InsertExternalEquationReference(document, formulaId);
            externalReference.Update();
            AssertEqual(
                ReadVisibleEquationNumber(document, formulaId).Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "The external body REF does not match the direct table SEQ target.");

            document.Fields.Update();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "02-after-f9");
            document.Save();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "03-after-save");

            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            Release(externalReference); externalReference = null;
            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "04-after-reopen");
            externalReference = document.Fields
                .Cast<Word.Field>()
                .FirstOrDefault(field =>
                {
                    Word.Range? code = null;
                    try
                    {
                        code = field.Code;
                        return (code.Text ?? string.Empty).IndexOf(
                            "REF VTEqNum_",
                            StringComparison.OrdinalIgnoreCase) >= 0;
                    }
                    finally { Release(code); }
                });
            AssertTrue(externalReference is not null,
                "The body equation reference disappeared after save/reopen.");
            externalReference!.Update();
            AssertEqual(
                ReadVisibleEquationNumber(document, formulaId).Trim('(', ')'),
                NormalizeEquationNumberText(externalReference.Result.Text),
                "The reopened body REF does not match the table SEQ target.");
            document.Save();

            Console.WriteLine(
                "Word OMML 1x3 lifecycle acceptance passed: true wdOMathDisplay remained isolated in center cell (1,2), the ordinary direct SEQ number remained in cell (1,3) behind a genuine right TabStop, no hidden VTEq caption paragraph existed outside the table, and F9/save/reopen plus an external REF all remained stable.");
        }
        finally
        {
            Release(externalReference);
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

    private static void ConfigureOmmlTableNumberPage(Word.Document document)
    {
        Word.Section? section = null;
        Word.PageSetup? setup = null;
        try
        {
            section = document.Sections[1];
            setup = section.PageSetup;
            setup.PageWidth = 595.3f;
            setup.PageHeight = 841.9f;
            setup.LeftMargin = 90f;
            setup.RightMargin = 90f;
            setup.TopMargin = 72f;
            setup.BottomMargin = 72f;
        }
        finally
        {
            Release(setup);
            Release(section);
        }
    }

    private static void AssertOmmlTableNumberLifecyclePhase(
        Word.Application application,
        Word.Document document,
        string formulaId,
        string phase)
    {
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Tables? tables = null;
        Word.Table? table = null;
        Word.Rows? rows = null;
        Word.Columns? columns = null;
        Word.Cell? centerCell = null;
        Word.Cell? numberCell = null;
        Word.Range? centerRange = null;
        Word.Range? numberCellRange = null;
        Word.Fields? numberFields = null;
        Word.Field? sequenceField = null;
        Word.Range? sequenceCode = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? visibleBookmark = null;
        Word.Bookmark? numberBookmark = null;
        Word.Bookmark? captionBookmark = null;
        Word.Range? visibleRange = null;
        Word.Range? numberRange = null;
        Word.Range? captionRange = null;
        Word.Paragraphs? numberParagraphs = null;
        Word.Paragraph? numberParagraph = null;
        Word.Range? numberParagraphRange = null;
        Word.ParagraphFormat? numberFormat = null;
        Word.Range? numberParagraphMark = null;
        Word.Font? numberParagraphMarkFont = null;
        Word.Font? visibleNumberFont = null;
        Word.Range? numberEnd = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? setup = null;
        Word.Window? window = null;
        Word.View? view = null;
        var restoreFieldCodes = false;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(phase + ": formula metadata is missing.");
            formulaRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            maths = formulaRange.OMaths;
            AssertEqual(1, maths.Count, phase + ": center formula is not exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                phase + ": center formula degraded from wdOMathDisplay.");
            AssertEqual(0, formulaRange.Fields.Count,
                phase + ": a number field leaked inside m:oMath.");
            AssertTrue((bool)formulaRange.get_Information(Word.WdInformation.wdWithInTable),
                phase + ": numbered OMML is no longer inside the managed table.");
            tables = formulaRange.Tables;
            AssertEqual(1, tables.Count, phase + ": formula does not own exactly one table.");
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            AssertEqual(1, rows.Count, phase + ": numbered table gained an extra row.");
            AssertEqual(3, columns.Count, phase + ": numbered table is not 1x3.");
            AssertEqual(0, document.Shapes.Count, phase + ": numbered OMML created a Shape/TextBox.");

            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerRange = centerCell.Range;
            numberCellRange = numberCell.Range;
            AssertEqual(1, centerRange.OMaths.Count,
                phase + ": center cell no longer contains exactly one OMath.");
            AssertEqual(0, centerRange.Fields.Count,
                phase + ": center cell contains a Word field.");
            AssertEqual(0, numberCellRange.OMaths.Count,
                phase + ": number cell unexpectedly contains OMML.");
            numberFields = numberCellRange.Fields;
            AssertEqual(1, numberFields.Count,
                phase + ": number cell must contain exactly one direct SEQ field.");
            sequenceField = numberFields[1];
            sequenceCode = sequenceField.Code;
            AssertTrue(WordEquationNumbering.IsVisualTeXSequenceFieldCode(sequenceCode.Text),
                phase + ": number-cell field is not SEQ VisualTeXEquation.");
            AssertTrue((sequenceCode.Text ?? string.Empty).IndexOf(
                    "REF VTEqNum_",
                    StringComparison.OrdinalIgnoreCase) < 0,
                phase + ": the number cell still contains the old generated REF indirection.");

            bookmarks = document.Bookmarks;
            visibleBookmark = bookmarks["VTEq_" + Guid.Parse(formulaId).ToString("N")];
            numberBookmark = bookmarks["VTEqNum_" + Guid.Parse(formulaId).ToString("N")];
            captionBookmark = bookmarks["VTEqCap_" + Guid.Parse(formulaId).ToString("N")];
            visibleRange = visibleBookmark.Range;
            numberRange = numberBookmark.Range;
            captionRange = captionBookmark.Range;
            foreach (var owned in new[] { visibleRange, numberRange, captionRange })
            {
                AssertTrue((bool)owned.get_Information(Word.WdInformation.wdWithInTable),
                    phase + ": a VTEq numbering bookmark escaped to a body paragraph.");
                AssertTrue(owned.Start >= numberCellRange.Start && owned.End <= numberCellRange.End,
                    phase + ": a VTEq numbering bookmark escaped the right number cell.");
            }
            AssertTrue((visibleRange.Text ?? string.Empty).StartsWith("(", StringComparison.Ordinal)
                       && (visibleRange.Text ?? string.Empty).EndsWith(")", StringComparison.Ordinal),
                phase + ": visible number bookmark is not the complete parenthesized label.");

            numberParagraphs = numberCellRange.Paragraphs;
            AssertEqual(1, numberParagraphs.Count,
                phase + ": number cell contains more than one paragraph.");
            numberParagraph = numberParagraphs[1];
            numberParagraphRange = numberParagraph.Range;
            AssertTrue((numberParagraphRange.Text ?? string.Empty).StartsWith("\t", StringComparison.Ordinal),
                phase + ": number paragraph does not start with a real TAB.");
            AssertTrue((numberParagraphRange.Text ?? string.Empty).EndsWith("\r\a", StringComparison.Ordinal),
                phase + ": number is not followed only by the cell paragraph/cell marks.");

            // Match the VisualTeX OLE numbering contract: the visible label is
            // ordinary Word text that inherits the host paragraph mark's typeface
            // and point size, with no manual baseline shift. Formula font-size edits
            // must therefore not resize or switch the number to the math font.
            numberParagraphMark = document.Range(
                Math.Max(numberParagraphRange.Start, numberParagraphRange.End - 2),
                Math.Max(numberParagraphRange.Start, numberParagraphRange.End - 1));
            numberParagraphMarkFont = numberParagraphMark.Font;
            visibleNumberFont = visibleRange.Font;
            AssertNear(0f, visibleNumberFont.Position, 0.1f,
                phase + ": number has a manual vertical font offset.");
            AssertNear(0f, numberParagraphMarkFont.Position, 0.1f,
                phase + ": number-cell paragraph mark has a manual vertical font offset.");
            AssertNear(numberParagraphMarkFont.Size, visibleNumberFont.Size, 0.1f,
                phase + ": number does not inherit the ordinary Word paragraph point size.");
            var paragraphFontName = numberParagraphMarkFont.NameAscii
                ?? numberParagraphMarkFont.Name
                ?? string.Empty;
            var visibleFontName = visibleNumberFont.NameAscii
                ?? visibleNumberFont.Name
                ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(paragraphFontName)
                && !string.IsNullOrWhiteSpace(visibleFontName))
                AssertEqual(paragraphFontName, visibleFontName,
                    phase + ": number does not inherit the ordinary Word paragraph typeface.");

            numberFormat = numberParagraphRange.ParagraphFormat;
            tabStops = numberFormat.TabStops;
            var sawRight = false;
            var rightPosition = 0f;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop); tabStop = tabStops[index];
                if (tabStop.Alignment != Word.WdTabAlignment.wdAlignTabRight) continue;
                sawRight = true;
                rightPosition = tabStop.Position;
            }
            AssertTrue(sawRight, phase + ": number cell lost its genuine right TabStop.");
            AssertNear(columns[3].Width, rightPosition, 0.8f,
                phase + ": right TabStop is not exactly at the right-cell boundary.");
            AssertNear(columns[1].Width, columns[3].Width, 0.8f,
                phase + ": left/right table gutters are not symmetric, so the formula center moved off the true center-tab axis.");

            sections = formulaRange.Sections;
            section = sections[1];
            setup = section.PageSetup;
            var writableWidth = setup.PageWidth - setup.LeftMargin - setup.RightMargin;
            var tableWidth = columns[1].Width + columns[2].Width + columns[3].Width;
            AssertNear(writableWidth, tableWidth, 1.2f,
                phase + ": the 1x3 table does not span the writable text width.");

            window = document.ActiveWindow;
            view = window.View;
            restoreFieldCodes = view.ShowFieldCodes;
            if (restoreFieldCodes)
            {
                // Physical number geometry is defined by the rendered SEQ result.
                // Alt+F9 / ShowFieldCodes deliberately lays out the long field-code
                // instruction instead, so measure with results visible and restore
                // the user's Word view preference in finally.
                view.ShowFieldCodes = false;
                System.Windows.Forms.Application.DoEvents();
            }
            object scrollStart = true;
            window.ScrollIntoView(formulaRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(100);
            var formulaBox = ReadVisibleMathInkBox(
                document,
                window,
                formulaRange,
                phase + " formula ink");
            var numberBox = ReadWordRangePixelBox(
                window,
                visibleRange,
                phase + " number label");

            // Match a genuine right-tab paragraph in physical coordinates too:
            // the last insertion position after ')' must land exactly at the body
            // right margin, and the visible number must sit on the same Word Y as
            // its ordinary paragraph mark. This catches a table that merely looks
            // right-aligned while using different horizontal/vertical metrics.
            numberEnd = visibleRange.Duplicate;
            numberEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var numberEndPageX = Convert.ToSingle(numberEnd.get_Information(
                Word.WdInformation.wdHorizontalPositionRelativeToPage));
            var expectedPageRight = setup.LeftMargin + writableWidth;
            AssertNear(expectedPageRight, numberEndPageX, 1.2f,
                phase + ": visible number does not physically end on the document right-tab boundary.");
            var numberY = Convert.ToSingle(visibleRange.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            var paragraphMarkY = Convert.ToSingle(numberParagraphMark.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            AssertNear(paragraphMarkY, numberY, 0.75f,
                phase + ": visible number and its ordinary paragraph mark are not on the same Word baseline/Y.");

            var formulaCenterY = formulaBox.Top + formulaBox.Height / 2.0;
            var numberCenterY = numberBox.Top + numberBox.Height / 2.0;
            AssertTrue(Math.Abs(numberCenterY - formulaCenterY) <= 5.0,
                phase + $": table number is not vertically centered with the display formula (delta={numberCenterY - formulaCenterY:0.###}px).");
            AssertTrue(numberBox.Left > formulaBox.Left + formulaBox.Width,
                phase + ": visible number overlaps the centered formula body.");

            Console.WriteLine(
                $"  {phase}: table=1x3 widths={columns[1].Width:0.###}/{columns[2].Width:0.###}/{columns[3].Width:0.###}pt, rightTab={rightPosition:0.###}pt, numberEndPageX={numberEndPageX:0.###}pt/{expectedPageRight:0.###}pt, numberY={numberY:0.###}, markY={paragraphMarkY:0.###}, formula={formulaBox.Left},{formulaBox.Top},{formulaBox.Width},{formulaBox.Height}, number={numberBox.Left},{numberBox.Top},{numberBox.Width},{numberBox.Height}, fieldsInMath={formulaRange.Fields.Count}, fieldsInNumberCell={numberFields.Count}.");
        }
        finally
        {
            if (view is not null && restoreFieldCodes)
            {
                try { view.ShowFieldCodes = true; } catch { }
            }
            Release(view);
            Release(window);
            Release(setup);
            Release(section);
            Release(sections);
            Release(tabStop);
            Release(tabStops);
            Release(numberEnd);
            Release(visibleNumberFont);
            Release(numberParagraphMarkFont);
            Release(numberParagraphMark);
            Release(numberFormat);
            Release(numberParagraphRange);
            Release(numberParagraph);
            Release(numberParagraphs);
            Release(captionRange);
            Release(numberRange);
            Release(visibleRange);
            Release(captionBookmark);
            Release(numberBookmark);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(sequenceCode);
            Release(sequenceField);
            Release(numberFields);
            Release(numberCellRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
            Release(math);
            Release(maths);
            Release(formulaRange);
        }
    }
}

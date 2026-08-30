using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlTableNavigableReferenceAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The 1x3 navigable-reference acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "word-omml-1x3-navigable-reference.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(document, EquationNumberFormat.ContinuousId);

            var formulaId = Guid.NewGuid().ToString("D");
            var service = new WordFormulaService(application);
            var insertion = document.Content.End - 1;
            application.Selection.SetRange(insertion, insertion);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion,
                insertion,
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());
            AssertOmmlTableNumberLifecyclePhase(application, document, formulaId, "01-before-navigable-reference");

            var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
            var target = targets.Single(item => string.Equals(
                item.FormulaId,
                formulaId,
                StringComparison.OrdinalIgnoreCase));
            var selection = application.Selection;
            selection.EndKey(Word.WdUnits.wdStory);
            selection.TypeParagraph();
            selection.TypeText("Reference: ");
            WordEquationNumbering.InsertEquationReference(
                document,
                selection,
                target,
                EquationReferenceStyle.Parenthesized);
            selection.TypeParagraph();

            AssertNavigableOmmlTableReference(
                application,
                document,
                formulaId,
                "02-after-reference-insert");

            // Exercise a real renumber that changes text length/content. The outer
            // GOTOBUTTON must survive while the nested REF follows VTEqNum_.
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);
            AssertTrue(WordEquationNumbering.UpdateEquationNumbers(document) >= 1,
                "The navigable-reference acceptance could not refresh the numbered OMML formula.");
            AssertTrue(WordEquationReferenceFields.UpdateNavigableReferences(document) >= 1,
                "The navigable-reference acceptance did not update its nested REF field.");
            AssertNavigableOmmlTableReference(
                application,
                document,
                formulaId,
                "03-after-heading-format-update");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertOmmlTableNumberLifecyclePhase(application, document, formulaId, "04-reference-reopen-host");
            AssertTrue(WordEquationReferenceFields.UpdateNavigableReferences(document) >= 1,
                "Save/reopen did not expose the navigable nested REF for refresh.");
            AssertNavigableOmmlTableReference(
                application,
                document,
                formulaId,
                "05-after-reference-reopen");

            Console.WriteLine(
                "Word OMML 1x3 navigable-reference acceptance passed: a real GOTOBUTTON VTEqNum field retained exactly one nested REF VTEqNum, its displayed number followed format changes, DoClick navigated into the same unique 1x3 equation row, and save/reopen preserved the complete field tree.");
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

    private static void AssertNavigableOmmlTableReference(
        Word.Application application,
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Field? goTo = null;
        Word.Range? goToCode = null;
        Word.Fields? nestedFields = null;
        Word.Field? nestedRef = null;
        Word.Range? nestedCode = null;
        Word.Range? nestedResult = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        Word.Table? table = null;
        Word.Cell? centerCell = null;
        Word.Range? centerRange = null;
        Word.Selection? selection = null;
        try
        {
            var targetName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                if (field.Type != Word.WdFieldType.wdFieldGoToButton) continue;
                code = field.Code;
                var text = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (!text.StartsWith("GOTOBUTTON " + targetName, StringComparison.OrdinalIgnoreCase))
                    continue;
                goTo = field;
                field = null;
                break;
            }
            AssertTrue(goTo is not null,
                context + ": GOTOBUTTON VTEqNum field is missing.");
            goToCode = goTo!.Code;
            nestedFields = goToCode.Fields;
            AssertEqual(1, nestedFields.Count,
                context + ": GOTOBUTTON does not contain exactly one nested field.");
            nestedRef = nestedFields[1];
            AssertEqual(Word.WdFieldType.wdFieldRef, nestedRef.Type,
                context + ": nested field is not REF.");
            nestedCode = nestedRef.Code;
            var nestedInstruction = (nestedCode.Text ?? string.Empty)
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('\t', ' ');
            AssertTrue(nestedInstruction.IndexOf(targetName, StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": nested REF does not target the same VTEqNum bookmark.");
            nestedRef.Update();
            nestedResult = nestedRef.Result;
            var expected = ReadVisibleEquationNumber(document, formulaId).Trim('(', ')');
            AssertEqual(expected, NormalizeEquationNumberText(nestedResult.Text),
                context + ": nested REF result does not match the current equation number.");

            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(targetName), context + ": VTEqNum target bookmark is missing.");
            numberBookmark = bookmarks[targetName];
            numberRange = numberBookmark.Range;
            AssertTrue((bool)numberRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": VTEqNum target escaped the managed table.");
            table = numberRange.Tables[1];
            AssertEqual(1, table.Rows.Count, context + ": navigation target table gained extra rows.");
            AssertEqual(3, table.Columns.Count, context + ": navigation target is not a 1x3 table.");
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            AssertEqual(1, centerRange.OMaths.Count,
                context + ": navigation target table does not own exactly one center OMath.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, centerRange.OMaths[1].Type,
                context + ": navigation target center OMath is not Display.");

            // Updating the nested REF can rematerialize GOTOBUTTON, so reacquire the
            // outer field immediately before simulating the user's double-click.
            Release(goTo); goTo = null;
            Release(goToCode); goToCode = null;
            Release(nestedFields); nestedFields = null;
            Release(nestedRef); nestedRef = null;
            Release(nestedCode); nestedCode = null;
            Release(nestedResult); nestedResult = null;
            Release(code); code = null;
            Release(field); field = null;
            Release(fields); fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = fields[index];
                if (field.Type != Word.WdFieldType.wdFieldGoToButton) continue;
                code = field.Code;
                var text = (code.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (!text.StartsWith("GOTOBUTTON " + targetName, StringComparison.OrdinalIgnoreCase))
                    continue;
                goTo = field;
                field = null;
                break;
            }
            AssertTrue(goTo is not null, context + ": live GOTOBUTTON disappeared after REF update.");
            goTo!.DoClick();
            selection = application.Selection;
            AssertTrue(selection.Start >= numberRange.Start && selection.Start <= numberRange.End,
                context + $": GOTOBUTTON did not navigate to VTEqNum; selection={selection.Start}:{selection.End}, target={numberRange.Start}:{numberRange.End}.");
            AssertTrue((bool)selection.Range.get_Information(Word.WdInformation.wdWithInTable),
                context + ": GOTOBUTTON navigation did not land in the equation table.");
            AssertEqual(table.Range.Start, selection.Range.Tables[1].Range.Start,
                context + ": GOTOBUTTON navigation landed in a different table from the center formula.");
            Console.WriteLine(
                $"  {context}: target={targetName}, number='{expected}', targetRange={numberRange.Start}:{numberRange.End}, selection={selection.Start}:{selection.End}, table={table.Range.Start}:{table.Range.End}.");
        }
        finally
        {
            Release(selection);
            Release(centerRange);
            Release(centerCell);
            Release(table);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(nestedResult);
            Release(nestedCode);
            Release(nestedRef);
            Release(nestedFields);
            Release(goToCode);
            Release(goTo);
            Release(code);
            Release(field);
            Release(fields);
        }
    }
}

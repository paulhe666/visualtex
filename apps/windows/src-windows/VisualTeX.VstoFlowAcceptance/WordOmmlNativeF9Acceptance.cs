using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeF9Acceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-f9.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Field? externalReference = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var firstId = Guid.NewGuid().ToString("D");
            var middleId = Guid.NewGuid().ToString("D");
            var lastId = Guid.NewGuid().ToString("D");

            InsertNativeF9FormulaAt(
                application,
                document,
                service,
                firstId,
                Math.Max(document.Content.Start, document.Content.End - 1),
                @"a=1");
            InsertNativeF9FormulaAt(
                application,
                document,
                service,
                lastId,
                Math.Max(document.Content.Start, document.Content.End - 1),
                @"c=3");

            Word.Range? middleInsertion = null;
            try
            {
                // Ask the product for the same ordinary, independently deletable
                // paragraph a user receives after pressing Enter at the first
                // formula's number. The returned paragraph-object Range carries
                // main-story affinity; recreating it from bare integer coordinates
                // at a 1-character table separator is ambiguous in Word and can
                // target either neighboring table.
                middleInsertion = WordEquationNumbering
                    .EnsureNormalTypingParagraphAfterNumberedDisplay(
                        document,
                        firstId)
                    ?? throw new InvalidDataException(
                        "Native F9 acceptance could not create a body insertion paragraph between the first and last formulas.");
                InsertNativeF9FormulaAtRange(
                    application,
                    document,
                    service,
                    middleId,
                    middleInsertion,
                    @"b=2");
            }
            finally { Release(middleInsertion); }
            AssertEqual(3, document.Tables.Count,
                "Inserting a numbered formula between two 1x3 hosts merged or nested a table.");
            externalReference = InsertExternalEquationReference(document, lastId);
            AssertEqual(3, WordEquationNumbering.UpdateEquationNumbers(document),
                "VisualTeX did not finalize all three direct-table numbers after middle insertion.");
            var firstSequenceCode = ReadNativeF9SequenceCode(document, firstId);
            var middleSequenceCode = ReadNativeF9SequenceCode(document, middleId);
            var lastSequenceCode = ReadNativeF9SequenceCode(document, lastId);

            UpdateMainStoryFieldsLikeSelectAllF9(document);
            AssertNativeF9Number(document, firstId, "1", "after middle insertion first");
            AssertNativeF9Number(document, middleId, "2", "after middle insertion second");
            AssertNativeF9Number(document, lastId, "3", "after middle insertion third");
            AssertEqual(
                firstSequenceCode,
                ReadNativeF9SequenceCode(document, firstId),
                "F9 rewrote the first direct-table SEQ instruction.");
            AssertEqual(
                middleSequenceCode,
                ReadNativeF9SequenceCode(document, middleId),
                "F9 rewrote the middle direct-table SEQ instruction.");
            AssertEqual(
                lastSequenceCode,
                ReadNativeF9SequenceCode(document, lastId),
                "F9 rewrote the last direct-table SEQ instruction.");
            AssertEqual(
                "3",
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Select-all F9 did not update the body REF to the third formula.");
            AssertExternalEquationReferenceHyperlink(
                externalReference,
                WordEquationNumbering.NativeNumberBookmarkName(lastId));

            Word.Table? middleTable = null;
            try
            {
                middleTable = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        middleId)
                    ?? throw new InvalidDataException(
                        "Native F9 acceptance could not locate the middle 1x3 table before deletion.");
                middleTable.Delete();
            }
            finally { Release(middleTable); }
            WordOmmlFormulaStore.Delete(document, middleId);
            AssertEqual(2, document.Tables.Count,
                "Deleting the middle numbered formula left an empty 1x3 table behind.");

            AssertEqual(2, WordEquationNumbering.UpdateEquationNumbers(document),
                "VisualTeX did not renumber both remaining direct-table formulas after middle deletion.");
            var firstSequenceCodeAfterDeletion =
                ReadNativeF9SequenceCode(document, firstId);
            var lastSequenceCodeAfterDeletion =
                ReadNativeF9SequenceCode(document, lastId);
            UpdateMainStoryFieldsLikeSelectAllF9(document);
            AssertNativeF9Number(document, firstId, "1", "after middle deletion first");
            AssertNativeF9Number(document, lastId, "2", "after middle deletion second");
            AssertEqual(
                firstSequenceCodeAfterDeletion,
                ReadNativeF9SequenceCode(document, firstId),
                "F9 rewrote the first direct-table SEQ instruction after VisualTeX renumbered the deletion.");
            AssertEqual(
                lastSequenceCodeAfterDeletion,
                ReadNativeF9SequenceCode(document, lastId),
                "F9 rewrote the last direct-table SEQ instruction after VisualTeX renumbered the deletion.");
            AssertEqual(
                "2",
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Select-all F9 did not update the existing body REF after middle deletion.");
            AssertExternalEquationReferenceHyperlink(
                externalReference,
                WordEquationNumbering.NativeNumberBookmarkName(lastId));
            AssertEqual(0, document.Shapes.Count,
                "Native F9 lifecycle created a floating Shape.");
            AssertEqual(2, document.Tables.Count,
                "Native F9 lifecycle did not retain one direct-SEQ 1x3 table per remaining formula.");

            document.Save();
            Release(externalReference); externalReference = null;
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
            UpdateMainStoryFieldsLikeSelectAllF9(document);
            AssertNativeF9Number(document, firstId, "1", "save/reopened F9 first");
            AssertNativeF9Number(document, lastId, "2", "save/reopened F9 second");
            AssertEqual(
                firstSequenceCodeAfterDeletion,
                ReadNativeF9SequenceCode(document, firstId),
                "Save/reopen/F9 rewrote the first direct-table SEQ instruction.");
            AssertEqual(
                lastSequenceCodeAfterDeletion,
                ReadNativeF9SequenceCode(document, lastId),
                "Save/reopen/F9 rewrote the last direct-table SEQ instruction.");
            externalReference = FindExternalEquationReference(document, lastId)
                ?? throw new InvalidDataException(
                    "Save/reopen lost the native F9 body REF field.");
            externalReference.Update();
            AssertEqual(
                "2",
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Save/reopen did not retain the F9-renumbered body REF.");
            AssertExternalEquationReferenceHyperlink(
                externalReference,
                WordEquationNumbering.NativeNumberBookmarkName(lastId));

            AssertEqual(2, document.Tables.Count,
                "Save/reopen/F9 did not retain the two remaining direct-SEQ 1x3 tables.");
            Console.WriteLine(
                "Native OMML F9 acceptance passed: VisualTeX planned 1/2/3 after middle insertion and 1/2 after middle deletion, direct main-story F9 preserved those right-cell SEQ instructions/results and body REF, and save/reopen retained zero Shape objects plus one 1x3 table per formula.");
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

    private static void InsertNativeF9FormulaAt(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        int position,
        string latex)
    {
        Word.Range? insertion = null;
        try
        {
            var clampedPosition = Math.Max(
                document.Content.Start,
                Math.Min(position, document.Content.End - 1));
            insertion = document.Range(clampedPosition, clampedPosition);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex,
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());
        }
        finally
        {
            Release(insertion);
        }
    }

    private static void InsertNativeF9FormulaAtRange(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        Word.Range insertionRange,
        string latex)
    {
        Word.Range? insertion = null;
        try
        {
            insertion = insertionRange.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertion.Select();
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertion.Start,
                insertion.End,
                latex,
                originalMetadata: null);
            service.InsertOmml(session, QuadraticFormulaMathMl());
        }
        finally
        {
            Release(insertion);
        }
    }

    private static void UpdateMainStoryFieldsLikeSelectAllF9(Word.Document document)
    {
        Word.Range? content = null;
        Word.Fields? fields = null;
        try
        {
            content = document.Content;
            fields = content.Fields;
            AssertTrue(fields.Count > 0,
                "The native F9 acceptance main story contains no fields.");
            fields.Update();
        }
        finally
        {
            Release(fields);
            Release(content);
        }
    }

    private static string ReadNativeF9SequenceCode(
        Word.Document document,
        string formulaId)
    {
        Word.Range? numberRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            numberRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    $"Native F9 formula {formulaId} lost its direct-table visible number.");
            AssertTrue((bool)numberRange.get_Information(
                    Word.WdInformation.wdWithInTable),
                $"Native F9 formula {formulaId} moved its number outside the 1x3 table.");
            fields = numberRange.Fields;
            AssertEqual(1, fields.Count,
                $"Native F9 formula {formulaId} does not contain one right-cell SEQ field.");
            field = fields[1];
            code = field.Code;
            return NormalizeFieldCode(code.Text ?? string.Empty);
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(numberRange);
        }
    }

    private static void AssertNativeF9Number(
        Word.Document document,
        string formulaId,
        string expected,
        string context)
    {
        AssertOmmlTabNumberingHost(
            document,
            formulaId,
            "native F9 " + context,
            updateReference: false);
        AssertEqual(
            expected,
            ReadNativeNumberBookmark(document, formulaId),
            "Native F9 numbering mismatch " + context + ".");
    }
}

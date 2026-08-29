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
        Word.Range? owner = null;
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

            owner = WordEquationNumbering.FindNumberingOwnerRange(document, lastId)
                ?? throw new InvalidDataException(
                    "Native F9 acceptance could not locate the last formula owner.");
            var middlePosition = owner.Start;
            Release(owner); owner = null;
            InsertNativeF9FormulaAt(
                application,
                document,
                service,
                middleId,
                middlePosition,
                @"b=2");
            externalReference = InsertExternalEquationReference(document, lastId);
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
                "F9 rewrote the first mathematical SEQ instruction.");
            AssertEqual(
                middleSequenceCode,
                ReadNativeF9SequenceCode(document, middleId),
                "F9 rewrote the middle mathematical SEQ instruction.");
            AssertEqual(
                lastSequenceCode,
                ReadNativeF9SequenceCode(document, lastId),
                "F9 rewrote the last mathematical SEQ instruction.");
            AssertEqual(
                "3",
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Select-all F9 did not update the body REF to the third formula.");
            AssertExternalEquationReferenceHyperlink(
                externalReference,
                WordEquationNumbering.NativeNumberBookmarkName(lastId));

            owner = WordEquationNumbering.FindNumberingOwnerRange(document, middleId)
                ?? throw new InvalidDataException(
                    "Native F9 acceptance could not locate the middle formula owner before deletion.");
            var ownerStart = owner.Start;
            var ownerEnd = owner.End;
            AssertTrue(ownerEnd > ownerStart,
                "Native F9 acceptance received an empty owner range.");
            owner.Delete();
            Release(owner); owner = null;
            WordOmmlFormulaStore.Delete(document, middleId);

            UpdateMainStoryFieldsLikeSelectAllF9(document);
            AssertNativeF9Number(document, firstId, "1", "after middle deletion first");
            AssertNativeF9Number(document, lastId, "2", "after middle deletion second");
            AssertEqual(
                firstSequenceCode,
                ReadNativeF9SequenceCode(document, firstId),
                "Middle deletion/F9 rewrote the first mathematical SEQ instruction.");
            AssertEqual(
                lastSequenceCode,
                ReadNativeF9SequenceCode(document, lastId),
                "Middle deletion/F9 rewrote the last mathematical SEQ instruction.");
            AssertEqual(
                "2",
                NormalizeEquationNumberText(externalReference.Result.Text),
                "Select-all F9 did not update the existing body REF after middle deletion.");
            AssertExternalEquationReferenceHyperlink(
                externalReference,
                WordEquationNumbering.NativeNumberBookmarkName(lastId));
            AssertEqual(0, document.Shapes.Count,
                "Native F9 lifecycle created a floating Shape.");
            AssertEqual(0, document.Tables.Count,
                "Native F9 lifecycle created a Word table.");

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
                firstSequenceCode,
                ReadNativeF9SequenceCode(document, firstId),
                "Save/reopen/F9 rewrote the first mathematical SEQ instruction.");
            AssertEqual(
                lastSequenceCode,
                ReadNativeF9SequenceCode(document, lastId),
                "Save/reopen/F9 rewrote the last mathematical SEQ instruction.");
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

            Console.WriteLine(
                "Native OMML F9 acceptance passed: direct main-story field update produced 1/2/3 after middle insertion and 1/2 after middle deletion, then survived save/reopen with no VisualTeX code rewrite, Shape or Table.");
        }
        finally
        {
            Release(externalReference);
            Release(owner);
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
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Native F9 formula {formulaId} lost its VTOMML identity.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            fields = formulaRange.Fields;
            AssertEqual(1, fields.Count,
                $"Native F9 formula {formulaId} does not contain one SEQ field.");
            field = fields[1];
            code = field.Code;
            return NormalizeFieldCode(code.Text ?? string.Empty);
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(formulaRange);
            Release(bookmark);
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

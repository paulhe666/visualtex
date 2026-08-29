using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string MixedSequenceBookmarkName = "VTMixedOleSequence";

    private static void RunWordMixedVisualTeXSequenceAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-mixed-visualtex-sequence.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Field? middleSequence = null;
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
            var firstFormulaId = Guid.NewGuid().ToString("D");
            var lastFormulaId = Guid.NewGuid().ToString("D");

            InsertMixedSequenceOmmlFormula(
                application,
                document,
                service,
                firstFormulaId,
                @"x=1");
            middleSequence = InsertOleLikeVisualTeXSequenceField(document);
            InsertMixedSequenceOmmlFormula(
                application,
                document,
                service,
                lastFormulaId,
                @"y=2");

            UpdateMixedVisualTeXSequenceInDocumentOrder(
                document,
                firstFormulaId,
                middleSequence,
                lastFormulaId);
            AssertMixedVisualTeXSequence(
                document,
                firstFormulaId,
                middleSequence,
                lastFormulaId,
                "initial mixed native/OLE-like sequence");

            document.Save();
            Release(middleSequence); middleSequence = null;
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
            middleSequence = FindMixedSequenceField(document)
                ?? throw new InvalidDataException(
                    "Save/reopen lost the OLE-like VisualTeXEquation SEQ field.");
            UpdateMixedVisualTeXSequenceInDocumentOrder(
                document,
                firstFormulaId,
                middleSequence,
                lastFormulaId);
            AssertMixedVisualTeXSequence(
                document,
                firstFormulaId,
                middleSequence,
                lastFormulaId,
                "save/reopened mixed native/OLE-like sequence");

            Console.WriteLine(
                "Mixed VisualTeX sequence acceptance passed: native #(SEQ), an ordinary OLE-style caption SEQ, and native #(SEQ) remained one dynamic 1/2/3 stream after save/reopen, with zero Shape/Table hosts.");
        }
        finally
        {
            Release(middleSequence);
            Release(insertion);
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

    private static void InsertMixedSequenceOmmlFormula(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string latex)
    {
        Word.Range? insertion = null;
        try
        {
            var position = Math.Max(document.Content.Start, document.Content.End - 1);
            insertion = document.Range(position, position);
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

    private static Word.Field InsertOleLikeVisualTeXSequenceField(
        Word.Document document)
    {
        Word.Range? insertion = null;
        Word.Range? resultRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Bookmarks? bookmarks = null;
        try
        {
            var position = Math.Max(document.Content.Start, document.Content.End - 1);
            insertion = document.Range(position, position);
            insertion.InsertParagraphAfter();
            position = Math.Max(document.Content.Start, document.Content.End - 1);
            Release(insertion);
            insertion = document.Range(position, position);
            fields = insertion.Fields;
            object fieldType = Word.WdFieldType.wdFieldEmpty;
            object fieldCode = "SEQ VisualTeXEquation \\* ARABIC";
            object preserveFormatting = true;
            field = fields.Add(
                insertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            field.Update();
            resultRange = field.Result;
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(MixedSequenceBookmarkName))
                bookmarks[MixedSequenceBookmarkName].Delete();
            bookmarks.Add(MixedSequenceBookmarkName, resultRange);
            var result = field;
            field = null;
            return result;
        }
        finally
        {
            Release(bookmarks);
            Release(resultRange);
            Release(field);
            Release(fields);
            Release(insertion);
        }
    }

    private static void UpdateMixedVisualTeXSequenceInDocumentOrder(
        Word.Document document,
        string firstFormulaId,
        Word.Field middleSequence,
        string lastFormulaId)
    {
        UpdateNativeOmmlSequenceField(document, firstFormulaId);
        middleSequence.Update();
        UpdateNativeOmmlSequenceField(document, lastFormulaId);
    }

    private static void UpdateNativeOmmlSequenceField(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Mixed sequence formula {formulaId} lost its VTOMML identity.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            fields = formulaRange.Fields;
            AssertEqual(1, fields.Count,
                $"Mixed sequence formula {formulaId} does not contain one mathematical SEQ.");
            field = fields[1];
            field.Update();
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static void AssertMixedVisualTeXSequence(
        Word.Document document,
        string firstFormulaId,
        Word.Field middleSequence,
        string lastFormulaId,
        string context)
    {
        AssertEqual("1", ReadNativeNumberBookmark(document, firstFormulaId),
            context + ": first native OMML number is not 1.");
        AssertEqual("2", NormalizeEquationNumberText(middleSequence.Result.Text),
            context + ": the OLE-like middle sequence number is not 2.");
        AssertEqual("3", ReadNativeNumberBookmark(document, lastFormulaId),
            context + ": last native OMML number is not 3.");
        AssertEqual(0, document.Shapes.Count,
            context + ": a floating Shape was created.");
        AssertEqual(0, document.Tables.Count,
            context + ": a Word table was created.");
        AssertOmmlTabNumberingHost(
            document,
            firstFormulaId,
            context + " first native host",
            updateReference: false);
        AssertOmmlTabNumberingHost(
            document,
            lastFormulaId,
            context + " last native host",
            updateReference: false);
    }

    private static string ReadNativeNumberBookmark(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertTrue(bookmarks.Exists(name),
                $"Mixed sequence formula {formulaId} lost its VTEqNum bookmark.");
            bookmark = bookmarks[name];
            range = bookmark.Range;
            return NormalizeEquationNumberText(range.Text);
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Word.Field? FindMixedSequenceField(Word.Document document)
    {
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        Word.Fields? fields = null;
        Word.Field? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(MixedSequenceBookmarkName)) return null;
            bookmark = bookmarks[MixedSequenceBookmarkName];
            bookmarkRange = bookmark.Range;
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                Word.Range? fieldResult = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    var instruction = code.Text ?? string.Empty;
                    if (instruction.IndexOf(
                            "SEQ VisualTeXEquation",
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    fieldResult = field.Result;
                    if (fieldResult.Start > bookmarkRange.Start
                        || fieldResult.End < bookmarkRange.End)
                        continue;
                    result = field;
                    field = null;
                    return result;
                }
                finally
                {
                    Release(fieldResult);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally
        {
            Release(fields);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }
}

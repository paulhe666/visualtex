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
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"mixed-sequence-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "formula.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"60\" viewBox=\"0 0 160 60\"><text x=\"4\" y=\"42\" font-size=\"32\">m = 2</text></svg>");
        var emfPath = VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview
            .CreateVectorEmfFromSvg(svgPath, 160, 60);
        var pngDataUrl =
            CreatePngDataUrl("mixed-sequence", 160, 60);
        var pngPath =
            Path.Combine(assetRoot, "formula.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(
                pngDataUrl.Substring(
                    pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
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
            var middleFormulaId = Guid.NewGuid().ToString("D");
            var lastFormulaId = Guid.NewGuid().ToString("D");

            // Create the native OLE first. In the isolated Office
            // acceptance host, initializing a new native OLE after an existing
            // OMath can be denied by Office's OLE ACL even though the same
            // product path is valid interactively. After the OLE exists, insert
            // native OMML on both sides and validate the final document order.
            InsertMixedSequenceVisualTeXFormula(
                application,
                document,
                service,
                middleFormulaId,
                pngPath,
                emfPath);
            InsertMixedSequenceOmmlFormulaAtDocumentStart(
                application,
                document,
                service,
                firstFormulaId,
                @"x=1");
            InsertMixedSequenceOmmlFormula(
                application,
                document,
                service,
                lastFormulaId,
                @"y=3");

            _ = WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                document);
            AssertCurrentMixedVisualTeXSequence(
                document,
                service,
                middleFormulaId,
                "initial mixed native/VisualTeX/native sequence");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            service = new WordFormulaService(application);
            _ = WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                document);
            AssertCurrentMixedVisualTeXSequence(
                document,
                service,
                middleFormulaId,
                "save/reopened mixed native/VisualTeX/native sequence");

            Console.WriteLine(
                "Mixed VisualTeX sequence acceptance passed: OMML keeps its native Word Equation stream (1/2) while VisualTeXPlaceRef keeps an independent VisualTeXEquation stream (1), before and after save/reopen.");
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
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void InsertMixedSequenceOmmlFormulaAtDocumentStart(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string latex)
    {
        Word.Range? boundary = null;
        Word.Range? insertion = null;
        try
        {
            boundary =
                document.Range(
                    document.Content.Start,
                    document.Content.Start);
            boundary.InsertBefore("\r");
            Release(boundary);
            boundary = null;

            insertion =
                document.Range(
                    document.Content.Start,
                    document.Content.Start);
            application.Selection.SetRange(
                insertion.Start,
                insertion.End);
            var session =
                CreateNumberedOmmlTabSession(
                    formulaId,
                    document.FullName,
                    insertion.Start,
                    insertion.End,
                    latex,
                    originalMetadata: null);
            service.InsertOmml(
                session,
                QuadraticFormulaMathMl());
        }
        finally
        {
            Release(insertion);
            Release(boundary);
        }
    }

    private static void InsertMixedSequenceVisualTeXFormula(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        string pngPath,
        string emfPath)
    {
        Word.Range? insertion = null;
        try
        {
            var position = Math.Max(
                document.Content.Start,
                document.Content.End - 1);
            insertion =
                document.Range(
                    position,
                    position);
            application.Selection.SetRange(
                insertion.Start,
                insertion.End);
            var session =
                CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(
                        insertion.Start,
                        insertion.End),
                    originalMetadata: null,
                    latex: @"m=2");
            session.Numbered = true;
            session.ExportResult =
                new VisualTeX.WindowsOffice.Contracts.OfficeExportDocument
                {
                    Width = 160,
                    Height = 60,
                    Baseline = 45,
                };
            service.InsertOle(
                session,
                pngPath,
                emfPath);
        }
        finally
        {
            Release(insertion);
        }
    }

    private static void AssertCurrentMixedVisualTeXSequence(
        Word.Document document,
        WordFormulaService service,
        string middleFormulaId,
        string context)
    {
        var targets =
            service.GetCanonicalEquationReferenceTargets(
                    document)
                .Where(target =>
                    target.Source is
                        EquationReferenceSource.WordOmml
                        or EquationReferenceSource.VisualTeX)
                .OrderBy(target =>
                    target.Position)
                .ToArray();
        AssertEqual(
            3,
            targets.Length,
            context
            + ": current mixed document does not expose exactly three numbered targets.");
        AssertEqual(
            "1",
            targets[0].NumberText,
            context
            + ": first native OMML number is not 1.");
        AssertEqual(
            "1",
            targets[1].NumberText,
            context
            + ": middle VisualTeX number is not 1 in its independent sequence.");
        AssertEqual(
            EquationReferenceSource.VisualTeX,
            targets[1].Source,
            context
            + ": middle numbered target is not VisualTeX.");
        AssertEqual(
            middleFormulaId,
            targets[1].FormulaId,
            context
            + ": middle VisualTeX numbered target changed identity.");
        AssertEqual(
            "2",
            targets[2].NumberText,
            context
            + ": last native OMML number is not 2 in Word's independent Equation sequence.");

        Word.Range? owner = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        try
        {
            owner =
                WordVisualTeXParagraphNumbering
                    .FindOwnerParagraphRange(
                        document,
                        middleFormulaId)
                ?? throw new InvalidDataException(
                    context
                    + ": middle VisualTeXPlaceRef owner paragraph is missing.");
            fields =
                owner.Fields;
            var nativeEquationSeqCount = 0;
            var visualTeXSequenceCount = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                code =
                    field.Code.Duplicate;
                var instruction =
                    code.Text
                    ?? string.Empty;
                if (instruction.IndexOf(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    visualTeXSequenceCount++;
                if (WordNativeOmmlNumbering
                    .IsEquationSequenceFieldCode(
                        document,
                        instruction))
                    nativeEquationSeqCount++;
            }
            AssertTrue(
                visualTeXSequenceCount >= 1,
                context
                + ": fresh VisualTeXPlaceRef does not own its VisualTeXEquation sequence.");
            AssertEqual(
                0,
                nativeEquationSeqCount,
                context
                + ": VisualTeXPlaceRef leaked into Word OMML's native Equation sequence.");
            AssertEqual(
                0,
                document.Shapes.Count,
                context
                + ": a floating Shape was created.");
            AssertEqual(
                0,
                document.Tables.Count,
                context
                + ": a Word table was created.");
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(owner);
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

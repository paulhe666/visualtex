using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlCopyPasteReeditAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? sourceBookmark = null;
        Word.Bookmark? copyBookmark = null;
        Word.Range? sourceRange = null;
        Word.Range? copyRange = null;
        var path = Path.Combine(artifactRoot, "word-omml-copy-paste-reedit.docx");
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "source: \rcopy: \r";
            document.Activate();
            var service = new WordFormulaService(application);
            var sourceFormulaId = Guid.NewGuid().ToString("D");
            var sourcePosition = "source: ".Length;
            application.Selection.SetRange(sourcePosition, sourcePosition);
            service.InsertOmml(
                CopyPasteOmmlSession(document, "create", sourceFormulaId, "a_1+b_1",
                    WordRangeReference(sourcePosition, sourcePosition), null),
                CopyPasteMathMl(1));

            sourceBookmark = WordOmmlFormulaStore.FindByFormulaId(document, sourceFormulaId)
                ?? throw new InvalidOperationException("OMML source did not receive VTOMML identity.");
            sourceRange = WordOmmlFormulaStore.GetEquationRange(sourceBookmark);
            sourceRange.Copy();
            var paragraph = document.Paragraphs[2];
            try
            {
                var pasteAt = paragraph.Range.Start + "copy: ".Length;
                application.Selection.SetRange(pasteAt, pasteAt);
                application.Selection.Paste();
            }
            finally { Release(paragraph); }
            AssertEqual(2, document.OMaths.Count, "Word did not paste OMML copy.");

            copyRange = document.OMaths[2].Range;
            copyRange.Select();
            var copied = service.ReadSelection();
            AssertTrue(copied.Metadata is not null && !string.IsNullOrWhiteSpace(copied.FormulaId),
                "Pasted OMML copy was not adopted as editable VisualTeX formula.");
            AssertEqual(FormulaOleContract.WordOmmlMode, copied.ObjectMode,
                "Pasted OMML copy lost Word OMML mode.");
            AssertTrue(!string.Equals(sourceFormulaId, copied.FormulaId, StringComparison.OrdinalIgnoreCase),
                "Pasted OMML copy reused source FormulaId.");
            var copyFormulaId = copied.FormulaId!;
            copyBookmark = WordOmmlFormulaStore.FindByFormulaId(document, copyFormulaId)
                ?? throw new InvalidOperationException("Adopted OMML copy did not receive durable VTOMML bookmark.");

            service.ReplaceOmml(
                CopyPasteOmmlSession(document, "edit", copyFormulaId, "a_2+b_2",
                    copied.ObjectId, copied.Metadata),
                CopyPasteMathMl(2));
            copied = SelectOmmlAndRead(document, service, copyFormulaId);
            AssertEqual(copyFormulaId, copied.FormulaId,
                "OMML copy changed FormulaId after first VisualTeX edit.");
            AssertTrue((copied.Metadata?.Lines.FirstOrDefault()?.Latex ?? string.Empty).Contains("2"),
                "First OMML copy edit did not persist updated source.");

            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            document.Activate();
            service = new WordFormulaService(application);
            copied = SelectOmmlAndRead(document, service, copyFormulaId);
            AssertEqual(copyFormulaId, copied.FormulaId,
                "OMML copy became uneditable after first edit and save/reopen.");
            AssertTrue(WordDoubleClickRouting.ShouldOpenVisualTeX(copied),
                "Reopened OMML copy would not route to VisualTeX.");

            service.ReplaceOmml(
                CopyPasteOmmlSession(document, "edit", copyFormulaId, "a_3+b_3",
                    copied.ObjectId, copied.Metadata),
                CopyPasteMathMl(3));
            copied = SelectOmmlAndRead(document, service, copyFormulaId);
            AssertEqual(copyFormulaId, copied.FormulaId,
                "OMML copy changed FormulaId after second VisualTeX edit.");
            AssertTrue((copied.Metadata?.Lines.FirstOrDefault()?.Latex ?? string.Empty).Contains("3"),
                "Second OMML copy edit did not persist updated source.");

            sourceBookmark = WordOmmlFormulaStore.FindByFormulaId(document, sourceFormulaId)
                ?? throw new InvalidOperationException("Editing OMML copy removed source bookmark.");
            var sourceMetadata = WordOmmlFormulaStore.TryRead(document, sourceBookmark)
                ?? throw new InvalidOperationException("Editing OMML copy removed source metadata.");
            AssertEqual(sourceFormulaId, sourceMetadata.FormulaId,
                "Editing OMML copy changed source formula identity.");
            document.Save();
            Console.WriteLine(
                $"Word OMML copy/paste re-edit passed: source={sourceFormulaId}, copy={copyFormulaId}, "
                + "two edits + save/reopen remained independently editable.");
        }
        finally
        {
            Release(copyRange);
            Release(sourceRange);
            Release(copyBookmark);
            Release(sourceBookmark);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSelection SelectOmmlAndRead(
        Word.Document document,
        WordFormulaService service,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidOperationException($"OMML formula {formulaId} has no VTOMML bookmark.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            range.Select();
            return service.ReadSelection();
        }
        finally
        {
            Release(range);
            Release(bookmark);
        }
    }

    private static OfficeSessionDocument CopyPasteOmmlSession(
        Word.Document document,
        string mode,
        string formulaId,
        string latex,
        string? sourceObjectId,
        FormulaMetadata? originalMetadata)
    {
        var lineId = originalMetadata?.Lines.FirstOrDefault()?.Id;
        if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString("D");
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Host = "word",
            Mode = mode,
            FormulaId = formulaId,
            SourceDocumentId = document.FullName,
            SourceObjectId = sourceObjectId,
            Title = "Word OMML copy/paste re-edit acceptance",
            CodeFormat = "latex",
            DisplayMode = "inline",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = false,
            FontSizePt = 14,
            OriginalMetadata = originalMetadata,
            Lines = new List<FormulaLine> { new() { Id = lineId!, Latex = latex } },
            ExportResult = new OfficeExportDocument
            {
                Width = 160,
                Height = 32,
                Baseline = 24,
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
    }

    private static string CopyPasteMathMl(int index) =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\">"
        + $"<msub><mi>a</mi><mn>{index}</mn></msub><mo>+</mo>"
        + $"<msub><mi>b</mi><mn>{index}</mn></msub></math>";
}

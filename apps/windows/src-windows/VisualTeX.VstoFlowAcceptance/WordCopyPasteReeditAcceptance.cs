using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordCopyPasteReeditAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? identity = null;
        Word.Bookmark? typingBoundary = null;
        Word.Range? typingBoundaryRange = null;
        Word.Range? followingProse = null;
        string? png1 = null;
        string? emf1 = null;
        string? png2 = null;
        string? emf2 = null;
        string? png3 = null;
        string? emf3 = null;
        var path = Path.Combine(artifactRoot, "word-copy-paste-reedit.docx");
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "source: 后文\rcopy: \r";
            document.Activate();
            var service = new WordFormulaService(application);
            var oleAssetRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp");
            Directory.CreateDirectory(oleAssetRoot);
            var assets1 = CreatePerformanceOleAssets(oleAssetRoot, 901, "x_1+y_1", PerformanceSvg(901));
            var assets2 = CreatePerformanceOleAssets(oleAssetRoot, 902, "x_2+y_2", PerformanceSvg(902));
            var assets3 = CreatePerformanceOleAssets(oleAssetRoot, 903, "x_3+y_3", PerformanceSvg(903));
            (png1, emf1) = assets1;
            (png2, emf2) = assets2;
            (png3, emf3) = assets3;

            var sourcePosition = "source: ".Length;
            application.Selection.SetRange(sourcePosition, sourcePosition);
            var sourceFormulaId = Guid.NewGuid().ToString("D");
            service.InsertOle(
                CopyPasteOleSession(document, "create", sourceFormulaId, "x_1+y_1",
                    WordRangeReference(sourcePosition, sourcePosition), null, 901),
                png1,
                emf1);
            AssertEqual(1, document.InlineShapes.Count, "OLE copy fixture did not create source formula.");

            shape = document.InlineShapes[1];
            bookmarks = document.Bookmarks;
            AssertTrue(bookmarks.Exists(WordFormulaMetadataReader.IdentityBookmarkName(sourceFormulaId)),
                "New OLE formula did not receive identity bookmark before first read.");
            var typingBoundaryName = "VTBL_" + Guid.Parse(sourceFormulaId).ToString("N");
            AssertTrue(bookmarks.Exists(typingBoundaryName),
                "Direct inline OLE insertion did not create its VTBL typing boundary.");
            typingBoundary = bookmarks[typingBoundaryName];
            typingBoundaryRange = typingBoundary.Range;
            AssertEqual("\u200C", typingBoundaryRange.Text,
                "Direct inline OLE insertion did not use the zero-width U+200C typing boundary.");
            followingProse = document.Range(
                typingBoundaryRange.End,
                typingBoundaryRange.End + 1);
            AssertEqual("后", followingProse.Text,
                "Direct inline OLE insertion added visible spacing before following prose.");
            range = shape.Range.Duplicate;
            range.Copy();
            var paragraph = document.Paragraphs[2];
            try
            {
                var pasteAt = paragraph.Range.Start + "copy: ".Length;
                application.Selection.SetRange(pasteAt, pasteAt);
                application.Selection.Paste();
            }
            finally { Release(paragraph); }
            AssertEqual(2, document.InlineShapes.Count, "Word did not paste OLE copy.");

            Release(shape);
            shape = document.InlineShapes[2];
            Release(range);
            range = shape.Range.Duplicate;
            range.Select();
            var copied = service.ReadSelection();
            AssertTrue(copied.Metadata is not null && !string.IsNullOrWhiteSpace(copied.FormulaId),
                "Pasted OLE copy was not recognized as editable.");
            AssertTrue(!string.Equals(sourceFormulaId, copied.FormulaId, StringComparison.OrdinalIgnoreCase),
                "Pasted OLE copy reused source FormulaId.");
            var copyFormulaId = copied.FormulaId!;
            var persistedCopy = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidOperationException("Re-keyed OLE copy has no persisted metadata.");
            AssertEqual(copyFormulaId, persistedCopy.FormulaId,
                "Re-keyed OLE FormulaId was not persisted to the pasted object.");

            // Simulate Word dropping/moving the identity bookmark. The copied
            // object's own metadata must still preserve its independent identity.
            var identityName = WordFormulaMetadataReader.IdentityBookmarkName(copyFormulaId);
            if (bookmarks.Exists(identityName))
            {
                identity = bookmarks[identityName];
                identity.Delete();
                Release(identity);
                identity = null;
            }
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            document.Activate();
            service = new WordFormulaService(application);

            copied = SelectOleAndRead(document, application, service, 2);
            AssertEqual(copyFormulaId, copied.FormulaId,
                "Copied OLE lost its independent FormulaId after save/reopen without identity bookmark.");
            AssertTrue(WordDoubleClickRouting.ShouldOpenVisualTeX(copied),
                "Reopened OLE copy would not route to VisualTeX.");
            service.ReplaceOle(
                CopyPasteOleSession(document, "edit", copyFormulaId, "x_2+y_2",
                    copied.ObjectId, copied.Metadata, 902),
                png2,
                emf2);
            copied = SelectOleAndRead(document, application, service, 2);
            AssertEqual(copyFormulaId, copied.FormulaId, "OLE copy changed FormulaId after first edit.");
            AssertEqual("x_2+y_2", copied.Metadata?.Lines.FirstOrDefault()?.Latex,
                "First OLE copy edit did not persist source metadata.");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false);
            document.Activate();
            service = new WordFormulaService(application);
            copied = SelectOleAndRead(document, application, service, 2);
            AssertEqual(copyFormulaId, copied.FormulaId,
                "OLE copy became uneditable after first edit and second reopen.");
            service.ReplaceOle(
                CopyPasteOleSession(document, "edit", copyFormulaId, "x_3+y_3",
                    copied.ObjectId, copied.Metadata, 903),
                png3,
                emf3);
            copied = SelectOleAndRead(document, application, service, 2);
            AssertEqual(copyFormulaId, copied.FormulaId, "OLE copy changed FormulaId after second edit.");
            AssertEqual("x_3+y_3", copied.Metadata?.Lines.FirstOrDefault()?.Latex,
                "Second OLE copy edit did not persist source metadata.");

            var source = SelectOleAndRead(document, application, service, 1);
            AssertEqual(sourceFormulaId, source.FormulaId,
                "Editing pasted OLE copy changed source formula identity.");
            document.Save();
            Console.WriteLine(
                $"Word OLE copy/paste re-edit passed: source={sourceFormulaId}, copy={copyFormulaId}, "
                + "identity survived bookmark loss, two reopens and two edits.");
        }
        finally
        {
            Release(followingProse);
            Release(typingBoundaryRange);
            Release(typingBoundary);
            Release(identity);
            Release(bookmarks);
            Release(range);
            Release(shape);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSelection SelectOleAndRead(
        Word.Document document,
        Word.Application application,
        WordFormulaService service,
        int index)
    {
        Word.InlineShape? shape = null;
        Word.Range? range = null;
        try
        {
            shape = document.InlineShapes[index];
            range = shape.Range.Duplicate;
            range.Select();
            return service.ReadSelection();
        }
        finally
        {
            Release(range);
            Release(shape);
        }
    }

    private static OfficeSessionDocument CopyPasteOleSession(
        Word.Document document,
        string mode,
        string formulaId,
        string latex,
        string? sourceObjectId,
        FormulaMetadata? originalMetadata,
        int renderIndex)
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
            Title = "Word copy/paste re-edit acceptance",
            CodeFormat = "latex",
            DisplayMode = "inline",
            ObjectMode = FormulaOleContract.NativeOleMode,
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
}

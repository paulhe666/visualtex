using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class WordFormulaService
{
    private const string RangeReferencePrefix = "visualtex-word-vsto-range:";
    private readonly Application _application;

    public WordFormulaService(Application application)
    {
        _application = application;
    }

    public OfficeSelection ReadSelection() => ReadSelection(null);

    public OfficeSelection ReadSelection(Selection? providedSelection)
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        InlineShapes? inlineShapes = null;
        InlineShape? shape = null;
        Bookmark? ommlBookmark = null;
        var ownsSelection = providedSelection is null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            selection = providedSelection ?? _application.Selection;
            range = selection.Range;
            inlineShapes = range.InlineShapes;
            FormulaMetadata? metadata = null;
            string? objectMode = null;
            if (inlineShapes.Count == 1)
            {
                shape = inlineShapes[1];
                metadata = WordFormulaMetadataReader.TryRead(shape);
                if (metadata is not null)
                    objectMode = WordFormulaMetadataReader.IsNativeOle(shape)
                        ? FormulaOleContract.NativeOleMode
                        : FormulaOleContract.CrossPlatformPictureMode;
            }
            if (metadata is null)
            {
                ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, range);
                if (ommlBookmark is not null)
                {
                    metadata = WordOmmlFormulaStore.TryRead(document, ommlBookmark);
                    if (metadata is not null)
                    {
                        metadata = WordOmmlNativeSource.RefreshForVisualTeX(
                            document,
                            ommlBookmark,
                            metadata);
                        objectMode = FormulaOleContract.WordOmmlMode;
                    }
                }
            }
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = metadata?.FormulaId ?? RangeReference(range),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
                ObjectMode = objectMode,
            };
        }
        finally
        {
            Release(ommlBookmark);
            Release(shape);
            Release(inlineShapes);
            Release(range);
            if (ownsSelection) Release(selection);
            Release(document);
        }
    }

    public bool IsSelectedNativeOle()
    {
        Selection? selection = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OLEFormat? format = null;
        try
        {
            selection = _application.Selection;
            range = selection.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1) return false;
            shape = shapes[1];
            if (shape.Type is not WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                and not WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                return false;
            format = shape.OLEFormat;
            return string.Equals(
                format.ProgID,
                FormulaOleContract.ProgId,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
            Release(shape);
            Release(shapes);
            Release(range);
            Release(selection);
        }
    }

    public string DeleteSelectedFormula()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Document? document = null;
        InlineShape? shape = null;
        Bookmark? ommlBookmark = null;
        Range? ommlRange = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Delete Formula");
            if (string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    requiredFormulaId)
                    ?? throw new InvalidOperationException(
                        "The selected Word OMML formula no longer exists.");
                ommlRange = WordOmmlFormulaStore.GetEquationRange(ommlBookmark);
                ommlBookmark.Delete();
                ommlRange.Delete();
                WordOmmlFormulaStore.Delete(document, requiredFormulaId);
            }
            else
            {
                shape = FindByFormulaId(document, requiredFormulaId)
                    ?? throw new InvalidOperationException(
                        "The selected Word formula no longer exists.");
                shape.Delete();
            }
            WordEquationNumbering.TryReconcile(document);
            return requiredFormulaId;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(ommlRange);
            Release(ommlBookmark);
            Release(shape);
            Release(document);
        }
    }

    public int UpdateEquationNumbers()
    {
        Document? document = null;
        UndoRecord? undoRecord = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Update Equation Numbers");
            return WordEquationNumbering.Reconcile(document);
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(document);
        }
    }

    public string ExportSelectedOleAsPicture()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Document? document = null;
        InlineShape? oldShape = null;
        OLEFormat? format = null;
        object? oleObject = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        string? pngPath = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            undoRecord = BeginUndoRecord("VisualTeX Export OLE Formula As Picture");
            oldShape = FindByFormulaId(document, requiredFormulaId)
                ?? throw new InvalidOperationException("The selected Word formula no longer exists.");
            var metadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? throw new InvalidDataException("The selected formula metadata is invalid.");
            format = oldShape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected formula is already a picture.");
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            pngPath = OlePngPreviewExtractor.MaterializePng(oleObject, requiredFormulaId);

            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                pngPath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                oldWidth,
                oldHeight,
                pngPath,
                (float)(metadata.RenderHeightPx ?? 0),
                metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                metadata.DisplayMode == "inline");
            oldShape.Delete();
            WordEquationNumbering.TryReconcile(document);
            return requiredFormulaId;
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            if (pngPath is not null)
            {
                try { File.Delete(pngPath); } catch { }
            }
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oleObject);
            Release(format);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult Insert(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        UndoRecord? undoRecord = null;
        var typingFontSize = 11f;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Inline Formula"
                    : "VisualTeX Insert Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            typingFontSize = CaptureTypingFontSize(selection);
            insertion = ResolveSourceRange(document, session.SourceObjectId, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            object link = false;
            object save = true;
            object rangeObject;
            if (session.DisplayMode == "inline")
            {
                rangeObject = insertion;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            else
            {
                insertion.InsertParagraphAfter();
                insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                paragraph = document.Paragraphs.Add(insertion);
                paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                paragraphRange = paragraph.Range;
                rangeObject = paragraphRange;
                shape = document.InlineShapes.AddPicture(
                    imagePath,
                    ref link,
                    ref save,
                    ref rangeObject);
            }
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                WordEquationNumbering.TryReconcile(document);
                Range? displayRange = null;
                try
                {
                    displayRange = shape.Range;
                    WordEquationNumbering.RestoreTypingAfterDisplayFormula(
                        document,
                        displayRange,
                        metadata.FormulaId,
                        selection,
                        typingFontSize);
                }
                finally { Release(displayRange); }
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult InsertOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShape? shape = null;
        Table? numberedTable = null;
        UndoRecord? undoRecord = null;
        var typingFontSize = 11f;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Native OLE Inline Formula"
                    : "VisualTeX Insert Native OLE Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            typingFontSize = CaptureTypingFontSize(selection);
            insertion = ResolveSourceRange(document, session.SourceObjectId, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (session.DisplayMode == "inline")
            {
                shape = AddOleObject(document, insertion);
            }
            else
            {
                if (session.Numbered)
                {
                    var tableInsertion = CreateNumberedDisplayTable(
                        document,
                        insertion,
                        out numberedTable);
                    Release(insertion);
                    insertion = tableInsertion;
                    shape = AddOleObject(document, insertion);
                }
                else
                {
                    insertion.InsertParagraphAfter();
                    insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                    paragraph = document.Paragraphs.Add(insertion);
                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraphRange = paragraph.Range;
                    shape = AddOleObject(document, paragraphRange);
                }
            }
            InitializeOle(shape, metadata, emfPath, pngPath);
            Configure(
                shape,
                metadata,
                (session.ExportResult?.Width ?? 200) * 0.75f,
                (session.ExportResult?.Height ?? 60) * 0.75f,
                pngPath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(shape);
            }
            else
            {
                WordEquationNumbering.TryReconcile(document);
                Range? displayRange = null;
                try
                {
                    displayRange = shape.Range;
                    WordEquationNumbering.RestoreTypingAfterDisplayFormula(
                        document,
                        displayRange,
                        metadata.FormulaId,
                        selection,
                        typingFontSize);
                }
                finally { Release(displayRange); }
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(shape);
            if (numberedTable is not null) TryDelete(numberedTable);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(shape);
            Release(numberedTable);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult InsertOmml(
        OfficeSessionDocument session,
        string mathMl)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        Selection? selection = null;
        Range? insertion = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? equationRange = null;
        Bookmark? bookmark = null;
        Table? numberedTable = null;
        UndoRecord? undoRecord = null;
        var metadataSaved = false;
        var typingFontSize = 11f;
        try
        {
            undoRecord = BeginUndoRecord(
                session.DisplayMode == "inline"
                    ? "VisualTeX Insert Word OMML Inline Formula"
                    : "VisualTeX Insert Word OMML Display Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            selection = _application.Selection;
            typingFontSize = CaptureTypingFontSize(selection);
            insertion = ResolveSourceRange(document, session.SourceObjectId, selection);
            insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (session.DisplayMode == "inline")
            {
                equationRange = WordOmmlConverter.Insert(
                    _application,
                    document,
                    insertion,
                    mathMl,
                    display: false);
            }
            else
            {
                if (session.Numbered)
                {
                    var tableInsertion = CreateNumberedDisplayTable(
                        document,
                        insertion,
                        out numberedTable);
                    Release(insertion);
                    insertion = tableInsertion;
                }
                else
                {
                    insertion.InsertParagraphAfter();
                    insertion.Collapse(WdCollapseDirection.wdCollapseEnd);
                    paragraph = document.Paragraphs.Add(insertion);
                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    paragraphRange = paragraph.Range;
                    insertion.SetRange(paragraphRange.Start, paragraphRange.Start);
                }
                equationRange = WordOmmlConverter.Insert(
                    _application,
                    document,
                    insertion,
                    mathMl,
                    display: true);
            }

            WordOmmlNativeSource.StampFingerprint(metadata, equationRange);
            bookmark = WordOmmlFormulaStore.Wrap(document, equationRange, metadata);
            WordOmmlFormulaStore.Save(document, metadata);
            metadataSaved = true;
            if (session.DisplayMode == "inline")
            {
                RestoreTypingBaselineAfter(bookmark);
            }
            else
            {
                WordEquationNumbering.TryReconcile(document);
                Range? displayRange = null;
                try
                {
                    displayRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    WordEquationNumbering.RestoreTypingAfterDisplayFormula(
                        document,
                        displayRange,
                        metadata.FormulaId,
                        selection,
                        typingFontSize);
                }
                finally { Release(displayRange); }
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(bookmark, deleteContents: true);
            if (bookmark is null) TryDelete(equationRange);
            if (numberedTable is not null) TryDelete(numberedTable);
            if (metadataSaved && document is not null)
            {
                try { WordOmmlFormulaStore.Delete(document, metadata.FormulaId); } catch { }
            }
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(bookmark);
            Release(equationRange);
            Release(numberedTable);
            Release(paragraphRange);
            Release(paragraph);
            Release(insertion);
            Release(selection);
            Release(document);
        }
    }

    public OfficeObjectResult ReplaceOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Bookmark? oldBookmark = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        Range? rollbackEquationRange = null;
        Bookmark? rollbackBookmark = null;
        UndoRecord? undoRecord = null;
        FormulaMetadata? originalMetadata = null;
        var oldStart = -1;
        var removedOmml = false;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert or Update Native OLE Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            oldShape = FindByFormulaId(document, session.FormulaId);
            float oldWidth;
            float oldHeight;
            if (oldShape is not null)
            {
                oldWidth = oldShape.Width;
                oldHeight = oldShape.Height;
                originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                    ?? session.OriginalMetadata;
            }
            else
            {
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(document, session.FormulaId)
                    ?? throw new InvalidOperationException(
                        "The target Word formula no longer exists.");
                originalMetadata = WordOmmlFormulaStore.TryRead(document, oldBookmark)
                    ?? session.OriginalMetadata;
                oldWidth = (float)Math.Max(
                    12,
                    (originalMetadata?.RenderWidthPx ?? session.ExportResult?.Width ?? 200) * 0.75);
                oldHeight = (float)Math.Max(
                    12,
                    (originalMetadata?.RenderHeightPx ?? session.ExportResult?.Height ?? 60) * 0.75);
            }
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f);

            if (oldShape is not null && TryUpdateOle(oldShape, metadata, emfPath, pngPath))
            {
                Configure(
                    oldShape,
                    metadata,
                    editedSize.Width,
                    editedSize.Height,
                    pngPath,
                    session.ExportResult?.Height ?? 0,
                    session.ExportResult?.Baseline,
                    session.DisplayMode == "inline");
                if (session.DisplayMode == "inline")
                    RestoreTypingBaselineAfter(oldShape);
                else
                    WordEquationNumbering.TryReconcile(document);
                return Result(session, document);
            }

            oldRange = oldShape is not null
                ? oldShape.Range
                : WordOmmlFormulaStore.GetEquationRange(oldBookmark!);
            oldStart = oldRange.Start;
            if (oldShape is not null)
            {
                insertion = oldRange.Duplicate;
                insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                // Remove the native equation before creating the OLE object.
                // Inserting at OMath.End first allows Word to expand the live
                // math container around the OLE object, leaving the large OMML
                // selection frame and shifting the replacement horizontally.
                oldBookmark!.Delete();
                oldRange.Delete();
                WordOmmlFormulaStore.Delete(document, session.FormulaId);
                removedOmml = true;
                object insertionStart = oldStart;
                object insertionEnd = oldStart;
                insertion = document.Range(ref insertionStart, ref insertionEnd);
            }
            replacement = AddOleObject(document, insertion);
            InitializeOle(replacement, metadata, emfPath, pngPath);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                pngPath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            if (oldShape is not null)
                oldShape.Delete();
            if (session.DisplayMode == "block" && session.Numbered)
                NormalizeNumberedDisplayCell(replacement);
            if (session.DisplayMode == "inline")
                RestoreTypingBaselineAfter(replacement);
            else
                WordEquationNumbering.TryReconcile(document);
            if (oldShape is null && session.DisplayMode != "inline")
            {
                try { replacement.Select(); } catch { }
            }
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            if (removedOmml
                && document is not null
                && originalMetadata is not null
                && oldStart >= 0
                && !string.IsNullOrWhiteSpace(session.ExportResult?.MathMl))
            {
                try
                {
                    object restoreStart = oldStart;
                    object restoreEnd = oldStart;
                    var restoreInsertion = document.Range(ref restoreStart, ref restoreEnd);
                    try
                    {
                        rollbackEquationRange = WordOmmlConverter.Insert(
                            _application,
                            document,
                            restoreInsertion,
                            session.ExportResult!.MathMl!,
                            session.DisplayMode == "block",
                            includeLeadingTab: false);
                    }
                    finally { Release(restoreInsertion); }
                    WordOmmlNativeSource.StampFingerprint(
                        originalMetadata,
                        rollbackEquationRange);
                    rollbackBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        rollbackEquationRange,
                        originalMetadata);
                    WordOmmlFormulaStore.Save(document, originalMetadata);
                    WordEquationNumbering.TryReconcile(document);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(rollbackBookmark);
            Release(rollbackEquationRange);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldBookmark);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult ReplaceOmml(
        OfficeSessionDocument session,
        string mathMl)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Bookmark? oldBookmark = null;
        Range? oldRange = null;
        Range? insertion = null;
        Range? equationRange = null;
        Bookmark? replacement = null;
        Table? numberedTable = null;
        Paragraph? replacementParagraph = null;
        Range? replacementParagraphRange = null;
        UndoRecord? undoRecord = null;
        FormulaMetadata? originalOmmlMetadata = null;
        var metadataSaved = false;
        var oldBookmarkRemoved = false;
        var oldNumberingArtifactsRemoved = false;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Convert or Update Word OMML Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            oldShape = FindByFormulaId(document, session.FormulaId);
            if (oldShape is not null)
            {
                // Remove old equation-number bookmarks before inserting the
                // adjacent replacement. Word expands a trailing bookmark when
                // content is inserted at its edge; deleting that old bookmark
                // during reconciliation can otherwise delete the new table.
                WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                    document,
                    session.FormulaId);
                oldNumberingArtifactsRemoved = true;
                oldRange = oldShape.Range;
                insertion = oldRange.Duplicate;
                // Insert immediately before the source OLE in its existing
                // paragraph/cell. Creating another paragraph or numbered table
                // here nests layout containers during OLE -> OMML -> OLE
                // round-trips and leaves visible empty paragraph marks behind.
                insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            }
            else
            {
                oldBookmark = WordOmmlFormulaStore.FindByFormulaId(document, session.FormulaId)
                    ?? throw new InvalidOperationException(
                        "The target Word formula no longer exists.");
                originalOmmlMetadata = WordOmmlFormulaStore.TryRead(document, oldBookmark);
                oldRange = WordOmmlFormulaStore.GetEquationRange(oldBookmark);
                insertion = oldRange.Duplicate;
            }

            equationRange = WordOmmlConverter.Insert(
                _application,
                document,
                insertion,
                mathMl,
                session.DisplayMode == "block",
                replaceTarget: oldShape is null);
            ValidateInsertedOmml(equationRange);
            if (oldBookmark is not null)
            {
                oldBookmark.Delete();
                oldBookmarkRemoved = true;
            }
            WordOmmlNativeSource.StampFingerprint(metadata, equationRange);
            replacement = WordOmmlFormulaStore.Wrap(document, equationRange, metadata);
            WordOmmlFormulaStore.Save(document, metadata);
            metadataSaved = true;

            // Keep the source OLE until replacement and metadata are valid.
            if (oldShape is not null)
                oldShape.Delete();
            if (session.DisplayMode == "block" && session.Numbered)
            {
                // Turning an OMath into display form while its source OLE is
                // still present makes Word insert manual line-break runs on
                // both sides. Once the OLE is deleted those hidden breaks stay
                // in the formula cell, so the cell is centered but the formula
                // is not. Remove everything outside the replacement equation,
                // then recreate its collapsed anchor at the normalized edge.
                NormalizeNumberedDisplayCell(equationRange);
                Release(replacement);
                replacement = WordOmmlFormulaStore.Wrap(
                    document,
                    equationRange,
                    metadata);
            }
            // Deleting the source object can move the selection and restore its
            // raised inline font position. Normalize the final caret only after
            // the old object is gone.
            if (session.DisplayMode == "inline")
                RestoreTypingBaselineAfter(replacement);
            if (session.DisplayMode == "block")
                WordEquationNumbering.TryReconcile(document);
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement, deleteContents: true);
            if (replacement is null) TryDelete(equationRange);
            if (numberedTable is not null) TryDelete(numberedTable);
            else if (oldShape is not null) TryDelete(replacementParagraphRange);
            if (document is not null)
            {
                try
                {
                    if (oldBookmarkRemoved
                        && oldRange is not null
                        && originalOmmlMetadata is not null)
                    {
                        var restored = WordOmmlFormulaStore.Wrap(
                            document,
                            oldRange,
                            originalOmmlMetadata);
                        Release(restored);
                        WordOmmlFormulaStore.Save(document, originalOmmlMetadata);
                    }
                    else if (metadataSaved)
                    {
                        WordOmmlFormulaStore.Delete(document, metadata.FormulaId);
                    }
                    if (oldNumberingArtifactsRemoved && oldShape is not null)
                        WordEquationNumbering.TryReconcile(document);
                }
                catch { }
            }
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(replacement);
            Release(equationRange);
            Release(replacementParagraphRange);
            Release(replacementParagraph);
            Release(numberedTable);
            Release(insertion);
            Release(oldRange);
            Release(oldBookmark);
            Release(oldShape);
            Release(document);
        }
    }

    public OfficeObjectResult Replace(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Document? document = null;
        InlineShape? oldShape = null;
        Range? oldRange = null;
        Range? insertion = null;
        InlineShape? replacement = null;
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = BeginUndoRecord("VisualTeX Replace Formula");
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, session.SourceDocumentId);
            oldShape = FindByFormulaId(document, session.FormulaId)
                ?? throw new InvalidOperationException("The target Word formula no longer exists.");
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var originalMetadata = WordFormulaMetadataReader.TryRead(oldShape)
                ?? session.OriginalMetadata;
            var editedSize = OfficeFormulaSizing.EditedSize(
                oldWidth,
                oldHeight,
                originalMetadata?.RenderWidthPx,
                originalMetadata?.RenderHeightPx,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f);
            oldRange = oldShape.Range;
            insertion = oldRange.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            object link = false;
            object save = true;
            object rangeObject = insertion;
            replacement = document.InlineShapes.AddPicture(
                imagePath,
                ref link,
                ref save,
                ref rangeObject);
            Configure(
                replacement,
                metadata,
                editedSize.Width,
                editedSize.Height,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            oldShape.Delete();
            if (session.DisplayMode == "inline")
                RestoreTypingBaselineAfter(replacement);
            else
                WordEquationNumbering.TryReconcile(document);
            return Result(session, document);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(replacement);
            Release(insertion);
            Release(oldRange);
            Release(oldShape);
            Release(document);
        }
    }

    private static InlineShape AddOleObject(Document document, Range range) =>
        document.InlineShapes.AddOLEObject(
            ClassType: FormulaOleContract.ProgId,
            LinkToFile: false,
            DisplayAsIcon: false,
            Range: range);

    private static void InitializeOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            if (oleObject is not IVisualTeXFormulaObject formula)
                throw new InvalidOperationException(
                    "The inserted Word object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        InlineShape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            try { format = shape.OLEFormat; }
            catch { return false; }
            try { oleObject = WordOleObjectAccessor.GetRunningObject(format); }
            catch { return false; }
            if (oleObject is not IVisualTeXFormulaObject formula) return false;
            FormulaOleInterop.Update(formula, metadata, emfPath, pngPath);
            return true;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static InlineShape? FindByFormulaId(Document document, string formulaId)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = shapes[index];
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata?.FormulaId == formulaId)
                    {
                        var result = shape;
                        shape = null;
                        return result;
                    }
                }
                finally { Release(shape); }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private UndoRecord? BeginUndoRecord(string name)
    {
        UndoRecord? undoRecord = null;
        try
        {
            undoRecord = _application.UndoRecord;
            undoRecord.StartCustomRecord(name);
            return undoRecord;
        }
        catch
        {
            Release(undoRecord);
            return null;
        }
    }

    private static void EndUndoRecord(UndoRecord? undoRecord)
    {
        if (undoRecord is null) return;
        try { undoRecord.EndCustomRecord(); } catch { }
    }

    private static void Configure(
        InlineShape shape,
        FormulaMetadata metadata,
        float maxWidth,
        float maxHeight,
        string imagePath,
        float exportedHeight,
        float? exportedBaseline,
        bool alignInline)
    {
        using var image = Image.FromFile(imagePath);
        var ratio = image.Width / (float)Math.Max(1, image.Height);
        var width = Math.Max(12f, maxWidth);
        var height = width / ratio;
        if (maxHeight > 0 && height > maxHeight)
        {
            height = maxHeight;
            width = height * ratio;
        }
        // An OLE object is initially created with the placeholder preview's 4:1
        // aspect ratio. Setting only Width while aspect-ratio locking is enabled
        // therefore distorts the real formula. Apply both natural dimensions
        // explicitly, then lock the resolved ratio for later user resizing.
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
        shape.Width = width;
        shape.Height = height;
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        if (!WordFormulaMetadataReader.IsNativeOle(shape))
        {
            var encoded = FormulaMetadataCodec.Encode(metadata);
            shape.Title = encoded;
            shape.AlternativeText = encoded;
        }
        if (alignInline)
            ApplyInlineBaseline(shape, shape.Height, exportedHeight, exportedBaseline);
    }

    private static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        Range? range = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            font = range.Font;
            font.Position = WordInlineAlignment.CalculateFontPosition(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline);
        }
        finally
        {
            Release(font);
            Release(range);
        }
    }

    private static float CaptureTypingFontSize(Selection selection)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = selection.Font;
            var size = font.Size;
            return WordEquationNumbering.IsNormalTextSize(size) ? size : 11f;
        }
        catch
        {
            return 11f;
        }
        finally { Release(font); }
    }

    private void RestoreTypingBaselineAfter(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            RestoreTypingBaselineAfter(range);
        }
        finally { Release(range); }
    }

    private void RestoreTypingBaselineAfter(Bookmark bookmark)
    {
        Range? range = null;
        try
        {
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            RestoreTypingBaselineAfter(range);
        }
        finally { Release(range); }
    }

    private void RestoreTypingBaselineAfter(Range formulaRange)
    {
        Range? caret = null;
        Selection? selection = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            caret = formulaRange.Duplicate;
            caret.Collapse(WdCollapseDirection.wdCollapseEnd);
            font = caret.Font;
            font.Position = 0;
            Release(font);
            font = null;

            selection = _application.Selection;
            selection.SetRange(caret.Start, caret.End);
            font = selection.Font;
            font.Position = 0;
        }
        finally
        {
            Release(font);
            Release(selection);
            Release(caret);
        }
    }

    private static OfficeObjectResult Result(OfficeSessionDocument session, Document document) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = DocumentIdentity(document),
            ObjectId = session.FormulaId,
        };

    private static string RangeReference(Range range) =>
        $"{RangeReferencePrefix}{range.Start}:{range.End}";

    private static Range ResolveSourceRange(
        Document document,
        string? sourceObjectId,
        Selection selection)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return selection.Range.Duplicate;
        Range? content = null;
        try
        {
            content = document.Content;
            if (start < 0 || end < start || end > content.End)
                throw new InvalidOperationException(
                    "The Word insertion range selected when the formula editor opened is no longer valid.");
            object startValue = start;
            object endValue = end;
            return document.Range(ref startValue, ref endValue);
        }
        finally { Release(content); }
    }

    private static bool TryParseRangeReference(
        string? value,
        out int start,
        out int end)
    {
        start = 0;
        end = 0;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var reference = value!;
        if (!reference.StartsWith(RangeReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = reference.Substring(RangeReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator <= 0 || separator >= payload.Length - 1) return false;
        return int.TryParse(payload.Substring(0, separator), out start)
            && int.TryParse(payload.Substring(separator + 1), out end);
    }

    private static void EnsureSourceDocument(
        Document document,
        string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(document);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active Word document changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(Document document)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(document.FullName)) return document.FullName;
        }
        catch { }
        return document.Name;
    }

    private static void EnsureWritable(Document document)
    {
        if (document.ReadOnly)
            throw new UnauthorizedAccessException("The active Word document is read-only.");
    }

    private static bool HasLeadingTab(Document document, Range formulaRange)
    {
        if (formulaRange.Start <= 0) return false;
        Range? preceding = null;
        try
        {
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return true;
            if (!string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                || formulaRange.Start <= 1)
                return false;
            preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
            return string.Equals(preceding.Text, "\t", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally { Release(preceding); }
    }

    private static Range CreateNumberedDisplayTable(
        Document document,
        Range anchor,
        out Table table)
    {
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Borders? borders = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        try
        {
            anchor.InsertParagraphAfter();
            anchor.Collapse(WdCollapseDirection.wdCollapseEnd);
            table = document.Tables.Add(anchor, 1, 3);
            table.AllowAutoFit = false;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            borders = table.Borders;
            borders.Enable = 0;
            columns = table.Columns;
            leftColumn = columns[1];
            centerColumn = columns[2];
            rightColumn = columns[3];
            leftColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            centerColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            rightColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            leftColumn.PreferredWidth = 20f;
            centerColumn.PreferredWidth = 60f;
            rightColumn.PreferredWidth = 20f;

            centerCell = table.Cell(1, 2);
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            centerCellRange = centerCell.Range;
            var insertion = centerCellRange.Duplicate;
            insertion.End = Math.Max(insertion.Start, insertion.End - 1);
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            return insertion;
        }
        finally
        {
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(borders);
            Release(centerCellRange);
            Release(centerCell);
        }
    }

    private static void NormalizeNumberedDisplayCell(Range formulaRange)
    {
        Document? document = null;
        Table? table = null;
        Columns? columns = null;
        Cell? centerCell = null;
        Range? cellRange = null;
        Range? character = null;
        try
        {
            if (!(bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                || formulaRange.Tables.Count == 0)
                return;
            document = formulaRange.Document;
            table = formulaRange.Tables[1];
            columns = table.Columns;
            if (columns.Count < 3) return;
            centerCell = table.Cell(1, 2);
            cellRange = centerCell.Range;

            // A display OMath inserted next to the source OLE can leave one
            // manual line break on each side. Delete only those exact control
            // characters, scanning backwards so Word's shifting ranges cannot
            // expand across and remove the replacement formula object.
            for (var position = cellRange.End - 2;
                 position >= cellRange.Start;
                 position--)
            {
                if (position >= formulaRange.Start
                    && position < formulaRange.End)
                    continue;
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(
                    ref characterStart,
                    ref characterEnd);
                if (string.Equals(character.Text, "\v", StringComparison.Ordinal))
                    character.Delete();
                Release(character);
                character = null;
            }
        }
        finally
        {
            Release(character);
            Release(cellRange);
            Release(centerCell);
            Release(columns);
            Release(table);
            Release(document);
        }
    }

    private static void NormalizeNumberedDisplayCell(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            NormalizeNumberedDisplayCell(range);
        }
        finally { Release(range); }
    }

    private static void ValidateInsertedOmml(Range equationRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        try
        {
            maths = equationRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create exactly one native OMML equation.");
            math = maths[1];
            mathRange = math.Range;
            var wordOpenXml = mathRange.WordOpenXML;
            if (mathRange.End <= mathRange.Start
                || string.IsNullOrWhiteSpace(wordOpenXml)
                || wordOpenXml.IndexOf("oMath", StringComparison.Ordinal) < 0)
                throw new InvalidOperationException(
                    "Word returned an empty native OMML equation.");
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
        }
    }

    private static void TryDelete(InlineShape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
    }

    private static void TryDelete(Table? table)
    {
        if (table is null) return;
        try { table.Delete(); } catch { }
    }

    private static void TryDelete(Bookmark? bookmark, bool deleteContents)
    {
        if (bookmark is null) return;
        Range? range = null;
        try
        {
            if (deleteContents) range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            bookmark.Delete();
            if (deleteContents) range?.Delete();
        }
        catch { }
        finally { Release(range); }
    }

    private static void TryDelete(Range? range)
    {
        if (range is null) return;
        try { range.Delete(); } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        // Office may return the same RCW to the host and to this service.
        // FinalReleaseComObject would invalidate every shared reference in the
        // add-in AppDomain, so release only the reference acquired here.
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

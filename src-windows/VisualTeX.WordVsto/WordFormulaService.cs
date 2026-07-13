using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

internal sealed class WordFormulaService
{
    private const string RangeReferencePrefix = "visualtex-word-vsto-range:";
    private readonly Application _application;

    public WordFormulaService(Application application)
    {
        _application = application;
    }

    public OfficeSelection ReadSelection()
    {
        Document? document = null;
        Selection? selection = null;
        Range? range = null;
        InlineShapes? inlineShapes = null;
        InlineShape? shape = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            selection = _application.Selection;
            range = selection.Range;
            inlineShapes = range.InlineShapes;
            FormulaMetadata? metadata = null;
            if (inlineShapes.Count == 1)
            {
                shape = inlineShapes[1];
                metadata = FormulaMetadataCodec.Decode(shape.AlternativeText)
                    ?? FormulaMetadataCodec.Decode(shape.Title);
            }
            return new OfficeSelection
            {
                Host = "word",
                DocumentId = DocumentIdentity(document),
                ObjectId = metadata?.FormulaId ?? RangeReference(range),
                ReadOnly = document.ReadOnly,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
            };
        }
        finally
        {
            Release(shape);
            Release(inlineShapes);
            Release(range);
            Release(selection);
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
                oldWidth,
                oldHeight,
                imagePath,
                session.ExportResult?.Height ?? 0,
                session.ExportResult?.Baseline,
                session.DisplayMode == "inline");
            oldShape.Delete();
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
                    var metadata = FormulaMetadataCodec.Decode(shape.AlternativeText)
                        ?? FormulaMetadataCodec.Decode(shape.Title);
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
        var encoded = FormulaMetadataCodec.Encode(metadata);
        using var image = Image.FromFile(imagePath);
        var ratio = image.Width / (float)Math.Max(1, image.Height);
        var width = Math.Max(12f, maxWidth);
        var height = width / ratio;
        if (maxHeight > 0 && height > maxHeight)
        {
            height = maxHeight;
            width = height * ratio;
        }
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        shape.Width = width;
        shape.Title = encoded;
        shape.AlternativeText = encoded;
        if (alignInline)
            ApplyInlineBaseline(shape, height, exportedHeight, exportedBaseline);
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
        if (string.IsNullOrWhiteSpace(value)
            || !value.StartsWith(RangeReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = value.Substring(RangeReferencePrefix.Length);
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

    private static void TryDelete(InlineShape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }
}

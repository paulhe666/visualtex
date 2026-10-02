using System.Runtime.InteropServices;
using System.Globalization;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Local ownership for an existing OLE edit. Word can collapse every range that
/// ended at the replaced paragraph mark. Only the end of an untouched following
/// character is an outside boundary; at story end reacquire the live story end.
/// No document-wide shape, paragraph, bookmark or revision enumeration is used.
/// </summary>
internal sealed class MathTypeLocalEditScope : IDisposable
{
    private readonly Document _document;
    private Range? _storyAnchor;
    private Range? _followingCharacter;
    private readonly List<BookmarkSpan> _bookmarks = new();
    private int _paragraphStart;
    private int _evidenceStart;
    private int _paragraphEnd;
    private int _precedingInsertionLength;
    private int _paragraphCount;
    private int _sourceLength;
    private string _paragraphText = string.Empty;
    private string _rightWhitespace = string.Empty;
    private string _originalSignature = string.Empty;
    private bool _insertionCapture;
    private int _insertionOffset;
    private string _insertionSerializedSignature = string.Empty;
    private string _insertionBoundaryGeometry = string.Empty;
    private bool _replacementBoundaryCaptured;
    private int _replacementParagraphCount;
    private int _replacementOleCount;
    private string _replacementLeftText = string.Empty;
    private string _replacementRightText = string.Empty;

    private sealed class BookmarkSpan
    {
        internal string Name = string.Empty;
        internal int Start;
        internal int End;
    }

    internal string ParagraphXml { get; private set; } = string.Empty;
    internal int SourceOleOrdinal { get; private set; }
    internal int SourceOleCount { get; private set; }
    internal string LocatorName { get; } = "VTMT_TX_" + Guid.NewGuid().ToString("N");

    private MathTypeLocalEditScope(Document document) { _document = document; }

    internal static MathTypeLocalEditScope Capture(Document document, Range source)
    {
        var result = new MathTypeLocalEditScope(document);
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        Range? evidence = null;
        Range? story = null;
        Range? cellRange = null;
        Tables? tables = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? right = null;
        Bookmarks? bookmarks = null;
        try
        {
            result._storyAnchor = source.Duplicate;
            result._sourceLength = source.End - source.Start;
            paragraphs = source.Paragraphs;
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            if (source.Start < owner.Start || source.End > owner.End)
                throw new InvalidOperationException("The native MathType object crosses its owning paragraph.");
            result._paragraphStart = owner.Start;
            result._paragraphEnd = owner.End;
            result._paragraphText = owner.Text ?? string.Empty;
            Paragraphs? ownerParagraphs = null;
            try { ownerParagraphs = owner.Paragraphs; result._paragraphCount = ownerParagraphs.Count; }
            finally { Release(ownerParagraphs); }
            story = source.Duplicate;
            story.Expand(WdUnits.wdStory);
            evidence = owner.Duplicate;
            // Export from an outside left boundary. Starting exactly at Word's
            // logical paragraph start can omit a deleted predecessor paragraph
            // even while InlineShapes still contains that predecessor's OLE.
            // One preceding paragraph-boundary character preserves that history;
            // this context is evidence only and is never imported as content.
            var contextStart = story.Start;
            // A cell/row boundary is not a paragraph-join history boundary.
            // Exporting one character into the preceding cell can serialize its
            // whole row (and OLEs outside the live range). Keep the read witness
            // within the source's actual cell, not just within the story.
            tables = source.Tables;
            if (tables.Count > 0)
            {
                cells = source.Cells;
                if (cells.Count != 1)
                    throw new InvalidDataException("The native OLE has more than one local table-cell owner.");
                cell = cells[1];
                cellRange = cell.Range;
                contextStart = Math.Max(contextStart, cellRange.Start);
            }
            if (owner.Start > contextStart) evidence.Start = owner.Start - 1;
            result._evidenceStart = evidence.Start;
            var tracePerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
                "1",
                StringComparison.Ordinal);
            var evidenceWatch = tracePerformance ? System.Diagnostics.Stopwatch.StartNew() : null;
            result.ParagraphXml = evidence.WordOpenXML;
            if (evidenceWatch is not null)
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-local-scope-perf stage=word-open-xml ms={evidenceWatch.ElapsedMilliseconds} chars={result.ParagraphXml.Length}");
            evidenceWatch?.Restart();
            var geometry = WordInlineObjectGeometry.Capture(evidence);
            if (evidenceWatch is not null)
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-local-scope-perf stage=geometry ms={evidenceWatch.ElapsedMilliseconds} objects={geometry.Count}");
            evidenceWatch?.Restart();
            result._originalSignature = RecoverySignature(evidence, result.ParagraphXml, geometry);
            if (evidenceWatch is not null)
                WordDoubleClickHook.TraceMessage(
                    $"mathtype-local-scope-perf stage=signature ms={evidenceWatch.ElapsedMilliseconds}");
            var localOle = geometry.Where(item => item.Type is 1 or 2).ToArray();
            result.SourceOleCount = localOle.Length;
            for (var index = 0; index < localOle.Length; index++)
            {
                if (localOle[index].Start != source.Start - evidence.Start
                    || localOle[index].End != source.End - evidence.Start) continue;
                if (result.SourceOleOrdinal != 0)
                    throw new InvalidDataException("The source OLE has more than one physical local owner.");
                result.SourceOleOrdinal = index + 1;
            }

            if (owner.End < story.End)
            {
                result._followingCharacter = source.Duplicate;
                result._followingCharacter.SetRange(owner.End, owner.End + 1);
            }
            right = source.Duplicate;
            right.SetRange(source.End, Math.Min(owner.End, source.End + 64));
            result._rightWhitespace = LeadingWhitespace(right.Text ?? string.Empty);

            // WordOpenXML on the exact object omits bookmarks on its outer
            // boundary. Discover names from the already captured LOCAL paragraph,
            // including hidden names, then bind only spans wholly inside the OLE.
            bookmarks = document.Bookmarks;
            foreach (var name in MathTypeLocalFragment.BookmarkNames(result.ParagraphXml))
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    if (!bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name];
                    range = bookmark.Range;
                    if (!range.InStory(source) || range.Start < source.Start || range.End > source.End)
                        continue;
                    result._bookmarks.Add(new BookmarkSpan
                    {
                        Name = name,
                        Start = range.Start - source.Start,
                        End = range.End - source.Start,
                    });
                }
                finally { Release(range); Release(bookmark); }
            }
            return result;
        }
        catch { result.Dispose(); throw; }
        finally
        {
            Release(bookmarks); Release(right); Release(cellRange); Release(cell); Release(cells); Release(tables); Release(story);
            Release(evidence); Release(owner); Release(paragraph); Release(paragraphs);
        }
    }

    // New insertion ownership is intentionally narrower than existing-OLE
    // edit ownership. The insertion does not own every formula already present in
    // its paragraph, so scanning live geometry for all of them on every create is
    // both semantically excessive and O(N). Preserve the complete serialized
    // paragraph/provenance evidence, and verify exact live geometry only in a
    // fixed boundary window around the insertion point.
    internal static MathTypeLocalEditScope CaptureInsertion(Document document, Range source)
    {
        if (source.Start != source.End)
            throw new InvalidOperationException("MathType insertion ownership requires a collapsed range.");

        var result = new MathTypeLocalEditScope(document);
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        Range? evidence = null;
        Range? story = null;
        Range? cellRange = null;
        Tables? tables = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? right = null;
        try
        {
            result._storyAnchor = source.Duplicate;
            paragraphs = source.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException("The MathType insertion point has no unique paragraph owner.");
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            if (source.Start < owner.Start || source.Start > owner.End)
                throw new InvalidOperationException("The MathType insertion point escaped its paragraph owner.");

            result._paragraphStart = owner.Start;
            result._paragraphEnd = owner.End;
            result._insertionOffset = source.Start - owner.Start;
            result._paragraphText = owner.Text ?? string.Empty;
            Paragraphs? ownerParagraphs = null;
            try { ownerParagraphs = owner.Paragraphs; result._paragraphCount = ownerParagraphs.Count; }
            finally { Release(ownerParagraphs); }

            story = source.Duplicate;
            story.Expand(WdUnits.wdStory);
            evidence = owner.Duplicate;
            var contextStart = story.Start;
            tables = source.Tables;
            if (tables.Count > 0)
            {
                cells = source.Cells;
                if (cells.Count != 1)
                    throw new InvalidDataException("The MathType insertion point has more than one table-cell owner.");
                cell = cells[1];
                cellRange = cell.Range;
                contextStart = Math.Max(contextStart, cellRange.Start);
            }
            if (owner.Start > contextStart) evidence.Start = owner.Start - 1;
            result._evidenceStart = evidence.Start;
            result.ParagraphXml = evidence.WordOpenXML;
            result._insertionSerializedSignature = WordLocalEditSnapshot.Signature(
                MathTypeLocalFragment.InsertionRecoveryEvidence(result.ParagraphXml));
            result._insertionBoundaryGeometry = CaptureInsertionBoundaryGeometry(
                document, owner, source.Start);

            if (owner.End < story.End)
            {
                result._followingCharacter = source.Duplicate;
                result._followingCharacter.SetRange(owner.End, owner.End + 1);
            }
            right = source.Duplicate;
            right.SetRange(source.End, Math.Min(owner.End, source.End + 64));
            result._rightWhitespace = LeadingWhitespace(right.Text ?? string.Empty);
            result._insertionCapture = true;
            return result;
        }
        catch
        {
            result.Dispose();
            throw;
        }
        finally
        {
            Release(right); Release(cellRange); Release(cell); Release(cells);
            Release(tables); Release(story); Release(evidence); Release(owner);
            Release(paragraph); Release(paragraphs);
        }
    }

    internal Range ResolveInsertedRange()
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? shapes = null;
        try
        {
            bookmarks = _document.Bookmarks;
            if (!bookmarks.Exists(LocatorName))
                throw new InvalidOperationException("Word did not materialize the local MathType transaction locator.");
            bookmark = bookmarks[LocatorName];
            range = bookmark.Range;
            shapes = range.InlineShapes;
            if (shapes.Count != 1 || !range.InStory(_storyAnchor!))
                throw new InvalidOperationException("The local MathType transaction does not own exactly one OLE.");
            var result = range;
            range = null;
            return result;
        }
        finally { Release(shapes); Release(range); Release(bookmark); Release(bookmarks); }
    }

    internal void RebaseAfterPrecedingInsertion(int position, int length)
    {
        if (length <= 0 || position > _paragraphStart + _precedingInsertionLength)
            throw new InvalidOperationException("Only a proven insertion before the owned MathType paragraph can rebase it.");
        // The outside COM witness moves automatically. Rebase just the bounded
        // working start; keep original positions/signature for native Undo proof.
        _precedingInsertionLength = checked(_precedingInsertionLength + length);
    }

    // Also used by standalone insertion: compare only the captured owning
    // paragraph(s), including a newly prepared display row, never the document.
    internal int ReadParagraphCount()
    {
        Range? scope = null;
        Paragraphs? paragraphs = null;
        try
        {
            scope = CurrentParagraph();
            paragraphs = scope.Paragraphs;
            return paragraphs.Count;
        }
        finally { Release(paragraphs); Release(scope); }
    }

    internal void RepairAndVerifyParagraph(Range objectRange)
    {
        Range? scope = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        Range? mark = null;
        Range? following = null;
        try
        {
            scope = CurrentParagraph();
            paragraphs = scope.Paragraphs;
            var delta = paragraphs.Count - _paragraphCount;
            if (delta == 1)
            {
                Release(paragraphs);
                paragraphs = objectRange.Paragraphs;
                paragraph = paragraphs[1];
                owner = paragraph.Range;
                mark = objectRange.Duplicate;
                mark.SetRange(owner.End - 1, owner.End);
                if (mark.Start < _paragraphStart + _precedingInsertionLength || mark.End >= scope.End || mark.Text != "\r")
                    throw new InvalidOperationException("Word's extra OLE paragraph boundary is not inside the owned local paragraph.");
                following = objectRange.Duplicate;
                following.SetRange(mark.End, Math.Min(scope.End, mark.End + 64));
                var actualWhitespace = LeadingWhitespace(following.Text ?? string.Empty);
                if (actualWhitespace.Length == 0 && _rightWhitespace.Length == 0) mark.Delete();
                else
                {
                    mark.SetRange(mark.Start, mark.End + actualWhitespace.Length);
                    mark.Text = _rightWhitespace;
                }
            }
            else if (delta != 0)
                throw new InvalidOperationException("The native OLE edit changed unexpected local paragraph boundaries.");

            Release(paragraphs); paragraphs = null;
            Release(scope); scope = CurrentParagraph();
            paragraphs = scope.Paragraphs;
            if (paragraphs.Count != _paragraphCount
                || !string.Equals(scope.Text ?? string.Empty, _paragraphText, StringComparison.Ordinal))
                throw new InvalidOperationException("The native OLE edit did not preserve its surrounding paragraph text.");
        }
        finally
        {
            Release(following); Release(mark); Release(owner); Release(paragraph);
            Release(paragraphs); Release(scope);
        }
    }

    // Opt-in for a cross-family replacement. Existing MathType editing does not
    // read these extra ranges and keeps its established performance and checks.
    internal void CaptureReplacementBoundary(Range selected)
    {
        if (_replacementBoundaryCaptured || !selected.InStory(_storyAnchor!))
            throw new InvalidOperationException("The replacement is outside its captured local paragraph.");
        Range? owner = null;
        Range? evidence = null;
        Range? left = null;
        Range? right = null;
        Range? story = null;
        Paragraph? paragraph = null;
        Paragraphs? paragraphs = null;
        try
        {
            // Preflight's next-character witness can itself belong to a later
            // batch target (for example its leading MathType center tab). That
            // target has already been replaced in reverse-order conversion.
            // Bind an outside witness from the CURRENT exact source paragraph
            // immediately before this mutation, not from stale preflight ranges.
            // Original positions/XML/signature remain immutable for native Undo.
            paragraphs = selected.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException("The replacement source no longer has one current paragraph owner.");
            paragraph = paragraphs[1];
            owner = paragraph.Range.Duplicate;
            if (owner.Start != checked(_paragraphStart + _precedingInsertionLength)
                || selected.Start < owner.Start || selected.End > owner.End)
                throw new InvalidOperationException("The prepared replacement escaped its current local paragraph.");
            var previousWitnessEnd = _followingCharacter?.End ?? -1;
            story = selected.Duplicate;
            story.Expand(WdUnits.wdStory);
            Release(_followingCharacter); _followingCharacter = null;
            if (owner.End < story.End)
            {
                _followingCharacter = selected.Duplicate;
                _followingCharacter.SetRange(owner.End, owner.End + 1);
            }
            WordFormulaMutationTransaction.TraceBoundary(_document.Application,
                $"replacement-current-boundary oldWitnessEnd={previousWitnessEnd} ownerEnd={owner.End}");
            Release(paragraphs); paragraphs = null;
            left = selected.Duplicate;
            left.SetRange(owner.Start, selected.Start);
            right = selected.Duplicate;
            right.SetRange(selected.End, owner.End);
            _replacementLeftText = left.Text ?? string.Empty;
            _replacementRightText = right.Text ?? string.Empty;
            paragraphs = owner.Paragraphs;
            _replacementParagraphCount = paragraphs.Count;
            evidence = owner.Duplicate;
            evidence.Start = checked(_evidenceStart + _precedingInsertionLength);
            _replacementOleCount = WordInlineObjectGeometry.Capture(evidence).Count(item => item.Type is 1 or 2);
            if (_replacementOleCount < 1)
                throw new InvalidOperationException("The prepared replacement no longer contains its source OLE.");
            _replacementBoundaryCaptured = true;
        }
        finally { Release(paragraphs); Release(paragraph); Release(story); Release(right); Release(left); Release(evidence); Release(owner); }
    }

    internal void VerifyReplacementBoundary(Range component)
    {
        if (!_replacementBoundaryCaptured)
            throw new InvalidOperationException("The exact replacement boundary was not captured before writing.");
        Range? owner = null;
        Range? left = null;
        Range? right = null;
        Range? mark = null;
        Range? evidence = null;
        Paragraphs? paragraphs = null;
        try
        {
            owner = CurrentParagraph();
            paragraphs = owner.Paragraphs;
            left = component.Duplicate;
            left.SetRange(owner.Start, component.Start);
            right = component.Duplicate;
            right.SetRange(component.End, owner.End);
            // InsertXML can materialize one transport paragraph between an
            // imported inline fragment and its untouched right-hand text. Only
            // that exact added mark, proven by the outside witness and suffix,
            // is ours to remove. Never trim user whitespace or another row.
            if (paragraphs.Count == _replacementParagraphCount + 1
                && string.Equals(left.Text ?? string.Empty, _replacementLeftText, StringComparison.Ordinal)
                && string.Equals(right.Text ?? string.Empty, "\r" + _replacementRightText, StringComparison.Ordinal))
            {
                mark = component.Duplicate;
                mark.SetRange(component.End, component.End + 1);
                if (mark.Text != "\r" || mark.End > owner.End)
                    throw new InvalidOperationException("The imported transport mark is not locally owned.");
                mark.Delete();
                Release(paragraphs); paragraphs = null;
                Release(owner); owner = CurrentParagraph();
                paragraphs = owner.Paragraphs;
                left.SetRange(owner.Start, component.Start);
                right.SetRange(component.End, owner.End);
            }
            if (paragraphs.Count != _replacementParagraphCount
                || !string.Equals(left.Text ?? string.Empty, _replacementLeftText, StringComparison.Ordinal)
                || !string.Equals(right.Text ?? string.Empty, _replacementRightText, StringComparison.Ordinal))
                throw new InvalidOperationException("The native replacement changed text or paragraph boundaries outside its exact source.");

            // This left context is read-only evidence, just as in Capture. It is
            // never used as the destination of InsertXML or a delete operation.
            evidence = owner.Duplicate;
            evidence.Start = checked(_evidenceStart + _precedingInsertionLength);
            var remaining = WordInlineObjectGeometry.Capture(evidence).Count(item => item.Type is 1 or 2);
            if (remaining != _replacementOleCount - 1)
                throw new InvalidOperationException("The native replacement did not remove exactly one physical local OLE.");
        }
        finally
        {
            Release(paragraphs); Release(evidence); Release(mark);
            Release(right); Release(left); Release(owner);
        }
    }

    internal void RestoreContainedBookmarks(Range replacement, string? retiredIdentityBookmark = null)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = _document.Bookmarks;
            foreach (var span in _bookmarks)
            {
                if (string.Equals(span.Name, retiredIdentityBookmark, StringComparison.OrdinalIgnoreCase)) continue;
                Range? range = null;
                Bookmark? bookmark = null;
                try
                {
                    range = replacement.Duplicate;
                    range.SetRange(Map(span.Start, replacement), Map(span.End, replacement));
                    bookmark = bookmarks.Add(span.Name, range);
                }
                finally { Release(bookmark); Release(range); }
            }
        }
        finally { Release(bookmarks); }
    }

    internal void RemoveLocator()
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = _document.Bookmarks;
            if (!bookmarks.Exists(LocatorName))
                throw new InvalidOperationException("The MathType transaction locator disappeared before commit.");
            bookmark = bookmarks[LocatorName];
            bookmark.Delete();
        }
        finally { Release(bookmark); Release(bookmarks); }
    }

    internal void VerifyRollback()
    {
        Range? range = null;
        Range? insertion = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        try
        {
            range = _storyAnchor!.Duplicate;
            range.SetRange(_evidenceStart, _paragraphEnd);
            if (_insertionCapture)
            {
                var serialized = WordLocalEditSnapshot.Signature(
                    MathTypeLocalFragment.InsertionRecoveryEvidence(range.WordOpenXML));
                if (!string.Equals(
                        serialized,
                        _insertionSerializedSignature,
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Word undo did not restore the insertion paragraph content, formatting, revisions and bookmarks.");

                var position = checked(_paragraphStart + _insertionOffset);
                insertion = _document.Range(position, position);
                paragraphs = insertion.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidOperationException(
                        "Word undo did not restore one insertion paragraph owner.");
                paragraph = paragraphs[1];
                owner = paragraph.Range;
                if (owner.Start != _paragraphStart
                    || owner.End != _paragraphEnd
                    || !string.Equals(owner.Text ?? string.Empty, _paragraphText, StringComparison.Ordinal)
                    || !string.Equals(
                        CaptureInsertionBoundaryGeometry(_document, owner, position),
                        _insertionBoundaryGeometry,
                        StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Word undo did not restore the exact local MathType insertion boundary.");
                return;
            }

            if (!string.Equals(RecoverySignature(range, range.WordOpenXML),
                    _originalSignature, StringComparison.Ordinal))
                throw new InvalidOperationException("Word undo did not restore the original local MathType content, layout and bookmarks.");
        }
        finally
        {
            Release(owner); Release(paragraph); Release(paragraphs);
            Release(insertion); Release(range);
        }
    }

    private static string RecoverySignature(Range range, string xml,
        IReadOnlyList<WordInlineObjectGeometry.Item>? capturedGeometry = null)
    {
        var geometry = capturedGeometry ?? WordInlineObjectGeometry.Capture(range);
        foreach (var item in geometry)
            if (!(item.Width > 0) || !(item.Height > 0)
                || float.IsInfinity(item.Width) || float.IsInfinity(item.Height))
                throw new InvalidDataException("Word returned invalid local recovery geometry.");
        var live = new XElement("geometry", geometry.Select(item => new XElement("object",
            new XAttribute("start", item.Start), new XAttribute("end", item.End),
            new XAttribute("type", item.Type), new XAttribute("progId", item.ProgId),
            new XAttribute("width", item.Width.ToString("R", CultureInfo.InvariantCulture)),
            new XAttribute("height", item.Height.ToString("R", CultureInfo.InvariantCulture)))));
        return WordLocalEditSnapshot.Signature(MathTypeLocalFragment.RecoveryEvidence(
            xml, live.ToString(SaveOptions.DisableFormatting), geometry.Count(item => item.Type is 1 or 2)));
    }

    private Range CurrentParagraph()
    {
        var range = _storyAnchor!.Duplicate;
        try
        {
            int end;
            if (_followingCharacter is not null) end = _followingCharacter.End - 1;
            else { range.Expand(WdUnits.wdStory); end = range.End; }
            var start = checked(_paragraphStart + _precedingInsertionLength);
            if (end < start)
                throw new InvalidOperationException("The native OLE edit lost its outside paragraph boundary.");
            range.SetRange(start, end);
            return range;
        }
        catch { Release(range); throw; }
    }

    private int Map(int offset, Range replacement)
    {
        if (offset == 0) return replacement.Start;
        if (offset == _sourceLength) return replacement.End;
        if (replacement.End - replacement.Start != _sourceLength)
            throw new InvalidOperationException("Word changed an internally bookmarked OLE field span.");
        return replacement.Start + offset;
    }

    private static string CaptureInsertionBoundaryGeometry(
        Document document,
        Range owner,
        int position)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        OLEFormat? ole = null;
        try
        {
            const int Radius = 8;
            var start = Math.Max(owner.Start, position - Radius);
            var end = Math.Min(owner.End, position + Radius);
            if (end <= start && owner.End > owner.Start)
            {
                start = Math.Max(owner.Start, Math.Min(position, owner.End - 1));
                end = Math.Min(owner.End, start + 1);
            }
            probe = document.Range(start, end);
            shapes = probe.InlineShapes;
            var rows = new List<string>();
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(ole); ole = null;
                Release(shapeRange); shapeRange = null;
                Release(shape); shape = shapes[index];
                shapeRange = shape.Range;
                if (shapeRange.StoryType != owner.StoryType
                    || shapeRange.Start < owner.Start
                    || shapeRange.End > owner.End)
                    throw new InvalidDataException(
                        "A boundary-adjacent inline object escaped the insertion paragraph.");
                if (shapeRange.End < start || shapeRange.Start > end)
                    continue;
                var type = shape.Type;
                var progId = string.Empty;
                if (type is WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                    or WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                {
                    ole = shape.OLEFormat;
                    progId = ole.ProgID;
                }
                rows.Add(string.Join(",",
                    shapeRange.Start - position,
                    shapeRange.End - position,
                    (int)type,
                    progId,
                    shape.Width.ToString("R", CultureInfo.InvariantCulture),
                    shape.Height.ToString("R", CultureInfo.InvariantCulture)));
            }
            return string.Join("|", rows);
        }
        finally
        {
            Release(ole); Release(shapeRange); Release(shape);
            Release(shapes); Release(probe);
        }
    }

    private static string LeadingWhitespace(string text)
    {
        var count = 0;
        while (count < text.Length && (text[count] == ' ' || text[count] == '\t' || text[count] == '\u00A0')) count++;
        return text.Substring(0, count);
    }

    public void Dispose()
    {
        Release(_followingCharacter); _followingCharacter = null;
        Release(_storyAnchor); _storyAnchor = null;
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}

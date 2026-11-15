using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordFormulaHostSetup;
using static VisualTeX.WordVsto.WordInlineOlePosition;

namespace VisualTeX.WordVsto;

/// <summary>
/// The baseline of inline formulas and the text around them: baseline bookmarks and sentinels, typing anchors after
/// an inline OLE, non-prose host ranges and the typing format source.
/// </summary>
internal static class WordInlineBaseline
{
    internal static void NormalizeFollowingInlineProseBaseline(Range formulaRange)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? precedingWindow = null;
        Range? preceding = null;
        Range? trailingHostWindow = null;
        Range? trailing = null;
        InlineShapes? shapes = null;
        InlineShape? nextShape = null;
        Range? nextShapeRange = null;
        OMaths? maths = null;
        OMath? nextMath = null;
        Range? nextMathRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);

            // Only the nearby prose can determine the insertion baseline. Capturing
            // every InlineShape/OMath in a paragraph made appending the Nth inline
            // formula rescan formulas 1..N-1 twice. Keep the format probe bounded;
            // hidden OLE field-code ranges inside this small window are still
            // excluded by the same host-aware logic used previously.
            const int maximumProbeCharacters = 256;
            var precedingStart = Math.Max(
                paragraphRange.Start,
                formulaRange.Start - maximumProbeCharacters);
            if (formulaRange.Start > precedingStart)
            {
                precedingWindow = document.Range(precedingStart, formulaRange.Start);
                var excludedHostRanges = CaptureNonProseHostRanges(precedingWindow);
                preceding = FindOrdinaryVisibleCharacterRange(
                    document,
                    formulaRange.Start - 1,
                    precedingStart,
                    step: -1,
                    excludedHostRanges);
            }

            var targetPosition = 0;
            if (preceding is not null)
            {
                font = preceding.Font;
                var precedingPosition = font.Position;
                if (precedingPosition != (int)WdConstants.wdUndefined
                    && precedingPosition >= -256
                    && precedingPosition <= 256)
                    targetPosition = precedingPosition;
                Release(font);
                font = null;
            }

            var trailingEnd = Math.Max(formulaRange.End, paragraphBodyEnd);
            if (trailingEnd <= formulaRange.End) return;
            trailingHostWindow = document.Range(formulaRange.End, trailingEnd);

            // Range collections are in document order. The first following host is
            // sufficient; do not enumerate every object in the paragraph.
            shapes = trailingHostWindow.InlineShapes;
            if (shapes.Count > 0)
            {
                nextShape = shapes[1];
                nextShapeRange = nextShape.Range;
                trailingEnd = Math.Min(trailingEnd, nextShapeRange.Start);
            }
            maths = trailingHostWindow.OMaths;
            if (maths.Count > 0)
            {
                nextMath = maths[1];
                nextMathRange = nextMath.Range;
                trailingEnd = Math.Min(trailingEnd, nextMathRange.Start);
            }

            if (trailingEnd <= formulaRange.End) return;
            trailing = document.Range(formulaRange.End, trailingEnd);
            if (!ContainsVisibleBodyText(trailing.Text)) return;
            font = trailing.Font;
            var currentPosition = font.Position;
            if (currentPosition == (int)WdConstants.wdUndefined
                || currentPosition != targetPosition)
                font.Position = targetPosition;
        }
        catch
        {
            // Baseline repair is best-effort and must not interrupt Word input.
        }
        finally
        {
            Release(font);
            Release(nextMathRange);
            Release(nextMath);
            Release(maths);
            Release(nextShapeRange);
            Release(nextShape);
            Release(shapes);
            Release(trailing);
            Release(trailingHostWindow);
            Release(preceding);
            Release(precedingWindow);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    internal static void ResetParagraphTypingPosition(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? nextCharacter = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End > paragraphRange.Start)
            {
                paragraphMark = paragraphRange.Duplicate;
                paragraphMark.SetRange(paragraphRange.End - 1, paragraphRange.End);
                ResetRangeFontPosition(paragraphMark);
            }

            if (formulaRange.End >= paragraphRange.End) return;
            nextCharacter = paragraphRange.Duplicate;
            nextCharacter.SetRange(
                formulaRange.End,
                Math.Min(formulaRange.End + 1, paragraphRange.End));
            if (nextCharacter.Text is "\v" or "\r" or "\n")
                ResetRangeFontPosition(nextCharacter);
        }
        catch
        {
            // Baseline restoration is best-effort and must not invalidate the
            // formula that has already been inserted or resized.
        }
        finally
        {
            Release(nextCharacter);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static string InlineBaselineBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX formulaId must be a UUID.");
        return InlineBaselineBookmarkPrefix + parsed.ToString("N");
    }


    private static bool IsKnownInlineBaselineSentinel(string? text) =>
        string.IsNullOrEmpty(text)
        || string.Equals(text, InlineOleTypingAnchor, StringComparison.Ordinal)
        || string.Equals(text, InlineBaselineSentinel, StringComparison.Ordinal)
        || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal)
        || string.Equals(
            text,
            LegacyInlineNonbreakingBaselineSentinel,
            StringComparison.Ordinal);

    private static bool IsHiddenTextRange(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            return font.Hidden != 0;
        }
        catch { return false; }
        finally { Release(font); }
    }

    private static bool RangeContainsMath(Range range)
    {
        OMaths? maths = null;
        try
        {
            maths = range.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
    }

    internal static void RemoveInlineOleTypingAnchorAfter(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            RemoveInlineOleTypingAnchorAfter(range);
        }
        finally { Release(range); }
    }

    internal static void RemoveInlineOleTypingAnchorAfter(Range formulaRange)
    {
        Document? document = null;
        Range? content = null;
        Range? anchor = null;
        try
        {
            document = formulaRange.Document;
            content = document.Content;
            var position = formulaRange.End;
            var contentStart = content.Start;
            var contentEnd = content.End;
            if (position < contentStart || position >= contentEnd) return;
            anchor = document.Range(position, Math.Min(position + 1, contentEnd));
            if (!string.Equals(
                    anchor.Text,
                    InlineOleTypingAnchor,
                    StringComparison.Ordinal))
                return;
            // Word can report an ordinary character immediately after an OLE as
            // math-affiliated because the adjacent object participates in the
            // same layout run. This helper is called only for a confirmed
            // VisualTeX inline OLE, whose first U+200C at Range.End is owned by us.
            anchor.Delete();
        }
        catch
        {
            // The formula may have been deleted by Word between range capture and
            // cleanup. Orphan-anchor removal is best-effort in rollback paths.
        }
        finally
        {
            Release(anchor);
            Release(content);
            Release(document);
        }
    }

    internal static void RemoveInlineBaselineSentinel(
        Document document, string formulaId, Range? protectedFollowingText = null)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? sentinel = null;
        try
        {
            var name = InlineBaselineBookmarkName(formulaId);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            sentinel = bookmark.Range;
            var sentinelStart = sentinel.Start;
            var ownedBoundaryHasWidth = sentinel.Start < sentinel.End
                && IsKnownInlineBaselineSentinel(sentinel.Text)
                && (protectedFollowingText is null || sentinel.End <= protectedFollowingText.Start);
            bookmark.Delete();
            if (ownedBoundaryHasWidth)
            {
                // Range.Delete at an OMath boundary can apply Word's smart
                // whitespace deletion and consume the next user-authored space.
                // Empty-text replacement addresses only the captured guard span.
                if (protectedFollowingText is not null) sentinel.Text = string.Empty;
                else sentinel.Delete();
            }

            // Remove a temporary guard from interrupted insertions. Probe both
            // sides of the former bookmark because deleting the bookmarked marker
            // shifts the following text one position to the left.
            TryDeleteTemporaryInlineBoundaryAt(document, sentinelStart - 1, protectedFollowingText);
            TryDeleteTemporaryInlineBoundaryAt(document, sentinelStart, protectedFollowingText);
        }
        catch
        {
            // A stale or externally edited sentinel must never block formula work.
        }
        finally
        {
            Release(sentinel);
            Release(bookmark);
            Release(bookmarks);
        }
    }


    private static bool TryDeleteTemporaryInlineBoundaryAt(
        Document document,
        int position,
        Range? protectedFollowingText = null)
    {
        Range? candidate = null;
        try
        {
            if (protectedFollowingText is not null && position >= protectedFollowingText.Start)
                return false;
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            if (position < contentStart || position >= contentEnd) return false;
            object candidateStart = position;
            object candidateEnd = Math.Min(contentEnd, position + 1);
            candidate = document.Range(ref candidateStart, ref candidateEnd);
            var text = candidate.Text;
            var removable =
                string.Equals(text, BulkInlineFormulaPlaceholder, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineMathGuard, StringComparison.Ordinal)
                || string.Equals(text, LegacyInlineBaselineSentinel, StringComparison.Ordinal)
                || string.Equals(text, InlineMathGuard, StringComparison.Ordinal)
                    && IsHiddenTextRange(candidate);
            // U+00A0 is also the correct Word representation of explicit LaTeX
            // spacing (for example `~` and `\ `). Delete it only when the VTBL
            // bookmark itself owns that legacy marker, never by proximity.
            if (!removable || RangeContainsMath(candidate)) return false;
            if (protectedFollowingText is not null) candidate.Text = string.Empty;
            else candidate.Delete();
            return true;
        }
        catch { return false; }
        finally { Release(candidate); }
    }

    internal static Range? FindInlineTypingFormatSource(
        Range formulaRange,
        Range typingAnchor)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? precedingWindow = null;
        Range? followingWindow = null;
        try
        {
            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            const int maximumProbeCharacters = 256;

            // The prose before an inline formula is the authoritative typing
            // style. Restrict host discovery to the same bounded region that can
            // actually be probed; scanning every OLE in a long formula paragraph
            // made caret restoration grow linearly with formula count.
            var precedingStart = Math.Max(
                paragraphRange.Start,
                formulaRange.Start - maximumProbeCharacters);
            if (formulaRange.Start > precedingStart)
            {
                precedingWindow = document.Range(precedingStart, formulaRange.Start);
                var precedingHosts = CaptureNonProseHostRanges(precedingWindow);
                var preceding = FindOrdinaryVisibleCharacterRange(
                    document,
                    formulaRange.Start - 1,
                    precedingStart,
                    step: -1,
                    precedingHosts);
                if (preceding is not null) return preceding;
            }

            // Paragraph-leading formulas have no preceding prose. In that case,
            // inherit from the first nearby ordinary character after the formula.
            var followingStart = Math.Max(typingAnchor.End, formulaRange.End);
            var followingEnd = Math.Min(
                paragraphBodyEnd,
                followingStart + maximumProbeCharacters);
            if (followingEnd <= followingStart) return null;
            followingWindow = document.Range(followingStart, followingEnd);
            var followingHosts = CaptureNonProseHostRanges(followingWindow);
            return FindOrdinaryVisibleCharacterRange(
                document,
                followingStart,
                followingEnd,
                step: 1,
                followingHosts);
        }
        finally
        {
            Release(followingWindow);
            Release(precedingWindow);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    private static Range? FindOrdinaryVisibleCharacterRange(
        Document document,
        int startPosition,
        int boundaryPosition,
        int step,
        IReadOnlyList<NonProseHostRange> excludedHostRanges)
    {
        if (step is not (-1 or 1)) return null;
        const int maximumProbeCharacters = 256;
        var position = startPosition;
        for (var probeIndex = 0;
             probeIndex < maximumProbeCharacters
             && (step > 0
                 ? position < boundaryPosition
                 : position >= boundaryPosition);
             probeIndex++, position += step)
        {
            Range? probe = null;
            try
            {
                if (position < 0) break;
                if (excludedHostRanges.Any(host => host.Contains(position)))
                    continue;
                probe = document.Range(position, position + 1);
                if (!ContainsVisibleBodyText(probe.Text)) continue;
                var result = probe.Duplicate;
                return result;
            }
            catch (COMException)
            {
                // Keep probing nearby ordinary prose.
            }
            finally
            {
                Release(probe);
            }
        }
        return null;
    }

    private static List<NonProseHostRange> CaptureNonProseHostRanges(
        Range paragraphRange)
    {
        var result = new List<NonProseHostRange>();
        InlineShapes? shapes = null;
        OMaths? maths = null;
        try
        {
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = shapes[index];
                    range = shape.Range;
                    result.Add(new NonProseHostRange(range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            maths = paragraphRange.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    result.Add(new NonProseHostRange(range.Start, range.End));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
            return result;
        }
        finally
        {
            Release(maths);
            Release(shapes);
        }
    }

    internal static void ResetRangeFontPosition(Range range)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = range.Font;
            // A redundant format assignment still dirties the Word run and can
            // repaginate the paragraph. Healthy OMML boundaries are normally
            // already at the neutral baseline.
            if (font.Position != 0)
                font.Position = 0;
        }
        finally { Release(font); }
    }

    internal static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        double semanticFontSizePoints) =>
        ApplyInlineBaseline(
            shape,
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            existingFontPosition: null,
            sourceSemanticFontSizePoints: semanticFontSizePoints,
            targetSemanticFontSizePoints: semanticFontSizePoints);

    internal static void ApplyInlineBaseline(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition,
        double sourceSemanticFontSizePoints,
        double targetSemanticFontSizePoints)
    {
        var position = CalculateVisualTeXInlinePosition(
            shape,
            actualHeightPoints,
            exportedHeight,
            exportedBaseline,
            existingFontPosition,
            sourceSemanticFontSizePoints,
            targetSemanticFontSizePoints);

        // Word stores an InlineShape OLE as an EMBED field whose instruction and
        // U+0001 object/result character share shape.Range. Applying Position to
        // the complete field lowers hidden instruction characters as well, which
        // inflates the line box and can leak the negative position into following
        // prose. Match Word/MathType behavior: keep the field instruction at the
        // paragraph baseline and move only the painted object character.
        SetInlineOleWordPosition(shape, position);
    }
}

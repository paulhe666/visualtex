using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using static VisualTeX.WordVsto.WordFormulaService;

namespace VisualTeX.WordVsto;

/// <summary>
/// Finding MathType OLE objects and their hosts in a document: by identity bookmark, by range, by inline-shape
/// index or near a position, and the object a command just created.
/// </summary>
internal static class WordMathTypeOleLookup
{
    private static InlineShape? FindNewMathTypeOleAtStart(
        Document document,
        int start,
        InlineShape sourceShape)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                var keepCandidate = false;
                try
                {
                    candidate = shapes[index];
                    range = candidate.Range;
                    if (range.Start != start
                        || IsSameComObject(candidate, sourceShape)
                        || !MathTypeOleInterop.IsMathTypeOle(candidate))
                        continue;
                    keepCandidate = true;
                    return candidate;
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                finally
                {
                    Release(range);
                    if (!keepCandidate) Release(candidate);
                }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private static bool IsSameComObject(object left, object right)
    {
        if (ReferenceEquals(left, right)) return true;
        if (!Marshal.IsComObject(left) || !Marshal.IsComObject(right)) return false;
        IntPtr leftIdentity = IntPtr.Zero;
        IntPtr rightIdentity = IntPtr.Zero;
        try
        {
            leftIdentity = Marshal.GetIUnknownForObject(left);
            rightIdentity = Marshal.GetIUnknownForObject(right);
            return leftIdentity == rightIdentity;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (leftIdentity != IntPtr.Zero) Marshal.Release(leftIdentity);
            if (rightIdentity != IntPtr.Zero) Marshal.Release(rightIdentity);
        }
    }

    internal static Range? FindMathTypeHostRangeByIdentityBookmark(
        Document document,
        string bookmarkName) =>
        FindMathTypeHostRangeByIdentityBookmark(document, bookmarkName, out _);

    // The single MathType OLE in the bounded insertion window, with the
    // identity bookmark set on it again; null unless there is exactly one.
    internal static Range? RecoverMathTypeHostRangeAtInsertion(
        Document document,
        string bookmarkName,
        int insertionStart)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? restored = null;
        try
        {
            shape = FindSingleInlineShapeInInsertionWindow(document, insertionStart);
            if (shape is null || !MathTypeOleInterop.IsMathTypeOle(shape))
                return null;
            shapeRange = shape.Range;
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(bookmarkName))
                bookmarks[bookmarkName].Delete();
            restored = bookmarks.Add(bookmarkName, shapeRange);
            // The shape range is the host the resolver above also accepts when
            // an insertion has no bookmark at all.
            return FindMathTypeHostRangeByIdentityBookmark(document, bookmarkName)
                ?? shapeRange.Duplicate;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(restored);
            Release(bookmarks);
            Release(shapeRange);
            Release(shape);
        }
    }

    internal static Range? FindMathTypeHostRangeByIdentityBookmark(
        Document document,
        string bookmarkName,
        out string failure)
    {
        failure = "missing";
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
                return null;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range.Duplicate;
            fields = range.Fields;
            if (fields.Count != 1)
            {
                failure = $"fields={fields.Count} range={range.Start}-{range.End}";
                return null;
            }
            field = fields[1];
            if (field.Type != WdFieldType.wdFieldEmbed)
            {
                failure = "field-type=" + field.Type;
                return null;
            }
            code = field.Code;
            var codeText = (code.Text ?? string.Empty).Trim();
            if (!codeText.StartsWith(
                    "EMBED Equation.DSMT4",
                    StringComparison.OrdinalIgnoreCase))
            {
                failure = "field-code";
                return null;
            }
            result = field.Result;
            var resultText = result.Text ?? string.Empty;
            if (!string.Equals(resultText, "\u0001", StringComparison.Ordinal)
                && !string.Equals(resultText, "\uFFFC", StringComparison.Ordinal))
            {
                failure = "field-result-length=" + resultText.Length;
                return null;
            }
            if (result.Start < range.Start || result.End > range.End)
            {
                failure = $"result={result.Start}-{result.End} outside {range.Start}-{range.End}";
                return null;
            }
            var resolved = range;
            range = null;
            return resolved;
        }
        catch (Exception error)
        {
            failure = error.GetType().Name;
            return null;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static InlineShape? FindMathTypeOleByIdentityBookmark(
        Document document,
        string bookmarkName)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            bookmarks = document.Bookmarks;
            // The bookmark name is generated from the current formula UUID and is
            // imported atomically in the same Flat OPC as this OLE. A direct Item
            // lookup is sufficient; Exists + Item doubles Word's global bookmark
            // lookup cost on every formula.
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range.Duplicate;
            shapes = range.InlineShapes;
            if (shapes.Count != 1)
                return null;
            candidate = shapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(candidate))
                return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(shapes);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static InlineShape? FindSingleInlineShapeAtPosition(
        Document document,
        int position)
    {
        Range? content = null;
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            content = document.Content;
            var start = Math.Max(content.Start, Math.Min(position, content.End));
            var end = Math.Min(content.End, start + 2);
            if (end <= start)
                return null;
            probe = document.Range(start, end);
            shapes = probe.InlineShapes;
            if (shapes.Count != 1)
                return null;
            candidate = shapes[1];
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(shapes);
            Release(probe);
            Release(content);
        }
    }

    private static InlineShape? FindSingleInlineShapeInExactRange(
        Range insertion)
    {
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            shapes = insertion.InlineShapes;
            if (shapes.Count != 1)
                return null;
            candidate = shapes[1];
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(shapes);
        }
    }

    internal static InlineShape? FindSingleInlineShapeInInsertionWindow(
        Document document,
        int insertionStart)
    {
        Range? content = null;
        Range? window = null;
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            content = document.Content;
            var start = Math.Max(content.Start, insertionStart - 1);
            var end = Math.Min(content.End, insertionStart + 8);
            if (end <= start)
                return null;
            window = document.Range(start, end);
            shapes = window.InlineShapes;
            if (shapes.Count != 1)
                return null;
            candidate = shapes[1];
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(shapes);
            Release(window);
            Release(content);
        }
    }

    internal static InlineShape? FindMathTypeOleByRange(
        Document document,
        string? sourceObjectId,
        bool allowGlobalFallback = true)
    {
        if (!TryParseRangeReference(sourceObjectId, out var start, out var end))
            return null;

        Range? hintedRange = null;
        InlineShapes? hintedShapes = null;
        Range? content = null;
        Range? localRange = null;
        InlineShapes? localShapes = null;
        InlineShapes? shapes = null;
        try
        {
            try
            {
                hintedRange = document.Range(start, end);
                hintedShapes = hintedRange.InlineShapes;
                for (var index = 1; index <= hintedShapes.Count; index++)
                {
                    InlineShape? candidate = null;
                    try
                    {
                        candidate = hintedShapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { Release(candidate); }
                }
            }
            catch (Exception ignored)
            {
                // Fall back to a document scan below. The source range can move
                // if surrounding prose changes while the VisualTeX editor is open.
                IgnoredErrors.Ignore(ignored);
            }

            // During whole-document format conversion targets are processed from
            // the end toward the start, so the captured MathType range remains a
            // reliable local locator. After the source has been deleted the exact
            // range is expected to contain no MathType object. Probe only a tiny
            // neighborhood for Word's occasional one/two-character range drift;
            // never enumerate the whole document merely to prove absence.
            try
            {
                content = document.Content;
                var localStart = Math.Max(content.Start, start - 2);
                var localEnd = Math.Min(content.End, Math.Max(end, start) + 2);
                if (localEnd > localStart)
                {
                    localRange = document.Range(localStart, localEnd);
                    localShapes = localRange.InlineShapes;
                    InlineShape? nearestLocal = null;
                    var nearestLocalDistance = int.MaxValue;
                    try
                    {
                        for (var index = 1; index <= localShapes.Count; index++)
                        {
                            InlineShape? candidate = null;
                            Range? candidateRange = null;
                            try
                            {
                                candidate = localShapes[index];
                                if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                                candidateRange = candidate.Range;
                                var distance = Math.Abs(candidateRange.Start - start);
                                if (distance >= nearestLocalDistance) continue;
                                Release(nearestLocal);
                                nearestLocal = candidate;
                                candidate = null;
                                nearestLocalDistance = distance;
                            }
                            finally
                            {
                                Release(candidateRange);
                                Release(candidate);
                            }
                        }
                        if (nearestLocal is not null && nearestLocalDistance <= 2)
                        {
                            var result = nearestLocal;
                            nearestLocal = null;
                            return result;
                        }
                    }
                    finally { Release(nearestLocal); }
                }
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            finally
            {
                Release(localShapes);
                localShapes = null;
                Release(localRange);
                localRange = null;
                Release(content);
                content = null;
            }

            if (!allowGlobalFallback) return null;

            shapes = document.InlineShapes;
            InlineShape? nearest = null;
            var nearestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    if (candidateRange.Start == start && candidateRange.End == end)
                    {
                        Release(nearest);
                        nearest = null;
                        var exact = candidate;
                        candidate = null;
                        return exact;
                    }
                    var distance = Math.Abs(candidateRange.Start - start);
                    if (distance >= nearestDistance) continue;
                    Release(nearest);
                    nearest = candidate;
                    candidate = null;
                    nearestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            // A range is only a transient locator for third-party OLE. Avoid
            // accidentally editing a different MathType equation after large
            // document changes: permit only a very small local shift.
            if (nearest is not null && nearestDistance <= 2)
            {
                var result = nearest;
                nearest = null;
                return result;
            }
            Release(nearest);
            return null;
        }
        finally
        {
            Release(shapes);
            Release(localShapes);
            Release(localRange);
            Release(content);
            Release(hintedShapes);
            Release(hintedRange);
        }
    }

    private static int ResolveInlineShapeInsertionIndex(Document document, int position)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        try
        {
            shapes = document.InlineShapes;
            var preceding = 0;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(range);
                range = null;
                Release(shape);
                shape = shapes[index];
                range = shape.Range;
                if (range.Start >= position) break;
                preceding++;
            }
            return preceding + 1;
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
        }
    }

    private static InlineShape? FindMathTypeOleAtIndex(Document document, int index)
    {
        InlineShapes? shapes = null;
        InlineShape? candidate = null;
        try
        {
            shapes = document.InlineShapes;
            if (index < 1 || index > shapes.Count) return null;
            candidate = shapes[index];
            if (!MathTypeOleInterop.IsMathTypeOle(candidate)) return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch { return null; }
        finally
        {
            Release(candidate);
            Release(shapes);
        }
    }

    internal static InlineShape? FindMathTypeOleInParagraphAtPosition(
        Document document,
        int position)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        try
        {
            var safePosition = Math.Max(
                document.Content.Start,
                Math.Min(document.Content.End - 1, position));
            probe = document.Range(safePosition, safePosition);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            InlineShape? match = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    if (match is not null)
                    {
                        Release(match);
                        return null;
                    }
                    match = candidate;
                    candidate = null;
                }
                finally { Release(candidate); }
            }
            return match;
        }
        catch { return null; }
        finally
        {
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    internal static InlineShape? FindMathTypeOleInLocalWindow(
        Document document,
        int position)
    {
        Range? window = null;
        InlineShapes? shapes = null;
        InlineShape? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            var start = Math.Max(contentStart, position - 8);
            var end = Math.Min(contentEnd, position + 256);
            if (end <= start) end = Math.Min(contentEnd, start + 1);
            window = document.Range(start, end);
            shapes = window.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    candidateRange = candidate.Range;
                    var distance = Math.Abs(candidateRange.Start - position);
                    if (distance >= bestDistance) continue;
                    Release(best);
                    best = candidate;
                    candidate = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            return best;
        }
        catch
        {
            Release(best);
            return null;
        }
        finally
        {
            Release(shapes);
            Release(window);
        }
    }

    internal static InlineShape? FindMathTypeOleNearPosition(
        Document document,
        int position)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            InlineShape? best = null;
            var bestDistance = int.MaxValue;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? range = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    range = candidate.Range;
                    var distance = Math.Abs(range.Start - position);
                    if (distance > 3 || distance >= bestDistance) continue;
                    Release(best);
                    best = candidate;
                    candidate = null;
                    bestDistance = distance;
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                finally
                {
                    Release(range);
                    Release(candidate);
                }
            }
            return best;
        }
        finally { Release(shapes); }
    }
}

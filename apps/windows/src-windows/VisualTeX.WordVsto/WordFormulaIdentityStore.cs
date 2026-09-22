using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Owns durable identity only for external OLE hosts. Pure Word OMML has no
/// VisualTeX-owned identity; VTOMML_* is retained only as a legacy name for
/// compatibility detection and historical acceptance fixtures.
/// </summary>
internal static class WordFormulaIdentityStore
{
    internal const string OmmlBookmarkPrefix = "VTOMML_";
    internal const string MathTypeBookmarkPrefix = "VTMATH_";

    internal static string OmmlBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException(
                "Legacy OMML bookmark naming requires a UUID FormulaId.");
        return OmmlBookmarkPrefix + parsed.ToString("N");
    }

    internal static string MathTypeBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException(
                "MathType host identity requires a UUID FormulaId.");
        return MathTypeBookmarkPrefix + parsed.ToString("N");
    }

    internal static bool TryParseMathTypeBookmarkName(
        string? name,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(name)
            || !name.StartsWith(
                MathTypeBookmarkPrefix,
                StringComparison.OrdinalIgnoreCase))
            return false;

        var suffix =
            name.Substring(
                MathTypeBookmarkPrefix.Length);
        if (!Guid.TryParseExact(
                suffix,
                "N",
                out var parsed))
            return false;

        formulaId = parsed.ToString("D");
        return true;
    }

    internal static void BindVisualTeX(InlineShape shape, string formulaId)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            BindExactRange(
                range.Document,
                WordFormulaMetadataReader.IdentityBookmarkName(formulaId),
                range);
        }
        finally { Release(range); }
    }

    internal static void BindMathType(
        InlineShape shape,
        string formulaId)
    {
        Range? range = null;
        try
        {
            range = shape.Range.Duplicate;
            BindExactRange(
                range.Document,
                MathTypeBookmarkName(formulaId),
                range);
        }
        finally { Release(range); }
    }

    internal static bool TryResolveMathType(
        InlineShape shape,
        out string formulaId)
    {
        formulaId = string.Empty;
        Range? shapeRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        string? resolved = null;
        try
        {
            shapeRange = shape.Range.Duplicate;
            bookmarks = shapeRange.Bookmarks;
            for (var index = 1;
                 index <= bookmarks.Count;
                 index++)
            {
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!TryParseMathTypeBookmarkName(
                        bookmark.Name,
                        out var candidate))
                    continue;

                bookmarkRange =
                    bookmark.Range.Duplicate;
                if (bookmarkRange.StoryType != shapeRange.StoryType
                    || bookmarkRange.Start != shapeRange.Start
                    || bookmarkRange.End != shapeRange.End)
                    continue;

                if (resolved is not null
                    && !string.Equals(
                        resolved,
                        candidate,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException(
                        "One MathType host owns multiple VisualTeX identities.");
                resolved = candidate;
            }

            if (resolved is null)
                return false;
            formulaId = resolved;
            return true;
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(shapeRange);
        }
    }

    internal static void RemoveMathType(
        Document document,
        string? formulaId)
    {
        if (!Guid.TryParse(formulaId, out _)) return;
        RemoveByName(
            document,
            MathTypeBookmarkName(formulaId!));
    }

    internal static void RemoveVisualTeX(Document document, string? formulaId)
    {
        if (!Guid.TryParse(formulaId, out _)) return;
        RemoveByName(
            document,
            WordFormulaMetadataReader.IdentityBookmarkName(formulaId!));
    }

    private static void BindExactRange(
        Document document,
        string name,
        Range owner)
    {
        if (document.ReadOnly)
            throw new UnauthorizedAccessException(
                "VisualTeX cannot bind formula identity in a read-only document.");

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? verified = null;
        try
        {
            bookmarks = document.Bookmarks;
            // Bookmarks.Add with an existing name retargets the same logical
            // bookmark. There is no need to scan or pre-delete unrelated names.
            bookmark = bookmarks.Add(name, owner);
            verified = bookmark.Range;
            if (verified.StoryType != owner.StoryType
                || verified.Start != owner.Start
                || verified.End != owner.End)
                throw new InvalidDataException(
                    $"Word did not retain the exact formula identity range for {name}.");
        }
        finally
        {
            Release(verified);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RemoveByName(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            bookmark.Delete();
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

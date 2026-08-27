using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordFormulaMetadataReader
{
    private const string IdentityBookmarkPrefix = "VTO_";

    public static FormulaMetadata? TryRead(InlineShape shape)
    {
        if (shape is null) return null;
        var cached = TryReadCached(shape);
        if (cached is not null) return cached;
        return TryReadAuthoritative(shape);
    }

    internal static FormulaMetadata? TryReadAuthoritative(InlineShape shape)
    {
        if (shape is null || !IsNativeOle(shape)) return null;
        // AlternativeText/Title are only a Word-side cache. They can survive a
        // copy, numbering-host migration or an older failed conversion after the
        // native OLE object itself has different current metadata. Operations that
        // change formula format must use the VisualTeX OLE object's authoritative
        // IVisualTeXFormulaObject payload, never a potentially stale cache.
        return ApplyIdentityBookmark(shape, TryReadNativeOle(shape));
    }

    internal static FormulaMetadata? TryReadEmbeddedNativeOle(InlineShape shape)
    {
        if (shape is null || !IsNativeOle(shape)) return null;
        // Structural batch migration must sometimes repair VTO_ bookmarks after
        // Word expands/moves them while a preceding numbered table is dismantled.
        // In that narrow window the embedded OLE payload is the only authoritative
        // identity. Do not apply a possibly-corrupted Word identity bookmark here.
        return TryReadNativeOle(shape);
    }

    internal static FormulaMetadata? TryReadCached(InlineShape shape) =>
        ApplyIdentityBookmark(shape, TryReadCachedPreview(shape));

    internal static FormulaMetadata? TryReadCachedPreview(InlineShape shape)
    {
        if (shape is null) return null;
        string? encoded = null;
        try { encoded = shape.AlternativeText; } catch { }
        var metadata = FormulaMetadataCodec.Decode(encoded);
        if (metadata is null)
        {
            try { encoded = shape.Title; } catch { encoded = null; }
            metadata = FormulaMetadataCodec.Decode(encoded);
        }
        return metadata;
    }

    internal static void CacheMetadata(InlineShape shape, FormulaMetadata metadata)
    {
        if (shape is null || metadata is null) return;
        var encoded = FormulaMetadataCodec.Encode(metadata);
        try { shape.AlternativeText = encoded; } catch { }
        try { shape.Title = encoded; } catch { }
    }

    internal static string IdentityBookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var value))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        return $"{IdentityBookmarkPrefix}{value:N}";
    }

    internal static bool TryFormulaIdFromIdentityBookmark(
        string? bookmarkName,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(bookmarkName)
            || !bookmarkName!.StartsWith(IdentityBookmarkPrefix, StringComparison.Ordinal)
            || !Guid.TryParseExact(
                bookmarkName.Substring(IdentityBookmarkPrefix.Length),
                "N",
                out var value))
            return false;
        formulaId = value.ToString("D");
        return true;
    }

    internal static FormulaMetadata CloneWithFormulaId(
        FormulaMetadata metadata,
        string formulaId)
    {
        var clone = FormulaMetadataCodec.DeserializeJson(
            FormulaMetadataCodec.SerializeJson(metadata))
            ?? throw new InvalidOperationException(
                "Unable to clone VisualTeX formula metadata.");
        clone.FormulaId = formulaId;
        clone.UpdatedWithVersion = "1.2.5";
        clone.UpdatedAt = DateTimeOffset.UtcNow.ToString("O");
        clone.Validate();
        return clone;
    }

    public static void Write(InlineShape shape, FormulaMetadata metadata)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        if (metadata is null) throw new ArgumentNullException(nameof(metadata));
        metadata.Validate();

        if (IsNativeOle(shape))
        {
            OLEFormat? format = null;
            object? oleObject = null;
            try
            {
                format = shape.OLEFormat;
                oleObject = WordOleObjectAccessor.GetRunningObject(format);
                if (oleObject is not IVisualTeXFormulaObject formula)
                    throw new InvalidOperationException(
                        "The VisualTeX native OLE object is unavailable.");
                FormulaOleInterop.UpdateMetadata(formula, metadata);
                CacheMetadata(shape, metadata);
                return;
            }
            finally
            {
                Release(oleObject);
                Release(format);
            }
        }

        CacheMetadata(shape, metadata);
    }

    public static bool IsNativeOle(InlineShape shape)
    {
        OLEFormat? format = null;
        try
        {
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
        }
    }

    private static FormulaMetadata? TryReadNativeOle(InlineShape shape)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = WordOleObjectAccessor.GetRunningObject(format);
            return oleObject is IVisualTeXFormulaObject formula
                ? FormulaOleInterop.ReadMetadata(formula)
                : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static FormulaMetadata? ApplyIdentityBookmark(
        InlineShape shape,
        FormulaMetadata? metadata)
    {
        if (metadata is null) return null;
        Range? range = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            range = shape.Range;
            bookmarks = range.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!TryFormulaIdFromIdentityBookmark(bookmark.Name, out var formulaId))
                    continue;

                // Word may expose a bookmark that begins exactly at the right
                // boundary of this InlineShape through shape.Range.Bookmarks.
                // That bookmark belongs to an adjacent formula and must never
                // override the current OLE's embedded FormulaId.  Accept only an
                // identity bookmark whose own range actually identifies/contains
                // this InlineShape (including Word's collapsed-at-start repair
                // form used after object replacement).
                bookmarkRange = bookmark.Range;
                if (!IdentityBookmarkOwnsInlineShape(bookmarkRange, range))
                    continue;

                if (string.Equals(
                        metadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    return metadata;
                return CloneWithFormulaId(metadata, formulaId);
            }
        }
        catch
        {
            // A missing/invalid local identity bookmark must not hide otherwise
            // valid VisualTeX metadata.
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(range);
        }
        return metadata;
    }

    private static bool IdentityBookmarkOwnsInlineShape(Range owner, Range candidate)
    {
        if (owner.Start == candidate.Start && owner.End == candidate.End)
            return true;
        if (owner.Start == owner.End && owner.Start == candidate.Start)
            return true;
        return owner.Start <= candidate.Start && owner.End >= candidate.End;
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

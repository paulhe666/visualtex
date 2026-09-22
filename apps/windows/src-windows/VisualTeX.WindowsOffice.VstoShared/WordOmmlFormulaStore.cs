using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlFormulaStore
{
    internal const string NamespaceUri = "urn:visualtex:word-omml:1";
    internal const string BookmarkPrefix = "VTOMML_";
    private const string InlineBaselineBookmarkPrefix = "VTBL_";

    private static readonly XNamespace VisualTeXNamespace = NamespaceUri;
    // Keep the static initializer independent of the Office PIAs so pure XML
    // metadata tests can load this type without an installed Office assembly.
    private static readonly ConditionalWeakTable<object, DocumentMetadataCache>
        MetadataCaches = new();

    private sealed class CachedMetadataPart
    {
        internal string PartId { get; set; } = string.Empty;
        internal FormulaMetadata Metadata { get; set; } = new();
    }

    private sealed class DocumentMetadataCache
    {
        internal object Gate { get; } = new();
        internal bool Hydrated { get; set; }
        internal Dictionary<string, CachedMetadataPart> Entries { get; } =
            new(StringComparer.OrdinalIgnoreCase);
    }

    internal static string BookmarkName(string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException("VisualTeX OMML formulaId must be a UUID.");
        return BookmarkPrefix + parsed.ToString("N");
    }

    internal static bool TryGetFormulaId(Bookmark? bookmark, out string formulaId)
    {
        formulaId = string.Empty;
        if (bookmark is null) return false;
        string name;
        try { name = bookmark.Name ?? string.Empty; }
        catch { return false; }
        if (!name.StartsWith(BookmarkPrefix, StringComparison.Ordinal)) return false;
        var candidate = name.Substring(BookmarkPrefix.Length);
        if (!Guid.TryParseExact(candidate, "N", out var parsed)) return false;
        formulaId = parsed.ToString();
        return true;
    }

    // Word's native equation editor can rebuild an OMath without preserving the
    // collapsed VTOMML anchor at its canonical start. Recovery here is deliberately
    // narrower than the normal identity resolver: callers must supply the user's
    // explicitly selected COMPLETE OMath, and exactly one managed VTOMML identity
    // must be physically inside or immediately adjacent to that same equation.
    // Content fingerprints are intentionally not used to choose the candidate,
    // because the native edit is precisely what changed the content.
    internal static bool TryResolveExplicitSelectedEquationIdentityForRebind(
        Document document,
        Range selectionRange,
        Range equationRange,
        out Bookmark? resolvedBookmark,
        out FormulaMetadata? resolvedMetadata)
    {
        resolvedBookmark = null;
        resolvedMetadata = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exactEquationRange = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        var candidateIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (selectionRange.StoryType != equationRange.StoryType
                || selectionRange.Start != equationRange.Start
                || selectionRange.End != equationRange.End)
                return false;

            maths = equationRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            exactEquationRange = math.Range.Duplicate;
            if (exactEquationRange.StoryType != equationRange.StoryType
                || exactEquationRange.Start != equationRange.Start
                || exactEquationRange.End != equationRange.End)
                return false;

            // Range.Bookmarks is not reliable for collapsed bookmarks
            // exactly at an OMath end boundary on Word 2021. Snapshot only managed
            // FormulaIds from the document, then inspect each anchor coordinate and
            // accept it only when it is physically inside or immediately adjacent
            // to this exact selected equation. This is still local identity proof;
            // no content/fingerprint or nearest-formula heuristic participates.
            foreach (var formulaId in BookmarkedFormulaIds(document))
            {
                Release(bookmarkRange); bookmarkRange = null;
                Release(bookmark); bookmark = null;
                bookmark = FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start != bookmarkRange.End
                    || bookmarkRange.StoryType != equationRange.StoryType)
                    continue;
                var anchorPosition = bookmarkRange.Start;
                if (anchorPosition < equationRange.Start - 1
                    || anchorPosition > equationRange.End + 1)
                    continue;
                var metadata = TryRead(document, formulaId);
                if (metadata is null) continue;
                if (metadata.Numbered
                    && !NumberedFormulaIdentityMatchesEquationRange(
                        document,
                        formulaId,
                        metadata,
                        equationRange))
                    continue;
                candidateIds.Add(formulaId);
            }

            // Adjacent equations can legitimately have a VTOMML anchor one
            // character away. Ambiguity is therefore a hard stop, never a
            // nearest-anchor heuristic.
            if (candidateIds.Count != 1) return false;
            var selectedFormulaId = candidateIds.Single();
            resolvedMetadata = TryRead(document, selectedFormulaId);
            if (resolvedMetadata is null) return false;
            resolvedBookmark = FindByFormulaId(document, selectedFormulaId);
            if (resolvedBookmark is null)
            {
                resolvedMetadata = null;
                return false;
            }
            return true;
        }
        catch
        {
            Release(resolvedBookmark);
            resolvedBookmark = null;
            resolvedMetadata = null;
            return false;
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(exactEquationRange);
            Release(math);
            Release(maths);
        }
    }

    internal static Range RebindExplicitSelectedEquationIdentityForStructuralEdit(
        Document document,
        Range selectedEquationRange,
        FormulaMetadata refreshedMetadata)
    {
        Bookmark? localBookmark = null;
        Bookmark? reboundBookmark = null;
        FormulaMetadata? storedMetadata = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? exactEquationRange = null;
        Range? verifiedRange = null;
        try
        {
            maths = selectedEquationRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "The selected OMML identity repair requires exactly one complete Word equation.");
            math = maths[1];
            exactEquationRange = math.Range.Duplicate;
            if (selectedEquationRange.StoryType != exactEquationRange.StoryType
                || selectedEquationRange.Start != exactEquationRange.Start
                || selectedEquationRange.End != exactEquationRange.End)
                throw new InvalidDataException(
                    "The selected OMML identity repair range is not the complete OMath.");

            if (!TryResolveExplicitSelectedEquationIdentityForRebind(
                    document,
                    selectedEquationRange,
                    exactEquationRange,
                    out localBookmark,
                    out storedMetadata)
                || storedMetadata is null
                || !string.Equals(
                    storedMetadata.FormulaId,
                    refreshedMetadata.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The selected OMML identity could not be rebound uniquely in its local equation.");

            if (storedMetadata.Numbered != refreshedMetadata.Numbered
                || !string.Equals(
                    storedMetadata.DisplayMode,
                    refreshedMetadata.DisplayMode,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The selected OMML identity changed structural metadata before rebinding.");

            var currentWordOpenXml =
                WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    document,
                    exactEquationRange,
                    refreshedMetadata.FormulaId);
            if (!WordOmmlConverter.MatchesStoredOmmlFingerprint(
                    currentWordOpenXml,
                    refreshedMetadata.NativeOmmlFingerprint))
                throw new InvalidDataException(
                    "The selected OMML equation changed again after conversion preview capture.");

            if (refreshedMetadata.Numbered
                && !NumberedFormulaIdentityMatchesEquationRange(
                    document,
                    refreshedMetadata.FormulaId,
                    refreshedMetadata,
                    exactEquationRange))
                throw new InvalidDataException(
                    "The selected numbered OMML equation no longer belongs to its managed numbering structure.");

            reboundBookmark = Wrap(
                document,
                exactEquationRange,
                refreshedMetadata,
                replaceExisting: true);
            Save(document, refreshedMetadata);

            verifiedRange = GetEquationRangeVerifiedForStructuralEdit(
                document,
                refreshedMetadata.FormulaId,
                refreshedMetadata);
            var result = verifiedRange;
            verifiedRange = null;
            return result;
        }
        finally
        {
            Release(verifiedRange);
            Release(exactEquationRange);
            Release(math);
            Release(maths);
            Release(reboundBookmark);
            Release(localBookmark);
        }
    }

    internal static Bookmark? FindAtRange(Document document, Range selectionRange)
    {
        var direct = FindAtRangeFast(document, selectionRange);
        if (direct is not null) return direct;

        // Slow recovery is used only when the local bookmark probe failed. Work
        // from a stable FormulaId snapshot instead of iterating the live Word
        // Bookmarks collection: GetEquationRange may repair a drifted VTOMML by
        // deleting/recreating that bookmark, which invalidates both collection
        // indexes and the old Bookmark RCW.
        var formulaIds = BookmarkedFormulaIds(document);
        string? bestFormulaId = null;
        var bestLength = int.MaxValue;
        foreach (var formulaId in formulaIds)
        {
            Bookmark? bookmark = null;
            Range? equationRange = null;
            try
            {
                bookmark = FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                try { equationRange = GetEquationRange(bookmark); }
                catch { continue; }
                var containsCaret = selectionRange.Start == selectionRange.End
                    && selectionRange.Start >= equationRange.Start
                    && selectionRange.Start <= equationRange.End;
                var overlaps = selectionRange.Start < equationRange.End
                    && selectionRange.End > equationRange.Start;
                if (!containsCaret && !overlaps) continue;
                var length = Math.Max(0, equationRange.End - equationRange.Start);
                if (length >= bestLength) continue;
                bestFormulaId = formulaId;
                bestLength = length;
            }
            finally
            {
                Release(equationRange);
                Release(bookmark);
            }
        }
        return bestFormulaId is null
            ? null
            : FindByFormulaId(document, bestFormulaId);
    }

    internal static Bookmark? FindAtRangeFast(
        Document document,
        Range selectionRange)
    {
        Range? equationRange = null;
        try
        {
            return FindAtRangeFast(
                document,
                selectionRange,
                out _,
                out equationRange);
        }
        finally { Release(equationRange); }
    }

    // Editor discovery needs the metadata and the complete equation range that
    // prove the local bookmark is live. Return those values from the same local
    // probe instead of throwing the proof away and resolving the identical
    // CustomXMLPart/OMath several more times during ReadSelection.
    internal static Bookmark? FindAtRangeFast(
        Document document,
        Range selectionRange,
        out FormulaMetadata? resolvedMetadata,
        out Range? resolvedEquationRange)
    {
        resolvedMetadata = null;
        resolvedEquationRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? equationRange = null;
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        try
        {
            content = document.Content;
            object start = Math.Max(content.Start, selectionRange.Start - 2);
            object end = Math.Min(content.End, selectionRange.End + 2);
            probe = document.Range(ref start, ref end);
            bookmarks = probe.Bookmarks;

            // VisualTeX OMML bookmarks are collapsed anchors immediately before
            // the equation. Looking only inside this tiny local probe avoids the
            // previous O(bookmarks × all OMaths) scan on documents with dozens of
            // formulas.
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out var formulaId)) continue;
                    bookmarkRange = bookmark.Range;
                    if (bookmarkRange.Start >= selectionRange.Start - 2
                        && bookmarkRange.Start <= selectionRange.End + 1)
                    {
                        // A VTOMML bookmark by itself is not proof that a live
                        // equation still exists. Older conversion/delete paths can
                        // leave a collapsed VTOMML_<id> anchor and metadata behind
                        // after the OMath has been removed. Treating that orphan as
                        // the current selection makes ordinary insertion points fail
                        // before the editor even opens when RefreshForVisualTeX later
                        // tries to recover a non-existent equation. Validate this one
                        // local candidate through the same read-only identity resolver
                        // used by edit/conversion; stale anchors are simply ignored and
                        // the slow recovery below remains available for genuinely
                        // drifted formulas.
                        Range? verifiedEquation = null;
                        try
                        {
                            var metadata = TryRead(document, formulaId);
                            if (metadata is null) continue;
                            verifiedEquation = GetEquationRangeForCurrentRead(
                                document,
                                bookmark,
                                metadata);
                            var containsOrTouchesCaret = selectionRange.Start == selectionRange.End
                                && DistanceFromAnchorToEquation(
                                    selectionRange.Start,
                                    verifiedEquation) <= 1;
                            var overlapsEquation = selectionRange.Start < verifiedEquation.End
                                && selectionRange.End > verifiedEquation.Start;
                            if (!containsOrTouchesCaret && !overlapsEquation)
                                continue;
                            resolvedMetadata = metadata;
                            resolvedEquationRange = verifiedEquation;
                            verifiedEquation = null;
                        }
                        catch
                        {
                            continue;
                        }
                        finally { Release(verifiedEquation); }

                        var result = bookmark;
                        bookmark = null;
                        return result;
                    }
                }
                catch { }
                finally
                {
                    Release(bookmarkRange);
                    Release(bookmark);
                }
            }

            // Some Word selections expose only a caret inside the OMath. Resolve
            // that one local OMath and compare its start with the nearby anchors;
            // never call GetEquationRange() for every bookmark here.
            maths = selectionRange.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            equationRange = math.Range;
            Release(bookmarks);
            bookmarks = null;
            Release(probe);
            probe = null;
            object anchorProbeStart = Math.Max(content.Start, equationRange.Start - 8);
            object anchorProbeEnd = Math.Min(content.End, equationRange.Start + 3);
            probe = document.Range(ref anchorProbeStart, ref anchorProbeEnd);
            bookmarks = probe.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? bookmarkRange = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryGetFormulaId(bookmark, out var formulaId)) continue;
                    bookmarkRange = bookmark.Range;
                    var distance = equationRange.Start - bookmarkRange.Start;
                    if (distance < 0 || distance > 2) continue;
                    var metadata = TryRead(document, formulaId);
                    if (metadata is null) continue;
                    Range? verifiedEquation = null;
                    try
                    {
                        verifiedEquation = GetEquationRangeForCurrentRead(
                            document,
                            bookmark,
                            metadata);
                        var containsOrTouchesCaret = selectionRange.Start == selectionRange.End
                            && DistanceFromAnchorToEquation(
                                selectionRange.Start,
                                verifiedEquation) <= 1;
                        var overlapsEquation = selectionRange.Start < verifiedEquation.End
                            && selectionRange.End > verifiedEquation.Start;
                        if (!containsOrTouchesCaret && !overlapsEquation)
                            continue;
                        resolvedMetadata = metadata;
                        resolvedEquationRange = verifiedEquation;
                        verifiedEquation = null;
                    }
                    finally { Release(verifiedEquation); }
                    var result = bookmark;
                    bookmark = null;
                    return result;
                }
                catch { }
                finally
                {
                    Release(bookmarkRange);
                    Release(bookmark);
                }
            }
            return null;
        }
        catch { return null; }
        finally
        {
            Release(bookmarks);
            Release(probe);
            Release(content);
            Release(equationRange);
            Release(math);
            Release(maths);
        }
    }

    internal static IReadOnlyList<string> FindOrphanedAnchorsAtExactPosition(
        Document document,
        int position)
    {
        var orphaned = new List<string>();
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        Range? equationRange = null;
        try
        {
            content = document.Content;
            if (position < content.Start || position > content.End)
                return orphaned;
            var probeStart = Math.Max(content.Start, position - 1);
            var probeEnd = Math.Min(content.End, position + 1);
            probe = document.Range(probeStart, probeEnd);
            bookmarks = probe.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(equationRange); equationRange = null;
                Release(bookmarkRange); bookmarkRange = null;
                Release(bookmark); bookmark = bookmarks[index];
                if (!TryGetFormulaId(bookmark, out _)) continue;
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start != bookmarkRange.End
                    || bookmarkRange.Start != position)
                    continue;
                try
                {
                    // This preflight deliberately runs before the insertion's
                    // Custom UndoRecord starts. A recoverable drifted formula is
                    // still live and must keep its anchor; only an identity that
                    // cannot resolve to any unique OMath is safe to detach while
                    // the new object is inserted at the same collapsed position.
                    equationRange = GetEquationRange(bookmark);
                }
                catch (InvalidDataException)
                {
                    orphaned.Add(bookmark.Name);
                }
            }
            return orphaned
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
        finally
        {
            Release(equationRange);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(probe);
            Release(content);
        }
    }

    internal static int DeleteCapturedOrphanedAnchorsAtExactPosition(
        Document document,
        int position,
        IReadOnlyCollection<string> bookmarkNames)
    {
        if (bookmarkNames.Count == 0) return 0;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        var deleted = 0;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var name in bookmarkNames.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                Release(range); range = null;
                Release(bookmark); bookmark = null;
                if (!bookmarks.Exists(name)) continue;
                bookmark = bookmarks[name];
                range = bookmark.Range;
                // The preflight proved this exact collapsed anchor was orphaned.
                // Re-check its physical coordinate before deletion so a user edit
                // while the external editor was open cannot make us delete a moved
                // or newly repaired live identity. Deletion occurs inside the
                // insertion Custom UndoRecord; Word then restores the original
                // orphan on Undo instead of losing it through bookmark gravity.
                if (range.Start != range.End || range.Start != position)
                    continue;
                bookmark.Delete();
                deleted++;
            }
            if (deleted > 0)
                WordDoubleClickHook.TraceMessage(
                    $"orphan-omml-anchor-detached-for-insert position={position} count={deleted}");
            return deleted;
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Bookmark? FindByFormulaId(Document document, string formulaId)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.FindByFormulaId");
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = BookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            return bookmarks[name];
        }
        finally { Release(bookmarks); }
    }

    internal static IReadOnlyList<string> BookmarkedFormulaIds(Document document)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.BookmarkedFormulaIds");
        var formulaIds = new List<string>();
        Bookmarks? bookmarks = null;
        try
        {
            // Snapshot only stable formula ids while holding the live Word
            // Bookmarks collection. No OMath/Range lookup is allowed here.
            bookmarks = document.Bookmarks;
            var bookmarkCount = bookmarks.Count;
            for (var index = 1; index <= bookmarkCount; index++)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (TryGetFormulaId(bookmark, out var formulaId))
                        formulaIds.Add(formulaId);
                }
                finally { Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
        return formulaIds
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    internal static IReadOnlyDictionary<string, FormulaMetadata> CaptureMetadataSnapshot(
        Document document,
        IReadOnlyCollection<string> formulaIds) =>
        CaptureMetadataSnapshot(
            document,
            formulaIds,
            out _);

    internal static IReadOnlyDictionary<string, FormulaMetadata> CaptureMetadataSnapshot(
        Document document,
        IReadOnlyCollection<string> formulaIds,
        out IReadOnlyDictionary<string, string> uniquePartIds)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure(
            "WordOmmlFormulaStore.CaptureMetadataSnapshot");
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (formulaIds is null) throw new ArgumentNullException(nameof(formulaIds));

        var requested = new HashSet<string>(
            formulaIds.Where(id => !string.IsNullOrWhiteSpace(id)),
            StringComparer.OrdinalIgnoreCase);
        if (requested.Count == 0)
        {
            uniquePartIds = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            return new Dictionary<string, FormulaMetadata>(
                StringComparer.OrdinalIgnoreCase);
        }

        var resolved = new Dictionary<string, FormulaMetadata>(
            StringComparer.OrdinalIgnoreCase);
        var partIds = new Dictionary<string, HashSet<string>>(
            StringComparer.OrdinalIgnoreCase);
        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            for (var index = 1; index <= count; index++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var metadata)
                        || !requested.Contains(metadata.FormulaId))
                        continue;

                    var partId = ReadPartId(part);
                    if (!string.IsNullOrWhiteSpace(partId))
                    {
                        if (!partIds.TryGetValue(metadata.FormulaId, out var ids))
                        {
                            ids = new HashSet<string>(
                                StringComparer.OrdinalIgnoreCase);
                            partIds.Add(metadata.FormulaId, ids);
                        }
                        ids.Add(partId!);
                    }

                    if (resolved.TryGetValue(metadata.FormulaId, out var existing)
                        && CompareMetadataFreshness(metadata, existing) < 0)
                    {
                        RememberPart(document, part, metadata);
                        continue;
                    }
                    resolved[metadata.FormulaId] = CloneMetadata(metadata);
                    RememberPart(document, part, metadata);
                }
                finally { Release(part); }
            }
        }
        finally
        {
            Release(selected);
            Release(parts);
        }

        uniquePartIds = partIds
            .Where(entry => entry.Value.Count == 1)
            .ToDictionary(
                entry => entry.Key,
                entry => entry.Value.Single(),
                StringComparer.OrdinalIgnoreCase);
        return resolved;
    }

    internal static void DeleteVerifiedMetadataPart(
        Document document,
        string formulaId,
        string partId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new ArgumentException("FormulaId is required.", nameof(formulaId));
        if (string.IsNullOrWhiteSpace(partId))
        {
            Delete(document, formulaId);
            return;
        }

        object? parts = null;
        object? part = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            part = ((dynamic)parts).SelectByID(partId);
            if (part is null)
            {
                Delete(document, formulaId);
                return;
            }
            var partXml = (string?)((dynamic)part).XML;
            if (!TryDecodePartXml(partXml, out var metadata)
                || !string.Equals(
                    metadata.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase))
            {
                Delete(document, formulaId);
                return;
            }

            ((dynamic)part).Delete();
            ForgetPart(document, formulaId);
        }
        catch
        {
            Delete(document, formulaId);
        }
        finally
        {
            Release(part);
            Release(parts);
        }
    }

    internal static IReadOnlyList<string> StoredFormulaIds(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var formulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            for (var index = 1; index <= count; index++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var metadata)
                        || !Guid.TryParse(metadata.FormulaId, out var parsed))
                        continue;
                    RememberPart(document, part, metadata);
                    formulaIds.Add(parsed.ToString("D"));
                }
                catch
                {
                    // An unreadable custom XML part is not a safe logical formula
                    // identity. Leave it untouched and exclude it from copy/dedup
                    // decisions rather than guessing from malformed metadata.
                }
                finally { Release(part); }
            }
            return formulaIds.ToArray();
        }
        finally
        {
            Release(selected);
            Release(parts);
        }
    }

    internal static IReadOnlyList<string> FormulaIds(Document document)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.FormulaIds");
        var resolved = new List<(string Id, int Start)>();
        var owners = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var id in BookmarkedFormulaIds(document))
        {
            Bookmark? bookmark = null;
            Range? range = null;
            try
            {
                bookmark = FindByFormulaId(document, id)
                    ?? throw new InvalidDataException($"The OMML identity {id} disappeared during enumeration.");
                var metadata = TryRead(document, id)
                    ?? throw new InvalidDataException($"The OMML metadata for {id} is missing.");
                range = GetEquationRangeForCurrentRead(document, bookmark, metadata);
                var key = FormulaRangeKey(range.Start, range.End);
                if (owners.TryGetValue(key, out var otherId))
                    throw new InvalidDataException($"OMML identities {otherId} and {id} both claim the same physical formula.");
                owners.Add(key, id);
                resolved.Add((id, range.Start));
            }
            finally { Release(range); Release(bookmark); }
        }
        // Discovery uses the same strict row/anchor/content resolver as editing.
        // Never choose a nearest owner, rebind bookmarks or hide unresolved items
        // as a side effect of enumerating formulas.
        return resolved.OrderBy(item => item.Start).Select(item => item.Id).ToArray();
    }
    internal static FormulaMetadata? TryRead(Document document, Bookmark bookmark)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.TryRead");
        if (!TryGetFormulaId(bookmark, out var formulaId)) return null;
        return TryRead(document, formulaId);
    }

    internal static FormulaMetadata? TryRead(Document document, string formulaId)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.TryRead");
        object? part = null;
        try
        {
            // A cached metadata entry is only an acceleration hint. Word users can
            // undo, paste, or externally delete a CustomXMLPart while the document
            // RCW remains alive; returning the ConditionalWeakTable value directly
            // then revives an orphan VTOMML bookmark as a second logical formula.
            // FindPart validates the cached part id first and forgets stale entries.
            part = FindPart(document, formulaId);
            if (part is null) return null;

            var cache = MetadataCaches.GetValue(document, _ => new DocumentMetadataCache());
            lock (cache.Gate)
            {
                if (cache.Entries.TryGetValue(formulaId, out var cached))
                    return CloneMetadata(cached.Metadata);
            }

            var partXml = (string?)((dynamic)part).XML;
            if (!TryDecodePartXml(partXml, out var metadata)
                || !string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return null;
            RememberPart(document, part, metadata);
            return CloneMetadata(metadata);
        }
        catch
        {
            return null;
        }
        finally { Release(part); }
    }

    internal static void Save(Document document, FormulaMetadata metadata)
    {
        metadata.Validate();
        var xml = BuildPartXml(metadata);
        object? existing = null;
        object? parts = null;
        object? added = null;
        try
        {
            existing = FindPart(document, metadata.FormulaId);
            if (existing is not null)
            {
                try
                {
                    if ((bool)((dynamic)existing).LoadXML(xml))
                    {
                        // Successful in-place replacement cannot create a duplicate
                        // CustomXMLPart. The previous unconditional duplicate sweep
                        // parsed every VisualTeX metadata part in the document after
                        // each formula edit, making one update O(N) in total formulas.
                        // Keep duplicate cleanup on the exceptional add-then-delete
                        // fallback and explicit Delete(), where duplicates can really
                        // be introduced or must be purged.
                        RememberPart(document, existing, metadata);
                        return;
                    }
                }
                catch
                {
                    // Some Word builds reject in-place replacement after an
                    // OMath rebuild. Fall back to add-then-delete below.
                }
            }
            parts = ((dynamic)document).CustomXMLParts;
            added = ((dynamic)parts).Add(xml);
            if (added is null)
                throw new InvalidOperationException("Word did not create the VisualTeX OMML metadata part.");

            // Word may reject CustomXMLPart.LoadXML after an OMath rebuild.
            // Add the replacement first, then remove the old part so a failed
            // update never destroys the last valid metadata copy.
            if (existing is not null)
            {
                try { ((dynamic)existing).Delete(); } catch { }
            }
            RememberPart(document, added, metadata);
            RemoveDuplicateMetadataParts(
                document,
                metadata.FormulaId,
                ReadPartId(added));
        }
        finally
        {
            Release(added);
            Release(parts);
            Release(existing);
        }
    }

    internal static bool TrySaveKnownCachedPart(
        Document document,
        FormulaMetadata metadata)
    {
        metadata.Validate();
        if (!MetadataCaches.TryGetValue(document, out var cache)) return false;
        string? cachedPartId = null;
        lock (cache.Gate)
        {
            if (cache.Entries.TryGetValue(metadata.FormulaId, out var cached))
                cachedPartId = cached.PartId;
        }
        if (string.IsNullOrWhiteSpace(cachedPartId)) return false;

        object? parts = null;
        object? cachedPart = null;
        object? addedPart = null;
        object? staleProbe = null;
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var traceWatch = tracePerformance ? Stopwatch.StartNew() : null;
        long traceCheckpoint = 0;
        void Trace(string stage)
        {
            if (traceWatch is null) return;
            var elapsed = traceWatch.ElapsedMilliseconds;
            Console.WriteLine(
                $"    [perf] OmmlStore.SaveCached.{stage}: +{elapsed - traceCheckpoint}ms ({elapsed}ms)");
            traceCheckpoint = elapsed;
        }
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            Trace("parts-get");
            cachedPart = ((dynamic)parts).SelectByID(cachedPartId!);
            Trace("select-by-id");
            if (cachedPart is null) return false;
            var xml = BuildPartXml(metadata);
            Trace("build-xml");
            var loadedInPlace = false;
            try { loadedInPlace = (bool)((dynamic)cachedPart).LoadXML(xml); }
            catch { loadedInPlace = false; }
            Trace("load-xml");
            if (!loadedInPlace)
            {
                // Word can reject LoadXML immediately after rebuilding an OMath.
                // The old fallback returned to Save(), which added a replacement
                // and then scanned every VisualTeX CustomXMLPart to remove possible
                // duplicates. In a 1000-formula document that made one local edit
                // O(N) even though this transaction already owns the exact cached
                // part id. Replace that one known part atomically instead: add the
                // new value first, delete the known old part, and prove the old id
                // disappeared before accepting the cache update. If Word refuses
                // either mutation, delete the provisional part and fall back to the
                // full defensive Save() path without changing metadata semantics.
                try
                {
                    addedPart = ((dynamic)parts).Add(xml);
                    if (addedPart is null) return false;
                    ((dynamic)cachedPart).Delete();
                    staleProbe = ((dynamic)parts).SelectByID(cachedPartId!);
                    if (staleProbe is not null)
                    {
                        try { ((dynamic)addedPart).Delete(); } catch { }
                        return false;
                    }
                    RememberPart(document, addedPart, metadata);
                    Trace("replace-known-part");
                    return true;
                }
                catch
                {
                    if (addedPart is not null)
                    {
                        try { ((dynamic)addedPart).Delete(); } catch { }
                    }
                    return false;
                }
            }

            // This fast path is intentionally transaction-local. The caller has
            // already read and validated this exact cached part earlier in the same
            // Apply operation, and no user/Undo turn can interleave while Word is
            // synchronously processing the edit. Avoid re-reading and decoding the
            // same XML merely to prove the identity again. Any missing/stale part or
            // LoadXML rejection returns false and the caller falls back to Save(),
            // which performs the full defensive cache/namespace validation.
            RememberPart(document, cachedPart, metadata);
            Trace("remember");
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(staleProbe);
            Release(addedPart);
            Release(cachedPart);
            Release(parts);
        }
    }

    internal static void SaveNew(Document document, FormulaMetadata metadata)
    {
        metadata.Validate();
        object? parts = null;
        object? added = null;
        try
        {
            // InsertOmml always supplies a fresh FormulaId. Do not call FindPart
            // here: hydrating the namespace for a brand-new id walks every prior
            // VisualTeX CustomXMLPart and made the Nth OMML insertion O(N).
            parts = ((dynamic)document).CustomXMLParts;
            added = ((dynamic)parts).Add(BuildPartXml(metadata));
            if (added is null)
                throw new InvalidOperationException(
                    "Word did not create the VisualTeX OMML metadata part.");
            RememberPart(document, added, metadata);
        }
        finally
        {
            Release(added);
            Release(parts);
        }
    }

    internal static void SaveNewBatch(
        Document document,
        IReadOnlyList<FormulaMetadata> metadataItems)
    {
        if (metadataItems is null)
            throw new ArgumentNullException(nameof(metadataItems));
        if (metadataItems.Count == 0) return;
        foreach (var metadata in metadataItems)
            metadata.Validate();

        object? parts = null;
        object? added = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            foreach (var metadata in metadataItems)
            {
                added = ((dynamic)parts).Add(BuildPartXml(metadata));
                if (added is null)
                    throw new InvalidOperationException(
                        "Word did not create the VisualTeX OMML metadata part.");
                RememberPart(document, added, metadata);
                Release(added);
                added = null;
            }
        }
        finally
        {
            Release(added);
            Release(parts);
        }
    }

    internal static void Delete(Document document, string formulaId)
    {
        // A Word OMath rebuild can make CustomXMLPart.LoadXML fall back to
        // add-then-delete. On a few builds the old part survives that replacement,
        // so deleting only FindPart's freshest match leaves a stale duplicate that
        // TryRead later revives as a ghost formula. Explicit formula deletion must
        // remove every metadata part with this FormulaId, then clear the cache.
        ForgetPart(document, formulaId);
        RemoveDuplicateMetadataParts(
            document,
            formulaId,
            keepPartId: null);
        ForgetPart(document, formulaId);
    }

    internal static int RemoveOrphanedMetadataParts(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var liveFormulaIds = new HashSet<string>(
            BookmarkedFormulaIds(document),
            StringComparer.OrdinalIgnoreCase);
        object? parts = null;
        object? selected = null;
        var removed = 0;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            for (var index = count; index >= 1; index--)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var metadata)
                        || !Guid.TryParse(metadata.FormulaId, out var parsed))
                        continue;
                    var formulaId = parsed.ToString("D");
                    if (liveFormulaIds.Contains(formulaId)) continue;
                    ((dynamic)part).Delete();
                    removed++;
                }
                catch
                {
                    // An unreadable or non-deletable part is not safe to mutate.
                    // Leave it untouched rather than guessing from malformed data.
                }
                finally { Release(part); }
            }
        }
        finally
        {
            Release(selected);
            Release(parts);
            InvalidateDocumentCache(document);
        }
        return removed;
    }

    internal static Bookmark WrapFreshOmmlReplacement(
        Document document,
        Range equationRange,
        FormulaMetadata metadata)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? anchorRange = null;
        try
        {
            // ReplaceOmml calls this only immediately after a managed inline or
            // standalone-display OMath has been replaced in place and
            // ValidateInsertedOmml has proved the live Range. The old VTOMML
            // bookmark was explicitly deleted before insertion, so the generic
            // table-affinity probe and Exists/Delete sweep are both redundant.
            // Keeping the anchor on a duplicate of the live OMath preserves its
            // native story/cell affinity without asking Word to enumerate table
            // geometry (~60-75ms in a 100-OMML document).
            anchorRange = equationRange.Duplicate;
            anchorRange.Collapse(WdCollapseDirection.wdCollapseStart);
            bookmarks = document.Bookmarks;
            bookmark = bookmarks.Add(BookmarkName(metadata.FormulaId), anchorRange);
            var result = bookmark;
            bookmark = null;
            return result;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(anchorRange);
        }
    }

    internal static Bookmark Wrap(
        Document document,
        Range equationRange,
        FormulaMetadata metadata,
        bool replaceExisting = true)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? anchorRange = null;
        Range? preceding = null;
        Tables? equationTables = null;
        Table? equationTable = null;
        Range? equationTableRange = null;
        Table? managedNumberTable = null;
        Range? managedNumberTableRange = null;
        Rows? equationRows = null;
        Columns? equationColumns = null;
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Paragraphs? centerParagraphs = null;
        Paragraph? centerParagraph = null;
        Range? centerParagraphRange = null;
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var traceWatch = tracePerformance ? Stopwatch.StartNew() : null;
        long traceCheckpoint = 0;
        void Trace(string stage)
        {
            if (traceWatch is null) return;
            var elapsed = traceWatch.ElapsedMilliseconds;
            Console.WriteLine(
                $"    [perf] OmmlStore.Wrap.{stage}: +{elapsed - traceCheckpoint}ms ({elapsed}ms)");
            traceCheckpoint = elapsed;
        }
        try
        {
            var anchorPosition = equationRange.Start;
            if (anchorPosition > 0)
            {
                object precedingStart = anchorPosition - 1;
                object precedingEnd = anchorPosition;
                preceding = document.Range(ref precedingStart, ref precedingEnd);
                if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal))
                    anchorPosition--;
            }
            Trace("preceding-probe");

            // A collapsed main-story Range at a Word cell boundary is ambiguous:
            // Word can serialize it between </w:tc> and the next <w:tc>, which made
            // VTOMML_<FormulaId> drift to the row level after MathType→OMML batch
            // conversion. For the managed native 1x3 host, derive the bookmark from
            // cell (1,2)'s own Paragraph.Range so the container affinity is explicit.
            var anchoredInCenterCell = false;
            try
            {
                // The exact cell/table ownership checks below are structural;
                // Information(wdWithInTable) unnecessarily requests page layout.
                equationTables = equationRange.Tables;
                if (equationTables.Count > 0)
                {
                    equationTable = equationTables[1];
                    managedNumberTable = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        metadata.FormulaId);
                    if (managedNumberTable is not null)
                    {
                        equationTableRange = equationTable.Range;
                        managedNumberTableRange = managedNumberTable.Range;
                        var sameManagedTable =
                            equationTableRange.StoryType == managedNumberTableRange.StoryType
                            && equationTableRange.Start == managedNumberTableRange.Start
                            && equationTableRange.End == managedNumberTableRange.End;
                        if (sameManagedTable)
                        {
                            equationRows = equationTable.Rows;
                            equationColumns = equationTable.Columns;
                            if (equationRows.Count >= 1
                                && equationColumns.Count == 3
                                && WordEquationNumbering.TryGetManagedNumberTableRowIndex(
                                    equationTable,
                                    equationRange,
                                    expectedColumnIndex: 2,
                                    out var equationRowIndex))
                            {
                                centerCell = equationTable.Cell(equationRowIndex, 2);
                                centerCellRange = centerCell.Range;
                                if (equationRange.Start >= centerCellRange.Start
                                    && equationRange.End <= centerCellRange.End)
                                {
                                    centerParagraphs = centerCellRange.Paragraphs;
                                    if (centerParagraphs.Count == 1)
                                    {
                                        centerParagraph = centerParagraphs[1];
                                        centerParagraphRange = centerParagraph.Range.Duplicate;
                                        centerParagraphRange.Collapse(WdCollapseDirection.wdCollapseStart);
                                        anchorRange = centerParagraphRange;
                                        centerParagraphRange = null;
                                        anchoredInCenterCell = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                anchoredInCenterCell = false;
                Release(anchorRange);
                anchorRange = null;
            }

            Trace("table-affinity");
            if (!anchoredInCenterCell)
            {
                // Outside the managed table, preserve the equation Range's native
                // story affinity. Only the legacy inline-baseline sentinel needs an
                // explicit one-character shift before the equation.
                anchorRange = equationRange.Duplicate;
                anchorRange.Collapse(WdCollapseDirection.wdCollapseStart);
                if (anchorPosition != equationRange.Start)
                    anchorRange.SetRange(anchorPosition, anchorPosition);
            }
            bookmarks = document.Bookmarks;
            Trace("bookmarks-get");
            var name = BookmarkName(metadata.FormulaId);
            if (replaceExisting && bookmarks.Exists(name))
                bookmarks[name].Delete();
            Trace("existing-delete");
            bookmark = bookmarks.Add(name, anchorRange);
            Trace("bookmark-add");
            var result = bookmark;
            bookmark = null;
            return result;
        }
        finally
        {
            Release(centerParagraphRange);
            Release(centerParagraph);
            Release(centerParagraphs);
            Release(centerCellRange);
            Release(centerCell);
            Release(equationColumns);
            Release(equationRows);
            Release(managedNumberTableRange);
            Release(managedNumberTable);
            Release(equationTableRange);
            Release(equationTable);
            Release(equationTables);
            Release(preceding);
            Release(anchorRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static bool IsCanonicalAnchor(Bookmark? bookmark, Range equationRange)
    {
        if (bookmark is null) return false;
        Range? anchorRange = null;
        try
        {
            anchorRange = bookmark.Range;
            return anchorRange.Start == anchorRange.End
                && anchorRange.StoryType == equationRange.StoryType
                && (anchorRange.Start == equationRange.Start
                    || anchorRange.Start == equationRange.Start - 1);
        }
        finally { Release(anchorRange); }
    }

    private static bool NumberedFormulaIdentityMatchesEquationRange(
        Document document,
        string formulaId,
        FormulaMetadata? metadata,
        Range equationRange)
    {
        if (metadata is null
            || !metadata.Numbered
            || !string.Equals(
                metadata.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase))
            return true;

        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;

            // Current 1x3 native OMML keeps the semantic OMath in cell (1,2) and
            // the durable VTEqNum_<FormulaId> identity in cell (1,3). This is a
            // stronger physical identity than the semantic fingerprint, especially
            // when two equations have identical mathematical content. Accept the
            // adjacent OMath only when both ranges belong to the exact same managed
            // three-column table.
            Tables? equationTables = null;
            Tables? numberTables = null;
            Table? equationTable = null;
            Table? numberTable = null;
            Range? equationTableRange = null;
            Range? numberTableRange = null;
            try
            {
                equationTables = equationRange.Tables;
                numberTables = numberRange.Tables;
                // Exact same-table/cell/row containment below is the ownership
                // proof; no layout-dependent Information() preflight is needed.
                if (equationTables.Count > 0 && numberTables.Count > 0)
                {
                    if (equationTables.Count > 0 && numberTables.Count > 0)
                    {
                        equationTable = equationTables[1];
                        numberTable = numberTables[1];
                        if (equationTable.Columns.Count == 3
                            && equationTable.Rows.Count >= 1
                            && numberTable.Columns.Count == 3
                            && numberTable.Rows.Count >= 1
                            && WordEquationNumbering.TryGetManagedNumberTableRowIndex(
                                equationTable,
                                equationRange,
                                expectedColumnIndex: 2,
                                out var equationRowIndex)
                            && WordEquationNumbering.TryGetManagedNumberTableRowIndex(
                                numberTable,
                                numberRange,
                                expectedColumnIndex: 3,
                                out var numberRowIndex)
                            && equationRowIndex == numberRowIndex)
                        {
                            equationTableRange = equationTable.Range;
                            numberTableRange = numberTable.Range;
                            if (equationTableRange.Start == numberTableRange.Start
                                && equationTableRange.End == numberTableRange.End)
                                return true;
                        }
                    }
                }
            }
            finally
            {
                Release(numberTableRange);
                Release(equationTableRange);
                Release(numberTable);
                Release(equationTable);
                Release(numberTables);
                Release(equationTables);
            }

            // The older #(SEQ) host owns its number inside the OMath. Inspect it
            // only after the common row identity; use the same COM/XML capture
            // as editing, since Word can return an empty scratch serialization
            // immediately after a structural replacement.
            if (WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                    WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                        document, equationRange, formulaId)))
                return numberRange.StoryType == WdStoryType.wdMainTextStory
                    && numberRange.Start >= equationRange.Start
                    && numberRange.End <= equationRange.End;

            // Shape-era hosts have no same-container VTEqNum identity and remain
            // migration input only; fall back to fingerprint recovery for them.
            return false;
        }
        finally
        {
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    internal static Range GetEquationRange(Bookmark bookmark)
    {
        Range? anchor = null;
        Document? document = null;
        try
        {
            anchor = bookmark.Range;
            document = anchor.Document;
            if (!TryGetFormulaId(bookmark, out var formulaId))
                throw new InvalidDataException("The bookmark is not a VisualTeX OMML identity.");
            var metadata = TryRead(document, formulaId)
                ?? throw new InvalidDataException($"The OMML metadata for {formulaId} is missing.");
            return ResolveEquationIdentity(document, bookmark, formulaId, metadata);
        }
        finally
        {
            Release(document);
            Release(anchor);
        }
    }

    internal static Range GetEquationRangeVerifiedForStructuralEdit(
        Document document,
        string formulaId,
        FormulaMetadata metadata)
    {
        var bookmark = FindByFormulaId(document, formulaId)
            ?? throw new InvalidDataException($"The VisualTeX OMML bookmark for {formulaId} is missing.");
        try { return ResolveEquationIdentity(document, bookmark, formulaId, metadata); }
        finally { Release(bookmark); }
    }

    // Opening an equation captures the user's current Word content. Only an
    // intact physical anchor and its own numbered row permit that capture;
    // drift recovery still requires the previously stored content fingerprint.
    // This never writes metadata. A later commit verifies the new session
    // fingerprint through GetEquationRangeVerifiedForStructuralEdit.
    internal static Range GetEquationRangeForCurrentRead(
        Document document,
        Bookmark bookmark,
        FormulaMetadata stored)
    {
        if (!TryGetFormulaId(bookmark, out var formulaId)
            || !string.Equals(formulaId, stored.FormulaId, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The OMML read request does not match its bookmark identity.");
        return ResolveEquationIdentity(document, bookmark, formulaId, stored,
            captureCurrentContent: true);
    }

    // Read/open/edit/conversion must agree on one physical formula. Resolution is
    // read-only: repairs belong to the caller's mutation transaction, never to a
    // discovery operation or a preflight check before its undo checkpoint.
    private static Range ResolveEquationIdentity(
        Document document,
        Bookmark bookmark,
        string formulaId,
        FormulaMetadata metadata,
        bool captureCurrentContent = false)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.ResolveEquationIdentity");
        Range? bookmarkRange = null;
        Range? candidate = null;
        Range? content = null;
        try
        {
            bookmarkRange = bookmark.Range;
            content = document.Content;
            var expectedFingerprint = metadata.NativeOmmlFingerprint;
            bool ContentMatches(Range range) => string.IsNullOrWhiteSpace(expectedFingerprint)
                || string.Equals(GetEquationFingerprint(range, expectedFingerprint!), expectedFingerprint, StringComparison.OrdinalIgnoreCase);

            // Read the exact canonical bookmark before broadening any range.
            // Word 2021 can omit or clip table OMaths in a probe spanning body /
            // table boundaries. Opening, applying and converting must not choose
            // a neighbouring equation merely because that widened collection did.
            candidate = TryGetExactAnchorEquationRange(bookmark)
                ?? FindAdjacentEquationRangeNearAnchor(document, content, bookmarkRange.Start);
            if (candidate is not null && IsCanonicalAnchor(bookmark, candidate)
                && ((NumberedFormulaIdentityMatchesEquationRange(document, formulaId, metadata, candidate)
                        && (captureCurrentContent || ContentMatches(candidate)))
                    || (metadata.Numbered && metadata.DisplayMode == "block"
                        && bookmarkRange.Start == candidate.Start
                        && !string.IsNullOrWhiteSpace(expectedFingerprint)
                        && ContentMatches(candidate)
                        && IsPendingStandaloneNumberHost(document, formulaId, candidate))))
            {
                var result = candidate;
                candidate = null;
                return result;
            }
            Release(candidate);
            candidate = null;

            // The same-row numeric identity distinguishes identical equations in
            // different rows. It cannot override a mismatch in captured content.
            if (metadata.Numbered && metadata.DisplayMode == "block")
            {
                candidate = FindNumberedEquationRangeByNumberIdentity(document, formulaId);
                if (candidate is not null)
                {
                    if (!ContentMatches(candidate))
                        throw new InvalidDataException($"The numbered OMML row for {formulaId} no longer matches its captured formula content.");
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }

            // Missing legacy fingerprints permit only a verified local/row host;
            // without content evidence there is no safe drift recovery by position.
            if (!string.IsNullOrWhiteSpace(expectedFingerprint))
                candidate = FindUniqueEquationRangeByFingerprint(document, expectedFingerprint!,
                    requireDisplay: !metadata.Numbered && metadata.DisplayMode == "block");
            if (candidate is null)
                throw new InvalidDataException($"The VisualTeX OMML identity for {formulaId} drifted and its equation could not be recovered uniquely.");
            var recovered = candidate;
            candidate = null;
            return recovered;
        }
        finally
        {
            Release(candidate);
            Release(content);
            Release(bookmarkRange);
        }
    }

    private static Range? TryGetExactAnchorEquationRange(Bookmark bookmark)
    {
        Range? anchor = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? range = null;
        try
        {
            anchor = bookmark.Range;
            if (anchor.Start != anchor.End) return null;
            maths = anchor.OMaths;
            if (maths.Count != 1) return null;
            math = maths[1];
            range = math.Range.Duplicate;
            if (!IsCanonicalAnchor(bookmark, range)) return null;
            var result = range;
            range = null;
            return result;
        }
        finally
        {
            Release(range); Release(math); Release(maths); Release(anchor);
        }
    }

    private static bool IsPendingStandaloneNumberHost(Document document, string formulaId, Range equation)
    {
        // Numbered metadata records intent before the common numbering pass has
        // built its host. That state still has an exact anchor and a required
        // matching fingerprint; it must not be confused with a displaced row.
        Bookmarks? bookmarks = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        Range? prefix = null;
        Range? suffix = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        Frames? frames = null;
        ContentControls? controls = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(WordEquationNumbering.NativeNumberBookmarkName(formulaId))
                || WordEquationNumbering.RangeIsWhollyWithinTable(equation)) return false;
            paragraphs = equation.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            maths = owner.OMaths;
            shapes = owner.InlineShapes;
            fields = owner.Fields;
            frames = owner.Frames;
            controls = owner.ContentControls;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0
                || frames.Count != 0 || controls.Count != 0) return false;
            prefix = document.Range(owner.Start, equation.Start);
            suffix = document.Range(equation.End, owner.End);
            return string.IsNullOrEmpty(prefix.Text) && suffix.Text == "\r";
        }
        finally
        {
            Release(controls); Release(frames); Release(fields); Release(shapes); Release(maths);
            Release(suffix); Release(prefix); Release(owner); Release(paragraph); Release(paragraphs); Release(bookmarks);
        }
    }

    internal static Range? FindNumberedEquationRangeByNumberIdentity(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        OMaths? localMaths = null;
        OMath? localMath = null;
        Range? localMathRange = null;
        OMaths? documentMaths = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return null;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            if (numberRange.StoryType != WdStoryType.wdMainTextStory)
                return null;

            // VTEqNum owns column 3 of one managed row. The semantic Display
            // OMath belongs to column 2 of that same row, for both 1x3 and Nx3.
            // Editing, conversion and numbering share this physical identity;
            // identical formula contents must never select a sibling row.
            Tables? numberTables = null;
            Table? numberTable = null;
            Cell? formulaCell = null;
            Range? formulaCellRange = null;
            OMaths? tableMaths = null;
            OMath? tableMath = null;
            Range? tableMathRange = null;
            try
            {
                numberTables = numberRange.Tables;
                if (numberTables.Count > 0)
                {
                    if (numberTables.Count > 0)
                    {
                        numberTable = numberTables[1];
                        if (numberTable.Rows.Count >= 1
                            && numberTable.Columns.Count == 3
                            && WordEquationNumbering.TryGetManagedNumberTableRowIndex(
                                numberTable,
                                numberRange,
                                expectedColumnIndex: 3,
                                out var numberRowIndex))
                        {
                            formulaCell = numberTable.Cell(numberRowIndex, 2);
                            formulaCellRange = formulaCell.Range;
                            tableMaths = formulaCellRange.OMaths;
                            if (tableMaths.Count == 1
                                && formulaCellRange.Fields.Count == 0
                                && formulaCellRange.InlineShapes.Count == 0)
                            {
                                tableMath = tableMaths[1];
                                tableMathRange = tableMath.Range.Duplicate;
                                if (tableMath.Type == WdOMathType.wdOMathDisplay
                                    && WordEquationNumbering.HasReusableNumberedNativeOmmlDirectTableHost(
                                        document, tableMathRange, formulaId))
                                {
                                    var tableResult = tableMathRange;
                                    tableMathRange = null;
                                    return tableResult;
                                }
                            }
                        }
                    }
                }
            }
            catch { }
            finally
            {
                Release(tableMathRange);
                Release(tableMath);
                Release(tableMaths);
                Release(formulaCellRange);
                Release(formulaCell);
                Release(numberTable);
                Release(numberTables);
            }

            // Fast path: Word normally exposes the containing professional OMath
            // directly from the number bookmark range.
            localMaths = numberRange.OMaths;
            if (localMaths.Count == 1)
            {
                localMath = localMaths[1];
                localMathRange = localMath.Range.Duplicate;
                if (numberRange.Start >= localMathRange.Start
                    && numberRange.End <= localMathRange.End
                    && WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        localMathRange.WordOpenXML))
                {
                    var result = localMathRange;
                    localMathRange = null;
                    return result;
                }
            }

            // Some Word builds expose OMaths.Count=0 for a bookmark that begins
            // exactly on the mathematical field result. Fall back to document OMath
            // containment, but still require the direct VisualTeXEquation SEQ host.
            documentMaths = document.OMaths;
            for (var index = 1; index <= documentMaths.Count; index++)
            {
                OMath? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = documentMaths[index];
                    candidateRange = candidate.Range.Duplicate;
                    if (numberRange.Start < candidateRange.Start
                        || numberRange.End > candidateRange.End)
                        continue;
                    if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                            candidateRange.WordOpenXML))
                        continue;
                    var result = candidateRange;
                    candidateRange = null;
                    return result;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }
            return null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(documentMaths);
            Release(localMathRange);
            Release(localMath);
            Release(localMaths);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static Range? FindAdjacentEquationRangeNearAnchor(
        Document document,
        Range content,
        int anchor)
    {
        var span = 64;
        while (span <= 65_536)
        {
            Range? probe = null;
            OMaths? maths = null;
            Range? candidate = null;
            try
            {
                object probeStart = Math.Max(content.Start, anchor - span);
                object probeEnd = Math.Min(content.End, anchor + span);
                if ((int)probeEnd <= (int)probeStart && content.End > (int)probeStart)
                    probeEnd = (int)probeStart + 1;
                probe = document.Range(ref probeStart, ref probeEnd);
                maths = probe.OMaths;
                candidate = FindAdjacentEquationRange(maths, anchor);
                if (candidate is null) return null;

                // A range touching either probe boundary may have been clipped
                // by Word. This matters when Word moves the collapsed VisualTeX
                // bookmark to the end of an equation after native OMath editing.
                var completeAtStart = candidate.Start > probe.Start
                    || probe.Start <= content.Start;
                var completeAtEnd = candidate.End < probe.End
                    || probe.End >= content.End;
                if (completeAtStart && completeAtEnd)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            finally
            {
                Release(candidate);
                Release(maths);
                Release(probe);
            }
            span *= 4;
        }
        return null;
    }

    private static int DistanceFromAnchorToEquation(int anchor, Range range)
    {
        if (anchor < range.Start) return range.Start - anchor;
        if (anchor > range.End) return anchor - range.End;
        return 0;
    }

    private static int AnchorRelationPriority(int anchor, Range range)
    {
        // Prefer a formula that actually contains the anchor. At an exact
        // boundary, prefer the equation after the bookmark because that is the
        // canonical VisualTeX layout; a preceding equation is the recovery path
        // for bookmarks moved by Word-native editing.
        if (anchor > range.Start && anchor < range.End) return 0;
        return range.Start >= anchor ? 1 : 2;
    }

    private static Range? FindAdjacentEquationRange(OMaths maths, int anchor)
    {
        Range? bestRange = null;
        var bestDistance = int.MaxValue;
        var bestPriority = int.MaxValue;
        for (var index = 1; index <= maths.Count; index++)
        {
            OMath? math = null;
            Range? range = null;
            try
            {
                math = maths[index];
                range = math.Range;
                var distance = DistanceFromAnchorToEquation(anchor, range);
                if (distance > 8) continue;
                var priority = AnchorRelationPriority(anchor, range);
                if (distance > bestDistance
                    || (distance == bestDistance && priority >= bestPriority))
                    continue;
                Release(bestRange);
                bestRange = TrimToNativeMath(range);
                bestDistance = distance;
                bestPriority = priority;
            }
            finally
            {
                Release(range);
                Release(math);
            }
        }
        return bestRange;
    }

    private static string FormulaRangeKey(int start, int end) =>
        start + ":" + end;

    private static string GetEquationFingerprint(
        Range equationRange,
        string expectedFingerprint,
        IDictionary<string, string>? cache = null)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.GetEquationFingerprint");
        var key = OmmlFingerprintFormat.GetVersion(expectedFingerprint) + ":"
            + FormulaRangeKey(equationRange.Start, equationRange.End);
        if (cache is not null && cache.TryGetValue(key, out var cached))
            return cached;
        Document? document = null;
        string fingerprint;
        try
        {
            document = equationRange.Document;
            fingerprint = WordOmmlConverter.ComputeOmmlFingerprintForExpectedVersion(
                WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    document, equationRange, string.Empty), expectedFingerprint);
        }
        finally { Release(document); }
        if (cache is not null) cache[key] = fingerprint;
        return fingerprint;
    }

    private static Range? FindUniqueEquationRangeByFingerprint(
        Document document,
        string expectedFingerprint,
        bool requireDisplay)
    {
        using var operationMetric = VisualTeX.WindowsOffice.Contracts.WordOperationMetrics.Measure("WordOmmlFormulaStore.FindUniqueEquationRangeByFingerprint");
        OMaths? maths = null;
        Range? match = null;
        var ambiguous = false;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                Range? trimmed = null;
                try
                {
                    math = maths[index];
                    if (requireDisplay && math.Type != WdOMathType.wdOMathDisplay)
                        continue;
                    range = math.Range;
                    trimmed = TrimToNativeMath(range);
                    string fingerprint;
                    try { fingerprint = GetEquationFingerprint(trimmed, expectedFingerprint); }
                    catch { continue; }
                    if (!string.Equals(
                            fingerprint,
                            expectedFingerprint,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (match is not null)
                    {
                        ambiguous = true;
                        break;
                    }
                    match = trimmed;
                    trimmed = null;
                }
                finally
                {
                    Release(trimmed);
                    Release(range);
                    Release(math);
                }
            }
            if (!ambiguous) return match;
            Release(match);
            match = null;
            return null;
        }
        finally { Release(maths); }
    }

    private static Range TrimToNativeMath(Range source)
    {
        Range? result = null;
        Range? probe = null;
        try
        {
            result = source.Duplicate;
            while (result.Start < result.End)
            {
                probe = result.Duplicate;
                probe.SetRange(result.Start, Math.Min(result.Start + 1, result.End));
                if (RangeContainsNativeMath(probe)) break;
                result.Start++;
                Release(probe);
                probe = null;
            }
            while (result.End > result.Start)
            {
                probe = result.Duplicate;
                probe.SetRange(Math.Max(result.Start, result.End - 1), result.End);
                if (RangeContainsNativeMath(probe)) break;
                result.End--;
                Release(probe);
                probe = null;
            }
            if (result.Start >= result.End)
                throw new InvalidDataException(
                    "Word returned an OMML range without native math content.");
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(probe);
            Release(result);
        }
    }

    private static bool RangeContainsNativeMath(Range range)
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

    internal static float EstimateHeightPoints(Bookmark bookmark)
    {
        Range? equationRange = null;
        try
        {
            equationRange = GetEquationRange(bookmark);
            return EstimateHeightPoints(equationRange);
        }
        finally { Release(equationRange); }
    }

    internal static float EstimateHeightPoints(Range equationRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = equationRange.Font;
            var size = 11f;
            try { size = font.Size; } catch { }
            if (float.IsNaN(size) || float.IsInfinity(size) || size <= 0 || size > 256)
                size = 11f;
            return Math.Max(11f, size * 1.5f);
        }
        finally { Release(font); }
    }

    internal static string BuildPartXml(FormulaMetadata metadata)
    {
        metadata.Validate();
        var encoded = FormulaMetadataCodec.Encode(metadata);
        return new XDocument(
            new XElement(
                VisualTeXNamespace + "formula",
                new XAttribute("formulaId", metadata.FormulaId),
                new XElement(VisualTeXNamespace + "metadata", encoded)))
            .ToString(SaveOptions.DisableFormatting);
    }

    internal static bool TryDecodePartXml(string? xml, out FormulaMetadata metadata)
    {
        metadata = null!;
        if (string.IsNullOrWhiteSpace(xml)) return false;
        try
        {
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
            };
            using var stringReader = new StringReader(xml);
            using var reader = XmlReader.Create(stringReader, settings);
            var document = XDocument.Load(reader, LoadOptions.None);
            var root = document.Root;
            if (root?.Name != VisualTeXNamespace + "formula") return false;
            var formulaId = (string?)root.Attribute("formulaId");
            var encoded = root.Element(VisualTeXNamespace + "metadata")?.Value;
            var decoded = FormulaMetadataCodec.Decode(encoded);
            if (decoded is null
                || !string.Equals(decoded.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return false;
            decoded.Validate();
            metadata = decoded;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static FormulaMetadata CloneMetadata(FormulaMetadata metadata) =>
        FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata))
        ?? throw new InvalidDataException("Unable to clone VisualTeX OMML metadata.");

    private static void RememberPart(
        Document document,
        object? part,
        FormulaMetadata metadata)
    {
        if (part is null) return;
        var partId = ReadPartId(part);
        if (string.IsNullOrWhiteSpace(partId)) return;
        var cache = MetadataCaches.GetValue(document, _ => new DocumentMetadataCache());
        lock (cache.Gate)
        {
            if (cache.Entries.TryGetValue(metadata.FormulaId, out var existing)
                && CompareMetadataFreshness(metadata, existing.Metadata) < 0)
                return;
            cache.Entries[metadata.FormulaId] = new CachedMetadataPart
            {
                PartId = partId!,
                Metadata = CloneMetadata(metadata),
            };
        }
    }

    private static string? ReadPartId(object? part)
    {
        if (part is null) return null;
        try { return (string?)((dynamic)part).Id; }
        catch { return null; }
    }

    private static int CompareMetadataFreshness(
        FormulaMetadata left,
        FormulaMetadata right)
    {
        static DateTimeOffset ParseTimestamp(string? value)
        {
            return DateTimeOffset.TryParse(
                    value,
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.RoundtripKind,
                    out var parsed)
                ? parsed
                : DateTimeOffset.MinValue;
        }

        var updated = ParseTimestamp(left.UpdatedAt)
            .CompareTo(ParseTimestamp(right.UpdatedAt));
        if (updated != 0) return updated;
        return ParseTimestamp(left.CreatedAt)
            .CompareTo(ParseTimestamp(right.CreatedAt));
    }

    private static void RemoveDuplicateMetadataParts(
        Document document,
        string formulaId,
        string? keepPartId)
    {
        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            for (var index = count; index >= 1; index--)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partId = ReadPartId(part);
                    if (!string.IsNullOrWhiteSpace(keepPartId)
                        && string.Equals(
                            partId,
                            keepPartId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var candidate)
                        || !string.Equals(
                            candidate.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    try { ((dynamic)part).Delete(); } catch { }
                }
                finally { Release(part); }
            }
        }
        finally
        {
            Release(selected);
            Release(parts);
        }
    }

    internal static void InvalidateDocumentCache(Document document)
    {
        if (document is null) return;
        MetadataCaches.Remove(document);
    }

    private static void ForgetPart(Document document, string formulaId)
    {
        if (!MetadataCaches.TryGetValue(document, out var cache)) return;
        lock (cache.Gate) cache.Entries.Remove(formulaId);
    }

    private static object? FindPart(Document document, string formulaId)
    {
        var cache = MetadataCaches.GetValue(document, _ => new DocumentMetadataCache());
        string? cachedPartId = null;
        lock (cache.Gate)
        {
            if (cache.Entries.TryGetValue(formulaId, out var cached))
                cachedPartId = cached.PartId;
        }

        object? parts = null;
        object? selected = null;
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            if (!string.IsNullOrWhiteSpace(cachedPartId))
            {
                object? cachedPart = null;
                try
                {
                    cachedPart = ((dynamic)parts).SelectByID(cachedPartId!);
                    if (cachedPart is not null)
                    {
                        var partXml = (string?)((dynamic)cachedPart).XML;
                        if (TryDecodePartXml(partXml, out var cachedMetadata)
                            && string.Equals(
                                cachedMetadata.FormulaId,
                                formulaId,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            var result = cachedPart;
                            cachedPart = null;
                            return result;
                        }
                    }
                }
                catch { }
                finally { Release(cachedPart); }
                ForgetPart(document, formulaId);
            }

            selected = ((dynamic)parts).SelectByNamespace(NamespaceUri);
            var count = (int)((dynamic)selected).Count;
            object? matched = null;
            FormulaMetadata? matchedMetadata = null;
            for (var index = 1; index <= count; index++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)selected)[index];
                    var partXml = (string?)((dynamic)part).XML;
                    if (!TryDecodePartXml(partXml, out var metadata)) continue;
                    RememberPart(document, part, metadata);
                    if (!string.Equals(
                            metadata.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (matched is null
                        || matchedMetadata is null
                        || CompareMetadataFreshness(metadata, matchedMetadata) >= 0)
                    {
                        Release(matched);
                        matched = part;
                        matchedMetadata = metadata;
                        part = null;
                    }
                }
                finally { Release(part); }
            }
            lock (cache.Gate) cache.Hydrated = true;
            if (matched is not null && matchedMetadata is not null)
                RememberPart(document, matched, matchedMetadata);
            return matched;
        }
        finally
        {
            Release(selected);
            Release(parts);
        }
    }

    internal static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

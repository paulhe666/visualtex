using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationNumberReferences;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordVisibleEquationNumbers;

namespace VisualTeX.WordVsto;

/// <summary>
/// Repairing numbering artifacts: external captions, orphaned artifacts and bookmarks, shared captions.
/// </summary>
internal static class WordNumberingCleanup
{
    internal static void RebindExternalCaptionToNativeNumberParagraph(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Range? captionRange = null;
        Range? numberRange = null;
        Paragraphs? numberParagraphs = null;
        Paragraph? numberParagraph = null;
        Range? numberParagraphRange = null;
        Fields? paragraphFields = null;
        Field? field = null;
        Range? code = null;
        Field? matchedSequence = null;
        Range? reboundRange = null;
        Bookmark? rebound = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName =
                NativeCaptionBookmarkName(
                    formulaId);
            var numberName =
                NativeNumberBookmarkName(
                    formulaId);
            if (!bookmarks.Exists(captionName)
                || !bookmarks.Exists(numberName))
                return;

            captionBookmark =
                bookmarks[captionName];
            numberBookmark =
                bookmarks[numberName];
            captionRange =
                captionBookmark.Range.Duplicate;
            numberRange =
                numberBookmark.Range.Duplicate;

            // Native #(SEQ) OMML keeps both aliases inside mathematics and is
            // handled by TryRemoveNativeHashSequenceCaptionBookmarks above.
            if (WordFormulaHost.IsWhollyInsideMath(
                    numberRange))
                return;

            numberParagraphs =
                numberRange.Paragraphs;
            if (numberParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The external equation number no longer belongs to exactly one caption paragraph.");
            numberParagraph =
                numberParagraphs[1];
            numberParagraphRange =
                numberParagraph.Range.Duplicate;

            if (numberRange.StoryType !=
                    numberParagraphRange.StoryType
                || numberRange.Start <
                    numberParagraphRange.Start
                || numberRange.End >
                    numberParagraphRange.End)
                throw new InvalidDataException(
                    "The external equation number escaped its caption paragraph.");

            // Prove this is the generated caption paragraph before rebinding any
            // bookmark. It must own exactly one VisualTeX equation SEQ field and
            // that field's result must contain the stable VTEqNum target.
            paragraphFields =
                numberParagraphRange.Fields;
            var sequenceCount = 0;
            for (var index = 1;
                 index <= paragraphFields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field =
                    paragraphFields[index];
                code =
                    field.Code.Duplicate;
                if (!IsVisualTeXSequenceFieldCode(
                        code.Text))
                    continue;

                Range? result = null;
                try
                {
                    result =
                        field.Result.Duplicate;
                    // A chapter number ("2." written before the field) puts
                    // the bookmark's start ahead of the SEQ result; it still
                    // has to end inside that result.
                    if (result.StoryType !=
                            numberRange.StoryType
                        || result.End <=
                            numberRange.Start
                        || result.End <
                            numberRange.End)
                        throw new InvalidDataException(
                            "The external equation number bookmark is not owned by its caption SEQ result.");
                }
                finally
                {
                    Release(result);
                }

                sequenceCount++;
                Release(matchedSequence);
                matchedSequence =
                    field;
                field = null;
            }

            if (sequenceCount != 1
                || matchedSequence is null)
                throw new InvalidDataException(
                    "The external equation number paragraph does not own exactly one VisualTeX SEQ field.");

            // A generated caption paragraph never owns mathematical hosts or OLEs.
            // Reject instead of trimming if the paragraph itself is contaminated;
            // only the stale outer bookmark is allowed to overlap user content.
            OMaths? maths = null;
            InlineShapes? shapes = null;
            Tables? tables = null;
            ContentControls? controls = null;
            try
            {
                maths =
                    numberParagraphRange.OMaths;
                shapes =
                    numberParagraphRange.InlineShapes;
                tables =
                    numberParagraphRange.Tables;
                controls =
                    numberParagraphRange.ContentControls;
                if (maths.Count != 0
                    || shapes.Count != 0
                    || tables.Count != 0
                    || controls.Count != 0)
                    throw new InvalidDataException(
                        "The verified external equation caption paragraph contains unowned document content.");
            }
            finally
            {
                Release(controls);
                Release(tables);
                Release(shapes);
                Release(maths);
            }

            var alreadyCanonical =
                captionRange.StoryType ==
                    numberParagraphRange.StoryType
                && captionRange.Start ==
                    numberParagraphRange.Start
                && captionRange.End ==
                    numberParagraphRange.End;
            if (alreadyCanonical)
                return;

            // Detach only the stale bookmark identity. No document text is
            // deleted here. Then bind VTEqCap_<id> to the exact paragraph proven
            // by VTEqNum_<id> + its unique SEQ field.
            captionBookmark.Delete();
            Release(captionBookmark);
            captionBookmark = null;

            reboundRange =
                numberParagraphRange.Duplicate;
            rebound =
                bookmarks.Add(
                    captionName,
                    reboundRange);

            Range? verified = null;
            try
            {
                verified =
                    rebound.Range.Duplicate;
                if (verified.StoryType !=
                        numberParagraphRange.StoryType
                    || verified.Start !=
                        numberParagraphRange.Start
                    || verified.End !=
                        numberParagraphRange.End)
                    throw new InvalidDataException(
                        "Word re-expanded the repaired external equation caption bookmark.");
            }
            finally
            {
                Release(verified);
            }
        }
        finally
        {
            Release(rebound);
            Release(reboundRange);
            Release(matchedSequence);
            Release(code);
            Release(field);
            Release(paragraphFields);
            Release(numberParagraphRange);
            Release(numberParagraph);
            Release(numberParagraphs);
            Release(numberRange);
            Release(captionRange);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static void DeleteBookmarkedRangeAndContainingFrame(
        Document document,
        string name,
        bool preserveParagraphSeparator = false)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Frames? frames = null;
        Frame? frame = null;
        var originalParagraphStart = -1;
        var originalParagraphEnd = -1;
        var originalOwnedWholeParagraph = false;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            Paragraphs? originalParagraphs = null;
            Paragraph? originalParagraph = null;
            Range? originalParagraphRange = null;
            try
            {
                originalParagraphs = range.Paragraphs;
                if (originalParagraphs.Count == 1)
                {
                    originalParagraph = originalParagraphs[1];
                    originalParagraphRange = originalParagraph.Range;
                    originalParagraphStart = originalParagraphRange.Start;
                    originalParagraphEnd = originalParagraphRange.End;
                    originalOwnedWholeParagraph = range.Start == originalParagraphStart
                        && range.End == originalParagraphEnd
                        && (range.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal);
                }
            }
            finally
            {
                Release(originalParagraphRange);
                Release(originalParagraph);
                Release(originalParagraphs);
            }
            OMaths? bookmarkMaths = null;
            try
            {
                bookmarkMaths = range.OMaths;
                if (WordFormulaHost.IsWhollyInsideMath(range))
                {
                    // VTEqCap can be an alias inside a native #(SEQ) result. A
                    // bookmark may transiently survive FormattedText replacement;
                    // deleting its Range would then remove part/all of the new
                    // equation. Mathematical aliases are identity only and are
                    // always removed without deleting content.
                    bookmark.Delete();
                    return;
                }
                for (var index = 1; index <= bookmarkMaths.Count; index++)
                {
                    OMath? math = null;
                    Range? mathRange = null;
                    try
                    {
                        math = bookmarkMaths[index]; mathRange = math.Range;
                        if (mathRange.StoryType == range.StoryType
                            && range.Start < mathRange.End && range.End > mathRange.Start)
                            throw new InvalidDataException("The external equation caption overlaps unowned mathematical content.");
                    }
                    finally { Release(mathRange); Release(math); }
                }
            }
            finally
            {
                Release(bookmarkMaths);
            }

            // VisualTeX keeps the native SEQ caption in a 0.1 pt clipping Frame
            // at the bottom-right page boundary so Word's native cross-reference
            // dialog can still discover the target. Deleting only the bookmarked
            // caption text leaves that empty Frame alive. A later numbered-formula
            // -> LaTeX conversion can then insert the restored source at the same
            // document position and Word silently adopts the new source into the
            // surviving clipping Frame, making the LaTeX text effectively invisible.
            // Remove the Frame formatting first; Frame.Delete preserves its text.
            // Detach the owned bookmark before changing the Frame. Frame.Delete
            // can shrink its live range. Native Undo can still retain a shortened
            // span; the transaction checkpoint verifies and restores owned spans.
            var captionStart = range.Start;
            var captionEnd = range.End;
            var captionText = range.Text;
            bookmark.Delete();
            frames = range.Frames;
            if (frames.Count > 0)
            {
                if (frames.Count != 1)
                    throw new InvalidDataException("An equation caption has ambiguous Frame ownership.");
                frame = frames[1];
                Range? frameRange = null;
                try
                {
                    frameRange = frame.Range;
                    // A shared caption Frame belongs to several equations. Removing
                    // it for one source destroys the other captions' layout and can
                    // rebind their bookmarks. Only detach an exclusively owned
                    // Frame; deleting this one caption paragraph shrinks a shared
                    // Frame normally without touching its neighbours.
                    if (frameRange.Start >= captionStart && frameRange.End <= captionEnd)
                        frame.Delete();
                }
                finally { Release(frameRange); }
                Release(frame);
                frame = null;
                Release(frames);
                frames = null;
                Release(range);
                range = document.Range(captionStart, captionEnd);
                if (range.Text != captionText)
                    throw new InvalidDataException("Removing the equation caption Frame changed its content.");
            }

            if (preserveParagraphSeparator
                && range.End > range.Start
                && (range.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal))
            {
                // The hidden caption paragraph is also the separator between two
                // adjacent numbered tables. Removing its final paragraph mark makes
                // Word merge those tables immediately. A numbered OLE -> OMML
                // replacement rebuilds the caption in the same table, so remove the
                // old SEQ/content while preserving that separator until the rebuild.
                var contentStart = range.Start;
                var contentEnd = range.End - 1;
                if (contentEnd > contentStart)
                {
                    Range? captionContents = null;
                    try
                    {
                        captionContents = document.Range(contentStart, contentEnd);
                        captionContents.Delete();
                    }
                    finally { Release(captionContents); }
                }
            }
            else
            {
                // Frame.Delete preserves the SEQ text but can shrink VTEqCap from
                // the original full paragraph to only the framed contents. Deleting
                // that mutated bookmark range leaves its paragraph mark behind as a
                // real empty line after VisualTeX -> MathType conversion. Freeze the
                // original whole-paragraph coordinates before touching the Frame and
                // remove that exact paragraph after detaching bookmark ownership.
                if (originalOwnedWholeParagraph
                    && originalParagraphStart >= document.Content.Start
                    && originalParagraphEnd <= document.Content.End
                    && originalParagraphEnd > originalParagraphStart)
                {
                    Range? completeCaptionParagraph = null;
                    try
                    {
                        completeCaptionParagraph = document.Range(
                            originalParagraphStart,
                            originalParagraphEnd);
                        completeCaptionParagraph.Delete();
                    }
                    finally { Release(completeCaptionParagraph); }
                }
                else
                {
                    range.Delete();
                }
            }
        }
        finally
        {
            Release(frame);
            Release(frames);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static void RemoveOrphanEquationArtifacts(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        RemoveOrphanBookmarks(
            document,
            EquationBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeCaptionBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeNumberBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: false);
        // Flow-tail bookmarks own generated empty paragraphs only while their
        // numbered formula exists. Orphan cleanup drops identity rather than text;
        // destructive reclamation is performed only by the strict local teardown
        // path that can prove the exact body->table geometry.
        RemoveOrphanBookmarks(
            document,
            NativeCaptionFlowTailBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: false);
    }

    private static void RemoveOrphanBookmarks(
        Document document,
        string prefix,
        ISet<string> activeFormulaIds,
        bool deleteRange)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = bookmarks.Count; index >= 1; index--)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                OMaths? maths = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(bookmark.Name, prefix, out var formulaId)
                        || activeFormulaIds.Contains(formulaId))
                        continue;
                    if (deleteRange)
                    {
                        range = bookmark.Range;
                        maths = range.OMaths;
                        if (maths.Count > 0)
                        {
                            // Native #(SEQ) aliases live inside the mathematical
                            // field result. Removing their text would corrupt the
                            // OMath/field tree. Drop only the orphan identity; an
                            // explicit unnumber/migration operation rebuilds the
                            // complete OMath atomically when content must disappear.
                            bookmark.Delete();
                        }
                        else
                        {
                            range.Delete();
                        }
                    }
                    else
                    {
                        bookmark.Delete();
                    }
                }
                finally
                {
                    Release(maths);
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    internal static void UpdateEquationNumberFields(
        Document document,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId)
    {
        // Native numbered OMML owns no generated REF label to restyle. Refresh its
        // in-math SEQ field only; touching the VTEq_ bookmark's Font/Position would
        // mutate the mathematical delimiter and can destabilize Word's #() host.
        if (IsNumberedNativeOmmlHashSequenceFormula(document, formulaId))
        {
            if (!TryRefreshNumberedNativeOmmlHashSequence(document, formulaId))
                throw new InvalidOperationException(
                    $"Native #(SEQ) OMML {formulaId} could not refresh its field without rewriting Field.Code.Text.");
            return;
        }

        UpdateFieldInBookmark(
            document,
            EquationBookmarkName(formulaId),
            code => IsReferenceToBookmark(code, NativeNumberBookmarkName(formulaId)));

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return;
            bookmark = bookmarks[visibleName];
            range = bookmark.Range;
            // The tab-layout number follows the paragraph baseline exactly, like
            // a native MathType display row. The VisualTeX OLE itself owns the
            // baseline offset; never move the number upward to compensate for the
            // object's height. Table numbers use the same body-font contract.
            AlignEquationNumberVertically(
                range,
                formulaHeightPoints: 0f,
                formulaFontSizePoints: formulaFontSizePoints);
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static IReadOnlyDictionary<string, FormulaDocumentLocation>
        CaptureNumberedFormulaLocations(
            Document document,
            InlineShapes inlineShapes,
            IReadOnlyList<string> ommlFormulaIds)
    {
        var result = new Dictionary<string, FormulaDocumentLocation>(
            StringComparer.OrdinalIgnoreCase);
        for (var index = 1; index <= inlineShapes.Count; index++)
        {
            InlineShape? shape = null;
            Range? range = null;
            try
            {
                shape = inlineShapes[index];
                var metadata = ReadMetadata(shape);
                if (metadata?.DisplayMode != "block" || !metadata.Numbered) continue;
                range = shape.Range;
                result[metadata.FormulaId] = new FormulaDocumentLocation(
                    range.Start,
                    range.End);
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }

        foreach (var formulaId in ommlFormulaIds)
        {
            Bookmark? bookmark = null;
            Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block" || !metadata.Numbered) continue;
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                result[formulaId] = new FormulaDocumentLocation(
                    range.Start,
                    range.End);
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }
        return result;
    }

    internal static void RepairSharedNativeCaptionArtifacts(
        Document document,
        IReadOnlyDictionary<string, FormulaDocumentLocation> formulaLocations)
    {
        var ownersByArtifactRange = new Dictionary<string, HashSet<string>>(
            StringComparer.Ordinal);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    string artifactKind;
                    string formulaId;
                    if (TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out formulaId))
                    {
                        artifactKind = "number";
                    }
                    else if (TryFormulaIdFromBookmark(
                                 bookmark.Name,
                                 NativeCaptionBookmarkPrefix,
                                 out formulaId))
                    {
                        artifactKind = "caption";
                    }
                    else
                    {
                        continue;
                    }
                    if (!formulaLocations.ContainsKey(formulaId)) continue;

                    range = bookmark.Range;
                    var key = artifactKind + ":" + range.Start + ":" + range.End;
                    if (!ownersByArtifactRange.TryGetValue(key, out var owners))
                    {
                        owners = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        ownersByArtifactRange[key] = owners;
                    }
                    owners.Add(formulaId);
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }

        var conflictedFormulaIds = ownersByArtifactRange.Values
            .Where(owners => owners.Count > 1)
            .SelectMany(owners => owners)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (conflictedFormulaIds.Count == 0) return;

        // A shared caption paragraph cannot safely keep one "owner": deleting or
        // rewriting that paragraph invalidates every overlapping VTEqCap/VTEqNum
        // bookmark. Remove the complete numbering scaffold for every participant
        // first, then the normal reconciliation loop rebuilds one independent
        // visible REF and one independent native SEQ target per formula.
        foreach (var formulaId in conflictedFormulaIds
                     .OrderByDescending(id => formulaLocations[id].Start))
        {
            RemoveVisibleEquationNumber(document, formulaId);
            RemoveNativeCaption(document, formulaId);
        }
    }

    internal static void RebuildNativeNumberBookmarksFromCaptions(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var formulaId in numberedFormulaIds)
            {
                Bookmark? captionBookmark = null;
                Range? captionRange = null;
                Field? sequenceField = null;
                Range? fieldResult = null;
                Range? completeNumberRange = null;
                Bookmark? rebuiltNumberBookmark = null;
                try
                {
                    var captionName = NativeCaptionBookmarkName(formulaId);
                    if (!bookmarks.Exists(captionName))
                        throw new InvalidOperationException(
                            $"VisualTeX formula {formulaId} has no native caption after reconciliation.");
                    captionBookmark = bookmarks[captionName];
                    captionRange = captionBookmark.Range;
                    sequenceField = FindNativeEquationFieldInCaption(
                        captionRange,
                        nativeSequenceName);
                    if (sequenceField is null)
                        throw new InvalidOperationException(
                            $"VisualTeX formula {formulaId} has no native SEQ field inside its caption.");
                    fieldResult = sequenceField.Result;
                    // After heading-aware renumbering, the chapter/section
                    // prefix is ordinary text immediately before the SEQ field.
                    // Rebuilding from Field.Result alone silently changes a
                    // target such as "1-1" back to "1". Preserve the complete
                    // visible caption number from the paragraph start through
                    // the field result; continuous numbering naturally has an
                    // empty prefix and uses the same range.
                    completeNumberRange = document.Range(
                        captionRange.Start,
                        fieldResult.End);

                    var numberName = NativeNumberBookmarkName(formulaId);
                    if (bookmarks.Exists(numberName))
                    {
                        Bookmark? existing = null;
                        try
                        {
                            existing = bookmarks[numberName];
                            existing.Delete();
                        }
                        finally { Release(existing); }
                    }
                    rebuiltNumberBookmark = bookmarks.Add(
                        numberName,
                        completeNumberRange);
                }
                finally
                {
                    Release(rebuiltNumberBookmark);
                    Release(completeNumberRange);
                    Release(fieldResult);
                    Release(sequenceField);
                    Release(captionRange);
                    Release(captionBookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static Field? FindNativeEquationFieldInCaption(
        Range captionRange,
        string nativeSequenceName)
    {
        Fields? fields = null;
        try
        {
            fields = captionRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(
                            code.Text,
                            nativeSequenceName))
                        continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    internal sealed class FormulaDocumentLocation
    {
        public FormulaDocumentLocation(int start, int end)
        {
            Start = start;
            End = end;
        }

        public int Start { get; }
        public int End { get; }
    }
}

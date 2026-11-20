using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationNumberFormat;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordVisibleEquationNumbers;

namespace VisualTeX.WordVsto;

/// <summary>
/// Native equation captions: creating, reusing, styling and removing them, and their flow tails.
/// </summary>
internal static class WordNativeCaptions
{
    internal static void EnsureNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName,
        bool restyleExisting = true,
        Table? knownNumberedTable = null,
        bool deferFieldUpdate = false,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool knownFreshConversionTarget = false)
    {
        if (ContainsNativeOmml(formulaRange))
        {
            if (ContainsVisualTeXSequenceInsideOmml(formulaRange))
                return;
            throw new InvalidOperationException(
                "VisualTeX refused to create an external caption for native OMML. Numbered OMML must be materialized atomically as wdOMathDisplay #(SEQ VisualTeXEquation).");
        }

        Range? captionRange = null;
        Range? numberRange = null;
        if (!knownFreshConversionTarget
            && TryGetNativeCaptionRanges(
                document,
                formulaId,
                nativeSequenceName,
                out captionRange,
                out numberRange)
            && captionRange is not null
            && numberRange is not null)
        {
            try
            {
                if (restyleExisting)
                    StyleNativeCaption(captionRange, numberRange);
            }
            finally
            {
                Release(numberRange);
                Release(captionRange);
            }
            return;
        }
        Release(numberRange);
        Release(captionRange);

        if (!knownFreshConversionTarget)
            RemoveNativeCaption(document, formulaId);
        CreateNativeCaption(
            document,
            formulaRange,
            formulaId,
            nativeSequenceName,
            knownNumberedTable,
            deferFieldUpdate,
            plannedOrdinal,
            plannedPrefix);
    }

    private static bool TryCaptureTopLevelTableAt(
        Document document,
        int position,
        WdStoryType storyType,
        out int tableEnd)
    {
        tableEnd = -1;
        Range? content = null;
        Range? probe = null;
        Tables? tables = null;
        Table? table = null;
        Range? tableRange = null;
        try
        {
            content = document.Content;
            if (position < content.Start || position >= content.End) return false;
            probe = document.Range(position, Math.Min(content.End, position + 1));
            tables = probe.Tables;
            for (var index = 1; index <= tables.Count; index++)
            {
                Release(tableRange); tableRange = null;
                Release(table); table = tables[index];
                if (table.NestingLevel != 1) continue;
                tableRange = table.Range;
                if (tableRange.StoryType != storyType || tableRange.Start != position)
                    continue;
                tableEnd = tableRange.End;
                return true;
            }
            return false;
        }
        finally
        {
            Release(tableRange);
            Release(table);
            Release(tables);
            Release(probe);
            Release(content);
        }
    }

    private static bool TryCaptureImmediateFollowingTopLevelTableBoundary(
        Document document,
        Range paragraphOwnedRange,
        out int tableStart,
        out int tableEnd)
    {
        tableStart = -1;
        tableEnd = -1;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = paragraphOwnedRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            tableStart = paragraphRange.End;
            if (!TryCaptureTopLevelTableAt(
                    document,
                    tableStart,
                    paragraphRange.StoryType,
                    out tableEnd))
            {
                tableStart = -1;
                tableEnd = -1;
                return false;
            }
            return true;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static bool IsExactOwnedNativeCaptionFlowTail(
        int expectedStart,
        int rangeStart,
        int rangeEnd,
        string? text,
        int paragraphStart,
        int paragraphEnd,
        bool withinTable,
        int tableCount,
        int shapeCount,
        int mathCount,
        int fieldCount,
        int frameCount)
        => rangeStart == expectedStart
            && rangeEnd == rangeStart + 1
            && paragraphStart == rangeStart
            && paragraphEnd == rangeEnd
            && string.Equals(text, "\r", StringComparison.Ordinal)
            && !withinTable
            && tableCount == 0
            && shapeCount == 0
            && mathCount == 0
            && fieldCount == 0
            && frameCount == 0;

    private static void MarkGeneratedNativeCaptionFlowTail(
        Document document,
        string formulaId,
        int followingTableStartBeforeStyle,
        int followingTableEndBeforeStyle)
    {
        Range? tail = null;
        Paragraphs? tailParagraphs = null;
        Paragraph? tailParagraph = null;
        Range? tailParagraphRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Frames? frames = null;
        Tables? localTables = null;
        Bookmarks? localBookmarks = null;
        Bookmark? localBookmark = null;
        Bookmarks? bookmarks = null;
        Bookmark? marker = null;
        Range? marked = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName))
                throw new InvalidDataException(
                    "The native caption identity disappeared while its Frame was created.");

            if (TryCaptureTopLevelTableAt(
                    document,
                    followingTableStartBeforeStyle,
                    WdStoryType.wdMainTextStory,
                    out var unchangedTableEnd))
            {
                if (unchangedTableEnd != followingTableEndBeforeStyle)
                    throw new InvalidDataException(
                        "The following user table changed while styling the native caption.");
                return; // This Word build did not materialize an extra flow paragraph.
            }

            tail = document.Range(
                followingTableStartBeforeStyle,
                followingTableStartBeforeStyle + 1);
            tailParagraphs = tail.Paragraphs;
            if (tailParagraphs.Count != 1)
                throw new InvalidDataException(
                    "The generated native-caption flow tail has ambiguous paragraph ownership.");
            tailParagraph = tailParagraphs[1];
            tailParagraphRange = tailParagraph.Range;
            shapes = tail.InlineShapes;
            maths = tail.OMaths;
            fields = WordFormulaHost.GetLocalFields(tail);
            frames = tail.Frames;
            localTables = tail.Tables;
            localBookmarks = tail.Bookmarks;
            var onlyOwnedTransientAliases = true;
            for (var bookmarkIndex = 1; bookmarkIndex <= localBookmarks.Count; bookmarkIndex++)
            {
                Release(localBookmark); localBookmark = localBookmarks[bookmarkIndex];
                if (!string.Equals(localBookmark.Name, captionName, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(localBookmark.Name, numberName, StringComparison.OrdinalIgnoreCase))
                {
                    onlyOwnedTransientAliases = false;
                    break;
                }
            }
            if (!IsExactOwnedNativeCaptionFlowTail(
                    followingTableStartBeforeStyle,
                    tail.Start,
                    tail.End,
                    tail.Text,
                    tailParagraphRange.Start,
                    tailParagraphRange.End,
                    RangeIsWhollyWithinTable(tail),
                    localTables.Count,
                    shapes.Count,
                    maths.Count,
                    fields.Count,
                    frames.Count)
                || !onlyOwnedTransientAliases)
                throw new InvalidDataException(
                    "Word inserted unexpected content between the native caption and following table.");

            if (!TryCaptureTopLevelTableAt(
                    document,
                    tail.End,
                    tail.StoryType,
                    out var shiftedTableEnd)
                || shiftedTableEnd != followingTableEndBeforeStyle + 1)
                throw new InvalidDataException(
                    "The following user table did not shift by exactly one generated caption-flow paragraph.");

            var name = NativeCaptionFlowTailBookmarkName(formulaId);
            if (bookmarks.Exists(name))
                throw new InvalidDataException(
                    "A native-caption flow-tail identity already exists before creation.");
            marker = bookmarks.Add(name, tail);
            marked = marker.Range;
            if (marked.Start != tail.Start
                || marked.End != tail.End
                || !string.Equals(marked.Text, "\r", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not retain the generated native-caption flow-tail identity.");
            WordDoubleClickHook.TraceMessage(
                $"native-caption-flow-tail-owned formulaId={formulaId} range={tail.Start}:{tail.End}");
        }
        finally
        {
            Release(marked);
            Release(marker);
            Release(bookmarks);
            Release(localBookmark);
            Release(localBookmarks);
            Release(localTables);
            Release(frames);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(tailParagraphRange);
            Release(tailParagraph);
            Release(tailParagraphs);
            Release(tail);
        }
    }

    internal static bool IsEmptyOwnedCaptionFlowTailBeforeConversion(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? marker = null;
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = NativeCaptionFlowTailBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return false;
            marker = bookmarks[name];
            range = marker.Range;
            if (range.End != range.Start + 1 || range.Text != "\r"
                || RangeIsWhollyWithinTable(range)) return false;
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.Start == range.Start && paragraphRange.End == range.End
                && TryCaptureTopLevelTableAt(document, range.End, range.StoryType, out _);
        }
        finally { Release(paragraphRange); Release(paragraph); Release(paragraphs); Release(range); Release(marker); Release(bookmarks); }
    }

    internal static void CompleteDeferredCaptionFlowTailCleanup(
        Document document, string sourceFormulaId, bool verifiedEmptyBeforeConversion)
        => CleanupOwnedNativeCaptionFlowTail(document, sourceFormulaId, deleteParagraph: true,
            recoverExpandedStart: verifiedEmptyBeforeConversion);

    internal static void CompleteDeferredOmmlSeparatorCleanup(Document document, string sourceFormulaId)
        => CleanupOwnedNativeCaptionFlowTail(document, sourceFormulaId, deleteParagraph: true,
            recoverExpandedStart: true,
            ownedMarkerName: NativeOmmlOwnedSeparatorBookmarkName(sourceFormulaId));

    internal static void CleanupOwnedNativeCaptionFlowTail(
        Document document,
        string formulaId,
        bool deleteParagraph,
        bool recoverExpandedStart = false,
        string? ownedMarkerName = null)
    {
        Bookmarks? bookmarks = null;
        Bookmark? marker = null;
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Frames? frames = null;
        Tables? localTables = null;
        Bookmarks? localBookmarks = null;
        Bookmark? localBookmark = null;
        Range? content = null;
        Range? previous = null;
        var tableEndBeforeDelete = -1;
        var safeToDelete = false;
        try
        {
            bookmarks = document.Bookmarks;
            var name = ownedMarkerName ?? NativeCaptionFlowTailBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return;
            marker = bookmarks[name];
            range = marker.Range;
            // A freshly inserted OLE can expand the START of the next paragraph's
            // bookmark backward. Only a transaction which proved this exact tag
            // owned one empty paragraph before mutation may recover its surviving
            // end. Persisted expanded/user-edited tags never receive this authority.
            if (recoverExpandedStart && range.End > range.Start + 1)
            {
                var survivingEnd = range.End;
                Release(range);
                range = document.Range(survivingEnd - 1, survivingEnd);
            }
            paragraphs = range.Paragraphs;
            if (paragraphs.Count == 1)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                shapes = range.InlineShapes;
                maths = range.OMaths;
                fields = WordFormulaHost.GetLocalFields(range);
                frames = range.Frames;
                localTables = range.Tables;
                localBookmarks = paragraphRange.Bookmarks;
                var ownsOnlyMarker = localBookmarks.Count == 1;
                if (ownsOnlyMarker)
                {
                    localBookmark = localBookmarks[1];
                    ownsOnlyMarker = string.Equals(
                        localBookmark.Name,
                        name,
                        StringComparison.OrdinalIgnoreCase);
                }
                content = document.Content;
                var previousIsTable = false;
                if (paragraphRange.Start > content.Start)
                {
                    previous = document.Range(
                        paragraphRange.Start - 1,
                        paragraphRange.Start);
                    previousIsTable = RangeIsWhollyWithinTable(previous);
                }
                var followedByTopLevelTable = TryCaptureTopLevelTableAt(
                    document,
                    paragraphRange.End,
                    paragraphRange.StoryType,
                    out tableEndBeforeDelete);
                safeToDelete = ownsOnlyMarker
                    && !previousIsTable
                    && followedByTopLevelTable
                    && IsExactOwnedNativeCaptionFlowTail(
                        range.Start,
                        range.Start,
                        range.End,
                        range.Text,
                        paragraphRange.Start,
                        paragraphRange.End,
                        RangeIsWhollyWithinTable(range),
                        localTables.Count,
                        shapes.Count,
                        maths.Count,
                        fields.Count,
                        frames.Count);
            }

            if (Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE") == "1")
                WordDoubleClickHook.TraceMessage(
                    $"native-caption-flow-tail-check formulaId={formulaId} range={range.Start}:{range.End} text=[{(range.Text ?? string.Empty).Replace("\r", "<CR>")}] paragraphs={paragraphs.Count} paragraph={paragraphRange?.Start}:{paragraphRange?.End} tables={localTables?.Count} shapes={shapes?.Count} maths={maths?.Count} fields={fields?.Count} frames={frames?.Count} bookmarks={localBookmarks?.Count} safe={safeToDelete}");
            marker.Delete();
            Release(marker); marker = null;
            if (!deleteParagraph || !safeToDelete)
            {
                WordDoubleClickHook.TraceMessage(
                    $"native-caption-flow-tail-released formulaId={formulaId} deleteRequested={deleteParagraph} safe={safeToDelete}");
                return;
            }

            var oldStart = paragraphRange!.Start;
            paragraphRange.Delete();
            if (!TryCaptureTopLevelTableAt(
                    document,
                    oldStart,
                    WdStoryType.wdMainTextStory,
                    out var tableEndAfterDelete)
                || tableEndAfterDelete != tableEndBeforeDelete - 1)
                throw new InvalidDataException(
                    "Removing the owned native-caption flow tail changed the following user table unexpectedly.");
            WordDoubleClickHook.TraceMessage(
                $"native-caption-flow-tail-removed formulaId={formulaId} start={oldStart}");
        }
        finally
        {
            Release(previous);
            Release(content);
            Release(localBookmark);
            Release(localBookmarks);
            Release(localTables);
            Release(frames);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
            Release(marker);
            Release(bookmarks);
        }
    }

    private static void CaptureBodyCaptionFollowingWitness(
        Document document,
        Range captionRange,
        out int followingStart,
        out int contentEnd,
        out int witnessSpan,
        out string witness)
    {
        Range? content = null;
        Range? probe = null;
        try
        {
            content = document.Content;
            followingStart = Math.Max(
                content.Start,
                Math.Min(captionRange.End, content.End));
            contentEnd = content.End;
            var witnessEnd = Math.Min(
                content.End,
                followingStart + 64);
            witnessSpan = Math.Max(
                0,
                witnessEnd - followingStart);
            if (witnessSpan == 0)
            {
                witness = string.Empty;
                return;
            }
            probe = document.Range(
                followingStart,
                witnessEnd);
            witness = probe.Text ?? string.Empty;
        }
        finally
        {
            Release(probe);
            Release(content);
        }
    }

    private static bool TryRemoveGeneratedBodyCaptionFlowParagraph(
        Document document,
        string formulaId,
        int followingStartBeforeStyle,
        int contentEndBeforeStyle,
        int followingWitnessSpan,
        string followingWitness)
    {
        Range? content = null;
        Range? candidate = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Frames? frames = null;
        Tables? tables = null;
        Bookmarks? bookmarks = null;
        Range? shiftedWitness = null;
        Range? restoredWitness = null;
        try
        {
            content = document.Content;
            if (content.End != contentEndBeforeStyle + 1
                || followingStartBeforeStyle < content.Start
                || followingStartBeforeStyle + 1 > content.End)
            {
                return false;
            }

            candidate = document.Range(
                followingStartBeforeStyle,
                followingStartBeforeStyle + 1);
            paragraphs = candidate.Paragraphs;
            if (paragraphs.Count != 1)
                return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            shapes = candidate.InlineShapes;
            maths = candidate.OMaths;
            fields = WordFormulaHost.GetLocalFields(candidate);
            frames = candidate.Frames;
            tables = candidate.Tables;
            bookmarks = candidate.Bookmarks;
            var exactGeneratedTail = IsExactOwnedNativeCaptionFlowTail(
                    followingStartBeforeStyle,
                    candidate.Start,
                    candidate.End,
                    candidate.Text,
                    paragraphRange.Start,
                    paragraphRange.End,
                    RangeIsWhollyWithinTable(candidate),
                    tables.Count,
                    shapes.Count,
                    maths.Count,
                    fields.Count,
                    frames.Count);
            var onlyOwnedAliases = true;
            Bookmark? localBookmark = null;
            try
            {
                var captionName = NativeCaptionBookmarkName(formulaId);
                var numberName = NativeNumberBookmarkName(formulaId);
                for (var bookmarkIndex = 1;
                     bookmarkIndex <= bookmarks.Count;
                     bookmarkIndex++)
                {
                    Release(localBookmark);
                    localBookmark = bookmarks[bookmarkIndex];
                    if (!string.Equals(
                            localBookmark.Name,
                            captionName,
                            StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(
                            localBookmark.Name,
                            numberName,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        onlyOwnedAliases = false;
                        break;
                    }
                }
            }
            finally { Release(localBookmark); }
            if (!exactGeneratedTail
                || !onlyOwnedAliases)
                return false;

            if (followingWitnessSpan == 0
                && followingStartBeforeStyle == contentEndBeforeStyle)
            {
                // A Word document must always retain one final body paragraph.
                // When a clipping Frame is created for the final equation caption,
                // Word materializes that mandatory trailing CR outside the Frame.
                // Deleting it simply makes Word recreate the same paragraph, so
                // preserve the structural terminator and compress only its visual
                // line box. This is safe only for the exact empty, object-free,
                // bookmark-free/owned-alias tail proven above.
                CompressNativeCaptionParagraphFlow(candidate);
                return true;
            }

            if (followingWitnessSpan > 0)
            {
                var shiftedEnd = Math.Min(
                    content.End,
                    followingStartBeforeStyle
                    + 1
                    + followingWitnessSpan);
                shiftedWitness = document.Range(
                    followingStartBeforeStyle + 1,
                    shiftedEnd);
                if (!string.Equals(
                        shiftedWitness.Text ?? string.Empty,
                        followingWitness,
                        StringComparison.Ordinal))
                {
                    return false;
                }
            }

            candidate.Delete();
            Release(content);
            content = document.Content;
            if (content.End != contentEndBeforeStyle)
                throw new InvalidDataException(
                    "Removing Word's generated equation-caption flow paragraph changed the document length unexpectedly.");

            if (followingWitnessSpan > 0)
            {
                var restoredEnd = Math.Min(
                    content.End,
                    followingStartBeforeStyle
                    + followingWitnessSpan);
                restoredWitness = document.Range(
                    followingStartBeforeStyle,
                    restoredEnd);
                if (!string.Equals(
                        restoredWitness.Text ?? string.Empty,
                        followingWitness,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "Removing Word's generated equation-caption flow paragraph changed following user content.");
            }
            return true;
        }
        finally
        {
            Release(restoredWitness);
            Release(shiftedWitness);
            Release(bookmarks);
            Release(tables);
            Release(frames);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(candidate);
            Release(content);
        }
    }

    private static int CreateNativeCaptionParagraphAfterBodyFormula(
        Document document,
        Range formulaParagraphRange)
    {
        Range? content = null;
        Range? followingProbe = null;
        Tables? followingTables = null;
        Table? followingTable = null;
        Range? followingTableRange = null;
        Range? insertion = null;
        Range? verification = null;
        Cell? verificationCell = null;
        try
        {
            content = document.Content;
            var captionStart = formulaParagraphRange.End;
            var followedImmediatelyByTable = false;
            if (captionStart < content.End)
            {
                followingProbe = document.Range(
                    captionStart,
                    Math.Min(content.End, captionStart + 1));
                followingTables = followingProbe.Tables;
                if (followingTables.Count > 0)
                {
                    followingTable = followingTables[1];
                    followingTableRange = followingTable.Range;
                    followedImmediatelyByTable =
                        followingTableRange.StoryType == formulaParagraphRange.StoryType
                        && followingTableRange.Start == captionStart;
                }
            }

            if (!followedImmediatelyByTable)
            {
                formulaParagraphRange.InsertParagraphAfter();
                return captionStart;
            }

            // Word 2021 treats Range.InsertParagraphAfter() at a body→table
            // boundary as an insertion into cell (1,1). The actual body boundary
            // is the preceding paragraph mark. Insert a NEW mark immediately before
            // that existing mark: the formula keeps the new mark, the old mark
            // becomes one empty body paragraph, and the whole table moves forward
            // by one story position without changing any cell.
            var terminatorLength =
                GetWordParagraphStructuralTerminatorLength(formulaParagraphRange);
            if (terminatorLength != 1
                || formulaParagraphRange.End - terminatorLength < formulaParagraphRange.Start)
                throw new InvalidDataException(
                    "The body equation before a Word table has no ordinary paragraph terminator.");
            insertion = document.Range(
                formulaParagraphRange.End - terminatorLength,
                formulaParagraphRange.End - terminatorLength);
            insertion.InsertBefore("\r");

            verification = document.Range(captionStart, captionStart + 1);
            if (!string.Equals(verification.Text, "\r", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word did not create the dedicated equation-caption paragraph before the following table.");
            verificationCell = WordFormulaHost.TryGetOwningCell(verification);
            if (verificationCell is not null)
                throw new InvalidDataException(
                    "Word placed the dedicated equation-caption paragraph inside the following table.");

            Release(followingTableRange);
            followingTableRange = followingTable.Range;
            if (followingTableRange.Start != captionStart + 1)
                throw new InvalidDataException(
                    "Creating the equation-caption paragraph changed the following table boundary unexpectedly.");
            return captionStart;
        }
        finally
        {
            Release(verificationCell);
            Release(verification);
            Release(insertion);
            Release(followingTableRange);
            Release(followingTable);
            Release(followingTables);
            Release(followingProbe);
            Release(content);
        }
    }

    private static bool CanReuseEmptyNativeCaptionParagraph(
        Document document,
        int position)
    {
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Frames? frames = null;
        Tables? tables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        try
        {
            content = document.Content;
            if (position < content.Start || position >= content.End) return false;
            probe = document.Range(position, Math.Min(content.End, position + 1));
            if ((bool)probe.get_Information(WdInformation.wdWithInTable)) return false;
            frames = probe.Frames;
            if (frames.Count > 0) return false;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start != position) return false;
            tables = paragraphRange.Tables;
            shapes = paragraphRange.InlineShapes;
            maths = paragraphRange.OMaths;
            fields = paragraphRange.Fields;
            if (tables.Count > 0 || shapes.Count > 0 || maths.Count > 0 || fields.Count > 0)
                return false;
            var text = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\n', '\t', '\v', ' ');
            return text.Length == 0;
        }
        finally
        {
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(tables);
            Release(frames);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
        }
    }

    internal static void CreateNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName,
        Table? knownNumberedTable = null,
        bool deferFieldUpdate = false,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool allowNativeOmmlAcceptanceFixture = false)
    {
        if (ContainsNativeOmml(formulaRange)
            && !allowNativeOmmlAcceptanceFixture)
            throw new InvalidOperationException(
                "VisualTeX refused to create a hidden SEQ caption beside native OMML. The SEQ field belongs inside the Word-native #() number slot.");
        if (allowNativeOmmlAcceptanceFixture
            && !string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The native-OMML legacy caption bypass is available only to isolated acceptance fixture construction.");

        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] NativeCaption.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Table? formulaTable = null;
        Range? formulaTableRange = null;
        Range? fieldInsertion = null;
        Fields? fields = null;
        Field? captionField = null;
        Range? numberRange = null;
        Range? captionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Bookmarks? bookmarks = null;
        var bodyCaption = false;
        var followingTableStartBeforeStyle = -1;
        var followingTableEndBeforeStyle = -1;
        try
        {
            // Word's Range.InsertCaption mutates the equation paragraph. For a
            // trailing inline OMath it moves the ordinary run after the equation
            // into a new paragraph, so the visible REF field is subsequently
            // absorbed into m:oMath. Build the native SEQ caption in a dedicated
            // hidden paragraph instead and leave the equation paragraph intact.
            int captionStart;
            if (knownNumberedTable is not null)
            {
                formulaTableRange = knownNumberedTable.Range;
                // A Word table's Range.InsertParagraphAfter() expands the table
                // range itself; its new End is then the start of the following
                // body paragraph. Adding the native SEQ field at that position
                // merges the caption with the user's next paragraph, and the
                // 0.1 pt clipping Frame subsequently hides that paragraph and any
                // inline formula in it. Capture the original table boundary and
                // create a dedicated paragraph there instead.
                captionStart = formulaTableRange.End;
                if (!CanReuseEmptyNativeCaptionParagraph(document, captionStart))
                {
                    fieldInsertion = document.Range(captionStart, captionStart);
                    fieldInsertion.InsertParagraphBefore();
                    Release(fieldInsertion);
                    fieldInsertion = null;
                }
            }
            else if (RangeIsWhollyWithinTable(formulaRange))
            {
                // All user-cell OLE captions live outside the enclosing table;
                // their visible REF stays beside the formula inside its own cell.
                // InsertParagraphAfter on a Cell paragraph would target the next
                // cell's content and is never a valid caption insertion boundary.
                formulaTable = formulaRange.Tables[1];
                formulaTableRange = formulaTable.Range;
                captionStart = ResolveCaptionAppendPositionAfterUserTable(document, formulaTableRange.End);
                if (!CanReuseEmptyNativeCaptionParagraph(document, captionStart))
                {
                    fieldInsertion = document.Range(captionStart, captionStart);
                    fieldInsertion.InsertParagraphBefore();
                    Release(fieldInsertion);
                    fieldInsertion = null;
                }
            }
            else
            {
                bodyCaption = true;
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                captionStart = CreateNativeCaptionParagraphAfterBodyFormula(
                    document,
                    formulaParagraphRange);
            }

            object insertionStart = captionStart;
            object insertionEnd = captionStart;
            fieldInsertion = document.Range(ref insertionStart, ref insertionEnd);
            if (!string.IsNullOrEmpty(plannedPrefix))
            {
                fieldInsertion.Text = plannedPrefix;
                var fieldPosition = fieldInsertion.End;
                Release(fieldInsertion);
                fieldInsertion = document.Range(fieldPosition, fieldPosition);
            }
            TraceStage("prepare-paragraph");
            // Add through the tiny target Range rather than document.Fields.
            // Materializing the full document field collection on every appended
            // equation makes Word's field bookkeeping grow with formula count.
            fields = fieldInsertion.Fields;
            object fieldType = WdFieldEmpty;
            // All newly produced VisualTeX OLE and OMML numbers share the exact
            // product-owned sequence used by Word's native #(SEQ) math host. The
            // localized built-in Equation label remains readable as legacy input,
            // but creating another localized SEQ would split mixed OLE/OMML order.
            object fieldCode = plannedOrdinal.HasValue
                ? $"SEQ {LegacyEquationSequenceName} \\r {plannedOrdinal.Value} \\* ARABIC"
                : $"SEQ {LegacyEquationSequenceName} \\* ARABIC";
            object preserveFormatting = true;
            captionField = fields.Add(
                fieldInsertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            TraceStage("add-seq");
            if (!deferFieldUpdate)
                captionField.Update();
            TraceStage(deferFieldUpdate ? "defer-update-seq" : "update-seq");
            numberRange = captionField.Result;
            paragraphs = numberRange.Paragraphs;
            paragraph = paragraphs[1];
            captionRange = paragraph.Range;
            try
            {
                object captionStyle = WdBuiltinStyle.wdStyleCaption;
                captionRange.set_Style(ref captionStyle);
            }
            catch (Exception ignored)
            {
                // Some locked/custom documents reject assigning the built-in
                // caption style. The SEQ field and bookmarks remain valid.
                IgnoredErrors.Ignore(ignored);
            }

            bookmarks = document.Bookmarks;
            if (!string.IsNullOrEmpty(plannedPrefix))
            {
                Range? plannedNumberRange = null;
                try
                {
                    plannedNumberRange = document.Range(captionRange.Start, numberRange.End);
                    bookmarks.Add(NativeNumberBookmarkName(formulaId), plannedNumberRange);
                }
                finally { Release(plannedNumberRange); }
            }
            else
            {
                bookmarks.Add(NativeNumberBookmarkName(formulaId), numberRange);
            }
            bookmarks.Add(NativeCaptionBookmarkName(formulaId), captionRange);
            if (bodyCaption)
            {
                TryCaptureImmediateFollowingTopLevelTableBoundary(
                    document,
                    captionRange,
                    out followingTableStartBeforeStyle,
                    out followingTableEndBeforeStyle);
            }
            TraceStage("bookmarks");
            var bodyFlowStartBeforeStyle = -1;
            var bodyContentEndBeforeStyle = -1;
            var bodyFollowingWitnessSpan = 0;
            var bodyFollowingWitness = string.Empty;
            if (bodyCaption
                && followingTableStartBeforeStyle < 0)
            {
                CaptureBodyCaptionFollowingWitness(
                    document,
                    captionRange,
                    out bodyFlowStartBeforeStyle,
                    out bodyContentEndBeforeStyle,
                    out bodyFollowingWitnessSpan,
                    out bodyFollowingWitness);
            }
            if (deferFieldUpdate
                && (knownNumberedTable is not null || plannedOrdinal.HasValue))
                StyleFreshConversionNativeCaption(captionRange, numberRange);
            else
                StyleNativeCaption(
                    captionRange,
                    numberRange,
                    cleanupLegacyFrames: false);
            if (bodyCaption
                && followingTableStartBeforeStyle < 0
                && bodyFlowStartBeforeStyle >= 0)
            {
                _ = TryRemoveGeneratedBodyCaptionFlowParagraph(
                    document,
                    formulaId,
                    bodyFlowStartBeforeStyle,
                    bodyContentEndBeforeStyle,
                    bodyFollowingWitnessSpan,
                    bodyFollowingWitness);
            }
            if (bodyCaption && followingTableStartBeforeStyle >= 0)
            {
                MarkGeneratedNativeCaptionFlowTail(
                    document,
                    formulaId,
                    followingTableStartBeforeStyle,
                    followingTableEndBeforeStyle);
            }
            TraceStage("style-frame");
        }
        finally
        {
            Release(bookmarks);
            Release(paragraph);
            Release(paragraphs);
            Release(captionRange);
            Release(numberRange);
            Release(captionField);
            Release(fields);
            Release(fieldInsertion);
            Release(formulaTableRange);
            Release(formulaTable);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static int ResolveCaptionAppendPositionAfterUserTable(Document document, int tableEnd)
    {
        Range? probe = null;
        Frames? frames = null;
        Frame? frame = null;
        Range? framed = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? lastParagraph = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarked = null;
        try
        {
            if (tableEnd >= document.Content.End) return tableEnd;
            probe = document.Range(tableEnd, tableEnd + 1);
            frames = probe.Frames;
            if (frames.Count != 1) return tableEnd;
            frame = frames[1];
            framed = frame.Range;
            if (framed.Start != tableEnd || framed.End <= tableEnd
                || frame.Width > 1f || frame.Height > 1f || !frame.LockAnchor || frame.TextWrap)
                return tableEnd;
            // Word coalesces identically formatted clipping captions into one
            // Frame. Append after its final paragraph, never before the first
            // caption: prepending at Table.End drags all existing number aliases
            // onto the new SEQ field through bookmark insertion gravity.
            probe.SetRange(framed.End - 1, framed.End);
            paragraphs = probe.Paragraphs;
            paragraph = paragraphs[1];
            lastParagraph = paragraph.Range;
            bookmarks = lastParagraph.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmarked); bookmarked = null;
                Release(bookmark); bookmark = bookmarks[index];
                if (!bookmark.Name.StartsWith(NativeCaptionBookmarkPrefix, StringComparison.Ordinal)
                    || !Guid.TryParseExact(bookmark.Name.Substring(NativeCaptionBookmarkPrefix.Length), "N", out _))
                    continue;
                bookmarked = bookmark.Range;
                if (bookmarked.StoryType == lastParagraph.StoryType
                    && WordFormulaHost.ContainsPhysicalRange(lastParagraph.Start, lastParagraph.End,
                        bookmarked.Start, bookmarked.End))
                    return lastParagraph.End;
            }
            throw new InvalidDataException("The existing clipped equation-caption block has lost its independent final identity; no new caption was inserted.");
        }
        finally
        {
            Release(bookmarked); Release(bookmark); Release(bookmarks);
            Release(lastParagraph); Release(paragraph); Release(paragraphs);
            Release(framed); Release(frame); Release(frames); Release(probe);
        }
    }

    private static Field? FindNewNativeEquationField(
        Document document,
        string nativeSequenceName,
        ISet<int> existingPositions,
        int formulaPosition)
    {
        Fields? fields = null;
        Field? result = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? fieldResult = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    fieldResult = field.Result;
                    if (existingPositions.Contains(fieldResult.Start)) continue;
                    var distance = Math.Abs(fieldResult.Start - formulaPosition);
                    if (distance >= bestDistance) continue;
                    Release(result);
                    result = field;
                    field = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(fieldResult);
                    Release(code);
                    Release(field);
                }
            }
            return result;
        }
        finally
        {
            Release(fields);
        }
    }

    private static Range GetNativeCaptionFrameContentRange(Range captionRange)
    {
        Range? frameRange = null;
        Cell? cell = null;
        try
        {
            frameRange = captionRange.Duplicate;
            var terminatorLength = GetWordParagraphStructuralTerminatorLength(frameRange);
            if (terminatorLength <= 0
                || frameRange.End - terminatorLength <= frameRange.Start)
                throw new InvalidDataException(
                    "The native equation caption has no frameable paragraph content.");
            frameRange.End -= terminatorLength;
            cell = WordFormulaHost.TryGetOwningCell(frameRange);
            if (cell is not null)
                throw new InvalidDataException(
                    "A native equation caption unexpectedly entered a user table cell.");
            var result = frameRange;
            frameRange = null;
            return result;
        }
        finally
        {
            Release(cell);
            Release(frameRange);
        }
    }

    private static void StyleFreshConversionNativeCaption(
        Range captionRange,
        Range numberRange)
    {
        AssertCaptionRangeIsOutsideOmml(captionRange);

        Microsoft.Office.Interop.Word.Font? font = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        ParagraphFormat? paragraph = null;
        ListFormat? listFormat = null;
        Range? frameRange = null;
        Frames? frames = null;
        Frame? frame = null;
        Borders? borders = null;
        try
        {
            // Fresh conversion captions have no legacy frames/list state to scan.
            // Keep normal visible formatting so native REF fields inherit black
            // 11pt text, but clip the SEQ paragraph into a 0.1pt frame using Word's
            // built-in page-edge positions. Avoid querying Section.PageSetup for
            // every formula: that forces pagination and becomes O(N^2) as fields
            // accumulate in large format-conversion batches.
            font = captionRange.Font;
            font.Hidden = 0;
            font.Size = 11f;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            numberFont = numberRange.Font;
            numberFont.Hidden = 0;
            numberFont.Size = 11f;
            numberFont.Color = WdColor.wdColorAutomatic;
            numberFont.Position = 0;

            paragraph = captionRange.ParagraphFormat;
            paragraph.SpaceBefore = 0f;
            paragraph.SpaceAfter = 0f;
            paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            paragraph.KeepTogether = 0;
            paragraph.KeepWithNext = 0;
            paragraph.PageBreakBefore = 0;
            paragraph.WidowControl = 0;
            try
            {
                listFormat = captionRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            // Keep the final paragraph mark in normal body flow. Word 2021 rejects
            // Frames.Add on a whole caption paragraph immediately before a user
            // table because the paragraph terminator has table-boundary affinity.
            frameRange = GetNativeCaptionFrameContentRange(captionRange);
            frames = frameRange.Frames;
            frame = frames.Count > 0 ? frames[1] : frames.Add(frameRange);
            const float clippedFrameSize = 0.1f;
            frame.WidthRule = WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = WdFrameSizeRule.wdFrameExact;
            frame.Width = clippedFrameSize;
            frame.Height = clippedFrameSize;
            frame.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = (float)WdFramePosition.wdFrameRight;
            frame.VerticalPosition = (float)WdFramePosition.wdFrameBottom;
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = captionRange.Borders;
            borders.Enable = 0;
            CompressNativeCaptionParagraphFlow(captionRange);
        }
        finally
        {
            Release(borders);
            Release(frame);
            Release(frames);
            Release(frameRange);
            Release(listFormat);
            Release(paragraph);
            Release(numberFont);
            Release(font);
        }
    }

    internal static void StyleNativeCaption(
        Range captionRange,
        Range numberRange,
        bool cleanupLegacyFrames = true)
    {
        AssertCaptionRangeIsOutsideOmml(captionRange);

        Microsoft.Office.Interop.Word.Font? font = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        ParagraphFormat? paragraph = null;
        ListFormat? listFormat = null;
        Range? frameRange = null;
        Frames? frames = null;
        Frame? frame = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        Borders? borders = null;
        try
        {
            // Keep the native SEQ target in ordinary visible formatting. Word's
            // own InsertCrossReference creates a plain REF ... \\h field, which
            // inherits the target formatting on every F9 update. Hiding this
            // paragraph with white 1 pt text therefore made native references
            // white and one point as well.
            font = captionRange.Font;
            font.Hidden = 0;
            font.Size = 11f;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            numberFont = numberRange.Font;
            numberFont.Hidden = 0;
            numberFont.Size = 11f;
            numberFont.Color = WdColor.wdColorAutomatic;
            numberFont.Position = 0;

            paragraph = captionRange.ParagraphFormat;
            paragraph.SpaceBefore = 0f;
            paragraph.SpaceAfter = 0f;
            paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            paragraph.KeepTogether = 0;
            paragraph.KeepWithNext = 0;
            paragraph.PageBreakBefore = 0;
            paragraph.WidowControl = 0;
            try
            {
                listFormat = captionRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            // A legacy Word Frame remains in the main document story, so Word's
            // native Cross-reference dialog and plain REF fields still recognize
            // the SEQ target. Negative frame coordinates are clamped by Word to
            // the top-left of the page, which exposed the number and frame edge.
            // Use an exact 0.1 pt clipping frame at the bottom-right page boundary
            // instead. The SEQ retains normal black 11 pt formatting for native
            // REF inheritance, while its rendered content lies beyond the page.
            var document = captionRange.Document;
            if (cleanupLegacyFrames)
                RemoveLegacyEmptyCaptionFrames(document, captionRange);
            // Frame the SEQ/prefix contents only. The paragraph mark must remain
            // outside the Frame, especially at the body-to-user-table boundary.
            frameRange = GetNativeCaptionFrameContentRange(captionRange);
            frames = frameRange.Frames;
            frame = frames.Count > 0 ? frames[1] : frames.Add(frameRange);
            sections = captionRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            const float clippedFrameSize = 0.1f;
            frame.WidthRule = WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = WdFrameSizeRule.wdFrameExact;
            frame.Width = clippedFrameSize;
            frame.Height = clippedFrameSize;
            frame.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = Math.Max(0f, pageSetup.PageWidth - clippedFrameSize);
            frame.VerticalPosition = Math.Max(0f, pageSetup.PageHeight - clippedFrameSize);
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = captionRange.Borders;
            borders.Enable = 0;
            CompressNativeCaptionParagraphFlow(captionRange);
        }
        finally
        {
            Release(borders);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(frame);
            Release(frames);
            Release(frameRange);
            Release(listFormat);
            Release(paragraph);
            Release(numberFont);
            Release(font);
        }
    }

    private static void CompressNativeCaptionParagraphFlow(Range captionRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        ParagraphFormat? format = null;
        Microsoft.Office.Interop.Word.Font? markFont = null;
        try
        {
            paragraphs = captionRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End <= paragraphRange.Start) return;

            // The SEQ target remains black 11pt text inside its 0.1pt clipping
            // Frame so Word's native REF fields inherit normal character styling.
            // The final paragraph mark sits outside that Frame; at ordinary 11pt
            // single spacing it creates a full visible blank line below every
            // numbered VisualTeX OLE. Compress only that body-flow paragraph.
            format = paragraph.Format;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = 1f;
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
            format.PageBreakBefore = 0;
            format.WidowControl = 0;
            try { format.DisableLineHeightGrid = -1; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

            paragraphMark = paragraphRange.Document.Range(
                paragraphRange.End - 1,
                paragraphRange.End);
            markFont = paragraphMark.Font;
            markFont.Hidden = 0;
            markFont.Size = 1f;
            markFont.Position = 0;
        }
        catch (Exception ignored)
        {
            // The clipping Frame still protects the SEQ target if a protected or
            // custom paragraph refuses the optional flow compaction.
            IgnoredErrors.Ignore(ignored);
        }
        finally
        {
            Release(markFont);
            Release(format);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void AssertCaptionRangeIsOutsideOmml(Range captionRange)
    {
        OMaths? maths = null;
        try
        {
            maths = captionRange.OMaths;
            if (maths.Count > 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to style or frame a caption range inside OMML. Native equation numbering is owned by #(SEQ) and must remain free of external caption formatting.");
        }
        finally
        {
            Release(maths);
        }
    }

    private static void RemoveLegacyEmptyCaptionFrames(
        Document document,
        Range keepRange)
    {
        Frames? frames = null;
        try
        {
            frames = document.Frames;
            for (var index = frames.Count; index >= 1; index--)
            {
                Frame? candidate = null;
                Range? range = null;
                Fields? fields = null;
                try
                {
                    candidate = frames[index];
                    range = candidate.Range;
                    if (range.Start <= keepRange.End && range.End >= keepRange.Start)
                        continue;
                    fields = range.Fields;
                    var text = (range.Text ?? string.Empty)
                        .Trim('\r', '\n', '\t', '\v', ' ');
                    var oldVisualTeXFrame =
                        candidate.HorizontalPosition <= -999f
                        && candidate.VerticalPosition <= -999f
                        && candidate.Width >= 70f
                        && candidate.Height >= 17f;
                    if (oldVisualTeXFrame
                        && fields.Count == 0
                        && string.IsNullOrEmpty(text))
                        candidate.Delete();
                }
                finally
                {
                    Release(fields);
                    Release(range);
                    Release(candidate);
                }
            }
        }
        finally { Release(frames); }
    }

    internal static bool TryGetNativeCaptionRanges(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out Range? captionRange,
        out Range? numberRange)
    {
        captionRange = null;
        numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Field? nativeField = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName)) return false;
            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            captionRange = captionBookmark.Range;
            numberRange = numberBookmark.Range;
            nativeField = FindNativeEquationFieldInRange(
                captionRange,
                nativeSequenceName);
            if (nativeField is null)
                nativeField = FindNativeEquationFieldAtRange(
                    document,
                    numberRange,
                    nativeSequenceName);
            if (nativeField is not null) return true;
            Release(numberRange);
            numberRange = null;
            Release(captionRange);
            captionRange = null;
            return false;
        }
        finally
        {
            Release(nativeField);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }
}

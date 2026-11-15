using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationHeadings;
using static VisualTeX.WordVsto.WordEquationNumberFormat;
using static VisualTeX.WordVsto.WordEquationNumberReferences;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;
using static VisualTeX.WordVsto.WordNumberingState;

namespace VisualTeX.WordVsto;

/// <summary>
/// Updating the SEQ fields of equation numbers, incrementally where it can.
/// </summary>
internal static class WordEquationSequenceFields
{
    internal static void UpdateNativeEquationSequenceFields(Document document)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        UpdateNativeEquationSequenceFields(
            document,
            nativeSequenceName,
            captions,
            formatOnly: false);
        WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
    }

    internal static bool TryUpdateAppendedNativeEquationSequenceField(
        Document document,
        string formulaId,
        out bool numberChanged)
    {
        numberChanged = false;
        var hasTailFormulaId = TryReadNativeEquationTailFormulaId(document, out var tailFormulaId);
        if (!hasTailFormulaId
            || string.Equals(tailFormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
            return false;

        Bookmarks? bookmarks = null;
        Bookmark? tailBookmark = null;
        Bookmark? currentBookmark = null;
        Range? tailRange = null;
        Range? currentRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var tailNumberName = NativeNumberBookmarkName(tailFormulaId);
            var currentNumberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(tailNumberName)
                || !bookmarks.Exists(currentNumberName)
                || !bookmarks.Exists(NativeCaptionBookmarkName(tailFormulaId))
                || !bookmarks.Exists(NativeCaptionBookmarkName(formulaId)))
                return false;

            tailBookmark = bookmarks[tailNumberName];
            currentBookmark = bookmarks[currentNumberName];
            tailRange = tailBookmark.Range;
            currentRange = currentBookmark.Range;
            if (currentRange.Start <= tailRange.Start)
                return false;

            var format = ReadEquationNumberFormat(document);
            if (!TryParseNativeEquationNumber(
                    tailRange.Text,
                    format,
                    out var prefix,
                    out var previousOrdinal))
                return false;

            var ordinal = previousOrdinal + 1;
            if (format.UsesHeading)
            {
                if (!TryResolveAppendHeadingPrefix(
                        document,
                        tailRange.End,
                        currentRange.Start,
                        format,
                        prefix,
                        out prefix,
                        out var startsNewScope))
                    return false;
                if (startsNewScope) ordinal = 1;
            }

            var expectedNumberText = prefix
                + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            var currentNumberText = NormalizeNativeEquationNumberText(currentRange.Text);
            if (!string.Equals(currentNumberText, expectedNumberText, StringComparison.Ordinal))
            {
                UpdateNativeEquationCaptionNumber(
                    document,
                    formulaId,
                    GetNativeEquationSequenceName(document),
                    ordinal,
                    prefix,
                    formatOnly: false,
                    cleanupLegacyFrames: false);
                numberChanged = true;
            }

            WriteNativeEquationTailFormulaId(document, formulaId);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(currentRange);
            Release(tailRange);
            Release(currentBookmark);
            Release(tailBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryParseNativeEquationNumber(
        string? value,
        EquationNumberFormat format,
        out string prefix,
        out int ordinal)
    {
        prefix = string.Empty;
        ordinal = 0;
        var text = NormalizeNativeEquationNumberText(value);
        if (!format.UsesHeading)
            return int.TryParse(
                    text,
                    System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out ordinal)
                && ordinal > 0;

        var separatorIndex = text.LastIndexOf(format.Separator, StringComparison.Ordinal);
        if (separatorIndex <= 0) return false;
        var ordinalText = text.Substring(separatorIndex + format.Separator.Length);
        if (!int.TryParse(
                ordinalText,
                System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture,
                out ordinal)
            || ordinal <= 0)
            return false;
        prefix = text.Substring(0, separatorIndex + format.Separator.Length);
        return true;
    }

    private static bool TryResolveAppendHeadingPrefix(
        Document document,
        int previousNumberEnd,
        int currentNumberStart,
        EquationNumberFormat format,
        string previousPrefix,
        out string prefix,
        out bool startsNewScope)
    {
        prefix = previousPrefix;
        startsNewScope = false;
        if (!format.UsesHeading || currentNumberStart <= previousNumberEnd)
            return true;

        Range? between = null;
        Paragraphs? paragraphs = null;
        try
        {
            between = document.Range(previousNumberEnd, currentNumberStart);
            paragraphs = between.Paragraphs;
            var sawHeadingCandidate = false;
            string? lastExplicitNumber = null;
            var lastExplicitLevel = 0;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Paragraph? paragraph = null;
                Range? range = null;
                ListFormat? listFormat = null;
                Frames? frames = null;
                try
                {
                    paragraph = paragraphs[index];
                    var outlineLevel = (int)paragraph.OutlineLevel;
                    if (outlineLevel < 1 || outlineLevel > format.HeadingLevel) continue;
                    range = paragraph.Range;
                    if (IsWithinTableLayoutFree(range)) continue;
                    try
                    {
                        frames = range.Frames;
                        if (frames.Count > 0) continue;
                    }
                    catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }

                    sawHeadingCandidate = true;
                    var listNumber = string.Empty;
                    try
                    {
                        listFormat = range.ListFormat;
                        listNumber = NormalizeHeadingNumberText(
                            listFormat.ListString,
                            outlineLevel);
                    }
                    catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                    var explicitNumber = !string.IsNullOrWhiteSpace(listNumber)
                        ? listNumber
                        : ParseHeadingNumberFromText(range.Text, outlineLevel);
                    if (string.IsNullOrWhiteSpace(explicitNumber)) continue;
                    lastExplicitNumber = explicitNumber;
                    lastExplicitLevel = outlineLevel;
                }
                finally
                {
                    Release(frames);
                    Release(listFormat);
                    Release(range);
                    Release(paragraph);
                }
            }

            if (!sawHeadingCandidate) return true;
            // If the new interval contains only unnumbered headings, the full
            // document pass is required to know whether they are real synthesized
            // chapter counters or titles skipped by an explicitly numbered scheme.
            if (string.IsNullOrWhiteSpace(lastExplicitNumber)) return false;
            if (lastExplicitLevel < format.HeadingLevel)
            {
                lastExplicitNumber += string.Concat(
                    Enumerable.Repeat(".0", format.HeadingLevel - lastExplicitLevel));
            }
            prefix = lastExplicitNumber + format.Separator;
            startsNewScope = true;
            return true;
        }
        finally
        {
            Release(paragraphs);
            Release(between);
        }
    }

    internal static bool TryUpdateInsertedContinuousEquationFieldRanges(
        Document document,
        string formulaId)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var rangeWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        long rangeCheckpoint = 0;
        void TraceRangeStage(string stage)
        {
            if (rangeWatch is null) return;
            var elapsed = rangeWatch.ElapsedMilliseconds;
            Console.WriteLine(
                $"      [perf] continuous-range.{stage}: +{elapsed - rangeCheckpoint}ms ({elapsed}ms)");
            rangeCheckpoint = elapsed;
        }

        var format = ReadEquationNumberFormat(document);
        if (format.UsesHeading) return false;

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Range? content = null;
        Range? suffixRange = null;
        Fields? suffixFields = null;
        Bookmarks? suffixBookmarks = null;
        Range? prefixRange = null;
        Fields? prefixFields = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName))
                return false;

            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            content = document.Content;
            if (captionRange.End >= content.End) return false;

            // A middle insertion only changes the sequence state after its native
            // caption. Let Word update that suffix twice: pass one recalculates the
            // later SEQ targets, pass two lets visible/body REF fields observe the
            // new target results. This is the same Fields.Update semantics used by
            // full reconciliation, but scoped to the affected half of the story.
            suffixRange = document.Range(captionRange.End, content.End);
            TraceRangeStage("prepare-suffix");
            var hasLaterNumberOwner = false;
            suffixBookmarks = suffixRange.Bookmarks;
            for (var index = 1; index <= suffixBookmarks.Count; index++)
            {
                Bookmark? candidate = null;
                try
                {
                    candidate = suffixBookmarks[index];
                    if (TryFormulaIdFromBookmark(
                            candidate.Name,
                            NativeNumberBookmarkPrefix,
                            out var candidateFormulaId)
                        && !string.Equals(
                            candidateFormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        hasLaterNumberOwner = true;
                        break;
                    }
                }
                finally { Release(candidate); }
            }
            // Never call Range.Fields.Update() here. A numbered document can also
            // contain EMBED/LINK/INCLUDE/DDE or unrelated user fields, and Word may
            // display its external-field confirmation dialog when those are swept
            // into a batch update. Refresh only VisualTeX's own SEQ/REF fields.
            if (UpdateVisualTeXOwnedFieldsInRange(suffixRange) == 0) return false;
            TraceRangeStage("suffix-pass-1");
            UpdateVisualTeXOwnedFieldsInRange(suffixRange);
            TraceRangeStage("suffix-pass-2");

            // References located before the insertion point can target a formula
            // whose number just shifted. Refresh only that prefix once after the
            // suffix SEQ values are stable. Previous formulas' own SEQ fields are
            // unchanged, and Word simply reproduces the same results for them.
            if (captionRange.Start > content.Start)
            {
                prefixRange = document.Range(content.Start, captionRange.Start);
                UpdateVisualTeXOwnedFieldsInRange(prefixRange);
            }
            TraceRangeStage("prefix-pass");
            if (!hasLaterNumberOwner)
                WriteNativeEquationTailFormulaId(document, formulaId);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(prefixFields);
            Release(prefixRange);
            Release(suffixBookmarks);
            Release(suffixFields);
            Release(suffixRange);
            Release(content);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    internal static bool ShouldPreferScopedSequenceSuffixUpdate(Document document)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            // Each healthy numbered VisualTeX formula owns at least three public
            // aliases (plus an OLE/OMML identity). This threshold keeps very small
            // documents on the simple path while avoiding repeated field walks once
            // the document reaches roughly a few dozen formulas.
            return bookmarks.Count >= 96;
        }
        catch { return true; }
        finally { Release(bookmarks); }
    }

    internal static bool TryUpdateInsertedEquationSequenceSuffixWithinScope(
        Document document,
        string formulaId,
        out Dictionary<string, string> changedFormulaNumbers,
        out int referencesAlreadyUpdatedFrom)
    {
        changedFormulaNumbers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        referencesAlreadyUpdatedFrom = -1;
        var format = ReadEquationNumberFormat(document);
        var scopeWatch = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                "1",
                StringComparison.Ordinal)
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var scopeCheckpoint = 0L;
        void TraceScopeStage(string stage)
        {
            if (scopeWatch is null) return;
            var elapsed = scopeWatch.ElapsedMilliseconds;
            Console.WriteLine(
                $"      [perf] continuous-scope.{stage}: +{elapsed - scopeCheckpoint}ms ({elapsed}ms)");
            scopeCheckpoint = elapsed;
        }

        Bookmarks? bookmarks = null;
        Bookmark? currentBookmark = null;
        Range? currentRange = null;
        Range? content = null;
        Range? suffixRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var currentNumberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(currentNumberName)
                || !bookmarks.Exists(NativeCaptionBookmarkName(formulaId)))
                return false;
            currentBookmark = bookmarks[currentNumberName];
            currentRange = currentBookmark.Range;
            if (!TryParseNativeEquationNumber(
                    currentRange.Text,
                    format,
                    out var currentPrefix,
                    out var currentOrdinal))
                return false;

            content = document.Content;
            if (currentRange.End < content.End)
                suffixRange = document.Range(currentRange.End, content.End);
            if (!TryReadHealthyEquationNumberArtifactsFromOpenXml(
                    document,
                    out var healthyEntries,
                    out var knownReferenceCounts))
                return false;
            TraceScopeStage("openxml-inventory");
            var orderedHealthyEntries = healthyEntries
                .OrderBy(entry => entry.Position)
                .ToArray();
            var currentEntryIndex = Array.FindIndex(
                orderedHealthyEntries,
                entry => string.Equals(
                    entry.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase));
            if (currentEntryIndex < 0) return false;
            var allOrderedSuffix = orderedHealthyEntries
                .Skip(currentEntryIndex + 1)
                .ToArray();
            var scopedSuffix = new List<NativeEquationCaptionEntry>();
            foreach (var entry in allOrderedSuffix)
            {
                if (!TryParseNativeEquationNumber(
                        entry.NumberText,
                        format,
                        out var entryPrefix,
                        out _))
                    return false;
                if (format.UsesHeading
                    && !string.Equals(entryPrefix, currentPrefix, StringComparison.Ordinal))
                    break;
                scopedSuffix.Add(entry);
            }
            var orderedSuffix = scopedSuffix.ToArray();
            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                if (!TryParseNativeEquationNumber(
                        orderedSuffix[index].NumberText,
                        format,
                        out var existingPrefix,
                        out var existingOrdinal)
                    || (format.UsesHeading
                        && !string.Equals(existingPrefix, currentPrefix, StringComparison.Ordinal)))
                    return false;
                var expectedOrdinal = currentOrdinal + index + 1;
                // A healthy continuous sequence after one middle insertion is
                // either still on the old ordinal (expected - 1) or has already
                // been refreshed by Word to the new ordinal. Anything else is a
                // structural anomaly and belongs on the conservative full path.
                if (existingOrdinal != expectedOrdinal
                    && existingOrdinal != expectedOrdinal - 1)
                    return false;
            }

            var nativeSequenceName = GetNativeEquationSequenceName(document);
            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                var expectedOrdinal = currentOrdinal + index + 1;
                if (!TryParseNativeEquationNumber(
                        orderedSuffix[index].NumberText,
                        format,
                        out _,
                        out var existingOrdinal))
                    return false;
                if (existingOrdinal == expectedOrdinal - 1)
                {
                    changedFormulaNumbers[orderedSuffix[index].FormulaId] =
                        currentPrefix
                        + expectedOrdinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                }
            }
            TraceScopeStage("plan-suffix");

            // A SEQ field result must be recalculated by Word; replacing that
            // result text directly can destroy the field/bookmark relationship.
            // Inventory the suffix once, update each affected SEQ exactly once,
            // then patch generated REF results directly. Word's Fields.Update on
            // the whole mixed OLE suffix repaginates every EMBED field even when it
            // is locked and is slower than this targeted single pass.
            var suffixHasOnlyGeneratedReferences = false;
            var suffixSequencesUpdated = !format.UsesHeading
                && changedFormulaNumbers.Count > 0
                && suffixRange is not null
                && TryUpdateHealthyVisualTeXSuffixFields(
                    suffixRange,
                    nativeSequenceName,
                    orderedSuffix,
                    changedFormulaNumbers,
                    knownReferenceCounts,
                    out suffixHasOnlyGeneratedReferences);
            TraceScopeStage("update-fields");
            if (format.UsesHeading)
            {
                foreach (var entry in orderedSuffix)
                {
                    if (!changedFormulaNumbers.ContainsKey(entry.FormulaId))
                        continue;
                    if (!TryParseNativeEquationNumber(
                            entry.NumberText,
                            format,
                            out _,
                            out var oldOrdinal))
                        return false;
                    var expectedOrdinal = oldOrdinal + 1;
                    if (!TryUpdateDirectTableSequenceNumber(
                            document,
                            entry.FormulaId,
                            nativeSequenceName,
                            expectedOrdinal,
                            currentPrefix,
                            formatOnly: false,
                            knownBookmarks: bookmarks))
                        return false;
                }
            }
            else if (!suffixSequencesUpdated)
            {
                // A legacy mathematical #(SEQ) host or an incomplete inventory
                // requires the conservative per-formula refresher.
                for (var index = 0; index < orderedSuffix.Length; index++)
                {
                    var expectedOrdinal = currentOrdinal + index + 1;
                    if (!changedFormulaNumbers.ContainsKey(orderedSuffix[index].FormulaId))
                        continue;
                    if (!TryRefreshExistingContinuousCaptionField(
                            document,
                            bookmarks,
                            orderedSuffix[index].FormulaId,
                            nativeSequenceName,
                            expectedOrdinal,
                            currentPrefix))
                        return false;
                }
            }

            var patchedGeneratedReferences = suffixSequencesUpdated;
            if (!patchedGeneratedReferences)
            {
                patchedGeneratedReferences = true;
                foreach (var changed in changedFormulaNumbers)
                {
                    if (TryPatchVisibleEquationNumberResult(
                            bookmarks,
                            changed.Key,
                            changed.Value))
                        continue;
                    patchedGeneratedReferences = false;
                    break;
                }
            }
            TraceScopeStage("patch-visible-results");
            if (patchedGeneratedReferences
                && suffixSequencesUpdated
                && suffixHasOnlyGeneratedReferences)
                referencesAlreadyUpdatedFrom = currentRange.End;

            for (var index = 0; index < orderedSuffix.Length; index++)
            {
                var expectedOrdinal = currentOrdinal + index + 1;
                Bookmark? refreshedNumberBookmark = null;
                Range? refreshedNumberRange = null;
                try
                {
                    var numberName = NativeNumberBookmarkName(orderedSuffix[index].FormulaId);
                    if (!bookmarks.Exists(numberName)) return false;
                    refreshedNumberBookmark = bookmarks[numberName];
                    refreshedNumberRange = refreshedNumberBookmark.Range;
                    if (!TryParseNativeEquationNumber(
                            refreshedNumberRange.Text,
                            format,
                            out var refreshedPrefix,
                            out var refreshedOrdinal)
                        || refreshedOrdinal != expectedOrdinal
                        || (format.UsesHeading
                            && !string.Equals(refreshedPrefix, currentPrefix, StringComparison.Ordinal)))
                        return false;
                }
                finally
                {
                    Release(refreshedNumberRange);
                    Release(refreshedNumberBookmark);
                }
            }
            TraceScopeStage("verify-results");

            if (orderedSuffix.Length == allOrderedSuffix.Length)
                WriteNativeEquationTailFormulaId(
                    document,
                    orderedSuffix.LastOrDefault()?.FormulaId ?? formulaId);
            return true;
        }
        catch
        {
            changedFormulaNumbers.Clear();
            referencesAlreadyUpdatedFrom = -1;
            return false;
        }
        finally
        {
            Release(suffixRange);
            Release(content);
            Release(currentRange);
            Release(currentBookmark);
            Release(bookmarks);
        }
    }

    internal static bool TryReadTargetedRemovalSnapshot(
        Document document,
        string removedFormulaId,
        out IReadOnlyList<NativeEquationCaptionEntry> captions,
        out HashSet<string> externalReferenceFormulaIds)
    {
        captions = Array.Empty<NativeEquationCaptionEntry>();
        externalReferenceFormulaIds = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            var current = GetNativeEquationCaptionEntries(
                document,
                GetNativeEquationSequenceName(document));
            if (!current.Any(item => string.Equals(
                    item.FormulaId,
                    removedFormulaId,
                    StringComparison.OrdinalIgnoreCase)))
                return false;

            // Current direct-SEQ number cells are ordinary SEQ fields; a wdFieldRef
            // here is therefore a real body/cross-reference, not the generated
            // visible number. Enumerate only those REF fields once instead of
            // serializing/parsing the whole 100-OMML document.
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = null;
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldRef) continue;
                code = field.Code;
                if (!TryResolveVisualTeXReferenceBookmark(
                        document,
                        code.Text ?? string.Empty,
                        out var bookmarkName)
                    || !TryFormulaIdFromBookmark(
                        bookmarkName,
                        NativeNumberBookmarkPrefix,
                        out var referencedFormulaId))
                    continue;
                externalReferenceFormulaIds.Add(referencedFormulaId);
            }
            if (externalReferenceFormulaIds.Contains(removedFormulaId))
                return false;
            captions = current;
            return true;
        }
        catch
        {
            captions = Array.Empty<NativeEquationCaptionEntry>();
            externalReferenceFormulaIds.Clear();
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    internal static bool TryUpdateRemovedEquationSequenceSuffixWithinScope(
        Document document,
        string removedFormulaId,
        IReadOnlyList<NativeEquationCaptionEntry> preRemovalCaptions,
        out Dictionary<string, string> changedFormulaNumbers)
    {
        changedFormulaNumbers = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        var removedIndex = -1;
        for (var index = 0; index < preRemovalCaptions.Count; index++)
        {
            if (!string.Equals(
                    preRemovalCaptions[index].FormulaId,
                    removedFormulaId,
                    StringComparison.OrdinalIgnoreCase))
                continue;
            removedIndex = index;
            break;
        }
        if (removedIndex < 0) return false;

        var format = ReadEquationNumberFormat(document);
        if (!TryParseNativeEquationNumber(
                preRemovalCaptions[removedIndex].NumberText,
                format,
                out var removedPrefix,
                out var removedOrdinal))
            return false;

        Bookmarks? bookmarks = null;
        Bookmark? verifyBookmark = null;
        Range? verifyRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var nativeSequenceName = GetNativeEquationSequenceName(document);
            for (var index = removedIndex + 1;
                 index < preRemovalCaptions.Count;
                 index++)
            {
                var item = preRemovalCaptions[index];
                if (!TryParseNativeEquationNumber(
                        item.NumberText,
                        format,
                        out var itemPrefix,
                        out var oldOrdinal))
                    return false;

                if (format.UsesHeading
                    && !string.Equals(
                        itemPrefix,
                        removedPrefix,
                        StringComparison.Ordinal))
                    break;

                var expectedOldOrdinal = removedOrdinal + (index - removedIndex);
                if (oldOrdinal != expectedOldOrdinal)
                    return false;
                var newOrdinal = oldOrdinal - 1;
                var expectedText = itemPrefix
                    + newOrdinal.ToString(
                        System.Globalization.CultureInfo.InvariantCulture);

                UpdateNativeEquationCaptionNumber(
                    document,
                    item.FormulaId,
                    nativeSequenceName,
                    newOrdinal,
                    itemPrefix,
                    formatOnly: false,
                    cleanupLegacyFrames: false,
                    knownBookmarks: bookmarks);

                var numberName = NativeNumberBookmarkName(item.FormulaId);
                if (!bookmarks.Exists(numberName)) return false;
                Release(verifyRange); verifyRange = null;
                Release(verifyBookmark); verifyBookmark = null;
                verifyBookmark = bookmarks[numberName];
                verifyRange = verifyBookmark.Range;
                if (!string.Equals(
                        NormalizeNativeEquationNumberText(verifyRange.Text),
                        expectedText,
                        StringComparison.Ordinal))
                    return false;
                changedFormulaNumbers[item.FormulaId] = expectedText;
            }

            WriteNativeEquationTailFormulaId(
                document,
                preRemovalCaptions
                    .Where(item => !string.Equals(
                        item.FormulaId,
                        removedFormulaId,
                        StringComparison.OrdinalIgnoreCase))
                    .LastOrDefault()?.FormulaId);
            return true;
        }
        catch
        {
            changedFormulaNumbers.Clear();
            return false;
        }
        finally
        {
            Release(verifyRange);
            Release(verifyBookmark);
            Release(bookmarks);
        }
    }

    internal static bool TryPatchVisibleEquationNumberResult(
        Bookmarks bookmarks,
        string formulaId,
        string expectedNumber)
    {
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        Document? visibleDocument = null;
        OMaths? maths = null;
        Tables? tables = null;
        Table? table = null;
        Fields? directFields = null;
        Field? directField = null;
        Range? directCode = null;
        try
        {
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return false;
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;
            visibleDocument = visibleRange.Document;
            maths = visibleRange.OMaths;

            // Current OMML numbering owns its visible label directly in cell
            // (1,3): TAB + ordinary SEQ + paragraph mark. There is no generated
            // REF result to patch. Prove that exact host locally and accept the
            // already-refreshed SEQ result instead of falling through to the
            // document-wide Fields scan (which made five numbers take ~1 s).
            tables = visibleRange.Tables;
            if (maths.Count == 0 && tables.Count > 0)
            {
                if (tables.Count == 1)
                {
                    table = tables[1];
                    if (TryGetManagedNumberTableRowIndex(table, visibleRange, 3, out _))
                    {
                        directFields = visibleRange.Fields;
                        if (directFields.Count == 1)
                        {
                            directField = directFields[1];
                            directCode = directField.Code;
                            if (IsVisualTeXSequenceFieldCode(directCode.Text))
                                return string.Equals(
                                    NormalizeNativeEquationNumberText(visibleRange.Text),
                                    expectedNumber,
                                    StringComparison.Ordinal);
                        }
                    }
                }
            }

            if (maths.Count > 0
                || HasManagedNativeOmmlHashSequenceHost(
                    visibleDocument,
                    formulaId))
            {
                // VTEq_ is an alias around the result of the mathematical SEQ.
                // Never replace that result text directly: Word must update the
                // unchanged field itself, or the entire numbered OMath must be
                // rebuilt when its format instruction changes.
                return string.Equals(
                    NormalizeNativeEquationNumberText(visibleRange.Text),
                    expectedNumber,
                    StringComparison.Ordinal);
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(directCode);
            Release(directField);
            Release(directFields);
            Release(table);
            Release(tables);
            Release(maths);
            Release(visibleDocument);
            Release(visibleRange);
            Release(visibleBookmark);
        }

        return TryPatchVisibleEquationNumberResultCore(
            bookmarks,
            formulaId,
            expectedNumber);
    }

    private static bool TryPatchVisibleEquationNumberResultCore(
        Bookmarks bookmarks,
        string formulaId,
        string expectedNumber)
    {
        // VTEqNum_* for current native OMML encloses the live mathematical
        // SEQ result. Never replace Range.Text there: doing so unlinks/destroys
        // the field. Recalculate the unchanged field and let ordinary body REF
        // fields read the same bookmark. Prefix/field-switch changes are handled
        // by an atomic OMath rebuild before this result-refresh path is entered.
        Bookmark? nativeHashBookmark = null;
        Range? nativeHashSequenceRange = null;
        OMaths? nativeHashMaths = null;
        Fields? nativeHashFields = null;
        try
        {
            var nativeHashName = NativeNumberBookmarkName(formulaId);
            if (bookmarks.Exists(nativeHashName))
            {
                nativeHashBookmark = bookmarks[nativeHashName];
                nativeHashSequenceRange = nativeHashBookmark.Range;
                nativeHashMaths = nativeHashSequenceRange.OMaths;
                if (nativeHashMaths.Count > 0)
                {
                    nativeHashFields = nativeHashSequenceRange.Fields;
                    for (var fieldIndex = 1;
                         fieldIndex <= nativeHashFields.Count;
                         fieldIndex++)
                    {
                        Field? nativeHashField = null;
                        Range? nativeHashCode = null;
                        try
                        {
                            nativeHashField = nativeHashFields[fieldIndex];
                            nativeHashCode = nativeHashField.Code;
                            if (!IsVisualTeXSequenceFieldCode(nativeHashCode.Text))
                                continue;
                            nativeHashField.Update();
                            return string.Equals(
                                NormalizeNativeEquationNumberText(
                                    nativeHashSequenceRange.Text),
                                expectedNumber,
                                StringComparison.Ordinal);
                        }
                        finally
                        {
                            Release(nativeHashCode);
                            Release(nativeHashField);
                        }
                    }
                    return false;
                }
            }
        }
        finally
        {
            Release(nativeHashFields);
            Release(nativeHashMaths);
            Release(nativeHashSequenceRange);
            Release(nativeHashBookmark);
        }


        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        try
        {
            var bookmarkName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            fields = range.Fields;
            var targetBookmarkName = NativeNumberBookmarkName(formulaId);
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsReferenceToBookmark(code.Text, targetBookmarkName))
                        continue;
                    result = field.Result;
                    if (!string.Equals(
                            NormalizeNativeEquationNumberText(result.Text),
                            expectedNumber,
                            StringComparison.Ordinal))
                        result.Text = expectedNumber;
                    if (range.StoryType == WdStoryType.wdMainTextStory)
                        ApplyParagraphEquationNumberFont(
                            range,
                            fallbackSize: 0f,
                            position: 0);
                    return true;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(fields);
            Release(range);
            Release(bookmark);
        }
    }

    private static bool TryUpdateHealthyVisualTeXSuffixFields(
        Range suffixRange,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> orderedSuffix,
        IReadOnlyDictionary<string, string> changedFormulaNumbers,
        IReadOnlyDictionary<string, int> knownReferenceCounts,
        out bool suffixHasOnlyGeneratedReferences)
    {
        suffixHasOnlyGeneratedReferences = false;
        Document? document = null;
        Fields? fields = null;
        var sequenceFieldCount = 0;
        var visualReferenceCount = 0;
        var externalReferenceCount = 0;
        var hasExternalReferences = false;
        var sequenceFieldOffset = -1;
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        bool RejectInventory(string reason)
        {
            if (tracePerformance)
                Console.WriteLine($"      [perf] continuous-suffix.inventory-rejected reason={reason}");
            return false;
        }
        try
        {
            foreach (var changedFormulaId in changedFormulaNumbers.Keys)
            {
                if (!knownReferenceCounts.TryGetValue(
                        changedFormulaId,
                        out var referenceCount))
                    return RejectInventory("missing-reference-count");
                // The healthy OpenXML inventory normalizes both OLE-generated REF
                // labels and direct-table/native visible slots to a baseline count
                // of one. Anything above one proves a real body reference exists.
                if (referenceCount <= 1) continue;
                hasExternalReferences = true;
                externalReferenceCount += referenceCount - 1;
            }

            document = suffixRange.Document;
            fields = suffixRange.Fields;
            if (fields.Count == 0) return RejectInventory("no-fields");

            // Inspect and update in the same forward walk. Materializing this mixed
            // field collection a second time is itself expensive in Word, while
            // Range.Fields.Update repaginates every EMBED even when it is locked.
            // Only owned SEQ fields are refreshed; REF ownership is inventoried so
            // their already-known visible results can be patched safely afterward.
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? sequenceResult = null;
                Range? referenceResult = null;
                try
                {
                    field = fields[index];
                    var isOwnedSequence = false;
                    if (field.Type == WdFieldType.wdFieldRef)
                    {
                        visualReferenceCount++;
                        code = field.Code;
                        if (TryResolveVisualTeXReferenceBookmark(
                                document,
                                code.Text ?? string.Empty,
                                out var bookmarkName)
                            && TryFormulaIdFromBookmark(
                                bookmarkName,
                                NativeNumberBookmarkPrefix,
                                out var referencedFormulaId)
                            && changedFormulaNumbers.TryGetValue(
                                referencedFormulaId,
                                out var expectedNumber))
                        {
                            referenceResult = field.Result;
                            if (!string.Equals(
                                    NormalizeNativeEquationNumberText(referenceResult.Text),
                                    expectedNumber,
                                    StringComparison.Ordinal))
                                referenceResult.Text = expectedNumber;
                        }
                        continue;
                    }
                    if (field.Type == WdFieldType.wdFieldEmbed) continue;
                    code = field.Code;
                    var codeText = code.Text ?? string.Empty;
                    isOwnedSequence = IsVisualTeXSequenceFieldCode(codeText);

                    if (isOwnedSequence)
                    {
                        if (field.Locked) return RejectInventory("locked-sequence");
                        sequenceFieldCount++;
                        sequenceResult = field.Result;
                        if (sequenceFieldOffset < 0)
                        {
                            sequenceFieldOffset = orderedSuffix.Count > 0
                                && sequenceResult.Start < orderedSuffix[0].Position
                                ? 1
                                : 0;
                        }
                        var entryIndex = sequenceFieldCount - 1 - sequenceFieldOffset;
                        if (entryIndex < 0)
                            continue;
                        if (entryIndex >= orderedSuffix.Count)
                            return RejectInventory("unexpected-extra-sequence");

                        var entry = orderedSuffix[entryIndex];
                        if (!changedFormulaNumbers.TryGetValue(
                                entry.FormulaId,
                                out var expectedNumber))
                            continue;
                        if (!int.TryParse(
                                expectedNumber,
                                System.Globalization.NumberStyles.None,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out _))
                            return RejectInventory("non-continuous-number");

                        if (Regex.IsMatch(
                                codeText,
                                @"\\r\s+\d+",
                                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                        {
                            if (IsFieldInsideOmml(field))
                                return RejectInventory("mathematical-sequence-host");
                            EnsureSequenceFieldCodeCanBeRewritten(field);
                            code!.Text = Regex.Replace(
                                codeText,
                                @"\\r\s+\d+",
                                "\\r " + expectedNumber,
                                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                        }
                        field.Update();
                    }
                }
                finally
                {
                    Release(referenceResult);
                    Release(sequenceResult);
                    Release(code);
                    Release(field);
                }
            }
            // Word can expose the field whose result ends exactly at
            // suffixRange.Start as the first member of Range.Fields. Re-updating
            // that already-correct current SEQ is harmless; any other count means
            // the suffix inventory is incomplete or contains an alien sequence.
            if (sequenceFieldCount != orderedSuffix.Count
                && sequenceFieldCount != orderedSuffix.Count + 1)
                return RejectInventory(
                    $"sequence-count-{sequenceFieldCount}-expected-{orderedSuffix.Count}");
            suffixHasOnlyGeneratedReferences = !hasExternalReferences;
            if (tracePerformance)
            {
                Console.WriteLine(
                    $"      [perf] continuous-suffix.inventory fields={fields.Count} "
                    + $"sequences={sequenceFieldCount} refs={visualReferenceCount} "
                    + $"externalRefs={externalReferenceCount} sequence={nativeSequenceName}");
            }
            return true;
        }
        catch (Exception error)
        {
            if (tracePerformance)
                Console.WriteLine(
                    $"      [perf] continuous-suffix.inventory-rejected reason=exception-{error.Message}");
            suffixHasOnlyGeneratedReferences = false;
            return false;
        }
        finally
        {
            Release(fields);
            Release(document);
        }
    }

    private static bool TryRefreshExistingContinuousCaptionField(
        Document document,
        Bookmarks bookmarks,
        string formulaId,
        string nativeSequenceName,
        int expectedOrdinal,
        string expectedPrefix)
    {
        // Current numbered OMML uses a proven 1x3 direct-table SEQ. Update its
        // explicit reset ordinal directly; this also gives ordinary OLE captions a
        // cheap bookmark-topology rejection without reading CustomXML/WordOpenXML.
        if (TryUpdateDirectTableSequenceNumber(
                document,
                formulaId,
                nativeSequenceName,
                expectedOrdinal,
                expectedPrefix,
                formatOnly: false,
                knownBookmarks: bookmarks,
                trustedHealthyDirectTable: true))
        {
            Bookmark? directNumberBookmark = null;
            Range? directNumberRange = null;
            try
            {
                var directNumberName = NativeNumberBookmarkName(formulaId);
                if (!bookmarks.Exists(directNumberName)) return false;
                directNumberBookmark = bookmarks[directNumberName];
                directNumberRange = directNumberBookmark.Range;
                return string.Equals(
                    NormalizeNativeEquationNumberText(directNumberRange.Text),
                    expectedOrdinal.ToString(
                        System.Globalization.CultureInfo.InvariantCulture),
                    StringComparison.Ordinal);
            }
            finally
            {
                Release(directNumberRange);
                Release(directNumberBookmark);
            }
        }

        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Range? captionRange = null;
        Range? numberRange = null;
        OMaths? captionMaths = null;
        Field? field = null;
        Range? code = null;
        try
        {
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName))
                return false;

            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            captionMaths = captionRange.OMaths;
            if (captionMaths.Count > 0)
            {
                // Retired mathematical #(SEQ) hosts are uncommon but must never
                // have Field.Code.Text rewritten in place. Their refresher performs
                // the full structural verification and atomic rebuild when needed.
                if (!TryRefreshNumberedNativeOmmlHashSequence(
                        document,
                        formulaId,
                        expectedPrefix))
                    return false;
                numberBookmark = bookmarks[numberName];
                numberRange = numberBookmark.Range;
                return string.Equals(
                    NormalizeNativeEquationNumberText(numberRange.Text),
                    expectedPrefix + expectedOrdinal.ToString(
                        System.Globalization.CultureInfo.InvariantCulture),
                    StringComparison.Ordinal);
            }
            field = FindNativeEquationFieldInRange(captionRange, nativeSequenceName);
            if (field is null) return false;

            code = field.Code;
            var codeText = code.Text ?? string.Empty;
            if (!IsNativeEquationSequenceFieldCode(codeText, nativeSequenceName))
                return false;

            // Older structural passes could freeze a continuous SEQ with \\r N.
            // Restore normal Word sequence semantics before the local update so
            // later insertions can again be recalculated without rewriting the
            // field or its bookmarks.
            if (Regex.IsMatch(
                    codeText,
                    @"\\r\s+\d+",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                EnsureSequenceFieldCodeCanBeRewritten(field);
                code.Text = $" SEQ {LegacyEquationSequenceName} \\* ARABIC ";
            }
            field.Update();

            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            var expectedText = expectedOrdinal.ToString(
                System.Globalization.CultureInfo.InvariantCulture);
            return string.Equals(
                NormalizeNativeEquationNumberText(numberRange.Text),
                expectedText,
                StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(captionMaths);
            Release(numberRange);
            Release(captionRange);
            Release(numberBookmark);
            Release(captionBookmark);
        }
    }

    internal static Dictionary<string, string> UpdateNativeEquationSequenceFieldsIncremental(
        Document document)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var captions = GetNativeEquationCaptionEntries(document, nativeSequenceName);
        return UpdateNativeEquationSequenceFieldsIncremental(
            document,
            nativeSequenceName,
            captions);
    }

    internal static Dictionary<string, string> UpdateNativeEquationSequenceFieldsIncremental(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions)
    {
        return UpdateNativeEquationSequenceFieldsIncremental(
            document,
            nativeSequenceName,
            captions,
            ReadEquationNumberFormat(document));
    }

    internal static Dictionary<string, string> UpdateNativeEquationSequenceFieldsIncremental(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions,
        EquationNumberFormat format,
        bool trustedHealthyDirectTables = false)
    {
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchorsForFormatBatch(
                document,
                format.HeadingLevel,
                captions)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();
        var changedFormulaNumbers = new Dictionary<string, string>(
            StringComparer.OrdinalIgnoreCase);
        Bookmarks? sharedBookmarks = null;
        try
        {
            sharedBookmarks = document.Bookmarks;

            foreach (var caption in captions)
            {
                var scope = ResolveEquationNumberScope(
                    caption.Position,
                    format,
                    headingAnchors);
                ordinalByScope.TryGetValue(scope.ScopePosition, out var localOrdinal);
                localOrdinal++;
                ordinalByScope[scope.ScopePosition] = localOrdinal;

                var expectedNumberText = scope.Prefix
                    + localOrdinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (string.Equals(
                        NormalizeNativeEquationNumberText(caption.NumberText),
                        expectedNumberText,
                        StringComparison.Ordinal))
                    continue;

                UpdateNativeEquationCaptionNumber(
                    document,
                    caption.FormulaId,
                    nativeSequenceName,
                    localOrdinal,
                    scope.Prefix,
                    formatOnly: false,
                    cleanupLegacyFrames: false,
                    knownBookmarks: sharedBookmarks,
                    trustedHealthyDirectTable:
                        trustedHealthyDirectTables);
                changedFormulaNumbers[caption.FormulaId] = expectedNumberText;
            }

            WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
            return changedFormulaNumbers;
        }
        finally { Release(sharedBookmarks); }
    }

    internal static bool IsFieldInsideOmml(Field field)
    {
        if (field is null) return false;
        Range? code = null;
        OMaths? codeMaths = null;
        try
        {
            // Rewriting safety is determined by the field instruction, not by its
            // rendered result. A perfectly ordinary body REF can target VTEqNum
            // inside an OMath and Word may then expose Field.Result.OMaths.Count>0;
            // that does not make the REF field itself mathematical. Conversely a
            // SEQ/REF whose Code is inside professional OMML must never have
            // Field.Code.Text assigned in place.
            code = field.Code;
            codeMaths = code.OMaths;
            return codeMaths.Count > 0;
        }
        catch
        {
            // An inaccessible field code is never safe for in-place rewriting.
            // Conservatively keep it on structural migration paths.
            return true;
        }
        finally
        {
            Release(codeMaths);
            Release(code);
        }
    }

    internal static void EnsureSequenceFieldCodeCanBeRewritten(Field field)
    {
        if (field is null) throw new ArgumentNullException(nameof(field));
        if (IsFieldInsideOmml(field))
            throw new InvalidOperationException(
                "VisualTeX refused to rewrite Field.Code.Text inside an OMML equation. Native #(SEQ) numbering must be rebuilt atomically or updated with F9/Field.Update only.");
    }

    internal static void EnsureReferenceFieldCodeCanBeRewritten(Field field)
    {
        if (field is null) throw new ArgumentNullException(nameof(field));
        if (IsFieldInsideOmml(field))
            throw new InvalidOperationException(
                "VisualTeX refused to rewrite a REF field inside OMML. Legacy #(REF) numbering must be migrated by atomically replacing the complete equation.");
    }

    internal static string NormalizeNativeEquationNumberText(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\n", string.Empty)
            .Replace("\a", string.Empty)
            .Trim();

    internal static void UpdateNativeEquationSequenceFields(
        Document document,
        string nativeSequenceName,
        IReadOnlyList<NativeEquationCaptionEntry> captions,
        bool formatOnly)
    {
        var format = ReadEquationNumberFormat(document);
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchors(document, format.HeadingLevel)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();

        foreach (var caption in captions)
        {
            var scope = ResolveEquationNumberScope(
                caption.Position,
                format,
                headingAnchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var localOrdinal);
            localOrdinal++;
            ordinalByScope[scope.ScopePosition] = localOrdinal;
            UpdateNativeEquationCaptionNumber(
                document,
                caption.FormulaId,
                nativeSequenceName,
                localOrdinal,
                scope.Prefix,
                formatOnly);
        }
        WriteNativeEquationTailFormulaId(document, captions.LastOrDefault()?.FormulaId);
    }
}

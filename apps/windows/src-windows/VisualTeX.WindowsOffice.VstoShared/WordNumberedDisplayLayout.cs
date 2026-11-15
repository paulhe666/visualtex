using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationNumberFormat;
using static VisualTeX.WordVsto.WordEquationParagraphLayout;
using static VisualTeX.WordVsto.WordEquationSequenceFields;
using static VisualTeX.WordVsto.WordNativeCaptions;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberedTypingParagraphs;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;
using static VisualTeX.WordVsto.WordVisibleEquationNumbers;

namespace VisualTeX.WordVsto;

/// <summary>
/// Laying out numbered and unnumbered display formulas: tab layouts, display shapes, numbered tables and OLE
/// baselines.
/// </summary>
internal static class WordNumberedDisplayLayout
{
    // A table cell is already a paragraph; a lone OLE is not evidence of
    // display mode. Require the native display tab scaffold as well.
    internal static bool IsTableMathTypeDisplayLayout(
        string? shapePrefixText,
        bool hasCenterTab,
        bool hasRightTab) =>
        string.Equals(shapePrefixText, "\t", StringComparison.Ordinal)
        && hasCenterTab
        && hasRightTab;

    internal static int RefreshNumberedOmmlTabLayouts(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        // Complete VTEq_/VTEqCap_/VTEqNum_ loss makes the normal structural
        // resolver reject an otherwise valid managed #(SEQ) OMath. Repair those
        // identities first from VTOMML + semantic fingerprint; the ordinary loop
        // below can then use its strict FormulaId-bound range checks unchanged.
        TryRepairDriftedManagedNativeHashSequenceHosts(document, out _);

        var refreshed = 0;
        var formulaIds = WordOmmlFormulaStore.BookmarkedFormulaIds(document);
        foreach (var formulaId in formulaIds)
        {
            Range? formulaRange = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata is null
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                formulaRange = WordOmmlFormulaStore
                    .GetEquationRangeVerifiedForStructuralEdit(
                        document,
                        formulaId,
                        metadata);
                if (!metadata.Numbered)
                {
                    // Older builds could dismantle a direct-SEQ 1x3 into a valid
                    // standalone Display OMath while accidentally retaining the
                    // adjacent VisualTeX table separator's Exactly-1pt paragraph
                    // metrics. Reuse this existing one-time document-open scan to
                    // self-heal only that unmistakable managed-OMML signature.
                    if (NormalizeCompactStandaloneNativeOmmlParagraph(formulaRange))
                        refreshed++;
                    continue;
                }

                // Bulk OMML insertion may persist adjacent numbered formulas as one
                // managed N x 3 table. Refresh the exact owning row in place. Do not
                // enter the legacy rebuild path, which first deletes this row's SEQ
                // aliases and historically left one random row empty after reopen.
                if (IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    refreshed++;
                    continue;
                }

                ConfigureNumberedDisplayFormula(
                    document,
                    formulaRange,
                    WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                    (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                    formulaId,
                    reuseExistingScaffold: true,
                    metadata: metadata,
                    deferNativeOmmlShapeCreation: true);
                // Repair documents produced by builds that retained the mandatory
                // separator between adjacent 1x3 tables at ordinary 10.5/12pt line
                // metrics. The paragraph cannot be deleted without merging the two
                // tables, but it can be made visually negligible and remains outside
                // every FormulaId/field/bookmark identity.
                CompactManagedNativeOmmlTableSeparatorBefore(document, formulaId);
                CleanupGeneratedNativeOmmlTypingTailBeforeFollowingTable(
                    document,
                    formulaId);
                refreshed++;
            }
            finally { Release(formulaRange); }
        }
        return refreshed;
    }

    internal static int FinalizeNumberedOmmlDisplayShapeLayouts(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        return FinalizeConvertedNumberedOmmlDisplayShapes(
            document,
            WordOmmlFormulaStore.BookmarkedFormulaIds(document));
    }

    internal static int FinalizeConvertedNumberedOmmlDisplayShapes(
        Document document,
        IReadOnlyList<string> formulaIds)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (formulaIds is null) throw new ArgumentNullException(nameof(formulaIds));
        var finalized = 0;
        foreach (var formulaId in formulaIds
                     .Where(id => !string.IsNullOrWhiteSpace(id))
                     .Distinct(StringComparer.OrdinalIgnoreCase))
        {
            FormulaMetadata? metadata = null;
            Range? formulaRange = null;
            try
            {
                metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata is null
                    || !metadata.Numbered
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                formulaRange = WordOmmlFormulaStore
                    .GetEquationRangeVerifiedForStructuralEdit(
                        document,
                        formulaId,
                        metadata);

                if (IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    finalized++;
                    continue;
                }

                if (!IsNativeHashSequenceHostForFinalizationV2(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    // This compatibility entry is still called after batch insert,
                    // document-open migration and old add-in dispatcher turns. It
                    // may repair/migrate a legacy host, but it must never create a
                    // floating object. ConfigureNumberedNativeOmmlDisplay now has
                    // only the native #() producer; old Shape/table/REF eqArr forms
                    // are consumed as migration input.
                    ConfigureNumberedNativeOmmlDisplay(
                        document,
                        formulaRange,
                        WordOmmlFormulaStore.EstimateHeightPoints(formulaRange),
                        (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                        formulaId,
                        reuseExistingScaffold: true,
                        metadata,
                        _ => { },
                        deferFieldUpdate: false,
                        deferExternalShapeCreation: true);
                    Release(formulaRange);
                    formulaRange = WordOmmlFormulaStore
                        .GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            formulaId,
                            metadata);
                }

                if (IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                {
                    finalized++;
                    continue;
                }
                if (!IsNativeHashSequenceHostForFinalizationV2(
                        document,
                        formulaRange,
                        formulaId,
                        updateField: true))
                    continue;

                // A just-migrated Shape host can expose a healthy native #(SEQ)
                // equation one Office turn before Word physically removes the old
                // one-point anchor paragraph. Word may also remove VTEqAnc_ while
                // retaining that empty paragraph, so completion checks both the
                // bookmark and the retired paragraph's exact structural/style
                // fingerprint before reporting the formula finalized.
                if (!EnsureRetiredNativeDisplayAnchorParagraphRemoved(
                        document,
                        formulaRange,
                        formulaId))
                    continue;
                ClearNativeDisplayAnchorCommitMarker(document, formulaId);
                finalized++;
            }
            finally
            {
                Release(formulaRange);
            }
        }
        return finalized;
    }

    private static void ScrollRangeIntoLayoutView(Window? window, Range range)
    {
        if (window is null) return;
        try
        {
            object start = true;
            window.ScrollIntoView(range, ref start);
        }
        catch
        {
            // Background/protected windows can reject scrolling. Width/geometry
            // readers retain their conservative fallback behavior in that case.
        }
    }

    private static bool HasNumberedOmmlLeftFormulaTab(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            tabStops = format.TabStops;
            var hasLeft = false;
            var hasRight = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                if (tabStop.Alignment == WdTabAlignment.wdAlignTabLeft)
                    hasLeft = true;
                else if (tabStop.Alignment == WdTabAlignment.wdAlignTabRight)
                    hasRight = true;
            }
            return hasLeft && hasRight;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static FormulaMetadata? ReadMetadata(InlineShape shape) =>
        WordFormulaMetadataReader.TryRead(shape);

    internal static bool ConfigureNumberedDisplayFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool reuseExistingScaffold = false,
        Table? knownNumberedTable = null,
        bool useConversionSafeVisibleNumber = false,
        FormulaMetadata? metadata = null,
        bool deferNativeOmmlShapeCreation = false,
        bool deferNativeOmmlMetadataPersistence = false,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool deferFieldUpdate = false,
        bool verifiedFreshBareOmml = false,
        string? preparedFreshBareOmml = null,
        bool knownFreshConversionTarget = false)
    {
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
                $"[perf] ConfigureNumbered.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }

        if (ContainsNativeOmml(formulaRange))
        {
            return ConfigureNumberedNativeOmmlDisplay(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                formulaId,
                reuseExistingScaffold,
                metadata,
                TraceStage,
                plannedOrdinal: plannedOrdinal,
                plannedPrefix: plannedPrefix,
                deferExternalShapeCreation: deferNativeOmmlShapeCreation,
                deferMetadataPersistence:
                    deferNativeOmmlMetadataPersistence,
                verifiedFreshBareOmml:
                    verifiedFreshBareOmml,
                preparedFreshBareOmml:
                    preparedFreshBareOmml);
        }

        Range? migratedFormulaRange = null;
        try
        {
            // Numbered VisualTeX OLE retains the accepted center/right-tab
            // paragraph contract. Native OMML has already branched to Word's own
            // display-equation number slot above. Migrate only the exact managed
            // legacy OLE table; arbitrary user tables remain untouched.
            if (IsManagedNumberedEquationTable(document, formulaRange, formulaId))
            {
                // A legacy 1x3/2x3 table is migration input only. Normalize the
                // known empty-row quirk, then require a complete conversion to the
                // final ordinary center/right-tab paragraph before doing anything
                // else. Retaining or repairing the table is no longer a supported
                // fallback.
                TrimBenignEmptyRowsFromNumberedTable(document, formulaRange, formulaId);
                migratedFormulaRange = TryConvertStandardNumberedOleTableToTabParagraph(
                    document,
                    formulaRange,
                    formulaId);
                if (migratedFormulaRange is null)
                    throw new InvalidOperationException(
                        "VisualTeX could not safely migrate this managed numbered OLE table to the table-free tab layout.");
                knownNumberedTable = null;
                TraceStage("migrate-ole-table-to-tabs");
            }

            var activeFormulaRange = migratedFormulaRange ?? formulaRange;
            if (IsManagedNumberedEquationTable(document, activeFormulaRange, formulaId))
                throw new InvalidOperationException(
                    "A numbered VisualTeX OLE formula remained inside a Word table after migration.");

            ConfigureEquationParagraph(activeFormulaRange, numbered: true);
            EnsureLeadingEquationTab(document, activeFormulaRange);
            TraceStage("tab-paragraph");

            var sequenceName = GetNativeEquationSequenceName(document);
            EnsureNativeCaption(
                document,
                activeFormulaRange,
                formulaId,
                sequenceName,
                restyleExisting: !reuseExistingScaffold,
                knownNumberedTable: null,
                deferFieldUpdate: deferFieldUpdate,
                plannedOrdinal: plannedOrdinal,
                plannedPrefix: plannedPrefix,
                knownFreshConversionTarget: knownFreshConversionTarget);
            TraceStage("native-caption");
            var visibleNumberCreated = EnsureVisibleEquationNumber(
                document,
                activeFormulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                formulaId,
                adoptExistingTableReference: false,
                knownNumberedTable: null,
                useConversionSafeVisibleNumber,
                deferFieldUpdate,
                knownFreshConversionTarget);
            TraceStage("visible-ref");

            if (metadata is not null && !ContainsNativeOmml(activeFormulaRange))
            {
                AlignDisplayOleObjectToParagraphBaseline(
                    document,
                    activeFormulaRange,
                    formulaHeightPoints,
                    metadata);
                TraceStage("ole-baseline");
            }
            return visibleNumberCreated;
        }
        finally { Release(migratedFormulaRange); }
    }

    internal static Range ResolveSingleNativeOmmlRange(Range formulaRange)
    {
        Document? document = null;
        Range? firstCharacter = null;
        Range? probe = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? range = null;
        try
        {
            document = formulaRange.Document;
            var start = formulaRange.Start;
            var end = formulaRange.End;

            // In the MathType-compatible numbering host Word tracks a Range across
            // text inserted immediately before it, so an old RCW can become
            // "TAB + OMath" even though the OMath itself still starts one character
            // later. Probe after that ordinary layout tab before asking Word for the
            // equation; otherwise OMath.Range is clipped to the caller's start and
            // VTOMML becomes anchored to the paragraph instead of the formula.
            if (start < end)
            {
                firstCharacter = document.Range(start, start + 1);
                if (string.Equals(firstCharacter.Text, "\t", StringComparison.Ordinal))
                    start++;
            }
            probe = document.Range(start, Math.Max(start, end));
            maths = probe.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    $"The managed OMML host contains {maths.Count} equations instead of exactly one.");
            math = maths[1];
            range = math.Range.Duplicate;
            var result = range;
            range = null;
            return result;
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
            Release(probe);
            Release(firstCharacter);
            Release(document);
        }
    }

    internal static Field? FindNativeOmmlNumberReferenceField(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        Field? positionalMatch = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleBookmarkName = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(visibleBookmarkName))
            {
                visibleBookmark = bookmarks[visibleBookmarkName];
                visibleRange = visibleBookmark.Range;
            }
            fields = formulaRange.Fields;
            var formulaText = formulaRange.Text ?? string.Empty;
            var separatorOffset = formulaText.LastIndexOf('#');
            var separatorPosition = separatorOffset >= 0
                ? formulaRange.Start + separatorOffset
                : formulaRange.Start;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? candidate = null;
                Range? candidateCode = null;
                Range? candidateResult = null;
                try
                {
                    candidate = fields[index];
                    candidateCode = candidate.Code;
                    candidateResult = candidate.Result;
                    if (IsReferenceToBookmark(candidateCode.Text, targetBookmarkName)
                        || (visibleRange is not null
                            && candidateResult.Start <= visibleRange.Start
                            && candidateResult.End >= visibleRange.End))
                    {
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    if (candidateResult.Start > separatorPosition)
                    {
                        Release(positionalMatch);
                        positionalMatch = candidate;
                        candidate = null;
                    }
                }
                finally
                {
                    Release(candidateResult);
                    Release(candidateCode);
                    Release(candidate);
                }
            }
            var fallback = positionalMatch;
            positionalMatch = null;
            return fallback;
        }
        finally
        {
            Release(positionalMatch);
            Release(visibleRange);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(fields);
        }
    }

    private static void RemoveOrdinaryTabsAfterNativeOmml(
        Document document,
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? probe = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var end = Math.Max(formulaRange.End, paragraphRange.End - 1);
            for (var position = end - 1; position >= formulaRange.End; position--)
            {
                Release(probe);
                probe = document.Range(position, position + 1);
                if (string.Equals(probe.Text, "\t", StringComparison.Ordinal))
                    probe.Delete();
            }
        }
        finally
        {
            Release(probe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void TrimBenignEmptyRowsFromNumberedTable(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Row? row = null;
        Range? rowRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null) return;
            columns = table.Columns;
            rows = table.Rows;
            if (columns.Count != 3 || rows.Count <= 1) return;

            var formulaRow = 0;
            for (var index = 1; index <= rows.Count; index++)
            {
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                if (formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End)
                {
                    if (formulaRow != 0) return;
                    formulaRow = index;
                    continue;
                }

                Release(shapes);
                shapes = rowRange.InlineShapes;
                if (shapes.Count != 0) return;
                Release(maths);
                maths = rowRange.OMaths;
                if (maths.Count != 0) return;
                Release(fields);
                fields = rowRange.Fields;
                if (fields.Count != 0) return;
                Release(bookmarks);
                bookmarks = rowRange.Bookmarks;
                if (bookmarks.Count != 0) return;
                foreach (var character in rowRange.Text ?? string.Empty)
                {
                    if (character is '\r' or '\a' or '\n' or '\v'
                        or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                        || char.IsWhiteSpace(character))
                        continue;
                    return;
                }
            }
            if (formulaRow == 0) return;

            for (var index = rows.Count; index >= 1; index--)
            {
                if (index == formulaRow) continue;
                Release(row);
                row = rows[index];
                row.Delete();
            }
        }
        catch
        {
            // This is a defensive cleanup for a Word table-conversion quirk.
            // Numbering itself is already durable, so never make formula creation
            // fail solely because an extra empty row could not be trimmed.
        }
        finally
        {
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(rowRange);
            Release(row);
            Release(columns);
            Release(rows);
            Release(table);
        }
    }

    internal static bool IsNumberingParagraphAdornment(string? text)
    {
        foreach (var character in text ?? string.Empty)
        {
            if (character is '\t' or '\r' or '\n' or '\v'
                or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                || char.IsWhiteSpace(character))
                continue;
            return false;
        }
        return true;
    }

    internal static void EnsureNumberedOmmlIsDisplay(Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? refreshed = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count == 0) return;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
                math.Type = WdOMathType.wdOMathDisplay;
            try { math.Justification = WdOMathJc.wdOMathJcCenterGroup; } catch { }
            refreshed = math.Range;
            formulaRange.SetRange(refreshed.Start, refreshed.End);
        }
        finally
        {
            Release(refreshed);
            Release(math);
            Release(maths);
        }
    }

    internal static bool ContainsNativeOmml(Range formulaRange)
    {
        OMaths? maths = null;
        try
        {
            maths = formulaRange.OMaths;
            return maths.Count > 0;
        }
        catch { return false; }
        finally { Release(maths); }
    }

    // Word's wdWithInTable answer, without its layout cost for ordinary body
    // ranges. Information() first brings the document layout up to date, which
    // costs hundreds of milliseconds after an edit in a formula-dense document.
    // A range that intersects no table cannot be in one; a range that does
    // keeps Word's exact answer.
    internal static bool IsWithinTableLayoutFree(Range range)
    {
        Tables? tables = null;
        try
        {
            tables = range.Tables;
            if (tables.Count == 0) return false;
        }
        catch (COMException) { }
        finally { Release(tables); }
        return (bool)range.get_Information(WdInformation.wdWithInTable);
    }

    internal static bool RangeIsWhollyWithinTable(Range range)
    {
        Tables? tables = null;
        Table? table = null;
        Range? owner = null;
        try
        {
            tables = range.Tables;
            var start = range.Start;
            var end = range.End;
            var story = range.StoryType;
            for (var index = 1; index <= tables.Count; index++)
            {
                Release(owner); owner = null;
                Release(table); table = tables[index];
                owner = table.Range;
                if (story == owner.StoryType && start >= owner.Start
                    && start < owner.End && end <= owner.End)
                    return true;
            }
            return false;
        }
        finally { Release(owner); Release(table); Release(tables); }
    }

    internal static bool IsManagedNumberedEquationTable(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        if (document is null
            || formulaRange is null
            || string.IsNullOrWhiteSpace(formulaId))
            return false;

        Table? table = null;
        Range? tableRange = null;
        Columns? columns = null;
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null) return false;
            tableRange = table.Range;
            columns = table.Columns;
            if (columns.Count != 3
                || tableRange.StoryType != formulaRange.StoryType
                || formulaRange.Start < tableRange.Start
                || formulaRange.End > tableRange.End
                || !TryGetManagedNumberTableRowIndex(
                    table,
                    formulaRange,
                    expectedColumnIndex: 2,
                    out var formulaRowIndex))
                return false;

            bookmarks = document.Bookmarks;
            var numberName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            return TryGetManagedNumberTableRowIndex(
                       table,
                       numberRange,
                       expectedColumnIndex: 3,
                       out var numberRowIndex)
                && numberRowIndex == formulaRowIndex;
        }
        catch { return false; }
        finally
        {
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(columns);
            Release(tableRange);
            Release(table);
        }
    }

    internal static bool IsNumberedEquationTable(Range formulaRange)
    {
        // Column count is not ownership. In particular, the third cell of an
        // ordinary three-column user table may contain another formula, and must
        // never become our number slot. Require the equation's actual identity
        // and its matching same-row third-column number bookmark.
        Document? document = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Bookmark? bookmark = null;
        try
        {
            if (!RangeIsWhollyWithinTable(formulaRange)) return false;
            document = formulaRange.Document;
            shapes = formulaRange.InlineShapes;
            FormulaMetadata? metadata = null;
            if (shapes.Count == 1)
            {
                shape = shapes[1];
                if (WordFormulaMetadataReader.IsNativeOle(shape))
                    metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape);
            }
            else
            {
                bookmark = WordOmmlFormulaStore.FindAtRange(document, formulaRange);
                if (bookmark is not null)
                    metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
            }
            return metadata?.Numbered == true
                && IsManagedNumberedEquationTable(document, formulaRange, metadata.FormulaId);
        }
        catch { return false; }
        finally { Release(bookmark); Release(shape); Release(shapes); Release(document); }
    }

    internal static int GetManagedNumberTableRowIndex(
        Table table, Range ownedRange, int expectedColumnIndex)
    {
        if (!TryGetManagedNumberTableRowIndex(table, ownedRange, expectedColumnIndex, out var rowIndex))
            throw new InvalidDataException(
                "The formula or number range is not owned by exactly one managed table row.");
        return rowIndex;
    }

    internal static bool TryGetManagedNumberTableRowIndex(
        Table table,
        Range ownedRange,
        int expectedColumnIndex,
        out int rowIndex)
    {
        rowIndex = 0;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? tableRange = null;
        Tables? cellTables = null;
        Table? cellTable = null;
        Range? cellTableRange = null;
        Rows? ownershipRows = null;
        Columns? ownershipColumns = null;
        var ownershipCheck = "initial";
        bool RejectOwnership(string reason)
        {
            TraceNumberingPerformance($"[perf] managed-row-ownership-rejected check={ownershipCheck} reason={reason}");
            return false;
        }
        try
        {
            if (table is null || ownedRange is null)
                return RejectOwnership("missing-table-or-range");
            ownershipColumns = table.Columns;
            ownershipRows = table.Rows;
            var columnCount = ownershipColumns.Count;
            var rowCount = ownershipRows.Count;
            // Cells, column, containment and exact owning table below prove
            // membership without Information(wdWithInTable), a layout query that
            // repaginates long formula documents even on a read-only fast path.
            if (columnCount != 3 || rowCount < 1)
                return RejectOwnership($"columns={columnCount} rows={rowCount} range={ownedRange.StoryType}/{ownedRange.Start}:{ownedRange.End}");

            ownershipCheck = "range-cells";
            cells = ownedRange.Cells;
            if (cells.Count != 1) return RejectOwnership($"cells={cells.Count} range={ownedRange.Start}:{ownedRange.End}");
            cell = cells[1];
            if (cell.ColumnIndex != expectedColumnIndex
                || cell.RowIndex < 1
                || cell.RowIndex > rowCount)
                return RejectOwnership($"cell={cell.RowIndex},{cell.ColumnIndex} expectedColumn={expectedColumnIndex} rows={rowCount}");

            ownershipCheck = "cell-containment";
            cellRange = cell.Range;
            if (ownedRange.StoryType != cellRange.StoryType
                || ownedRange.Start < cellRange.Start
                || ownedRange.End > cellRange.End)
                return RejectOwnership($"range={ownedRange.StoryType}/{ownedRange.Start}:{ownedRange.End} cell={cellRange.StoryType}/{cellRange.Start}:{cellRange.End}");
            ownershipCheck = "cell-table";
            cellTables = cellRange.Tables;
            if (cellTables.Count != 1) return RejectOwnership($"tables={cellTables.Count} cell={cellRange.Start}:{cellRange.End}");
            cellTable = cellTables[1];
            tableRange = table.Range;
            cellTableRange = cellTable.Range;
            if (tableRange.Start != cellTableRange.Start
                || tableRange.End != cellTableRange.End)
                return RejectOwnership($"table={tableRange.Start}:{tableRange.End} cellTable={cellTableRange.Start}:{cellTableRange.End}");

            rowIndex = cell.RowIndex;
            return true;
        }
        catch (Exception error)
        {
            rowIndex = 0;
            return RejectOwnership(error.ToString());
        }
        finally
        {
            Release(cellTableRange);
            Release(cellTable);
            Release(cellTables);
            Release(tableRange);
            Release(cellRange);
            Release(cell);
            Release(cells);
            Release(ownershipColumns);
            Release(ownershipRows);
        }
    }

    internal static Table IsolateManagedNumberTableRow(Table table, Range formulaRange)
    {
        var rowIndex = GetManagedNumberTableRowIndex(table, formulaRange, 2);
        Rows? rows = null;
        Table? trailing = null;
        try
        {
            rows = table.Rows;
            if (rowIndex < rows.Count)
            {
                object firstTrailingRow = rowIndex + 1;
                trailing = table.Split(ref firstTrailingRow);
            }
            if (rowIndex > 1)
            {
                object firstOwnedRow = rowIndex;
                return table.Split(ref firstOwnedRow);
            }
            // Return an independently owned RCW reference, including for row 1.
            Tables? ownedTables = null;
            try
            {
                ownedTables = formulaRange.Tables;
                return ownedTables[1];
            }
            finally { Release(ownedTables); }
        }
        finally
        {
            Release(trailing);
            Release(rows);
        }
    }

    internal static Range? TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph(
        Document document,
        Range formulaRange,
        string formulaId,
        FormulaMetadata? metadata,
        string? preparedStandaloneOmml = null)
    {
        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        OMaths? tableMaths = null;
        InlineShapes? tableShapes = null;
        Cell? leftCell = null;
        Cell? formulaCell = null;
        Cell? numberCell = null;
        Range? leftRange = null;
        Range? formulaCellRange = null;
        Range? numberCellRange = null;
        OMaths? formulaCellMaths = null;
        OMath? formulaCellMath = null;
        InlineShapes? formulaCellShapes = null;
        Fields? formulaCellFields = null;
        OMaths? numberCellMaths = null;
        InlineShapes? numberCellShapes = null;
        Fields? numberCellFields = null;
        Range? visibleNumberRange = null;
        Fields? visibleNumberFields = null;
        Range? equationRange = null;
        Microsoft.Office.Interop.Word.Application? application = null;
        Document? stagingDocument = null;
        Range? stagingInsertion = null;
        OMaths? stagingMaths = null;
        OMath? stagingMath = null;
        Range? stagingMathRange = null;
        Range? convertedRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? editableRange = null;
        Range? placeholderRange = null;
        OMaths? migratedMaths = null;
        OMath? migratedMath = null;
        Range? migratedMathRange = null;
        Bookmark? oldFormulaBookmark = null;
        var migrationWatch = System.Diagnostics.Stopwatch.StartNew();
        long migrationCheckpoint = 0;
        void TraceMigration(string stage)
        {
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
                return;
            var elapsed = migrationWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"      [perf] unnumber-local.{stage}: +{elapsed - migrationCheckpoint}ms ({elapsed}ms)");
            migrationCheckpoint = elapsed;
        }
        Range? Fail(string reason)
        {
            TraceNumberingPerformance(
                $"[migration] legacy-omml-table-rejected formulaId={formulaId} reason={reason}");
            return null;
        }
        try
        {
            if (!IsNumberedEquationTable(formulaRange) || !ContainsNativeOmml(formulaRange))
                return Fail("range-not-numbered-omml-table");
            TraceMigration("initial-host-probe");

            tables = formulaRange.Tables;
            if (tables.Count == 0) return Fail("range-table-collection-empty");
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            // Flatten only the exact VisualTeX-managed legacy host. User-created
            // equation tables, nested tables and malformed structures are never
            // rewritten merely because they happen to contain an OMath. Historical
            // builds may have appended wholly empty rows; those are validated and
            // trimmed below before conversion.
            if (rows.Count < 1 || columns.Count != 3)
                return Fail($"unexpected-dimensions rows={rows.Count} columns={columns.Count}");

            if (rows.Count > 1)
            {
                // Unnumbering changes the selected row into a paragraph. Isolate
                // that proven owner with Word's native table split; other rows,
                // their formulas, fields and reference targets remain live.
                if (!HasReusableNumberedNativeOmmlDirectTableHost(document, formulaRange, formulaId))
                    return Fail("multirow-host-ownership-not-proven");
                var isolated = IsolateManagedNumberTableRow(table, formulaRange);
                Release(columns);
                Release(rows);
                Release(table);
                table = isolated;
                rows = table.Rows;
                columns = table.Columns;
            }

            tableRange = table.Range;
            tableMaths = tableRange.OMaths;
            tableShapes = tableRange.InlineShapes;
            if (tableMaths.Count != 1 || tableShapes.Count != 0)
                return Fail($"table-payload maths={tableMaths.Count} shapes={tableShapes.Count}");

            leftCell = table.Cell(1, 1);
            formulaCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            leftRange = leftCell.Range.Duplicate;
            formulaCellRange = formulaCell.Range.Duplicate;
            numberCellRange = numberCell.Range.Duplicate;
            if (!ContainsOnlyStructuralWordText(leftRange.Text))
                return Fail("left-cell-has-nonstructural-text");
            TraceMigration("table-and-cells");

            formulaCellMaths = formulaCellRange.OMaths;
            formulaCellShapes = formulaCellRange.InlineShapes;
            formulaCellFields = formulaCellRange.Fields;
            if (formulaCellMaths.Count != 1
                || formulaCellShapes.Count != 0
                || formulaCellFields.Count != 0)
                return Fail(
                    $"formula-cell-payload maths={formulaCellMaths.Count} shapes={formulaCellShapes.Count} fields={formulaCellFields.Count}");

            metadata ??= WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null
                || !string.Equals(
                    metadata.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase))
                return Fail("metadata-missing-or-formulaid-mismatch");
            // The caller already holds the just-created live OMath Range. During
            // OLE→OMML replacement Word can temporarily expand VTOMML across the
            // surrounding legacy table, so resolving through that bookmark again
            // may report a non-unique fingerprint even though this exact center cell
            // contains one and only one native equation. Resolve that unique OMath
            // directly from the fully validated cell and use the caller range only
            // as an ownership cross-check.
            formulaCellMath = formulaCellMaths[1];
            equationRange = formulaCellMath.Range.Duplicate;
            if (equationRange.Start < formulaCellRange.Start
                || equationRange.End > formulaCellRange.End)
                return Fail(
                    $"equation-outside-center-cell equation={equationRange.Start}:{equationRange.End} cell={formulaCellRange.Start}:{formulaCellRange.End}");
            var callerOverlapsEquation = formulaRange.Start < equationRange.End
                && formulaRange.End > equationRange.Start;
            var callerMatchesCollapsedBoundary = formulaRange.Start == equationRange.Start
                && formulaRange.End == equationRange.End;
            if (!callerOverlapsEquation && !callerMatchesCollapsedBoundary)
                return Fail(
                    $"caller-range-does-not-own-center-equation caller={formulaRange.Start}:{formulaRange.End} equation={equationRange.Start}:{equationRange.End}");
            TraceMigration("center-equation");
            // The converter-side fingerprint was captured before Word imported and
            // normalized this OMath, so it can legitimately differ at this exact
            // intermediate point. Identity is already proven by the caller-range
            // overlap plus the one-OMath-only managed center cell; the durable live
            // fingerprint is stamped after the completed #(SEQ) replacement.

            // Legacy callers may need to preserve the existing live OMath exactly,
            // so keep the mature hidden-Word FormattedText staging fallback. The
            // interactive ReplaceOmml path already owns the clean target semantic
            // OMML, however; staging that same equation in another document adds a
            // costly Activate/Documents.Add/FormattedText round-trip for no value.
            application = document.Application;
            if (string.IsNullOrWhiteSpace(preparedStandaloneOmml))
            {
                stagingDocument = application.Documents.Add(Visible: false);
                stagingInsertion = stagingDocument.Content;
                stagingInsertion.FormattedText = equationRange.FormattedText;
                stagingMaths = stagingDocument.OMaths;
                if (stagingMaths.Count != 1)
                    throw new InvalidOperationException(
                        "Word could not stage exactly one native OMath before numbering-table migration.");
                stagingMath = stagingMaths[1];
                stagingMathRange = stagingMath.Range.Duplicate;
                try { stagingDocument.Saved = true; } catch { }
                document.Activate();
            }

            numberCellMaths = numberCellRange.OMaths;
            numberCellShapes = numberCellRange.InlineShapes;
            numberCellFields = numberCellRange.Fields;
            if (numberCellMaths.Count != 0 || numberCellShapes.Count != 0)
                return Fail(
                    $"number-cell-object-payload maths={numberCellMaths.Count} shapes={numberCellShapes.Count}");
            visibleNumberRange = FindVisibleEquationNumberRange(document, formulaId);
            if (visibleNumberRange is null)
            {
                if (numberCellFields.Count != 0
                    || !ContainsOnlyStructuralWordText(numberCellRange.Text))
                    return Fail(
                        $"number-cell-without-owned-ref fields={numberCellFields.Count} text='{NormalizeNativeEquationNumberText(numberCellRange.Text)}'");
            }
            else
            {
                if (visibleNumberRange.Start < numberCellRange.Start
                    || visibleNumberRange.End > numberCellRange.End
                    || !ContainsOnlyStructuralWordTextOutsideRange(
                        document,
                        numberCellRange,
                        visibleNumberRange))
                    return Fail(
                        $"owned-ref-outside-number-cell visible={visibleNumberRange.Start}:{visibleNumberRange.End} cell={numberCellRange.Start}:{numberCellRange.End}");
                visibleNumberFields = visibleNumberRange.Fields;
                if (numberCellFields.Count != visibleNumberFields.Count)
                    return Fail(
                        $"number-cell-field-count-mismatch cell={numberCellFields.Count} visible={visibleNumberFields.Count}");
            }
            TraceMigration("number-cell-ownership");

            // Remove only the generated visible REF label before flattening. The
            // hidden VTEqNum_* SEQ target and every external REF remain intact.
            RemoveVisibleEquationNumber(document, formulaId);
            // Table.ConvertToText can expand its returned range across the hidden
            // caption paragraph that immediately follows the table. Reusing that
            // live SEQ/bookmark while replacing the converted range can make the
            // visible REF point to itself. Remove only this formula's hidden caption
            // now; the normal reconciliation path recreates the same VTEqCap/VTEqNum
            // identities before any external reference is updated.
            RemoveNativeCaption(document, formulaId);
            oldFormulaBookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                formulaId);
            try { oldFormulaBookmark?.Delete(); } catch { }
            TraceMigration("remove-number-artifacts");

            object separator = WdTableFieldSeparator.wdSeparateByTabs;
            object nestedTables = false;
            convertedRange = table.ConvertToText(ref separator, ref nestedTables);
            TraceMigration("convert-to-text");
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
            {
                OMaths? convertedMathsProbe = null;
                OMath? convertedMathProbe = null;
                try
                {
                    convertedMathsProbe = convertedRange.OMaths;
                    var convertedType = -1;
                    if (convertedMathsProbe.Count == 1)
                    {
                        convertedMathProbe = convertedMathsProbe[1];
                        convertedType = (int)convertedMathProbe.Type;
                    }
                    Console.WriteLine(
                        $"      [perf] unnumber-convert-to-text maths={convertedMathsProbe.Count} type={convertedType} range={convertedRange.Start}:{convertedRange.End}");
                }
                finally
                {
                    Release(convertedMathProbe);
                    Release(convertedMathsProbe);
                }
            }

            var paragraphStart = convertedRange.Start;
            if (!string.IsNullOrWhiteSpace(preparedStandaloneOmml))
            {
                // Current direct-SEQ 1x3 hosts are already fully owned and were
                // validated above. Word's ConvertToText preserves the center OMath
                // itself and only demotes it to Inline while inserting the two cell
                // separator TABs. Keep that exact live OMath: remove only structural
                // text around it, then restore Display type. Replacing the whole
                // converted range with a placeholder and importing the same OMath a
                // second time cost ~0.5-0.8s per checkbox toggle.
                Release(migratedMaths); migratedMaths = convertedRange.OMaths;
                if (migratedMaths.Count != 1)
                    throw new InvalidOperationException(
                        "Word did not preserve exactly one OMath while removing its numbering table.");
                Release(migratedMath); migratedMath = migratedMaths[1];
                Release(migratedMathRange); migratedMathRange = migratedMath.Range.Duplicate;

                editableRange = document.Range(
                    migratedMathRange.End,
                    Math.Max(migratedMathRange.End, convertedRange.End - 1));
                if (!ContainsOnlyStructuralWordText(editableRange.Text))
                    throw new InvalidOperationException(
                        "The converted number-table suffix contains user text.");
                editableRange.Delete();
                Release(editableRange); editableRange = null;

                placeholderRange = document.Range(
                    convertedRange.Start,
                    migratedMathRange.Start);
                if (!ContainsOnlyStructuralWordText(placeholderRange.Text))
                    throw new InvalidOperationException(
                        "The converted number-table prefix contains user text.");
                placeholderRange.Delete();
                Release(placeholderRange); placeholderRange = null;
                TraceMigration("strip-cell-separators");

                var directProbe = document.Range(
                    paragraphStart,
                    Math.Min(document.Content.End, paragraphStart + 1));
                try
                {
                    Release(paragraphs); paragraphs = directProbe.Paragraphs;
                    if (paragraphs.Count != 1)
                        throw new InvalidOperationException(
                            "Word did not leave one standalone paragraph after removing the numbering table.");
                    Release(paragraph); paragraph = paragraphs[1];
                    Release(paragraphRange); paragraphRange = paragraph.Range;
                    Release(migratedMaths); migratedMaths = paragraphRange.OMaths;
                    if (migratedMaths.Count != 1)
                        throw new InvalidOperationException(
                            "Word lost the preserved OMath after removing table separators.");
                    Release(migratedMath); migratedMath = migratedMaths[1];
                    if (migratedMath.Type != WdOMathType.wdOMathDisplay)
                        migratedMath.Type = WdOMathType.wdOMathDisplay;
                    Release(migratedMathRange); migratedMathRange = migratedMath.Range.Duplicate;
                    Release(equationRange); equationRange = migratedMathRange.Duplicate;
                    TraceMigration("restore-display-range");
                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        ParagraphFormat? probeFormat = null;
                        TabStops? probeTabs = null;
                        try
                        {
                            probeFormat = paragraph.Format;
                            probeTabs = probeFormat.TabStops;
                            TraceNumberingPerformance(
                                $"      [perf] unnumber-direct-paragraph alignment={(int)probeFormat.Alignment} left={probeFormat.LeftIndent:0.###} right={probeFormat.RightIndent:0.###} first={probeFormat.FirstLineIndent:0.###} lineRule={(int)probeFormat.LineSpacingRule} line={probeFormat.LineSpacing:0.###} before={probeFormat.SpaceBefore:0.###} after={probeFormat.SpaceAfter:0.###} tabs={probeTabs.Count} keep={probeFormat.KeepTogether}/{probeFormat.KeepWithNext} pageBreak={probeFormat.PageBreakBefore} widow={probeFormat.WidowControl}");
                        }
                        finally
                        {
                            Release(probeTabs);
                            Release(probeFormat);
                        }
                    }
                    var directResult = equationRange.Duplicate;
                    return directResult;
                }
                finally { Release(directProbe); }
            }

            // Legacy compatibility path: rebuild the complete range returned by
            // Table.ConvertToText from a staged OMath.
            const string placeholderText = "\uE000";
            convertedRange.Text = placeholderText + "\r";
            var paragraphProbe = document.Range(
                paragraphStart,
                Math.Min(document.Content.End, paragraphStart + 1));
            try
            {
                paragraphs = paragraphProbe.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidOperationException(
                        "Word did not create one stable paragraph while rebuilding the numbered OMML host.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }
            finally { Release(paragraphProbe); }

            // Restore the staged professional equation directly into one ordinary
            // table-free paragraph. Do not create the retired TAB + inline OMath
            // intermediate host: numbered OMML must remain genuine Word display math
            // throughout migration, before the outer reconciler adds the dedicated
            // external REF Shape/anchor structure.
            var editableEnd = Math.Max(paragraphStart, paragraphRange.End - 1);
            editableRange = document.Range(paragraphStart, editableEnd);
            editableRange.Text = placeholderText;
            placeholderRange = document.Range(
                paragraphStart,
                paragraphStart + placeholderText.Length);
            if (!string.IsNullOrWhiteSpace(preparedStandaloneOmml))
            {
                Range? preparedRange = null;
                try
                {
                    string mathFontName;
                    try { mathFontName = document.OMathFontName ?? string.Empty; }
                    catch { mathFontName = string.Empty; }
                    preparedRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                        application,
                        document,
                        placeholderRange,
                        preparedStandaloneOmml!,
                        display: true,
                        mathFontName);
                }
                finally { Release(preparedRange); }
            }
            else
            {
                placeholderRange.FormattedText = stagingMathRange.FormattedText;
            }

            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            migratedMaths = paragraphRange.OMaths;
            if (migratedMaths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not restore exactly one native OMath while migrating its numbering table.");
            migratedMath = migratedMaths[1];
            if (migratedMath.Type != WdOMathType.wdOMathDisplay)
                migratedMath.Type = WdOMathType.wdOMathDisplay;
            migratedMathRange = migratedMath.Range;
            equationRange = migratedMathRange.Duplicate;

            // Do not bind metadata here. The outer numbered-OMML reconciler still
            // creates the final anchor/Shape/REF scaffold and owns the one durable
            // bookmark/fingerprint save after that physical host is complete.
            var result = equationRange.Duplicate;
            return result;
        }
        finally
        {
            Release(oldFormulaBookmark);
            Release(migratedMathRange);
            Release(migratedMath);
            Release(migratedMaths);
            Release(placeholderRange);
            Release(editableRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(convertedRange);
            Release(equationRange);
            Release(stagingMathRange);
            Release(stagingMath);
            Release(stagingMaths);
            Release(stagingInsertion);
            if (stagingDocument is not null)
            {
                try
                {
                    stagingDocument.Saved = true;
                    stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(stagingDocument);
            Release(application);
            Release(visibleNumberFields);
            Release(visibleNumberRange);
            Release(numberCellFields);
            Release(numberCellShapes);
            Release(numberCellMaths);
            Release(formulaCellFields);
            Release(formulaCellShapes);
            Release(formulaCellMath);
            Release(formulaCellMaths);
            Release(numberCellRange);
            Release(formulaCellRange);
            Release(leftRange);
            Release(numberCell);
            Release(formulaCell);
            Release(leftCell);
            Release(tableShapes);
            Release(tableMaths);
            Release(tableRange);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
        }
    }

    internal static bool ContainsOnlyStructuralWordText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var character in text!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a'
                or '\u0001' or '\u200B' or '\u200C' or '\u2060' or '\uFEFF')
                continue;
            if (!char.IsWhiteSpace(character)) return false;
        }
        return true;
    }

    private static bool ContainsOnlyStructuralWordTextOutsideRange(
        Document document,
        Range container,
        Range allowed)
    {
        Range? before = null;
        Range? after = null;
        try
        {
            if (allowed.Start < container.Start || allowed.End > container.End)
                return false;
            if (container.Start < allowed.Start)
            {
                before = document.Range(container.Start, allowed.Start);
                if (!ContainsOnlyStructuralWordText(before.Text)) return false;
            }
            if (allowed.End < container.End)
            {
                after = document.Range(allowed.End, container.End);
                if (!ContainsOnlyStructuralWordText(after.Text)) return false;
            }
            return true;
        }
        finally
        {
            Release(after);
            Release(before);
        }
    }

    private static Range? TryConvertStandardNumberedOleTableToTabParagraph(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        InlineShapes? tableShapes = null;
        InlineShape? tableShape = null;
        Range? convertedRange = null;
        InlineShapes? convertedShapes = null;
        InlineShape? convertedShape = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? cleanupRange = null;
        try
        {
            if (!IsNumberedEquationTable(formulaRange) || ContainsNativeOmml(formulaRange))
                return null;

            tables = formulaRange.Tables;
            if (tables.Count == 0) return null;
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            // Historical 2x3 hosts can carry one or more entirely empty trailing
            // rows. Normalize only those rows on this exact table before requiring
            // the final 1x3 migration input. This does not rely on numbering
            // bookmarks, which some oldest OLE documents never stored.
            if (columns.Count == 3 && rows.Count > 1)
            {
                RemoveEmptyTrailingNumberedTableRows(table);
                Release(rows);
                rows = table.Rows;
            }
            // Restrict the compatibility migration to the exact VisualTeX scaffold.
            // Custom or malformed user tables are left untouched rather than flattened.
            if (rows.Count != 1 || columns.Count != 3) return null;

            tableRange = table.Range;
            tableShapes = tableRange.InlineShapes;
            if (tableShapes.Count != 1) return null;
            tableShape = tableShapes[1];
            var metadata = ReadMetadata(tableShape);
            if (metadata is null
                || !string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                return null;

            object separator = WdTableFieldSeparator.wdSeparateByTabs;
            object nestedTables = false;
            convertedRange = table.ConvertToText(ref separator, ref nestedTables);
            convertedShapes = convertedRange.InlineShapes;
            for (var index = 1; index <= convertedShapes.Count; index++)
            {
                Release(convertedShape);
                convertedShape = convertedShapes[index];
                var convertedMetadata = ReadMetadata(convertedShape);
                if (convertedMetadata is not null
                    && string.Equals(
                        convertedMetadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    break;
                Release(convertedShape);
                convertedShape = null;
            }
            if (convertedShape is null)
                throw new InvalidOperationException(
                    "Word converted the numbered VisualTeX table, but the OLE formula could not be resolved afterwards.");

            // Discard the table-cell separators and the old visible REF label, then let
            // the normal paragraph path rebuild exactly one leading and one trailing tab.
            RemoveVisibleEquationNumber(document, formulaId);
            shapeRange = convertedShape.Range;
            paragraphs = shapeRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (shapeRange.End < paragraphRange.End - 1)
            {
                cleanupRange = document.Range(shapeRange.End, paragraphRange.End - 1);
                cleanupRange.Delete();
                Release(cleanupRange);
                cleanupRange = null;
            }

            Release(shapeRange);
            shapeRange = convertedShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start < shapeRange.Start)
            {
                cleanupRange = document.Range(paragraphRange.Start, shapeRange.Start);
                cleanupRange.Delete();
                Release(cleanupRange);
                cleanupRange = null;
            }

            Release(shapeRange);
            shapeRange = convertedShape.Range;
            var result = shapeRange.Duplicate;
            return result;
        }
        finally
        {
            Release(cleanupRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(convertedShape);
            Release(convertedShapes);
            Release(convertedRange);
            Release(tableShape);
            Release(tableShapes);
            Release(tableRange);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
        }
    }

    internal static void ConfigureUnnumberedDisplayFormula(
        Document document,
        Range formulaRange,
        string formulaId,
        FormulaMetadata? knownMetadata = null,
        float? formulaHeightPoints = null,
        string? preparedStandaloneOmml = null,
        bool deferMetadataPersistence = false)
    {
        Range? activeRange = null;
        Bookmark? repairedBookmark = null;
        var unnumberWatch = System.Diagnostics.Stopwatch.StartNew();
        long unnumberCheckpoint = 0;
        void TraceUnnumber(string stage)
        {
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
                return;
            var elapsed = unnumberWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"      [perf] unnumber-post.{stage}: +{elapsed - unnumberCheckpoint}ms ({elapsed}ms)");
            unnumberCheckpoint = elapsed;
        }
        try
        {
            if (ContainsNativeOmml(formulaRange))
            {
                activeRange = ResolveSingleNativeOmmlRange(formulaRange);
                TraceUnnumber("initial-resolve");
                var metadata = knownMetadata
                    ?? WordOmmlFormulaStore.TryRead(document, formulaId);
                if (IsNumberedEquationTable(activeRange))
                {
                    Range? standaloneRange = null;
                    try
                    {
                        standaloneRange = TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph(
                            document,
                            activeRange,
                            formulaId,
                            metadata,
                            preparedStandaloneOmml);
                        if (standaloneRange is null)
                            throw new InvalidOperationException(
                                "VisualTeX could not safely remove the managed 1x3 OMML numbering table.");
                        Release(activeRange);
                        // TryConvertStandardNumberedOmmlTableToStandaloneDisplayParagraph
                        // returns the exact preserved live Display OMath Range after
                        // ConvertToText and separator cleanup. Re-resolving the same
                        // OMath through Range.OMaths costs ~100ms in a 100-OMML
                        // document and adds no identity information here.
                        activeRange = standaloneRange;
                        standaloneRange = null;
                        TraceUnnumber("convert-live-range");
                    }
                    finally { Release(standaloneRange); }
                    if (!string.IsNullOrWhiteSpace(preparedStandaloneOmml))
                    {
                        FinalizeDirectUnnumberedOmmlParagraph(activeRange);
                        TraceUnnumber("direct-paragraph-finalize");
                    }
                    else
                    {
                        NormalizeCompactStandaloneNativeOmmlParagraph(activeRange);
                        TraceUnnumber("compact-paragraph");
                        ConfigureEquationParagraph(activeRange, numbered: false);
                        TraceUnnumber("configure-paragraph");
                        EnsureNumberedOmmlIsDisplay(activeRange);
                        TraceUnnumber("ensure-display");
                    }
                    if (metadata is not null && !deferMetadataPersistence)
                    {
                        repairedBookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            activeRange,
                            metadata,
                            replaceExisting: true);
                        TraceUnnumber("wrap-anchor");
                        WordOmmlNativeSource.StampFingerprintFromResolvedRange(metadata, activeRange);
                        WordOmmlFormulaStore.Save(document, metadata);
                    }
                    formulaRange.SetRange(activeRange.Start, activeRange.End);
                    // Table dismantling stages the professional OMath through a
                    // temporary hidden Word document. Re-activate the real target
                    // after that staging document is closed so the next immediate
                    // VisualTeX edit does not observe a transient/no ActiveDocument.
                    try { document.Activate(); } catch { }
                    return;
                }
                if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                        WordOmmlConverter.ReadLocalEquationXml(activeRange))
                    || HasManagedNativeOmmlHashSequenceHost(
                        document,
                        formulaId))
                {
                    RemoveVisibleEquationNumber(document, formulaId);
                    var semanticOmml = WordOmmlConverter
                        .StripManagedVisualTeXNativeEquationNumber(
                            WordOmmlConverter.ReadLocalEquationXml(activeRange));
                    Microsoft.Office.Interop.Word.Application? application = null;
                    Range? replacementRange = null;
                    try
                    {
                        application = document.Application;
                        string mathFontName;
                        try { mathFontName = document.OMathFontName ?? string.Empty; }
                        catch { mathFontName = string.Empty; }
                        replacementRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                            application,
                            document,
                            activeRange,
                            semanticOmml,
                            display: true,
                            mathFontName);
                        Release(activeRange);
                        activeRange = replacementRange;
                        replacementRange = null;
                    }
                    finally
                    {
                        Release(replacementRange);
                        Release(application);
                    }
                }
                else
                {
                    RemoveVisibleEquationNumber(document, formulaId);
                }

                RemoveNativeCaption(document, formulaId);
                RemoveLeadingEquationTab(document, activeRange);
                RemoveOrdinaryTabsAfterNativeOmml(document, activeRange);
                ConfigureEquationParagraph(activeRange, numbered: false);
                EnsureNumberedOmmlIsDisplay(activeRange);
                if (metadata is not null)
                {
                    repairedBookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        activeRange,
                        metadata,
                        replaceExisting: true);
                    if (!deferMetadataPersistence)
                    {
                        WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                            metadata,
                            activeRange);
                        WordOmmlFormulaStore.Save(document, metadata);
                    }
                }
                formulaRange.SetRange(activeRange.Start, activeRange.End);
                return;
            }

            RemoveVisibleEquationNumber(document, formulaId);
            RemoveNativeCaption(document, formulaId);
            ConfigureEquationParagraph(formulaRange, numbered: false);
            EnsureLeadingEquationTab(document, formulaRange);
            if (knownMetadata is not null && formulaHeightPoints is > 0f)
            {
                AlignDisplayOleObjectToParagraphBaseline(
                    document,
                    formulaRange,
                    formulaHeightPoints.Value,
                    knownMetadata);
            }
            else
            {
                // Metadata-free legacy cleanup cannot infer the formula's exported
                // baseline safely. Retain the old neutral fallback for that narrow
                // repair path; normal create/edit/reconcile always supplies both.
                ResetOleObjectBaselinePosition(document, formulaRange);
            }
        }
        finally
        {
            Release(repairedBookmark);
            Release(activeRange);
        }
    }

    internal static void AlignDisplayOleObjectToParagraphBaseline(
        Document document,
        Range formulaRange,
        float actualHeightPoints,
        FormulaMetadata metadata)
    {
        Microsoft.Office.Interop.Word.Font? rangeFont = null;
        Range? character = null;
        Microsoft.Office.Interop.Word.Font? characterFont = null;
        try
        {
            // VisualTeX OLE and MathType are different objects, but Word lays both
            // out as an EMBED field whose U+0001 result character owns the visible
            // object box. Keep the field instruction at the ordinary paragraph
            // baseline and move only that object-result character. The offset is
            // calculated from VisualTeX's own exported MathJax baseline, not from
            // MathType's storage or an arbitrary half-height approximation.
            rangeFont = formulaRange.Font;
            rangeFont.Position = 0;

            var semanticFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
            var exportedHeight = metadata.RenderHeightPx.HasValue
                ? (float)metadata.RenderHeightPx.Value
                : 0f;
            float? exportedBaseline = metadata.Baseline.HasValue
                ? (float)metadata.Baseline.Value
                : null;
            // Preview-ink readback used to run here for every unnumbered display
            // OLE even though the resulting ratios are not used by the baseline
            // calculation below. That diagnostic cache forces Word to materialize
            // EnhMetaFileBits and parse the preview during every batch insertion.
            // Keep insertion dependent only on the already-exported geometry.

            var position = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                actualHeightPoints,
                exportedHeight,
                exportedBaseline,
                existingFontPosition: null,
                sourceSemanticFontSizePoints: semanticFontSize,
                targetSemanticFontSizePoints: semanticFontSize);
            var clamped = Math.Max(-256, Math.Min(256, position));

            // A Word inline OLE is a wdFieldEmbed whose Result owns the painted
            // U+0001 object character. Resolve that one local field directly on
            // the normal path instead of creating a COM Range for every character
            // in the complete EMBED field span. The legacy character walk remains
            // the compatibility fallback when Word exposes an unusual field tree.
            if (TrySetEmbeddedOleResultPosition(formulaRange, clamped))
                return;

            for (var index = formulaRange.Start; index < formulaRange.End; index++)
            {
                Release(characterFont);
                characterFont = null;
                Release(character);
                character = document.Range(index, index + 1);
                if (!string.Equals(character.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                characterFont = character.Font;
                characterFont.Position = clamped;
                return;
            }
        }
        finally
        {
            Release(characterFont);
            Release(character);
            Release(rangeFont);
        }
    }

    private static bool TrySetEmbeddedOleResultPosition(
        Range formulaRange,
        int position)
    {
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            // formulaRange is the exact InlineShape.Range supplied by the host
            // writer. Direct Fields access is bounded to this one OLE and avoids
            // the arbitrary-range compatibility work used elsewhere.
            fields = formulaRange.Fields;
            if (fields.Count != 1)
                return false;
            field = fields[1];
            if (field.Type != WdFieldType.wdFieldEmbed)
                return false;

            result = field.Result.Duplicate;
            var text = result.Text ?? string.Empty;
            if (!string.Equals(text, "\u0001", StringComparison.Ordinal)
                && !string.Equals(text, "\uFFFC", StringComparison.Ordinal))
                return false;
            if (result.Start < formulaRange.Start
                || result.End > formulaRange.End)
                return false;

            font = result.Font;
            font.Position = position;
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(font);
            Release(result);
            Release(field);
            Release(fields);
        }
    }

    private static void ResetOleObjectBaselinePosition(Document document, Range formulaRange)
    {
        Microsoft.Office.Interop.Word.Font? rangeFont = null;
        Range? character = null;
        Microsoft.Office.Interop.Word.Font? characterFont = null;
        try
        {
            rangeFont = formulaRange.Font;
            rangeFont.Position = 0;
            for (var index = formulaRange.Start; index < formulaRange.End; index++)
            {
                Release(characterFont);
                characterFont = null;
                Release(character);
                character = document.Range(index, index + 1);
                if (!string.Equals(character.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                characterFont = character.Font;
                characterFont.Position = 0;
                return;
            }
        }
        finally
        {
            Release(characterFont);
            Release(character);
            Release(rangeFont);
        }
    }

    private static void FinalizeDirectUnnumberedOmmlParagraph(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The direct-SEQ table did not flatten to one formula paragraph.");
            paragraph = paragraphs[1];
            format = paragraph.Format;

            // The managed 1x3 host already normalized indents, spacing, list state
            // and pagination flags when it was created. ConvertToText preserves
            // those values; the only structural formatting it carries into the
            // standalone paragraph is the table's tab-stop set and left alignment.
            // Clear exactly those two artifacts. Retain the old compact-line guard
            // for historical tables whose mandatory separator paragraph used an
            // Exactly-1pt line box.
            tabStops = format.TabStops;
            if (tabStops.Count > 0) tabStops.ClearAll();
            format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && format.LineSpacing <= 2.01f)
            {
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                try { format.DisableLineHeightGrid = -1; } catch { }
            }
        }
        finally
        {
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool NormalizeCompactStandaloneNativeOmmlParagraph(
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        try
        {
            if ((bool)formulaRange.get_Information(WdInformation.wdWithInTable))
                return false;
            OMaths? maths = null;
            try
            {
                maths = formulaRange.OMaths;
                if (maths.Count != 1) return false;
            }
            finally { Release(maths); }
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            // A direct-SEQ 1x3 is separated from adjacent Word tables by a mandatory
            // VisualTeX structural paragraph using Exactly 1pt line spacing. When
            // Word dismantles the table it can reuse that paragraph's pPr for the
            // newly restored standalone Display OMath. Preserve ordinary/custom
            // paragraph formatting, but never let this unmistakable compact-table
            // sentinel become the visible formula's line box.
            if (format.LineSpacingRule != WdLineSpacing.wdLineSpaceExactly
                || format.LineSpacing > 2.01f)
                return false;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            try { format.DisableLineHeightGrid = -1; } catch { }
            return true;
        }
        finally
        {
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }
    }
}

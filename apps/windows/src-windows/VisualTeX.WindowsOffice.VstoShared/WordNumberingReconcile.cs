using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationHeadings;
using static VisualTeX.WordVsto.WordEquationNumberFormat;
using static VisualTeX.WordVsto.WordEquationNumberReferences;
using static VisualTeX.WordVsto.WordEquationParagraphLayout;
using static VisualTeX.WordVsto.WordEquationSequenceFields;
using static VisualTeX.WordVsto.WordNativeCaptions;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberedTypingParagraphs;
using static VisualTeX.WordVsto.WordNumberingCleanup;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingState;
using static VisualTeX.WordVsto.WordVisibleEquationNumbers;

namespace VisualTeX.WordVsto;

/// <summary>
/// Reconciling equation numbers with the document: the whole-document pass, a formula's number, scaffolds for
/// conversion, new-number batches, finding the visible number and the numbered table.
/// </summary>
internal static class WordNumberingReconcile
{
    private static bool IsCustomUndoRecordActive(Document document)
    {
        UndoRecord? record = null;
        try
        {
            record = document.Application.UndoRecord;
            return record.IsRecordingCustomRecord || record.CustomRecordLevel > 0;
        }
        catch
        {
            return false;
        }
        finally { Release(record); }
    }

    public static void TryReconcile(Document document)
    {
        try { Reconcile(document); }
        catch
        {
            // Formula insertion/update is already durable. The user can retry
            // only the numbering command without duplicating or losing it.
        }
    }

    internal static void BuildFormulaNumberingScaffoldForConversion(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata,
        Table? knownNumberedTable = null,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool deferFieldUpdate = false,
        bool knownFreshConversionTarget = false)
    {
        if (!metadata.Numbered
            || !string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            return;
        EnsureDocumentEquationNumberFormatPreference(document);
        var formulaFontSizePoints = (float)FormulaFontSize.ResolveSemanticFontSize(metadata);
        if (ContainsNativeOmml(formulaRange))
        {
            ConfigureNumberedNativeOmmlDisplay(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId,
                reuseExistingScaffold: false,
                metadata,
                _ => { },
                plannedOrdinal,
                plannedPrefix,
                deferFieldUpdate: true);
            return;
        }

        // A conversion may still hand us the RCW for a legacy 1x3/2x3 owner,
        // but it is migration input only. No conversion path is allowed to create,
        // preserve or populate a numbered table; ConfigureNumberedDisplayFormula
        // converts a proven legacy OLE table to the final center/right-tab paragraph
        // before constructing its caption and external REF.

        ConfigureNumberedDisplayFormula(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            metadata.FormulaId,
            reuseExistingScaffold: false,
            knownNumberedTable: null,
            useConversionSafeVisibleNumber: true,
            metadata: metadata,
            plannedOrdinal: plannedOrdinal,
            plannedPrefix: plannedPrefix,
            deferFieldUpdate: deferFieldUpdate,
            knownFreshConversionTarget: knownFreshConversionTarget);
    }

    internal static IReadOnlyDictionary<string, (int Ordinal, string Prefix)> PlanNewEquationNumberBatch(
        Document document,
        IReadOnlyCollection<(string FormulaId, int Position)> newFormulaPositions,
        out bool hasExistingManagedCaptions)
    {
        var distinctNew = newFormulaPositions
            .Where(item => !string.IsNullOrWhiteSpace(item.FormulaId))
            .GroupBy(item => item.FormulaId, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.OrderBy(item => item.Position).First())
            .OrderBy(item => item.Position)
            .ToArray();
        var result = new Dictionary<string, (int Ordinal, string Prefix)>(
            StringComparer.OrdinalIgnoreCase);
        if (distinctNew.Length == 0)
        {
            hasExistingManagedCaptions = false;
            return result;
        }

        var newIds = distinctNew
            .Select(item => item.FormulaId)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var existingCaptions = GetNativeEquationCaptionEntries(
                document,
                GetNativeEquationSequenceName(document))
            .Where(item => !newIds.Contains(item.FormulaId))
            .ToArray();
        hasExistingManagedCaptions = existingCaptions.Length > 0;

        var syntheticCaptions = existingCaptions
            .Concat(distinctNew.Select(item => new NativeEquationCaptionEntry(
                item.FormulaId,
                item.Position,
                string.Empty)))
            .OrderBy(item => item.Position)
            .ToArray();
        var format = ReadEquationNumberFormat(document);
        var headingAnchors = format.UsesHeading
            ? GetHeadingNumberAnchorsForFormatBatch(
                document,
                format.HeadingLevel,
                syntheticCaptions)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();
        foreach (var item in syntheticCaptions)
        {
            var scope = ResolveEquationNumberScope(
                item.Position,
                format,
                headingAnchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var ordinal);
            ordinalByScope[scope.ScopePosition] = ++ordinal;
            if (newIds.Contains(item.FormulaId))
                result[item.FormulaId] = (ordinal, scope.Prefix);
        }
        return result;
    }

    internal static bool VerifyFreshPlannedEquationNumberBatch(
        Document document,
        IReadOnlyDictionary<string, (int Ordinal, string Prefix)> plannedNumbers)
    {
        if (plannedNumbers.Count == 0) return true;
        if (!TryReadHealthyEquationNumberArtifactsFromOpenXml(
                document,
                out var captions,
                out var referenceCounts))
            return false;
        if (captions.Count != plannedNumbers.Count
            || referenceCounts.Count != plannedNumbers.Count)
            return false;
        var byId = captions.ToDictionary(
            item => item.FormulaId,
            item => item.NumberText,
            StringComparer.OrdinalIgnoreCase);
        foreach (var item in plannedNumbers)
        {
            var expected = item.Value.Prefix
                + item.Value.Ordinal.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            if (!byId.TryGetValue(item.Key, out var actual)
                || !string.Equals(actual, expected, StringComparison.Ordinal))
                return false;
            if (!referenceCounts.TryGetValue(item.Key, out var count)
                || count != 1)
                return false;
        }
        WriteNativeEquationTailFormulaId(
            document,
            captions.OrderBy(item => item.Position).LastOrDefault()?.FormulaId);
        return true;
    }

    internal static int BuildNewOmmlNumberingBatch(
        Document document,
        IReadOnlyCollection<string> formulaIds,
        bool deferFinalIdentityPersistence = false,
        bool cleanupCommittedSourceSpacing = false,
        IReadOnlyDictionary<string, FormulaMetadata>? verifiedFreshMetadata = null,
        WordOmmlNativeSource.LiveInsertionOwners? verifiedFreshInsertions = null,
        bool verifiedFreshBareOmml = false,
        IReadOnlyDictionary<string, string>? verifiedFreshSourceOmml = null,
        bool allowExternalCleanTableTemplate = true)
    {
        var built = 0;
        var entries = new List<(string FormulaId, FormulaMetadata Metadata, int Position, int SourceParagraphStart)>();
        var committedTableStarts = cleanupCommittedSourceSpacing
            ? new List<(string FormulaId, int TableStart)>()
            : null;
        Document? cleanTableTemplateDocument = null;
        Range? cleanTableTemplateRange = null;
        var allowCleanTableTemplate = false;
        Sections? batchSections = null;
        try
        {
            if (cleanupCommittedSourceSpacing && allowExternalCleanTableTemplate)
            {
                batchSections = document.Sections;
                allowCleanTableTemplate = batchSections.Count == 1;
            }
        if (verifiedFreshBareOmml
            && (verifiedFreshSourceOmml is null
                || verifiedFreshSourceOmml.Count != formulaIds
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Count()
                || formulaIds.Any(formulaId =>
                    !verifiedFreshSourceOmml.ContainsKey(formulaId))))
            throw new InvalidDataException(
                "Verified fresh OMML numbering requires the exact prepared source inventory.");
        }
        finally { Release(batchSections); }
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var watch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        long checkpoint = 0;
        void TraceStage(string stage)
        {
            if (watch is null) return;
            var elapsed = watch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] ConvertedOmmlBatch.{stage}: +{elapsed - checkpoint}ms ({elapsed}ms)");
            checkpoint = elapsed;
        }
        Range ResolveBatchSourceRange(FormulaMetadata metadata, out bool usedVerifiedOwner)
        {
            usedVerifiedOwner = false;
            Range? candidate;
            if (verifiedFreshInsertions is not null)
            {
                candidate = verifiedFreshInsertions.TryReadVerifiedOwner(metadata.FormulaId)
                    ?? throw new InvalidDataException(
                        $"Verified fresh OMML owner {metadata.FormulaId} disappeared before numbering.");
                usedVerifiedOwner = true;
            }
            else
            {
                candidate = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    metadata.FormulaId,
                    metadata);
            }

            if (!verifiedFreshBareOmml)
                return candidate;

            try
            {
                // Fresh OMML may legitimately live inside an ordinary user table.
                // The old assertion rejected every table-owned formula before the
                // numbering code could choose its cell-safe native #(SEQ) host.
                // Reject only a pre-existing VisualTeX numbering table: a fresh bare
                // source must not already own our generated numbering scaffold.
                if (IsManagedNumberedEquationTable(
                        document,
                        candidate,
                        metadata.FormulaId))
                    throw new InvalidDataException(
                        $"Verified fresh OMML {metadata.FormulaId} unexpectedly entered a managed numbering table before numbering.");
                return candidate;
            }
            catch
            {
                Release(candidate);
                throw;
            }
        }
        try
        {
            var verifiedInventoryOwners = 0;
            foreach (var formulaId in formulaIds.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                // Format conversion may provide the exact metadata objects that the
                // preceding document-wide fingerprint pass just validated and saved.
                // Other callers keep the established durable-store path.
                FormulaMetadata metadata;
                if (verifiedFreshMetadata is not null
                    && verifiedFreshMetadata.TryGetValue(formulaId, out var verifiedMetadata))
                    metadata = verifiedMetadata;
                else
                    metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                        ?? throw new InvalidDataException($"The new OMML identity {formulaId} is missing.");
                if (!metadata.Numbered || metadata.DisplayMode != "block")
                    throw new InvalidDataException($"The new OMML identity {formulaId} is not a numbered display formula.");
                Range? range = null;
                Paragraphs? sourceParagraphs = null;
                Paragraph? sourceParagraph = null;
                Range? sourceParagraphRange = null;
                try
                {
                    range = ResolveBatchSourceRange(metadata, out var usedVerifiedOwner);
                    if (usedVerifiedOwner) verifiedInventoryOwners++;
                    sourceParagraphs = range.Paragraphs;
                    if (sourceParagraphs.Count != 1)
                        throw new InvalidOperationException(
                            $"Converted OMML formula {metadata.FormulaId} does not occupy one source paragraph before numbering migration.");
                    sourceParagraph = sourceParagraphs[1];
                    sourceParagraphRange = sourceParagraph.Range;
                    entries.Add((
                        metadata.FormulaId,
                        metadata,
                        range.Start,
                        sourceParagraphRange.Start));
                }
                finally
                {
                    Release(sourceParagraphRange);
                    Release(sourceParagraph);
                    Release(sourceParagraphs);
                    Release(range);
                }
            }

            TraceStage(verifiedInventoryOwners == entries.Count && entries.Count > 0
                ? "inventory-verified-owner"
                : "inventory");

            // Plan the final number for every converted formula before any tab,
            // caption or REF structure is inserted. This is the same fast-path
            // principle used for large numbered documents: discover heading scope
            // once, then carry known local ordinals through the mutation phase.
            var orderedEntries = entries.OrderBy(item => item.Position).ToArray();
            var plannedNumbers = new Dictionary<string, (int Ordinal, string Prefix)>(
                StringComparer.OrdinalIgnoreCase);
            if (orderedEntries.Length == 1)
            {
                Range? source = null;
                try
                {
                    var entry = orderedEntries[0];
                    source = document.Range(entry.Position, entry.Position);
                    // Use the same positional planner as ordinary insertion.
                    // A single new row needs no whole-document caption XML map.
                    plannedNumbers[entry.FormulaId] = ResolveDirectTableNumberPlan(
                        document, source, entry.FormulaId, null, null);
                }
                finally { Release(source); }
            }
            else
            {
                var format = ReadEquationNumberFormat(document);
                var newIds = new HashSet<string>(orderedEntries.Select(item => item.FormulaId), StringComparer.OrdinalIgnoreCase);
                var syntheticCaptions = GetNativeEquationCaptionEntries(
                        document, GetNativeEquationSequenceName(document))
                    .Where(item => !newIds.Contains(item.FormulaId))
                    .Concat(orderedEntries.Select(item => new NativeEquationCaptionEntry(
                        item.FormulaId, item.Position, string.Empty)))
                    .OrderBy(item => item.Position).ToArray();
                var headingAnchors = format.UsesHeading
                    ? GetHeadingNumberAnchorsForFormatBatch(document, format.HeadingLevel, syntheticCaptions)
                    : Array.Empty<HeadingNumberAnchor>();
                var ordinalByScope = new Dictionary<int, int>();
                foreach (var item in syntheticCaptions)
                {
                    var scope = ResolveEquationNumberScope(item.Position, format, headingAnchors);
                    ordinalByScope.TryGetValue(scope.ScopePosition, out var ordinal);
                    ordinalByScope[scope.ScopePosition] = ++ordinal;
                    plannedNumbers[item.FormulaId] = (ordinal, scope.Prefix);
                }
            }
            TraceStage("number-plan");

            // Use the same managed row as normal insertion: pure mathematical
            // content in the center cell, ordinary Word numbering in the last
            // cell. Resolve each identity again after every structural change.
            // Every fresh formula is re-resolved through its durable FormulaId before
            // mutation, so the captured numeric position is needed only for ordering
            // and the already-computed number plan. Build rows from start to end: a
            // later table insertion then leaves all previously materialized 1x3 rows
            // before the mutation point. The old end-to-start order repeatedly shifted
            // and repaginated an ever-growing suffix of complex tables, making a 100
            // row batch effectively O(N²) in Word's layout engine.
            foreach (var entry in entries.OrderBy(item => item.Position))
            {
                Range? range = null;
                Range? refreshedRange = null;
                Bookmark? repairedBookmark = null;
                try
                {
                    var metadata = entry.Metadata;
                    range = ResolveBatchSourceRange(metadata, out var usedVerifiedOwner);
                    TraceStage(usedVerifiedOwner
                        ? "locate-target-verified-owner"
                        : "locate-target");

                    var plannedNumber = plannedNumbers[entry.FormulaId];
                    var replaceSourceParagraphFromTemplate =
                        allowCleanTableTemplate
                        && cleanTableTemplateRange is not null
                        && CanReplaceNativeOmmlSourceParagraphFromTemplate(document, range);
                    ConfigureNumberedNativeOmmlDisplay(
                        document,
                        range,
                        WordOmmlFormulaStore.EstimateHeightPoints(range),
                        (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                        entry.FormulaId,
                        reuseExistingScaffold: false,
                        metadata,
                        stage => TraceStage("true-display-" + stage),
                        plannedNumber.Ordinal,
                        plannedNumber.Prefix,
                        deferFieldUpdate: true,
                        deferMetadataPersistence: true,
                        cleanTableTemplateRange: cleanTableTemplateRange,
                        replaceSourceParagraphFromTemplate:
                            replaceSourceParagraphFromTemplate,
                        verifiedFreshBareOmml:
                            verifiedFreshBareOmml,
                        preparedFreshBareOmml:
                            verifiedFreshBareOmml
                                && verifiedFreshSourceOmml!.TryGetValue(
                                    entry.FormulaId,
                                    out var preparedFreshBareOmml)
                                    ? preparedFreshBareOmml
                                    : null);
                    TraceStage("true-display-host");
                    if (usedVerifiedOwner)
                        verifiedFreshInsertions!.RebindAfterVerifiedNumbering(entry.FormulaId, range);

                    if (committedTableStarts is not null)
                    {
                        // A numbered OMML formula inside a user table intentionally
                        // remains in that cell and uses the native #(SEQ) wrapper.
                        // Do not mistake the outer user table for the 1x3 table that
                        // this cleanup list is designed to compact/delete around.
                        if (IsHealthyNumberedNativeOmmlHashSequenceHost(
                                document,
                                range,
                                entry.FormulaId,
                                requireCurrentFieldResult: false,
                                skipWordOpenXmlSequenceProof: verifiedFreshBareOmml))
                        {
                            TraceStage("capture-table-cell-hash-host");
                        }
                        else
                        {
                            Tables? committedTables = null;
                            Table? committedTable = null;
                            Range? committedTableRange = null;
                            try
                            {
                                committedTables = range.Tables;
                                if (committedTables.Count != 1)
                                    throw new InvalidDataException(
                                        $"The newly numbered OMML {entry.FormulaId} does not own exactly one committed table.");
                                committedTable = committedTables[1];
                                committedTableRange = committedTable.Range;
                                if (!IsManagedNumberedEquationTable(
                                        document,
                                        range,
                                        entry.FormulaId))
                                    throw new InvalidDataException(
                                        $"The newly numbered OMML {entry.FormulaId} resolved to an unrelated Word table.");
                                if (!replaceSourceParagraphFromTemplate)
                                    committedTableStarts.Add((entry.FormulaId, committedTableRange.Start));
                                else
                                    TraceStage("atomic-source-paragraph-replaced");
                                if (allowCleanTableTemplate && cleanTableTemplateRange is null)
                                {
                                    try
                                    {
                                        cleanTableTemplateRange = CreateCleanNativeOmmlNumberTableTemplate(
                                            document,
                                            committedTable,
                                            out cleanTableTemplateDocument);
                                        TraceStage("clean-table-template-created");
                                    }
                                    catch (Exception templateError)
                                    {
                                        allowCleanTableTemplate = false;
                                        TraceNumberingPerformance(
                                            $"[perf] ConvertedOmmlBatch.clean-table-template-fallback error={templateError.GetType().Name}:{templateError.Message}");
                                    }
                                }
                            }
                            finally
                            {
                                Release(committedTableRange);
                                Release(committedTable);
                                Release(committedTables);
                            }
                            TraceStage("capture-committed-table");
                        }
                    }

                    if (!deferFinalIdentityPersistence)
                    {
                        refreshedRange = ResolveSingleNativeOmmlRange(range);
                        repairedBookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            refreshedRange,
                            metadata,
                            replaceExisting: true);
                        WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                            metadata,
                            refreshedRange);
                        WordOmmlFormulaStore.Save(document, metadata);
                        TraceStage("repair-identity");
                    }
                    else
                    {
                        // ConfigureNumberedNativeOmmlDisplay already rebound the
                        // FormulaId to the center cell. Defer the expensive per-row
                        // WordOpenXML fingerprint and CustomXML write until every
                        // new table exists; the caller verifies the complete batch
                        // against one document XML snapshot before commit.
                        TraceStage("repair-identity-deferred");
                    }
                    built++;
                }
                finally
                {
                    Release(repairedBookmark);
                    Release(refreshedRange);
                    Release(range);
                }
            }

            if (committedTableStarts is not null)
            {
                foreach (var committed in committedTableStarts
                             .OrderByDescending(item => item.TableStart))
                {
                    Range? probe = null;
                    Tables? tables = null;
                    Table? table = null;
                    Range? tableRange = null;
                    try
                    {
                        var contentEnd = document.Content.End;
                        if (committed.TableStart < document.Content.Start
                            || committed.TableStart >= contentEnd)
                            throw new InvalidDataException(
                                $"The committed numbered OMML table {committed.FormulaId} moved outside the document.");
                        probe = document.Range(
                            committed.TableStart,
                            Math.Min(contentEnd, committed.TableStart + 1));
                        tables = probe.Tables;
                        if (tables.Count != 1)
                            throw new InvalidDataException(
                                $"The committed numbered OMML table {committed.FormulaId} cannot be re-resolved by its stable start.");
                        table = tables[1];
                        tableRange = table.Range;
                        if (tableRange.Start != committed.TableStart
                            || !IsManagedNativeOmmlDirectTable(table))
                            throw new InvalidDataException(
                                $"The committed numbered OMML table {committed.FormulaId} changed before batch cleanup.");
                        CleanupNumberedDisplayInsertionSpacing(
                            document,
                            committed.FormulaId,
                            knownDirectTable: table,
                            cleanupFreshOmmlSource: true);
                    }
                    finally
                    {
                        Release(tableRange);
                        Release(table);
                        Release(tables);
                        Release(probe);
                    }
                }
                TraceStage("cleanup-committed-source-spacing");
            }
            return built;
        }
        catch (Exception error)
        {
            TraceNumberingPerformance(
                $"[perf] converted-omml-numbering-batch.failed built={built}/{entries.Count} error={error}");
            throw new InvalidDataException(
                $"Converted OMML numbering failed after {built}/{entries.Count} rows: {error.Message}", error);
        }
        finally
        {
            Release(cleanTableTemplateRange);
            if (cleanTableTemplateDocument is not null)
            {
                try { cleanTableTemplateDocument.Close(WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(cleanTableTemplateDocument);
            try { document.Activate(); } catch { }
        }
    }

    internal static void ReconcileFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null,
        bool deferNativeOmmlShapeFinalization = false,
        bool deferNativeOmmlShapeCreation = false,
        bool deferNativeOmmlMetadataPersistence = false,
        string? preparedUnnumberedOmml = null,
        bool verifiedFreshBareOmml = false,
        string? preparedFreshBareOmml = null,
        int? plannedNumberOrdinal = null,
        string? plannedNumberPrefix = null)
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
            Console.WriteLine(
                $"      [perf] ReconcileFormula.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }

        var hadNumberingArtifacts = HasFormulaNumberingArtifacts(
            document,
            metadata.FormulaId);
        TraceStage("artifact-probe");
        if (metadata.DisplayMode != "block")
        {
            if (!hadNumberingArtifacts) return;
            RemoveFormulaNumberingArtifacts(document, metadata.FormulaId);
            UpdateMainStoryFields(document);
            UpdateNativeCrossReferences(document);
            return;
        }

        var formulaFontSizePoints = (float)FormulaFontSize.ResolveSemanticFontSize(metadata);
        if (metadata.Numbered)
        {
            var hadCompleteOwnedArtifacts =
                HasCompleteFormulaNumberingArtifacts(document, metadata.FormulaId)
                && FormulaRangeOwnsNumberingArtifacts(
                    document,
                    formulaRange,
                    metadata.FormulaId);
            TraceStage("owned-artifacts");

            EnsureDocumentEquationNumberFormatPreference(document);
            TraceStage("format-preference");
            var rebuiltStableNumberingSlot =
                !hadCompleteOwnedArtifacts
                && knownNumberedTable is not null
                && reuseExistingNumberedTableFormatting
                && !numberingOrderMayHaveChanged;
            var rebuiltWithPreservedNumberPlan =
                !hadCompleteOwnedArtifacts
                && !numberingOrderMayHaveChanged
                && plannedNumberOrdinal.HasValue
                && plannedNumberPrefix is not null;
            var deferLegacyNativeOmmlShapeCreation =
                deferNativeOmmlShapeCreation;
            if (ContainsNativeOmml(formulaRange))
            {
                try
                {
                    deferLegacyNativeOmmlShapeCreation =
                        deferLegacyNativeOmmlShapeCreation
                        || IsManagedNumberedEquationTable(
                            document,
                            formulaRange,
                            metadata.FormulaId)
                        || (!verifiedFreshBareOmml
                            && WordOmmlEquationNumbers.HasVisualTeXNativeEquationNumber(
                                WordOmmlConverter.ReadLocalEquationXml(formulaRange)));
                }
                catch
                {
                    // Any ambiguous legacy native host is handled conservatively by
                    // the normal structural path; fresh table-free insertions keep
                    // their immediate Shape creation behavior.
                }
            }
            var visibleNumberCreated = ConfigureNumberedDisplayFormula(
                document,
                formulaRange,
                formulaHeightPoints,
                formulaFontSizePoints,
                metadata.FormulaId,
                reuseExistingScaffold:
                    knownNumberedTable is not null
                    || (hadCompleteOwnedArtifacts && !ContainsNativeOmml(formulaRange))
                    || ((hadCompleteOwnedArtifacts || reuseExistingNumberedTableFormatting)
                        && IsManagedNumberedEquationTable(
                            document,
                            formulaRange,
                            metadata.FormulaId)),
                knownNumberedTable: knownNumberedTable,
                metadata: metadata,
                deferNativeOmmlShapeCreation:
                    deferLegacyNativeOmmlShapeCreation,
                deferNativeOmmlMetadataPersistence:
                    deferNativeOmmlMetadataPersistence,
                plannedOrdinal:
                    plannedNumberOrdinal,
                plannedPrefix:
                    plannedNumberPrefix,
                verifiedFreshBareOmml:
                    verifiedFreshBareOmml,
                preparedFreshBareOmml:
                    preparedFreshBareOmml);
            if (ContainsNativeOmml(formulaRange))
            {
                // Do not repaginate here. SEQ/REF updates below can still change the
                // Shape label width and Word's display paragraph geometry. One final
                // measurement is performed at the end of this reconciliation.
                // Legacy table/eqArr migration intentionally stops with a committed
                // pure-display formula and anchor; its posted Office turn creates and
                // finalizes the external REF Shape after replacement RCWs unwind.
            }
            TraceStage("configure-scaffold");

            // OLE -> OMML replacement deliberately removes this formula's three
            // numbering bookmarks before inserting the adjacent native equation,
            // then rebuilds them in the same known table. The formula did not move,
            // so the newly updated local SEQ/REF is already the correct ordinal;
            // treating the missing pre-rebuild bookmarks as an insertion would
            // renumber the entire suffix of a large document for no reason.
            if (rebuiltStableNumberingSlot)
                TraceStage("stable-slot-rebuild");

            // Editing a healthy numbered formula does not change its document
            // position or sequence ordinal. The previous path still enumerated
            // every SEQ/REF/field in the document for every Apply, which made a
            // six-formula document visibly block Word for one or two seconds.
            // Only insert/copy/numbering-structure changes need an order refresh.
            var currentVisibleNumberRefreshed = false;
            if (rebuiltWithPreservedNumberPlan)
            {
                // OLE -> OMML conversion replaces the physical host but not the
                // formula's document position or logical sequence slot. The caller
                // captured the old SEQ ordinal/prefix before removing its aliases,
                // and ConfigureNumberedDisplayFormula used that exact plan for the
                // new direct-SEQ cell. No suffix sequence or body REF changed, so a
                // document-wide caption/field inventory here is pure overhead.
                TraceStage("preserved-number-plan");
            }
            else if (rebuiltStableNumberingSlot)
            {
                // OLE -> OMML keeps the existing 1x3 table but deliberately removes
                // this formula's old VTEq/VTEqCap/VTEqNum artifacts before rebuilding
                // them. CreateNativeCaption therefore starts from a bare SEQ field.
                // The formula has not moved, but its *format* still has to be applied:
                // otherwise a document using chapter/section numbering silently falls
                // back to a continuous number after conversion. Re-plan the existing
                // captions against the document's current format and patch only the
                // formulas whose rendered number actually changes.
                var stableSlotChangedNumbers =
                    UpdateNativeEquationSequenceFieldsIncremental(document);
                TraceStage("stable-slot-format-refresh");
                if (stableSlotChangedNumbers.Count > 0)
                {
                    UpdateHealthyNativeCrossReferencesAfterRenumbering(
                        document,
                        stableSlotChangedNumbers);
                    currentVisibleNumberRefreshed =
                        stableSlotChangedNumbers.ContainsKey(metadata.FormulaId);
                    TraceStage("stable-slot-references");
                }
            }
            else if (numberingOrderMayHaveChanged || !hadCompleteOwnedArtifacts)
            {
                var appendedSequenceUpdated = TryUpdateAppendedNativeEquationSequenceField(
                    document,
                    metadata.FormulaId,
                    out var appendedNumberChanged);
                if (appendedSequenceUpdated)
                {
                    TraceStage("append-sequence-fast");
                    // A newly appended FormulaId cannot have pre-existing body
                    // cross-references. If its heading-aware caption changed,
                    // refresh only this formula owner's visible REF instead of scanning
                    // every Field in the document.
                    if (appendedNumberChanged)
                    {
                        UpdateEquationNumberFields(
                            document,
                            formulaHeightPoints,
                            formulaFontSizePoints,
                            metadata.FormulaId);
                        currentVisibleNumberRefreshed = true;
                    }
                }
                else
                {
                    // Two individual suffix field passes are acceptable for a tiny
                    // document, but become quadratic once dozens of numbered
                    // formulas are present. Large documents use the validated
                    // single-batch suffix path below first; the range path remains
                    // a compatibility fallback for small/legacy hosts.
                    var continuousRangeUpdated =
                        !ShouldPreferScopedSequenceSuffixUpdate(document)
                        && TryUpdateInsertedContinuousEquationFieldRanges(
                            document,
                            metadata.FormulaId);
                    if (continuousRangeUpdated)
                    {
                        TraceStage("continuous-range-fast");
                    }
                    else
                    {
                        var continuousSuffixUpdated = TryUpdateInsertedEquationSequenceSuffixWithinScope(
                            document,
                            metadata.FormulaId,
                            out var changedFormulaNumbers,
                            out var referencesAlreadyUpdatedFrom);
                        if (continuousSuffixUpdated)
                        {
                            TraceStage("continuous-suffix-fast");
                            if (changedFormulaNumbers.Count > 0)
                            {
                                UpdateHealthyNativeCrossReferencesAfterRenumbering(
                                    document,
                                    changedFormulaNumbers,
                                    referencesAlreadyUpdatedFrom >= 0
                                        ? referencesAlreadyUpdatedFrom
                                        : null);
                                TraceStage("continuous-suffix-references");
                            }
                        }
                        else
                        {
                            Dictionary<string, string> fallbackChangedFormulaNumbers;
                            IReadOnlyDictionary<string, int>? fallbackReferenceCounts = null;
                            IReadOnlyList<NativeEquationCaptionEntry> fallbackOpenXmlCaptions =
                                Array.Empty<NativeEquationCaptionEntry>();
                            IReadOnlyDictionary<string, int> knownFallbackReferenceCounts =
                                new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                            // Word 2021 terminates an active Custom UndoRecord when a
                            // main-story Range.WordOpenXML snapshot is requested. A
                            // create/edit transaction must therefore prove numbering
                            // ownership from the already-created bookmarks and fields
                            // instead of invoking the document-wide XML health pass.
                            // Outside a custom transaction, retain the mature OpenXML
                            // validation and its faster reference-count inventory.
                            var customUndoActive = IsCustomUndoRecordActive(document);
                            var healthyOpenXmlRead = !customUndoActive
                                && TryReadHealthyEquationNumberArtifactsFromOpenXml(
                                    document,
                                    out fallbackOpenXmlCaptions,
                                    out knownFallbackReferenceCounts);
                            if (customUndoActive)
                            {
                                var nativeSequenceName = GetNativeEquationSequenceName(document);
                                if (!TryReadCompleteEquationNumberArtifactsFromBookmarks(
                                        document,
                                        nativeSequenceName,
                                        out var transactionCaptions)
                                    || !transactionCaptions.Any(caption => string.Equals(
                                        caption.FormulaId,
                                        metadata.FormulaId,
                                        StringComparison.OrdinalIgnoreCase)))
                                    throw new InvalidDataException(
                                        "VisualTeX could not prove a complete equation-number bookmark scaffold without ending Word's undo transaction.");

                                var transactionFormat = ReadEquationNumberFormat(document);
                                fallbackChangedFormulaNumbers =
                                    UpdateNativeEquationSequenceFieldsIncremental(
                                        document,
                                        nativeSequenceName,
                                        transactionCaptions,
                                        transactionFormat,
                                        trustedHealthyDirectTables: false);
                                TraceStage("sequence-fallback-custom-undo");
                            }
                            else if (healthyOpenXmlRead)
                            {
                                var fallbackFormat = ReadEquationNumberFormat(document);
                                var fallbackCaptions = fallbackFormat.UsesHeading
                                    ? ResolveNativeEquationCaptionPositions(
                                        document,
                                        fallbackOpenXmlCaptions)
                                    : fallbackOpenXmlCaptions;
                                if (fallbackCaptions.Count == fallbackOpenXmlCaptions.Count)
                                {
                                    fallbackChangedFormulaNumbers =
                                        UpdateNativeEquationSequenceFieldsIncremental(
                                            document,
                                            GetNativeEquationSequenceName(document),
                                            fallbackCaptions,
                                            fallbackFormat,
                                            trustedHealthyDirectTables: true);
                                    fallbackReferenceCounts = knownFallbackReferenceCounts;
                                    TraceStage("sequence-fallback-healthy");
                                }
                                else
                                {
                                    fallbackChangedFormulaNumbers =
                                        UpdateNativeEquationSequenceFieldsIncremental(document);
                                    TraceStage("sequence-fallback");
                                }
                            }
                            else
                            {
                                fallbackChangedFormulaNumbers =
                                    UpdateNativeEquationSequenceFieldsIncremental(document);
                                TraceStage("sequence-fallback");
                            }
                            if (fallbackChangedFormulaNumbers.Count > 0)
                            {
                                UpdateHealthyNativeCrossReferencesAfterRenumbering(
                                    document,
                                    fallbackChangedFormulaNumbers,
                                    knownReferenceCounts: fallbackReferenceCounts);
                                TraceStage("cross-reference-fallback");
                            }
                        }
                    }
                }
            }

            var containsNativeOmml = ContainsNativeOmml(formulaRange);
            // Current numbered OMML is the direct-SEQ 1x3 host. Validate that
            // local structure first; only non-table legacy formulas need the much
            // more expensive retired mathematical #(SEQ) metadata/OpenXML probe.
            // On a 100-OMML document this avoids over a second of work after every
            // newly numbered formula.
            var isNativeDirectTable = containsNativeOmml
                && (visibleNumberCreated
                    || IsHealthyNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        metadata.FormulaId,
                        updateField: false));
            var isNativeHashSequence = containsNativeOmml
                && !isNativeDirectTable
                && IsNumberedNativeOmmlHashSequenceFormula(
                    document,
                    metadata.FormulaId);
            // Legacy OLE/table/Shape hosts display their number through a generated
            // REF range and therefore still need local REF formatting. Native OMML
            // displays the SEQ result directly inside its managed host; applying
            // the old REF path only enumerates/normalizes the equation again.
            if (!visibleNumberCreated
                && !currentVisibleNumberRefreshed
                && !isNativeHashSequence
                && !isNativeDirectTable)
            {
                UpdateEquationNumberFields(
                    document,
                    formulaHeightPoints,
                    formulaFontSizePoints,
                    metadata.FormulaId);
            }
            if (containsNativeOmml
                && !deferNativeOmmlShapeFinalization
                && !isNativeHashSequence
                && !isNativeDirectTable)
                TryFinalizeNativeDisplayNumberShapeLayout(
                    document,
                    metadata.FormulaId);
            TraceStage("visible-number");
            return;
        }

        if (!hadNumberingArtifacts)
        {
            // A VisualTeX display OLE uses the same center/right tab paragraph as
            // a numbered display formula; the unnumbered form simply omits the
            // trailing tab/number artifacts. Keeping one paragraph-layout model
            // prevents Word from giving the paragraph mark a different line-box
            // placement when numbering is toggled.
            var containsNativeOmml = ContainsNativeOmml(formulaRange);
            ConfigureEquationParagraph(formulaRange, numbered: false);
            if (!containsNativeOmml)
            {
                EnsureLeadingEquationTab(document, formulaRange);
                AlignDisplayOleObjectToParagraphBaseline(
                    document,
                    formulaRange,
                    formulaHeightPoints,
                    metadata);
            }
            return;
        }

        IReadOnlyList<NativeEquationCaptionEntry>? removalSnapshot = null;
        HashSet<string>? removalExternalReferenceIds = null;
        var canUseTargetedDirectRemoval = false;
        if (!string.IsNullOrWhiteSpace(preparedUnnumberedOmml)
            && TryReadTargetedRemovalSnapshot(
                document,
                metadata.FormulaId,
                out var targetedRemovalCaptions,
                out var targetedExternalReferenceIds))
        {
            removalSnapshot = targetedRemovalCaptions;
            removalExternalReferenceIds = targetedExternalReferenceIds;
            canUseTargetedDirectRemoval = true;
        }
        TraceStage("unnumber-snapshot");

        ConfigureUnnumberedDisplayFormula(
            document,
            formulaRange,
            metadata.FormulaId,
            metadata,
            formulaHeightPoints,
            preparedUnnumberedOmml,
            deferMetadataPersistence: deferNativeOmmlMetadataPersistence);
        TraceStage("unnumber-local-structure");

        if (canUseTargetedDirectRemoval
            && removalSnapshot is not null
            && removalExternalReferenceIds is not null
            && TryUpdateRemovedEquationSequenceSuffixWithinScope(
                document,
                metadata.FormulaId,
                removalSnapshot,
                out var removalChangedNumbers))
        {
            TraceStage("unnumber-suffix-fast");
            if (removalChangedNumbers.Keys.Any(
                    removalExternalReferenceIds.Contains))
            {
                UpdateHealthyNativeCrossReferencesAfterRenumbering(
                    document,
                    removalChangedNumbers);
                TraceStage("unnumber-references-fast");
            }
            return;
        }

        // Legacy, damaged, externally referenced, or otherwise ambiguous hosts
        // retain the mature comprehensive reconciliation behavior.
        UpdateNativeEquationSequenceFields(document);
        UpdateMainStoryFields(document);
        UpdateNativeCrossReferences(document);
    }

    private static bool HasFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId))
                || bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                || bookmarks.Exists(NativeNumberBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    internal static bool HasCompleteFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            return bookmarks.Exists(EquationBookmarkName(formulaId))
                && bookmarks.Exists(NativeCaptionBookmarkName(formulaId))
                && bookmarks.Exists(NativeNumberBookmarkName(formulaId));
        }
        catch { return false; }
        finally { Release(bookmarks); }
    }

    internal static Range? FindNativeEquationCaptionRange(
        Document document,
        string formulaId)
    {
        Range? captionRange = null;
        Range? numberRange = null;
        try
        {
            if (!TryGetNativeCaptionRanges(
                    document,
                    formulaId,
                    GetNativeEquationSequenceName(document),
                    out captionRange,
                    out numberRange)
                || captionRange is null)
                return null;
            var result = captionRange.Duplicate;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(numberRange);
            Release(captionRange);
        }
    }

    internal static Range? FindVisibleEquationNumberRange(
        Document document,
        string formulaId)
    {
        // New VisualTeX body numbering is self-contained in the same paragraph
        // as the OLE and intentionally owns no always-on visible-number bookmark.
        // Resolve that field tree first; the legacy bookmark paths below remain
        // only for older documents and OMML migration compatibility.
        var selfContained =
            WordVisualTeXParagraphNumbering
                .FindVisibleLabelRange(
                    document,
                    formulaId);
        if (selfContained is not null)
            return selfContained;

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Bookmark? formulaBookmark = null;
        Bookmark? repairedBookmark = null;
        Range? formulaRange = null;
        Field? nativeNumberField = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(name))
            {
                bookmark = bookmarks[name];
                range = bookmark.Range.Duplicate;

                // Word 2021 can expand a bookmark ending at the paragraph boundary
                // across the paragraph terminator. Around fields, Range.Text and
                // story positions are not one-to-one, so trimming End by character
                // count is unsafe. Canonicalize against the first owning paragraph's
                // exact editable-body boundary instead. A visible equation-number
                // label is never allowed to own a paragraph/cell terminator.
                Paragraphs? visibleParagraphs = null;
                Paragraph? visibleParagraph = null;
                Range? visibleParagraphRange = null;
                Range? trailing = null;
                try
                {
                    visibleParagraphs = range.Paragraphs;
                    if (visibleParagraphs.Count > 0)
                    {
                        visibleParagraph = visibleParagraphs[1];
                        visibleParagraphRange = visibleParagraph.Range.Duplicate;
                        var terminatorLength =
                            GetWordParagraphStructuralTerminatorLength(
                                visibleParagraphRange);
                        var editableEnd = Math.Max(
                            visibleParagraphRange.Start,
                            visibleParagraphRange.End - terminatorLength);
                        if (range.Start < editableEnd
                            && range.End > editableEnd)
                            range.End = editableEnd;
                    }

                    // Defensive normalization for unusual Word stories whose
                    // paragraph owner does not expose a structural terminator.
                    while (range.End > range.Start)
                    {
                        Release(trailing);
                        trailing = document.Range(range.End - 1, range.End);
                        var trailingText = trailing.Text ?? string.Empty;
                        if (!string.Equals(
                                trailingText,
                                "\r",
                                StringComparison.Ordinal)
                            && !string.Equals(
                                trailingText,
                                "\a",
                                StringComparison.Ordinal))
                            break;
                        range.End -= 1;
                    }
                }
                finally
                {
                    Release(trailing);
                    Release(visibleParagraphRange);
                    Release(visibleParagraph);
                    Release(visibleParagraphs);
                }

                var bookmarkedResult = range;
                range = null;
                return bookmarkedResult;
            }

            // Word may discard a bookmark nested inside a professional OMath when
            // it normalizes or serializes the equation. The generated REF field is
            // still durable, so recover its result from the formula identity rather
            // than treating the missing convenience bookmark as number loss.
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is null) return null;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            if (!WordOmmlEquationNumbers.HasVisualTeXNativeEquationNumber(
                    WordOmmlConverter.ReadLocalEquationXml(formulaRange)))
                return null;
            nativeNumberField = FindNativeOmmlNumberReferenceField(
                document,
                formulaRange,
                formulaId);
            if (nativeNumberField is null) return null;
            range = nativeNumberField.Result.Duplicate;
            if (!document.ReadOnly)
            {
                try { repairedBookmark = bookmarks.Add(name, range); }
                catch { }
            }
            var recoveredResult = range;
            range = null;
            return recoveredResult;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(range);
            Release(nativeNumberField);
            Release(formulaRange);
            Release(repairedBookmark);
            Release(formulaBookmark);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Range? FindVisibleEquationNumberTextRange(
        Document document,
        string formulaId)
    {
        Range? range = null;
        try
        {
            range = FindVisibleEquationNumberRange(document, formulaId);
            if (range is null) return null;
            var text = range.Text ?? string.Empty;
            var trimStart = 0;
            while (trimStart < text.Length && text[trimStart] == '\t')
                trimStart++;
            var trimEnd = 0;
            while (trimEnd < text.Length - trimStart)
            {
                var character = text[text.Length - 1 - trimEnd];
                if (character != '\r' && character != '\a') break;
                trimEnd++;
            }
            if (trimStart > 0 || trimEnd > 0)
            {
                var start = Math.Min(range.End, range.Start + trimStart);
                var end = Math.Max(start, range.End - trimEnd);
                range.SetRange(start, end);
            }
            var result = range;
            range = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(range);
        }
    }

    internal static Range? FindNumberingOwnerRange(
        Document document,
        string formulaId)
    {
        var selfContainedOwner =
            WordVisualTeXParagraphNumbering
                .FindOwnerParagraphRange(
                    document,
                    formulaId);
        if (selfContainedOwner is not null)
            return selfContainedOwner;

        Range? visibleRange = null;
        Tables? tables = null;
        Table? table = null;
        Columns? columns = null;
        Range? ownerRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        try
        {
            visibleRange = FindVisibleEquationNumberRange(document, formulaId);
            if (visibleRange is null) return null;
            var nativeHashSequenceHost =
                HasManagedNativeOmmlHashSequenceHost(document, formulaId);
            if (visibleRange.StoryType == WdStoryType.wdTextFrameStory)
            {
                ownerRange = TryResolveNativeDisplayFormulaOwnerRange(
                    document,
                    formulaId);
                if (ownerRange is not null)
                {
                    var trueDisplayResult = ownerRange;
                    ownerRange = null;
                    return trueDisplayResult;
                }
            }
            try
            {
                if (!nativeHashSequenceHost
                    && RangeIsWhollyWithinTable(visibleRange))
                {
                    // A user cell in column three can own a normal OLE + REF
                    // paragraph. Only a proven managed number-only third cell
                    // may widen its owner to the whole legacy equation row.
                    table = FindNumberedEquationTable(document, formulaId);
                    if (table is not null)
                    {
                        columns = table.Columns;
                        if (columns.Count == 3
                            && TryGetManagedNumberTableRowIndex(table, visibleRange, 3, out var rowIndex))
                        {
                            Row? row = null;
                            try
                            {
                                row = table.Rows[rowIndex];
                                ownerRange = row.Range.Duplicate;
                            }
                            finally { Release(row); }
                            var tableResult = ownerRange;
                            ownerRange = null;
                            return tableResult;
                        }
                    }
                }
            }
            catch
            {
                // Fall through to the paragraph owner. This is the native
                // MathType-style layout used by numbered VisualTeX OLE formulas.
            }

            paragraphs = visibleRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range.Duplicate;
            var paragraphResult = ownerRange;
            ownerRange = null;
            return paragraphResult;
        }
        catch (Exception error)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                    "1",
                    StringComparison.Ordinal))
            {
                TraceNumberingPerformance(
                    $"[perf] numbering-owner-range-failed formulaId={formulaId} "
                    + $"error={error.GetType().Name}:{error.Message}");
            }
            return null;
        }
        finally
        {
            Release(paragraph);
            Release(paragraphs);
            Release(ownerRange);
            Release(columns);
            Release(table);
            Release(tables);
            Release(visibleRange);
        }
    }

    internal static Table? FindNumberedEquationTable(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Tables? tables = null;
        Table? table = null;
        Columns? columns = null;
        Cell? numberCell = null;
        Range? numberCellRange = null;
        OMaths? numberCellMaths = null;
        InlineShapes? numberCellShapes = null;
        Cell? centerCell = null;
        Range? centerCellRange = null;
        OMaths? centerMaths = null;
        InlineShapes? centerShapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            tables = range.Tables;
            if (tables.Count == 0) return null;
            table = tables[1];
            columns = table.Columns;
            if (columns.Count != 3
                || !TryGetManagedNumberTableRowIndex(
                    table,
                    range,
                    expectedColumnIndex: 3,
                    out var rowIndex))
                return null;

            numberCell = table.Cell(rowIndex, 3);
            numberCellRange = numberCell.Range;
            numberCellMaths = numberCellRange.OMaths;
            numberCellShapes = numberCellRange.InlineShapes;
            if (numberCellMaths.Count != 0 || numberCellShapes.Count != 0)
                return null;

            centerCell = table.Cell(rowIndex, 2);
            centerCellRange = centerCell.Range;
            centerMaths = centerCellRange.OMaths;
            centerShapes = centerCellRange.InlineShapes;
            if (centerMaths.Count + centerShapes.Count != 1)
                return null;

            var result = table;
            table = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(centerShapes);
            Release(centerMaths);
            Release(centerCellRange);
            Release(centerCell);
            Release(numberCellShapes);
            Release(numberCellMaths);
            Release(numberCellRange);
            Release(numberCell);
            Release(columns);
            Release(table);
            Release(tables);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static Table? FindNumberedEquationTableByFormulaOwner(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        if (document is null || formulaRange is null
            || string.IsNullOrWhiteSpace(formulaId))
            return null;

        // The visible VTEq_* bookmark is the normal and cheapest ownership proof.
        // Word can nevertheless expand that bookmark into column 2 after an
        // equation object is replaced or a neighbouring invisible typing anchor is
        // removed. In that damaged-document state the generated 1x3 row is still
        // unambiguous: the formula's own durable identity owns the sole payload in
        // column 2 and column 3 owns exactly one VisualTeX SEQ field. Recover that
        // row so callers can remove it instead of leaving an empty table behind.
        var bookmarked = FindNumberedEquationTable(document, formulaId);
        if (bookmarked is not null) return bookmarked;

        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Cell? leftCell = null;
        Cell? formulaCell = null;
        Cell? numberCell = null;
        Range? leftRange = null;
        Range? formulaCellRange = null;
        Range? numberCellRange = null;
        OMaths? formulaMaths = null;
        InlineShapes? formulaShapes = null;
        InlineShape? formulaShape = null;
        OMaths? numberMaths = null;
        InlineShapes? numberShapes = null;
        Fields? numberFields = null;
        Field? numberField = null;
        Range? numberCode = null;
        Bookmark? ommlBookmark = null;
        Bookmarks? bookmarks = null;
        Bookmark? driftedVisibleBookmark = null;
        Range? driftedVisibleRange = null;
        Table? Fail(string reason)
        {
            TraceNumberingPerformance(
                $"[recovery] numbered-table-owner-rejected formulaId={formulaId} reason={reason}");
            return null;
        }
        try
        {
            tables = formulaRange.Tables;
            if (tables.Count != 1) return Fail($"formula-tables={tables.Count}");
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            if (columns.Count != 3
                || !TryGetManagedNumberTableRowIndex(
                    table,
                    formulaRange,
                    expectedColumnIndex: 2,
                    out var rowIndex))
                return Fail($"table-columns={columns.Count}-or-formula-row");

            leftCell = table.Cell(rowIndex, 1);
            formulaCell = table.Cell(rowIndex, 2);
            numberCell = table.Cell(rowIndex, 3);
            leftRange = leftCell.Range.Duplicate;
            formulaCellRange = formulaCell.Range.Duplicate;
            numberCellRange = numberCell.Range.Duplicate;
            if (!ContainsOnlyStructuralWordText(leftRange.Text))
                return Fail("left-cell-user-text");

            formulaMaths = formulaCellRange.OMaths;
            formulaShapes = formulaCellRange.InlineShapes;
            if (formulaMaths.Count + formulaShapes.Count != 1)
                return Fail(
                    $"center-payload-maths={formulaMaths.Count}-shapes={formulaShapes.Count}");
            if (formulaShapes.Count == 1)
            {
                formulaShape = formulaShapes[1];
                var metadata = WordFormulaMetadataReader.IsNativeOle(formulaShape)
                    ? WordFormulaMetadataReader.TryReadEmbeddedNativeOle(formulaShape)
                    : null;
                if (metadata is null
                    || !string.Equals(
                        metadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    return Fail(
                        $"ole-identity={(metadata?.FormulaId ?? "missing")}");
            }
            else
            {
                ommlBookmark = WordOmmlFormulaStore.FindAtRange(document, formulaRange);
                var metadata = ommlBookmark is null
                    ? null
                    : WordOmmlFormulaStore.TryRead(document, ommlBookmark);
                if (metadata is null
                    || !string.Equals(
                        metadata.FormulaId,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                    return Fail(
                        $"omml-identity={(metadata?.FormulaId ?? "missing")}");
            }

            numberMaths = numberCellRange.OMaths;
            numberShapes = numberCellRange.InlineShapes;
            numberFields = numberCellRange.Fields;
            if (numberMaths.Count != 0
                || numberShapes.Count != 0)
                return Fail(
                    $"number-payload-maths={numberMaths.Count}-shapes={numberShapes.Count}-fields={numberFields.Count}");
            if (numberFields.Count == 1)
            {
                numberField = numberFields[1];
                numberCode = numberField.Code;
                if (!IsVisualTeXSequenceFieldCode(numberCode.Text))
                    return Fail($"number-field-code={numberCode.Text}");
                var visibleNumberText = NormalizeNativeEquationNumberText(
                    numberCellRange.Text);
                if (!visibleNumberText.StartsWith("(", StringComparison.Ordinal)
                    || !visibleNumberText.EndsWith(")", StringComparison.Ordinal))
                    return Fail($"number-text={visibleNumberText}");
            }
            else if (numberFields.Count == 0)
            {
                // The same Word mutation that expands VTEq_* into the formula cell
                // can consume the generated right-cell field first. Recover only
                // the exact one-row scaffold whose left/right cells are otherwise
                // empty and whose formula-specific visible bookmark now owns the
                // same column-2 row. This conjunction cannot match an ordinary
                // user table merely because it contains an equation.
                if (rows.Count != 1
                    || !ContainsOnlyStructuralWordText(numberCellRange.Text))
                    return Fail(
                        $"number-field-missing-rows={rows.Count}-text={NormalizeNativeEquationNumberText(numberCellRange.Text)}");
                bookmarks = document.Bookmarks;
                var visibleNumberName = EquationBookmarkName(formulaId);
                if (!bookmarks.Exists(visibleNumberName))
                    return Fail("number-field-and-visible-bookmark-missing");
                driftedVisibleBookmark = bookmarks[visibleNumberName];
                driftedVisibleRange = driftedVisibleBookmark.Range.Duplicate;
                if (!TryGetManagedNumberTableRowIndex(
                        table,
                        driftedVisibleRange,
                        expectedColumnIndex: 2,
                        out var driftedRowIndex)
                    || driftedRowIndex != rowIndex)
                    return Fail("missing-field-visible-bookmark-not-in-formula-cell");
            }
            else
            {
                return Fail($"number-fields={numberFields.Count}");
            }

            TraceNumberingPerformance(
                $"[recovery] numbered-table-owned-by-formula formulaId={formulaId} row={rowIndex} reason=visible-bookmark-drift");
            var result = table;
            table = null;
            return result;
        }
        catch
        {
            return Fail("com-error");
        }
        finally
        {
            Release(driftedVisibleRange);
            Release(driftedVisibleBookmark);
            Release(bookmarks);
            Release(ommlBookmark);
            Release(numberCode);
            Release(numberField);
            Release(numberFields);
            Release(numberShapes);
            Release(numberMaths);
            Release(formulaShape);
            Release(formulaShapes);
            Release(formulaMaths);
            Release(numberCellRange);
            Release(formulaCellRange);
            Release(leftRange);
            Release(numberCell);
            Release(formulaCell);
            Release(leftCell);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
        }
    }

    internal static bool FormulaRangeOwnsNumberingArtifacts(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        Tables? formulaTables = null;
        Tables? numberTables = null;
        Table? formulaTable = null;
        Table? numberTable = null;
        Range? formulaTableRange = null;
        Range? numberTableRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var numberName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(numberName)) return false;
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;

            if (ContainsNativeOmml(formulaRange)
                && NativeDisplayNumberShapeBelongsToFormula(
                    document,
                    formulaRange,
                    formulaId))
                return true;

            formulaTables = formulaRange.Tables;
            numberTables = numberRange.Tables;
            if (formulaTables.Count == 0 || numberTables.Count == 0)
                return (formulaRange.Start <= numberRange.Start
                        && formulaRange.End >= numberRange.End)
                    || (formulaTables.Count == 0 && numberTables.Count == 0
                        && OwnsLocalOleNumberReference(formulaRange, numberRange, formulaId));

            formulaTable = formulaTables[1];
            numberTable = numberTables[1];
            formulaTableRange = formulaTable.Range;
            numberTableRange = numberTable.Range;
            return formulaTableRange.Start == numberTableRange.Start;
        }
        catch { return false; }
        finally
        {
            Release(numberTableRange);
            Release(formulaTableRange);
            Release(numberTable);
            Release(formulaTable);
            Release(numberTables);
            Release(formulaTables);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static bool OwnsLocalOleNumberReference(Range formulaRange, Range numberRange, string formulaId)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? host = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? actual = null;
        Fields? fields = null;
        try
        {
            // In the current tab layout the visible REF is beside the OLE, not
            // inside its Range. Requiring containment misclassified every healthy
            // tab host as a new insertion, causing full-document numbering on Edit.
            if (formulaRange.StoryType != numberRange.StoryType) return false;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1]; host = paragraph.Range;
            if (numberRange.Start < formulaRange.End || numberRange.End > host.End) return false;
            shapes = host.InlineShapes;
            if (shapes.Count != 1) return false;
            shape = shapes[1]; actual = shape.Range;
            if (actual.Start != formulaRange.Start || actual.End != formulaRange.End
                || !WordFormulaMetadataReader.IsNativeOle(shape)) return false;
            fields = host.Fields;
            for (var i = 1; i <= fields.Count; i++)
            {
                Field? field = null; Range? code = null; Range? result = null;
                try
                {
                    field = fields[i];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!IsReferenceToBookmark(code.Text, NativeNumberBookmarkName(formulaId))) continue;
                    result = field.Result;
                    if (result.Start >= numberRange.Start && result.End <= numberRange.End) return true;
                }
                finally { Release(result); Release(code); Release(field); }
            }
            return false;
        }
        catch { return false; }
        finally { Release(fields); Release(actual); Release(shape); Release(shapes); Release(host); Release(paragraph); Release(paragraphs); }
    }

    internal static void UpdateMainStoryFields(Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            UpdateVisualTeXOwnedFieldsInRange(content);
        }
        finally { Release(content); }
    }

    internal static int UpdateVisualTeXOwnedFieldsInRange(
        Range range,
        out int visualTeXReferenceCount)
    {
        visualTeXReferenceCount = 0;
        if (range is null) return 0;

        Document? document = null;
        Fields? fields = null;
        var updated = 0;
        try
        {
            document = range.Document;
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];

                    // Fast-reject the common OLE object itself before reading its
                    // code. In Word an embedded VisualTeX/MathType object is exposed
                    // as an EMBED field; updating it is unnecessary and can trigger
                    // the "fields may refer to other files" confirmation dialog.
                    if (field.Type == WdFieldType.wdFieldEmbed)
                        continue;

                    code = field.Code;
                    var codeText = code.Text ?? string.Empty;
                    if (IsVisualTeXSequenceFieldCode(codeText))
                    {
                        field.Update();
                        updated++;
                        continue;
                    }

                    if (field.Type != WdFieldType.wdFieldRef
                        || !TryResolveVisualTeXReferenceBookmark(
                            document,
                            codeText,
                            out var bookmarkName)
                        || !TryFormulaIdFromBookmark(
                            bookmarkName,
                            NativeNumberBookmarkPrefix,
                            out _))
                        continue;

                    WordEquationReferenceFields.RefreshFieldPreservingFormatting(field);
                    visualTeXReferenceCount++;
                    updated++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return updated;
        }
        finally
        {
            Release(fields);
            Release(document);
        }
    }

    internal static int UpdateVisualTeXOwnedFieldsInRange(Range range) =>
        UpdateVisualTeXOwnedFieldsInRange(range, out _);

    public static int Reconcile(Document document)
    {
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
            "1",
            StringComparison.Ordinal);
        var performanceWatch = tracePerformance
            ? System.Diagnostics.Stopwatch.StartNew()
            : null;
        var performanceCheckpoint = 0L;
        void TraceReconcileStage(string stage)
        {
            if (performanceWatch is null) return;
            var elapsed = performanceWatch.ElapsedMilliseconds;
            TraceNumberingPerformance(
                $"[perf] Reconcile.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }

        var numberedFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var adoptedNativeHashCopies =
            AdoptUnownedCopiesOfManagedNativeHashSequenceHosts(document);
        if (adoptedNativeHashCopies > 0)
            TraceNumberingPerformance(
                $"[perf] Reconcile.adopt-native-hash-copies count={adoptedNativeHashCopies}");
        TraceReconcileStage("adopt-native-hash-copies");

        // Freeze and verify every display-OMML identity before the first numbering
        // paragraph mutation, then reconcile from the end of the document toward
        // the start. Later tab/caption insertions therefore cannot drag the collapsed
        // bookmark of an equation that has not yet been processed.
        var ommlFormulaIds = GetOmmlDisplayFormulaIdsForStructuralEdit(document);
        TraceReconcileStage("capture-omml-ids");
        InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            var inlineCount = inlineShapes.Count;
            var numberedFormulaLocations = CaptureNumberedFormulaLocations(
                document,
                inlineShapes,
                ommlFormulaIds);
            TraceReconcileStage("capture-numbered-locations");
            RepairSharedNativeCaptionArtifacts(document, numberedFormulaLocations);
            TraceReconcileStage("repair-shared-captions");

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                Range? formulaRange = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange = shape.Range;
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            shape.Height,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            metadata.FormulaId,
                            metadata: metadata);
                        numberedFormulaIds.Add(metadata.FormulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(
                            document,
                            formulaRange,
                            metadata.FormulaId,
                            metadata,
                            shape.Height);
                    }
                }
                finally
                {
                    Release(formulaRange);
                    Release(shape);
                }
            }
            TraceReconcileStage("configure-inline-formulas");

            foreach (var formulaId in ommlFormulaIds)
            {
                Range? formulaRange = null;
                try
                {
                    var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange =
                        WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            formulaId,
                            metadata);
                    var formulaHeightPoints =
                        WordOmmlFormulaStore.EstimateHeightPoints(formulaRange);
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            formulaHeightPoints,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            formulaId,
                            metadata: metadata);
                        numberedFormulaIds.Add(formulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(document, formulaRange, formulaId);
                    }
                }
                finally { Release(formulaRange); }
            }
            TraceReconcileStage("configure-omml-formulas");

            if (numberedFormulaIds.Count > 0)
                EnsureDocumentEquationNumberFormatPreference(document);
            TraceReconcileStage("number-format-preference");

            RemoveOrphanEquationArtifacts(document, numberedFormulaIds);
            TraceReconcileStage("remove-orphan-artifacts");
            RebuildNativeNumberBookmarksFromCaptions(
                document,
                numberedFormulaIds);
            TraceReconcileStage("rebuild-number-bookmarks-1");

            // Word caches SEQ results independently from REF results. After a
            // numbered formula is deleted, refresh every native Equation SEQ
            // field in document order before updating any visible or body REF
            // field. Otherwise a REF can continue displaying the removed
            // formula's old ordinal until Word performs a later global update.
            UpdateNativeEquationSequenceFields(document);
            TraceReconcileStage("update-native-sequences");
            // Writing chapter/section prefixes mutates the hidden caption
            // paragraphs. Word can invalidate a number bookmark whose range was
            // created before that mutation even though the SEQ field and caption
            // bookmark survive. Re-wrap every final SEQ result after all caption
            // text is stable so each formula keeps an independent REF target.
            RebuildNativeNumberBookmarksFromCaptions(
                document,
                numberedFormulaIds);
            TraceReconcileStage("rebuild-number-bookmarks-2");
            // Newly batch-numbered formulas create their visible REF fields
            // before the native number bookmarks are rewritten with the final
            // chapter/section prefix. Word can temporarily cache "reference
            // source not found" for those fresh fields. Refresh the complete
            // main story once after all target bookmarks are stable, then apply
            // formula-specific alignment below.
            UpdateMainStoryFields(document);
            TraceReconcileStage("update-main-story-fields");

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            shape.Height,
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            metadata.FormulaId);
                }
                finally { Release(shape); }
            }
            TraceReconcileStage("update-inline-number-fields");
            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            WordOmmlFormulaStore.EstimateHeightPoints(bookmark),
                            (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                            formulaId);
                }
                finally { Release(bookmark); }
            }
            TraceReconcileStage("update-omml-number-fields");
            UpdateNativeCrossReferences(document);
            TraceReconcileStage("update-cross-references");
        }
        finally { Release(inlineShapes); }

        return numberedFormulaIds.Count;
    }

    private sealed class OmmlStructuralFormulaEntry
    {
        public OmmlStructuralFormulaEntry(string formulaId, int position)
        {
            FormulaId = formulaId;
            Position = position;
        }

        public string FormulaId { get; }
        public int Position { get; }
    }

    private static IReadOnlyList<string> GetOmmlDisplayFormulaIdsForStructuralEdit(
        Document document)
    {
        var entries = new List<OmmlStructuralFormulaEntry>();
        var formulaIds = WordOmmlFormulaStore.BookmarkedFormulaIds(document);
        foreach (var formulaId in formulaIds)
        {
            Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.DisplayMode != "block") continue;
                range = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                entries.Add(new OmmlStructuralFormulaEntry(formulaId, range.Start));
            }
            finally { Release(range); }
        }
        var ordered = entries
            .OrderByDescending(entry => entry.Position)
            .ToArray();
        return ordered
            .Select(entry => entry.FormulaId)
            .ToArray();
    }
}

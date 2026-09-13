using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private sealed class NewVisualTeXNumberingEntry
    {
        internal FormulaMetadata Metadata { get; set; } = new();
        internal int Position { get; set; }
        internal string RangeHint { get; set; } = string.Empty;
    }

    private static void CaptureTableInlineFollowingBoundary(
        Document document,
        Range formulaRange,
        FormulaMetadata metadata,
        WordFormulaFormatConversionTarget target)
    {
        if (!string.Equals(metadata.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase)
            || !WordEquationNumbering.RangeIsWhollyWithinTable(formulaRange))
            return;

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? suffix = null;
        Range? firstCharacter = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (formulaRange.End >= bodyEnd) return;
            var end = Math.Min(bodyEnd, formulaRange.End + 16);
            suffix = document.Range(formulaRange.End, end);
            // Freeze only ordinary prose between objects. A later formula in the
            // same cell/paragraph is converted first and legitimately changes its
            // Word text/field encoding. It must never become the comparison
            // context for this formula's whitespace boundary.
            OMaths? followingMaths = null;
            InlineShapes? followingShapes = null;
            Fields? followingFields = null;
            try
            {
                followingMaths = suffix.OMaths;
                for (var index = 1; index <= followingMaths.Count; index++)
                {
                    OMath? math = null;
                    Range? range = null;
                    try
                    {
                        math = followingMaths[index]; range = math.Range;
                        if (range.Start >= formulaRange.End) end = Math.Min(end, range.Start);
                    }
                    finally { Release(range); Release(math); }
                }
                followingShapes = suffix.InlineShapes;
                for (var index = 1; index <= followingShapes.Count; index++)
                {
                    InlineShape? shape = null;
                    Range? range = null;
                    try
                    {
                        shape = followingShapes[index]; range = shape.Range;
                        if (range.Start >= formulaRange.End) end = Math.Min(end, range.Start);
                    }
                    finally { Release(range); Release(shape); }
                }
                followingFields = WordFormulaHost.GetLocalFields(suffix);
                for (var index = 1; index <= followingFields.Count; index++)
                {
                    Field? field = null;
                    Range? code = null;
                    try
                    {
                        field = followingFields[index]; code = field.Code;
                        if (code.Start - 1 >= formulaRange.End) end = Math.Min(end, code.Start - 1);
                    }
                    finally { Release(code); Release(field); }
                }
            }
            finally { Release(followingFields); Release(followingShapes); Release(followingMaths); }
            if (end <= formulaRange.End) return;
            suffix.SetRange(formulaRange.End, end);
            var text = suffix.Text ?? string.Empty;
            if (text.Length == 0) return;
            var first = text[0];
            if (!char.IsWhiteSpace(first)
                || first is '\r' or '\n' or '\a' or '\v')
                return;
            firstCharacter = document.Range(formulaRange.End, formulaRange.End + 1);
            target.FollowingInlineTableText = text;
            target.FollowingInlineTableFormatting = WordCharacterFormatting.Capture(firstCharacter);
        }
        finally
        {
            Release(firstCharacter);
            Release(suffix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RestoreTableInlineFollowingBoundary(
        Document document,
        Range targetRange,
        WordFormulaFormatConversionTarget target)
    {
        var expected = target.FollowingInlineTableText;
        if (string.IsNullOrEmpty(expected)) return;

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? probe = null;
        Range? insertion = null;
        Range? restored = null;
        try
        {
            paragraphs = targetRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("The converted inline table formula no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            var start = targetRange.End;
            var probeEnd = Math.Min(bodyEnd, start + expected.Length);
            var actual = string.Empty;
            if (probeEnd > start)
            {
                probe = document.Range(start, probeEnd);
                actual = probe.Text ?? string.Empty;
            }
            if (actual.StartsWith(expected, StringComparison.Ordinal)) return;

            var expectedAfterFirst = expected.Length > 1 ? expected.Substring(1) : string.Empty;
            if (expected.Length > 1
                && !actual.StartsWith(expectedAfterFirst, StringComparison.Ordinal))
                throw new InvalidDataException("Text following the converted inline table formula changed by more than one boundary character.");

            insertion = document.Range(start, start);
            insertion.Text = expected.Substring(0, 1);
            restored = document.Range(start, start + 1);
            target.FollowingInlineTableFormatting?.Apply(restored);

            Release(probe);
            probe = null;
            var verifyEnd = Math.Min(bodyEnd + 1, start + expected.Length);
            if (verifyEnd > start)
            {
                probe = document.Range(start, verifyEnd);
                if (!(probe.Text ?? string.Empty).StartsWith(expected, StringComparison.Ordinal))
                    throw new InvalidDataException("The inline table boundary character could not be restored exactly.");
            }
        }
        finally
        {
            Release(restored);
            Release(insertion);
            Release(probe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RestoreTableInlineFollowingBoundaryForOleTarget(
        Document document,
        WordFormulaFormatConversionTarget target,
        string targetMode,
        string targetFormulaId,
        string? mathTypeBookmarkName)
    {
        if (string.IsNullOrEmpty(target.FollowingInlineTableText)) return;

        InlineShape? shape = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? targetRange = null;
        try
        {
            if (string.Equals(
                    targetMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindByFormulaId(
                        document,
                        targetFormulaId,
                        sourceObjectIdHint: null,
                        allowGlobalFallback: false)
                    ?? throw new InvalidDataException(
                        "The converted VisualTeX inline table formula lost its physical owner before boundary restoration.");
                targetRange = shape.Range.Duplicate;
            }
            else if (string.Equals(
                         targetMode,
                         FormulaOleContract.MathTypeOleMode,
                         StringComparison.Ordinal))
            {
                if (string.IsNullOrWhiteSpace(mathTypeBookmarkName))
                    throw new InvalidDataException(
                        "The converted MathType inline table formula has no temporary owner bookmark for boundary restoration.");
                bookmarks = document.Bookmarks;
                if (!bookmarks.Exists(mathTypeBookmarkName))
                    throw new InvalidDataException(
                        "The converted MathType inline table formula lost its temporary owner bookmark before boundary restoration.");
                bookmark = bookmarks[mathTypeBookmarkName];
                targetRange = bookmark.Range.Duplicate;
            }
            else
            {
                return;
            }

            RestoreTableInlineFollowingBoundary(document, targetRange, target);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-table-inline-boundary-restored targetMode={targetMode} formulaId={targetFormulaId} textLength={target.FollowingInlineTableText!.Length}");
        }
        finally
        {
            Release(targetRange);
            Release(bookmark);
            Release(bookmarks);
            Release(shape);
        }
    }

    private int BuildNewVisualTeXNumberingBatch(
        Document document,
        IReadOnlyCollection<FormulaMetadata> expectedMetadata,
        bool normalizeMathTypeStyle = false) =>
        BuildNewVisualTeXNumberingBatch(
            document,
            expectedMetadata,
            normalizeMathTypeStyle,
            usePlannedFieldFastPath: false,
            out _);

    private int BuildNewVisualTeXNumberingBatch(
        Document document,
        IReadOnlyCollection<FormulaMetadata> expectedMetadata,
        bool normalizeMathTypeStyle,
        bool usePlannedFieldFastPath,
        out bool numberingFinalizedInBatch)
    {
        numberingFinalizedInBatch = false;
        var convertedFormulaCount = expectedMetadata.Count;
        var entries = new List<NewVisualTeXNumberingEntry>();
        var targetByFormulaId = expectedMetadata.ToDictionary(
            metadata => metadata.FormulaId, StringComparer.OrdinalIgnoreCase);

        // VTO_ bookmarks are excellent single-operation locators, but migrating an
        // earlier legacy numbered host can make Word expand a later target bookmark
        // across the replaced host/caption boundary. Inventory the actual OLE hosts
        // once, before any numbering structure is created, using their embedded
        // FormulaId as the source of truth. No InlineShapes collection is retained
        // across a structural edit.
        InlineShapes? inventoryShapes = null;
        try
        {
            inventoryShapes = document.InlineShapes;
            var inventoryCount = inventoryShapes.Count;
            for (var index = 1; index <= inventoryCount; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = inventoryShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape);
                    if (metadata is null
                        || !targetByFormulaId.TryGetValue(
                            metadata.FormulaId,
                            out var target))
                        continue;
                    if (metadata.Numbered != target.Numbered
                        || !string.Equals(
                            metadata.DisplayMode,
                            target.DisplayMode,
                            StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException(
                            $"Converted VisualTeX formula {metadata.FormulaId} lost its display/numbered state before numbering finalization.");
                    range = shape.Range;
                    Paragraphs? inventoryParagraphs = null;
                    Paragraph? inventoryParagraph = null;
                    Range? inventoryParagraphRange = null;
                    try
                    {
                        inventoryParagraphs = range.Paragraphs;
                        if (inventoryParagraphs.Count != 1)
                            throw new InvalidOperationException(
                                $"Converted VisualTeX formula {metadata.FormulaId} spans {inventoryParagraphs.Count} paragraphs before numbering finalization.");
                        inventoryParagraph = inventoryParagraphs[1];
                        inventoryParagraphRange = inventoryParagraph.Range;
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-visualtex-numbering-inventory formulaId={metadata.FormulaId} range={range.Start}:{range.End} paragraph={inventoryParagraphRange.Start}:{inventoryParagraphRange.End} numbered={metadata.Numbered}");
                    }
                    finally
                    {
                        Release(inventoryParagraphRange);
                        Release(inventoryParagraph);
                        Release(inventoryParagraphs);
                    }
                    entries.Add(new NewVisualTeXNumberingEntry
                    {
                        Metadata = metadata,
                        Position = range.Start,
                        RangeHint = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                    });
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }
        }
        finally { Release(inventoryShapes); }

        if (entries.Count != convertedFormulaCount
            || entries.Select(entry => entry.Metadata.FormulaId)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count() != convertedFormulaCount)
            throw new InvalidDataException(
                $"Converted VisualTeX target inventory mismatch before numbering: expected {convertedFormulaCount} unique formulas, found {entries.Count} hosts / {entries.Select(entry => entry.Metadata.FormulaId).Distinct(StringComparer.OrdinalIgnoreCase).Count()} identities.");

        var numberedEntries = entries
            .Where(entry => entry.Metadata.Numbered
                && string.Equals(
                    entry.Metadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal))
            .ToArray();
        IReadOnlyDictionary<string, (int Ordinal, string Prefix)> plannedNumbers =
            new Dictionary<string, (int Ordinal, string Prefix)>(StringComparer.OrdinalIgnoreCase);
        var canFinalizeFreshPlannedBatch = false;
        if (usePlannedFieldFastPath && numberedEntries.Length > 0)
        {
            plannedNumbers = WordEquationNumbering.PlanNewEquationNumberBatch(
                document,
                numberedEntries
                    .Select(entry => (entry.Metadata.FormulaId, entry.Position))
                    .ToArray(),
                out var hasExistingManagedCaptions);
            canFinalizeFreshPlannedBatch = !hasExistingManagedCaptions
                && plannedNumbers.Count == numberedEntries.Length;
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-visualtex-numbering-plan formulas={numberedEntries.Length} planned={plannedNumbers.Count} existingCaptions={hasExistingManagedCaptions} freshFastPath={canFinalizeFreshPlannedBatch}");
        }

        var builtNumbered = 0;
        // Numbered display creation is a structural Word edit. Never hold or walk a
        // live InlineShapes collection while finalizing tab paragraphs or migrating
        // legacy numbered hosts: Word can invalidate/reorder that COM collection
        // after the first structural replacement. Freeze FormulaId + position first,
        // then resolve each durable VTO_ identity afresh and work from the end of the
        // document toward the start. Unlike the OMML table batch, each OLE hidden
        // caption inserts a real paragraph after its formula; start-to-end creation
        // shifts every later OLE and measured slower in real Word.
        foreach (var entry in entries.OrderByDescending(item => item.Position))
        {
            InlineShape? shape = null;
            Range? range = null;
            try
            {
                shape = ResolveConvertedVisualTeXByEmbeddedIdentity(
                        document,
                        entry.Metadata.FormulaId,
                        entry.RangeHint)
                    ?? throw new InvalidOperationException(
                        $"Converted VisualTeX formula {entry.Metadata.FormulaId} could not be resolved from its frozen range before numbering finalization.");
                range = shape.Range;
                if (normalizeMathTypeStyle)
                    NormalizeConvertedVisualTeXParagraphFromMathTypeStyle(range);
                Paragraphs? liveParagraphs = null;
                Paragraph? liveParagraph = null;
                Range? liveParagraphRange = null;
                try
                {
                    liveParagraphs = range.Paragraphs;
                    if (liveParagraphs.Count > 0)
                    {
                        liveParagraph = liveParagraphs[1];
                        liveParagraphRange = liveParagraph.Range;
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-visualtex-numbering-before-scaffold formulaId={entry.Metadata.FormulaId} frozenRange={entry.RangeHint} liveRange={range.Start}:{range.End} paragraph={liveParagraphRange.Start}:{liveParagraphRange.End}");
                    }
                }
                finally
                {
                    Release(liveParagraphRange);
                    Release(liveParagraph);
                    Release(liveParagraphs);
                }
                if (string.Equals(
                        entry.Metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal))
                {
                    RemoveInlineBaselineSentinel(document, entry.Metadata.FormulaId);
                    ResetDisplayFormulaPosition(range);
                }

                if (entry.Metadata.Numbered
                    && string.Equals(
                        entry.Metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal))
                {
                    if (canFinalizeFreshPlannedBatch
                        && plannedNumbers.TryGetValue(
                            entry.Metadata.FormulaId,
                            out var plannedNumber))
                    {
                        WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                            document,
                            range,
                            shape.Height,
                            entry.Metadata,
                            plannedOrdinal: plannedNumber.Ordinal,
                            plannedPrefix: plannedNumber.Prefix,
                            deferFieldUpdate: true);
                    }
                    else
                    {
                        WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                            document,
                            range,
                            shape.Height,
                            entry.Metadata);
                    }
                    builtNumbered++;
                }
                else
                {
                    // DeferNumberingLayout also skips the normal display-paragraph
                    // reconciliation for unnumbered targets. Reapply that local
                    // formatting here without touching unrelated formulas.
                    WordEquationNumbering.ReconcileFormula(
                        document,
                        range,
                        shape.Height,
                        entry.Metadata,
                        numberingOrderMayHaveChanged: false);
                }
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }

        // Structural migration is now complete. Rebind every converted OLE identity
        // bookmark from the actual embedded FormulaId so later edits do not inherit
        // a VTO_ bookmark that Word expanded while an earlier OMML table was removed.
        InlineShapes? finalShapes = null;
        try
        {
            var convertedIds = new HashSet<string>(
                entries.Select(entry => entry.Metadata.FormulaId),
                StringComparer.OrdinalIgnoreCase);
            finalShapes = document.InlineShapes;
            var finalCount = finalShapes.Count;
            var rebound = 0;
            for (var index = 1; index <= finalCount; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = finalShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                    var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(shape);
                    if (metadata is null || !convertedIds.Contains(metadata.FormulaId))
                        continue;
                    BindOleIdentityBookmark(shape, metadata.FormulaId);
                    rebound++;
                }
                finally { Release(shape); }
            }
            if (rebound != entries.Count)
                throw new InvalidDataException(
                    $"Converted VisualTeX identity repair mismatch: expected {entries.Count}, rebound {rebound}.");
        }
        finally { Release(finalShapes); }

        if (canFinalizeFreshPlannedBatch)
        {
            var verifyWatch = System.Diagnostics.Stopwatch.StartNew();
            if (WordEquationNumbering.VerifyFreshPlannedEquationNumberBatch(
                    document,
                    plannedNumbers))
            {
                numberingFinalizedInBatch = true;
                verifyWatch.Stop();
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-visualtex-numbering-plan-verified formulas={plannedNumbers.Count} elapsedMs={verifyWatch.ElapsedMilliseconds}");
            }
            else
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-visualtex-numbering-plan-fallback formulas={plannedNumbers.Count}");
                WordEquationNumbering.UpdateEquationNumbers(
                    document,
                    plannedNumbers.Keys.ToArray());
                numberingFinalizedInBatch = true;
                verifyWatch.Stop();
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-visualtex-numbering-plan-fallback-complete formulas={plannedNumbers.Count} elapsedMs={verifyWatch.ElapsedMilliseconds}");
            }
        }

        return builtNumbered;
    }

    private static void NormalizeConvertedVisualTeXParagraphFromMathTypeStyle(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Style? style = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            object styleObject = paragraphRange.get_Style();
            style = styleObject as Style;
            var styleName = style?.NameLocal
                ?? Convert.ToString(styleObject)
                ?? string.Empty;
            if (styleName.IndexOf(
                    "MTDisplayEquation",
                    StringComparison.OrdinalIgnoreCase) < 0)
                return;

            // MathType's display style owns its center/right tabs at the style
            // level. When a MathType paragraph is reused as the exact insertion
            // boundary for a VisualTeX OLE target, Word may optimize away the
            // direct TabStops that VisualTeX subsequently writes because the same
            // effective tabs are inherited from MTDisplayEquation. That leaves a
            // VisualTeX formula semantically dependent on MathType's style and the
            // strict OpenXML health check sees TAB characters with no direct
            // center/right definitions. Detach only this known MathType style;
            // ConfigureEquationParagraph immediately reapplies VisualTeX's own
            // direct paragraph geometry afterwards.
            object normalStyle = WdBuiltinStyle.wdStyleNormal;
            paragraphRange.set_Style(ref normalStyle);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-visualtex-detached-mathtype-style range={paragraphRange.Start}:{paragraphRange.End} style={styleName}");
        }
        finally
        {
            Release(style);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void NormalizeConvertedOmmlParagraphFromMathTypeStyle(
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Style? style = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        try
        {
            // Resolve the paragraph only after Word has materialized the native
            // OMath. A collapsed insertion range at a paragraph boundary can point
            // to the wrong side, while formulaRange.Paragraphs[1] is the definitive
            // owner that will later be converted into the 1x3 table.
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            object styleObject = paragraphRange.get_Style();
            style = styleObject as Style;
            var styleName = style?.NameLocal
                ?? Convert.ToString(styleObject)
                ?? string.Empty;
            if (styleName.IndexOf(
                    "MTDisplayEquation",
                    StringComparison.OrdinalIgnoreCase) < 0)
                return;

            // The source MathType paragraph has already been proven to contain no
            // user prose or second object. Detach its generated paragraph style now,
            // before numbered OMML converts this paragraph into a table. Otherwise
            // Word retains an MTDisplayEquation live layout origin and places the
            // table at page X=0 even though the serialized tblPr has no indent.
            object normalStyle = WdBuiltinStyle.wdStyleNormal;
            paragraphRange.set_Style(ref normalStyle);
            format = paragraphRange.ParagraphFormat;
            try { format.Reset(); } catch { }
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            tabs = format.TabStops;
            tabs.ClearAll();
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-omml-detached-mathtype-style range={paragraphRange.Start}:{paragraphRange.End} style={styleName}");
        }
        finally
        {
            Release(tabs);
            Release(format);
            Release(style);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ConvertedOmmlDirectTablesAreAlreadyComplete(
        Document document,
        IReadOnlyCollection<string> formulaIds,
        IReadOnlyDictionary<string, FormulaMetadata>? freshMetadata = null)
    {
        foreach (var formulaId in formulaIds)
        {
            // The producer stamps the final native #SEQ content into this
            // transaction's verified metadata before persistence. The stored copy
            // still fingerprints bare OMath until the batch commits; never use it
            // to reject a correctly materialized new number wrapper.
            var metadata = freshMetadata is not null && freshMetadata.TryGetValue(formulaId, out var fresh)
                ? fresh
                : WordOmmlFormulaStore.TryRead(document, formulaId)
                    ?? throw new InvalidDataException($"Converted OMML metadata {formulaId} is missing.");
            Range? formulaRange = null;
            try
            {
                formulaRange = WordOmmlFormulaStore
                    .GetEquationRangeVerifiedForStructuralEdit(
                        document,
                        metadata.FormulaId,
                        metadata);
                var directTableHost = WordEquationNumbering
                    .HasReusableNumberedNativeOmmlDirectTableHost(
                        document,
                        formulaRange,
                        metadata.FormulaId);
                var tableCellHashHost = WordEquationNumbering
                    .HasReusableNumberedNativeOmmlHashSequenceHost(
                        document,
                        formulaRange,
                        metadata.FormulaId);
                if (!directTableHost && !tableCellHashHost)
                {
                    WordDoubleClickHook.TraceMessage($"format-conversion-number-owner-incomplete formulaId={formulaId} fresh={freshMetadata is not null} range={formulaRange.Start}:{formulaRange.End}");
                    return false;
                }
            }
            catch (Exception error)
            {
                WordDoubleClickHook.TraceMessage($"format-conversion-number-owner-unresolved formulaId={formulaId} fresh={freshMetadata is not null} error={error.Message}");
                return false;
            }
            finally { Release(formulaRange); }
        }
        return true;
    }

    private static InlineShape? ResolveConvertedVisualTeXByEmbeddedIdentity(
        Document document,
        string formulaId,
        string rangeHint)
    {
        Range? hintedRange = null;
        InlineShapes? hintedShapes = null;
        Range? content = null;
        InlineShapes? allShapes = null;
        try
        {
            if (TryParseRangeReference(rangeHint, out var start, out var end))
            {
                content = document.Content;
                if (start >= content.Start && end >= start && end <= content.End)
                {
                    hintedRange = document.Range(start, end);
                    hintedShapes = hintedRange.InlineShapes;
                    for (var index = 1; index <= hintedShapes.Count; index++)
                    {
                        InlineShape? candidate = null;
                        try
                        {
                            candidate = hintedShapes[index];
                            if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                            var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(candidate);
                            if (!string.Equals(
                                    metadata?.FormulaId,
                                    formulaId,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            var result = candidate;
                            candidate = null;
                            return result;
                        }
                        finally { Release(candidate); }
                    }
                }
            }

            // End-to-start migration should keep every earlier frozen range stable.
            // Retain a raw embedded-identity fallback for unusual Word range drift,
            // but never consult VTO_ here because those are exactly what this batch
            // is repairing.
            allShapes = document.InlineShapes;
            for (var index = 1; index <= allShapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = allShapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var metadata = WordFormulaMetadataReader.TryReadEmbeddedNativeOle(candidate);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var result = candidate;
                    candidate = null;
                    return result;
                }
                finally { Release(candidate); }
            }
            return null;
        }
        finally
        {
            Release(allShapes);
            Release(content);
            Release(hintedShapes);
            Release(hintedRange);
        }
    }


    private sealed class BlockMathTypeOmmlGroup
    {
        internal int Start { get; set; }
        internal int End { get; set; }
        internal IReadOnlyList<WordFormulaFormatConversionTarget> Targets { get; set; } =
            Array.Empty<WordFormulaFormatConversionTarget>();
    }

    private IReadOnlyList<BlockMathTypeOmmlGroup> BuildContiguousBlockMathTypeOmmlGroups(
        Document document,
        IReadOnlyList<WordFormulaFormatConversionTarget> targets)
    {
        var entries = new List<(
            WordFormulaFormatConversionTarget Target,
            int Start,
            int End)>();
        foreach (var target in targets.Where(item =>
                     !item.SourceWithinTable
                     && !item.SourceHasMathTypeSectionPrefix
                     && string.Equals(
                         item.DisplayMode,
                         "block",
                         StringComparison.Ordinal)))
        {
            InlineShape? shape = null;
            Range? shapeRange = null;
            Paragraphs? paragraphs = null;
            Paragraph? paragraph = null;
            Range? paragraphRange = null;
            try
            {
                shape = FindMathTypeOleByRange(
                        document,
                        target.SourceObjectId,
                        allowGlobalFallback: false)
                    ?? throw new InvalidOperationException(
                        "A MathType block source moved before grouped OMML conversion was planned.");
                shapeRange = shape.Range;
                paragraphs = shapeRange.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidOperationException(
                        "A MathType block source no longer occupies one paragraph.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                    throw new InvalidOperationException(
                        "A MathType block source paragraph contains ordinary user text.");
                entries.Add((target, paragraphRange.Start, paragraphRange.End));
            }
            finally
            {
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(shapeRange);
                Release(shape);
            }
        }

        var groups = new List<BlockMathTypeOmmlGroup>();
        var current = new List<(
            WordFormulaFormatConversionTarget Target,
            int Start,
            int End)>();
        foreach (var entry in entries.OrderBy(item => item.Start))
        {
            if (current.Count == 0
                || current[current.Count - 1].End == entry.Start)
            {
                current.Add(entry);
                continue;
            }
            if (current.Count > 1)
            {
                groups.Add(new BlockMathTypeOmmlGroup
                {
                    Start = current[0].Start,
                    End = current[current.Count - 1].End,
                    Targets = current.Select(item => item.Target).ToArray(),
                });
            }
            current.Clear();
            current.Add(entry);
        }
        if (current.Count > 1)
        {
            groups.Add(new BlockMathTypeOmmlGroup
            {
                Start = current[0].Start,
                End = current[current.Count - 1].End,
                Targets = current.Select(item => item.Target).ToArray(),
            });
        }
        return groups;
    }

    private static IReadOnlyList<IReadOnlyList<WordFormulaFormatConversionTarget>>
        BuildAdjacentInlineOmmlGroups(
            IReadOnlyList<WordFormulaFormatConversionTarget> targets)
    {
        var sorted = targets
            .Where(target =>
                !target.SourceWithinTable
                && string.Equals(
                    target.DisplayMode,
                    "inline",
                    StringComparison.Ordinal))
            .Select(target =>
            {
                var valid = TryParseRangeReference(
                    target.SourceObjectId,
                    out var start,
                    out var end);
                return (Target: target, Valid: valid, Start: start, End: end);
            })
            .Where(item => item.Valid)
            .OrderBy(item => item.Start)
            .ToArray();
        var groups = new List<IReadOnlyList<WordFormulaFormatConversionTarget>>();
        var current = new List<(WordFormulaFormatConversionTarget Target, int Start, int End)>();
        foreach (var item in sorted)
        {
            if (current.Count == 0)
            {
                current.Add((item.Target, item.Start, item.End));
                continue;
            }
            var previous = current[current.Count - 1];
            if (previous.End == item.Start)
            {
                current.Add((item.Target, item.Start, item.End));
                continue;
            }
            if (current.Count > 1)
                groups.Add(current.Select(entry => entry.Target).ToArray());
            current.Clear();
            current.Add((item.Target, item.Start, item.End));
        }
        if (current.Count > 1)
            groups.Add(current.Select(entry => entry.Target).ToArray());
        return groups;
    }

    private int ApplyBlockMathTypeToOmmlGroup(
        Document document,
        WordFormulaFormatConversionPlan plan,
        BlockMathTypeOmmlGroup group,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        WordOmmlConverter.BatchSource ommlBatchSource)
    {
        if (group.Targets.Count < 2)
            throw new ArgumentOutOfRangeException(nameof(group));

        // The public transaction processes source hosts from end to start.
        // Resolve each member afresh and verify its actual paragraph before the
        // grouped replacement; a captured coordinate is never sufficient alone.
        var liveEntries = new List<(
            WordFormulaFormatConversionTarget Target,
            int Start,
            int End)>();
        foreach (var target in group.Targets)
        {
            InlineShape? shape = null;
            Range? shapeRange = null;
            Paragraphs? paragraphs = null;
            Paragraph? paragraph = null;
            Range? paragraphRange = null;
            try
            {
                shape = FindMathTypeOleByRange(
                        document,
                        target.SourceObjectId,
                        allowGlobalFallback: false)
                    ?? throw new InvalidOperationException(
                        "The MathType source formula moved before conversion started.");
                shapeRange = shape.Range;
                paragraphs = shapeRange.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidOperationException(
                        "A MathType block source no longer occupies one paragraph.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                    throw new InvalidOperationException(
                        "A MathType block source paragraph contains ordinary user text.");
                liveEntries.Add((target, paragraphRange.Start, paragraphRange.End));
            }
            finally
            {
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(shapeRange);
                Release(shape);
            }
        }

        var orderedEntries = liveEntries
            .OrderBy(entry => entry.Start)
            .ToArray();
        if (orderedEntries.Length != group.Targets.Count)
            throw new InvalidOperationException(
                "Not every MathType source in the consecutive group could be resolved.");
        for (var index = 1; index < orderedEntries.Length; index++)
        {
            if (orderedEntries[index - 1].End != orderedEntries[index].Start)
                throw new InvalidOperationException(
                    "The MathType display formulas are no longer a contiguous paragraph group.");
        }
        var ordered = orderedEntries
            .Select(entry => entry.Target)
            .ToArray();
        var liveGroupStart = orderedEntries[0].Start;
        var liveGroupEnd = orderedEntries[orderedEntries.Length - 1].End;
        Range? targetRange = null;
        IReadOnlyList<Range>? insertedRanges = null;
        try
        {
            foreach (var target in ordered)
                ValidateSimpleSourceHost(
                    document,
                    plan.SourceMode,
                    target);
            var targetCountBefore = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            var sourceCountBefore = CountSimpleFormatObjects(
                document,
                plan.SourceMode);
            targetRange = document.Range(liveGroupStart, liveGroupEnd);
            var formulaIds = ordered
                .Select(target => prepared[target.Id].Session.FormulaId)
                .ToArray();
            insertedRanges = ommlBatchSource.ReplaceDisplayParagraphGroup(
                _application,
                document,
                targetRange,
                formulaIds);
            document.Activate();
            if (insertedRanges.Count != ordered.Length)
                throw new InvalidOperationException(
                    $"Word retained {insertedRanges.Count}/{ordered.Length} block OMML formulas after atomic MathType replacement.");

            for (var index = 0; index < ordered.Length; index++)
            {
                var target = ordered[index];
                var formula = prepared[target.Id];
                var session = formula.Session;
                session.Mode = "create";
                session.SourceDocumentId = plan.DocumentId;
                session.SourceObjectId =
                    $"{RangeReferencePrefix}{insertedRanges[index].Start}:{insertedRanges[index].End}";
                session.DisplayMode = "block";
                session.ObjectMode = FormulaOleContract.WordOmmlMode;
                session.Numbered = target.Numbered;
                session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                session.FontSizePt = target.FontSizePt;
                session.OriginalMetadata = null;

                var metadata = session.ToMetadata();
                metadata.NativeOmmlFingerprint =
                    ommlBatchSource.GetSourceFingerprint(session.FormulaId);
                ApplyDocumentOmmlMathFont(document, metadata);
                ApplyOmmlTypography(
                    insertedRanges[index],
                    session.FontSizePt,
                    metadata);
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        insertedRanges[index],
                        metadata);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(
                            bookmark,
                            insertedRanges[index]))
                    {
                        Release(bookmark);
                        bookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            insertedRanges[index],
                            metadata,
                            replaceExisting: true);
                    }
                    WordOmmlFormulaStore.SaveNew(document, metadata);
                }
                finally { Release(bookmark); }
            }

            var targetCountAfter = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            if (targetCountAfter != targetCountBefore + ordered.Length)
                throw new InvalidOperationException(
                    $"Word retained {targetCountAfter - targetCountBefore}/{ordered.Length} block OMML targets after atomic replacement.");
            var sourceCountAfter = CountSimpleFormatObjects(
                document,
                plan.SourceMode);
            var expectedSourceCountAfter = Math.Max(0, sourceCountBefore - ordered.Length);
            if (sourceCountAfter != expectedSourceCountAfter)
                throw new InvalidOperationException(
                    $"Word retained an unexpected number of MathType sources after grouped display replacement: expected {expectedSourceCountAfter}, actual {sourceCountAfter}.");

            // Do not verify grouped MathType removal through each target's old
            // Range. Replacing several complete paragraphs changes their aggregate
            // character length, so a later, unrelated MathType OLE can legitimately
            // slide into one member's captured range. The group replacement itself
            // already proves exactly N Display OMaths were materialized; the global
            // source/target deltas above prove exactly N MathType hosts disappeared.
            // Keep these checks inside the custom UndoRecord so any genuine count or
            // materialization failure is still one atomic rollback operation.
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-block-omml-group-complete count={ordered.Length} range={liveGroupStart}:{liveGroupEnd} sourceCount={sourceCountBefore}->{sourceCountAfter} targetCount={targetCountBefore}->{targetCountAfter}");
            return ordered.Length;
        }
        finally
        {
            if (insertedRanges is not null)
                foreach (var range in insertedRanges) Release(range);
            Release(targetRange);
        }
    }

    private int ApplyAdjacentMathTypeToOmmlGroup(
        Document document,
        WordFormulaFormatConversionPlan plan,
        IReadOnlyList<WordFormulaFormatConversionTarget> group,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        WordOmmlConverter.BatchSource ommlBatchSource)
    {
        if (group.Count < 2)
            throw new ArgumentOutOfRangeException(nameof(group));
        var ordered = group.OrderBy(target => target.SourceStart).ToArray();
        if (ordered.Any(target => !string.Equals(
                target.DisplayMode,
                "inline",
                StringComparison.Ordinal)))
            throw new InvalidOperationException(
                "Only adjacent inline MathType formulas can use grouped OMML conversion.");

        Range? insertion = null;
        IReadOnlyList<Range>? insertedRanges = null;
        try
        {
            var targetCountBefore = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            var sourceCountBefore = CountSimpleFormatObjects(
                document,
                plan.SourceMode);
            var insertionStart = ordered[0].SourceStart;
            foreach (var source in ordered.OrderByDescending(target => target.SourceStart))
            {
                _ = DeleteSimpleSourceHost(
                    document,
                    plan.SourceMode,
                    source,
                    knownReferenceCounts: null);
                EnsureSimpleFormatSourceRemoved(
                    document,
                    plan.SourceMode,
                    source,
                    "adjacent-group-delete-source");
                ThrowIfSimpleFormatConversionFailureInjected(source);
            }

            Range? content = null;
            try
            {
                content = document.Content;
                insertionStart = Math.Max(
                    content.Start,
                    Math.Min(insertionStart, content.End));
            }
            finally { Release(content); }
            insertion = document.Range(insertionStart, insertionStart);

            var formulaIds = ordered
                .Select(target => prepared[target.Id].Session.FormulaId)
                .ToArray();
            insertedRanges = ommlBatchSource.InsertAdjacentInlineGroup(
                _application,
                document,
                insertion,
                formulaIds);
            document.Activate();
            if (insertedRanges.Count != ordered.Length)
                throw new InvalidOperationException(
                    $"Word retained {insertedRanges.Count}/{ordered.Length} adjacent OMML formulas.");

            for (var index = 0; index < ordered.Length; index++)
            {
                var target = ordered[index];
                var formula = prepared[target.Id];
                var session = formula.Session;
                session.Mode = "create";
                session.SourceDocumentId = plan.DocumentId;
                session.SourceObjectId =
                    $"{RangeReferencePrefix}{insertedRanges[index].Start}:{insertedRanges[index].End}";
                session.DisplayMode = "inline";
                session.ObjectMode = FormulaOleContract.WordOmmlMode;
                session.Numbered = false;
                session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                session.FontSizePt = target.FontSizePt;
                session.OriginalMetadata = null;

                var metadata = session.ToMetadata();
                metadata.NativeOmmlFingerprint =
                    ommlBatchSource.GetSourceFingerprint(session.FormulaId);
                ApplyOmmlTypography(
                    insertedRanges[index],
                    session.FontSizePt,
                    metadata);
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.Wrap(
                        document,
                        insertedRanges[index],
                        metadata);
                    if (!WordOmmlFormulaStore.IsCanonicalAnchor(
                            bookmark,
                            insertedRanges[index]))
                    {
                        Release(bookmark);
                        bookmark = WordOmmlFormulaStore.Wrap(
                            document,
                            insertedRanges[index],
                            metadata,
                            replaceExisting: true);
                    }
                    WordOmmlFormulaStore.SaveNew(document, metadata);
                }
                finally { Release(bookmark); }
            }

            var targetCountAfter = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            if (targetCountAfter != targetCountBefore + ordered.Length)
                throw new InvalidOperationException(
                    $"Word retained {targetCountAfter - targetCountBefore}/{ordered.Length} OMML targets after adjacent-group conversion.");
            var sourceCountAfter = CountSimpleFormatObjects(
                document,
                plan.SourceMode);
            var expectedSourceCountAfter = Math.Max(0, sourceCountBefore - ordered.Length);
            if (sourceCountAfter != expectedSourceCountAfter)
                throw new InvalidOperationException(
                    $"Word retained an unexpected number of MathType sources after adjacent inline replacement: expected {expectedSourceCountAfter}, actual {sourceCountAfter}.");

            // After rebuilding an adjacent inline run, the document-length delta can
            // move the next unrelated MathType object into an old member's captured
            // range. Do not use those stale ranges as residual-source identity.
            // Exact inserted OMath cardinality plus source/target count deltas are
            // the stable proof that this grouped transaction replaced N hosts.
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-adjacent-omml-group-complete count={ordered.Length} start={ordered[0].SourceStart} end={ordered[ordered.Length - 1].SourceStart} sourceCount={sourceCountBefore}->{sourceCountAfter} targetCount={targetCountBefore}->{targetCountAfter}");
            return ordered.Length;
        }
        finally
        {
            if (insertedRanges is not null)
                foreach (var range in insertedRanges) Release(range);
            Release(insertion);
        }
    }

    public WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlan(
        bool wholeDocument,
        string sourceMode,
        string targetMode)
    {
        ValidateSimpleFormatConversionPair(sourceMode, targetMode);
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        InlineShapes? shapes = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                var repairedCaptionFrames =
                    WordEquationNumbering.RepairLeakedNativeCaptionFrames(document);
                if (repairedCaptionFrames > 0)
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-preflight-repaired-caption-frames count={repairedCaptionFrames}");
            }
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;

            var plan = new WordFormulaFormatConversionPlan
            {
                DocumentId = DocumentIdentity(document),
                WritableValidated = true,
                SourceMode = sourceMode,
                TargetMode = targetMode,
                WholeDocument = wholeDocument,
                NumberFormatId = WordEquationNumbering.GetEquationNumberFormatId(document),
            };

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                CaptureOmmlFormulaFormatConversionTargets(
                    document,
                    scope,
                    wholeDocument,
                    plan);
                return plan;
            }

            var cacheWholeDocumentOleCounts = wholeDocument
                && !string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal);
            if (cacheWholeDocumentOleCounts)
            {
                plan.InitialSourceObjectCount = 0;
                if (!string.Equals(
                        targetMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                    plan.InitialTargetObjectCount = 0;
            }

            IReadOnlyList<MathTypeSectionColumnLayout>? mathTypeTargetColumnLayouts = null;
            if (string.Equals(
                    targetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                // Freeze display-column geometry while the source document is still
                // intact. Re-reading PageSetup/TextColumns after every OLE replacement
                // makes large VisualTeX -> MathType conversions depend on Word's
                // partially repaginated intermediate state and can return invalid tab
                // coordinates late in a 1000-formula batch.
                mathTypeTargetColumnLayouts = CaptureMathTypeSectionColumnLayouts(document);
            }

            shapes = document.InlineShapes;
            var tracePlanPerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
                "1",
                StringComparison.Ordinal);
            IReadOnlyList<MathTypeWordOpenXml.OleSnapshot>? bulkOleSnapshots = null;
            var bulkOleSnapshotIndex = 0;
            if (wholeDocument && string.Equals(
                    sourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                Range? packageRange = null;
                var bulkWatch = tracePlanPerformance
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                try
                {
                    packageRange = document.Content;
                    bulkOleSnapshots = MathTypeWordOpenXml.ReadOleSnapshots(packageRange);
                    if (bulkWatch is not null)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-plan-bulk-ole-snapshot count={bulkOleSnapshots.Count} elapsedMs={bulkWatch.ElapsedMilliseconds}");
                    }
                }
                catch (Exception bulkError)
                {
                    bulkOleSnapshots = null;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-plan-bulk-ole-snapshot-fallback error={bulkError.GetType().Name}:{bulkError.Message}");
                }
                finally { Release(packageRange); }
            }
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                var planPerfWatch = tracePlanPerformance
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                long planPerfLastMs = 0;
                void TracePlanPerf(string perfStage)
                {
                    if (planPerfWatch is null) return;
                    var totalMs = planPerfWatch.ElapsedMilliseconds;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-plan-perf stage={perfStage} source={sourceMode} index={index}/{shapes.Count} deltaMs={totalMs - planPerfLastMs} totalMs={totalMs}");
                    planPerfLastMs = totalMs;
                }
                try
                {
                    shape = shapes[index];
                    TracePlanPerf("shape-access");
                    MathTypeWordOpenXml.OleSnapshot? bulkOleSnapshot = null;
                    if (bulkOleSnapshots is not null
                        && shape.Type is WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                            or WdInlineShapeType.wdInlineShapeLinkedOLEObject)
                    {
                        if (bulkOleSnapshotIndex < bulkOleSnapshots.Count)
                            bulkOleSnapshot = bulkOleSnapshots[bulkOleSnapshotIndex];
                        bulkOleSnapshotIndex++;
                    }
                    // Scope is a physical range test. Perform it before reading
                    // embedded metadata/MathML so selecting one formula never
                    // activates or decodes unrelated equations in the document.
                    range = shape.Range;
                    TracePlanPerf("shape-range");
                    if (!FormulaRangeMatchesScope(range, scope, wholeDocument))
                        continue;
                    FormulaMetadata? metadata = null;
                    string? sourceMathMl = null;
                    var mathTypeNumberPosition = "right";
                    string sourceFormulaId;

                    if (string.Equals(
                            sourceMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal))
                    {
                        var isSourceNativeOle = WordFormulaMetadataReader.IsNativeOle(shape);
                        if (cacheWholeDocumentOleCounts)
                        {
                            if (isSourceNativeOle)
                                plan.InitialSourceObjectCount++;
                            else if (plan.InitialTargetObjectCount.HasValue
                                && string.Equals(
                                    targetMode,
                                    FormulaOleContract.MathTypeOleMode,
                                    StringComparison.Ordinal)
                                && MathTypeOleInterop.IsMathTypeOle(shape))
                                plan.InitialTargetObjectCount++;
                        }
                        if (!isSourceNativeOle) continue;
                        metadata = WordFormulaMetadataReader.TryRead(shape);
                        if (metadata is null) continue;
                        sourceFormulaId = metadata.FormulaId;
                    }
                    else
                    {
                        var isSourceMathTypeOle = MathTypeOleInterop.IsMathTypeOle(shape);
                        if (cacheWholeDocumentOleCounts)
                        {
                            if (isSourceMathTypeOle)
                                plan.InitialSourceObjectCount++;
                            else if (plan.InitialTargetObjectCount.HasValue
                                && string.Equals(
                                    targetMode,
                                    FormulaOleContract.NativeOleMode,
                                    StringComparison.Ordinal)
                                && WordFormulaMetadataReader.IsNativeOle(shape))
                                plan.InitialTargetObjectCount++;
                        }
                        if (!isSourceMathTypeOle) continue;
                        TracePlanPerf("identify-mathtype");
                        byte[]? knownNativeCompoundFile = null;
                        if (bulkOleSnapshot is not null
                            && bulkOleSnapshot.CompoundFile.Length > 0
                            && MathTypeOleStorage.LooksLikeMathTypeCompoundFile(
                                bulkOleSnapshot.CompoundFile))
                        {
                            try
                            {
                                sourceMathMl = MathTypeOleStorage.ReadMathMl(
                                    bulkOleSnapshot.CompoundFile);
                                knownNativeCompoundFile = bulkOleSnapshot.CompoundFile;
                                TracePlanPerf("read-mathml-bulk");
                            }
                            catch
                            {
                                sourceMathMl = MathTypeOleStorage.ReadMathMl(shape);
                                TracePlanPerf("read-mathml-fallback");
                            }
                        }
                        else
                        {
                            sourceMathMl = MathTypeOleStorage.ReadMathMl(shape);
                            TracePlanPerf("read-mathml-fallback");
                        }
                        metadata = MathTypeOleInterop.ReadMetadata(
                            _application,
                            shape,
                            sourceMathMl,
                            knownNativeCompoundFile);
                        TracePlanPerf("read-metadata");
                        if (MathTypeOleInterop.TryReadDisplayNumberPosition(
                                shape,
                                out var detectedPosition))
                            mathTypeNumberPosition = detectedPosition;
                        TracePlanPerf("number-position");
                        sourceFormulaId = metadata.FormulaId;
                    }

                    var latex = string.IsNullOrWhiteSpace(metadata.Latex)
                        ? string.Join("\n", metadata.Lines.Select(line => line.Latex))
                        : metadata.Latex;
                    latex = (latex ?? string.Empty).Trim();
                    if (latex.Length == 0)
                        throw new InvalidDataException(
                            "A source formula has no recoverable LaTeX and was not converted.");

                    if (string.Equals(
                            sourceMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal)
                        && metadata.Numbered
                        && string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                    {
                        // VisualTeX numbered formulas do not carry MathType's side
                        // preference in their own metadata.  Conversion must follow
                        // the current document's native MathType preference instead
                        // of silently forcing every target to the right.
                        mathTypeNumberPosition = ReadMathTypeNumberPositionPreference(document);
                    }

                    float? mathTypeDisplayColumnWidth = null;
                    if (mathTypeTargetColumnLayouts is not null
                        && string.Equals(
                            metadata.DisplayMode,
                            "block",
                            StringComparison.Ordinal))
                    {
                        mathTypeDisplayColumnWidth = ResolveMathTypeDisplayColumnWidthFromLayouts(
                            document,
                            range,
                            mathTypeTargetColumnLayouts);
                    }

                    var conversionTarget = new WordFormulaFormatConversionTarget
                    {
                        Id = Guid.NewGuid().ToString("D"),
                        SourceFormulaId = sourceFormulaId,
                        SourceObjectId = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                        SourceStart = range.Start,
                        Latex = latex,
                        SourceMathMl = sourceMathMl,
                        SourceIsManagedOmml = false,
                        SourceWithinTable = WordEquationNumbering.RangeIsWhollyWithinTable(range),
                        DisplayMode = metadata.DisplayMode,
                        Numbered = metadata.Numbered,
                        // Preserve user-authored paragraph spacing for
                        // both OLE source families. Restricting this capture to
                        // VisualTeX made a MathType→OMML return leg treat the real
                        // blank before the first display formula as generated
                        // scaffolding and delete it during finalization.
                        PrecedingPlainBlankParagraphCount =
                            (string.Equals(
                                 sourceMode,
                                 FormulaOleContract.NativeOleMode,
                                 StringComparison.Ordinal)
                             || string.Equals(
                                 sourceMode,
                                 FormulaOleContract.MathTypeOleMode,
                                 StringComparison.Ordinal))
                            ? CountPlainBlankParagraphsImmediatelyBeforeFormulaHost(
                                document,
                                range,
                                metadata,
                                sourceMode)
                            : 0,
                        MathTypeNumberPosition = mathTypeNumberPosition,
                        FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                        MathTypeDisplayColumnWidth = mathTypeDisplayColumnWidth,
                        Metadata = metadata,
                    };
                    CaptureTableInlineFollowingBoundary(document, range, metadata, conversionTarget);
                    plan.Targets.Add(conversionTarget);
                    TracePlanPerf("target-added");
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            if (cacheWholeDocumentOleCounts)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-plan-object-counts source={plan.InitialSourceObjectCount?.ToString() ?? "<deferred>"} target={plan.InitialTargetObjectCount?.ToString() ?? "<deferred>"}");
            }
            return plan;
        }
        finally
        {
            Release(shapes);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    private static int CountPlainBlankParagraphsImmediatelyBeforeFormulaHost(
        Document document,
        Range formulaRange,
        FormulaMetadata metadata,
        string sourceMode)
    {
        Range? numberingOwnerRange = null;
        Paragraphs? ownerParagraphs = null;
        Paragraph? ownerParagraph = null;
        Range? ownerParagraphRange = null;
        try
        {
            var anchorStart = formulaRange.Start;
            if (metadata.Numbered
                && string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            {
                if (string.Equals(
                        sourceMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    numberingOwnerRange =
                        WordEquationNumbering.FindNumberingOwnerRange(
                            document,
                            metadata.FormulaId);
                    if (numberingOwnerRange is not null)
                        anchorStart = numberingOwnerRange.Start;
                }
                else if (string.Equals(
                             sourceMode,
                             FormulaOleContract.MathTypeOleMode,
                             StringComparison.Ordinal))
                {
                    // A MathType number is a nested MTPlaceRef field later in the
                    // same paragraph. FindNumberingOwnerRange therefore starts
                    // inside the field, and counting from there always reports
                    // zero. The display formula's true structural owner is its
                    // paragraph; count from that paragraph boundary so a real
                    // source blank before the first formula survives a round trip.
                    ownerParagraphs = formulaRange.Paragraphs;
                    if (ownerParagraphs.Count != 1)
                        throw new InvalidDataException(
                            "A numbered MathType source does not occupy exactly one paragraph.");
                    ownerParagraph = ownerParagraphs[1];
                    ownerParagraphRange = ownerParagraph.Range.Duplicate;
                    anchorStart = ownerParagraphRange.Start;
                }
            }
            return CountConsecutivePlainBlankParagraphsBeforePosition(
                document,
                anchorStart);
        }
        finally
        {
            Release(ownerParagraphRange);
            Release(ownerParagraph);
            Release(ownerParagraphs);
            Release(numberingOwnerRange);
        }
    }

    private static int CountConsecutivePlainBlankParagraphsBeforePosition(
        Document document,
        int position)
    {
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Tables? tables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        try
        {
            content = document.Content;
            var cursor = position;
            var count = 0;
            while (cursor > content.Start)
            {
                Release(frames); frames = null;
                Release(bookmarks); bookmarks = null;
                Release(fields); fields = null;
                Release(maths); maths = null;
                Release(shapes); shapes = null;
                Release(tables); tables = null;
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = null;
                Release(paragraphs); paragraphs = null;
                Release(probe); probe = null;

                probe = document.Range(cursor - 1, cursor);
                if (WordEquationNumbering.RangeIsWhollyWithinTable(probe)) break;
                paragraphs = probe.Paragraphs;
                if (paragraphs.Count != 1) break;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                if (paragraphRange.End != cursor
                    || !string.Equals(paragraphRange.Text, "\r", StringComparison.Ordinal))
                    break;

                tables = paragraphRange.Tables;
                shapes = paragraphRange.InlineShapes;
                maths = paragraphRange.OMaths;
                fields = paragraphRange.Fields;
                bookmarks = paragraphRange.Bookmarks;
                frames = paragraphRange.Frames;
                if (tables.Count != 0
                    || shapes.Count != 0
                    || maths.Count != 0
                    || fields.Count != 0
                    || bookmarks.Count != 0
                    || frames.Count != 0)
                    break;

                count++;
                cursor = paragraphRange.Start;
            }
            return count;
        }
        finally
        {
            Release(frames);
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(tables);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
        }
    }

    private static int RestoreConvertedOmmlPrecedingBlankParagraphCounts(
        Document document,
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        IReadOnlyCollection<string> convertedTargetIds)
    {
        var convertedTargetIdSet = new HashSet<string>(
            convertedTargetIds,
            StringComparer.Ordinal);
        var removed = 0;
        foreach (var target in plan.Targets
                     .Where(item => convertedTargetIdSet.Contains(item.Id)
                         && item.Numbered
                         && string.Equals(item.DisplayMode, "block", StringComparison.Ordinal)))
        {
            if (!prepared.TryGetValue(target.Id, out var converted)
                || string.IsNullOrWhiteSpace(converted.Session.FormulaId))
                continue;
            var formulaId = converted.Session.FormulaId;
            while (true)
            {
                Table? table = null;
                Range? tableRange = null;
                try
                {
                    tableRange = WordEquationNumbering.FindNumberingOwnerRange(
                        document,
                        formulaId)
                        ?? throw new InvalidOperationException(
                            $"Converted OMML formula {formulaId} lost its numbering owner during blank-paragraph restoration.");
                    var actual = CountConsecutivePlainBlankParagraphsBeforePosition(
                        document,
                        tableRange.Start);
                    // Two independent Word tables can never be truly adjacent:
                    // deleting the final intervening paragraph makes Word merge
                    // their rows into one 2x3 table. MathType source paragraphs can
                    // legitimately report zero preceding blanks, but once both
                    // formulas have become 1x3 hosts one structural separator must
                    // remain. Delete only surplus source paragraphs above that floor.
                    var requiresTableSeparator =
                        PlainBlankRunImmediatelyFollowsTable(
                            document,
                            tableRange.Start,
                            actual);
                    var expected = Math.Max(
                        target.PrecedingPlainBlankParagraphCount,
                        requiresTableSeparator ? 1 : 0);
                    if (actual <= expected)
                        break;
                    if (!DeletePlainBlankParagraphImmediatelyBeforePosition(
                            document,
                            tableRange.Start))
                        throw new InvalidOperationException(
                            $"Converted OMML formula {formulaId} retained an unexpected blank source paragraph that could not be removed safely.");
                    removed++;
                }
                finally
                {
                    Release(tableRange);
                    Release(table);
                }
            }

            // Independent 1x3 Word tables require one intervening body paragraph;
            // deleting that final paragraph merges the tables into a forbidden 2x3.
            // MathType source paragraphs use ordinary 10.5/12pt metrics, so after
            // surplus blanks have been removed, collapse only the mandatory table-
            // to-table separator to the same 1pt structural line used by ordinary
            // VisualTeX OMML insertion. User-authored extra blank paragraphs remain
            // untouched because this helper requires a managed direct-SEQ table on
            // both sides of the paragraph.
            WordEquationNumbering.CompactManagedNativeOmmlTableSeparatorBefore(
                document,
                formulaId);
        }
        return removed;
    }

    private static bool PlainBlankRunImmediatelyFollowsTable(
        Document document,
        int position,
        int plainBlankCount)
    {
        if (plainBlankCount <= 0) return false;
        Range? content = null;
        Range? probe = null;
        try
        {
            content = document.Content;
            var runStart = position - plainBlankCount;
            if (runStart <= content.Start || runStart > content.End)
                return false;
            probe = document.Range(runStart - 1, runStart);
            return WordEquationNumbering.RangeIsWhollyWithinTable(probe);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(probe);
            Release(content);
        }
    }

    private static bool DeletePlainBlankParagraphImmediatelyBeforePosition(
        Document document,
        int position)
    {
        Range? content = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Tables? tables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        try
        {
            content = document.Content;
            if (position <= content.Start || position > content.End)
                return false;
            probe = document.Range(position - 1, position);
            if (WordEquationNumbering.RangeIsWhollyWithinTable(probe))
                return false;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (paragraphRange.End != position
                || !string.Equals(paragraphRange.Text, "\r", StringComparison.Ordinal))
                return false;

            tables = paragraphRange.Tables;
            shapes = paragraphRange.InlineShapes;
            maths = paragraphRange.OMaths;
            fields = paragraphRange.Fields;
            bookmarks = paragraphRange.Bookmarks;
            frames = paragraphRange.Frames;
            if (tables.Count != 0
                || shapes.Count != 0
                || maths.Count != 0
                || fields.Count != 0
                || bookmarks.Count != 0
                || frames.Count != 0)
                return false;

            paragraphRange.Delete();
            return true;
        }
        finally
        {
            Release(frames);
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(tables);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(content);
        }
    }

    private static void CaptureOmmlFormulaFormatConversionTargets(
        Document document,
        Range scope,
        bool wholeDocument,
        WordFormulaFormatConversionPlan plan)
    {
        if (!wholeDocument
            && TryCaptureSingleOmmlFormatConversionTarget(document, scope, plan))
            return;

        List<FormulaToLatexTarget>? sources = null;
        try
        {
            sources = CaptureFormulaToLatexTargets(
                document,
                scope,
                wholeDocument,
                FormulaOleContract.WordOmmlMode,
                refreshOmmlMetadata: true);
            foreach (var source in sources)
            {
                var metadata = source.Metadata;
                var range = source.FormulaRange;
                var wordOpenXml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    document,
                    range,
                    metadata.FormulaId);
                var sourceMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                    wordOpenXml,
                    display: string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.Ordinal));
                var latex = MathMlToLatexConverter.Convert(sourceMathMl).Trim();
                if (latex.Length == 0)
                    throw new InvalidDataException(
                        "A Word OMML source formula could not be converted to editable LaTeX for preview rendering.");

                var conversionTarget = new WordFormulaFormatConversionTarget
                {
                    Id = Guid.NewGuid().ToString("D"),
                    SourceFormulaId = metadata.FormulaId,
                    SourceObjectId = $"{RangeReferencePrefix}{range.Start}:{range.End}",
                    SourceStart = range.Start,
                    Latex = latex,
                    SourceMathMl = sourceMathMl,
                    SourceIsManagedOmml = source.OmmlBookmark is not null,
                    SourceWithinTable = WordEquationNumbering.RangeIsWhollyWithinTable(range),
                    DisplayMode = metadata.DisplayMode,
                    Numbered = metadata.Numbered,
                    MathTypeNumberPosition = metadata.Numbered
                        ? ReadMathTypeNumberPositionPreference(document)
                        : "right",
                    FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                    Metadata = metadata,
                };
                CaptureTableInlineFollowingBoundary(document, range, metadata, conversionTarget);
                plan.Targets.Add(conversionTarget);
            }
        }
        finally
        {
            ReleaseFormulaToLatexTargets(sources);
        }
    }

    private static bool TryCaptureSingleOmmlFormatConversionTarget(
        Document document,
        Range scope,
        WordFormulaFormatConversionPlan plan)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? equationRange = null;
        Bookmark? bookmark = null;
        string? capturedWordOpenXml = null;
        try
        {
            maths = scope.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            equationRange = math.Range.Duplicate;
            bookmark = WordOmmlFormulaStore.FindAtRange(document, equationRange);

            FormulaMetadata metadata;
            var sourceIsManagedOmml = bookmark is not null;
            if (bookmark is not null)
            {
                var stored = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (stored is null) return false;
                // This is preview/source capture only. Resolve the exact current
                // OMath once and reuse that same WordOpenXML for metadata refresh
                // and MathML conversion. The mutation path performs the strict
                // structural/fingerprint resolver again immediately before delete.
                Release(equationRange);
                equationRange = WordOmmlFormulaStore.GetEquationRangeForCurrentRead(
                    document, bookmark, stored);
                capturedWordOpenXml = WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    document,
                    equationRange,
                    stored.FormulaId);
                metadata = WordOmmlNativeSource.RefreshForVisualTeXFromCapturedWordOpenXml(
                    stored,
                    capturedWordOpenXml);
            }
            else
            {
                metadata = WordOmmlNativeSource.CreateForNative(
                    document,
                    equationRange);
            }

            var wordOpenXml = capturedWordOpenXml
                ?? WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                    document,
                    equationRange,
                    metadata.FormulaId);
            var sourceMathMl = WordOmmlConverter.TransformOmmlToMathMl(
                wordOpenXml,
                display: string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.Ordinal));
            var latex = MathMlToLatexConverter.Convert(sourceMathMl).Trim();
            if (latex.Length == 0)
                throw new InvalidDataException(
                    "The selected Word OMML source could not be converted to editable LaTeX.");

            var conversionTarget = new WordFormulaFormatConversionTarget
            {
                Id = Guid.NewGuid().ToString("D"),
                SourceFormulaId = metadata.FormulaId,
                SourceObjectId = $"{RangeReferencePrefix}{equationRange.Start}:{equationRange.End}",
                SourceStart = equationRange.Start,
                Latex = latex,
                SourceMathMl = sourceMathMl,
                SourceIsManagedOmml = sourceIsManagedOmml,
                SourceWithinTable = WordEquationNumbering.RangeIsWhollyWithinTable(equationRange),
                DisplayMode = metadata.DisplayMode,
                Numbered = metadata.Numbered,
                MathTypeNumberPosition = metadata.Numbered
                    ? ReadMathTypeNumberPositionPreference(document)
                    : "right",
                FontSizePt = FormulaFontSize.Normalize(metadata.FontSizePt),
                Metadata = metadata,
            };
            CaptureTableInlineFollowingBoundary(document, equationRange, metadata, conversionTarget);
            plan.Targets.Add(conversionTarget);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-single-omml-local-capture managed={sourceIsManagedOmml} range={equationRange.Start}:{equationRange.End}");
            return true;
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Fall back to the mature document-level enumerator when Word does
            // not expose exactly one stable local OMath at the current selection.
            return false;
        }
        finally
        {
            Release(bookmark);
            Release(equationRange);
            Release(math);
            Release(maths);
        }
    }

    private static void DeleteConversionLocator(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            bookmark.Delete();
            if (bookmarks.Exists(name))
                throw new InvalidDataException($"The conversion locator {name} was not removed.");
        }
        finally { Release(bookmark); Release(bookmarks); }
    }

    public WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlan(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        if (prepared is null) throw new ArgumentNullException(nameof(prepared));
        ValidateSimpleFormatConversionPair(plan.SourceMode, plan.TargetMode);

        Document? document = null;
        WordDocumentEditSnapshot? rollbackSnapshot = null;
        var isolatedConversion = false;
        UndoRecord? conversionUndo = null;
        var conversionUndoEnded = false;
        FormulaMetadata? conversionMathFontMetadata = null;
        Document? mathTypeVisualTeXCleanParagraphBuffer = null;
        Selection? selection = null;
        WordOmmlConverter.BatchSource? ommlBatchSource = null;
        using var liveOmmlInsertions = new WordOmmlNativeSource.LiveInsertionOwners();
        var retainedMathTypeTargetBookmarks = new Dictionary<string, string>(StringComparer.Ordinal);
        var userCellOwners = new List<CellConversionOwner>();
        var verifiedEmptyCaptionTails = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var deferredOwnedOmmlSeparators = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var deferredSingleVisualTeXNumbering = false;
        var userCellOwnerByTargetId = new Dictionary<string, CellConversionOwner>(StringComparer.Ordinal);
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        Window? batchWindow = null;
        var previousViewType = WdViewType.wdPrintView;
        var batchViewSuspended = false;
        var previousPagination = true;
        var paginationSuspended = false;
        var prepareWatch = System.Diagnostics.Stopwatch.StartNew();
        long prepareCheckpoint = 0;
        void TracePreparation(string stage) => TraceAcceptancePerformance("ConversionPrepare", stage, prepareWatch, ref prepareCheckpoint);
        void FinishBatchPresentation()
        {
            // Word's CustomRecord belongs to the application, while closing a
            // hidden source changes its active document. Commit the owned record
            // before that close; otherwise Word can discard the still-open record
            // and expose the unchanged pre-conversion undo history after failure.
            if (conversionUndo is not null && !conversionUndoEnded)
            {
                TraceConversionUndoState("before-source-window-close-commit");
                conversionUndo.EndCustomRecord();
                conversionUndoEnded = true;
                TraceConversionUndoState("after-source-window-close-commit");
            }
            // Hidden source windows and draft layout must not remain active while
            // Word serializes or restores the real document. Draft layout rounds
            // OLE extents; use the user's original presentation for both snapshots.
            ommlBatchSource?.Dispose();
            document?.Activate();
            if (paginationSuspended)
            {
                _application.Options.Pagination = previousPagination;
                paginationSuspended = false;
            }
            if (batchViewSuspended && batchWindow is not null)
            {
                batchWindow.View.Type = previousViewType;
                batchViewSuspended = false;
            }
            // Keep screen updating suspended through numbering/reference
            // finalization as well. The source and user's view are now restored;
            // repainting every intermediate row added seconds and exposed the
            // partially converted document. The outer finally restores this flag
            // on success, cancellation and verified rollback alike.
        }
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureSourceDocument(document, plan.DocumentId);
            // The capture step already validated ReadOnly for this same gated
            // conversion. Re-reading it can be an expensive Word state query on
            // large documents; retain the defensive check for any future plan
            // that was not produced by CaptureFormulaFormatConversionPlan.
            if (!plan.WritableValidated)
                EnsureWritable(document);
            if (plan.Targets.Count > 0)
            {
                try
                {
                    previousScreenUpdating = _application.ScreenUpdating;
                    _application.ScreenUpdating = false;
                    screenUpdatingSuspended = true;
                }
                catch { }
            }
            selection = _application.Selection;
            foreach (var target in plan.Targets)
            {
                if (!prepared.TryGetValue(target.Id, out var formula))
                    throw new InvalidDataException(
                        $"Missing rendered payload for formula '{target.Latex}'.");
                ValidatePreparedFormatConversionTarget(plan.TargetMode, target, formula);
                // A single OMML deletion resolves and validates the exact source
                // fingerprint immediately before deleting it below. An earlier
                // duplicate check cannot strengthen that proof. Batches still
                // preflight every source before changing the first one; OLE
                // sources retain their existing extra validation as well.
                if (plan.Targets.Count != 1 || plan.SourceMode != FormulaOleContract.WordOmmlMode)
                    ValidateSimpleSourceHost(document, plan.SourceMode, target);
            }

            var result = new WordFormulaFormatConversionResult();
            // Source replacement and all numbering/reference finalizers share one
            // owned Word undo record. No provisional target is a committed result.
            var successfulTargetIds = new HashSet<string>(StringComparer.Ordinal);
            var targetIsMathType = string.Equals(
                plan.TargetMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal);
            var targetIsVisualTeX = string.Equals(
                plan.TargetMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal);
            var targetIsOmml = string.Equals(
                plan.TargetMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal);
            var canFinalizeSingleMathTypeNumberLocally =
                targetIsMathType
                && plan.Targets.Count == 1
                && plan.Targets[0].Numbered
                && (MathTypeEquationNumbering.CanProveNoPlaceRefFields(document)
                    || MathTypeEquationNumbering.CountPlaceRefFields(document) == 0);
            var needsAtomicMathTypeVisualTeXParagraphReplacement =
                targetIsVisualTeX
                && plan.Targets.Count > 1
                && string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                && plan.Targets.Any(target => target.Numbered
                    && string.Equals(
                        target.DisplayMode,
                        "block",
                        StringComparison.Ordinal));
            if (needsAtomicMathTypeVisualTeXParagraphReplacement)
            {
                // A completed MathType row owns one Equation.DSMT4 object and one
                // nested MTPlaceRef field tree. Word can invalidate a later row's
                // paragraph RCW after the first row is deleted/replaced in the same
                // batch, producing 0x800A1710 ("Cannot edit Range") even though a
                // freshly resolved copy of that exact paragraph is writable. Keep
                // one pristine paragraph in an isolated document and replace each
                // validated MathType owner as a complete FormattedText transaction.
                mathTypeVisualTeXCleanParagraphBuffer =
                    _application.Documents.Add(Visible: false);
                ValidateCleanFormatConversionParagraphBuffer(
                    mathTypeVisualTeXCleanParagraphBuffer);
                document.Activate();
                Release(selection);
                selection = _application.Selection;
                WordDoubleClickHook.TraceMessage(
                    "format-conversion-mathtype-visualtex-atomic-buffer-created");
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-runtime assembly={typeof(WordFormulaService).Assembly.Location} source={plan.SourceMode} target={plan.TargetMode} targetIsVisualTeX={targetIsVisualTeX} targetIsMathType={targetIsMathType}");
            TracePreparation("source-validation");
            var initialTargetObjectCount = plan.InitialTargetObjectCount
                ?? CountSimpleFormatObjects(document, plan.TargetMode);
            var initialSourceObjectCount = plan.InitialSourceObjectCount
                ?? CountSimpleFormatObjects(document, plan.SourceMode);
            var canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath =
                plan.WholeDocument
                && targetIsMathType
                && string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal)
                && plan.InitialSourceObjectCount.HasValue
                && plan.InitialTargetObjectCount.HasValue
                && initialSourceObjectCount == plan.Targets.Count
                && initialTargetObjectCount == 0;
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-object-count-baseline source={initialSourceObjectCount} target={initialTargetObjectCount} cachedSource={plan.InitialSourceObjectCount.HasValue} cachedTarget={plan.InitialTargetObjectCount.HasValue} fastNativeToMathType={canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath}");
            TracePreparation("object-counts");
            IReadOnlyDictionary<string, int>? sourceReferenceCounts = null;
            var sourceHasManagedNumbering =
                (string.Equals(plan.SourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal)
                    || string.Equals(plan.SourceMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal))
                && plan.Targets.Any(target => target.Numbered);
            if (sourceHasManagedNumbering)
            {
                // Reference health is a one-pass WordOpenXML operation. Do it before
                // per-formula alias discovery: a healthy generated number owns exactly
                // one internal REF (or zero for the current direct-SEQ single-OMML
                // host). Counts <= 1 therefore prove there is no external reference
                // whose alias must be migrated. The old order discovered aliases first,
                // causing CaptureNumberAliases to enumerate document.Fields once per
                // formula and turning a 100-number conversion into ~100 full scans.
                if (plan.Targets.Count == 1)
                {
                    sourceReferenceCounts = WordEquationNumbering.ReadSingleFormulaReferenceCounts(
                        document, plan.Targets[0].SourceFormulaId);
                }
                else if (WordEquationNumbering.TryGetHealthyEquationReferenceCounts(
                             document,
                             out var capturedReferenceCounts))
                {
                    sourceReferenceCounts = capturedReferenceCounts;
                }
                if (sourceReferenceCounts is not null)
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-reference-counts-preflight source={plan.SourceMode} formulas={sourceReferenceCounts.Count}");
                TracePreparation("reference-counts-preflight");
            }
            var referenceAliasesByTargetId =
                new Dictionary<string, IReadOnlyList<EquationReferenceBookmarkAlias>>(StringComparer.Ordinal);
            foreach (var target in plan.Targets.Where(item => item.Numbered))
            {
                IReadOnlyList<EquationReferenceBookmarkAlias> aliases;
                if (string.Equals(
                        plan.SourceMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal))
                {
                    InlineShape? sourceShape = null;
                    try
                    {
                        sourceShape = FindMathTypeOleByRange(
                            document,
                            target.SourceObjectId,
                            allowGlobalFallback: false);
                        if (sourceShape is null) continue;
                        aliases = MathTypeEquationReferences.CaptureFormatConversionAliasesFromMathType(
                            document,
                            sourceShape);
                    }
                    finally { Release(sourceShape); }
                }
                else
                {
                    if (sourceReferenceCounts is not null
                        && sourceReferenceCounts.TryGetValue(target.SourceFormulaId, out var knownReferences)
                        && knownReferences <= 1)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-reference-alias-scan-skipped formulaId={target.SourceFormulaId} references={knownReferences}");
                        continue;
                    }
                    aliases = MathTypeEquationReferences.CaptureFormatConversionAliasesFromVisualTeX(
                        document,
                        target.SourceFormulaId);
                }
                if (aliases.Count == 0) continue;
                referenceAliasesByTargetId[target.Id] = aliases;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-reference-aliases-captured sourceMode={plan.SourceMode} formulaId={target.SourceFormulaId} aliases={string.Join(",", aliases.Select(alias => alias.Name))}");
            }
            TracePreparation("number-aliases");
            var capturedReferenceFormatting =
                MathTypeEquationReferences.CaptureReferenceCharacterFormatting(
                    document,
                    referenceAliasesByTargetId.Values
                        .SelectMany(items => items)
                        .Select(alias => alias.Name)
                        .Distinct(StringComparer.OrdinalIgnoreCase));
            TracePreparation("reference-formatting");
            if (targetIsOmml && plan.Targets.Count > 0)
            {
                var ommlPreparedByFormulaId = plan.Targets
                    .Select(target => prepared[target.Id])
                    .ToDictionary(
                        formula => formula.Session.FormulaId,
                        formula => formula,
                        StringComparer.OrdinalIgnoreCase);
                var requestedMathFonts = ommlPreparedByFormulaId.Values
                    .Select(item => ResolveDocumentOmmlMathFont(
                        item.Session.ToMetadata().FormulaLetterFont))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                if (requestedMathFonts.Length != 1)
                    throw new InvalidDataException(
                        "One Word OMML conversion batch cannot request multiple document-level math fonts.");
                var documentMathMetadata = ommlPreparedByFormulaId.Values
                    .First()
                    .Session
                    .ToMetadata();
                documentMathMetadata.Validate();
                conversionMathFontMetadata = documentMathMetadata;
                var targetMathFontName = requestedMathFonts[0];

                var ommlFormulas = plan.Targets.Select(target =>
                {
                    var formula = prepared[target.Id];
                    return (
                        FormulaId: formula.Session.FormulaId,
                        MathMl: formula.MathMl
                            ?? throw new InvalidDataException(
                                $"Missing MathML for OMML conversion target '{target.Latex}'."));
                }).ToList();
                var batchSourceWatch = System.Diagnostics.Stopwatch.StartNew();
                ommlBatchSource = WordOmmlConverter.CreateBatchSource(
                    _application,
                    ommlFormulas,
                    (formulaId, omml) =>
                    {
                        if (!ommlPreparedByFormulaId.TryGetValue(
                                formulaId,
                                out var preparedFormula))
                            throw new InvalidDataException(
                                $"Missing OMML typography configuration for formula {formulaId}.");
                        var typographyMetadata = preparedFormula.Session.ToMetadata();
                        typographyMetadata.Validate();
                        return ApplyOmmlTypographyXml(
                            omml,
                            preparedFormula.Session.FontSizePt,
                            typographyMetadata);
                    },
                    mathFontName: targetMathFontName);
                document.Activate();
                Release(selection);
                selection = _application.Selection;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-omml-batch-source-created formulas={ommlFormulas.Count} elapsedMs={batchSourceWatch.ElapsedMilliseconds}");
            }
            TracePreparation("omml-source");
            IReadOnlyDictionary<int, ResolvedEquationHeadingScope>? batchHeadingScopes = null;
            HashSet<int>? preparedHeadingScopeStarts = null;
            if (targetIsMathType && plan.Targets.Any(target => target.Numbered))
            {
                var batchNumberFormatId = EquationNumberFormat.Resolve(plan.NumberFormatId).Id;
                var batchNumberFormat = EquationNumberFormat.Resolve(batchNumberFormatId);
                if (batchNumberFormat.UsesHeading)
                {
                    batchHeadingScopes = WordEquationNumbering.CaptureHeadingScopesAtPositions(
                        document,
                        batchNumberFormatId,
                        plan.Targets
                            .Where(target => target.Numbered)
                            .Select(target => target.SourceStart));
                    preparedHeadingScopeStarts = new HashSet<int>();
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-heading-scope-cache format={batchNumberFormatId} formulas={batchHeadingScopes.Count} scopes={batchHeadingScopes.Values.Select(scope => scope.ScopeStart).Distinct().Count()}");
                }
            }
            TracePreparation("heading-scopes");
            if (sourceHasManagedNumbering && sourceReferenceCounts is null)
            {
                // Preserve the conservative fallback for legacy/malformed numbering.
                // This normally remains null because the healthy one-pass preflight
                // above has already supplied the counts before alias discovery.
                if (plan.Targets.Count == 1)
                    sourceReferenceCounts = WordEquationNumbering.ReadSingleFormulaReferenceCounts(
                        document, plan.Targets[0].SourceFormulaId);
                else if (WordEquationNumbering.TryGetHealthyEquationReferenceCounts(
                             document,
                             out var capturedReferenceCounts))
                    sourceReferenceCounts = capturedReferenceCounts;
                if (sourceReferenceCounts is not null)
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-reference-counts-fallback source={plan.SourceMode} formulas={sourceReferenceCounts.Count}");
            }

            // A numbered VisualTeX/OMML formula has one generated REF inside its
            // own number cell. Any larger count proves that an external equation
            // reference exists. Such a reference must have a bookmark alias that
            // can be rebound to the target format; silently unlinking it would turn
            // a live, navigable reference into plain text. Fail before the first
            // document mutation instead.
            if (sourceReferenceCounts is not null)
            {
                foreach (var target in plan.Targets.Where(item => item.Numbered))
                {
                    if (!sourceReferenceCounts.TryGetValue(
                            target.SourceFormulaId,
                            out var referenceCount)
                        || referenceCount <= 1)
                        continue;
                    if (referenceAliasesByTargetId.ContainsKey(target.Id))
                        continue;
                    throw new InvalidDataException(
                        $"公式“{target.Latex}”存在 {referenceCount - 1} 个外部动态引用，但未能建立可跨格式恢复的编号书签。为避免把引用降级为普通文本，本次转换未执行。");
                }
            }
            Range? isolatedRecoveryScope = null;
            try
            {
                // The first MathType target has no sibling MTEqn fields to
                // renumber. With no external reference aliases, replacing one
                // validated direct-SEQ OMML row is an isolated body edit, just
                // like the existing local OMML content-edit transaction. Keep
                // full-document evidence for every operation that can touch
                // sibling numbering or external references.
                if (canFinalizeSingleMathTypeNumberLocally
                    && plan.SourceMode == FormulaOleContract.WordOmmlMode
                    && referenceAliasesByTargetId.Count == 0
                    && sourceReferenceCounts is not null
                    && sourceReferenceCounts.TryGetValue(plan.Targets[0].SourceFormulaId, out var localReferences)
                    && localReferences == 0)
                {
                    Table? table = null;
                    Range? owner = null;
                    Range? body = null;
                    Rows? rows = null;
                    try
                    {
                        table = WordEquationNumbering.FindNumberedEquationTable(document, plan.Targets[0].SourceFormulaId);
                        if (table is not null)
                        {
                            rows = table.Rows;
                            if (rows.Count == 1)
                            {
                                owner = table.Range;
                                body = document.Content;
                                isolatedRecoveryScope = document.Range(Math.Max(body.Start, owner.Start - 1),
                                    Math.Min(body.End, owner.End + 1));
                            }
                        }
                    }
                    finally { Release(rows); Release(body); Release(owner); Release(table); }
                }
                isolatedConversion = isolatedRecoveryScope is not null;
                rollbackSnapshot = new WordDocumentEditSnapshot(document,
                    plan.Targets.SelectMany(target => WordBookmarkRecoverySnapshot.NamesForFormula(target.SourceFormulaId))
                        .Concat(referenceAliasesByTargetId.Values.SelectMany(aliases => aliases.Select(alias => alias.Name))),
                    isolatedRecoveryScope);
            }
            finally { Release(isolatedRecoveryScope); }
            if (targetIsOmml && plan.Targets.Count > 1)
            {
                try
                {
                    batchWindow = _application.ActiveWindow;
                    if (batchWindow is not null)
                    {
                        previousViewType = batchWindow.View.Type;
                        if (previousViewType != WdViewType.wdNormalView)
                        {
                            batchWindow.View.Type = WdViewType.wdNormalView;
                            batchViewSuspended = true;
                        }
                    }
                    previousPagination = _application.Options.Pagination;
                    if (previousPagination)
                    {
                        _application.Options.Pagination = false;
                        paginationSuspended = true;
                    }
                }
                catch
                {
                    // View/pagination suspension is only a batch performance aid.
                    // Conversion correctness must not depend on it.
                }
            }

            conversionUndo = BeginUndoRecord("VisualTeX Convert Formula Format");
            if (conversionUndo is null)
                throw new InvalidOperationException("Word could not establish an independent conversion transaction.");

            if (plan.SourceMode == FormulaOleContract.NativeOleMode)
            {
                foreach (var target in plan.Targets.Where(item => item.Numbered))
                    if (WordEquationNumbering.IsEmptyOwnedCaptionFlowTailBeforeConversion(
                            document, target.SourceFormulaId))
                        verifiedEmptyCaptionTails.Add(target.SourceFormulaId);
            }

            // Capture every user-table structural owner before the first document
            // mutation. A later formula in the same cell must retain the ordinal it
            // had in the original source, not an ordinal recomputed after an earlier
            // source was replaced by its target. The VTFCC_* bookmark spans the full
            // cell and survives object replacement without retaining a mutable COM
            // Cell/Table/Range RCW.
            foreach (var target in plan.Targets.Where(item => item.SourceWithinTable))
            {
                if (!prepared.TryGetValue(target.Id, out var tablePrepared))
                    throw new InvalidDataException(
                        $"Missing rendered payload for table formula '{target.Latex}'.");
                var owner = CaptureUserCellForConversion(
                    document,
                    plan.SourceMode,
                    target,
                    tablePrepared.Session.FormulaId);
                if (owner is null) continue;
                userCellOwners.Add(owner);
                userCellOwnerByTargetId.Add(target.Id, owner);
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-user-cell-owners-captured count={userCellOwners.Count}");

            TraceConversionUndoState("before-preferences");
            WordEquationNumbering.RestoreEquationNumberFormatForConversion(document, plan.NumberFormatId);
            if (conversionMathFontMetadata is not null)
                ApplyDocumentOmmlMathFont(document, conversionMathFontMetadata);
            WordDoubleClickHook.TraceMessage($"format-conversion-transaction-start targets={plan.Targets.Count}");
            TraceConversionUndoState("after-preferences");

            // All source formats use one structural order. Temporary forward
            // bookmarks expand when an earlier paragraph becomes a table; they
            // are not an independent identity for a following inline object.
            var blockMathTypeOmmlGroups = targetIsOmml
                && string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                ? BuildContiguousBlockMathTypeOmmlGroups(document, plan.Targets)
                : Array.Empty<BlockMathTypeOmmlGroup>();
            var blockOmmlGroupByMember =
                new Dictionary<string, BlockMathTypeOmmlGroup>(StringComparer.Ordinal);
            foreach (var group in blockMathTypeOmmlGroups)
            {
                foreach (var member in group.Targets)
                    blockOmmlGroupByMember[member.Id] = group;
            }
            var processedBlockOmmlTargets = new HashSet<string>(StringComparer.Ordinal);
            if (blockMathTypeOmmlGroups.Count > 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-block-omml-groups groups={blockMathTypeOmmlGroups.Count} formulas={blockOmmlGroupByMember.Count}");
            }

            var adjacentInlineOmmlGroups = targetIsOmml
                && string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                ? BuildAdjacentInlineOmmlGroups(plan.Targets)
                : Array.Empty<IReadOnlyList<WordFormulaFormatConversionTarget>>();
            var adjacentOmmlGroupByMember =
                new Dictionary<string, IReadOnlyList<WordFormulaFormatConversionTarget>>(
                    StringComparer.Ordinal);
            foreach (var group in adjacentInlineOmmlGroups)
            {
                foreach (var member in group)
                    adjacentOmmlGroupByMember[member.Id] = group;
            }
            var processedAdjacentOmmlTargets = new HashSet<string>(StringComparer.Ordinal);
            if (adjacentInlineOmmlGroups.Count > 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-adjacent-omml-groups groups={adjacentInlineOmmlGroups.Count} formulas={adjacentOmmlGroupByMember.Count}");
            }

            var traceObjectCounts = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_COUNTS"),
                "1",
                StringComparison.Ordinal);
            var tracePerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
                "1",
                StringComparison.Ordinal);
            var orderedTargets = plan.Targets.OrderByDescending(item => item.SourceStart);
            foreach (var target in orderedTargets)
            {
                string? createdTargetBookmarkName = null;
                string? localMathTypeSourceBookmarkName = null;
                var stage = "capture-rollback-snapshot";
                var perfWatch = tracePerformance
                    ? System.Diagnostics.Stopwatch.StartNew()
                    : null;
                long perfLastMs = 0;
                void TracePerf(string perfStage)
                {
                    if (perfWatch is null) return;
                    var totalMs = perfWatch.ElapsedMilliseconds;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-perf stage={perfStage} formulaId={target.SourceFormulaId} numbered={target.Numbered} deltaMs={totalMs - perfLastMs} totalMs={totalMs}");
                    perfLastMs = totalMs;
                }
                try
                {
                    if (blockOmmlGroupByMember.TryGetValue(
                            target.Id,
                            out var blockGroup))
                    {
                        if (processedBlockOmmlTargets.Contains(target.Id))
                            continue;
                        stage = "block-omml-group";
                        if (ommlBatchSource is null)
                            throw new InvalidOperationException(
                                "The OMML batch source is unavailable for a consecutive MathType display group.");
                        var convertedBlockGroupCount = ApplyBlockMathTypeToOmmlGroup(
                            document,
                            plan,
                            blockGroup,
                            prepared,
                            ommlBatchSource);
                        result.FormulaCount += convertedBlockGroupCount;
                        foreach (var member in blockGroup.Targets)
                        {
                            processedBlockOmmlTargets.Add(member.Id);
                            successfulTargetIds.Add(member.Id);
                        }
                        continue;
                    }

                    if (adjacentOmmlGroupByMember.TryGetValue(target.Id, out var adjacentGroup))
                    {
                        if (processedAdjacentOmmlTargets.Contains(target.Id))
                            continue;
                        var groupAnchor = adjacentGroup
                            .OrderByDescending(member => member.SourceStart)
                            .First();
                        if (!string.Equals(target.Id, groupAnchor.Id, StringComparison.Ordinal))
                            continue;
                        stage = "adjacent-omml-group";
                        if (ommlBatchSource is null)
                            throw new InvalidOperationException(
                                "The OMML batch source is unavailable for an adjacent MathType formula group.");
                        var convertedGroupCount = ApplyAdjacentMathTypeToOmmlGroup(
                            document,
                            plan,
                            adjacentGroup,
                            prepared,
                            ommlBatchSource);
                        result.FormulaCount += convertedGroupCount;
                        foreach (var member in adjacentGroup)
                        {
                            processedAdjacentOmmlTargets.Add(member.Id);
                            successfulTargetIds.Add(member.Id);
                        }
                        continue;
                    }

                    if (string.Equals(
                            plan.SourceMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal))
                    {
                        // A MathType formula has no durable VisualTeX FormulaId.
                        // Its captured numeric Word range is only a location hint:
                        // deleting an earlier display paragraph can move the next
                        // Equation.DSMT4 into exactly the same Start/End coordinates.
                        // Bind this concrete source object before mutation and use the
                        // bookmark itself for all source-removal checks in this one
                        // formula transaction.
                        if (target.SourceHasMathTypeSectionPrefix)
                            IsolateMathTypeSectionPrefix(document, target);
                        localMathTypeSourceBookmarkName =
                            CreateMathTypeSourceIdentityBookmark(document, target);
                    }
                    var formula = prepared[target.Id];
                    userCellOwnerByTargetId.TryGetValue(target.Id, out var userCellOwner);
                    // MathType replacements have an exact per-object VTMT bookmark,
                    // while VisualTeX OLE replacements bind their durable VTO_
                    // identity bookmark before InsertOle returns. Both can therefore
                    // be verified locally. Re-enumerating document.InlineShapes for
                    // every target is O(N^2) and eventually destabilizes Word after
                    // dozens of delete/insert mutations. OMML still uses the older
                    // count stabilization until it gets an equivalent local anchor.
                    var targetObjectCountBefore = targetIsMathType || targetIsVisualTeX
                        ? -1
                        : CountSimpleFormatObjects(document, plan.TargetMode);
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "before-delete");
                    stage = "delete-source";
                    var sourceIsOmml = string.Equals(
                        plan.SourceMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal);
                    var useDirectUnnumberedOmmlDelete = sourceIsOmml && !target.Numbered;
                    var useDirectSingleNativeOmmlDelete = false;
                    var useDirectSingleManagedDisplayOmmlDelete =
                        (targetIsMathType || targetIsVisualTeX)
                        && plan.Targets.Count == 1
                        && sourceIsOmml
                        && !target.Numbered
                        && target.SourceIsManagedOmml
                        && !target.SourceWithinTable
                        && string.Equals(
                            target.DisplayMode,
                            "block",
                            StringComparison.Ordinal);
                    var useDirectSingleManagedNumberedOmmlDelete =
                        (targetIsMathType || targetIsVisualTeX)
                        && sourceIsOmml
                        && target.Numbered
                        && target.SourceIsManagedOmml
                        && string.Equals(
                            target.DisplayMode,
                            "block",
                            StringComparison.Ordinal);
                    var preserveFormulaCrossReferences =
                        referenceAliasesByTargetId.ContainsKey(target.Id);
                    var useAtomicMathTypeParagraphReplacement =
                        targetIsOmml
                        && ommlBatchSource is not null
                        && !target.SourceWithinTable
                        && string.Equals(
                            plan.SourceMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal)
                        && string.Equals(
                            target.DisplayMode,
                            "block",
                            StringComparison.Ordinal);
                    var useAtomicMathTypeVisualTeXParagraphReplacement =
                        targetIsVisualTeX
                        && mathTypeVisualTeXCleanParagraphBuffer is not null
                        && !target.SourceWithinTable
                        && !string.IsNullOrWhiteSpace(localMathTypeSourceBookmarkName)
                        && target.Numbered
                        && string.Equals(
                            plan.SourceMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal)
                        && string.Equals(
                            target.DisplayMode,
                            "block",
                            StringComparison.Ordinal);
                    TraceConversionUndoState("before-source-delete");
                    var insertionStart = useDirectUnnumberedOmmlDelete
                        ? DeleteUnnumberedOmmlSourceExactly(document, target)
                        : useDirectSingleNativeOmmlDelete
                        ? DeleteSingleNativeOmmlSourceDirect(document, target)
                        : useDirectSingleManagedDisplayOmmlDelete
                            ? DeleteSingleManagedDisplayOmmlSourceDirect(
                                document,
                                target)
                        : useDirectSingleManagedNumberedOmmlDelete
                            ? DeleteSingleManagedNumberedOmmlSourceDirect(
                                document,
                                target,
                                sourceReferenceCounts,
                                preserveFormulaCrossReferences,
                                deferredOwnedOmmlSeparators)
                            : useAtomicMathTypeParagraphReplacement
                                ? ReplaceMathTypeDisplaySourceParagraphAtomically(
                                    document,
                                    target,
                                    ommlBatchSource!)
                            : useAtomicMathTypeVisualTeXParagraphReplacement
                                ? ReplaceNumberedMathTypeDisplaySourceParagraphForVisualTeX(
                                    document,
                                    target,
                                    mathTypeVisualTeXCleanParagraphBuffer!,
                                    localMathTypeSourceBookmarkName!)
                                : DeleteSimpleSourceHost(
                                    document,
                                    plan.SourceMode,
                                    target,
                                    sourceReferenceCounts,
                                    preserveFormulaCrossReferences,
                                    useIdentityBookmarkSourceCheck:
                                        canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath,
                                    deferCaptionFlowTailCleanup: true);
                    if (useDirectUnnumberedOmmlDelete)
                        WordDoubleClickHook.TraceMessage($"format-conversion-exact-omml-delete formulaId={target.SourceFormulaId} start={insertionStart}");
                    else if (useDirectSingleNativeOmmlDelete)
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-omml-delete formulaId={target.SourceFormulaId} start={insertionStart}");
                    else if (useDirectSingleManagedDisplayOmmlDelete)
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-managed-display-omml-delete formulaId={target.SourceFormulaId} start={insertionStart}");
                    else if (useDirectSingleManagedNumberedOmmlDelete)
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-numbered-omml-delete formulaId={target.SourceFormulaId} start={insertionStart}");
                    TracePerf("delete-source-host");
                    TraceConversionUndoState("after-source-delete");
                    if (!string.IsNullOrWhiteSpace(localMathTypeSourceBookmarkName))
                        EnsureForwardMathTypeSourceRemoved(
                            document,
                            localMathTypeSourceBookmarkName!,
                            target,
                            "delete-source");
                    else if (!canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
                        EnsureSimpleFormatSourceRemoved(document, plan.SourceMode, target, "delete-source");
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "after-delete");
                    ThrowIfSimpleFormatConversionFailureInjected(target);
                    var content = document.Content;
                    try
                    {
                        insertionStart = Math.Max(
                            content.Start,
                            Math.Min(insertionStart, content.End));
                    }
                    finally { Release(content); }

                    stage = "prepare-target-selection";
                    selection.SetRange(insertionStart, insertionStart);
                    selection.Collapse(WdCollapseDirection.wdCollapseStart);

                    var session = formula.Session;
                    session.Mode = "create";
                    session.SourceDocumentId = plan.DocumentId;
                    session.SourceObjectId =
                        $"{RangeReferencePrefix}{insertionStart}:{insertionStart}";
                    session.DisplayMode = target.DisplayMode;
                    session.ObjectMode = plan.TargetMode;
                    session.Numbered = target.Numbered;
                    session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                    session.FontSizePt = target.FontSizePt;
                    session.OriginalMetadata = null;

                    stage = "insert-target";
                    if (targetIsMathType)
                    {
                        createdTargetBookmarkName =
                            "VTMT_" + target.Id.Replace("-", string.Empty);
                        ResolvedEquationHeadingScope? preResolvedHeadingScope = null;
                        if (target.Numbered && batchHeadingScopes is not null)
                            batchHeadingScopes.TryGetValue(
                                target.SourceStart,
                                out preResolvedHeadingScope);
                        var batchNativePreview = formula.MathTypeNativePreview;
                        InsertMathTypeOle(
                            session,
                            formula.MathMl!,
                            formula.EmfPath!,
                            createdTargetBookmarkName,
                            preResolvedHeadingScope,
                            preparedHeadingScopeStarts,
                            batchNativePreview?.WmfPath,
                            batchNativePreview?.WidthPt ?? 0,
                            batchNativePreview?.HeightPt ?? 0,
                            batchNativePreview?.WordPosition ?? 0,
                            formula.MathTypeNativePreviewAttempted,
                            reuseExistingInlineTypingBoundary:
                                useDirectSingleNativeOmmlDelete,
                            updateCreatedMathTypeNumberFields:
                                canFinalizeSingleMathTypeNumberLocally,
                            preserveExistingDisplayParagraphBoundary:
                                string.Equals(
                                    target.DisplayMode,
                                    "block",
                                    StringComparison.Ordinal),
                            knownDisplayColumnWidth:
                                target.MathTypeDisplayColumnWidth,
                            preserveCapturedInsertion: true);
                    }
                    else if (string.Equals(
                                 plan.TargetMode,
                                 FormulaOleContract.WordOmmlMode,
                                 StringComparison.Ordinal))
                    {
                        // Keep format conversion on the same performance architecture as the
                        // pre-MathType numbered-formula work: materialize formula content first,
                        // then repair numbering once after the batch. Building a table + hidden
                        // SEQ caption + visible REF while MathType numbering structures still
                        // remain later in the document forces Word to repaginate/maintain fields
                        // on every item and turns a 50-formula conversion into O(N^2) work.
                        // Metadata still records Numbered=true, so the final numbering refresh
                        // below can rebuild the standard VisualTeX structure without changing
                        // formula semantics or the target OMML payload.
                        InsertOmml(
                            session,
                            formula.MathMl!,
                            // Keep numbered targets as pure OMath until every
                            // MathType source and MTPlaceRef field has been removed.
                            // Mixing a new VisualTeX direct SEQ table with a remaining
                            // numbered Equation.DSMT4 paragraph crashes Word 2021.
                            deferNumberingLayout: true,
                            deferFinalFingerprint: true,
                            ommlBatchSource: ommlBatchSource,
                            preserveExistingDisplayParagraphBoundary: true,
                            normalizeMathTypeDisplayParagraph:
                                string.Equals(
                                    plan.SourceMode,
                                    FormulaOleContract.MathTypeOleMode,
                                    StringComparison.Ordinal)
                                && string.Equals(
                                    target.DisplayMode,
                                    "block",
                                    StringComparison.Ordinal),
                            retainInsertedMath: range => liveOmmlInsertions.Capture(session.FormulaId, range),
                            preserveCapturedInsertion: true);
                    }
                    else
                    {
                        // A single OMML/MathType -> VisualTeX replacement does not
                        // need the deferred batch-numbering scaffold. Let the mature
                        // InsertOle path build the complete center/right-tab + hidden
                        // SEQ + visible REF structure inside the same per-formula
                        // transaction. Deferred numbering remains reserved for real
                        // multi-formula batches where it avoids O(N^2) field work.
                        var deferVisualTeXNumbering = targetIsVisualTeX
                            && (plan.Targets.Count > 1
                                || deferredOwnedOmmlSeparators.Contains(target.SourceFormulaId));
                        deferredSingleVisualTeXNumbering |= deferVisualTeXNumbering && plan.Targets.Count == 1;
                        InsertOle(
                            session,
                            formula.PngPath!,
                            formula.EmfPath!,
                            deferNumberingLayout: deferVisualTeXNumbering,
                            preserveExistingDisplayParagraphBoundary:
                                string.Equals(
                                    target.DisplayMode,
                                    "block",
                                    StringComparison.Ordinal),
                            preserveCapturedInsertion: true);
                    }
                    TracePerf("insert-target");
                    TraceConversionUndoState("after-target-insert");
                    if (!string.IsNullOrWhiteSpace(localMathTypeSourceBookmarkName))
                        EnsureForwardMathTypeSourceRemoved(
                            document,
                            localMathTypeSourceBookmarkName!,
                            target,
                            "insert-target");
                    else if (!canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
                        EnsureSimpleFormatSourceRemoved(document, plan.SourceMode, target, "insert-target");
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "after-insert-before-commit");
                    stage = "commit-undo";
                    stage = "verify-target";
                    if (!string.IsNullOrWhiteSpace(createdTargetBookmarkName))
                    {
                        EnsureBookmarkedMathTypeTargetSurvived(
                            document,
                            createdTargetBookmarkName,
                            target);
                        if (userCellOwner is not null)
                            userCellOwner.MathTypeBookmark = createdTargetBookmarkName;
                        // Body targets need an exact locator through final host
                        // verification too. Counting five OLEs cannot detect a body
                        // equation which Word has pulled into a user's first cell.
                        retainedMathTypeTargetBookmarks[target.Id] = createdTargetBookmarkName;
                    }
                    else if (targetIsVisualTeX)
                    {
                        EnsureBookmarkedVisualTeXTargetSurvived(
                            document,
                            session.FormulaId,
                            target,
                            insertionStart);
                        if (plan.Targets.Count == 1
                            && !deferredSingleVisualTeXNumbering
                            && target.Numbered
                            && string.Equals(
                                target.DisplayMode,
                                "block",
                                StringComparison.Ordinal))
                        {
                            EnsureConvertedVisualTeXNumberingHostHealthy(
                                document,
                                session.FormulaId,
                                target);
                        }
                    }
                    else
                    {
                        WaitForSimpleTargetObjectCountToStabilize(
                            document,
                            plan.TargetMode,
                            targetObjectCountBefore + 1,
                            target);
                    }
                    if ((targetIsVisualTeX || targetIsMathType)
                        && !string.IsNullOrEmpty(target.FollowingInlineTableText))
                    {
                        stage = "restore-table-inline-boundary";
                        RestoreTableInlineFollowingBoundaryForOleTarget(
                            document,
                            target,
                            plan.TargetMode,
                            session.FormulaId,
                            createdTargetBookmarkName);
                    }
                    if (!string.IsNullOrWhiteSpace(localMathTypeSourceBookmarkName))
                        EnsureForwardMathTypeSourceRemoved(
                            document,
                            localMathTypeSourceBookmarkName!,
                            target,
                            "post-transaction");
                    else if (!canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
                        EnsureSimpleFormatSourceRemoved(
                            document,
                            plan.SourceMode,
                            target,
                            "post-transaction");
                    if ((useDirectUnnumberedOmmlDelete && target.SourceIsManagedOmml)
                        || useDirectSingleManagedDisplayOmmlDelete
                        || useDirectSingleManagedNumberedOmmlDelete)
                    {
                        // Keep the managed OMML metadata intact until the new
                        // Equation.DSMT4 object has survived its Word transaction.
                        // That preserves the mature Undo rollback path if insertion
                        // fails; only the now-dead source identity is removed here.
                        stage = "delete-source-metadata";
                        WordOmmlFormulaStore.Delete(
                            document,
                            target.SourceFormulaId);
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-direct-managed-omml-metadata-deleted formulaId={target.SourceFormulaId} numbered={target.Numbered}");
                    }
                    if (userCellOwner is not null)
                        VerifyConvertedUserCell(document, plan.TargetMode, userCellOwner, liveOmmlInsertions);
                    TracePerf("verify-target");
                    if (traceObjectCounts)
                        TraceSimpleFormatObjectCounts(document, target, "after-commit-stable");
                    result.FormulaCount++;
                    successfulTargetIds.Add(target.Id);
                }
                catch (Exception error)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-item-failed stage={stage} formulaId={target.SourceFormulaId} latex={target.Latex} error={error}");
                    throw;
                }
                finally
                {
                    if (document is not null
                        && !string.IsNullOrWhiteSpace(createdTargetBookmarkName)
                        && !retainedMathTypeTargetBookmarks.ContainsKey(target.Id))
                        TryDeleteBookmark(document, createdTargetBookmarkName);
                    if (document is not null
                        && !string.IsNullOrWhiteSpace(localMathTypeSourceBookmarkName))
                        TryDeleteBookmark(document, localMathTypeSourceBookmarkName!);
                }
            }

            var finalizeWatch = tracePerformance
                ? System.Diagnostics.Stopwatch.StartNew()
                : null;
            long finalizeLastMs = 0;
            void TraceFinalize(string finalizeStage)
            {
                if (finalizeWatch is null) return;
                var totalMs = finalizeWatch.ElapsedMilliseconds;
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-finalize-perf stage={finalizeStage} sourceMode={plan.SourceMode} targetMode={plan.TargetMode} deltaMs={totalMs - finalizeLastMs} totalMs={totalMs}");
                finalizeLastMs = totalMs;
            }

            var useWholeDocumentVisualTeXToMathTypeCardinalityFastPath =
                canUseWholeDocumentVisualTeXToMathTypeCardinalityFastPath
                && successfulTargetIds.Count == plan.Targets.Count;
            (int VisualTeX, int MathType)? canonicalOleCountsAfterInsert = null;
            if (useWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
            {
                canonicalOleCountsAfterInsert = ReadCanonicalOleObjectCounts(document);
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-canonical-ole-counts stage=after-insert visualtex={canonicalOleCountsAfterInsert.Value.VisualTeX} mathtype={canonicalOleCountsAfterInsert.Value.MathType}");
            }

            // Validate the state that was actually committed, even when a later
            // formula failed. Unprocessed source hosts are expected to remain; only
            // a source belonging to successfulTargetIds is evidence of rollback or
            // Word resurrecting an object after commit.
            if (string.Equals(
                    plan.SourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                // MathType sources are third-party OLEs without a durable VisualTeX
                // FormulaId. A surviving unrelated Equation.DSMT4 can shift into the
                // deleted source's old numeric range, so final verification must use
                // the exact per-item bookmark checks above plus the source object
                // count, never a stale captured range.
                var finalSourceObjectCount = CountSimpleFormatObjects(document, plan.SourceMode);
                var expectedSourceObjectCount = Math.Max(
                    0,
                    initialSourceObjectCount - successfulTargetIds.Count);
                if (finalSourceObjectCount != expectedSourceObjectCount)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"Source formula count mismatch after partial conversion: expected {expectedSourceObjectCount}, actual {finalSourceObjectCount}.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-source-count-mismatch sourceMode={plan.SourceMode} initial={initialSourceObjectCount} committed={successfulTargetIds.Count} expected={expectedSourceObjectCount} actual={finalSourceObjectCount}");
                }
            }
            else if (useWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
            {
                var expectedSourceObjectCount = Math.Max(
                    0,
                    initialSourceObjectCount - successfulTargetIds.Count);
                var actualSourceObjectCount = canonicalOleCountsAfterInsert!.Value.VisualTeX;
                if (actualSourceObjectCount != expectedSourceObjectCount)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"Source formula count mismatch after conversion: expected {expectedSourceObjectCount}, actual {actualSourceObjectCount}.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-source-count-mismatch sourceMode={plan.SourceMode} initial={initialSourceObjectCount} committed={successfulTargetIds.Count} expected={expectedSourceObjectCount} actual={actualSourceObjectCount} source=flat-opc");
                }
            }
            else
            {
                foreach (var target in plan.Targets.Where(item => successfulTargetIds.Contains(item.Id)).ToArray())
                {
                    if (!IsSimpleFormatSourcePresent(document, plan.SourceMode, target))
                        continue;
                    successfulTargetIds.Remove(target.Id);
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"{target.Latex}: Word restored the source formula after the conversion transaction completed.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-source-reappeared formulaId={target.SourceFormulaId} latex={target.Latex}");
                }
            }
            result.FormulaCount = successfulTargetIds.Count;
            TraceFinalize("source-residual-check");

            var finalTargetObjectCount = useWholeDocumentVisualTeXToMathTypeCardinalityFastPath
                ? canonicalOleCountsAfterInsert!.Value.MathType
                : CountSimpleFormatObjects(document, plan.TargetMode);
            var expectedTargetObjectCount = initialTargetObjectCount + successfulTargetIds.Count;
            if (finalTargetObjectCount != expectedTargetObjectCount)
            {
                result.FailedFormulaCount++;
                result.Failures.Add(
                    $"Target formula count mismatch after partial conversion: expected {expectedTargetObjectCount}, actual {finalTargetObjectCount}. Word removed or failed to retain a converted formula.");
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-target-count-mismatch targetMode={plan.TargetMode} initial={initialTargetObjectCount} committed={successfulTargetIds.Count} expected={expectedTargetObjectCount} actual={finalTargetObjectCount}");
            }
            TraceFinalize("target-count-check");

            var successfulTargets = plan.Targets
                .Where(target => successfulTargetIds.Contains(target.Id))
                .ToArray();
            var convertedOmmlFormulaIds = targetIsOmml
                ? successfulTargets
                    .Select(target => prepared[target.Id].Session.FormulaId)
                    .Where(formulaId => !string.IsNullOrWhiteSpace(formulaId))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray()
                : Array.Empty<string>();
            var convertedOmmlNumberingFormulaIds = targetIsOmml
                ? successfulTargets
                    .Where(target => target.Numbered
                        && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                    .Select(target => prepared[target.Id].Session.FormulaId)
                    .ToArray()
                : Array.Empty<string>();
            var verifiedConvertedOmmlMetadata = new Dictionary<string, FormulaMetadata>(
                StringComparer.OrdinalIgnoreCase);
            if (targetIsVisualTeX && deferredOwnedOmmlSeparators.Count > 0)
            {
                // VisualTeX OLE numbering is a separate phase. Retire the old
                // OMML separator after bare OLE insertion and BEFORE its new
                // caption is created, so the caption can own its own flow tail.
                VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                    prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
                foreach (var sourceId in deferredOwnedOmmlSeparators)
                    WordEquationNumbering.CompleteDeferredOmmlSeparatorCleanup(document, sourceId);
                deferredOwnedOmmlSeparators.Clear();
                VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                    prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
            }

            // InsertOmml deliberately defers final WordOpenXML fingerprinting during
            // a batch. Word can normalize a freshly inserted OMath before numbering
            // is built, so the converter-side source fingerprint is provisional.
            // Refresh the successful subset from one document snapshot before any
            // numbered structural edit. This makes a bare Numbered=true OMath a
            // first-class conversion state without weakening the strict resolver
            // used for already-finalized managed documents.
            if (targetIsOmml && convertedOmmlFormulaIds.Length > 0)
            {
                FinishBatchPresentation();
                Release(conversionUndo);
                conversionUndo = BeginUndoRecord("VisualTeX Finalize Formula Conversion")
                    ?? throw new InvalidOperationException("Word could not establish the conversion finalization record.");
                conversionUndoEnded = false;
                var convertedOmmlDisplayModes = successfulTargets.ToDictionary(
                    target => prepared[target.Id].Session.FormulaId,
                    target => target.DisplayMode,
                    StringComparer.OrdinalIgnoreCase);
                liveOmmlInsertions.NormalizeStableTableOwnerTypes(
                    convertedOmmlDisplayModes,
                    formulaId => TryResolveConvertedOmmlUserCellOwner(
                        document,
                        userCellOwners,
                        formulaId));
                var indexed = WordOmmlNativeSource.RefreshFingerprintsFromDocumentOpenXml(
                    document, convertedOmmlFormulaIds,
                    convertedOmmlFormulaIds.ToDictionary(id => id, id => ommlBatchSource!.GetSourceOmml(id), StringComparer.OrdinalIgnoreCase),
                    liveOmmlInsertions,
                    verifiedConvertedOmmlMetadata);
                if (indexed != convertedOmmlFormulaIds.Length)
                    throw new InvalidDataException("The conversion identity index is incomplete before numbering.");
                WordDoubleClickHook.TraceMessage($"format-conversion-omml-identities-before-numbering count={indexed}");
                if (plan.SourceMode == FormulaOleContract.NativeOleMode)
                {
                    // Bare OMML is already materialized and content-verified here.
                    // Retire the old OLE caption tail before building the NEW OMML
                    // table separator; otherwise its old bookmark spans the newly
                    // inserted anchor and looks like protected unrelated content.
                    VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                        prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
                    foreach (var target in successfulTargets.Where(item => item.Numbered))
                        WordEquationNumbering.CompleteDeferredCaptionFlowTailCleanup(
                            document, target.SourceFormulaId,
                            verifiedEmptyCaptionTails.Contains(target.SourceFormulaId));
                    VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                        prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
                }
                foreach (var owner in userCellOwners)
                    owner.ReanchorAtCurrentCellTerminator(document);
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-owners-reanchored-before-numbering count={userCellOwners.Count}");
            }
            if (result.FormulaCount > 0)
            {
                try
                {
                    if (targetIsMathType)
                    {
                        var hasConvertedNumberedMathType = successfulTargets.Any(target =>
                            target.Numbered
                            && string.Equals(
                                target.DisplayMode,
                                "block",
                                StringComparison.Ordinal));
                        if (!hasConvertedNumberedMathType)
                        {
                            WordDoubleClickHook.TraceMessage(
                                "format-conversion-numbering-mathtype-skipped reason=no-numbered-targets");
                        }
                        else if (!canFinalizeSingleMathTypeNumberLocally)
                        {
                            MathTypeEquationNumbering.UpdateNumberFields(document);
                        }
                        else
                        {
                            WordDoubleClickHook.TraceMessage(
                                "format-conversion-numbering-local-mathtype-single finalized=1");
                        }
                    }
                    else if (targetIsOmml)
                    {
                    WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                        document,
                        plan.NumberFormatId);
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-number-format-restored format={plan.NumberFormatId}");
                    // Finalize only the known converted rows. A partial local
                    // failure must roll back; a document-wide legacy rebuild can
                    // overwrite unrelated numbering and refresh compatibility
                    // references before their captured aliases have been rebound.
                    var builtNumbered = convertedOmmlNumberingFormulaIds.Length;
                    if (!ConvertedOmmlDirectTablesAreAlreadyComplete(document, convertedOmmlNumberingFormulaIds))
                        builtNumbered = WordEquationNumbering.BuildNewOmmlNumberingBatch(
                            document,
                            convertedOmmlNumberingFormulaIds,
                            deferFinalIdentityPersistence: true,
                            cleanupCommittedSourceSpacing: true,
                            verifiedFreshMetadata: verifiedConvertedOmmlMetadata,
                            verifiedFreshInsertions: liveOmmlInsertions,
                            // MathType numbering is foreign MTPlaceRef/MT* state and
                            // cannot collide with a fresh VisualTeX FormulaId. A
                            // VisualTeX source, however, has just dismantled its own
                            // SEQ/REF scaffold in the same Word transaction; retain
                            // the mature generated-artifact cleanup for that path.
                            verifiedFreshBareOmml: string.Equals(
                                plan.SourceMode,
                                FormulaOleContract.MathTypeOleMode,
                                StringComparison.Ordinal));
                    if (builtNumbered != convertedOmmlNumberingFormulaIds.Length
                        || !ConvertedOmmlDirectTablesAreAlreadyComplete(document, convertedOmmlNumberingFormulaIds,
                            verifiedConvertedOmmlMetadata))
                        throw new InvalidDataException("The converted OMML row identities or number fields are incomplete.");
                    var removedBlankParagraphs = RestoreConvertedOmmlPrecedingBlankParagraphCounts(
                        document, plan, prepared, successfulTargetIds);
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-numbering-direct-table-complete targetMode={plan.TargetMode} built={builtNumbered} removedBlankParagraphs={removedBlankParagraphs}");

                    var orderedConvertedNumberedOmmlFormulaIds = successfulTargets
                        .Where(target => target.Numbered
                            && string.Equals(
                                target.DisplayMode,
                                "block",
                                StringComparison.Ordinal))
                        .OrderBy(target => target.SourceStart)
                        .Select(target => prepared[target.Id].Session.FormulaId)
                        .Where(formulaId => !string.IsNullOrWhiteSpace(formulaId))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .Where(formulaId =>
                        {
                            // Only body 1x3 hosts participate in row compaction.
                            // Numbered equations inside a user cell are complete
                            // native #SEQ paragraphs, not missing managed tables.
                            var table = WordEquationNumbering.FindNumberedEquationTable(document, formulaId);
                            try { return table is not null; }
                            finally { Release(table); }
                        })
                        .ToArray();
                    if (orderedConvertedNumberedOmmlFormulaIds.Length > 1)
                    {
                        var repairedOleAnchors = RepairExpandedOleIdentityBookmarks(document);
                        WordDoubleClickHook.TraceMessage($"format-conversion-adjacent-ole-identities-rebound count={repairedOleAnchors}");
                        var mergedNumberTableBoundaries =
                            WordEquationNumbering.MergeAdjacentManagedNativeOmmlNumberTableRows(
                                document,
                                orderedConvertedNumberedOmmlFormulaIds);
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-numbering-table-rows-grouped boundaries={mergedNumberTableBoundaries} formulas={orderedConvertedNumberedOmmlFormulaIds.Length}");
                    }
                }
                else if (targetIsVisualTeX)
                {
                    WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                        document,
                        plan.NumberFormatId);
                    if (plan.Targets.Count == 1 && !deferredSingleVisualTeXNumbering)
                    {
                        // Single-target VisualTeX conversion was finalized by
                        // InsertOle inside the formula transaction and was locally
                        // validated before that transaction was accepted. Do not
                        // dismantle/rebuild it a second time and do not make this
                        // successful conversion depend on unrelated pre-existing
                        // numbering damage elsewhere in the document.
                        WordDoubleClickHook.TraceMessage(
                            "format-conversion-numbering-visualtex-single-reused");
                    }
                    else
                    {
                        var convertedVisualTeXNumberedIds = successfulTargets
                            .Where(target => target.Numbered && target.DisplayMode == "block")
                            .Select(target => prepared[target.Id].Session.FormulaId)
                            .Where(id => !string.IsNullOrWhiteSpace(id))
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToArray();
                        var rebuiltVisualTeXNumbered =
                            BuildNewVisualTeXNumberingBatch(
                                document,
                                successfulTargets.Select(target => prepared[target.Id].Session.ToMetadata()).ToArray(),
                                normalizeMathTypeStyle: plan.SourceMode == FormulaOleContract.MathTypeOleMode,
                                usePlannedFieldFastPath: convertedVisualTeXNumberedIds.Length > 1,
                                out var visualTeXNumberingFinalizedInBatch);
                        if (rebuiltVisualTeXNumbered > 0)
                        {
                            var finalizedVisualTeXNumbered = rebuiltVisualTeXNumbered;
                            if (!visualTeXNumberingFinalizedInBatch)
                            {
                                if (!WordEquationNumbering.TryFinalizeHealthyConversionNumbering(
                                        document,
                                        out finalizedVisualTeXNumbered,
                                        convertedVisualTeXNumberedIds))
                                    throw new InvalidOperationException(
                                        "Converted VisualTeX numbering scaffolds could not be finalized safely.");
                            }
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-numbering-visualtex-batch built={rebuiltVisualTeXNumbered} finalized={finalizedVisualTeXNumbered} plannedFastPath={visualTeXNumberingFinalizedInBatch}");
                        }
                    }
                }
                }
                catch (Exception targetFinalizeError)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"Converted target finalization failed: {targetFinalizeError.Message}");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-target-finalize-failed sourceMode={plan.SourceMode} targetMode={plan.TargetMode} committed={successfulTargetIds.Count} error={targetFinalizeError}");
                    throw;
                }
            }
            var successfulReferenceAliasesByTargetId = referenceAliasesByTargetId
                .Where(entry => successfulTargetIds.Contains(entry.Key))
                .ToArray();
            if (successfulReferenceAliasesByTargetId.Length > 0)
            {
                try
                {
                    var restoredAliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var entry in successfulReferenceAliasesByTargetId)
                    {
                        int restored;
                        if (targetIsMathType)
                        {
                            if (!retainedMathTypeTargetBookmarks.TryGetValue(
                                    entry.Key,
                                    out var targetBookmarkName))
                                throw new InvalidDataException(
                                    $"Missing converted MathType target locator for reference aliases on {entry.Key}.");
                            restored = MathTypeEquationReferences.RestoreFormatConversionAliasesToMathType(
                                document,
                                targetBookmarkName,
                                entry.Value);
                        }
                        else
                        {
                            if (!prepared.TryGetValue(entry.Key, out var convertedFormula))
                                throw new InvalidDataException(
                                    $"Missing converted VisualTeX payload for reference aliases on {entry.Key}.");
                            restored = MathTypeEquationReferences.RestoreFormatConversionAliasesToVisualTeX(
                                document,
                                convertedFormula.Session.FormulaId,
                                entry.Value);
                        }
                        if (restored != entry.Value.Count)
                            throw new InvalidDataException(
                                $"Restored {restored}/{entry.Value.Count} equation-reference aliases for conversion target {entry.Key}.");
                        foreach (var alias in entry.Value)
                            restoredAliases.Add(alias.Name);
                    }
                    var refreshedReferences = MathTypeEquationReferences.RefreshReferences(
                        document,
                        restoredAliases);
                    var successfulCapturedReferenceFormatting =
                        capturedReferenceFormatting
                            .Where(entry => restoredAliases.Contains(entry.Key))
                            .ToDictionary(
                                entry => entry.Key,
                                entry => entry.Value,
                                StringComparer.OrdinalIgnoreCase);
                    var expectedFormattingCount =
                        successfulCapturedReferenceFormatting.Values.Sum(items => items.Count);
                    var restoredFormattingCount =
                        MathTypeEquationReferences.RestoreReferenceCharacterFormatting(
                            document,
                            successfulCapturedReferenceFormatting);
                    if (restoredFormattingCount != expectedFormattingCount)
                        throw new InvalidDataException(
                            $"Restored {restoredFormattingCount}/{expectedFormattingCount} equation-reference formatting snapshots.");
                    WordEquationReferenceFields.ValidateReferences(document,
                        successfulCapturedReferenceFormatting.ToDictionary(
                            entry => entry.Key, entry => entry.Value.Count,
                            StringComparer.OrdinalIgnoreCase));
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-reference-aliases-restored sourceMode={plan.SourceMode} targetMode={plan.TargetMode} aliases={restoredAliases.Count} refreshedReferences={refreshedReferences} restoredFormatting={restoredFormattingCount}");
                }
                catch (Exception referenceAliasError)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"Equation-reference compatibility finalization failed: {referenceAliasError.Message}");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-reference-alias-finalize-failed error={referenceAliasError}");
                    throw;
                }
            }
            // Every destination has its number and every captured alias has
            // been rebound before a document-wide equation-reference refresh.
            // An entirely unnumbered conversion cannot affect any VisualTeX/MathType
            // equation-reference field. Scanning the full field tree in that case
            // cost about 40 seconds in the 1000-OLE Word 2021 acceptance document
            // without changing a single field. Keep the global refresh whenever a
            // successful numbered target or a captured compatibility alias exists.
            var needsGlobalEquationReferenceRefresh =
                !isolatedConversion
                && (successfulReferenceAliasesByTargetId.Length > 0
                    || successfulTargets.Any(target => target.Numbered));
            if (needsGlobalEquationReferenceRefresh)
            {
                WordEquationReferenceFields.UpdateReferences(document);
            }
            else
            {
                WordDoubleClickHook.TraceMessage(
                    "format-conversion-reference-refresh-skipped reason=no-affected-numbering-or-aliases");
            }
            TraceFinalize("numbering-reconcile");
            // Do not recover a user-cell locator before proving that every body
            // target stayed in the body (or in its own managed OMML number row).
            // Otherwise a displaced body OLE can occupy the user's original slot
            // and a row/column/ordinal match would certify the wrong equation.
            var verifyBodyTableBoundary = document.Tables.Count > 0;
            if (verifyBodyTableBoundary)
                VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                    prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
            if (plan.SourceMode == FormulaOleContract.NativeOleMode && !targetIsOmml)
            {
                foreach (var target in successfulTargets.Where(item => item.Numbered))
                    WordEquationNumbering.CompleteDeferredCaptionFlowTailCleanup(
                        document, target.SourceFormulaId,
                        verifiedEmptyCaptionTails.Contains(target.SourceFormulaId));
                if (verifyBodyTableBoundary)
                    VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                        prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
            }
            if (deferredOwnedOmmlSeparators.Count > 0)
            {
                foreach (var sourceId in deferredOwnedOmmlSeparators)
                    WordEquationNumbering.CompleteDeferredOmmlSeparatorCleanup(document, sourceId);
                deferredOwnedOmmlSeparators.Clear();
                if (verifyBodyTableBoundary)
                    VerifyConvertedBodyHosts(document, plan.TargetMode, successfulTargets,
                        prepared, retainedMathTypeTargetBookmarks, userCellOwnerByTargetId);
            }
            TraceFinalize("physical-body-hosts-and-deferred-tails");
            if (targetIsOmml)
            {
                // Word can make a collapsed cell-terminator bookmark ambiguous when
                // a managed 1x3 numbering table is inserted at the same boundary:
                // both sides may legitimately report row=1/column=1. Rebuild every
                // user-cell structural anchor from the formula's own VTFCO owner,
                // which was content-verified before numbering, and verify its OMML
                // fingerprint again before allowing it to define the cell.
                foreach (var owner in userCellOwners)
                {
                    Range? verifiedTarget = null;
                    try
                    {
                        verifiedTarget = liveOmmlInsertions.TryReadVerifiedOwner(owner.FormulaId)
                            ?? throw new InvalidDataException(
                                $"Verified OMML owner {owner.FormulaId} disappeared after numbering.");
                        if (!verifiedConvertedOmmlMetadata.TryGetValue(owner.FormulaId, out var verifiedMetadata)
                            || string.IsNullOrWhiteSpace(verifiedMetadata.NativeOmmlFingerprint))
                            throw new InvalidDataException(
                                $"Verified OMML metadata {owner.FormulaId} is unavailable after numbering.");
                        var actualFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                            WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                                document,
                                verifiedTarget,
                                owner.FormulaId));
                        if (!string.Equals(
                                actualFingerprint,
                                verifiedMetadata.NativeOmmlFingerprint,
                                StringComparison.OrdinalIgnoreCase))
                            throw new InvalidDataException(
                                $"Verified OMML owner {owner.FormulaId} changed content during numbering.");
                        owner.ReanchorAtVerifiedTargetCellTerminator(document, verifiedTarget);
                    }
                    finally { Release(verifiedTarget); }
                }
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-owners-reanchored-after-numbering count={userCellOwners.Count} source=verified-vtfco");

                // Keep the two table domains physically separate after numbering.
                // A formula that originated in a user table may be rebound only via
                // its now-repaired VTFCC structural owner. A body formula that
                // numbering moved into a VisualTeX-managed 1x3 host has no VTFCC
                // owner at all; its freshly recaptured VTFCO remains authoritative.
                var reboundTableOwners = liveOmmlInsertions.RebindStableTableOwners(
                    formulaId => TryResolveConvertedOmmlUserCellOwner(
                        document,
                        userCellOwners,
                        formulaId));
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-omml-table-owners-post-numbering-rebound count={reboundTableOwners} source=user-cell-structural-owner");
            }
            else if ((targetIsMathType || targetIsVisualTeX) && userCellOwners.Count > 0)
            {
                // VTFCC is intentionally only a cheap transaction-local point. Word
                // can pull that collapsed cell terminator onto an earlier body host
                // when a numbered VisualTeX caption/table scaffold is removed or
                // rebuilt. At this point every OLE target has already survived its
                // own insertion check. Recover from the target object itself, prove
                // it still occupies the original row/column + cell-local ordinal,
                // then rebuild VTFCC from the verified physical cell. MathType's
                // temporary VTMT locator can expand across later structural edits,
                // so shrink it back to the uniquely recovered Equation.DSMT4 too.
                foreach (var owner in userCellOwners)
                {
                    Range? verifiedTarget = null;
                    try
                    {
                        verifiedTarget = ResolveVerifiedConvertedOleUserCellTarget(
                            document,
                            plan.TargetMode,
                            owner);
                        owner.ReanchorAtVerifiedTargetCellTerminator(document, verifiedTarget);
                        if (targetIsMathType)
                            owner.RebindMathTypeBookmarkToVerifiedTarget(document, verifiedTarget);
                    }
                    finally { Release(verifiedTarget); }
                }
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-user-cell-owners-reanchored-after-numbering count={userCellOwners.Count} source={(targetIsMathType ? "verified-vtmt-cell-slot" : "embedded-visualtex-formula-id")}");
            }
            // Single-selection conversion can drag a non-target trailing OLE just
            // as a full batch does. Verify/rebind surviving physical identities on
            // every conversion boundary, including operations without row grouping.
            if (useWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
            {
                WordDoubleClickHook.TraceMessage(
                    "format-conversion-ole-identity-repair-skipped reason=all-visualtex-sources-converted");
            }
            else
            {
                RepairExpandedOleIdentityBookmarks(document);
            }
            if (targetIsOmml
                && result.FormulaCount > 0)
            {
                try
                {
                    var refreshed = WordOmmlNativeSource.RefreshFingerprintsFromDocumentOpenXml(
                        document,
                        convertedOmmlFormulaIds,
                        liveInsertions: liveOmmlInsertions,
                        finalizedMetadata: verifiedConvertedOmmlMetadata);
                    if (refreshed != convertedOmmlFormulaIds.Length)
                        throw new InvalidDataException(
                            $"OMML fingerprint finalization refreshed {refreshed}/{convertedOmmlFormulaIds.Length} converted formulas.");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-omml-fingerprints-finalized count={refreshed}");
                }
                catch (Exception fingerprintError)
                {
                    result.FailedFormulaCount++;
                    result.Failures.Add(
                        $"OMML fingerprint finalization failed: {fingerprintError.Message}");
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-omml-fingerprint-finalize-failed error={fingerprintError}");
                }
            }
            TraceFinalize("omml-fingerprint-refresh");
            if (targetIsOmml
                && string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_STAGE"),
                    "after-final-omml-fingerprint-refresh",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "Injected format-conversion failure after final OMML fingerprint refresh.");
            if (result.FailedFormulaCount > 0 || result.FormulaCount != plan.Targets.Count)
                throw new InvalidDataException(string.Join("\n", result.Failures));
            foreach (var owner in userCellOwners)
                VerifyConvertedUserCell(document, plan.TargetMode, owner, liveOmmlInsertions);
            foreach (var owner in userCellOwners)
                owner.DeleteTemporaryBookmark(document);
            IReadOnlyDictionary<string, string>? finalOmmlDisplayModes = null;
            var verifiedFinalOmmlTypes = false;
            if (targetIsOmml && convertedOmmlFormulaIds.Length > 0)
            {
                finalOmmlDisplayModes = successfulTargets.ToDictionary(
                    target => prepared[target.Id].Session.FormulaId,
                    target => target.DisplayMode,
                    StringComparer.OrdinalIgnoreCase);
                // RefreshFingerprintsFromDocumentOpenXml immediately above has
                // already proved the complete document XML, final fingerprints and
                // canonical anchors, and marked these exact live owners verified.
                // Before deleting VTFCO transport bookmarks, cheaply prove that Word
                // already retained every requested inline/display type. If any owner
                // is uncertain or any type differs, keep the mature full per-formula
                // normalization below as the conservative fallback.
                verifiedFinalOmmlTypes =
                    liveOmmlInsertions.VerifiedOwnersAlreadyHaveExpectedTypes(
                        finalOmmlDisplayModes);
            }
            // Table-owned fresh OMML uses temporary full-range bookmarks to survive
            // Word's OMath RCW rebinding across later cell mutations. Remove those
            // transaction-only owners before the successful custom record is closed;
            // VTOMML_* remains the sole durable formula identity in the document.
            liveOmmlInsertions.ClearTemporaryOwners();
            if (finalOmmlDisplayModes is not null)
            {
                if (verifiedFinalOmmlTypes)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-final-omml-type-normalization-skipped reason=verified-live-owners count={convertedOmmlFormulaIds.Length}");
                }
                else
                {
                    var normalizedFinalTypes = WordOmmlNativeSource.NormalizeFinalManagedEquationTypes(
                        document,
                        finalOmmlDisplayModes);
                    if (normalizedFinalTypes != convertedOmmlFormulaIds.Length)
                        throw new InvalidDataException(
                            $"Final OMML type normalization completed {normalizedFinalTypes}/{convertedOmmlFormulaIds.Length} formulas.");
                }
            }
            foreach (var name in retainedMathTypeTargetBookmarks.Values)
                DeleteConversionLocator(document, name);
            retainedMathTypeTargetBookmarks.Clear();
            conversionUndo.EndCustomRecord();
            conversionUndoEnded = true;
            if (useWholeDocumentVisualTeXToMathTypeCardinalityFastPath)
            {
                var committedCounts = ReadCanonicalOleObjectCounts(document);
                if (committedCounts.MathType != initialTargetObjectCount + plan.Targets.Count
                    || committedCounts.VisualTeX != initialSourceObjectCount - plan.Targets.Count)
                    throw new InvalidDataException("The source/target object counts changed when Word committed the conversion.");
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-canonical-ole-counts stage=commit visualtex={committedCounts.VisualTeX} mathtype={committedCounts.MathType}");
            }
            else if (CountSimpleFormatObjects(document, plan.TargetMode) != initialTargetObjectCount + plan.Targets.Count
                || CountSimpleFormatObjects(document, plan.SourceMode) != initialSourceObjectCount - plan.Targets.Count)
            {
                throw new InvalidDataException("The source/target object counts changed when Word committed the conversion.");
            }
            WordDoubleClickHook.TraceMessage($"format-conversion-transaction-completed targets={plan.Targets.Count}");
            return result;
        }
        catch (Exception conversionError)
        {
            WordDoubleClickHook.TraceMessage($"format-conversion-transaction-failed error={conversionError}");
            if (document is not null && conversionUndo is not null && rollbackSnapshot is not null)
            {
                try
                {
                    TraceConversionUndoState("failure-before-end");
                    if (!conversionUndoEnded)
                    {
                        conversionUndo.EndCustomRecord();
                        conversionUndoEnded = true;
                    }
                    TraceConversionUndoState("failure-after-end");
                    FinishBatchPresentation();
                    rollbackSnapshot.Restore(document);
                    WordDoubleClickHook.TraceMessage($"format-conversion-rollback-verified targets={plan.Targets.Count} error={conversionError}");
                }
                catch (Exception rollbackError)
                {
                    throw new AggregateException("Conversion failed and Word undo recovery could not be verified.", conversionError, rollbackError);
                }
            }
            throw;
        }
        finally
        {
            if (document is not null && retainedMathTypeTargetBookmarks.Count > 0)
            {
                foreach (var bookmarkName in retainedMathTypeTargetBookmarks.Values)
                {
                    try { TryDeleteBookmark(document, bookmarkName); } catch { }
                }
            }
            // The hidden OMML batch source can become Word's ActiveDocument even
            // though it is opened Visible=false. Closing it after activating the
            // real target leaves some Word builds with no ActiveDocument at all;
            // the next VisualTeX edit then fails before resolving its formula.
            // Dispose every hidden source first, then restore the user's document.
            ommlBatchSource?.Dispose();
            if (mathTypeVisualTeXCleanParagraphBuffer is not null)
            {
                try
                {
                    mathTypeVisualTeXCleanParagraphBuffer.Close(
                        WdSaveOptions.wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(mathTypeVisualTeXCleanParagraphBuffer);
            try { document?.Activate(); } catch { }
            if (paginationSuspended)
            {
                try { _application.Options.Pagination = previousPagination; } catch { }
            }
            if (batchViewSuspended && batchWindow is not null)
            {
                try { batchWindow.View.Type = previousViewType; } catch { }
            }
            Release(batchWindow);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            if (document is not null)
            {
                foreach (var owner in userCellOwners)
                {
                    try { owner.DeleteTemporaryBookmark(document); } catch { }
                }
            }
            foreach (var owner in userCellOwners) owner.Dispose();
            if (!conversionUndoEnded) EndUndoRecord(conversionUndo);
            Release(conversionUndo);
            Release(selection);
            Release(document);
        }
    }


    private static bool IsRollbackBridgeResidual(string? text, string bridge)
    {
        var normalizedText = NormalizeFormulaToLatexVerificationText(text ?? string.Empty).Trim();
        var normalizedBridge = NormalizeFormulaToLatexVerificationText(bridge).Trim();
        return normalizedText.Length > 0
            && string.Equals(normalizedText, normalizedBridge, StringComparison.Ordinal);
    }

    private static bool IsRollbackParagraphAdornment(string? text)
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

    private static string DescribeWordCharacters(string text) =>
        string.Join(",", (text ?? string.Empty).Take(64).Select(ch => $"U+{(int)ch:X4}"));

    private static (int VisualTeX, int MathType) ReadCanonicalOleObjectCounts(Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            var wordOpenXml = content.WordOpenXML ?? string.Empty;
            return (
                CountFlatOpcProgId(wordOpenXml, FormulaOleContract.ProgId),
                CountFlatOpcProgId(wordOpenXml, MathTypeOleInterop.CanonicalProgId));
        }
        finally { Release(content); }
    }

    private static int CountFlatOpcProgId(string wordOpenXml, string progId)
    {
        if (string.IsNullOrEmpty(wordOpenXml) || string.IsNullOrWhiteSpace(progId))
            return 0;
        // Word's generated Flat OPC serializes OLE ProgID as a double-quoted
        // attribute. Counting the exact canonical value avoids activating 1000
        // OLEFormat RCWs merely for a cardinality assertion. This fast path is used
        // only when capture proved there were no pre-existing MathType targets, so
        // registered legacy MathType aliases cannot be hidden by the canonical test.
        var needle = "ProgID=\"" + progId + "\"";
        var count = 0;
        var offset = 0;
        while (offset < wordOpenXml.Length)
        {
            var found = wordOpenXml.IndexOf(
                needle,
                offset,
                StringComparison.OrdinalIgnoreCase);
            if (found < 0) break;
            count++;
            offset = found + needle.Length;
        }
        return count;
    }

    private static int CountSimpleFormatObjects(Document document, string mode)
    {
        // Native math has its own COM collection. Counting it must not wake or
        // enumerate unrelated embedded OLE servers before returning OMaths.Count.
        if (string.Equals(mode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal))
        {
            OMaths? maths = null;
            try { maths = document.OMaths; return maths.Count; }
            finally { Release(maths); }
        }
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            var count = 0;
            shapes = document.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (string.Equals(mode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
                {
                    if (WordFormulaMetadataReader.IsNativeOle(shape)) count++;
                }
                else if (string.Equals(mode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
                         && MathTypeOleInterop.IsMathTypeOle(shape))
                {
                    count++;
                }
            }
            return count;
        }
        finally
        {
            Release(shape);
            Release(shapes);
        }
    }

    private static void EnsureBookmarkedMathTypeTargetSurvived(
        Document document,
        string bookmarkName,
        WordFormulaFormatConversionTarget target)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Exception? lastError = null;
        while (watch.ElapsedMilliseconds < 750)
        {
            Bookmark? bookmark = null;
            Range? bookmarkRange = null;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                if (!document.Bookmarks.Exists(bookmarkName))
                {
                    lastError = new InvalidOperationException(
                        "The temporary MathType identity bookmark disappeared after the formula transaction.");
                }
                else
                {
                    bookmark = document.Bookmarks[bookmarkName];
                    bookmarkRange = bookmark.Range;
                    shapes = bookmarkRange.InlineShapes;
                    var mathTypeCount = 0;
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        Release(shape);
                        shape = shapes[index];
                        if (MathTypeOleInterop.IsMathTypeOle(shape))
                            mathTypeCount++;
                    }
                    if (mathTypeCount == 1)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-bookmarked-target-stable formulaId={target.SourceFormulaId} bookmark={bookmarkName} range={bookmarkRange.Start}:{bookmarkRange.End} elapsedMs={watch.ElapsedMilliseconds}");
                        return;
                    }
                    lastError = new InvalidOperationException(
                        $"The temporary MathType identity bookmark contains {mathTypeCount} MathType objects instead of exactly one.");
                }
            }
            catch (Exception error)
            {
                lastError = error;
            }
            finally
            {
                Release(shape);
                Release(shapes);
                Release(bookmarkRange);
                Release(bookmark);
            }
            Thread.Sleep(25);
        }

        throw new InvalidOperationException(
            "Word did not retain the exact MathType replacement object after the formula transaction.",
            lastError);
    }

    private static void EnsureBookmarkedVisualTeXTargetSurvived(
        Document document,
        string formulaId,
        WordFormulaFormatConversionTarget target,
        int insertionStart)
    {
        var watch = System.Diagnostics.Stopwatch.StartNew();
        Exception? lastError = null;
        while (watch.ElapsedMilliseconds < 750)
        {
            InlineShape? shape = null;
            try
            {
                shape = FindByFormulaId(
                    document,
                    formulaId,
                    sourceObjectIdHint: null,
                    allowGlobalFallback: false);
                if (shape is null)
                {
                    // Word can collapse a bookmark boundary when two embedded OLE
                    // objects are directly adjacent and the left object is replaced
                    // after the right one. The freshly inserted VisualTeX OLE still
                    // owns the correct cache FormulaId, so recover only in a small
                    // neighborhood around the exact insertion point and immediately
                    // rebind the durable VTO_ bookmark. The session FormulaId is a
                    // newly generated GUID, therefore a unique native OLE carrying
                    // that exact cached id is sufficient proof that the target
                    // survived; do not let a second unstable bookmark COM read turn
                    // a successful adjacent replacement into a false rollback.
                    shape = FindFreshVisualTeXTargetNearPosition(
                        document,
                        insertionStart,
                        formulaId);
                    if (shape is not null)
                    {
                        BindOleIdentityBookmark(shape, formulaId);
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-bookmarked-visualtex-target-stable sourceFormulaId={target.SourceFormulaId} formulaId={formulaId} localAnchorRepair=True elapsedMs={watch.ElapsedMilliseconds}");
                        return;
                    }
                }
                if (shape is not null && WordFormulaMetadataReader.IsNativeOle(shape))
                {
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    if (metadata is not null
                        && string.Equals(metadata.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-bookmarked-visualtex-target-stable sourceFormulaId={target.SourceFormulaId} formulaId={formulaId} localAnchorRepair=False elapsedMs={watch.ElapsedMilliseconds}");
                        return;
                    }
                }
                lastError = new InvalidOperationException(
                    "The VisualTeX identity bookmark did not resolve to the exact converted OLE object.");
            }
            catch (Exception error)
            {
                lastError = error;
            }
            finally { Release(shape); }
            Thread.Sleep(25);
        }
        throw new InvalidOperationException(
            "Word did not retain the exact VisualTeX replacement object after the formula transaction.",
            lastError);
    }

    private static void EnsureConvertedVisualTeXNumberingHostHealthy(
        Document document,
        string formulaId,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Range? ownerRange = null;
        Range? visibleRange = null;
        InlineShapes? ownerShapes = null;
        ParagraphFormat? ownerFormat = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        Fields? visibleFields = null;
        Field? visibleField = null;
        Range? visibleCode = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Range? visibleBookmarkRange = null;
        Range? captionRange = null;
        Range? numberRange = null;
        Paragraphs? captionParagraphs = null;
        Frames? captionFrames = null;
        Fields? captionFields = null;
        Field? captionField = null;
        Range? captionCode = null;
        try
        {
            shape = FindByFormulaId(
                    document,
                    formulaId,
                    sourceObjectIdHint: null,
                    allowGlobalFallback: false)
                ?? throw new InvalidOperationException(
                    "The converted VisualTeX OLE disappeared before numbering validation.");
            if (!WordFormulaMetadataReader.IsNativeOle(shape))
                throw new InvalidOperationException(
                    "The converted numbering owner is not a VisualTeX native OLE.");
            var metadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidOperationException(
                    "The converted VisualTeX OLE has no readable metadata.");
            if (!string.Equals(
                    metadata.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase)
                || !metadata.Numbered
                || !string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase)
                || !target.Numbered
                || !string.Equals(
                    target.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    "The converted VisualTeX OLE lost its numbered display identity.");

            shapeRange = shape.Range;
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                    document,
                    formulaId)
                ?? throw new InvalidOperationException(
                    "The converted VisualTeX OLE has no numbering owner paragraph.");
            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(
                    document,
                    formulaId)
                ?? throw new InvalidOperationException(
                    "The converted VisualTeX OLE has no visible equation number.");
            if ((bool)ownerRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    "The converted VisualTeX OLE numbering owner unexpectedly remained inside a table.");
            if (shapeRange.Start < ownerRange.Start
                || shapeRange.End > ownerRange.End
                || visibleRange.Start < shapeRange.End
                || visibleRange.End > ownerRange.End)
                throw new InvalidOperationException(
                    "The converted VisualTeX OLE and its visible number do not share one safe owner paragraph.");

            ownerShapes = ownerRange.InlineShapes;
            if (ownerShapes.Count != 1)
                throw new InvalidOperationException(
                    $"The converted VisualTeX numbering paragraph owns {ownerShapes.Count} OLE objects instead of exactly one.");
            var ownerText = ownerRange.Text ?? string.Empty;
            if (ownerText.Count(character => character == '\t') < 2)
                throw new InvalidOperationException(
                    "The converted VisualTeX numbering paragraph lost one of its two layout TAB characters.");

            ownerFormat = ownerRange.ParagraphFormat;
            tabStops = ownerFormat.TabStops;
            var hasCenterTab = false;
            var hasRightTab = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                // Word does not report TabStop.CustomTab consistently for direct
                // center/right stops on a freshly rebuilt conversion paragraph.
                // The long-standing production acceptance therefore identifies the
                // two required stops by alignment and position, not CustomTab.
                if (tabStop.Alignment == WdTabAlignment.wdAlignTabCenter)
                    hasCenterTab = true;
                else if (tabStop.Alignment == WdTabAlignment.wdAlignTabRight)
                    hasRightTab = true;
            }
            if (!hasCenterTab || !hasRightTab)
                throw new InvalidOperationException(
                    $"The converted VisualTeX numbering paragraph lost its direct tab stops (center={hasCenterTab}, right={hasRightTab}).");

            var expectedNumberBookmark =
                WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            visibleFields = visibleRange.Fields;
            var hasVisibleReference = false;
            for (var index = 1; index <= visibleFields.Count; index++)
            {
                Release(visibleCode);
                visibleCode = null;
                Release(visibleField);
                visibleField = visibleFields[index];
                visibleCode = visibleField.Code;
                var code = visibleCode.Text ?? string.Empty;
                if (code.IndexOf(
                        "REF " + expectedNumberBookmark,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    hasVisibleReference = true;
                    break;
                }
            }
            if (!hasVisibleReference)
                throw new InvalidOperationException(
                    "The converted VisualTeX visible number is not a REF to its own VTEqNum bookmark.");

            bookmarks = document.Bookmarks;
            var visibleName = WordEquationNumbering.EquationBookmarkName(formulaId);
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)
                || !bookmarks.Exists(captionName)
                || !bookmarks.Exists(expectedNumberBookmark))
                throw new InvalidOperationException(
                    "The converted VisualTeX numbering scaffold is missing VTEq/VTEqCap/VTEqNum identity.");
            visibleBookmark = bookmarks[visibleName];
            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[expectedNumberBookmark];
            visibleBookmarkRange = visibleBookmark.Range;
            captionRange = captionBookmark.Range;
            numberRange = numberBookmark.Range;
            if (visibleBookmarkRange.Start < shapeRange.End
                || visibleBookmarkRange.End > ownerRange.End)
                throw new InvalidOperationException(
                    "The converted VTEq bookmark crosses the VisualTeX OLE boundary.");
            if (captionRange.Start < ownerRange.End
                || numberRange.Start < captionRange.Start
                || numberRange.End > captionRange.End)
                throw new InvalidOperationException(
                    "The converted hidden caption bookmarks overlap the formula paragraph.");

            captionParagraphs = captionRange.Paragraphs;
            if (captionParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The converted hidden SEQ caption spans more than one paragraph.");
            captionFrames = captionRange.Frames;
            if (captionFrames.Count != 1)
                throw new InvalidOperationException(
                    "The converted hidden SEQ caption is not isolated in its clipping frame.");
            captionFields = captionRange.Fields;
            var sequenceCount = 0;
            for (var index = 1; index <= captionFields.Count; index++)
            {
                Release(captionCode);
                captionCode = null;
                Release(captionField);
                captionField = captionFields[index];
                captionCode = captionField.Code;
                if ((captionCode.Text ?? string.Empty).IndexOf(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    sequenceCount++;
            }
            if (sequenceCount != 1)
                throw new InvalidOperationException(
                    $"The converted VisualTeX hidden caption owns {sequenceCount} VisualTeX SEQ fields instead of exactly one.");

            WordDoubleClickHook.TraceMessage(
                $"format-conversion-visualtex-single-numbering-healthy formulaId={formulaId} owner={ownerRange.Start}:{ownerRange.End} caption={captionRange.Start}:{captionRange.End}");
        }
        finally
        {
            Release(captionCode);
            Release(captionField);
            Release(captionFields);
            Release(captionFrames);
            Release(captionParagraphs);
            Release(numberRange);
            Release(captionRange);
            Release(visibleBookmarkRange);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(visibleCode);
            Release(visibleField);
            Release(visibleFields);
            Release(tabStop);
            Release(tabStops);
            Release(ownerFormat);
            Release(ownerShapes);
            Release(visibleRange);
            Release(ownerRange);
            Release(shapeRange);
            Release(shape);
        }
    }

    private static InlineShape? FindFreshVisualTeXTargetNearPosition(
        Document document,
        int insertionStart,
        string formulaId)
    {
        Range? content = null;
        Range? localRange = null;
        InlineShapes? shapes = null;
        InlineShape? result = null;
        try
        {
            content = document.Content;
            var start = Math.Max(content.Start, insertionStart - 2);
            var end = Math.Min(content.End, insertionStart + 64);
            if (end <= start) return null;
            localRange = document.Range(start, end);
            shapes = localRange.InlineShapes;
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-local-vt-probe formulaId={formulaId} insertionStart={insertionStart} range={start}:{end} shapes={shapes.Count}");
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var cached = WordFormulaMetadataReader.TryReadCachedPreview(candidate);
                    Range? candidateRange = null;
                    try
                    {
                        candidateRange = candidate.Range;
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-local-vt-candidate formulaId={formulaId} index={index}/{shapes.Count} range={candidateRange.Start}:{candidateRange.End} cachedFormulaId={cached?.FormulaId ?? "<null>"}");
                    }
                    finally { Release(candidateRange); }
                    if (!string.Equals(
                            cached?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    if (result is not null)
                        throw new InvalidOperationException(
                            "More than one fresh VisualTeX target near the insertion point carries the same FormulaId.");
                    result = candidate;
                    candidate = null;
                }
                finally { Release(candidate); }
            }
            return result;
        }
        catch
        {
            Release(result);
            throw;
        }
        finally
        {
            Release(shapes);
            Release(localRange);
            Release(content);
        }
    }

    private static void WaitForSimpleTargetObjectCountToStabilize(
        Document document,
        string targetMode,
        int expectedCount,
        WordFormulaFormatConversionTarget target)
    {
        const int requiredStableSamples = 2;
        var stableSamples = 0;
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var lastCount = -1;
        while (watch.ElapsedMilliseconds < 1200)
        {
            lastCount = CountSimpleFormatObjects(document, targetMode);
            if (lastCount == expectedCount)
            {
                stableSamples++;
                if (stableSamples >= requiredStableSamples)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-target-stable formulaId={target.SourceFormulaId} targetMode={targetMode} count={lastCount} elapsedMs={watch.ElapsedMilliseconds}");
                    return;
                }
            }
            else
            {
                stableSamples = 0;
            }
            Thread.Sleep(25);
        }

        throw new InvalidOperationException(
            $"Word did not stabilize the converted target formula after the formula transaction. Expected target count {expectedCount}, actual {lastCount}.");
    }

    private static void TraceSimpleFormatObjectCounts(
        Document document,
        WordFormulaFormatConversionTarget target,
        string stage)
    {
        try
        {
            var visualTeXCount = CountSimpleFormatObjects(
                document,
                FormulaOleContract.NativeOleMode);
            var mathTypeCount = CountSimpleFormatObjects(
                document,
                FormulaOleContract.MathTypeOleMode);
            var ommlCount = CountSimpleFormatObjects(
                document,
                FormulaOleContract.WordOmmlMode);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-counts stage={stage} formulaId={target.SourceFormulaId} VT={visualTeXCount} MT={mathTypeCount} OMML={ommlCount} latex={target.Latex}");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-counts-failed stage={stage} formulaId={target.SourceFormulaId} error={error.Message}");
        }
    }

    private bool IsSimpleFormatSourcePresent(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                // Batch conversion processes targets from the end of the document
                // toward the start, so the captured source range remains a strong
                // local hint. Resolve the durable VTO_ bookmark first and then a
                // bounded neighborhood around that hint; never enumerate every OLE
                // just to prove one already-deleted source is absent.
                shape = FindByFormulaId(
                    document,
                    target.SourceFormulaId,
                    target.SourceObjectId,
                    allowGlobalFallback: false);
                if (shape is null) return false;
                Range? liveRange = null;
                try
                {
                    liveRange = shape.Range;
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-source-match formulaId={target.SourceFormulaId} liveRange={liveRange.Start}:{liveRange.End} sourceHint={target.SourceObjectId} latex={target.Latex}");
                }
                catch { }
                finally { Release(liveRange); }
                return true;
            }
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                Range? ommlRange = null;
                try
                {
                    ommlRange = ResolveSimpleOmmlSourceRange(document, target);
                    return ommlRange is not null;
                }
                finally { Release(ommlRange); }
            }
            shape = FindMathTypeOleByRange(
                document,
                target.SourceObjectId,
                allowGlobalFallback: false);
            return shape is not null && MathTypeOleInterop.IsMathTypeOle(shape);
        }
        catch { return false; }
        finally
        {
            Release(shape);
            Release(shapes);
        }
    }

    private static Range? ResolveSimpleOmmlSourceRange(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Bookmark? bookmark = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(
                document,
                target.SourceFormulaId);
            if (bookmark is not null)
                return WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document, target.SourceFormulaId, target.Metadata);
            if (target.SourceIsManagedOmml)
                return null;
            return TryResolveOmmlRangeReference(document, target.SourceObjectId);
        }
        finally { Release(bookmark); }
    }

    private string CreateMathTypeSourceIdentityBookmark(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? range = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            shape = FindMathTypeOleByRange(
                    document,
                    target.SourceObjectId,
                    allowGlobalFallback: false)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before identity binding.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "The source object is no longer a MathType Equation.DSMT4 formula.");
            range = shape.Range.Duplicate;
            var bookmarkName = "VTFCI_" + target.Id.Replace("-", string.Empty);
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(bookmarkName))
            {
                bookmark = bookmarks[bookmarkName];
                bookmark.Delete();
                Release(bookmark);
                bookmark = null;
            }
            bookmark = bookmarks.Add(bookmarkName, range);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-mathtype-source-identity-bound formulaId={target.SourceFormulaId} bookmark={bookmarkName} range={range.Start}:{range.End}");
            return bookmarkName;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(range);
            Release(shape);
        }
    }

    private static void EnsureForwardMathTypeSourceRemoved(
        Document document,
        string bookmarkName,
        WordFormulaFormatConversionTarget target,
        string stage)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? shapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-forward-source-check stage={stage} formulaId={target.SourceFormulaId} bookmarkMissing=True present=False");
                return;
            }

            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            shapes = range.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate)) continue;
                    throw new InvalidOperationException(
                        $"Word still contains the MathType source formula after stage '{stage}'. The current formula conversion was rolled back.");
                }
                finally { Release(candidate); }
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-forward-source-check stage={stage} formulaId={target.SourceFormulaId} range={range.Start}:{range.End} present=False");
        }
        finally
        {
            Release(shapes);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static InlineShape? FindVisualTeXSourceByIdentityBookmark(
        Document document,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        InlineShapes? shapes = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = WordFormulaMetadataReader.IdentityBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return null;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            shapes = range.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                try
                {
                    candidate = shapes[index];
                    if (!WordFormulaMetadataReader.IsNativeOle(candidate)) continue;
                    var metadata = WordFormulaMetadataReader.TryRead(candidate);
                    if (!string.Equals(
                            metadata?.FormulaId,
                            formulaId,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var result = candidate;
                    candidate = null;
                    return result;
                }
                finally { Release(candidate); }
            }
            return null;
        }
        finally
        {
            Release(shapes);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private void EnsureSimpleFormatSourceRemoved(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target,
        string stage)
    {
        var present = IsSimpleFormatSourcePresent(document, sourceMode, target);
        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-source-check stage={stage} formulaId={target.SourceFormulaId} present={present} latex={target.Latex}");
        }
        if (present)
            throw new InvalidOperationException(
                $"Word still contains the source formula after stage '{stage}'. The current formula conversion was rolled back.");
    }

    private static void ThrowIfSimpleFormatConversionFailureInjected(
        WordFormulaFormatConversionTarget target)
    {
        if (!string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var requested = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
        if (string.IsNullOrWhiteSpace(requested)) return;
        var matches = string.Equals(requested, "1", StringComparison.Ordinal)
            || string.Equals(requested, target.SourceFormulaId, StringComparison.OrdinalIgnoreCase)
            || string.Equals(requested, "numbered", StringComparison.OrdinalIgnoreCase)
                && target.Numbered;
        if (matches)
            throw new InvalidOperationException(
                "Injected format-conversion failure after deleting the source host.");
    }

    private static void ValidateSimpleFormatConversionPair(
        string sourceMode,
        string targetMode)
    {
        var visualTeXToMathType =
            string.Equals(sourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal);
        var mathTypeToVisualTeX =
            string.Equals(sourceMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal);
        var ommlToMathType =
            string.Equals(sourceMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal);
        var mathTypeToOmml =
            string.Equals(sourceMode, FormulaOleContract.MathTypeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal);
        var visualTeXToOmml =
            string.Equals(sourceMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal);
        var ommlToVisualTeX =
            string.Equals(sourceMode, FormulaOleContract.WordOmmlMode, StringComparison.Ordinal)
            && string.Equals(targetMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal);
        if (!visualTeXToMathType
            && !mathTypeToVisualTeX
            && !ommlToMathType
            && !mathTypeToOmml
            && !visualTeXToOmml
            && !ommlToVisualTeX)
            throw new ArgumentOutOfRangeException(
                nameof(targetMode),
                $"Unsupported simple format conversion: {sourceMode} -> {targetMode}.");
    }

    private static void ValidatePreparedFormatConversionTarget(
        string targetMode,
        WordFormulaFormatConversionTarget target,
        PreparedWordBulkFormula formula)
    {
        if (string.IsNullOrWhiteSpace(formula.MathMl))
            throw new InvalidDataException(
                $"Formula '{target.Latex}' did not produce MathML.");
        if (string.Equals(
                targetMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
            return;
        if (string.Equals(
                targetMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
            && formula.MathTypeNativePreviewAttempted
            && formula.MathTypeNativePreview is not null)
            return;
        if (string.IsNullOrWhiteSpace(formula.EmfPath)
            || !File.Exists(formula.EmfPath))
            throw new FileNotFoundException(
                $"Formula '{target.Latex}' did not produce an EMF preview.",
                formula.EmfPath);
        if (string.Equals(
                targetMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            && (string.IsNullOrWhiteSpace(formula.PngPath)
                || !File.Exists(formula.PngPath)))
            throw new FileNotFoundException(
                $"Formula '{target.Latex}' did not produce a PNG preview.",
                formula.PngPath);
    }

    private void ValidateSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Table? table = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId,
                        allowGlobalFallback: false)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula moved before conversion started.");
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                    throw new InvalidOperationException(
                        "The source object is no longer a VisualTeX OLE formula.");
                shapeRange = shape.Range;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        target.SourceFormulaId);
                    if (table is not null)
                    {
                        if (!IsSafeVisualTeXNumberingTableForConversion(table, shapeRange))
                            throw new InvalidOperationException(
                                "The legacy VisualTeX numbering table contains nonempty extra rows or another formula object; conversion was refused before modifying the document.");
                    }
                    else if (!IsSafeVisualTeXNumberingParagraphForConversion(
                                 document,
                                 target.SourceFormulaId,
                                 shapeRange))
                    {
                        throw new InvalidOperationException(
                            "The numbered VisualTeX source no longer owns one safe MathType-style tab paragraph.");
                    }
                }
                return;
            }

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                shapeRange = ResolveSimpleOmmlSourceRange(document, target)
                    ?? throw new InvalidOperationException(
                        "The Word OMML source formula moved before conversion started.");
                OMaths? maths = null;
                try
                {
                    maths = shapeRange.OMaths;
                    if (maths.Count != 1)
                        throw new InvalidOperationException(
                            "The source range no longer contains exactly one Word OMath equation.");
                }
                finally { Release(maths); }
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata);
                    if (table is not null)
                    {
                        var currentDirectTable = WordEquationNumbering
                            .HasReusableNumberedNativeOmmlDirectTableHost(
                                document,
                                shapeRange,
                                target.SourceFormulaId);
                        if (!currentDirectTable
                            && !IsSafeOmmlNumberingTableForConversion(table, shapeRange))
                            throw new InvalidOperationException(
                                "The legacy OMML numbering table contains ordinary user content or another formula; conversion was refused before modifying the document.");
                    }
                    else if (!IsSafeOmmlNumberingParagraphForConversion(
                                 document,
                                 target.SourceFormulaId,
                                 shapeRange))
                    {
                        throw new InvalidOperationException(
                            "The numbered OMML source no longer owns one safe center/right-tab paragraph.");
                    }
                }
                return;
            }

            shape = FindMathTypeOleByRange(
                    document,
                    target.SourceObjectId,
                    allowGlobalFallback: false)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before conversion started.");
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "The source object is no longer Equation.DSMT4.");
            if (!string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                return;
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display formula no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.InlineShapes.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display paragraph contains another inline object; conversion was refused.");
            target.SourceHasMathTypeSectionPrefix = false;
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
            {
                if (!IsSafeMathTypeRowWithSectionPrefix(document, paragraphRange, shapeRange))
                    throw new InvalidOperationException(
                        "The MathType display paragraph contains ordinary user text; conversion was refused to avoid deleting prose.");
                target.SourceHasMathTypeSectionPrefix = true;
            }
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(table);
            Release(shape);
        }
    }

    private static bool HasLocalVisualTeXOmmlAnchor(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Range? content = null;
        Range? probe = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            content = document.Content;
            var start = Math.Max(content.Start, target.SourceStart - 3);
            var end = Math.Min(
                content.End,
                Math.Max(target.SourceStart + 3, target.SourceStart + 1));
            probe = document.Range(start, end);
            bookmarks = probe.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (name.StartsWith(
                        WordOmmlFormulaStore.BookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
        catch
        {
            // Direct deletion is only a performance optimization. If Word refuses
            // a local bookmark safety probe, fall back to the mature managed-OMML
            // replacement path rather than risk orphaning VisualTeX metadata.
            return true;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(probe);
            Release(content);
        }
    }

    private static void RepairManagedOmmlIdentityAfterDirectRollback(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Bookmark? existing = null;
        Range? candidate = null;
        Bookmark? repaired = null;
        try
        {
            existing = WordOmmlFormulaStore.FindByFormulaId(
                document,
                target.SourceFormulaId);
            if (existing is not null) return;

            candidate = TryResolveOmmlRangeReference(
                document,
                target.SourceObjectId)
                ?? throw new InvalidOperationException(
                    "Word restored the managed OMML after rollback, but its original range could not be resolved for VTOMML identity repair.");
            OMaths? maths = null;
            OMath? math = null;
            Range? exactRange = null;
            try
            {
                maths = candidate.OMaths;
                if (maths.Count != 1)
                    throw new InvalidOperationException(
                        "Word rollback restored an ambiguous OMML range while repairing VTOMML identity.");
                math = maths[1];
                if (math.Type != WdOMathType.wdOMathDisplay)
                    throw new InvalidOperationException(
                        "Word rollback restored the managed block OMML with the wrong OMath type.");
                exactRange = math.Range.Duplicate;
                Release(candidate);
                candidate = exactRange;
                exactRange = null;
            }
            finally
            {
                Release(exactRange);
                Release(math);
                Release(maths);
            }

            repaired = WordOmmlFormulaStore.Wrap(
                document,
                candidate,
                target.Metadata,
                replaceExisting: true);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-managed-omml-rollback-id-repaired formulaId={target.SourceFormulaId} range={candidate.Start}:{candidate.End}");
        }
        finally
        {
            Release(repaired);
            Release(candidate);
            Release(existing);
        }
    }

    private int DeleteUnnumberedOmmlSourceExactly(
        Document document, WordFormulaFormatConversionTarget target)
    {
        if (target.Numbered)
            throw new InvalidDataException("Exact bare OMML replacement cannot delete a numbered host.");
        Range? source = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? actual = null;
        try
        {
            source = ResolveSimpleOmmlSourceRange(document, target)
                ?? throw new InvalidDataException("The OMML source lost its verified identity.");
            maths = source.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException("Exact OMML replacement requires one physical equation.");
            math = maths[1]; actual = math.Range;
            if (actual.StoryType != source.StoryType || actual.Start != source.Start || actual.End != source.End)
                throw new InvalidDataException("The OMML replacement range includes content outside its equation.");
            var start = actual.Start;
            if (target.SourceIsManagedOmml)
                TryDeleteBookmark(document, WordOmmlFormulaStore.BookmarkName(target.SourceFormulaId));
            // Range.Delete invokes Word's smart whitespace joining. Replacing the
            // exact equation span preserves the original adjacent prose and avoids
            // the former two-step OMML -> LaTeX -> empty bridge entirely.
            actual.Text = string.Empty;
            return start;
        }
        finally { Release(actual); Release(math); Release(maths); Release(source); }
    }

    private Range DeleteSingleNativeOmmlSourceRange(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        var sourceRange = ResolveSimpleOmmlSourceRange(document, target)
            ?? throw new InvalidOperationException(
                "The native Word OMML source moved before direct replacement.");
        if (target.SourceIsManagedOmml || target.Numbered)
        {
            Release(sourceRange);
            throw new InvalidOperationException(
                "Direct OMML deletion is restricted to one unnumbered native Word equation.");
        }
        return sourceRange;
    }

    private int DeleteSingleNativeOmmlSourceDirect(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Range? sourceRange = null;
        try
        {
            sourceRange = DeleteSingleNativeOmmlSourceRange(document, target);
            var start = sourceRange.Start;
            sourceRange.Delete();
            return start;
        }
        finally { Release(sourceRange); }
    }

    private int DeleteSingleManagedDisplayOmmlSourceDirect(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        if (!target.SourceIsManagedOmml
            || target.Numbered
            || !string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The managed display OMML fast path requires one unnumbered VisualTeX display OMath.");

        Range? sourceRange = null;
        OMaths? maths = null;
        OMath? math = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        try
        {
            sourceRange = ResolveSimpleOmmlSourceRange(document, target)
                ?? throw new InvalidOperationException(
                    "The managed display OMML source moved before direct replacement.");
            if ((bool)sourceRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    "The managed unnumbered display OMML unexpectedly belongs to a table.");
            maths = sourceRange.OMaths;
            if (maths.Count != 1)
                throw new InvalidOperationException(
                    "The managed display OMML source no longer contains exactly one OMath.");
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
                throw new InvalidOperationException(
                    "The managed OMML direct replacement source is no longer Word display math.");

            paragraphs = sourceRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The managed display OMML source spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            prefix = document.Range(paragraphRange.Start, sourceRange.Start);
            suffix = document.Range(
                sourceRange.End,
                Math.Max(sourceRange.End, paragraphRange.End - 1));
            if (!IsRollbackParagraphAdornment(prefix.Text)
                || !IsRollbackParagraphAdornment(suffix.Text))
                throw new InvalidOperationException(
                    "The managed display OMML paragraph contains ordinary user text and cannot use direct replacement.");

            var start = sourceRange.Start;
            // The VTOMML identity belongs to the source host. Keep its CustomXML
            // metadata until the MathType target has survived the Word transaction,
            // but remove the bookmark and the native OMath inside the same Undo
            // record. Word can then restore both atomically if insertion fails.
            TryDeleteBookmark(
                document,
                WordOmmlFormulaStore.BookmarkName(target.SourceFormulaId));
            sourceRange.Delete();
            return start;
        }
        finally
        {
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(math);
            Release(maths);
            Release(sourceRange);
        }
    }

    private bool HasLegacyNumberedOmmlTable(
        Document document,
        WordFormulaFormatConversionTarget target)
    {
        Range? sourceRange = null;
        Table? table = null;
        try
        {
            sourceRange = ResolveSimpleOmmlSourceRange(document, target);
            if (sourceRange is null) return false;
            if (WordEquationNumbering.HasManagedNativeOmmlHashSequenceHost(
                    document,
                    target.SourceFormulaId))
                return true;
            table = TryGetVisualTeXNumberedTable(sourceRange, target.Metadata);
            return table is not null;
        }
        catch
        {
            // This method only selects an optional fast path. A failed probe must
            // fall back to the mature managed replacement path, never fail the
            // conversion before its normal source validation has run.
            return false;
        }
        finally
        {
            Release(table);
            Release(sourceRange);
        }
    }

    private int DeleteSingleManagedNumberedOmmlSourceDirect(
        Document document,
        WordFormulaFormatConversionTarget target,
        IReadOnlyDictionary<string, int>? knownReferenceCounts,
        bool preserveCrossReferences,
        ISet<string> deferredOwnedSeparators)
    {
        if (!target.SourceIsManagedOmml
            || !target.Numbered
            || !string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The managed numbered OMML fast path requires one numbered VisualTeX display OMath.");

        Range? sourceRange = null;
        Table? table = null;
        Table? splitTable = null;
        Range? ownerRange = null;
        Range? convertedTableRange = null;
        Range? clearedParagraphProbe = null;
        Paragraphs? clearedParagraphs = null;
        Paragraph? clearedParagraph = null;
        Range? clearedParagraphRange = null;
        var numberedNativeHashSequence = false;
        var currentDirectTable = false;
        try
        {
            var bodyFormatting = WordEquationNumbering.CaptureNumberedBodyCharacterFormatting(
                document, target.SourceFormulaId);
            sourceRange = ResolveSimpleOmmlSourceRange(document, target)
                ?? throw new InvalidOperationException(
                    "The managed numbered OMML source moved before direct replacement.");
            using var paragraphFormatting = WordParagraphFormatting.Capture(sourceRange);
            table = TryGetVisualTeXNumberedTable(sourceRange, target.Metadata);
            if (table is not null)
            {
                currentDirectTable = WordEquationNumbering
                    .HasReusableNumberedNativeOmmlDirectTableHost(
                        document,
                        sourceRange,
                        target.SourceFormulaId);
                if (!currentDirectTable
                    && !IsSafeOmmlNumberingTableForConversion(table, sourceRange))
                    throw new InvalidOperationException(
                        "The legacy managed OMML table contains user content or another formula and cannot use the direct replacement path.");

                // Direct bulk OMML insertion compacts adjacent numbered equations
                // into one managed N x 3 table so Word has no undeletable separator
                // paragraphs. A format replacement must dismantle only the row that
                // owns this formula. Split before and after that row, then re-resolve
                // every COM range through the durable VTOMML identity.
                if (table.Rows.Count > 1)
                {
                    if (!WordEquationNumbering.TryGetManagedNumberTableRowIndex(
                            table,
                            sourceRange,
                            expectedColumnIndex: 2,
                            out var rowIndex))
                        throw new InvalidOperationException(
                            "The managed numbered OMML source row could not be resolved before replacement.");

                    if (rowIndex > 1)
                    {
                        splitTable = table.Split(rowIndex);
                        Release(splitTable);
                        splitTable = null;
                        Release(sourceRange);
                        sourceRange = ResolveSimpleOmmlSourceRange(document, target)
                            ?? throw new InvalidOperationException(
                                "The managed numbered OMML source moved while isolating its table row.");
                        Release(table);
                        table = TryGetVisualTeXNumberedTable(sourceRange, target.Metadata)
                            ?? throw new InvalidOperationException(
                                "The managed numbered OMML table disappeared while isolating its row.");
                    }

                    if (table.Rows.Count > 1)
                    {
                        splitTable = table.Split(2);
                        Release(splitTable);
                        splitTable = null;
                        Release(sourceRange);
                        sourceRange = ResolveSimpleOmmlSourceRange(document, target)
                            ?? throw new InvalidOperationException(
                                "The managed numbered OMML source moved while detaching its following rows.");
                        Release(table);
                        table = TryGetVisualTeXNumberedTable(sourceRange, target.Metadata)
                            ?? throw new InvalidOperationException(
                                "The isolated managed numbered OMML row lost its table host.");
                    }

                    if (table.Rows.Count != 1
                        || !WordEquationNumbering.HasReusableNumberedNativeOmmlDirectTableHost(
                            document,
                            sourceRange,
                            target.SourceFormulaId))
                        throw new InvalidOperationException(
                            "The managed numbered OMML row was not a healthy isolated 1x3 host.");
                }
                if (currentDirectTable) ownerRange = table.Range;
            }
            else
            {
                numberedNativeHashSequence =
                    WordEquationNumbering
                        .HasReusableNumberedNativeOmmlHashSequenceHost(
                            document,
                            sourceRange,
                            target.SourceFormulaId);
                if (!numberedNativeHashSequence
                    && !IsSafeOmmlNumberingParagraphForConversion(
                        document,
                        target.SourceFormulaId,
                        sourceRange))
                    throw new InvalidOperationException(
                        "The managed numbered OMML source is neither a healthy native #(SEQ) display nor a safe legacy tab paragraph.");
                ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                        document,
                        target.SourceFormulaId)
                    ?? throw new InvalidOperationException(
                        "The managed numbered OMML source lost its paragraph owner before direct replacement.");
            }

            if (!preserveCrossReferences)
            {
                WordEquationNumbering.FreezeFormulaCrossReferences(
                    document,
                    target.SourceFormulaId,
                    knownReferenceCounts);
            }
            if (numberedNativeHashSequence)
            {
                WordEquationNumbering.RemoveNativeOmmlHashSequenceAliasesForReplacement(
                    document,
                    target.SourceFormulaId);
            }
            else
            {
                WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                    document,
                    target.SourceFormulaId,
                    replacingCompleteHost: currentDirectTable ? ownerRange : null);
            }
            TryDeleteBookmark(
                document,
                WordOmmlFormulaStore.BookmarkName(target.SourceFormulaId));
            WordOmmlFormulaStore.Delete(document, target.SourceFormulaId);

            if (table is not null)
            {
                // Compatibility path for documents created by older VisualTeX
                // versions. Dismantle the managed 1x3 host and leave one ordinary
                // paragraph at the exact source position for the MathType target.
                var reuseOwnedSeparator = currentDirectTable
                    && WordEquationNumbering.HasExactOwnedNativeOmmlTableSeparator(
                        document, table, target.SourceFormulaId);
                object separator = WdTableFieldSeparator.wdSeparateByParagraphs;
                object nestedTables = false;
                convertedTableRange = table.ConvertToText(
                    ref separator,
                    ref nestedTables);
                var tableStart = convertedTableRange.Start;
                // Keep a dedicated BODY insertion paragraph until the replacement
                // OLE is materialized. Reusing the following table separator here
                // leaves Word's insertion range with stale cell affinity. The same
                // deferred lifetime used by caption tails removes the extra mark
                // only after a real target has proven its physical body owner.
                convertedTableRange.Text = "\r";
                var formattingRange = convertedTableRange;
                if (reuseOwnedSeparator)
                {
                    deferredOwnedSeparators.Add(target.SourceFormulaId);
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-owned-separator-deferred formulaId={target.SourceFormulaId} start={tableStart}");
                }
                else
                {
                    // Edited/unmarked user content is never reclaimed.
                    TryDeleteBookmark(document,
                        WordEquationNumbering.NativeOmmlOwnedSeparatorBookmarkName(target.SourceFormulaId));
                }
                // Word can give this new paragraph the style of the following
                // heading when the final table row is removed. Restore the source
                // owner's style as well as its body font before inserting a target;
                // otherwise an equation becomes an extra chapter in SEQ numbering.
                paragraphFormatting.Apply(formattingRange);
                bodyFormatting.ApplyToParagraphMark(formattingRange);
                return tableStart;
            }

            // A current #(SEQ) equation and the older tab host each own one proven
            // safe paragraph with no prose or second object. Replacing that complete
            // owner with one paragraph mark deletes the mathematical field atomically
            // and leaves an exact insertion host for the MathType target.
            var paragraphStart = ownerRange!.Start;
            if (target.SourceWithinTable)
            {
                // Preserve the cell's terminal \r\a. Replacing the complete owner
                // with a plain \r would consume the cell end and can merge/damage
                // the user's table. Clear only editable paragraph content, then
                // re-resolve the surviving cell paragraph for formatting repair.
                SetRangeToEditableParagraphBody(ownerRange);
                ownerRange.Text = string.Empty;
                clearedParagraphProbe = document.Range(paragraphStart, paragraphStart);
                clearedParagraphs = clearedParagraphProbe.Paragraphs;
                if (clearedParagraphs.Count != 1)
                    throw new InvalidDataException(
                        "The table cell paragraph disappeared while replacing numbered OMML.");
                clearedParagraph = clearedParagraphs[1];
                clearedParagraphRange = clearedParagraph.Range.Duplicate;
                paragraphFormatting.Apply(clearedParagraphRange);
                bodyFormatting.ApplyToParagraphMark(clearedParagraphRange);
            }
            else
            {
                ownerRange.Text = "\r";
                paragraphFormatting.Apply(ownerRange);
                bodyFormatting.ApplyToParagraphMark(ownerRange);
            }
            return paragraphStart;
        }
        finally
        {
            Release(clearedParagraphRange);
            Release(clearedParagraph);
            Release(clearedParagraphs);
            Release(clearedParagraphProbe);
            Release(convertedTableRange);
            Release(ownerRange);
            Release(splitTable);
            Release(table);
            Release(sourceRange);
        }
    }

    private static void ValidateCleanFormatConversionParagraphBuffer(
        Document bufferDocument)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = bufferDocument.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The MathType→VisualTeX atomic replacement buffer must contain exactly one clean paragraph.");
            paragraph = paragraphs[1];
            range = paragraph.Range.Duplicate;
            if (!string.Equals(range.Text, "\r", StringComparison.Ordinal)
                || range.OMaths.Count != 0
                || range.InlineShapes.Count != 0
                || range.Fields.Count != 0
                || range.Tables.Count != 0)
                throw new InvalidDataException(
                    "The MathType→VisualTeX atomic replacement buffer is not structurally empty.");
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private int ReplaceNumberedMathTypeDisplaySourceParagraphForVisualTeX(
        Document document,
        WordFormulaFormatConversionTarget target,
        Document cleanParagraphBuffer,
        string sourceIdentityBookmarkName)
    {
        Bookmarks? bookmarks = null;
        Bookmark? identityBookmark = null;
        Range? identityRange = null;
        InlineShapes? identityShapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Paragraphs? cleanParagraphs = null;
        Paragraph? cleanParagraph = null;
        Range? cleanRange = null;
        Range? replacementRange = null;
        try
        {
            ValidateCleanFormatConversionParagraphBuffer(cleanParagraphBuffer);
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(sourceIdentityBookmarkName))
                throw new InvalidOperationException(
                    "The MathType source identity bookmark disappeared before atomic replacement.");
            identityBookmark = bookmarks[sourceIdentityBookmarkName];
            identityRange = identityBookmark.Range;
            identityShapes = identityRange.InlineShapes;
            if (identityShapes.Count != 1)
                throw new InvalidOperationException(
                    "The MathType source identity no longer owns exactly one inline object.");
            shape = identityShapes[1];
            if (!MathTypeOleInterop.IsMathTypeOle(shape))
                throw new InvalidOperationException(
                    "The MathType source identity no longer owns Equation.DSMT4.");

            // Re-read every COM range from the identity bookmark immediately before
            // the structural mutation. Never reuse the capture-time numeric range or
            // a paragraph RCW retained across a previous formula conversion.
            shapeRange = shape.Range.Duplicate;
            if (identityRange.Start < shapeRange.Start
                || identityRange.End > shapeRange.End)
                throw new InvalidOperationException(
                    "The MathType source identity bookmark escaped its live Equation.DSMT4 range.");
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered MathType display source no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was stopped.");

            cleanParagraphs = cleanParagraphBuffer.Paragraphs;
            cleanParagraph = cleanParagraphs[1];
            cleanRange = cleanParagraph.Range.Duplicate;
            var start = paragraphRange.Start;
            var end = paragraphRange.End;
            var paragraphText = paragraphRange.Text ?? string.Empty;
            if (!paragraphText.EndsWith("\r", StringComparison.Ordinal)
                || end <= start)
                throw new InvalidDataException(
                    "The numbered MathType source does not end in one ordinary body paragraph mark.");

            // Keep the ORIGINAL body paragraph mark. Replacing the complete
            // paragraph (including ¶) immediately before a user table makes Word
            // 2021 adopt the replacement ¶ as cell (1,1); the converted formula is
            // then inserted into that user table. The validated MathType owner has
            // no user content, so one atomic text replacement of the editable
            // paragraph contents removes Equation.DSMT4, MTPlaceRef and its nested
            // SEQ fields while leaving the body/table boundary physically intact.
            var bodyFormatting = WordCharacterFormatting.CaptureParagraphMark(paragraphRange);
            replacementRange = document.Range(start, end - 1);
            replacementRange.Text = string.Empty;
            Range? emptyParagraph = null;
            Cell? emptyParagraphCell = null;
            InlineShapes? emptyShapes = null;
            OMaths? emptyMaths = null;
            Fields? emptyFields = null;
            try
            {
                emptyParagraph = document.Range(start, start + 1);
                emptyShapes = emptyParagraph.InlineShapes;
                emptyMaths = emptyParagraph.OMaths;
                emptyFields = WordFormulaHost.GetLocalFields(emptyParagraph);
                if (emptyParagraph.Text != "\r"
                    || emptyShapes.Count != 0
                    || emptyMaths.Count != 0
                    || emptyFields.Count != 0)
                    throw new InvalidDataException(
                        "The MathType replacement did not leave one structurally empty body paragraph.");
                emptyParagraphCell = WordFormulaHost.TryGetOwningCell(emptyParagraph);
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-mathtype-visualtex-empty-owner start={start} range={emptyParagraph.Start}:{emptyParagraph.End} cell={(emptyParagraphCell is null ? "body" : emptyParagraphCell.RowIndex + "," + emptyParagraphCell.ColumnIndex)}");
                if (emptyParagraphCell is not null)
                    throw new InvalidDataException(
                        "Word merged the cleared MathType body paragraph into the following user table.");
                bodyFormatting.ApplyToParagraphMark(emptyParagraph);
            }
            finally
            {
                Release(emptyFields);
                Release(emptyMaths);
                Release(emptyShapes);
                Release(emptyParagraphCell);
                Release(emptyParagraph);
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-mathtype-visualtex-paragraph-atomically-replaced formulaId={target.SourceFormulaId} bookmark={sourceIdentityBookmarkName} range={start}:{end} start={start}");
            return start;
        }
        finally
        {
            Release(replacementRange);
            Release(cleanRange);
            Release(cleanParagraph);
            Release(cleanParagraphs);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
            Release(identityShapes);
            Release(identityRange);
            Release(identityBookmark);
            Release(bookmarks);
        }
    }

    private int ReplaceMathTypeDisplaySourceParagraphAtomically(
        Document document,
        WordFormulaFormatConversionTarget target,
        WordOmmlConverter.BatchSource batchSource)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? cleanParagraph = null;
        try
        {
            shape = FindMathTypeOleByRange(
                    document,
                    target.SourceObjectId,
                    allowGlobalFallback: false)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before atomic paragraph replacement.");
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display source no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was stopped.");
            var start = paragraphRange.Start;
            cleanParagraph = batchSource
                .ReplaceTargetParagraphAtomicallyWithCleanParagraph(
                    document,
                    paragraphRange);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-mathtype-paragraph-atomically-replaced formulaId={target.SourceFormulaId} range={paragraphRange.Start}:{paragraphRange.End} start={cleanParagraph.Start}");
            return start;
        }
        finally
        {
            Release(cleanParagraph);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
            Release(shape);
        }
    }

    private static Range ResolveTemporaryLatexBridgeRange(
        Document document,
        int expectedStart,
        string latexSource)
    {
        Range? content = null;
        Range? candidate = null;
        try
        {
            content = document.Content;
            var normalizedExpected = NormalizeFormulaToLatexVerificationText(latexSource);
            // Word paragraph cleanup can move the verified temporary source by a
            // handful of visible characters, but a merged hidden MathType section
            // field can contribute more than a hundred *story* characters to the
            // old coordinate even though it renders nothing. Search only a bounded
            // neighborhood around this source host, never the full document.
            for (var distance = 0; distance <= 512; distance++)
            {
                foreach (var direction in distance == 0 ? new[] { 0 } : new[] { -1, 1 })
                {
                    var start = expectedStart + distance * direction;
                    if (start < content.Start || start > content.End) continue;
                    var end = Math.Min(content.End, start + latexSource.Length);
                    if (end <= start) continue;
                    Release(candidate);
                    candidate = document.Range(start, end);
                    if (!string.Equals(
                            NormalizeFormulaToLatexVerificationText(candidate.Text ?? string.Empty),
                            normalizedExpected,
                            StringComparison.Ordinal))
                        continue;
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            throw new InvalidDataException(
                "The temporary LaTeX bridge changed before target insertion.");
        }
        finally
        {
            Release(candidate);
            Release(content);
        }
    }

    private void ReplaceNumberedVisualTeXParagraphWithLatex(
        Document document,
        FormulaToLatexTarget target,
        ref bool documentMutationStarted,
        IReadOnlyDictionary<string, int>? knownReferenceCounts,
        bool preserveCrossReferences,
        bool deferCaptionFlowTailCleanup = false)
    {
        var formulaId = target.Metadata.FormulaId;
        var sourceWatch = System.Diagnostics.Stopwatch.StartNew();
        long sourceCheckpoint = 0;
        void TraceSource(string stage) => TraceAcceptancePerformance("NativeSource", stage, sourceWatch, ref sourceCheckpoint);
        InlineShape? shape = null;
        Range? shapeRange = null;
        Range? owner = null;
        Range? prefix = null;
        Range? suffix = null;
        Range? inserted = null;
        Fields? fields = null;
        Field? embed = null;
        try
        {
            shape = FindByFormulaId(document, formulaId,
                $"{RangeReferencePrefix}{target.FormulaRange.Start}:{target.FormulaRange.End}", allowGlobalFallback: false)
                ?? throw new InvalidDataException("The VisualTeX paragraph lost its formula identity.");
            documentMutationStarted = true;
            BindOleIdentityBookmark(shape, formulaId);
            shapeRange = shape.Range;
            if (TrySplitEmbeddedMathTypeSectionStateFromVisualTeXParagraph(document, formulaId, shapeRange))
            {
                documentMutationStarted = true;
                Release(shapeRange);
                shapeRange = null;
                Release(shape);
                shape = FindByFormulaId(document, formulaId, null, allowGlobalFallback: false)
                    ?? throw new InvalidDataException("The VisualTeX formula moved while isolating section state.");
                shapeRange = shape.Range;
            }
            if (!IsSafeVisualTeXNumberingParagraphForConversion(document, formulaId, shapeRange))
                throw new InvalidDataException("The VisualTeX numbered paragraph contains unowned content.");
            using var paragraphFormatting = WordParagraphFormatting.Capture(shapeRange);
            var bodyFormatting = WordEquationNumbering.CaptureNumberedBodyCharacterFormatting(document, formulaId);
            TraceSource("validate-and-format");
            documentMutationStarted = true;
            if (!preserveCrossReferences)
                WordEquationNumbering.FreezeFormulaCrossReferences(document, formulaId, knownReferenceCounts);
            WordEquationNumbering.RemoveFormulaNumberingArtifacts(document, formulaId,
                deferCaptionFlowTailCleanup: deferCaptionFlowTailCleanup);
            RemoveInlineBaselineSentinel(document, formulaId);
            TraceSource("remove-numbering");

            // Removing the caption and REF changes live Word ranges. Resolve the
            // same OLE again through its durable identity before replacing anything.
            Release(shapeRange);
            shapeRange = null;
            Release(shape);
            shape = FindByFormulaId(document, formulaId, null, allowGlobalFallback: false)
                ?? throw new InvalidDataException("The VisualTeX formula disappeared during unnumbering.");
            shapeRange = shape.Range;
            owner = DuplicateContainingParagraphRange(shapeRange);
            if (owner.Tables.Count != 0 || owner.OMaths.Count != 0 || owner.InlineShapes.Count != 1
                || owner.Frames.Count != 0 || owner.ContentControls.Count != 0)
                throw new InvalidDataException("The unnumbered VisualTeX owner contains another structure.");
            fields = owner.Fields;
            if (fields.Count != 1)
                throw new InvalidDataException("The unnumbered VisualTeX owner has fields outside its OLE.");
            embed = fields[1];
            if (embed.Type != WdFieldType.wdFieldEmbed)
                throw new InvalidDataException("The VisualTeX owner no longer contains its EMBED field.");
            prefix = document.Range(owner.Start, shapeRange.Start);
            suffix = document.Range(shapeRange.End, owner.End);
            static bool IsHostWhitespace(string? text) => text is not null
                && text.All(character => character is '\t' or '\v' or '\r' or ' ');
            if (!IsHostWhitespace(prefix.Text) || !IsHostWhitespace(suffix.Text))
                throw new InvalidDataException("The VisualTeX paragraph has text outside the identified OLE.");

            // Force removal of the complete proven EMBED field in one edit, but
            // retain this original body paragraph's final mark. Replacing that mark
            // also retargets the next paragraph's bookmarks at a user-table edge.
            // This is the same editable-body contract used by native MathType input.
            var start = owner.Start;
            if (owner.End <= start || !(owner.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal))
                throw new InvalidDataException("The VisualTeX body owner has no ordinary paragraph terminator.");
            TraceSource("verify-unumbered-owner");
            ThrowIfFormulaToLatexFailureInjected(target);
            SetRangeToEditableParagraphBody(owner);
            owner.Text = target.LatexSource;
            inserted = document.Range(start, start + target.LatexSource.Length + 1);
            if (inserted.Text != target.LatexSource + "\r")
                throw new InvalidDataException("Word did not replace the complete VisualTeX paragraph.");
            if (inserted.InlineShapes.Count != 0 || inserted.OMaths.Count != 0
                || inserted.Fields.Count != 0 || inserted.Frames.Count != 0)
                throw new InvalidDataException("The VisualTeX replacement retained part of its generated host.");
            paragraphFormatting.Apply(inserted);
            bodyFormatting.ApplyToParagraphMark(inserted);
            inserted.SetRange(start, start + target.LatexSource.Length);
            VerifyLatexSourceRange(inserted, target.LatexSource, formulaId);
            NormalizeFormerNumberedFormulaParagraph(document, inserted);
            NormalizeLatexSourceRange(inserted, target.Metadata);
            TraceSource("replace-and-format");
            WordDoubleClickHook.TraceMessage(
                $"formula-to-latex-visualtex-paragraph-replaced formulaId={formulaId} range={inserted.Start}:{inserted.End}");
        }
        finally
        {
            Release(embed);
            Release(fields);
            Release(inserted);
            Release(suffix);
            Release(prefix);
            Release(owner);
            Release(shapeRange);
            Release(shape);
        }
    }

    private int DeleteSimpleSourceHost(
        Document document,
        string sourceMode,
        WordFormulaFormatConversionTarget target,
        IReadOnlyDictionary<string, int>? knownReferenceCounts = null,
        bool preserveCrossReferences = false,
        bool useIdentityBookmarkSourceCheck = false,
        bool deferCaptionFlowTailCleanup = false)
    {
        InlineShape? shape = null;
        Range? shapeRange = null;
        Table? table = null;
        Range? tableRange = null;
        Range? numberingOwnerRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? contentRange = null;
        Bookmark? ommlBookmark = null;
        try
        {
            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                // Do not treat InlineShape.Delete() as a committed replacement.
                // Word can defer OLE removal until something is actually written
                // at the same range, which caused real documents to report success
                // while leaving one VisualTeX object behind. Reuse the mature
                // VisualTeX OLE -> LaTeX replacement path: it deletes the complete
                // source host/numbering, writes the recoverable source text, and
                // verifies that Word materialized that text. The text is only a
                // transaction-local bridge; Numbered/DisplayMode were captured
                // before this call and are used independently for target creation.
                shape = FindByFormulaId(
                        document,
                        target.SourceFormulaId,
                        target.SourceObjectId,
                        allowGlobalFallback: false)
                    ?? throw new InvalidOperationException(
                        "The VisualTeX source formula moved before replacement.");
                shapeRange = shape.Range;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal)
                    && TrySplitEmbeddedMathTypeSectionStateFromVisualTeXParagraph(
                        document,
                        target.SourceFormulaId,
                        shapeRange))
                {
                    // Some real documents have a valid MTEditEquationSection2 state
                    // paragraph accidentally merged into the first VisualTeX OLE
                    // display row.  The preflight accepts only that exact generated
                    // prefix.  Split it back out before the LaTeX bridge is created,
                    // then reacquire the OLE because Word shifted its live range.
                    Release(shapeRange);
                    shapeRange = null;
                    Release(shape);
                    shape = FindByFormulaId(
                            document,
                            target.SourceFormulaId,
                            target.SourceObjectId,
                            allowGlobalFallback: false)
                        ?? throw new InvalidOperationException(
                            "The VisualTeX source formula disappeared while separating its MathType section state.");
                    shapeRange = shape.Range;
                }
                var start = shapeRange.Start;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata);
                    if (table is not null)
                    {
                        if (table.Rows.Count != 1)
                        {
                            if (!IsSafeVisualTeXNumberingTableForConversion(table, shapeRange))
                                throw new InvalidOperationException(
                                    "The legacy VisualTeX numbering table contains nonempty extra rows or another formula object; conversion was refused before modifying the document.");
                            TrimEmptyVisualTeXNumberingRows(table, shapeRange);
                            Release(shapeRange);
                            shapeRange = null;
                            Release(shape);
                            shape = FindByFormulaId(
                                    document,
                                    target.SourceFormulaId,
                                    target.SourceObjectId,
                                    allowGlobalFallback: false)
                                ?? throw new InvalidOperationException(
                                    "The VisualTeX source formula disappeared while normalizing an empty numbering-table row.");
                            shapeRange = shape.Range;
                            Release(table);
                            table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata)
                                ?? throw new InvalidOperationException(
                                    "The numbered VisualTeX source lost its legacy table after empty-row normalization.");
                        }
                        tableRange = table.Range.Duplicate;
                        start = tableRange.Start;
                    }
                    else
                    {
                        if (!IsSafeVisualTeXNumberingParagraphForConversion(
                                document,
                                target.SourceFormulaId,
                                shapeRange))
                            throw new InvalidOperationException(
                                "The numbered VisualTeX source no longer owns one safe MathType-style tab paragraph.");
                        numberingOwnerRange = WordEquationNumbering.FindNumberingOwnerRange(
                                document,
                                target.SourceFormulaId)
                            ?? throw new InvalidOperationException(
                                "The numbered VisualTeX source lost its tab paragraph before replacement.");
                        start = numberingOwnerRange.Start;
                    }
                }

                var latexSource = BuildFormulaLatexSource(target.Metadata);
                var sourceTarget = new FormulaToLatexTarget
                {
                    Metadata = target.Metadata,
                    ObjectMode = FormulaOleContract.NativeOleMode,
                    LatexSource = latexSource,
                    Start = shapeRange.Start,
                    End = shapeRange.End,
                    FormulaRange = shapeRange,
                    OleShape = shape,
                };
                shapeRange = null;
                shape = null;
                var mutationStarted = false;
                try
                {
                    ConvertFormulaTargetToLatex(
                        document,
                        sourceTarget,
                        ref mutationStarted,
                        knownReferenceCounts,
                        preserveCrossReferences,
                        deferCaptionFlowTailCleanup: deferCaptionFlowTailCleanup);

                    // Word can acknowledge InlineShape.Delete() yet keep the EMBED
                    // field alive until a later structural edit. In that case the
                    // verified LaTeX bridge is inserted immediately before the OLE,
                    // shifting the same FormulaId forward by exactly the bridge text
                    // length. Once the recoverable LaTeX text has been verified it is
                    // safe to force-commit removal by replacing the live OLE field
                    // range with empty text inside the same custom Undo transaction.
                    InlineShape? deferredShape = null;
                    try
                    {
                        var deferredSourceStillPresent = false;
                        if (useIdentityBookmarkSourceCheck)
                        {
                            // Whole-document VisualTeX -> MathType has two independent
                            // Flat OPC cardinality checks at commit. Here we only need to
                            // detect Word's rare deferred-delete case, not rediscover the
                            // source through a 16K-character neighborhood on every item.
                            // The VTO_ identity bookmark, when it survives the delete, is
                            // an exact local owner and turns this normal-path check into O(1).
                            deferredShape = FindVisualTeXSourceByIdentityBookmark(
                                document,
                                target.SourceFormulaId);
                            deferredSourceStillPresent = deferredShape is not null;
                        }
                        else
                        {
                            deferredSourceStillPresent =
                                IsSimpleFormatSourcePresent(document, sourceMode, target);
                        }

                        if (deferredSourceStillPresent)
                        {
                            Range? deferredRange = null;
                            try
                            {
                                deferredShape ??= FindByFormulaId(
                                        document,
                                        target.SourceFormulaId,
                                        target.SourceObjectId,
                                        allowGlobalFallback: false)
                                    ?? throw new InvalidOperationException(
                                        "Word deferred VisualTeX OLE removal, but the live source object could not be resolved.");
                                deferredRange = deferredShape.Range;
                                WordDoubleClickHook.TraceMessage(
                                    $"format-conversion-force-remove-deferred-source formulaId={target.SourceFormulaId} range={deferredRange.Start}:{deferredRange.End} latexLength={latexSource.Length}");
                                deferredRange.Text = string.Empty;
                            }
                            finally { Release(deferredRange); }

                            if (useIdentityBookmarkSourceCheck)
                            {
                                Release(deferredShape);
                                deferredShape = FindVisualTeXSourceByIdentityBookmark(
                                    document,
                                    target.SourceFormulaId);
                                if (deferredShape is not null)
                                    throw new InvalidOperationException(
                                        "Word retained the VisualTeX source formula after forced local removal.");
                            }
                            else
                            {
                                // A second lookup adds proof only when Word actually
                                // deferred removal. The common path stops after the
                                // first absence proof above.
                                EnsureSimpleFormatSourceRemoved(
                                    document,
                                    sourceMode,
                                    target,
                                    "latex-bridge-force-remove");
                            }
                        }
                    }
                    finally { Release(deferredShape); }

                    // Remove only the verified temporary LaTeX characters. For a
                    // former numbered table the mature converter deliberately adds
                    // one paragraph mark; keep that paragraph as the clean insertion
                    // host for the new MathType display equation.
                    Range? latexRange = null;
                    try
                    {
                        latexRange = ResolveTemporaryLatexBridgeRange(
                            document,
                            start,
                            latexSource);
                        // Paragraph-prefix cleanup can delete VisualTeX's leading
                        // TAB after the temporary source text was written.  The
                        // bridge itself is still verified; return its live start so
                        // target insertion cannot be one character off after Word's
                        // structural normalization.
                        start = latexRange.Start;
                        latexRange.Text = string.Empty;
                    }
                    finally { Release(latexRange); }
                }
                finally
                {
                    ReleaseFormulaToLatexTargets(new[] { sourceTarget });
                }
                return start;
            }

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                shapeRange = ResolveSimpleOmmlSourceRange(document, target)
                    ?? throw new InvalidOperationException(
                        "The Word OMML source formula moved before replacement.");
                var start = shapeRange.Start;
                if (target.Numbered
                    && string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
                {
                    table = TryGetVisualTeXNumberedTable(shapeRange, target.Metadata);
                    if (table is not null)
                    {
                        var currentDirectTable = WordEquationNumbering
                            .HasReusableNumberedNativeOmmlDirectTableHost(
                                document,
                                shapeRange,
                                target.SourceFormulaId);
                        if (!currentDirectTable
                            && !IsSafeOmmlNumberingTableForConversion(table, shapeRange))
                            throw new InvalidOperationException(
                                "The legacy OMML numbering table contains ordinary user content or another formula; conversion was refused before modifying the document.");
                        tableRange = table.Range.Duplicate;
                        start = tableRange.Start;
                    }
                    else
                    {
                        if (!IsSafeOmmlNumberingParagraphForConversion(
                                document,
                                target.SourceFormulaId,
                                shapeRange))
                            throw new InvalidOperationException(
                                "The numbered OMML source lost its safe center/right-tab paragraph before replacement.");
                        numberingOwnerRange = WordEquationNumbering.FindNumberingOwnerRange(
                                document,
                                target.SourceFormulaId)
                            ?? throw new InvalidOperationException(
                                "The numbered OMML source lost its paragraph owner before replacement.");
                        start = numberingOwnerRange.Start;
                    }
                }

                ommlBookmark = WordOmmlFormulaStore.FindByFormulaId(
                    document,
                    target.SourceFormulaId);
                var latexSource = BuildFormulaLatexSource(target.Metadata);
                var sourceTarget = new FormulaToLatexTarget
                {
                    Metadata = target.Metadata,
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                    LatexSource = latexSource,
                    Start = shapeRange.Start,
                    End = shapeRange.End,
                    FormulaRange = shapeRange,
                    OmmlBookmark = ommlBookmark,
                };
                shapeRange = null;
                ommlBookmark = null;
                var mutationStarted = false;
                try
                {
                    ConvertFormulaTargetToLatex(
                        document,
                        sourceTarget,
                        ref mutationStarted,
                        knownReferenceCounts,
                        preserveCrossReferences);
                    EnsureSimpleFormatSourceRemoved(
                        document,
                        sourceMode,
                        target,
                        "omml-latex-bridge");

                    Range? latexRange = null;
                    try
                    {
                        latexRange = document.Range(
                            start,
                            Math.Min(document.Content.End, start + latexSource.Length));
                        if (!string.Equals(
                                NormalizeFormulaToLatexVerificationText(latexRange.Text ?? string.Empty),
                                NormalizeFormulaToLatexVerificationText(latexSource),
                                StringComparison.Ordinal))
                            throw new InvalidDataException(
                                "The temporary OMML LaTeX bridge changed before target insertion.");
                        latexRange.Text = string.Empty;
                    }
                    finally { Release(latexRange); }
                }
                finally
                {
                    ReleaseFormulaToLatexTargets(new[] { sourceTarget });
                }
                return start;
            }

            shape = FindMathTypeOleByRange(
                    document,
                    target.SourceObjectId,
                    allowGlobalFallback: false)
                ?? throw new InvalidOperationException(
                    "The MathType source formula moved before replacement.");
            shapeRange = shape.Range;
            var mathTypeStart = shapeRange.Start;
            if (!string.Equals(target.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineOleTypingAnchorAfter(shape);
                // Word 2021 can defer InlineShape.Delete() for an embedded OLE field
                // that sits inside ordinary paragraph text. The source identity
                // bookmark then still contains Equation.DSMT4 when the caller
                // performs its immediate post-delete check, even though display
                // paragraph deletion is already atomic. Replace the exact OLE field
                // range with empty text instead; this commits removal synchronously
                // without touching surrounding prose.
                shapeRange.Text = string.Empty;
                return mathTypeStart;
            }

            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The MathType display source no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The MathType display paragraph contains ordinary user text; conversion was stopped.");
            // Delete the complete MathType paragraph body in one Word operation,
            // retaining only its final paragraph mark. Word 2021 can crash when the
            // outer MTPlaceRef tree, its nested fields and the Equation.DSMT4 OLE
            // are dismantled in separate mutations; replacing the full body lets
            // Word tear down the entire field/object owner atomically.
            contentRange = paragraphRange.Duplicate;
            mathTypeStart = contentRange.Start;
            SetRangeToEditableParagraphBody(contentRange);
            contentRange.Delete();
            return mathTypeStart;
        }
        finally
        {
            Release(ommlBookmark);
            Release(contentRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(numberingOwnerRange);
            Release(tableRange);
            Release(table);
            Release(shapeRange);
            Release(shape);
        }
    }

    private static void SetRangeToEditableParagraphBody(Range range)
    {
        var text = range.Text ?? string.Empty;
        var editableEnd = range.End;
        // Ordinary Word paragraphs end in \r. A table cell is exposed through
        // Range.Text as \r\a, but those two text characters occupy ONE Word story
        // position (Cell.Range.End - Cell.Range.Start is therefore one smaller than
        // Cell.Range.Text.Length). Range.End must consequently move back by exactly
        // one position for either representation. Subtracting two for \r\a slices
        // into the preceding OLE/field and Word rejects the edit with 0x800A1710.
        if (text.EndsWith("\r\a", StringComparison.Ordinal)
            || text.EndsWith("\r", StringComparison.Ordinal)
            || text.EndsWith("\a", StringComparison.Ordinal))
            editableEnd -= 1;
        range.SetRange(range.Start, Math.Max(range.Start, editableEnd));
    }

    private static bool IsSafeVisualTeXNumberingTableForConversion(
        Table table,
        Range formulaRange)
    {
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        InlineShapes? tableShapes = null;
        Row? row = null;
        Range? rowRange = null;
        try
        {
            columns = table.Columns;
            if (columns.Count != 3) return false;
            tableRange = table.Range;
            tableShapes = tableRange.InlineShapes;
            if (tableShapes.Count != 1) return false;

            rows = table.Rows;
            var formulaRowFound = false;
            for (var index = 1; index <= rows.Count; index++)
            {
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                var ownsFormula = formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End;
                if (ownsFormula)
                {
                    if (formulaRowFound) return false;
                    formulaRowFound = true;
                    continue;
                }
                if (!IsStructurallyEmptyVisualTeXNumberingRow(rowRange))
                    return false;
            }
            return formulaRowFound;
        }
        finally
        {
            Release(rowRange);
            Release(row);
            Release(tableShapes);
            Release(tableRange);
            Release(columns);
            Release(rows);
        }
    }

    private static bool TryGetEmbeddedMathTypeSectionStatePrefix(
        Document document,
        Range ownerRange,
        Range formulaRange,
        out int splitPosition,
        out int sectionCodeStart,
        out int sectionCodeEnd)
    {
        splitPosition = -1;
        sectionCodeStart = -1;
        sectionCodeEnd = -1;
        Range? precedingFormula = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        var sectionFieldCount = 0;
        try
        {
            if (formulaRange.Start <= ownerRange.Start) return false;
            var candidateSplit = formulaRange.Start - 1;
            precedingFormula = document.Range(candidateSplit, formulaRange.Start);
            if (!string.Equals(precedingFormula.Text, "\t", StringComparison.Ordinal))
                return false;

            fields = WordFormulaHost.GetLocalFields(ownerRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                sectionFieldCount++;
                if (sectionFieldCount != 1) return false;
                result = field.Result;
                // The section-state field is document-level MathType bookkeeping.
                // It is legal here only as a complete prefix before VisualTeX's
                // leading center TAB.  Anything overlapping the formula or its TAB
                // is not a generated state we can separate safely.
                if (code.Start < ownerRange.Start
                    || code.End > candidateSplit
                    || result.End > candidateSplit)
                    return false;
                sectionCodeStart = code.Start;
                sectionCodeEnd = code.End;
            }

            if (sectionFieldCount != 1) return false;
            splitPosition = candidateSplit;
            return true;
        }
        catch
        {
            splitPosition = -1;
            sectionCodeStart = -1;
            sectionCodeEnd = -1;
            return false;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(precedingFormula);
        }
    }

    private static bool TrySplitEmbeddedMathTypeSectionStateFromVisualTeXParagraph(
        Document document,
        string formulaId,
        Range formulaRange)
    {
        Range? ownerRange = null;
        Range? split = null;
        Range? stateProbe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        try
        {
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId);
            if (ownerRange is null) return false;
            if (!TryGetEmbeddedMathTypeSectionStatePrefix(
                    document,
                    ownerRange,
                    formulaRange,
                    out var splitPosition,
                    out _,
                    out _))
                return false;

            // Split immediately before VisualTeX's leading center TAB.  This keeps
            // the complete MTEditEquationSection2 field in the preceding paragraph
            // and moves the TAB + OLE + right-tab REF row intact into the next one.
            // No field code or bookmark is rewritten.
            split = document.Range(splitPosition, splitPosition);
            split.Text = "\r";

            EnsureMathTypeNativeStyles(document);
            stateProbe = document.Range(
                ownerRange.Start,
                Math.Min(document.Content.End, ownerRange.Start + 1));
            paragraphs = stateProbe.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The embedded MathType section state did not separate into one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            if (shapes.Count != 0 || !IsMathTypeSectionStateParagraph(paragraphRange))
                throw new InvalidOperationException(
                    "The separated MathType section state still shares its paragraph with a formula object.");
            object sectionStyle = "MTEquationSection";
            paragraphRange.set_Style(ref sectionStyle);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-separated-mathtype-section-state formulaId={formulaId} split={splitPosition}");
            return true;
        }
        finally
        {
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(stateProbe);
            Release(split);
            Release(ownerRange);
        }
    }

    private static bool IsSafeVisualTeXNumberingParagraphForConversion(
        Document document,
        string formulaId,
        Range formulaRange)
    {
        Range? ownerRange = null;
        Range? visibleRange = null;
        Range? visibleStructureRange = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                document,
                formulaId);
            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(
                document,
                formulaId);
            if (ownerRange is null || visibleRange is null) return false;
            if (WordEquationNumbering.RangeIsWhollyWithinTable(formulaRange)
                && !WordFormulaHost.IsWithinSingleCell(ownerRange))
                return false;
            if (formulaRange.Start < ownerRange.Start
                || formulaRange.End > ownerRange.End
                || visibleRange.Start < formulaRange.End
                || visibleRange.Start < ownerRange.Start
                || visibleRange.End > ownerRange.End)
                return false;

            var hasEmbeddedSectionState =
                TryGetEmbeddedMathTypeSectionStatePrefix(
                    document,
                    ownerRange,
                    formulaRange,
                    out var sectionSplitPosition,
                    out var sectionCodeStart,
                    out var sectionCodeEnd);
            visibleStructureRange = hasEmbeddedSectionState
                ? document.Range(sectionSplitPosition, ownerRange.End)
                : ownerRange.Duplicate;

            shapes = ownerRange.InlineShapes;
            if (shapes.Count != 1) return false;
            maths = ownerRange.OMaths;
            if (maths.Count != 0) return false;
            if ((visibleStructureRange.Text ?? string.Empty).Count(character => character == '\t') < 2)
                return false;

            var expectedTarget = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            // This is a VisualTeX OLE host, not a MathType paragraph. Validate only
            // the two field families that are legal here: the one embedded OLE and
            // this formula's visible REF. Reusing IsSafeMathTypeDisplayParagraph
            // rejected every healthy table-free VisualTeX number because that helper
            // intentionally accepts only Equation.DSMT4/MTPlaceRef fields.
            fields = WordFormulaHost.GetLocalFields(ownerRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var fieldCode = code.Text ?? string.Empty;
                if (field.Type == WdFieldType.wdFieldEmbed)
                    continue;
                if (field.Type == WdFieldType.wdFieldRef
                    && fieldCode.IndexOf(
                        "REF " + expectedTarget,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;
                if (hasEmbeddedSectionState
                    && code.Start >= sectionCodeStart
                    && code.End <= sectionCodeEnd)
                    continue;
                return false;
            }

            // Reject ordinary prose even when field codes are hidden. When a
            // generated MathType section state was merged into this row, validate
            // only the actual VisualTeX TAB/OLE/REF suffix here; the state prefix
            // has already been structurally bounded above.
            var ownerText = visibleStructureRange.Text ?? string.Empty;
            var fieldDepth = 0;
            foreach (var character in ownerText)
            {
                if (character == '\u0013')
                {
                    fieldDepth++;
                    continue;
                }
                if (character == '\u0015')
                {
                    if (fieldDepth > 0) fieldDepth--;
                    continue;
                }
                if (fieldDepth > 0 || character == '\u0014') continue;
                if (char.IsWhiteSpace(character) || char.IsDigit(character)) continue;
                if (character < ' ' || character is '\u0001' or '\uFFFC') continue;
                if ("()[]{}.-–—_:;,+/\\".IndexOf(character) >= 0) continue;
                return false;
            }
            if (fieldDepth != 0) return false;

            Release(code);
            code = null;
            Release(field);
            field = null;
            Release(fields);
            fields = null;
            fields = WordFormulaHost.GetLocalFields(visibleRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var fieldCode = code.Text ?? string.Empty;
                if (fieldCode.IndexOf(
                        "REF " + expectedTarget,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(visibleStructureRange);
            Release(visibleRange);
            Release(ownerRange);
        }
    }

    private static bool IsSafeOmmlNumberingParagraphForConversion(
        Document document,
        string formulaId,
        Range formulaRange)
    {
        if (WordEquationNumbering.IsSafeNativeHashSequenceOmmlForConversion(
                document,
                formulaRange,
                formulaId))
            return true;

        // Current numbered OMML is one genuine display OMath whose legal m:eqArr
        // contains the formula plus #(SEQ VisualTeXEquation). It deliberately has
        // no paragraph TABs and no generated REF field, so the retired tab-layout
        // safety test below must not reject it. The strict native-host validator
        // proves FormulaId ownership, live SEQ, internal VTEqNum/VTEq/VTEqCap
        // aliases, one table-free paragraph and absence of Shape/TextBox content.
        if (WordEquationNumbering.HasReusableNumberedNativeOmmlHashSequenceHost(
                document,
                formulaRange,
                formulaId))
            return true;

        // Current numbered OMML is not a center/right-tab paragraph. It is one
        // genuine display OMath whose legal m:eqArr/#() delimiter owns a direct
        // SEQ VisualTeXEquation field. Reuse the same strict production health
        // check here so only that exact structure is accepted; arbitrary user
        // eqArr formulas and the retired artificial #(...REF...) wrapper still
        // fall through to the legacy migration validation below.
        if (WordEquationNumbering.HasReusableNumberedNativeOmmlHashSequenceHost(
                document,
                formulaRange,
                formulaId))
            return true;

        Range? ownerRange = null;
        Range? visibleRange = null;
        Range? beforeFormula = null;
        Range? betweenFormulaAndNumber = null;
        Range? afterNumber = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? ownerFields = null;
        Fields? visibleFields = null;
        Field? field = null;
        Range? code = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                document,
                formulaId);
            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(
                document,
                formulaId);
            if (ownerRange is null || visibleRange is null) return false;
            if ((bool)ownerRange.get_Information(WdInformation.wdWithInTable))
                return false;
            if (formulaRange.Start < ownerRange.Start
                || formulaRange.End > ownerRange.End
                || visibleRange.Start < formulaRange.End
                || visibleRange.Start < ownerRange.Start
                || visibleRange.End > ownerRange.End)
                return false;

            shapes = ownerRange.InlineShapes;
            if (shapes.Count != 0) return false;
            maths = ownerRange.OMaths;
            if (maths.Count != 1) return false;
            paragraphs = ownerRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            if (format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                return false;
            tabStops = format.TabStops;
            var hasFormulaTab = false;
            var hasRightTab = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                if (tabStop.Alignment is WdTabAlignment.wdAlignTabLeft
                    or WdTabAlignment.wdAlignTabCenter)
                    hasFormulaTab = true;
                else if (tabStop.Alignment == WdTabAlignment.wdAlignTabRight)
                    hasRightTab = true;
            }
            if (!hasFormulaTab || !hasRightTab) return false;

            beforeFormula = document.Range(ownerRange.Start, formulaRange.Start);
            betweenFormulaAndNumber = document.Range(
                formulaRange.End,
                visibleRange.Start);
            afterNumber = document.Range(visibleRange.End, ownerRange.End);
            if (!ContainsOnlyOmmlNumberingLayoutText(beforeFormula.Text)
                || !ContainsOnlyOmmlNumberingLayoutText(betweenFormulaAndNumber.Text)
                || !ContainsOnlyOmmlNumberingLayoutText(afterNumber.Text))
                return false;
            var ownerText = ownerRange.Text ?? string.Empty;
            if (ownerText.Count(character => character == '\t') < 2)
                return false;

            ownerFields = ownerRange.Fields;
            visibleFields = visibleRange.Fields;
            if (visibleFields.Count == 0 || ownerFields.Count != visibleFields.Count)
                return false;
            var expectedTarget = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            for (var index = 1; index <= visibleFields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = visibleFields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "REF " + expectedTarget,
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
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
            Release(code);
            Release(field);
            Release(visibleFields);
            Release(ownerFields);
            Release(maths);
            Release(shapes);
            Release(afterNumber);
            Release(betweenFormulaAndNumber);
            Release(beforeFormula);
            Release(visibleRange);
            Release(ownerRange);
        }
    }

    private static bool ContainsOnlyOmmlNumberingLayoutText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var character in text!)
        {
            if (character is '\t' or '\v' or '\r' or '\n' or '\a'
                or '\u0001' or '\u0013' or '\u0014' or '\u0015'
                or '\u200B' or '\u200C' or '\u2060' or '\uFEFF')
                continue;
            if (!char.IsWhiteSpace(character)) return false;
        }
        return true;
    }

    private static bool IsSafeOmmlNumberingTableForConversion(
        Table table,
        Range formulaRange)
    {
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        OMaths? tableMaths = null;
        InlineShapes? tableShapes = null;
        Row? row = null;
        Range? rowRange = null;
        OMaths? rowMaths = null;
        try
        {
            columns = table.Columns;
            if (columns.Count != 3) return false;
            tableRange = table.Range;
            tableMaths = tableRange.OMaths;
            if (tableMaths.Count != 1) return false;
            tableShapes = tableRange.InlineShapes;
            if (tableShapes.Count != 0) return false;

            rows = table.Rows;
            var formulaRowFound = false;
            for (var index = 1; index <= rows.Count; index++)
            {
                Release(rowMaths);
                rowMaths = null;
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                var ownsFormula = formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End;
                if (!ownsFormula)
                {
                    if (!IsStructurallyEmptyVisualTeXNumberingRow(rowRange))
                        return false;
                    continue;
                }

                if (formulaRowFound) return false;
                formulaRowFound = true;
                rowMaths = rowRange.OMaths;
                if (rowMaths.Count != 1) return false;
                if (!IsSafeOmmlNumberingRowOutsideFormula(rowRange, formulaRange))
                    return false;
            }
            return formulaRowFound;
        }
        finally
        {
            Release(rowMaths);
            Release(rowRange);
            Release(row);
            Release(tableShapes);
            Release(tableMaths);
            Release(tableRange);
            Release(columns);
            Release(rows);
        }
    }

    private static bool IsSafeOmmlNumberingRowOutsideFormula(
        Range rowRange,
        Range formulaRange)
    {
        if (formulaRange.Start < rowRange.Start || formulaRange.End > rowRange.End)
            return false;
        Document? document = null;
        Range? before = null;
        Range? after = null;
        try
        {
            document = rowRange.Document;
            before = document.Range(rowRange.Start, formulaRange.Start);
            after = document.Range(formulaRange.End, rowRange.End);
            return IsSafeMathTypeDisplayParagraph(before)
                && IsSafeMathTypeDisplayParagraph(after);
        }
        finally
        {
            Release(after);
            Release(before);
            Release(document);
        }
    }

    private static bool IsStructurallyEmptyVisualTeXNumberingRow(Range rowRange)
    {
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            shapes = rowRange.InlineShapes;
            if (shapes.Count != 0) return false;
            maths = rowRange.OMaths;
            if (maths.Count != 0) return false;
            fields = rowRange.Fields;
            if (fields.Count != 0) return false;
            bookmarks = rowRange.Bookmarks;
            if (bookmarks.Count != 0) return false;
            foreach (var character in rowRange.Text ?? string.Empty)
            {
                if (character is '\r' or '\a' or '\n' or '\v'
                    or '\u200b' or '\u200c' or '\u200d' or '\ufeff'
                    || char.IsWhiteSpace(character))
                    continue;
                return false;
            }
            return true;
        }
        finally
        {
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
        }
    }

    private static void TrimEmptyVisualTeXNumberingRows(
        Table table,
        Range formulaRange)
    {
        if (!IsSafeVisualTeXNumberingTableForConversion(table, formulaRange))
            throw new InvalidOperationException(
                "The VisualTeX numbering table contains nonempty extra rows or another formula object.");

        Rows? rows = null;
        Row? row = null;
        Range? rowRange = null;
        try
        {
            rows = table.Rows;
            for (var index = rows.Count; index >= 1; index--)
            {
                Release(rowRange);
                rowRange = null;
                Release(row);
                row = rows[index];
                rowRange = row.Range;
                if (formulaRange.Start >= rowRange.Start
                    && formulaRange.Start < rowRange.End)
                    continue;
                row.Delete();
            }
            if (table.Rows.Count != 1 || table.Columns.Count != 3)
                throw new InvalidOperationException(
                    "Word did not normalize the VisualTeX numbering table to one 1x3 formula row.");
        }
        finally
        {
            Release(rowRange);
            Release(row);
            Release(rows);
        }
    }

    private static bool IsSafeMathTypeDisplayParagraph(Range paragraphRange)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            // A display paragraph that carries an unrelated user field is not safe
            // to replace atomically. Only MathType's own embedded-equation and
            // MTPlaceRef/MTChap/MTSec/MTEqn field family may be ignored below.
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsKnownMathTypeDisplayFieldCode(code.Text))
                    return false;
            }

            var text = paragraphRange.Text ?? string.Empty;
            var fieldDepth = 0;
            foreach (var character in text)
            {
                // In Word's field-code view the literal instruction text becomes
                // part of Paragraph.Range.Text. Strip complete nested field trees
                // by their native control characters before looking for user prose.
                if (character == '\u0013')
                {
                    fieldDepth++;
                    continue;
                }
                if (character == '\u0015')
                {
                    if (fieldDepth > 0) fieldDepth--;
                    continue;
                }
                if (fieldDepth > 0 || character == '\u0014') continue;

                if (char.IsWhiteSpace(character) || char.IsDigit(character)) continue;
                if (character < ' ') continue;
                if (character is '\u0001' or '\uFFFC') continue;
                if ("()[]{}.-–—_:;,+/\\".IndexOf(character) >= 0) continue;
                return false;
            }
            return fieldDepth == 0;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool IsKnownMathTypeDisplayFieldCode(string? value)
    {
        var code = (value ?? string.Empty).Trim();
        if (code.Length == 0) return false;
        if (code.IndexOf(
                "EMBED Equation.DSMT4",
                StringComparison.OrdinalIgnoreCase) >= 0)
            return true;
        if (code.StartsWith(
                "MACROBUTTON MTPlaceRef",
                StringComparison.OrdinalIgnoreCase))
            return true;
        return code.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
            || code.StartsWith("SEQ MTChap ", StringComparison.OrdinalIgnoreCase)
            || code.StartsWith("SEQ MTSec ", StringComparison.OrdinalIgnoreCase);
    }

    private static void RemoveDetachedVisualTeXNumberingArtifacts(
        Document document,
        string formulaId)
    {
        // The visible 1x3 numbered host has already been deleted. Only the hidden
        // native SEQ caption/frame remains. Delete that detached structure without
        // touching any former table/cell Range.
        TryDeleteBookmark(document, WordEquationNumbering.NativeNumberBookmarkName(formulaId));

        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Range? captionRange = null;
        Frames? frames = null;
        Frame? frame = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)) return;
            captionBookmark = bookmarks[captionName];
            captionRange = captionBookmark.Range;
            try
            {
                frames = captionRange.Frames;
                if (frames.Count > 0)
                {
                    frame = frames[1];
                    frame.Delete();
                    Release(frame);
                    frame = null;
                    Release(frames);
                    frames = null;
                    Release(captionRange);
                    captionRange = captionBookmark.Range;
                }
            }
            catch
            {
                // The caption may already have lost its clipping frame. Deleting
                // the bookmarked caption contents is sufficient in that case.
            }
            captionRange.Delete();
        }
        finally
        {
            Release(frame);
            Release(frames);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static void TryDeleteBookmark(Document document, string name)
    {
        if (document is null || string.IsNullOrWhiteSpace(name)) return;
        try
        {
            if (!document.Bookmarks.Exists(name)) return;
            Bookmark? bookmark = null;
            try
            {
                bookmark = document.Bookmarks[name];
                bookmark.Delete();
            }
            finally { Release(bookmark); }
        }
        catch
        {
            // The source host itself has already been removed. A collapsed stale
            // bookmark is harmless and must never turn cleanup into a conversion failure.
        }
    }
}

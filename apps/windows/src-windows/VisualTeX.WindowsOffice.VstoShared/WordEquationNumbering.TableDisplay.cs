using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationNumbering
{
    private const float NativeOmmlTableSideWidthPoints = 60f;
    private const float NativeOmmlTableHeightSafetyPoints = 0.9f;

    private static bool ConfigureNumberedNativeOmmlDisplay(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool reuseExistingScaffold,
        FormulaMetadata? metadata,
        Action<string> traceStage,
        int? plannedOrdinal = null,
        string? plannedPrefix = null,
        bool deferFieldUpdate = false,
        bool deferExternalShapeCreation = false,
        bool deferMetadataPersistence = false)
    {
        _ = reuseExistingScaffold;
        _ = deferFieldUpdate;
        _ = deferExternalShapeCreation;

        Range? activeRange = null;
        Range? replacementRange = null;
        Table? table = null;
        Bookmark? repairedBookmark = null;
        Microsoft.Office.Interop.Word.Application? application = null;
        try
        {
            metadata ??= WordOmmlFormulaStore.TryRead(document, formulaId);
            activeRange = ResolveSingleNativeOmmlRange(formulaRange);
            EnsureNumberedOmmlIsDisplay(activeRange);
            traceStage("native-resolve-display");

            // A document created by the retired #(SEQ) route is valid migration
            // input, but its mathematical number must be stripped before the OMath
            // is placed in the center cell. The center cell owns formula semantics
            // only; every number field remains ordinary Word content in cell (1,3).
            var activeXml = activeRange.WordOpenXML;
            if (WordOmmlConverter.HasVisualTeXNativeEquationNumber(activeXml))
            {
                var semanticOmml = metadata is not null
                    ? WordOmmlConverter.StripVisualTeXNativeEquationNumberForManagedRepair(activeXml)
                    : WordOmmlConverter.StripVisualTeXNativeEquationNumber(activeXml);
                RemoveVisibleEquationNumber(document, formulaId);
                RemoveNativeCaption(document, formulaId);
                application = document.Application;
                replacementRange = WordOmmlConverter.ReplaceWithPreparedOmml(
                    application,
                    document,
                    activeRange,
                    semanticOmml,
                    display: true,
                    mathFontName: document.OMathFontName);
                Release(activeRange);
                activeRange = replacementRange;
                replacementRange = null;
                EnsureNumberedOmmlIsDisplay(activeRange);
                traceStage("strip-retired-native-hash");
            }
            else
            {
                // Remove only generated numbering artifacts. For an existing 1x3
                // host this clears an older hidden-caption + visible-REF scaffold;
                // the formula itself remains untouched in the center cell.
                RemoveVisibleEquationNumber(document, formulaId);
                RemoveNativeCaption(document, formulaId);
            }

            DeleteBookmarkOnly(document, EquationBookmarkName(formulaId));
            DeleteBookmarkOnly(document, NativeCaptionBookmarkName(formulaId));
            DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
            traceStage("native-clear-artifacts");

            float? nativeDisplayHeightPoints = null;
            var activeRangeAlreadyInTable = false;
            try
            {
                activeRangeAlreadyInTable = (bool)activeRange.get_Information(
                    WdInformation.wdWithInTable);
            }
            catch { }
            if (!activeRangeAlreadyInTable)
                nativeDisplayHeightPoints = TryMeasureNativeDisplayHeightPoints(
                    document,
                    activeRange);
            traceStage("native-measure-display");

            table = EnsureNativeOmmlNumberTableHost(
                document,
                activeRange,
                formulaId,
                out var removeGeneratedPostTableParagraph);
            traceStage("native-ensure-table");
            var refreshedTableFormula = ResolveSingleNativeOmmlRange(activeRange);
            Release(activeRange);
            activeRange = refreshedTableFormula;
            EnsureNumberedOmmlIsDisplay(activeRange);
            traceStage("native-resolve-table-formula");
            var needsHeightRepair = activeRangeAlreadyInTable
                && NeedsNativeOmmlTableHeightRepair(table);
            if (!nativeDisplayHeightPoints.HasValue
                && needsHeightRepair)
            {
                // Upgrade only legacy Auto-height 1x3 hosts. New hosts persist an
                // AtLeast row height, so ordinary F9/open/layout refreshes never
                // reopen a scratch document or pay this compatibility cost.
                try
                {
                    nativeDisplayHeightPoints =
                        WordOmmlConverter.MeasurePreparedDisplayHeightPoints(
                            document.Application,
                            document,
                            activeRange.WordOpenXML ?? string.Empty,
                            document.OMathFontName);
                }
                catch { }
            }
            traceStage("native-height-repair");
            var minimumDisplayHeightPoints = nativeDisplayHeightPoints
                ?? (!activeRangeAlreadyInTable || needsHeightRepair
                    ? formulaHeightPoints
                    : (float?)null);
            ConfigureNativeOmmlNumberTableGeometry(
                document,
                table,
                activeRange,
                minimumDisplayHeightPoints);
            traceStage("native-geometry");
            traceStage("native-1x3-table");

            var numberPlan = ResolveDirectTableNumberPlan(
                document,
                activeRange,
                formulaId,
                plannedOrdinal,
                plannedPrefix);
            traceStage("direct-number-plan");
            EnsureDirectTableSequenceNumber(
                document,
                table,
                activeRange,
                formulaId,
                formulaFontSizePoints,
                numberPlan.Ordinal,
                numberPlan.Prefix);
            traceStage("direct-visible-seq");

            if (removeGeneratedPostTableParagraph)
            {
                RemoveGeneratedPostTableParagraphBeforeKnownContent(
                    document,
                    table);
                traceStage("native-post-table-spacing");
            }

            if (metadata is not null)
            {
                repairedBookmark = WrapNativeOmmlTableFormulaIdentity(
                    document,
                    table,
                    activeRange,
                    metadata);
                if (!deferMetadataPersistence)
                {
                    WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                        metadata,
                        activeRange);
                    WordOmmlFormulaStore.Save(document, metadata);
                }
            }
            traceStage("native-wrap-identity");

            formulaRange.SetRange(activeRange.Start, activeRange.End);
            return true;
        }
        finally
        {
            Release(application);
            Release(repairedBookmark);
            Release(table);
            Release(replacementRange);
            Release(activeRange);
        }
    }

    private static Bookmark WrapNativeOmmlTableFormulaIdentity(
        Document document,
        Table table,
        Range formulaRange,
        FormulaMetadata metadata)
    {
        Cell? centerCell = null;
        Range? centerRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? anchor = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range;
            paragraphs = centerRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML center cell must contain exactly one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (formulaRange.Start < paragraphRange.Start
                || formulaRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OMML formula escaped its center-cell paragraph.");

            // A collapsed range rebuilt from absolute story coordinates can be
            // serialized by Word between two <w:tc> elements when its position is
            // exactly the center-cell boundary. Collapse the *paragraph object's*
            // own Range instead; this carries an unambiguous w:p container affinity
            // and serializes VTOMML_* immediately before m:oMathPara inside cell 2.
            anchor = paragraphRange.Duplicate;
            anchor.Collapse(WdCollapseDirection.wdCollapseStart);
            bookmarks = document.Bookmarks;
            var name = WordOmmlFormulaStore.BookmarkName(metadata.FormulaId);
            if (bookmarks.Exists(name))
                bookmarks[name].Delete();
            bookmark = bookmarks.Add(name, anchor);
            var result = bookmark;
            bookmark = null;
            return result;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(anchor);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(centerRange);
            Release(centerCell);
        }
    }

    private static (int Ordinal, string Prefix) ResolveDirectTableNumberPlan(
        Document document,
        Range formulaRange,
        string formulaId,
        int? plannedOrdinal,
        string? plannedPrefix)
    {
        if (plannedOrdinal.HasValue && plannedPrefix is not null)
            return (Math.Max(1, plannedOrdinal.Value), plannedPrefix);

        var format = ReadEquationNumberFormat(document);
        if (!format.UsesHeading
            && TryResolveContinuousDirectTableNumberPlan(
                document,
                formulaRange,
                out var continuousPlan))
            return continuousPlan;
        if (format.UsesHeading
            && TryResolveHeadingDirectTableNumberPlan(
                document,
                formulaRange,
                format,
                out var headingPlan))
            return headingPlan;

        var sequenceName = GetNativeEquationSequenceName(document);
        var existing = GetNativeEquationCaptionEntries(document, sequenceName)
            .Where(item => !string.Equals(
                item.FormulaId,
                formulaId,
                StringComparison.OrdinalIgnoreCase))
            .ToList();
        existing.Add(new NativeEquationCaptionEntry(
            formulaId,
            formulaRange.Start,
            string.Empty));
        var ordered = existing.OrderBy(item => item.Position).ToArray();
        var anchors = format.UsesHeading
            ? GetHeadingNumberAnchorsForFormatBatch(
                document,
                format.HeadingLevel,
                ordered)
            : Array.Empty<HeadingNumberAnchor>();
        var ordinalByScope = new Dictionary<int, int>();
        foreach (var item in ordered)
        {
            var scope = ResolveEquationNumberScope(item.Position, format, anchors);
            ordinalByScope.TryGetValue(scope.ScopePosition, out var ordinal);
            ordinal++;
            ordinalByScope[scope.ScopePosition] = ordinal;
            if (!string.Equals(
                    item.FormulaId,
                    formulaId,
                    StringComparison.OrdinalIgnoreCase))
                continue;
            return (
                plannedOrdinal ?? ordinal,
                plannedPrefix ?? scope.Prefix);
        }
        return (plannedOrdinal ?? 1, plannedPrefix ?? string.Empty);
    }

    private static bool TryResolveHeadingDirectTableNumberPlan(
        Document document,
        Range formulaRange,
        EquationNumberFormat format,
        out (int Ordinal, string Prefix) plan)
    {
        plan = (1, string.Empty);
        if (!format.UsesHeading) return false;

        Range? content = null;
        Range? headingSearch = null;
        Find? find = null;
        ParagraphFormat? findFormat = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ListFormat? listFormat = null;
        Frames? frames = null;
        Range? scopeRange = null;
        Bookmarks? scopeBookmarks = null;
        try
        {
            content = document.Content;
            var formulaStart = Math.Max(content.Start, formulaRange.Start);
            var bestHeadingPosition = int.MinValue;
            string? bestHeadingNumber = null;

            for (var outlineLevel = 1; outlineLevel <= format.HeadingLevel; outlineLevel++)
            {
                var searchEnd = formulaStart;
                while (searchEnd > content.Start)
                {
                    Release(frames); frames = null;
                    Release(listFormat); listFormat = null;
                    Release(paragraphRange); paragraphRange = null;
                    Release(paragraph); paragraph = null;
                    Release(paragraphs); paragraphs = null;
                    Release(findFormat); findFormat = null;
                    Release(find); find = null;
                    Release(headingSearch); headingSearch = null;

                    headingSearch = document.Range(content.Start, searchEnd);
                    find = headingSearch.Find;
                    find.ClearFormatting();
                    find.Text = "^p";
                    find.Forward = false;
                    find.Wrap = WdFindWrap.wdFindStop;
                    find.Format = true;
                    findFormat = find.ParagraphFormat;
                    findFormat.OutlineLevel = (WdOutlineLevel)outlineLevel;
                    if (!find.Execute()) break;

                    paragraphs = headingSearch.Paragraphs;
                    if (paragraphs.Count == 0)
                    {
                        searchEnd = Math.Max(content.Start, headingSearch.Start - 1);
                        continue;
                    }
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range;
                    searchEnd = Math.Max(content.Start, paragraphRange.Start);
                    if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                        continue;
                    try
                    {
                        frames = paragraphRange.Frames;
                        if (frames.Count > 0) continue;
                    }
                    catch { }

                    var listNumber = string.Empty;
                    try
                    {
                        listFormat = paragraphRange.ListFormat;
                        listNumber = NormalizeHeadingNumberText(listFormat.ListString);
                    }
                    catch { }
                    var explicitNumber = !string.IsNullOrWhiteSpace(listNumber)
                        ? listNumber
                        : ParseHeadingNumberFromText(paragraphRange.Text, outlineLevel);
                    if (string.IsNullOrWhiteSpace(explicitNumber))
                    {
                        // An unnumbered heading is ambiguous: it can be a synthesized
                        // chapter in an otherwise unnumbered document or a title that
                        // the explicit-numbering scheme intentionally skips. Preserve
                        // the mature full-document resolver for that uncommon case.
                        return false;
                    }
                    if (paragraphRange.Start <= bestHeadingPosition) break;
                    if (outlineLevel < format.HeadingLevel)
                        explicitNumber += string.Concat(
                            Enumerable.Repeat(".0", format.HeadingLevel - outlineLevel));
                    bestHeadingPosition = paragraphRange.Start;
                    bestHeadingNumber = explicitNumber;
                    break;
                }
            }

            var prefix = bestHeadingNumber is null
                ? string.Join(".", Enumerable.Repeat("0", format.HeadingLevel))
                    + format.Separator
                : bestHeadingNumber + format.Separator;
            var scopeStart = bestHeadingPosition == int.MinValue
                ? content.Start
                : bestHeadingPosition;
            if (formulaStart <= scopeStart)
            {
                plan = (1, prefix);
                return true;
            }

            scopeRange = document.Range(scopeStart, formulaStart);
            scopeBookmarks = scopeRange.Bookmarks;
            var priorNumberCount = 0;
            for (var index = 1; index <= scopeBookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = scopeBookmarks[index];
                    if (TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeNumberBookmarkPrefix,
                            out _))
                        priorNumberCount++;
                }
                finally { Release(bookmark); }
            }
            plan = (priorNumberCount + 1, prefix);
            return true;
        }
        catch
        {
            plan = (1, string.Empty);
            return false;
        }
        finally
        {
            Release(scopeBookmarks);
            Release(scopeRange);
            Release(frames);
            Release(listFormat);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(findFormat);
            Release(find);
            Release(headingSearch);
            Release(content);
        }
    }

    private static bool TryResolveContinuousDirectTableNumberPlan(
        Document document,
        Range formulaRange,
        out (int Ordinal, string Prefix) plan)
    {
        plan = (1, string.Empty);
        Range? content = null;
        Range? prefix = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            content = document.Content;
            if (formulaRange.Start <= content.Start)
                return true;
            prefix = document.Range(content.Start, formulaRange.Start);
            fields = prefix.Fields;
            var priorSequenceCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code); code = null;
                Release(field); field = null;
                field = fields[index];
                // OLE objects themselves are EMBED fields; their equation number is
                // a separate SEQ. Reject EMBED before touching Field.Code because
                // reading embedded-object field code is disproportionately costly.
                if (field.Type == WdFieldType.wdFieldEmbed) continue;
                code = field.Code;
                if (IsVisualTeXSequenceFieldCode(code.Text))
                    priorSequenceCount++;
            }
            plan = (priorSequenceCount + 1, string.Empty);
            return true;
        }
        catch
        {
            plan = (1, string.Empty);
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(prefix);
            Release(content);
        }
    }

    private static Table EnsureNativeOmmlNumberTableHost(
        Document document,
        Range formulaRange,
        string formulaId,
        out bool removeGeneratedPostTableParagraph)
    {
        removeGeneratedPostTableParagraph = false;
        if (IsNumberedEquationTable(formulaRange))
        {
            TrimBenignEmptyRowsFromNumberedTable(document, formulaRange, formulaId);
            Tables? existingTables = null;
            Table? existing = null;
            Rows? existingRows = null;
            Columns? existingColumns = null;
            try
            {
                existingTables = formulaRange.Tables;
                if (existingTables.Count == 0)
                    throw new InvalidOperationException(
                        "The numbered OMML range reports table ownership without a table.");
                existing = existingTables[1];
                existingRows = existing.Rows;
                existingColumns = existing.Columns;
                if (existingColumns.Count == 3 && existingRows.Count > 1)
                {
                    // The oldest VisualTeX OMML hosts can be 2x3 (or carry more
                    // benign empty trailing rows) before any VTEq_* numbering alias
                    // exists. Trim only rows independently proven empty from this
                    // exact table; do not depend on FindNumberedEquationTable(),
                    // which intentionally resolves through the later visible-number
                    // bookmark and therefore cannot see this pre-scaffold fixture.
                    RemoveEmptyTrailingNumberedTableRows(existing);
                    Release(existingRows);
                    existingRows = existing.Rows;
                }
                if (existingRows.Count != 1 || existingColumns.Count != 3)
                    throw new InvalidOperationException(
                        $"The managed numbered OMML table must converge to exactly 1x3, not {existingRows.Count}x{existingColumns.Count}.");
                NormalizeNativeOmmlCenterCellParagraphs(
                    document,
                    existing,
                    formulaRange,
                    formulaId);
                var result = existing;
                existing = null;
                return result;
            }
            finally
            {
                Release(existingColumns);
                Release(existingRows);
                Release(existing);
                Release(existingTables);
            }
        }

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefixRange = null;
        Range? suffixRange = null;
        Range? sourceContent = null;
        Range? formattedContent = null;
        Range? tableAnchor = null;
        Range? documentContent = null;
        Range? followingContentProbe = null;
        Paragraphs? followingParagraphs = null;
        Paragraph? followingParagraph = null;
        Range? followingParagraphRange = null;
        Paragraphs? anchorParagraphs = null;
        Paragraph? anchorParagraph = null;
        Range? anchorParagraphRange = null;
        ParagraphFormat? anchorFormat = null;
        TabStops? anchorTabs = null;
        Microsoft.Office.Interop.Word.Font? anchorFont = null;
        Frames? anchorFrames = null;
        Frame? anchorFrame = null;
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Range? centerInsertion = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Range? copiedFormulaRange = null;
        var sourceDeleted = false;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "VisualTeX cannot number native OMML spanning multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable))
                throw new InvalidOperationException(
                    "VisualTeX refused to nest a numbered OMML table inside another table.");

            prefixRange = document.Range(paragraphRange.Start, formulaRange.Start);
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            suffixRange = document.Range(formulaRange.End, editableEnd);
            if (!IsNumberingParagraphAdornment(prefixRange.Text)
                || !IsNumberingParagraphAdornment(suffixRange.Text))
                throw new InvalidOperationException(
                    "A numbered display OMML formula must occupy its own paragraph.");

            // Copy only the resolved professional OMath, never the surrounding
            // source paragraph adornments. OLE→OMML staging can legally contain
            // TAB/line-break/TAB around the temporary formula; copying the whole
            // paragraph would persist those runs inside center cell (1,2), where
            // they shift the table to page X=0 after save/reopen and violate the
            // invariant that the center cell owns mathematical content only.
            sourceContent = formulaRange.Duplicate;
            formattedContent = sourceContent.FormattedText;

            // Create the table after the source paragraph first. This keeps the
            // original professional OMath intact until the center-cell copy has
            // been verified, so a COM failure cannot destroy the user's formula.
            var sourceParagraphStart = paragraphRange.Start;
            var tablePosition = paragraphRange.End;
            documentContent = document.Content;
            var sourceFollowedByTable = false;
            var sourceFollowedByNonEmptyBodyParagraph = false;
            if (tablePosition < documentContent.End)
            {
                followingContentProbe = document.Range(
                    tablePosition,
                    Math.Min(documentContent.End, tablePosition + 1));
                sourceFollowedByTable = (bool)followingContentProbe.get_Information(
                    WdInformation.wdWithInTable);
                if (!sourceFollowedByTable)
                {
                    followingParagraphs = followingContentProbe.Paragraphs;
                    if (followingParagraphs.Count == 1)
                    {
                        followingParagraph = followingParagraphs[1];
                        followingParagraphRange = followingParagraph.Range.Duplicate;
                        sourceFollowedByNonEmptyBodyParagraph =
                            followingParagraphRange.Start == tablePosition
                            && !IsNumberingParagraphAdornment(
                                followingParagraphRange.Text);
                    }
                }
            }
            removeGeneratedPostTableParagraph =
                sourceFollowedByNonEmptyBodyParagraph;
            Release(followingParagraphRange);
            followingParagraphRange = null;
            Release(followingParagraph);
            followingParagraph = null;
            Release(followingParagraphs);
            followingParagraphs = null;
            Release(followingContentProbe);
            followingContentProbe = null;
            Release(documentContent);
            documentContent = null;

            paragraphRange.InsertParagraphAfter();
            documentContent = document.Content;
            tablePosition = Math.Max(
                documentContent.Start,
                Math.Min(tablePosition, documentContent.End - 1));
            tableAnchor = document.Range(
                tablePosition,
                Math.Min(documentContent.End, tablePosition + 1));
            anchorParagraphs = tableAnchor.Paragraphs;
            if (anchorParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "Word did not create one empty paragraph for the OMML number table.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range.Duplicate;

            if (sourceFollowedByTable)
            {
                // Tables.Add consumes the anchor paragraph. If the original source
                // paragraph was immediately before another table, consuming the only
                // new paragraph would make Word merge the new 1x3 host with that
                // following table into a 2x3 structure. Insert a second body paragraph
                // *before* the complete anchor (InsertParagraphAfter at this boundary
                // is interpreted as editing the next table's first cell), then use
                // the new first paragraph as the table anchor. The original one stays
                // behind the new table as its mandatory independent separator.
                anchorParagraphRange.InsertParagraphBefore();
                Release(anchorParagraphRange);
                anchorParagraphRange = null;
                Release(anchorParagraph);
                anchorParagraph = null;
                Release(anchorParagraphs);
                anchorParagraphs = null;
                Release(tableAnchor);
                tableAnchor = null;
                Release(documentContent);
                documentContent = document.Content;
                tableAnchor = document.Range(
                    tablePosition,
                    Math.Min(documentContent.End, tablePosition + 1));
                anchorParagraphs = tableAnchor.Paragraphs;
                if (anchorParagraphs.Count != 1)
                    throw new InvalidOperationException(
                        "Word did not preserve a dedicated table anchor before the following table.");
                anchorParagraph = anchorParagraphs[1];
                anchorParagraphRange = anchorParagraph.Range.Duplicate;
            }

            // InsertParagraphAfter inherits more than serialized pPr from the
            // source paragraph. In particular, a paragraph that just hosted an
            // MTDisplayEquation OLE can pass a stale live layout origin to the new
            // paragraph, making an otherwise normal inline table start at page X=0.
            // This paragraph was created by VisualTeX and is still empty, so detach
            // every paragraph-level layout characteristic before Tables.Add.
            try
            {
                anchorFrames = anchorParagraphRange.Frames;
                while (anchorFrames.Count > 0)
                {
                    Release(anchorFrame);
                    anchorFrame = anchorFrames[1];
                    anchorFrame.Delete();
                    Release(anchorFrame);
                    anchorFrame = null;
                    Release(anchorFrames);
                    anchorFrames = anchorParagraphRange.Frames;
                }
            }
            catch { }
            object normalStyle = WdBuiltinStyle.wdStyleNormal;
            anchorParagraphRange.set_Style(ref normalStyle);
            anchorFormat = anchorParagraphRange.ParagraphFormat;
            try { anchorFormat.Reset(); } catch { }
            anchorFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            anchorFormat.LeftIndent = 0f;
            anchorFormat.RightIndent = 0f;
            anchorFormat.FirstLineIndent = 0f;
            anchorFormat.SpaceBefore = 0f;
            anchorFormat.SpaceAfter = 0f;
            anchorFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            anchorTabs = anchorFormat.TabStops;
            anchorTabs.ClearAll();
            anchorFont = anchorParagraphRange.Font;
            try { anchorFont.Reset(); } catch { }
            anchorFont.Hidden = 0;
            anchorFont.Position = 0;
            tableAnchor.SetRange(
                anchorParagraphRange.Start,
                anchorParagraphRange.Start);
            table = document.Tables.Add(tableAnchor, 1, 3);
            rows = table.Rows;
            columns = table.Columns;
            if (rows.Count != 1 || columns.Count != 3)
                throw new InvalidOperationException(
                    "Word did not create the required 1x3 OMML numbering table.");

            centerCell = table.Cell(1, 2);
            centerCellRange = centerCell.Range;
            centerInsertion = centerCellRange.Duplicate;
            centerInsertion.End = Math.Max(
                centerInsertion.Start,
                centerInsertion.End - 1);
            centerInsertion.Collapse(WdCollapseDirection.wdCollapseStart);
            centerInsertion.FormattedText = formattedContent;

            OMaths? centerMaths = null;
            OMath? centerMath = null;
            try
            {
                Release(centerCellRange);
                centerCellRange = centerCell.Range;
                centerMaths = centerCellRange.OMaths;
                if (centerMaths.Count != 1)
                    throw new InvalidOperationException(
                        "Word did not preserve exactly one OMath in the center numbering cell.");
                centerMath = centerMaths[1];
                if (centerMath.Type != WdOMathType.wdOMathDisplay)
                    centerMath.Type = WdOMathType.wdOMathDisplay;
                copiedFormulaRange = centerMath.Range.Duplicate;
            }
            finally
            {
                Release(centerMath);
                Release(centerMaths);
            }

            // Only after the copied OMath is proven healthy do we remove the old
            // source content. Keep the now-empty source paragraph until the entire
            // 1x3 host, direct SEQ and FormulaId identity have been committed. Word
            // merges adjacent tables when this separator paragraph is deleted too
            // early; the post-commit spacing cleanup handles it safely.
            var oldEditable = document.Range(
                sourceParagraphStart,
                Math.Max(sourceParagraphStart, table.Range.Start - 1));
            try
            {
                oldEditable.Delete();
                sourceDeleted = true;
            }
            finally { Release(oldEditable); }

            // Deleting source content shifts the table. Re-resolve the live center
            // OMath from its cell before returning, matching the proven 1x3 path.
            Release(copiedFormulaRange);
            copiedFormulaRange = null;
            Release(centerCellRange);
            centerCellRange = centerCell.Range.Duplicate;
            OMaths? refreshedMaths = null;
            OMath? refreshedMath = null;
            try
            {
                refreshedMaths = centerCellRange.OMaths;
                if (refreshedMaths.Count != 1)
                    throw new InvalidOperationException(
                        "Word lost the center OMath after removing the source content.");
                refreshedMath = refreshedMaths[1];
                if (refreshedMath.Type != WdOMathType.wdOMathDisplay)
                    refreshedMath.Type = WdOMathType.wdOMathDisplay;
                copiedFormulaRange = refreshedMath.Range.Duplicate;
            }
            finally
            {
                Release(refreshedMath);
                Release(refreshedMaths);
            }

            formulaRange.SetRange(
                copiedFormulaRange.Start,
                copiedFormulaRange.End);
            var result = table;
            table = null;
            return result;
        }
        catch
        {
            if (!sourceDeleted && table is not null)
            {
                Range? rollback = null;
                try
                {
                    rollback = table.Range;
                    rollback.Delete();
                }
                catch { }
                finally { Release(rollback); }
            }
            throw;
        }
        finally
        {
            Release(copiedFormulaRange);
            Release(columns);
            Release(rows);
            Release(table);
            Release(centerInsertion);
            Release(centerCellRange);
            Release(centerCell);
            Release(anchorFrame);
            Release(anchorFrames);
            Release(anchorFont);
            Release(anchorTabs);
            Release(anchorFormat);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(followingParagraphRange);
            Release(followingParagraph);
            Release(followingParagraphs);
            Release(followingContentProbe);
            Release(documentContent);
            Release(tableAnchor);
            Release(formattedContent);
            Release(sourceContent);
            Release(suffixRange);
            Release(prefixRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void NormalizeNativeOmmlCenterCellParagraphs(
        Document document,
        Table table,
        Range formulaRange,
        string formulaId)
    {
        Cell? centerCell = null;
        Range? centerCellRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        Range? deleteRange = null;
        OMaths? refreshedMaths = null;
        OMath? refreshedMath = null;
        Range? refreshedRange = null;
        try
        {
            centerCell = table.Cell(1, 2);
            centerCellRange = centerCell.Range;
            paragraphs = centerCellRange.Paragraphs;
            if (paragraphs.Count <= 1) return;

            var formulaParagraphIndex = 0;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = null;
                Release(maths); maths = null;
                paragraph = paragraphs[index];
                paragraphRange = paragraph.Range.Duplicate;
                maths = paragraphRange.OMaths;
                var ownsFormula = formulaRange.Start >= paragraphRange.Start
                    && formulaRange.End <= paragraphRange.End;
                if (maths.Count > 0 || ownsFormula)
                {
                    if (maths.Count != 1 || formulaParagraphIndex != 0)
                        throw new InvalidOperationException(
                            "The legacy numbered OMML center cell contains more than one mathematical paragraph.");
                    formulaParagraphIndex = index;
                }
            }
            if (formulaParagraphIndex == 0)
                throw new InvalidOperationException(
                    "The legacy numbered OMML center cell has no paragraph owning the formula.");

            // Validate every extra paragraph before mutating anything. The final
            // paragraph of a Word table cell is mandatory and cannot itself be
            // deleted; once all extras are proven empty, normalization is performed
            // by deleting only the paragraph separators so the mathematical
            // paragraph merges into the cell's mandatory final paragraph.
            var expectedFormulaBookmark = WordOmmlFormulaStore.BookmarkName(formulaId);
            var detachFormulaIdentity = false;
            for (var index = paragraphs.Count; index >= 1; index--)
            {
                if (index == formulaParagraphIndex) continue;
                Release(frames); frames = null;
                Release(bookmarks); bookmarks = null;
                Release(fields); fields = null;
                Release(shapes); shapes = null;
                Release(maths); maths = null;
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = null;
                paragraph = paragraphs[index];
                paragraphRange = paragraph.Range.Duplicate;
                maths = paragraphRange.OMaths;
                shapes = paragraphRange.InlineShapes;
                fields = paragraphRange.Fields;
                bookmarks = paragraphRange.Bookmarks;
                frames = paragraphRange.Frames;
                var carriesOnlyFormulaIdentity = true;
                for (var bookmarkIndex = 1; bookmarkIndex <= bookmarks.Count; bookmarkIndex++)
                {
                    Bookmark? candidateBookmark = null;
                    try
                    {
                        candidateBookmark = bookmarks[bookmarkIndex];
                        if (!string.Equals(
                                candidateBookmark.Name,
                                expectedFormulaBookmark,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            carriesOnlyFormulaIdentity = false;
                            break;
                        }
                    }
                    finally { Release(candidateBookmark); }
                }
                if (maths.Count != 0
                    || shapes.Count != 0
                    || fields.Count != 0
                    || !carriesOnlyFormulaIdentity
                    || frames.Count != 0
                    || !ContainsOnlyStructuralWordText(paragraphRange.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to remove a non-empty legacy paragraph from the OMML number table center cell.");
                detachFormulaIdentity |= bookmarks.Count > 0;
            }
            if (detachFormulaIdentity)
                DeleteBookmarkOnly(document, expectedFormulaBookmark);

            while (true)
            {
                Release(paragraphs); paragraphs = null;
                Release(centerCellRange); centerCellRange = null;
                centerCellRange = centerCell.Range;
                paragraphs = centerCellRange.Paragraphs;
                if (paragraphs.Count <= 1) break;

                formulaParagraphIndex = 0;
                for (var index = 1; index <= paragraphs.Count; index++)
                {
                    Release(maths); maths = null;
                    Release(paragraphRange); paragraphRange = null;
                    Release(paragraph); paragraph = null;
                    paragraph = paragraphs[index];
                    paragraphRange = paragraph.Range.Duplicate;
                    maths = paragraphRange.OMaths;
                    if (maths.Count == 0) continue;
                    if (maths.Count != 1 || formulaParagraphIndex != 0)
                        throw new InvalidOperationException(
                            "The legacy numbered OMML center cell changed to multiple mathematical paragraphs while merging separators.");
                    formulaParagraphIndex = index;
                }
                if (formulaParagraphIndex == 0)
                    throw new InvalidOperationException(
                        "The legacy numbered OMML center formula disappeared while merging empty paragraphs.");

                // If an empty paragraph follows the formula, remove the formula
                // paragraph's final ¶ so it merges forward. If all empty paragraphs
                // precede the formula, remove the immediately preceding ¶ instead.
                var separatorOwnerIndex = formulaParagraphIndex < paragraphs.Count
                    ? formulaParagraphIndex
                    : formulaParagraphIndex - 1;
                if (separatorOwnerIndex < 1)
                    throw new InvalidOperationException(
                        "Word exposed multiple center-cell paragraphs without a removable paragraph separator.");
                Release(paragraphRange); paragraphRange = null;
                Release(paragraph); paragraph = null;
                paragraph = paragraphs[separatorOwnerIndex];
                paragraphRange = paragraph.Range.Duplicate;
                var separatorStart = Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
                deleteRange = document.Range(separatorStart, paragraphRange.End);
                if (!string.Equals(deleteRange.Text, "\r", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "The legacy OMML center-cell separator was not an ordinary paragraph mark.");
                deleteRange.Delete();
                Release(deleteRange); deleteRange = null;
            }

            Release(centerCellRange); centerCellRange = null;
            centerCellRange = centerCell.Range;
            refreshedMaths = centerCellRange.OMaths;
            if (refreshedMaths.Count != 1)
                throw new InvalidOperationException(
                    "Word did not retain exactly one center OMath after legacy paragraph cleanup.");
            refreshedMath = refreshedMaths[1];
            refreshedRange = refreshedMath.Range.Duplicate;
            formulaRange.SetRange(refreshedRange.Start, refreshedRange.End);

            Release(paragraphs); paragraphs = null;
            paragraphs = centerCellRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    $"The legacy OMML center cell did not converge to one paragraph; remaining={paragraphs.Count}.");
        }
        finally
        {
            Release(refreshedRange);
            Release(refreshedMath);
            Release(refreshedMaths);
            Release(deleteRange);
            Release(frames);
            Release(bookmarks);
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(centerCellRange);
            Release(centerCell);
        }
    }

    private static bool CleanupNativeOmmlTablePrecedingParagraph(
        Document document,
        string formulaId)
    {
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null) return false;
            rows = table.Rows;
            columns = table.Columns;
            if (rows.Count != 1 || columns.Count != 3)
                return false;
            RemoveEmptyBodyParagraphImmediatelyBeforeTable(
                document,
                table,
                formulaId);
            return true;
        }
        catch
        {
            // The direct-table host is already durable. A spacing cleanup failure
            // must never route that formula into retired caption/Shape cleanup or
            // turn successful insertion into a destructive failure.
            return true;
        }
        finally
        {
            Release(columns);
            Release(rows);
            Release(table);
        }
    }

    private static void RemoveEmptyBodyParagraphImmediatelyBeforeTable(
        Document document,
        Table table,
        string formulaId)
    {
        Range? tableRange = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphContent = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        Range? documentContent = null;
        Range? previousTableProbe = null;
        try
        {
            var traceSpacing = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                "1",
                StringComparison.Ordinal);
            tableRange = table.Range;
            if (traceSpacing)
                TraceNumberingPerformance(
                    $"[perf] cleanup-before-table enter formulaId={formulaId} table={tableRange.Start}:{tableRange.End}");
            if (tableRange.Start <= document.Content.Start)
            {
                if (traceSpacing) TraceNumberingPerformance("[perf] cleanup-before-table return-at-document-start");
                return;
            }
            probe = document.Range(tableRange.Start - 1, tableRange.Start);
            if ((bool)probe.get_Information(WdInformation.wdWithInTable))
            {
                if (traceSpacing) TraceNumberingPerformance("[perf] cleanup-before-table return-probe-in-table");
                return;
            }
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1)
            {
                if (traceSpacing) TraceNumberingPerformance($"[perf] cleanup-before-table return-paragraph-count={paragraphs.Count}");
                return;
            }
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (paragraphRange.End != tableRange.Start
                || !IsNumberingParagraphAdornment(paragraphRange.Text))
            {
                if (traceSpacing)
                    TraceNumberingPerformance(
                        $"[perf] cleanup-before-table return-paragraph-shape paragraph={paragraphRange.Start}:{paragraphRange.End} text='{(paragraphRange.Text ?? string.Empty).Replace("\r", "<CR>").Replace("\a", "<CELL>")}'");
                return;
            }
            // Word treats Paragraph.Range.End at a table boundary as having affinity
            // with the following table. Querying OMaths/Fields/InlineShapes on the
            // full one-character "\r" paragraph can therefore report objects that
            // actually live in the next table and falsely protect this generated
            // blank source paragraph. Inspect only the paragraph content before its
            // final paragraph mark; for a truly empty paragraph this is a collapsed
            // range at paragraphRange.Start with no affinity to the following table.
            paragraphContent = document.Range(
                paragraphRange.Start,
                Math.Max(paragraphRange.Start, paragraphRange.End - 1));
            shapes = paragraphContent.InlineShapes;
            maths = paragraphContent.OMaths;
            fields = paragraphContent.Fields;
            bookmarks = paragraphContent.Bookmarks;
            frames = paragraphContent.Frames;
            if (traceSpacing)
                TraceNumberingPerformance(
                    $"[perf] cleanup-before-table formulaId={formulaId} paragraph={paragraphRange.Start}:{paragraphRange.End} content={paragraphContent.Start}:{paragraphContent.End} shapes={shapes.Count} maths={maths.Count} fields={fields.Count} bookmarks={bookmarks.Count} frames={frames.Count}");
            if (shapes.Count != 0
                || maths.Count != 0
                || fields.Count != 0
                || frames.Count != 0)
            {
                if (traceSpacing) TraceNumberingPerformance("[perf] cleanup-before-table protected-by-content");
                return;
            }

            var expectedIdentity = WordOmmlFormulaStore.BookmarkName(formulaId);
            var carriesExpectedIdentity = false;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? candidate = null;
                try
                {
                    candidate = bookmarks[index];
                    if (!string.Equals(
                            candidate.Name,
                            expectedIdentity,
                            StringComparison.OrdinalIgnoreCase))
                        return;
                    carriesExpectedIdentity = true;
                }
                finally { Release(candidate); }
            }
            if (carriesExpectedIdentity)
                DeleteBookmarkOnly(document, expectedIdentity);

            // A paragraph between two Word tables is structural: deleting it makes
            // Word merge the two independent 1x3 hosts into one 2x3 table. This case
            // occurs when the user inserts the next numbered equation at the typing
            // paragraph immediately after an earlier numbered equation. Keep that
            // separator, but collapse it to a visually negligible 1pt zero-width
            // line. Ordinary empty body paragraphs (including the user's reported
            // blank line between text paragraphs) are still removed completely.
            documentContent = document.Content;
            if (paragraphRange.Start > documentContent.Start)
            {
                previousTableProbe = document.Range(
                    paragraphRange.Start - 1,
                    paragraphRange.Start);
                if ((bool)previousTableProbe.get_Information(
                        WdInformation.wdWithInTable))
                {
                    CompactNativeOmmlTableSeparatorParagraph(paragraphRange);
                    return;
                }
            }
            if (traceSpacing) TraceNumberingPerformance("[perf] cleanup-before-table delete-empty-source-paragraph");
            paragraphRange.Delete();
        }
        finally
        {
            Release(previousTableProbe);
            Release(documentContent);
            Release(frames);
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(paragraphContent);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(tableRange);
        }
    }

    internal static bool CompactManagedNativeOmmlTableSeparatorBefore(
        Document document,
        string formulaId)
    {
        Table? table = null;
        Range? tableRange = null;
        Range? separatorProbe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Tables? paragraphTables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        Range? previousTableProbe = null;
        Tables? previousTables = null;
        Table? previousTable = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null || table.Rows.Count != 1 || table.Columns.Count != 3)
                return false;
            tableRange = table.Range;
            if (tableRange.Start <= document.Content.Start) return false;

            separatorProbe = document.Range(tableRange.Start - 1, tableRange.Start);
            if ((bool)separatorProbe.get_Information(WdInformation.wdWithInTable))
                return false;
            paragraphs = separatorProbe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (paragraphRange.End != tableRange.Start
                || !IsNumberingParagraphAdornment(paragraphRange.Text))
                return false;

            paragraphTables = paragraphRange.Tables;
            shapes = paragraphRange.InlineShapes;
            maths = paragraphRange.OMaths;
            fields = paragraphRange.Fields;
            bookmarks = paragraphRange.Bookmarks;
            frames = paragraphRange.Frames;
            if (paragraphTables.Count != 0
                || shapes.Count != 0
                || maths.Count != 0
                || fields.Count != 0
                || bookmarks.Count != 0
                || frames.Count != 0
                || paragraphRange.Start <= document.Content.Start)
                return false;

            previousTableProbe = document.Range(
                paragraphRange.Start - 1,
                paragraphRange.Start);
            if (!(bool)previousTableProbe.get_Information(
                    WdInformation.wdWithInTable))
                return false;
            previousTables = previousTableProbe.Tables;
            if (previousTables.Count != 1) return false;
            previousTable = previousTables[1];
            if (!IsManagedNativeOmmlDirectTable(previousTable)) return false;

            CompactNativeOmmlTableSeparatorParagraph(paragraphRange);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(previousTable);
            Release(previousTables);
            Release(previousTableProbe);
            Release(frames);
            Release(bookmarks);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(paragraphTables);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(separatorProbe);
            Release(tableRange);
            Release(table);
        }
    }

    private static void RemoveGeneratedPostTableParagraphBeforeKnownContent(
        Document document,
        Table table)
    {
        Range? tableRange = null;
        Range? emptyProbe = null;
        Paragraphs? emptyParagraphs = null;
        Paragraph? emptyParagraph = null;
        Range? emptyRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        try
        {
            tableRange = table.Range;
            var contentEnd = document.Content.End;
            if (tableRange.End >= contentEnd) return;

            emptyProbe = document.Range(
                tableRange.End,
                Math.Min(contentEnd, tableRange.End + 1));
            if ((bool)emptyProbe.get_Information(WdInformation.wdWithInTable))
                return;
            emptyParagraphs = emptyProbe.Paragraphs;
            if (emptyParagraphs.Count != 1) return;
            emptyParagraph = emptyParagraphs[1];
            emptyRange = emptyParagraph.Range.Duplicate;
            if (emptyRange.Start != tableRange.End
                || !IsPlainNativeOmmlBodyParagraph(
                    document,
                    emptyRange,
                    allowCompactTailBookmark: false)
                || emptyRange.End >= contentEnd)
                return;

            // The caller recorded that the source formula was immediately followed
            // by a non-empty ordinary paragraph before Tables.Add. Therefore the
            // empty paragraph now sitting between the new 1x3 host and that content
            // is Word-generated insertion residue, not a user-authored blank line.
            // Re-prove the surviving next paragraph is still ordinary non-empty
            // content before deleting anything. A following table, another empty
            // paragraph, or the terminal document paragraph keeps the conservative
            // legacy behavior.
            nextProbe = document.Range(
                emptyRange.End,
                Math.Min(contentEnd, emptyRange.End + 1));
            if ((bool)nextProbe.get_Information(WdInformation.wdWithInTable))
                return;
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1) return;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            if (nextRange.Start != emptyRange.End
                || IsNumberingParagraphAdornment(nextRange.Text))
                return;

            emptyRange.Delete();
        }
        catch
        {
            // Numbering is already durable. Spacing cleanup must never make an
            // otherwise valid direct-SEQ formula fail.
        }
        finally
        {
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(emptyRange);
            Release(emptyParagraph);
            Release(emptyParagraphs);
            Release(emptyProbe);
            Release(tableRange);
        }
    }

    private static void CleanupGeneratedNativeOmmlTypingTailBeforeFollowingTable(
        Document document,
        string formulaId,
        Table? knownDirectTable = null)
    {
        var releaseTable = knownDirectTable is null;
        Table? table = knownDirectTable;
        Range? tableRange = null;
        Range? typingProbe = null;
        Paragraphs? typingParagraphs = null;
        Paragraph? typingParagraph = null;
        Range? typingRange = null;
        Range? separatorProbe = null;
        Paragraphs? separatorParagraphs = null;
        Paragraph? separatorParagraph = null;
        Range? separatorRange = null;
        Range? nextTableProbe = null;
        Tables? nextTables = null;
        Table? nextTable = null;
        try
        {
            if (table is null)
                table = FindNumberedEquationTable(document, formulaId);
            if (table is null || !IsManagedNativeOmmlDirectTable(table)) return;
            tableRange = table.Range;
            var contentEnd = document.Content.End;
            if (tableRange.End >= contentEnd) return;

            typingProbe = document.Range(
                tableRange.End,
                Math.Min(contentEnd, tableRange.End + 1));
            if ((bool)typingProbe.get_Information(WdInformation.wdWithInTable))
                return;
            typingParagraphs = typingProbe.Paragraphs;
            if (typingParagraphs.Count != 1) return;
            typingParagraph = typingParagraphs[1];
            typingRange = typingParagraph.Range.Duplicate;
            if (typingRange.Start != tableRange.End
                || !IsPlainNativeOmmlBodyParagraph(
                    document,
                    typingRange,
                    allowCompactTailBookmark: true))
                return;

            // A single paragraph after the last table is the user's ordinary typing
            // line and must remain. Remove it only when a second, already compacted
            // structural paragraph follows and that paragraph is immediately before
            // another managed direct-SEQ 1x3 table.
            if (typingRange.End >= contentEnd) return;
            separatorProbe = document.Range(
                typingRange.End,
                Math.Min(contentEnd, typingRange.End + 1));
            if ((bool)separatorProbe.get_Information(WdInformation.wdWithInTable))
                return;
            separatorParagraphs = separatorProbe.Paragraphs;
            if (separatorParagraphs.Count != 1) return;
            separatorParagraph = separatorParagraphs[1];
            separatorRange = separatorParagraph.Range.Duplicate;
            if (separatorRange.Start != typingRange.End
                || !IsPlainNativeOmmlBodyParagraph(
                    document,
                    separatorRange,
                    allowCompactTailBookmark: true)
                || !IsCompactNativeOmmlBodyParagraph(separatorRange)
                || separatorRange.End >= contentEnd)
                return;

            nextTableProbe = document.Range(
                separatorRange.End,
                Math.Min(contentEnd, separatorRange.End + 1));
            if (!(bool)nextTableProbe.get_Information(WdInformation.wdWithInTable))
                return;
            nextTables = nextTableProbe.Tables;
            if (nextTables.Count != 1) return;
            nextTable = nextTables[1];
            if (!IsManagedNativeOmmlDirectTable(nextTable)) return;

            typingRange.Delete();
        }
        catch
        {
            // The numbered tables and their mandatory compact separator are already
            // durable. Failure to remove one generated empty tail must not damage
            // either formula or route into a legacy numbering path.
        }
        finally
        {
            Release(nextTable);
            Release(nextTables);
            Release(nextTableProbe);
            Release(separatorRange);
            Release(separatorParagraph);
            Release(separatorParagraphs);
            Release(separatorProbe);
            Release(typingRange);
            Release(typingParagraph);
            Release(typingParagraphs);
            Release(typingProbe);
            Release(tableRange);
            if (releaseTable) Release(table);
        }
    }

    private static bool IsManagedNativeOmmlDirectTable(Table table)
    {
        Cell? centerCell = null;
        Cell? numberCell = null;
        Range? centerRange = null;
        Range? numberRange = null;
        OMaths? centerMaths = null;
        OMath? centerMath = null;
        Fields? numberFields = null;
        Field? numberField = null;
        Range? numberCode = null;
        try
        {
            if (table.Rows.Count != 1 || table.Columns.Count != 3) return false;
            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerRange = centerCell.Range;
            numberRange = numberCell.Range;
            centerMaths = centerRange.OMaths;
            if (centerMaths.Count != 1 || centerRange.Fields.Count != 0)
                return false;
            centerMath = centerMaths[1];
            if (centerMath.Type != WdOMathType.wdOMathDisplay) return false;
            if (numberRange.OMaths.Count != 0) return false;
            numberFields = numberRange.Fields;
            if (numberFields.Count != 1) return false;
            numberField = numberFields[1];
            numberCode = numberField.Code;
            return IsVisualTeXSequenceFieldCode(numberCode.Text);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(numberCode);
            Release(numberField);
            Release(numberFields);
            Release(centerMath);
            Release(centerMaths);
            Release(numberRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
        }
    }

    private static void CompactNativeOmmlTableSeparatorParagraph(
        Range paragraphRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                paragraphRange.set_Style(ref normalStyle);
            }
            catch { }
            font = paragraphRange.Font;
            try { font.Reset(); } catch { }
            font.Size = CompactTypingTailFontSizePoints;
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = paragraphRange.ParagraphFormat;
            try { format.Reset(); } catch { }
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            format.LineSpacing = CompactTypingTailLineSpacingPoints;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
        }
        finally
        {
            Release(format);
            Release(font);
        }
    }

    private static float? TryMeasureNativeDisplayHeightPoints(
        Document document,
        Range formulaRange)
    {
        Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        Zoom? zoom = null;
        try
        {
            window = document.ActiveWindow;
            // GetPoint succeeds for the formula currently being edited without
            // moving Word's viewport. The old path called ScrollIntoView followed
            // by Repaginate, which visibly jumped a long document to another page
            // (often through page 1) and dominated an otherwise local add-number
            // operation. Off-screen/compatibility-mode ranges simply fall back to
            // the semantic height estimate supplied by the caller.
            window.GetPoint(
                out _,
                out _,
                out _,
                out var heightPixels,
                formulaRange);
            view = window.View;
            zoom = view.Zoom;
            var zoomPercentage = zoom.Percentage;
            var dpi = 96u;
            try
            {
                var detected = GetDpiForWindow(new IntPtr(window.Hwnd));
                if (detected > 0) dpi = detected;
            }
            catch (EntryPointNotFoundException) { }
            if (heightPixels <= 0 || zoomPercentage <= 0 || dpi == 0)
                return null;
            var heightPoints = heightPixels
                * 72f
                * 100f
                / dpi
                / zoomPercentage;
            return heightPoints > 0f
                && !float.IsNaN(heightPoints)
                && !float.IsInfinity(heightPoints)
                    ? heightPoints
                    : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(zoom);
            Release(view);
            Release(window);
        }
    }

    private static bool NeedsNativeOmmlTableHeightRepair(Table table)
    {
        Rows? rows = null;
        Row? row = null;
        try
        {
            rows = table.Rows;
            if (rows.Count != 1) return false;
            row = rows[1];
            return row.HeightRule == WdRowHeightRule.wdRowHeightAuto
                || row.Height <= 0f
                || row.Height >= 1000000f;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(row);
            Release(rows);
        }
    }

    internal static void ApplyNativeOmmlTableMinimumDisplayHeight(
        Table table,
        float minimumDisplayHeightPoints)
    {
        Rows? rows = null;
        try
        {
            rows = table.Rows;
            ApplyNativeOmmlTableMinimumDisplayHeight(
                rows,
                minimumDisplayHeightPoints);
        }
        finally { Release(rows); }
    }

    private static void ApplyNativeOmmlTableMinimumDisplayHeight(
        Rows rows,
        float? minimumDisplayHeightPoints)
    {
        if (!minimumDisplayHeightPoints.HasValue
            || float.IsNaN(minimumDisplayHeightPoints.Value)
            || float.IsInfinity(minimumDisplayHeightPoints.Value)
            || minimumDisplayHeightPoints.Value <= 0f)
            return;
        Row? row = null;
        try
        {
            if (rows.Count != 1) return;
            row = rows[1];
            row.HeightRule = WdRowHeightRule.wdRowHeightAtLeast;
            row.Height = Math.Max(
                1f,
                minimumDisplayHeightPoints.Value
                + NativeOmmlTableHeightSafetyPoints);
        }
        finally { Release(row); }
    }

    private static void ConfigureNativeOmmlNumberTableGeometry(
        Document document,
        Table table,
        Range formulaRange,
        float? minimumDisplayHeightPoints = null)
    {
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        Rows? rows = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        Cell? leftCell = null;
        Cell? centerCell = null;
        Cell? rightCell = null;
        Range? centerRange = null;
        Range? rightRange = null;
        ParagraphFormat? centerFormat = null;
        ParagraphFormat? rightFormat = null;
        TabStops? rightTabs = null;
        TabStop? rightTab = null;
        Borders? borders = null;
        try
        {
            sections = formulaRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            var writableWidth = Math.Max(
                120f,
                pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
            var sideWidth = Math.Min(
                NativeOmmlTableSideWidthPoints,
                Math.Max(24f, (writableWidth - 96f) / 2f));
            var centerWidth = Math.Max(96f, writableWidth - 2f * sideWidth);

            table.AllowAutoFit = false;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            table.TopPadding = 0f;
            table.BottomPadding = 0f;
            try { table.Spacing = 0f; } catch { }
            try { table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed); } catch { }
            borders = table.Borders;
            borders.Enable = 0;

            rows = table.Rows;
            // The fixed-width host starts at the section's writable text boundary.
            // Keep the row explicitly left-aligned; Word 2021 can keep a newly
            // converted table at page X=0 for the current live layout when center
            // alignment is applied before the former MathType paragraph has fully
            // settled, even though the saved tblW is correct.
            try { rows.Alignment = WdRowAlignment.wdAlignRowLeft; } catch { }
            try { rows.SetLeftIndent(0f, WdRulerStyle.wdAdjustNone); } catch { }
            try { rows.AllowBreakAcrossPages = 0; } catch { }
            ApplyNativeOmmlTableMinimumDisplayHeight(
                rows,
                minimumDisplayHeightPoints);

            columns = table.Columns;
            if (columns.Count != 3 || rows.Count != 1)
                throw new InvalidOperationException(
                    "The native OMML number host is no longer exactly 1x3.");
            leftColumn = columns[1];
            centerColumn = columns[2];
            rightColumn = columns[3];
            leftColumn.SetWidth(sideWidth, WdRulerStyle.wdAdjustNone);
            centerColumn.SetWidth(centerWidth, WdRulerStyle.wdAdjustNone);
            rightColumn.SetWidth(sideWidth, WdRulerStyle.wdAdjustNone);

            leftCell = table.Cell(1, 1);
            centerCell = table.Cell(1, 2);
            rightCell = table.Cell(1, 3);
            leftCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            rightCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            centerRange = centerCell.Range;
            rightRange = rightCell.Range;
            centerFormat = centerRange.ParagraphFormat;
            rightFormat = rightRange.ParagraphFormat;
            centerFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rightFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            centerFormat.LeftIndent = centerFormat.RightIndent = 0f;
            centerFormat.FirstLineIndent = 0f;
            rightFormat.LeftIndent = rightFormat.RightIndent = 0f;
            rightFormat.FirstLineIndent = 0f;
            centerFormat.SpaceBefore = centerFormat.SpaceAfter = 0f;
            rightFormat.SpaceBefore = rightFormat.SpaceAfter = 0f;
            centerFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            rightFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            try { centerFormat.DisableLineHeightGrid = -1; } catch { }
            try { rightFormat.DisableLineHeightGrid = -1; } catch { }

            // The number cell deliberately uses a *real* right TabStop. Because the
            // three-column host is symmetric and has zero padding, this tab lands at
            // the same body-right coordinate as a MathType/VisualTeX OLE right tab,
            // while the center cell midpoint remains exactly at bodyWidth / 2.
            rightTabs = rightFormat.TabStops;
            rightTabs.ClearAll();
            rightTab = rightTabs.Add(
                rightColumn.Width,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);

            // Word can carry a stale MTDisplayEquation layout origin through the
            // source paragraph while MathType→OMML is still inside its conversion
            // transaction. Recommit the completed table's inline-row positioning
            // after widths, cells and tab stops are final. This writes no floating
            // table properties; it only forces the live row back to the body text
            // boundary with zero direct indent.
            try { rows.WrapAroundText = 0; } catch { }
            try { rows.LeftIndent = 0f; } catch { }
            try { rows.SetLeftIndent(0f, WdRulerStyle.wdAdjustNone); } catch { }
            try { rows.Alignment = WdRowAlignment.wdAlignRowLeft; } catch { }

            // AutoFitBehavior, SetWidth, and late row normalization can reset the
            // aggregate table width to w:type="auto" on a paragraph that used to
            // host MathType. Persist tblW only after every grid/layout mutation so
            // the final DOCX stores the exact body width (for the acceptance page,
            // 8306 twips / 415.3pt) instead of an auto-width table at page X=0.
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            table.PreferredWidth = writableWidth;
        }
        finally
        {
            Release(borders);
            Release(rightTab);
            Release(rightTabs);
            Release(rightFormat);
            Release(centerFormat);
            Release(rightRange);
            Release(centerRange);
            Release(rightCell);
            Release(centerCell);
            Release(leftCell);
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(rows);
            Release(pageSetup);
            Release(section);
            Release(sections);
        }
    }

    private static void EnsureDirectTableSequenceNumber(
        Document document,
        Table table,
        Range formulaRange,
        string formulaId,
        float formulaFontSizePoints,
        int ordinal,
        string prefix)
    {
        Cell? numberCell = null;
        Range? cellRange = null;
        Range? editableRange = null;
        Range? scaffoldRange = null;
        Range? fieldRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? fieldCode = null;
        Range? fieldResult = null;
        Range? captionRange = null;
        Range? numberRange = null;
        Range? labelRange = null;
        Range? closingRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            numberCell = table.Cell(1, 3);
            cellRange = numberCell.Range.Duplicate;
            editableRange = cellRange.Duplicate;
            editableRange.End = Math.Max(editableRange.Start, editableRange.End - 1);
            editableRange.Text = string.Empty;

            scaffoldRange = document.Range(editableRange.Start, editableRange.Start);
            scaffoldRange.Text = "\t(" + prefix;
            var labelStart = editableRange.Start + 1;
            var prefixStart = labelStart + 1;
            var fieldStart = editableRange.Start + scaffoldRange.Text!.Length;
            fieldRange = document.Range(fieldStart, fieldStart);
            fields = fieldRange.Fields;
            object fieldType = WdFieldType.wdFieldEmpty;
            object fieldText = $"SEQ {LegacyEquationSequenceName} \\r {Math.Max(1, ordinal)} \\* ARABIC";
            object preserveFormatting = true;
            field = fields.Add(
                fieldRange,
                ref fieldType,
                ref fieldText,
                ref preserveFormatting);
            field.Update();
            fieldResult = field.Result;
            var closingPosition = Math.Min(
                Math.Max(fieldResult.End + 1, fieldStart + 1),
                Math.Max(editableRange.Start, numberCell.Range.End - 1));
            closingRange = document.Range(closingPosition, closingPosition);
            closingRange.Text = ")";

            Release(fieldCode);
            fieldCode = field.Code;
            Release(fieldResult);
            fieldResult = field.Result;
            Release(cellRange);
            cellRange = numberCell.Range.Duplicate;
            var cellEditableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            var fieldBegin = Math.Max(prefixStart, fieldCode.Start - 1);
            var fieldEnd = Math.Min(cellEditableEnd, fieldResult.End + 1);
            var labelEnd = Math.Min(cellEditableEnd, fieldEnd + 1);
            captionRange = document.Range(prefixStart, fieldEnd);
            numberRange = document.Range(prefixStart, fieldResult.End);
            labelRange = document.Range(labelStart, labelEnd);
            if (!(labelRange.Text ?? string.Empty).StartsWith("(", StringComparison.Ordinal)
                || !(labelRange.Text ?? string.Empty).EndsWith(")", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word did not preserve the complete direct table equation number label.");

            bookmarks = document.Bookmarks;
            foreach (var name in new[]
                     {
                         EquationBookmarkName(formulaId),
                         NativeCaptionBookmarkName(formulaId),
                         NativeNumberBookmarkName(formulaId),
                     })
            {
                if (!bookmarks.Exists(name)) continue;
                Release(bookmark);
                bookmark = bookmarks[name];
                bookmark.Delete();
            }
            Release(bookmark);
            bookmark = null;
            bookmark = bookmarks.Add(NativeCaptionBookmarkName(formulaId), captionRange);
            Release(bookmark); bookmark = null;
            bookmark = bookmarks.Add(NativeNumberBookmarkName(formulaId), numberRange);
            Release(bookmark); bookmark = null;
            bookmark = bookmarks.Add(EquationBookmarkName(formulaId), labelRange);
            Release(bookmark); bookmark = null;

            ApplyParagraphEquationNumberFont(
                labelRange,
                formulaFontSizePoints,
                position: 0);

            // Keep the right cell structurally sterile: exactly one real TAB,
            // one visible SEQ label and the cell's own final paragraph/cell marks.
            OMaths? numberMaths = null;
            InlineShapes? numberShapes = null;
            try
            {
                numberMaths = cellRange.OMaths;
                numberShapes = cellRange.InlineShapes;
                if (numberMaths.Count != 0 || numberShapes.Count != 0)
                    throw new InvalidOperationException(
                        "The direct OMML number cell contains a mathematical or embedded object.");
            }
            finally
            {
                Release(numberShapes);
                Release(numberMaths);
            }
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(closingRange);
            Release(labelRange);
            Release(numberRange);
            Release(captionRange);
            Release(fieldResult);
            Release(fieldCode);
            Release(field);
            Release(fields);
            Release(fieldRange);
            Release(scaffoldRange);
            Release(editableRange);
            Release(cellRange);
            Release(numberCell);
        }
    }

    private static bool TryUpdateDirectTableSequenceNumber(
        Document document,
        string formulaId,
        string nativeSequenceName,
        int ordinal,
        string prefix,
        bool formatOnly,
        Bookmarks? knownBookmarks = null,
        bool trustedHealthyDirectTable = false)
    {
        _ = formatOnly;
        var releaseBookmarks = knownBookmarks is null;
        Table? table = null;
        Cell? numberCell = null;
        Range? cellRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Bookmark? visibleBookmark = null;
        Range? captionRange = null;
        Range? visibleRange = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? prefixRange = null;
        Range? closing = null;
        try
        {
            bookmarks = knownBookmarks ?? document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)
                || !bookmarks.Exists(numberName)
                || !bookmarks.Exists(visibleName))
                return false;
            visibleBookmark = bookmarks[visibleName];
            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            visibleRange = visibleBookmark.Range;
            captionRange = captionBookmark.Range;
            var numberIdentityRange = numberBookmark.Range;
            var bookmarkTopologyProvesDirectNumber = false;
            try
            {
                var visibleText = visibleRange.Text ?? string.Empty;
                bookmarkTopologyProvesDirectNumber =
                    visibleRange.StoryType == WdStoryType.wdMainTextStory
                    && captionRange.StoryType == WdStoryType.wdMainTextStory
                    && numberIdentityRange.StoryType == WdStoryType.wdMainTextStory
                    && visibleRange.Start <= captionRange.Start
                    && visibleRange.End >= captionRange.End
                    && captionRange.Start == numberIdentityRange.Start
                    && captionRange.End >= numberIdentityRange.End
                    && visibleText.StartsWith("(", StringComparison.Ordinal)
                    && visibleText.EndsWith(")", StringComparison.Ordinal);
            }
            finally { Release(numberIdentityRange); }

            // `trustedHealthyDirectTable` is a batch-level optimization hint, not
            // per-formula proof that every healthy caption belongs to a native OMML
            // 1x3 direct-SEQ host. A mixed document can contain a healthy VisualTeX
            // OLE hidden caption in the same OpenXML inventory. Only the nested
            // VTEq/VTEqCap/VTEqNum bookmark topology proves that this particular
            // formula may use the table-specific range rewrite without COM table
            // discovery. Otherwise return false so the caller uses the ordinary OLE
            // hidden-caption updater.
            if (trustedHealthyDirectTable && !bookmarkTopologyProvesDirectNumber)
                return false;
            var useDirectBookmarkTopology = bookmarkTopologyProvesDirectNumber;
            if (!useDirectBookmarkTopology)
            {
                if (!(bool)visibleRange.get_Information(WdInformation.wdWithInTable)
                    || visibleRange.Tables.Count == 0)
                    return false;
                table = visibleRange.Tables[1];
                if (table.Rows.Count != 1 || table.Columns.Count != 3)
                    return false;
                numberCell = table.Cell(1, 3);
                cellRange = numberCell.Range.Duplicate;
                if (visibleRange.Start < cellRange.Start || visibleRange.End > cellRange.End)
                    return false;
            }
            else if (visibleRange.StoryType != WdStoryType.wdMainTextStory)
            {
                return false;
            }
            field = FindNativeEquationFieldInRange(captionRange, nativeSequenceName);
            if (field is null) return false;
            code = field.Code;
            result = field.Result;
            var expectedOrdinal = Math.Max(1, ordinal).ToString(
                System.Globalization.CultureInfo.InvariantCulture);
            var codeText = code.Text ?? string.Empty;
            var restartMatches = System.Text.RegularExpressions.Regex.IsMatch(
                codeText,
                $@"\\r\s+{System.Text.RegularExpressions.Regex.Escape(expectedOrdinal)}\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
                | System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            var resultMatches = string.Equals(
                NormalizeNativeEquationNumberText(result.Text),
                expectedOrdinal,
                StringComparison.Ordinal);
            var fieldInstructionChanged = false;
            var fieldResultMayHaveChanged = false;
            if (!IsNativeEquationSequenceFieldCode(codeText, nativeSequenceName)
                || !restartMatches)
            {
                EnsureSequenceFieldCodeCanBeRewritten(field);
                code.Text = $" SEQ {LegacyEquationSequenceName} \\r {expectedOrdinal} \\* ARABIC ";
                field.Update();
                fieldInstructionChanged = true;
                fieldResultMayHaveChanged = true;
            }
            else if (!resultMatches)
            {
                // A format-only switch normally changes the literal heading
                // prefix, not the SEQ ordinal. Avoiding an unnecessary Field.Update
                // is important in large Word documents, but still refresh a stale
                // result when Word has not evaluated the existing instruction.
                field.Update();
                fieldResultMayHaveChanged = true;
            }

            // Preserve the cell's leading TAB and opening parenthesis. Only the
            // literal heading prefix between '(' and the field-begin control is
            // replaced when the numbering format changes.
            Release(code);
            code = field.Code;
            var labelStart = useDirectBookmarkTopology
                ? visibleRange.Start
                : cellRange!.Start + 1;
            var prefixStart = labelStart + 1;
            var fieldBegin = Math.Max(prefixStart, code.Start - 1);
            prefixRange = document.Range(prefixStart, fieldBegin);
            var previousPrefix = prefixRange.Text ?? string.Empty;
            var prefixLengthChanged = previousPrefix.Length != prefix.Length;
            if (!string.Equals(previousPrefix, prefix, StringComparison.Ordinal))
                prefixRange.Text = prefix;

            Release(result);
            result = field.Result;
            var editableEnd = int.MaxValue;
            if (!useDirectBookmarkTopology)
            {
                Release(cellRange);
                cellRange = numberCell!.Range.Duplicate;
                editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            }
            var closePosition = useDirectBookmarkTopology
                ? result.End + 1
                : Math.Min(editableEnd, result.End + 1);
            closing = useDirectBookmarkTopology
                ? document.Range(closePosition, closePosition + 1)
                : document.Range(closePosition, Math.Min(editableEnd, closePosition + 1));
            var closingChanged = false;
            if (!string.Equals(closing.Text, ")", StringComparison.Ordinal))
            {
                closing.SetRange(closePosition, closePosition);
                closing.Text = ")";
                closingChanged = true;
            }

            // The common format switch changes only one same-width separator in
            // the literal prefix (for example 2.1. -> 2.1-). The three existing
            // bookmarks already surround the same character positions and Word
            // preserves them across an equal-length text replacement. Rebuilding
            // all three aliases per formula was the dominant cost of changing a
            // five-formula document's number format.
            if (!fieldInstructionChanged
                && !fieldResultMayHaveChanged
                && !prefixLengthChanged
                && !closingChanged)
                return true;

            DeleteBookmarkOnly(document, captionName);
            DeleteBookmarkOnly(document, numberName);
            DeleteBookmarkOnly(document, visibleName);
            Release(code);
            code = field.Code;
            Release(result);
            result = field.Result;
            if (!useDirectBookmarkTopology)
            {
                Release(cellRange);
                cellRange = numberCell!.Range.Duplicate;
                editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            }
            var fieldEnd = useDirectBookmarkTopology
                ? result.End + 1
                : Math.Min(editableEnd, result.End + 1);
            var labelEnd = useDirectBookmarkTopology
                ? fieldEnd + 1
                : Math.Min(editableEnd, fieldEnd + 1);
            var newCaption = document.Range(prefixStart, fieldEnd);
            var newNumber = document.Range(prefixStart, result.End);
            var newVisible = document.Range(labelStart, labelEnd);
            try
            {
                bookmarks.Add(captionName, newCaption);
                bookmarks.Add(numberName, newNumber);
                bookmarks.Add(visibleName, newVisible);
                ApplyParagraphEquationNumberFont(newVisible, fallbackSize: 11f, position: 0);
            }
            finally
            {
                Release(newVisible);
                Release(newNumber);
                Release(newCaption);
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(closing);
            Release(prefixRange);
            Release(result);
            Release(code);
            Release(field);
            Release(visibleRange);
            Release(captionRange);
            Release(visibleBookmark);
            Release(numberBookmark);
            Release(captionBookmark);
            if (releaseBookmarks) Release(bookmarks);
            Release(cellRange);
            Release(numberCell);
            Release(table);
        }
    }

    internal static bool HasReusableNumberedNativeOmmlDirectTableHost(
        Document document,
        Range formulaRange,
        string formulaId) =>
        IsHealthyNativeOmmlDirectTableHost(
            document,
            formulaRange,
            formulaId,
            updateField: false);

    internal static bool HasReusableNumberedNativeOmmlDirectTableCenterHost(
        Table knownTable,
        Range formulaRange)
    {
        Rows? rows = null;
        Columns? columns = null;
        Cell? centerCell = null;
        Range? centerRange = null;
        Fields? centerFields = null;
        OMaths? centerMaths = null;
        OMath? centerMath = null;
        Range? centerMathRange = null;
        try
        {
            bool Fail(string reason)
            {
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE"),
                        "1",
                        StringComparison.Ordinal))
                    TraceNumberingPerformance($"[perf] direct-center-health-fail {reason}");
                return false;
            }

            // A same-numbering content edit mutates only cell (1,2). The direct
            // VTEq_ lookup already resolved the owning table through the right-cell
            // visible-number bookmark, so re-walking VTEq_/VTEqCap_/VTEqNum_, the
            // SEQ field and right-cell text is redundant and costs ~150-200ms in a
            // 100-OMML document. Prove exactly the part we are about to replace:
            // genuine 1x3 topology, one field-free Display OMath, and no prefix or
            // suffix content between that OMath and the center cell markers.
            rows = knownTable.Rows;
            columns = knownTable.Columns;
            if (rows.Count != 1 || columns.Count != 3)
                return Fail($"shape={rows.Count}x{columns.Count}");

            centerCell = knownTable.Cell(1, 2);
            centerRange = centerCell.Range;
            // Word exposes the textual cell terminator as "\r\a", but those two
            // control glyphs occupy one story position. A formula that exactly
            // fills the center cell therefore ends at centerRange.End - 1.
            if (centerRange.End - centerRange.Start < 1
                || formulaRange.Start != centerRange.Start
                || formulaRange.End != centerRange.End - 1)
                return Fail(
                    $"bounds formula={formulaRange.Start}:{formulaRange.End} center={centerRange.Start}:{centerRange.End}");

            centerFields = centerRange.Fields;
            if (centerFields.Count != 0) return Fail($"fields={centerFields.Count}");
            centerMaths = centerRange.OMaths;
            if (centerMaths.Count != 1) return Fail($"maths={centerMaths.Count}");
            centerMath = centerMaths[1];
            if (centerMath.Type != WdOMathType.wdOMathDisplay)
                return Fail($"type={centerMath.Type}");
            centerMathRange = centerMath.Range;
            if (centerMathRange.Start != formulaRange.Start
                || centerMathRange.End != formulaRange.End)
                return Fail(
                    $"math-bounds formula={formulaRange.Start}:{formulaRange.End} math={centerMathRange.Start}:{centerMathRange.End}");
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(centerMathRange);
            Release(centerMath);
            Release(centerMaths);
            Release(centerFields);
            Release(centerRange);
            Release(centerCell);
            Release(columns);
            Release(rows);
        }
    }

    private static bool IsHealthyNativeOmmlDirectTableHost(
        Document document,
        Range formulaRange,
        string formulaId,
        bool updateField)
    {
        Tables? tables = null;
        Table? table = null;
        Cell? centerCell = null;
        Cell? numberCell = null;
        Range? centerRange = null;
        Range? numberRange = null;
        OMaths? centerMaths = null;
        OMath? centerMath = null;
        Range? centerMathRange = null;
        Range? centerPrefix = null;
        Range? centerSuffix = null;
        Fields? numberFields = null;
        Field? sequenceField = null;
        Range? sequenceCode = null;
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Bookmark? numberBookmark = null;
        Bookmark? captionBookmark = null;
        Range? visibleRange = null;
        Range? identityRange = null;
        Range? captionRange = null;
        try
        {
            if (!(bool)formulaRange.get_Information(WdInformation.wdWithInTable))
                return false;
            tables = formulaRange.Tables;
            if (tables.Count != 1) return false;
            table = tables[1];
            if (table.Rows.Count != 1 || table.Columns.Count != 3) return false;
            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerRange = centerCell.Range;
            numberRange = numberCell.Range;
            centerMaths = centerRange.OMaths;
            if (centerMaths.Count != 1 || centerRange.Fields.Count != 0) return false;
            centerMath = centerMaths[1];
            if (centerMath.Type != WdOMathType.wdOMathDisplay) return false;
            centerMathRange = centerMath.Range.Duplicate;
            if (formulaRange.Start != centerMathRange.Start
                || formulaRange.End != centerMathRange.End)
                return false;
            centerPrefix = document.Range(centerRange.Start, centerMathRange.Start);
            centerSuffix = document.Range(centerMathRange.End, centerRange.End);
            if (!string.IsNullOrEmpty(centerPrefix.Text)
                || !string.Equals(
                    centerSuffix.Text,
                    "\r\a",
                    StringComparison.Ordinal))
                return false;
            if (numberRange.OMaths.Count != 0) return false;
            numberFields = numberRange.Fields;
            if (numberFields.Count != 1) return false;
            sequenceField = numberFields[1];
            sequenceCode = sequenceField.Code;
            if (!IsVisualTeXSequenceFieldCode(sequenceCode.Text)) return false;

            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)
                || !bookmarks.Exists(numberName)
                || !bookmarks.Exists(captionName))
                return false;
            visibleBookmark = bookmarks[visibleName];
            numberBookmark = bookmarks[numberName];
            captionBookmark = bookmarks[captionName];
            visibleRange = visibleBookmark.Range;
            identityRange = numberBookmark.Range;
            captionRange = captionBookmark.Range;
            foreach (var owned in new[] { visibleRange, identityRange, captionRange })
            {
                if (!(bool)owned.get_Information(WdInformation.wdWithInTable)) return false;
                if (owned.Start < numberRange.Start || owned.End > numberRange.End) return false;
            }
            var visibleText = visibleRange.Text ?? string.Empty;
            if (!visibleText.StartsWith("(", StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;
            var cellText = numberRange.Text ?? string.Empty;
            if (!cellText.StartsWith("\t", StringComparison.Ordinal)
                || !cellText.EndsWith("\r\a", StringComparison.Ordinal))
                return false;
            if (updateField)
            {
                // MathType→OMML can retain the source paragraph's live layout
                // origin until the enclosing custom Undo transaction is committed.
                // Reapply the already-proven 1x3 geometry here, after that commit,
                // so the table is anchored to the real body text boundary rather
                // than page X=0. Ordinary insertion/edit takes the same idempotent
                // targeted path and does not enumerate any sibling formulas.
                ConfigureNativeOmmlNumberTableGeometry(
                    document,
                    table,
                    centerMathRange);
                sequenceField.Update();
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(captionRange);
            Release(identityRange);
            Release(visibleRange);
            Release(captionBookmark);
            Release(numberBookmark);
            Release(visibleBookmark);
            Release(bookmarks);
            Release(sequenceCode);
            Release(sequenceField);
            Release(numberFields);
            Release(centerSuffix);
            Release(centerPrefix);
            Release(centerMathRange);
            Release(centerMath);
            Release(centerMaths);
            Release(numberRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
            Release(table);
            Release(tables);
        }
    }

    internal static bool TryRedirectManagedNativeOmmlNumberEndEnter(
        Selection selection)
    {
        if (selection is null || selection.Start != selection.End) return false;

        Document? document = null;
        Range? selectionRange = null;
        Tables? selectionTables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Cell? numberCell = null;
        Range? numberCellRange = null;
        Paragraphs? numberParagraphs = null;
        Paragraph? trailingParagraph = null;
        Range? trailingRange = null;
        Paragraph? extraParagraph = null;
        Range? extraRange = null;
        Fields? numberFields = null;
        Field? sequenceField = null;
        Range? sequenceCode = null;
        Cell? centerCell = null;
        Range? centerRange = null;
        OMaths? centerMaths = null;
        OMath? centerMath = null;
        Range? centerMathRange = null;
        Range? typingRange = null;
        var stage = "validate-selection";
        try
        {
            document = selection.Document;
            selectionRange = selection.Range.Duplicate;
            if (!(bool)selectionRange.get_Information(WdInformation.wdWithInTable))
                return false;
            selectionTables = selectionRange.Tables;
            if (selectionTables.Count != 1) return false;
            table = selectionTables[1];
            rows = table.Rows;
            columns = table.Columns;
            if (rows.Count != 1 || columns.Count != 3) return false;

            numberCell = table.Cell(1, 3);
            numberCellRange = numberCell.Range.Duplicate;
            if (selection.Start < numberCellRange.Start
                || selection.Start > numberCellRange.End)
                return false;
            numberParagraphs = numberCellRange.Paragraphs;
            if (numberParagraphs.Count <= 1) return false;

            trailingParagraph = numberParagraphs[numberParagraphs.Count];
            trailingRange = trailingParagraph.Range.Duplicate;
            if (selection.Start < trailingRange.Start
                || selection.Start > trailingRange.End
                || !ContainsOnlyStructuralWordText(trailingRange.Text))
                return false;

            // Word's native Enter at the end of the right-cell label splits that
            // cell into two paragraphs and extends the SEQ/bookmark tree across the
            // new empty paragraph. Accept only that exact structural mutation; a
            // user who has typed visible text in any extra paragraph is left alone.
            for (var index = 2; index <= numberParagraphs.Count; index++)
            {
                Release(extraRange); extraRange = null;
                Release(extraParagraph); extraParagraph = null;
                extraParagraph = numberParagraphs[index];
                extraRange = extraParagraph.Range.Duplicate;
                if (!ContainsOnlyStructuralWordText(extraRange.Text)
                    || extraRange.OMaths.Count != 0
                    || extraRange.InlineShapes.Count != 0)
                    return false;
            }

            numberFields = numberCellRange.Fields;
            if (numberFields.Count != 1) return false;
            sequenceField = numberFields[1];
            sequenceCode = sequenceField.Code;
            if (!IsVisualTeXSequenceFieldCode(sequenceCode.Text)) return false;
            if (!TryResolveDirectTableFormulaIdFromNumberCell(
                    document,
                    numberCellRange,
                    out var formulaId))
                return false;
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is null
                || !metadata.Numbered
                || !string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase))
                return false;

            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range.Duplicate;
            centerMaths = centerRange.OMaths;
            if (centerMaths.Count != 1 || centerRange.Fields.Count != 0)
                return false;
            centerMath = centerMaths[1];
            if (centerMath.Type != WdOMathType.wdOMathDisplay) return false;
            centerMathRange = centerMath.Range.Duplicate;

            stage = "resolve-number-plan";
            var plan = ResolveDirectTableNumberPlan(
                document,
                centerMathRange,
                formulaId,
                plannedOrdinal: null,
                plannedPrefix: null);
            stage = "rebuild-number-cell";
            EnsureDirectTableSequenceNumber(
                document,
                table,
                centerMathRange,
                formulaId,
                (float)FormulaFontSize.ResolveSemanticFontSize(metadata),
                plan.Ordinal,
                plan.Prefix);
            stage = "configure-table-geometry";
            ConfigureNativeOmmlNumberTableGeometry(
                document,
                table,
                centerMathRange);
            stage = "validate-rebuilt-host";
            if (!IsHealthyNativeOmmlDirectTableHost(
                    document,
                    centerMathRange,
                    formulaId,
                    updateField: false))
                throw new InvalidOperationException(
                    "Word did not restore the direct-SEQ number cell after Enter.");

            stage = "create-body-typing-paragraph";
            typingRange = EnsureDeletableTypingParagraphAfterNativeOmmlTable(
                document,
                table);
            if (typingRange is null)
                throw new InvalidOperationException(
                    "Word did not expose an ordinary typing paragraph after the numbered OMML table.");
            stage = "move-selection";
            selection.SetRange(typingRange.Start, typingRange.Start);
            selection.Collapse(WdCollapseDirection.wdCollapseStart);
            return true;
        }
        catch (Exception error)
        {
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
                Console.WriteLine(
                    $"[NUMBER END ENTER REPAIR FAILED] stage={stage} type={error.GetType().Name} hresult=0x{error.HResult:X8} message={error.Message}");
            // This is a narrowly scoped user-input repair. Never turn an unusual
            // table or protected document into a SelectionChange failure; the
            // ordinary Word behavior remains available when validation is not exact.
            return false;
        }
        finally
        {
            Release(typingRange);
            Release(centerMathRange);
            Release(centerMath);
            Release(centerMaths);
            Release(centerRange);
            Release(centerCell);
            Release(sequenceCode);
            Release(sequenceField);
            Release(numberFields);
            Release(extraRange);
            Release(extraParagraph);
            Release(trailingRange);
            Release(trailingParagraph);
            Release(numberParagraphs);
            Release(numberCellRange);
            Release(numberCell);
            Release(columns);
            Release(rows);
            Release(table);
            Release(selectionTables);
            Release(selectionRange);
            Release(document);
        }
    }

    private static bool TryResolveDirectTableFormulaIdFromNumberCell(
        Document document,
        Range numberCellRange,
        out string formulaId)
    {
        formulaId = string.Empty;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(range); range = null;
                Release(bookmark); bookmark = bookmarks[index];
                var name = bookmark.Name;
                if (!TryFormulaIdFromEquationBookmark(name, out var candidate)
                    && !TryFormulaIdFromBookmark(
                        name,
                        NativeNumberBookmarkPrefix,
                        out candidate)
                    && !TryFormulaIdFromBookmark(
                        name,
                        NativeCaptionBookmarkPrefix,
                        out candidate))
                    continue;
                range = bookmark.Range;
                if (range.Start < numberCellRange.Start
                    || range.End > numberCellRange.End)
                    continue;
                if (formulaId.Length > 0
                    && !string.Equals(
                        formulaId,
                        candidate,
                        StringComparison.OrdinalIgnoreCase))
                    return false;
                formulaId = candidate;
            }
            return formulaId.Length > 0;
        }
        catch
        {
            formulaId = string.Empty;
            return false;
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Range? EnsureDeletableTypingParagraphAfterNativeOmmlTable(
        Document document,
        Table table)
    {
        Range? tableRange = null;
        Range? immediateProbe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? followingProbe = null;
        Paragraphs? followingParagraphs = null;
        Paragraph? followingParagraph = null;
        Range? followingRange = null;
        try
        {
            tableRange = table.Range;
            var paragraphStart = tableRange.End;
            var contentEnd = document.Content.End;
            if (paragraphStart >= contentEnd) return null;
            immediateProbe = document.Range(
                paragraphStart,
                Math.Min(contentEnd, paragraphStart + 1));
            if ((bool)immediateProbe.get_Information(WdInformation.wdWithInTable))
                return null;
            paragraphs = immediateProbe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (paragraphRange.Start != paragraphStart) return null;

            var immediateIsEmpty = IsPlainNativeOmmlBodyParagraph(
                document,
                paragraphRange,
                allowCompactTailBookmark: true);
            var immediateIsCompact = immediateIsEmpty
                && IsCompactNativeOmmlBodyParagraph(paragraphRange);
            var followingStartsTable = false;
            if (paragraphRange.End < contentEnd)
            {
                followingProbe = document.Range(
                    paragraphRange.End,
                    Math.Min(contentEnd, paragraphRange.End + 1));
                followingStartsTable = (bool)followingProbe.get_Information(
                    WdInformation.wdWithInTable);
            }

            if (immediateIsEmpty && followingStartsTable)
            {
                // This is the one paragraph Word requires between independent
                // tables. Keep it as a 1pt internal separator and create a second,
                // ordinary paragraph for the user's Enter. Deleting the user line
                // later therefore cannot merge the two 1x3 tables into a 2x3 host.
                return SplitNativeOmmlBodySeparatorForTyping(
                    document,
                    paragraphRange);
            }

            if (immediateIsCompact)
            {
                // A prior number-end Enter can already have created the ordinary
                // line after the compact separator. Reuse it instead of accumulating
                // a new paragraph on every Enter press.
                if (paragraphRange.End < contentEnd)
                {
                    Release(followingProbe); followingProbe = null;
                    followingProbe = document.Range(
                        paragraphRange.End,
                        Math.Min(contentEnd, paragraphRange.End + 1));
                    if (!(bool)followingProbe.get_Information(
                            WdInformation.wdWithInTable))
                    {
                        followingParagraphs = followingProbe.Paragraphs;
                        if (followingParagraphs.Count == 1)
                        {
                            followingParagraph = followingParagraphs[1];
                            followingRange = followingParagraph.Range.Duplicate;
                            if (followingRange.Start == paragraphRange.End
                                && IsPlainNativeOmmlBodyParagraph(
                                    document,
                                    followingRange,
                                    allowCompactTailBookmark: true))
                                return NormalizeNativeOmmlTypingParagraph(
                                    document,
                                    followingRange);
                        }
                    }
                }

                if (paragraphRange.End >= contentEnd)
                    return NormalizeNativeOmmlTypingParagraph(
                        document,
                        paragraphRange);
                return SplitNativeOmmlBodySeparatorForTyping(
                    document,
                    paragraphRange);
            }

            if (immediateIsEmpty)
                return NormalizeNativeOmmlTypingParagraph(
                    document,
                    paragraphRange);

            // Preserve existing text/content after the formula. Insert a new normal
            // paragraph before it rather than moving the user's caret into that text.
            return CreateNormalBodyParagraphAt(document, paragraphStart);
        }
        finally
        {
            Release(followingRange);
            Release(followingParagraph);
            Release(followingParagraphs);
            Release(followingProbe);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(immediateProbe);
            Release(tableRange);
        }
    }

    private static Range? SplitNativeOmmlBodySeparatorForTyping(
        Document document,
        Range separatorParagraphRange)
    {
        Range? source = null;
        Range? typingProbe = null;
        Paragraphs? typingParagraphs = null;
        Paragraph? typingParagraph = null;
        Range? typingRange = null;
        Range? compactProbe = null;
        Paragraphs? compactParagraphs = null;
        Paragraph? compactParagraph = null;
        Range? compactRange = null;
        try
        {
            source = separatorParagraphRange.Duplicate;
            if (!string.Equals(source.Text, "\r", StringComparison.Ordinal)
                || !IsPlainNativeOmmlBodyParagraph(
                    document,
                    source,
                    allowCompactTailBookmark: true))
                return null;
            var start = source.Start;

            // Inserting *after* this paragraph targets the following table's first
            // cell. Insert before the complete body paragraph instead: Word keeps
            // the new paragraph in the main story, while the original 1pt separator
            // remains immediately before the next table. The user paragraph is thus
            // independently deletable without merging the two 1x3 hosts.
            source.InsertParagraphBefore();

            typingProbe = document.Range(start, start + 1);
            if ((bool)typingProbe.get_Information(WdInformation.wdWithInTable))
                return null;
            typingParagraphs = typingProbe.Paragraphs;
            if (typingParagraphs.Count != 1) return null;
            typingParagraph = typingParagraphs[1];
            typingRange = typingParagraph.Range.Duplicate;
            if (typingRange.Start != start
                || !IsPlainNativeOmmlBodyParagraph(
                    document,
                    typingRange,
                    allowCompactTailBookmark: true))
                return null;
            var normalizedTyping = NormalizeNativeOmmlTypingParagraph(
                document,
                typingRange);

            compactProbe = document.Range(start + 1, start + 2);
            if ((bool)compactProbe.get_Information(WdInformation.wdWithInTable))
            {
                Release(normalizedTyping);
                return null;
            }
            compactParagraphs = compactProbe.Paragraphs;
            if (compactParagraphs.Count != 1)
            {
                Release(normalizedTyping);
                return null;
            }
            compactParagraph = compactParagraphs[1];
            compactRange = compactParagraph.Range.Duplicate;
            if (compactRange.Start != start + 1
                || !string.Equals(compactRange.Text, "\r", StringComparison.Ordinal))
            {
                Release(normalizedTyping);
                return null;
            }
            CompactNativeOmmlTableSeparatorParagraph(compactRange);
            return normalizedTyping;
        }
        finally
        {
            Release(compactRange);
            Release(compactParagraph);
            Release(compactParagraphs);
            Release(compactProbe);
            Release(typingRange);
            Release(typingParagraph);
            Release(typingParagraphs);
            Release(typingProbe);
            Release(source);
        }
    }

    private static Range? CreateNormalBodyParagraphAt(
        Document document,
        int position)
    {
        Range? insertion = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            var content = document.Content;
            try
            {
                position = Math.Max(
                    content.Start,
                    Math.Min(position, content.End - 1));
            }
            finally { Release(content); }
            insertion = document.Range(position, position);
            insertion.InsertParagraphBefore();
            var contentEnd = document.Content.End;
            probe = document.Range(position, Math.Min(contentEnd, position + 1));
            if ((bool)probe.get_Information(WdInformation.wdWithInTable))
                return null;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (paragraphRange.Start != position
                || !IsPlainNativeOmmlBodyParagraph(
                    document,
                    paragraphRange,
                    allowCompactTailBookmark: true))
                return null;
            return NormalizeNativeOmmlTypingParagraph(document, paragraphRange);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(insertion);
        }
    }

    private static bool IsPlainNativeOmmlBodyParagraph(
        Document document,
        Range paragraphRange,
        bool allowCompactTailBookmark)
    {
        Tables? tables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Frames? frames = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable)
                || !IsNumberingParagraphAdornment(paragraphRange.Text))
                return false;
            tables = paragraphRange.Tables;
            shapes = paragraphRange.InlineShapes;
            maths = paragraphRange.OMaths;
            fields = paragraphRange.Fields;
            frames = paragraphRange.Frames;
            if (tables.Count != 0
                || shapes.Count != 0
                || maths.Count != 0
                || fields.Count != 0
                || frames.Count != 0)
                return false;
            bookmarks = paragraphRange.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark); bookmark = bookmarks[index];
                if (!allowCompactTailBookmark
                    || !string.Equals(
                        bookmark.Name,
                        CompactTypingTailBookmarkName,
                        StringComparison.Ordinal))
                    return false;
            }
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(frames);
            Release(fields);
            Release(maths);
            Release(shapes);
            Release(tables);
        }
    }

    private static bool IsCompactNativeOmmlBodyParagraph(Range paragraphRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            font = paragraphRange.Font;
            format = paragraphRange.ParagraphFormat;
            return font.Size > 0f
                && font.Size <= CompactTypingTailFontSizePoints + 0.25f
                && format.LineSpacingRule == WdLineSpacing.wdLineSpaceExactly
                && format.LineSpacing > 0f
                && format.LineSpacing
                    <= CompactTypingTailLineSpacingPoints + 0.25f;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(format);
            Release(font);
        }
    }

    private static Range NormalizeNativeOmmlTypingParagraph(
        Document document,
        Range paragraphRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(CompactTypingTailBookmarkName))
            {
                bookmark = bookmarks[CompactTypingTailBookmarkName];
                var bookmarkRange = bookmark.Range;
                try
                {
                    if (bookmarkRange.Start >= paragraphRange.Start
                        && bookmarkRange.End <= paragraphRange.End)
                        bookmark.Delete();
                }
                finally { Release(bookmarkRange); }
            }
            try
            {
                object normalStyle = WdBuiltinStyle.wdStyleNormal;
                paragraphRange.set_Style(ref normalStyle);
            }
            catch { }
            font = paragraphRange.Font;
            try { font.Reset(); } catch { }
            font.Hidden = 0;
            font.Position = 0;
            font.Color = WdColor.wdColorAutomatic;
            format = paragraphRange.ParagraphFormat;
            try { format.Reset(); } catch { }
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            result = paragraphRange.Duplicate;
            result.Collapse(WdCollapseDirection.wdCollapseStart);
            var returned = result;
            result = null;
            return returned;
        }
        finally
        {
            Release(result);
            Release(format);
            Release(font);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Range? EnsureNormalTypingParagraphAfterNativeOmmlTable(
        Document document,
        string formulaId,
        out bool directTableMatched)
    {
        directTableMatched = false;
        Table? table = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null || !IsManagedNativeOmmlDirectTable(table))
                return null;
            directTableMatched = true;
            return EnsureDeletableTypingParagraphAfterNativeOmmlTable(
                document,
                table);
        }
        catch
        {
            // Once an exact direct-SEQ 1x3 host has been recognized, its typing
            // boundary must never fall through to the retired caption/Frame path.
            // Returning null lets the caller stop safely without mutating legacy
            // artifacts that do not exist for this formula.
            return null;
        }
        finally
        {
            Release(table);
        }
    }
}

using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationNumbering
{
    private const float NativeOmmlTableSideWidthPoints = 60f;

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
        bool deferExternalShapeCreation = false)
    {
        _ = formulaHeightPoints;
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

            table = EnsureNativeOmmlNumberTableHost(
                document,
                activeRange,
                formulaId);
            var refreshedTableFormula = ResolveSingleNativeOmmlRange(activeRange);
            Release(activeRange);
            activeRange = refreshedTableFormula;
            EnsureNumberedOmmlIsDisplay(activeRange);
            ConfigureNativeOmmlNumberTableGeometry(document, table, activeRange);
            traceStage("native-1x3-table");

            var numberPlan = ResolveDirectTableNumberPlan(
                document,
                activeRange,
                formulaId,
                plannedOrdinal,
                plannedPrefix);
            EnsureDirectTableSequenceNumber(
                document,
                table,
                activeRange,
                formulaId,
                formulaFontSizePoints,
                numberPlan.Ordinal,
                numberPlan.Prefix);
            traceStage("direct-visible-seq");

            if (metadata is not null)
            {
                repairedBookmark = WrapNativeOmmlTableFormulaIdentity(
                    document,
                    table,
                    activeRange,
                    metadata);
                WordOmmlNativeSource.StampFingerprintFromResolvedRange(
                    metadata,
                    activeRange);
                WordOmmlFormulaStore.Save(document, metadata);
            }

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

    private static Table EnsureNativeOmmlNumberTableHost(
        Document document,
        Range formulaRange,
        string formulaId)
    {
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
        Tables? tables = null;
        InlineShapes? shapes = null;
        OMaths? maths = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        Range? documentContent = null;
        Range? previousTableProbe = null;
        try
        {
            tableRange = table.Range;
            if (tableRange.Start <= document.Content.Start) return;
            probe = document.Range(tableRange.Start - 1, tableRange.Start);
            if ((bool)probe.get_Information(WdInformation.wdWithInTable)) return;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (paragraphRange.End != tableRange.Start
                || !IsNumberingParagraphAdornment(paragraphRange.Text))
                return;
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
                || frames.Count != 0)
                return;

            // The provisional VTOMML_<FormulaId> anchor is created before the
            // display formula is moved into its permanent 1x3 host. Word keeps a
            // collapsed bookmark on the source paragraph mark even after the OMath
            // content has been copied and deleted. Treat only that exact generated
            // identity as removable; any user/foreign bookmark still protects the
            // paragraph from deletion.
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
            Release(tables);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(tableRange);
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

    private static void ConfigureNativeOmmlNumberTableGeometry(
        Document document,
        Table table,
        Range formulaRange)
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
        bool formatOnly)
    {
        _ = formatOnly;
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
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName)
                || !bookmarks.Exists(numberName)
                || !bookmarks.Exists(visibleName))
                return false;
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;
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

            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            captionRange = captionBookmark.Range;
            field = FindNativeEquationFieldInRange(captionRange, nativeSequenceName);
            if (field is null) return false;
            EnsureSequenceFieldCodeCanBeRewritten(field);
            code = field.Code;
            code.Text = $" SEQ {LegacyEquationSequenceName} \\r {Math.Max(1, ordinal)} \\* ARABIC ";
            field.Update();

            // Preserve the cell's leading TAB and opening parenthesis. Only the
            // literal heading prefix between '(' and the field-begin control is
            // replaced when the numbering format changes.
            Release(code);
            code = field.Code;
            var labelStart = cellRange.Start + 1;
            var prefixStart = labelStart + 1;
            var fieldBegin = Math.Max(prefixStart, code.Start - 1);
            prefixRange = document.Range(prefixStart, fieldBegin);
            prefixRange.Text = prefix;

            Release(result);
            result = field.Result;
            Release(cellRange);
            cellRange = numberCell.Range.Duplicate;
            var editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            var closePosition = Math.Min(editableEnd, result.End + 1);
            closing = document.Range(closePosition, Math.Min(editableEnd, closePosition + 1));
            if (!string.Equals(closing.Text, ")", StringComparison.Ordinal))
            {
                closing.SetRange(closePosition, closePosition);
                closing.Text = ")";
            }

            DeleteBookmarkOnly(document, captionName);
            DeleteBookmarkOnly(document, numberName);
            DeleteBookmarkOnly(document, visibleName);
            Release(code);
            code = field.Code;
            Release(result);
            result = field.Result;
            Release(cellRange);
            cellRange = numberCell.Range.Duplicate;
            editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            var fieldEnd = Math.Min(editableEnd, result.End + 1);
            var labelEnd = Math.Min(editableEnd, fieldEnd + 1);
            var newCaption = document.Range(prefixStart, fieldEnd);
            var newNumber = document.Range(prefixStart, result.End);
            var newVisible = document.Range(labelStart, labelEnd);
            try
            {
                bookmarks = document.Bookmarks;
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
            Release(bookmarks);
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

    private static Range? EnsureNormalTypingParagraphAfterNativeOmmlTable(
        Document document,
        string formulaId)
    {
        Table? table = null;
        Range? tableRange = null;
        Range? insertion = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        ParagraphFormat? format = null;
        try
        {
            table = FindNumberedEquationTable(document, formulaId);
            if (table is null) return null;
            if (table.Rows.Count != 1 || table.Columns.Count != 3) return null;

            Cell? centerCell = null;
            Range? centerRange = null;
            OMaths? maths = null;
            OMath? math = null;
            try
            {
                centerCell = table.Cell(1, 2);
                centerRange = centerCell.Range;
                maths = centerRange.OMaths;
                if (maths.Count != 1) return null;
                math = maths[1];
                if (math.Type != WdOMathType.wdOMathDisplay) return null;
            }
            finally
            {
                Release(math);
                Release(maths);
                Release(centerRange);
                Release(centerCell);
            }

            tableRange = table.Range;
            var paragraphStart = tableRange.End;
            if (!CanReuseEmptyNativeCaptionParagraph(document, paragraphStart))
            {
                insertion = document.Range(paragraphStart, paragraphStart);
                insertion.InsertParagraphBefore();
            }

            var contentEnd = document.Content.End;
            if (paragraphStart >= contentEnd) return null;
            probe = document.Range(
                paragraphStart,
                Math.Min(contentEnd, paragraphStart + 1));
            if ((bool)probe.get_Information(WdInformation.wdWithInTable)) return null;
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (!IsNumberingParagraphAdornment(paragraphRange.Text)) return null;

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
            paragraphRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var result = paragraphRange;
            paragraphRange = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(format);
            Release(font);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(insertion);
            Release(tableRange);
            Release(table);
        }
    }
}

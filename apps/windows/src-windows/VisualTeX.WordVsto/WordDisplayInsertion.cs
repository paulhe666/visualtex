using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordFormulaHostSetup;
using static VisualTeX.WordVsto.WordMathTypeDisplayLayout;

namespace VisualTeX.WordVsto;

/// <summary>
/// Where a display formula goes: the insertion range and its own paragraph, MathType display rows and their
/// rollback, the placeholders numbered OMML replaces, and the selection after the formula.
/// </summary>
internal static class WordDisplayInsertion
{
    internal static void CompactParagraphBeforeOleDisplayFormula(
        Document document,
        Range insertion)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Min(Math.Max(insertion.Start, contentStart), contentEnd);
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A caret at the start of an empty paragraph belongs to that empty
            // paragraph, while the display formula will still be visually tied
            // to the preceding prose. Resolve the preceding paragraph mark in
            // that case so the local spacing adjustment is applied to the text
            // the reader actually sees above the equation.
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && position > contentStart)
            {
                Release(paragraphRange);
                paragraphRange = null;
                Release(paragraph);
                paragraph = null;
                Release(paragraphs);
                paragraphs = null;
                Release(anchor);
                anchor = null;

                anchor = document.Range(position - 1, position - 1);
                paragraphs = anchor.Paragraphs;
                if (paragraphs.Count == 0) return;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }

            if (!ContainsVisibleBodyText(paragraphRange.Text)) return;
            format = paragraph.Format;
            if (Math.Abs(format.SpaceAfter - ParagraphBeforeOleDisplaySpaceAfterPoints) > 0.01f)
                format.SpaceAfter = ParagraphBeforeOleDisplaySpaceAfterPoints;
        }
        finally
        {
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    private static Range DuplicateContainingParagraphRange(Range sourceRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = sourceRange.Paragraphs;
            if (paragraphs.Count == 0)
                throw new InvalidDataException("Word 未能定位行间公式所在段落。");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var result = paragraphRange;
            paragraphRange = null;
            return result;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range? DuplicateFollowingParagraphRange(
        Document document,
        Range paragraphRange)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? result = null;
        try
        {
            if (paragraphRange.End >= document.Content.End) return null;
            var start = paragraphRange.End;
            var end = Math.Min(document.Content.End, start + 1);
            probe = document.Range(start, end);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            result = paragraph.Range.Duplicate;
            var value = result;
            result = null;
            return value;
        }
        finally
        {
            Release(result);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    internal static void RemoveGeneratedParagraphAfterMathTypeRedraw(
        Document document,
        Range preservedFormulaParagraphRange,
        string? preservedFollowingParagraphText)
    {
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        InlineShapes? nextShapes = null;
        Fields? nextFields = null;
        Range? afterNextProbe = null;
        Paragraphs? afterNextParagraphs = null;
        Paragraph? afterNextParagraph = null;
        Range? afterNextRange = null;
        var tracePerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF"),
            "1",
            StringComparison.Ordinal);
        var cleanupWatch = tracePerformance ? Stopwatch.StartNew() : null;
        long cleanupCheckpoint = 0;
        void TraceCleanup(string stage)
        {
            if (cleanupWatch is null) return;
            var elapsed = cleanupWatch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage(
                $"mathtype-redraw-cleanup-perf stage={stage} deltaMs={elapsed - cleanupCheckpoint} totalMs={elapsed}");
            cleanupCheckpoint = elapsed;
        }
        try
        {
            formulaParagraphs = preservedFormulaParagraphRange.Paragraphs;
            if (formulaParagraphs.Count != 1) return;
            formulaParagraph = formulaParagraphs[1];
            formulaRange = formulaParagraph.Range.Duplicate;
            TraceCleanup("formula-range");
            // This cleanup is reached only immediately after InsertMathTypeOle
            // successfully materialized and validated Equation.DSMT4 in this exact
            // preserved redraw paragraph. Re-enumerating InlineShapes and probing
            // OLEFormat.ProgID here repeats an expensive OLE identity/layout query
            // for every formula in a large redraw batch. The structural checks below
            // still require the generated paragraph to be empty and the following
            // paragraph to match the text captured before insertion exactly.
            if (formulaRange.End >= document.Content.End) return;

            nextProbe = document.Range(
                formulaRange.End,
                Math.Min(document.Content.End, formulaRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count != 1) return;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range.Duplicate;
            TraceCleanup("next-range");
            if (ContainsVisibleBodyText(nextRange.Text)) return;
            TraceCleanup("next-text");
            nextShapes = nextRange.InlineShapes;
            nextFields = WordFormulaHost.GetLocalFields(nextRange);
            TraceCleanup("next-collections");
            if (nextShapes.Count > 0 || nextFields.Count > 0) return;
            if (preservedFollowingParagraphText is null) return;

            // If the source already had a blank paragraph immediately after the
            // display formula, this empty paragraph is user-authored and must stay.
            if (string.Equals(
                    nextRange.Text ?? string.Empty,
                    preservedFollowingParagraphText,
                    StringComparison.Ordinal))
                return;
            if (nextRange.End >= document.Content.End) return;

            afterNextProbe = document.Range(
                nextRange.End,
                Math.Min(document.Content.End, nextRange.End + 1));
            afterNextParagraphs = afterNextProbe.Paragraphs;
            if (afterNextParagraphs.Count != 1) return;
            afterNextParagraph = afterNextParagraphs[1];
            afterNextRange = afterNextParagraph.Range.Duplicate;
            TraceCleanup("after-next-range");
            if (!string.Equals(
                    afterNextRange.Text ?? string.Empty,
                    preservedFollowingParagraphText,
                    StringComparison.Ordinal))
                return;
            TraceCleanup("after-next-text");

            // Word's MathType Flat OPC generated exactly one extra empty paragraph
            // between the formula and the paragraph that originally followed the
            // LaTeX source. Delete only that proven generated paragraph.
            nextRange.Delete();
            TraceCleanup("delete-generated-paragraph");
        }
        finally
        {
            Release(afterNextRange);
            Release(afterNextParagraph);
            Release(afterNextParagraphs);
            Release(afterNextProbe);
            Release(nextFields);
            Release(nextShapes);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(formulaRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static bool IsIncompleteMathTypeNumberRow(Range paragraphRange)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var placeRefCount = 0;
        var onlyMathTypeNumberFields = true;
        try
        {
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            }

            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var fieldCode = (code.Text ?? string.Empty).Trim();
                if (fieldCode.IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    placeRefCount++;
                    continue;
                }
                if (fieldCode.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTSec ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTChap ", StringComparison.OrdinalIgnoreCase))
                    continue;
                onlyMathTypeNumberFields = false;
            }
            if (placeRefCount != 1) return false;
            // A failed constructor can leave the nested MathType SEQ fields as
            // document-level siblings. Their numeric results make Paragraph.Text
            // look non-empty even though the row owns no user content at all. If
            // every field in an OLE-free row belongs to this native numbering
            // scaffold, the row is incomplete by construction and can be removed
            // atomically on rollback/retry.
            if (onlyMathTypeNumberFields && fields.Count > 0) return true;

            var text = paragraphRange.Text ?? string.Empty;
            foreach (var character in text)
            {
                if (character == '\r' || character == '\t'
                    || character == '\u0013' || character == '\u0014' || character == '\u0015'
                    || character == '\u0001' || character == '\uFFFC'
                    || char.IsWhiteSpace(character))
                    continue;
                return false;
            }
            return true;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(shape);
            Release(shapes);
        }
    }

    private static (int Start, int End) GetIncompleteMathTypeTemplateExclusion(Range insertion)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        try
        {
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count != 1) return (-1, -1);
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            return insertion.Start >= owner.Start && insertion.Start <= owner.End
                && IsIncompleteMathTypeNumberRow(owner) ? (owner.Start, owner.End) : (-1, -1);
        }
        finally { Release(owner); Release(paragraph); Release(paragraphs); }
    }

    internal static bool ClearIncompleteMathTypeNumberRowAtInsertion(
        Document document,
        Range insertion)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        try
        {
            paragraphs = insertion.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (insertion.Start < paragraphRange.Start || insertion.Start > paragraphRange.End)
                return false;
            if (!IsIncompleteMathTypeNumberRow(paragraphRange)) return false;

            var start = paragraphRange.Start;
            var bodyEnd = Math.Max(start, paragraphRange.End - 1);
            body = document.Range(start, bodyEnd);
            body.Delete();
            insertion.SetRange(start, start);
            return true;
        }
        finally
        {
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsRollbackOwnedMathTypeDisplayRow(Range paragraphRange)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            shapes = paragraphRange.InlineShapes;
            if (shapes.Count == 0) return false;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) return false;
            }

            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            if (fields.Count == 0) return false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var fieldCode = (code.Text ?? string.Empty).Trim();
                if (fieldCode.IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    || fieldCode.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTSec ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("SEQ MTChap ", StringComparison.OrdinalIgnoreCase)
                    || fieldCode.StartsWith("EMBED Equation.DSMT4", StringComparison.OrdinalIgnoreCase))
                    continue;
                return false;
            }

            // Detached native SEQ results can expose only digits and the supported
            // equation-number punctuation outside field control characters. The OLE
            // itself is U+0001/U+FFFC. Reject any ordinary prose before treating the
            // row as transaction-owned rollback state.
            foreach (var character in paragraphRange.Text ?? string.Empty)
            {
                if (character == '\r' || character == '\t'
                    || character == '\u0013' || character == '\u0014' || character == '\u0015'
                    || character == '\u0001' || character == '\uFFFC'
                    || char.IsWhiteSpace(character)
                    || char.IsDigit(character)
                    || character is '(' or ')' or '.' or '-')
                    continue;
                return false;
            }
            return true;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(shape);
            Release(shapes);
        }
    }

    internal static void RollbackStandaloneMathTypeDisplayInsertion(
        Document document,
        int insertionStart,
        int paragraphCountBeforePreparation)
    {
        Range? anchor = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? body = null;
        Range? paragraphMark = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            var contentStart = document.Content.Start;
            var contentEnd = Math.Max(contentStart, document.Content.End - 1);
            var position = Math.Max(contentStart, Math.Min(insertionStart, contentEnd));
            anchor = document.Range(position, position);
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            // A half-relocated left number can temporarily contain two complete
            // MTPlaceRef trees plus the OLE. A detached-field failure can contain
            // one outer field plus sibling SEQ fields. If the entire row is proven
            // to consist only of MathType transaction artifacts, delete its body in
            // one Word range operation. This is substantially more reliable than
            // deleting OLE/field COM objects one by one after Word has invalidated
            // some of their ranges.
            if (IsRollbackOwnedMathTypeDisplayRow(paragraphRange))
            {
                var start = paragraphRange.Start;
                var end = Math.Max(start, paragraphRange.End - 1);
                body = document.Range(start, end);
                body.Delete();
            }
            else
            {
                // Conservative compatibility fallback for failure states that do
                // not prove full transaction ownership of the row.
                shapes = paragraphRange.InlineShapes;
                for (var index = shapes.Count; index >= 1; index--)
                {
                    Release(shape);
                    shape = shapes[index];
                    if (MathTypeOleInterop.IsMathTypeOle(shape))
                        shape.Delete();
                }
                Release(shape);
                shape = null;
                Release(shapes);
                shapes = null;

                if (IsIncompleteMathTypeNumberRow(paragraphRange)
                    || !ContainsVisibleBodyText(paragraphRange.Text))
                {
                    var start = paragraphRange.Start;
                    var end = Math.Max(start, paragraphRange.End - 1);
                    body = document.Range(start, end);
                    body.Delete();
                }
            }

            // ResolveStandaloneMathTypeDisplayInsertionRange may have created one
            // dedicated blank paragraph before InsertXML. Restore the original
            // paragraph count when that paragraph was created solely for the failed
            // transaction; preserve a blank paragraph that already existed.
            if (paragraphCountBeforePreparation >= 0
                && ReadDocumentParagraphCount(document) > paragraphCountBeforePreparation)
            {
                Release(paragraphRange);
                paragraphRange = paragraph.Range;
                if (!ContainsVisibleBodyText(paragraphRange.Text))
                {
                    paragraphMark = document.Range(
                        Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                        paragraphRange.End);
                    if (string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                        paragraphMark.Delete();
                }
            }
        }
        catch
        {
            // Rollback is best-effort and must never mask the original Word/OLE
            // failure. The caller rethrows the original exception with stage data.
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphMark);
            Release(body);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(anchor);
        }
    }

    internal static bool CanUseExactBatchMathTypeDisplayInsertion(
        Document document,
        Range insertion)
    {
        Range? content = null;
        Range? current = null;
        Range? following = null;
        Tables? followingTables = null;
        try
        {
            content = document.Content;
            var position = insertion.Start;
            if (position < content.Start || position >= content.End)
                return false;
            current = document.Range(
                position,
                Math.Min(content.End, position + 1));
            if (!string.Equals(current.Text, "\r", StringComparison.Ordinal))
                return false;
            if (WordFormulaHostLookup.IsWhollyInsideTable(current))
                return false;

            // Immediately before a table, Word 2021 can adopt Flat OPC into the
            // first table cell even though the collapsed insertion itself reports
            // body affinity. Keep the established guarded resolver for that rare
            // boundary; ordinary empty body paragraphs can use the exact range.
            if (position + 1 < content.End)
            {
                following = document.Range(
                    position + 1,
                    Math.Min(content.End, position + 2));
                followingTables = following.Tables;
                if (followingTables.Count > 0)
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
            Release(followingTables);
            Release(following);
            Release(current);
            Release(content);
        }
    }

    internal static Range ResolveStandaloneMathTypeDisplayInsertionRange(
        Document document,
        Range anchor,
        bool replaceAtExactInsertion = false)
    {
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? paragraphShapes = null;
        Fields? paragraphFields = null;
        try
        {
            // A collapsed Range at a paragraph/table boundary is frequently
            // reported by Word as belonging to the PREVIOUS paragraph. That is
            // catastrophic during right-to-left format conversion: a following
            // block equation can then be appended to the previous MathType row.
            // Resolve ownership with a one-character forward probe so the anchor
            // is interpreted in the paragraph that actually owns the insertion
            // position in the remaining document.
            var contentStart = document.Content.Start;
            var probeContentEnd = document.Content.End;
            var position = Math.Max(
                contentStart,
                Math.Min(anchor.Start, Math.Max(contentStart, probeContentEnd - 1)));
            var probeEnd = Math.Min(probeContentEnd, position + 1);
            if (replaceAtExactInsertion && position < probeContentEnd)
            {
                Range? exactCharacter = null;
                Cell? exactCell = null;
                try
                {
                    // Immediately before a user table, Word 2021 can give a
                    // forward probe at the BODY paragraph mark affinity to the
                    // first table-cell paragraph. Calling InsertParagraphBefore on
                    // that cell then extends the user table backwards and absorbs
                    // the redrawn display equations. Prove that the exact character
                    // is a real body CR and reuse that physical paragraph mark.
                    exactCharacter = document.Range(position, probeEnd);
                    if (string.Equals(exactCharacter.Text, "\r", StringComparison.Ordinal))
                    {
                        exactCell = WordFormulaHost.TryGetOwningCell(exactCharacter);
                        WordDoubleClickHook.TraceMessage(
                            $"mathtype-display-exact-owner position={position} cell={(exactCell is null ? "body" : exactCell.RowIndex + "," + exactCell.ColumnIndex)}");
                        if (exactCell is null)
                        {
                            Range? followingProbe = null;
                            Tables? followingTables = null;
                            Table? followingTable = null;
                            Range? followingTableRange = null;
                            Range? placeholderInsertion = null;
                            Range? placeholderRange = null;
                            try
                            {
                                // Word 2021 can adopt a Flat OPC <w:p> into the
                                // following table whenever InsertXML targets the last
                                // collapsed BODY position before that table. Make the
                                // replacement target unambiguously body-owned: insert
                                // one private-use character into the existing empty
                                // paragraph and return that exact non-collapsed range.
                                // InsertXML replaces the character atomically, so no
                                // user text or table structure participates.
                                var followingStart = position + 1;
                                if (followingStart < probeContentEnd)
                                {
                                    followingProbe = document.Range(
                                        followingStart,
                                        Math.Min(probeContentEnd, followingStart + 1));
                                    followingTables = followingProbe.Tables;
                                    if (followingTables.Count > 0)
                                    {
                                        followingTable = followingTables[1];
                                        followingTableRange = followingTable.Range;
                                        if (followingTableRange.Start == followingStart)
                                        {
                                            placeholderInsertion = document.Range(position, position);
                                            placeholderInsertion.Text = BulkInlineFormulaPlaceholder;
                                            placeholderRange = document.Range(
                                                position,
                                                position + BulkInlineFormulaPlaceholder.Length);
                                            WordDoubleClickHook.TraceMessage(
                                                $"mathtype-display-table-boundary-placeholder position={position} tableStart={followingStart}");
                                            var result = placeholderRange;
                                            placeholderRange = null;
                                            return result;
                                        }
                                    }
                                }
                            }
                            finally
                            {
                                Release(placeholderRange);
                                Release(placeholderInsertion);
                                Release(followingTableRange);
                                Release(followingTable);
                                Release(followingTables);
                                Release(followingProbe);
                            }
                            return document.Range(position, position);
                        }
                    }
                }
                finally
                {
                    Release(exactCell);
                    Release(exactCharacter);
                }
            }
            probe = document.Range(position, probeEnd);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count == 0) return anchor.Duplicate;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphShapes = paragraphRange.InlineShapes;
            paragraphFields = WordFormulaHost.GetLocalFields(paragraphRange);
            var hasInlineShapes = paragraphShapes.Count > 0;
            var hasFields = paragraphFields.Count > 0;
            if (IsMathTypeSectionStateParagraph(paragraphRange)
                && !hasInlineShapes
                && !ContainsVisibleBodyText(paragraphRange.Text))
            {
                // A collapsed Range exactly on a paragraph boundary is often
                // reported by Word as belonging to the previous paragraph.  When
                // that previous paragraph is MathType's hidden section state, the
                // user caret is logically in the following paragraph, so never
                // snap back to the hidden paragraph start.
                var nextParagraphStart = Math.Max(
                    document.Content.Start,
                    Math.Min(
                        paragraphRange.End,
                        Math.Max(document.Content.Start, document.Content.End - 1)));
                return document.Range(nextParagraphStart, nextParagraphStart);
            }
            if (!ContainsVisibleBodyText(paragraphRange.Text)
                && !hasInlineShapes
                && !hasFields)
                return document.Range(paragraphRange.Start, paragraphRange.Start);

            if (replaceAtExactInsertion)
            {
                // Format conversion has already deleted the original display host;
                // anchor.Start is the exact replacement position. At a paragraph
                // boundary Word reports the following paragraph as the owner of the
                // collapsed range. Inserting after that paragraph reverses adjacent
                // equations during right-to-left conversion. Reserve a fresh row
                // immediately BEFORE the following content instead.
                var exactStart = Math.Max(
                    document.Content.Start,
                    Math.Min(anchor.Start, Math.Max(document.Content.Start, document.Content.End - 1)));
                paragraphRange.InsertParagraphBefore();
                return document.Range(exactStart, exactStart);
            }

            // Flat OPC always carries its own <w:p>. Inserting it immediately before
            // the final paragraph mark does not automatically create a clean display
            // paragraph; Word can put the OLE beside existing prose. Create the one
            // required blank paragraph first, then materialize the MathType object
            // into that paragraph.
            var newParagraphStart = paragraphRange.End;
            paragraphRange.InsertParagraphAfter();
            var contentEnd = document.Content.End;
            newParagraphStart = Math.Max(
                document.Content.Start,
                Math.Min(newParagraphStart, Math.Max(document.Content.Start, contentEnd - 1)));
            return document.Range(newParagraphStart, newParagraphStart);
        }
        finally
        {
            Release(paragraphFields);
            Release(paragraphShapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static bool IsMathTypeSectionStateParagraph(Range range)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = WordFormulaHost.GetLocalFields(range);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTEditEquationSection2",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static Range PrepareNumberedOmmlReplacementTabPlaceholderPreservingOle(
        Document document,
        InlineShape sourceShape)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? shapeRange = null;
        Range? before = null;
        Range? after = null;
        Range? scaffold = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            shapeRange = sourceShape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OLE replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (shapeRange.Start < paragraphRange.Start
                || shapeRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OLE replacement range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            if (string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                Console.WriteLine(
                    $"    [diagnostic] numbered OLE→OMML host: paragraph={paragraphRange.Start}:{paragraphRange.End}, shape={shapeRange.Start}:{shapeRange.End}, maths={maths.Count}, shapes={shapes.Count}, fields={fields.Count}.");
                for (var diagnosticIndex = 1; diagnosticIndex <= fields.Count; diagnosticIndex++)
                {
                    Field? diagnosticField = null;
                    Range? diagnosticCode = null;
                    Range? diagnosticResult = null;
                    try
                    {
                        diagnosticField = fields[diagnosticIndex];
                        diagnosticCode = diagnosticField.Code;
                        diagnosticResult = diagnosticField.Result;
                        Console.WriteLine(
                            $"    [diagnostic] numbered OLE→OMML field#{diagnosticIndex}: type={diagnosticField.Type}, code={diagnosticCode.Start}:{diagnosticCode.End} '{diagnosticCode.Text}', result={diagnosticResult.Start}:{diagnosticResult.End} '{diagnosticResult.Text}'.");
                    }
                    finally
                    {
                        Release(diagnosticResult);
                        Release(diagnosticCode);
                        Release(diagnosticField);
                    }
                }
            }
            if (maths.Count != 0
                || shapes.Count != 1
                || !ContainsOnlySourceOleFields(fields, sourceShape))
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OLE paragraph containing another formula, object, or unrelated field.");

            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (shapeRange.End < editableEnd)
            {
                after = document.Range(shapeRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the numbered OLE formula.");
                after.Text = string.Empty;
            }

            // Delete structural TAB/zero-width runs before the OLE only after the
            // trailing side has been removed, because deleting the prefix shifts the
            // InlineShape's live Word coordinates.
            Release(shapeRange);
            shapeRange = sourceShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start < shapeRange.Start)
            {
                before = document.Range(paragraphRange.Start, shapeRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the numbered OLE formula.");
                before.Text = string.Empty;
            }

            Release(shapeRange);
            shapeRange = sourceShape.Range;
            Release(paragraphRange);
            paragraphRange = paragraph.Range;
            var start = paragraphRange.Start;
            scaffold = document.Range(start, start);
            scaffold.Text = "\t" + BulkInlineFormulaPlaceholder + "\t";
            placeholder = document.Range(
                start + 1,
                start + 1 + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(scaffold);
            Release(after);
            Release(before);
            Release(shapeRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool ContainsOnlySourceOleFields(
        Fields fields,
        InlineShape sourceShape)
    {
        if (fields.Count == 0) return true;
        Range? shapeRange = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            shapeRange = sourceShape.Range;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                result = field.Result;
                var codeText = code.Text ?? string.Empty;
                var isEmbed = field.Type == WdFieldType.wdFieldEmbed
                    || codeText.IndexOf(
                        "EMBED ",
                        StringComparison.OrdinalIgnoreCase) >= 0;
                // For embedded OLE Word reports InlineShape.Range as the complete
                // field span including the begin/separator/end characters, while
                // Field.Code/Result exclude those boundary characters. Reconstruct
                // that full span instead of expecting Result alone to contain the
                // InlineShape range.
                var fullFieldStart = Math.Max(0, code.Start - 1);
                var fullFieldEnd = result.End + 1;
                var ownsSourceShape = fullFieldStart <= shapeRange.Start
                    && fullFieldEnd >= shapeRange.End;
                if (!isEmbed || !ownsSourceShape)
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
            Release(result);
            Release(code);
            Release(field);
            Release(shapeRange);
        }
    }

    private static Range PrepareNumberedOmmlTrueDisplayReplacementPlaceholder(
        Document document,
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The true-display numbered OMML replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if ((bool)paragraphRange.get_Information(WdInformation.wdWithInTable)
                || equationRange.Start < paragraphRange.Start
                || equationRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The true-display numbered OMML replacement escaped its table-free formula paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to replace a true-display OMML paragraph containing another formula, object, or field.");
            var paragraphXml = paragraphRange.WordOpenXML ?? string.Empty;
            if (paragraphXml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) < 0
                || paragraphXml.IndexOf("<m:eqArr", StringComparison.OrdinalIgnoreCase) >= 0
                || paragraphXml.IndexOf("<w:fldChar", StringComparison.OrdinalIgnoreCase) >= 0)
                throw new InvalidOperationException(
                    "The numbered OMML source is not a pure genuine-display paragraph.");

            if (paragraphRange.Start < equationRange.Start)
            {
                before = document.Range(paragraphRange.Start, equationRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the true-display OMML formula.");
            }
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (equationRange.End < editableEnd)
            {
                after = document.Range(equationRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the true-display OMML formula.");
            }

            var start = paragraphRange.Start;
            editableRange = document.Range(start, editableEnd);
            editableRange.Text = BulkInlineFormulaPlaceholder;
            placeholder = document.Range(
                start,
                start + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range PrepareNumberedOmmlReplacementTabPlaceholder(
        Document document,
        Range equationRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        Range? editableRange = null;
        Range? placeholder = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The numbered OMML replacement no longer occupies one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (equationRange.Start < paragraphRange.Start
                || equationRange.End > paragraphRange.End)
                throw new InvalidOperationException(
                    "The numbered OMML replacement range escaped its paragraph.");

            maths = paragraphRange.OMaths;
            shapes = paragraphRange.InlineShapes;
            fields = paragraphRange.Fields;
            if (maths.Count != 1 || shapes.Count != 0 || fields.Count != 0)
                throw new InvalidOperationException(
                    "VisualTeX refused to rebuild a numbered OMML paragraph containing another formula, object, or field.");

            if (paragraphRange.Start < equationRange.Start)
            {
                before = document.Range(paragraphRange.Start, equationRange.Start);
                if (ContainsVisibleBodyText(before.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text before the numbered OMML formula.");
            }
            var editableEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (equationRange.End < editableEnd)
            {
                after = document.Range(equationRange.End, editableEnd);
                if (ContainsVisibleBodyText(after.Text))
                    throw new InvalidOperationException(
                        "VisualTeX refused to replace ordinary text after the numbered OMML formula.");
            }

            editableRange = document.Range(paragraphRange.Start, editableEnd);
            editableRange.Text = "\t" + BulkInlineFormulaPlaceholder + "\t";
            placeholder = document.Range(
                paragraphRange.Start + 1,
                paragraphRange.Start + 1 + BulkInlineFormulaPlaceholder.Length);
            var result = placeholder;
            placeholder = null;
            return result;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
            Release(placeholder);
            Release(editableRange);
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static Range ResolveDisplayInsertionRange(
        Document document,
        Range anchor,
        bool replaceAtExactInsertion = false)
    {
        Range? probe = null;
        Range? exactCharacter = null;
        Cell? exactCell = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? content = null;
        try
        {
            content = document.Content;
            var lastInsertPosition = Math.Max(content.Start, content.End - 1);
            var exactPosition = Math.Max(
                content.Start,
                Math.Min(anchor.Start, lastInsertPosition));
            if (replaceAtExactInsertion)
            {
                // An atomic source replacement can leave one real empty BODY
                // paragraph immediately before a user table. Word 2021 gives a
                // one-character forward probe at that boundary table affinity, so
                // probe.Paragraphs can become the first cell paragraph and move the
                // converted display formula into the table. Prove the exact CR is
                // not owned by a cell and reuse that body paragraph directly.
                if (exactPosition < content.End)
                {
                    exactCharacter = document.Range(
                        exactPosition,
                        Math.Min(content.End, exactPosition + 1));
                    var exactText = exactCharacter.Text ?? string.Empty;
                    WordDoubleClickHook.TraceMessage(
                        $"display-insertion-exact-probe position={exactPosition} chars=[{string.Join(",", exactText.Select(character => ((int)character).ToString("X4")))}] range={exactCharacter.Start}:{exactCharacter.End}");
                    if (string.Equals(
                            exactText,
                            "\r",
                            StringComparison.Ordinal))
                    {
                        exactCell = WordFormulaHost.TryGetOwningCell(exactCharacter);
                        WordDoubleClickHook.TraceMessage(
                            $"display-insertion-exact-owner position={exactPosition} cell={(exactCell is null ? "body" : exactCell.RowIndex + "," + exactCell.ColumnIndex)}");
                        if (exactCell is null)
                            return document.Range(exactPosition, exactPosition);
                    }
                    Release(exactCell);
                    exactCell = null;
                    Release(exactCharacter);
                    exactCharacter = null;
                }

                // At other paragraph boundaries a collapsed Word range can report
                // the paragraph on either side. Probe one character forward so an
                // in-place format conversion never appends the target formula to
                // the paragraph that originally followed the source equation.
                probe = document.Range(
                    exactPosition,
                    Math.Min(content.End, exactPosition + 1));
                paragraphs = probe.Paragraphs;
            }
            else
            {
                paragraphs = anchor.Paragraphs;
            }
            if (paragraphs.Count == 0)
                return anchor.Duplicate;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;

            var hasVisibleText = ContainsVisibleBodyText(paragraphRange.Text);
            if (replaceAtExactInsertion)
            {
                if (!hasVisibleText)
                    return document.Range(paragraphRange.Start, paragraphRange.Start);

                // The source display host has been removed completely and the
                // exact replacement point now coincides with the start of the
                // following user paragraph. Reserve one paragraph before that
                // content; this is the formula host, not an additional blank line.
                if (exactPosition <= paragraphRange.Start)
                {
                    paragraphRange.InsertParagraphBefore();
                    return document.Range(exactPosition, exactPosition);
                }
                throw new InvalidOperationException(
                    "VisualTeX refused to insert a converted display formula inside user paragraph text.");
            }

            // Reuse an existing empty paragraph instead of creating another one.
            // At the physical document end, however, paragraphRange.End lies one
            // position beyond the last legal insertion point. Clamping it back to
            // Content.End - 1 places the caret at the preceding OMath boundary and
            // Word absorbs the next display formula into that same equation. Create
            // the following paragraph explicitly and insert at its start instead.
            if (hasVisibleText && paragraphRange.End > lastInsertPosition)
            {
                var nextParagraphStart = paragraphRange.End;
                paragraphRange.InsertParagraphAfter();
                return document.Range(nextParagraphStart, nextParagraphStart);
            }

            // For ordinary body text with a following paragraph, its existing
            // boundary remains a stable display insertion point.
            var position = hasVisibleText
                ? paragraphRange.End
                : paragraphRange.Start;
            position = Math.Max(
                content.Start,
                Math.Min(position, lastInsertPosition));
            return document.Range(position, position);
        }
        finally
        {
            Release(exactCell);
            Release(exactCharacter);
            Release(content);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
        }
    }

    private static void FormatExistingDisplayParagraph(
        Range paragraphRange,
        bool preserveNativeOmmlSpacing)
    {
        ParagraphFormat? format = null;
        try
        {
            format = paragraphRange.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                format.SpaceBefore = 0;
                format.SpaceAfter = 0;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
            try { paragraphRange.ListFormat.RemoveNumbers(); } catch { }
        }
        finally { Release(format); }
    }

    private static bool CanUseDedicatedDisplayParagraphFastPath(
        Document document,
        Range source)
    {
        Range? previous = null;
        Range? next = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        try
        {
            if (source.StoryType != WdStoryType.wdMainTextStory)
                return false;

            // A Word table cell serializes its final structural position as CR+BEL
            // in Range.Text even though that pair occupies one story coordinate.
            // Therefore probing source.End with a one-character Range does not
            // reliably return a lone '\r' or '\a'. If the LaTeX source owns the
            // complete editable cell body, it already has the exact display host we
            // need: deleting only source.Start..source.End leaves the cell's single
            // paragraph/cell terminator in place. Do not insert synthetic CRs around
            // it, otherwise one source cell becomes three paragraphs after redraw.
            cell = WordFormulaHost.TryGetOwningCell(source);
            if (cell is not null)
            {
                cellRange = cell.Range;
                var editableCellEnd = Math.Max(cellRange.Start, cellRange.End - 1);
                if (source.Start == cellRange.Start
                    && source.End == editableCellEnd)
                    return true;
            }

            var contentStart = document.Content.Start;
            var contentEnd = document.Content.End;
            if (source.End >= contentEnd)
                return false;
            next = document.Range(source.End, Math.Min(contentEnd, source.End + 1));
            if (WordFormulaHost.ParagraphTerminatorStoryLength(next.Text) == 0)
                return false;
            if (source.Start <= contentStart)
                return true;
            previous = document.Range(source.Start - 1, source.Start);
            return WordFormulaHost.ParagraphTerminatorStoryLength(previous.Text) == 1;
        }
        finally
        {
            Release(cellRange);
            Release(cell);
            Release(cells);
            Release(next);
            Release(previous);
        }
    }

    // A display source owns its formula text and the immediately adjacent line
    // separators. Split both sides before deleting the source so every format
    // writes into the same exact empty paragraph, and existing prose keeps its
    // character/paragraph formatting. CR, manual BR and cell ends are distinct.
    internal static Range IsolateDisplaySourceParagraph(Document document, Range source)
    {
        var start = source.Start;
        var end = source.End;
        var expected = source.Text;
        Range? boundary = null;
        Range? isolated = null;
        try
        {
            if (source.StoryType != WdStoryType.wdMainTextStory)
                throw new InvalidDataException("Display redraw requires a main-story source range.");
            boundary = document.Range(end, Math.Min(document.Content.End, end + 1));
            if (boundary.Text == "\v") PromoteDisplayLineBoundary(boundary);
            else if (WordFormulaHost.ParagraphTerminatorStoryLength(boundary.Text) == 0)
            {
                boundary.SetRange(end, end);
                boundary.Text = "\r";
            }
            Release(boundary); boundary = null;
            if (start > document.Content.Start)
            {
                boundary = document.Range(start - 1, start);
                if (boundary.Text == "\v") PromoteDisplayLineBoundary(boundary);
                else if (WordFormulaHost.ParagraphTerminatorStoryLength(boundary.Text) == 0)
                {
                    boundary.SetRange(start, start);
                    boundary.Text = "\r";
                    start++; end++;
                }
            }
            isolated = document.Range(start, end);
            if (!string.Equals(isolated.Text, expected, StringComparison.Ordinal))
                throw new InvalidDataException("The display source changed while isolating its paragraph.");

            // The explicit CR/cell-end checks above already prove that this
            // source starts and ends on paragraph boundaries. Re-enumerating
            // Range.Paragraphs here is redundant, and Word compatibility mode
            // makes that COM query progressively slower as more OLE objects are
            // inserted, turning a large redraw into an O(N^2) paragraph walk.
            var result = isolated; isolated = null;
            return result;
        }
        finally
        {
            Release(isolated); Release(boundary);
        }
    }

    private static void PromoteDisplayLineBoundary(Range boundary)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        Frames? frames = null;
        try
        {
            maths = boundary.OMaths; shapes = boundary.InlineShapes;
            fields = boundary.Fields; bookmarks = boundary.Bookmarks; frames = boundary.Frames;
            if (boundary.Text != "\v" || maths.Count != 0 || shapes.Count != 0
                || fields.Count != 0 || bookmarks.Count != 0 || frames.Count != 0)
                throw new InvalidDataException("A display line boundary contains another object or identity.");
            boundary.Text = "\r";
        }
        finally
        {
            Release(frames); Release(bookmarks); Release(fields); Release(shapes); Release(maths);
        }
    }

    private static void EnsureBlankDisplayParagraph(
        Selection selection,
        bool preserveNativeOmmlSpacing)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        try
        {
            paragraphs = selection.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                range = paragraph.Range;
                if (ContainsVisibleBodyText(range.Text))
                    selection.TypeParagraph();
            }
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (!preserveNativeOmmlSpacing)
            {
                selection.ParagraphFormat.SpaceBefore = 0;
                selection.ParagraphFormat.SpaceAfter = 0;
                selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            }
        }
        finally
        {
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void MoveSelectionAfterNumberedDisplayFormula(
        Document document,
        Selection selection,
        Range equationRange,
        string formulaId,
        Table? knownDirectTable = null)
    {
        Range? typingRange = null;
        Range? ownerRange = null;
        Range? captionRange = null;
        Range? content = null;
        try
        {
            typingRange =
                WordNumberedTypingParagraphs.EnsureNormalTypingParagraphAfterNumberedDisplay(
                    document,
                    formulaId,
                    knownDirectTable);
            if (typingRange is not null)
            {
                selection.SetRange(typingRange.Start, typingRange.Start);
                selection.Collapse(WdCollapseDirection.wdCollapseStart);
                return;
            }

            ownerRange = WordNumberingReconcile.FindNumberingOwnerRange(
                document,
                formulaId);
            captionRange = WordNumberingReconcile.FindNativeEquationCaptionRange(
                document,
                formulaId);
            content = document.Content;
            var target = equationRange.End;
            if (ownerRange is not null) target = Math.Max(target, ownerRange.End);
            if (captionRange is not null) target = Math.Max(target, captionRange.End);
            target = Math.Max(content.Start, Math.Min(target, content.End));
            selection.SetRange(target, target);
            selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        }
        finally
        {
            Release(content);
            Release(captionRange);
            Release(ownerRange);
            Release(typingRange);
        }
    }

    private static int ReadParagraphStart(Range anchor)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count == 0) return anchor.Start;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.Start;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsStructurallyEmptyParagraph(Range range)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            maths = range.OMaths;
            shapes = range.InlineShapes;
            fields = WordFormulaHost.GetLocalFields(range);
            var text = (range.Text ?? string.Empty)
                .Trim('\r', '\a', '\v', '\f', '\t', ' ');
            return text.Length == 0
                && maths.Count == 0
                && shapes.Count == 0
                && fields.Count == 0;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(maths);
        }
    }

    private static void RepairPreservedDisplayParagraphBoundary(
        Document document,
        Range formulaRange,
        int sourceParagraphStart,
        int sourceParagraphCount)
    {
        if (sourceParagraphStart < 0 || sourceParagraphCount < 0) return;
        var currentParagraphCount = ReadDocumentParagraphCount(document);
        if (currentParagraphCount == sourceParagraphCount) return;
        if (currentParagraphCount != sourceParagraphCount + 1)
            throw new InvalidOperationException(
                $"Word changed the paragraph count unexpectedly while converting a display formula to OMML: before={sourceParagraphCount}, after={currentParagraphCount}.");

        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Range? content = null;
        Range? probe = null;
        Paragraphs? candidateParagraphs = null;
        Paragraph? candidateParagraph = null;
        Range? candidateRange = null;
        try
        {
            formulaParagraphs = formulaRange.Paragraphs;
            if (formulaParagraphs.Count != 1)
                throw new InvalidOperationException(
                    "The converted OMML display formula spans multiple paragraphs.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            content = document.Content;

            if (formulaParagraphRange.Start > sourceParagraphStart)
            {
                var previousStart = Math.Max(content.Start, formulaParagraphRange.Start - 1);
                probe = document.Range(previousStart, formulaParagraphRange.Start);
                candidateParagraphs = probe.Paragraphs;
                if (candidateParagraphs.Count == 0)
                    throw new InvalidOperationException(
                        "Word inserted an OMML paragraph before the source boundary, but the residual source paragraph could not be resolved.");
                candidateParagraph = candidateParagraphs[1];
                candidateRange = candidateParagraph.Range;
                if (candidateRange.Start != sourceParagraphStart
                    || !IsStructurallyEmptyParagraph(candidateRange))
                    throw new InvalidOperationException(
                        "The paragraph before the converted OMML formula is not the empty source paragraph VisualTeX expected to repair.");
            }
            else if (formulaParagraphRange.Start == sourceParagraphStart)
            {
                if (formulaParagraphRange.End >= content.End)
                    throw new InvalidOperationException(
                        "Word added a display-formula paragraph split at the document end, but no residual paragraph is available to repair.");
                probe = document.Range(
                    formulaParagraphRange.End,
                    Math.Min(content.End, formulaParagraphRange.End + 1));
                candidateParagraphs = probe.Paragraphs;
                if (candidateParagraphs.Count == 0)
                    throw new InvalidOperationException(
                        "Word inserted an OMML paragraph after the source boundary, but the residual paragraph could not be resolved.");
                candidateParagraph = candidateParagraphs[1];
                candidateRange = candidateParagraph.Range;
                if (!IsStructurallyEmptyParagraph(candidateRange))
                    throw new InvalidOperationException(
                        "The paragraph after the converted OMML formula contains user content and was not removed.");
            }
            else
            {
                throw new InvalidOperationException(
                    "The converted OMML display formula moved before its captured source paragraph boundary.");
            }

            candidateRange.Delete();
        }
        finally
        {
            Release(candidateRange);
            Release(candidateParagraph);
            Release(candidateParagraphs);
            Release(probe);
            Release(content);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }

        var repairedParagraphCount = ReadDocumentParagraphCount(document);
        if (repairedParagraphCount != sourceParagraphCount)
            throw new InvalidOperationException(
                $"VisualTeX could not restore the display formula's original paragraph structure: expected={sourceParagraphCount}, actual={repairedParagraphCount}.");
    }

    internal static void MoveSelectionAfterDisplayFormula(
        Selection selection,
        Range formulaRange)
    {
        selection.SetRange(formulaRange.End, formulaRange.End);
        selection.TypeParagraph();
        selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
        selection.ParagraphFormat.LeftIndent = 0;
        selection.ParagraphFormat.FirstLineIndent = 0;
        object normal = WdBuiltinStyle.wdStyleNormal;
        try { selection.Range.set_Style(ref normal); } catch { }
        try { selection.Range.ListFormat.RemoveNumbers(); } catch { }
        ResetSelectionTransientFormatting(selection);
    }

    internal static void ResetSelectionTransientFormatting(Selection selection)
    {
        Range? caret = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        Microsoft.Office.Interop.Word.Font? paragraphMarkFont = null;
        try
        {
            selectionFont = selection.Font;
            ResetTransientFont(selectionFont);

            // A collapsed Word Selection can report neutral formatting while the
            // paragraph mark still stores the old direct italic/bold state. The
            // next typed character inherits the paragraph mark, not the temporary
            // Selection.Font value. Clear that mark only when the caret is at the
            // paragraph boundary used for subsequent typing.
            if (selection.Start != selection.End) return;
            caret = selection.Range;
            paragraphRange = WordCharacterFormatting.ResolveParagraphAtPosition(caret, selection.Start);
            if (selection.Start < paragraphRange.End - 1) return;
            paragraphMark = paragraphRange.Duplicate;
            paragraphMark.SetRange(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                paragraphRange.End);
            paragraphMarkFont = paragraphMark.Font;
            ResetTransientFont(paragraphMarkFont);
        }
        finally
        {
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(caret);
            Release(selectionFont);
        }
    }

    // Every Font write is a formatting change Word applies at once; beside OLE
    // formulas each one costs several ms (bulk import: ~85 ms per reset). Only
    // properties that do not already read neutral are written, so the result
    // is the same formatting. Mixed values (wdUndefined) read non-neutral and
    // are written as before.
    internal static void ResetTransientFont(Microsoft.Office.Interop.Word.Font font)
    {
        if (font.Bold != 0) font.Bold = 0;
        if (font.Italic != 0) font.Italic = 0;
        if (font.StrikeThrough != 0) font.StrikeThrough = 0;
        try { if (font.DoubleStrikeThrough != 0) font.DoubleStrikeThrough = 0; } catch { }
        if (font.Underline != WdUnderline.wdUnderlineNone) font.Underline = WdUnderline.wdUnderlineNone;
        if (font.Hidden != 0) font.Hidden = 0;
        if (font.Subscript != 0) font.Subscript = 0;
        if (font.Superscript != 0) font.Superscript = 0;
        if (font.Position != 0) font.Position = 0;
        try { if (font.AllCaps != 0) font.AllCaps = 0; } catch { }
        try { if (font.SmallCaps != 0) font.SmallCaps = 0; } catch { }
    }
}

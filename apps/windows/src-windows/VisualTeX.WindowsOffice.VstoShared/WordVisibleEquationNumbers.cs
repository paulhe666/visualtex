using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordEquationNumberFormat;
using static VisualTeX.WordVsto.WordEquationNumberReferences;
using static VisualTeX.WordVsto.WordEquationParagraphLayout;
using static VisualTeX.WordVsto.WordNativeCaptions;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberedTableDisplay;
using static VisualTeX.WordVsto.WordNumberedTrueDisplay;
using static VisualTeX.WordVsto.WordNumberingCleanup;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;

namespace VisualTeX.WordVsto;

/// <summary>
/// The visible equation number: inserting, adopting, rebinding and removing it (with a formula's other numbering
/// artifacts), and its label.
/// </summary>
internal static class WordVisibleEquationNumbers
{
    internal static void RemoveVisibleEquationNumberForFormula(
        Document document,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
    }

    internal static void RemoveFormulaNumberingArtifacts(
        Document document,
        string formulaId,
        bool preserveNativeCaptionParagraph = false,
        Range? replacingCompleteHost = null,
        bool deferCaptionFlowTailCleanup = false,
        bool completeHostAlreadyVerified = false)
    {
        Bookmark? formulaBookmark = null;
        Range? formulaRange = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is not null)
            {
                formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
                if (replacingCompleteHost is not null)
                {
                    // A conversion replacing a complete managed row owns its
                    // formula and ordinary number together. The native row
                    // replacement removes both atomically; legacy bookmark-text
                    // deletion can touch a cell terminator or a retargeted range.
                    Tables? tables = null;
                    Table? table = null;
                    Range? tableRange = null;
                    try
                    {
                        if (!completeHostAlreadyVerified
                            && !HasReusableNumberedNativeOmmlDirectTableHost(
                                document,
                                formulaRange,
                                formulaId))
                            throw new InvalidDataException("Complete-host replacement requires a verified numbered OMML row.");
                        tables = formulaRange.Tables;
                        table = tables[1];
                        tableRange = table.Range;
                        if (table.Rows.Count != 1 || table.Columns.Count != 3
                            || replacingCompleteHost.StoryType != tableRange.StoryType
                            || replacingCompleteHost.Start != tableRange.Start || replacingCompleteHost.End != tableRange.End)
                            throw new InvalidDataException("The replacement scope is not exactly the formula's isolated managed row.");
                        return;
                    }
                    finally { Release(tableRange); Release(table); Release(tables); }
                }
                if (WordOmmlEquationNumbers.HasVisualTeXDirectSequenceEquationNumber(
                        WordOmmlConverter.ReadLocalEquationXml(formulaRange)))
                {
                    // VTEqCap_/VTEqNum_ live inside the professional OMath in the
                    // native #SEQ design. Never delete their bookmarked ranges as
                    // ordinary caption text; strip the complete number wrapper in
                    // one OMML replacement instead.
                    ConfigureUnnumberedDisplayFormula(
                        document,
                        formulaRange,
                        formulaId);
                    return;
                }
            }
        }
        finally
        {
            Release(formulaRange);
            Release(formulaBookmark);
        }

        if (replacingCompleteHost is not null)
            throw new InvalidDataException("The complete replacement host has no formula identity.");

        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(
            document,
            formulaId,
            preserveParagraphSeparator: preserveNativeCaptionParagraph,
            deferFlowTailCleanup: deferCaptionFlowTailCleanup);
    }

    internal static bool EnsureVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        bool adoptExistingTableReference = false,
        Table? knownNumberedTable = null,
        bool useConversionSafeVisibleNumber = false,
        bool deferFieldUpdate = false,
        bool knownFreshConversionTarget = false)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        if (!knownFreshConversionTarget
            && knownNumberedTable is null
            && !IsNumberedEquationTable(formulaRange)
            && TryRebindHealthyBodyTabVisibleNumber(
                document,
                formulaRange,
                formulaId,
                targetBookmarkName))
            return false;
        if (!knownFreshConversionTarget
            && HasVisibleEquationNumber(
                document,
                formulaRange,
                formulaId,
                targetBookmarkName,
                knownNumberedTable)) return false;
        if (!knownFreshConversionTarget
            && adoptExistingTableReference
            && TryAdoptExistingTableEquationNumber(
                document,
                formulaRange,
                formulaFontSizePoints,
                formulaId,
                targetBookmarkName))
            return true;
        if (!knownFreshConversionTarget)
            RemoveVisibleEquationNumber(document, formulaId);
        InsertVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaFontSizePoints,
            formulaId,
            targetBookmarkName,
            knownNumberedTable,
            useConversionSafeVisibleNumber,
            deferFieldUpdate,
            knownFreshConversionTarget);
        return true;
    }

    internal static bool TryRebindHealthyBodyTabVisibleNumber(
        Document document,
        Range formulaRange,
        string formulaId,
        string targetBookmarkName)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? labelRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? existing = null;
        Bookmark? rebound = null;
        try
        {
            if (ContainsNativeOmml(formulaRange)
                || RangeIsWhollyWithinTable(formulaRange))
                return false;

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var editableEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (formulaRange.End < paragraphRange.Start
                || formulaRange.End >= editableEnd)
                return false;

            labelRange = document.Range(
                formulaRange.End,
                editableEnd);
            var labelText = labelRange.Text ?? string.Empty;
            if (!labelText.StartsWith("\t(", StringComparison.Ordinal)
                || !labelText.EndsWith(")", StringComparison.Ordinal)
                || labelText.IndexOf('\r') >= 0
                || labelText.IndexOf('\a') >= 0)
                return false;

            fields = WordFormulaHost.GetLocalFields(labelRange);
            var targetReferenceCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldRef)
                    return false;
                code = field.Code;
                if (!IsReferenceToBookmark(
                        code.Text,
                        targetBookmarkName))
                    return false;
                targetReferenceCount++;
            }
            if (targetReferenceCount != 1)
                return false;

            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(visibleName))
            {
                existing = bookmarks[visibleName];
                var existingRange = existing.Range;
                try
                {
                    if (existingRange.StoryType == labelRange.StoryType
                        && existingRange.Start == labelRange.Start
                        && existingRange.End == labelRange.End
                        && string.Equals(
                            existingRange.Text,
                            labelText,
                            StringComparison.Ordinal))
                        return true;
                }
                finally { Release(existingRange); }
                existing.Delete();
                Release(existing);
                existing = null;
            }

            rebound = bookmarks.Add(
                visibleName,
                labelRange);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(rebound);
            Release(existing);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
            Release(labelRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void RebindFreshVisibleEquationNumberBookmark(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? visibleBookmark = null;
        Range? visibleRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? resultRange = null;
        Range? exactLabel = null;
        Range? editableLabel = null;
        Range? nextParagraphPrefix = null;
        Range? suffixInsertion = null;
        Bookmark? rebound = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName))
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} lost its visible bookmark.");
            visibleBookmark = bookmarks[visibleName];
            visibleRange = visibleBookmark.Range;
            var labelStart = visibleRange.Start;

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} has no single formula paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            var targetBookmark = NativeNumberBookmarkName(formulaId);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(resultRange);
                resultRange = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                if (field.Type != WdFieldType.wdFieldRef)
                    continue;
                code = field.Code;
                if (!IsReferenceToBookmark(code.Text, targetBookmark))
                    continue;
                resultRange = field.Result.Duplicate;
                break;
            }
            if (resultRange is null)
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} has no local REF result.");

            if (labelStart < paragraphRange.Start
                || labelStart >= paragraphRange.End)
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} has an invalid visible label start.");

            var editableEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            editableLabel = document.Range(labelStart, editableEnd);
            var editableText = editableLabel.Text ?? string.Empty;
            if (!editableText.StartsWith("\t(", StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} lost its TAB/'(' prefix.");

            if (!editableText.EndsWith(")", StringComparison.Ordinal))
            {
                // Deferred REF updates can either push the generated ')' across the
                // paragraph mark or consume it while replacing an initially-empty
                // field result. Repair only a proven fresh scaffold: TAB + '(' must
                // be followed by exactly this local REF result, with no user text
                // between the result and the paragraph terminator.
                var resultText = resultRange.Text ?? string.Empty;
                var expectedWithoutSuffix = "\t(" + resultText;
                if (!string.Equals(
                        editableText,
                        expectedWithoutSuffix,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        $"Fresh VisualTeX number {formulaId} lost its closing parenthesis "
                        + "but the remaining paragraph is not exactly TAB + (REF).");

                var movedSuffix = false;
                Range? content = null;
                try
                {
                    content = document.Content;
                    if (paragraphRange.End < content.End)
                    {
                        nextParagraphPrefix = document.Range(
                            paragraphRange.End,
                            Math.Min(content.End, paragraphRange.End + 1));
                        if (string.Equals(
                                nextParagraphPrefix.Text,
                                ")",
                                StringComparison.Ordinal))
                        {
                            nextParagraphPrefix.Delete();
                            movedSuffix = true;
                        }
                    }
                }
                finally { Release(content); }

                Release(paragraphRange);
                paragraphRange = paragraph.Range.Duplicate;
                editableEnd = Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
                Release(resultRange);
                resultRange = field.Result.Duplicate;

                // Whether ')' was moved or consumed, materialize it only after the
                // deferred REF result is now final. Use the paragraph's editable
                // end, not Result.End, because Word 2021 can report Result.End at
                // the structural paragraph position.
                var suffixPosition = editableEnd;
                if (suffixPosition <= labelStart + 1)
                    throw new InvalidDataException(
                        $"Fresh VisualTeX number {formulaId} has no legal suffix insertion point.");
                suffixInsertion = document.Range(
                    suffixPosition,
                    suffixPosition);
                suffixInsertion.Text = ")";

                if (movedSuffix)
                {
                    // The old generated suffix was removed from the next paragraph
                    // above; no other content was touched.
                }
            }

            Release(editableLabel);
            editableLabel = null;
            Release(paragraphRange);
            paragraphRange = paragraph.Range.Duplicate;
            editableEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            exactLabel = document.Range(labelStart, editableEnd);
            var exactText = exactLabel.Text ?? string.Empty;
            if (!exactText.StartsWith("\t(", StringComparison.Ordinal)
                || !exactText.EndsWith(")", StringComparison.Ordinal)
                || exactText.IndexOf('\r') >= 0
                || exactText.IndexOf('\a') >= 0
                || exactLabel.Paragraphs.Count != 1)
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} did not normalize to one TAB/(REF) label.");

            var exactStart = exactLabel.Start;
            var exactEnd = exactLabel.End;
            // Field.Result.End is not a visible-text boundary in Word 2021. For a
            // REF at paragraph end it may resolve at the paragraph-structure
            // position even though Result.Text is wholly inside TAB + (REF).
            // Prove the target REF locally, but keep VTEq_* bound only to the
            // visible label. The public range reader canonicalizes any bookmark
            // gravity that Word applies after Bookmarks.Add.
            if (resultRange.Start < exactStart
                || resultRange.Start >= exactEnd
                || resultRange.End > paragraphRange.End)
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} has a REF owner outside its visible label paragraph.");
            visibleBookmark.Delete();
            Release(visibleBookmark);
            visibleBookmark = null;

            // A Word Range is live. Deleting the old bookmark can expand a Range
            // that shared its boundary back across the paragraph mark. Freeze the
            // numeric boundary first, then reacquire a fresh Range only after the
            // old VTEq_* owner is gone.
            Release(exactLabel);
            exactLabel = document.Range(exactStart, exactEnd);
            var reboundSourceText = exactLabel.Text ?? string.Empty;
            if (!string.Equals(
                    reboundSourceText,
                    exactText,
                    StringComparison.Ordinal)
                || reboundSourceText.IndexOf('\r') >= 0
                || reboundSourceText.IndexOf('\a') >= 0
                || exactLabel.Paragraphs.Count != 1)
            {
                static string DescribeRebindCodes(string value) => string.Join(
                    ",",
                    value.Select(character => $"U+{(int)character:X4}"));
                throw new InvalidDataException(
                    $"Fresh VisualTeX number {formulaId} changed while rebinding its visible label. "
                    + $"range={exactStart}:{exactEnd}, "
                    + $"fieldResult={resultRange.Start}:{resultRange.End}, "
                    + $"source=[{reboundSourceText.Replace("\r", "<CR>").Replace("\a", "<CELL>").Replace("\t", "<TAB>")}], "
                    + $"expected=[{exactText.Replace("\r", "<CR>").Replace("\a", "<CELL>").Replace("\t", "<TAB>")}], "
                    + $"sourceCodes=[{DescribeRebindCodes(reboundSourceText)}], "
                    + $"expectedCodes=[{DescribeRebindCodes(exactText)}].");
            }
            rebound = bookmarks.Add(visibleName, exactLabel);
            Range? reboundRange = null;
            Range? canonicalRebound = null;
            try
            {
                reboundRange = rebound.Range;
                canonicalRebound = FindVisibleEquationNumberRange(
                    document,
                    formulaId)
                    ?? throw new InvalidDataException(
                        $"Fresh VisualTeX number {formulaId} lost its rebound visible label.");
                var canonicalText = canonicalRebound.Text ?? string.Empty;
                if (canonicalRebound.Start != exactLabel.Start
                    || !string.Equals(
                        canonicalText,
                        exactText,
                        StringComparison.Ordinal)
                    || canonicalRebound.Paragraphs.Count != 1)
                {
                    var rawText = reboundRange.Text ?? string.Empty;
                    static string DescribeCodes(string value) => string.Join(
                        ",",
                        value.Select(character => $"U+{(int)character:X4}"));
                    throw new InvalidDataException(
                        $"Fresh VisualTeX number {formulaId} could not rebind its exact visible label. "
                        + $"raw={reboundRange.Start}:{reboundRange.End}, "
                        + $"canonical={canonicalRebound.Start}:{canonicalRebound.End}, "
                        + $"expected={exactLabel.Start}:{exactLabel.End}, "
                        + $"fieldCode={code?.Start}:{code?.End}, "
                        + $"fieldResult={resultRange?.Start}:{resultRange?.End}, "
                        + $"canonicalParagraphs={canonicalRebound.Paragraphs.Count}, "
                        + $"rawText=[{rawText.Replace("\r", "<CR>").Replace("\a", "<CELL>").Replace("\t", "<TAB>")}], "
                        + $"canonicalText=[{canonicalText.Replace("\r", "<CR>").Replace("\a", "<CELL>").Replace("\t", "<TAB>")}], "
                        + $"expectedText=[{exactText.Replace("\r", "<CR>").Replace("\a", "<CELL>").Replace("\t", "<TAB>")}], "
                        + $"rawCodes=[{DescribeCodes(rawText)}], "
                        + $"canonicalCodes=[{DescribeCodes(canonicalText)}], "
                        + $"expectedCodes=[{DescribeCodes(exactText)}].");
                }
            }
            finally
            {
                Release(canonicalRebound);
                Release(reboundRange);
            }
        }
        finally
        {
            Release(rebound);
            Release(suffixInsertion);
            Release(nextParagraphPrefix);
            Release(editableLabel);
            Release(exactLabel);
            Release(resultRange);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(visibleRange);
            Release(visibleBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryAdoptExistingTableEquationNumber(
        Document document,
        Range formulaRange,
        float formulaFontSizePoints,
        string formulaId,
        string targetBookmarkName)
    {
        if (!IsNumberedEquationTable(formulaRange)) return false;
        Table? table = null;
        Cell? numberCell = null;
        Range? cellRange = null;
        Range? editableRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? labelRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? existingBookmark = null;
        ParagraphFormat? paragraph = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            table = formulaRange.Tables[1];
            numberCell = table.Cell(1, 3);
            cellRange = numberCell.Range;
            editableRange = cellRange.Duplicate;
            editableRange.End = Math.Max(editableRange.Start, editableRange.End - 1);
            var visibleText = editableRange.Text ?? string.Empty;
            if (!visibleText.StartsWith("(", StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;

            fields = editableRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                Release(code);
                code = field.Code;
                if (IsReferenceFieldCode(code.Text)) break;
                Release(field);
                field = null;
            }
            if (field is null) return false;

            if (!IsReferenceToBookmark(code?.Text, targetBookmarkName))
                code!.Text = $" REF {targetBookmarkName} \\h ";
            NormalizeReferenceResult(field, formulaFontSizePoints);
            result = field.Result;
            var labelEnd = Math.Min(editableRange.End, result.End + 1);
            if (labelEnd <= editableRange.Start) return false;
            labelRange = document.Range(editableRange.Start, labelEnd);
            var labelText = labelRange.Text ?? string.Empty;
            if (!labelText.StartsWith("(", StringComparison.Ordinal)
                || !labelText.EndsWith(")", StringComparison.Ordinal))
                return false;

            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (bookmarks.Exists(visibleName))
            {
                existingBookmark = bookmarks[visibleName];
                existingBookmark.Delete();
                Release(existingBookmark);
                existingBookmark = null;
            }
            bookmarks.Add(visibleName, labelRange);
            paragraph = labelRange.ParagraphFormat;
            paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            ApplyParagraphEquationNumberFont(labelRange, formulaFontSizePoints, position: 0);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(font);
            Release(paragraph);
            Release(existingBookmark);
            Release(bookmarks);
            Release(labelRange);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(editableRange);
            Release(cellRange);
            Release(numberCell);
            Release(table);
        }
    }

    private static bool HasVisibleEquationNumber(
        Document document,
        Range formulaRange,
        string formulaId,
        string targetBookmarkName,
        Table? knownNumberedTable = null)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        OMaths? maths = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Paragraphs? numberParagraphs = null;
        Paragraph? numberParagraph = null;
        Range? numberParagraphRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return false;
            bookmark = bookmarks[name];
            range = bookmark.Range;

            // Older builds inserted the visible REF field at OMath.Range.End.
            // Word still considers that position part of m:oMath, so the tab,
            // parentheses and number became equation content. A valid number is
            // an ordinary Word run after the formula, never a child of OMML.
            if (range.Start < formulaRange.End) return false;
            maths = range.OMaths;
            if (maths.Count > 0) return false;
            var tableLayout = knownNumberedTable is not null
                || IsNumberedEquationTable(formulaRange);
            var visibleText = range.Text ?? string.Empty;
            var expectedPrefix = tableLayout ? "(" : "\t(";
            if (!visibleText.StartsWith(expectedPrefix, StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;

            if (tableLayout)
            {
                if (!(bool)range.get_Information(WdInformation.wdWithInTable)
                    || range.Tables.Count == 0)
                    return false;
                Range? expectedTableRange = null;
                Range? visibleTableRange = null;
                try
                {
                    expectedTableRange = (knownNumberedTable ?? formulaRange.Tables[1]).Range;
                    visibleTableRange = range.Tables[1].Range;
                    if (visibleTableRange.Start != expectedTableRange.Start)
                        return false;
                }
                finally
                {
                    Release(visibleTableRange);
                    Release(expectedTableRange);
                }
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                numberParagraphs = range.Paragraphs;
                numberParagraph = numberParagraphs[1];
                numberParagraphRange = numberParagraph.Range;
                if (formulaParagraphRange.Start != numberParagraphRange.Start) return false;
            }

            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (IsReferenceToBookmark(code.Text, targetBookmarkName)) return true;
                    if (tableLayout && IsVisualTeXSequenceFieldCode(code.Text)) return true;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return false;
        }
        finally
        {
            Release(numberParagraphRange);
            Release(numberParagraph);
            Release(numberParagraphs);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(maths);
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void InsertVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        float formulaFontSizePoints,
        string formulaId,
        string targetBookmarkName,
        Table? knownNumberedTable = null,
        bool useConversionSafeVisibleNumber = false,
        bool deferFieldUpdate = false,
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
                $"[perf] VisibleRef.{stage}: +{elapsed - performanceCheckpoint}ms ({elapsed}ms)");
            performanceCheckpoint = elapsed;
        }
        Range? scaffoldRange = null;
        Range? fieldRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? fieldResult = null;
        Range? bookmarkRange = null;
        Bookmarks? bookmarks = null;
        try
        {
            var tableLayout = knownNumberedTable is not null
                || IsNumberedEquationTable(formulaRange);
            var suffixStart = PrepareEquationNumberInsertionPosition(
                formulaRange,
                knownNumberedTable);
            var reusesExistingRightTab = !tableLayout
                && HasTabImmediatelyBeforePosition(document, suffixStart);
            var labelStart = reusesExistingRightTab
                ? suffixStart - 1
                : suffixStart;
            // Normal VisualTeX insertion/edit keeps the long-standing stable
            // "()" scaffold. Only MathType/OMML -> VisualTeX conversion uses the
            // isolated REF-first path because migrated MathType paragraphs can
            // make Word consume a pre-seeded parenthesis while the field is born.
            var conversionSafeBody =
                useConversionSafeVisibleNumber && !tableLayout;
            var scaffold = tableLayout
                ? useConversionSafeVisibleNumber
                    ? (Text: string.Empty, FieldOffset: 0)
                    : (Text: "()", FieldOffset: 1)
                : conversionSafeBody
                    ? reusesExistingRightTab
                        ? (Text: "(", FieldOffset: 1)
                        : (Text: "\t(", FieldOffset: 2)
                    : reusesExistingRightTab
                        ? (Text: "()", FieldOffset: 1)
                        : EquationNumberScaffold();
            object suffixStartObject = suffixStart;
            object suffixEndObject = suffixStart;
            scaffoldRange = document.Range(ref suffixStartObject, ref suffixEndObject);
            if (scaffold.Text.Length > 0)
                scaffoldRange.Text = scaffold.Text;
            TraceStage("scaffold");

            var fieldStart = suffixStart + scaffold.FieldOffset;
            object fieldStartObject = fieldStart;
            object fieldEndObject = fieldStart;
            fieldRange = document.Range(ref fieldStartObject, ref fieldEndObject);
            // Keep field creation local to the number cell. document.Fields.Add
            // forces Word to maintain the complete story field collection even
            // though this REF belongs to one collapsed insertion range.
            fields = fieldRange.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"REF {targetBookmarkName} \\h";
            object preserveFormatting = true;
            field = fields.Add(
                fieldRange,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            TraceStage("add-ref");
            // Normal single-formula insertion resolves the REF immediately. Batch
            // format conversion keeps ScreenUpdating disabled and defers the field
            // result refresh until all 35 numbered targets exist; updating each REF
            // while the document field collection is still growing is O(N^2).
            if (!deferFieldUpdate)
                NormalizeReferenceResult(
                    field,
                    formulaFontSizePoints);
            TraceStage(deferFieldUpdate ? "defer-update-ref" : "update-ref");
            fieldResult = field.Result;
            if (!tableLayout)
            {
                // Word 2021 can discard direct pPr tab-stop definitions while
                // inserting the REF field. Repair only TabStops here. Re-running
                // the full paragraph configurator at this point also rewrites the
                // field/result boundary and can pull the closing parenthesis or
                // paragraph mark into Field.Result.
                ConfigureEquationParagraphTabStops(
                    formulaRange,
                    numbered: true);
                TraceStage("final-tab-stops");
                Release(fieldResult);
                fieldResult = field.Result;
            }

            // Do not rely on scaffoldRange.End. Word 2013/2016 perpetual,
            // Microsoft 365 and compatibility mode expand a Range around an
            // inserted field differently. Resolve the actual closing parenthesis
            // from the document text so the bookmark always contains the tab,
            // both brackets and the complete REF field.
            bookmarkRange = tableLayout && useConversionSafeVisibleNumber
                ? CompleteTableEquationNumberLabelAfterField(
                    document,
                    formulaRange,
                    field,
                    knownNumberedTable)
                : conversionSafeBody
                    ? CompleteParagraphEquationNumberLabelAfterField(
                        document,
                        formulaRange,
                        field,
                        labelStart)
                    : ResolveEquationNumberLabelRangeFast(
                        document,
                        formulaRange,
                        labelStart,
                        scaffoldRange,
                        fieldResult,
                        tableLayout,
                        knownNumberedTable);
            bookmarks = document.Bookmarks;
            bookmarks.Add(EquationBookmarkName(formulaId), bookmarkRange);
            if (tableLayout)
            {
                ParagraphFormat? format = null;
                try
                {
                    format = bookmarkRange.ParagraphFormat;
                    if (format.Alignment != WdParagraphAlignment.wdAlignParagraphRight)
                        format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                }
                finally { Release(format); }
            }
            TraceStage("bookmark-format");
            AlignEquationNumberVertically(
                bookmarkRange,
                formulaHeightPoints: 0f,
                formulaFontSizePoints);
            TraceStage("vertical-align");
        }
        finally
        {
            Release(bookmarks);
            Release(bookmarkRange);
            Release(fieldResult);
            Release(field);
            Release(fields);
            Release(fieldRange);
            Release(scaffoldRange);
        }
    }

    private static Range CompleteTableEquationNumberLabelAfterField(
        Document document,
        Range formulaRange,
        Field field,
        Table? knownNumberedTable)
    {
        Table? discoveredTable = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? prefix = null;
        Range? suffix = null;
        Range? candidate = null;
        try
        {
            discoveredTable = knownNumberedTable is null
                && formulaRange.Tables.Count > 0
                    ? formulaRange.Tables[1]
                    : null;
            var table = knownNumberedTable ?? discoveredTable
                ?? throw new InvalidOperationException(
                    "VisualTeX could not resolve the numbered equation table after creating its REF field.");
            cell = table.Cell(1, 3);

            // Prefix: insert at the cell's first editable position, which is now
            // guaranteed to be before the already-materialized REF field.
            cellRange = cell.Range;
            prefix = document.Range(cellRange.Start, cellRange.Start);
            prefix.Text = "(";
            Release(cellRange);
            cellRange = null;

            // Prefix insertion shifts the REF field. Re-read its live Result and
            // insert ')' *after the field-end control*, not at Result.End itself.
            // Inserting at Result.End is exactly what made Word absorb ')' into
            // Field.Result; the next field update then erased the parenthesis.
            Release(cellRange);
            cellRange = null;
            Release(candidate);
            candidate = null;
            Range? liveResult = null;
            try
            {
                liveResult = field.Result;
                cellRange = cell.Range;
                var suffixPosition = Math.Min(
                    Math.Max(cellRange.Start + 1, liveResult.End + 1),
                    Math.Max(cellRange.Start + 1, cellRange.End - 1));
                suffix = document.Range(suffixPosition, suffixPosition);
                suffix.Text = ")";

                Release(liveResult);
                liveResult = field.Result;
                if ((liveResult.Text ?? string.Empty).EndsWith(")", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "Word absorbed the VisualTeX closing parenthesis into the REF field result.");
            }
            finally { Release(liveResult); }

            Release(cellRange);
            cellRange = cell.Range;
            var editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            candidate = document.Range(cellRange.Start, editableEnd);
            var text = candidate.Text ?? string.Empty;
            if (!text.StartsWith("(", StringComparison.Ordinal)
                || !text.EndsWith(")", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word did not materialize the complete VisualTeX equation-number label after the REF field was finalized.");
            var result = candidate;
            candidate = null;
            return result;
        }
        finally
        {
            Release(candidate);
            Release(suffix);
            Release(prefix);
            Release(cellRange);
            Release(cell);
            Release(discoveredTable);
        }
    }

    private static Range CompleteParagraphEquationNumberLabelAfterField(
        Document document,
        Range formulaRange,
        Field field,
        int labelStart)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        Range? candidate = null;
        Range? liveResult = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The conversion-safe equation number no longer belongs to one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            var editableEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (labelStart < paragraphRange.Start
                || labelStart >= editableEnd)
                throw new InvalidOperationException(
                    "The conversion-safe equation-number label start escaped its paragraph.");

            prefix = document.Range(
                labelStart,
                Math.Min(editableEnd, labelStart + 2));
            if (!string.Equals(prefix.Text, "\t(", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "The conversion-safe equation-number prefix is not TAB + '('.");

            liveResult = field.Result;
            if (liveResult.Start < paragraphRange.Start
                || liveResult.Start > editableEnd
                || liveResult.End > paragraphRange.End)
                throw new InvalidOperationException(
                    $"The conversion-safe REF result escaped its equation paragraph: "
                    + $"result={liveResult.Start}:{liveResult.End}, "
                    + $"paragraph={paragraphRange.Start}:{paragraphRange.End}.");
            var suffixPosition = Math.Min(
                Math.Max(labelStart + 2, liveResult.End + 1),
                editableEnd);
            if (suffixPosition <= labelStart + 1)
                throw new InvalidOperationException(
                    "The conversion-safe REF field has no legal closing-parenthesis position.");
            suffix = document.Range(suffixPosition, suffixPosition);
            suffix.Text = ")";

            Release(liveResult);
            liveResult = field.Result;
            if ((liveResult.Text ?? string.Empty)
                .EndsWith(")", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word absorbed the VisualTeX closing parenthesis into the REF field result.");

            Release(paragraphRange);
            paragraphRange = paragraph.Range.Duplicate;
            candidate = document.Range(
                labelStart,
                Math.Min(paragraphRange.End - 1, suffixPosition + 1));
            var text = candidate.Text ?? string.Empty;
            if (!text.StartsWith("\t(", StringComparison.Ordinal)
                || !text.EndsWith(")", StringComparison.Ordinal)
                || text.IndexOf('\r') >= 0
                || text.IndexOf('\a') >= 0)
                throw new InvalidOperationException(
                    "Word did not materialize one single-paragraph VisualTeX equation-number label.");

            var result = candidate;
            candidate = null;
            return result;
        }
        finally
        {
            Release(liveResult);
            Release(candidate);
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static Range ResolveEquationNumberLabelRangeFast(
        Document document,
        Range formulaRange,
        int labelStart,
        Range scaffoldRange,
        Range fieldResult,
        bool tableLayout,
        Table? knownNumberedTable)
    {
        Range? candidate = null;
        try
        {
            var contentEnd = document.Content.End;
            var expectedEnd = Math.Min(contentEnd, fieldResult.End + 1);
            if (expectedEnd > labelStart)
            {
                candidate = document.Range(labelStart, expectedEnd);
                var text = candidate.Text ?? string.Empty;
                var expectedPrefix = tableLayout ? "(" : "\t(";
                if (text.StartsWith(expectedPrefix, StringComparison.Ordinal)
                    && text.EndsWith(")", StringComparison.Ordinal)
                    && text.IndexOf('\r') < 0
                    && text.IndexOf('\a') < 0)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
        }
        finally { Release(candidate); }

        // Compatibility fallback for older Word builds that report a field
        // Result.End outside the scaffold in a different way.
        return ResolveEquationNumberLabelRange(
            document,
            formulaRange,
            labelStart,
            scaffoldRange,
            fieldResult,
            tableLayout,
            knownNumberedTable);
    }

    private static Range ResolveEquationNumberLabelRange(
        Document document,
        Range formulaRange,
        int labelStart,
        Range scaffoldRange,
        Range fieldResult,
        bool tableLayout,
        Table? knownNumberedTable)
    {
        Range? character = null;
        Range? candidate = null;
        try
        {
            var searchEnd = Math.Min(document.Content.End, labelStart + 512);
            for (var position = labelStart + 1; position < searchEnd; position++)
            {
                object characterStart = position;
                object characterEnd = position + 1;
                character = document.Range(ref characterStart, ref characterEnd);
                if (!string.Equals(character.Text, ")", StringComparison.Ordinal))
                {
                    Release(character);
                    character = null;
                    continue;
                }

                object candidateStart = labelStart;
                object candidateEnd = character.End;
                candidate = document.Range(ref candidateStart, ref candidateEnd);
                var text = candidate.Text ?? string.Empty;
                var expectedPrefix = tableLayout ? "(" : "\t(";
                if (text.StartsWith(expectedPrefix, StringComparison.Ordinal)
                    && text.EndsWith(")", StringComparison.Ordinal)
                    && text.IndexOf('\r') < 0
                    && text.IndexOf('\a') < 0)
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
                Release(candidate);
                candidate = null;
                Release(character);
                character = null;
            }

            object fallbackStart = labelStart;
            object fallbackEnd = scaffoldRange.End;
            candidate = document.Range(ref fallbackStart, ref fallbackEnd);
            var fallbackText = candidate.Text ?? string.Empty;
            var fallbackPrefix = tableLayout ? "(" : "\t(";
            if (fallbackText.StartsWith(fallbackPrefix, StringComparison.Ordinal)
                && fallbackText.EndsWith(")", StringComparison.Ordinal)
                && fallbackText.IndexOf('\r') < 0
                && fallbackText.IndexOf('\a') < 0)
            {
                var fallback = candidate;
                candidate = null;
                return fallback;
            }

            // Some Word builds consume the literal ')' that was placed after the
            // zero-width REF insertion point while materializing/updating the
            // field.  A numbered equation table reserves cell (1,3) exclusively
            // for VisualTeX, so repair only that missing suffix locally instead of
            // failing the whole MathType/OMML conversion or running a document-wide
            // rebuild.  The opening parenthesis and field must already be intact.
            Release(candidate);
            candidate = null;
            if (tableLayout)
            {
                var repaired = TryRepairEquationNumberClosingParenthesis(
                    document,
                    formulaRange,
                    labelStart,
                    fieldResult,
                    knownNumberedTable);
                if (repaired is not null) return repaired;
            }

            var diagnostic = string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                ? DescribeEquationNumberLabelFailure(
                    document,
                    formulaRange,
                    labelStart,
                    scaffoldRange,
                    fieldResult,
                    knownNumberedTable)
                : string.Empty;
            throw new InvalidOperationException(
                "Word did not preserve the complete VisualTeX equation-number label."
                + diagnostic);
        }
        finally
        {
            Release(candidate);
            Release(character);
        }
    }

    private static string DescribeEquationNumberLabelFailure(
        Document document,
        Range formulaRange,
        int labelStart,
        Range scaffoldRange,
        Range fieldResult,
        Table? knownNumberedTable)
    {
        Table? table = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? probe = null;
        try
        {
            table = knownNumberedTable is null
                && formulaRange.Tables.Count > 0
                    ? formulaRange.Tables[1]
                    : null;
            var owner = knownNumberedTable ?? table;
            if (owner is null)
            {
                return $" [labelStart={labelStart}; scaffold=[{scaffoldRange.Start},{scaffoldRange.End}]; fieldResult=[{fieldResult.Start},{fieldResult.End}]; no-number-table]";
            }
            cell = owner.Cell(1, 3);
            cellRange = cell.Range;
            var probeStart = Math.Max(cellRange.Start, Math.Min(labelStart, cellRange.End - 1));
            var probeEnd = Math.Min(cellRange.End, Math.Max(probeStart, cellRange.End));
            probe = document.Range(probeStart, probeEnd);
            var text = (probe.Text ?? string.Empty)
                .Replace("\r", "<P>")
                .Replace("\a", "<CELL>")
                .Replace("\t", "<TAB>")
                .Replace("\v", "<BR>");
            return $" [labelStart={labelStart}; scaffold=[{scaffoldRange.Start},{scaffoldRange.End}]; fieldResult=[{fieldResult.Start},{fieldResult.End}]; cell=[{cellRange.Start},{cellRange.End}]; text='{text}']";
        }
        catch (Exception error)
        {
            return $" [label diagnostic failed: {error.GetType().Name}: {error.Message}]";
        }
        finally
        {
            Release(probe);
            Release(cellRange);
            Release(cell);
            Release(table);
        }
    }

    private static Range? TryRepairEquationNumberClosingParenthesis(
        Document document,
        Range formulaRange,
        int labelStart,
        Range fieldResult,
        Table? knownNumberedTable)
    {
        Tables? tables = null;
        Table? table = null;
        Cell? cell = null;
        Range? cellRange = null;
        Range? prefix = null;
        Range? insertion = null;
        Range? candidate = null;
        try
        {
            if (knownNumberedTable is null)
            {
                tables = formulaRange.Tables;
                if (tables.Count == 0) return null;
                table = tables[1];
            }
            var owner = knownNumberedTable ?? table;
            if (owner is null) return null;
            cell = owner.Cell(1, 3);
            cellRange = cell.Range;
            if (labelStart < cellRange.Start || labelStart >= cellRange.End)
                return null;

            var prefixEnd = Math.Min(cellRange.End, labelStart + 1);
            prefix = document.Range(labelStart, prefixEnd);
            if (!string.Equals(prefix.Text, "(", StringComparison.Ordinal))
                return null;

            // field.Result excludes Word's field-end control character; +1 is the
            // first legal character position after the complete REF field. Clamp
            // it to this reserved number cell so a malformed field can never make
            // the repair touch adjacent document content.
            var editableEnd = Math.Max(cellRange.Start, cellRange.End - 1);
            var insertionPosition = Math.Max(labelStart + 1, fieldResult.End + 1);
            insertionPosition = Math.Min(insertionPosition, editableEnd);
            if (insertionPosition <= labelStart) return null;

            insertion = document.Range(insertionPosition, insertionPosition);
            insertion.Text = ")";
            candidate = document.Range(labelStart, insertionPosition + 1);
            var text = candidate.Text ?? string.Empty;
            if (!text.StartsWith("(", StringComparison.Ordinal)
                || !text.EndsWith(")", StringComparison.Ordinal))
                return null;
            var result = candidate;
            candidate = null;
            return result;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(candidate);
            Release(insertion);
            Release(prefix);
            Release(cellRange);
            Release(cell);
            Release(table);
            Release(tables);
        }
    }

    private static bool HasTabImmediatelyBeforePosition(
        Document document,
        int position)
    {
        Range? preceding = null;
        try
        {
            if (position <= document.Content.Start) return false;
            preceding = document.Range(position - 1, position);
            return string.Equals(preceding.Text, "\t", StringComparison.Ordinal);
        }
        catch
        {
            return false;
        }
        finally { Release(preceding); }
    }

    private static int PrepareEquationNumberInsertionPosition(
        Range formulaRange,
        Table? knownNumberedTable = null)
    {
        if (knownNumberedTable is not null || IsNumberedEquationTable(formulaRange))
        {
            Table? table = null;
            Cell? cell = null;
            Range? cellRange = null;
            Range? editableRange = null;
            try
            {
                table = knownNumberedTable is null
                    ? formulaRange.Tables[1]
                    : null;
                cell = (knownNumberedTable ?? table!).Cell(1, 3);
                cellRange = cell.Range;
                // This cell is reserved exclusively for the generated number.
                // Word can leave an empty paragraph behind when a REF field is
                // removed/reconciled. Centering that empty paragraph together
                // with the new number pushes the visible number downward.
                // Clear everything except the structural cell mark, which
                // normalizes the cell to exactly one paragraph, then insert at
                // the beginning of that paragraph.
                editableRange = cellRange.Duplicate;
                editableRange.End = Math.Max(
                    editableRange.Start,
                    editableRange.End - 1);
                editableRange.Text = string.Empty;
                return cellRange.Start;
            }
            finally
            {
                Release(editableRange);
                Release(cellRange);
                Release(cell);
                Release(table);
            }
        }
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return Math.Max(paragraphRange.Start, paragraphRange.End - 1);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void RemoveVisibleEquationNumber(Document document, string formulaId)
    {
        var removedTrueDisplayHost = RemoveNativeDisplayNumberShapeAndAnchor(
            document,
            formulaId);
        if (!removedTrueDisplayHost
            || HasVisibleEquationNumberBookmark(document, formulaId))
            RemoveVisibleEquationNumberLegacy(document, formulaId);
    }

    internal static void RemoveVisibleEquationNumberLegacy(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Range? trailing = null;
        OMaths? maths = null;
        OMath? containingMath = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            var start = range.Start;
            var text = range.Text ?? string.Empty;
            try
            {
                maths = range.OMaths;
                if (maths.Count > 0) containingMath = maths[1];
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            if (containingMath is not null)
            {
                Range? containingRange = null;
                try
                {
                    containingRange = containingMath.Range;
                    if (WordOmmlEquationNumbers.HasVisualTeXNativeEquationNumber(
                            WordOmmlConverter.ReadLocalEquationXml(containingRange))
                        || HasManagedNativeOmmlHashSequenceHost(
                            document,
                            formulaId))
                    {
                        // The visible result is part of one generated display OMath.
                        // Deleting its result range would leave a damaged field and
                        // can make Word linearize the complete equation. Remove only
                        // the lookup bookmark; replacement/unnumbering rebuilds the
                        // professional OMath atomically.
                        bookmark.Delete();
                        return;
                    }
                }
                finally { Release(containingRange); }
            }
            range.Delete();

            // Legacy OMML numbering bookmarks stopped immediately before the
            // closing parenthesis. Remove that orphan as part of migration.
            if (!text.EndsWith(")", StringComparison.Ordinal)
                && start < document.Content.End)
            {
                object trailingStart = start;
                object trailingEnd = Math.Min(document.Content.End, start + 1);
                trailing = document.Range(ref trailingStart, ref trailingEnd);
                if (string.Equals(trailing.Text, ")", StringComparison.Ordinal))
                    trailing.Delete();
            }
            try { containingMath?.BuildUp(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        }
        finally
        {
            Release(containingMath);
            Release(maths);
            Release(trailing);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static void RemoveNativeCaption(
        Document document,
        string formulaId,
        bool preserveParagraphSeparator = false,
        bool deferFlowTailCleanup = false)
    {
        if (Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_CAPTION_OWNERS") == "1")
        {
            var diagnosticBookmarks = document.Bookmarks;
            try
            {
                foreach (var diagnosticName in new[] { NativeCaptionBookmarkName(formulaId), NativeNumberBookmarkName(formulaId) })
                {
                    if (!diagnosticBookmarks.Exists(diagnosticName))
                    {
                        WordDoubleClickHook.TraceMessage($"caption-owner-probe name={diagnosticName} missing=True");
                        continue;
                    }
                    var diagnosticBookmark = diagnosticBookmarks[diagnosticName];
                    var diagnosticRange = diagnosticBookmark.Range;
                    try { WordDoubleClickHook.TraceMessage($"caption-owner-probe name={diagnosticName} range={diagnosticRange.Start}:{diagnosticRange.End} math={WordFormulaHost.IsWhollyInsideMath(diagnosticRange)} text=[{diagnosticRange.Text}]"); }
                    finally { Release(diagnosticRange); Release(diagnosticBookmark); }
                }
            }
            finally { Release(diagnosticBookmarks); }
        }
        if (TryRemoveNativeHashSequenceCaptionBookmarks(document, formulaId))
        {
            if (!deferFlowTailCleanup)
                CleanupOwnedNativeCaptionFlowTail(
                    document,
                    formulaId,
                    deleteParagraph: !preserveParagraphSeparator);
            return;
        }

        // Word bookmark insertion gravity can expand VTEqCap_<id> when the user
        // later inserts ordinary text/formulas exactly at the hidden caption's
        // former start boundary. VTEqNum_<id>, however, remains attached to the
        // actual SEQ result. Never trust a widened caption bookmark for deletion:
        // re-derive its ownership from the number bookmark's one verified SEQ
        // paragraph first. This keeps the overlap guard below strict while
        // preventing user content / unrelated OMath from being mistaken for
        // VisualTeX-owned caption content.
        RebindExternalCaptionToNativeNumberParagraph(
            document,
            formulaId);

        DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
        DeleteBookmarkedRangeAndContainingFrame(
            document,
            NativeCaptionBookmarkName(formulaId),
            preserveParagraphSeparator);
        // Conversion still needs a stable body paragraph while the replacement
        // object is inserted. Its transaction owns this marker until the new
        // target's physical host is verified; ordinary unnumbering is unchanged.
        if (!deferFlowTailCleanup)
            CleanupOwnedNativeCaptionFlowTail(
                document,
                formulaId,
                deleteParagraph: !preserveParagraphSeparator);
    }
}

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal WordBulkInsertResult InsertBulkDocumentHostCore(
        WordBulkImportDocument source,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared,
        string? expectedDocumentId,
        string? sourceObjectId)
    {
        using var numberFormatScope =
            WordEquationNumberFormatAuthority.BeginOperationScope();
        if (source is null)
            throw new ArgumentNullException(nameof(source));
        if (prepared is null)
            throw new ArgumentNullException(nameof(prepared));

        Document? document = null;
        Selection? selection = null;
        Range? captured = null;
        // Diagnostics only (same switch as the redraw perf log): where the
        // bulk write spends its time.
        var bulkMetrics = string.IsNullOrWhiteSpace(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_PERF_LOG"))
            ? null
            : new WordOperationMetrics();
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                expectedDocumentId);

            captured =
                WordFormulaOperationLocator.ResolveCapturedRange(
                    document,
                    sourceObjectId);

            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Bulk Import",
                () =>
                {
                    selection = _application.Selection;
                    selection.SetRange(
                        captured.Start,
                        captured.End);
                    if (selection.Start != selection.End)
                        selection.Text = string.Empty;
                    selection.Collapse(
                        WdCollapseDirection.wdCollapseStart);

                    var insertedIds =
                        new List<string>();
                    var deferredVisualTeXNumbers = 0;
                    // As in the redraw: display tab stops once per section.
                    var displayTabStopsBySection =
                        new Dictionary<int, (float Center, float Right)>();

                    for (var blockIndex = 0;
                         blockIndex < source.Blocks.Count;
                         blockIndex++)
                    {
                        var block =
                            source.Blocks[blockIndex];
                        ReportOperationProgress(OfficePluginLanguage.Text("正在写入导入内容…", "Writing the imported content…"), insertedIds.Count, prepared.Count);
                        var nextKind =
                            blockIndex + 1
                                < source.Blocks.Count
                                ? source.Blocks[
                                    blockIndex + 1].Kind
                                : (WordBulkBlockKind?)null;

                        if (block.Kind ==
                            WordBulkBlockKind.DisplayFormula)
                        {
                            var formulaRun =
                                block.Runs.Single(
                                    run => run.IsFormula);
                            if (!prepared.TryGetValue(
                                    formulaRun.Id,
                                    out var formula))
                                throw new InvalidDataException(OfficePluginLanguage.Text(
                                    $"行间公式“{WordFormulaRenderException.Excerpt(formulaRun.Latex, 120)}”缺少渲染结果。",
                                    $"The display formula “{WordFormulaRenderException.Excerpt(formulaRun.Latex, 120)}” is missing its render result."));

                            formula.Session.DisplayMode =
                                "block";
                            formula.Session.Numbered =
                                source.NumberDisplayFormulas;
                            using (WordOperationMetrics.Measure("Bulk.DisplayFormula"))
                                if (InsertPreparedHostCore(
                                        document,
                                        selection,
                                        formula,
                                        display: true,
                                        insertedIds,
                                        displayTabStopsBySection))
                                    deferredVisualTeXNumbers++;

                            using (WordOperationMetrics.Measure("Bulk.ResetNextParagraph"))
                                ResetNextParagraphFormatting(
                                    selection,
                                    block.Kind,
                                    nextKind);
                            continue;
                        }

                        EnsureWritableParagraph(
                            selection);
                        var paragraphStart =
                            selection.Start;
                        var pending =
                            new List<(
                                int Start,
                                PreparedWordBulkFormula Formula)>();

                        foreach (var run in block.Runs)
                        {
                            if (!run.IsFormula)
                            {
                                using (WordOperationMetrics.Measure("Bulk.TextRun"))
                                    InsertNativeTextRun(
                                        document,
                                        selection,
                                        run);
                                continue;
                            }

                            if (!prepared.TryGetValue(
                                    run.Id,
                                    out var formula))
                                throw new InvalidDataException(OfficePluginLanguage.Text(
                                    $"行内公式“{WordFormulaRenderException.Excerpt(run.Latex, 120)}”缺少渲染结果。",
                                    $"The inline formula “{WordFormulaRenderException.Excerpt(run.Latex, 120)}” is missing its render result."));

                            var placeholderStart =
                                selection.Start;
                            using (WordOperationMetrics.Measure("Bulk.Placeholder"))
                                selection.TypeText(
                                    BulkInlineFormulaPlaceholder);
                            pending.Add((
                                placeholderStart,
                                formula));
                        }

                        using (WordOperationMetrics.Measure("Bulk.TypeParagraph"))
                            selection.TypeParagraph();
                        var paragraphEnd =
                            selection.Start;
                        using (WordOperationMetrics.Measure("Bulk.ParagraphFormatting"))
                            ApplyBulkParagraphFormatting(
                                document,
                                paragraphStart,
                                paragraphEnd,
                                block);

                        for (var index =
                                 pending.Count - 1;
                             index >= 0;
                             index--)
                        {
                            var item =
                                pending[index];
                            Range? placeholder = null;
                            Range? insertion = null;
                            try
                            {
                                placeholder =
                                    document.Range(
                                        item.Start,
                                        item.Start
                                        + BulkInlineFormulaPlaceholder.Length);
                                if (!string.Equals(
                                        placeholder.Text,
                                        BulkInlineFormulaPlaceholder,
                                        StringComparison.Ordinal))
                                    throw new InvalidDataException(OfficePluginLanguage.Text(
                                        "批量导入的行内公式占位符发生了变化。",
                                        "The inline-equation placeholder changed during bulk import."));
                                placeholder.Text =
                                    string.Empty;
                                insertion =
                                    document.Range(
                                        item.Start,
                                        item.Start);

                                item.Formula.Session.DisplayMode =
                                    "inline";
                                item.Formula.Session.Numbered =
                                    false;
                                var request =
                                    BuildHostWriteRequest(
                                        item.Formula.Session,
                                        item.Formula.MathMl,
                                        item.Formula.PngPath,
                                        item.Formula.EmfPath);
                                // Like the LaTeX redraw: a batch writer inserting at a
                                // collapsed caret imports the prepared equation directly
                                // and skips the duplicate generic read-back validation
                                // (the writers prove the host they wrote).
                                request.BatchMutation = true;
                                WordFormulaHostDescriptor inserted;
                                using (WordOperationMetrics.Measure("Bulk.InlineFormula"))
                                    inserted =
                                        WordFormulaHostMutationKernel
                                            .InsertInActiveTransaction(
                                                _application,
                                                document,
                                                insertion,
                                                request);
                                insertedIds.Add(
                                    inserted.FormulaId
                                    ?? item.Formula.Session.FormulaId);
                            }
                            finally
                            {
                                Release(insertion);
                                Release(placeholder);
                            }
                        }

                        using (WordOperationMetrics.Measure("Bulk.MoveSelection"))
                            MoveSelectionAfterBulkParagraph(
                                document,
                                selection,
                                paragraphStart);
                        using (WordOperationMetrics.Measure("Bulk.ResetNextParagraph"))
                            ResetNextParagraphFormatting(
                                selection,
                                block.Kind,
                                nextKind);
                    }

                    // Numbered VisualTeX hosts were written with their field
                    // trees but without per-formula evaluation: one ordered field
                    // pass gives every number, as after a batch redraw/conversion.
                    // Inside the import transaction, so a failure rolls it back.
                    if (deferredVisualTeXNumbers > 0)
                        using (WordOperationMetrics.Measure("Bulk.RefreshNumbers"))
                            _ = WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                                document,
                                kindFilter: WordFormulaHostKind.VisualTeX,
                                minimumExpectedNumbered: deferredVisualTeXNumbers);

                    return new WordBulkInsertResult
                    {
                        BlockCount =
                            source.Blocks.Count,
                        FormulaCount =
                            insertedIds.Count,
                        FormulaIds =
                            insertedIds,
                    };
                });
        }
        finally
        {
            if (bulkMetrics is not null)
            {
                bulkMetrics.Dispose();
                WordDoubleClickHook.TraceMessage(
                    "bulk-insert-metrics " + string.Join(" | ", bulkMetrics.Entries
                        .OrderByDescending(item => item.Value.Milliseconds)
                        .Take(30)
                        .Select(item => $"{item.Key}={item.Value.Milliseconds:0}/{item.Value.Calls}")));
            }
            Release(captured);
            Release(selection);
            Release(document);
        }
    }

    // Returns true when the formula is a numbered VisualTeX host whose number
    // is left to the caller's single ordered refresh.
    private bool InsertPreparedHostCore(
        Document document,
        Selection selection,
        PreparedWordBulkFormula formula,
        bool display,
        ICollection<string> insertedIds,
        IDictionary<int, (float Center, float Right)> displayTabStopsBySection)
    {
        Range? insertion = null;
        Range? current = null;
        try
        {
            current = selection.Range.Duplicate;
            current.Collapse(
                WdCollapseDirection.wdCollapseStart);

            if (display)
            {
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                try
                {
                    paragraphs = current.Paragraphs;
                    if (paragraphs.Count != 1)
                        throw new InvalidDataException(OfficePluginLanguage.Text(
                            "批量导入的行间公式插入点不属于唯一段落。",
                            "The display-equation insertion point does not belong to exactly one paragraph during bulk import."));
                    paragraph = paragraphs[1];
                    paragraphRange =
                        paragraph.Range.Duplicate;
                    var editableEnd =
                        Math.Max(
                            paragraphRange.Start,
                            paragraphRange.End - 1);
                    Range? body = null;
                    try
                    {
                        body = document.Range(
                            paragraphRange.Start,
                            editableEnd);
                        if (ContainsVisibleBodyText(
                                body.Text))
                        {
                            var next =
                                paragraphRange.End;
                            paragraphRange
                                .InsertParagraphAfter();
                            Release(current);
                            current =
                                document.Range(
                                    next,
                                    next);
                        }
                        else
                        {
                            current.SetRange(
                                paragraphRange.Start,
                                paragraphRange.Start);
                        }
                    }
                    finally
                    {
                        Release(body);
                    }
                }
                finally
                {
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                }
            }

            insertion = current.Duplicate;
            var request =
                BuildHostWriteRequest(
                    formula.Session,
                    formula.MathMl,
                    formula.PngPath,
                    formula.EmfPath);
            // Batch write as the redraw does (see the inline case). A numbered
            // VisualTeX host is attached without evaluation; its number comes
            // from the caller's ordered refresh after the import. Numbered native
            // OMML keeps per-formula numbering finalization.
            var deferVisualTeXNumber = formula.Session.Numbered
                && request.Kind == WordFormulaHostKind.VisualTeX;
            request.BatchMutation = !formula.Session.Numbered || deferVisualTeXNumber;
            if (display && request.Kind == WordFormulaHostKind.VisualTeX)
            {
                using (WordOperationMetrics.Measure("Bulk.DisplayTabStops"))
                {
                    var tabStops = ResolveBatchDisplayTabStops(
                        document,
                        insertion,
                        displayTabStopsBySection);
                    if (tabStops.HasValue)
                    {
                        request.DisplayCenterTabPoints = tabStops.Value.Center;
                        request.DisplayRightTabPoints = tabStops.Value.Right;
                    }
                }
            }
            var inserted =
                WordFormulaHostMutationKernel
                    .InsertInActiveTransaction(
                        _application,
                        document,
                        insertion,
                        request);
            insertedIds.Add(
                inserted.FormulaId
                ?? formula.Session.FormulaId);

            Range? insertedRange = null;
            try
            {
                insertedRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        inserted.Range);
                RestoreCaretAfterCoreMutation(
                    document,
                    inserted,
                    insertedRange,
                    bodyNumberingPending: deferVisualTeXNumber && !inserted.WithinTable);
            }
            finally
            {
                Release(insertedRange);
            }
            return deferVisualTeXNumber;
        }
        finally
        {
            Release(insertion);
            Release(current);
        }
    }
}

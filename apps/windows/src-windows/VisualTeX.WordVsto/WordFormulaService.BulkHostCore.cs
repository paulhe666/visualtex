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
        if (source is null)
            throw new ArgumentNullException(nameof(source));
        if (prepared is null)
            throw new ArgumentNullException(nameof(prepared));

        Document? document = null;
        Selection? selection = null;
        Range? captured = null;
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

                    for (var blockIndex = 0;
                         blockIndex < source.Blocks.Count;
                         blockIndex++)
                    {
                        var block =
                            source.Blocks[blockIndex];
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
                                throw new InvalidDataException(
                                    $"缺少行间公式 {formulaRun.Id} 的渲染结果。");

                            formula.Session.DisplayMode =
                                "block";
                            formula.Session.Numbered =
                                source.NumberDisplayFormulas;
                            InsertPreparedHostCore(
                                document,
                                selection,
                                formula,
                                display: true,
                                insertedIds);

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
                                InsertNativeTextRun(
                                    document,
                                    selection,
                                    run);
                                continue;
                            }

                            if (!prepared.TryGetValue(
                                    run.Id,
                                    out var formula))
                                throw new InvalidDataException(
                                    $"缺少行内公式 {run.Id} 的渲染结果。");

                            var placeholderStart =
                                selection.Start;
                            selection.TypeText(
                                BulkInlineFormulaPlaceholder);
                            pending.Add((
                                placeholderStart,
                                formula));
                        }

                        selection.TypeParagraph();
                        var paragraphEnd =
                            selection.Start;
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
                                    throw new InvalidDataException(
                                        "批量导入的行内公式占位符发生了变化。");
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
                                var inserted =
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

                        MoveSelectionAfterBulkParagraph(
                            document,
                            selection,
                            paragraphStart);
                        ResetNextParagraphFormatting(
                            selection,
                            block.Kind,
                            nextKind);
                    }

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
            Release(captured);
            Release(selection);
            Release(document);
        }
    }

    private void InsertPreparedHostCore(
        Document document,
        Selection selection,
        PreparedWordBulkFormula formula,
        bool display,
        ICollection<string> insertedIds)
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
                        throw new InvalidDataException(
                            "批量导入的行间公式插入点不属于唯一段落。");
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
                    insertedRange);
            }
            finally
            {
                Release(insertedRange);
            }
        }
        finally
        {
            Release(insertion);
            Release(current);
        }
    }
}

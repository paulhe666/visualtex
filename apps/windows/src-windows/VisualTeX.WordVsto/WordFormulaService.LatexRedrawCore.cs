using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    internal WordLatexRedrawPlan CaptureLatexRedrawPlanCore(
        bool wholeDocument)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);

            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;
            if (!wholeDocument && scope.Start == scope.End)
                throw new InvalidOperationException(
                    "请先选择包含 LaTeX 代码的 Word 内容。");

            var sourceText = scope.Text ?? string.Empty;
            var spans = WordBulkImportParser.FindFormulaSpans(
                sourceText);
            if (spans.Count == 0)
                throw new InvalidDataException(
                    wholeDocument
                        ? "当前 Word 文档中没有找到可重绘的 LaTeX 公式。"
                        : "所选内容中没有找到可重绘的 LaTeX 公式。");

            var plan = new WordLatexRedrawPlan
            {
                DocumentId = DocumentIdentity(document),
                ScopeStart = scope.Start,
                ScopeEnd = scope.End,
                SourceText = sourceText,
                Targets = spans.Select(span =>
                    new WordLatexRedrawTarget
                    {
                        Id = span.Id,
                        RelativeStart = span.Start,
                        SourceLength = span.Length,
                        Latex = span.Latex,
                        DisplayMode = span.DisplayMode,
                    }).ToList(),
            };

            var offsets = BuildWordStoryOffsetIndex(
                sourceText);
            var contexts =
                BuildLatexRedrawSourceContexts(
                    sourceText,
                    plan.Targets);
            var sourceFontSizes =
                TryBuildWordOpenXmlFontSizeIndex(
                    scope,
                    sourceText);

            foreach (var target in plan.Targets
                         .OrderBy(item => item.RelativeStart))
            {
                var relativeEnd =
                    target.RelativeStart
                    + target.SourceLength;
                if (target.RelativeStart < 0
                    || relativeEnd > sourceText.Length)
                    throw new InvalidDataException(
                        "LaTeX formula span escaped the captured Word scope.");

                var sourceStart =
                    plan.ScopeStart
                    + offsets[target.RelativeStart];
                var sourceEnd =
                    plan.ScopeStart
                    + offsets[relativeEnd];
                if (sourceEnd <= sourceStart
                    || sourceStart < plan.ScopeStart
                    || sourceEnd > plan.ScopeEnd)
                    throw new InvalidDataException(
                        "LaTeX formula span could not be mapped to one Word range.");

                Range? exact = null;
                try
                {
                    exact = document.Range(
                        sourceStart,
                        sourceEnd);
                    var expected = sourceText.Substring(
                        target.RelativeStart,
                        target.SourceLength);
                    if (!string.Equals(
                            exact.Text ?? string.Empty,
                            expected,
                            StringComparison.Ordinal))
                        throw new InvalidDataException(
                            "Word text coordinates differ from the captured LaTeX span; redraw stopped before mutation.");

                    target.AbsoluteStart = sourceStart;
                    target.AbsoluteEnd = sourceEnd;

                    var context =
                        contexts[target.Id];
                    var display = string.Equals(
                        target.DisplayMode,
                        "block",
                        StringComparison.Ordinal);
                    target.PreserveDisplayParagraphBoundary =
                        display;
                    if (display)
                    {
                        var paragraphStart =
                            FindSourceParagraphStart(
                                sourceText,
                                target.RelativeStart);
                        var paragraphEnd =
                            FindSourceParagraphEnd(
                                sourceText,
                                relativeEnd);
                        target.DedicatedDisplayParagraph =
                            !context.HasVisibleSurroundingText
                            && target.RelativeStart
                                == paragraphStart
                            && relativeEnd
                                == paragraphEnd;
                    }

                    var contextPosition =
                        context.FontContextRelativePosition;
                    if (sourceFontSizes is not null
                        && contextPosition >= 0
                        && contextPosition
                            < sourceFontSizes.Length
                        && sourceFontSizes[
                            contextPosition] is double indexedSize
                        && indexedSize > 0)
                    {
                        target.FontSizePt = indexedSize;
                    }
                    else
                    {
                        target.FontSizePt =
                            ResolveSourceFormulaFontSize(
                                document,
                                exact,
                                display);
                    }
                }
                finally { Release(exact); }
            }

            return plan;
        }
        finally
        {
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    internal WordLatexRedrawResult ApplyLatexRedrawPlanCore(
        WordLatexRedrawPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (plan is null)
            throw new ArgumentNullException(nameof(plan));
        if (prepared is null)
            throw new ArgumentNullException(nameof(prepared));

        Document? document = null;
        Range? validation = null;
        WordViewState? viewState = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                plan.DocumentId);

            validation = document.Range(
                plan.ScopeStart,
                plan.ScopeEnd);
            if (!string.Equals(
                    validation.Text ?? string.Empty,
                    plan.SourceText,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "渲染期间 Word 内容发生了变化。为避免替换错误位置，本次重绘已停止。");

            var missing = plan.Targets
                .Where(target =>
                    !prepared.ContainsKey(target.Id))
                .Select(target => target.Id)
                .ToArray();
            if (missing.Length > 0)
                throw new InvalidDataException(
                    "LaTeX 重绘缺少一个或多个已渲染目标。");

            foreach (var target in plan.Targets)
            {
                var formula = prepared[target.Id];
                formula.Session.Numbered =
                    plan.NumberDisplayFormulas
                    && string.Equals(
                        target.DisplayMode,
                        "block",
                        StringComparison.Ordinal);
                formula.Session.DisplayMode =
                    target.DisplayMode;
                formula.Session.FontSizePt =
                    FormulaFontSize.Normalize(
                        target.FontSizePt);
            }

            viewState = CaptureViewState();

            var result =
                WordFormulaMutationTransaction.Execute(
                    _application,
                    document,
                    "VisualTeX Redraw LaTeX",
                    () => ApplyLatexRedrawInsideTransaction(
                        document,
                        plan,
                        prepared));

            try
            {
                RestoreViewState(
                    document,
                    viewState,
                    preferredSelection: null);
            }
            catch
            {
                // View restoration is transient UI state and cannot invalidate a
                // successfully committed document mutation.
            }

            return result;
        }
        finally
        {
            Release(validation);
            Release(document);
        }
    }

    private WordLatexRedrawResult
        ApplyLatexRedrawInsideTransaction(
            Document document,
            WordLatexRedrawPlan plan,
            IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        var result = new WordLatexRedrawResult();
        var total = 0L;
        var maximum = 0L;

        foreach (var target in plan.Targets
                     .OrderByDescending(
                         item => item.AbsoluteStart))
        {
            var formula = prepared[target.Id];
            Range? source = null;
            Range? isolated = null;
            Range? insertion = null;
            try
            {
                source = document.Range(
                    target.AbsoluteStart,
                    target.AbsoluteEnd);
                var expected = plan.SourceText.Substring(
                    target.RelativeStart,
                    target.SourceLength);
                if (!string.Equals(
                        source.Text ?? string.Empty,
                        expected,
                        StringComparison.Ordinal))
                    throw new InvalidDataException(
                        $"LaTeX source changed before redraw: {target.Latex}");

                var display = string.Equals(
                    target.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase);

                if (display
                    && !target.DedicatedDisplayParagraph)
                {
                    isolated =
                        IsolateDisplaySourceParagraph(
                            document,
                            source);
                    Release(source);
                    source = isolated;
                    isolated = null;
                }

                var insertionStart = source.Start;
                source.Delete();
                insertion = document.Range(
                    insertionStart,
                    insertionStart);

                var watch = Stopwatch.StartNew();
                var targetKind =
                    ObjectModeToHostKind(
                        formula.Session.ObjectMode);

                if (targetKind is
                    WordFormulaHostKind.Omml
                    or WordFormulaHostKind.VisualTeX)
                {
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
                    result.FormulaIds.Add(
                        inserted.FormulaId
                        ?? formula.Session.FormulaId);
                }
                else if (targetKind ==
                         WordFormulaHostKind.MathType)
                {
                    var mathMl = formula.MathMl
                        ?? throw new InvalidDataException(
                            $"MathType redraw target '{target.Latex}' has no MathML.");
                    var session = formula.Session;
                    var oldSourceObjectId =
                        session.SourceObjectId;
                    try
                    {
                        session.SourceObjectId =
                            RangeReference(insertion);
                        var native =
                            formula.MathTypeNativePreview;
                        _ = InsertMathTypeOle(
                            session,
                            mathMl,
                            formula.EmfPath,
                            isolatedNativePreviewWmfPath:
                                native?.WmfPath,
                            isolatedNativePreviewWidthPt:
                                native?.WidthPt ?? 0,
                            isolatedNativePreviewHeightPt:
                                native?.HeightPt ?? 0,
                            isolatedNativePreviewWordPosition:
                                native?.WordPosition ?? 0,
                            isolatedNativePreviewAttempted:
                                formula.MathTypeNativePreviewAttempted,
                            preserveExistingDisplayParagraphBoundary:
                                display,
                            preserveCapturedInsertion:
                                true);
                        result.FormulaIds.Add(
                            session.FormulaId);
                    }
                    finally
                    {
                        session.SourceObjectId =
                            oldSourceObjectId;
                    }
                }
                else
                {
                    throw new NotSupportedException(
                        $"Unsupported redraw target {targetKind}.");
                }

                watch.Stop();
                total += watch.ElapsedMilliseconds;
                maximum = Math.Max(
                    maximum,
                    watch.ElapsedMilliseconds);
                result.FormulaCount++;
            }
            finally
            {
                Release(insertion);
                Release(isolated);
                Release(source);
            }
        }

        result.TotalInsertMilliseconds = total;
        result.MaxInsertMilliseconds = maximum;
        return result;
    }
}

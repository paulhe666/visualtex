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
        var captureWatch = Stopwatch.StartNew();
        var captureStageStart = 0L;
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
            plan.CaptureTimings["snapshot"] = captureWatch.ElapsedMilliseconds;

            captureStageStart = captureWatch.ElapsedMilliseconds;
            ResolveLatexRedrawTargetCoordinates(
                document,
                scope,
                plan,
                allowStoredCoordinates: false);
            plan.CaptureTimings["coordinates"] =
                captureWatch.ElapsedMilliseconds - captureStageStart;

            captureStageStart = captureWatch.ElapsedMilliseconds;
            var contexts =
                BuildLatexRedrawSourceContexts(
                    sourceText,
                    plan.Targets);
            plan.CaptureTimings["contexts"] =
                captureWatch.ElapsedMilliseconds - captureStageStart;

            captureStageStart = captureWatch.ElapsedMilliseconds;
            // Most ordinary Word redraw scopes use one body font size. Prove that
            // cheaply once through Range.Font before asking Word to serialize the
            // complete scope as Flat OPC merely to recover per-character sizes.
            // Mixed-format scopes still fall back to the exact WordOpenXML index.
            float? uniformScopeFontSize = null;
            Microsoft.Office.Interop.Word.Font? scopeFont = null;
            try
            {
                scopeFont = scope.Font;
                if (TryResolveWordFontSize(
                        scopeFont.Size,
                        out var resolvedUniformSize))
                    uniformScopeFontSize = resolvedUniformSize;
            }
            catch { }
            finally { Release(scopeFont); }

            var fontIndexReconstructedLength = -1;
            var fontIndexMismatchIndex = -1;
            var sourceFontSizes = uniformScopeFontSize.HasValue
                ? null
                : TryBuildWordOpenXmlFontSizeIndex(
                    scope,
                    sourceText,
                    out fontIndexReconstructedLength,
                    out fontIndexMismatchIndex);
            plan.CaptureTimings["fontIndex"] =
                captureWatch.ElapsedMilliseconds - captureStageStart;
            plan.CaptureTimings["fontIndexReconstructedLength"] =
                fontIndexReconstructedLength;
            plan.CaptureTimings["fontIndexSourceLength"] = sourceText.Length;
            plan.CaptureTimings["fontIndexMismatchIndex"] =
                fontIndexMismatchIndex;

            captureStageStart = captureWatch.ElapsedMilliseconds;
            var fontFallbackCount = 0L;
            var fontFallbackMilliseconds = 0L;
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
                if (uniformScopeFontSize.HasValue)
                {
                    target.FontSizePt = uniformScopeFontSize.Value;
                }
                else if (sourceFontSizes is not null
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
                    // The complete gap/source/tail chain above already proved the
                    // target's exact text and Word coordinates. Create a COM Range
                    // only when mixed formatting prevents the bulk font indexes
                    // from resolving this target; do not re-read every source text
                    // merely to repeat the same validation.
                    Range? exact = null;
                    var fallbackWatch = Stopwatch.StartNew();
                    try
                    {
                        fontFallbackCount++;
                        exact = document.Range(
                            target.AbsoluteStart,
                            target.AbsoluteEnd);
                        target.FontSizePt =
                            ResolveSourceFormulaFontSize(
                                document,
                                exact,
                                display);
                    }
                    finally
                    {
                        fallbackWatch.Stop();
                        fontFallbackMilliseconds += fallbackWatch.ElapsedMilliseconds;
                        Release(exact);
                    }
                }
            }
            plan.CaptureTimings["targets"] =
                captureWatch.ElapsedMilliseconds - captureStageStart;
            plan.CaptureTimings["fontFallbackCount"] = fontFallbackCount;
            plan.CaptureTimings["fontFallback"] = fontFallbackMilliseconds;
            plan.CaptureTimings["total"] = captureWatch.ElapsedMilliseconds;

            return plan;
        }
        finally
        {
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    // LaTeX is parsed from Range.Text, whose logical-character indices
    // are not Word story coordinates when the scope contains hidden host structure
    // such as fields, native math, or other non-text story positions. Never infer
    // those hidden positions here. Use direct story coordinates only when the whole
    // scope proves both coordinate spaces have identical cardinality; otherwise ask
    // Word's Characters collection for the logical boundaries. In both cases the
    // mapped formulas are accepted only if the complete gap/source/tail text chain
    // exactly reproduces the captured Range.Text snapshot.
    private static void ResolveLatexRedrawTargetCoordinates(
        Document document,
        Range scope,
        WordLatexRedrawPlan plan,
        bool allowStoredCoordinates)
    {
        var liveText =
            scope.Text ?? string.Empty;
        if (!string.Equals(
                liveText,
                plan.SourceText,
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The Word text changed after the LaTeX redraw plan was captured.");

        if (allowStoredCoordinates
            && TryValidateLatexRedrawCoordinateChain(
                document,
                scope,
                plan,
                plan.Targets
                    .Select(target =>
                        (
                            Target: target,
                            Start: target.AbsoluteStart,
                            End: target.AbsoluteEnd))
                    .ToArray()))
            return;

        var logicalOffsets =
            BuildWordStoryOffsetIndex(
                plan.SourceText);
        var logicalCharacterCount =
            logicalOffsets[
                plan.SourceText.Length];
        var directCoordinatesProven =
            scope.End - scope.Start
            == logicalCharacterCount;

        Characters? characters = null;
        var mapped =
            new List<(
                WordLatexRedrawTarget Target,
                int Start,
                int End)>(
                plan.Targets.Count);
        try
        {
            if (!directCoordinatesProven)
            {
                characters =
                    scope.Characters;
                if (characters.Count
                    != logicalCharacterCount)
                    throw new InvalidDataException(
                        "Word logical-character indexing does not match the captured redraw text.");
            }

            foreach (var target in plan.Targets
                         .OrderBy(item =>
                             item.RelativeStart))
            {
                var relativeEnd =
                    target.RelativeStart
                    + target.SourceLength;
                if (target.RelativeStart < 0
                    || target.SourceLength <= 0
                    || relativeEnd
                        > plan.SourceText.Length)
                    throw new InvalidDataException(
                        "LaTeX formula span escaped the captured Word scope.");

                var logicalStart =
                    logicalOffsets[
                        target.RelativeStart];
                var logicalEnd =
                    logicalOffsets[
                        relativeEnd];
                if (logicalEnd <= logicalStart)
                    throw new InvalidDataException(
                        "LaTeX formula span has no Word logical-character extent.");

                int sourceStart;
                int sourceEnd;
                Range? first = null;
                Range? last = null;
                try
                {
                    if (directCoordinatesProven)
                    {
                        sourceStart =
                            scope.Start
                            + logicalStart;
                        sourceEnd =
                            scope.Start
                            + logicalEnd;
                    }
                    else
                    {
                        first =
                            characters![
                                logicalStart + 1];
                        last =
                            characters[
                                logicalEnd];
                        sourceStart =
                            first.Start;
                        sourceEnd =
                            last.End;
                    }
                }
                finally
                {
                    Release(last);
                    Release(first);
                }

                if (sourceStart < scope.Start
                    || sourceEnd <= sourceStart
                    || sourceEnd > scope.End)
                    throw new InvalidDataException(
                        "Word returned an invalid range for a captured LaTeX formula.");

                mapped.Add(
                    (
                        Target: target,
                        Start: sourceStart,
                        End: sourceEnd));
            }

            if (!TryValidateLatexRedrawCoordinateChain(
                    document,
                    scope,
                    plan,
                    mapped))
                throw new InvalidDataException(
                    "Word logical-character mapping did not reproduce the captured LaTeX source sequence exactly.");

            foreach (var item in mapped)
            {
                item.Target.AbsoluteStart =
                    item.Start;
                item.Target.AbsoluteEnd =
                    item.End;
            }
        }
        finally
        {
            Release(characters);
        }
    }

    private static bool TryValidateLatexRedrawCoordinateChain(
        Document document,
        Range scope,
        WordLatexRedrawPlan plan,
        IReadOnlyList<(
            WordLatexRedrawTarget Target,
            int Start,
            int End)> mapped)
    {
        if (mapped.Count
            != plan.Targets.Count)
            return false;

        var logicalCursor = 0;
        var wordCursor =
            scope.Start;
        Range? gap = null;
        Range? source = null;
        Range? tail = null;
        try
        {
            foreach (var item in mapped
                         .OrderBy(candidate =>
                             candidate.Target
                                 .RelativeStart))
            {
                var target =
                    item.Target;
                var relativeEnd =
                    target.RelativeStart
                    + target.SourceLength;
                if (target.RelativeStart
                        < logicalCursor
                    || target.RelativeStart < 0
                    || target.SourceLength <= 0
                    || relativeEnd
                        > plan.SourceText.Length
                    || item.Start < wordCursor
                    || item.Start < scope.Start
                    || item.End <= item.Start
                    || item.End > scope.End)
                    return false;

                Release(gap);
                gap = document.Range(
                    wordCursor,
                    item.Start);
                var expectedGap =
                    plan.SourceText.Substring(
                        logicalCursor,
                        target.RelativeStart
                        - logicalCursor);
                if (!string.Equals(
                        gap.Text ?? string.Empty,
                        expectedGap,
                        StringComparison.Ordinal))
                    return false;

                Release(source);
                source = document.Range(
                    item.Start,
                    item.End);
                var expectedSource =
                    plan.SourceText.Substring(
                        target.RelativeStart,
                        target.SourceLength);
                if (!string.Equals(
                        source.Text ?? string.Empty,
                        expectedSource,
                        StringComparison.Ordinal))
                    return false;

                logicalCursor =
                    relativeEnd;
                wordCursor =
                    item.End;
            }

            tail = document.Range(
                wordCursor,
                scope.End);
            var expectedTail =
                plan.SourceText.Substring(
                    logicalCursor);
            return string.Equals(
                tail.Text ?? string.Empty,
                expectedTail,
                StringComparison.Ordinal);
        }
        catch (COMException)
        {
            return false;
        }
        finally
        {
            Release(tail);
            Release(source);
            Release(gap);
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

        var applyWatch = Stopwatch.StartNew();
        var applyStageStart = 0L;
        var batchTimings = new Dictionary<string, long>(StringComparer.Ordinal);
        Document? document = null;
        Range? validation = null;
        WordViewState? viewState = null;
        NativeOmmlScreenUpdatingScope? screenUpdatingScope = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                plan.DocumentId);

            applyStageStart = applyWatch.ElapsedMilliseconds;
            validation = document.Range(
                plan.ScopeStart,
                plan.ScopeEnd);
            if (!string.Equals(
                    validation.Text ?? string.Empty,
                    plan.SourceText,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "渲染期间 Word 内容发生了变化。为避免替换错误位置，本次重绘已停止。");
            batchTimings["snapshotValidation"] =
                applyWatch.ElapsedMilliseconds - applyStageStart;

            // Capture already proved the complete logical gap/source/tail chain
            // and stored exact Word story coordinates. Rendering does not mutate
            // Word. Before mutation we prove the whole captured scope text is still
            // byte-for-byte identical above, and the ordered writer verifies
            // every target's exact source text immediately before clearing it.
            // Re-running the full coordinate-chain COM walk here duplicated the
            // same proof once per batch without adding a stronger postcondition.
            batchTimings["coordinateValidation"] = 0;

            applyStageStart = applyWatch.ElapsedMilliseconds;
            var missing = plan.Targets
                .Where(target =>
                    !prepared.ContainsKey(target.Id))
                .Select(target => target.Id)
                .ToArray();
            if (missing.Length > 0)
                throw new InvalidDataException(
                    "LaTeX 重绘缺少一个或多个已渲染目标。");

            // MathType's native materialization can temporarily activate helper
            // documents, so retain its existing screen-update guard. VisualTeX
            // and OMML mutate only the target document; freezing their visible
            // window provided little benefit once mutations run in document order
            // and can leave older Word builds without a rebuilt document surface.
            var usesMathType = plan.Targets.Any(target =>
                ObjectModeToHostKind(
                    prepared[target.Id].Session.ObjectMode)
                == WordFormulaHostKind.MathType);
            if (usesMathType)
                screenUpdatingScope =
                    NativeOmmlScreenUpdatingScope.Suspend(
                        _application);

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
            batchTimings["prepareTargets"] =
                applyWatch.ElapsedMilliseconds - applyStageStart;

            applyStageStart = applyWatch.ElapsedMilliseconds;
            viewState = CaptureViewState();
            batchTimings["captureView"] =
                applyWatch.ElapsedMilliseconds - applyStageStart;

            applyStageStart = applyWatch.ElapsedMilliseconds;
            var result =
                ApplyLatexRedrawWithIndependentTransactions(
                    document,
                    plan,
                    prepared);
            batchTimings["transaction"] =
                applyWatch.ElapsedMilliseconds - applyStageStart;

            applyStageStart = applyWatch.ElapsedMilliseconds;
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
            batchTimings["restoreView"] =
                applyWatch.ElapsedMilliseconds - applyStageStart;
            batchTimings["total"] = applyWatch.ElapsedMilliseconds;
            foreach (var timing in batchTimings)
                result.BatchTimings[timing.Key] = timing.Value;

            return result;
        }
        finally
        {
            screenUpdatingScope?.Dispose();
            Release(validation);
            Release(document);
        }
    }

    private WordLatexRedrawResult
        ApplyLatexRedrawWithIndependentTransactions(
            Document document,
            WordLatexRedrawPlan plan,
            IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        var result = new WordLatexRedrawResult();
        var successfulTargetIds =
            new HashSet<string>(StringComparer.Ordinal);
        var sourceMutationTotal = 0L;
        var total = 0L;
        var maximum = 0L;
        var coordinateRebaseTotal = 0L;
        var cumulativeCoordinateDelta = 0;
        var previousMainStoryEnd = ReadMainStoryEnd(document);
        var displayTabStopsBySection =
            new Dictionary<int, (float Center, float Right)>();
        var mathTypeDisplayWidthsBySection =
            new Dictionary<int, float>();
        var mathTypeNativeStylesPrepared = false;
        var mathTypeInlineFinalizationTotal = 0L;
        var mathTypeInlineIdentityIndexTotal = 0L;
        var mathTypeInlineBoundaryMergeTotal = 0L;
        var mathTypeInlineTypingResetTotal = 0L;

        // Process in document order. Every formula owns one independent Word undo
        // transaction. A failed target is rolled back to its exact source and does
        // not alter accumulated batch state; a successful target contributes its
        // result only after Word has committed the custom undo record.
        foreach (var target in plan.Targets
                     .OrderBy(item => item.AbsoluteStart))
        {
            var formula = prepared[target.Id];

            // Source identity and target kind are pure preflight checks. Keep
            // them outside the custom undo record: if the document changed under
            // us, there is no redraw mutation to undo and therefore no risk of
            // accidentally undoing an earlier user edit.
            ValidateLatexRedrawTargetSourceBeforeTransaction(
                document,
                plan,
                target,
                cumulativeCoordinateDelta);
            var targetKind =
                ObjectModeToHostKind(
                    formula.Session.ObjectMode);

            LatexRedrawTargetMutationResult targetResult;
            try
            {
                targetResult =
                    WordFormulaMutationTransaction.Execute(
                        _application,
                        document,
                        "VisualTeX Redraw LaTeX Formula",
                        () => ApplyLatexRedrawTargetInsideTransaction(
                            document,
                            target,
                            formula,
                            targetKind,
                            cumulativeCoordinateDelta,
                            displayTabStopsBySection,
                            mathTypeDisplayWidthsBySection,
                            mathTypeNativeStylesPrepared));
            }
            catch (Exception error)
            {
                VerifyFailedLatexRedrawTargetRollback(
                    document,
                    plan,
                    target,
                    cumulativeCoordinateDelta,
                    previousMainStoryEnd,
                    error);

                var detail = string.IsNullOrWhiteSpace(error.Message)
                    ? error.GetType().Name
                    : error.Message.Trim();
                result.Failures.Add(
                    (target.Id, target.Latex, detail));
                continue;
            }

            cumulativeCoordinateDelta = checked(
                cumulativeCoordinateDelta
                + targetResult.MainStoryEnd
                - previousMainStoryEnd);
            previousMainStoryEnd = targetResult.MainStoryEnd;
            coordinateRebaseTotal +=
                targetResult.CoordinateRebaseMilliseconds;

            sourceMutationTotal +=
                targetResult.SourceMutationMilliseconds;
            total += targetResult.InsertMilliseconds;
            maximum = Math.Max(
                maximum,
                targetResult.InsertMilliseconds);

            result.SourceMutationMilliseconds.Add(
                targetResult.SourceMutationMilliseconds);
            result.InsertMilliseconds.Add(
                targetResult.InsertMilliseconds);
            if (targetResult.InsertMetricDetail is not null)
                result.InsertMetricDetails.Add(
                    targetResult.InsertMetricDetail);
            result.FormulaIds.Add(
                targetResult.FormulaId);
            result.FormulaCount++;
            successfulTargetIds.Add(target.Id);

            if (targetResult.MathTypeNativeStylesPrepared)
                mathTypeNativeStylesPrepared = true;
            mathTypeInlineFinalizationTotal +=
                targetResult.MathTypeInlineFinalizationMilliseconds;
            mathTypeInlineIdentityIndexTotal +=
                targetResult.MathTypeInlineIdentityIndexMilliseconds;
            mathTypeInlineBoundaryMergeTotal +=
                targetResult.MathTypeInlineBoundaryMergeMilliseconds;
            mathTypeInlineTypingResetTotal +=
                targetResult.MathTypeInlineTypingResetMilliseconds;
        }

        if (mathTypeInlineFinalizationTotal > 0)
        {
            result.BatchTimings["mathTypeInlineFinalization"] =
                mathTypeInlineFinalizationTotal;
            result.BatchTimings["mathTypeInlineIdentityIndex"] =
                mathTypeInlineIdentityIndexTotal;
            result.BatchTimings["mathTypeInlineBoundaryMerge"] =
                mathTypeInlineBoundaryMergeTotal;
            result.BatchTimings["mathTypeInlineTypingReset"] =
                mathTypeInlineTypingResetTotal;
        }

        // Number only successfully committed display formulas. Number refresh is
        // itself isolated: if Word rejects the refresh, only that refresh is
        // rolled back while all successfully redrawn formulas remain intact.
        if (plan.NumberDisplayFormulas
            && result.FormulaCount > 0)
        {
            try
            {
                WordFormulaMutationTransaction.Execute(
                    _application,
                    document,
                    "VisualTeX Finalize Redraw Numbering",
                    () =>
                    {
                        var targetKinds =
                            plan.Targets
                                .Where(target =>
                                    successfulTargetIds.Contains(target.Id)
                                    && string.Equals(
                                        target.DisplayMode,
                                        "block",
                                        StringComparison.OrdinalIgnoreCase))
                                .Select(target =>
                                    ObjectModeToHostKind(
                                        prepared[target.Id]
                                            .Session
                                            .ObjectMode))
                                .Distinct()
                                .ToArray();

                        if (targetKinds.Any(kind =>
                                kind == WordFormulaHostKind.MathType))
                        {
                            _ =
                                MathTypeEquationNumbering
                                    .UpdateEquationNumbers(
                                        document);
                        }

                        if (targetKinds.Any(kind =>
                                kind is
                                    WordFormulaHostKind.Omml
                                    or WordFormulaHostKind.VisualTeX))
                        {
                            _ =
                                WordFormulaNumberingKernel
                                    .RefreshCanonicalNumbers(
                                        document);
                        }

                        return true;
                    });
            }
            catch (Exception error)
            {
                var detail = string.IsNullOrWhiteSpace(error.Message)
                    ? error.GetType().Name
                    : error.Message.Trim();
                result.Warnings.Add(
                    $"Equation numbering finalization failed and was rolled back: {detail}");
            }
        }

        result.TotalSourceMutationMilliseconds =
            sourceMutationTotal;
        result.BatchTimings["coordinateRebase"] =
            coordinateRebaseTotal;
        result.TotalInsertMilliseconds = total;
        result.MaxInsertMilliseconds = maximum;
        return result;
    }

    private sealed class LatexRedrawTargetMutationResult
    {
        internal string FormulaId { get; set; } = string.Empty;
        internal int MainStoryEnd { get; set; }
        internal long CoordinateRebaseMilliseconds { get; set; }
        internal long SourceMutationMilliseconds { get; set; }
        internal long InsertMilliseconds { get; set; }
        internal string? InsertMetricDetail { get; set; }
        internal bool MathTypeNativeStylesPrepared { get; set; }
        internal long MathTypeInlineFinalizationMilliseconds { get; set; }
        internal long MathTypeInlineIdentityIndexMilliseconds { get; set; }
        internal long MathTypeInlineBoundaryMergeMilliseconds { get; set; }
        internal long MathTypeInlineTypingResetMilliseconds { get; set; }
    }

    private LatexRedrawTargetMutationResult
        ApplyLatexRedrawTargetInsideTransaction(
            Document document,
            WordLatexRedrawTarget target,
            PreparedWordBulkFormula formula,
            WordFormulaHostKind targetKind,
            int cumulativeCoordinateDelta,
            Dictionary<int, (float Center, float Right)> displayTabStopsBySection,
            Dictionary<int, float> mathTypeDisplayWidthsBySection,
            bool mathTypeNativeStylesPrepared)
    {
        Range? source = null;
        Range? isolated = null;
        Range? insertion = null;
        WordOperationMetrics? insertionMetrics = null;
        var deferredMathTypeInlineParagraphRepairs =
            new List<(string FormulaId, string RightBoundaryWhitespace)>();
        try
        {
            if (!string.IsNullOrWhiteSpace(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_VSTO_REDRAW_PERF_LOG")))
                insertionMetrics = new WordOperationMetrics();

            var sourceMutationWatch = Stopwatch.StartNew();
            using (WordOperationMetrics.Measure(
                       "Redraw.Source.CreateRange"))
            {
                var adjustedStart = checked(
                    target.AbsoluteStart
                    + cumulativeCoordinateDelta);
                var adjustedEnd = checked(
                    target.AbsoluteEnd
                    + cumulativeCoordinateDelta);
                source = document.Range(
                    adjustedStart,
                    adjustedEnd);
            }

            var display = string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase);
            if (display
                && !target.DedicatedDisplayParagraph)
            {
                using (WordOperationMetrics.Measure(
                           "Redraw.Source.IsolateDisplay"))
                {
                    isolated =
                        IsolateDisplaySourceParagraph(
                            document,
                            source);
                }
                Release(source);
                source = isolated;
                isolated = null;
            }

            (float Center, float Right)? displayTabStops = null;
            float? mathTypeDisplayColumnWidth = null;
            if (display
                && targetKind == WordFormulaHostKind.VisualTeX)
            {
                using (WordOperationMetrics.Measure(
                           "Redraw.Source.ResolveDisplayTabStops"))
                {
                    displayTabStops = ResolveBatchDisplayTabStops(
                        document,
                        source,
                        displayTabStopsBySection);
                }
            }
            else if (display
                     && targetKind == WordFormulaHostKind.MathType)
            {
                using (WordOperationMetrics.Measure(
                           "Redraw.Source.ResolveMathTypeDisplayWidth"))
                {
                    mathTypeDisplayColumnWidth =
                        ResolveBatchMathTypeDisplayColumnWidth(
                            document,
                            source,
                            mathTypeDisplayWidthsBySection);
                }
            }

            var insertionStart = source.Start;
            // Word Range.Delete can consume a following inline character under
            // Smart Cut/Paste. The captured range is exact, so clear only it.
            using (WordOperationMetrics.Measure(
                       "Redraw.Source.Clear"))
            {
                source.Text = string.Empty;
            }
            using (WordOperationMetrics.Measure(
                       "Redraw.Source.CreateInsertion"))
            {
                insertion = document.Range(
                    insertionStart,
                    insertionStart);
            }
            sourceMutationWatch.Stop();

            var insertWatch = Stopwatch.StartNew();
            string formulaId;
            var preparedMathTypeNativeStyles = false;
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
                request.BatchMutation = true;
                if (displayTabStops.HasValue)
                {
                    request.DisplayCenterTabPoints =
                        displayTabStops.Value.Center;
                    request.DisplayRightTabPoints =
                        displayTabStops.Value.Right;
                }

                var inserted =
                    WordFormulaHostMutationKernel
                        .InsertInActiveTransaction(
                            _application,
                            document,
                            insertion,
                            request);
                formulaId =
                    inserted.FormulaId
                    ?? formula.Session.FormulaId;
            }
            else if (targetKind == WordFormulaHostKind.MathType)
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
                            true,
                        boundedBatchMutation:
                            true,
                        exactInsertionRange:
                            insertion,
                        mathTypeNativeStylesPrepared:
                            mathTypeNativeStylesPrepared,
                        deferInlineParagraphRepair:
                            display
                                ? null
                                : (id, rightWhitespace) =>
                                    deferredMathTypeInlineParagraphRepairs.Add(
                                        (id, rightWhitespace)),
                        knownDisplayColumnWidth:
                            mathTypeDisplayColumnWidth,
                        preparedStandaloneGenerated:
                            formula.MathTypeStandaloneGenerated,
                        preparedStandaloneCompoundFile:
                            formula.MathTypeStandaloneCompoundFile,
                        preparedStandaloneSemanticSignature:
                            formula.MathTypeStandaloneSemanticSignature);
                    formulaId = session.FormulaId;
                    preparedMathTypeNativeStyles = display;
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
            insertWatch.Stop();

            long finalizationMilliseconds = 0;
            long identityIndexMilliseconds = 0;
            long boundaryMergeMilliseconds = 0;
            long typingResetMilliseconds = 0;
            if (deferredMathTypeInlineParagraphRepairs.Count > 0)
            {
                var finalizationWatch = Stopwatch.StartNew();
                FinalizeDeferredMathTypeInlineParagraphRepairs(
                    document,
                    deferredMathTypeInlineParagraphRepairs,
                    out identityIndexMilliseconds,
                    out boundaryMergeMilliseconds,
                    out typingResetMilliseconds);
                finalizationWatch.Stop();
                finalizationMilliseconds =
                    finalizationWatch.ElapsedMilliseconds;
            }

            var coordinateRebaseWatch = Stopwatch.StartNew();
            var mainStoryEnd = ReadMainStoryEnd(document);
            coordinateRebaseWatch.Stop();

            string? insertMetricDetail = null;
            if (insertionMetrics is not null)
            {
                insertionMetrics.Dispose();
                insertMetricDetail =
                    string.Join(
                        " | ",
                        insertionMetrics.Entries
                            .OrderBy(entry => entry.Key)
                            .Select(entry =>
                                $"{entry.Key}={entry.Value.Milliseconds:0.0}ms/{entry.Value.Calls}"));
                insertionMetrics = null;
            }

            return new LatexRedrawTargetMutationResult
            {
                FormulaId = formulaId,
                MainStoryEnd = mainStoryEnd,
                CoordinateRebaseMilliseconds =
                    coordinateRebaseWatch.ElapsedMilliseconds,
                SourceMutationMilliseconds =
                    sourceMutationWatch.ElapsedMilliseconds,
                InsertMilliseconds =
                    insertWatch.ElapsedMilliseconds,
                InsertMetricDetail = insertMetricDetail,
                MathTypeNativeStylesPrepared =
                    preparedMathTypeNativeStyles,
                MathTypeInlineFinalizationMilliseconds =
                    finalizationMilliseconds,
                MathTypeInlineIdentityIndexMilliseconds =
                    identityIndexMilliseconds,
                MathTypeInlineBoundaryMergeMilliseconds =
                    boundaryMergeMilliseconds,
                MathTypeInlineTypingResetMilliseconds =
                    typingResetMilliseconds,
            };
        }
        finally
        {
            insertionMetrics?.Dispose();
            Release(insertion);
            Release(isolated);
            Release(source);
        }
    }

    private static void ValidateLatexRedrawTargetSourceBeforeTransaction(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        int cumulativeCoordinateDelta)
    {
        Range? source = null;
        try
        {
            var adjustedStart = checked(
                target.AbsoluteStart
                + cumulativeCoordinateDelta);
            var adjustedEnd = checked(
                target.AbsoluteEnd
                + cumulativeCoordinateDelta);
            source = document.Range(
                adjustedStart,
                adjustedEnd);
            var expected = plan.SourceText.Substring(
                target.RelativeStart,
                target.SourceLength);
            if (!string.Equals(
                    source.Text ?? string.Empty,
                    expected,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"LaTeX source changed before redraw: {target.Latex}");
        }
        finally
        {
            Release(source);
        }
    }

    private static void VerifyFailedLatexRedrawTargetRollback(
        Document document,
        WordLatexRedrawPlan plan,
        WordLatexRedrawTarget target,
        int cumulativeCoordinateDelta,
        int expectedMainStoryEnd,
        Exception primaryError)
    {
        Range? restoredSource = null;
        try
        {
            var restoredMainStoryEnd =
                ReadMainStoryEnd(document);
            if (restoredMainStoryEnd != expectedMainStoryEnd)
                throw new InvalidDataException(
                    $"Failed LaTeX redraw rollback changed the Word story length for '{target.Latex}'.");

            var adjustedStart = checked(
                target.AbsoluteStart
                + cumulativeCoordinateDelta);
            var adjustedEnd = checked(
                target.AbsoluteEnd
                + cumulativeCoordinateDelta);
            restoredSource = document.Range(
                adjustedStart,
                adjustedEnd);
            var expected = plan.SourceText.Substring(
                target.RelativeStart,
                target.SourceLength);
            if (!string.Equals(
                    restoredSource.Text ?? string.Empty,
                    expected,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    $"Failed LaTeX redraw rollback did not restore the exact source for '{target.Latex}'.");
        }
        catch (Exception verificationError)
        {
            throw new AggregateException(
                $"LaTeX redraw failed for '{target.Latex}', and its rollback could not be verified. The batch was stopped to protect later formula positions.",
                primaryError,
                verificationError);
        }
        finally
        {
            Release(restoredSource);
        }
    }

    private static int ReadMainStoryEnd(Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            return content.End;
        }
        finally
        {
            Release(content);
        }
    }

    private static void FinalizeDeferredMathTypeInlineParagraphRepairs(
        Document document,
        IReadOnlyList<(string FormulaId, string RightBoundaryWhitespace)> repairs,
        out long identityIndexMilliseconds,
        out long boundaryMergeMilliseconds,
        out long typingResetMilliseconds)
    {
        identityIndexMilliseconds = 0;
        boundaryMergeMilliseconds = 0;
        typingResetMilliseconds = 0;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? hostRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var repair in repairs)
            {
                Release(hostRange);
                hostRange = null;
                Release(bookmark);
                bookmark = null;

                var identityWatch = Stopwatch.StartNew();
                var bookmarkName =
                    WordFormulaIdentityStore.MathTypeBookmarkName(
                        repair.FormulaId);
                bookmark = bookmarks[bookmarkName];
                hostRange = bookmark.Range.Duplicate;
                if (hostRange.End <= hostRange.Start)
                    throw new InvalidDataException(
                        $"Deferred MathType inline repair {repair.FormulaId} has an empty identity range.");
                identityWatch.Stop();
                identityIndexMilliseconds += identityWatch.ElapsedMilliseconds;

                var mergeWatch = Stopwatch.StartNew();
                MergeDeferredMathTypeInlineParagraphBoundary(
                    document,
                    hostRange,
                    repair.RightBoundaryWhitespace);
                mergeWatch.Stop();
                boundaryMergeMilliseconds += mergeWatch.ElapsedMilliseconds;

                // The bounded batch path sets only the EMBED result character's
                // vertical position. Its field instruction and surrounding prose
                // remain at baseline, and the merge above removes the temporary
                // InsertXML paragraph mark. The surviving paragraph mark belongs
                // to the original prose paragraph and was never mutated, so a
                // second paragraph/typing-format walk here adds no postcondition.

                // Do not retain live Range COM objects across the next structural
                // paragraph merge. Releasing each exact host immediately avoids a
                // large Word resynchronization cost when many stale ranges are
                // released together at the end of the batch.
                Release(hostRange);
                hostRange = null;
                Release(bookmark);
                bookmark = null;
            }
        }
        finally
        {
            Release(hostRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void MergeDeferredMathTypeInlineParagraphBoundary(
        Document document,
        Range hostRange,
        string preservedRightBoundaryWhitespace)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? followingProbe = null;
        Range? mergeRange = null;
        Range? content = null;
        try
        {
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A deferred MathType inline OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End <= paragraphRange.Start)
                throw new InvalidDataException(
                    "A deferred MathType inline OLE paragraph has no terminal mark.");

            paragraphMark = document.Range(
                paragraphRange.End - 1,
                paragraphRange.End);
            if (!string.Equals(
                    paragraphMark.Text,
                    "\r",
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "A deferred MathType inline OLE lost its generated paragraph boundary.");

            content = document.Content;
            var probeStart = paragraphRange.End;
            var probeEnd = Math.Min(
                content.End,
                probeStart + 64);
            if (probeEnd > probeStart)
                followingProbe = document.Range(probeStart, probeEnd);

            var followingText = followingProbe?.Text ?? string.Empty;
            var actualLeadingLength = 0;
            while (actualLeadingLength < followingText.Length)
            {
                var character = followingText[actualLeadingLength];
                if (character is not (' ' or '\t' or '\u00A0'))
                    break;
                actualLeadingLength++;
            }

            if (actualLeadingLength == 0
                && preservedRightBoundaryWhitespace.Length == 0)
            {
                paragraphMark.Delete();
            }
            else
            {
                mergeRange = document.Range(
                    paragraphMark.Start,
                    paragraphRange.End + actualLeadingLength);
                mergeRange.Text = preservedRightBoundaryWhitespace;
            }
        }
        finally
        {
            Release(content);
            Release(mergeRange);
            Release(followingProbe);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static float? ResolveBatchMathTypeDisplayColumnWidth(
        Document document,
        Range source,
        IDictionary<int, float> sectionCache)
    {
        if (WordFormulaHost.TryGetCellContentWidth(
                source,
                out var cellContentWidth))
            return cellContentWidth;

        int sectionNumber;
        try
        {
            sectionNumber = Convert.ToInt32(
                source.get_Information(
                    WdInformation.wdActiveEndSectionNumber));
        }
        catch
        {
            return null;
        }
        if (sectionNumber <= 0)
            return null;
        if (sectionCache.TryGetValue(
                sectionNumber,
                out var cached))
            return cached;

        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        TextColumns? columns = null;
        try
        {
            sections = document.Sections;
            if (sectionNumber > sections.Count)
                return null;
            section = sections[sectionNumber];
            pageSetup = section.PageSetup;
            columns = pageSetup.TextColumns;
            // A multi-column section can have unequal columns and therefore needs
            // the live paragraph-position resolver. Cache only the unambiguous
            // single-column case; table cells were handled above.
            if (columns is not null
                && columns.Count > 1)
                return null;
            var width = Math.Max(
                72f,
                pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin);
            sectionCache[sectionNumber] = width;
            return width;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(columns);
            Release(pageSetup);
            Release(section);
            Release(sections);
        }
    }

    private static (float Center, float Right)? ResolveBatchDisplayTabStops(
        Document document,
        Range source,
        IDictionary<int, (float Center, float Right)> sectionCache)
    {
        if (WordFormulaHost.TryGetCellContentWidth(source, out var cellContentWidth))
            return (cellContentWidth / 2f, cellContentWidth);

        int sectionNumber;
        try
        {
            sectionNumber = Convert.ToInt32(
                source.get_Information(
                    WdInformation.wdActiveEndSectionNumber));
        }
        catch
        {
            return null;
        }
        if (sectionNumber <= 0)
            return null;
        if (sectionCache.TryGetValue(sectionNumber, out var cached))
            return cached;

        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        try
        {
            sections = document.Sections;
            if (sectionNumber > sections.Count)
                return null;
            section = sections[sectionNumber];
            pageSetup = section.PageSetup;
            var positions = WordEquationNumbering.CalculateEquationTabStops(
                pageSetup.PageWidth,
                pageSetup.LeftMargin,
                pageSetup.RightMargin,
                0,
                0);
            sectionCache[sectionNumber] = positions;
            return positions;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
        }
    }
}

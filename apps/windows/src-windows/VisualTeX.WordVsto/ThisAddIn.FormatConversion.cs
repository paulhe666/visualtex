using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

public sealed partial class ThisAddIn
{
    public void OnConvertVisualTeXToMathTypeSelection(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: false,
            FormulaOleContract.NativeOleMode,
            FormulaOleContract.MathTypeOleMode);

    public void OnConvertVisualTeXToMathTypeDocument(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: true,
            FormulaOleContract.NativeOleMode,
            FormulaOleContract.MathTypeOleMode);

    public void OnConvertMathTypeToVisualTeXSelection(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: false,
            FormulaOleContract.MathTypeOleMode,
            FormulaOleContract.NativeOleMode);

    public void OnConvertMathTypeToVisualTeXDocument(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: true,
            FormulaOleContract.MathTypeOleMode,
            FormulaOleContract.NativeOleMode);

    public void OnConvertOmmlToMathTypeSelection(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: false,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.MathTypeOleMode);

    public void OnConvertOmmlToMathTypeDocument(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.MathTypeOleMode);

    public void OnConvertMathTypeToOmmlSelection(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: false,
            FormulaOleContract.MathTypeOleMode,
            FormulaOleContract.WordOmmlMode);

    public void OnConvertMathTypeToOmmlDocument(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: true,
            FormulaOleContract.MathTypeOleMode,
            FormulaOleContract.WordOmmlMode);

    public void OnConvertVisualTeXToOmmlSelection(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: false,
            FormulaOleContract.NativeOleMode,
            FormulaOleContract.WordOmmlMode);

    public void OnConvertVisualTeXToOmmlDocument(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: true,
            FormulaOleContract.NativeOleMode,
            FormulaOleContract.WordOmmlMode);

    public void OnConvertOmmlToVisualTeXSelection(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: false,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.NativeOleMode);

    public void OnConvertOmmlToVisualTeXDocument(object control) =>
        _ = ConvertFormulaFormatAsync(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.NativeOleMode);

    private async Task ConvertFormulaFormatAsync(
        bool wholeDocument,
        string sourceMode,
        string targetMode)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = _sessionClient;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || client is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-gate-timeout sourceMode={sourceMode} targetMode={targetMode} wholeDocument={wholeDocument}");
            SetStatus(T("VisualTeX 正在执行其他 Word 操作，请稍候再试。",
                "VisualTeX is busy with another Word operation. Try again in a moment."));
            return;
        }
        WordDoubleClickHook.TraceMessage(
            $"format-conversion-start sourceMode={sourceMode} targetMode={targetMode} wholeDocument={wholeDocument}");

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(StringComparer.Ordinal);
        var progress = await dispatcher.InvokeAsync(
            () => BeginOperationProgress(ConversionTitle, service)).ConfigureAwait(false);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        WordFormulaService.WordViewState? originalViewState = null;
        // Word is mutated only inside ApplyFormulaFormatConversionPlan, whose
        // undo record rolls the whole batch back; before it nothing changed.
        var applyStarted = false;
        try
        {
            // Snapshot selection and viewport before even capturing the source
            // formulas. MathType inspection and OMML preparation can themselves
            // touch Word's live Selection, so taking this snapshot inside Apply is
            // already too late for selection-only conversions in long documents.
            if (!wholeDocument)
            {
                originalViewState = await dispatcher.InvokeAsync(
                        () => service.CaptureFormulaFormatConversionViewState())
                    .ConfigureAwait(false);
            }

            var plan = await dispatcher.InvokeAsync(
                    () => service.CaptureFormulaFormatConversionPlan(
                        wholeDocument,
                        sourceMode,
                        targetMode))
                .ConfigureAwait(false);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-plan sourceMode={sourceMode} targetMode={targetMode} targets={plan.Targets.Count}");
            if (plan.Targets.Count == 0)
            {
                var sourceLabel = FormatConversionModeLabel(sourceMode);
                throw new InvalidDataException(
                    wholeDocument
                        ? T($"当前 Word 文档中没有找到可转换的 {sourceLabel} 公式。",
                            $"No convertible {sourceLabel} equations were found in the current Word document.")
                        : T($"所选范围中没有找到可转换的 {sourceLabel} 公式。",
                            $"No convertible {sourceLabel} equations were found in the current selection."));
            }

            var sourceName = FormatConversionModeLabel(sourceMode);
            var targetName = FormatConversionModeLabel(targetMode);

            if (wholeDocument
                && !string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && !string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                progress.Dispose();
                var confirmed = await dispatcher.InvokeAsync(() =>
                    System.Windows.Forms.MessageBox.Show(
                        progress.Owner,
                        T(
                            $"将把全文 {plan.Targets.Count} 个 {sourceName} 公式重新绘制为 {targetName}。\r\n\r\n旧公式宿主和旧编号会直接删除，目标编号将按当前 Word 编号设置重新创建。是否继续？",
                            $"Redraw all {plan.Targets.Count} {sourceName} equations in the document as {targetName}?\r\n\r\nThe old equation hosts and numbering will be removed, and target numbering will be recreated using the current Word numbering settings."),
                        ConversionTitle,
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button2)
                    == System.Windows.Forms.DialogResult.Yes).ConfigureAwait(false);
                if (!confirmed)
                {
                    SetStatus(T("已取消公式格式转换，Word 文档未修改。",
                        "Equation format conversion was cancelled. The Word document was not changed."));
                    return;
                }
                progress = await dispatcher.InvokeAsync(
                    () => BeginOperationProgress(ConversionTitle, service,
                        T("正在准备公式转换…", "Preparing the equation conversion…"))).ConfigureAwait(false);
            }

            // Document-order ordinal of every target ("Equation 3"): plan targets
            // are captured in document order within the converted scope.
            var ordinals = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var index = 0; index < plan.Targets.Count; index++)
                ordinals[plan.Targets[index].Id] = index + 1;
            // Every formula that could not be prepared, with its own reason. The
            // conversion is all-or-nothing, so all of them are reported together
            // before Word is touched.
            var conversionFailures = new List<(int Ordinal, string Latex, string Reason)>();
            // A formula that fails before Word is written keeps its original
            // format; the others are still converted and the skipped ones are
            // listed afterwards. Returns how many targets were left out.
            var skipped = new List<(int Ordinal, string Latex, string Reason)>();
            var totalTargets = plan.Targets.Count;
            int SkipFailedTargets()
            {
                if (conversionFailures.Count == 0) return 0;
                var failed = new HashSet<int>(conversionFailures.Select(failure => failure.Ordinal));
                skipped.AddRange(conversionFailures);
                conversionFailures.Clear();
                var removed = plan.Targets.RemoveAll(target => failed.Contains(ordinals[target.Id]));
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-skipped formulas={removed} remaining={plan.Targets.Count}");
                if (plan.Targets.Count == 0)
                    throw new FormatConversionReport(
                        T("公式格式转换失败：没有可以转换的公式，Word 文档未修改。",
                            "Equation format conversion failed: no equation could be converted. The Word document was not changed."),
                        BuildFormatConversionFailureReport(skipped, totalTargets, null));
                return removed;
            }

            SetStatus(T($"正在准备 {plan.Targets.Count} 个公式的 {targetName} 重绘结果…",
                $"Preparing {targetName} output for {plan.Targets.Count} equations…"));
            var targetAcceptsDirectSourceMathMl = string.Equals(
                    targetMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal)
                || string.Equals(
                    targetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal);
            var allTargetsHaveDirectMathMl = targetAcceptsDirectSourceMathMl
                && plan.Targets.All(target => !string.IsNullOrWhiteSpace(target.SourceMathMl));
            if (!allTargetsHaveDirectMathMl)
            {
                await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
                await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);
            }
            else
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-render-bypass sourceMode={sourceMode} targetMode={targetMode} targets={plan.Targets.Count} reason=source-mathml-ready");
            }

            var targetKeys = new Dictionary<string, string>(StringComparer.Ordinal);
            var pendingKeys = new HashSet<string>(StringComparer.Ordinal);
            var pending = new List<(
                string Key,
                WordBulkRun Run,
                WordFormulaFormatConversionTarget Target,
                CreateVstoSessionRequest Request)>();

            foreach (var target in plan.Targets)
            {
                var run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                };
                var key = string.Join(
                    "\u001F",
                    targetMode,
                    target.DisplayMode,
                    target.FontSizePt.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    target.Latex);
                targetKeys[target.Id] = key;
                if (targetAcceptsDirectSourceMathMl
                    && !string.IsNullOrWhiteSpace(target.SourceMathMl))
                    continue;
                if (!pendingKeys.Add(key)) continue;

                var request = CreateBulkFormulaConversionRequest(
                    run, targetMode, plan.DocumentId, target.FontSizePt);
                pending.Add((key, run, target, request));
            }

            if (pending.Count > 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-render-start sourceMode={sourceMode} targetMode={targetMode} formulas={pending.Count}");
                // Bounded submissions expose actual completed counts and avoid
                // one opaque multi-minute request for an entire document. The
                // next chunk renders in the companion while this chunk's
                // previews are built concurrently in the add-in.
                Task<RenderBatchResult>? SubmitChunk(int offset) =>
                    offset >= pending.Count
                        ? null
                        : client.RenderBatchAsync(
                            pending.Skip(offset).Take(20).Select(item => new RenderBatchItemRequest
                            { Id = item.Run.Id, Request = item.Request }).ToList(),
                            TimeSpan.FromMinutes(3), lifetime.Token);
                // Identical formulas share one render; a failure names each of them.
                var targetsByKey = plan.Targets
                    .Where(target => !(targetAcceptsDirectSourceMathMl
                        && !string.IsNullOrWhiteSpace(target.SourceMathMl)))
                    .GroupBy(target => targetKeys[target.Id], StringComparer.Ordinal)
                    .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
                var nextChunk = SubmitChunk(0);
                for (var offset = 0; offset < pending.Count; offset += 20)
                {
                    progress.Report(T("正在渲染公式…", "Rendering equations…"), offset, pending.Count);
                    var chunk = pending.Skip(offset).Take(20).ToArray();
                    RenderBatchResult batch;
                    try
                    {
                        batch = await nextChunk!.ConfigureAwait(false);
                    }
                    catch (Exception error) when (!IsUserCancellation(error, lifetime.Token))
                    {
                        // The whole batch failed (VisualTeX unreachable, timed
                        // out, converter not started...): every later batch
                        // would wait for the same cause, so stop here and report
                        // it once, with the formulas it covered and every
                        // formula that already failed on its own.
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-render-batch-failed offset={offset} items={chunk.Length} error={error}");
                        var covered = chunk
                            .SelectMany(item => targetsByKey[item.Key])
                            .Select(target => ordinals[target.Id])
                            .OrderBy(ordinal => ordinal)
                            .ToArray();
                        throw new FormatConversionReport(
                            T($"公式格式转换失败：VisualTeX 无法渲染第 {covered.First()}–{covered.Last()} 个公式所在的一批公式，Word 文档未修改。",
                                $"Equation format conversion failed: VisualTeX could not render the batch containing equations {covered.First()}–{covered.Last()}. The Word document was not changed."),
                            BuildFormatConversionFailureReport(
                                conversionFailures,
                                plan.Targets.Count,
                                T($"VisualTeX 无法渲染第 {covered.First()}–{covered.Last()} 个公式所在的一批公式（{covered.Length} 个公式）：",
                                    $"VisualTeX could not render the batch containing equations {covered.First()}–{covered.Last()} ({covered.Length} equations): ")
                                    + OfficeErrorText.Describe(error)));
                    }
                    nextChunk = SubmitChunk(offset + 20);
                    var materialized = MaterializeBulkFormulaTemplates(
                        client,
                        chunk.Select((item, index) => (item.Run, batch.Items[index].Session)).ToList(),
                        targetMode,
                        materializePreview: targetMode == FormulaOleContract.NativeOleMode);
                    for (var index = 0; index < chunk.Length; index++)
                    {
                        var item = chunk[index];
                        var completed = batch.Items[index].Session;
                        if (materialized[index].Error is { } materializeError)
                        {
                            var reason = materializeError is WordFormulaRenderException renderError
                                ? renderError.Reason
                                : OfficeErrorText.Describe(materializeError);
                            foreach (var target in targetsByKey[item.Key])
                                conversionFailures.Add((ordinals[target.Id], target.Latex, reason));
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-render-failed sourceMode={sourceMode} targetMode={targetMode} itemId={item.Run.Id} status={completed.Status} error={materializeError}");
                            continue;
                        }
                        rendered[item.Key] = materialized[index].Template!;
                        WordDoubleClickHook.TraceMessage(
                            $"format-conversion-rendered sourceMode={sourceMode} targetMode={targetMode} itemId={item.Run.Id} status={completed.Status}");
                    }
                    progress.Report(T("正在渲染公式…", "Rendering equations…"), offset + chunk.Length, pending.Count);
                }
            }

            SkipFailedTargets();

            foreach (var target in plan.Targets)
            {
                var key = targetKeys[target.Id];
                var run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                };
                if (targetAcceptsDirectSourceMathMl
                    && !string.IsNullOrWhiteSpace(target.SourceMathMl))
                {
                    var directSession = CloneBulkFormulaSession(
                        new OfficeSessionDocument(),
                        run,
                        plan.DocumentId,
                        target.FontSizePt,
                        targetMode);
                    directSession.Numbered = target.Numbered;
                    directSession.MathTypeNumberPosition = target.MathTypeNumberPosition;
                    directSession.OriginalMetadata = target.Metadata;
                    prepared[target.Id] = new PreparedWordBulkFormula
                    {
                        Run = run,
                        Session = directSession,
                        MathMl = target.SourceMathMl,
                    };
                    continue;
                }

                if (!rendered.TryGetValue(key, out var template))
                    throw new InvalidDataException(T(
                        $"第 {ordinals[target.Id]} 个公式（{WordFormulaRenderException.Excerpt(target.Latex, 120)}）缺少渲染结果。",
                        $"Equation {ordinals[target.Id]} ({WordFormulaRenderException.Excerpt(target.Latex, 120)}) has no render result."));
                var session = CloneBulkFormulaSession(
                    template.Session,
                    run,
                    plan.DocumentId,
                    target.FontSizePt,
                    targetMode);
                session.Numbered = target.Numbered;
                session.MathTypeNumberPosition = target.MathTypeNumberPosition;
                session.OriginalMetadata = target.Metadata;
                prepared[target.Id] = new PreparedWordBulkFormula
                {
                    Run = run,
                    Session = session,
                    MathMl = string.IsNullOrWhiteSpace(target.SourceMathMl)
                        ? template.MathMl
                        : target.SourceMathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                };
            }

            if (string.Equals(
                    targetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                foreach (var target in plan.Targets)
                    prepared[target.Id].MathTypeNativePreviewAttempted = true;

                var nativePreviewAvailable = MathTypeNativePreviewRenderer.IsAvailable;
                if (!nativePreviewAvailable)
                {
                    throw new InvalidOperationException(MathTypeNativePreviewRenderer.UnavailableMessage);
                }
                else
                {
                    SetStatus(T($"正在批量生成 {plan.Targets.Count} 个 MathType 原生预览…",
                        $"Generating {plan.Targets.Count} native MathType previews…"));
                    var nativePreviewInputs =
                        new Dictionary<string, byte[]>(StringComparer.Ordinal);
                    var encodingAttempts = 0;
                    foreach (var target in plan.Targets)
                    {
                        progress.Report(T("正在编码 MathType 公式…", "Encoding MathType equations…"), encodingAttempts++, plan.Targets.Count);
                        var formula = prepared[target.Id];
                        try
                        {
                            var mathMl = formula.MathMl
                                ?? throw new InvalidDataException(T(
                                    "该公式没有可写入 MathType 的 MathML。",
                                    "This equation has no MathML to write as MathType."));
                            var inline = string.Equals(
                                target.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase);
                            var standalone = WordFormulaService.PrepareStandaloneMathTypeOleData(
                                mathMl, inline, formula.Session.FontSizePt, target.Latex);
                            formula.MathTypeStandaloneGenerated = standalone.Generated;
                            formula.MathTypeStandaloneCompoundFile = standalone.CompoundFile;
                            formula.MathTypeStandaloneSemanticSignature = standalone.SemanticSignature;
                            nativePreviewInputs[target.Id] = standalone.Generated.Mtef;
                        }
                        catch (Exception error) when (!IsUserCancellation(error, lifetime.Token))
                        {
                            // Encode every formula first: all failures are listed.
                            WordDoubleClickHook.TraceMessage(
                                $"format-conversion-mathtype-encode-failed ordinal={ordinals[target.Id]} error={error}");
                            conversionFailures.Add((
                                ordinals[target.Id],
                                target.Latex,
                                T("无法编码为 MathType 公式：", "Could not encode it as a MathType equation: ")
                                    + OfficeErrorText.Describe(error)));
                        }
                    }
                    SkipFailedTargets();

                    var nativePreviewRoot = prepared.Values
                        .Select(formula => string.IsNullOrWhiteSpace(formula.EmfPath)
                            ? null
                            : Path.GetDirectoryName(formula.EmfPath))
                        .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path))
                        ?? Path.GetTempPath();
                    var nativePreviewWatch = System.Diagnostics.Stopwatch.StartNew();
                    progress.Report(T("正在生成 MathType 原生预览…", "Generating native MathType previews…"), 0, plan.Targets.Count);
                    var renderedAllNativePreviews =
                        MathTypeNativePreviewRenderer.TryRenderBatchWithRetry(
                            nativePreviewInputs,
                            nativePreviewRoot,
                            out var nativePreviews);
                    progress.Report(T("MathType 原生预览已生成", "Native MathType previews generated"), nativePreviews.Count, plan.Targets.Count);
                    var missingNativePreviews = plan.Targets
                        .Where(target => !nativePreviews.ContainsKey(target.Id))
                        .ToArray();
                    if (missingNativePreviews.Length > 0
                        && (MathTypeNativePreviewRenderer.LastFailureWasFonts || nativePreviews.Count == 0))
                    {
                        // The renderer itself is unusable (no MathType fonts, nothing
                        // rendered): no formula can be converted.
                        foreach (var preview in nativePreviews.Values)
                            preview.Dispose();
                        nativePreviewWatch.Stop();
                        string? OrdinalOf(string id) =>
                            ordinals.TryGetValue(id, out var ordinal)
                                ? T($"第 {ordinal} 个公式", $"Equation {ordinal}")
                                : null;
                        throw new InvalidOperationException(
                            MathTypeNativePreviewRenderer.DescribeFailures(OrdinalOf));
                    }
                    // A formula the renderer could not draw keeps its source format.
                    var previewFailures = MathTypeNativePreviewRenderer.LastFailures;
                    foreach (var target in missingNativePreviews)
                        conversionFailures.Add((
                            ordinals[target.Id],
                            target.Latex,
                            previewFailures.TryGetValue(target.Id, out var previewReason)
                                ? OfficePluginLanguage.SafeErrorMessage(previewReason)
                                : MathTypeNativePreviewRenderer.NoResultMessage));
                    SkipFailedTargets();
                    foreach (var target in plan.Targets)
                        prepared[target.Id].MathTypeNativePreview =
                            nativePreviews[target.Id];
                    nativePreviewWatch.Stop();
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-native-preview-batch-complete formulas={nativePreviews.Count} all={renderedAllNativePreviews} elapsedMs={nativePreviewWatch.ElapsedMilliseconds}");
                }
            }

            SetStatus(T("公式已全部渲染，正在原位写入 Word…",
                "All equations are rendered. Writing them into Word in place…"));
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-word-apply-start sourceMode={sourceMode} targetMode={targetMode} targets={plan.Targets.Count}");
            applyStarted = true;
            WordFormulaFormatConversionResult result;
            while (true)
            {
                BeginFormulaFormatMutation();
                try
                {
                    result = await dispatcher.InvokeAsync(
                            () => service.ApplyFormulaFormatConversionPlan(plan, prepared))
                        .ConfigureAwait(false);
                    break;
                }
                catch (Exception error) when (FindPreflightFailure(error) is not null)
                {
                    // Nothing was written: leave these formulas in their source
                    // format and convert the others in a new pass.
                    foreach (var failure in FindPreflightFailure(error)!.Failures)
                    {
                        var failedTarget = plan.Targets.FirstOrDefault(item => item.Id == failure.TargetId);
                        if (failedTarget is not null)
                            conversionFailures.Add((ordinals[failedTarget.Id], failedTarget.Latex, failure.Reason));
                    }
                    if (SkipFailedTargets() == 0) throw;
                }
                finally
                {
                    EndFormulaFormatMutation();
                }
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-word-apply-finished sourceMode={sourceMode} targetMode={targetMode} converted={result.FormulaCount} failed={result.FailedFormulaCount}");

            // The conversion applies all-or-nothing: Apply either converts
            // every target or throws after rolling the batch back, so there is
            // no partial-success result to report here.
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-complete source={sourceName} target={targetName} converted={result.FormulaCount} skipped={skipped.Count}");
            if (skipped.Count == 0)
            {
                SetStatus(T(
                    $"公式格式转换完成：{result.FormulaCount} 个公式已重新绘制为 {targetName}。",
                    $"Equation format conversion finished: {result.FormulaCount} equation(s) were redrawn as {targetName}."));
            }
            else
            {
                SetStatus(T(
                    $"公式格式转换完成：{result.FormulaCount} 个公式已重新绘制为 {targetName}，{skipped.Count} 个公式无法转换，保持原格式。",
                    $"Equation format conversion finished: {result.FormulaCount} equation(s) were redrawn as {targetName}; {skipped.Count} could not be converted and keep their original format."));
                progress.Dispose();
                if (!IsConversionAcceptanceRun)
                {
                    var report = BuildFormatConversionFailureReport(skipped, totalTargets, null, documentChanged: true);
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            report,
                            ConversionTitle,
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return true;
                    }).ConfigureAwait(false);
                }
            }
        }
        catch (Exception error) when (IsUserCancellation(error, lifetime.Token))
        {
            // Only the operation's own token is a cancellation; a timed-out or
            // released render batch is an error reported below.
            WordDoubleClickHook.TraceMessage($"format-conversion-cancelled error={error}");
            SetStatus(T("公式格式转换已取消，Word 文档未修改。",
                "Equation format conversion was cancelled. The Word document was not changed."));
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-failed sourceMode={sourceMode} targetMode={targetMode} applyStarted={applyStarted} error={error}");
            string statusMessage;
            string failureMessage;
            if (error is FormatConversionReport report)
            {
                statusMessage = report.Status;
                failureMessage = report.Message;
            }
            else
            {
                failureMessage = OfficePluginLanguage.FormatOperationError(
                    error,
                    "公式格式转换失败",
                    "Equation format conversion failed");
                // Before Apply nothing was written; Apply rolls its one undo
                // record back and reports a failed rollback as an aggregate.
                if (!applyStarted || error is not AggregateException)
                    failureMessage += T("\r\n\r\nWord 文档未修改。", "\r\n\r\nThe Word document was not changed.");
                statusMessage = failureMessage;
            }
            SetStatus(statusMessage);
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal)
                && dispatcher is not null)
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            failureMessage,
                            ConversionTitle,
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch (Exception displayError)
                {
                    WordDoubleClickHook.TraceMessage($"format-conversion-error-display-failed error={displayError}");
                }
            }
        }
        finally
        {
            if (originalViewState is not null)
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        service.RestoreFormulaFormatConversionViewState(originalViewState);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch
                {
                    // Selection/view restoration is best-effort and must never
                    // hide the conversion result or prevent gate release.
                }
            }
            foreach (var template in rendered.Values)
            {
                TryDeleteFile(template.EmfPath);
                TryDeleteFile(template.SvgPath);
                TryDeleteFile(template.PngPath);
            }
            foreach (var preview in prepared.Values
                         .Select(formula => formula.MathTypeNativePreview)
                         .Where(preview => preview is not null)
                         .Distinct())
                preview!.Dispose();
            progress.Dispose();
            _operationGate.Release();
        }
    }

    private static async Task<Exception> RecordSessionFailureAsync(
        VisualTeXSessionClient client,
        IReadOnlyCollection<string> sessionIds,
        Exception error)
    {
        if (sessionIds.Count == 0) return error;
        try
        {
            await FinalizeOfficeSessionsAsync(client, sessionIds, completed: false, error.ToString()).ConfigureAwait(false);
            return error;
        }
        catch (Exception recordingError)
        {
            return new AggregateException(
                T("Word 操作失败，且未能向 VisualTeX 记录失败状态。",
                    "The Word operation failed, and its failure could not be recorded in VisualTeX."),
                error,
                recordingError);
        }
    }

    private static async Task CleanupConverterSessionsBestEffortAsync(
        VisualTeXSessionClient client,
        IReadOnlyCollection<string> sessionIds,
        CancellationToken lifetimeToken)
    {
        if (sessionIds.Count == 0) return;
        try
        {
            using var cleanupCts = CancellationTokenSource.CreateLinkedTokenSource(lifetimeToken);
            cleanupCts.CancelAfter(TimeSpan.FromSeconds(5));
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-session-cleanup-start sessions={sessionIds.Count}");
            await client.DeleteSessionsBatchAsync(
                    sessionIds,
                    cleanupCts.Token)
                .ConfigureAwait(false);
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-session-cleanup-complete sessions={sessionIds.Count}");
        }
        catch (Exception cleanupError)
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-session-cleanup-skipped sessions={sessionIds.Count} error={cleanupError.GetType().Name}");
        }
    }

    private static string ConversionTitle =>
        T("VisualTeX 公式格式转换", "VisualTeX Equation Format Conversion");

    /// <summary>
    /// The user's own cancellation: an OperationCanceledException while the
    /// operation's token is cancelled. VisualTeX request timeouts are
    /// TimeoutException and a batch VisualTeX released is an error, so neither
    /// is ever shown as "cancelled".
    /// </summary>
    private static bool IsUserCancellation(Exception error, CancellationToken operationToken) =>
        operationToken.IsCancellationRequested
        && (error is OperationCanceledException
            || (error is AggregateException aggregate
                && aggregate.Flatten().InnerExceptions.Any(inner => inner is OperationCanceledException)));

    /// <summary>
    /// A conversion stopped before Word was changed, with its complete report
    /// (every failed formula and its own reason) already worded for the user.
    /// </summary>
    private sealed class FormatConversionReport : Exception
    {
        internal FormatConversionReport(string status, string message)
            : base(message)
        {
            Status = status;
        }

        internal string Status { get; }
    }

    private const int MaximumListedConversionFailures = 15;

    private static bool IsConversionAcceptanceRun =>
        string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"), "1", StringComparison.Ordinal)
        || string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE"), "1", StringComparison.Ordinal);

    // The sources of a conversion's write step are prepared in one call on the
    // Word thread; its failure arrives through the dispatcher, possibly wrapped.
    private static WordFormulaFormatConversionPreflightException? FindPreflightFailure(Exception? error)
    {
        for (; error is not null; error = error.InnerException)
            if (error is WordFormulaFormatConversionPreflightException preflight)
                return preflight;
        return null;
    }

    private static string BuildFormatConversionFailureReport(
        IReadOnlyList<(int Ordinal, string Latex, string Reason)> failures,
        int totalFormulas,
        string? batchFailure,
        bool documentChanged = false)
    {
        var builder = new System.Text.StringBuilder();
        if (failures.Count > 0)
        {
            builder.Append(T(
                $"共 {totalFormulas} 个公式，其中 {failures.Count} 个无法转换：",
                $"{failures.Count} of {totalFormulas} equation(s) could not be converted:"));
            builder.Append("\r\n");
            foreach (var failure in failures
                         .OrderBy(item => item.Ordinal)
                         .Take(MaximumListedConversionFailures))
            {
                builder.Append("\r\n");
                builder.Append(T(
                    $"第 {failure.Ordinal} 个公式：{WordFormulaRenderException.Excerpt(failure.Latex, 120)}\r\n   原因：{failure.Reason}",
                    $"Equation {failure.Ordinal}: {WordFormulaRenderException.Excerpt(failure.Latex, 120)}\r\n   Reason: {failure.Reason}"));
            }
            if (failures.Count > MaximumListedConversionFailures)
            {
                builder.Append("\r\n\r\n");
                builder.Append(T(
                    $"……另有 {failures.Count - MaximumListedConversionFailures} 个失败的公式未列出。",
                    $"…and {failures.Count - MaximumListedConversionFailures} more failed equation(s) not listed."));
            }
            builder.Append("\r\n\r\n");
        }
        if (!string.IsNullOrWhiteSpace(batchFailure))
        {
            builder.Append(batchFailure);
            builder.Append("\r\n\r\n");
        }
        builder.Append(documentChanged
            ? T("其余公式已转换；上述公式保持原格式，可修改后再单独转换。",
                "The other equations were converted; the equations above keep their original format and can be converted separately after correcting them.")
            : T("Word 文档未修改。", "The Word document was not changed."));
        return builder.ToString();
    }

    private static string FormatConversionModeLabel(string mode)
    {
        if (string.Equals(
                mode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal))
            return "MathType";
        if (string.Equals(
                mode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal))
            return "OMML";
        return "VisualTeX";
    }
}

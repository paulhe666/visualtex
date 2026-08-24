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
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }
        WordDoubleClickHook.TraceMessage(
            $"format-conversion-start sourceMode={sourceMode} targetMode={targetMode} wholeDocument={wholeDocument}");

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        var converterSessionIds = new List<string>();
        WordFormulaService.WordViewState? originalViewState = null;
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
                        ? $"当前 Word 文档中没有找到可转换的 {sourceLabel} 公式。"
                        : $"所选范围中没有找到可转换的 {sourceLabel} 公式。");
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
                var confirmed = await dispatcher.InvokeAsync(() =>
                    System.Windows.Forms.MessageBox.Show(
                        $"将把全文 {plan.Targets.Count} 个 {sourceName} 公式重新绘制为 {targetName}。\r\n\r\n"
                        + "旧公式宿主和旧编号会直接删除，目标编号将按当前 Word 编号设置重新创建。是否继续？",
                        "VisualTeX 公式格式转换",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button2)
                    == System.Windows.Forms.DialogResult.Yes).ConfigureAwait(false);
                if (!confirmed)
                {
                    SetStatus("已取消公式格式转换，Word 文档未修改。");
                    return;
                }
            }

            SetStatus($"正在准备 {plan.Targets.Count} 个公式的 {targetName} 重绘结果…");
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
                OfficeSessionDocument Session)>();

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

                var conversionSession = await CreateBulkFormulaConversionSessionAsync(
                        client,
                        run,
                        targetMode,
                        plan.DocumentId,
                        target.FontSizePt,
                        lifetime.Token)
                    .ConfigureAwait(false);
                pending.Add((key, run, target, conversionSession));
                converterSessionIds.Add(conversionSession.Id);
            }

            if (pending.Count > 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-render-start sourceMode={sourceMode} targetMode={targetMode} sessions={pending.Count}");
                await client.OpenConverterBatchAsync(
                        pending.Select(item => item.Session.Id).ToList(),
                        lifetime.Token)
                    .ConfigureAwait(false);
                foreach (var item in pending)
                {
                    var completed = await client.WaitForCommitAsync(
                            item.Session.Id,
                            TimeSpan.FromMinutes(3),
                            lifetime.Token)
                        .ConfigureAwait(false);
                    rendered[item.Key] = MaterializeBulkFormulaTemplate(
                        client,
                        item.Run,
                        targetMode,
                        completed);
                    WordDoubleClickHook.TraceMessage(
                        $"format-conversion-rendered sourceMode={sourceMode} targetMode={targetMode} sessionId={item.Session.Id} status={completed.Status}");
                }
            }

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
                    throw new InvalidDataException(
                        $"缺少公式“{target.Latex}”的重绘结果。");
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
                SetStatus($"正在批量生成 {plan.Targets.Count} 个 MathType 原生预览…");
                var nativePreviewInputs =
                    new Dictionary<string, byte[]>(StringComparer.Ordinal);
                foreach (var target in plan.Targets)
                {
                    var formula = prepared[target.Id];
                    var mathMl = formula.MathMl
                        ?? throw new InvalidDataException(
                            $"缺少公式“{target.Latex}”的 MathType MathML。");
                    var inline = string.Equals(
                        target.DisplayMode,
                        "inline",
                        StringComparison.OrdinalIgnoreCase);
                    var generated = MathTypeMtefCodec.CreateEquationNative(
                        mathMl,
                        inline);
                    nativePreviewInputs[target.Id] = generated.Mtef;
                    // Once the batch is attempted, InsertMathTypeOle must never
                    // start one MathPage sidecar per formula or silently switch
                    // the visible result back to frontend/MathJax geometry.
                    formula.MathTypeNativePreviewAttempted = true;
                }

                var nativePreviewRoot = prepared.Values
                    .Select(formula => string.IsNullOrWhiteSpace(formula.EmfPath)
                        ? null
                        : Path.GetDirectoryName(formula.EmfPath))
                    .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path))
                    ?? Path.GetTempPath();
                var nativePreviewWatch = System.Diagnostics.Stopwatch.StartNew();
                var renderedAllNativePreviews =
                    MathTypeNativePreviewRenderer.TryRenderBatch(
                        nativePreviewInputs,
                        nativePreviewRoot,
                        out var nativePreviews);
                var missingNativePreviewIds = plan.Targets
                    .Where(target => !nativePreviews.ContainsKey(target.Id))
                    .Select(target => target.Id)
                    .ToArray();
                if (!renderedAllNativePreviews
                    || missingNativePreviewIds.Length > 0)
                {
                    foreach (var preview in nativePreviews.Values)
                        preview.Dispose();
                    throw new InvalidOperationException(
                        $"MathType 原生预览批量渲染失败（成功 {nativePreviews.Count}/{plan.Targets.Count}）。"
                        + "为避免回退到 VisualTeX 前端几何，Word 文档尚未开始转换。");
                }

                foreach (var target in plan.Targets)
                    prepared[target.Id].MathTypeNativePreview =
                        nativePreviews[target.Id];
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-native-preview-batch-complete formulas={nativePreviews.Count} elapsedMs={nativePreviewWatch.ElapsedMilliseconds}");
            }

            SetStatus("公式已全部渲染，正在用正常新建公式路径原位重绘…");
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-word-apply-start sourceMode={sourceMode} targetMode={targetMode} targets={plan.Targets.Count}");
            BeginFormulaFormatMutation();
            WordFormulaFormatConversionResult result;
            try
            {
                result = await dispatcher.InvokeAsync(
                        () => service.ApplyFormulaFormatConversionPlan(plan, prepared))
                    .ConfigureAwait(false);
            }
            finally
            {
                EndFormulaFormatMutation();
            }
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-word-apply-finished sourceMode={sourceMode} targetMode={targetMode} converted={result.FormulaCount} failed={result.FailedFormulaCount}");

            // Converter sessions are disposable render workspaces. Word has already
            // committed the document result, so SessionStore cleanup must not keep
            // the Ribbon operation/gate blocked if the Companion mutex is busy.
            _ = CleanupConverterSessionsBestEffortAsync(
                client,
                converterSessionIds.ToArray(),
                lifetime.Token);

            if (result.FailedFormulaCount == 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-complete source={sourceName} target={targetName} converted={result.FormulaCount} failed=0");
                SetStatus(
                    $"公式格式转换完成：{result.FormulaCount} 个公式已重新绘制为 {targetName}。");
            }
            else
            {
                var detail = result.Failures.FirstOrDefault() ?? "未知 Word 写入错误。";
                WordDoubleClickHook.TraceMessage(
                    $"format-conversion-stopped source={sourceName} target={targetName} converted={result.FormulaCount} failed={result.FailedFormulaCount} detail={detail}");
                SetStatus(
                    $"已转换 {result.FormulaCount} 个公式，随后停止：{detail}");
                if (!string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            $"已成功转换 {result.FormulaCount} 个公式，随后停止。\r\n\r\n{detail}",
                            "VisualTeX 公式格式转换",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                        return true;
                    }).ConfigureAwait(false);
                }
            }
        }
        catch (OperationCanceledException)
        {
            SetStatus("公式格式转换已取消。");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"format-conversion-failed sourceMode={sourceMode} targetMode={targetMode} error={error}");
            SetStatus($"公式格式转换失败：{error.Message}");
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
                            error.Message,
                            "VisualTeX 公式格式转换",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch { }
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
            _operationGate.Release();
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

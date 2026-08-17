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
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        var converterSessionIds = new List<string>();
        try
        {
            var plan = await dispatcher.InvokeAsync(
                    () => service.CaptureFormulaFormatConversionPlan(
                        wholeDocument,
                        sourceMode,
                        targetMode))
                .ConfigureAwait(false);
            if (plan.Targets.Count == 0)
            {
                var sourceLabel = string.Equals(
                    sourceMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                    ? "MathType"
                    : "VisualTeX";
                throw new InvalidDataException(
                    wholeDocument
                        ? $"当前 Word 文档中没有找到可转换的 {sourceLabel} 公式。"
                        : $"所选范围中没有找到可转换的 {sourceLabel} 公式。");
            }

            var sourceName = string.Equals(
                sourceMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
                ? "MathType"
                : "VisualTeX";
            var targetName = string.Equals(
                targetMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
                ? "MathType"
                : "VisualTeX";

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
            await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
            await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);

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
                }
            }

            foreach (var target in plan.Targets)
            {
                var key = targetKeys[target.Id];
                if (!rendered.TryGetValue(key, out var template))
                    throw new InvalidDataException(
                        $"缺少公式“{target.Latex}”的重绘结果。");
                var run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                };
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
                    MathMl = template.MathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                };
            }

            SetStatus("公式已全部渲染，正在用正常新建公式路径原位重绘…");
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

            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.CompleteAsync(sessionId, lifetime.Token).ConfigureAwait(false);
                }
                catch { }
            }

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
            foreach (var template in rendered.Values)
            {
                TryDeleteFile(template.EmfPath);
                TryDeleteFile(template.SvgPath);
                TryDeleteFile(template.PngPath);
            }
            _operationGate.Release();
        }
    }
}

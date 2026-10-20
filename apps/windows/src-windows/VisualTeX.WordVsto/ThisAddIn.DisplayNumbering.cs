using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Task = System.Threading.Tasks.Task;

namespace VisualTeX.WordVsto;

public sealed partial class ThisAddIn
{
    public void OnNumberDisplayFormulasSelection(object control) =>
        _ = NumberDisplayFormulasAsync(wholeDocument: false);

    public void OnNumberDisplayFormulasDocument(object control) =>
        _ = NumberDisplayFormulasAsync(wholeDocument: true);

    private async Task NumberDisplayFormulasAsync(bool wholeDocument)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            SetStatus(T(
                "VisualTeX 正在执行其他 Word 操作，请稍候再试。",
                "VisualTeX is busy with another Word operation. Please try again shortly."));
            return;
        }

        var title = T("VisualTeX 批量编号", "VisualTeX Equation Numbering");
        var scopeText = wholeDocument ? T("全文", "the document") : T("所选内容", "the selection");
        WordOperationProgress? progress = null;
        try
        {
            progress = await dispatcher.InvokeAsync(() => BeginOperationProgress(
                    title, service, T("正在查找行间公式…", "Finding display equations…")))
                .ConfigureAwait(false);
            var plan = await dispatcher.InvokeAsync(
                    () => TraceDisplayNumberingStages(
                        "plan", () => service.PlanDisplayFormulaNumberingCore(wholeDocument)))
                .ConfigureAwait(false);
            // The confirmation must not compete with the progress window.
            progress.Dispose();
            progress = null;
            WriteRedrawAcceptanceLog(
                $"display-numbering-plan scope={(wholeDocument ? "document" : "selection")} "
                + $"omml={plan.Omml.Count} visualtex={plan.VisualTeX.Count} mathtype={plan.MathType.Count} "
                + $"numbered={plan.AlreadyNumbered} skipped={plan.Skipped}");

            var format = EquationNumberFormat.Resolve(service.GetEquationNumberFormatId());
            var details = DescribeDisplayNumberingPlan(plan);
            var confirmed = await dispatcher.InvokeAsync(() =>
            {
                if (plan.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        T(
                            $"{scopeText}中没有需要编号的行间公式。" + details,
                            $"There is no unnumbered display equation in {scopeText}." + details),
                        title,
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                    return true;
                return System.Windows.Forms.MessageBox.Show(
                        T(
                            $"将为{scopeText}中 {plan.Count} 个未编号的行间公式添加编号"
                            + $"（VisualTeX {plan.VisualTeX.Count} 个，OMML {plan.Omml.Count} 个，MathType {plan.MathType.Count} 个），"
                            + $"编号格式：{format.DisplayName}。\r\n\r\n"
                            + "已有编号保持不变，其后的编号会自动顺延，交叉引用同步更新。" + details
                            + "\r\n\r\n可通过一次 Ctrl+Z 整体撤销。是否继续？",
                            $"Add equation numbers to {plan.Count} unnumbered display equations in {scopeText} "
                            + $"(VisualTeX {plan.VisualTeX.Count}, OMML {plan.Omml.Count}, MathType {plan.MathType.Count}), "
                            + $"number format: {format.DisplayName}.\r\n\r\n"
                            + "Existing numbers stay as they are; the numbers after them follow on and cross-references are updated." + details
                            + "\r\n\r\nOne Ctrl+Z undoes it. Continue?"),
                        title,
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button1)
                    == System.Windows.Forms.DialogResult.Yes;
            }).ConfigureAwait(false);
            if (!confirmed)
            {
                SetStatus(plan.Count == 0
                    ? T("没有需要编号的行间公式。", "No display equation needed a number.")
                    : T("已取消批量编号，Word 文档未修改。", "Batch numbering was cancelled; the Word document was not changed."));
                WriteRedrawAcceptanceLog("display-numbering-cancelled");
                return;
            }

            SetStatus(T($"正在为 {plan.Count} 个行间公式添加编号…", $"Numbering {plan.Count} display equations…"));
            progress = await dispatcher.InvokeAsync(() => BeginOperationProgress(
                    title, service, T("正在添加公式编号…", "Adding equation numbers…")))
                .ConfigureAwait(false);
            BeginFormulaFormatMutation();
            WordDisplayNumberingResult result;
            try
            {
                result = await dispatcher.InvokeAsync(
                        () => TraceDisplayNumberingStages(
                            "apply", () => service.NumberDisplayFormulasCore(plan)))
                    .ConfigureAwait(false);
            }
            finally { EndFormulaFormatMutation(); }
            progress.Dispose();
            progress = null;

            WriteRedrawAcceptanceLog(
                $"display-numbering-complete omml={result.Omml} visualtex={result.VisualTeX} "
                + $"mathtype={result.MathType} seconds={result.ElapsedSeconds:0.0}");
            SetStatus(T(
                $"已为 {result.Count} 个行间公式添加编号（{result.ElapsedSeconds:0.0} 秒）。",
                $"Numbered {result.Count} display equations ({result.ElapsedSeconds:0.0} s)."));
        }
        catch (OperationCanceledException)
        {
            WriteRedrawAcceptanceLog("display-numbering-cancelled");
            SetStatus(T("VisualTeX 批量编号已取消。", "VisualTeX batch numbering was cancelled."));
        }
        catch (Exception error)
        {
            progress?.Dispose();
            progress = null;
            WriteRedrawAcceptanceLog("display-numbering-failed " + error);
            SetStatus(T(
                $"批量编号失败，文档已恢复原状：{error.Message}",
                $"Batch numbering failed and the document was restored: {error.Message}"));
            await ShowWordOperationErrorAsync(
                    dispatcher,
                    T("批量编号", "Batch Numbering"),
                    error)
                .ConfigureAwait(false);
        }
        finally
        {
            progress?.Dispose();
            _operationGate.Release();
        }
    }

    // With VISUALTEX_WORD_HOOK_TRACE_PATH set, report where the time went, per
    // stage, so the performance contract can be checked on a real document.
    private static T TraceDisplayNumberingStages<T>(string phase, Func<T> run)
    {
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH")))
            return run();
        var watch = System.Diagnostics.Stopwatch.StartNew();
        using var metrics = new WordOperationMetrics();
        try { return run(); }
        finally
        {
            WordDoubleClickHook.TraceMessage(
                $"[perf] DisplayNumbering.{phase}: {watch.Elapsed.TotalMilliseconds:0.0}ms");
            foreach (var entry in metrics.Entries.OrderByDescending(item => item.Value.Milliseconds))
                WordDoubleClickHook.TraceMessage(
                    $"[perf] DisplayNumbering.{phase}-stage {entry.Key}: calls={entry.Value.Calls} inclusive={entry.Value.Milliseconds:0.0}ms");
        }
    }

    private static string DescribeDisplayNumberingPlan(WordDisplayNumberingPlan plan)
    {
        var lines = new List<string>();
        if (plan.AlreadyNumbered > 0)
            lines.Add(T($"已有编号的行间公式：{plan.AlreadyNumbered} 个（不变）",
                $"Display equations that already have a number: {plan.AlreadyNumbered} (unchanged)"));
        if (plan.SkippedInTable > 0)
            lines.Add(T($"位于表格中，跳过：{plan.SkippedInTable} 个",
                $"Inside a table, skipped: {plan.SkippedInTable}"));
        if (plan.SkippedOtherContent > 0)
            lines.Add(T($"公式所在段落还有其他内容，跳过：{plan.SkippedOtherContent} 个",
                $"The equation's paragraph holds other content, skipped: {plan.SkippedOtherContent}"));
        if (plan.SkippedTypedNumber > 0)
            lines.Add(T($"已手动输入编号文字，跳过：{plan.SkippedTypedNumber} 个",
                $"A number was typed as text, skipped: {plan.SkippedTypedNumber}"));
        if (plan.SkippedOldNumbering > 0)
            lines.Add(T($"使用旧版编号结构，跳过：{plan.SkippedOldNumbering} 个（可先点“更新公式格式”）",
                $"Old numbering structure, skipped: {plan.SkippedOldNumbering} (use Update Equation Format first)"));
        return lines.Count == 0 ? string.Empty : "\r\n\r\n" + string.Join("\r\n", lines);
    }
}

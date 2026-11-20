using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Office = Microsoft.Office.Core;

namespace VisualTeX.WordVsto;

/// <summary>
/// What one "convert every equation to X" command did over its passes (one pass per source format).
/// </summary>
internal sealed class UnifiedFormulaConversionRun
{
    internal int Converted { get; set; }
    /// <summary>Formulas left in their format: the source format's name, the LaTeX (may be empty), why.</summary>
    internal List<(string Source, string Latex, string Reason)> Skipped { get; } = new();
    /// <summary>A pass failed, was cancelled or could not start: the passes after it do not run.</summary>
    internal bool Stopped { get; set; }
}

public sealed partial class ThisAddIn
{
    private static readonly string[] UnifiedFormulaModes =
    {
        FormulaOleContract.NativeOleMode,
        FormulaOleContract.MathTypeOleMode,
        FormulaOleContract.WordOmmlMode,
    };

    /// <summary>
    /// "Convert every equation to X" (tag "target|scope", scope "selection" or "document"): the user names only the
    /// format they want. Every equation of another format in the range becomes it, equations already in it stay as
    /// they are. Each source format is one pass through the same pipeline as the directional conversions (its
    /// numbering, references and table handling); every pass renders from the formula's LaTeX, so the same formula
    /// gives the same result whatever it was, at the document's sizes and number format.
    /// </summary>
    public void OnConvertFormulasTo(Office.IRibbonControl control)
    {
        var parts = (control?.Tag ?? string.Empty).Split('|');
        if (parts.Length != 2 || !UnifiedFormulaModes.Contains(parts[0], StringComparer.Ordinal)) return;
        if (parts[1] is not ("selection" or "document")) return;
        _ = ConvertAllFormulasToAsync(parts[1] == "document", parts[0]);
    }

    private async Task ConvertAllFormulasToAsync(bool wholeDocument, string targetMode)
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null || _formulaService is null || _lifetime is null || _lifetime.IsCancellationRequested)
            return;
        var targetName = FormatConversionModeLabel(targetMode);
        var others = UnifiedFormulaModes.Where(mode => mode != targetMode).ToArray();
        WordDoubleClickHook.TraceMessage(
            $"format-conversion-unified-start targetMode={targetMode} wholeDocument={wholeDocument}");

        if (wholeDocument && !IsConversionAcceptanceRun)
        {
            var otherNames = string.Join(T("、", ", "), others.Select(FormatConversionModeLabel));
            var confirmed = await dispatcher.InvokeAsync(() =>
                AutomationMessageBox.Show(
                    T($"将把全文中的 {otherNames} 公式全部转为 {targetName}；已是 {targetName} 的公式保持不变。\r\n\r\n旧公式和旧编号会直接删除，编号按当前 Word 编号设置重新创建。是否继续？",
                        $"Turn every {otherNames} equation in the document into {targetName}? Equations that already are {targetName} stay as they are.\r\n\r\nThe old equations and numbering are removed, and numbering is recreated using the current Word numbering settings."),
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
        }

        var run = new UnifiedFormulaConversionRun();
        foreach (var source in others)
        {
            await ConvertFormulaFormatAsync(wholeDocument, source, targetMode, run).ConfigureAwait(false);
            // The failed pass has reported itself; the formulas of the passes before it stay converted.
            if (run.Stopped) return;
        }
        WordDoubleClickHook.TraceMessage(
            $"format-conversion-unified-complete targetMode={targetMode} converted={run.Converted} skipped={run.Skipped.Count}");

        if (run.Converted == 0 && run.Skipped.Count == 0)
        {
            SetStatus(wholeDocument
                ? T($"全文没有需要转换的公式：公式已经都是 {targetName}。",
                    $"Nothing to convert: every equation in the document already is {targetName}.")
                : T($"所选范围没有需要转换的公式：没有公式，或公式已经都是 {targetName}。",
                    $"Nothing to convert in the selection: it has no equations, or they already are {targetName}."));
            return;
        }
        if (run.Skipped.Count == 0)
        {
            SetStatus(T($"公式格式转换完成：{run.Converted} 个公式已转为 {targetName}。",
                $"Equation format conversion finished: {run.Converted} equation(s) are now {targetName}."));
            return;
        }
        var summary = T(
            $"公式格式转换完成：{run.Converted} 个公式已转为 {targetName}，{run.Skipped.Count} 个公式无法转换，保持原格式。",
            $"Equation format conversion finished: {run.Converted} equation(s) are now {targetName}; {run.Skipped.Count} could not be converted and keep their format.");
        SetStatus(summary);
        if (IsConversionAcceptanceRun) return;
        var details = string.Join("\r\n", run.Skipped.Take(15).Select(item =>
            $"{item.Source} {item.Reason}"
            + (string.IsNullOrEmpty(item.Latex) ? string.Empty : $"（{WordFormulaRenderException.Excerpt(item.Latex, 60)}）")));
        await dispatcher.InvokeAsync(() =>
        {
            AutomationMessageBox.Show(
                summary + "\r\n\r\n" + details,
                ConversionTitle,
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Warning);
            return true;
        }).ConfigureAwait(false);
    }

    /// <summary>
    /// Every equation in the selection or the document back to LaTeX code, whatever its format (tag "selection" or
    /// "document"): one confirmation and one summary; each format is restored by its own service call (one undo
    /// step each), which keeps the selection over the same text for the next.
    /// </summary>
    public void OnFormulasToLatex(Office.IRibbonControl control)
    {
        var scope = control?.Tag ?? string.Empty;
        if (scope is not ("selection" or "document")) return;
        _ = ConvertAllFormulaObjectsToLatexAsync(scope == "document");
    }

    private async Task ConvertAllFormulaObjectsToLatexAsync(bool wholeDocument)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var lifetime = _lifetime;
        if (dispatcher is null || service is null || lifetime is null || lifetime.IsCancellationRequested)
            return;
        if (!await _operationGate.WaitAsync(TimeSpan.FromSeconds(2), lifetime.Token).ConfigureAwait(false))
        {
            SetStatus(T("VisualTeX 正在执行其他 Word 操作，请稍候再试。",
                "VisualTeX is busy with another Word operation. Try again in a moment."));
            return;
        }
        var title = T("VisualTeX 公式转为 LaTeX", "VisualTeX Equation to LaTeX");
        var progress = await dispatcher.InvokeAsync(() => BeginOperationProgress(title, service)).ConfigureAwait(false);
        try
        {
            var counts = new List<(string Mode, int Count)>();
            foreach (var mode in UnifiedFormulaModes)
            {
                var count = await dispatcher.InvokeAsync(() => service.CountFormulaObjectsForLatex(wholeDocument, mode))
                    .ConfigureAwait(false);
                if (count > 0) counts.Add((mode, count));
            }
            var total = counts.Sum(item => item.Count);
            if (total == 0)
            {
                SetStatus(wholeDocument
                    ? T("当前 Word 文档中没有公式。", "The current Word document has no equations.")
                    : T("所选内容中没有公式。", "The selection has no equations."));
                return;
            }
            if (wholeDocument && !IsConversionAcceptanceRun
                && !string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"), "1", StringComparison.Ordinal))
            {
                progress.Dispose();
                var confirmed = await dispatcher.InvokeAsync(() =>
                    AutomationMessageBox.Show(
                        progress.Owner,
                        T($"将把当前文档中的 {total} 个公式（不论格式）原位恢复为 LaTeX 代码。是否继续？",
                            $"Restore all {total} equations in the current document (any format) to LaTeX code in place?"),
                        title,
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button2)
                    == System.Windows.Forms.DialogResult.Yes).ConfigureAwait(false);
                if (!confirmed)
                {
                    SetStatus(T("已取消全文公式转为 LaTeX，Word 文档未修改。",
                        "Converting all equations to LaTeX was cancelled. The Word document was not changed."));
                    return;
                }
                progress = await dispatcher.InvokeAsync(() => BeginOperationProgress(title, service)).ConfigureAwait(false);
            }

            var restored = 0;
            var skipped = new List<string>();
            foreach (var (mode, count) in counts)
            {
                var label = FormatConversionModeLabel(mode);
                SetStatus(T($"正在把 {count} 个 {label} 公式恢复为 LaTeX 代码…",
                    $"Restoring {count} {label} equation(s) to LaTeX code…"));
                var result = await dispatcher.InvokeAsync(() => service.ConvertFormulaObjectsToLatex(wholeDocument, mode))
                    .ConfigureAwait(false);
                restored += result.FormulaCount;
                skipped.AddRange(result.Skipped.OrderBy(item => item.Ordinal).Select(item =>
                    T($"{label} 第 {item.Ordinal} 个：{item.Reason}", $"{label} No. {item.Ordinal}: {item.Reason}")));
            }
            WriteRedrawAcceptanceLog(
                $"formula-to-latex-unified-complete scope={(wholeDocument ? "document" : "selection")} formulas={restored} skipped={skipped.Count}");
            if (skipped.Count == 0)
            {
                SetStatus(T($"公式转为 LaTeX 完成：{restored} 个公式已恢复为源码。",
                    $"Equation to LaTeX finished: {restored} equation(s) were restored to source."));
                return;
            }
            var summary = T($"公式转为 LaTeX 完成：{restored} 个公式已恢复为源码，{skipped.Count} 个无法读取，保持原样。",
                $"Equation to LaTeX finished: {restored} equation(s) were restored to source; {skipped.Count} could not be read and were left as they are.");
            SetStatus(summary);
            if (string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"), "1", StringComparison.Ordinal)) return;
            progress.Dispose();
            await dispatcher.InvokeAsync(() =>
            {
                AutomationMessageBox.Show(summary + "\r\n\r\n" + string.Join("\r\n", skipped.Take(15)), title,
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                return true;
            }).ConfigureAwait(false);
        }
        catch (Exception error) when (IsUserCancellation(error, lifetime.Token))
        {
            SetStatus(T("VisualTeX 公式转为 LaTeX 已取消。", "VisualTeX equation-to-LaTeX conversion was cancelled."));
        }
        catch (Exception error)
        {
            var failureMessage = OfficePluginLanguage.FormatOperationError(
                error, "VisualTeX 公式转为 LaTeX 失败", "VisualTeX equation-to-LaTeX conversion failed");
            SetStatus(failureMessage);
            if (!string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"), "1", StringComparison.Ordinal))
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        AutomationMessageBox.Show(failureMessage, title,
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
        }
        finally { progress.Dispose(); _operationGate.Release(); }
    }
}

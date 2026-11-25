using System.Windows.Forms;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Office = Microsoft.Office.Core;

namespace VisualTeX.PowerPointVsto;

// Update Equation Format in PowerPoint and WPS Presentation (PowerPointFormulaService.FontRefresh.cs): the fonts the
// main app draws in now come from the companion; the VisualTeX formulas drawn in other fonts are drawn again from
// their LaTeX by the hidden converter (render batches, as conversions use) and put back by the service.
public sealed partial class ThisAddIn
{
    public void OnUpdateEquationFormat(Office.IRibbonControl control)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        // off PowerPoint's UI thread, as a conversion runs: PowerPoint keeps painting while formulas are drawn
        FormatConversionTask = Task.Run(() => UpdateEquationFormatAsync(lifetime.Token));
    }

    private async Task UpdateEquationFormatAsync(CancellationToken lifetimeToken)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = SessionClient;
        var application = _application;
        if (dispatcher is null || service is null || client is null || application is null) return;
        if (!await _operationGate.WaitAsync(TimeSpan.FromSeconds(2), lifetimeToken).ConfigureAwait(false))
        {
            ReportError(T(
                "VisualTeX 正在执行其他 PowerPoint 操作，请稍候再试。",
                "VisualTeX is busy with another PowerPoint operation. Try again in a moment."));
            return;
        }
        var title = T("VisualTeX 更新公式格式", "VisualTeX Update Equation Format");
        OfficeOperationProgress? progress = null;
        var temporaryFiles = new List<string>();
        var owner = IntPtr.Zero;
        try
        {
            owner = await dispatcher.InvokeAsync(() => new IntPtr(application.HWND)).ConfigureAwait(false);
            progress = NewConversionProgress(title, owner, T("正在读取 VisualTeX 的字体设置…", "Reading VisualTeX's font settings…"));
            await client.EnsureHealthyAsync(lifetimeToken).ConfigureAwait(false);
            var fonts = await client.GetFormulaFontSettingsAsync(lifetimeToken).ConfigureAwait(false);
            if (fonts is null)
            {
                // VisualTeX started without its main window (by an add-in) has not published its fonts: nothing to
                // compare with, so nothing is drawn again
                progress.Dispose();
                progress = null;
                PowerPointOperationLog.Write("font-refresh-unpublished");
                await ShowConversionMessageAsync(dispatcher, owner,
                    T("请先打开 VisualTeX 主窗口，再更新公式格式（本次未检查公式字体）。",
                        "Open the VisualTeX main window first, then update the equation format (formula fonts were not checked this time)."),
                    title, MessageBoxIcon.Information).ConfigureAwait(false);
                return;
            }
            progress.Report(T("正在扫描公式…", "Scanning equations…"));
            var targets = await dispatcher.InvokeAsync(() => service.CaptureFontRefreshTargets(fonts)).ConfigureAwait(false);
            progress.Dispose();
            progress = null;
            PowerPointOperationLog.Write($"font-refresh-plan formulas={targets.Count} fonts={fonts.Describe(true)}");
            if (targets.Count == 0)
            {
                await ShowConversionMessageAsync(dispatcher, owner,
                    T($"本演示文稿的 VisualTeX 公式已经使用当前字体（{fonts.Describe(false)}），无需更新。",
                        $"The VisualTeX equations in this presentation already use the current fonts ({fonts.Describe(true)})."),
                    title, MessageBoxIcon.Information).ConfigureAwait(false);
                return;
            }
            var slideCount = targets.Select(target => target.SlideId).Distinct().Count();
            if (!IsAcceptanceRun)
            {
                var message = T(
                    $"将用当前字体（{fonts.Describe(false)}）重绘 {targets.Count} 个字体与之不同的 VisualTeX 公式（{slideCount} 张幻灯片）。\r\n\r\n每个公式保持位置、显示大小、旋转和层级，公式内容不变。是否继续？",
                    $"Draw {targets.Count} VisualTeX equation(s) in other fonts again in the current fonts ({fonts.Describe(true)}), on {slideCount} slide(s)?\r\n\r\nEach equation keeps its position, size, rotation and layer; its content does not change.");
                var confirmed = await dispatcher.InvokeAsync(() =>
                        AutomationMessageBox.Show(new WindowOwner(owner), message, title, MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    .ConfigureAwait(false);
                if (!confirmed)
                {
                    SetStatus(T("已取消更新公式格式，演示文稿未修改。", "Update Equation Format cancelled. The presentation was not changed."));
                    return;
                }
            }
            progress = NewConversionProgress(title, owner, T("正在用当前字体重绘公式…", "Drawing equations in the current fonts…"));
            using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(lifetimeToken, progress.CancellationToken);
            var failures = new List<string>();
            var writes = await RenderFontRefreshAsync(client, targets, temporaryFiles, progress, failures, cancellation.Token)
                .ConfigureAwait(false);
            progress.Report(T("正在写入公式…", "Writing equations…"), 0, writes.Count);
            var written = writes.Count == 0
                ? 0
                : await dispatcher.InvokeAsync(() => service.ApplyFontRefresh(writes, failures)).ConfigureAwait(false);
            PowerPointOperationLog.Write($"font-refresh-complete formulas={targets.Count} written={written} failures={failures.Count}");
            var summary = T($"已用当前字体重绘 {written} 个 VisualTeX 公式。", $"{written} VisualTeX equation(s) drawn again in the current fonts.");
            SetStatus(summary);
            progress.Dispose();
            progress = null;
            if (failures.Count > 0)
                await ShowConversionMessageAsync(dispatcher, owner,
                    summary + "\r\n\r\n" + T($"{failures.Count} 个公式保持原来的字体：", $"{failures.Count} equation(s) keep their fonts:")
                        + "\r\n" + string.Join("\r\n", failures.Take(15)),
                    title, MessageBoxIcon.Warning).ConfigureAwait(false);
        }
        catch (Exception error) when (IsCancellation(error))
        {
            SetStatus(T("已停止更新公式格式，演示文稿未修改。", "Update Equation Format stopped. The presentation was not changed."));
        }
        catch (Exception error)
        {
            PowerPointOperationLog.Write("font-refresh-failed " + error);
            ReportError(T($"更新公式格式失败：{OfficeErrorText.Describe(error)}", $"Update Equation Format failed: {OfficeErrorText.Describe(error)}"));
        }
        finally
        {
            progress?.Dispose();
            _activeConversionProgress = null;
            _stopRequestedBeforeProgress = false;
            foreach (var path in temporaryFiles.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try { File.Delete(path); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            _operationGate.Release();
        }
    }

    /// <summary>
    /// Draws the formulas again in the current fonts at the size each shows at (20 to a render batch, no fonts named,
    /// so the converter takes the current ones) and prepares their files and sessions for the service.
    /// </summary>
    private async Task<List<PowerPointFormulaService.FontRefreshWrite>> RenderFontRefreshAsync(
        VisualTeXSessionClient client,
        IReadOnlyList<PowerPointFormulaService.FontRefreshTarget> targets,
        List<string> temporaryFiles,
        OfficeOperationProgress progress,
        List<string> failures,
        CancellationToken token)
    {
        var writes = new List<PowerPointFormulaService.FontRefreshWrite>();
        await client.PrewarmConverterAsync(token).ConfigureAwait(false);
        for (var offset = 0; offset < targets.Count; offset += RenderChunkSize)
        {
            token.ThrowIfCancellationRequested();
            progress.Report(T("正在用当前字体重绘公式…", "Drawing equations in the current fonts…"), offset, targets.Count);
            var chunk = targets.Skip(offset).Take(RenderChunkSize).ToList();
            var sizes = chunk.Select(target => PowerPointConversionRules.ResolveRenderFontSize(target.EffectiveFontSizePt)).ToList();
            var requests = chunk.Select((target, index) => new RenderBatchItemRequest
            {
                Id = Guid.NewGuid().ToString("N"),
                Request = CreateFontRefreshRenderRequest(target, sizes[index].RenderFontSizePt),
            }).ToList();
            var batch = await client.RenderBatchAsync(requests, TimeSpan.FromMinutes(3), token).ConfigureAwait(false);
            for (var index = 0; index < chunk.Count; index++)
            {
                var target = chunk[index];
                try
                {
                    var session = batch.Items.FirstOrDefault(item => item.Id == requests[index].Id)?.Session
                        ?? throw new InvalidOperationException(T("VisualTeX 没有返回该公式的渲染结果。", "VisualTeX returned no render result for this equation."));
                    if (session.Status == "failed")
                        throw new InvalidOperationException(OfficeErrorText.DescribeItemError(session.Error));
                    var export = session.ExportResult
                        ?? throw new InvalidOperationException(T("VisualTeX 没有返回该公式的导出结果。", "VisualTeX returned no export result for this equation."));
                    // the formula as it was: its id, form, rows and title; fonts and extent come with the drawing
                    var original = target.Metadata;
                    session.Mode = "edit";
                    session.FormulaId = original.FormulaId;
                    session.ObjectMode = target.ObjectMode;
                    session.DisplayMode = original.DisplayMode;
                    session.FontSizePt = sizes[index].RenderFontSizePt;
                    session.Title = string.IsNullOrWhiteSpace(original.Title) ? session.Title : original.Title;
                    if (original.Lines is { Count: > 0 })
                        session.Lines = original.Lines.Select(line => new FormulaLine { Id = line.Id, Latex = line.Latex }).ToList();
                    if (!string.IsNullOrWhiteSpace(original.CodeFormat)) session.CodeFormat = original.CodeFormat;
                    session.OriginalMetadata = original;
                    var write = new PowerPointFormulaService.FontRefreshWrite
                    {
                        Target = target,
                        Session = session,
                        PlacementScale = sizes[index].PlacementScale,
                        SvgPath = client.MaterializeSvg(session),
                    };
                    temporaryFiles.Add(write.SvgPath);
                    if (target.ObjectMode == FormulaOleContract.NativeOleMode)
                    {
                        write.PngPath = client.MaterializePng(session);
                        temporaryFiles.Add(write.PngPath);
                        write.EmfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                            write.SvgPath, export.Width, export.Height, usePowerPointStablePhysicalFrame: true);
                        temporaryFiles.Add(write.EmfPath);
                    }
                    writes.Add(write);
                }
                catch (Exception error) when (!IsCancellation(error))
                {
                    failures.Add(T($"第 {target.SlideIndex} 张幻灯片（{Shorten(target.Metadata.Latex)}）：", $"Slide {target.SlideIndex} ({Shorten(target.Metadata.Latex)}): ")
                        + OfficeErrorText.Describe(error));
                }
            }
        }
        progress.Report(T("正在用当前字体重绘公式…", "Drawing equations in the current fonts…"), targets.Count, targets.Count);
        return writes;
    }

    // As a conversion's render request, without the source's fonts: the converter draws in the current ones.
    private static CreateVstoSessionRequest CreateFontRefreshRenderRequest(
        PowerPointFormulaService.FontRefreshTarget target,
        float renderFontSizePt)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var source = target.Metadata;
        var lines = source.Lines is { Count: > 0 }
            ? source.Lines.Select(line => new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(line.Latex),
            }).ToList()
            : new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(source.Latex) },
            };
        var now = DateTimeOffset.UtcNow.ToString("O");
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "PowerPoint Formula",
            Latex = string.Join("\n", lines.Select(line => line.Latex)),
            Lines = lines,
            CodeFormat = source.CodeFormat,
            DisplayMode = source.DisplayMode,
            Numbered = false,
            FontSizePt = renderFontSizePt,
            RenderFontSizePt = renderFontSizePt,
            CreatedWithVersion = "1.3.1",
            UpdatedWithVersion = "1.3.1",
            CreatedAt = now,
            UpdatedAt = now,
        };
        return new CreateVstoSessionRequest
        {
            Mode = "create",
            Host = "powerpoint",
            FormulaId = formulaId,
            SourceDocumentId = target.DocumentId,
            Title = "PowerPoint Formula",
            Lines = lines,
            ActiveLineId = lines[0].Id,
            CodeFormat = source.CodeFormat,
            DisplayMode = source.DisplayMode,
            ObjectMode = target.ObjectMode,
            Numbered = false,
            FontSizePt = renderFontSizePt,
            OriginalMetadata = metadata,
            AutoCommitOnClose = false,
        };
    }
}

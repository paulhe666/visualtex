using System.Windows.Forms;
using Microsoft.Win32;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Office = Microsoft.Office.Core;

namespace VisualTeX.PowerPointVsto;

// Batch conversion between VisualTeX, MathType and native PowerPoint equations
// for the selection, the current slide or the whole presentation.
public sealed partial class ThisAddIn
{
    private const string FormatConversionRegistryPath = @"Software\VisualTeX\PowerPointVsto\FormatConversion";
    private const string VisualTeXTargetValueName = "VisualTeXTargetMode";
    private const int RenderChunkSize = 20;
    private const int MathTypeChunkSize = 12;
    private const int StopAfterConsecutiveFailures = 3;

    private static readonly string[] VisualTeXTargetControlIds =
    {
        "VisualTeX.PowerPointVsto.VisualTeXTargetForm",
        "VisualTeX.PowerPointVsto.VisualTeXTargetOle",
        "VisualTeX.PowerPointVsto.VisualTeXTargetSvg",
    };

    private string? _visualTeXTargetMode;
    private PowerPointModalGuard? _activeModalGuard;

    // The batch started by the last Ribbon click and how it ended; observed by
    // the VSTO flow acceptance, never by product code.
    internal Task? FormatConversionTask { get; private set; }
    internal PowerPointConversionOutcome? LastFormatConversionOutcome { get; private set; }

    internal sealed class PowerPointConversionOutcome
    {
        internal int Converted { get; set; }
        internal int Attempted { get; set; }
        internal int Total { get; set; }
        internal bool Cancelled { get; set; }
        internal string? Stopped { get; set; }
        internal string? Failure { get; set; }
        internal IReadOnlyList<PowerPointConversionSkip> Skips { get; set; } = Array.Empty<PowerPointConversionSkip>();
    }

    public void OnConvertFormulaFormat(object control)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        if (control is not Office.IRibbonControl ribbonControl
            || !PowerPointConversionRequest.TryParse(ribbonControl.Tag, VisualTeXTargetMode, out var request))
        {
            ReportError(T("无法识别公式格式转换命令。", "Unrecognized equation format conversion command."));
            return;
        }
        // Run the batch off PowerPoint's UI thread. Every PowerPoint call goes
        // through the dispatcher one unit at a time, so PowerPoint keeps
        // painting between equations instead of turning "Not Responding" on a
        // long batch that needs no rendering (e.g. MathType -> native).
        FormatConversionTask = Task.Run(() => ConvertFormulaFormatAsync(request, lifetime.Token));
    }

    public bool GetVisualTeXTargetFormPressed(Office.IRibbonControl control) =>
        string.Equals(control?.Tag, VisualTeXTargetMode, StringComparison.Ordinal);

    public void OnVisualTeXTargetFormChanged(Office.IRibbonControl control, bool pressed)
    {
        _ = pressed;
        var mode = control?.Tag;
        if (mode != FormulaOleContract.NativeOleMode && mode != FormulaOleContract.CrossPlatformPictureMode) return;
        _visualTeXTargetMode = mode;
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(FormatConversionRegistryPath, writable: true);
            key?.SetValue(VisualTeXTargetValueName, mode!, RegistryValueKind.String);
        }
        catch
        {
            // Keep the in-process choice even if the profile rejects persistence.
        }
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            foreach (var id in VisualTeXTargetControlIds) ui.InvalidateControl(id);
        }
        catch { }
    }

    public string GetVisualTeXTargetFormLabel(Office.IRibbonControl control) =>
        VisualTeXTargetMode == FormulaOleContract.CrossPlatformPictureMode
            ? T("VisualTeX 形式：SVG", "VisualTeX as: SVG")
            : T("VisualTeX 形式：OLE", "VisualTeX as: OLE");

    // What "-> VisualTeX" produces: a double-click-editable OLE object (the
    // default, like Word) or an SVG picture that displays on every device.
    private string VisualTeXTargetMode
    {
        get
        {
            if (_visualTeXTargetMode is not null) return _visualTeXTargetMode;
            var mode = FormulaOleContract.NativeOleMode;
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(FormatConversionRegistryPath, writable: false);
                if (key?.GetValue(VisualTeXTargetValueName) is string stored
                    && stored == FormulaOleContract.CrossPlatformPictureMode)
                    mode = stored;
            }
            catch { }
            _visualTeXTargetMode = mode;
            return mode;
        }
    }

    private async Task ConvertFormulaFormatAsync(PowerPointConversionRequest request, CancellationToken lifetimeToken)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = _sessionClient;
        var application = _application;
        if (dispatcher is null || service is null || client is null || application is null) return;
        if (!await _operationGate.WaitAsync(TimeSpan.FromSeconds(2), lifetimeToken).ConfigureAwait(false))
        {
            ReportError(T(
                "VisualTeX 正在执行其他 PowerPoint 操作，请稍候再试。",
                "VisualTeX is busy with another PowerPoint operation. Try again in a moment."));
            return;
        }

        var english = OfficePluginLanguage.IsEnglish;
        var title = T("VisualTeX 公式格式转换", "VisualTeX Equation Format Conversion");
        var targetLabel = PowerPointConversionRules.TargetLabel(request.TargetMode, english);
        OfficeOperationProgress? progress = null;
        PowerPointFormulaConversionPlan? plan = null;
        PowerPointMathTypeStaging? staging = null;
        var prepared = new Dictionary<string, PreparedPowerPointConversion>(StringComparer.Ordinal);
        var temporaryFiles = new List<string>();
        var replacements = new List<(int SlideId, int ShapeId)>();
        var owner = IntPtr.Zero;
        var converted = 0;
        var attempted = 0;
        var total = 0;
        var started = false;
        var cancelled = false;
        string? stoppedReason = null;
        Exception? failure = null;
        PowerPointOperationLog.Write($"format-conversion-start source={request.SourceFamily} target={request.TargetMode} scope={request.Scope}");
        // The outer finally owns the operation gate and the progress window, so
        // every early return (nothing found, confirmation declined) releases them.
        try
        {
            try
            {
                owner = await dispatcher.InvokeAsync(() => new IntPtr(application.HWND)).ConfigureAwait(false);
                progress = NewConversionProgress(title, owner, T("正在扫描公式…", "Scanning equations…"));
                plan = await dispatcher.InvokeAsync(() => service.CaptureFormulaConversionPlan(request)).ConfigureAwait(false);
                PowerPointOperationLog.Write(
                    $"format-conversion-plan slides={plan.ScopeSlideCount} formulas={plan.ItemCount} skipped={PowerPointConversionRules.CountSkips(plan.Skips)}");
                if (plan.ItemCount == 0)
                {
                    progress.Dispose();
                    progress = null;
                    await ShowConversionMessageAsync(dispatcher, owner, NothingToConvertMessage(plan, english), title, MessageBoxIcon.Information)
                        .ConfigureAwait(false);
                    return;
                }

                if (request.Scope == PowerPointConversionScope.Presentation && !IsAcceptanceRun)
                {
                    progress.Dispose();
                    progress = null;
                    var message = ConfirmationMessage(plan, targetLabel, english);
                    var confirmed = await dispatcher.InvokeAsync(() =>
                            MessageBox.Show(new WindowOwner(owner), message, title, MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                        .ConfigureAwait(false);
                    if (!confirmed)
                    {
                        SetStatus(T("已取消公式格式转换，演示文稿未修改。", "Equation format conversion cancelled. The presentation was not changed."));
                        return;
                    }
                    progress = NewConversionProgress(title, owner, T("正在准备转换…", "Preparing conversion…"));
                }

                // Like a modal dialog: no user input reaches PowerPoint until the
                // batch has finished, so nothing can change under the writer.
                await dispatcher.InvokeAsync(() =>
                {
                    _activeModalGuard?.Dispose();
                    _activeModalGuard = PowerPointModalGuard.Enter();
                    return true;
                }).ConfigureAwait(false);
                started = true;
                await dispatcher.InvokeAsync(() =>
                {
                    service.VerifyFormulaConversionSources(plan);
                    return true;
                }).ConfigureAwait(false);

                using var cancellation = CancellationTokenSource.CreateLinkedTokenSource(lifetimeToken, progress.CancellationToken);
                var token = cancellation.Token;
                await PrepareFormulaConversionAsync(client, plan, prepared, temporaryFiles, progress, token).ConfigureAwait(false);

                var work = new List<PowerPointConversionUnit>();
                foreach (var unit in plan.Units)
                {
                    var preparationFailure = unit.Items
                        .Select(item => prepared.TryGetValue(item.Id, out var formula) ? formula.Failure : "missing")
                        .FirstOrDefault(value => value is not null);
                    if (preparationFailure is null) work.Add(unit);
                    else plan.Skips.Add(new PowerPointConversionSkip(
                        PowerPointConversionSkipReason.Failed, unit.SlideIndex, unit.Items.Count, preparationFailure));
                }
                total = work.Sum(unit => unit.Items.Count);
                if (token.IsCancellationRequested) throw new OperationCanceledException(token);

                if (work.Count > 0 && request.TargetMode == FormulaOleContract.WordOmmlMode)
                {
                    progress.Report(T("正在准备 PowerPoint 原生公式…", "Preparing native PowerPoint equations…"), 0, total);
                    await dispatcher.InvokeAsync(() =>
                    {
                        PowerPointNativeOmmlKernel.Prewarm();
                        return true;
                    }).ConfigureAwait(false);
                }
                if (work.Count > 0 && request.TargetMode == FormulaOleContract.MathTypeOleMode)
                {
                    progress.Report(T("正在载入 MathType 公式…", "Loading MathType equations…"), 0, total);
                    var formulas = work.SelectMany(unit => unit.Items).Select(item => (item, prepared[item.Id])).ToList();
                    staging = await dispatcher.InvokeAsync(() => service.StageMathTypeConversions(formulas)).ConfigureAwait(false);
                }

                var consecutiveFailures = 0;
                var next = 0;
                while (next < work.Count)
                {
                    if (token.IsCancellationRequested)
                    {
                        cancelled = true;
                        break;
                    }
                    var chunk = new List<PowerPointConversionUnit>();
                    if (request.TargetMode == FormulaOleContract.MathTypeOleMode)
                    {
                        var formulasInChunk = 0;
                        while (next < work.Count && (chunk.Count == 0 || formulasInChunk + work[next].Items.Count <= MathTypeChunkSize))
                        {
                            formulasInChunk += work[next].Items.Count;
                            chunk.Add(work[next++]);
                        }
                    }
                    else
                    {
                        chunk.Add(work[next++]);
                    }
                    progress.Report(
                        T($"正在转换第 {attempted + 1}/{total} 个公式（第 {chunk[0].SlideIndex} 张幻灯片）…",
                            $"Converting equation {attempted + 1} of {total} (slide {chunk[0].SlideIndex})…"),
                        attempted, total);
                    List<PowerPointConversionUnitResult> results;
                    if (request.TargetMode == FormulaOleContract.MathTypeOleMode)
                    {
                        var stagingInUse = staging!;
                        var units = chunk;
                        results = await dispatcher.InvokeAsync(() =>
                            service.ApplyMathTypeConversionChunk(plan, units, prepared, stagingInUse)).ConfigureAwait(false);
                    }
                    else
                    {
                        var unit = chunk[0];
                        var single = await dispatcher.InvokeAsync(() => service.ApplyConversionUnit(plan, unit, prepared)).ConfigureAwait(false);
                        results = new List<PowerPointConversionUnitResult> { single };
                    }
                    for (var index = 0; index < results.Count; index++)
                    {
                        var unit = chunk[index];
                        var result = results[index];
                        attempted += unit.Items.Count;
                        if (result.Converted)
                        {
                            converted += unit.Items.Count;
                            consecutiveFailures = 0;
                            foreach (var shapeId in result.ReplacementShapeIds) replacements.Add((result.SlideId, shapeId));
                            continue;
                        }
                        var reason = result.Reason ?? PowerPointConversionSkipReason.Failed;
                        plan.Skips.Add(new PowerPointConversionSkip(reason, unit.SlideIndex, unit.Items.Count, result.Detail));
                        PowerPointOperationLog.Write(
                            $"format-conversion-unit-not-converted slide={unit.SlideIndex} shape={unit.ShapeId} reason={reason} detail={result.Detail}");
                        if (reason == PowerPointConversionSkipReason.Failed) consecutiveFailures++;
                    }
                    if (consecutiveFailures >= StopAfterConsecutiveFailures)
                    {
                        stoppedReason = T(
                            "连续多个公式转换失败，为避免继续出错已停止。",
                            "Several equations in a row failed to convert, so the conversion stopped.");
                        break;
                    }
                }
                progress.Report(T("正在完成…", "Finishing…"), attempted, total);
            }
            catch (OperationCanceledException) when (!lifetimeToken.IsCancellationRequested)
            {
                cancelled = true;
            }
            catch (Exception error)
            {
                failure = error;
                PowerPointOperationLog.Write($"format-conversion-failed error={error}");
            }
            finally
            {
                if (staging is not null)
                {
                    var closing = staging;
                    try { await dispatcher.InvokeAsync(() => { service.CloseMathTypeStaging(closing); return true; }).ConfigureAwait(false); }
                    catch { }
                }
                if (started && plan is not null)
                {
                    var finishing = plan;
                    var selected = replacements.ToArray();
                    try { await dispatcher.InvokeAsync(() => { service.FinishFormulaConversion(finishing, selected); return true; }).ConfigureAwait(false); }
                    catch { }
                }
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        _activeModalGuard?.Dispose();
                        _activeModalGuard = null;
                        return true;
                    }).ConfigureAwait(false);
                }
                catch
                {
                    _activeModalGuard?.Dispose();
                    _activeModalGuard = null;
                }
                foreach (var path in temporaryFiles.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try { File.Delete(path); } catch { }
                }
            }

            PowerPointOperationLog.Write(
                $"format-conversion-finished converted={converted} attempted={attempted} total={total} cancelled={cancelled} stopped={stoppedReason is not null} failed={failure is not null}");
            LastFormatConversionOutcome = new PowerPointConversionOutcome
            {
                Converted = converted,
                Attempted = attempted,
                Total = total,
                Cancelled = cancelled,
                Stopped = stoppedReason,
                Failure = failure?.ToString(),
                Skips = plan?.Skips.ToArray() ?? Array.Empty<PowerPointConversionSkip>(),
            };
            await ReportConversionOutcomeAsync(dispatcher, progress, owner, title, plan, targetLabel, english,
                converted, attempted, total, started, cancelled, stoppedReason, failure).ConfigureAwait(false);
        }
        finally
        {
            progress?.Dispose();
            _activeConversionProgress = null;
            _stopRequestedBeforeProgress = false;
            _operationGate.Release();
        }
    }

    private OfficeOperationProgress? _activeConversionProgress;

    private volatile bool _stopRequestedBeforeProgress;

    private OfficeOperationProgress NewConversionProgress(string title, IntPtr owner, string message)
    {
        var progress = new OfficeOperationProgress(title, owner, cancellable: true, message, PowerPointOperationLog.Write);
        _activeConversionProgress = progress;
        if (_stopRequestedBeforeProgress)
        {
            _stopRequestedBeforeProgress = false;
            progress.RequestCancel();
        }
        return progress;
    }

    // What the Stop button does, for the VSTO flow acceptance. The batch starts
    // on a worker thread, so a stop can arrive before its progress window exists.
    internal void RequestFormatConversionStop()
    {
        var progress = _activeConversionProgress;
        if (progress is null) _stopRequestedBeforeProgress = true;
        else progress.RequestCancel();
    }

    private static bool IsAcceptanceRun =>
        string.Equals(Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"), "1", StringComparison.Ordinal);

    private async Task ReportConversionOutcomeAsync(
        OfficeUiDispatcher dispatcher,
        OfficeOperationProgress? progress,
        IntPtr owner,
        string title,
        PowerPointFormulaConversionPlan? plan,
        string targetLabel,
        bool english,
        int converted,
        int attempted,
        int total,
        bool started,
        bool cancelled,
        string? stoppedReason,
        Exception? failure)
    {
        var skips = plan?.Skips ?? new List<PowerPointConversionSkip>();
        var skipText = PowerPointConversionRules.DescribeSkips(skips, english);
        if (failure is not null)
        {
            var detail = OfficePluginLanguage.FormatOperationError(failure, "公式格式转换失败", "Equation format conversion failed");
            var text = converted > 0
                ? detail + "\r\n\r\n" + T($"出错前已转换 {converted} 个公式，其余公式保持原样。",
                    $"{converted} equation(s) were converted before the error; the rest were left unchanged.")
                : detail + "\r\n\r\n" + T("演示文稿中的公式保持原样。", "The equations in the presentation were left unchanged.");
            progress?.Dispose();
            SetStatus(detail);
            await ShowConversionMessageAsync(dispatcher, owner, text, title, MessageBoxIcon.Error).ConfigureAwait(false);
            return;
        }
        if (plan is null) return;

        string headline;
        if (cancelled && converted == 0)
            headline = T("已停止，演示文稿未修改。", "Stopped. The presentation was not changed.");
        else if (cancelled)
            headline = T(
                $"已停止：{converted} 个公式已转换为 {targetLabel}，其余 {Math.Max(0, total - attempted)} 个保持原样。",
                $"Stopped: {converted} equation(s) were converted to {targetLabel}; the other {Math.Max(0, total - attempted)} were left unchanged.");
        else if (stoppedReason is not null)
            headline = stoppedReason + "\r\n" + T(
                $"已转换 {converted} 个公式，其余 {Math.Max(0, total - attempted)} 个未处理。",
                $"{converted} equation(s) were converted; {Math.Max(0, total - attempted)} were not attempted.");
        else
            headline = T($"已将 {converted} 个公式转换为 {targetLabel}。", $"Converted {converted} equation(s) to {targetLabel}.");
        SetStatus(headline);

        if (!cancelled && stoppedReason is null && skips.Count == 0 && started)
        {
            // Nothing needs the user's attention: show the result in place.
            progress?.Complete(headline, 2200);
            return;
        }
        progress?.Dispose();
        var body = skipText.Length == 0
            ? headline
            : headline + "\r\n\r\n" + T("以下公式未转换：", "Not converted:") + "\r\n" + skipText;
        var icon = converted == 0 && !cancelled ? MessageBoxIcon.Warning : MessageBoxIcon.Information;
        await ShowConversionMessageAsync(dispatcher, owner, body, title, icon).ConfigureAwait(false);
    }

    private static string ScopeText(PowerPointConversionScope scope, bool english) => scope switch
    {
        PowerPointConversionScope.Selection => english ? "the selection" : "所选内容",
        PowerPointConversionScope.Slide => english ? "this slide" : "当前幻灯片",
        _ => english ? "the presentation" : "整个演示文稿",
    };

    private static string SourceText(PowerPointConversionRequest request, bool english) =>
        request.Unify
            ? (english ? "equation(s) in other formats" : "其他格式的公式")
            : (english
                ? $"{PowerPointConversionRules.FamilyLabel(request.SourceFamily)} equation(s)"
                : $"{PowerPointConversionRules.FamilyLabel(request.SourceFamily)} 公式");

    private static string NothingToConvertMessage(PowerPointFormulaConversionPlan plan, bool english)
    {
        var request = plan.Request;
        var none = english
            ? $"No {SourceText(request, english)} to convert were found in {ScopeText(request.Scope, english)}."
            : $"{ScopeText(request.Scope, english)}中没有可转换的{SourceText(request, english)}。";
        if (plan.Skips.Count == 0) return none;
        return none + "\r\n\r\n" + (english ? "Found but not convertible:" : "找到但无法转换的公式：") + "\r\n"
            + PowerPointConversionRules.DescribeSkips(plan.Skips, english);
    }

    private static string ConfirmationMessage(PowerPointFormulaConversionPlan plan, string targetLabel, bool english)
    {
        var request = plan.Request;
        var slides = plan.Units.Select(unit => unit.SlideId).Distinct().Count();
        var text = english
            ? $"Convert {plan.ItemCount} {SourceText(request, english)} on {slides} slide(s) of the presentation to {targetLabel}?\r\n\r\n"
                + "Each equation is replaced in place; its position, visible size, animations, hyperlinks and alt text are kept. "
                + "PowerPoint can undo only a limited number of steps, so save the presentation first if you may want to go back."
            : $"将把整个演示文稿中 {slides} 张幻灯片上的 {plan.ItemCount} 个{SourceText(request, english)}转换为 {targetLabel}。\r\n\r\n"
                + "每个公式都在原位置替换，并保留位置、显示大小、动画、超链接和替代文字。PowerPoint 的撤销步数有限，如需随时回到原样，请先保存文件。";
        if (plan.Skips.Count > 0)
            text += "\r\n\r\n" + (english ? "These will be left unchanged:" : "以下公式不会转换：") + "\r\n"
                + PowerPointConversionRules.DescribeSkips(plan.Skips, english);
        return text + "\r\n\r\n" + (english ? "Continue?" : "是否继续？");
    }

    private static Task ShowConversionMessageAsync(
        OfficeUiDispatcher dispatcher,
        IntPtr owner,
        string message,
        string title,
        MessageBoxIcon icon)
    {
        if (IsAcceptanceRun)
        {
            PowerPointOperationLog.Write("format-conversion-message " + message.Replace("\r\n", " | "));
            return Task.CompletedTask;
        }
        return dispatcher.InvokeAsync(() =>
        {
            MessageBox.Show(new WindowOwner(owner), message, title, MessageBoxButtons.OK, icon);
            return true;
        });
    }

    private async Task PrepareFormulaConversionAsync(
        VisualTeXSessionClient client,
        PowerPointFormulaConversionPlan plan,
        Dictionary<string, PreparedPowerPointConversion> prepared,
        List<string> temporaryFiles,
        OfficeOperationProgress progress,
        CancellationToken token)
    {
        var target = plan.Request.TargetMode;
        var items = plan.Units.SelectMany(unit => unit.Items).ToList();
        foreach (var item in items)
        {
            var (render, scale) = PowerPointConversionRules.ResolveRenderFontSize(item.EffectiveFontSizePt);
            prepared[item.Id] = new PreparedPowerPointConversion
            {
                ItemId = item.Id,
                RenderFontSizePt = render,
                PlacementScale = scale,
            };
        }

        // MathType and native-equation sources already carry exact MathML; a
        // MathType or native target uses it directly instead of a LaTeX round trip.
        var acceptsSourceMathMl = target is FormulaOleContract.WordOmmlMode or FormulaOleContract.MathTypeOleMode;
        var toRender = new List<PowerPointConversionItem>();
        foreach (var item in items)
        {
            if (acceptsSourceMathMl && !string.IsNullOrWhiteSpace(item.SourceMathMl))
                prepared[item.Id].MathMl = item.SourceMathMl;
            else
                toRender.Add(item);
        }

        token.ThrowIfCancellationRequested();
        if (toRender.Count > 0)
        {
            progress.Report(T("正在连接 VisualTeX…", "Connecting to VisualTeX…"));
            await client.EnsureHealthyAsync(token).ConfigureAwait(false);
            await client.PrewarmConverterAsync(token).ConfigureAwait(false);
            var groups = toRender
                .GroupBy(item => RenderKeyFor(item, prepared[item.Id], target))
                .Select(group => group.ToList())
                .ToList();
            var rendered = 0;
            for (var offset = 0; offset < groups.Count; offset += RenderChunkSize)
            {
                token.ThrowIfCancellationRequested();
                progress.Report(T("正在渲染公式…", "Rendering equations…"), rendered, toRender.Count);
                var chunk = groups.Skip(offset).Take(RenderChunkSize).ToList();
                var requests = chunk.Select(group => new RenderBatchItemRequest
                {
                    Id = Guid.NewGuid().ToString("N"),
                    Request = CreateConversionRenderRequest(group[0], prepared[group[0].Id], target, plan.DocumentId),
                }).ToList();
                var batch = await client.RenderBatchAsync(requests, TimeSpan.FromMinutes(3), token).ConfigureAwait(false);
                for (var index = 0; index < chunk.Count; index++)
                {
                    var session = batch.Items.FirstOrDefault(item => item.Id == requests[index].Id)?.Session;
                    ApplyRenderedTemplate(client, chunk[index], session, target, prepared, temporaryFiles, plan.DocumentId);
                    rendered += chunk[index].Count;
                }
                progress.Report(T("正在渲染公式…", "Rendering equations…"), rendered, toRender.Count);
            }
        }

        if (target == FormulaOleContract.MathTypeOleMode)
            await Task.Run(() => PrepareMathTypeTargets(items, prepared, temporaryFiles, progress, token), token).ConfigureAwait(false);
    }

    private static string RenderKeyFor(PowerPointConversionItem item, PreparedPowerPointConversion formula, string target)
    {
        var lines = item.Lines is { Count: > 0 }
            ? string.Join("\u001E", item.Lines.Select(line => line.Latex))
            : item.Latex;
        return PowerPointConversionRules.RenderKey(
            target,
            item.DisplayMode,
            formula.RenderFontSizePt,
            item.SourceMetadata?.FormulaLetterFont,
            item.SourceMetadata?.FormulaChineseFont,
            item.CodeFormat + "\u001D" + lines);
    }

    private static List<FormulaLine> ConversionLines(PowerPointConversionItem item) =>
        item.Lines is { Count: > 0 }
            ? item.Lines.Select(line => new FormulaLine
            {
                Id = Guid.NewGuid().ToString("D"),
                Latex = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(line.Latex),
            }).ToList()
            : new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(item.Latex) },
            };

    private static CreateVstoSessionRequest CreateConversionRenderRequest(
        PowerPointConversionItem item,
        PreparedPowerPointConversion formula,
        string target,
        string documentId)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var lines = ConversionLines(item);
        var now = DateTimeOffset.UtcNow.ToString("O");
        var metadata = new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "PowerPoint Formula",
            Latex = string.Join("\n", lines.Select(line => line.Latex)),
            Lines = lines,
            CodeFormat = item.CodeFormat,
            DisplayMode = item.DisplayMode,
            Numbered = false,
            FontSizePt = formula.RenderFontSizePt,
            RenderFontSizePt = formula.RenderFontSizePt,
            FormulaLetterFont = item.SourceMetadata?.FormulaLetterFont,
            FormulaChineseFont = item.SourceMetadata?.FormulaChineseFont,
            CreatedWithVersion = "1.2.8",
            UpdatedWithVersion = "1.2.8",
            CreatedAt = now,
            UpdatedAt = now,
        };
        return new CreateVstoSessionRequest
        {
            Mode = "create",
            Host = "powerpoint",
            FormulaId = formulaId,
            SourceDocumentId = documentId,
            Title = "PowerPoint Formula",
            Lines = lines,
            ActiveLineId = lines[0].Id,
            CodeFormat = item.CodeFormat,
            DisplayMode = item.DisplayMode,
            ObjectMode = target,
            Numbered = false,
            FontSizePt = formula.RenderFontSizePt,
            OriginalMetadata = metadata,
            AutoCommitOnClose = false,
        };
    }

    private static void ApplyRenderedTemplate(
        VisualTeXSessionClient client,
        IReadOnlyList<PowerPointConversionItem> group,
        OfficeSessionDocument? session,
        string target,
        Dictionary<string, PreparedPowerPointConversion> prepared,
        List<string> temporaryFiles,
        string documentId)
    {
        void FailAll(string reason)
        {
            foreach (var item in group) prepared[item.Id].Failure = reason;
        }

        if (session is null)
        {
            FailAll(T("VisualTeX 没有返回该公式的渲染结果。", "VisualTeX returned no render result for this equation."));
            return;
        }
        if (session.Status == "cancelled" || session.ExplicitCancel)
            throw new OperationCanceledException();
        if (session.Status == "failed")
        {
            var detail = string.IsNullOrWhiteSpace(session.Error) || session.Error == "[object Object]"
                ? T("VisualTeX 无法渲染该公式。", "VisualTeX could not render this equation.")
                : session.Error!.Trim();
            FailAll(detail + T("（公式：", " (equation: ") + Shorten(group[0].Latex) + T("）", ")"));
            return;
        }
        var export = session.ExportResult;
        if (export is null || string.IsNullOrWhiteSpace(export.MathMl))
        {
            FailAll(T("VisualTeX 没有生成该公式的 MathML。", "VisualTeX produced no MathML for this equation."));
            return;
        }

        try
        {
            switch (target)
            {
                case FormulaOleContract.WordOmmlMode:
                case FormulaOleContract.MathTypeOleMode:
                    foreach (var item in group) prepared[item.Id].MathMl = export.MathMl;
                    return;
                case FormulaOleContract.NativeOleMode:
                {
                    var png = client.MaterializePng(session);
                    temporaryFiles.Add(png);
                    var svg = client.MaterializeSvg(session);
                    temporaryFiles.Add(svg);
                    var emf = OfficeOlePreview.CreateVectorEmfFromSvg(svg, export.Width, export.Height, usePowerPointStablePhysicalFrame: true);
                    temporaryFiles.Add(emf);
                    foreach (var item in group)
                    {
                        var formula = prepared[item.Id];
                        formula.PngPath = png;
                        formula.EmfPath = emf;
                        AssignVisualTeXSession(item, formula, export, target, documentId);
                    }
                    return;
                }
                default:
                {
                    var svg = client.MaterializeSvg(session);
                    temporaryFiles.Add(svg);
                    foreach (var item in group)
                    {
                        var formula = prepared[item.Id];
                        formula.SvgPath = svg;
                        AssignVisualTeXSession(item, formula, export, target, documentId);
                    }
                    return;
                }
            }
        }
        catch (Exception error) when (error is not OperationCanceledException)
        {
            FailAll(error.Message);
        }
    }

    private static void AssignVisualTeXSession(
        PowerPointConversionItem item,
        PreparedPowerPointConversion formula,
        OfficeExportDocument export,
        string target,
        string documentId)
    {
        formula.WidthPt = Math.Max(1f, export.Width * 0.75f) * formula.PlacementScale;
        formula.HeightPt = Math.Max(1f, export.Height * 0.75f) * formula.PlacementScale;
        formula.Session = new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "powerpoint",
            FormulaId = Guid.NewGuid().ToString("D"),
            SourceDocumentId = documentId,
            Title = "PowerPoint Formula",
            Lines = ConversionLines(item),
            CodeFormat = item.CodeFormat,
            DisplayMode = item.DisplayMode,
            ObjectMode = target,
            Numbered = false,
            FontSizePt = formula.RenderFontSizePt,
            Status = "committing",
            Dirty = true,
            ExportResult = export,
        };
    }

    // Same encoder, lossless check and MathType-native preview as a single
    // MathType edit, done once for the whole batch.
    private static void PrepareMathTypeTargets(
        IReadOnlyList<PowerPointConversionItem> items,
        Dictionary<string, PreparedPowerPointConversion> prepared,
        List<string> temporaryFiles,
        OfficeOperationProgress progress,
        CancellationToken token)
    {
        var inputs = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        var encoded = 0;
        foreach (var item in items)
        {
            token.ThrowIfCancellationRequested();
            var formula = prepared[item.Id];
            progress.Report(T("正在编码 MathType 公式…", "Encoding MathType equations…"), encoded++, items.Count);
            if (formula.Failure is not null) continue;
            if (string.IsNullOrWhiteSpace(formula.MathMl))
            {
                formula.Failure = T("缺少用于 MathType 的 MathML。", "No MathML is available for MathType.");
                continue;
            }
            try
            {
                var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(formula.MathMl!, inline: false, formula.RenderFontSizePt);
                var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
                if (!string.Equals(
                        MathTypeMtefCodec.SemanticSignature(formula.MathMl!),
                        MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(compound)),
                        StringComparison.Ordinal))
                    throw new InvalidDataException(T(
                        "该公式无法在不改变内容的情况下写成 MathType。",
                        "This equation cannot be written to MathType without changing its content."));
                formula.MathTypeCompoundFile = MathTypeOleStorage.RemovePresentationCaches(compound);
                formula.MathTypeEquationNative = generated.EquationNative;
                inputs[item.Id] = generated.Mtef;
            }
            catch (Exception error)
            {
                formula.Failure = error.Message + T("（公式：", " (equation: ") + Shorten(item.Latex) + T("）", ")");
            }
        }
        if (inputs.Count == 0) return;
        if (!MathTypeNativePreviewRenderer.IsAvailable)
            throw new InvalidOperationException(MathTypeNativePreviewRenderer.UnavailableMessage);

        var previews = RenderMathTypePreviews(inputs, progress, token, out var fontsMissing, out var previewFailures);
        try
        {
            if (fontsMissing) throw new InvalidOperationException(MathTypeNativePreviewRenderer.FontsUnavailableMessage);
            if (previews.Count == 0) throw new InvalidOperationException(MathTypeNativePreviewRenderer.FailureMessage);
            foreach (var pair in inputs)
            {
                var formula = prepared[pair.Key];
                if (!previews.TryGetValue(pair.Key, out var preview))
                {
                    formula.Failure = previewFailures.TryGetValue(pair.Key, out var reason)
                        ? MathTypeNativePreviewRenderer.IncompleteMessage + "（原因：" + reason + "）"
                        : MathTypeNativePreviewRenderer.IncompleteMessage;
                    continue;
                }
                try
                {
                    var emf = PowerPointMathTypePreview.FromNativeWmf(preview.WmfPath, preview.WidthPt, preview.HeightPt);
                    temporaryFiles.Add(emf);
                    formula.MathTypePreviewEmf = File.ReadAllBytes(emf);
                    formula.MathTypeNaturalWidthPt = preview.WidthPt;
                    formula.MathTypeNaturalHeightPt = preview.HeightPt;
                    formula.WidthPt = Math.Max(1f, preview.WidthPt * formula.PlacementScale);
                    formula.HeightPt = Math.Max(1f, preview.HeightPt * formula.PlacementScale);
                }
                catch (Exception error)
                {
                    formula.Failure = error.Message;
                }
            }
        }
        finally
        {
            foreach (var preview in previews.Values) preview.Dispose();
        }
    }

    // Grouped rendering with retries (NativePreviewRetry): a renderer crash
    // costs time, not the batch, and a broken renderer fails fast.
    private static Dictionary<string, MathTypeNativePreviewRenderer.Result> RenderMathTypePreviews(
        IReadOnlyDictionary<string, byte[]> inputs,
        OfficeOperationProgress progress,
        CancellationToken token,
        out bool fontsMissing,
        out Dictionary<string, string> failures)
    {
        var fonts = false;
        var reasons = new Dictionary<string, string>(StringComparer.Ordinal);
        var message = T("正在生成 MathType 排版预览…", "Rendering MathType previews…");
        var rendered = NativePreviewRetry.Render<MathTypeNativePreviewRenderer.Result>(
            inputs.Keys.ToList(),
            group =>
            {
                token.ThrowIfCancellationRequested();
                MathTypeNativePreviewRenderer.TryRenderBatch(
                    group.ToDictionary(id => id, id => inputs[id], StringComparer.Ordinal),
                    Path.GetTempPath(),
                    out var results);
                fonts |= MathTypeNativePreviewRenderer.LastFailureWasFonts;
                foreach (var failure in MathTypeNativePreviewRenderer.LastFailures) reasons[failure.Key] = failure.Value;
                foreach (var key in results.Keys) reasons.Remove(key);
                return results;
            },
            new NativePreviewRetryOptions
            {
                Pause = pause =>
                {
                    if (token.WaitHandle.WaitOne(pause)) token.ThrowIfCancellationRequested();
                },
                StopNow = () => fonts,
                Progress = (done, total) => progress.Report(message, done, total),
                Retrying = (pass, missing, size) => PowerPointOperationLog.Write(
                    $"format-conversion-mathtype-preview-retry pass={pass} missing={missing} group={size}"),
            });
        fontsMissing = fonts;
        failures = reasons;
        return rendered;
    }

    private static string Shorten(string latex)
    {
        var single = (latex ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        return single.Length <= 120 ? single : single.Substring(0, 117) + "...";
    }

    private sealed class WindowOwner : IWin32Window
    {
        internal WindowOwner(IntPtr handle) => Handle = handle;
        public IntPtr Handle { get; }
    }
}

// %LOCALAPPDATA%\VisualTeX\office\logs\powerpoint-operations.log, written off
// the Office thread and size-capped like the double-click trace.
internal static class PowerPointOperationLog
{
    private const long LimitBytes = 512 * 1024;
    private static readonly object Gate = new();
    private static readonly int ProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;

    internal static void Write(string message)
    {
        var line = $"{DateTimeOffset.Now:O} pid={ProcessId} {message}{Environment.NewLine}";
        ThreadPool.QueueUserWorkItem(_ =>
        {
            try
            {
                var directory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "VisualTeX", "office", "logs");
                var path = Path.Combine(directory, "powerpoint-operations.log");
                lock (Gate)
                {
                    Directory.CreateDirectory(directory);
                    var file = new FileInfo(path);
                    if (file.Exists && file.Length > LimitBytes)
                    {
                        var previous = path + ".1";
                        if (File.Exists(previous)) File.Delete(previous);
                        File.Move(path, previous);
                    }
                    File.AppendAllText(path, line);
                }
            }
            catch { }
        });
    }
}

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Extensibility;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.PowerPointVsto;

[ComVisible(true)]
[Guid("29C64025-AB17-4F25-9B89-6E1D8D22C2D7")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IPowerPointRibbonCallbacks
{
    [DispId(1)]
    void OnRibbonLoad(object ribbonUi);

    [DispId(2)]
    void OnNewFormula(object control);

    [DispId(3)]
    void OnEditSelected(object control);

    [DispId(4)]
    void OnConvertSelected(object control);

    [DispId(5)]
    void OnExportSelectedAsPicture(object control);

    [DispId(6)]
    void OnDeleteSelected(object control);

    [DispId(7)]
    void OnOpenDesktop(object control);

    [DispId(8)]
    object? GetRibbonImage(Office.IRibbonControl control);

    [DispId(9)]
    string GetFormulaFontSizeText(Office.IRibbonControl control);

    [DispId(10)]
    bool GetFormulaFontSizeEnabled(Office.IRibbonControl control);

    [DispId(11)]
    void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value);

    [DispId(12)]
    void OnDecreaseFormulaFontSize(Office.IRibbonControl control);

    [DispId(13)]
    void OnIncreaseFormulaFontSize(Office.IRibbonControl control);

    [DispId(14)]
    void OnConvertSelectedOmml(object control);
}

[ComVisible(true)]
[Guid("7E586D2D-57B0-4D14-AB24-EBA9021A5E6D")]
[ProgId("VisualTeX.PowerPointVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IPowerPointRibbonCallbacks))]
public sealed class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IPowerPointRibbonCallbacks
{
    private const int AllowAnyProcessToSetForeground = -1;

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AllowSetForegroundWindow(int processId);

    static ThisAddIn() => VstoDependencyResolver.Install();

    private static void GrantVisualTeXForegroundActivation()
    {
        try { _ = AllowSetForegroundWindow(AllowAnyProcessToSetForeground); }
        catch { }
    }

    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="VisualTeX">
          <button id="VisualTeX.PowerPointVsto.New" label="新建公式" size="large" tag="insertFormula" getImage="GetRibbonImage" onAction="OnNewFormula" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <button id="VisualTeX.PowerPointVsto.ConvertOmml" label="转为 OMML" screentip="转为 PowerPoint 原生公式" supertip="转换为 PowerPoint 原生 Office Math（PPTX 内部以 OMML 保存），可继续使用 PowerPoint 公式工具编辑。" imageMso="EquationInsertNew" onAction="OnConvertSelectedOmml" />
          <button id="VisualTeX.PowerPointVsto.ConvertSelected" label="转为原生 OLE" screentip="转为可嵌入编辑的原生 OLE" supertip="转换后外观应保持不变，但对象会嵌入 PowerPoint 文件，并可通过 VisualTeX 双击重新编辑。" tag="convertToOle" getImage="GetRibbonImage" onAction="OnConvertSelected" />
          <button id="VisualTeX.PowerPointVsto.ExportPicture" label="转为 SVG 图片" imageMso="PictureInsertFromFile" onAction="OnExportSelectedAsPicture" />
          <button id="VisualTeX.PowerPointVsto.Delete" label="删除所选公式" imageMso="Delete" onAction="OnDeleteSelected" />
          <button id="VisualTeX.PowerPointVsto.OpenDesktop" label="打开 VisualTeX" imageMso="FileOpen" onAction="OnOpenDesktop" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FontSizeGroup" label="公式字号">
          <button id="VisualTeX.PowerPointVsto.FontSizeDecrease" label="减小" imageMso="FontSizeDecrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
          <comboBox id="VisualTeX.PowerPointVsto.FontSize" label="字号" sizeString="初号（42 磅）" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
            <item id="VisualTeX.PowerPointVsto.FontSizeChu" label="初号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoChu" label="小初" />
            <item id="VisualTeX.PowerPointVsto.FontSizeYi" label="一号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoYi" label="小一" />
            <item id="VisualTeX.PowerPointVsto.FontSizeEr" label="二号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoEr" label="小二" />
            <item id="VisualTeX.PowerPointVsto.FontSizeSan" label="三号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoSan" label="小三" />
            <item id="VisualTeX.PowerPointVsto.FontSizeSi" label="四号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoSi" label="小四" />
            <item id="VisualTeX.PowerPointVsto.FontSizeWu" label="五号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoWu" label="小五" />
            <item id="VisualTeX.PowerPointVsto.FontSizeLiu" label="六号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoLiu" label="小六" />
            <item id="VisualTeX.PowerPointVsto.FontSizeQi" label="七号" />
            <item id="VisualTeX.PowerPointVsto.FontSizeBa" label="八号" />
            <item id="VisualTeX.PowerPointVsto.FontSize8" label="8" />
            <item id="VisualTeX.PowerPointVsto.FontSize9" label="9" />
            <item id="VisualTeX.PowerPointVsto.FontSize10" label="10" />
            <item id="VisualTeX.PowerPointVsto.FontSize10_5" label="10.5" />
            <item id="VisualTeX.PowerPointVsto.FontSize11" label="11" />
            <item id="VisualTeX.PowerPointVsto.FontSize12" label="12" />
            <item id="VisualTeX.PowerPointVsto.FontSize14" label="14" />
            <item id="VisualTeX.PowerPointVsto.FontSize16" label="16" />
            <item id="VisualTeX.PowerPointVsto.FontSize18" label="18" />
            <item id="VisualTeX.PowerPointVsto.FontSize20" label="20" />
            <item id="VisualTeX.PowerPointVsto.FontSize24" label="24" />
            <item id="VisualTeX.PowerPointVsto.FontSize28" label="28" />
            <item id="VisualTeX.PowerPointVsto.FontSize36" label="36" />
            <item id="VisualTeX.PowerPointVsto.FontSize48" label="48" />
            <item id="VisualTeX.PowerPointVsto.FontSize72" label="72" />
          </comboBox>
          <button id="VisualTeX.PowerPointVsto.FontSizeIncrease" label="增大" imageMso="FontSizeIncrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    private Application? _application;
    private PowerPointFormulaService? _formulaService;
    private OfficeUiDispatcher? _dispatcher;
    private VisualTeXSessionClient? _sessionClient;
    private PowerPointDoubleClickHook? _doubleClickHook;
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private CancellationTokenSource? _lifetime;
    private string _lastDoubleClickFormulaId = string.Empty;
    private DateTimeOffset _lastDoubleClickAt;
    private string? _activeSessionId;
    private OfficeSelection? _lastFormulaSelection;
    private int _formulaFontInvalidationPending;
    private object? _ribbonUi;
    private Office.COMAddIn? _comAddIn;

    public string DiagnosticLastError { get; private set; } = string.Empty;

    public string GetCustomUI(string ribbonId) => RibbonXml;

    public void OnConnection(
        object application,
        ext_ConnectMode connectMode,
        object addInInstance,
        ref Array custom)
    {
        _application = (Application)application;
        _comAddIn = addInInstance as Office.COMAddIn;
        if (_comAddIn is not null)
        {
            try { _comAddIn.Object = this; } catch { }
        }
        _dispatcher = new OfficeUiDispatcher();
        _formulaService = new PowerPointFormulaService(
            _application,
            _dispatcher.Post,
            _dispatcher.PostDelayed);
        _sessionClient = new VisualTeXSessionClient();
        _lifetime = new CancellationTokenSource();
        _ = PrewarmCompanionAsync(_sessionClient, _lifetime.Token);
        _application.WindowSelectionChange += OnWindowSelectionChange;
        _application.PresentationBeforeClose += OnPresentationBeforeClose;
        string? doubleClickError = null;
        try
        {
            _doubleClickHook = new PowerPointDoubleClickHook(OnNativeDoubleClick);
            _doubleClickHook.Start();
        }
        catch (Exception error)
        {
            try { _doubleClickHook?.Dispose(); } catch { }
            _doubleClickHook = null;
            doubleClickError = error.Message;
        }
        SetStatus(doubleClickError is null
            ? "VisualTeX PowerPoint VSTO 已就绪。"
            : $"VisualTeX 已就绪，但双击监听不可用：{doubleClickError}");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) => Dispose();

    public void OnRibbonLoad(object ribbonUi)
    {
        _ribbonUi = ribbonUi;
        InvalidateFormulaFontControls();
    }
    public object? GetRibbonImage(Office.IRibbonControl control) =>
        RibbonIconProvider.GetImage(control?.Tag);
    public string GetFormulaFontSizeText(Office.IRibbonControl control)
    {
        try
        {
            var size = _formulaService?.GetSelectedFormulaFontSize();
            return size.HasValue
                ? FormulaFontSize.FormatDisplay(size.Value)
                : string.Empty;
        }
        catch { return string.Empty; }
    }
    public bool GetFormulaFontSizeEnabled(Office.IRibbonControl control)
    {
        try { return _formulaService?.GetSelectedFormulaFontSize().HasValue == true; }
        catch { return false; }
    }
    public void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value) =>
        ApplyFormulaFontSize(ParseFontSize(value));
    public void OnDecreaseFormulaFontSize(Office.IRibbonControl control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException("请先选择一个公式。");
            ApplyFormulaFontSize(FormulaFontSize.PreviousPreset(current));
        }
        catch (Exception error) { ReportError($"无法设置公式字号：{error.Message}"); }
    }
    public void OnIncreaseFormulaFontSize(Office.IRibbonControl control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException("请先选择一个公式。");
            ApplyFormulaFontSize(FormulaFontSize.NextPreset(current));
        }
        catch (Exception error) { ReportError($"无法设置公式字号：{error.Message}"); }
    }
    public void OnNewFormula(object control) => BeginSession("create", "crossPlatformPicture", null);
    public void OnEditSelected(object control) => BeginSelectedSession(null);
    public void OnConvertSelectedOmml(object control) =>
        BeginSelectedSession("wordOmml", conversionOnly: true);
    public void OnConvertSelected(object control) =>
        BeginSelectedSession("nativeOle", conversionOnly: true);
    public void OnExportSelectedAsPicture(object control) =>
        BeginSelectedSession("crossPlatformPicture", conversionOnly: true);
    public void OnDeleteSelected(object control) => _ = DeleteSelectedAsync();
    public void OnOpenDesktop(object control)
    {
        try
        {
            (_sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable."))
                .OpenDesktop();
            SetStatus("VisualTeX 已打开。");
        }
        catch (Exception error)
        {
            SetStatus($"无法打开 VisualTeX：{error.Message}");
        }
    }

    private static async Task PrewarmCompanionAsync(
        VisualTeXSessionClient client,
        CancellationToken cancellationToken)
    {
        try
        {
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            await client.PrewarmConverterAsync(cancellationToken).ConfigureAwait(false);
        }
        catch
        {
            // Keep PowerPoint startup non-blocking. Any explicit VisualTeX action
            // will retry the companion path and report a real error to the user.
        }
    }

    private static double ParseFontSize(string value) => FormulaFontSize.Parse(value);

    private void ApplyFormulaFontSize(double value)
    {
        try
        {
            var applied = (_formulaService
                    ?? throw new InvalidOperationException("PowerPoint formula service is unavailable."))
                .SetSelectedFormulaFontSize(value);
            SetStatus($"公式字号已设置为 {FormulaFontSize.Describe(applied)}。");
        }
        catch (Exception error)
        {
            ReportError($"无法设置公式字号：{error.Message}");
        }
        finally { InvalidateFormulaFontControls(); }
    }

    private void ScheduleFormulaFontControlsInvalidation()
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null
            || Interlocked.Exchange(ref _formulaFontInvalidationPending, 1) != 0)
            return;
        dispatcher.Post(() =>
        {
            Interlocked.Exchange(ref _formulaFontInvalidationPending, 0);
            InvalidateFormulaFontControls();
        });
    }

    private void InvalidateFormulaFontControls()
    {
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            ui.InvalidateControl("VisualTeX.PowerPointVsto.FontSize");
            ui.InvalidateControl("VisualTeX.PowerPointVsto.FontSizeDecrease");
            ui.InvalidateControl("VisualTeX.PowerPointVsto.FontSizeIncrease");
        }
        catch { }
    }

    private void OnWindowSelectionChange(Selection selection)
    {
        ScheduleFormulaFontControlsInvalidation();
        var service = _formulaService;
        if (service is null) return;
        try
        {
            var current = service.ReadSelection(selection);
            if (current.Metadata is not null)
            {
                _lastFormulaSelection = current;
                return;
            }
            // Clicking the Ribbon can temporarily expose ppSelectionNone. Keep
            // the last formula only for that transient state; any explicit
            // non-formula selection invalidates the cache.
            if (selection.Type == PpSelectionType.ppSelectionShapes)
                _lastFormulaSelection = null;
        }
        catch
        {
            // Selection notifications must never destabilize PowerPoint.
        }
    }

    private void OnPresentationBeforeClose(Presentation presentation, ref bool cancel)
    {
        _lastFormulaSelection = null;
    }

    private void OnNativeDoubleClick(int screenX, int screenY)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        dispatcher.Post(() =>
        {
            var selected = service.ReadFormulaAtScreenPoint(screenX, screenY);
            if (selected?.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
                return;
            _lastFormulaSelection = selected;
            var now = DateTimeOffset.UtcNow;
            if (selected.FormulaId == _lastDoubleClickFormulaId
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
                return;
            _lastDoubleClickFormulaId = selected.FormulaId!;
            _lastDoubleClickAt = now;
            BeginSession("edit", null, selected);
        });
    }

    private void BeginSelectedSession(
        string? requestedObjectMode,
        bool conversionOnly = false)
    {
        DiagnosticLastError = string.Empty;
        try
        {
            var service = _formulaService
                ?? throw new InvalidOperationException("PowerPoint formula service is unavailable.");
            var selection = ResolveFormulaSelection(service);
            if (selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");
            BeginSession("edit", requestedObjectMode, selection, conversionOnly);
        }
        catch (Exception error)
        {
            ReportError($"VisualTeX PowerPoint 写入失败：{error.Message}");
        }
    }

    private void BeginSession(
        string mode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
        bool conversionOnly = false)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(
            mode,
            requestedObjectMode,
            capturedSelection,
            conversionOnly,
            lifetime.Token);
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
        bool conversionOnly,
        CancellationToken cancellationToken)
    {
        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                cancellationToken).ConfigureAwait(false))
        {
            var activeSessionId = Volatile.Read(ref _activeSessionId);
            if (!string.IsNullOrWhiteSpace(activeSessionId) && _sessionClient is not null)
            {
                try
                {
                    GrantVisualTeXForegroundActivation();
                    await _sessionClient.OpenEditorAsync(activeSessionId!, cancellationToken)
                        .ConfigureAwait(false);
                    SetStatus("已有 VisualTeX 编辑任务，已将编辑窗口切换到前台。");
                }
                catch (Exception error)
                {
                    SetStatus($"已有编辑任务，但无法置前窗口：{error.Message}");
                }
            }
            else
            {
                SetStatus("VisualTeX 正在准备编辑窗口，请稍候再试。");
            }
            return;
        }

        DiagnosticLastError = string.Empty;
        string? sessionId = null;
        string? imagePath = null;
        string? svgPath = null;
        string? emfPath = null;
        try
        {
            var dispatcher = _dispatcher ?? throw new InvalidOperationException("PowerPoint dispatcher is unavailable.");
            var service = _formulaService ?? throw new InvalidOperationException("PowerPoint formula service is unavailable.");
            var client = _sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable.");
            SetStatus("正在连接 VisualTeX 本地服务…");
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            if (conversionOnly)
                await client.PrewarmConverterAsync(cancellationToken).ConfigureAwait(false);
            var selection = capturedSelection?.Metadata is not null
                ? capturedSelection
                : await dispatcher.InvokeAsync(
                    () => ResolveFormulaSelection(service)).ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 PowerPoint 演示文稿为只读状态。");
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

            // PowerPoint commonly leaves the just-inserted formula selected.
            // Do not treat that selection as initial content for New Formula;
            // only an explicit edit command may reuse existing metadata.
            var metadata = mode == "edit" ? selection.Metadata : null;
            var targetObjectMode = requestedObjectMode
                ?? (mode == "create" ? "crossPlatformPicture" : selection.ObjectMode)
                ?? "crossPlatformPicture";
            var requiresObjectModeChange = mode == "edit"
                && !string.Equals(
                    selection.ObjectMode,
                    targetObjectMode,
                    StringComparison.Ordinal);
            var lines = metadata?.Lines ?? new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = string.Empty },
            };
            var fontSizePt = metadata?.FontSizePt
                ?? await dispatcher.InvokeAsync(service.ReadCurrentTypingFontSize)
                    .ConfigureAwait(false);
            var request = new CreateVstoSessionRequest
            {
                Mode = mode,
                Host = "powerpoint",
                FormulaId = metadata?.FormulaId,
                SourceDocumentId = selection.DocumentId,
                SourceObjectId = mode == "edit" ? selection.ObjectId : null,
                Title = metadata?.Title ?? "PowerPoint Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = "block",
                ObjectMode = targetObjectMode,
                Numbered = false,
                FontSizePt = FormulaFontSize.Normalize(fontSizePt, 20f),
                OriginalMetadata = metadata,
                AutoCommitOnClose = true,
            };
            var session = await client.CreateSessionAsync(request, cancellationToken).ConfigureAwait(false);
            sessionId = session.Id;
            Volatile.Write(ref _activeSessionId, session.Id);
            if (conversionOnly)
            {
                await client.OpenConverterAsync(session.Id, cancellationToken)
                    .ConfigureAwait(false);
                SetStatus("正在直接转换 PowerPoint 公式格式…");
            }
            else
            {
                GrantVisualTeXForegroundActivation();
                await client.OpenEditorAsync(session.Id, cancellationToken)
                    .ConfigureAwait(false);
                SetStatus("VisualTeX 编辑器已打开。");
            }
            session = await client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromMinutes(30),
                cancellationToken).ConfigureAwait(false);
            if (session.Status == "cancelled" || session.ExplicitCancel)
            {
                SetStatus("已取消，PowerPoint 未修改。");
                return;
            }
            if (session.Status == "failed")
                throw new InvalidOperationException(session.Error ?? "VisualTeX Session 失败。");
            if (session.Mode == "edit"
                && !session.Dirty
                && (!requiresObjectModeChange || session.ExportResult is null))
            {
                await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
                SetStatus(requiresObjectModeChange
                    ? "未执行对象格式转换。"
                    : "公式内容未变化。");
                return;
            }

            if (session.ObjectMode == "nativeOle")
            {
                imagePath = client.MaterializePng(session);
                var export = session.ExportResult
                    ?? throw new InvalidOperationException("VisualTeX Session has no vector export result.");
                svgPath = client.MaterializeSvg(session);
                emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                    svgPath,
                    export.Width,
                    export.Height,
                    usePowerPointStablePhysicalFrame: true);
            }
            else if (session.ObjectMode == "crossPlatformPicture")
            {
                imagePath = client.MaterializeSvg(session);
            }
            else if (session.ObjectMode != "wordOmml")
            {
                throw new InvalidOperationException($"Unsupported PowerPoint object mode: {session.ObjectMode}");
            }
            var writeResult = await dispatcher.InvokeAsync(() =>
            {
                var current = service.ReadSelection();
                if (!string.Equals(
                        current.DocumentId,
                        session.SourceDocumentId,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("活动演示文稿已切换，未写入公式。");
                if (session.ObjectMode == "nativeOle")
                {
                    if (imagePath is null || emfPath is null)
                        throw new InvalidOperationException("VisualTeX native OLE preview is unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceOle(session, imagePath, emfPath)
                        : service.InsertOle(session, imagePath, emfPath);
                }
                if (session.ObjectMode == "wordOmml")
                {
                    return session.Mode == "edit"
                        ? service.ReplaceOmml(session)
                        : service.InsertOmml(session);
                }
                if (imagePath is null)
                    throw new InvalidOperationException("VisualTeX picture export is unavailable.");
                return session.Mode == "edit"
                    ? service.Replace(session, imagePath)
                    : service.Insert(session, imagePath);
            }).ConfigureAwait(false);
            _lastFormulaSelection = new OfficeSelection
            {
                Host = "powerpoint",
                DocumentId = writeResult.DocumentId,
                ObjectId = writeResult.ObjectId,
                ReadOnly = false,
                FormulaId = session.FormulaId,
                Metadata = session.ToMetadata(),
                ObjectMode = session.ObjectMode,
            };
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            SetStatus(requiresObjectModeChange && session.ObjectMode == "nativeOle"
                ? "已转换为原生 OLE：外观保持不变，可双击编辑，并嵌入 PowerPoint 文件。"
                : requiresObjectModeChange && session.ObjectMode == "wordOmml"
                    ? "已转换为 PowerPoint 原生 Office Math（OMML），可继续使用 PowerPoint 公式工具编辑。"
                    : requiresObjectModeChange && session.ObjectMode == "crossPlatformPicture"
                        ? "已转换为嵌入式 SVG 图片，可跨平台显示并保持矢量清晰度。"
                        : session.Mode == "edit" ? "PowerPoint 公式已更新。" : "PowerPoint 公式已插入。");
        }
        catch (OperationCanceledException)
        {
            SetStatus("VisualTeX 操作已取消。");
        }
        catch (Exception error)
        {
            if (sessionId is not null && _sessionClient is not null)
            {
                try
                {
                    await _sessionClient.FailAsync(sessionId, error.Message, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            ReportError($"VisualTeX PowerPoint 写入失败：{error.Message}");
        }
        finally
        {
            if (emfPath is not null)
            {
                try { File.Delete(emfPath); } catch { }
            }
            if (svgPath is not null)
            {
                try { File.Delete(svgPath); } catch { }
            }
            if (imagePath is not null)
            {
                try { File.Delete(imagePath); } catch { }
            }
            if (sessionId is not null
                && string.Equals(
                    Volatile.Read(ref _activeSessionId),
                    sessionId,
                    StringComparison.Ordinal))
                Volatile.Write(ref _activeSessionId, null);
            _operationGate.Release();
        }
    }

    private async Task DeleteSelectedAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            await dispatcher.InvokeAsync(service.DeleteSelectedFormula).ConfigureAwait(false);
            SetStatus("PowerPoint 公式已删除。");
        }
        catch (Exception error)
        {
            ReportError($"删除 PowerPoint 公式失败：{error.Message}");
        }
    }

    private void SetStatus(string message)
    {
        var dispatcher = _dispatcher;
        var application = _application;
        if (dispatcher is null || application is null) return;
        dispatcher.Post(() =>
        {
            try { ((dynamic)application).StatusBar = message; } catch { }
        });
    }

    private OfficeSelection ResolveFormulaSelection(PowerPointFormulaService service)
    {
        var current = service.ReadSelection();
        if (current.Metadata is not null)
        {
            _lastFormulaSelection = current;
            return current;
        }
        var cached = _lastFormulaSelection;
        if (cached?.Metadata is null
            || !string.Equals(
                cached.DocumentId,
                current.DocumentId,
                StringComparison.OrdinalIgnoreCase))
            return current;
        return new OfficeSelection
        {
            Host = current.Host,
            DocumentId = current.DocumentId,
            ObjectId = cached.ObjectId,
            ReadOnly = current.ReadOnly,
            FormulaId = cached.FormulaId,
            Metadata = cached.Metadata,
            ObjectMode = cached.ObjectMode,
        };
    }

    private void ReportError(string message)
    {
        DiagnosticLastError = message;
        SetStatus(message);
        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;
        var dispatcher = _dispatcher;
        if (dispatcher is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            try
            {
                System.Windows.Forms.MessageBox.Show(
                    message,
                    "VisualTeX PowerPoint",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
            catch { }
            return true;
        });
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }

    private void Dispose()
    {
        _lifetime?.Cancel();
        if (_application is not null)
        {
            try { _application.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
            try { _application.PresentationBeforeClose -= OnPresentationBeforeClose; } catch { }
        }
        _doubleClickHook?.Dispose();
        _sessionClient?.Dispose();
        _dispatcher?.Dispose();
        _lifetime?.Dispose();
        _doubleClickHook = null;
        _sessionClient = null;
        _dispatcher = null;
        _formulaService = null;
        _lastFormulaSelection = null;
        Volatile.Write(ref _activeSessionId, null);
        _lifetime = null;
        _ribbonUi = null;
        if (_comAddIn is not null)
        {
            try { _comAddIn.Object = null; } catch { }
            ReleaseComObject(_comAddIn);
            _comAddIn = null;
        }
        _application = null;
    }
}

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
}

[ComVisible(true)]
[Guid("7E586D2D-57B0-4D14-AB24-EBA9021A5E6D")]
[ProgId("VisualTeX.PowerPointVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IPowerPointRibbonCallbacks))]
public sealed class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IPowerPointRibbonCallbacks
{
    static ThisAddIn() => VstoDependencyResolver.Install();

    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="VisualTeX">
          <button id="VisualTeX.PowerPointVsto.New" label="新建公式" size="large" tag="insertFormula" getImage="GetRibbonImage" onAction="OnNewFormula" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <button id="VisualTeX.PowerPointVsto.ConvertSelected" label="转为原生 OLE" screentip="转为可嵌入编辑的原生 OLE" supertip="转换后外观应保持不变，但对象会嵌入 PowerPoint 文件，并可通过 VisualTeX 双击重新编辑。" tag="convertToOle" getImage="GetRibbonImage" onAction="OnConvertSelected" />
          <button id="VisualTeX.PowerPointVsto.ExportPicture" label="导出所选为图片" imageMso="PictureInsertFromFile" onAction="OnExportSelectedAsPicture" />
          <button id="VisualTeX.PowerPointVsto.Delete" label="删除所选公式" imageMso="Delete" onAction="OnDeleteSelected" />
          <button id="VisualTeX.PowerPointVsto.OpenDesktop" label="打开 VisualTeX" imageMso="FileOpen" onAction="OnOpenDesktop" />
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
    private object? _ribbonUi;

    public string DiagnosticLastError { get; private set; } = string.Empty;

    public string GetCustomUI(string ribbonId) => RibbonXml;

    public void OnConnection(
        object application,
        ext_ConnectMode connectMode,
        object addInInstance,
        ref Array custom)
    {
        _application = (Application)application;
        _formulaService = new PowerPointFormulaService(_application);
        _dispatcher = new OfficeUiDispatcher();
        _sessionClient = new VisualTeXSessionClient();
        _lifetime = new CancellationTokenSource();
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

    public void OnRibbonLoad(object ribbonUi) => _ribbonUi = ribbonUi;
    public object? GetRibbonImage(Office.IRibbonControl control) =>
        RibbonIconProvider.GetImage(control?.Tag);
    public void OnNewFormula(object control) => BeginSession("create", "crossPlatformPicture", null);
    public void OnEditSelected(object control) => BeginSelectedSession(null);
    public void OnConvertSelected(object control) => BeginSelectedSession("nativeOle");
    public void OnExportSelectedAsPicture(object control) => _ = ExportSelectedAsPictureAsync();
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

    private void OnWindowSelectionChange(Selection selection)
    {
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

    private void OnNativeDoubleClick()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            var selected = ResolveFormulaSelection(service);
            if (string.IsNullOrWhiteSpace(selected.FormulaId)) return false;
            var now = DateTimeOffset.UtcNow;
            if (selected.FormulaId == _lastDoubleClickFormulaId
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
                return false;
            _lastDoubleClickFormulaId = selected.FormulaId!;
            _lastDoubleClickAt = now;
            BeginSession("edit", null, selected);
            return true;
        });
    }

    private void BeginSelectedSession(string? requestedObjectMode)
    {
        DiagnosticLastError = string.Empty;
        try
        {
            var service = _formulaService
                ?? throw new InvalidOperationException("PowerPoint formula service is unavailable.");
            var selection = ResolveFormulaSelection(service);
            if (selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");
            BeginSession("edit", requestedObjectMode, selection);
        }
        catch (Exception error)
        {
            ReportError($"VisualTeX PowerPoint 写入失败：{error.Message}");
        }
    }

    private void BeginSession(
        string mode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(mode, requestedObjectMode, capturedSelection, lifetime.Token);
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
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
                OriginalMetadata = metadata,
                AutoCommitOnClose = true,
            };
            var session = await client.CreateSessionAsync(request, cancellationToken).ConfigureAwait(false);
            sessionId = session.Id;
            Volatile.Write(ref _activeSessionId, session.Id);
            await client.OpenEditorAsync(session.Id, cancellationToken).ConfigureAwait(false);
            SetStatus("VisualTeX 编辑器已打开。");
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
                    export.Height);
            }
            else
            {
                // PowerPoint supports SVG as a native picture format. Insert the
                // original vector export directly instead of rasterizing it to
                // PNG, so scaling and PDF export retain sharp formula glyphs.
                imagePath = client.MaterializeSvg(session);
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
                    if (emfPath is null)
                        throw new InvalidOperationException("VisualTeX native OLE vector preview is unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceOle(session, imagePath, emfPath)
                        : service.InsertOle(session, imagePath, emfPath);
                }
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

    private async Task ExportSelectedAsPictureAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            await dispatcher.InvokeAsync(service.ExportSelectedOleAsPicture)
                .ConfigureAwait(false);
            SetStatus("PowerPoint OLE 公式已导出为跨平台图片。");
        }
        catch (Exception error)
        {
            ReportError($"导出 PowerPoint OLE 公式失败：{error.Message}");
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
        _ = dispatcher.InvokeAsync(() =>
        {
            try { ((dynamic)application).StatusBar = message; } catch { }
            return true;
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
        _application = null;
    }
}

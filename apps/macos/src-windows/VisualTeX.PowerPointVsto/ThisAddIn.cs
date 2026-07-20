using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Extensibility;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

[ComVisible(true)]
[Guid("7E586D2D-57B0-4D14-AB24-EBA9021A5E6D")]
[ProgId("VisualTeX.PowerPointVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IDTExtensibility2))]
public sealed class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility
{
    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="VisualTeX">
          <button id="VisualTeX.PowerPointVsto.New" label="新建公式" size="large" imageMso="EquationInsertNew" onAction="OnNewFormula" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="编辑所选公式" size="large" imageMso="ObjectEdit" onAction="OnEditSelected" />
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
    private object? _ribbonUi;

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
        // Keep the previous global low-level mouse hook disabled. Editing remains
        // available through the native "编辑所选公式" Ribbon command.
        _doubleClickHook = null;
        SetStatus("VisualTeX PowerPoint VSTO 已就绪。");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) => Dispose();

    public void OnRibbonLoad(object ribbonUi) => _ribbonUi = ribbonUi;
    public void OnNewFormula(object control) => BeginSession("create");
    public void OnEditSelected(object control) => BeginSession("edit");

    private void OnWindowSelectionChange(Selection selection)
    {
        // The current selection is read only when a command starts.
    }

    private void OnPresentationBeforeClose(Presentation presentation, ref bool cancel)
    {
        // The lifetime token is process-wide. Session source identities prevent
        // a result from being applied to another presentation after switching.
    }

    private void OnNativeDoubleClick()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            var selected = service.ReadSelection();
            if (string.IsNullOrWhiteSpace(selected.FormulaId)) return false;
            var now = DateTimeOffset.UtcNow;
            if (selected.FormulaId == _lastDoubleClickFormulaId
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
                return false;
            _lastDoubleClickFormulaId = selected.FormulaId;
            _lastDoubleClickAt = now;
            BeginSession("edit");
            return true;
        });
    }

    private void BeginSession(string mode)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(mode, lifetime.Token);
    }

    private async Task RunSessionAsync(string mode, CancellationToken cancellationToken)
    {
        if (!await _operationGate.WaitAsync(0, cancellationToken).ConfigureAwait(false))
        {
            SetStatus("VisualTeX 编辑窗口已经打开。");
            return;
        }

        string? sessionId = null;
        string? imagePath = null;
        try
        {
            var dispatcher = _dispatcher ?? throw new InvalidOperationException("PowerPoint dispatcher is unavailable.");
            var service = _formulaService ?? throw new InvalidOperationException("PowerPoint formula service is unavailable.");
            var client = _sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable.");
            SetStatus("正在连接 VisualTeX 本地服务…");
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            var selection = await dispatcher.InvokeAsync(service.ReadSelection).ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 PowerPoint 演示文稿为只读状态。");
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

            var metadata = selection.Metadata;
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
                SourceObjectId = selection.ObjectId,
                Title = metadata?.Title ?? "PowerPoint Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = "block",
                Numbered = false,
                OriginalMetadata = metadata,
                AutoCommitOnClose = true,
            };
            var session = await client.CreateSessionAsync(request, cancellationToken).ConfigureAwait(false);
            sessionId = session.Id;
            client.OpenEditor(session.Id);
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
            if (session.Mode == "edit" && !session.Dirty)
            {
                await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
                SetStatus("公式内容未变化。");
                return;
            }

            imagePath = client.MaterializePng(session);
            await dispatcher.InvokeAsync(() =>
            {
                var current = service.ReadSelection();
                if (!string.Equals(
                        current.DocumentId,
                        session.SourceDocumentId,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("活动演示文稿已切换，未写入公式。");
                return session.Mode == "edit"
                    ? service.Replace(session, imagePath)
                    : service.Insert(session, imagePath);
            }).ConfigureAwait(false);
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            SetStatus(session.Mode == "edit" ? "PowerPoint 公式已更新。" : "PowerPoint 公式已插入。");
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
            SetStatus($"VisualTeX PowerPoint 写入失败：{error.Message}");
        }
        finally
        {
            if (imagePath is not null)
            {
                try { File.Delete(imagePath); } catch { }
            }
            _operationGate.Release();
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
        _lifetime = null;
        _ribbonUi = null;
        _application = null;
    }
}

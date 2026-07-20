using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Extensibility;
using Office = Microsoft.Office.Core;
using Task = System.Threading.Tasks.Task;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

[ComVisible(true)]
[Guid("F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1")]
[ProgId("VisualTeX.WordVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IDTExtensibility2))]
public sealed class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility
{
    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab idMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="行内公式" size="large" imageMso="EquationInsertNew" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="行间公式" size="large" imageMso="EquationInsertNew" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.Edit" label="编辑所选公式" size="large" imageMso="ObjectEdit" onAction="OnEditSelected" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    private Application? _application;
    private WordFormulaService? _formulaService;
    private OfficeUiDispatcher? _dispatcher;
    private VisualTeXSessionClient? _sessionClient;
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
        _formulaService = new WordFormulaService(_application);
        _dispatcher = new OfficeUiDispatcher();
        _sessionClient = new VisualTeXSessionClient();
        _lifetime = new CancellationTokenSource();
        _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
        _application.WindowSelectionChange += OnWindowSelectionChange;
        SetStatus("VisualTeX Word VSTO 已就绪。");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) => Dispose();

    public void OnRibbonLoad(object ribbonUi) => _ribbonUi = ribbonUi;
    public void OnInsertInline(object control) => BeginSession("create", "inline");
    public void OnInsertDisplay(object control) => BeginSession("create", "block");
    public void OnEditSelected(object control) => BeginSession("edit", null);

    private void OnWindowSelectionChange(Selection selection)
    {
        // Selection is read lazily by the command to avoid retaining COM objects.
    }

    private void OnWindowBeforeDoubleClick(Selection selection, ref bool cancel)
    {
        try
        {
            var selected = _formulaService?.ReadSelection();
            var formulaId = selected?.FormulaId;
            if (string.IsNullOrWhiteSpace(formulaId)) return;
            var now = DateTimeOffset.UtcNow;
            if (formulaId == _lastDoubleClickFormulaId
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
                return;
            _lastDoubleClickFormulaId = formulaId;
            _lastDoubleClickAt = now;
            cancel = true;
            BeginSession("edit", null);
        }
        catch (Exception error)
        {
            SetStatus($"VisualTeX 双击检测失败：{error.Message}");
        }
    }

    private void BeginSession(string mode, string? displayMode)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(mode, displayMode, lifetime.Token);
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedDisplayMode,
        CancellationToken cancellationToken)
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
            var dispatcher = _dispatcher ?? throw new InvalidOperationException("Word dispatcher is unavailable.");
            var service = _formulaService ?? throw new InvalidOperationException("Word formula service is unavailable.");
            var client = _sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable.");
            SetStatus("正在连接 VisualTeX 本地服务…");
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            var selection = await dispatcher.InvokeAsync(service.ReadSelection).ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
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
                Host = "word",
                FormulaId = metadata?.FormulaId,
                SourceDocumentId = selection.DocumentId,
                SourceObjectId = selection.ObjectId,
                Title = metadata?.Title ?? "Word Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = requestedDisplayMode ?? metadata?.DisplayMode ?? "inline",
                Numbered = (requestedDisplayMode ?? metadata?.DisplayMode) == "block"
                    && (metadata?.Numbered ?? false),
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
                SetStatus("已取消，Word 文档未修改。");
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
                    throw new InvalidOperationException("活动 Word 文档已切换，未写入公式。");
                return session.Mode == "edit"
                    ? service.Replace(session, imagePath)
                    : service.Insert(session, imagePath);
            }).ConfigureAwait(false);
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            SetStatus(session.Mode == "edit" ? "Word 公式已更新。" : "Word 公式已插入。");
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
            SetStatus($"VisualTeX Word 写入失败：{error.Message}");
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
            try { application.StatusBar = message; } catch { }
            return true;
        });
    }

    private void Dispose()
    {
        _lifetime?.Cancel();
        if (_application is not null)
        {
            try { _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick; } catch { }
            try { _application.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
        }
        _sessionClient?.Dispose();
        _dispatcher?.Dispose();
        _lifetime?.Dispose();
        _sessionClient = null;
        _dispatcher = null;
        _formulaService = null;
        _lifetime = null;
        _ribbonUi = null;
        _application = null;
    }
}

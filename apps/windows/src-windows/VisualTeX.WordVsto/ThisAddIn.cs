using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using Extensibility;
using Office = Microsoft.Office.Core;
using Task = System.Threading.Tasks.Task;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

[ComVisible(true)]
[Guid("D4A1A3CB-0ED7-4B2F-8A2B-5CB0B1E25421")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IWordRibbonCallbacks
{
    [DispId(1)]
    void OnRibbonLoad(object ribbonUi);

    [DispId(2)]
    void OnInsertInline(object control);

    [DispId(3)]
    void OnInsertDisplay(object control);

    [DispId(4)]
    void OnEditSelected(object control);

    [DispId(5)]
    void OnConvertSelected(object control);

    [DispId(6)]
    void OnUpdateEquationNumbers(object control);

    [DispId(7)]
    void OnExportSelectedAsPicture(object control);

    [DispId(8)]
    void OnDeleteSelected(object control);

    [DispId(9)]
    void OnOpenDesktop(object control);

    [DispId(10)]
    void OnInsertEquationReference(object control);

    [DispId(11)]
    void OnInsertInlineOmml(object control);

    [DispId(12)]
    void OnInsertDisplayOmml(object control);

    [DispId(13)]
    void OnConvertSelectedToOmml(object control);

    [DispId(14)]
    object? GetRibbonImage(Office.IRibbonControl control);
}

[ComVisible(true)]
[Guid("F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1")]
[ProgId("VisualTeX.WordVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IWordRibbonCallbacks))]
public sealed class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IWordRibbonCallbacks
{
    static ThisAddIn() => VstoDependencyResolver.Install();

    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="OLE 行内公式" size="large" tag="oleInline" getImage="GetRibbonImage" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="OLE 行间公式" size="large" tag="oleDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.InlineOmml" label="OMML 行内公式" size="large" screentip="插入 Word 原生公式" supertip="插入可由 Word 原生公式工具直接编辑、同时保留 VisualTeX LaTeX 元数据的 OMML 行内公式。" tag="ommlInline" getImage="GetRibbonImage" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.WordVsto.DisplayOmml" label="OMML 行间公式" size="large" screentip="插入 Word 原生公式" supertip="插入可由 Word 原生公式工具直接编辑、同时保留 VisualTeX LaTeX 元数据的 OMML 行间公式。" tag="ommlDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.WordVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <button id="VisualTeX.WordVsto.ConvertSelected" label="转为原生 OLE" screentip="转为可嵌入编辑的原生 OLE" supertip="转换后对象随 Word 文档保存，并可通过 VisualTeX 双击重新编辑。" tag="convertToOle" getImage="GetRibbonImage" onAction="OnConvertSelected" />
          <button id="VisualTeX.WordVsto.ConvertSelectedToOmml" label="转为 Word OMML" screentip="转为 Word 原生公式" supertip="将所选 VisualTeX 公式转换为 Word 原生 OMML；可在 Word 中直接编辑，也可继续用 VisualTeX 编辑。" tag="convertToOmml" getImage="GetRibbonImage" onAction="OnConvertSelectedToOmml" />
          <button id="VisualTeX.WordVsto.UpdateNumbers" label="更新公式编号" tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationNumbers" />
          <button id="VisualTeX.WordVsto.InsertReference" label="插入公式引用" screentip="引用带编号公式" supertip="从当前文档的带编号公式中选择目标，并插入可自动更新的 Word REF 字段。" imageMso="HyperlinkInsert" onAction="OnInsertEquationReference" />
          <button id="VisualTeX.WordVsto.ExportPicture" label="导出所选为图片" imageMso="PictureInsertFromFile" onAction="OnExportSelectedAsPicture" />
          <button id="VisualTeX.WordVsto.Delete" label="删除所选公式" imageMso="Delete" onAction="OnDeleteSelected" />
          <button id="VisualTeX.WordVsto.OpenDesktop" label="打开 VisualTeX" imageMso="FileOpen" onAction="OnOpenDesktop" />
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
    private WordDoubleClickHook? _doubleClickHook;
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private readonly object _nativeOleTargetGate = new();
    private CancellationTokenSource? _lifetime;
    private string _lastDoubleClickFormulaId = string.Empty;
    private DateTimeOffset _lastDoubleClickAt;
    private string? _activeSessionId;
    private bool _nativeOleTargetActive;
    private int _nativeOleTargetLeft;
    private int _nativeOleTargetTop;
    private int _nativeOleTargetRight;
    private int _nativeOleTargetBottom;
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
        string? doubleClickError = null;
        try
        {
            _doubleClickHook = new WordDoubleClickHook(
                ShouldInterceptNativeOleDoubleClick,
                OnNativeOleDoubleClick);
            _doubleClickHook.Start();
        }
        catch (Exception error)
        {
            try { _doubleClickHook?.Dispose(); } catch { }
            _doubleClickHook = null;
            doubleClickError = error.Message;
        }
        SetStatus(doubleClickError is null
            ? "VisualTeX Word VSTO 已就绪。"
            : $"VisualTeX 已就绪，但 OLE 双击监听不可用：{doubleClickError}");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
    public void OnBeginShutdown(ref Array custom) => Dispose();

    public void OnRibbonLoad(object ribbonUi) => _ribbonUi = ribbonUi;
    public object? GetRibbonImage(Office.IRibbonControl control) =>
        RibbonIconProvider.GetImage(control?.Tag);
    public void OnInsertInline(object control) =>
        BeginSession("create", "inline", FormulaOleContract.NativeOleMode);
    public void OnInsertDisplay(object control) =>
        BeginSession("create", "block", FormulaOleContract.NativeOleMode);
    public void OnInsertInlineOmml(object control) =>
        BeginSession("create", "inline", FormulaOleContract.WordOmmlMode);
    public void OnInsertDisplayOmml(object control) =>
        BeginSession("create", "block", FormulaOleContract.WordOmmlMode);
    public void OnEditSelected(object control) => BeginSession("edit", null, null);
    public void OnConvertSelected(object control) =>
        BeginSession(
            "edit",
            null,
            FormulaOleContract.NativeOleMode,
            conversionOnly: true);
    public void OnConvertSelectedToOmml(object control) =>
        BeginSession(
            "edit",
            null,
            FormulaOleContract.WordOmmlMode,
            conversionOnly: true);
    public void OnUpdateEquationNumbers(object control) => _ = UpdateEquationNumbersAsync();
    public void OnExportSelectedAsPicture(object control) => _ = ExportSelectedAsPictureAsync();
    public void OnDeleteSelected(object control) => _ = DeleteSelectedAsync();
    public void OnInsertEquationReference(object control) => _ = InsertEquationReferenceAsync();
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
        var application = _application;
        if (service is null || application is null)
        {
            ClearNativeOleTarget();
            return;
        }

        Range? range = null;
        Window? window = null;
        try
        {
            // Do not inspect OMML metadata here. Word fires SelectionChange while
            // entering its native equation editor, and touching the OMath at that
            // point can disturb the caret state. Only perform the heavier metadata
            // read after the fast OLE type check succeeds.
            if (!service.IsSelectedNativeOle())
            {
                ClearNativeOleTarget();
                return;
            }
            var selected = service.ReadSelection(selection);
            if (!WordDoubleClickRouting.ShouldOpenVisualTeX(selected)
                || !string.Equals(
                    selected.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
            {
                ClearNativeOleTarget();
                return;
            }

            range = selection.Range;
            window = application.ActiveWindow;
            window.GetPoint(
                out var left,
                out var top,
                out var width,
                out var height,
                range);
            if (width <= 0 || height <= 0)
            {
                ClearNativeOleTarget();
                return;
            }
            const int padding = 4;
            lock (_nativeOleTargetGate)
            {
                _nativeOleTargetLeft = left - padding;
                _nativeOleTargetTop = top - padding;
                _nativeOleTargetRight = left + width + padding;
                _nativeOleTargetBottom = top + height + padding;
                _nativeOleTargetActive = true;
            }
            WordDoubleClickHook.TraceMessage(
                $"cache-active formulaId={selected.FormulaId} rect={left - padding},{top - padding},{left + width + padding},{top + height + padding}");
        }
        catch
        {
            ClearNativeOleTarget();
        }
        finally
        {
            ReleaseComObject(window);
            ReleaseComObject(range);
        }
    }

    private void OnWindowBeforeDoubleClick(Selection selection, ref bool cancel)
    {
        try
        {
            var selected = _formulaService?.ReadSelection(selection);
            if (selected?.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
                return;

            if (!WordDoubleClickRouting.ShouldOpenVisualTeX(selected)) return;

            cancel = true;
            ClearNativeOleTarget();
            TryBeginDoubleClickSession(selected);
        }
        catch (Exception error)
        {
            SetStatus($"VisualTeX 双击检测失败：{error.Message}");
        }
    }

    private bool ShouldInterceptNativeOleDoubleClick(int screenX, int screenY)
    {
        lock (_nativeOleTargetGate)
        {
            return _nativeOleTargetActive
                && screenX >= _nativeOleTargetLeft
                && screenX <= _nativeOleTargetRight
                && screenY >= _nativeOleTargetTop
                && screenY <= _nativeOleTargetBottom;
        }
    }

    private void OnNativeOleDoubleClick()
    {
        WordDoubleClickHook.TraceMessage("addin-callback-received");
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            try
            {
                var selected = service.ReadSelection();
                WordDoubleClickHook.TraceMessage(
                    $"addin-selection formulaId={selected.FormulaId ?? "<null>"} objectMode={selected.ObjectMode ?? "<null>"}");
                if (!string.Equals(
                        selected.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                    return false;
                var started = TryBeginDoubleClickSession(selected);
                WordDoubleClickHook.TraceMessage($"addin-session-started={started}");
                return started;
            }
            catch (Exception error)
            {
                SetStatus($"VisualTeX OLE 双击检测失败：{error.Message}");
                return false;
            }
        });
    }

    private bool TryBeginDoubleClickSession(OfficeSelection? selected)
    {
        var formulaId = selected?.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId)
            || !WordDoubleClickRouting.ShouldOpenVisualTeX(selected))
            return false;

        var now = DateTimeOffset.UtcNow;
        if (formulaId == _lastDoubleClickFormulaId
            && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
            return false;
        _lastDoubleClickFormulaId = formulaId!;
        _lastDoubleClickAt = now;
        BeginSession("edit", null, null, capturedSelection: selected);
        return true;
    }

    private void ClearNativeOleTarget()
    {
        lock (_nativeOleTargetGate)
        {
            _nativeOleTargetActive = false;
            _nativeOleTargetLeft = 0;
            _nativeOleTargetTop = 0;
            _nativeOleTargetRight = 0;
            _nativeOleTargetBottom = 0;
        }
    }

    private void BeginSession(
        string mode,
        string? displayMode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection = null,
        bool conversionOnly = false)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(
            mode,
            displayMode,
            requestedObjectMode,
            capturedSelection,
            conversionOnly,
            lifetime.Token);
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedDisplayMode,
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

        string? sessionId = null;
        string? imagePath = null;
        string? svgPath = null;
        string? emfPath = null;
        string? mathMl = null;
        try
        {
            var dispatcher = _dispatcher ?? throw new InvalidOperationException("Word dispatcher is unavailable.");
            var service = _formulaService ?? throw new InvalidOperationException("Word formula service is unavailable.");
            var client = _sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable.");
            SetStatus("正在连接 VisualTeX 本地服务…");
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            var selection = capturedSelection?.Metadata is not null
                ? capturedSelection
                : await dispatcher.InvokeAsync(service.ReadSelection).ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

            // A create command may be invoked while the previous formula is
            // still selected. Only edit commands are allowed to seed the new
            // Session from that selection; every create Session starts blank.
            var metadata = mode == "edit" ? selection.Metadata : null;
            var targetObjectMode = requestedObjectMode
                ?? (mode == "create" ? FormulaOleContract.NativeOleMode : selection.ObjectMode)
                ?? FormulaOleContract.NativeOleMode;
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
                Host = "word",
                FormulaId = metadata?.FormulaId,
                SourceDocumentId = selection.DocumentId,
                SourceObjectId = mode == "edit" ? selection.ObjectId : null,
                Title = metadata?.Title ?? "Word Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = requestedDisplayMode ?? metadata?.DisplayMode ?? "inline",
                ObjectMode = targetObjectMode,
                Numbered = (requestedDisplayMode ?? metadata?.DisplayMode) == "block"
                    && (metadata?.Numbered ?? false),
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
                SetStatus("正在直接转换 Word 公式格式…");
            }
            else
            {
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
                SetStatus("已取消，Word 文档未修改。");
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

            var export = session.ExportResult
                ?? throw new InvalidOperationException("VisualTeX Session has no export result.");
            if (string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                var requiredMathMl = export.MathMl;
                if (string.IsNullOrWhiteSpace(requiredMathMl)
                    || !requiredMathMl!.TrimStart().StartsWith("<math", StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "VisualTeX Session has no valid MathML result for Word OMML.");
                mathMl = requiredMathMl;
            }
            else
            {
                imagePath = client.MaterializePng(session);
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    svgPath = client.MaterializeSvg(session);
                    emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                        svgPath,
                        export.Width,
                        export.Height);
                }
            }
            await dispatcher.InvokeAsync(() =>
            {
                var current = service.ReadSelection();
                if (!string.Equals(
                        current.DocumentId,
                        session.SourceDocumentId,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("活动 Word 文档已切换，未写入公式。");
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    if (mathMl is null)
                        throw new InvalidOperationException(
                            "VisualTeX Word OMML MathML payload is unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceOmml(session, mathMl)
                        : service.InsertOmml(session, mathMl);
                }
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    if (emfPath is null || imagePath is null)
                        throw new InvalidOperationException(
                            "VisualTeX native OLE previews are unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceOle(session, imagePath, emfPath)
                        : service.InsertOle(session, imagePath, emfPath);
                }
                if (imagePath is null)
                    throw new InvalidOperationException(
                        "VisualTeX picture preview is unavailable.");
                return session.Mode == "edit"
                    ? service.Replace(session, imagePath)
                    : service.Insert(session, imagePath);
            }).ConfigureAwait(false);
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            if (requiresObjectModeChange
                && string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                SetStatus("已转换为 Word 原生 OMML：可在 Word 中直接编辑，也可继续用 VisualTeX 编辑。");
            else if (requiresObjectModeChange
                && string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.NativeOleMode,
                    StringComparison.Ordinal))
                SetStatus("已转换为原生 OLE：可双击使用 VisualTeX 编辑，并随 Word 文档保存。");
            else
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

    private async Task UpdateEquationNumbersAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            var count = await dispatcher.InvokeAsync(service.UpdateEquationNumbers)
                .ConfigureAwait(false);
            SetStatus($"已更新 {count} 个 Word 公式编号。");
        }
        catch (Exception error)
        {
            SetStatus($"更新 Word 公式编号失败：{error.Message}");
        }
    }

    private async Task InsertEquationReferenceAsync()
    {
        var dispatcher = _dispatcher;
        var application = _application;
        if (dispatcher is null || application is null) return;
        try
        {
            var inserted = await dispatcher.InvokeAsync(() =>
            {
                Document? document = null;
                Selection? selection = null;
                Window? window = null;
                try
                {
                    document = application.ActiveDocument;
                    selection = application.Selection;
                    if (document.ReadOnly)
                        throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
                    var targets = WordEquationNumbering.GetEquationReferenceTargets(document);
                    if (targets.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "当前文档没有带编号的 VisualTeX 行间公式。请先插入行间公式并勾选“添加公式编号”。",
                            "VisualTeX",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return string.Empty;
                    }

                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        var requestedIndex = 0;
                        _ = int.TryParse(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX"),
                            out requestedIndex);
                        requestedIndex = Math.Max(0, Math.Min(targets.Count - 1, requestedIndex));
                        var target = targets[requestedIndex];
                        WordEquationNumbering.InsertEquationReference(
                            document,
                            selection,
                            target,
                            EquationReferenceStyle.Parenthesized);
                        return target.NumberText;
                    }

                    using var dialog = new EquationReferenceDialog(targets);
                    System.Windows.Forms.DialogResult result;
                    try
                    {
                        window = application.ActiveWindow;
                        result = dialog.ShowDialog(new NativeWindowOwner(new IntPtr(window.Hwnd)));
                    }
                    catch
                    {
                        result = dialog.ShowDialog();
                    }
                    if (result != System.Windows.Forms.DialogResult.OK
                        || dialog.SelectedTarget is null)
                        return string.Empty;
                    WordEquationNumbering.InsertEquationReference(
                        document,
                        selection,
                        dialog.SelectedTarget,
                        dialog.SelectedStyle);
                    return dialog.SelectedTarget.NumberText;
                }
                finally
                {
                    ReleaseComObject(window);
                    ReleaseComObject(selection);
                    ReleaseComObject(document);
                }
            }).ConfigureAwait(false);
            if (!string.IsNullOrWhiteSpace(inserted))
                SetStatus($"已插入公式 ({inserted}) 的交叉引用；更新编号时引用会同步刷新。");
        }
        catch (Exception error)
        {
            SetStatus($"插入公式引用失败：{error.Message}");
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
            SetStatus("Word OLE 公式已导出为跨平台图片。");
        }
        catch (Exception error)
        {
            SetStatus($"导出 Word OLE 公式失败：{error.Message}");
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
            SetStatus("Word 公式已删除。");
        }
        catch (Exception error)
        {
            SetStatus($"删除 Word 公式失败：{error.Message}");
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
            try { _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick; } catch { }
            try { _application.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
        }
        try { _doubleClickHook?.Dispose(); } catch { }
        _doubleClickHook = null;
        ClearNativeOleTarget();
        _sessionClient?.Dispose();
        _dispatcher?.Dispose();
        _lifetime?.Dispose();
        _sessionClient = null;
        _dispatcher = null;
        _formulaService = null;
        _lifetime = null;
        Volatile.Write(ref _activeSessionId, null);
        _ribbonUi = null;
        _application = null;
    }
}

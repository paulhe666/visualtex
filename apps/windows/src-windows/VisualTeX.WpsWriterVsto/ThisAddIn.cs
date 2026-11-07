using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Application = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WpsWriterVsto;

[ComVisible(true)]
[Guid("2B7C9D1E-4F5A-4B6C-8D9E-0A1B2C3D4E5F")]
[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
public interface IWpsRibbonCallbacks
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
    void OnUpdateEquationNumbers(object control);

    [DispId(6)]
    bool GetEquationNumberFormatPressed(object control);

    [DispId(7)]
    void OnEquationNumberFormatChanged(object control, bool pressed);

    [DispId(8)]
    object? GetRibbonImage(object control);
}

[ComVisible(true)]
[Guid("5D6E7F80-91A2-4B3C-8D4E-5F60718293A4")]
[ProgId("VisualTeX.WpsWriterVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IWpsRibbonCallbacks))]
public sealed class ThisAddIn : IDTExtensibility2, IRibbonExtensibility, IWpsRibbonCallbacks
{
    private const string RibbonXml = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WpsWriterVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WpsWriterVsto.FormulaGroup" label="VisualTeX 公式">
          <button id="VisualTeX.WpsWriterVsto.Inline" label="MathType 行内公式" size="large" tag="oleInline" getImage="GetRibbonImage" onAction="OnInsertInline" screentip="插入 MathType OLE 行内公式" supertip="在光标处插入可双击编辑的 MathType Equation.DSMT4 行内公式。" />
          <button id="VisualTeX.WpsWriterVsto.Display" label="MathType 行间公式" size="large" tag="oleDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplay" screentip="插入 MathType OLE 行间公式" supertip="在光标处插入可双击编辑的 MathType Equation.DSMT4 行间公式。" />
          <button id="VisualTeX.WpsWriterVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" screentip="编辑所选 MathType 公式" supertip="读取所选 MathType 公式并在 VisualTeX 中重新编辑。" />
        </group>
        <group id="VisualTeX.WpsWriterVsto.NumberingGroup" label="公式编号">
          <button id="VisualTeX.WpsWriterVsto.UpdateNumbers" label="更新公式编号" size="large" tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationNumbers" screentip="更新公式编号" supertip="刷新当前文档中由 VisualTeX 插入的公式编号。" />
          <menu id="VisualTeX.WpsWriterVsto.NumberFormat" label="编号格式" screentip="设置公式编号格式" supertip="选择后应用于后续插入的行间公式，并立即更新当前文档中的公式编号。">
            <toggleButton id="VisualTeX.WpsWriterVsto.NumberFormatOff" label="不自动编号" tag="off" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WpsWriterVsto.NumberFormatContinuous" label="全文连续编号（1）" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WpsWriterVsto.NumberFormatHeading1Dot" label="按章编号（1.1）" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WpsWriterVsto.NumberFormatHeading1Dash" label="按章编号（1-1）" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WpsWriterVsto.NumberFormatHeading2Dot" label="按节编号（1.1.1）" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            <toggleButton id="VisualTeX.WpsWriterVsto.NumberFormatHeading2Dash" label="按节编号（1.1-1）" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
          </menu>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    static ThisAddIn() => VstoDependencyResolver.Install();

    private Application? _application;
    private OfficeUiDispatcher? _dispatcher;
    private VisualTeXSessionClient? _sessionClient;
    private WpsWriterFormulaService? _service;
    private CancellationTokenSource? _lifetime;
    private object? _ribbonUi;
    private string _lastDoubleClickIdentity = string.Empty;
    private DateTimeOffset _lastDoubleClickAt;

    public string GetCustomUI(string ribbonId) => RibbonXml;

    public void OnConnection(
        object application,
        ext_ConnectMode connectMode,
        object addInInstance,
        ref Array custom)
    {
        if (application is not Application word)
        {
            WpsLog.Write("OnConnection ignored: application object is not Word-compatible.");
            return;
        }

        try
        {
            _application = word;
            _lifetime = new CancellationTokenSource();
            _dispatcher = new OfficeUiDispatcher();
            _sessionClient = new VisualTeXSessionClient();
            _service = new WpsWriterFormulaService(word, _sessionClient, _dispatcher);

            try
            {
                _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
            }
            catch (Exception error)
            {
                WpsLog.WriteException("WindowBeforeDoubleClick subscription failed", error);
            }

            WpsLog.Write(
                $"Connected to {_application.Name} {_application.Version} under WPS registration.");
        }
        catch (Exception error)
        {
            WpsLog.WriteException("OnConnection failed", error);
            try
            {
                MessageBox.Show(
                    OfficePluginLanguage.SafeErrorMessage(error.Message),
                    "VisualTeX",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch { }
        }
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => DisposeAddIn();

    public void OnBeginShutdown(ref Array custom) => DisposeAddIn();

    public void OnAddInsUpdate(ref Array custom) { }

    public void OnStartupComplete(ref Array custom) { }

    private void DisposeAddIn()
    {
        var lifetime = Interlocked.Exchange(ref _lifetime, null);
        try { lifetime?.Cancel(); } catch { }
        try { lifetime?.Dispose(); } catch { }

        if (_application is not null)
        {
            try { _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick; } catch { }
        }

        try { _service?.Dispose(); } catch { }
        try { _dispatcher?.Dispose(); } catch { }
        _sessionClient = null;
        _service = null;
        _dispatcher = null;
        _application = null;
    }

    public void OnRibbonLoad(object ribbonUi)
    {
        _ribbonUi = ribbonUi;
        WpsLog.Write("Ribbon loaded.");
    }

    public void OnInsertInline(object control) =>
        StartSession(displayMode: "inline", editSelected: false);

    public void OnInsertDisplay(object control) =>
        StartSession(displayMode: "block", editSelected: false);

    public void OnEditSelected(object control) =>
        StartSession(displayMode: null, editSelected: true);

    private void StartSession(string? displayMode, bool editSelected)
    {
        var service = _service;
        var lifetime = _lifetime;
        if (service is null || lifetime is null)
        {
            MessageBox.Show("VisualTeX 尚未完成初始化。", "VisualTeX");
            return;
        }

        _ = Task.Run(async () =>
        {
            try
            {
                if (editSelected)
                    await service.EditSelectedAsync(lifetime.Token).ConfigureAwait(false);
                else
                    await service.CreateAsync(displayMode ?? "block", lifetime.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException)
            {
                WpsLog.Write("Session operation cancelled.");
            }
            catch (Exception error)
            {
                WpsLog.WriteException("Formula session failed", error);
                ShowError(OfficePluginLanguage.SafeErrorMessage(error.Message));
            }
        }, lifetime.Token);
    }

    public void OnUpdateEquationNumbers(object control)
    {
        try
        {
            _service?.UpdateEquationNumbers();
        }
        catch (Exception error)
        {
            WpsLog.WriteException("UpdateEquationNumbers failed", error);
            ShowError(OfficePluginLanguage.SafeErrorMessage(error.Message));
        }
    }

    public bool GetEquationNumberFormatPressed(object control)
    {
        var tag = ReadControlTag(control);
        try
        {
            return string.Equals(
                _service?.ResolveNumberFormatId(),
                tag,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    public void OnEquationNumberFormatChanged(object control, bool pressed)
    {
        if (!pressed) return;
        try
        {
            _service?.SetNumberFormat(ReadControlTag(control));
            _service?.UpdateEquationNumbers();
        }
        catch (Exception error)
        {
            WpsLog.WriteException("SetNumberFormat failed", error);
            ShowError(OfficePluginLanguage.SafeErrorMessage(error.Message));
        }
    }

    public object? GetRibbonImage(object control) =>
        RibbonIconProvider.GetImage(ReadControlTag(control));

    private static string ReadControlTag(object control)
    {
        if (control is null) return string.Empty;
        try
        {
            // The callback parameter is an IDispatch RCW, not an IRibbonControl
            // typed proxy, so read the property late-bound.
            return control.GetType().InvokeMember(
                "Tag",
                System.Reflection.BindingFlags.GetProperty,
                null,
                control,
                null) as string ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private void OnWindowBeforeDoubleClick(Selection selection, ref bool cancel)
    {
        try
        {
            if (_service is null || _lifetime is null) return;
            if (!_service.IsMathTypeSelection(selection)) return;

            var identity = _service.DescribeSelection(selection);
            var now = DateTimeOffset.UtcNow;
            if (string.Equals(identity, _lastDoubleClickIdentity, StringComparison.Ordinal)
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1.5))
                return;

            _lastDoubleClickIdentity = identity;
            _lastDoubleClickAt = now;
            cancel = true;
            StartSession(displayMode: null, editSelected: true);
        }
        catch (Exception error)
        {
            WpsLog.WriteException("Double-click routing failed", error);
        }
    }

    private static void ShowError(string message)
    {
        try
        {
            MessageBox.Show(message, "VisualTeX", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        catch { }
    }
}

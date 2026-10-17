using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Extensibility;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;

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

    [DispId(15)]
    void OnNewSvgFormula(object control);

    [DispId(16)]
    void OnNewOmmlFormula(object control);

    [DispId(17)]
    void OnNewOleFormula(object control);

    [DispId(18)]
    void OnInsertInlineOmml(object control);

    [DispId(19)]
    void OnInsertDisplayOmml(object control);

    [DispId(20)]
    void OnNewMathTypeFormula(object control);

    [DispId(21)]
    void OnConvertFormulaFormat(object control);

    [DispId(22)]
    bool GetVisualTeXTargetFormPressed(Office.IRibbonControl control);

    [DispId(23)]
    void OnVisualTeXTargetFormChanged(Office.IRibbonControl control, bool pressed);

    [DispId(24)]
    string GetVisualTeXTargetFormLabel(Office.IRibbonControl control);
}

[ComVisible(true)]
[Guid("7E586D2D-57B0-4D14-AB24-EBA9021A5E6D")]
[ProgId("VisualTeX.PowerPointVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IPowerPointRibbonCallbacks))]
public sealed partial class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IPowerPointRibbonCallbacks
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
          <button id="VisualTeX.PowerPointVsto.New" label="新建 SVG" size="large" tag="insertFormula" getImage="GetRibbonImage" onAction="OnNewSvgFormula" />
          <button id="VisualTeX.PowerPointVsto.NewOle" label="新建 OLE" size="large" tag="convertToOle" getImage="GetRibbonImage" onAction="OnNewOleFormula" />
          <button id="VisualTeX.PowerPointVsto.NewMathType" label="新建 MathType" size="large" tag="convertToOle" getImage="GetRibbonImage" onAction="OnNewMathTypeFormula" />
          <button id="VisualTeX.PowerPointVsto.InsertInlineOmml" label="插入行内 OMML 公式" size="large" screentip="插入行内 PowerPoint 原生公式" supertip="在当前文本框的光标或文本选区处插入行内原生 Office Math。" imageMso="EquationInsertNew" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.PowerPointVsto.InsertDisplayOmml" label="插入行间 OMML 公式" size="large" screentip="插入行间 PowerPoint 原生公式" supertip="若正在编辑文本框，则在当前光标或文本选区处插入独立公式段落；否则在当前幻灯片创建独立原生 Office Math。" imageMso="EquationInsertNew" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="公式格式转换">
          <box id="VisualTeX.PowerPointVsto.FormatConversionLabelsA" boxStyle="vertical">
            <labelControl id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeLabel" label="VisualTeX → MathType" />
            <labelControl id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXLabel" label="MathType → VisualTeX" />
            <labelControl id="VisualTeX.PowerPointVsto.OmmlToMathTypeLabel" label="OMML → MathType" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSelectionA" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeSelection" label="所选" screentip="VisualTeX → MathType（所选）" supertip="转换所选对象中的 VisualTeX 公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部 VisualTeX 公式，结果为 MathType。" tag="visualtex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXSelection" label="所选" screentip="MathType → VisualTeX（所选）" supertip="转换所选对象中的 MathType 公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部 MathType 公式，结果为 VisualTeX。" tag="mathTypeOle|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToMathTypeSelection" label="所选" screentip="OMML → MathType（所选）" supertip="转换所选对象中的 PowerPoint 原生（OMML） 公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部 PowerPoint 原生（OMML） 公式，结果为 MathType。" tag="wordOmml|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSlideA" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeSlide" label="本页" screentip="VisualTeX → MathType（本页）" supertip="把当前幻灯片上的全部 VisualTeX 公式转换为 MathType。" tag="visualtex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXSlide" label="本页" screentip="MathType → VisualTeX（本页）" supertip="把当前幻灯片上的全部 MathType 公式转换为 VisualTeX。" tag="mathTypeOle|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToMathTypeSlide" label="本页" screentip="OMML → MathType（本页）" supertip="把当前幻灯片上的全部 PowerPoint 原生（OMML） 公式转换为 MathType。" tag="wordOmml|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionAllA" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeAll" label="全部" screentip="VisualTeX → MathType（全部）" supertip="把整个演示文稿中的全部 VisualTeX 公式转换为 MathType；开始前会确认。" tag="visualtex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXAll" label="全部" screentip="MathType → VisualTeX（全部）" supertip="把整个演示文稿中的全部 MathType 公式转换为 VisualTeX；开始前会确认。" tag="mathTypeOle|visualtex|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToMathTypeAll" label="全部" screentip="OMML → MathType（全部）" supertip="把整个演示文稿中的全部 PowerPoint 原生（OMML） 公式转换为 MathType；开始前会确认。" tag="wordOmml|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
          </box>
          <separator id="VisualTeX.PowerPointVsto.FormatConversionSeparator1" />
          <box id="VisualTeX.PowerPointVsto.FormatConversionLabelsB" boxStyle="vertical">
            <labelControl id="VisualTeX.PowerPointVsto.MathTypeToOmmlLabel" label="MathType → OMML" />
            <labelControl id="VisualTeX.PowerPointVsto.VisualTeXToOmmlLabel" label="VisualTeX → OMML" />
            <labelControl id="VisualTeX.PowerPointVsto.OmmlToVisualTeXLabel" label="OMML → VisualTeX" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSelectionB" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.MathTypeToOmmlSelection" label="所选" screentip="MathType → OMML（所选）" supertip="转换所选对象中的 MathType 公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部 MathType 公式，结果为 OMML。" tag="mathTypeOle|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToOmmlSelection" label="所选" screentip="VisualTeX → OMML（所选）" supertip="转换所选对象中的 VisualTeX 公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部 VisualTeX 公式，结果为 OMML。" tag="visualtex|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToVisualTeXSelection" label="所选" screentip="OMML → VisualTeX（所选）" supertip="转换所选对象中的 PowerPoint 原生（OMML） 公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部 PowerPoint 原生（OMML） 公式，结果为 VisualTeX。" tag="wordOmml|visualtex|selection" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSlideB" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.MathTypeToOmmlSlide" label="本页" screentip="MathType → OMML（本页）" supertip="把当前幻灯片上的全部 MathType 公式转换为 OMML。" tag="mathTypeOle|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToOmmlSlide" label="本页" screentip="VisualTeX → OMML（本页）" supertip="把当前幻灯片上的全部 VisualTeX 公式转换为 OMML。" tag="visualtex|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToVisualTeXSlide" label="本页" screentip="OMML → VisualTeX（本页）" supertip="把当前幻灯片上的全部 PowerPoint 原生（OMML） 公式转换为 VisualTeX。" tag="wordOmml|visualtex|slide" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionAllB" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.MathTypeToOmmlAll" label="全部" screentip="MathType → OMML（全部）" supertip="把整个演示文稿中的全部 MathType 公式转换为 OMML；开始前会确认。" tag="mathTypeOle|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToOmmlAll" label="全部" screentip="VisualTeX → OMML（全部）" supertip="把整个演示文稿中的全部 VisualTeX 公式转换为 OMML；开始前会确认。" tag="visualtex|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToVisualTeXAll" label="全部" screentip="OMML → VisualTeX（全部）" supertip="把整个演示文稿中的全部 PowerPoint 原生（OMML） 公式转换为 VisualTeX；开始前会确认。" tag="wordOmml|visualtex|presentation" onAction="OnConvertFormulaFormat" />
          </box>
          <separator id="VisualTeX.PowerPointVsto.FormatConversionSeparator2" />
          <menu id="VisualTeX.PowerPointVsto.UnifyFormat" label="一键 统一" size="large" screentip="把所有公式统一为同一种格式" supertip="把当前幻灯片或整个演示文稿中其他格式的公式全部转换为所选格式，已是该格式的公式保持不变。" tag="batchImport" getImage="GetRibbonImage">
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifySlideHeader" title="当前幻灯片" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="本页公式全部转为 MathType" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="本页公式全部转为 OMML" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="本页公式全部转为 VisualTeX" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifyAllHeader" title="整个演示文稿" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="全部公式转为 MathType" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="全部公式转为 OMML" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="全部公式转为 VisualTeX" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
          <menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" size="large" screentip="转为 VisualTeX 时使用的形式" supertip="选择“→ VisualTeX”和“统一为 VisualTeX”生成的对象：OLE 对象可双击编辑并嵌入文件；SVG 图片在任何设备上都能清晰显示，同样可双击编辑。" tag="convertToOle" getImage="GetRibbonImage">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE 对象（双击编辑，嵌入文件）" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG 图片（任何设备都能显示）" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu>
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
    private OfficeFormulaShortcutKeyboardHook? _keyboardHook;
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private CancellationTokenSource? _lifetime;
    private string _lastDoubleClickFormulaId = string.Empty;
    private DateTimeOffset _lastDoubleClickAt;
    // The formula the double-click hook claimed, handed to its callback.
    private string? _claimedOleReference;
    private bool _claimedOleIsMathType;
    private string? _activeSessionId;
    private int _formulaFontInvalidationPending;
    private object? _ribbonUi;
    private Office.COMAddIn? _comAddIn;

    public string DiagnosticLastError { get; private set; } = string.Empty;

    public string GetCustomUI(string ribbonId) =>
        OfficePluginLanguage.IsEnglish ? RibbonXmlEnglish : RibbonXml;

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
        string? doubleClickError = null;
        try
        {
            _doubleClickHook = new PowerPointDoubleClickHook(ClaimOleDoubleClick, OnNativeDoubleClick);
            _doubleClickHook.Start();
        }
        catch (Exception error)
        {
            try { _doubleClickHook?.Dispose(); } catch { }
            _doubleClickHook = null;
            doubleClickError = error.Message;
        }
        string? keyboardError = null;
        try
        {
            _keyboardHook = new OfficeFormulaShortcutKeyboardHook(
                OfficeFormulaShortcutHost.PowerPoint,
                OnPowerPointFormulaShortcut);
            try { _keyboardHook.BindOwnerWindow(_application.HWND); } catch { }
            _keyboardHook.Start();
        }
        catch (Exception error)
        {
            try { _keyboardHook?.Dispose(); } catch { }
            _keyboardHook = null;
            keyboardError = error.Message;
        }
        SetStatus(doubleClickError is not null
            ? $"VisualTeX 已就绪，但双击监听不可用：{doubleClickError}"
            : keyboardError is not null
                ? $"VisualTeX 已就绪，但公式快捷键不可用：{keyboardError}"
                : "VisualTeX PowerPoint VSTO 已就绪。");
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
                ? FormulaFontSize.FormatDisplay(size.Value, OfficePluginLanguage.IsEnglish)
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
                ?? throw new InvalidOperationException(T("请先选择一个公式。", "Select a formula first."));
            ApplyFormulaFontSize(FormulaFontSize.PreviousPreset(current));
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
        }
    }
    public void OnIncreaseFormulaFontSize(Office.IRibbonControl control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException(T("请先选择一个公式。", "Select a formula first."));
            ApplyFormulaFontSize(FormulaFontSize.NextPreset(current));
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
        }
    }
    public void OnNewFormula(object control) => OnNewSvgFormula(control);
    public void OnNewSvgFormula(object control) =>
        BeginSession("create", "crossPlatformPicture", null);
    public void OnNewOmmlFormula(object control) =>
        BeginContextAwareOmmlSession();
    public void OnNewOleFormula(object control) =>
        BeginSession("create", "nativeOle", null);
    public void OnNewMathTypeFormula(object control) =>
        BeginSession("create", FormulaOleContract.MathTypeOleMode, null);
    public void OnInsertInlineOmml(object control) =>
        BeginOmmlTextInsertion("inline");
    public void OnInsertDisplayOmml(object control) =>
        BeginOmmlTextInsertion("block");
    private void OnPowerPointFormulaShortcut(OfficeFormulaShortcutCommand command)
    {
        switch (command.Kind)
        {
            case OfficeFormulaShortcutKind.PowerPointSvg:
                OnNewSvgFormula(this);
                break;
            case OfficeFormulaShortcutKind.PowerPointOle:
                OnNewOleFormula(this);
                break;
            case OfficeFormulaShortcutKind.PowerPointInlineOmml:
                OnInsertInlineOmml(this);
                break;
            case OfficeFormulaShortcutKind.PowerPointDisplayOmml:
                OnInsertDisplayOmml(this);
                break;
        }
    }
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
            SetStatus(T("VisualTeX 已打开。", "VisualTeX is open."));
        }
        catch (Exception error)
        {
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "无法打开 VisualTeX",
                "Unable to open VisualTeX"));
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

    private static double ParseFontSize(string value) =>
        FormulaFontSize.Parse(value, OfficePluginLanguage.IsEnglish);

    private void ApplyFormulaFontSize(double value)
    {
        try
        {
            var applied = (_formulaService
                    ?? throw new InvalidOperationException("PowerPoint formula service is unavailable."))
                .SetSelectedFormulaFontSize(value);
            SetStatus(T(
                $"公式字号已设置为 {FormulaFontSize.Describe(applied)}。",
                $"Formula font size set to {FormulaFontSize.Describe(applied, english: true)}."));
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
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
        // A notification is not an insertion command. In particular None also
        // means explicit deselection, not proof of a transient Ribbon handoff.
        // Capture text/source only at the command; never hydrate or cache here.
    }

    // Office UI thread, inside the mouse hook, when PowerPoint retrieves the
    // second click (the first one is already processed). A VisualTeX OLE formula
    // always opens VisualTeX; letting PowerPoint activate it instead starts the
    // OLE server concurrently. A MathType equation is edited in VisualTeX when
    // the user chose that, and always when MathType itself is not installed.
    private bool ClaimOleDoubleClick(int screenX, int screenY)
    {
        _claimedOleReference = null;
        var service = _formulaService;
        if (service is null
            || !service.TryHitOleFormula(screenX, screenY, out var reference, out var isMathType))
            return false;
        if (isMathType && !MathTypeDoubleClickPreference.VisualTeXEditsMathType())
            return false;
        _claimedOleReference = reference;
        _claimedOleIsMathType = isMathType;
        return true;
    }

    private void OnNativeDoubleClick(bool claimedOle, int screenX, int screenY)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        if (claimedOle)
        {
            var claimed = _claimedOleReference;
            var isMathType = _claimedOleIsMathType;
            _claimedOleReference = null;
            dispatcher.Post(() => OpenOleFormula(service, claimed, isMathType));
            return;
        }
        dispatcher.Post(() =>
        {
            OfficeSelection? selected;
            try { selected = service.ReadFormulaAtScreenPoint(screenX, screenY); }
            catch (Exception error)
            {
                PowerPointDoubleClickHook.TraceMessage($"hit-test-error {error.GetType().Name}: {error.Message}");
                return;
            }
            PowerPointDoubleClickHook.TraceMessage(
                $"hit-test mode={selected?.ObjectMode ?? "none"} metadata={selected?.Metadata is not null}");
            if (selected is null || (selected.Metadata is null
                && !PowerPointFormulaService.IsNativeOmmlReference(selected.ObjectId)))
                return;
            var identity = selected.FormulaId ?? selected.ObjectId!;
            var now = DateTimeOffset.UtcNow;
            if (identity == _lastDoubleClickFormulaId
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
                return;
            _lastDoubleClickFormulaId = identity;
            _lastDoubleClickAt = now;
            BeginSession("edit", null, selected);
        });
    }

    private void OpenOleFormula(PowerPointFormulaService service, string? reference, bool isMathType)
    {
        try
        {
            var selected = service.ReadSelection();
            var expectedMode = isMathType
                ? FormulaOleContract.MathTypeOleMode
                : FormulaOleContract.NativeOleMode;
            if (selected.ObjectMode != expectedMode
                || !string.Equals(selected.ObjectId, reference, StringComparison.Ordinal)
                || !isMathType && selected.Metadata is null)
            {
                PowerPointDoubleClickHook.TraceMessage(
                    $"claimed-target-changed expected={expectedMode} selected={selected.ObjectMode} sameObject={string.Equals(selected.ObjectId, reference, StringComparison.Ordinal)} metadata={selected.Metadata is not null}");
                return;
            }
            var now = DateTimeOffset.UtcNow;
            var identity = selected.FormulaId ?? selected.ObjectId!;
            if (identity == _lastDoubleClickFormulaId
                && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
                return;
            _lastDoubleClickFormulaId = identity;
            _lastDoubleClickAt = now;
            PowerPointDoubleClickHook.TraceMessage($"open-editor mode={selected.ObjectMode}");
            BeginSession("edit", null, selected);
        }
        catch (Exception error)
        {
            PowerPointDoubleClickHook.TraceMessage($"open-error {error.GetType().Name}: {error.Message}");
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                isMathType ? "无法打开 MathType 公式" : "无法打开 VisualTeX 公式",
                isMathType ? "Unable to open the MathType equation" : "Unable to open the VisualTeX formula"));
        }
    }

    private void BeginContextAwareOmmlSession()
    {
        DiagnosticLastError = string.Empty;
        try
        {
            var service = _formulaService
                ?? throw new InvalidOperationException(
                    "PowerPoint formula service is unavailable.");
            var target = ResolveOmmlTextInsertionTarget(service);
            if (target is not null)
            {
                BeginSession(
                    "create",
                    "wordOmml",
                    target,
                    requestedDisplayMode: "inline");
                return;
            }

            BeginSession(
                "create",
                "wordOmml",
                null,
                requestedDisplayMode: "block");
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX PowerPoint 写入失败",
                "VisualTeX PowerPoint formula write failed"));
        }
    }

    private void BeginOmmlTextInsertion(string displayMode)
    {
        DiagnosticLastError = string.Empty;
        try
        {
            var service = _formulaService
                ?? throw new InvalidOperationException("PowerPoint formula service is unavailable.");
            var target = ResolveOmmlTextInsertionTarget(service);
            if (target is null
                && !string.Equals(
                    displayMode,
                    "block",
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "插入行内 OMML 公式前，请先把文本光标或文本选区放到一个现有文本框中。");
            BeginSession(
                "create",
                "wordOmml",
                target,
                requestedDisplayMode: displayMode);
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX PowerPoint 写入失败",
                "VisualTeX PowerPoint formula write failed"));
        }
    }

    private OfficeSelection? ResolveOmmlTextInsertionTarget(
        PowerPointFormulaService service)
    {
        return service.TryCaptureTextInsertionTarget();
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
            if (selection.Metadata is null
                && !PowerPointFormulaService.IsNativeOmmlReference(selection.ObjectId)
                && selection.ObjectMode != FormulaOleContract.MathTypeOleMode)
                throw new InvalidOperationException("没有选中公式。请先选中一个 VisualTeX 公式、PowerPoint 原生公式或 MathType 公式。");
            BeginSession("edit", requestedObjectMode, selection, conversionOnly);
        }
        catch (Exception error)
        {
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX PowerPoint 写入失败",
                "VisualTeX PowerPoint formula write failed"));
        }
    }

    private void BeginSession(
        string mode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
        bool conversionOnly = false,
        string? requestedDisplayMode = null)
    {
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        _ = RunSessionAsync(
            mode,
            requestedObjectMode,
            capturedSelection,
            conversionOnly,
            requestedDisplayMode,
            lifetime.Token);
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
        bool conversionOnly,
        string? requestedDisplayMode,
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
        PowerPointFormulaService.PreparedMathType? preparedMathType = null;
        try
        {
            var dispatcher = _dispatcher ?? throw new InvalidOperationException("PowerPoint dispatcher is unavailable.");
            var service = _formulaService ?? throw new InvalidOperationException("PowerPoint formula service is unavailable.");
            var client = _sessionClient ?? throw new InvalidOperationException("VisualTeX Session client is unavailable.");
            SetStatus("正在连接 VisualTeX 本地服务…");
            await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
            if (conversionOnly)
                await client.PrewarmConverterAsync(cancellationToken).ConfigureAwait(false);
            var selection = capturedSelection
                ?? await dispatcher.InvokeAsync(
                    () => mode == "create" ? service.CaptureSlideInsertionTarget() : ResolveFormulaSelection(service)).ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 PowerPoint 演示文稿为只读状态。");
            // Source snapshots belong to the acquired operation, not to a
            // repeated Ribbon click that only refocuses an existing editor.
            if (mode == "edit")
                selection = await dispatcher.InvokeAsync(() =>
                    PowerPointFormulaService.IsNativeOmmlReference(selection.ObjectId)
                        ? service.LoadNativeOmmlSelection(selection)
                        : selection.ObjectMode == FormulaOleContract.MathTypeOleMode
                            ? service.LoadMathTypeSelection(selection)
                            : service.PrepareManagedSelectionForEdit(selection)).ConfigureAwait(false);
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("没有选中公式。请先选中一个 VisualTeX 公式、PowerPoint 原生公式或 MathType 公式。");

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
            var targetDisplayMode = requestedDisplayMode
                ?? metadata?.DisplayMode
                ?? "block";
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
                SourceObjectId = selection.ObjectId,
                Title = metadata?.Title ?? "PowerPoint Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = targetDisplayMode,
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

                if (string.Equals(
                        targetObjectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    // Hide loading the shared OMML writer (the Word add-in's
                    // assembly) behind the time the user spends editing the
                    // formula.
                    _ = dispatcher.InvokeAsync(() =>
                    {
                        try { PowerPointNativeOmmlKernel.Prewarm(); }
                        catch { }
                        return true;
                    });
                }
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

            if (session.ObjectMode == FormulaOleContract.MathTypeOleMode)
            {
                preparedMathType = await Task.Run(() => service.PrepareMathType(session), cancellationToken)
                    .ConfigureAwait(false);
            }
            else if (session.ObjectMode == "nativeOle")
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
            await dispatcher.InvokeAsync(() =>
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
                if (session.ObjectMode == FormulaOleContract.MathTypeOleMode)
                    return service.ApplyMathType(session, preparedMathType
                        ?? throw new InvalidOperationException("The MathType equation was not prepared."));
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
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            SetStatus(requiresObjectModeChange && session.ObjectMode == "nativeOle"
                ? "已转换为原生 OLE：外观保持不变，可双击编辑，并嵌入 PowerPoint 文件。"
                : requiresObjectModeChange && session.ObjectMode == "wordOmml"
                    ? "已转换为 PowerPoint 原生 Office Math（OMML），可继续使用 PowerPoint 公式工具编辑。"
                    : requiresObjectModeChange && session.ObjectMode == "crossPlatformPicture"
                        ? "已转换为嵌入式 SVG 图片，可跨平台显示并保持矢量清晰度。"
                        : session.Mode == "edit"
                            ? "PowerPoint 公式已更新。"
                            : session.ObjectMode == FormulaOleContract.WordOmmlMode
                                ? "已插入 PowerPoint 原生 OMML；未写入 VisualTeX FormulaId 或元数据。"
                                : "PowerPoint 公式已插入。");
        }
        catch (OperationCanceledException)
        {
            SetStatus("VisualTeX 操作已取消。");
        }
        catch (Exception error)
        {
            var failureMessage = OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX PowerPoint 写入失败",
                "VisualTeX PowerPoint formula write failed");
            if (sessionId is not null && _sessionClient is not null)
            {
                try
                {
                    await _sessionClient.FailAsync(sessionId, failureMessage, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            ReportError(failureMessage);
        }
        finally
        {
            preparedMathType?.Dispose();
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
            ReportError(OfficePluginLanguage.FormatOperationError(
                error,
                "删除 PowerPoint 公式失败",
                "Failed to delete the PowerPoint formula"));
        }
    }

    private void SetStatus(string message)
    {
        var dispatcher = _dispatcher;
        var application = _application;
        if (dispatcher is null || application is null) return;
        var userMessage = OfficePluginLanguage.SafeStatusMessage(message);
        dispatcher.Post(() =>
        {
            try { ((dynamic)application).StatusBar = userMessage; } catch { }
        });
    }

    private OfficeSelection ResolveFormulaSelection(PowerPointFormulaService service)
    {
        // A blank-slide click and a transient Ribbon state can both be
        // ppSelectionNone. Neither authorizes editing a previous object.
        // Capture the current target now; hydrate its source only after the
        // operation gate has been acquired.
        return service.ReadSelection();
    }

    private void ReportError(string message)
    {
        DiagnosticLastError = message;
        var userMessage = OfficePluginLanguage.SafeErrorMessage(message);
        SetStatus(userMessage);
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
                    userMessage,
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
        // Never leave PowerPoint without input if it shuts down mid-batch.
        try { _activeModalGuard?.Dispose(); } catch { }
        _activeModalGuard = null;
        if (_application is not null)
        {
            try { _application.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
        }
        _doubleClickHook?.Dispose();
        _keyboardHook?.Dispose();
        _sessionClient?.Dispose();
        try { _formulaService?.Dispose(); } catch { }
        _dispatcher?.Dispose();
        _lifetime?.Dispose();
        _doubleClickHook = null;
        _keyboardHook = null;
        _sessionClient = null;
        _dispatcher = null;
        _formulaService = null;
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

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
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

    [DispId(15)]
    string GetFormulaFontSizeText(Office.IRibbonControl control);

    [DispId(16)]
    bool GetFormulaFontSizeEnabled(Office.IRibbonControl control);

    [DispId(17)]
    void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value);

    [DispId(18)]
    void OnDecreaseFormulaFontSize(object control);

    [DispId(19)]
    void OnIncreaseFormulaFontSize(object control);

    [DispId(20)]
    void OnBulkImport(object control);

    [DispId(21)]
    void OnRedrawSelectionToOmml(object control);

    [DispId(22)]
    void OnRedrawSelectionToOle(object control);

    [DispId(23)]
    void OnRedrawDocumentToOmml(object control);

    [DispId(24)]
    void OnRedrawDocumentToOle(object control);

    [DispId(25)]
    bool GetEquationNumberFormatPressed(Office.IRibbonControl control);

    [DispId(26)]
    void OnEquationNumberFormatChanged(Office.IRibbonControl control, bool pressed);

    [DispId(27)]
    void OnRedrawSelectionOleToLatex(object control);

    [DispId(28)]
    void OnRedrawSelectionOmmlToLatex(object control);

    [DispId(29)]
    void OnRedrawDocumentOleToLatex(object control);

    [DispId(30)]
    void OnRedrawDocumentOmmlToLatex(object control);

    [DispId(31)]
    void OnConvertVisualTeXToMathTypeSelection(object control);

    [DispId(32)]
    void OnConvertVisualTeXToMathTypeDocument(object control);

    [DispId(33)]
    void OnConvertMathTypeToVisualTeXSelection(object control);

    [DispId(34)]
    void OnConvertMathTypeToVisualTeXDocument(object control);

    [DispId(35)]
    void OnConvertOmmlToMathTypeSelection(object control);

    [DispId(36)]
    void OnConvertOmmlToMathTypeDocument(object control);

    [DispId(37)]
    void OnConvertMathTypeToOmmlSelection(object control);

    [DispId(38)]
    void OnConvertMathTypeToOmmlDocument(object control);

    [DispId(39)]
    void OnRedrawSelectionToMathType(object control);

    [DispId(40)]
    void OnRedrawDocumentToMathType(object control);

    [DispId(41)]
    void OnRedrawSelectionMathTypeToLatex(object control);

    [DispId(42)]
    void OnRedrawDocumentMathTypeToLatex(object control);

    [DispId(43)]
    void OnConvertVisualTeXToOmmlSelection(object control);

    [DispId(44)]
    void OnConvertVisualTeXToOmmlDocument(object control);

    [DispId(45)]
    void OnConvertOmmlToVisualTeXSelection(object control);

    [DispId(46)]
    void OnConvertOmmlToVisualTeXDocument(object control);

}

[ComVisible(true)]
[Guid("F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1")]
[ProgId("VisualTeX.WordVsto")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IWordRibbonCallbacks))]
public sealed partial class ThisAddIn : IDTExtensibility2, Office.IRibbonExtensibility, IWordRibbonCallbacks
{
    private const int AllowAnyProcessToSetForeground = -1;
    private const float MathTypePreviewHorizontalSafetyInsetPixels = 4.0f;

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
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="OLE 行内公式" size="large" tag="oleInline" getImage="GetRibbonImage" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="OLE 行间公式" size="large" tag="oleDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.InlineOmml" label="OMML 行内公式" size="large" screentip="插入 Word 原生公式" supertip="插入可由 Word 原生公式工具直接编辑、同时保留 VisualTeX LaTeX 元数据的 OMML 行内公式。" tag="ommlInline" getImage="GetRibbonImage" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.WordVsto.DisplayOmml" label="OMML 行间公式" size="large" screentip="插入 Word 原生公式" supertip="插入可由 Word 原生公式工具直接编辑、同时保留 VisualTeX LaTeX 元数据的 OMML 行间公式。" tag="ommlDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.WordVsto.Edit" label="编辑所选公式" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <box id="VisualTeX.WordVsto.FormatConversionBox" boxStyle="vertical">
            <menu id="VisualTeX.WordVsto.VisualTeXToMathType" label="VisualTeX → MathType" screentip="重新绘制为 MathType OLE" supertip="删除原 VisualTeX 宿主与旧编号，再用正常新建 MathType 公式的同一条成熟路径重新绘制。">
              <button id="VisualTeX.WordVsto.VisualTeXToMathTypeSelection" label="转换选中部分" onAction="OnConvertVisualTeXToMathTypeSelection" />
              <button id="VisualTeX.WordVsto.VisualTeXToMathTypeDocument" label="全文批量转换" onAction="OnConvertVisualTeXToMathTypeDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.MathTypeToVisualTeX" label="MathType → VisualTeX" screentip="重新绘制为 VisualTeX OLE" supertip="删除原 MathType 宿主与旧编号，再用正常新建 VisualTeX OLE 的同一条成熟路径重新绘制。">
              <button id="VisualTeX.WordVsto.MathTypeToVisualTeXSelection" label="转换选中部分" onAction="OnConvertMathTypeToVisualTeXSelection" />
              <button id="VisualTeX.WordVsto.MathTypeToVisualTeXDocument" label="全文批量转换" onAction="OnConvertMathTypeToVisualTeXDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.OmmlToMathType" label="OMML → MathType" screentip="将 Word 原生公式转换为 MathType OLE" supertip="读取 Word 原生 OMath/OMathPara 的真实 MathML，删除原 OMML 宿主与旧编号，再用 VisualTeX 自包含 MathType Equation.DSMT4 路径原位重绘。">
              <button id="VisualTeX.WordVsto.OmmlToMathTypeSelection" label="转换选中部分" onAction="OnConvertOmmlToMathTypeSelection" />
              <button id="VisualTeX.WordVsto.OmmlToMathTypeDocument" label="全文批量转换" onAction="OnConvertOmmlToMathTypeDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.MathTypeToOmml" label="MathType → OMML" screentip="将 MathType OLE 转换为 Word 原生公式" supertip="直接读取 Equation Native 中的 MathML，不启动 MathType，再用 VisualTeX 现有 Word OMML 插入与编号路径原位重绘。">
              <button id="VisualTeX.WordVsto.MathTypeToOmmlSelection" label="转换选中部分" onAction="OnConvertMathTypeToOmmlSelection" />
              <button id="VisualTeX.WordVsto.MathTypeToOmmlDocument" label="全文批量转换" onAction="OnConvertMathTypeToOmmlDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.VisualTeXToOmml" label="VisualTeX → OMML" screentip="将 VisualTeX OLE 转换为 Word OMML" supertip="使用统一格式转换管线，将选中范围或全文中的 VisualTeX OLE 原位转换为 Word 原生 OMML，并保留编号和公式引用。">
              <button id="VisualTeX.WordVsto.VisualTeXToOmmlSelection" label="转换选中部分" onAction="OnConvertVisualTeXToOmmlSelection" />
              <button id="VisualTeX.WordVsto.VisualTeXToOmmlDocument" label="全文批量转换" onAction="OnConvertVisualTeXToOmmlDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.OmmlToVisualTeX" label="OMML → VisualTeX" screentip="将 Word OMML 转换为 VisualTeX OLE" supertip="使用统一格式转换管线，将选中范围或全文中的 Word OMML 原位转换为 VisualTeX OLE，并保留编号和公式引用。">
              <button id="VisualTeX.WordVsto.OmmlToVisualTeXSelection" label="转换选中部分" onAction="OnConvertOmmlToVisualTeXSelection" />
              <button id="VisualTeX.WordVsto.OmmlToVisualTeXDocument" label="全文批量转换" onAction="OnConvertOmmlToVisualTeXDocument" />
            </menu>
          </box>
          <box id="VisualTeX.WordVsto.NumberingBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.UpdateNumbers" label="更新公式编号" screentip="更新 VisualTeX 与 MathType 公式编号" supertip="刷新当前文档中的 VisualTeX 编号，以及 MathType 原生 MTChap/MTSec/MTEqn 编号和对应公式引用。" tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationNumbers" />
            <menu id="VisualTeX.WordVsto.NumberFormat" label="编号格式" screentip="设置当前文档的公式编号格式" supertip="选择后同步更新当前文档已有的 VisualTeX 编号和 MathType 原生 MTPlaceRef 编号，并应用于后续新插入的带编号公式。">
              <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="全文连续编号（1）" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="按章编号（1.1）" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="按章编号（1-1）" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="按节编号（1.1.1）" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="按节编号（1.1-1）" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            </menu>
            <button id="VisualTeX.WordVsto.InsertReference" label="插入公式引用" screentip="引用带编号公式" supertip="从当前文档的 VisualTeX 或 MathType 带编号公式中选择目标；VisualTeX 使用 Word REF，MathType 保留原生 ZEqnNum/GOTOBUTTON/REF 引用结构。" imageMso="HyperlinkInsert" onAction="OnInsertEquationReference" />
          </box>
          <button id="VisualTeX.WordVsto.BulkImport" label="批量导入" size="large" screentip="批量导入 LaTeX / Markdown" supertip="将 Markdown 或 LaTeX 文档解析为 Word 原生文字，以及可单独编辑和调整字号的行内/行间公式。" tag="batchImport" getImage="GetRibbonImage" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.RedrawGroup" label="LaTeX 重绘">
          <menu id="VisualTeX.WordVsto.RedrawSelection" label="重绘所选" size="large" screentip="重绘所选公式或 LaTeX 代码" supertip="可将所选 LaTeX 代码原位重绘为 VisualTeX OLE、Word OMML 或 MathType，也可将这三种公式恢复为 LaTeX 代码。" tag="batchImport" getImage="GetRibbonImage">
            <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="LaTeX 重绘为 Word OMML" screentip="原位替换为 Word 原生公式" onAction="OnRedrawSelectionToOmml" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="LaTeX 重绘为 VisualTeX OLE" screentip="原位替换为可双击编辑的 VisualTeX OLE" onAction="OnRedrawSelectionToOle" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="LaTeX 重绘为 MathType" screentip="原位替换为 MathType Equation.DSMT4 公式" onAction="OnRedrawSelectionToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawSelectionSeparator" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOleToLatex" label="所选 VisualTeX OLE 转为 LaTeX 代码" screentip="恢复所选 VisualTeX OLE 的 LaTeX 源码" onAction="OnRedrawSelectionOleToLatex" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOmmlToLatex" label="所选 OMML 公式转为 LaTeX 代码" screentip="恢复所选 Word OMML 的 LaTeX 源码" onAction="OnRedrawSelectionOmmlToLatex" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathTypeToLatex" label="所选 MathType 公式转为 LaTeX 代码" screentip="直接读取 MathType Equation Native 并恢复为 LaTeX 源码" onAction="OnRedrawSelectionMathTypeToLatex" />
          </menu>
          <menu id="VisualTeX.WordVsto.RedrawDocument" label="重绘全文" size="large" screentip="重绘全文公式或 LaTeX 代码" supertip="可将全文 LaTeX 代码原位重绘为 VisualTeX OLE、Word OMML 或 MathType，也可将全文这三种公式恢复为 LaTeX 代码；开始前会再次确认。" imageMso="RefreshAll">
            <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="全文 LaTeX 重绘为 Word OMML" onAction="OnRedrawDocumentToOmml" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="全文 LaTeX 重绘为 VisualTeX OLE" onAction="OnRedrawDocumentToOle" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="全文 LaTeX 重绘为 MathType" onAction="OnRedrawDocumentToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawDocumentSeparator" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOleToLatex" label="全文 VisualTeX OLE 转为 LaTeX 代码" onAction="OnRedrawDocumentOleToLatex" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOmmlToLatex" label="全文 OMML 公式转为 LaTeX 代码" onAction="OnRedrawDocumentOmmlToLatex" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathTypeToLatex" label="全文 MathType 公式转为 LaTeX 代码" onAction="OnRedrawDocumentMathTypeToLatex" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="公式字号">
          <button id="VisualTeX.WordVsto.FontSizeDecrease" label="减小" imageMso="FontSizeDecrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
          <comboBox id="VisualTeX.WordVsto.FontSize" label="字号" sizeString="初号（42 磅）" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
            <item id="VisualTeX.WordVsto.FontSizeChu" label="初号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoChu" label="小初" />
            <item id="VisualTeX.WordVsto.FontSizeYi" label="一号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoYi" label="小一" />
            <item id="VisualTeX.WordVsto.FontSizeEr" label="二号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoEr" label="小二" />
            <item id="VisualTeX.WordVsto.FontSizeSan" label="三号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoSan" label="小三" />
            <item id="VisualTeX.WordVsto.FontSizeSi" label="四号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoSi" label="小四" />
            <item id="VisualTeX.WordVsto.FontSizeWu" label="五号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoWu" label="小五" />
            <item id="VisualTeX.WordVsto.FontSizeLiu" label="六号" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoLiu" label="小六" />
            <item id="VisualTeX.WordVsto.FontSizeQi" label="七号" />
            <item id="VisualTeX.WordVsto.FontSizeBa" label="八号" />
            <item id="VisualTeX.WordVsto.FontSize8" label="8" />
            <item id="VisualTeX.WordVsto.FontSize9" label="9" />
            <item id="VisualTeX.WordVsto.FontSize10" label="10" />
            <item id="VisualTeX.WordVsto.FontSize10_5" label="10.5" />
            <item id="VisualTeX.WordVsto.FontSize11" label="11" />
            <item id="VisualTeX.WordVsto.FontSize12" label="12" />
            <item id="VisualTeX.WordVsto.FontSize14" label="14" />
            <item id="VisualTeX.WordVsto.FontSize16" label="16" />
            <item id="VisualTeX.WordVsto.FontSize18" label="18" />
            <item id="VisualTeX.WordVsto.FontSize20" label="20" />
            <item id="VisualTeX.WordVsto.FontSize24" label="24" />
            <item id="VisualTeX.WordVsto.FontSize28" label="28" />
            <item id="VisualTeX.WordVsto.FontSize36" label="36" />
            <item id="VisualTeX.WordVsto.FontSize48" label="48" />
            <item id="VisualTeX.WordVsto.FontSize72" label="72" />
          </comboBox>
          <button id="VisualTeX.WordVsto.FontSizeIncrease" label="增大" imageMso="FontSizeIncrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
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
    private WordCopyPasteHook? _copyPasteHook;
    private readonly object _copyPasteGate = new();
    private WordFormulaService.WordFormulaCopySnapshot? _formulaCopySnapshot;
    private uint _formulaCopyClipboardSequence;
    private int _copyPasteRepairGeneration;
    private int _copyPasteWatchGeneration;
    private int _copyPasteSelectionRepairPending;
    private long _lastSuccessfulPasteRepairStartedAt;
    private long _lastSuccessfulPasteRepairCompletedAt;
    private static readonly object BulkAcceptanceLogGate = new();
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private readonly object _activeSessionOperationGate = new();
    private readonly object _nativeOleTargetGate = new();
    private readonly object _mouseDoubleClickGate = new();
    private CancellationTokenSource? _lifetime;
    private string _lastDoubleClickIdentity = string.Empty;
    private DateTimeOffset _lastDoubleClickAt;
    private string? _activeSessionId;
    private CancellationTokenSource? _activeSessionCancellation;
    private bool _nativeOleTargetActive;
    private bool _nativeOleTargetIsMathType;
    private int _nativeOleTargetLeft;
    private int _nativeOleTargetTop;
    private int _nativeOleTargetRight;
    private int _nativeOleTargetBottom;
    private int _nativeOleTargetRangeStart = -1;
    private int _nativeOleTargetRangeEnd = -1;
    private int _pendingNativeMathTypeRangeStart = -1;
    private int _pendingNativeMathTypeRangeEnd = -1;
    private int _claimedNativeMathTypeRangeStart = -1;
    private int _claimedNativeMathTypeRangeEnd = -1;
    private DateTimeOffset _claimedNativeMathTypeAt;
    private bool _mouseDoubleClickPointActive;
    private int _mouseDoubleClickX;
    private int _mouseDoubleClickY;
    private DateTimeOffset _mouseDoubleClickAt;
    private int _formulaFontInvalidationPending;
    private double _cachedSelectedFormulaFontSize = double.NaN;
    private string? _cachedEquationNumberFormatId;
    private int _lastFormulaRibbonOwnerStart = int.MinValue;
    private int _lastFormulaRibbonOwnerEnd = int.MinValue;
    private int _normalizingTypingCaret;
    private int _typingCaretNormalizationPending;
    private int _typingCaretNormalizationGeneration;
    private int _ommlRedrawUndoCleanupPending;
    private readonly object _ommlRedrawUndoCleanupWatcherGate = new();
    private System.Threading.Timer? _ommlRedrawUndoCleanupWatcher;
    private int _formulaFormatMutationDepth;
    private int _formulaFontReadsDeferredDuringMutation;
    private bool _acceptanceSelectionDiagnostics;
    private int _acceptanceSelectionChangeCount;
    private int _acceptanceFormulaStateReadCount;
    private int _acceptanceDeferredCaretPassCount;
    private int _acceptanceEquationFormatReadCount;
    private object? _ribbonUi;
    private Office.COMAddIn? _comAddIn;
    private bool _mathTypePreviewSessionAcquired;

    public string GetCustomUI(string ribbonId) =>
        OfficePluginLanguage.IsEnglish ? RibbonXmlEnglish : RibbonXml;

    public void OnConnection(
        object application,
        ext_ConnectMode connectMode,
        object addInInstance,
        ref Array custom)
    {
        // Real VSTO-flow acceptances host the current source assembly manually
        // while Word may also auto-load the installed COM add-in into the same
        // application. Keep that installed instance inert only in acceptance so
        // one physical double-click cannot be handled by two event subscribers.
        // Production Word startup never sets this environment variable.
        if (addInInstance is Office.COMAddIn
            && string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            WordDoubleClickHook.TraceMessage(
                "installed-addin-suppressed-for-manual-acceptance");
            return;
        }

        _application = (Application)application;
        var loadedAssembly = typeof(ThisAddIn).Assembly;
        WordDoubleClickHook.TraceMessage(
            $"word-addin-loaded pid={Process.GetCurrentProcess().Id} assembly={loadedAssembly.Location} mvid={loadedAssembly.ManifestModule.ModuleVersionId}");
        _comAddIn = addInInstance as Office.COMAddIn;
        if (_comAddIn is not null)
        {
            try { _comAddIn.Object = this; } catch { }
        }
        var officeMathFontReady = WordOfficeMathFontLoader.TryEnsureLoaded(
            out var officeMathFontError);
        _formulaService = new WordFormulaService(_application);
        _dispatcher = new OfficeUiDispatcher();
        _acceptanceSelectionDiagnostics = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal);
        _sessionClient = new VisualTeXSessionClient();
        _lifetime = new CancellationTokenSource();
        _ = PrewarmCompanionAsync(_sessionClient, _lifetime.Token);
        MathTypeNativePreviewRenderer.AcquireSharedSession();
        _mathTypePreviewSessionAcquired = true;
        _application.WindowBeforeDoubleClick += OnWindowBeforeDoubleClick;
        _application.WindowSelectionChange += OnWindowSelectionChange;
        _application.WindowActivate += OnWindowActivate;
        _application.DocumentOpen += OnDocumentOpen;
        _application.DocumentBeforeSave += OnDocumentBeforeSave;
        string? doubleClickError = null;
        try
        {
            _doubleClickHook = new WordDoubleClickHook(
                ShouldInterceptNativeOleDoubleClick,
                OnNativeWordDoubleClick);
            Window? ownerWindow = null;
            try
            {
                ownerWindow = _application.ActiveWindow;
                if (ownerWindow is not null) _doubleClickHook.BindOwnerWindow(ownerWindow.Hwnd);
            }
            catch { }
            finally { ReleaseComObject(ownerWindow); }
            _doubleClickHook.Start();
        }
        catch (Exception error)
        {
            try { _doubleClickHook?.Dispose(); } catch { }
            _doubleClickHook = null;
            doubleClickError = error.Message;
        }
        try
        {
            _copyPasteHook = new WordCopyPasteHook(OnWordClipboardGesture);
            Window? ownerWindow = null;
            try
            {
                ownerWindow = _application.ActiveWindow;
                if (ownerWindow is not null) _copyPasteHook.BindOwnerWindow(ownerWindow.Hwnd);
            }
            catch { }
            finally { ReleaseComObject(ownerWindow); }
            _copyPasteHook.Start();
        }
        catch (Exception error)
        {
            try { _copyPasteHook?.Dispose(); } catch { }
            _copyPasteHook = null;
            WordDoubleClickHook.TraceMessage(
                $"copy-paste-hook-start-failed {error.GetType().Name}: {error.Message}");
        }
        SetStatus(!officeMathFontReady
            ? $"VisualTeX 已就绪，但 Word 数学字体不可用：{officeMathFontError}"
            : doubleClickError is null
                ? "VisualTeX Word VSTO 已就绪。"
                : $"VisualTeX 已就绪，但 OLE 双击监听不可用：{doubleClickError}");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom)
    {
        Document? document = null;
        try
        {
            document = _application?.ActiveDocument;
            if (document is not null)
                RefreshNumberedOmmlTabLayoutsAfterOpen(document);
        }
        catch { }
        finally { ReleaseComObject(document); }
    }
    public void OnBeginShutdown(ref Array custom) => Dispose();

    public void OnRibbonLoad(object ribbonUi)
    {
        _ribbonUi = ribbonUi;
        Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
        _cachedEquationNumberFormatId = null;
        InvalidateFormulaFontControls();
        InvalidateEquationNumberFormatControls();
    }
    public object? GetRibbonImage(Office.IRibbonControl control) =>
        RibbonIconProvider.GetImage(control?.Tag);
    public string GetFormulaFontSizeText(Office.IRibbonControl control)
    {
        try
        {
            var size = GetCachedSelectedFormulaFontSize();
            return size.HasValue
                ? FormulaFontSize.FormatDisplay(size.Value, OfficePluginLanguage.IsEnglish)
                : string.Empty;
        }
        catch { return string.Empty; }
    }
    public bool GetFormulaFontSizeEnabled(Office.IRibbonControl control)
    {
        try { return GetCachedSelectedFormulaFontSize().HasValue; }
        catch { return false; }
    }
    public void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value) =>
        ApplyFormulaFontSize(ParseFontSize(value));
    public void OnDecreaseFormulaFontSize(object control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException(T("请先选择一个 VisualTeX 公式。", "Select a VisualTeX formula first."));
            ApplyFormulaFontSize(FormulaFontSize.PreviousPreset(current));
        }
        catch (Exception error)
        {
            SetStatus(T($"无法设置公式字号：{error.Message}", $"Unable to set formula font size: {error.Message}"));
        }
    }
    public void OnIncreaseFormulaFontSize(object control)
    {
        try
        {
            var current = _formulaService?.GetSelectedFormulaFontSize()
                ?? throw new InvalidOperationException(T("请先选择一个 VisualTeX 公式。", "Select a VisualTeX formula first."));
            ApplyFormulaFontSize(FormulaFontSize.NextPreset(current));
        }
        catch (Exception error)
        {
            SetStatus(T($"无法设置公式字号：{error.Message}", $"Unable to set formula font size: {error.Message}"));
        }
    }
    public void OnInsertInline(object control) =>
        BeginSession("create", "inline", null);
    public void OnInsertDisplay(object control) =>
        BeginSession("create", "block", null);
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
    public bool GetEquationNumberFormatPressed(Office.IRibbonControl control)
    {
        try
        {
            var current = _cachedEquationNumberFormatId;
            if (string.IsNullOrWhiteSpace(current))
            {
                if (_acceptanceSelectionDiagnostics)
                    Interlocked.Increment(ref _acceptanceEquationFormatReadCount);
                current = _formulaService?.GetEquationNumberFormatId()
                    ?? EquationNumberFormat.ContinuousId;
                _cachedEquationNumberFormatId = current;
            }
            return string.Equals(current, control?.Tag, StringComparison.Ordinal);
        }
        catch { return false; }
    }
    public void OnEquationNumberFormatChanged(
        Office.IRibbonControl control,
        bool pressed)
    {
        if (!pressed)
        {
            InvalidateEquationNumberFormatControls();
            return;
        }
        _ = SetEquationNumberFormatAsync(control?.Tag);
    }
    public void OnBulkImport(object control) => _ = BulkImportAsync();
    public void OnRedrawSelectionToOmml(object control) =>
        _ = RedrawLatexAsync(wholeDocument: false, FormulaOleContract.WordOmmlMode);
    public void OnRedrawSelectionToOle(object control) =>
        _ = RedrawLatexAsync(wholeDocument: false, FormulaOleContract.NativeOleMode);
    public void OnRedrawDocumentToOmml(object control) =>
        _ = RedrawLatexAsync(wholeDocument: true, FormulaOleContract.WordOmmlMode);
    public void OnRedrawDocumentToOle(object control) =>
        _ = RedrawLatexAsync(wholeDocument: true, FormulaOleContract.NativeOleMode);
    public void OnRedrawSelectionToMathType(object control) =>
        _ = RedrawLatexAsync(wholeDocument: false, FormulaOleContract.MathTypeOleMode);
    public void OnRedrawDocumentToMathType(object control) =>
        _ = RedrawLatexAsync(wholeDocument: true, FormulaOleContract.MathTypeOleMode);
    public void OnRedrawSelectionOleToLatex(object control) =>
        _ = ConvertFormulaObjectsToLatexAsync(
            wholeDocument: false,
            FormulaOleContract.NativeOleMode);
    public void OnRedrawSelectionOmmlToLatex(object control) =>
        _ = ConvertFormulaObjectsToLatexAsync(
            wholeDocument: false,
            FormulaOleContract.WordOmmlMode);
    public void OnRedrawDocumentOleToLatex(object control) =>
        _ = ConvertFormulaObjectsToLatexAsync(
            wholeDocument: true,
            FormulaOleContract.NativeOleMode);
    public void OnRedrawDocumentOmmlToLatex(object control) =>
        _ = ConvertFormulaObjectsToLatexAsync(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode);
    public void OnRedrawSelectionMathTypeToLatex(object control) =>
        _ = ConvertFormulaObjectsToLatexAsync(
            wholeDocument: false,
            FormulaOleContract.MathTypeOleMode);
    public void OnRedrawDocumentMathTypeToLatex(object control) =>
        _ = ConvertFormulaObjectsToLatexAsync(
            wholeDocument: true,
            FormulaOleContract.MathTypeOleMode);
    public void OnDeleteSelected(object control) => _ = DeleteSelectedAsync();
    public void OnInsertEquationReference(object control) => _ = InsertEquationReferenceAsync();
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
            SetStatus(T($"无法打开 VisualTeX：{error.Message}", $"Unable to open VisualTeX: {error.Message}"));
        }
    }

    private static double ParseFontSize(string value) =>
        FormulaFontSize.Parse(value, OfficePluginLanguage.IsEnglish);

    private void ApplyFormulaFontSize(double value)
    {
        try
        {
            var applied = (_formulaService
                    ?? throw new InvalidOperationException("Word formula service is unavailable."))
                .SetSelectedFormulaFontSize(value);
            Volatile.Write(ref _cachedSelectedFormulaFontSize, applied);
            WordDoubleClickHook.TraceMessage($"ribbon-font-size-completed requested={value} applied={applied}");
            SetStatus(T(
                $"公式字号已设置为 {FormulaFontSize.Describe(applied)}。",
                $"Formula font size set to {FormulaFontSize.Describe(applied, english: true)}."));
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage($"ribbon-font-size-failed requested={value} error={error}");
            SetStatus(T($"无法设置公式字号：{error.Message}", $"Unable to set formula font size: {error.Message}"));
        }
        finally { InvalidateFormulaFontControls(); }
    }

    private float? GetCachedSelectedFormulaFontSize()
    {
        // Ribbon can ask for its value while a COM call pumps Word events inside
        // an edit. No selection/OMath inspection is safe until that write ends.
        if (Volatile.Read(ref _formulaFormatMutationDepth) > 0)
        {
            Interlocked.Increment(ref _formulaFontReadsDeferredDuringMutation);
            return null;
        }
        // Keep the baseline positive-only Ribbon cache. Prose is now rejected
        // by the service's local range probe, so it needs no persistent negative
        // cache that could outlive Word's native selection notifications.
        var cached = Volatile.Read(ref _cachedSelectedFormulaFontSize);
        if (!double.IsNaN(cached)) return (float)cached;
        var size = _formulaService?.GetSelectedFormulaFontSize();
        Volatile.Write(
            ref _cachedSelectedFormulaFontSize,
            size.HasValue ? size.Value : double.NaN);
        return size;
    }

    private void ScheduleFormulaFontControlsInvalidation(Selection selection)
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null || Volatile.Read(ref _formulaFormatMutationDepth) > 0) return;
        var generation = Volatile.Read(ref _typingCaretNormalizationGeneration);

        if (!TryResolveFormulaRibbonOwnerBounds(selection, out var ownerStart, out var ownerEnd))
        {
            _lastFormulaRibbonOwnerStart = int.MinValue;
            _lastFormulaRibbonOwnerEnd = int.MinValue;
            // Ordinary prose-to-prose caret movement cannot change formula controls.
            // If the controls are already disabled, do absolutely no Word/OMML work.
            if (double.IsNaN(Volatile.Read(ref _cachedSelectedFormulaFontSize)))
                return;
            Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
            InvalidateFormulaFontControls();
            return;
        }

        if (ownerStart == _lastFormulaRibbonOwnerStart
            && ownerEnd == _lastFormulaRibbonOwnerEnd
            && !double.IsNaN(Volatile.Read(ref _cachedSelectedFormulaFontSize)))
            return;
        _lastFormulaRibbonOwnerStart = ownerStart;
        _lastFormulaRibbonOwnerEnd = ownerEnd;
        Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
        if (Interlocked.Exchange(ref _formulaFontInvalidationPending, 1) != 0)
            return;
        dispatcher.Post(() =>
        {
            using var perf = WordSelectionPerformance.Start("deferred-ribbon-font");
            Interlocked.Exchange(ref _formulaFontInvalidationPending, 0);
            if (generation != Volatile.Read(ref _typingCaretNormalizationGeneration)
                || Volatile.Read(ref _formulaFormatMutationDepth) > 0)
                return;
            // A queued refresh may execute after the caret left the formula.
            // The current owner sentinel is sufficient; no negative result cache.
            if (_lastFormulaRibbonOwnerStart == int.MinValue) return;
            try
            {
                if (_acceptanceSelectionDiagnostics)
                    Interlocked.Increment(ref _acceptanceFormulaStateReadCount);
                var size = _formulaService?.GetSelectedFormulaFontSize();
                Volatile.Write(
                    ref _cachedSelectedFormulaFontSize,
                    size.HasValue ? size.Value : double.NaN);
            }
            catch
            {
                Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
            }
            InvalidateFormulaFontControls();
        });
    }

    private static bool TryResolveFormulaRibbonOwnerBounds(
        Selection selection,
        out int ownerStart,
        out int ownerEnd)
    {
        ownerStart = int.MinValue;
        ownerEnd = int.MinValue;
        Range? selectionRange = null;
        Document? document = null;
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        Tables? tables = null;
        Table? table = null;
        Rows? tableRows = null;
        Columns? tableColumns = null;
        Cell? centerCell = null;
        Range? centerRange = null;
        Cell? numberCell = null;
        Range? numberRange = null;
        Fields? numberFields = null;
        Field? numberField = null;
        Range? numberCode = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Range? content = null;
        try
        {
            selectionRange = selection.Range;
            shapes = selectionRange.InlineShapes;
            if (shapes.Count == 1)
            {
                shape = shapes[1];
                shapeRange = shape.Range;
                ownerStart = shapeRange.Start;
                ownerEnd = shapeRange.End;
                return true;
            }

            document = selectionRange.Document;

            // A numbered OMML 1x3 contains many nested OMath ranges. A collapsed
            // caret can therefore appear to move between different inner OMaths
            // even though the user is still in the same formula. Resolve the
            // stable center-cell Display OMath first so Ribbon state is read only
            // once when entering this managed equation host.
            // Ribbon owner discovery is structural. Information() requests a
            // full layout pass after an edit even for an ordinary prose caret.
            if (WordEquationNumbering.RangeIsWhollyWithinTable(selectionRange))
            {
                tables = selectionRange.Tables;
                if (tables.Count == 1)
                {
                    table = tables[1];
                    tableRows = table.Rows;
                    tableColumns = table.Columns;
                    if (tableRows.Count == 1 && tableColumns.Count == 3)
                    {
                        numberCell = table.Cell(1, 3);
                        numberRange = numberCell.Range;
                        numberFields = numberRange.Fields;
                        if (numberFields.Count == 1)
                        {
                            numberField = numberFields[1];
                            numberCode = numberField.Code;
                            var codeText = numberCode.Text ?? string.Empty;
                            if (codeText.IndexOf(
                                    "SEQ VisualTeXEquation",
                                    StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                centerCell = table.Cell(1, 2);
                                centerRange = centerCell.Range;
                                maths = centerRange.OMaths;
                                if (maths.Count == 1)
                                {
                                    math = maths[1];
                                    mathRange = math.Range;
                                    ownerStart = mathRange.Start;
                                    ownerEnd = mathRange.End;
                                    return true;
                                }
                            }
                        }
                    }
                }
                ReleaseComObject(mathRange); mathRange = null;
                ReleaseComObject(math); math = null;
                ReleaseComObject(maths); maths = null;
            }

            content = document.Content;
            var contentStart = content.Start;
            var contentEnd = content.End;
            var probeStart = Math.Max(contentStart, selectionRange.Start - 1);
            var probeEnd = Math.Min(
                contentEnd,
                Math.Max(selectionRange.End + 1, selectionRange.Start + 1));
            probe = document.Range(probeStart, probeEnd);
            maths = probe.OMaths;
            var bestSpan = -1;
            for (var index = 1; index <= maths.Count; index++)
            {
                ReleaseComObject(mathRange);
                mathRange = null;
                ReleaseComObject(math);
                math = maths[index];
                mathRange = math.Range;
                if (selectionRange.Start < mathRange.Start - 1
                    || selectionRange.Start > mathRange.End + 1)
                    continue;
                var span = mathRange.End - mathRange.Start;
                if (span <= bestSpan) continue;
                bestSpan = span;
                ownerStart = mathRange.Start;
                ownerEnd = mathRange.End;
            }
            return bestSpan >= 0;
        }
        catch
        {
            ownerStart = int.MinValue;
            ownerEnd = int.MinValue;
            return false;
        }
        finally
        {
            ReleaseComObject(mathRange);
            ReleaseComObject(math);
            ReleaseComObject(maths);
            ReleaseComObject(numberCode);
            ReleaseComObject(numberField);
            ReleaseComObject(numberFields);
            ReleaseComObject(numberRange);
            ReleaseComObject(numberCell);
            ReleaseComObject(centerRange);
            ReleaseComObject(centerCell);
            ReleaseComObject(tableColumns);
            ReleaseComObject(tableRows);
            ReleaseComObject(table);
            ReleaseComObject(tables);
            ReleaseComObject(shapeRange);
            ReleaseComObject(shape);
            ReleaseComObject(shapes);
            ReleaseComObject(probe);
            ReleaseComObject(content);
            ReleaseComObject(document);
            ReleaseComObject(selectionRange);
        }
    }

    private void InvalidateFormulaFontControls()
    {
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            ui.InvalidateControl("VisualTeX.WordVsto.FontSize");
            ui.InvalidateControl("VisualTeX.WordVsto.FontSizeDecrease");
            ui.InvalidateControl("VisualTeX.WordVsto.FontSizeIncrease");
        }
        catch { }
    }

    private void InvalidateEquationNumberFormatControls()
    {
        var ribbon = _ribbonUi;
        if (ribbon is null) return;
        try
        {
            dynamic ui = ribbon;
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormat");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatContinuous");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading1Dot");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading1Dash");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading2Dot");
            ui.InvalidateControl("VisualTeX.WordVsto.NumberFormatHeading2Dash");
        }
        catch { }
    }

    private void ScheduleOmmlRedrawUndoMetadataCleanup()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var application = _application;
        if (dispatcher is null
            || service is null
            || application is null
            || !service.HasPendingOmmlRedrawUndoCleanup
            || Volatile.Read(ref _formulaFormatMutationDepth) > 0
            || Interlocked.Exchange(ref _ommlRedrawUndoCleanupPending, 1) != 0)
            return;

        dispatcher.Post(() =>
        {
            Interlocked.Exchange(ref _ommlRedrawUndoCleanupPending, 0);
            if (Volatile.Read(ref _formulaFormatMutationDepth) > 0
                || !service.HasPendingOmmlRedrawUndoCleanup)
            {
                StopOmmlRedrawUndoCleanupWatcherIfIdle();
                return;
            }
            Document? document = null;
            try
            {
                document = application.ActiveDocument;
                if (document is null) return;
                _ = service.TryCleanupPendingOmmlRedrawUndoMetadata(document);
            }
            catch (Exception error)
            {
                WordDoubleClickHook.TraceMessage(
                    $"redraw-omml-undo-orphan-metadata-cleanup-failed error={error.GetType().Name}:{error.Message}");
            }
            finally
            {
                ReleaseComObject(document);
                StopOmmlRedrawUndoCleanupWatcherIfIdle();
            }
        });
    }

    private void EnsureOmmlRedrawUndoCleanupWatcher()
    {
        lock (_ommlRedrawUndoCleanupWatcherGate)
        {
            if (_ommlRedrawUndoCleanupWatcher is not null) return;
            _ommlRedrawUndoCleanupWatcher = new System.Threading.Timer(
                _ =>
                {
                    var lifetime = _lifetime;
                    var service = _formulaService;
                    if (lifetime is null || lifetime.IsCancellationRequested
                        || service is null || !service.HasPendingOmmlRedrawUndoCleanup)
                    {
                        StopOmmlRedrawUndoCleanupWatcherIfIdle(force: lifetime is null || lifetime.IsCancellationRequested);
                        return;
                    }
                    ScheduleOmmlRedrawUndoMetadataCleanup();
                },
                null,
                dueTime: 2000,
                period: 2000);
        }
    }

    private void StopOmmlRedrawUndoCleanupWatcherIfIdle(bool force = false)
    {
        System.Threading.Timer? timer = null;
        lock (_ommlRedrawUndoCleanupWatcherGate)
        {
            if (!force && _formulaService?.HasPendingOmmlRedrawUndoCleanup == true)
                return;
            timer = _ommlRedrawUndoCleanupWatcher;
            _ommlRedrawUndoCleanupWatcher = null;
        }
        try { timer?.Dispose(); } catch { }
    }

    private void ScheduleTypingCaretNormalization()
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null
            || Interlocked.Exchange(ref _typingCaretNormalizationPending, 1) != 0)
            return;
        var generation = Volatile.Read(ref _typingCaretNormalizationGeneration);
        if (_acceptanceSelectionDiagnostics)
            Interlocked.Increment(ref _acceptanceDeferredCaretPassCount);
        dispatcher.Post(() =>
        {
            using var perf = WordSelectionPerformance.Start("deferred-caret");
            Interlocked.Exchange(ref _typingCaretNormalizationPending, 0);
            if (generation != Volatile.Read(ref _typingCaretNormalizationGeneration)
                || Volatile.Read(ref _formulaFormatMutationDepth) > 0)
                return;
            var service = _formulaService;
            var application = _application;
            if (service is null || application is null) return;
            Selection? currentSelection = null;
            try
            {
                currentSelection = application.Selection;
                if (Interlocked.CompareExchange(ref _normalizingTypingCaret, 1, 0) != 0)
                    return;
                try
                {
                    // SelectionChange can arrive one Word UI turn before the new
                    // right-cell paragraph/field boundary is fully materialized.
                    // Reuse this existing single deferred caret pass as a targeted
                    // retry; it never enumerates documents or sibling formulas.
                    var redirected = WordEquationNumbering
                        .TryRedirectManagedNativeOmmlNumberEndEnter(currentSelection);
                    if (!redirected)
                        service.NormalizeTypingCaretAfterInlineFormula(currentSelection);
                    else
                        ClearNativeOleTarget();
                }
                finally { Interlocked.Exchange(ref _normalizingTypingCaret, 0); }
            }
            catch { }
            finally { ReleaseComObject(currentSelection); }
        });
    }

    private void BeginFormulaFormatMutation(bool copyPasteRepair = false)
    {
        if (!copyPasteRepair)
        {
            // Explicit Ribbon inserts/edits are not clipboard pastes. Cancel any
            // queued copy probe before the new objects become visible.
            Interlocked.Increment(ref _copyPasteRepairGeneration);
            Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
        }
        Interlocked.Increment(ref _typingCaretNormalizationGeneration);
        Interlocked.Increment(ref _formulaFormatMutationDepth);
    }

    private void EndFormulaFormatMutation(bool copyPasteRepair = false)
    {
        var depth = Interlocked.Decrement(ref _formulaFormatMutationDepth);
        if (depth < 0)
            Interlocked.Exchange(ref _formulaFormatMutationDepth, 0);
        if (depth > 0) return;
        var deferred = Interlocked.Exchange(ref _formulaFontReadsDeferredDuringMutation, 0);
        if (deferred > 0)
            WordDoubleClickHook.TraceMessage($"formula-mutation-ribbon-reads-deferred count={deferred}");
        var generation = Volatile.Read(ref _typingCaretNormalizationGeneration);
        // End can run on a continuation thread. Refresh once on Word's dispatcher
        // after the completed write, cancelling it if another write has started.
        _dispatcher?.Post(() =>
        {
            if (generation != Volatile.Read(ref _typingCaretNormalizationGeneration)
                || Volatile.Read(ref _formulaFormatMutationDepth) > 0)
                return;
            if (!copyPasteRepair)
            {
                WordFormulaService.WordFormulaCopySnapshot? copy;
                lock (_copyPasteGate) { copy = _formulaCopySnapshot; }
                if (copy is not null) _formulaService?.RefreshCopySnapshotAfterExplicitMutation(copy);
            }
            _lastFormulaRibbonOwnerStart = int.MinValue;
            _lastFormulaRibbonOwnerEnd = int.MinValue;
            Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
            InvalidateFormulaFontControls();
        });
    }

    private void OnWordClipboardGesture(WordClipboardGestureEvent gesture)
    {
        if (gesture.Gesture != WordClipboardGesture.Copy) return;
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;

        dispatcher.Post(() =>
        {
            try
            {
                var snapshot = service.CaptureSelectedFormulaForCopy();
                WordFormulaService.WordFormulaCopySnapshot? existingSnapshot;
                lock (_copyPasteGate) { existingSnapshot = _formulaCopySnapshot; }
                if (snapshot is null && existingSnapshot is not null
                    && (IsClipboardChangeFromRecentPasteRepair(gesture.ObservedTimestamp)
                        || service.HasPotentialCopiedInsertion(existingSnapshot)))
                {
                    // Word's delayed OLE rendering can advance the clipboard
                    // sequence during Paste, before our repair has run. Keep the
                    // captured source only for a verifiable local insertion; the
                    // repair still checks the physical FormulaId/native content.
                    var continuedSequence = WordCopyPasteHook.CurrentClipboardSequence;
                    lock (_copyPasteGate)
                    {
                        if (ReferenceEquals(_formulaCopySnapshot, existingSnapshot))
                            _formulaCopyClipboardSequence = continuedSequence;
                    }
                    var continuedWatchGeneration =
                        Interlocked.Increment(ref _copyPasteWatchGeneration);
                    ScheduleCopiedFormulaInsertionWatch(
                        existingSnapshot,
                        continuedSequence,
                        continuedWatchGeneration,
                        existingSnapshot.KnownDocumentEnd,
                        attempt: 0);
                    WordDoubleClickHook.TraceMessage(
                        $"copy-snapshot-sequence-continued mode={existingSnapshot.ObjectMode} "
                        + $"clipboardSequence={continuedSequence}");
                    TrySchedulePastedFormulaRepairFromSelectionChange();
                    return;
                }

                lock (_copyPasteGate)
                {
                    _formulaCopySnapshot = snapshot;
                    _formulaCopyClipboardSequence = snapshot is null
                        ? 0
                        : gesture.ClipboardSequence;
                }
                var watchGeneration = Interlocked.Increment(ref _copyPasteWatchGeneration);
                if (snapshot is not null)
                {
                    Interlocked.Exchange(ref _lastSuccessfulPasteRepairStartedAt, 0);
                    Interlocked.Exchange(ref _lastSuccessfulPasteRepairCompletedAt, 0);
                    ScheduleCopiedFormulaInsertionWatch(
                        snapshot,
                        gesture.ClipboardSequence,
                        watchGeneration,
                        snapshot.KnownDocumentEnd,
                        attempt: 0);
                }
                Interlocked.Increment(ref _copyPasteRepairGeneration);
                Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
                WordDoubleClickHook.TraceMessage(
                    snapshot is null
                        ? $"copy-snapshot-cleared clipboardSequence={gesture.ClipboardSequence}"
                        : $"copy-snapshot-captured mode={snapshot.ObjectMode} "
                            + $"source={snapshot.SourceStart}:{snapshot.SourceEnd} "
                            + $"numbered={snapshot.Metadata.Numbered} "
                            + $"inline={snapshot.KnownInlineShapeCount} omml={snapshot.KnownOmmlCount} "
                            + $"clipboardSequence={gesture.ClipboardSequence}");
            }
            catch (Exception error)
            {
                lock (_copyPasteGate)
                {
                    _formulaCopySnapshot = null;
                    _formulaCopyClipboardSequence = 0;
                }
                Interlocked.Increment(ref _copyPasteWatchGeneration);
                Interlocked.Increment(ref _copyPasteRepairGeneration);
                Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
                WordDoubleClickHook.TraceMessage(
                    $"copy-snapshot-failed {error.GetType().Name}: {error.Message}");
            }
        });
    }

    private void ScheduleCopiedFormulaInsertionWatch(
        WordFormulaService.WordFormulaCopySnapshot snapshot,
        uint clipboardSequence,
        int generation,
        int lastDocumentEnd,
        int attempt)
    {
        // Fast cadence only around the user's immediate Copy/Paste gesture, then
        // back off to one O(1) Content.End read per second. The watch is bounded;
        // WindowSelectionChange remains available for much later Paste operations.
        if (attempt >= 320) return;
        var lifetime = _lifetime;
        var dispatcher = _dispatcher;
        if (lifetime is null || dispatcher is null || lifetime.IsCancellationRequested)
            return;
        var delayMilliseconds = attempt < 42
            ? 120
            : attempt < 72
                ? 350
                : 1000;
        var token = lifetime.Token;
        _ = Task.Delay(delayMilliseconds, token).ContinueWith(
            task =>
            {
                if (task.IsCanceled || token.IsCancellationRequested) return;
                dispatcher.Post(() =>
                {
                    if (generation != Volatile.Read(ref _copyPasteWatchGeneration))
                        return;
                    lock (_copyPasteGate)
                    {
                        if (!ReferenceEquals(_formulaCopySnapshot, snapshot)
                            || _formulaCopyClipboardSequence != clipboardSequence)
                            return;
                    }
                    if (WordCopyPasteHook.CurrentClipboardSequence != clipboardSequence)
                        return;
                    var service = _formulaService;
                    if (service is null)
                        return;
                    var currentDocumentEnd = service.ReadCopyWatchDocumentEnd(snapshot);
                    if (currentDocumentEnd < 0)
                        return;
                    if (currentDocumentEnd != lastDocumentEnd)
                    {
                        if (service.HasPotentialCopiedInsertion(snapshot))
                        {
                            WordDoubleClickHook.TraceMessage(
                                $"paste-watch-detected mode={snapshot.ObjectMode} "
                                + $"documentEnd={lastDocumentEnd}->{currentDocumentEnd} attempt={attempt}");
                            TrySchedulePastedFormulaRepairFromSelectionChange();
                        }
                        lastDocumentEnd = currentDocumentEnd;
                    }
                    ScheduleCopiedFormulaInsertionWatch(
                        snapshot,
                        clipboardSequence,
                        generation,
                        lastDocumentEnd,
                        attempt + 1);
                });
            },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private bool IsClipboardChangeFromRecentPasteRepair(long observedTimestamp)
    {
        var startedAt = Interlocked.Read(ref _lastSuccessfulPasteRepairStartedAt);
        var completedAt = Interlocked.Read(ref _lastSuccessfulPasteRepairCompletedAt);
        if (startedAt <= 0 || completedAt < startedAt) return false;
        var graceTicks = (long)(Stopwatch.Frequency * 0.5d);
        return observedTimestamp >= startedAt
            && observedTimestamp <= completedAt + graceTicks;
    }

    private void TrySchedulePastedFormulaRepairFromSelectionChange()
    {
        WordFormulaService.WordFormulaCopySnapshot? snapshot;
        uint clipboardSequence;
        lock (_copyPasteGate)
        {
            snapshot = _formulaCopySnapshot;
            clipboardSequence = _formulaCopyClipboardSequence;
        }
        if (snapshot is null || clipboardSequence == 0) return;

        var currentClipboardSequence = WordCopyPasteHook.CurrentClipboardSequence;
        if (currentClipboardSequence != clipboardSequence)
        {
            if (_formulaService?.HasPotentialCopiedInsertion(snapshot) == true)
            {
                lock (_copyPasteGate)
                {
                    if (!ReferenceEquals(_formulaCopySnapshot, snapshot)) return;
                    _formulaCopyClipboardSequence = currentClipboardSequence;
                }
                clipboardSequence = currentClipboardSequence;
            }
            else
            {
                lock (_copyPasteGate)
                {
                    if (ReferenceEquals(_formulaCopySnapshot, snapshot))
                    {
                        _formulaCopySnapshot = null;
                        _formulaCopyClipboardSequence = 0;
                    }
                }
                Interlocked.Increment(ref _copyPasteWatchGeneration);
                Interlocked.Increment(ref _copyPasteRepairGeneration);
                Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
                return;
            }
        }
        if (Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 1) != 0)
            return;

        var generation = Interlocked.Increment(ref _copyPasteRepairGeneration);
        SchedulePastedFormulaRepair(
            snapshot,
            clipboardSequence,
            generation,
            attempt: 0);
    }

    private void SchedulePastedFormulaRepair(
        WordFormulaService.WordFormulaCopySnapshot snapshot,
        uint clipboardSequence,
        int generation,
        int attempt)
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null) return;
        dispatcher.Post(() =>
        {
            if (generation != Volatile.Read(ref _copyPasteRepairGeneration))
            {
                Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
                return;
            }
            lock (_copyPasteGate)
            {
                if (!ReferenceEquals(_formulaCopySnapshot, snapshot))
                {
                    Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
                    return;
                }
            }
            var service = _formulaService;
            if (service is null || Volatile.Read(ref _formulaFormatMutationDepth) > 0)
            {
                RetryPastedFormulaRepair(snapshot, clipboardSequence, generation, attempt);
                return;
            }

            WordFormulaService.PastedFormulaRepairResult result;
            var repairStartedAt = Stopwatch.GetTimestamp();
            BeginFormulaFormatMutation(copyPasteRepair: true);
            try
            {
                result = service.RepairPastedFormula(snapshot);
            }
            catch (Exception error)
            {
                Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
                WordDoubleClickHook.TraceMessage(
                    $"paste-repair-failed mode={snapshot.ObjectMode} "
                    + $"attempt={attempt} error={error.GetType().Name}:{error.Message}");
                return;
            }
            finally { EndFormulaFormatMutation(copyPasteRepair: true); }

            if (result == WordFormulaService.PastedFormulaRepairResult.NotReady)
            {
                RetryPastedFormulaRepair(snapshot, clipboardSequence, generation, attempt);
                return;
            }
            if (result == WordFormulaService.PastedFormulaRepairResult.Repaired)
            {
                var repairCompletedAt = Stopwatch.GetTimestamp();
                Interlocked.Exchange(ref _lastSuccessfulPasteRepairStartedAt, repairStartedAt);
                Interlocked.Exchange(ref _lastSuccessfulPasteRepairCompletedAt, repairCompletedAt);
                lock (_copyPasteGate)
                {
                    if (ReferenceEquals(_formulaCopySnapshot, snapshot))
                        _formulaCopyClipboardSequence = WordCopyPasteHook.CurrentClipboardSequence;
                }
            }
            Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
            WordDoubleClickHook.TraceMessage(
                $"paste-repair-complete mode={snapshot.ObjectMode} result={result} attempt={attempt}");
        });
    }

    private void RetryPastedFormulaRepair(
        WordFormulaService.WordFormulaCopySnapshot snapshot,
        uint clipboardSequence,
        int generation,
        int attempt)
    {
        if (attempt >= 6)
        {
            Interlocked.Exchange(ref _copyPasteSelectionRepairPending, 0);
            WordDoubleClickHook.TraceMessage(
                $"paste-repair-timeout mode={snapshot.ObjectMode} attempts={attempt + 1}");
            return;
        }
        var lifetime = _lifetime;
        if (lifetime is null || lifetime.IsCancellationRequested) return;
        var token = lifetime.Token;
        _ = Task.Delay(90, token).ContinueWith(
            task =>
            {
                if (task.IsCanceled || token.IsCancellationRequested) return;
                SchedulePastedFormulaRepair(
                    snapshot,
                    clipboardSequence,
                    generation,
                    attempt + 1);
            },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private void OnWindowActivate(Document document, Window window)
    {
        try { _doubleClickHook?.BindOwnerWindow(window.Hwnd); } catch { }
        try { _copyPasteHook?.BindOwnerWindow(window.Hwnd); } catch { }
        WordFormulaService.WordFormulaCopySnapshot? copySnapshot = null;
        uint copySequence = 0;
        lock (_copyPasteGate)
        {
            copySnapshot = _formulaCopySnapshot;
            copySequence = _formulaCopyClipboardSequence;
        }
        if (copySnapshot is not null && copySequence != 0)
        {
            if (WordCopyPasteHook.CurrentClipboardSequence == copySequence)
            {
                try
                {
                    _formulaService?.TrackCopySnapshotDocument(copySnapshot);
                    var watchGeneration =
                        Interlocked.Increment(ref _copyPasteWatchGeneration);
                    ScheduleCopiedFormulaInsertionWatch(
                        copySnapshot,
                        copySequence,
                        watchGeneration,
                        copySnapshot.KnownDocumentEnd,
                        attempt: 0);
                }
                catch { }
            }
            else
            {
                lock (_copyPasteGate)
                {
                    if (_formulaCopyClipboardSequence == copySequence)
                    {
                        _formulaCopySnapshot = null;
                        _formulaCopyClipboardSequence = 0;
                        Interlocked.Increment(ref _copyPasteWatchGeneration);
                    }
                }
            }
        }
        ClearNativeOleTarget();
        _cachedEquationNumberFormatId = null;
        Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
        _lastFormulaRibbonOwnerStart = int.MinValue;
        _lastFormulaRibbonOwnerEnd = int.MinValue;
        InvalidateEquationNumberFormatControls();
        Selection? selection = null;
        try
        {
            selection = _application?.Selection;
            if (selection is not null)
                ScheduleFormulaFontControlsInvalidation(selection);
        }
        catch { }
        finally { ReleaseComObject(selection); }

        ScheduleOmmlRedrawUndoMetadataCleanup();

        // Returning from the VisualTeX Office editor does not necessarily move
        // Word's Selection, so WindowSelectionChange may never fire. Word can
        // still rebuild the collapsed caret's character format from the adjacent
        // MathType OLE field run when this window becomes active again. Reuse the
        // existing deferred O(1) caret normalization after activation so the first
        // real keystroke inherits nearby prose rather than Equation.DSMT4.
        ScheduleTypingCaretNormalization();
    }

    private void OnDocumentOpen(Document document)
    {
        _cachedEquationNumberFormatId = null;
        InvalidateEquationNumberFormatControls();
        RefreshNumberedOmmlTabLayoutsAfterOpen(document);
    }

    private void RefreshNumberedOmmlTabLayoutsAfterOpen(Document document)
    {
        try
        {
            document.Repaginate();
            var refreshed = WordEquationNumbering.RefreshNumberedOmmlTabLayouts(
                document);
            if (refreshed > 0)
            {
                // RefreshNumberedOmmlTabLayouts now completes the direct-SEQ 1x3
                // host synchronously. The retired Shape/TextBox design needed five
                // later dispatcher turns; scheduling those scans for current tables
                // only keeps Word's UI thread busy after the document is usable.
                WordDoubleClickHook.TraceMessage(
                    $"document-open-omml-tab-layout-refreshed formulas={refreshed}");
            }
        }
        catch (Exception error)
        {
            // A protected/read-only document can reject paragraph-format changes.
            // Opening the document must remain successful; the explicit update-
            // number command can retry after editing is enabled.
            WordDoubleClickHook.TraceMessage(
                $"document-open-omml-tab-layout-refresh-failed: {error}");
        }
    }

    private void OnDocumentBeforeSave(
        Document document,
        ref bool saveAsUi,
        ref bool cancel)
    {
        if (cancel) return;
        try
        {
            if (_formulaService?.HasPendingOmmlRedrawUndoCleanup == true)
                _ = _formulaService.TryCleanupPendingOmmlRedrawUndoMetadata(document);
            _formulaService?.NormalizeInlineOleParagraphBaselinesBeforeSave(
                document);
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"document-before-save-baseline-normalization-failed: {error}");
        }
    }

    private void OnWindowSelectionChange(Selection selection)
    {
        using var perf = WordSelectionPerformance.Start("selection-change");
        if (Volatile.Read(ref _normalizingTypingCaret) != 0) return;
        if (_acceptanceSelectionDiagnostics)
            Interlocked.Increment(ref _acceptanceSelectionChangeCount);
        if (Volatile.Read(ref _formulaFormatMutationDepth) > 0)
        {
            Interlocked.Increment(ref _formulaFontReadsDeferredDuringMutation);
            ClearNativeOleTarget();
            return;
        }
        ScheduleOmmlRedrawUndoMetadataCleanup();
        // The guard must precede the Ribbon owner probe as well as normalization:
        // owner discovery itself reads OMaths/Cells and can reenter an incomplete
        // table write. The deferred callback also checks the mutation generation.
        ScheduleFormulaFontControlsInvalidation(selection);
        perf?.Mark("ribbon-owner");
        var service = _formulaService;
        var application = _application;
        if (service is null || application is null)
        {
            ClearNativeOleTarget();
            return;
        }
        TrySchedulePastedFormulaRepairFromSelectionChange();

        Range? range = null;
        Window? window = null;
        try
        {
            var redirectedNumberEndEnter = false;
            if (Interlocked.CompareExchange(ref _normalizingTypingCaret, 1, 0) == 0)
            {
                try
                {
                    // Pressing Enter at the end of a direct-SEQ 1x3 number does not
                    // leave the table in Word: Word splits the right cell and drags
                    // the SEQ/VTEq* tree into a second cell paragraph. Repair only
                    // that exact structural state, then place the caret in an ordinary
                    // body paragraph after the table. This is O(1) for normal clicks;
                    // no OMML metadata is touched unless the right cell has already
                    // acquired an extra empty paragraph.
                    redirectedNumberEndEnter =
                        WordEquationNumbering
                            .TryRedirectManagedNativeOmmlNumberEndEnter(selection);
                    if (!redirectedNumberEndEnter)
                        service.NormalizeTypingCaretAfterInlineFormula(selection);
                }
                finally { Interlocked.Exchange(ref _normalizingTypingCaret, 0); }
            }
            perf?.Mark("caret-normalization");
            if (redirectedNumberEndEnter)
            {
                ClearNativeOleTarget();
                return;
            }
            // Do not inspect OMML metadata here. Word fires SelectionChange while
            // entering its native equation editor, and touching the OMath at that
            // point can disturb the caret state. MathType is even more sensitive:
            // activating its IDataObject on a single click can synchronously open
            // the external OLE server. Cache MathType's rectangle using only its
            // ProgID/CLSID and defer content import until the actual double-click.
            // Cache MathType geometry regardless of the preference value. The
            // first physical click selects the Equation.DSMT4 OLE; defer any content
            // import until the actual double-click and retain only its range/rect.
            var isMathTypeOle = service.IsSelectedMathTypeOle();
            perf?.Mark("ole-kind");
            OfficeSelection? selected = null;
            if (!isMathTypeOle)
            {
                if (!service.IsSelectedNativeOle())
                {
                    ClearNativeOleTarget();
                    return;
                }
                // A single click only caches the physical OLE rectangle, just
                // like MathType. Metadata import and copy-identity repair still
                // run in ReadSelection when an actual edit/double-click begins.
                // Do not activate the companion OLE server merely to select it.
            }

            // The deferred caret retry exists only for Word's OLE selection→caret
            // transition. Ordinary prose and native OMML clicks already completed
            // their synchronous O(1) checks above and must not queue a second COM
            // pass after every mouse click.
            ScheduleTypingCaretNormalization();

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
                _nativeOleTargetRangeStart = range.Start;
                _nativeOleTargetRangeEnd = range.End;
                _nativeOleTargetIsMathType = isMathTypeOle;
                _nativeOleTargetActive = true;
            }
            WordDoubleClickHook.TraceMessage(
                $"cache-active formulaId={selected?.FormulaId ?? "<mathType-pending>"} "
                + $"objectMode={(isMathTypeOle ? FormulaOleContract.MathTypeOleMode : selected?.ObjectMode)} "
                + $"rect={left - padding},{top - padding},{left + width + padding},{top + height + padding}");
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
            // MathType OLE double-click has two explicit owners. When VisualTeX
            // editing is enabled, the hook observes the native OLE and this Word
            // event is a synchronous fallback that opens VisualTeX. When disabled,
            // VisualTeX must not cancel, suppress, replay or asynchronously inspect
            // the double-click: Word and MathType receive the native gesture alone.
            var selectedMathTypeOle = _formulaService?.IsSelectedMathTypeOle() == true;
            if (selectedMathTypeOle)
            {
                if (!MathTypeDoubleClickPreference.IsEnabled())
                {
                    // The hook object merely being alive does not prove that it
                    // claimed this physical double-click. After inserting another
                    // formula its cached rectangle can still refer to the newly
                    // created Equation.DSMT4; a direct double-click on an older
                    // formula then passes through the hook. Cancelling Word here in
                    // that state creates a black hole: neither Word nor the hook
                    // opens MathType. Cancel only when this selected equation
                    // overlaps the short-lived range claim recorded by the hook for
                    // the current gesture. Otherwise release Word's native behavior.
                    var hookOwnsCurrentGesture =
                        _doubleClickHook is not null
                        && HasRecentClaimedNativeMathTypeGesture(selection);
                    cancel = hookOwnsCurrentGesture;
                    if (!cancel) ClearNativeOleTarget();
                    WordDoubleClickHook.TraceMessage(
                        cancel
                            ? "window-before-double-click-mathtype-native-owned-by-hook"
                            : "window-before-double-click-mathtype-released-to-native-editor-unclaimed");
                    return;
                }

                // Do not make successful MathType double-click editing depend on
                // the low-level mouse hook completing its second-click callback.
                // Some real Word sessions keep the hook object alive but fail to
                // deliver that callback; the old code still cancelled Word's native
                // double-click here and then returned, leaving the user with no
                // editor at all. WindowBeforeDoubleClick already proves that Word
                // resolved the click onto the selected Equation.DSMT4 object, so
                // use it as a deterministic fallback (and, in practice, primary
                // completion path). The hook may have started the same session a
                // few milliseconds earlier; TryBeginDoubleClickSession deduplicates
                // by FormulaId for one second, so the two routes cannot open two
                // editors.
                var selectedMathType = _formulaService?.ReadSelection(selection);
                if (selectedMathType?.Metadata is null
                    || string.IsNullOrWhiteSpace(selectedMathType.FormulaId)
                    || !string.Equals(
                        selectedMathType.ObjectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal)
                    || !WordDoubleClickRouting.ShouldOpenVisualTeX(selectedMathType))
                {
                    // If the object cannot be safely imported into VisualTeX, do
                    // not swallow Word's normal MathType activation.
                    cancel = false;
                    ClearNativeOleTarget();
                    WordDoubleClickHook.TraceMessage(
                        "window-before-double-click-mathtype-fallback-to-native");
                    return;
                }

                cancel = true;
                ClearNativeOleTarget();
                var started = TryBeginDoubleClickSession(selectedMathType);
                WordDoubleClickHook.TraceMessage(
                    $"window-before-double-click-mathtype-visualtex started={started} hookPresent={_doubleClickHook is not null}");
                return;
            }
            var selected = _formulaService?.ReadSelection(selection);
            if (selected?.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
                return;

            var shouldOpenVisualTeX = WordDoubleClickRouting.ShouldOpenVisualTeX(selected);
            var hasMousePoint = TryGetRecentMouseDoubleClickPoint(
                out var screenX,
                out var screenY);
            var pointHitsFormula = !hasMousePoint
                ? _doubleClickHook is null
                : _formulaService?.IsFormulaAtScreenPoint(
                    selected,
                    screenX,
                    screenY) == true;
            WordDoubleClickHook.TraceMessage(
                $"window-before-double-click formulaId={selected.FormulaId} "
                + $"objectMode={selected.ObjectMode ?? "<null>"} "
                + $"shouldOpenVisualTeX={shouldOpenVisualTeX} "
                + $"hasMousePoint={hasMousePoint} pointHitsFormula={pointHitsFormula}");
            if (!shouldOpenVisualTeX || !pointHitsFormula) return;

            cancel = true;
            ClearNativeOleTarget();
            TryBeginDoubleClickSession(selected);
        }
        catch (Exception error)
        {
            SetStatus($"VisualTeX 双击检测失败：{error.Message}");
        }
    }

    private WordDoubleClickHook.NativeOleDecision ShouldInterceptNativeOleDoubleClick(
        int screenX,
        int screenY)
    {
        RememberMouseDoubleClickPoint(screenX, screenY);
        lock (_nativeOleTargetGate)
        {
            var hitsCachedOle = _nativeOleTargetActive
                && screenX >= _nativeOleTargetLeft
                && screenX <= _nativeOleTargetRight
                && screenY >= _nativeOleTargetTop
                && screenY <= _nativeOleTargetBottom;
            if (!hitsCachedOle)
                return WordDoubleClickHook.NativeOleDecision.PassThrough;

            if (_nativeOleTargetIsMathType)
            {
                if (!MathTypeDoubleClickPreference.IsEnabled())
                {
                    // Word 2021 can consume a released Equation.DSMT4 double-click
                    // without opening MathType or raising WindowBeforeDoubleClick.
                    // Suppress only the second button-down so Word cannot start a
                    // competing OLE activation; the callback below schedules exactly
                    // one native Open verb on a later Office UI turn. Freeze the
                    // cached Word range now: numbered MathType paragraphs can move
                    // Selection onto their REF/field text after the first click, so
                    // the later callback must not rediscover identity from Selection.
                    _pendingNativeMathTypeRangeStart = _nativeOleTargetRangeStart;
                    _pendingNativeMathTypeRangeEnd = _nativeOleTargetRangeEnd;
                    _claimedNativeMathTypeRangeStart = _nativeOleTargetRangeStart;
                    _claimedNativeMathTypeRangeEnd = _nativeOleTargetRangeEnd;
                    _claimedNativeMathTypeAt = DateTimeOffset.UtcNow;
                    return WordDoubleClickHook.NativeOleDecision.OpenNativeOle;
                }

                // MathType must never enter the old black-hole state where the
                // low-level hook suppresses Word's second button-down and then an
                // asynchronous callback fails to open VisualTeX. Observe and
                // dispatch the MathType OLE here, but let Word receive the click so
                // WindowBeforeDoubleClick can synchronously cancel native MathType
                // activation and start the same edit Session as a redundant path.
                return WordDoubleClickHook.NativeOleDecision.ObserveNativeOle;
            }

            // VisualTeX's own native OLE still needs the historical suppression:
            // allowing Word to receive its complete double-click can activate the
            // embedded server before VisualTeX has a chance to open the editor.
            return WordDoubleClickHook.NativeOleDecision.InterceptNativeOle;
        }
    }

    private void OnNativeWordDoubleClick(
        bool interceptedNativeOle,
        int screenX,
        int screenY)
    {
        WordDoubleClickHook.TraceMessage(
            $"addin-callback-received interceptedNativeOle={interceptedNativeOle} "
            + $"x={screenX} y={screenY}");
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        _ = dispatcher.InvokeAsync(() =>
        {
            try
            {
                if (interceptedNativeOle
                    && !MathTypeDoubleClickPreference.IsEnabled())
                {
                    int targetStart;
                    int targetEnd;
                    lock (_nativeOleTargetGate)
                    {
                        targetStart = _pendingNativeMathTypeRangeStart;
                        targetEnd = _pendingNativeMathTypeRangeEnd;
                        if (targetStart >= 0 && targetEnd > targetStart)
                        {
                            _pendingNativeMathTypeRangeStart = -1;
                            _pendingNativeMathTypeRangeEnd = -1;
                        }
                    }
                    if (targetStart >= 0 && targetEnd > targetStart)
                    {
                        ClearNativeOleTarget();
                        QueueNativeMathTypeEditorOpen(
                            dispatcher,
                            service,
                            targetStart,
                            targetEnd,
                            screenX,
                            screenY);
                        return true;
                    }
                    // `interceptedNativeOle` also covers VisualTeX.Formula.1. A
                    // disabled MathType preference must never swallow VisualTeX's
                    // own OLE double-click just because both use the same low-level
                    // hook callback. Only a frozen Equation.DSMT4 range proves that
                    // this callback belongs to native MathType; otherwise continue
                    // through the ordinary VisualTeX OLE edit route below.
                    WordDoubleClickHook.TraceMessage(
                        "addin-native-mathtype-open-not-applicable continue-native-ole-route");
                }

                var selected = interceptedNativeOle
                    ? service.ReadSelection()
                    : service.ReadVisualTeXOmmlAtScreenPoint(screenX, screenY);
                WordDoubleClickHook.TraceMessage(
                    $"addin-coordinate-selection formulaId={selected?.FormulaId ?? "<null>"} "
                    + $"objectMode={selected?.ObjectMode ?? "<null>"}");
                if (!WordDoubleClickRouting.ShouldOpenVisualTeX(selected))
                    return false;

                if (interceptedNativeOle)
                {
                    var supportedOleMode = string.Equals(
                            selected!.ObjectMode,
                            FormulaOleContract.NativeOleMode,
                            StringComparison.Ordinal)
                        || string.Equals(
                            selected.ObjectMode,
                            FormulaOleContract.MathTypeOleMode,
                            StringComparison.Ordinal);
                    if (!supportedOleMode
                        || !service.IsFormulaAtScreenPoint(
                            selected,
                            screenX,
                            screenY))
                    {
                        WordDoubleClickHook.TraceMessage(
                            "addin-coordinate-selection-rejected-native-ole-miss");
                        return false;
                    }
                }
                else if (!string.Equals(
                             selected!.ObjectMode,
                             FormulaOleContract.WordOmmlMode,
                             StringComparison.Ordinal))
                {
                    return false;
                }

                var started = TryBeginDoubleClickSession(selected);
                WordDoubleClickHook.TraceMessage($"addin-session-started={started}");
                return started;
            }
            catch (Exception error)
            {
                SetStatus($"VisualTeX Word 双击检测失败：{error.Message}");
                return false;
            }
        });
    }

    private static void QueueNativeMathTypeEditorOpen(
        OfficeUiDispatcher dispatcher,
        WordFormulaService service,
        int targetStart,
        int targetEnd,
        int screenX,
        int screenY)
    {
        const int delayMilliseconds = 300;
        WordDoubleClickHook.TraceMessage(
            $"addin-native-mathtype-open-scheduled delayMs={delayMilliseconds} "
            + $"range={targetStart}:{targetEnd} x={screenX} y={screenY}");
        ThreadPool.QueueUserWorkItem(_ =>
        {
            Thread.Sleep(delayMilliseconds);
            dispatcher.Post(() =>
            {
                try
                {
                    // The user may change the preference while Word's message queue
                    // settles, but Selection is deliberately irrelevant here. A
                    // numbered MathType paragraph can move Selection onto its REF or
                    // field text after the first click even though the double-click
                    // unquestionably hit the OLE. Re-resolve the frozen range and
                    // verify the original screen point before invoking the native verb.
                    if (MathTypeDoubleClickPreference.IsEnabled())
                    {
                        WordDoubleClickHook.TraceMessage(
                            "addin-native-mathtype-open-skipped preference-changed");
                        return;
                    }

                    var openedNative = service.OpenMathTypeNativeEditorAtRange(
                        targetStart,
                        targetEnd,
                        screenX,
                        screenY);
                    WordDoubleClickHook.TraceMessage(
                        $"addin-native-mathtype-open started={openedNative}");
                }
                catch (Exception error)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"addin-native-mathtype-open-error {error.GetType().Name}: {error.Message}");
                }
            });
        });
    }

    private bool TryBeginDoubleClickSession(OfficeSelection? selected)
    {
        var formulaId = selected?.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId)
            || !WordDoubleClickRouting.ShouldOpenVisualTeX(selected))
            return false;

        // MathType OLE is third-party owned, so ReadMetadata() deliberately creates
        // a fresh transient FormulaId on every read. The low-level mouse hook and
        // WindowBeforeDoubleClick can therefore observe the same Equation.DSMT4
        // object with two different FormulaIds a few milliseconds apart. Deduplicate
        // MathType by its stable Word range ObjectId; VisualTeX-owned OLE/OMML keeps
        // using the durable FormulaId identity.
        var identity = string.Equals(
                selected!.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
            && !string.IsNullOrWhiteSpace(selected.ObjectId)
            ? $"mathtype:{selected.ObjectId}"
            : $"formula:{formulaId}";
        var now = DateTimeOffset.UtcNow;
        if (string.Equals(identity, _lastDoubleClickIdentity, StringComparison.Ordinal)
            && now - _lastDoubleClickAt < TimeSpan.FromSeconds(1))
            return false;
        _lastDoubleClickIdentity = identity;
        _lastDoubleClickAt = now;
        BeginSession("edit", null, null, capturedSelection: selected);
        return true;
    }

    private bool HasRecentClaimedNativeMathTypeGesture(Selection selection)
    {
        int targetStart;
        int targetEnd;
        DateTimeOffset claimedAt;
        lock (_nativeOleTargetGate)
        {
            targetStart = _claimedNativeMathTypeRangeStart;
            targetEnd = _claimedNativeMathTypeRangeEnd;
            claimedAt = _claimedNativeMathTypeAt;
        }
        if (targetStart < 0
            || targetEnd <= targetStart
            || DateTimeOffset.UtcNow - claimedAt > TimeSpan.FromSeconds(1))
            return false;

        Range? range = null;
        try
        {
            range = selection.Range;
            if (range.Start == range.End)
                return range.Start >= targetStart && range.Start <= targetEnd;
            return range.Start < targetEnd && range.End > targetStart;
        }
        catch { return false; }
        finally { ReleaseComObject(range); }
    }

    private void RememberMouseDoubleClickPoint(int screenX, int screenY)
    {
        lock (_mouseDoubleClickGate)
        {
            _mouseDoubleClickX = screenX;
            _mouseDoubleClickY = screenY;
            _mouseDoubleClickAt = DateTimeOffset.UtcNow;
            _mouseDoubleClickPointActive = true;
        }
    }

    private bool TryGetRecentMouseDoubleClickPoint(
        out int screenX,
        out int screenY)
    {
        lock (_mouseDoubleClickGate)
        {
            screenX = _mouseDoubleClickX;
            screenY = _mouseDoubleClickY;
            if (!_mouseDoubleClickPointActive
                || DateTimeOffset.UtcNow - _mouseDoubleClickAt
                    > TimeSpan.FromSeconds(1))
            {
                _mouseDoubleClickPointActive = false;
                return false;
            }
            return true;
        }
    }

    private void ClearNativeOleTarget()
    {
        lock (_nativeOleTargetGate)
        {
            _nativeOleTargetActive = false;
            _nativeOleTargetIsMathType = false;
            _nativeOleTargetLeft = 0;
            _nativeOleTargetTop = 0;
            _nativeOleTargetRight = 0;
            _nativeOleTargetBottom = 0;
            _nativeOleTargetRangeStart = -1;
            _nativeOleTargetRangeEnd = -1;
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
            // Startup must remain non-blocking. The first explicit Office action
            // retries the full diagnostic/startup path and reports any failure.
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
        if (lifetime is null || lifetime.IsCancellationRequested)
        {
            SetStatus("VisualTeX Word 插件正在重新初始化，请稍后再试。");
            WordDoubleClickHook.TraceMessage("ribbon-session-rejected-addin-lifetime-unavailable");
            return;
        }
        _ = ObserveRibbonSessionTaskAsync(RunSessionAsync(
            mode,
            displayMode,
            requestedObjectMode,
            capturedSelection,
            conversionOnly,
            lifetime.Token));
    }

    private async Task ObserveRibbonSessionTaskAsync(Task task)
    {
        try { await task.ConfigureAwait(false); }
        catch (OperationCanceledException)
        {
            // Add-in shutdown can cancel before RunSessionAsync acquires its gate.
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage($"ribbon-session-unobserved-failure: {error}");
            SetStatus($"VisualTeX Word 插件操作失败：{error.Message}");
        }
    }

    private void RegisterActiveSessionOperation(CancellationTokenSource operationCancellation)
    {
        lock (_activeSessionOperationGate)
        {
            _activeSessionCancellation = operationCancellation;
            _activeSessionId = null;
        }
    }

    private void SetActiveSessionOperationId(
        CancellationTokenSource operationCancellation,
        string sessionId)
    {
        lock (_activeSessionOperationGate)
        {
            if (!ReferenceEquals(_activeSessionCancellation, operationCancellation)) return;
            _activeSessionId = sessionId;
        }
    }

    private void ClearActiveSessionOperation(CancellationTokenSource operationCancellation)
    {
        lock (_activeSessionOperationGate)
        {
            if (!ReferenceEquals(_activeSessionCancellation, operationCancellation)) return;
            _activeSessionCancellation = null;
            _activeSessionId = null;
        }
    }

    private bool CancelStaleActiveSessionOperation(string expectedSessionId)
    {
        CancellationTokenSource? cancellation = null;
        lock (_activeSessionOperationGate)
        {
            if (!string.Equals(_activeSessionId, expectedSessionId, StringComparison.Ordinal)
                || _activeSessionCancellation is null)
                return false;
            cancellation = _activeSessionCancellation;
        }
        try { cancellation.Cancel(); }
        catch (ObjectDisposedException) { }
        return true;
    }

    private static bool IsMissingSessionError(Exception error) =>
        error.Message.IndexOf("(404)", StringComparison.OrdinalIgnoreCase) >= 0;

    private async Task<bool> TryRecoverBusySessionAsync(CancellationToken cancellationToken)
    {
        string? activeSessionId;
        lock (_activeSessionOperationGate) activeSessionId = _activeSessionId;
        var client = _sessionClient;
        if (string.IsNullOrWhiteSpace(activeSessionId) || client is null)
        {
            SetStatus("VisualTeX 正在准备编辑窗口，请稍候再试。");
            return false;
        }

        OfficeSessionDocument? activeSession = null;
        try
        {
            using var probeTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            probeTimeout.CancelAfter(TimeSpan.FromSeconds(2));
            activeSession = await client.GetSessionAsync(activeSessionId!, probeTimeout.Token)
                .ConfigureAwait(false);
        }
        catch (Exception error) when (IsMissingSessionError(error))
        {
            WordDoubleClickHook.TraceMessage(
                $"ribbon-stale-session-missing sessionId={activeSessionId}");
            if (!CancelStaleActiveSessionOperation(activeSessionId!)) return false;
            SetStatus("检测到已失效的 VisualTeX 编辑任务，正在自动恢复当前操作…");
            return await _operationGate.WaitAsync(
                    TimeSpan.FromSeconds(4),
                    cancellationToken)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            SetStatus("已有 VisualTeX 编辑任务正在响应，请稍候再试。");
            return false;
        }
        catch (Exception error)
        {
            SetStatus($"已有 VisualTeX 编辑任务，但状态检查失败：{error.Message}");
            return false;
        }

        if (activeSession.Status is "completed" or "cancelled" or "failed")
        {
            WordDoubleClickHook.TraceMessage(
                $"ribbon-stale-session-terminal sessionId={activeSessionId} status={activeSession.Status}");
            if (!CancelStaleActiveSessionOperation(activeSessionId!)) return false;
            SetStatus("检测到已结束但未释放的 VisualTeX 编辑任务，正在自动恢复当前操作…");
            return await _operationGate.WaitAsync(
                    TimeSpan.FromSeconds(4),
                    cancellationToken)
                .ConfigureAwait(false);
        }

        if (string.Equals(activeSession.Status, "committing", StringComparison.Ordinal))
        {
            SetStatus("VisualTeX 正在把上一条公式写入 Word，请稍候再试。");
            return false;
        }

        try
        {
            GrantVisualTeXForegroundActivation();
            using var openTimeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            openTimeout.CancelAfter(TimeSpan.FromSeconds(3));
            await client.OpenEditorAsync(activeSessionId!, openTimeout.Token).ConfigureAwait(false);
            SetStatus("已有 VisualTeX 编辑任务，已将编辑窗口切换到前台。");
        }
        catch (Exception error) when (IsMissingSessionError(error))
        {
            WordDoubleClickHook.TraceMessage(
                $"ribbon-stale-session-disappeared-before-open sessionId={activeSessionId}");
            if (!CancelStaleActiveSessionOperation(activeSessionId!)) return false;
            SetStatus("检测到编辑任务已失效，正在自动恢复当前操作…");
            return await _operationGate.WaitAsync(
                    TimeSpan.FromSeconds(4),
                    cancellationToken)
                .ConfigureAwait(false);
        }
        catch (Exception error)
        {
            // The Session still exists, so never cancel it merely because its
            // window failed to foreground: it may contain unsaved user edits.
            SetStatus($"已有编辑任务，但无法置前窗口：{error.Message}");
        }
        return false;
    }

    private async Task RunSessionAsync(
        string mode,
        string? requestedDisplayMode,
        string? requestedObjectMode,
        OfficeSelection? capturedSelection,
        bool conversionOnly,
        CancellationToken cancellationToken)
    {
        var lifetimeCancellationToken = cancellationToken;
        var operationGateAcquired = await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                cancellationToken)
            .ConfigureAwait(false);
        if (!operationGateAcquired)
        {
            operationGateAcquired = await TryRecoverBusySessionAsync(cancellationToken)
                .ConfigureAwait(false);
            if (!operationGateAcquired) return;
        }

        var operationCancellation =
            CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        RegisterActiveSessionOperation(operationCancellation);
        cancellationToken = operationCancellation.Token;

        var openPerformance = string.Equals(
            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
            "1",
            StringComparison.Ordinal)
            ? Stopwatch.StartNew()
            : null;
        long openPerformanceCheckpoint = 0;
        void TraceOpenPerformance(string stage)
        {
            if (openPerformance is null) return;
            var elapsed = openPerformance.ElapsedMilliseconds;
            Console.WriteLine(
                $"    [perf] OpenSession.{stage}: +{elapsed - openPerformanceCheckpoint}ms ({elapsed}ms total)");
            openPerformanceCheckpoint = elapsed;
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
            TraceOpenPerformance("health");
            var selection = capturedSelection?.Metadata is not null
                ? capturedSelection
                : await dispatcher.InvokeAsync(service.ReadSelection).ConfigureAwait(false);
            TraceOpenPerformance("read-selection");
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("请先选择一个 VisualTeX 公式。");

            // A create command may be invoked while the previous formula is
            // still selected. Only edit commands are allowed to seed the new
            // Session from that selection; every create Session starts blank.
            var metadata = mode == "edit"
                ? NormalizeEditableMetadata(selection.Metadata)
                : null;
            var targetObjectMode = requestedObjectMode
                ?? (mode == "create"
                    ? WordEquationNumbering.GetDefaultCreateObjectMode()
                    : selection.ObjectMode)
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
            var fontSizePt = metadata?.FontSizePt
                ?? await dispatcher.InvokeAsync(service.ReadCurrentTypingFontSize)
                    .ConfigureAwait(false);
            var effectiveDisplayMode =
                requestedDisplayMode ?? metadata?.DisplayMode ?? "inline";
            var mathTypeNumberPosition = mode == "create"
                ? await dispatcher.InvokeAsync(service.GetMathTypeNumberPositionPreference)
                    .ConfigureAwait(false)
                : string.Equals(
                        selection.ObjectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal)
                    ? await dispatcher.InvokeAsync(
                            () => service.GetMathTypeNumberPositionForRange(selection.ObjectId))
                        .ConfigureAwait(false)
                    : "right";
            var request = new CreateVstoSessionRequest
            {
                Mode = mode,
                Host = "word",
                FormulaId = metadata?.FormulaId,
                SourceDocumentId = selection.DocumentId,
                // Preserve the Word range captured when the editor opens for both
                // create and edit sessions. Office 2019 can move the live Selection
                // back into the preceding equation table while the external editor
                // owns focus; resolving the insertion point only at commit time then
                // inserts the new formula into the old numbered structure.
                SourceObjectId = selection.ObjectId,
                Title = metadata?.Title ?? "Word Formula",
                Lines = lines,
                ActiveLineId = lines.FirstOrDefault()?.Id,
                CodeFormat = metadata?.CodeFormat ?? "latex",
                DisplayMode = effectiveDisplayMode,
                ObjectMode = targetObjectMode,
                Numbered = effectiveDisplayMode == "block"
                    && (mode == "create"
                        ? WordEquationNumbering.GetDefaultDisplayEquationNumbered()
                        : metadata?.Numbered ?? false),
                MathTypeNumberPosition = mathTypeNumberPosition,
                FontSizePt = FormulaFontSize.Normalize(fontSizePt),
                OriginalMetadata = metadata,
                AutoCommitOnClose = true,
            };
            TraceOpenPerformance("build-request");
            var session = await client.CreateSessionAsync(request, cancellationToken).ConfigureAwait(false);
            TraceOpenPerformance("create-session");
            sessionId = session.Id;
            SetActiveSessionOperationId(operationCancellation, session.Id);
            if (conversionOnly)
            {
                await client.OpenConverterAsync(session.Id, cancellationToken)
                    .ConfigureAwait(false);
                SetStatus("正在直接转换 Word 公式格式…");
            }
            else
            {
                GrantVisualTeXForegroundActivation();
                await client.OpenEditorAsync(session.Id, cancellationToken)
                    .ConfigureAwait(false);
                TraceOpenPerformance("open-editor-window");
                SetStatus("VisualTeX 编辑器已打开。");
            }
            session = await client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromMinutes(30),
                cancellationToken).ConfigureAwait(false);
            if (session.Mode == "create" && session.DisplayMode == "block")
                WordEquationNumbering.SetDefaultDisplayEquationNumbered(session.Numbered);
            if (session.Mode == "create"
                && (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal)
                    || string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal)))
                WordEquationNumbering.SetDefaultCreateObjectMode(session.ObjectMode);
            if (session.Mode == "create"
                && session.Numbered
                && string.Equals(session.DisplayMode, "block", StringComparison.Ordinal)
                && string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
                WordEquationNumbering.SetDefaultMathTypeNumberPosition(
                    session.MathTypeNumberPosition);
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
                    StringComparison.Ordinal)
                || string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                var requiredMathMl = export.MathMl;
                if (string.IsNullOrWhiteSpace(requiredMathMl)
                    || !requiredMathMl!.TrimStart().StartsWith("<math", StringComparison.Ordinal))
                    throw new InvalidDataException(
                        $"VisualTeX Session has no valid MathML result for {session.ObjectMode}.");
                mathMl = requiredMathMl;
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal))
                {
                    svgPath = client.MaterializeSvg(session);
                    var preview = OfficeOlePreview.CreateInkSafeVectorPreviewFromSvg(
                        svgPath,
                        export.Width,
                        export.Height,
                        export.Baseline,
                        safetyPaddingPixels: MathTypePreviewHorizontalSafetyInsetPixels);
                    emfPath = preview.EmfPath;
                    // MathType fallback SVGs replace MathJax's original TeX glyphs
                    // with Times outlines after MathJax has already chosen its SVG
                    // viewBox.  Re-measure the actual GDI+ outlines before creating
                    // the offline Equation.DSMT4 presentation; otherwise wider Times
                    // letters can be clipped or visually squeezed into the old TeX
                    // geometry on machines without a native MathPage renderer.
                    export.Width = preview.WidthPixels;
                    export.Height = preview.HeightPixels;
                    export.Baseline = preview.BaselinePixels;
                }
            }
            else if (string.Equals(
                         session.ObjectMode,
                         FormulaOleContract.NativeOleMode,
                         StringComparison.Ordinal))
            {
                imagePath = client.MaterializePng(session);
                svgPath = client.MaterializeSvg(session);
                var originalImagePath = imagePath;
                var preview = OfficeOlePreview.CreateInkSafePreviewFromSvg(
                    svgPath,
                    export.Width,
                    export.Height,
                    export.Baseline,
                    fallbackPngPath: originalImagePath);
                if (!string.Equals(
                        preview.PngPath,
                        originalImagePath,
                        StringComparison.OrdinalIgnoreCase))
                    TryDeleteFile(originalImagePath);
                emfPath = preview.EmfPath;
                imagePath = preview.PngPath;
                // Keep the OLE physical frame, vector/raster previews and
                // mathematical baseline on one geometry after final GDI+ font
                // outline measurement.
                export.Width = preview.WidthPixels;
                export.Height = preview.HeightPixels;
                export.Baseline = preview.BaselinePixels;
            }
            else
            {
                imagePath = client.MaterializePng(session);
            }
            // Word can fire WindowSelectionChange synchronously while an edit moves
            // an OLE formula into/out of the numbered 1x3 host.  SelectionChange
            // performs formula-identity repair, which is correct for user copy/paste
            // but must not inspect a formula in the middle of this structural write:
            // the VTO_ identity bookmark and the live InlineShape can temporarily
            // occupy different ranges while ConvertToTable is committing.  Reuse
            // the same mutation guard as format conversion so event callbacks stay
            // read-free until the formula write has fully settled.
            BeginFormulaFormatMutation();
            try
            {
                await dispatcher.InvokeAsync(() =>
                {
                    var activeDocumentId = service.ReadActiveDocumentId();
                if (!string.Equals(
                        activeDocumentId,
                        session.SourceDocumentId,
                        StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("活动 Word 文档已切换，未写入公式。");
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal))
                {
                    if (mathMl is null)
                        throw new InvalidOperationException(
                            "VisualTeX MathType OLE MathML payload is unavailable.");
                    if (emfPath is null)
                        throw new InvalidOperationException(
                            "VisualTeX MathType OLE vector preview is unavailable.");
                    return session.Mode == "edit"
                        ? service.ReplaceMathTypeOle(session, mathMl, emfPath)
                        : service.InsertMathTypeOle(session, mathMl, emfPath);
                }
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    if (mathMl is null)
                        throw new InvalidOperationException(
                            "VisualTeX Word OMML MathML payload is unavailable.");
                    // InsertOmml/ReplaceOmml synchronously validate the exact
                    // FormulaId-bound 1x3 host before returning. Do not enqueue the
                    // retired five-turn whole-document Shape finalizer: every later
                    // mouse click would otherwise contend with those UI-thread scans.
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
            }
            finally
            {
                EndFormulaFormatMutation();
            }
            await client.CompleteAsync(session.Id, cancellationToken).ConfigureAwait(false);
            if (string.Equals(
                    session.ObjectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
                SetStatus(session.Mode == "edit"
                    ? "MathType OLE 公式已原位更新，仍可继续用 MathType 编辑。"
                    : "MathType OLE 公式已插入，可继续用 MathType 或 VisualTeX 编辑。");
            else if (requiresObjectModeChange
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
            if (lifetimeCancellationToken.IsCancellationRequested)
                SetStatus("VisualTeX 操作已取消。");
            else if (operationCancellation.IsCancellationRequested)
                WordDoubleClickHook.TraceMessage(
                    $"ribbon-stale-session-local-waiter-cancelled sessionId={sessionId ?? "<pending>"}");
            else
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
            ClearActiveSessionOperation(operationCancellation);
            operationCancellation.Dispose();
            _operationGate.Release();
            WordDoubleClickHook.TraceMessage(
                $"ribbon-session-operation-released sessionId={sessionId ?? "<pending>"}");
        }
    }

    private static FormulaMetadata? NormalizeEditableMetadata(FormulaMetadata? source)
    {
        if (source is null) return null;
        var metadata = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(source))
            ?? throw new InvalidDataException("Unable to clone VisualTeX formula metadata.");
        if (metadata.Lines.Count == 0) return metadata;

        var last = metadata.Lines[metadata.Lines.Count - 1];
        var split = FormulaEquationTag.Extract(last.Latex);
        if (!string.Equals(last.Latex, split.Latex, StringComparison.Ordinal))
            last.Latex = split.Latex;
        metadata.EquationTag ??= split.EquationTag;
        metadata.Latex = string.Join("\n", metadata.Lines.Select(line => line.Latex));
        if (!string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            metadata.EquationTag = null;
        metadata.Validate();
        return metadata;
    }

    private async Task RedrawLatexAsync(bool wholeDocument, string objectMode)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = _sessionClient;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || client is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(
            StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(
            StringComparer.Ordinal);
        var mathTypePreviews =
            new Dictionary<string, MathTypeNativePreviewRenderer.Result>(
                StringComparer.Ordinal);
        var converterSessionIds = new List<string>();
        var renderFailures = new Dictionary<string, string>(StringComparer.Ordinal);
        var skippedTargets = new List<(WordLatexRedrawTarget Target, string Error)>();
        var maxRenderMilliseconds = 0L;
        var totalRenderMilliseconds = 0L;
        try
        {
            var plan = await dispatcher.InvokeAsync(
                    () => service.CaptureLatexRedrawPlan(wholeDocument))
                .ConfigureAwait(false);
            var modeLabel = string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
                ? "VisualTeX OLE"
                : string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                    ? "MathType"
                    : "Word OMML";
            var displayFormulaCount = plan.Targets.Count(target => string.Equals(
                target.DisplayMode,
                "block",
                StringComparison.Ordinal));
            var acceptanceMode = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal);
            var allowRedrawNumbering = !string.Equals(
                objectMode,
                FormulaOleContract.WordOmmlMode,
                StringComparison.Ordinal);
            var numberDisplayFormulas = allowRedrawNumbering
                && acceptanceMode
                && IsEnabledEnvironmentOption(
                    "VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS");
            if (!acceptanceMode)
            {
                var options = await dispatcher.InvokeAsync(() =>
                {
                    using var dialog = new LatexRedrawDialog(
                        wholeDocument,
                        plan.Targets.Count,
                        displayFormulaCount,
                        modeLabel,
                        service.GetEquationNumberFormatDisplayName(),
                        allowRedrawNumbering);
                    var accepted = dialog.ShowDialog()
                        == System.Windows.Forms.DialogResult.OK;
                    return (
                        Accepted: accepted,
                        NumberDisplayFormulas: accepted
                            && dialog.NumberDisplayFormulas);
                }).ConfigureAwait(false);
                if (!options.Accepted)
                {
                    SetStatus(
                        wholeDocument
                            ? "已取消全文 LaTeX 重绘，Word 文档未修改。"
                            : "已取消所选 LaTeX 重绘，Word 文档未修改。");
                    return;
                }
                numberDisplayFormulas = options.NumberDisplayFormulas;
            }
            plan.NumberDisplayFormulas = allowRedrawNumbering && numberDisplayFormulas;
            var mathTypeNumberPosition = string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                ? await dispatcher.InvokeAsync(
                        service.GetMathTypeNumberPositionPreference)
                    .ConfigureAwait(false)
                : "right";

            WriteRedrawAcceptanceLog(
                $"redraw-start scope={(wholeDocument ? "document" : "selection")} "
                + $"mode={objectMode} formulas={plan.Targets.Count} "
                + $"numberDisplayFormulas={plan.NumberDisplayFormulas}");
            SetStatus($"正在准备重绘 {plan.Targets.Count} 个 LaTeX 公式为 {modeLabel}…");
            await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
            await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);

            // Redraw used to switch the single hidden converter WebView from one
            // Session to the next and wait after every switch. React could finish
            // the previous conversion after the URL/session state had already
            // advanced, causing the first formula's exportResult (especially its
            // MathML) to be written into later Sessions. Bulk import already avoids
            // that race by queuing explicit Session ids and patching each id
            // directly; use the same path for redraw.
            var formulaKeys = new Dictionary<string, string>(StringComparer.Ordinal);
            var pendingKeys = new HashSet<string>(StringComparer.Ordinal);
            var pendingTemplates = new List<(
                string Key,
                WordBulkRun Run,
                WordLatexRedrawTarget Target,
                OfficeSessionDocument Session,
                int Index)>();

            for (var index = 0; index < plan.Targets.Count; index++)
            {
                lifetime.Token.ThrowIfCancellationRequested();
                var target = plan.Targets[index];
                var run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                };
                var key = string.Join(
                    "\u001F",
                    objectMode,
                    target.DisplayMode,
                    target.FontSizePt.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    target.Latex);
                formulaKeys.Add(target.Id, key);
                if (rendered.ContainsKey(key)
                    || renderFailures.ContainsKey(key)
                    || !pendingKeys.Add(key))
                    continue;

                SetStatus($"正在准备公式 {index + 1}/{plan.Targets.Count}…");
                try
                {
                    var conversionSession = await CreateBulkFormulaConversionSessionAsync(
                            client,
                            run,
                            objectMode,
                            plan.DocumentId,
                            target.FontSizePt,
                            lifetime.Token)
                        .ConfigureAwait(false);
                    pendingTemplates.Add((key, run, target, conversionSession, index));
                    converterSessionIds.Add(conversionSession.Id);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception error)
                {
                    var detail = string.IsNullOrWhiteSpace(error.Message)
                        ? error.GetType().Name
                        : error.Message.Trim();
                    renderFailures[key] = detail;
                    WriteRedrawAcceptanceLog(
                        $"render-skipped index={index + 1} elapsedMs=0 "
                        + $"fontSizePt={target.FontSizePt:0.##} display={target.DisplayMode} "
                        + $"latex={target.Latex} error={detail}");
                }
            }

            if (pendingTemplates.Count > 0)
            {
                WriteRedrawAcceptanceLog(
                    $"render-batch-start unique={pendingTemplates.Count} total={plan.Targets.Count}");
                SetStatus($"正在批量渲染 {pendingTemplates.Count} 个独立公式…");
                await client.OpenConverterBatchAsync(
                        pendingTemplates.Select(item => item.Session.Id).ToList(),
                        lifetime.Token)
                    .ConfigureAwait(false);

                foreach (var pending in pendingTemplates)
                {
                    lifetime.Token.ThrowIfCancellationRequested();
                    var stopwatch = Stopwatch.StartNew();
                    try
                    {
                        var completedSession = await client.WaitForCommitAsync(
                                pending.Session.Id,
                                TimeSpan.FromMinutes(3),
                                lifetime.Token)
                            .ConfigureAwait(false);
                        var template = MaterializeBulkFormulaTemplate(
                            client,
                            pending.Run,
                            objectMode,
                            completedSession);
                        stopwatch.Stop();
                        totalRenderMilliseconds += stopwatch.ElapsedMilliseconds;
                        maxRenderMilliseconds = Math.Max(
                            maxRenderMilliseconds,
                            stopwatch.ElapsedMilliseconds);
                        rendered.Add(pending.Key, template);
                        WriteRedrawAcceptanceLog(
                            $"render index={pending.Index + 1} elapsedMs={stopwatch.ElapsedMilliseconds} "
                            + $"fontSizePt={pending.Target.FontSizePt:0.##} display={pending.Target.DisplayMode} "
                            + $"latex={pending.Target.Latex}");
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception error)
                    {
                        stopwatch.Stop();
                        var detail = string.IsNullOrWhiteSpace(error.Message)
                            ? error.GetType().Name
                            : error.Message.Trim();
                        renderFailures[pending.Key] = detail;
                        WriteRedrawAcceptanceLog(
                            $"render-skipped index={pending.Index + 1} elapsedMs={stopwatch.ElapsedMilliseconds} "
                            + $"fontSizePt={pending.Target.FontSizePt:0.##} display={pending.Target.DisplayMode} "
                            + $"latex={pending.Target.Latex} error={detail}");
                    }
                }
            }

            if (string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                && rendered.Count > 0)
            {
                // Prefer MathType's native MathPage geometry when the optional
                // renderer is available. Machines without MathType already have a
                // vector EMF from the converter, so the whole batch falls back to
                // that existing presentation instead of failing or launching one
                // native renderer attempt per formula during Word insertion.
                var nativePreviewAvailable = MathTypeNativePreviewRenderer.IsAvailable;
                if (!nativePreviewAvailable)
                {
                    SetStatus("未检测到 MathType 原生预览器，正在使用 VisualTeX 矢量预览写入 MathType 公式…");
                    WriteRedrawAcceptanceLog(
                        $"mathtype-preview-fallback reason=native-unavailable templates={rendered.Count} "
                        + $"formulas={plan.Targets.Count}");
                }
                else
                {
                    SetStatus($"正在批量生成 {rendered.Count} 个 MathType 原生预览…");
                    var nativePreviewInputs =
                        new Dictionary<string, byte[]>(StringComparer.Ordinal);
                    foreach (var item in rendered)
                    {
                        var standalone = WordFormulaService.PrepareStandaloneMathTypeOleData(
                            item.Value.MathMl
                                ?? throw new InvalidDataException(
                                    $"MathType 重绘模板 {item.Key} 缺少 MathML。"),
                            string.Equals(
                                item.Value.Session.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase),
                            item.Value.Session.FontSizePt,
                            item.Value.Session.ToMetadata().Latex);
                        nativePreviewInputs[item.Key] = standalone.Generated.Mtef;
                    }

                    var nativePreviewRoot = rendered.Values
                        .Select(template => string.IsNullOrWhiteSpace(template.EmfPath)
                            ? null
                            : Path.GetDirectoryName(template.EmfPath))
                        .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path))
                        ?? Path.GetTempPath();
                    var nativePreviewWatch = Stopwatch.StartNew();
                    var renderedAllNativePreviews =
                        MathTypeNativePreviewRenderer.TryRenderBatch(
                            nativePreviewInputs,
                            nativePreviewRoot,
                            out var nativePreviews);
                    var missingPreviewKeys = rendered.Keys
                        .Where(key => !nativePreviews.ContainsKey(key))
                        .ToArray();
                    if (!renderedAllNativePreviews || missingPreviewKeys.Length > 0)
                    {
                        foreach (var preview in nativePreviews.Values)
                            preview.Dispose();
                        nativePreviewWatch.Stop();
                        SetStatus("MathType 原生预览未完整生成，正在整批使用 VisualTeX 矢量预览…");
                        WriteRedrawAcceptanceLog(
                            $"mathtype-preview-fallback reason=native-batch-incomplete "
                            + $"native={nativePreviews.Count}/{rendered.Count} formulas={plan.Targets.Count} "
                            + $"elapsedMs={nativePreviewWatch.ElapsedMilliseconds}");
                    }
                    else
                    {
                        foreach (var preview in nativePreviews)
                            mathTypePreviews.Add(preview.Key, preview.Value);
                        nativePreviewWatch.Stop();
                        WriteRedrawAcceptanceLog(
                            $"mathtype-native-preview-batch templates={nativePreviews.Count} "
                            + $"formulas={plan.Targets.Count} elapsedMs={nativePreviewWatch.ElapsedMilliseconds}");
                    }
                }
            }

            for (var index = 0; index < plan.Targets.Count; index++)
            {
                var target = plan.Targets[index];
                var key = formulaKeys[target.Id];
                if (renderFailures.TryGetValue(key, out var failure))
                {
                    skippedTargets.Add((target, failure));
                    WriteRedrawAcceptanceLog(
                        $"render-skip-cache-hit index={index + 1} display={target.DisplayMode} "
                        + $"latex={target.Latex} error={failure}");
                    continue;
                }
                if (!rendered.TryGetValue(key, out var template))
                    throw new InvalidDataException($"缺少公式 {target.Id} 的批量渲染结果。");

                var run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                };
                var independentSession = CloneBulkFormulaSession(
                    template.Session,
                    run,
                    plan.DocumentId,
                    target.FontSizePt,
                    objectMode,
                    mathTypeNumberPosition);
                var mathTypePreview = mathTypePreviews.TryGetValue(key, out var preview)
                    ? preview
                    : null;
                prepared.Add(target.Id, new PreparedWordBulkFormula
                {
                    Run = run,
                    Session = independentSession,
                    MathMl = template.MathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                    MathTypeNativePreview = mathTypePreview,
                    MathTypeNativePreviewAttempted = string.Equals(
                        objectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal),
                });
            }

            if (prepared.Count == 0)
            {
                WriteRedrawAcceptanceLog(
                    $"redraw-complete formulas=0 unique=0 skipped={skippedTargets.Count} "
                    + "renderAverageMs=0 renderMaxMs=0 insertTotalMs=0 insertMaxMs=0");
                SetStatus(
                    $"LaTeX 重绘未转换任何公式：{skippedTargets.Count} 个公式无法解析，均已保留原代码。");
                return;
            }

            plan.Targets = plan.Targets
                .Where(target => prepared.ContainsKey(target.Id))
                .ToList();
            SetStatus("公式渲染完成，正在原位写入 Word…");
            BeginFormulaFormatMutation();
            WordLatexRedrawResult result;
            try
            {
                result = await dispatcher.InvokeAsync(
                        () => service.ApplyLatexRedrawPlan(plan, prepared))
                    .ConfigureAwait(false);
            }
            finally { EndFormulaFormatMutation(); }
            if (string.Equals(
                    objectMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                service.TrackCompletedOmmlRedrawForUndo(
                    plan.DocumentId,
                    result.FormulaIds);
                EnsureOmmlRedrawUndoCleanupWatcher();
            }
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.CompleteAsync(sessionId, lifetime.Token)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            var uniqueRenderCount = Math.Max(1, rendered.Count);
            var averageRenderMilliseconds = totalRenderMilliseconds / uniqueRenderCount;
            WriteRedrawAcceptanceLog(
                $"redraw-complete formulas={result.FormulaCount} unique={rendered.Count} "
                + $"skipped={skippedTargets.Count} "
                + $"renderAverageMs={averageRenderMilliseconds} renderMaxMs={maxRenderMilliseconds} "
                + $"insertTotalMs={result.TotalInsertMilliseconds} insertMaxMs={result.MaxInsertMilliseconds}");
            var performanceSuffix = maxRenderMilliseconds <= 250
                ? $"渲染最大 {maxRenderMilliseconds} ms/公式"
                : $"渲染最大 {maxRenderMilliseconds} ms/公式（本机超过 250 ms 目标）";
            var skippedSuffix = skippedTargets.Count == 0
                ? string.Empty
                : $"；{skippedTargets.Count} 个无法解析的公式已保留原 LaTeX";
            SetStatus(
                $"LaTeX 重绘完成：{result.FormulaCount} 个公式已转换为 {modeLabel}{skippedSuffix}；{performanceSuffix}。");
        }
        catch (OperationCanceledException error)
        {
            WriteRedrawAcceptanceLog("redraw-cancelled " + error);
            SetStatus("VisualTeX LaTeX 重绘已取消。");
        }
        catch (Exception error)
        {
            WriteRedrawAcceptanceLog("redraw-failed " + error);
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.FailAsync(sessionId, error.Message, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            SetStatus($"VisualTeX LaTeX 重绘失败：{error.Message}");
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            OfficePluginLanguage.SafeErrorMessage(error.Message),
                            T("VisualTeX LaTeX 重绘", "VisualTeX LaTeX Redraw"),
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch { }
            }
        }
        finally
        {
            foreach (var preview in mathTypePreviews.Values.Distinct())
                preview.Dispose();
            foreach (var template in rendered.Values)
            {
                TryDeleteFile(template.EmfPath);
                TryDeleteFile(template.SvgPath);
                TryDeleteFile(template.PngPath);
            }
            _operationGate.Release();
        }
    }

    private async Task ConvertFormulaObjectsToLatexAsync(
        bool wholeDocument,
        string objectMode)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }

        var modeLabel = string.Equals(
            objectMode,
            FormulaOleContract.NativeOleMode,
            StringComparison.Ordinal)
            ? "VisualTeX OLE"
            : string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
                ? "MathType"
                : "Word OMML";
        try
        {
            var count = await dispatcher.InvokeAsync(
                    () => service.CountFormulaObjectsForLatex(
                        wholeDocument,
                        objectMode))
                .ConfigureAwait(false);
            if (count == 0)
                throw new InvalidDataException(
                    wholeDocument
                        ? $"当前 Word 文档中没有找到 {modeLabel} 公式。"
                        : $"所选内容中没有找到 {modeLabel} 公式。");

            if (wholeDocument
                && !string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                var confirmed = await dispatcher.InvokeAsync(() =>
                    System.Windows.Forms.MessageBox.Show(
                        T(
                            $"将把当前文档中的 {count} 个 {modeLabel} 公式原位恢复为 LaTeX 代码。\r\n\r\n另一种公式对象不会被修改；该操作可通过一次 Ctrl+Z 整体撤销。是否继续？",
                            $"Restore {count} {modeLabel} equations in the current document to LaTeX source in place?\r\n\r\nOther equation object types will not be changed. The operation can be undone with a single Ctrl+Z."),
                        T("VisualTeX 公式转为 LaTeX", "VisualTeX Equation to LaTeX"),
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxDefaultButton.Button2)
                    == System.Windows.Forms.DialogResult.Yes).ConfigureAwait(false);
                if (!confirmed)
                {
                    SetStatus($"已取消全文 {modeLabel} 公式转为 LaTeX，Word 文档未修改。");
                    return;
                }
            }

            WriteRedrawAcceptanceLog(
                $"formula-to-latex-start scope={(wholeDocument ? "document" : "selection")} "
                + $"mode={objectMode} formulas={count}");
            SetStatus($"正在把 {count} 个 {modeLabel} 公式恢复为 LaTeX 代码…");
            var result = await dispatcher.InvokeAsync(
                    () => service.ConvertFormulaObjectsToLatex(
                        wholeDocument,
                        objectMode))
                .ConfigureAwait(false);
            WriteRedrawAcceptanceLog(
                $"formula-to-latex-complete scope={(wholeDocument ? "document" : "selection")} "
                + $"mode={objectMode} formulas={result.FormulaCount}");
            SetStatus(
                $"公式转为 LaTeX 完成：{result.FormulaCount} 个 {modeLabel} 公式已恢复为源码。");
        }
        catch (OperationCanceledException error)
        {
            WriteRedrawAcceptanceLog("formula-to-latex-cancelled " + error);
            SetStatus("VisualTeX 公式转为 LaTeX 已取消。");
        }
        catch (Exception error)
        {
            WriteRedrawAcceptanceLog("formula-to-latex-failed " + error);
            SetStatus($"VisualTeX 公式转为 LaTeX 失败：{error.Message}");
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            OfficePluginLanguage.SafeErrorMessage(error.Message),
                            T("VisualTeX 公式转为 LaTeX", "VisualTeX Equation to LaTeX"),
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch { }
            }
        }
        finally { _operationGate.Release(); }
    }

    private async Task BulkImportAsync()
    {
        WriteBulkAcceptanceLog("bulk-import-start");
        var dispatcher = _dispatcher;
        var service = _formulaService;
        var client = _sessionClient;
        var application = _application;
        var lifetime = _lifetime;
        if (dispatcher is null
            || service is null
            || client is null
            || application is null
            || lifetime is null
            || lifetime.IsCancellationRequested)
            return;

        if (!await _operationGate.WaitAsync(
                TimeSpan.FromSeconds(2),
                lifetime.Token).ConfigureAwait(false))
        {
            SetStatus("VisualTeX 正在执行其他 Word 操作，请稍候再试。");
            return;
        }

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(
            StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(
            StringComparer.Ordinal);
        var mathTypePreviews =
            new Dictionary<string, MathTypeNativePreviewRenderer.Result>(
                StringComparer.Ordinal);
        var ownedSessionIds = new List<string>();
        var operationStopwatch = Stopwatch.StartNew();
        var operationGateHeld = true;
        void ReleaseOperationGate()
        {
            if (!operationGateHeld) return;
            operationGateHeld = false;
            _operationGate.Release();
            WriteBulkAcceptanceLog("bulk-operation-gate-released");
        }
        try
        {
            var selection = await dispatcher.InvokeAsync(service.ReadSelection)
                .ConfigureAwait(false);
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
            var sourceDocumentId = selection.DocumentId;
            var fontSizePt = FormulaFontSize.Normalize(
                await dispatcher.InvokeAsync(service.ReadCurrentTypingFontSize)
                    .ConfigureAwait(false));
            var resolvedImport = await ResolveBulkImportDocumentAsync(
                    client,
                    sourceDocumentId,
                    selection.ObjectId,
                    fontSizePt,
                    ownedSessionIds.Add,
                    lifetime.Token)
                .ConfigureAwait(false);
            var document = resolvedImport.Document;
            if (document is null)
            {
                WriteBulkAcceptanceLog("bulk-import-cancelled-no-document");
                SetStatus("已取消批量导入，Word 文档未修改。");
                return;
            }

            WriteBulkAcceptanceLog(
                $"parsed blocks={document.Blocks.Count} formulas={document.FormulaCount} "
                + $"mode={document.FormulaObjectMode} fontSizePt={fontSizePt:0.##}");
            SetStatus(
                $"正在准备批量导入：{document.Blocks.Count} 个块，{document.FormulaCount} 个公式…");
            await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
            await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);
            var formulaRuns = document.Blocks
                .SelectMany(block => block.Runs)
                .Where(run => run.IsFormula)
                .ToList();
            var objectMode = document.FormulaObjectMode switch
            {
                WordBulkFormulaObjectMode.Ole => FormulaOleContract.NativeOleMode,
                WordBulkFormulaObjectMode.MathType => FormulaOleContract.MathTypeOleMode,
                _ => FormulaOleContract.WordOmmlMode,
            };
            var mathTypeNumberPosition = string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                ? await dispatcher.InvokeAsync(
                        service.GetMathTypeNumberPositionPreference)
                    .ConfigureAwait(false)
                : "right";
            var pendingKeys = new HashSet<string>(StringComparer.Ordinal);
            var formulaKeys = new Dictionary<string, string>(StringComparer.Ordinal);
            var pendingTemplates = new List<(
                string Key,
                WordBulkRun Run,
                OfficeSessionDocument Session)>();

            for (var index = 0; index < formulaRuns.Count; index++)
            {
                lifetime.Token.ThrowIfCancellationRequested();
                var run = formulaRuns[index];
                var key = string.Join(
                    "\u001F",
                    objectMode,
                    run.DisplayMode,
                    fontSizePt.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    run.Latex,
                    run.EquationTag ?? string.Empty);
                formulaKeys.Add(run.Id, key);
                if (rendered.ContainsKey(key) || !pendingKeys.Add(key))
                    continue;

                WriteBulkAcceptanceLog(
                    $"render-prepare index={index + 1}/{formulaRuns.Count} "
                    + $"display={run.DisplayMode} latex={run.Latex}");
                SetStatus($"正在准备公式 {index + 1}/{formulaRuns.Count}…");
                var conversionSession = await CreateBulkFormulaConversionSessionAsync(
                        client,
                        run,
                        objectMode,
                        sourceDocumentId,
                        fontSizePt,
                        lifetime.Token)
                    .ConfigureAwait(false);
                pendingTemplates.Add((key, run, conversionSession));
                ownedSessionIds.Add(conversionSession.Id);
            }

            if (pendingTemplates.Count > 0)
            {
                WriteBulkAcceptanceLog(
                    $"render-batch-start unique={pendingTemplates.Count} total={formulaRuns.Count}");
                SetStatus($"正在批量渲染 {pendingTemplates.Count} 个独立公式…");
                var renderStopwatch = Stopwatch.StartNew();
                await client.OpenConverterBatchAsync(
                        pendingTemplates.Select(item => item.Session.Id).ToList(),
                        lifetime.Token)
                    .ConfigureAwait(false);
                foreach (var pending in pendingTemplates)
                {
                    var completedSession = await client.WaitForCommitAsync(
                            pending.Session.Id,
                            TimeSpan.FromMinutes(3),
                            lifetime.Token)
                        .ConfigureAwait(false);
                    var template = MaterializeBulkFormulaTemplate(
                        client,
                        pending.Run,
                        objectMode,
                        completedSession);
                    rendered.Add(pending.Key, template);
                    WriteBulkAcceptanceLog(
                        $"render-batch-item sessionId={completedSession.Id} "
                        + $"status={completedSession.Status} display={pending.Run.DisplayMode}");
                }
                renderStopwatch.Stop();
                WriteBulkAcceptanceLog(
                    $"render-batch-complete unique={pendingTemplates.Count} "
                    + $"elapsedMs={renderStopwatch.ElapsedMilliseconds}");
            }

            if (string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal))
            {
                var nativePreviewAvailable = MathTypeNativePreviewRenderer.IsAvailable;
                if (!nativePreviewAvailable)
                {
                    SetStatus("未检测到 MathType 原生预览器，正在使用 VisualTeX 矢量预览导入 MathType 公式…");
                    WriteBulkAcceptanceLog(
                        $"mathtype-preview-fallback reason=native-unavailable templates={rendered.Count} "
                        + $"formulas={formulaRuns.Count}");
                }
                else
                {
                    SetStatus($"正在批量生成 {formulaRuns.Count} 个 MathType 原生预览…");
                    var nativePreviewInputs =
                        new Dictionary<string, byte[]>(StringComparer.Ordinal);
                    foreach (var item in rendered)
                    {
                        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
                            item.Value.MathMl,
                            string.Equals(
                                item.Value.Session.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase),
                            item.Value.Session.FontSizePt);
                        nativePreviewInputs[item.Key] = generated.Mtef;
                    }

                    var nativePreviewRoot = rendered.Values
                        .Select(template => string.IsNullOrWhiteSpace(template.EmfPath)
                            ? null
                            : Path.GetDirectoryName(template.EmfPath))
                        .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path))
                        ?? Path.GetTempPath();
                    var nativePreviewWatch = Stopwatch.StartNew();
                    var renderedAllNativePreviews =
                        MathTypeNativePreviewRenderer.TryRenderBatch(
                            nativePreviewInputs,
                            nativePreviewRoot,
                            out var nativePreviews);
                    var missingPreviewKeys = rendered.Keys
                        .Where(key => !nativePreviews.ContainsKey(key))
                        .ToArray();
                    if (!renderedAllNativePreviews || missingPreviewKeys.Length > 0)
                    {
                        foreach (var preview in nativePreviews.Values)
                            preview.Dispose();
                        nativePreviewWatch.Stop();
                        SetStatus("MathType 原生预览未完整生成，正在整批使用 VisualTeX 矢量预览…");
                        WriteBulkAcceptanceLog(
                            $"mathtype-preview-fallback reason=native-batch-incomplete "
                            + $"native={nativePreviews.Count}/{rendered.Count} formulas={formulaRuns.Count} "
                            + $"elapsedMs={nativePreviewWatch.ElapsedMilliseconds}");
                    }
                    else
                    {
                        foreach (var preview in nativePreviews)
                            mathTypePreviews.Add(preview.Key, preview.Value);
                        nativePreviewWatch.Stop();
                        WriteBulkAcceptanceLog(
                            $"mathtype-native-preview-batch templates={nativePreviews.Count} "
                            + $"formulas={formulaRuns.Count} elapsedMs={nativePreviewWatch.ElapsedMilliseconds}");
                    }
                }
            }

            foreach (var run in formulaRuns)
            {
                var key = formulaKeys[run.Id];
                var template = rendered[key];
                var independentSession = CloneBulkFormulaSession(
                    template.Session,
                    run,
                    sourceDocumentId,
                    fontSizePt,
                    objectMode,
                    mathTypeNumberPosition);
                var mathTypePreview = mathTypePreviews.TryGetValue(key, out var preview)
                    ? preview
                    : null;
                prepared.Add(run.Id, new PreparedWordBulkFormula
                {
                    Run = run,
                    Session = independentSession,
                    MathMl = template.MathMl,
                    PngPath = template.PngPath,
                    EmfPath = template.EmfPath,
                    MathTypeNativePreview = mathTypePreview,
                    MathTypeNativePreviewAttempted = string.Equals(
                        objectMode,
                        FormulaOleContract.MathTypeOleMode,
                        StringComparison.Ordinal),
                });
            }

            SetStatus("公式渲染完成，正在写入 Word…");
            var insertStopwatch = Stopwatch.StartNew();
            var result = await dispatcher.InvokeAsync(() =>
                    service.InsertBulkDocument(
                        document,
                        prepared,
                        sourceDocumentId,
                        selection.ObjectId))
                .ConfigureAwait(false);
            insertStopwatch.Stop();
            WriteBulkAcceptanceLog(
                $"bulk-insert-complete blocks={result.BlockCount} "
                + $"formulas={result.FormulaCount} elapsedMs={insertStopwatch.ElapsedMilliseconds}");
            operationStopwatch.Stop();
            WriteBulkAcceptanceLog(
                $"bulk-import-complete blocks={result.BlockCount} formulas={result.FormulaCount} "
                + $"elapsedMs={operationStopwatch.ElapsedMilliseconds}");
            SetStatus(
                $"批量导入完成：{result.BlockCount} 个内容块，{result.FormulaCount} 个独立公式；"
                + $"耗时 {operationStopwatch.Elapsed.TotalSeconds:0.0} 秒。");

            // Session finalization and temporary-file cleanup are companion-side
            // bookkeeping. They must never keep Word's single-operation gate
            // locked after the document has already been modified successfully.
            // A stalled local PATCH previously made every later Ribbon command
            // report that another Word operation was still running forever.
            ReleaseOperationGate();
            try
            {
                await FinalizeOfficeSessionsAsync(client, ownedSessionIds, completed: true, error: null)
                    .ConfigureAwait(false);
            }
            catch (Exception completionError)
            {
                var message = $"Word 批量导入已完成，但会话完成状态未能记录：{completionError.Message}";
                WriteBulkAcceptanceLog("bulk-import-session-finalization-failed " + completionError);
                SetStatus(message);
                await ShowBulkImportMessageAsync(dispatcher, message, warning: true).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException error)
        {
            var reported = await RecordSessionFailureAsync(client, ownedSessionIds, error).ConfigureAwait(false);
            WriteBulkAcceptanceLog("bulk-import-cancelled " + reported);
            SetStatus(ReferenceEquals(reported, error) ? "VisualTeX 批量导入已取消。" : reported.Message);
        }
        catch (Exception error)
        {
            var reported = await RecordSessionFailureAsync(client, ownedSessionIds, error).ConfigureAwait(false);
            WriteBulkAcceptanceLog("bulk-import-failed " + reported);
            SetStatus($"VisualTeX 批量导入失败：{reported.Message}");
            ReleaseOperationGate();
            await ShowBulkImportMessageAsync(dispatcher, reported.Message, warning: false).ConfigureAwait(false);
        }
        finally
        {
            ReleaseOperationGate();
            foreach (var template in rendered.Values)
            {
                TryDeleteFile(template.EmfPath);
                TryDeleteFile(template.SvgPath);
                TryDeleteFile(template.PngPath);
            }
            foreach (var preview in mathTypePreviews.Values.Distinct())
                preview.Dispose();
        }
    }

    private static async Task FinalizeOfficeSessionsAsync(
        VisualTeXSessionClient client,
        IReadOnlyCollection<string> ownedSessionIds,
        bool completed,
        string? error)
    {
        var sessionIds = ownedSessionIds
            .Where(sessionId => !string.IsNullOrWhiteSpace(sessionId))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (sessionIds.Length == 0) return;

        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        var tasks = sessionIds.Select(sessionId => completed
            ? client.CompleteAsync(sessionId, timeout.Token)
            : client.FailAsync(sessionId, error ?? "Word operation failed.", timeout.Token));
        await Task.WhenAll(tasks).ConfigureAwait(false);
        WordDoubleClickHook.TraceMessage(
            $"office-session-finalization-complete sessions={sessionIds.Length} completed={completed}");
    }

    private static async Task ShowBulkImportMessageAsync(
        OfficeUiDispatcher dispatcher,
        string message,
        bool warning)
    {
        if (Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE") == "1") return;
        try
        {
            await dispatcher.InvokeAsync(() =>
            {
                System.Windows.Forms.MessageBox.Show(
                    OfficePluginLanguage.SafeUserMessage(
                        message,
                        warning ? "The VisualTeX bulk import could not continue." : "The VisualTeX bulk import failed."),
                    T("VisualTeX 批量导入", "VisualTeX Bulk Import"),
                    System.Windows.Forms.MessageBoxButtons.OK,
                    warning ? System.Windows.Forms.MessageBoxIcon.Warning : System.Windows.Forms.MessageBoxIcon.Error);
                return true;
            }).ConfigureAwait(false);
        }
        catch (Exception displayError)
        {
            WriteBulkAcceptanceLog("bulk-import-error-display-failed " + displayError);
        }
    }

    private async Task<(WordBulkImportDocument? Document, string? SessionId)>
        ResolveBulkImportDocumentAsync(
            VisualTeXSessionClient client,
            string? sourceDocumentId,
            string? sourceObjectId,
            double fontSizePt,
            Action<string> sessionCreated,
            CancellationToken cancellationToken)
    {
        if (string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
        {
            var sourcePath = Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_BULK_SOURCE_PATH");
            WriteBulkAcceptanceLog(
                $"resolve-acceptance-source path={sourcePath ?? "<null>"} "
                + $"format={Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT") ?? "<null>"} "
                + $"mode={Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE") ?? "<null>"}");
            var acceptanceSource = !string.IsNullOrWhiteSpace(sourcePath)
                ? File.ReadAllText(sourcePath!, Encoding.UTF8)
                : Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE")
                  ?? throw new InvalidOperationException(
                      "Acceptance bulk import requires VISUALTEX_VSTO_BULK_SOURCE_PATH or VISUALTEX_VSTO_BULK_SOURCE.");
            var acceptanceFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT")
                ?.Trim().ToLowerInvariant() switch
            {
                "markdown" => WordBulkSourceFormat.Markdown,
                "latex" => WordBulkSourceFormat.Latex,
                _ => WordBulkSourceFormat.Auto,
            };
            var configuredObjectMode = Environment.GetEnvironmentVariable(
                "VISUALTEX_VSTO_BULK_OBJECT_MODE");
            var acceptanceObjectMode = string.Equals(
                    configuredObjectMode,
                    "ole",
                    StringComparison.OrdinalIgnoreCase)
                ? WordBulkFormulaObjectMode.Ole
                : string.Equals(
                    configuredObjectMode,
                    "mathtype",
                    StringComparison.OrdinalIgnoreCase)
                    ? WordBulkFormulaObjectMode.MathType
                    : WordBulkFormulaObjectMode.Omml;
            var acceptanceDocument = WordBulkImportParser.Parse(
                acceptanceSource,
                acceptanceFormat,
                acceptanceObjectMode);
            acceptanceDocument.NumberDisplayFormulas = IsEnabledEnvironmentOption(
                "VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS");
            return (acceptanceDocument, null);
        }

        await client.EnsureHealthyAsync(cancellationToken).ConfigureAwait(false);
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = string.Empty,
        };
        var session = await client.CreateSessionAsync(
            new CreateVstoSessionRequest
            {
                Mode = "create",
                Host = "word",
                SourceDocumentId = sourceDocumentId,
                SourceObjectId = sourceObjectId,
                Title = "Word 文档批量导入",
                Lines = new List<FormulaLine> { line },
                ActiveLineId = line.Id,
                CodeFormat = "auto-document",
                DisplayMode = "block",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = false,
                FontSizePt = fontSizePt,
                AutoCommitOnClose = false,
            },
            cancellationToken).ConfigureAwait(false);
        sessionCreated(session.Id);
        WriteBulkAcceptanceLog($"bulk-import-ui-created sessionId={session.Id}");
        GrantVisualTeXForegroundActivation();
        await client.OpenBulkImportAsync(session.Id, cancellationToken)
            .ConfigureAwait(false);
        WriteBulkAcceptanceLog($"bulk-import-ui-opened sessionId={session.Id}");
        session = await client.WaitForCommitAsync(
                session.Id,
                TimeSpan.FromHours(1),
                cancellationToken)
            .ConfigureAwait(false);
        WriteBulkAcceptanceLog(
            $"bulk-import-ui-finished sessionId={session.Id} status={session.Status} "
            + $"error={session.Error ?? "<null>"}");
        if (session.Status == "cancelled" || session.ExplicitCancel)
            return (null, session.Id);
        if (session.Status == "failed")
            throw new InvalidOperationException(
                session.Error ?? "VisualTeX 文档导入窗口返回失败状态。");
        if (session.Status is not ("committing" or "completed"))
            throw new InvalidOperationException(
                $"VisualTeX 文档导入窗口返回了意外状态：{session.Status}。");

        var source = string.Join("\n", session.Lines.Select(item => item.Latex));
        var objectMode = string.Equals(
                session.ObjectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            ? WordBulkFormulaObjectMode.Ole
            : string.Equals(
                session.ObjectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal)
                ? WordBulkFormulaObjectMode.MathType
                : WordBulkFormulaObjectMode.Omml;
        WordBulkImportDocument parsedDocument;
        if (string.Equals(
                session.CodeFormat,
                "visualtex-document-json",
                StringComparison.OrdinalIgnoreCase))
        {
            parsedDocument = WordBulkImportParser.ParseSerialized(source, objectMode);
        }
        else
        {
            var format = session.CodeFormat.Trim().ToLowerInvariant() switch
            {
                "markdown-document" => WordBulkSourceFormat.Markdown,
                "latex-document" => WordBulkSourceFormat.Latex,
                _ => WordBulkSourceFormat.Auto,
            };
            parsedDocument = WordBulkImportParser.Parse(source, format, objectMode);
        }
        parsedDocument.NumberDisplayFormulas = session.Numbered;
        return (parsedDocument, session.Id);
    }

    private static async Task<OfficeSessionDocument> CreateBulkFormulaConversionSessionAsync(
        VisualTeXSessionClient client,
        WordBulkRun run,
        string objectMode,
        string? sourceDocumentId,
        double fontSizePt,
        CancellationToken cancellationToken)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = run.Latex,
        };
        var originalMetadata = CreateBulkFormulaMetadata(
            formulaId,
            line,
            run,
            fontSizePt);
        WriteBulkAcceptanceLog(
            $"converter-create display={run.DisplayMode} mode={objectMode} latex={run.Latex}");
        var session = await client.CreateSessionAsync(
            new CreateVstoSessionRequest
            {
                Mode = "create",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = sourceDocumentId,
                Title = "Bulk imported Word formula",
                Lines = new List<FormulaLine> { line },
                ActiveLineId = line.Id,
                CodeFormat = "latex",
                DisplayMode = run.DisplayMode,
                ObjectMode = objectMode,
                Numbered = false,
                FontSizePt = fontSizePt,
                OriginalMetadata = originalMetadata,
                AutoCommitOnClose = false,
            },
            cancellationToken).ConfigureAwait(false);
        WriteBulkAcceptanceLog($"converter-created sessionId={session.Id}");
        return session;
    }

    private static RenderedWordBulkFormulaTemplate MaterializeBulkFormulaTemplate(
        VisualTeXSessionClient client,
        WordBulkRun run,
        string objectMode,
        OfficeSessionDocument session)
    {
        if (session.Status == "failed")
        {
            var detail = string.IsNullOrWhiteSpace(session.Error)
                || string.Equals(session.Error, "[object Object]", StringComparison.Ordinal)
                ? "MathJax 无法解析该公式，转换窗口没有返回有效错误文本。"
                : session.Error!.Trim();
            var formula = run.Latex.Length <= 500
                ? run.Latex
                : run.Latex.Substring(0, 500) + "…";
            throw new InvalidOperationException(
                $"公式渲染失败：{formula}\r\n原因：{detail}");
        }
        if (session.Status == "cancelled" || session.ExplicitCancel)
            throw new OperationCanceledException("批量公式渲染已取消。" );
        var export = session.ExportResult
            ?? throw new InvalidOperationException(
                $"公式 {run.Latex} 没有生成导出结果。" );
        if (string.IsNullOrWhiteSpace(export.MathMl))
            throw new InvalidDataException(
                $"公式 {run.Latex} 没有生成 MathML。" );

        var template = new RenderedWordBulkFormulaTemplate
        {
            Session = session,
            MathMl = export.MathMl,
        };
        var needsVectorPreview = string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            || string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal);
        if (needsVectorPreview)
        {
            template.SvgPath = client.MaterializeSvg(session);
            try
            {
                if (string.Equals(
                        objectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    template.PngPath = client.MaterializePng(session);
                    var originalPngPath = template.PngPath;
                    var preview = OfficeOlePreview.CreateInkSafePreviewFromSvg(
                        template.SvgPath,
                        export.Width,
                        export.Height,
                        export.Baseline,
                        fallbackPngPath: originalPngPath);
                    if (!string.Equals(
                            preview.PngPath,
                            originalPngPath,
                            StringComparison.OrdinalIgnoreCase))
                        TryDeleteFile(originalPngPath);
                    template.EmfPath = preview.EmfPath;
                    template.PngPath = preview.PngPath;
                    export.Width = preview.WidthPixels;
                    export.Height = preview.HeightPixels;
                    export.Baseline = preview.BaselinePixels;
                }
                else
                {
                    var preview = OfficeOlePreview.CreateInkSafeVectorPreviewFromSvg(
                        template.SvgPath,
                        export.Width,
                        export.Height,
                        export.Baseline,
                        safetyPaddingPixels: MathTypePreviewHorizontalSafetyInsetPixels);
                    template.EmfPath = preview.EmfPath;
                    export.Width = preview.WidthPixels;
                    export.Height = preview.HeightPixels;
                    export.Baseline = preview.BaselinePixels;
                }
            }
            catch (Exception error)
            {
                var formula = run.Latex
                    .Replace("\r", " ")
                    .Replace("\n", " ")
                    .Trim();
                if (formula.Length > 240)
                    formula = formula.Substring(0, 237) + "...";
                throw new InvalidDataException(
                    $"OLE 公式预览生成失败：{formula}\r\n原因：{error.Message}",
                    error);
            }
        }
        return template;
    }

    private static FormulaMetadata CreateBulkFormulaMetadata(
        string formulaId,
        FormulaLine line,
        WordBulkRun run,
        double fontSizePt)
    {
        var now = DateTimeOffset.UtcNow.ToString("O");
        return new FormulaMetadata
        {
            FormulaId = formulaId,
            Title = "Bulk imported Word formula",
            Latex = line.Latex,
            Lines = new List<FormulaLine> { line },
            CodeFormat = "latex",
            DisplayMode = run.DisplayMode,
            Numbered = false,
            EquationTag = run.DisplayMode == "block" ? run.EquationTag : null,
            FontSizePt = FormulaFontSize.Normalize(fontSizePt),
            RenderFontSizePt = FormulaFontSize.Normalize(fontSizePt),
            CreatedWithVersion = "1.2.7",
            UpdatedWithVersion = "1.2.7",
            CreatedAt = now,
            UpdatedAt = now,
        };
    }

    private static OfficeSessionDocument CloneBulkFormulaSession(
        OfficeSessionDocument template,
        WordBulkRun run,
        string? sourceDocumentId,
        double fontSizePt,
        string objectMode,
        string mathTypeNumberPosition = "right")
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = run.Latex,
        };
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = sourceDocumentId,
            Title = "Bulk imported Word formula",
            Lines = new List<FormulaLine> { line },
            CodeFormat = "latex",
            DisplayMode = run.DisplayMode,
            ObjectMode = objectMode,
            Numbered = false,
            MathTypeNumberPosition = string.Equals(
                mathTypeNumberPosition,
                "left",
                StringComparison.OrdinalIgnoreCase)
                    ? "left"
                    : "right",
            FontSizePt = fontSizePt,
            Status = "committing",
            Dirty = true,
            OriginalMetadata = CreateBulkFormulaMetadata(
                formulaId,
                line,
                run,
                fontSizePt),
            ExportResult = template.ExportResult,
        };
    }

    private static bool IsEnabledEnvironmentOption(string name)
    {
        var value = Environment.GetEnvironmentVariable(name)?.Trim();
        return string.Equals(value, "1", StringComparison.Ordinal)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
            || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteRedrawAcceptanceLog(string message)
    {
        var path = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG");
        if (string.IsNullOrWhiteSpace(path)) return;
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
            lock (BulkAcceptanceLogGate)
            {
                File.AppendAllText(
                    path!,
                    $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch { }
    }

    private static void WriteBulkAcceptanceLog(string message)
    {
        WordDoubleClickHook.TraceMessage("bulk-import " + message);
        var path = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");
        if (string.IsNullOrWhiteSpace(path)) return;
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory!);
            lock (BulkAcceptanceLogGate)
            {
                File.AppendAllText(
                    path!,
                    $"{DateTimeOffset.Now:O} {message}{Environment.NewLine}",
                    Encoding.UTF8);
            }
        }
        catch { }
    }

    private static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;
        try { File.Delete(path!); } catch { }
    }

    private async Task UpdateEquationNumbersAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            var count = await dispatcher.InvokeAsync(
                    service.UpdateEquationNumbers)
                .ConfigureAwait(false);
            SetStatus($"已更新 {count} 个 VisualTeX / MathType 公式编号及相关引用。");
        }
        catch (Exception error)
        {
            SetStatus($"更新 Word 公式编号失败：{error.Message}");
            await ShowWordOperationErrorAsync(dispatcher, "更新公式编号", error).ConfigureAwait(false);
        }
    }

    private async Task SetEquationNumberFormatAsync(string? requestedFormatId)
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        var format = EquationNumberFormat.Resolve(requestedFormatId);
        try
        {
            // Always apply an explicit selection. Older documents may already
            // store this VisualTeX format id while their native MathType
            // MTPlaceRef fields still use a different template; short-circuiting
            // here would make choosing the same menu item appear to do nothing.
            var count = await dispatcher.InvokeAsync(
                    () => service.SetEquationNumberFormat(format.Id))
                .ConfigureAwait(false);
            // Publish the future-document default only after the actual document
            // edit has committed. A failed rewrite must leave both unchanged.
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(format.Id);
            _cachedEquationNumberFormatId = format.Id;
            SetStatus($"公式编号格式已设置为“{format.DisplayName}”，并同步更新了 {count} 个 VisualTeX / MathType 带编号公式。");
        }
        catch (Exception error)
        {
            SetStatus($"设置公式编号格式失败：{error.Message}");
            await ShowWordOperationErrorAsync(dispatcher, "编号格式", error).ConfigureAwait(false);
        }
        finally { InvalidateEquationNumberFormatControls(); }
    }

    private async Task InsertEquationReferenceAsync()
    {
        var dispatcher = _dispatcher;
        var application = _application;
        var service = _formulaService;
        if (dispatcher is null || application is null || service is null) return;
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
                    // Freeze the real Word insertion state before the modal target
                    // picker steals focus. Desktop Word can temporarily expose
                    // GOTOBUTTON's built-in red typing format after that focus
                    // transition; MathType references must inherit what the user
                    // had at the insertion point before opening the picker.
                    var referenceInsertionStart = selection.Start;
                    var referenceInsertionEnd = selection.End;
                    var referenceInsertionColor = selection.Font.Color;
                    var visualTexTargets = WordEquationNumbering.GetEquationReferenceTargets(document);
                    var mathTypeTargets = MathTypeEquationReferences.GetTargets(document);
                    if (visualTexTargets.Count == 0 && mathTypeTargets.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            T(
                                "当前文档没有可引用的带编号公式。请先插入带编号的 VisualTeX 或 MathType 行间公式。",
                                "The current document has no numbered equations to reference. Insert a numbered VisualTeX or MathType display equation first."),
                            "VisualTeX",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return string.Empty;
                    }

                    static string DescribeReferenceTarget(EquationReferenceTarget target) =>
                        target.Source == EquationReferenceSource.MathType
                            ? $"MathType 公式 {target.NumberText}"
                            : $"VisualTeX 公式 ({target.NumberText})";

                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        var requestedSource = Environment.GetEnvironmentVariable(
                            "VISUALTEX_VSTO_REFERENCE_SOURCE");
                        var targets = string.Equals(
                                requestedSource,
                                "mathtype",
                                StringComparison.OrdinalIgnoreCase)
                            ? mathTypeTargets
                            : visualTexTargets.Count > 0
                                ? visualTexTargets
                                : mathTypeTargets;
                        if (targets.Count == 0)
                            throw new InvalidOperationException(
                                $"Acceptance requested equation-reference source '{requestedSource}' but that source has no targets.");
                        var requestedIndex = 0;
                        _ = int.TryParse(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REFERENCE_TARGET_INDEX"),
                            out requestedIndex);
                        requestedIndex = Math.Max(0, Math.Min(targets.Count - 1, requestedIndex));
                        var target = targets[requestedIndex];
                        if (target.Source == EquationReferenceSource.MathType)
                        {
                            selection.SetRange(referenceInsertionStart, referenceInsertionEnd);
                            MathTypeEquationReferences.InsertReference(
                                document,
                                selection,
                                target,
                                referenceInsertionColor);
                        }
                        else
                        {
                            WordEquationNumbering.InsertEquationReference(
                                document,
                                selection,
                                target,
                                EquationReferenceStyle.Parenthesized,
                                referenceInsertionColor);
                        }
                        return DescribeReferenceTarget(target);
                    }

                    using var dialog = new EquationReferenceDialog(
                        visualTexTargets,
                        mathTypeTargets);
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
                    selection.SetRange(referenceInsertionStart, referenceInsertionEnd);
                    service.InsertEquationReference(document, selection, dialog.SelectedTarget,
                        dialog.SelectedStyle, referenceInsertionColor);
                    return DescribeReferenceTarget(dialog.SelectedTarget);
                }
                finally
                {
                    ReleaseComObject(window);
                    ReleaseComObject(selection);
                    ReleaseComObject(document);
                }
            }).ConfigureAwait(false);
            if (!string.IsNullOrWhiteSpace(inserted))
            {
                WordDoubleClickHook.TraceMessage($"reference-insertion-complete target={inserted}");
                SetStatus($"已插入 {inserted} 的交叉引用；更新编号时引用会同步刷新。");
            }
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage($"reference-insertion-failed error={error}");
            SetStatus($"插入公式引用失败：{error.Message}");
            await ShowWordOperationErrorAsync(dispatcher, "插入公式引用", error).ConfigureAwait(false);
        }
    }

    private async Task ShowWordOperationErrorAsync(OfficeUiDispatcher dispatcher, string operation, Exception error)
    {
        WordDoubleClickHook.TraceMessage($"word-operation-failed operation={operation} error={error}");
        try
        {
            await dispatcher.InvokeAsync(() =>
            {
                Window? window = null;
                try
                {
                    window = _application?.ActiveWindow
                        ?? throw new InvalidOperationException("Word's error-reporting window is unavailable.");
                    System.Windows.Forms.MessageBox.Show(new NativeWindowOwner(new IntPtr(window.Hwnd)),
                        OfficePluginLanguage.SafeErrorMessage(error.Message),
                        OfficePluginLanguage.IsEnglish ? "VisualTeX Word" : "VisualTeX " + operation,
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return true;
                }
                finally { ReleaseComObject(window); }
            }).ConfigureAwait(false);
        }
        catch (Exception displayError)
        {
            // Preserve both errors when Word is shutting down. This only reports
            // an already-failed operation; it never converts that failure to success.
            WordDoubleClickHook.TraceMessage(
                $"word-operation-error-display-failed operation={operation} original={error} display={displayError}");
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
        var userMessage = OfficePluginLanguage.SafeStatusMessage(message);
        dispatcher.Post(() =>
        {
            try { application.StatusBar = userMessage; } catch { }
        });
    }

    private static void ReleaseComObject(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }

    internal void ResetSelectionChangeDiagnosticsForAcceptance()
    {
        Interlocked.Exchange(ref _acceptanceSelectionChangeCount, 0);
        Interlocked.Exchange(ref _acceptanceFormulaStateReadCount, 0);
        Interlocked.Exchange(ref _acceptanceDeferredCaretPassCount, 0);
        Interlocked.Exchange(ref _acceptanceEquationFormatReadCount, 0);
    }

    internal (int SelectionChanges, int FormulaStateReads, int DeferredCaretPasses, int EquationFormatReads)
        ReadSelectionChangeDiagnosticsForAcceptance() =>
        (
            Volatile.Read(ref _acceptanceSelectionChangeCount),
            Volatile.Read(ref _acceptanceFormulaStateReadCount),
            Volatile.Read(ref _acceptanceDeferredCaretPassCount),
            Volatile.Read(ref _acceptanceEquationFormatReadCount));

    private void Dispose()
    {
        _lifetime?.Cancel();
        StopOmmlRedrawUndoCleanupWatcherIfIdle(force: true);
        CancellationTokenSource? activeOperationCancellation = null;
        lock (_activeSessionOperationGate)
        {
            activeOperationCancellation = _activeSessionCancellation;
            _activeSessionCancellation = null;
            _activeSessionId = null;
        }
        try { activeOperationCancellation?.Cancel(); }
        catch (ObjectDisposedException) { }
        if (_application is not null)
        {
            try { _application.WindowBeforeDoubleClick -= OnWindowBeforeDoubleClick; } catch { }
            try { _application.WindowSelectionChange -= OnWindowSelectionChange; } catch { }
            try { _application.WindowActivate -= OnWindowActivate; } catch { }
            try { _application.DocumentOpen -= OnDocumentOpen; } catch { }
            try { _application.DocumentBeforeSave -= OnDocumentBeforeSave; } catch { }
        }
        try { _doubleClickHook?.Dispose(); } catch { }
        _doubleClickHook = null;
        try { _copyPasteHook?.Dispose(); } catch { }
        _copyPasteHook = null;
        lock (_copyPasteGate)
        {
            _formulaCopySnapshot = null;
            _formulaCopyClipboardSequence = 0;
        }
        Interlocked.Increment(ref _copyPasteWatchGeneration);
        Interlocked.Increment(ref _copyPasteRepairGeneration);
        ClearNativeOleTarget();
        if (_mathTypePreviewSessionAcquired)
        {
            _mathTypePreviewSessionAcquired = false;
            MathTypeNativePreviewRenderer.ReleaseSharedSession();
        }
        _sessionClient?.Dispose();
        _dispatcher?.Dispose();
        _lifetime?.Dispose();
        _sessionClient = null;
        _dispatcher = null;
        _formulaService = null;
        WordOfficeMathFontLoader.UnloadSessionRegistration();
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

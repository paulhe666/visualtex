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

    [DispId(47)]
    void OnRepairInlineFormulaBaselines(object control);

    // DispIds 48-50 belonged to the retired split-button default-scope
    // callbacks. Each conversion row now has explicit selection and document
    // buttons; the ids stay unused so no callback is ever renumbered.

    [DispId(51)]
    void OnRepairMathTypeDisplays(object control);

    [DispId(52)]
    void OnUpdateEquationFormat(object control);

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
          <box id="VisualTeX.WordVsto.NumberingBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.UpdateNumbers" label="更新公式编号" screentip="更新 VisualTeX 与 MathType 公式编号" supertip="刷新当前文档中的 VisualTeX 编号，以及 MathType 原生 MTChap/MTSec/MTEqn 编号和对应公式引用。" tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationNumbers" />
            <menu id="VisualTeX.WordVsto.NumberFormat" label="编号格式" screentip="设置当前文档的公式编号格式" supertip="选择后同步更新当前文档已有的 VisualTeX 编号和 MathType 原生 MTPlaceRef 编号，并应用于后续新插入的带编号公式。">
              <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="全文连续编号（1）" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="按章编号（1.1）" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="按章编号（1-1）" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="按节编号（1.1.1）" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="按节编号（1.1-1）" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            </menu>
            <button id="VisualTeX.WordVsto.InsertReference" label="插入公式引用" screentip="引用带编号公式" supertip="分别从当前文档的 VisualTeX、Word OMML 或 MathType 带编号公式中选择目标；各格式沿用各自的原生编号与引用结构。" imageMso="HyperlinkInsert" onAction="OnInsertEquationReference" />
          </box>
          <button id="VisualTeX.WordVsto.UpdateEquationFormat" label="更新公式格式" size="large" screentip="更新公式格式" supertip="把本文档中旧版 VisualTeX 写入的公式编号结构一次升级为当前格式（与 Mac 版相同）。公式内容与类型不变。" tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationFormat" />
          <button id="VisualTeX.WordVsto.BulkImport" label="批量导入" size="large" screentip="批量导入 LaTeX / Markdown" supertip="将 Markdown 或 LaTeX 文档解析为 Word 原生文字，以及可单独编辑和调整字号的行内/行间公式。" tag="batchImport" getImage="GetRibbonImage" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.FormatConversionGroup" label="公式格式转换">
          <box id="VisualTeX.WordVsto.FormatConversionLabelsA" boxStyle="vertical">
            <labelControl id="VisualTeX.WordVsto.VisualTeXToMathTypeLabel" label="VisualTeX → MathType" />
            <labelControl id="VisualTeX.WordVsto.MathTypeToVisualTeXLabel" label="MathType → VisualTeX" />
            <labelControl id="VisualTeX.WordVsto.OmmlToMathTypeLabel" label="OMML → MathType" />
          </box>
          <box id="VisualTeX.WordVsto.FormatConversionSelectionA" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.VisualTeXToMathTypeSelection" label="所选" screentip="所选 VisualTeX → MathType" supertip="把所选范围中的 VisualTeX OLE 公式重新绘制为 MathType OLE。" onAction="OnConvertVisualTeXToMathTypeSelection" />
            <button id="VisualTeX.WordVsto.MathTypeToVisualTeXSelection" label="所选" screentip="所选 MathType → VisualTeX" supertip="把所选范围中的 MathType 公式重新绘制为 VisualTeX OLE。" onAction="OnConvertMathTypeToVisualTeXSelection" />
            <button id="VisualTeX.WordVsto.OmmlToMathTypeSelection" label="所选" screentip="所选 OMML → MathType" supertip="把所选范围中的 Word 原生公式转换为 MathType OLE。" onAction="OnConvertOmmlToMathTypeSelection" />
          </box>
          <box id="VisualTeX.WordVsto.FormatConversionDocumentA" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.VisualTeXToMathTypeDocument" label="全文" screentip="全文 VisualTeX → MathType" supertip="把全文所有 VisualTeX OLE 公式重新绘制为 MathType OLE；开始前会再次确认。" onAction="OnConvertVisualTeXToMathTypeDocument" />
            <button id="VisualTeX.WordVsto.MathTypeToVisualTeXDocument" label="全文" screentip="全文 MathType → VisualTeX" supertip="把全文所有 MathType 公式重新绘制为 VisualTeX OLE；开始前会再次确认。" onAction="OnConvertMathTypeToVisualTeXDocument" />
            <button id="VisualTeX.WordVsto.OmmlToMathTypeDocument" label="全文" screentip="全文 OMML → MathType" supertip="把全文所有 Word 原生公式转换为 MathType OLE；开始前会再次确认。" onAction="OnConvertOmmlToMathTypeDocument" />
          </box>
          <separator id="VisualTeX.WordVsto.FormatConversionSeparator" />
          <box id="VisualTeX.WordVsto.FormatConversionLabelsB" boxStyle="vertical">
            <labelControl id="VisualTeX.WordVsto.MathTypeToOmmlLabel" label="MathType → OMML" />
            <labelControl id="VisualTeX.WordVsto.VisualTeXToOmmlLabel" label="VisualTeX → OMML" />
            <labelControl id="VisualTeX.WordVsto.OmmlToVisualTeXLabel" label="OMML → VisualTeX" />
          </box>
          <box id="VisualTeX.WordVsto.FormatConversionSelectionB" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.MathTypeToOmmlSelection" label="所选" screentip="所选 MathType → OMML" supertip="把所选范围中的 MathType 公式转换为 Word 原生公式。" onAction="OnConvertMathTypeToOmmlSelection" />
            <button id="VisualTeX.WordVsto.VisualTeXToOmmlSelection" label="所选" screentip="所选 VisualTeX → OMML" supertip="把所选范围中的 VisualTeX OLE 公式转换为 Word 原生公式。" onAction="OnConvertVisualTeXToOmmlSelection" />
            <button id="VisualTeX.WordVsto.OmmlToVisualTeXSelection" label="所选" screentip="所选 OMML → VisualTeX" supertip="把所选范围中的 Word 原生公式转换为 VisualTeX OLE。" onAction="OnConvertOmmlToVisualTeXSelection" />
          </box>
          <box id="VisualTeX.WordVsto.FormatConversionDocumentB" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.MathTypeToOmmlDocument" label="全文" screentip="全文 MathType → OMML" supertip="把全文所有 MathType 公式转换为 Word 原生公式；开始前会再次确认。" onAction="OnConvertMathTypeToOmmlDocument" />
            <button id="VisualTeX.WordVsto.VisualTeXToOmmlDocument" label="全文" screentip="全文 VisualTeX → OMML" supertip="把全文所有 VisualTeX OLE 公式转换为 Word 原生公式；开始前会再次确认。" onAction="OnConvertVisualTeXToOmmlDocument" />
            <button id="VisualTeX.WordVsto.OmmlToVisualTeXDocument" label="全文" screentip="全文 OMML → VisualTeX" supertip="把全文所有 Word 原生公式转换为 VisualTeX OLE；开始前会再次确认。" onAction="OnConvertOmmlToVisualTeXDocument" />
          </box>
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
        <group id="VisualTeX.WordVsto.CompatibilityRepairGroup" label="兼容修复">
          <button id="VisualTeX.WordVsto.RepairInlineBaselines" label="修复公式偏移" size="large" screentip="扫描并修复行内公式垂直偏移" supertip="扫描全文中的 MathType / VisualTeX 行内 OLE；VisualTeX 公式优先按自身保存的渲染基线校准，MathType 公式按文档稳定基线检测上浮或下沉。确认后仅修复垂直偏移，不修改公式内容、字号、编号或引用。" tag="repairBaseline" getImage="GetRibbonImage" onAction="OnRepairInlineFormulaBaselines" />
          <button id="VisualTeX.WordVsto.RepairMathTypeDisplays" label="修复公式显示" size="large" screentip="用 MathType 自身排版重新生成公式图像" supertip="扫描全文 MathType 公式，用 VisualTeX 内置的 MathType 排版服务按公式原有数据重新生成显示图像，修复缺少 MathType 字体时生成的错误符号。只替换显示图像并按 MathType 基线重新对齐，公式数据、编号和引用保持不变。" tag="repairBaseline" getImage="GetRibbonImage" onAction="OnRepairMathTypeDisplays" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    private sealed class NumberedHostDeleteViewState
    {
        internal WordFormulaService.NumberedHostDeleteGuard Guard { get; set; } = null!;
        internal int VerticalPercentScrolled { get; set; }
        internal int HorizontalPercentScrolled { get; set; }
    }

    private Application? _application;
    private WordFormulaService? _formulaService;
    private OfficeUiDispatcher? _dispatcher;
    private VisualTeXSessionClient? _sessionClient;
    private WordDoubleClickHook? _doubleClickHook;
    private WordCopyPasteHook? _copyPasteHook;
    private OfficeFormulaShortcutKeyboardHook? _keyboardHook;
    private readonly HashSet<OfficeFormulaShortcut> _reportedShortcutConflicts = new();
    private WordDeleteKeyHook? _deleteKeyHook;
    private readonly object _copyPasteGate = new();
    private WordFormulaService.WordFormulaCopySnapshot? _formulaCopySnapshot;
    private uint _formulaCopyClipboardSequence;
    private int _copyPasteRepairGeneration;
    private int _copyPasteWatchGeneration;
    private int _copyPasteSelectionRepairPending;
    private long _lastSuccessfulPasteRepairStartedAt;
    private long _lastSuccessfulPasteRepairCompletedAt;
    private UndoRecord? _activePasteUndoRecord;
    private WordFormulaService.WordFormulaCopySnapshot? _activePasteUndoSnapshot;
    private int _activePasteUndoGeneration;
    private NumberedHostDeleteViewState? _numberedHostDeleteViewState;
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
    private int _formulaFormatMutationDepth;
    private int _formulaFontReadsDeferredDuringMutation;
    private bool _acceptanceSelectionDiagnostics;
    private int _acceptanceSelectionChangeCount;
    private int _acceptanceFormulaStateReadCount;
    private int _acceptanceDeferredCaretPassCount;
    private int _acceptanceEquationFormatReadCount;
    private object? _ribbonUi;
    private Office.COMAddIn? _comAddIn;

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
        try
        {
            _keyboardHook = new OfficeFormulaShortcutKeyboardHook(
                OfficeFormulaShortcutHost.Word,
                OnWordFormulaShortcut,
                OnWordPasteKeyGesture,
                WordDoubleClickHook.TraceMessage);
            Window? ownerWindow = null;
            try
            {
                ownerWindow = _application.ActiveWindow;
                if (ownerWindow is not null) _keyboardHook.BindOwnerWindow(ownerWindow.Hwnd);
            }
            catch { }
            finally { ReleaseComObject(ownerWindow); }
            _keyboardHook.Start();
        }
        catch (Exception error)
        {
            try { _keyboardHook?.Dispose(); } catch { }
            _keyboardHook = null;
            WordDoubleClickHook.TraceMessage(
                $"paste-key-hook-start-failed {error.GetType().Name}: {error.Message}");
        }
        try
        {
            _deleteKeyHook = new WordDeleteKeyHook(OnWordDeleteKeyGesture);
            Window? ownerWindow = null;
            try
            {
                ownerWindow = _application.ActiveWindow;
                if (ownerWindow is not null) _deleteKeyHook.BindOwnerWindow(ownerWindow.Hwnd);
            }
            catch { }
            finally { ReleaseComObject(ownerWindow); }
            _deleteKeyHook.Start();
        }
        catch (Exception error)
        {
            try { _deleteKeyHook?.Dispose(); } catch { }
            _deleteKeyHook = null;
            WordDoubleClickHook.TraceMessage(
                $"delete-key-hook-start-failed {error.GetType().Name}: {error.Message}");
        }
        SetStatus(!officeMathFontReady
            ? $"VisualTeX 已就绪，但 Word 数学字体不可用：{officeMathFontError}"
            : doubleClickError is null
                ? "VisualTeX Word VSTO 已就绪。"
                : $"VisualTeX 已就绪，但 OLE 双击监听不可用：{doubleClickError}");
    }

    public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom) => Dispose();
    public void OnAddInsUpdate(ref Array custom) { }
    public void OnStartupComplete(ref Array custom) { }
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
        if (Volatile.Read(ref _formulaFormatMutationDepth) > 0) return false;
        try
        {
            // A selected OLE is an inspection candidate even when its host
            // identity is unavailable. Only the explicit command may read and
            // validate its MTEF; enabling this control does not claim the OLE.
            return GetCachedSelectedFormulaFontSize().HasValue
                || _formulaService?.IsSelectedOleFormulaCommandCandidate() == true;
        }
        catch { return false; }
    }
    public void OnFormulaFontSizeChanged(Office.IRibbonControl control, string value) =>
        ApplyFormulaFontSize(ParseFontSize(value));
    public void OnDecreaseFormulaFontSize(object control)
    {
        try
        {
            var current = _formulaService?.ReadSelectedFormulaFontSizeForCommand()
                ?? throw new InvalidOperationException(T("请先选择一个公式。", "Select a formula first."));
            ApplyFormulaFontSize(FormulaFontSize.PreviousPreset(current));
        }
        catch (Exception error)
        {
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
        }
    }
    public void OnIncreaseFormulaFontSize(object control)
    {
        try
        {
            var current = _formulaService?.ReadSelectedFormulaFontSizeForCommand()
                ?? throw new InvalidOperationException(T("请先选择一个公式。", "Select a formula first."));
            ApplyFormulaFontSize(FormulaFontSize.NextPreset(current));
        }
        catch (Exception error)
        {
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
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
    private bool OnWordFormulaShortcut(OfficeFormulaShortcut shortcut,
        OfficeFormulaShortcut? previous, OfficeFormulaShortcutCommand command)
    {
        if (!WordFormulaShortcutPolicy.CanExecute(_application, shortcut, previous, out var conflict))
        {
            if (conflict.HasValue) ReportWordShortcutConflict(conflict.Value);
            return false;
        }
        WordShortcutConflictRecord.Clear(shortcut);
        return BeginSession("create",
            command.Kind == OfficeFormulaShortcutKind.Inline ? "inline" : "block",
            command.Format == OfficeFormulaShortcutFormat.Omml
                ? FormulaOleContract.WordOmmlMode : null);
    }

    // Runs inside the keyboard hook: record for the settings page, and explain
    // once per chord in this Word session through the (posted) status bar.
    private void ReportWordShortcutConflict(WordShortcutConflict conflict)
    {
        WordShortcutConflictRecord.Write(conflict);
        lock (_reportedShortcutConflicts)
        {
            if (!_reportedShortcutConflicts.Add(conflict.Shortcut)) return;
        }
        var chord = conflict.Shortcut.Format();
        var name = conflict.Command;
        var owner = conflict.Category switch
        {
            WdKeyCategory.wdKeyCategoryDisable => T("Word（此键在 Word 中被设为禁用）", "Word (the key is disabled in Word)"),
            WdKeyCategory.wdKeyCategoryPrefix => T("Word 两段式快捷键的第一段", "the first stroke of a Word key sequence"),
            WdKeyCategory.wdKeyCategoryMacro => T($"Word 宏「{name}」", $"the Word macro “{name}”"),
            WdKeyCategory.wdKeyCategoryFont => T($"Word 字体「{name}」", $"the Word font “{name}”"),
            WdKeyCategory.wdKeyCategoryStyle => T($"Word 样式「{name}」", $"the Word style “{name}”"),
            WdKeyCategory.wdKeyCategoryAutoText => T($"Word 自动图文集「{name}」", $"the Word AutoText “{name}”"),
            WdKeyCategory.wdKeyCategorySymbol => T($"Word 插入符号「{name}」", $"the Word symbol “{name}”"),
            _ => string.IsNullOrWhiteSpace(name)
                ? T("Word 命令", "a Word command")
                : T($"Word 命令「{name}」", $"the Word command “{name}”"),
        };
        SetStatus(T(
            $"VisualTeX：{chord} 已被 {owner}占用，已交给 Word 处理。请在 VisualTeX 设置里换一个未占用的快捷键。",
            $"VisualTeX: {chord} is used by {owner}, so Word handled it. Choose an unused shortcut in VisualTeX settings."));
    }
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
    public void OnUpdateEquationFormat(object control) => _ = UpdateEquationFormatAsync();
    public void OnRepairInlineFormulaBaselines(object control) => _ = RepairInlineFormulaBaselinesAsync();
    public void OnRepairMathTypeDisplays(object control) => _ = RepairMathTypeDisplaysAsync();
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
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "无法打开 VisualTeX",
                "Unable to open VisualTeX"));
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
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "无法设置公式字号",
                "Unable to set formula font size"));
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

    private void ScheduleFormulaFontControlsInvalidation(
        Selection selection,
        bool contextChanged = false)
    {
        var dispatcher = _dispatcher;
        if (dispatcher is null || Volatile.Read(ref _formulaFormatMutationDepth) > 0) return;
        var generation = Volatile.Read(ref _typingCaretNormalizationGeneration);

        if (!TryResolveFormulaRibbonOwnerBounds(selection, out var ownerStart, out var ownerEnd))
        {
            var hadOwner = _lastFormulaRibbonOwnerStart != int.MinValue;
            _lastFormulaRibbonOwnerStart = int.MinValue;
            _lastFormulaRibbonOwnerEnd = int.MinValue;
            // Unknown semantic size is not disabled state: an OLE command can
            // remain enabled without a payload read. Invalidate when its target
            // is left or the window/document context changes, but keep ordinary
            // prose-to-prose movement free of redundant Ribbon refreshes.
            if (!contextChanged && !hadOwner
                && double.IsNaN(Volatile.Read(ref _cachedSelectedFormulaFontSize)))
                return;
            Volatile.Write(ref _cachedSelectedFormulaFontSize, double.NaN);
            InvalidateFormulaFontControls();
            return;
        }

        if (!contextChanged
            && ownerStart == _lastFormulaRibbonOwnerStart
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
            _numberedHostDeleteViewState = null;
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

    private void OnWordPasteKeyGesture()
    {
        var application = _application;
        if (application is null) return;

        WordFormulaService.WordFormulaCopySnapshot? snapshot;
        uint clipboardSequence;
        lock (_copyPasteGate)
        {
            snapshot = _formulaCopySnapshot;
            clipboardSequence = _formulaCopyClipboardSequence;
        }
        if (snapshot is null || clipboardSequence == 0) return;
        if (WordCopyPasteHook.CurrentClipboardSequence != clipboardSequence) return;

        // MathType keeps its mature native paste behavior. OMML and VisualTeX now
        // share one explicit Word Custom UndoRecord covering native Paste + local
        // identity/number repair.
        if (!snapshot.UsesHostCore
            && !string.Equals(
                snapshot.ObjectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal))
            return;

        // Native OMML paste is now deliberately mutation-free on the VisualTeX
        // side. Word copies the complete m:eqArr/#(SEQ) host and owns its own
        // Paste/Undo record. Keeping a Custom UndoRecord open across Word's
        // delayed professional-math normalization is actively harmful: Word can
        // close it asynchronously before our read-only validation runs.
        if (snapshot.UsesHostCore
            && snapshot.CoreHostKind == WordFormulaHostKind.Omml)
        {
            _formulaService?.ArmHostCorePaste(snapshot);
            WordDoubleClickHook.TraceMessage(
                "paste-native-omml-armed-with-word-undo");
            return;
        }

        if (_activePasteUndoRecord is not null) return;

        UndoRecord? record = null;
        try
        {
            record = application.UndoRecord;
            if (record.IsRecordingCustomRecord || record.CustomRecordLevel > 0)
            {
                WordDoubleClickHook.TraceMessage(
                    $"paste-undo-not-started reason=existing-custom-record level={record.CustomRecordLevel}");
                return;
            }

            // This callback runs from a WH_KEYBOARD hook installed on Word's own
            // UI thread, before the native Ctrl+V / Shift+Insert command executes.
            // Opening the record here is what makes Word's Paste and the later
            // identity/number repair one native Ctrl+Z operation. We do not replace
            // or synthesize the user's paste command.
            record.StartCustomRecord("VisualTeX Paste Formula");

            // Capture the exact pre-paste Selection only after this transaction is
            // known to exist. The post-paste caret will delimit the sole physical
            // region the host core is allowed to inspect.
            if (snapshot.UsesHostCore)
                _formulaService?.ArmHostCorePaste(snapshot);

            _activePasteUndoRecord = record;
            _activePasteUndoSnapshot = snapshot;
            record = null;
            var generation = Interlocked.Increment(ref _activePasteUndoGeneration);
            WordDoubleClickHook.TraceMessage(
                $"paste-undo-started mode={snapshot.ObjectMode} clipboardSequence={clipboardSequence} generation={generation}");
            ScheduleActivePasteUndoTimeout(snapshot, generation);
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"paste-undo-start-failed {error.GetType().Name}:{error.Message}");
        }
        finally { ReleaseComObject(record); }
    }

    private void ScheduleActivePasteUndoTimeout(
        WordFormulaService.WordFormulaCopySnapshot snapshot,
        int generation)
    {
        var lifetime = _lifetime;
        var dispatcher = _dispatcher;
        if (lifetime is null || dispatcher is null || lifetime.IsCancellationRequested)
            return;
        var token = lifetime.Token;
        _ = Task.Delay(2000, token).ContinueWith(
            task =>
            {
                if (task.IsCanceled || token.IsCancellationRequested) return;
                dispatcher.Post(() =>
                {
                    if (generation != Volatile.Read(ref _activePasteUndoGeneration))
                        return;
                    if (snapshot.UsesHostCore
                        && snapshot.CoreHostKind == WordFormulaHostKind.VisualTeX
                        && snapshot.CorePendingVisualTeXPhysicalPasteObserved)
                    {
                        // Word has already materialized the exact copied
                        // VisualTeX InlineShape.  A slow OLE server must never turn
                        // that successful native paste into an automatic Undo.
                        // Close only our custom record; keep the pending snapshot
                        // alive so a later selection change can finish the
                        // authoritative FormulaId re-key once COM is ready.
                        EndActivePasteUndoRecord(
                            snapshot,
                            "watchdog-timeout-physical-paste-preserved");
                    }
                    else if (snapshot.UsesHostCore)
                    {
                        RollbackActivePasteUndoRecord(
                            snapshot,
                            "watchdog-timeout");
                    }
                    else
                    {
                        EndActivePasteUndoRecord(
                            snapshot,
                            "watchdog-timeout");
                    }
                });
            },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
    }

    private void EndActivePasteUndoRecord(
        WordFormulaService.WordFormulaCopySnapshot? snapshot,
        string reason)
    {
        if (_activePasteUndoRecord is null) return;
        if (snapshot is not null
            && !ReferenceEquals(snapshot, _activePasteUndoSnapshot))
            return;

        var record = _activePasteUndoRecord;
        _activePasteUndoRecord = null;
        _activePasteUndoSnapshot = null;
        Interlocked.Increment(ref _activePasteUndoGeneration);
        try
        {
            if (record.IsRecordingCustomRecord || record.CustomRecordLevel > 0)
                record.EndCustomRecord();
            WordDoubleClickHook.TraceMessage(
                $"paste-undo-ended reason={reason}");
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"paste-undo-end-failed reason={reason} {error.GetType().Name}:{error.Message}");
        }
        finally { ReleaseComObject(record); }
    }

    private void RollbackActivePasteUndoRecord(
        WordFormulaService.WordFormulaCopySnapshot snapshot,
        string reason)
    {
        if (!snapshot.UsesHostCore)
        {
            EndActivePasteUndoRecord(snapshot, reason);
            return;
        }

        if (_activePasteUndoRecord is null)
        {
            _formulaService?.CancelHostCorePaste(snapshot);
            WordDoubleClickHook.TraceMessage(
                $"paste-undo-rollback-skipped reason={reason} no-active-record");
            return;
        }
        if (!ReferenceEquals(
                snapshot,
                _activePasteUndoSnapshot))
            return;

        var record = _activePasteUndoRecord;
        _activePasteUndoRecord = null;
        _activePasteUndoSnapshot = null;
        Interlocked.Increment(
            ref _activePasteUndoGeneration);
        try
        {
            if (record.IsRecordingCustomRecord
                || record.CustomRecordLevel > 0)
                record.EndCustomRecord();
        }
        catch (Exception endError)
        {
            _formulaService?.CancelHostCorePaste(snapshot);
            WordDoubleClickHook.TraceMessage(
                $"paste-undo-rollback-end-failed reason={reason} "
                + $"{endError.GetType().Name}:{endError.Message}");
            return;
        }
        finally
        {
            ReleaseComObject(record);
        }

        string? rollbackError = null;
        var rolledBack = false;
        if (_formulaService is not null)
        {
            rolledBack =
                _formulaService.RollbackHostCorePaste(
                    snapshot,
                    out rollbackError);
        }
        WordDoubleClickHook.TraceMessage(
            rolledBack
                ? $"paste-undo-rolled-back reason={reason}"
                : $"paste-undo-rollback-failed reason={reason} error={rollbackError}");
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
                if (snapshot.UsesHostCore
                    && snapshot.CoreHostKind == WordFormulaHostKind.VisualTeX
                    && snapshot.CorePendingVisualTeXPhysicalPasteObserved)
                {
                    // The exact copied InlineShape is already present. Identity
                    // repair is secondary bookkeeping and must never erase a
                    // successful user paste if a legacy/slow OLE server rejects a
                    // follow-up metadata write. Keep the snapshot armed so later
                    // selection activity can retry the repair.
                    EndActivePasteUndoRecord(
                        snapshot,
                        "repair-error-physical-paste-preserved");
                }
                else if (snapshot.UsesHostCore)
                {
                    RollbackActivePasteUndoRecord(
                        snapshot,
                        "repair-error");
                }
                else
                {
                    EndActivePasteUndoRecord(
                        snapshot,
                        "repair-error");
                }
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
                EndActivePasteUndoRecord(snapshot, "repair-complete");
            }
            else if (result == WordFormulaService.PastedFormulaRepairResult.NotApplicable)
            {
                // A tracked formula can legitimately be pasted into a context that
                // does not need identity repair. The native paste still belongs to
                // this custom record, but it must be closed immediately so later
                // user typing is never captured by the paste transaction watchdog.
                if (snapshot.UsesHostCore)
                    RollbackActivePasteUndoRecord(
                        snapshot,
                        "repair-not-applicable");
                else
                    EndActivePasteUndoRecord(
                        snapshot,
                        "repair-not-applicable");
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
            if (snapshot.UsesHostCore
                && snapshot.CoreHostKind == WordFormulaHostKind.VisualTeX
                && snapshot.CorePendingVisualTeXPhysicalPasteObserved)
            {
                EndActivePasteUndoRecord(
                    snapshot,
                    "repair-timeout-physical-paste-preserved");
                WordDoubleClickHook.TraceMessage(
                    $"paste-repair-timeout-preserved mode={snapshot.ObjectMode} "
                    + $"attempts={attempt + 1}");
                return;
            }

            if (snapshot.UsesHostCore)
                RollbackActivePasteUndoRecord(
                    snapshot,
                    "repair-timeout");
            else
                EndActivePasteUndoRecord(
                    snapshot,
                    "repair-timeout");
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
        try { _keyboardHook?.BindOwnerWindow(window.Hwnd); } catch { }
        try { _deleteKeyHook?.BindOwnerWindow(window.Hwnd); } catch { }
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
                ScheduleFormulaFontControlsInvalidation(selection, contextChanged: true);
        }
        catch { }
        finally { ReleaseComObject(selection); }

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

        // Save/reopen correctness is a persistence invariant. Opening a document
        // must not mutate or globally "repair" OMML/VisualTeX numbering. Explicit
        // migration/refresh commands remain available for legacy documents.
    }

    private void OnDocumentBeforeSave(
        Document document,
        ref bool saveAsUi,
        ref bool cancel)
    {
        // OMML/VisualTeX persistence is established when the mutation commits.
        // Saving a document must not enumerate or repair formula hosts.
    }

    private NumberedHostDeleteViewState? CaptureNumberedHostDeleteViewState(Selection selection)
    {
        var service = _formulaService;
        var application = _application;
        if (service is null || application is null) return null;
        var guard = service.CaptureNumberedHostDeleteGuard(selection);
        if (guard is null) return null;
        Window? window = null;
        Pane? pane = null;
        try
        {
            window = application.ActiveWindow;
            pane = window.ActivePane;
            return new NumberedHostDeleteViewState
            {
                Guard = guard,
                VerticalPercentScrolled = pane.VerticalPercentScrolled,
                HorizontalPercentScrolled = pane.HorizontalPercentScrolled,
            };
        }
        catch
        {
            return null;
        }
        finally
        {
            ReleaseComObject(pane);
            ReleaseComObject(window);
        }
    }

    private bool TryRestoreNumberedHostDeleteViewport(
        Selection selection,
        NumberedHostDeleteViewState? capturedState = null)
    {
        var state = capturedState ?? _numberedHostDeleteViewState;
        var service = _formulaService;
        var application = _application;
        if (state is null || service is null || application is null
            || !service.IsCompletedNumberedHostDeletion(selection, state.Guard))
            return false;

        if (ReferenceEquals(_numberedHostDeleteViewState, state))
            _numberedHostDeleteViewState = null;
        Window? window = null;
        Pane? pane = null;
        try
        {
            window = application.ActiveWindow;
            pane = window.ActivePane;
            pane.VerticalPercentScrolled = state.VerticalPercentScrolled;
            pane.HorizontalPercentScrolled = state.HorizontalPercentScrolled;
            WordDoubleClickHook.TraceMessage(
                $"numbered-host-delete-view-restored formulaId={state.Guard.FormulaId} "
                + $"caret={selection.Start} vertical={state.VerticalPercentScrolled} "
                + $"horizontal={state.HorizontalPercentScrolled}");
            return true;
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"numbered-host-delete-view-restore-failed formulaId={state.Guard.FormulaId} "
                + $"error={error.GetType().Name}:{error.Message}");
            return false;
        }
        finally
        {
            ReleaseComObject(pane);
            ReleaseComObject(window);
        }
    }

    private void OnWordDeleteKeyGesture()
    {
        NumberedHostDeleteViewState? capturedState = null;
        Selection? currentSelection = null;
        try
        {
            var application = _application;
            if (application is not null)
            {
                currentSelection = application.Selection;
                capturedState = CaptureNumberedHostDeleteViewState(currentSelection);
                if (capturedState is not null)
                    _numberedHostDeleteViewState = capturedState;
            }
        }
        catch (Exception error)
        {
            WordDoubleClickHook.TraceMessage(
                $"numbered-host-delete-key-live-capture-failed error={error.GetType().Name}:{error.Message}");
        }
        finally { ReleaseComObject(currentSelection); }

        capturedState ??= _numberedHostDeleteViewState;
        if (capturedState is null) return;
        WordDoubleClickHook.TraceMessage(
            $"numbered-host-delete-key-armed formulaId={capturedState.Guard.FormulaId} "
            + $"selection={capturedState.Guard.SelectionStart}:{capturedState.Guard.SelectionEnd} "
            + $"liveCapture={(ReferenceEquals(capturedState, _numberedHostDeleteViewState) ? "true" : "false")}");
        ScheduleNumberedHostDeleteViewportVerification(capturedState, attempt: 0);
    }

    private void ScheduleNumberedHostDeleteViewportVerification(
        NumberedHostDeleteViewState capturedState,
        int attempt)
    {
        if (attempt >= 6) return;
        var dispatcher = _dispatcher;
        var lifetime = _lifetime;
        if (dispatcher is null || lifetime is null || lifetime.IsCancellationRequested) return;
        var token = lifetime.Token;
        _ = Task.Delay(attempt == 0 ? 25 : 40, token).ContinueWith(
            task =>
            {
                if (task.IsCanceled || token.IsCancellationRequested) return;
                dispatcher.Post(() =>
                {
                    var application = _application;
                    if (application is null) return;
                    Selection? currentSelection = null;
                    try
                    {
                        currentSelection = application.Selection;
                        if (TryRestoreNumberedHostDeleteViewport(currentSelection, capturedState))
                            return;
                    }
                    catch (Exception error)
                    {
                        WordDoubleClickHook.TraceMessage(
                            $"numbered-host-delete-view-verify-failed attempt={attempt} "
                            + $"error={error.GetType().Name}:{error.Message}");
                    }
                    finally { ReleaseComObject(currentSelection); }
                    ScheduleNumberedHostDeleteViewportVerification(capturedState, attempt + 1);
                });
            },
            CancellationToken.None,
            TaskContinuationOptions.ExecuteSynchronously,
            TaskScheduler.Default);
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
        var previousDeleteViewState = _numberedHostDeleteViewState;
        var restoredDeletedHostViewport = TryRestoreNumberedHostDeleteViewport(selection);
        if (!restoredDeletedHostViewport)
        {
            if (previousDeleteViewState is not null
                && service.IsPendingNumberedHostDeletionTransition(
                    selection,
                    previousDeleteViewState.Guard))
            {
                WordDoubleClickHook.TraceMessage(
                    $"numbered-host-delete-view-guard-retained formulaId={previousDeleteViewState.Guard.FormulaId} "
                    + $"caret={selection.Start}");
            }
            else
            {
                _numberedHostDeleteViewState = CaptureNumberedHostDeleteViewState(selection);
                if (_numberedHostDeleteViewState is not null)
                    WordDoubleClickHook.TraceMessage(
                        $"numbered-host-delete-view-guard-captured formulaId={_numberedHostDeleteViewState.Guard.FormulaId} "
                        + $"selection={_numberedHostDeleteViewState.Guard.SelectionStart}:{_numberedHostDeleteViewState.Guard.SelectionEnd} "
                        + $"vertical={_numberedHostDeleteViewState.VerticalPercentScrolled}");
            }
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
                if (!MathTypeDoubleClickPreference.VisualTeXEditsMathType())
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
                if (!MathTypeDoubleClickPreference.VisualTeXEditsMathType())
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
                    && !MathTypeDoubleClickPreference.VisualTeXEditsMathType())
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
                    if (MathTypeDoubleClickPreference.VisualTeXEditsMathType())
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
            var ommlReadPrewarm = Task.Run(
                WordOmmlConverter.PrewarmOmmlReadPipeline,
                cancellationToken);
            // Pay the reusable Office editor WebView cold-start cost while Word
            // itself is starting, not on the user's first Edit Formula click.
            // The companion remains WebView-free until Word explicitly requests
            // this prewarm, preserving office-background/bootstrap startup.
            await client.PrewarmEditorAsync(cancellationToken).ConfigureAwait(false);
            await client.PrewarmConverterAsync(cancellationToken).ConfigureAwait(false);
            await ommlReadPrewarm.ConfigureAwait(false);
        }
        catch
        {
            // Startup must remain non-blocking. The first explicit Office action
            // retries the full diagnostic/startup path and reports any failure.
        }
    }

    private bool BeginSession(
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
            return false;
        }
        if (string.Equals(mode, "create", StringComparison.Ordinal)
            && capturedSelection is null)
        {
            try
            {
                capturedSelection = _formulaService?.ReadCreateSelection()
                    ?? throw new InvalidOperationException("Word formula service is unavailable.");
            }
            catch (Exception error)
            {
                SetStatus($"无法读取当前插入位置：{error.Message}");
                WordDoubleClickHook.TraceMessage(
                    $"ribbon-create-selection-capture-failed: {error}");
                return false;
            }
        }
        _ = ObserveRibbonSessionTaskAsync(RunSessionAsync(
            mode,
            displayMode,
            requestedObjectMode,
            capturedSelection,
            conversionOnly,
            lifetime.Token));
        return true;
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

        var traceOpenPerformanceToHook = !string.IsNullOrWhiteSpace(
            Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH"));
        var openPerformance = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal)
            || traceOpenPerformanceToHook
            ? Stopwatch.StartNew()
            : null;
        long openPerformanceCheckpoint = 0;
        void TraceOpenPerformance(string stage)
        {
            if (openPerformance is null) return;
            var elapsed = openPerformance.ElapsedMilliseconds;
            var message =
                $"[perf] OpenSession.{stage}: +{elapsed - openPerformanceCheckpoint}ms ({elapsed}ms total)";
            Console.WriteLine("    " + message);
            if (traceOpenPerformanceToHook)
                WordDoubleClickHook.TraceMessage(message);
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
            var selection = capturedSelection
                ?? await dispatcher.InvokeAsync(service.ReadSelection).ConfigureAwait(false);
            TraceOpenPerformance("read-selection");
            if (selection.ReadOnly)
                throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");
            if (mode == "edit" && selection.Metadata is null)
                throw new InvalidOperationException("没有选中公式。请先选中一个 VisualTeX 公式、Word 原生公式或 MathType 公式。");

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
                    return service.ApplyMathTypeTargetHostSession(
                        session,
                        selection.ObjectMode,
                        mathMl,
                        emfPath);
                }
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    if (mathMl is null)
                        throw new InvalidOperationException(
                            "VisualTeX Word OMML MathML payload is unavailable.");
                    if (!traceOpenPerformanceToHook)
                    {
                        return service.ApplyOmmlVisualTeXHostSession(
                            session,
                            selection.ObjectMode,
                            mathMl,
                            pngPath: null,
                            emfPath: null);
                    }

                    var ommlApplyWatch = Stopwatch.StartNew();
                    using var ommlMetrics = new WordOperationMetrics();
                    try
                    {
                        return service.ApplyOmmlVisualTeXHostSession(
                            session,
                            selection.ObjectMode,
                            mathMl,
                            pngPath: null,
                            emfPath: null);
                    }
                    finally
                    {
                        ommlApplyWatch.Stop();
                        WordDoubleClickHook.TraceMessage(
                            $"[perf] ApplySession.omml-write: {ommlApplyWatch.Elapsed.TotalMilliseconds:0.0}ms");
                        foreach (var entry in ommlMetrics.Entries
                                     .OrderByDescending(item => item.Value.Milliseconds))
                        {
                            WordDoubleClickHook.TraceMessage(
                                $"[perf] ApplySession.omml-stage {entry.Key}: "
                                + $"calls={entry.Value.Calls} "
                                + $"inclusive={entry.Value.Milliseconds:0.0}ms");
                        }
                    }
                }
                if (string.Equals(
                        session.ObjectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    if (emfPath is null || imagePath is null)
                        throw new InvalidOperationException(
                            "VisualTeX native OLE previews are unavailable.");
                    return service.ApplyOmmlVisualTeXHostSession(
                        session,
                        selection.ObjectMode,
                        mathMl: null,
                        pngPath: imagePath,
                        emfPath: emfPath);
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
            WordDoubleClickHook.TraceMessage(
                $"ribbon-session-failed sessionId={sessionId ?? "<pending>"} mode={mode} "
                + $"error={error.GetType().FullName}: {error}");
            var failureMessage = OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX Word 写入失败",
                "VisualTeX Word formula write failed");
            if (sessionId is not null && _sessionClient is not null)
            {
                try
                {
                    await _sessionClient.FailAsync(sessionId, failureMessage, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            SetStatus(failureMessage);
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

        // The Office Session API accepts only canonical UUID text for line ids.
        // Older metadata and one former native-OMML adapter could contain the
        // same UUID in "N" form (32 hex digits without dashes). Normalize the
        // editable clone at the Session boundary instead of rejecting or
        // rewriting the persisted source metadata merely because it is opened.
        var usedLineIds = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase);
        foreach (var line in metadata.Lines)
        {
            string canonical;
            if (Guid.TryParse(
                    line.Id,
                    out var parsed))
            {
                canonical =
                    parsed.ToString("D");
            }
            else
            {
                canonical =
                    Guid.NewGuid().ToString("D");
            }

            while (!usedLineIds.Add(
                       canonical))
            {
                canonical =
                    Guid.NewGuid().ToString("D");
            }
            line.Id = canonical;
        }

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

        var progress = await dispatcher.InvokeAsync(
            () => BeginOperationProgress("VisualTeX LaTeX 重绘", service)).ConfigureAwait(false);

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(
            StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(
            StringComparer.Ordinal);
        var mathTypePreviews =
            new Dictionary<string, MathTypeNativePreviewRenderer.Result>(
                StringComparer.Ordinal);
        var mathTypeStandalone =
            new Dictionary<string, (
                MathTypeMtefCodec.RewriteResult Generated,
                byte[] CompoundFile,
                string SemanticSignature)>(StringComparer.Ordinal);
        var converterSessionIds = new List<string>();
        var renderFailures = new Dictionary<string, string>(StringComparer.Ordinal);
        var skippedTargets = new List<(WordLatexRedrawTarget Target, string Error)>();
        var maxRenderMilliseconds = 0L;
        var totalRenderMilliseconds = 0L;
        var redrawPhaseWatch = Stopwatch.StartNew();
        try
        {
            var phaseStartMilliseconds = redrawPhaseWatch.ElapsedMilliseconds;
            var plan = await dispatcher.InvokeAsync(
                    () => service.CaptureLatexRedrawPlan(wholeDocument))
                .ConfigureAwait(false);
            WriteRedrawAcceptanceLog(
                $"redraw-phase captureMs={redrawPhaseWatch.ElapsedMilliseconds - phaseStartMilliseconds} "
                + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds} "
                + string.Join(
                    " ",
                    plan.CaptureTimings.Select(item =>
                        $"capture.{item.Key}Ms={item.Value}")));
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
            var allowRedrawNumbering = true;
            var numberDisplayFormulas = allowRedrawNumbering
                && acceptanceMode
                && IsEnabledEnvironmentOption(
                    "VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS");
            if (!acceptanceMode)
            {
                // Capture has finished. Waiting for the user's choice is not
                // scanning, and its dialog must not compete with an owned HUD.
                progress.Dispose();
                var options = await dispatcher.InvokeAsync(() =>
                {
                    using var dialog = new LatexRedrawDialog(
                        wholeDocument,
                        plan.Targets.Count,
                        displayFormulaCount,
                        modeLabel,
                        service.GetEquationNumberFormatDisplayName(),
                        allowRedrawNumbering);
                    var accepted = dialog.ShowDialog(progress.Owner)
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
                progress = await dispatcher.InvokeAsync(
                    () => BeginOperationProgress("VisualTeX LaTeX 重绘", service, "正在准备重绘…")).ConfigureAwait(false);
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
            phaseStartMilliseconds = redrawPhaseWatch.ElapsedMilliseconds;
            progress.Report("正在连接公式转换服务…");
            await client.EnsureHealthyAsync(lifetime.Token).ConfigureAwait(false);
            progress.Report("正在准备公式渲染器…");
            await client.PrewarmConverterAsync(lifetime.Token).ConfigureAwait(false);
            WriteRedrawAcceptanceLog(
                $"redraw-phase healthPrewarmMs={redrawPhaseWatch.ElapsedMilliseconds - phaseStartMilliseconds} "
                + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds}");

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
            var mathTypePreflightPrepared =
                new System.Collections.Concurrent.ConcurrentDictionary<
                    string,
                    (MathTypeMtefCodec.RewriteResult Generated,
                     byte[] CompoundFile,
                     string SemanticSignature)>(StringComparer.Ordinal);
            var mathTypePreflightFailures =
                new System.Collections.Concurrent.ConcurrentDictionary<
                    string,
                    (string Latex, string Error)>(StringComparer.Ordinal);
            var mathTypePreflightTasks = new List<Task>();
            var mathTypePreflightGate =
                new System.Threading.SemaphoreSlim(
                    Math.Min(4, Math.Max(1, Environment.ProcessorCount)));

            phaseStartMilliseconds = redrawPhaseWatch.ElapsedMilliseconds;
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
            WriteRedrawAcceptanceLog(
                $"redraw-phase createSessionsMs={redrawPhaseWatch.ElapsedMilliseconds - phaseStartMilliseconds} "
                + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds} sessions={pendingTemplates.Count}");

            if (pendingTemplates.Count > 0)
            {
                phaseStartMilliseconds = redrawPhaseWatch.ElapsedMilliseconds;
                WriteRedrawAcceptanceLog(
                    $"render-batch-start unique={pendingTemplates.Count} total={plan.Targets.Count}");
                SetStatus($"正在批量渲染 {pendingTemplates.Count} 个独立公式…");
                await client.OpenConverterBatchAsync(
                        pendingTemplates.Select(item => item.Session.Id).ToList(),
                        lifetime.Token)
                    .ConfigureAwait(false);

                var renderCompleted = 0;
                progress.Report("正在渲染公式…", 0, pendingTemplates.Count);
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
                        if (string.Equals(
                                objectMode,
                                FormulaOleContract.MathTypeOleMode,
                                StringComparison.Ordinal))
                        {
                            var preflightKey = pending.Key;
                            var preflightMathMl = template.MathMl
                                ?? throw new InvalidDataException(
                                    $"MathType 重绘模板 {pending.Key} 缺少 MathML。");
                            var preflightInline = string.Equals(
                                template.Session.DisplayMode,
                                "inline",
                                StringComparison.OrdinalIgnoreCase);
                            var preflightFontSizePt = template.Session.FontSizePt;
                            var preflightLatex = template.Session.ToMetadata().Latex;
                            mathTypePreflightTasks.Add(Task.Run(
                                async () =>
                                {
                                    await mathTypePreflightGate
                                        .WaitAsync(lifetime.Token)
                                        .ConfigureAwait(false);
                                    try
                                    {
                                        mathTypePreflightPrepared[preflightKey] =
                                            WordFormulaService.PrepareStandaloneMathTypeOleData(
                                                preflightMathMl,
                                                preflightInline,
                                                preflightFontSizePt,
                                                preflightLatex);
                                    }
                                    catch (Exception error)
                                    {
                                        var detail = string.IsNullOrWhiteSpace(error.Message)
                                            ? error.GetType().Name
                                            : error.Message.Trim();
                                        mathTypePreflightFailures[preflightKey] =
                                            (preflightLatex, detail);
                                    }
                                    finally
                                    {
                                        mathTypePreflightGate.Release();
                                    }
                                },
                                lifetime.Token));
                        }
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
                    progress.Report("正在渲染公式…", ++renderCompleted, pendingTemplates.Count);
                }
                WriteRedrawAcceptanceLog(
                    $"redraw-phase renderBatchMs={redrawPhaseWatch.ElapsedMilliseconds - phaseStartMilliseconds} "
                    + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds}");
            }

            if (string.Equals(
                    objectMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                && rendered.Count > 0)
            {
                // Each MathType preflight starts as soon as that formula's
                // converter Session finishes. This overlaps pure MTEF/CFB codec
                // work with the remaining frontend conversions while preserving
                // the rule that every formula is fully validated before Word is
                // mutated.
                var preflightAwaitStart = redrawPhaseWatch.ElapsedMilliseconds;
                progress.Report("正在生成并校验 MathType 数据…");
                await Task.WhenAll(mathTypePreflightTasks)
                    .ConfigureAwait(false);
                WriteRedrawAcceptanceLog(
                    $"redraw-phase mathtypePreflightAwaitMs={redrawPhaseWatch.ElapsedMilliseconds - preflightAwaitStart} "
                    + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds} tasks={mathTypePreflightTasks.Count}");

                foreach (var item in rendered)
                {
                    if (mathTypePreflightPrepared.TryGetValue(
                            item.Key,
                            out var standalone))
                    {
                        mathTypeStandalone[item.Key] = standalone;
                        continue;
                    }
                    if (!mathTypePreflightFailures.TryGetValue(
                            item.Key,
                            out var failure))
                        continue;
                    renderFailures[item.Key] = failure.Error;
                    WriteRedrawAcceptanceLog(
                        $"mathtype-preflight-skipped latex={failure.Latex} error={failure.Error}");
                }

                // Complete native layout before mutating Word. A missing private
                // service must not silently change the equation's layout engine.
                var nativePreviewAvailable = MathTypeNativePreviewRenderer.IsAvailable;
                if (!nativePreviewAvailable)
                {
                    throw new InvalidOperationException(MathTypeNativePreviewRenderer.UnavailableMessage);
                }
                else if (mathTypeStandalone.Count > 0)
                {
                    SetStatus($"正在批量生成 {mathTypeStandalone.Count} 个 MathType 原生预览…");
                    var nativePreviewInputs =
                        mathTypeStandalone.ToDictionary(
                            item => item.Key,
                            item => item.Value.Generated.Mtef,
                            StringComparer.Ordinal);

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
                    var missingPreviewKeys = mathTypeStandalone.Keys
                        .Where(key => !nativePreviews.ContainsKey(key))
                        .ToArray();
                    if (!renderedAllNativePreviews || missingPreviewKeys.Length > 0)
                    {
                        foreach (var preview in nativePreviews.Values)
                            preview.Dispose();
                        nativePreviewWatch.Stop();
                        throw new InvalidOperationException(MathTypeNativePreviewRenderer.FailureMessage);
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

            phaseStartMilliseconds = redrawPhaseWatch.ElapsedMilliseconds;
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
                var hasMathTypeStandalone =
                    mathTypeStandalone.TryGetValue(
                        key,
                        out var standalone);
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
                    MathTypeStandaloneGenerated = hasMathTypeStandalone
                        ? standalone.Generated
                        : null,
                    MathTypeStandaloneCompoundFile = hasMathTypeStandalone
                        ? standalone.CompoundFile
                        : null,
                    MathTypeStandaloneSemanticSignature = hasMathTypeStandalone
                        ? standalone.SemanticSignature
                        : null,
                });
            }
            WriteRedrawAcceptanceLog(
                $"redraw-phase prepareTargetsMs={redrawPhaseWatch.ElapsedMilliseconds - phaseStartMilliseconds} "
                + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds} prepared={prepared.Count}");

            if (prepared.Count == 0)
            {
                WriteRedrawAcceptanceLog(
                    $"redraw-complete formulas=0 unique=0 skipped={skippedTargets.Count} "
                    + "renderAverageMs=0 renderMaxMs=0 insertTotalMs=0 insertMaxMs=0");
                SetStatus(
                    $"LaTeX 重绘未转换任何公式：{skippedTargets.Count} 个公式无法重绘，均已保留原代码。");
                progress.Dispose();
                await ShowLatexRedrawFailuresAsync(
                        dispatcher,
                        skippedTargets)
                    .ConfigureAwait(false);
                _ = CleanupConverterSessionsBestEffortAsync(
                    client,
                    converterSessionIds.ToArray(),
                    lifetime.Token);
                return;
            }

            plan.Targets = plan.Targets
                .Where(target => prepared.ContainsKey(target.Id))
                .ToList();
            SetStatus("公式渲染完成，正在原位写入 Word…");
            BeginFormulaFormatMutation();
            phaseStartMilliseconds = redrawPhaseWatch.ElapsedMilliseconds;
            WordLatexRedrawResult result;
            try
            {
                result = await dispatcher.InvokeAsync(
                        () => service.ApplyLatexRedrawPlan(plan, prepared))
                    .ConfigureAwait(false);
            }
            finally { EndFormulaFormatMutation(); }
            foreach (var failure in result.Failures)
            {
                var failedTarget = plan.Targets.FirstOrDefault(
                    target => string.Equals(
                        target.Id,
                        failure.TargetId,
                        StringComparison.Ordinal));
                if (failedTarget is not null)
                    skippedTargets.Add((failedTarget, failure.Error));
                WriteRedrawAcceptanceLog(
                    $"apply-skipped targetId={failure.TargetId} "
                    + $"latex={failure.Latex} error={failure.Error}");
            }
            foreach (var warning in result.Warnings)
                WriteRedrawAcceptanceLog($"apply-warning {warning}");

            WriteRedrawAcceptanceLog(
                $"redraw-phase wordApplyMs={redrawPhaseWatch.ElapsedMilliseconds - phaseStartMilliseconds} "
                + $"elapsedMs={redrawPhaseWatch.ElapsedMilliseconds} "
                + $"sourceMutationTotalMs={result.TotalSourceMutationMilliseconds} "
                + $"insertTotalMs={result.TotalInsertMilliseconds} "
                + string.Join(
                    " ",
                    result.BatchTimings.Select(item =>
                        $"apply.{item.Key}Ms={item.Value}")));
            for (var timingIndex = 0;
                 timingIndex < result.InsertMilliseconds.Count;
                 timingIndex++)
            {
                if (timingIndex < result.SourceMutationMilliseconds.Count)
                {
                    WriteRedrawAcceptanceLog(
                        $"source-mutation-timing index={timingIndex + 1}/{result.InsertMilliseconds.Count} "
                        + $"elapsedMs={result.SourceMutationMilliseconds[timingIndex]}");
                }
                WriteRedrawAcceptanceLog(
                    $"insert-timing index={timingIndex + 1}/{result.InsertMilliseconds.Count} "
                    + $"elapsedMs={result.InsertMilliseconds[timingIndex]}");
                if (timingIndex < result.InsertMetricDetails.Count
                    && !string.IsNullOrWhiteSpace(
                        result.InsertMetricDetails[timingIndex]))
                {
                    WriteRedrawAcceptanceLog(
                        $"insert-metrics index={timingIndex + 1}/{result.InsertMilliseconds.Count} "
                        + result.InsertMetricDetails[timingIndex]);
                }
            }
            // Converter sessions are temporary render workspaces. The Word
            // transaction is already committed, so do not serialize one local
            // completion request per formula on the user-visible redraw path.
            // Format conversion uses the same batched best-effort cleanup model.
            _ = CleanupConverterSessionsBestEffortAsync(
                client,
                converterSessionIds.ToArray(),
                lifetime.Token);
            var uniqueRenderCount = Math.Max(1, rendered.Count);
            var averageRenderMilliseconds = totalRenderMilliseconds / uniqueRenderCount;
            WriteRedrawAcceptanceLog(
                $"redraw-complete formulas={result.FormulaCount} unique={rendered.Count} "
                + $"skipped={skippedTargets.Count} "
                + $"renderAverageMs={averageRenderMilliseconds} renderMaxMs={maxRenderMilliseconds} "
                + $"insertTotalMs={result.TotalInsertMilliseconds} insertMaxMs={result.MaxInsertMilliseconds} "
                + $"sourceMutationTotalMs={result.TotalSourceMutationMilliseconds} "
                + $"sourceMutationMaxMs={(result.SourceMutationMilliseconds.Count == 0 ? 0 : result.SourceMutationMilliseconds.Max())}");
            var performanceSuffix = maxRenderMilliseconds <= 250
                ? $"渲染最大 {maxRenderMilliseconds} ms/公式"
                : $"渲染最大 {maxRenderMilliseconds} ms/公式（本机超过 250 ms 目标）";
            var skippedSuffix = skippedTargets.Count == 0
                ? string.Empty
                : $"；{skippedTargets.Count} 个无法重绘的公式已保留原 LaTeX";
            var warningSuffix = result.Warnings.Count == 0
                ? string.Empty
                : $"；另有 {result.Warnings.Count} 个收尾警告";
            SetStatus(
                $"LaTeX 重绘完成：{result.FormulaCount} 个公式已转换为 {modeLabel}{skippedSuffix}{warningSuffix}；{performanceSuffix}。");
            progress.Dispose();
            await ShowLatexRedrawFailuresAsync(
                    dispatcher,
                    skippedTargets,
                    result.Warnings)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException error)
        {
            WriteRedrawAcceptanceLog("redraw-cancelled " + error);
            SetStatus("VisualTeX LaTeX 重绘已取消。");
        }
        catch (Exception error)
        {
            WriteRedrawAcceptanceLog("redraw-failed " + error);
            var failureMessage = OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX LaTeX 重绘失败",
                "VisualTeX LaTeX redraw failed");
            foreach (var sessionId in converterSessionIds)
            {
                try
                {
                    await client.FailAsync(sessionId, failureMessage, CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch { }
            }
            SetStatus(failureMessage);
            if (!string.Equals(
                    Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                progress.Dispose();
                try
                {
                    await dispatcher.InvokeAsync(() =>
                    {
                        System.Windows.Forms.MessageBox.Show(
                            failureMessage,
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
            progress.Dispose();
            _operationGate.Release();
        }
    }

    private async Task ShowLatexRedrawFailuresAsync(
        OfficeUiDispatcher dispatcher,
        IReadOnlyList<(WordLatexRedrawTarget Target, string Error)> failures,
        IReadOnlyList<string>? warnings = null)
    {
        warnings ??= Array.Empty<string>();
        if ((failures.Count == 0 && warnings.Count == 0)
            || string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                "1",
                StringComparison.Ordinal))
            return;

        static string Clip(string value, int maximum)
        {
            var normalized = (value ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            return normalized.Length <= maximum
                ? normalized
                : normalized.Substring(0, maximum) + "…";
        }

        const int maximumFailures = 12;
        const int maximumWarnings = 6;
        var failureLines = failures
            .Take(maximumFailures)
            .Select((failure, index) =>
                $"{index + 1}. {Clip(failure.Target.Latex, 160)}\r\n"
                + $"   {Clip(failure.Error, 320)}")
            .ToArray();
        var warningLines = warnings
            .Take(maximumWarnings)
            .Select((warning, index) =>
                $"{index + 1}. {Clip(warning, 320)}")
            .ToArray();

        var chineseParts = new List<string>();
        var englishParts = new List<string>();
        if (failureLines.Length > 0)
        {
            chineseParts.Add(
                $"以下 {failures.Count} 个公式无法重绘，原 LaTeX 已保留；其余公式已继续处理：\r\n\r\n"
                + string.Join("\r\n\r\n", failureLines)
                + (failures.Count > maximumFailures
                    ? $"\r\n\r\n……另有 {failures.Count - maximumFailures} 个失败公式未展开。"
                    : string.Empty));
            englishParts.Add(
                $"{failures.Count} formula(s) could not be redrawn. Their original LaTeX was preserved and the remaining formulas continued:\r\n\r\n"
                + string.Join("\r\n\r\n", failureLines)
                + (failures.Count > maximumFailures
                    ? $"\r\n\r\n…plus {failures.Count - maximumFailures} additional failed formula(s)."
                    : string.Empty));
        }

        if (warningLines.Length > 0)
        {
            chineseParts.Add(
                $"收尾警告：\r\n\r\n"
                + string.Join("\r\n", warningLines)
                + (warnings.Count > maximumWarnings
                    ? $"\r\n……另有 {warnings.Count - maximumWarnings} 条警告未展开。"
                    : string.Empty));
            englishParts.Add(
                $"Finalization warning(s):\r\n\r\n"
                + string.Join("\r\n", warningLines)
                + (warnings.Count > maximumWarnings
                    ? $"\r\n…plus {warnings.Count - maximumWarnings} additional warning(s)."
                    : string.Empty));
        }

        var message = T(
            string.Join("\r\n\r\n", chineseParts),
            string.Join("\r\n\r\n", englishParts));
        try
        {
            await dispatcher.InvokeAsync(() =>
            {
                System.Windows.Forms.MessageBox.Show(
                    message,
                    T("VisualTeX LaTeX 重绘", "VisualTeX LaTeX Redraw"),
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return true;
            }).ConfigureAwait(false);
        }
        catch
        {
            // Redraw has already finished. A transient UI notification failure
            // must not turn successful document mutations into an operation error.
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

        var progress = await dispatcher.InvokeAsync(
            () => BeginOperationProgress("VisualTeX 公式转为 LaTeX", service)).ConfigureAwait(false);

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
                progress.Dispose();
                var confirmed = await dispatcher.InvokeAsync(() =>
                    System.Windows.Forms.MessageBox.Show(
                        progress.Owner,
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
                progress = await dispatcher.InvokeAsync(
                    () => BeginOperationProgress("VisualTeX 公式转为 LaTeX", service, "正在准备恢复 LaTeX…")).ConfigureAwait(false);
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
            var failureMessage = OfficePluginLanguage.FormatOperationError(
                error,
                "VisualTeX 公式转为 LaTeX 失败",
                "VisualTeX equation-to-LaTeX conversion failed");
            SetStatus(failureMessage);
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
                            failureMessage,
                            T("VisualTeX 公式转为 LaTeX", "VisualTeX Equation to LaTeX"),
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error);
                        return true;
                    }).ConfigureAwait(false);
                }
                catch { }
            }
        }
        finally { progress.Dispose(); _operationGate.Release(); }
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

        using var progress = await dispatcher.InvokeAsync(
            () => BeginOperationProgress("VisualTeX 批量导入", service)).ConfigureAwait(false);

        var rendered = new Dictionary<string, RenderedWordBulkFormulaTemplate>(
            StringComparer.Ordinal);
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(
            StringComparer.Ordinal);
        var mathTypePreviews =
            new Dictionary<string, MathTypeNativePreviewRenderer.Result>(
                StringComparer.Ordinal);
        var mathTypeStandalone =
            new Dictionary<string, (
                MathTypeMtefCodec.RewriteResult Generated,
                byte[] CompoundFile,
                string SemanticSignature)>(
                StringComparer.Ordinal);
        var ownedSessionIds = new List<string>();
        var operationStopwatch = Stopwatch.StartNew();
        var operationGateHeld = true;
        void ReleaseOperationGate()
        {
            if (!operationGateHeld) return;
            operationGateHeld = false;
            progress.Dispose();
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
                    throw new InvalidOperationException(MathTypeNativePreviewRenderer.UnavailableMessage);
                }
                else
                {
                    SetStatus($"正在批量生成 {formulaRuns.Count} 个 MathType 原生预览…");
                    var nativePreviewInputs =
                        new Dictionary<string, byte[]>(StringComparer.Ordinal);
                    foreach (var item in rendered)
                    {
                        var standalone =
                            WordFormulaService.PrepareStandaloneMathTypeOleData(
                                item.Value.MathMl,
                                string.Equals(
                                    item.Value.Session.DisplayMode,
                                    "inline",
                                    StringComparison.OrdinalIgnoreCase),
                                item.Value.Session.FontSizePt,
                                item.Value.Session.ToMetadata().Latex);
                        mathTypeStandalone[item.Key] = standalone;
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
                        throw new InvalidOperationException(MathTypeNativePreviewRenderer.FailureMessage);
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
                var hasMathTypeStandalone =
                    mathTypeStandalone.TryGetValue(
                        key,
                        out var standalone);
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
                    MathTypeStandaloneGenerated = hasMathTypeStandalone
                        ? standalone.Generated
                        : null,
                    MathTypeStandaloneCompoundFile = hasMathTypeStandalone
                        ? standalone.CompoundFile
                        : null,
                    MathTypeStandaloneSemanticSignature = hasMathTypeStandalone
                        ? standalone.SemanticSignature
                        : null,
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
        VisualTeXSessionClient client, WordBulkRun run, string objectMode,
        string? sourceDocumentId, double fontSizePt, CancellationToken cancellationToken,
        string? formulaLetterFont = null, string? formulaChineseFont = null)
    {
        var request = CreateBulkFormulaConversionRequest(run, objectMode, sourceDocumentId,
            fontSizePt, formulaLetterFont, formulaChineseFont);
        WriteBulkAcceptanceLog($"converter-create display={run.DisplayMode} mode={objectMode} latex={run.Latex}");
        var session = await client.CreateSessionAsync(request, cancellationToken).ConfigureAwait(false);
        WriteBulkAcceptanceLog($"converter-created sessionId={session.Id}");
        return session;
    }

    private static CreateVstoSessionRequest CreateBulkFormulaConversionRequest(
        WordBulkRun run, string objectMode, string? sourceDocumentId, double fontSizePt,
        string? formulaLetterFont = null, string? formulaChineseFont = null)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        var line = new FormulaLine
        {
            Id = Guid.NewGuid().ToString("D"),
            Latex = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(run.Latex),
        };
        var originalMetadata = CreateBulkFormulaMetadata(
            formulaId,
            line,
            run,
            fontSizePt);
        originalMetadata.FormulaLetterFont = formulaLetterFont;
        originalMetadata.FormulaChineseFont = formulaChineseFont;
        return new CreateVstoSessionRequest
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
            };
    }

    private static RenderedWordBulkFormulaTemplate MaterializeBulkFormulaTemplate(
        VisualTeXSessionClient client,
        WordBulkRun run,
        string objectMode,
        OfficeSessionDocument session,
        bool materializePreview = true)
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
            RawRenderWidthPx = export.Width,
            RawRenderHeightPx = export.Height,
            RawBaselinePx = export.Baseline,
        };
        var needsVectorPreview = materializePreview && (string.Equals(
                objectMode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
            || string.Equals(
                objectMode,
                FormulaOleContract.MathTypeOleMode,
                StringComparison.Ordinal));
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
            Latex = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(run.Latex),
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

    private async Task RepairInlineFormulaBaselinesAsync()
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
            SetStatus(T(
                "VisualTeX 正在执行其他 Word 操作，请稍候再试。",
                "VisualTeX is busy with another Word operation. Please try again shortly."));
            return;
        }

        WordInlineBaselineRepairReport? report = null;
        try
        {
            report = await dispatcher.InvokeAsync(service.ScanInlineBaselineRepairs)
                .ConfigureAwait(false);
            var decision = await dispatcher.InvokeAsync(() =>
            {
                if (report.TotalRepairCount == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        T(
                            $"扫描完成。已检查 {report.ScannedOleCount} 个行内对象，无需调整公式位置。{report.SkippedCount} 个对象缺少可确定的基线信息，未作推测性修改。",
                            $"Scan complete. {report.ScannedOleCount} inline objects were checked; no position change is needed. {report.SkippedCount} objects without a determinable baseline were left unchanged."),
                        T("VisualTeX 修复公式偏移", "VisualTeX Formula Offset Repair"),
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return 0;
                }

                var answer = System.Windows.Forms.MessageBox.Show(
                    T(
                        $"需要调整 {report.TotalRepairCount} 个行内公式的位置（VisualTeX {report.VisualTeXRepairCount}，MathType {report.MathTypeRepairCount}）。\r\n\r\n"
                        + "按当前段落的字符对齐规则校正公式：普通基线段落使用公式渲染基线，字符中心对齐段落不重复下移。只调整公式对象字符的位置，不重绘、不改字号或对象尺寸，不改段落格式、编号和引用。\r\n\r\n"
                        + $"无法确定的对象：{report.SkippedCount} 个，保持不变。是否继续？",
                        $"Adjust {report.TotalRepairCount} inline equation positions (VisualTeX {report.VisualTeXRepairCount}, MathType {report.MathTypeRepairCount}).\r\n\r\n"
                        + "Use the current paragraph's character-alignment rule: render baselines for baseline-aligned text, without an extra descent for character-center alignment. Only equation-character positions are changed. No rendering, font-size, extent, paragraph-format, numbering or reference changes.\r\n\r\n"
                        + $"Objects without a determinable baseline: {report.SkippedCount}, left unchanged. Continue?"),
                    T("VisualTeX 修复公式偏移", "VisualTeX Formula Offset Repair"),
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question,
                    System.Windows.Forms.MessageBoxDefaultButton.Button2);
                return answer == System.Windows.Forms.DialogResult.Yes ? 1 : -1;
            }).ConfigureAwait(false);

            if (decision <= 0)
            {
                if (decision == -1)
                {
                    SetStatus(T(
                        "已取消公式偏移修复，Word 文档未修改。",
                        "Formula offset repair was cancelled; the Word document was not changed."));
                }
                else if (decision == 0)
                {
                    SetStatus(T(
                        "公式偏移扫描完成，未发现需要修复的垂直布局异常。",
                        "Formula offset scan complete; no vertical-layout repair was needed."));
                }
                return;
            }

            SetStatus(T("正在校正公式字符的位置…", "Correcting equation character positions…"));

            BeginFormulaFormatMutation();
            int repaired;
            try
            {
                repaired = await dispatcher.InvokeAsync(
                        () => service.RepairInlineBaselineRepairs(report))
                    .ConfigureAwait(false);
            }
            finally { EndFormulaFormatMutation(); }

            SetStatus(T(
                $"已校正 {repaired} 个行内公式的位置，未重绘公式或修改尺寸。",
                $"Corrected {repaired} inline equation positions without rendering or resizing."));
        }
        catch (OperationCanceledException)
        {
            SetStatus(T("VisualTeX 公式偏移修复已取消。", "VisualTeX formula offset repair was cancelled."));
        }
        catch (Exception error)
        {
            SetStatus(T(
                $"修复公式偏移失败：{error.Message}",
                $"Formula offset repair failed: {error.Message}"));
            await ShowWordOperationErrorAsync(
                    dispatcher,
                    T("修复公式偏移", "Formula Offset Repair"),
                    error)
                .ConfigureAwait(false);
        }
        finally
        {
            _operationGate.Release();
        }
    }

    private async Task RepairMathTypeDisplaysAsync()
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
            SetStatus(T(
                "VisualTeX 正在执行其他 Word 操作，请稍候再试。",
                "VisualTeX is busy with another Word operation. Please try again shortly."));
            return;
        }

        WordOperationProgress? progress = null;
        try
        {
            SetStatus(T("正在扫描 MathType 公式的显示图像…", "Scanning MathType equation pictures…"));
            progress = await dispatcher.InvokeAsync(() => BeginOperationProgress(
                    T("VisualTeX 修复公式显示", "VisualTeX Formula Display Repair"), service,
                    T("正在读取 MathType 公式…", "Reading MathType equations…")))
                .ConfigureAwait(false);
            var report = await dispatcher.InvokeAsync(service.ScanMathTypeDisplayRepairs)
                .ConfigureAwait(false);
            // The confirmation must not compete with the progress window.
            progress.Dispose();
            progress = null;
            var skipped = report.UnreadableCount + report.UnrenderableCount;
            var unreadable = skipped == 0
                ? string.Empty
                : T(
                    $"\r\n\r\n{skipped} 个 MathType 公式无法读取或无法由 MathType 重新排版，保持不变。",
                    $"\r\n\r\n{skipped} MathType equations could not be read or laid out again by MathType and were left unchanged.");
            var decision = await dispatcher.InvokeAsync(() =>
            {
                if (report.Targets.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        T(
                            $"扫描完成。已检查 {report.MathTypeCount} 个 MathType 公式，它们的显示图像与 MathType 自身的排版一致，无需修复。" + unreadable,
                            $"Scan complete. {report.MathTypeCount} MathType equations were checked; their pictures already match MathType's own layout." + unreadable),
                        T("VisualTeX 修复公式显示", "VisualTeX Formula Display Repair"),
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return 0;
                }

                var answer = System.Windows.Forms.MessageBox.Show(
                    T(
                        $"{report.Targets.Count} 个 MathType 公式的显示图像与 MathType 自身的排版不一致（例如缺少 MathType 字体时生成的错误符号）。\r\n\r\n"
                        + "将用 VisualTeX 内置的 MathType 排版服务按公式原有数据重新生成这些图像。只替换显示图像并按 MathType 基线重新对齐，公式数据、编号和引用保持不变；每个公式可单独撤销。"
                        + unreadable + "\r\n\r\n是否继续？",
                        $"The pictures of {report.Targets.Count} MathType equations differ from MathType's own layout (for example, wrong symbols drawn while MathType's fonts were missing).\r\n\r\n"
                        + "VisualTeX's built-in MathType renderer will redraw these pictures from each equation's stored data. Only the pictures are replaced and realigned on MathType's baseline; equation data, numbering and references stay unchanged, and each equation can be undone separately."
                        + unreadable + "\r\n\r\nContinue?"),
                    T("VisualTeX 修复公式显示", "VisualTeX Formula Display Repair"),
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question,
                    System.Windows.Forms.MessageBoxDefaultButton.Button2);
                return answer == System.Windows.Forms.DialogResult.Yes ? 1 : -1;
            }).ConfigureAwait(false);

            if (decision <= 0)
            {
                SetStatus(decision == -1
                    ? T("已取消公式显示修复，Word 文档未修改。", "Formula display repair was cancelled; the Word document was not changed.")
                    : T("公式显示扫描完成，无需修复。", "Formula display scan complete; no repair was needed."));
                return;
            }

            SetStatus(T("正在用 MathType 排版重新生成公式图像…", "Redrawing MathType equation pictures…"));
            progress = await dispatcher.InvokeAsync(() => BeginOperationProgress(
                    T("VisualTeX 修复公式显示", "VisualTeX Formula Display Repair"), service,
                    T("正在重新生成公式图像…", "Redrawing equation pictures…")))
                .ConfigureAwait(false);
            BeginFormulaFormatMutation();
            int repaired;
            try
            {
                repaired = await dispatcher.InvokeAsync(
                        () => service.RepairMathTypeDisplays(report))
                    .ConfigureAwait(false);
            }
            finally { EndFormulaFormatMutation(); }
            progress.Dispose();
            progress = null;

            SetStatus(T(
                $"已用 MathType 排版重新生成 {repaired} 个公式的显示图像，公式数据未修改。",
                $"Redrew the pictures of {repaired} MathType equations; equation data was not changed."));
        }
        catch (OperationCanceledException)
        {
            SetStatus(T("VisualTeX 公式显示修复已取消。", "VisualTeX formula display repair was cancelled."));
        }
        catch (Exception error)
        {
            progress?.Dispose();
            progress = null;
            SetStatus(T(
                $"修复公式显示失败：{error.Message}",
                $"Formula display repair failed: {error.Message}"));
            await ShowWordOperationErrorAsync(
                    dispatcher,
                    T("修复公式显示", "Formula Display Repair"),
                    error)
                .ConfigureAwait(false);
        }
        finally
        {
            progress?.Dispose();
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

    // Update Equation Format: scan (read-only), show what will change, then
    // bring every older number structure into the current format.
    private async Task UpdateEquationFormatAsync()
    {
        var dispatcher = _dispatcher;
        var service = _formulaService;
        if (dispatcher is null || service is null) return;
        try
        {
            var report = await dispatcher.InvokeAsync(service.ScanEquationFormat)
                .ConfigureAwait(false);
            var proceed = await dispatcher.InvokeAsync(() =>
            {
                var title = T("VisualTeX 更新公式格式", "VisualTeX Update Equation Format");
                if (report.Total == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        T("本文档的公式已经是当前格式，无需更新。",
                          "The equations in this document already use the current format."),
                        title,
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }
                var zh = "将本文档的公式更新为当前 VisualTeX 格式：\r\n\r\n";
                var en = "Update this document's equations to the current VisualTeX format:\r\n\r\n";
                if (report.RetiredLayouts > 0)
                {
                    zh += $"旧版编号版式：{report.RetiredLayouts} 个\r\n";
                    en += $"Older number layouts: {report.RetiredLayouts}\r\n";
                }
                if (report.NonCanonicalNumbers > 0)
                {
                    zh += $"需重写的公式编号：{report.NonCanonicalNumbers} 个\r\n";
                    en += $"Equation numbers to rewrite: {report.NonCanonicalNumbers}\r\n";
                }
                if (report.LegacyFormatVariable)
                {
                    zh += "旧版编号格式变量：1 个\r\n";
                    en += "Older numbering-style variable: 1\r\n";
                }
                zh += "\r\n公式内容和类型不变。是否继续？";
                en += "\r\nEquation content and kinds do not change. Continue?";
                return System.Windows.Forms.MessageBox.Show(
                    T(zh, en),
                    title,
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question,
                    System.Windows.Forms.MessageBoxDefaultButton.Button2)
                    == System.Windows.Forms.DialogResult.Yes;
            }).ConfigureAwait(false);
            if (!proceed) return;
            var count = await dispatcher.InvokeAsync(service.UpgradeEquationFormat)
                .ConfigureAwait(false);
            SetStatus(T($"公式格式已更新（{count} 个编号）。", $"Equation format updated ({count} numbers)."));
        }
        catch (Exception error)
        {
            SetStatus(T($"更新公式格式失败：{error.Message}", $"Equation format update failed: {error.Message}"));
            await ShowWordOperationErrorAsync(dispatcher, "更新公式格式", error).ConfigureAwait(false);
        }
        finally { InvalidateEquationNumberFormatControls(); }
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
                    var canonicalTargets = service.GetCanonicalEquationReferenceTargets(document);
                    var visualTexTargets = canonicalTargets
                        .Where(target => target.Source == EquationReferenceSource.VisualTeX)
                        .ToArray();
                    var wordOmmlTargets = canonicalTargets
                        .Where(target => target.Source == EquationReferenceSource.WordOmml)
                        .ToArray();
                    var mathTypeTargets = MathTypeEquationReferences.GetTargets(document);
                    if (visualTexTargets.Length == 0
                        && wordOmmlTargets.Length == 0
                        && mathTypeTargets.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            T(
                                "当前文档没有可引用的带编号公式。请先插入带编号的 VisualTeX、Word OMML 或 MathType 行间公式。",
                                "The current document has no numbered equations to reference. Insert a numbered VisualTeX, Word OMML, or MathType display equation first."),
                            "VisualTeX",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return string.Empty;
                    }

                    static string DescribeReferenceTarget(EquationReferenceTarget target) =>
                        target.Source switch
                        {
                            EquationReferenceSource.MathType =>
                                $"MathType 公式 {target.NumberText}",
                            EquationReferenceSource.WordOmml =>
                                $"Word OMML 公式 ({target.NumberText})",
                            _ =>
                                $"VisualTeX 公式 ({target.NumberText})",
                        };

                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        var requestedSource = Environment.GetEnvironmentVariable(
                            "VISUALTEX_VSTO_REFERENCE_SOURCE");
                        IReadOnlyList<EquationReferenceTarget> targets =
                            string.Equals(
                                requestedSource,
                                "mathtype",
                                StringComparison.OrdinalIgnoreCase)
                                ? mathTypeTargets
                                : string.Equals(
                                      requestedSource,
                                      "omml",
                                      StringComparison.OrdinalIgnoreCase)
                                  || string.Equals(
                                      requestedSource,
                                      "wordomml",
                                      StringComparison.OrdinalIgnoreCase)
                                    ? wordOmmlTargets
                                    : string.Equals(
                                          requestedSource,
                                          "visualtex",
                                          StringComparison.OrdinalIgnoreCase)
                                        ? visualTexTargets
                                        : visualTexTargets.Length > 0
                                            ? visualTexTargets
                                            : wordOmmlTargets.Length > 0
                                                ? wordOmmlTargets
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
                            service.InsertEquationReference(
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
                        wordOmmlTargets,
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
        var englishOperation = operation switch
        {
            "更新公式编号" => "equation-number update",
            "更新公式格式" => "equation-format update",
            "编号格式" => "equation-number format update",
            "插入公式引用" => "equation-reference insertion",
            _ when !OfficePluginLanguage.ContainsCjk(operation) => operation,
            _ => "operation",
        };
        var userMessage = OfficePluginLanguage.FormatOperationError(
            error,
            $"VisualTeX {operation}失败",
            $"VisualTeX Word {englishOperation} failed");
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
                        userMessage,
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
            SetStatus(OfficePluginLanguage.FormatOperationError(
                error,
                "删除 Word 公式失败",
                "Failed to delete the Word formula"));
        }
    }

    private void SetStatus(string message)
    {
        _operationProgress?.Report(OfficePluginLanguage.SafeStatusMessage(message));
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
        try { _keyboardHook?.Dispose(); } catch { }
        _keyboardHook = null;
        EndActivePasteUndoRecord(null, "addin-dispose");
        try { _deleteKeyHook?.Dispose(); } catch { }
        _deleteKeyHook = null;
        lock (_copyPasteGate)
        {
            _formulaCopySnapshot = null;
            _formulaCopyClipboardSequence = 0;
        }
        Interlocked.Increment(ref _copyPasteWatchGeneration);
        Interlocked.Increment(ref _copyPasteRepairGeneration);
        ClearNativeOleTarget();
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

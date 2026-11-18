namespace VisualTeX.WordVsto;

// The VisualTeX tab as WPS Writer shows it: Word's tab (RibbonXml) control for control - the same groups, names, sizes
// and order, large where Word's are large, and Word's three-row columns as three-row columns - with WPS's own
// equations called "WPS 公式" where Word says OMML. What WPS needs on top (12.1, Windows and Mac):
// - WPS shows no group titles and no menu section titles, so a menu item that Word names under a section title
//   carries that title in its own name ("所选公式转为 WPS 公式").
// - WPS does not stack small controls by itself: it packs consecutive vertical boxes into one column, so each column
//   is a vertical box of its own and side-by-side columns stand in a horizontal box. WPS stretches every control to
//   its column's width; a column shorter than its neighbour gets empty labels (U+3000) as its last rows, as WPS
//   centres a column's controls otherwise. Columns of three make the tab taller than WPS's own two-row tabs.
// - Icons are imageMso names WPS ships, chosen by what they show: WPS draws an add-in picture in a fixed 16 px box
//   without transparency and shows a "?" for an Office icon name it lacks.
// - WPS shows a large control's label on one line, however long, unless the control says enableLineBreak="true"
//   (Windows WPS 12.1); then a line break in the label starts a second line, as Word wraps its own large labels.
//   Every large label breaks once between words, after the space where the words are spaced: a WPS without
//   enableLineBreak drops the break and shows the label as it reads on one line.
// Same callbacks as Word's tab; the formula-size boxes take their items from Word's tab.
public sealed partial class ThisAddIn
{
    private const string RibbonXmlWps = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="OLE &#10;行内公式" size="large" enableLineBreak="true" screentip="插入 VisualTeX OLE 行内公式" supertip="在光标处插入 VisualTeX 行内公式（OLE 对象）；双击即可再次编辑。" imageMso="EquationScriptGallery" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="OLE &#10;行间公式" size="large" enableLineBreak="true" screentip="插入 VisualTeX OLE 行间公式" supertip="在单独的居中段落插入 VisualTeX 行间公式（OLE 对象）；双击即可再次编辑。" imageMso="EquationFractionGallery" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.InlineOmml" label="WPS &#10;行内公式" size="large" enableLineBreak="true" screentip="插入 WPS 行内公式" supertip="在 VisualTeX 编辑器中写好后，在光标处插入 WPS 自己的公式（OMML），之后也能用 WPS 的公式工具编辑。" imageMso="EquationProfessional" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.WordVsto.DisplayOmml" label="WPS &#10;行间公式" size="large" enableLineBreak="true" screentip="插入 WPS 行间公式" supertip="在 VisualTeX 编辑器中写好后，插入 WPS 自己的行间公式（OMML），之后也能用 WPS 的公式工具编辑。" imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.WordVsto.Edit" label="编辑&#10;所选公式" size="large" enableLineBreak="true" screentip="编辑所选公式" supertip="在 VisualTeX 编辑器中编辑所选的 VisualTeX 公式、WPS 公式或 MathType 公式；也可以双击公式。" imageMso="KSEquationEditor" onAction="OnEditSelected" />
          <box id="VisualTeX.WordVsto.NumberingBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.UpdateNumbers" label="更新公式编号" screentip="更新公式编号" supertip="重新计算当前文档中 VisualTeX 公式、WPS 公式和 MathType 公式的编号，并更新对它们的引用。" imageMso="UpdateField" onAction="OnUpdateEquationNumbers" />
            <menu id="VisualTeX.WordVsto.NumberFormat" label="编号格式" screentip="公式编号格式" supertip="选择后立即重排当前文档已有的公式编号，并用于之后插入的带编号公式。" imageMso="page_number_format_kd">
              <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="全文连续编号（1）" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="按章编号（1.1）" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="按章编号（1-1）" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="按节编号（1.1.1）" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="按节编号（1.1-1）" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            </menu>
            <button id="VisualTeX.WordVsto.InsertReference" label="插入公式引用" screentip="插入公式引用" supertip="从当前文档带编号的公式中选择一个，在光标处插入随编号自动更新的引用。" imageMso="CrossReference" onAction="OnInsertEquationReference" />
          </box>
          <menu id="VisualTeX.WordVsto.BatchNumbering" label="批量&#10;编号" size="large" enableLineBreak="true" screentip="为行间公式批量添加编号" supertip="按当前编号格式，为所选内容或全文中所有未编号的行间公式（VisualTeX OLE、WPS 公式、MathType）添加编号；已有编号保持不变并自动顺延，引用随之更新，可一步撤销。" imageMso="Numbering">
            <button id="VisualTeX.WordVsto.BatchNumberingSelection" label="为所选行间公式编号" onAction="OnNumberDisplayFormulasSelection" />
            <button id="VisualTeX.WordVsto.BatchNumberingDocument" label="为全文行间公式编号" onAction="OnNumberDisplayFormulasDocument" />
          </menu>
          <button id="VisualTeX.WordVsto.UpdateEquationFormat" label="更新&#10;公式格式" size="large" enableLineBreak="true" screentip="更新公式格式" supertip="把本文档中旧版 VisualTeX 写入的公式编号结构一次升级为当前格式，并让 WPS 公式使用“WPS 公式字体”选定的字体。公式内容与类型不变。" imageMso="Refresh" onAction="OnUpdateEquationFormat" />
          <button id="VisualTeX.WordVsto.BulkImport" label="批量&#10;导入" size="large" enableLineBreak="true" screentip="批量导入 LaTeX / Markdown" supertip="选择一个 LaTeX 或 Markdown 文档，正文插入为文字，公式插入为可编辑、可调字号的行内和行间公式。" imageMso="OpenFile" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.FormatConversionGroup" label="公式格式转换">
          <menu id="VisualTeX.WordVsto.ConvertFormulas" label="格式&#10;转换" size="large" enableLineBreak="true" screentip="公式格式转换" supertip="把范围内其他格式的公式（VisualTeX OLE、WPS 公式、MathType）全部转为所选格式，已是该格式的公式不变。不同来源的同一公式转换结果相同；字号、编号与引用按当前文档设置重建，可撤销。" imageMso="EquationTransferOMML">
            <button id="VisualTeX.WordVsto.ConvertSelectionToOmml" label="所选公式转为 WPS 公式" tag="wordOmml|selection" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToMathType" label="所选公式转为 MathType" tag="mathTypeOle|selection" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToVisualTeX" label="所选公式转为 VisualTeX OLE" tag="nativeOle|selection" onAction="OnConvertFormulasTo" />
            <menuSeparator id="VisualTeX.WordVsto.ConvertDocumentHeader" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToOmml" label="全部转为 WPS 公式" tag="wordOmml|document" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToMathType" label="全部转为 MathType" tag="mathTypeOle|document" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToVisualTeX" label="全部转为 VisualTeX OLE" tag="nativeOle|document" onAction="OnConvertFormulasTo" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.RedrawGroup" label="LaTeX 重绘">
          <menu id="VisualTeX.WordVsto.RedrawSelection" label="重绘&#10;所选" size="large" enableLineBreak="true" screentip="重绘所选公式或 LaTeX 代码" supertip="可将所选 LaTeX 代码原位重绘为 WPS 公式、VisualTeX OLE 或 MathType，也可将这三种公式恢复为 LaTeX 代码。" imageMso="EquationInsertLaTex">
            <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="LaTeX 重绘为 WPS 公式" screentip="原位替换为 WPS 公式" onAction="OnRedrawSelectionToOmml" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="LaTeX 重绘为 VisualTeX OLE" screentip="原位替换为可双击编辑的 VisualTeX OLE" onAction="OnRedrawSelectionToOle" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="LaTeX 重绘为 MathType" screentip="原位替换为 MathType 公式" onAction="OnRedrawSelectionToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawSelectionSeparator" />
            <button id="VisualTeX.WordVsto.FormulasToLatexSelection" label="所选公式转为 LaTeX 代码" screentip="所选公式恢复为 LaTeX 代码" supertip="把所选范围中的公式（VisualTeX OLE、WPS 公式、MathType 均可）原位恢复为 LaTeX 代码。" tag="selection" onAction="OnFormulasToLatex" />
          </menu>
          <menu id="VisualTeX.WordVsto.RedrawDocument" label="重绘&#10;全文" size="large" enableLineBreak="true" screentip="重绘全文公式或 LaTeX 代码" supertip="可将全文 LaTeX 代码原位重绘为 WPS 公式、VisualTeX OLE 或 MathType，也可将全文这三种公式恢复为 LaTeX 代码；开始前会再次确认。" imageMso="EquationTransferLaTex">
            <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="全文 LaTeX 重绘为 WPS 公式" onAction="OnRedrawDocumentToOmml" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="全文 LaTeX 重绘为 VisualTeX OLE" onAction="OnRedrawDocumentToOle" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="全文 LaTeX 重绘为 MathType" onAction="OnRedrawDocumentToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawDocumentSeparator" />
            <button id="VisualTeX.WordVsto.FormulasToLatexDocument" label="全文公式转为 LaTeX 代码" tag="document" onAction="OnFormulasToLatex" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="公式字号">
          <box id="VisualTeX.WordVsto.SizeColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.SelectedSizeBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.FontSizeDecrease" label="减小" imageMso="DecreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
              <comboBox id="VisualTeX.WordVsto.FontSize" label="字号" sizeString="初号（42 磅）" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
                __FONT_SIZE_ITEMS__
              </comboBox>
              <button id="VisualTeX.WordVsto.FontSizeIncrease" label="增大" imageMso="IncreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
            </box>
            <box id="VisualTeX.WordVsto.KindSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.WordVsto.InlineFontSize" label="行内" sizeString="初号（42 磅）" screentip="行内公式字号" supertip="选择字号后，当前文档中所有行内公式立即改为这个字号，可撤销；之后 LaTeX 重绘、批量导入和格式转换生成的行内公式也使用这个字号。选“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __INLINE_FONT_SIZE_ITEMS__
              </comboBox>
              <comboBox id="VisualTeX.WordVsto.DisplayFontSize" label="行间" sizeString="初号（42 磅）" screentip="行间公式字号" supertip="选择字号后，当前文档中所有行间公式（含编号公式）立即改为这个字号，可撤销；之后 LaTeX 重绘、批量导入和格式转换生成的行间公式（含编号公式）也使用这个字号。选“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __DISPLAY_FONT_SIZE_ITEMS__
              </comboBox>
              <menu id="VisualTeX.WordVsto.OmmlMathFont" label="WPS 公式字体" imageMso="batch_set_font_kd" tag="menu" screentip="WPS 公式字体" supertip="VisualTeX 写入 WPS 公式（OMML）时使用的数学字体。WPS 没有文档数学字体：所选字体用于之后 VisualTeX 插入、编辑或转换得到的 WPS 公式，已有公式在再次编辑后改用它。" getEnabled="GetOmmlMathFontEnabled">
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontFollow" label="跟随文档（不更改）" tag="follow" screentip="跟随文档" supertip="VisualTeX 不指定字体，WPS 公式使用 Cambria Math（默认）。" getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontCambria" label="Cambria Math" tag="cambria" screentip="Cambria Math" supertip="Office 默认的数学字体。" getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontStix" getLabel="GetOmmlMathFontLabel" tag="stix" screentip="STIX Two Math（内置）" supertip="VisualTeX 自带的 STIX Two Math（官方 TrueType 版，Times 风格，另存为 PDF 时可嵌入）。" getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontTermes" getLabel="GetOmmlMathFontLabel" tag="termes" screentip="TeX Gyre Termes Math（内置）" supertip="VisualTeX 自带的 TeX Gyre Termes Math（Times 风格，接近 LaTeX 的 newtxmath；TrueType 版，另存为 PDF 时可嵌入）。" getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontLatinModern" getLabel="GetOmmlMathFontLabel" tag="latin-modern" screentip="Latin Modern Math（内置）" supertip="VisualTeX 自带的 Latin Modern Math（TrueType 版，另存为 PDF 时可嵌入）。" getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
              </menu>
            </box>
          </box>
        </group>
        <group id="VisualTeX.WordVsto.CompatibilityRepairGroup" label="兼容修复">
          <button id="VisualTeX.WordVsto.RepairInlineBaselines" label="修复&#10;公式偏移" size="large" enableLineBreak="true" screentip="扫描并修复行内公式垂直偏移" supertip="扫描全文中的 MathType / VisualTeX 行内公式；VisualTeX 公式按自身保存的渲染基线校准，MathType 公式按文档稳定基线检测上浮或下沉。确认后仅修复垂直偏移，不修改公式内容、字号、编号或引用。" imageMso="VAlignCenter" onAction="OnRepairInlineFormulaBaselines" />
          <button id="VisualTeX.WordVsto.RepairMathTypeDisplays" label="修复&#10;公式显示" size="large" enableLineBreak="true" screentip="用 MathType 自身排版重新生成公式图像" supertip="扫描全文 MathType 公式，按公式原有数据重新生成显示图像，修复缺少 MathType 字体时出现的错误符号。只替换显示图像并按 MathType 基线重新对齐，公式数据、编号和引用保持不变。" imageMso="exporttopicture32" onAction="OnRepairMathTypeDisplays" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    private const string RibbonXmlWpsEnglish = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="Inline &#10;OLE" size="large" enableLineBreak="true" screentip="Insert an inline VisualTeX OLE equation" supertip="Insert a VisualTeX inline equation (OLE object) at the cursor; double-click it to edit it again." imageMso="EquationScriptGallery" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="Display &#10;OLE" size="large" enableLineBreak="true" screentip="Insert a display VisualTeX OLE equation" supertip="Insert a VisualTeX display equation (OLE object) in its own centered paragraph; double-click it to edit it again." imageMso="EquationFractionGallery" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.InlineOmml" label="Inline WPS &#10;Equation" size="large" enableLineBreak="true" screentip="Insert an inline WPS equation" supertip="Write it in the VisualTeX editor, then insert it at the cursor as WPS's own equation (OMML), which WPS's equation tools can edit too." imageMso="EquationProfessional" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.WordVsto.DisplayOmml" label="Display WPS &#10;Equation" size="large" enableLineBreak="true" screentip="Insert a display WPS equation" supertip="Write it in the VisualTeX editor, then insert it as WPS's own display equation (OMML), which WPS's equation tools can edit too." imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.WordVsto.Edit" label="Edit Selected &#10;Formula" size="large" enableLineBreak="true" screentip="Edit the selected equation" supertip="Edit the selected VisualTeX, WPS or MathType equation in the VisualTeX editor; double-clicking it does the same." imageMso="KSEquationEditor" onAction="OnEditSelected" />
          <box id="VisualTeX.WordVsto.NumberingBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.UpdateNumbers" label="Update Equation Numbers" screentip="Update equation numbers" supertip="Renumber the VisualTeX, WPS and MathType equations of the current document and update the references to them." imageMso="UpdateField" onAction="OnUpdateEquationNumbers" />
            <menu id="VisualTeX.WordVsto.NumberFormat" label="Number Format" screentip="Equation number format" supertip="Renumber the existing equations of this document in the chosen format right away and use it for numbered equations inserted later." imageMso="page_number_format_kd">
              <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="Continuous (1)" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="By chapter (1.1)" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="By chapter (1-1)" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="By section (1.1.1)" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="By section (1.1-1)" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            </menu>
            <button id="VisualTeX.WordVsto.InsertReference" label="Insert Equation Reference" screentip="Insert an equation reference" supertip="Choose a numbered equation of this document and insert, at the cursor, a reference that follows its number." imageMso="CrossReference" onAction="OnInsertEquationReference" />
          </box>
          <menu id="VisualTeX.WordVsto.BatchNumbering" label="Number &#10;Equations" size="large" enableLineBreak="true" screentip="Number display equations in bulk" supertip="Give every unnumbered display equation (VisualTeX OLE, WPS equation, MathType) in the selection or the whole document a number in the document's number format. Existing numbers are kept and the following ones follow on; references are updated, and one Undo reverts it." imageMso="Numbering">
            <button id="VisualTeX.WordVsto.BatchNumberingSelection" label="Number Display Equations in Selection" onAction="OnNumberDisplayFormulasSelection" />
            <button id="VisualTeX.WordVsto.BatchNumberingDocument" label="Number All Display Equations" onAction="OnNumberDisplayFormulasDocument" />
          </menu>
          <button id="VisualTeX.WordVsto.UpdateEquationFormat" label="Update &#10;Equation Format" size="large" enableLineBreak="true" screentip="Update equation format" supertip="Bring every equation number structure an earlier VisualTeX wrote in this document into the current format and give WPS equations the font chosen under WPS Equation Font. Equation content and kinds do not change." imageMso="Refresh" onAction="OnUpdateEquationFormat" />
          <button id="VisualTeX.WordVsto.BulkImport" label="Bulk &#10;Import" size="large" enableLineBreak="true" screentip="Import LaTeX / Markdown" supertip="Choose a LaTeX or Markdown document: its text becomes text and its equations become editable inline and display equations." imageMso="OpenFile" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.FormatConversionGroup" label="Format Conversion">
          <menu id="VisualTeX.WordVsto.ConvertFormulas" label="Convert &#10;Equations" size="large" enableLineBreak="true" screentip="Convert equation format" supertip="Turn every equation of another format (VisualTeX OLE, WPS equation, MathType) in the range into the chosen format; equations already in it stay as they are. The same equation gives the same result whatever its source; sizes, numbers and references follow this document's settings, and it can be undone." imageMso="EquationTransferOMML">
            <button id="VisualTeX.WordVsto.ConvertSelectionToOmml" label="Selected to WPS Equations" tag="wordOmml|selection" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToMathType" label="Selected to MathType" tag="mathTypeOle|selection" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToVisualTeX" label="Selected to VisualTeX OLE" tag="nativeOle|selection" onAction="OnConvertFormulasTo" />
            <menuSeparator id="VisualTeX.WordVsto.ConvertDocumentHeader" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToOmml" label="All to WPS Equations" tag="wordOmml|document" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToMathType" label="All to MathType" tag="mathTypeOle|document" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToVisualTeX" label="All to VisualTeX OLE" tag="nativeOle|document" onAction="OnConvertFormulasTo" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.RedrawGroup" label="LaTeX Redraw">
          <menu id="VisualTeX.WordVsto.RedrawSelection" label="Redraw &#10;Selection" size="large" enableLineBreak="true" screentip="Redraw selected equations or LaTeX" supertip="Redraw selected LaTeX in place as WPS equations, VisualTeX OLE or MathType, or restore any of those three equation types to LaTeX source." imageMso="EquationInsertLaTex">
            <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="LaTeX → WPS Equation" screentip="Replace in place with a WPS equation" onAction="OnRedrawSelectionToOmml" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="LaTeX → VisualTeX OLE" screentip="Replace in place with a double-click editable VisualTeX OLE" onAction="OnRedrawSelectionToOle" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="LaTeX → MathType" screentip="Replace in place with a MathType equation" onAction="OnRedrawSelectionToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawSelectionSeparator" />
            <button id="VisualTeX.WordVsto.FormulasToLatexSelection" label="Selected Equations to LaTeX Code" screentip="Turn the selected equations back into LaTeX code" supertip="Turn the equations in the selection (VisualTeX OLE, WPS equation or MathType) back into LaTeX code in place." tag="selection" onAction="OnFormulasToLatex" />
          </menu>
          <menu id="VisualTeX.WordVsto.RedrawDocument" label="Redraw Entire &#10;Document" size="large" enableLineBreak="true" screentip="Redraw equations or LaTeX in the entire document" supertip="Redraw all LaTeX in place as WPS equations, VisualTeX OLE or MathType, or restore all three equation types to LaTeX source. A confirmation is shown before changes begin." imageMso="EquationTransferLaTex">
            <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="All LaTeX → WPS Equations" onAction="OnRedrawDocumentToOmml" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="All LaTeX → VisualTeX OLE" onAction="OnRedrawDocumentToOle" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="All LaTeX → MathType" onAction="OnRedrawDocumentToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawDocumentSeparator" />
            <button id="VisualTeX.WordVsto.FormulasToLatexDocument" label="All Equations to LaTeX Code" tag="document" onAction="OnFormulasToLatex" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="Formula Font Size">
          <box id="VisualTeX.WordVsto.SizeColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.SelectedSizeBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.FontSizeDecrease" label="Decrease" imageMso="DecreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
              <comboBox id="VisualTeX.WordVsto.FontSize" label="Font Size" sizeString="Follow Text" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
                __FONT_SIZE_ITEMS__
              </comboBox>
              <button id="VisualTeX.WordVsto.FontSizeIncrease" label="Increase" imageMso="IncreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
            </box>
            <box id="VisualTeX.WordVsto.KindSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.WordVsto.InlineFontSize" label="Inline" sizeString="Follow Text" screentip="Inline formula size" supertip="Once chosen, every inline formula of the current document takes this size at once (undoable), and LaTeX redraw, bulk import and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __INLINE_FONT_SIZE_ITEMS__
              </comboBox>
              <comboBox id="VisualTeX.WordVsto.DisplayFontSize" label="Display" sizeString="Follow Text" screentip="Display formula size" supertip="Once chosen, every display formula (numbered ones included) of the current document takes this size at once (undoable), and LaTeX redraw, bulk import and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __DISPLAY_FONT_SIZE_ITEMS__
              </comboBox>
              <menu id="VisualTeX.WordVsto.OmmlMathFont" label="WPS Equation Font" imageMso="batch_set_font_kd" tag="menu" screentip="WPS equation font" supertip="The math font of the WPS equations (OMML) VisualTeX writes. WPS has no document math font: the chosen font is used for the WPS equations VisualTeX inserts, edits or converts from now on; an existing equation takes it when it is edited again." getEnabled="GetOmmlMathFontEnabled">
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontFollow" label="Follow Document (No Change)" tag="follow" screentip="Follow document" supertip="VisualTeX names no font; WPS equations use Cambria Math (default)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontCambria" label="Cambria Math" tag="cambria" screentip="Cambria Math" supertip="Office's default math font." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontStix" getLabel="GetOmmlMathFontLabel" tag="stix" screentip="STIX Two Math (bundled)" supertip="The STIX Two Math that ships with VisualTeX (the official TrueType build, Times-style, embedded by Save as PDF)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontTermes" getLabel="GetOmmlMathFontLabel" tag="termes" screentip="TeX Gyre Termes Math (bundled)" supertip="The TeX Gyre Termes Math that ships with VisualTeX (Times-style, close to LaTeX's newtxmath; TrueType build, embedded by Save as PDF)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
                <toggleButton id="VisualTeX.WordVsto.OmmlMathFontLatinModern" getLabel="GetOmmlMathFontLabel" tag="latin-modern" screentip="Latin Modern Math (bundled)" supertip="The Latin Modern Math that ships with VisualTeX (TrueType build, embedded by Save as PDF)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
              </menu>
            </box>
          </box>
        </group>
        <group id="VisualTeX.WordVsto.CompatibilityRepairGroup" label="Compatibility Repair">
          <button id="VisualTeX.WordVsto.RepairInlineBaselines" label="Repair Formula &#10;Baselines" size="large" enableLineBreak="true" screentip="Scan and repair inline equation offsets" supertip="Scan the inline MathType / VisualTeX equations of the document; VisualTeX equations are calibrated from their own saved render baseline, MathType equations against a stable document baseline. After you confirm, only their vertical offsets are repaired, not content, size, numbers or references." imageMso="VAlignCenter" onAction="OnRepairInlineFormulaBaselines" />
          <button id="VisualTeX.WordVsto.RepairMathTypeDisplays" label="Repair Formula &#10;Display" size="large" enableLineBreak="true" screentip="Redraw MathType equation pictures with MathType's own layout" supertip="Scan the MathType equations of the document and redraw their pictures from their data, fixing wrong symbols drawn when MathType fonts were missing. Only the pictures are replaced and realigned on MathType's baseline; equation data, numbers and references stay as they are." imageMso="exporttopicture32" onAction="OnRepairMathTypeDisplays" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";
}

public sealed partial class ThisAddIn
{
    /// <summary>
    /// WPS Writer's tab (Chinese or English): the three formula-size boxes take their items from Word's tab, so
    /// both list the same sizes.
    /// </summary>
    internal static string BuildWpsRibbonXml(bool english)
    {
        var wps = english ? RibbonXmlWpsEnglish : RibbonXmlWps;
        var word = System.Xml.Linq.XDocument.Parse(english ? RibbonXmlEnglish : RibbonXml);
        string Items(string comboBoxId)
        {
            var box = word.Descendants().First(element =>
                element.Name.LocalName == "comboBox" && (string?)element.Attribute("id") == comboBoxId);
            return string.Concat(box.Elements().Select(item =>
                item.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"", string.Empty)));
        }
        return wps
            .Replace("__FONT_SIZE_ITEMS__", Items("VisualTeX.WordVsto.FontSize"))
            .Replace("__INLINE_FONT_SIZE_ITEMS__", Items("VisualTeX.WordVsto.InlineFontSize"))
            .Replace("__DISPLAY_FONT_SIZE_ITEMS__", Items("VisualTeX.WordVsto.DisplayFontSize"));
    }
}

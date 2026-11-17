namespace VisualTeX.WordVsto;

// The VisualTeX tab as WPS Writer shows it, in WPS's own style. WPS shows no group titles and no menu section
// titles, so every name says what it does: a verb, what it acts on, and into what ("转为 WPS 公式", "LaTeX 代码转为
// 公式"). The formats are named the same everywhere: VisualTeX 公式, WPS 公式 (WPS's own equations, OMML), MathType
// 公式, LaTeX 代码. Groups, left to right: insert and edit, numbering and references, formula format conversion,
// LaTeX code, sizes and tools. Everyday actions are direct buttons; the batch conversions, which also choose a scope,
// are menus with the scope spelled out in each item.
// How WPS lays the columns out (12.1, Windows and Mac): it packs consecutive vertical boxes into one column of up to
// four (the tab then grows taller than WPS's own), so a group's columns stand in a horizontal box; it centres a lone
// control between the two rows, so a lone control's column gets an empty label as its second row; a box of only size
// boxes gets tighter rows than buttons, so the size boxes share their horizontal box with a column of buttons. The
// WPS equation font is a drop-down showing the current font, like WPS's own font box (a menu under a size box is
// stretched to its width); the size box above it has a longer sizeString, because WPS draws a size box narrower than
// a drop-down with the same one. Icons are imageMso names WPS ships, chosen by what they show (an add-in picture is
// drawn in a fixed 16 px box without transparency, an Office icon name WPS lacks shows a "?"). Same callbacks as
// Word's tab.
public sealed partial class ThisAddIn
{
    private const string RibbonXmlWps = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="插入与编辑">
          <button id="VisualTeX.WordVsto.Inline" label="行内公式" size="large" screentip="插入 VisualTeX 行内公式" supertip="在光标处插入 VisualTeX 行内公式；双击即可再次编辑。" imageMso="EquationScriptGallery" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="行间公式" size="large" screentip="插入 VisualTeX 行间公式" supertip="在单独的居中段落插入 VisualTeX 行间公式；双击即可再次编辑。" imageMso="EquationFractionGallery" onAction="OnInsertDisplay" />
          <box id="VisualTeX.WordVsto.OmmlBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.InlineOmml" label="WPS 行内公式" screentip="插入 WPS 行内公式" supertip="在 VisualTeX 编辑器中写好后，在光标处插入 WPS 自己的公式（OMML），之后也能用 WPS 的公式工具编辑。" imageMso="EquationProfessional" onAction="OnInsertInlineOmml" />
            <button id="VisualTeX.WordVsto.DisplayOmml" label="WPS 行间公式" screentip="插入 WPS 行间公式" supertip="在 VisualTeX 编辑器中写好后，插入 WPS 自己的行间公式（OMML），之后也能用 WPS 的公式工具编辑。" imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" />
          </box>
          <button id="VisualTeX.WordVsto.Edit" label="编辑公式" size="large" screentip="编辑所选公式" supertip="在 VisualTeX 编辑器中编辑所选的 VisualTeX 公式、WPS 公式或 MathType 公式；也可以双击公式。" imageMso="KSEquationEditor" onAction="OnEditSelected" />
        </group>
        <group id="VisualTeX.WordVsto.NumberingGroup" label="编号与引用">
          <box id="VisualTeX.WordVsto.NumberingColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.NumbersBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.UpdateNumbers" label="更新编号" screentip="更新公式编号" supertip="重新计算当前文档中 VisualTeX 公式、WPS 公式和 MathType 公式的编号，并更新对它们的引用。" imageMso="UpdateField" onAction="OnUpdateEquationNumbers" />
              <menu id="VisualTeX.WordVsto.NumberFormat" label="编号格式" screentip="公式编号格式" supertip="选择后立即重排当前文档已有的公式编号，并用于之后插入的带编号公式。" imageMso="MoreNumbering">
                <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="全文连续编号（1）" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="按章编号（1.1）" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="按章编号（1-1）" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="按节编号（1.1.1）" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="按节编号（1.1-1）" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              </menu>
            </box>
            <box id="VisualTeX.WordVsto.ReferenceBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.InsertReference" label="插入公式引用" screentip="插入公式引用" supertip="从当前文档带编号的公式中选择一个，在光标处插入随编号自动更新的引用。" imageMso="CrossReference" onAction="OnInsertEquationReference" />
              <menu id="VisualTeX.WordVsto.BatchNumbering" label="添加编号" screentip="给行间公式添加编号" supertip="按当前编号格式，给所选范围或全文中没有编号的行间公式添加编号；已有编号保持不变并自动顺延，引用随之更新，可一步撤销。" imageMso="Numbering">
                <button id="VisualTeX.WordVsto.BatchNumberingSelection" label="给所选行间公式编号" onAction="OnNumberDisplayFormulasSelection" />
                <button id="VisualTeX.WordVsto.BatchNumberingDocument" label="给全文行间公式编号" onAction="OnNumberDisplayFormulasDocument" />
              </menu>
            </box>
          </box>
        </group>
        <group id="VisualTeX.WordVsto.FormatConversionGroup" label="格式转换">
          <box id="VisualTeX.WordVsto.ConvertColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.ToOmmlVisualTeXBox" boxStyle="vertical">
              <menu id="VisualTeX.WordVsto.ToOmmlMenu" label="转为 WPS 公式" screentip="公式转为 WPS 公式" supertip="把其他格式的公式（VisualTeX、MathType）转为 WPS 自己的公式（OMML），已是 WPS 公式的不变；同一公式不论原来是什么格式，结果都相同。字号、编号与引用按当前文档设置，可撤销。" imageMso="EquationTransferOMML">
                <button id="VisualTeX.WordVsto.ConvertSelectionToOmml" label="所选公式转为 WPS 公式" tag="wordOmml|selection" onAction="OnConvertFormulasTo" />
                <button id="VisualTeX.WordVsto.ConvertDocumentToOmml" label="全文公式转为 WPS 公式" tag="wordOmml|document" onAction="OnConvertFormulasTo" />
              </menu>
              <menu id="VisualTeX.WordVsto.ToVisualTeXMenu" label="转为 VisualTeX 公式" screentip="公式转为 VisualTeX 公式" supertip="把其他格式的公式（WPS、MathType）转为 VisualTeX 公式，已是 VisualTeX 公式的不变；同一公式不论原来是什么格式，结果都相同。字号、编号与引用按当前文档设置，可撤销。" imageMso="exporttopicture32">
                <button id="VisualTeX.WordVsto.ConvertSelectionToVisualTeX" label="所选公式转为 VisualTeX 公式" tag="nativeOle|selection" onAction="OnConvertFormulasTo" />
                <button id="VisualTeX.WordVsto.ConvertDocumentToVisualTeX" label="全文公式转为 VisualTeX 公式" tag="nativeOle|document" onAction="OnConvertFormulasTo" />
              </menu>
            </box>
            <box id="VisualTeX.WordVsto.ToMathTypeBox" boxStyle="vertical">
              <menu id="VisualTeX.WordVsto.ToMathTypeMenu" label="转为 MathType 公式" screentip="公式转为 MathType 公式" supertip="把其他格式的公式（VisualTeX、WPS）转为 MathType 公式，已是 MathType 公式的不变；同一公式不论原来是什么格式，结果都相同。字号与编号按当前文档设置，可撤销。" imageMso="Convert">
                <button id="VisualTeX.WordVsto.ConvertSelectionToMathType" label="所选公式转为 MathType 公式" tag="mathTypeOle|selection" onAction="OnConvertFormulasTo" />
                <button id="VisualTeX.WordVsto.ConvertDocumentToMathType" label="全文公式转为 MathType 公式" tag="mathTypeOle|document" onAction="OnConvertFormulasTo" />
              </menu>
              <labelControl id="VisualTeX.WordVsto.ToMathTypeRow2" label="　" />
            </box>
          </box>
        </group>
        <group id="VisualTeX.WordVsto.LatexGroup" label="LaTeX 代码">
          <box id="VisualTeX.WordVsto.LatexBox" boxStyle="vertical">
            <menu id="VisualTeX.WordVsto.Redraw" label="LaTeX 代码转为公式" screentip="LaTeX 代码转为公式" supertip="把文档里写成文字的 LaTeX 公式（$...$、$$...$$、\(...\)、\[...\] 和公式环境）原位换成公式；全文操作开始前会再次确认。" imageMso="EquationInsertLaTex">
              <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="所选 LaTeX 代码转为 VisualTeX 公式" onAction="OnRedrawSelectionToOle" />
              <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="所选 LaTeX 代码转为 WPS 公式" onAction="OnRedrawSelectionToOmml" />
              <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="所选 LaTeX 代码转为 MathType 公式" onAction="OnRedrawSelectionToMathType" />
              <menuSeparator id="VisualTeX.WordVsto.RedrawSeparator" />
              <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="全文 LaTeX 代码转为 VisualTeX 公式" onAction="OnRedrawDocumentToOle" />
              <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="全文 LaTeX 代码转为 WPS 公式" onAction="OnRedrawDocumentToOmml" />
              <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="全文 LaTeX 代码转为 MathType 公式" onAction="OnRedrawDocumentToMathType" />
            </menu>
            <menu id="VisualTeX.WordVsto.ToLatex" label="公式转为 LaTeX 代码" screentip="公式转为 LaTeX 代码" supertip="把公式（VisualTeX、WPS、MathType 公式均可）原位换回 LaTeX 代码，之后可以再转回公式。" imageMso="EquationTransferLaTex">
              <button id="VisualTeX.WordVsto.FormulasToLatexSelection" label="所选公式转为 LaTeX 代码" tag="selection" onAction="OnFormulasToLatex" />
              <button id="VisualTeX.WordVsto.FormulasToLatexDocument" label="全文公式转为 LaTeX 代码" tag="document" onAction="OnFormulasToLatex" />
            </menu>
          </box>
          <button id="VisualTeX.WordVsto.BulkImport" label="导入文档" size="large" screentip="导入 LaTeX 或 Markdown 文档" supertip="选择一个 LaTeX 或 Markdown 文档，正文插入为文字，公式插入为可编辑、可调字号的行内和行间公式。" imageMso="OpenFile" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="字号与工具">
          <box id="VisualTeX.WordVsto.SizeColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.KindSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.WordVsto.InlineFontSize" label="行内公式字号" sizeString="初号（42 磅）" screentip="行内公式字号" supertip="选择字号后，当前文档中所有行内公式立即改为这个字号，可撤销；之后转换、LaTeX 代码转为公式和导入文档生成的行内公式也用这个字号。“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __INLINE_FONT_SIZE_ITEMS__
              </comboBox>
              <comboBox id="VisualTeX.WordVsto.DisplayFontSize" label="行间公式字号" sizeString="初号（42 磅）" screentip="行间公式字号" supertip="选择字号后，当前文档中所有行间公式（含编号公式）立即改为这个字号，可撤销；之后转换、LaTeX 代码转为公式和导入文档生成的行间公式也用这个字号。“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __DISPLAY_FONT_SIZE_ITEMS__
              </comboBox>
            </box>
            <box id="VisualTeX.WordVsto.SelectedSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.WordVsto.FontSize" label="所选公式字号" sizeString="TeX Gyre Termes Math" screentip="所选公式字号" supertip="先选中一个或多个 VisualTeX 公式，再输入或选择字号。" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
                __FONT_SIZE_ITEMS__
              </comboBox>
              <dropDown id="VisualTeX.WordVsto.OmmlMathFont" label="WPS 公式字体" sizeString="Latin Modern Math" screentip="WPS 公式字体" supertip="VisualTeX 插入或编辑的 WPS 公式（OMML）使用的数学字体。WPS 没有文档数学字体：所选字体用于之后写入的公式。" getItemCount="GetOmmlMathFontItemCount" getItemLabel="GetOmmlMathFontItemLabel" getSelectedItemIndex="GetOmmlMathFontSelectedIndex" onAction="OnOmmlMathFontSelected" />
            </box>
            <box id="VisualTeX.WordVsto.MaintenanceBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.UpdateEquationFormat" label="升级旧版公式" screentip="升级旧版公式" supertip="把本文档中旧版 VisualTeX 写入的公式编号结构一次升级为当前格式，并让 WPS 公式使用所选数学字体。公式内容与类型不变。" imageMso="Refresh" onAction="OnUpdateEquationFormat" />
              <menu id="VisualTeX.WordVsto.Repair" label="修复公式显示" imageMso="VAlignCenter" screentip="修复公式偏移与显示">
                <button id="VisualTeX.WordVsto.RepairInlineBaselines" label="修复行内公式偏移" screentip="扫描并修复行内公式垂直偏移" supertip="扫描全文中的 MathType / VisualTeX 行内公式；确认后仅修复垂直偏移，不修改公式内容、字号、编号或引用。" onAction="OnRepairInlineFormulaBaselines" />
                <button id="VisualTeX.WordVsto.RepairMathTypeDisplays" label="修复 MathType 公式显示" screentip="用 MathType 自身排版重新生成公式图像" supertip="扫描全文 MathType 公式，按公式原有数据重新生成显示图像，修复缺少 MathType 字体时出现的错误符号；公式数据、编号和引用保持不变。" onAction="OnRepairMathTypeDisplays" />
              </menu>
            </box>
          </box>
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
        <group id="VisualTeX.WordVsto.Group" label="Insert and Edit">
          <button id="VisualTeX.WordVsto.Inline" label="Inline" size="large" screentip="Insert an inline VisualTeX equation" supertip="Insert a VisualTeX inline equation at the cursor; double-click it to edit it again." imageMso="EquationScriptGallery" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="Display" size="large" screentip="Insert a display VisualTeX equation" supertip="Insert a VisualTeX display equation in its own centered paragraph; double-click it to edit it again." imageMso="EquationFractionGallery" onAction="OnInsertDisplay" />
          <box id="VisualTeX.WordVsto.OmmlBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.InlineOmml" label="WPS Inline Equation" screentip="Insert an inline WPS equation" supertip="Write it in the VisualTeX editor, then insert it at the cursor as WPS's own equation (OMML), which WPS's equation tools can edit too." imageMso="EquationProfessional" onAction="OnInsertInlineOmml" />
            <button id="VisualTeX.WordVsto.DisplayOmml" label="WPS Display Equation" screentip="Insert a display WPS equation" supertip="Write it in the VisualTeX editor, then insert it as WPS's own display equation (OMML), which WPS's equation tools can edit too." imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" />
          </box>
          <button id="VisualTeX.WordVsto.Edit" label="Edit" size="large" screentip="Edit the selected equation" supertip="Edit the selected VisualTeX, WPS or MathType equation in the VisualTeX editor; double-clicking it does the same." imageMso="KSEquationEditor" onAction="OnEditSelected" />
        </group>
        <group id="VisualTeX.WordVsto.NumberingGroup" label="Numbering and References">
          <box id="VisualTeX.WordVsto.NumberingColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.NumbersBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.UpdateNumbers" label="Update Numbers" screentip="Update equation numbers" supertip="Renumber the VisualTeX, WPS and MathType equations of the current document and update the references to them." imageMso="UpdateField" onAction="OnUpdateEquationNumbers" />
              <menu id="VisualTeX.WordVsto.NumberFormat" label="Number Format" screentip="Equation number format" supertip="Renumber the existing equations of this document in the chosen format right away and use it for numbered equations inserted later." imageMso="MoreNumbering">
                <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="Continuous (1)" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="By chapter (1.1)" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="By chapter (1-1)" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="By section (1.1.1)" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
                <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="By section (1.1-1)" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              </menu>
            </box>
            <box id="VisualTeX.WordVsto.ReferenceBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.InsertReference" label="Insert Reference" screentip="Insert an equation reference" supertip="Choose a numbered equation of this document and insert, at the cursor, a reference that follows its number." imageMso="CrossReference" onAction="OnInsertEquationReference" />
              <menu id="VisualTeX.WordVsto.BatchNumbering" label="Add Numbers" screentip="Number display equations" supertip="Give the unnumbered display equations in the selection or the document a number in the current format. Existing numbers are kept and the following ones follow on; references are updated, and one Undo reverts it." imageMso="Numbering">
                <button id="VisualTeX.WordVsto.BatchNumberingSelection" label="Number Display Equations in Selection" onAction="OnNumberDisplayFormulasSelection" />
                <button id="VisualTeX.WordVsto.BatchNumberingDocument" label="Number All Display Equations" onAction="OnNumberDisplayFormulasDocument" />
              </menu>
            </box>
          </box>
        </group>
        <group id="VisualTeX.WordVsto.FormatConversionGroup" label="Convert Formats">
          <box id="VisualTeX.WordVsto.ConvertColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.ToOmmlVisualTeXBox" boxStyle="vertical">
              <menu id="VisualTeX.WordVsto.ToOmmlMenu" label="To WPS Equations" screentip="Equations to WPS equations" supertip="Turn equations of other formats (VisualTeX, MathType) into WPS's own equations (OMML); WPS equations stay as they are. The same equation gives the same result whatever its format. Sizes, numbers and references follow this document's settings; it can be undone." imageMso="EquationTransferOMML">
                <button id="VisualTeX.WordVsto.ConvertSelectionToOmml" label="Selected Equations to WPS Equations" tag="wordOmml|selection" onAction="OnConvertFormulasTo" />
                <button id="VisualTeX.WordVsto.ConvertDocumentToOmml" label="All Equations to WPS Equations" tag="wordOmml|document" onAction="OnConvertFormulasTo" />
              </menu>
              <menu id="VisualTeX.WordVsto.ToVisualTeXMenu" label="To VisualTeX Equations" screentip="Equations to VisualTeX equations" supertip="Turn equations of other formats (WPS, MathType) into VisualTeX equations; VisualTeX equations stay as they are. The same equation gives the same result whatever its format. Sizes, numbers and references follow this document's settings; it can be undone." imageMso="exporttopicture32">
                <button id="VisualTeX.WordVsto.ConvertSelectionToVisualTeX" label="Selected Equations to VisualTeX Equations" tag="nativeOle|selection" onAction="OnConvertFormulasTo" />
                <button id="VisualTeX.WordVsto.ConvertDocumentToVisualTeX" label="All Equations to VisualTeX Equations" tag="nativeOle|document" onAction="OnConvertFormulasTo" />
              </menu>
            </box>
            <box id="VisualTeX.WordVsto.ToMathTypeBox" boxStyle="vertical">
              <menu id="VisualTeX.WordVsto.ToMathTypeMenu" label="To MathType Equations" screentip="Equations to MathType equations" supertip="Turn equations of other formats (VisualTeX, WPS) into MathType equations; MathType equations stay as they are. The same equation gives the same result whatever its format. Sizes and numbers follow this document's settings; it can be undone." imageMso="Convert">
                <button id="VisualTeX.WordVsto.ConvertSelectionToMathType" label="Selected Equations to MathType Equations" tag="mathTypeOle|selection" onAction="OnConvertFormulasTo" />
                <button id="VisualTeX.WordVsto.ConvertDocumentToMathType" label="All Equations to MathType Equations" tag="mathTypeOle|document" onAction="OnConvertFormulasTo" />
              </menu>
              <labelControl id="VisualTeX.WordVsto.ToMathTypeRow2" label="　" />
            </box>
          </box>
        </group>
        <group id="VisualTeX.WordVsto.LatexGroup" label="LaTeX Code">
          <box id="VisualTeX.WordVsto.LatexBox" boxStyle="vertical">
            <menu id="VisualTeX.WordVsto.Redraw" label="LaTeX Code to Equations" screentip="LaTeX code to equations" supertip="Turn LaTeX written as text in the document ($...$, $$...$$, \(...\), \[...\] and environments) into equations in place; whole-document operations ask for confirmation first." imageMso="EquationInsertLaTex">
              <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="Selected LaTeX Code to VisualTeX Equations" onAction="OnRedrawSelectionToOle" />
              <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="Selected LaTeX Code to WPS Equations" onAction="OnRedrawSelectionToOmml" />
              <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="Selected LaTeX Code to MathType Equations" onAction="OnRedrawSelectionToMathType" />
              <menuSeparator id="VisualTeX.WordVsto.RedrawSeparator" />
              <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="All LaTeX Code to VisualTeX Equations" onAction="OnRedrawDocumentToOle" />
              <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="All LaTeX Code to WPS Equations" onAction="OnRedrawDocumentToOmml" />
              <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="All LaTeX Code to MathType Equations" onAction="OnRedrawDocumentToMathType" />
            </menu>
            <menu id="VisualTeX.WordVsto.ToLatex" label="Equations to LaTeX Code" screentip="Equations to LaTeX code" supertip="Turn equations (VisualTeX, WPS or MathType) back into LaTeX code in place; they can be turned into equations again." imageMso="EquationTransferLaTex">
              <button id="VisualTeX.WordVsto.FormulasToLatexSelection" label="Selected Equations to LaTeX Code" tag="selection" onAction="OnFormulasToLatex" />
              <button id="VisualTeX.WordVsto.FormulasToLatexDocument" label="All Equations to LaTeX Code" tag="document" onAction="OnFormulasToLatex" />
            </menu>
          </box>
          <button id="VisualTeX.WordVsto.BulkImport" label="Import Document" size="large" screentip="Import a LaTeX or Markdown document" supertip="Choose a LaTeX or Markdown document: its text becomes text and its equations become editable inline and display equations." imageMso="OpenFile" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="Size and Tools">
          <box id="VisualTeX.WordVsto.SizeColumns" boxStyle="horizontal">
            <box id="VisualTeX.WordVsto.KindSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.WordVsto.InlineFontSize" label="Inline Size" sizeString="Follow text (512 pt)" screentip="Inline equation size" supertip="Choosing a size changes every inline equation in this document to it at once (undoable); inline equations made later by conversion, LaTeX code to equations and import use it too. Follow text leaves existing equations alone." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __INLINE_FONT_SIZE_ITEMS__
              </comboBox>
              <comboBox id="VisualTeX.WordVsto.DisplayFontSize" label="Display Size" sizeString="Follow text (512 pt)" screentip="Display equation size" supertip="Choosing a size changes every display equation (numbered ones too) in this document to it at once (undoable); display equations made later by conversion, LaTeX code to equations and import use it too. Follow text leaves existing equations alone." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __DISPLAY_FONT_SIZE_ITEMS__
              </comboBox>
            </box>
            <box id="VisualTeX.WordVsto.SelectedSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.WordVsto.FontSize" label="Selected Size" sizeString="TeX Gyre Termes Math" screentip="Size of the selected equations" supertip="Select one or more VisualTeX equations, then type or choose a size." getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
                __FONT_SIZE_ITEMS__
              </comboBox>
              <dropDown id="VisualTeX.WordVsto.OmmlMathFont" label="WPS Equation Font" sizeString="Latin Modern Math" screentip="WPS equation font" supertip="The math font of the WPS equations (OMML) VisualTeX inserts or edits. WPS has no document math font: the chosen font is used for equations written from now on." getItemCount="GetOmmlMathFontItemCount" getItemLabel="GetOmmlMathFontItemLabel" getSelectedItemIndex="GetOmmlMathFontSelectedIndex" onAction="OnOmmlMathFontSelected" />
            </box>
            <box id="VisualTeX.WordVsto.MaintenanceBox" boxStyle="vertical">
              <button id="VisualTeX.WordVsto.UpdateEquationFormat" label="Upgrade Old Equations" screentip="Upgrade old equations" supertip="Bring every equation number structure an earlier VisualTeX wrote in this document into the current format and give WPS equations the chosen math font. Equation content and kinds do not change." imageMso="Refresh" onAction="OnUpdateEquationFormat" />
              <menu id="VisualTeX.WordVsto.Repair" label="Repair Equation Display" imageMso="VAlignCenter" screentip="Repair equation offsets and display">
                <button id="VisualTeX.WordVsto.RepairInlineBaselines" label="Repair Inline Equation Offsets" screentip="Find and repair vertical offsets of inline equations" supertip="Scan the inline MathType / VisualTeX equations of the document; after you confirm, only their vertical offsets are repaired, not content, size, numbers or references." onAction="OnRepairInlineFormulaBaselines" />
                <button id="VisualTeX.WordVsto.RepairMathTypeDisplays" label="Repair MathType Equation Display" screentip="Redraw MathType equation pictures with MathType's own typesetting" supertip="Scan the MathType equations of the document and redraw their pictures from their data, fixing wrong symbols drawn when MathType fonts were missing; equation data, numbers and references stay as they are." onAction="OnRepairMathTypeDisplays" />
              </menu>
            </box>
          </box>
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

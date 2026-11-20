namespace VisualTeX.PowerPointVsto;

// The VisualTeX tab as WPS Presentation shows it: PowerPoint's tab (RibbonXml) control for control - the same groups,
// names, sizes and order, and PowerPoint's three-row columns as three-row columns - with WPS's own equations called
// "WPS 公式" where PowerPoint says OMML. As in WPS Writer's tab (VisualTeX.WordVsto/WordRibbonWps.cs): WPS shows no
// menu section titles, so an item PowerPoint names under a section title carries that title in its own name
// ("所选重绘为 WPS 公式"); each column is a vertical box of its own, side-by-side columns stand in a horizontal box,
// and a column shorter than its neighbour gets empty labels as its last rows; icons are imageMso names WPS ships;
// every large label breaks once onto a second line (enableLineBreak).
// Same callbacks as PowerPoint's tab; the formula-size boxes take their items from PowerPoint's tab.
public sealed partial class ThisAddIn
{
    private const string RibbonXmlWps = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="VisualTeX">
          <button id="VisualTeX.PowerPointVsto.New" label="新建 &#10;SVG" size="large" enableLineBreak="true" screentip="新建 VisualTeX SVG 公式" supertip="在当前幻灯片中央插入一个 VisualTeX 公式（SVG 图片）；编辑完成后可以像图片一样移动和缩放，双击即可再次编辑。" imageMso="EquationFractionGallery" onAction="OnNewSvgFormula" />
          <button id="VisualTeX.PowerPointVsto.NewOle" label="新建 &#10;OLE" size="large" enableLineBreak="true" screentip="新建 VisualTeX OLE 公式" supertip="插入一个 VisualTeX 公式（OLE 对象，嵌入文件，双击即可再次编辑）。" imageMso="EquationScriptGallery" onAction="OnNewOleFormula" />
          <button id="VisualTeX.PowerPointVsto.NewMathType" label="新建 &#10;MathType" size="large" enableLineBreak="true" screentip="新建 MathType 公式" supertip="在 VisualTeX 编辑器中写好后，插入一个 MathType 公式。" imageMso="EquationOptions" onAction="OnNewMathTypeFormula" />
          <button id="VisualTeX.PowerPointVsto.InsertInlineOmml" label="插入行内 &#10;WPS 公式" size="large" enableLineBreak="true" screentip="插入行内 WPS 公式" supertip="在 VisualTeX 编辑器中写好后，在文本框的光标或文本选区处插入 WPS 自己的行内公式（OMML），前后文字保留。" imageMso="EquationProfessional" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.PowerPointVsto.InsertDisplayOmml" label="插入行间 &#10;WPS 公式" size="large" enableLineBreak="true" screentip="插入行间 WPS 公式" supertip="正在编辑文本框时，在光标或文本选区处插入单独成段的 WPS 公式（OMML）；否则在当前幻灯片新建一个只含公式的文本框。" imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="编辑&#10;所选公式" size="large" enableLineBreak="true" screentip="编辑所选公式" supertip="在 VisualTeX 编辑器中编辑所选的 VisualTeX、WPS 或 MathType 公式；也可以双击公式。" imageMso="KSEquationEditor" onAction="OnEditSelected" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="公式格式转换">
          <menu id="VisualTeX.PowerPointVsto.UnifyFormat" label="格式&#10;转换" size="large" enableLineBreak="true" screentip="公式格式转换" supertip="把范围内其他格式的公式（VisualTeX、WPS 公式、MathType）全部转为所选格式，已是该格式的公式不变。不同来源的同一公式转换结果相同；若在左侧缩略图中选中了幻灯片，“所选”转换这些幻灯片上的公式。" imageMso="EquationTransferOMML">
            <button id="VisualTeX.PowerPointVsto.UnifySelectionOMML" label="所选公式转为 WPS 公式" tag="any|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySelectionMathType" label="所选公式转为 MathType" tag="any|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySelectionVisualTeX" label="所选公式转为 VisualTeX" tag="any|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifySlideHeader" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="当前幻灯片转为 WPS 公式" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="当前幻灯片转为 MathType" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="当前幻灯片转为 VisualTeX" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifyAllHeader" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="整个演示文稿转为 WPS 公式" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="整个演示文稿转为 MathType" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="整个演示文稿转为 VisualTeX" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
          <menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" size="large" enableLineBreak="true" screentip="转为 VisualTeX 时使用的形式" supertip="选择“转为 VisualTeX”和“重绘为 VisualTeX”生成的对象：OLE 对象可双击编辑并嵌入文件；SVG 图片在任何设备上都能清晰显示，同样可双击编辑。" imageMso="InsertPictureFromFile">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE 对象（双击编辑，嵌入文件）" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG 图片（任何设备都能显示）" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu>
          <button id="VisualTeX.PowerPointVsto.UpdateEquationFormat" label="更新&#10;公式格式" size="large" enableLineBreak="true" screentip="更新公式格式" supertip="用当前设置的字体（数学字体、中文字体以及字母、希腊字母和数字的字体）重绘字体与之不同的 VisualTeX 公式（SVG 图片和 OLE 对象）；每个公式保持位置、显示大小、旋转和层级，公式内容不变。" imageMso="Refresh" onAction="OnUpdateEquationFormat" />
        </group>
        <group id="VisualTeX.PowerPointVsto.LatexRedrawGroup" label="LaTeX 重绘">
          <menu id="VisualTeX.PowerPointVsto.LatexRedraw" label="LaTeX &#10;重绘" size="large" enableLineBreak="true" screentip="LaTeX 代码重绘为公式" supertip="把 LaTeX 公式（$…$、$$…$$、\(…\)、\[…\] 和 equation 等环境）重绘为公式。重绘为 WPS 公式在文本框中原位进行；重绘为 VisualTeX 或 MathType 时，只含一个 LaTeX 公式的文本框换成同一位置的公式。" imageMso="EquationInsertLaTex">
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlSelection" label="所选重绘为 WPS 公式" tag="latex|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlSlide" label="本页重绘为 WPS 公式" tag="latex|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlAll" label="全部重绘为 WPS 公式" tag="latex|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.LatexToVisualTeXHeader" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSelection" label="所选重绘为 VisualTeX" tag="latex|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSlide" label="本页重绘为 VisualTeX" tag="latex|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXAll" label="全部重绘为 VisualTeX" tag="latex|visualtex|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.LatexToMathTypeHeader" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeSelection" label="所选重绘为 MathType" tag="latex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeSlide" label="本页重绘为 MathType" tag="latex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeAll" label="全部重绘为 MathType" tag="latex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
          <menu id="VisualTeX.PowerPointVsto.FormulaToLatex" label="转为 &#10;LaTeX 代码" size="large" enableLineBreak="true" screentip="把公式转为 LaTeX 代码" supertip="把 VisualTeX、MathType 和 WPS 公式恢复为 LaTeX 代码：文字中的 WPS 公式原位替换为 $…$（单独成段的为 $$…$$），其余文字不变；公式对象替换为中心相同、字号相同的文本框，内容为 $$…$$。之后可以用“LaTeX 重绘”重新绘制。" imageMso="EquationTransferLaTex">
            <button id="VisualTeX.PowerPointVsto.AnyToLatexSelection" label="所选公式转为 LaTeX 代码" screentip="公式 → LaTeX 代码（所选）" supertip="把所选对象中的 VisualTeX、MathType 和 WPS 公式转为 LaTeX 代码；若只选中了文本框中的部分文字，则只转换其中的公式；若在左侧缩略图中选中了幻灯片，则转换这些幻灯片上的全部公式。" tag="any|latex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexSlide" label="本页公式全部转为 LaTeX 代码" screentip="公式 → LaTeX 代码（本页）" supertip="把当前幻灯片上的全部 VisualTeX、MathType 和 WPS 公式转为 LaTeX 代码。" tag="any|latex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexAll" label="全部公式转为 LaTeX 代码" screentip="公式 → LaTeX 代码（全部）" supertip="把整个演示文稿中的全部 VisualTeX、MathType 和 WPS 公式转为 LaTeX 代码；开始前会确认。" tag="any|latex|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
        </group>
        <group id="VisualTeX.PowerPointVsto.FontSizeGroup" label="公式字号">
          <box id="VisualTeX.PowerPointVsto.SizeColumns" boxStyle="horizontal">
            <box id="VisualTeX.PowerPointVsto.SelectedSizeBox" boxStyle="vertical">
              <button id="VisualTeX.PowerPointVsto.FontSizeDecrease" label="减小" imageMso="DecreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
              <comboBox id="VisualTeX.PowerPointVsto.FontSize" label="字号" sizeString="初号（42 磅）" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
                __FONT_SIZE_ITEMS__
              </comboBox>
              <button id="VisualTeX.PowerPointVsto.FontSizeIncrease" label="增大" imageMso="IncreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
            </box>
            <box id="VisualTeX.PowerPointVsto.KindSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.PowerPointVsto.InlineFontSize" label="行内" sizeString="初号（42 磅）" screentip="行内公式字号" supertip="选择字号后，当前演示文稿中所有行内公式立即改为这个字号，可撤销；之后 LaTeX 重绘和格式转换生成的行内公式也使用这个字号。选“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __INLINE_FONT_SIZE_ITEMS__
              </comboBox>
              <comboBox id="VisualTeX.PowerPointVsto.DisplayFontSize" label="行间" sizeString="初号（42 磅）" screentip="行间公式字号" supertip="选择字号后，当前演示文稿中所有行间公式立即改为这个字号，可撤销；之后 LaTeX 重绘和格式转换生成的行间公式也使用这个字号。选“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __DISPLAY_FONT_SIZE_ITEMS__
              </comboBox>
              <labelControl id="VisualTeX.PowerPointVsto.KindSizeRow3" label="　" />
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
      <tab id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="VisualTeX">
          <button id="VisualTeX.PowerPointVsto.New" label="New &#10;SVG" size="large" enableLineBreak="true" screentip="New VisualTeX SVG equation" supertip="Insert a VisualTeX equation (SVG picture) in the middle of the current slide; move and resize it like a picture, double-click it to edit it again." imageMso="EquationFractionGallery" onAction="OnNewSvgFormula" />
          <button id="VisualTeX.PowerPointVsto.NewOle" label="New &#10;OLE" size="large" enableLineBreak="true" screentip="New VisualTeX OLE equation" supertip="Insert a VisualTeX equation as an OLE object (embedded, double-click to edit it again)." imageMso="EquationScriptGallery" onAction="OnNewOleFormula" />
          <button id="VisualTeX.PowerPointVsto.NewMathType" label="New &#10;MathType" size="large" enableLineBreak="true" screentip="New MathType equation" supertip="Write it in the VisualTeX editor, then insert it as a MathType equation." imageMso="EquationOptions" onAction="OnNewMathTypeFormula" />
          <button id="VisualTeX.PowerPointVsto.InsertInlineOmml" label="Insert Inline &#10;WPS Equation" size="large" enableLineBreak="true" screentip="Insert an inline WPS equation" supertip="Write it in the VisualTeX editor, then insert it at the text-box caret or text selection as WPS's own inline equation (OMML), keeping the text around it." imageMso="EquationProfessional" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.PowerPointVsto.InsertDisplayOmml" label="Insert Display &#10;WPS Equation" size="large" enableLineBreak="true" screentip="Insert a display WPS equation" supertip="While editing a text box, insert a WPS equation (OMML) paragraph of its own at the caret or text selection; otherwise create a text box holding only the equation on the current slide." imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="Edit Selected &#10;Formula" size="large" enableLineBreak="true" screentip="Edit the selected equation" supertip="Edit the selected VisualTeX, WPS or MathType equation in the VisualTeX editor; double-clicking it does the same." imageMso="KSEquationEditor" onAction="OnEditSelected" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="Equation Format Conversion">
          <menu id="VisualTeX.PowerPointVsto.UnifyFormat" label="Convert &#10;Equations" size="large" enableLineBreak="true" screentip="Convert equation format" supertip="Turn every equation of another format (VisualTeX, WPS equation, MathType) in the range into the chosen format; equations already in it stay as they are. The same equation gives the same result whatever its source. With slides selected in the thumbnail pane, Selected converts the equations on those slides." imageMso="EquationTransferOMML">
            <button id="VisualTeX.PowerPointVsto.UnifySelectionOMML" label="Selected to WPS Equations" tag="any|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySelectionMathType" label="Selected to MathType" tag="any|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySelectionVisualTeX" label="Selected to VisualTeX" tag="any|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifySlideHeader" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="Current Slide to WPS Equations" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="Current Slide to MathType" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="Current Slide to VisualTeX" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifyAllHeader" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="Whole Presentation to WPS Equations" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="Whole Presentation to MathType" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="Whole Presentation to VisualTeX" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
          <menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" size="large" enableLineBreak="true" screentip="Form used when converting to VisualTeX" supertip="Choose what converting or redrawing to VisualTeX creates: an OLE object is embedded and double-click editable; an SVG picture displays sharply on every device and is also double-click editable." imageMso="InsertPictureFromFile">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE object (double-click to edit, embedded)" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG picture (displays on every device)" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu>
          <button id="VisualTeX.PowerPointVsto.UpdateEquationFormat" label="Update &#10;Equation Format" size="large" enableLineBreak="true" screentip="Update equation format" supertip="Draw the VisualTeX equations (SVG pictures and OLE objects) drawn in other fonts again in the current fonts (the math font, the Chinese font and the fonts for letters, Greek letters and digits); each keeps its position, size, rotation and layer, and its content does not change." imageMso="Refresh" onAction="OnUpdateEquationFormat" />
        </group>
        <group id="VisualTeX.PowerPointVsto.LatexRedrawGroup" label="LaTeX Redraw">
          <menu id="VisualTeX.PowerPointVsto.LatexRedraw" label="LaTeX &#10;Redraw" size="large" enableLineBreak="true" screentip="Redraw LaTeX code as equations" supertip="Redraw LaTeX equations ($…$, $$…$$, \(…\), \[…\] and environments such as equation). Redrawing as WPS equations happens in place in the text box; redrawing as VisualTeX or MathType replaces a text box holding a single LaTeX equation with an equation at the same place." imageMso="EquationInsertLaTex">
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlSelection" label="Selection as WPS Equations" tag="latex|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlSlide" label="Current Slide as WPS Equations" tag="latex|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlAll" label="Whole Presentation as WPS Equations" tag="latex|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.LatexToVisualTeXHeader" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSelection" label="Selection as VisualTeX" tag="latex|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSlide" label="Current Slide as VisualTeX" tag="latex|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXAll" label="Whole Presentation as VisualTeX" tag="latex|visualtex|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.LatexToMathTypeHeader" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeSelection" label="Selection as MathType" tag="latex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeSlide" label="Current Slide as MathType" tag="latex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeAll" label="Whole Presentation as MathType" tag="latex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
          <menu id="VisualTeX.PowerPointVsto.FormulaToLatex" label="To LaTeX &#10;Code" size="large" enableLineBreak="true" screentip="Turn equations into LaTeX code" supertip="Turn VisualTeX, MathType and WPS equations back into LaTeX code: a WPS equation in text becomes $…$ where it stands ($$…$$ when it is alone in its paragraph) and the other text stays as it is; an equation object becomes a text box with $$…$$ at the same center and font size. LaTeX Redraw draws them again." imageMso="EquationTransferLaTex">
            <button id="VisualTeX.PowerPointVsto.AnyToLatexSelection" label="Selected Equations → LaTeX Code" screentip="Equations → LaTeX Code (Selected)" supertip="Turn the VisualTeX, MathType and WPS equations among the selected objects into LaTeX code. With part of a text box's text selected, only the equations in it are converted. With slides selected in the thumbnail pane, converts every equation on those slides." tag="any|latex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexSlide" label="This Slide → LaTeX Code" screentip="Equations → LaTeX Code (Slide)" supertip="Turn every VisualTeX, MathType and WPS equation on the current slide into LaTeX code." tag="any|latex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexAll" label="All Slides → LaTeX Code" screentip="Equations → LaTeX Code (All)" supertip="Turn every VisualTeX, MathType and WPS equation in the presentation into LaTeX code. You are asked to confirm first." tag="any|latex|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
        </group>
        <group id="VisualTeX.PowerPointVsto.FontSizeGroup" label="Formula Font Size">
          <box id="VisualTeX.PowerPointVsto.SizeColumns" boxStyle="horizontal">
            <box id="VisualTeX.PowerPointVsto.SelectedSizeBox" boxStyle="vertical">
              <button id="VisualTeX.PowerPointVsto.FontSizeDecrease" label="Decrease" imageMso="DecreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
              <comboBox id="VisualTeX.PowerPointVsto.FontSize" label="Font Size" sizeString="Follow Text" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
                __FONT_SIZE_ITEMS__
              </comboBox>
              <button id="VisualTeX.PowerPointVsto.FontSizeIncrease" label="Increase" imageMso="IncreaseFontSize" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
            </box>
            <box id="VisualTeX.PowerPointVsto.KindSizeBox" boxStyle="vertical">
              <comboBox id="VisualTeX.PowerPointVsto.InlineFontSize" label="Inline" sizeString="Follow Text" screentip="Inline formula size" supertip="Once chosen, every inline formula of the current presentation takes this size at once (undoable), and LaTeX redraw and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __INLINE_FONT_SIZE_ITEMS__
              </comboBox>
              <comboBox id="VisualTeX.PowerPointVsto.DisplayFontSize" label="Display" sizeString="Follow Text" screentip="Display formula size" supertip="Once chosen, every display formula of the current presentation takes this size at once (undoable), and LaTeX redraw and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
                __DISPLAY_FONT_SIZE_ITEMS__
              </comboBox>
              <labelControl id="VisualTeX.PowerPointVsto.KindSizeRow3" label="　" />
            </box>
          </box>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";

    /// <summary>
    /// WPS Presentation's tab (Chinese or English): the three formula-size boxes take their items from PowerPoint's
    /// tab, so both list the same sizes.
    /// </summary>
    internal static string BuildWpsRibbonXml(bool english)
    {
        var wps = english ? RibbonXmlWpsEnglish : RibbonXmlWps;
        var powerPoint = System.Xml.Linq.XDocument.Parse(english ? RibbonXmlEnglish : RibbonXml);
        string Items(string comboBoxId)
        {
            var box = powerPoint.Descendants().First(element =>
                element.Name.LocalName == "comboBox" && (string?)element.Attribute("id") == comboBoxId);
            return string.Concat(box.Elements().Select(item =>
                item.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
                    .Replace(" xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"", string.Empty)));
        }
        return wps
            .Replace("__FONT_SIZE_ITEMS__", Items("VisualTeX.PowerPointVsto.FontSize"))
            .Replace("__INLINE_FONT_SIZE_ITEMS__", Items("VisualTeX.PowerPointVsto.InlineFontSize"))
            .Replace("__DISPLAY_FONT_SIZE_ITEMS__", Items("VisualTeX.PowerPointVsto.DisplayFontSize"));
    }
}

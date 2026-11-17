namespace VisualTeX.PowerPointVsto;

// The VisualTeX tab as WPS Presentation shows it, in WPS's own style (see VisualTeX.WordVsto/WordRibbonWps.cs for the
// naming and for how WPS lays columns out): every name says what it does, the formats are VisualTeX 公式 / WPS 公式
// (OMML) / MathType 公式 / LaTeX 代码, columns of two in a horizontal box per group, the size boxes beside the LaTeX
// menus, one menu per conversion target with the scopes spelled out in its items (WPS shows no menu section titles),
// WPS's own icons chosen by what they show. Built from PowerPoint's tab: the same controls and callbacks.
public sealed partial class ThisAddIn
{
    private const string RibbonXmlWps = """

<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="插入与编辑">
        <button id="VisualTeX.PowerPointVsto.New" label="插入公式" onAction="OnNewSvgFormula"  imageMso="EquationFractionGallery"  screentip="插入 VisualTeX 公式"  supertip="在当前幻灯片中央插入一个 VisualTeX 公式（SVG 图片）；编辑完成后可以像图片一样移动和缩放，双击即可再次编辑。"  size="large" />
        <box id="VisualTeX.PowerPointVsto.InsertColumns" boxStyle="horizontal">
        <box id="VisualTeX.PowerPointVsto.NewObjectBox" boxStyle="vertical"><button id="VisualTeX.PowerPointVsto.NewOle" label="OLE 公式" onAction="OnNewOleFormula"  imageMso="EquationScriptGallery"  screentip="插入 VisualTeX OLE 公式"  supertip="插入一个 VisualTeX 公式（OLE 对象，嵌入文件，双击即可再次编辑）。" /><button id="VisualTeX.PowerPointVsto.NewMathType" label="MathType 公式" onAction="OnNewMathTypeFormula"  imageMso="EquationOptions"  screentip="插入 MathType 公式"  supertip="在 VisualTeX 编辑器中写好后，插入一个 MathType 公式。" /></box>
        <box id="VisualTeX.PowerPointVsto.OmmlBox" boxStyle="vertical"><button id="VisualTeX.PowerPointVsto.InsertInlineOmml" label="WPS 行内公式" screentip="插入 WPS 行内公式" supertip="在 VisualTeX 编辑器中写好后，在文本框的光标处插入 WPS 自己的公式（OMML），前后文字保留；没有文本光标时新建一个只含公式的文本框。" imageMso="EquationProfessional" onAction="OnInsertInlineOmml" /><button id="VisualTeX.PowerPointVsto.InsertDisplayOmml" label="WPS 行间公式" screentip="插入 WPS 行间公式" supertip="在 VisualTeX 编辑器中写好后，在当前文本框中插入单独居中的 WPS 公式（OMML）段落；没有文本光标时新建一个只含公式的文本框。" imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" /></box>
        </box>
        <button id="VisualTeX.PowerPointVsto.Edit" label="编辑公式" onAction="OnEditSelected"  imageMso="KSEquationEditor"  screentip="编辑所选公式"  supertip="在 VisualTeX 编辑器中编辑所选的 VisualTeX、WPS 或 MathType 公式；也可以双击公式。"  size="large" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="格式转换">
        <box id="VisualTeX.PowerPointVsto.ConvertColumns" boxStyle="horizontal">
        <box id="VisualTeX.PowerPointVsto.ToOmmlMathTypeBox" boxStyle="vertical"><menu id="VisualTeX.PowerPointVsto.UnifyToOMML" label="转为 WPS 公式" screentip="转为 WPS 公式" supertip="把范围内其他格式的公式转为 WPS 公式，已是 WPS 公式的不变；同一公式不论原来是什么格式，结果都相同。若在左侧缩略图中选中了幻灯片，“所选”转换这些幻灯片上的公式。" imageMso="EquationTransferOMML"><button id="VisualTeX.PowerPointVsto.UnifySelectionOMML" label="所选公式转为 WPS 公式" tag="any|wordOmml|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="本页公式转为 WPS 公式" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="全部公式转为 WPS 公式" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.UnifyToMathType" label="转为 MathType 公式" screentip="转为 MathType 公式" supertip="把范围内其他格式的公式转为 MathType 公式，已是 MathType 公式的不变；同一公式不论原来是什么格式，结果都相同。若在左侧缩略图中选中了幻灯片，“所选”转换这些幻灯片上的公式。" imageMso="Convert"><button id="VisualTeX.PowerPointVsto.UnifySelectionMathType" label="所选公式转为 MathType 公式" tag="any|mathTypeOle|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="本页公式转为 MathType 公式" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="全部公式转为 MathType 公式" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" /></menu></box>
        <box id="VisualTeX.PowerPointVsto.ToVisualTeXBox" boxStyle="vertical"><menu id="VisualTeX.PowerPointVsto.UnifyToVisualTeX" label="转为 VisualTeX 公式" screentip="转为 VisualTeX 公式" supertip="把范围内其他格式的公式转为 VisualTeX 公式，已是 VisualTeX 公式的不变；同一公式不论原来是什么格式，结果都相同。若在左侧缩略图中选中了幻灯片，“所选”转换这些幻灯片上的公式。" imageMso="exporttopicture32"><button id="VisualTeX.PowerPointVsto.UnifySelectionVisualTeX" label="所选公式转为 VisualTeX 公式" tag="any|visualtex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="本页公式转为 VisualTeX 公式" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="全部公式转为 VisualTeX 公式" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" screentip="转为 VisualTeX 时使用的形式" supertip="选择“→ VisualTeX”和“统一为 VisualTeX”生成的对象：OLE 对象可双击编辑并嵌入文件；SVG 图片在任何设备上都能清晰显示，同样可双击编辑。" imageMso="InsertPictureFromFile">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE 对象（双击编辑，嵌入文件）" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG 图片（任何设备都能显示）" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu></box>
        </box>
        </group>
        <group id="VisualTeX.PowerPointVsto.LatexGroup" label="LaTeX 代码与字号">
        <box id="VisualTeX.PowerPointVsto.LatexSizeColumns" boxStyle="horizontal">
        <box id="VisualTeX.PowerPointVsto.LatexBox" boxStyle="vertical"><menu id="VisualTeX.PowerPointVsto.LatexRedraw" label="LaTeX 代码转为公式" screentip="LaTeX 代码转为公式" supertip="把文本框中写成文字的 LaTeX 公式（$…$、$$…$$、\(…\)、\[…\] 和 equation 等环境）换成公式。转为 WPS 公式时在文本框中原位替换；转为 VisualTeX 或 MathType 公式时，只含一个 LaTeX 公式的文本框换成同一位置的公式。" imageMso="EquationInsertLaTex"><button id="VisualTeX.PowerPointVsto.LatexToOmmlSelection" label="所选 LaTeX 代码转为 WPS 公式" tag="latex|wordOmml|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlSlide" label="本页 LaTeX 代码转为 WPS 公式" tag="latex|wordOmml|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlAll" label="全部 LaTeX 代码转为 WPS 公式" tag="latex|wordOmml|presentation" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.LatexToOmmlSeparator" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSelection" label="所选 LaTeX 代码转为 VisualTeX 公式" tag="latex|visualtex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSlide" label="本页 LaTeX 代码转为 VisualTeX 公式" tag="latex|visualtex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXAll" label="全部 LaTeX 代码转为 VisualTeX 公式" tag="latex|visualtex|presentation" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.LatexToVisualTeXSeparator" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeSelection" label="所选 LaTeX 代码转为 MathType 公式" tag="latex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeSlide" label="本页 LaTeX 代码转为 MathType 公式" tag="latex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeAll" label="全部 LaTeX 代码转为 MathType 公式" tag="latex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.FormulaToLatex" label="公式转为 LaTeX 代码" screentip="公式转为 LaTeX 代码" supertip="把公式（不论格式）换回 LaTeX 代码，之后可以再转回公式。" imageMso="EquationTransferLaTex"><button id="VisualTeX.PowerPointVsto.AnyToLatexSelection" label="所选公式转为 LaTeX 代码" tag="any|latex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.AnyToLatexSlide" label="本页公式转为 LaTeX 代码" tag="any|latex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.AnyToLatexAll" label="全部公式转为 LaTeX 代码" tag="any|latex|presentation" onAction="OnConvertFormulaFormat" /></menu></box>
        <box id="VisualTeX.PowerPointVsto.KindSizeBox" boxStyle="vertical"><comboBox id="VisualTeX.PowerPointVsto.InlineFontSize" label="行内公式字号" sizeString="初号（42 磅）" screentip="行内公式字号" supertip="选择字号后，当前演示文稿中所有行内公式立即改为这个字号，可撤销；之后 LaTeX 重绘和格式转换生成的行内公式也使用这个字号。选“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeFollow" label="跟随正文" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeChu" label="初号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoChu" label="小初" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeYi" label="一号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoYi" label="小一" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeEr" label="二号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoEr" label="小二" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeSan" label="三号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoSan" label="小三" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeSi" label="四号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoSi" label="小四" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeWu" label="五号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoWu" label="小五" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeLiu" label="六号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoLiu" label="小六" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeQi" label="七号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeBa" label="八号" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize8" label="8" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize9" label="9" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize10" label="10" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize10_5" label="10.5" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize11" label="11" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize12" label="12" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize14" label="14" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize16" label="16" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize18" label="18" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize20" label="20" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize24" label="24" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize28" label="28" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize36" label="36" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize48" label="48" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize72" label="72" />
          </comboBox><comboBox id="VisualTeX.PowerPointVsto.DisplayFontSize" label="行间公式字号" sizeString="初号（42 磅）" screentip="行间公式字号" supertip="选择字号后，当前演示文稿中所有行间公式立即改为这个字号，可撤销；之后 LaTeX 重绘和格式转换生成的行间公式也使用这个字号。选“跟随正文”不改动现有公式。" getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeFollow" label="跟随正文" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeChu" label="初号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoChu" label="小初" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeYi" label="一号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoYi" label="小一" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeEr" label="二号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoEr" label="小二" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeSan" label="三号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoSan" label="小三" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeSi" label="四号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoSi" label="小四" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeWu" label="五号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoWu" label="小五" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeLiu" label="六号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoLiu" label="小六" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeQi" label="七号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeBa" label="八号" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize8" label="8" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize9" label="9" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize10" label="10" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize10_5" label="10.5" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize11" label="11" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize12" label="12" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize14" label="14" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize16" label="16" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize18" label="18" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize20" label="20" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize24" label="24" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize28" label="28" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize36" label="36" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize48" label="48" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize72" label="72" />
          </comboBox></box>
        <box id="VisualTeX.PowerPointVsto.SelectedSizeBox" boxStyle="vertical"><comboBox id="VisualTeX.PowerPointVsto.FontSize" label="所选公式字号" sizeString="初号（42 磅）" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged" screentip="所选公式字号">
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
          </comboBox><labelControl id="VisualTeX.PowerPointVsto.SelectedSizeRow2" label="　" /></box>
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
        <group id="VisualTeX.PowerPointVsto.Group" label="Insert and Edit">
        <button id="VisualTeX.PowerPointVsto.New" label="Insert Equation" onAction="OnNewSvgFormula"  imageMso="EquationFractionGallery"  screentip="Insert a VisualTeX equation"  supertip="Insert a VisualTeX equation (SVG picture) in the middle of the current slide; move and resize it like a picture, double-click it to edit it again."  size="large" />
        <box id="VisualTeX.PowerPointVsto.InsertColumns" boxStyle="horizontal">
        <box id="VisualTeX.PowerPointVsto.NewObjectBox" boxStyle="vertical"><button id="VisualTeX.PowerPointVsto.NewOle" label="OLE Equation" onAction="OnNewOleFormula"  imageMso="EquationScriptGallery"  screentip="Insert a VisualTeX OLE equation"  supertip="Insert a VisualTeX equation as an OLE object (embedded, double-click to edit it again)." /><button id="VisualTeX.PowerPointVsto.NewMathType" label="MathType Equation" onAction="OnNewMathTypeFormula"  imageMso="EquationOptions"  screentip="Insert a MathType equation"  supertip="Write it in the VisualTeX editor, then insert it as a MathType equation." /></box>
        <box id="VisualTeX.PowerPointVsto.OmmlBox" boxStyle="vertical"><button id="VisualTeX.PowerPointVsto.InsertInlineOmml" label="WPS Inline Equation" screentip="Insert an inline WPS equation" supertip="Write it in the VisualTeX editor, then insert it at the text cursor as WPS's own equation (OMML), keeping the text around it; without a text cursor it gets a text box of its own." imageMso="EquationProfessional" onAction="OnInsertInlineOmml" /><button id="VisualTeX.PowerPointVsto.InsertDisplayOmml" label="WPS Display Equation" screentip="Insert a display WPS equation" supertip="Write it in the VisualTeX editor, then insert it as a centered WPS equation (OMML) paragraph in the current text box; without a text cursor it gets a text box of its own." imageMso="EquationInsertGallery" onAction="OnInsertDisplayOmml" /></box>
        </box>
        <button id="VisualTeX.PowerPointVsto.Edit" label="Edit" onAction="OnEditSelected"  imageMso="KSEquationEditor"  screentip="Edit the selected equation"  supertip="Edit the selected VisualTeX, WPS or MathType equation in the VisualTeX editor; double-clicking it does the same."  size="large" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="Convert Formats">
        <box id="VisualTeX.PowerPointVsto.ConvertColumns" boxStyle="horizontal">
        <box id="VisualTeX.PowerPointVsto.ToOmmlMathTypeBox" boxStyle="vertical"><menu id="VisualTeX.PowerPointVsto.UnifyToOMML" label="To WPS Equations" screentip="To WPS Equations" supertip="Turn the equations of other formats in the range into WPS Equations; equations already in that format stay as they are. The same equation gives the same result whatever its format. With slides selected in the thumbnail pane, Selected converts the equations on those slides." imageMso="EquationTransferOMML"><button id="VisualTeX.PowerPointVsto.UnifySelectionOMML" label="Selected Equations to WPS Equations" tag="any|wordOmml|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="Equations on This Slide to WPS Equations" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="All Equations to WPS Equations" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.UnifyToMathType" label="To MathType Equations" screentip="To MathType Equations" supertip="Turn the equations of other formats in the range into MathType Equations; equations already in that format stay as they are. The same equation gives the same result whatever its format. With slides selected in the thumbnail pane, Selected converts the equations on those slides." imageMso="Convert"><button id="VisualTeX.PowerPointVsto.UnifySelectionMathType" label="Selected Equations to MathType Equations" tag="any|mathTypeOle|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="Equations on This Slide to MathType Equations" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="All Equations to MathType Equations" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" /></menu></box>
        <box id="VisualTeX.PowerPointVsto.ToVisualTeXBox" boxStyle="vertical"><menu id="VisualTeX.PowerPointVsto.UnifyToVisualTeX" label="To VisualTeX Equations" screentip="To VisualTeX Equations" supertip="Turn the equations of other formats in the range into VisualTeX Equations; equations already in that format stay as they are. The same equation gives the same result whatever its format. With slides selected in the thumbnail pane, Selected converts the equations on those slides." imageMso="exporttopicture32"><button id="VisualTeX.PowerPointVsto.UnifySelectionVisualTeX" label="Selected Equations to VisualTeX Equations" tag="any|visualtex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="Equations on This Slide to VisualTeX Equations" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="All Equations to VisualTeX Equations" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" screentip="Form used when converting to VisualTeX" supertip="Choose what “→ VisualTeX” and “Unify → VisualTeX” create: an OLE object is embedded and double-click editable; an SVG picture displays sharply on every device and is also double-click editable." imageMso="InsertPictureFromFile">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE object (double-click to edit, embedded)" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG picture (displays on every device)" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu></box>
        </box>
        </group>
        <group id="VisualTeX.PowerPointVsto.LatexGroup" label="LaTeX Code and Size">
        <box id="VisualTeX.PowerPointVsto.LatexSizeColumns" boxStyle="horizontal">
        <box id="VisualTeX.PowerPointVsto.LatexBox" boxStyle="vertical"><menu id="VisualTeX.PowerPointVsto.LatexRedraw" label="LaTeX Code to Equations" screentip="LaTeX Code to Equations" supertip="Turn LaTeX written as text in text boxes ($…$, $$…$$, \(…\), \[…\] and environments such as equation) into equations. To WPS equations happens in place in the text box; to VisualTeX or MathType equations replaces a text box holding a single LaTeX equation with an equation at the same place." imageMso="EquationInsertLaTex"><button id="VisualTeX.PowerPointVsto.LatexToOmmlSelection" label="Selected LaTeX Code to WPS Equations" tag="latex|wordOmml|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlSlide" label="LaTeX Code on This Slide to WPS Equations" tag="latex|wordOmml|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlAll" label="All LaTeX Code to WPS Equations" tag="latex|wordOmml|presentation" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.LatexToOmmlSeparator" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSelection" label="Selected LaTeX Code to VisualTeX Equations" tag="latex|visualtex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSlide" label="LaTeX Code on This Slide to VisualTeX Equations" tag="latex|visualtex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXAll" label="All LaTeX Code to VisualTeX Equations" tag="latex|visualtex|presentation" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.LatexToVisualTeXSeparator" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeSelection" label="Selected LaTeX Code to MathType Equations" tag="latex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeSlide" label="LaTeX Code on This Slide to MathType Equations" tag="latex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeAll" label="All LaTeX Code to MathType Equations" tag="latex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.FormulaToLatex" label="Equations to LaTeX Code" screentip="Equations to LaTeX Code" supertip="Turn equations (any format) back into LaTeX code; they can be turned into equations again." imageMso="EquationTransferLaTex"><button id="VisualTeX.PowerPointVsto.AnyToLatexSelection" label="Selected Equations to LaTeX Code" tag="any|latex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.AnyToLatexSlide" label="Equations on This Slide to LaTeX Code" tag="any|latex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.AnyToLatexAll" label="All Equations to LaTeX Code" tag="any|latex|presentation" onAction="OnConvertFormulaFormat" /></menu></box>
        <box id="VisualTeX.PowerPointVsto.KindSizeBox" boxStyle="vertical"><comboBox id="VisualTeX.PowerPointVsto.InlineFontSize" label="Inline Size" sizeString="Follow Text" screentip="Inline Size" supertip="Once chosen, every inline formula of the current presentation takes this size at once (undoable), and LaTeX redraw and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeFollow" label="Follow Text" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeChu" label="42 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoChu" label="36 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeYi" label="26 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoYi" label="24 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeEr" label="22 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoEr" label="18 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeSan" label="16 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoSan" label="15 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeSi" label="14 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoSi" label="12 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeWu" label="10.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoWu" label="9 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeLiu" label="7.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeXiaoLiu" label="6.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeQi" label="5.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSizeBa" label="5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize8" label="8" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize9" label="9" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize10" label="10" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize10_5" label="10.5" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize11" label="11" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize12" label="12" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize14" label="14" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize16" label="16" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize18" label="18" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize20" label="20" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize24" label="24" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize28" label="28" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize36" label="36" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize48" label="48" />
            <item id="VisualTeX.PowerPointVsto.InlineFontSize72" label="72" />
          </comboBox><comboBox id="VisualTeX.PowerPointVsto.DisplayFontSize" label="Display Size" sizeString="Follow Text" screentip="Display Size" supertip="Once chosen, every display formula of the current presentation takes this size at once (undoable), and LaTeX redraw and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeFollow" label="Follow Text" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeChu" label="42 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoChu" label="36 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeYi" label="26 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoYi" label="24 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeEr" label="22 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoEr" label="18 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeSan" label="16 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoSan" label="15 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeSi" label="14 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoSi" label="12 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeWu" label="10.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoWu" label="9 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeLiu" label="7.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeXiaoLiu" label="6.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeQi" label="5.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSizeBa" label="5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize8" label="8" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize9" label="9" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize10" label="10" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize10_5" label="10.5" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize11" label="11" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize12" label="12" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize14" label="14" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize16" label="16" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize18" label="18" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize20" label="20" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize24" label="24" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize28" label="28" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize36" label="36" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize48" label="48" />
            <item id="VisualTeX.PowerPointVsto.DisplayFontSize72" label="72" />
          </comboBox></box>
        <box id="VisualTeX.PowerPointVsto.SelectedSizeBox" boxStyle="vertical"><comboBox id="VisualTeX.PowerPointVsto.FontSize" label="Selected Size" sizeString="42 pt" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged" screentip="Selected Size">
            <item id="VisualTeX.PowerPointVsto.FontSizeChu" label="42 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoChu" label="36 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeYi" label="26 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoYi" label="24 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeEr" label="22 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoEr" label="18 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeSan" label="16 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoSan" label="15 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeSi" label="14 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoSi" label="12 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeWu" label="10.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoWu" label="9 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeLiu" label="7.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeXiaoLiu" label="6.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeQi" label="5.5 pt (CN preset)" />
            <item id="VisualTeX.PowerPointVsto.FontSizeBa" label="5 pt (CN preset)" />
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
          </comboBox><labelControl id="VisualTeX.PowerPointVsto.SelectedSizeRow2" label="　" /></box>
        </box>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";
}

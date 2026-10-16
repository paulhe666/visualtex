namespace VisualTeX.PowerPointVsto;

public sealed partial class ThisAddIn
{
    private static string T(string chinese, string english) =>
        VisualTeX.WindowsOffice.VstoShared.OfficePluginLanguage.Text(chinese, english);

    private const string RibbonXmlEnglish = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.PowerPointVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.PowerPointVsto.Group" label="VisualTeX">
          <button id="VisualTeX.PowerPointVsto.New" label="New SVG" size="large" tag="insertFormula" getImage="GetRibbonImage" onAction="OnNewSvgFormula" />
          <button id="VisualTeX.PowerPointVsto.NewOle" label="New OLE" size="large" tag="convertToOle" getImage="GetRibbonImage" onAction="OnNewOleFormula" />
          <button id="VisualTeX.PowerPointVsto.NewMathType" label="New MathType" size="large" tag="convertToOle" getImage="GetRibbonImage" onAction="OnNewMathTypeFormula" />
          <button id="VisualTeX.PowerPointVsto.InsertInlineOmml" label="Insert Inline OMML" size="large" screentip="Insert an inline native PowerPoint equation" supertip="Insert native Office Math at the current text-box caret or text selection as inline math." imageMso="EquationInsertNew" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.PowerPointVsto.InsertDisplayOmml" label="Insert Display OMML" size="large" screentip="Insert a display native PowerPoint equation" supertip="Insert a separate equation paragraph at the current text caret when editing a text box; otherwise create a standalone native Office Math equation on the current slide." imageMso="EquationInsertNew" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="Edit Selected Formula" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
        </group>
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="Equation Format Conversion">
          <box id="VisualTeX.PowerPointVsto.FormatConversionLabelsA" boxStyle="vertical">
            <labelControl id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeLabel" label="VisualTeX → MathType" />
            <labelControl id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXLabel" label="MathType → VisualTeX" />
            <labelControl id="VisualTeX.PowerPointVsto.OmmlToMathTypeLabel" label="OMML → MathType" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSelectionA" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeSelection" label="Selected" screentip="VisualTeX → MathType (Selected)" supertip="Convert the VisualTeX equations among the selected objects to MathType. With slides selected in the thumbnail pane, converts every VisualTeX equation on those slides." tag="visualtex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXSelection" label="Selected" screentip="MathType → VisualTeX (Selected)" supertip="Convert the MathType equations among the selected objects to VisualTeX. With slides selected in the thumbnail pane, converts every MathType equation on those slides." tag="mathTypeOle|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToMathTypeSelection" label="Selected" screentip="OMML → MathType (Selected)" supertip="Convert the native PowerPoint (OMML) equations among the selected objects to MathType. With slides selected in the thumbnail pane, converts every native PowerPoint (OMML) equation on those slides." tag="wordOmml|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSlideA" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeSlide" label="Slide" screentip="VisualTeX → MathType (Slide)" supertip="Convert every VisualTeX equation on the current slide to MathType." tag="visualtex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXSlide" label="Slide" screentip="MathType → VisualTeX (Slide)" supertip="Convert every MathType equation on the current slide to VisualTeX." tag="mathTypeOle|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToMathTypeSlide" label="Slide" screentip="OMML → MathType (Slide)" supertip="Convert every native PowerPoint (OMML) equation on the current slide to MathType." tag="wordOmml|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionAllA" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.VisualTeXToMathTypeAll" label="All" screentip="VisualTeX → MathType (All)" supertip="Convert every VisualTeX equation in the presentation to MathType. You are asked to confirm first." tag="visualtex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToVisualTeXAll" label="All" screentip="MathType → VisualTeX (All)" supertip="Convert every MathType equation in the presentation to VisualTeX. You are asked to confirm first." tag="mathTypeOle|visualtex|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToMathTypeAll" label="All" screentip="OMML → MathType (All)" supertip="Convert every native PowerPoint (OMML) equation in the presentation to MathType. You are asked to confirm first." tag="wordOmml|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
          </box>
          <separator id="VisualTeX.PowerPointVsto.FormatConversionSeparator1" />
          <box id="VisualTeX.PowerPointVsto.FormatConversionLabelsB" boxStyle="vertical">
            <labelControl id="VisualTeX.PowerPointVsto.MathTypeToOmmlLabel" label="MathType → OMML" />
            <labelControl id="VisualTeX.PowerPointVsto.VisualTeXToOmmlLabel" label="VisualTeX → OMML" />
            <labelControl id="VisualTeX.PowerPointVsto.OmmlToVisualTeXLabel" label="OMML → VisualTeX" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSelectionB" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.MathTypeToOmmlSelection" label="Selected" screentip="MathType → OMML (Selected)" supertip="Convert the MathType equations among the selected objects to OMML. With slides selected in the thumbnail pane, converts every MathType equation on those slides." tag="mathTypeOle|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToOmmlSelection" label="Selected" screentip="VisualTeX → OMML (Selected)" supertip="Convert the VisualTeX equations among the selected objects to OMML. With slides selected in the thumbnail pane, converts every VisualTeX equation on those slides." tag="visualtex|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToVisualTeXSelection" label="Selected" screentip="OMML → VisualTeX (Selected)" supertip="Convert the native PowerPoint (OMML) equations among the selected objects to VisualTeX. With slides selected in the thumbnail pane, converts every native PowerPoint (OMML) equation on those slides." tag="wordOmml|visualtex|selection" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionSlideB" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.MathTypeToOmmlSlide" label="Slide" screentip="MathType → OMML (Slide)" supertip="Convert every MathType equation on the current slide to OMML." tag="mathTypeOle|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToOmmlSlide" label="Slide" screentip="VisualTeX → OMML (Slide)" supertip="Convert every VisualTeX equation on the current slide to OMML." tag="visualtex|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToVisualTeXSlide" label="Slide" screentip="OMML → VisualTeX (Slide)" supertip="Convert every native PowerPoint (OMML) equation on the current slide to VisualTeX." tag="wordOmml|visualtex|slide" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.FormatConversionAllB" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.MathTypeToOmmlAll" label="All" screentip="MathType → OMML (All)" supertip="Convert every MathType equation in the presentation to OMML. You are asked to confirm first." tag="mathTypeOle|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToOmmlAll" label="All" screentip="VisualTeX → OMML (All)" supertip="Convert every VisualTeX equation in the presentation to OMML. You are asked to confirm first." tag="visualtex|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToVisualTeXAll" label="All" screentip="OMML → VisualTeX (All)" supertip="Convert every native PowerPoint (OMML) equation in the presentation to VisualTeX. You are asked to confirm first." tag="wordOmml|visualtex|presentation" onAction="OnConvertFormulaFormat" />
          </box>
          <separator id="VisualTeX.PowerPointVsto.FormatConversionSeparator2" />
          <menu id="VisualTeX.PowerPointVsto.UnifyFormat" label="Unify Format" size="large" screentip="Convert every equation to one format" supertip="Convert the equations in other formats on the current slide or in the whole presentation to the chosen format. Equations already in that format are left as they are." tag="batchImport" getImage="GetRibbonImage">
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifySlideHeader" title="This slide" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="This Slide → MathType" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="This Slide → OMML" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="This Slide → VisualTeX" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.UnifyAllHeader" title="Whole presentation" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="All Slides → MathType" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="All Slides → OMML" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="All Slides → VisualTeX" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" />
          </menu>
          <menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" size="large" screentip="Form used when converting to VisualTeX" supertip="Choose what “→ VisualTeX” and “Unify → VisualTeX” create: an OLE object is embedded and double-click editable; an SVG picture displays sharply on every device and is also double-click editable." tag="convertToOle" getImage="GetRibbonImage">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE object (double-click to edit, embedded)" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG picture (displays on every device)" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu>
        </group>
        <group id="VisualTeX.PowerPointVsto.FontSizeGroup" label="Formula Font Size">
          <button id="VisualTeX.PowerPointVsto.FontSizeDecrease" label="Decrease" imageMso="FontSizeDecrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
          <comboBox id="VisualTeX.PowerPointVsto.FontSize" label="Font Size" sizeString="42 pt" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
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
          </comboBox>
          <button id="VisualTeX.PowerPointVsto.FontSizeIncrease" label="Increase" imageMso="FontSizeIncrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";
}

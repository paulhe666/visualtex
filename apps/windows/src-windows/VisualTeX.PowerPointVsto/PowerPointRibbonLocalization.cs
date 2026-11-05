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
        <group id="VisualTeX.PowerPointVsto.LatexRedrawGroup" label="LaTeX Redraw">
          <box id="VisualTeX.PowerPointVsto.LatexRedrawLabels" boxStyle="vertical">
            <labelControl id="VisualTeX.PowerPointVsto.LatexToOmmlLabel" label="LaTeX → OMML" />
            <labelControl id="VisualTeX.PowerPointVsto.LatexToVisualTeXLabel" label="LaTeX → VisualTeX" />
            <labelControl id="VisualTeX.PowerPointVsto.LatexToMathTypeLabel" label="LaTeX → MathType" />
          </box>
          <box id="VisualTeX.PowerPointVsto.LatexRedrawSelection" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlSelection" label="Selected" screentip="LaTeX → OMML (Selected)" supertip="Redraw the LaTeX formulas ($…$, $$…$$, \(…\), \[…\] and environments such as equation) in the selected text boxes as native PowerPoint (OMML) equations where they stand; the other text stays as it is. With part of a text box's text selected, only the complete formulas in it are redrawn. With slides selected in the thumbnail pane, redraws every LaTeX formula on those slides." tag="latex|wordOmml|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSelection" label="Selected" screentip="LaTeX → VisualTeX (Selected)" supertip="Replace each selected text box that holds just one LaTeX formula with a VisualTeX equation (OLE or SVG, see “VisualTeX as”), keeping its center and font size. With slides selected in the thumbnail pane, converts every such text box on those slides. For formulas inside text, use LaTeX → OMML." tag="latex|visualtex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeSelection" label="Selected" screentip="LaTeX → MathType (Selected)" supertip="Replace each selected text box that holds just one LaTeX formula with a MathType equation, keeping its center and font size. With slides selected in the thumbnail pane, converts every such text box on those slides. For formulas inside text, use LaTeX → OMML." tag="latex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.LatexRedrawSlide" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlSlide" label="Slide" screentip="LaTeX → OMML (Slide)" supertip="Redraw every LaTeX formula in the text boxes of the current slide as a native PowerPoint (OMML) equation where it stands; the other text stays as it is." tag="latex|wordOmml|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSlide" label="Slide" screentip="LaTeX → VisualTeX (Slide)" supertip="Replace every text box on the current slide that holds just one LaTeX formula with a VisualTeX equation." tag="latex|visualtex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeSlide" label="Slide" screentip="LaTeX → MathType (Slide)" supertip="Replace every text box on the current slide that holds just one LaTeX formula with a MathType equation." tag="latex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" />
          </box>
          <box id="VisualTeX.PowerPointVsto.LatexRedrawAll" boxStyle="vertical">
            <button id="VisualTeX.PowerPointVsto.LatexToOmmlAll" label="All" screentip="LaTeX → OMML (All)" supertip="Redraw every LaTeX formula in the text boxes of the presentation as a native PowerPoint (OMML) equation where it stands. You are asked to confirm first." tag="latex|wordOmml|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToVisualTeXAll" label="All" screentip="LaTeX → VisualTeX (All)" supertip="Replace every text box in the presentation that holds just one LaTeX formula with a VisualTeX equation. You are asked to confirm first." tag="latex|visualtex|presentation" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.LatexToMathTypeAll" label="All" screentip="LaTeX → MathType (All)" supertip="Replace every text box in the presentation that holds just one LaTeX formula with a MathType equation. You are asked to confirm first." tag="latex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" />
          </box>
          <separator id="VisualTeX.PowerPointVsto.LatexRedrawSeparator" />
          <menu id="VisualTeX.PowerPointVsto.FormulaToLatex" label="To LaTeX Code" size="large" screentip="Turn equations into LaTeX code" supertip="Turn VisualTeX, MathType and native PowerPoint (OMML) equations back into LaTeX code: a native equation in text becomes $…$ where it stands ($$…$$ when it is alone in its paragraph) and the other text stays as it is; an equation object becomes a text box with $$…$$ at the same center and font size. LaTeX → … draws them again." imageMso="TextBoxInsert">
            <menuSeparator id="VisualTeX.PowerPointVsto.AnyToLatexHeader" title="Equations in every format" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexSelection" label="Selected Equations → LaTeX Code" screentip="Equations → LaTeX Code (Selected)" supertip="Turn the VisualTeX, MathType and OMML equations among the selected objects into LaTeX code. With part of a text box's text selected, only the equations in it are converted. With slides selected in the thumbnail pane, converts every equation on those slides." tag="any|latex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexSlide" label="This Slide → LaTeX Code" screentip="Equations → LaTeX Code (Slide)" supertip="Turn every VisualTeX, MathType and OMML equation on the current slide into LaTeX code." tag="any|latex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.AnyToLatexAll" label="All Slides → LaTeX Code" screentip="Equations → LaTeX Code (All)" supertip="Turn every VisualTeX, MathType and OMML equation in the presentation into LaTeX code. You are asked to confirm first." tag="any|latex|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.VisualTeXToLatexHeader" title="VisualTeX equations only" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToLatexSelection" label="Selected VisualTeX Equations" screentip="VisualTeX → LaTeX Code (Selected)" supertip="Replace the VisualTeX equations among the selected objects with text boxes holding $$…$$. With slides selected in the thumbnail pane, converts every VisualTeX equation on those slides." tag="visualtex|latex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToLatexSlide" label="VisualTeX Equations on This Slide" screentip="VisualTeX → LaTeX Code (Slide)" supertip="Replace every VisualTeX equation on the current slide with a text box holding $$…$$." tag="visualtex|latex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.VisualTeXToLatexAll" label="All VisualTeX Equations" screentip="VisualTeX → LaTeX Code (All)" supertip="Replace every VisualTeX equation in the presentation with a text box holding $$…$$. You are asked to confirm first." tag="visualtex|latex|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.MathTypeToLatexHeader" title="MathType equations only" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToLatexSelection" label="Selected MathType Equations" screentip="MathType → LaTeX Code (Selected)" supertip="Replace the MathType equations among the selected objects with text boxes holding $$…$$. With slides selected in the thumbnail pane, converts every MathType equation on those slides." tag="mathTypeOle|latex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToLatexSlide" label="MathType Equations on This Slide" screentip="MathType → LaTeX Code (Slide)" supertip="Replace every MathType equation on the current slide with a text box holding $$…$$." tag="mathTypeOle|latex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.MathTypeToLatexAll" label="All MathType Equations" screentip="MathType → LaTeX Code (All)" supertip="Replace every MathType equation in the presentation with a text box holding $$…$$. You are asked to confirm first." tag="mathTypeOle|latex|presentation" onAction="OnConvertFormulaFormat" />
            <menuSeparator id="VisualTeX.PowerPointVsto.OmmlToLatexHeader" title="OMML equations only" />
            <button id="VisualTeX.PowerPointVsto.OmmlToLatexSelection" label="Selected OMML Equations" screentip="OMML → LaTeX Code (Selected)" supertip="Replace the native PowerPoint (OMML) equations in the selected text boxes (or the selected text) with $…$ or $$…$$ where they stand; the other text stays as it is. With slides selected in the thumbnail pane, converts every OMML equation on those slides." tag="wordOmml|latex|selection" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToLatexSlide" label="OMML Equations on This Slide" screentip="OMML → LaTeX Code (Slide)" supertip="Replace every native PowerPoint (OMML) equation on the current slide with $…$ or $$…$$ where it stands; the other text stays as it is." tag="wordOmml|latex|slide" onAction="OnConvertFormulaFormat" />
            <button id="VisualTeX.PowerPointVsto.OmmlToLatexAll" label="All OMML Equations" screentip="OMML → LaTeX Code (All)" supertip="Replace every native PowerPoint (OMML) equation in the presentation with $…$ or $$…$$ where it stands. You are asked to confirm first." tag="wordOmml|latex|presentation" onAction="OnConvertFormulaFormat" />
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
          <comboBox id="VisualTeX.PowerPointVsto.InlineFontSize" label="Inline" sizeString="Follow Text" screentip="Inline formula size" supertip="Once chosen, every inline formula of the current presentation takes this size at once (undoable), and LaTeX redraw and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
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
          </comboBox>
          <comboBox id="VisualTeX.PowerPointVsto.DisplayFontSize" label="Display" sizeString="Follow Text" screentip="Display formula size" supertip="Once chosen, every display formula of the current presentation takes this size at once (undoable), and LaTeX redraw and format conversion give their new ones this size. Follow Text leaves existing formulas as they are." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
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
          </comboBox>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";
}

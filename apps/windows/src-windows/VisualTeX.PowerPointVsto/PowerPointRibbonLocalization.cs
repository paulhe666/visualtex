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
        <group id="VisualTeX.PowerPointVsto.FormatConversionGroup" label="Equation Format Conversion"><menu id="VisualTeX.PowerPointVsto.UnifyFormat" label="Convert Equations" size="large" screentip="Convert equation format" supertip="Turn every equation of another format (VisualTeX, OMML, MathType) in the range into the chosen format; equations already in it stay as they are. The same equation gives the same result whatever its source. With slides selected in the thumbnail pane, Selection converts the equations on those slides." tag="batchImport" getImage="GetRibbonImage"><menuSeparator id="VisualTeX.PowerPointVsto.UnifySelectionHeader" title="Selected equations" /><button id="VisualTeX.PowerPointVsto.UnifySelectionOMML" label="To OMML" tag="any|wordOmml|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySelectionMathType" label="To MathType" tag="any|mathTypeOle|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySelectionVisualTeX" label="To VisualTeX" tag="any|visualtex|selection" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.UnifySlideHeader" title="Current slide" /><button id="VisualTeX.PowerPointVsto.UnifySlideOMML" label="To OMML" tag="any|wordOmml|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideMathType" label="To MathType" tag="any|mathTypeOle|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifySlideVisualTeX" label="To VisualTeX" tag="any|visualtex|slide" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.UnifyAllHeader" title="Whole presentation" /><button id="VisualTeX.PowerPointVsto.UnifyAllOMML" label="To OMML" tag="any|wordOmml|presentation" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllMathType" label="To MathType" tag="any|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.UnifyAllVisualTeX" label="To VisualTeX" tag="any|visualtex|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.VisualTeXTargetForm" getLabel="GetVisualTeXTargetFormLabel" size="large" screentip="Form used when converting to VisualTeX" supertip="Choose what “→ VisualTeX” and “Unify → VisualTeX” create: an OLE object is embedded and double-click editable; an SVG picture displays sharply on every device and is also double-click editable." tag="convertToOle" getImage="GetRibbonImage">
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetOle" label="OLE object (double-click to edit, embedded)" tag="nativeOle" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
            <toggleButton id="VisualTeX.PowerPointVsto.VisualTeXTargetSvg" label="SVG picture (displays on every device)" tag="crossPlatformPicture" getPressed="GetVisualTeXTargetFormPressed" onAction="OnVisualTeXTargetFormChanged" />
          </menu><button id="VisualTeX.PowerPointVsto.UpdateEquationFormat" label="Update Equation Format" size="large" screentip="Update equation format" supertip="Draw the VisualTeX equations (SVG pictures and OLE objects) drawn in other fonts again in the current fonts (the math font, the Chinese font and the fonts for letters, Greek letters and digits); each keeps its position, size, rotation and layer, and its content does not change." imageMso="Refresh" onAction="OnUpdateEquationFormat" /></group>
        <group id="VisualTeX.PowerPointVsto.LatexRedrawGroup" label="LaTeX Redraw"><menu id="VisualTeX.PowerPointVsto.LatexRedraw" label="LaTeX Redraw" size="large" screentip="Redraw LaTeX code as equations" supertip="Redraw LaTeX equations ($…$, $$…$$, \(…\), \[…\] and environments such as equation). Redrawing as OMML happens in place in the text box; redrawing as VisualTeX or MathType replaces a text box holding a single LaTeX equation with an equation at the same place." tag="batchImport" getImage="GetRibbonImage"><menuSeparator id="VisualTeX.PowerPointVsto.LatexToOmmlHeader" title="Redraw as OMML" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlSelection" label="Selection" tag="latex|wordOmml|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlSlide" label="Current slide" tag="latex|wordOmml|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToOmmlAll" label="Whole presentation" tag="latex|wordOmml|presentation" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.LatexToVisualTeXHeader" title="Redraw as VisualTeX" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSelection" label="Selection" tag="latex|visualtex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXSlide" label="Current slide" tag="latex|visualtex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToVisualTeXAll" label="Whole presentation" tag="latex|visualtex|presentation" onAction="OnConvertFormulaFormat" /><menuSeparator id="VisualTeX.PowerPointVsto.LatexToMathTypeHeader" title="Redraw as MathType" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeSelection" label="Selection" tag="latex|mathTypeOle|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeSlide" label="Current slide" tag="latex|mathTypeOle|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.LatexToMathTypeAll" label="Whole presentation" tag="latex|mathTypeOle|presentation" onAction="OnConvertFormulaFormat" /></menu><menu id="VisualTeX.PowerPointVsto.FormulaToLatex" label="To LaTeX Code" size="large" screentip="Turn equations into LaTeX code" supertip="Turn VisualTeX, MathType and native PowerPoint (OMML) equations back into LaTeX code: a native equation in text becomes $…$ where it stands ($$…$$ when it is alone in its paragraph) and the other text stays as it is; an equation object becomes a text box with $$…$$ at the same center and font size. LaTeX → … draws them again." imageMso="TextBoxInsert"><menuSeparator id="VisualTeX.PowerPointVsto.AnyToLatexHeader" title="Equations in every format" /><button id="VisualTeX.PowerPointVsto.AnyToLatexSelection" label="Selected Equations → LaTeX Code" screentip="Equations → LaTeX Code (Selected)" supertip="Turn the VisualTeX, MathType and OMML equations among the selected objects into LaTeX code. With part of a text box's text selected, only the equations in it are converted. With slides selected in the thumbnail pane, converts every equation on those slides." tag="any|latex|selection" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.AnyToLatexSlide" label="This Slide → LaTeX Code" screentip="Equations → LaTeX Code (Slide)" supertip="Turn every VisualTeX, MathType and OMML equation on the current slide into LaTeX code." tag="any|latex|slide" onAction="OnConvertFormulaFormat" /><button id="VisualTeX.PowerPointVsto.AnyToLatexAll" label="All Slides → LaTeX Code" screentip="Equations → LaTeX Code (All)" supertip="Turn every VisualTeX, MathType and OMML equation in the presentation into LaTeX code. You are asked to confirm first." tag="any|latex|presentation" onAction="OnConvertFormulaFormat" /></menu></group>
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

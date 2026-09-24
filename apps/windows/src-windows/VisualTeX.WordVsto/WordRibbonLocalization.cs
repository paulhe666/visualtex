namespace VisualTeX.WordVsto;

public sealed partial class ThisAddIn
{
    private static string T(string chinese, string english) =>
        VisualTeX.WindowsOffice.VstoShared.OfficePluginLanguage.Text(chinese, english);

    private const string RibbonXmlEnglish = """
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="VisualTeX.WordVsto.Tab" label="VisualTeX" insertAfterMso="TabHome">
        <group id="VisualTeX.WordVsto.Group" label="VisualTeX">
          <button id="VisualTeX.WordVsto.Inline" label="Inline OLE" size="large" tag="oleInline" getImage="GetRibbonImage" onAction="OnInsertInline" />
          <button id="VisualTeX.WordVsto.Display" label="Display OLE" size="large" tag="oleDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplay" />
          <button id="VisualTeX.WordVsto.InlineOmml" label="Inline OMML" size="large" screentip="Insert a native Word equation" supertip="Insert an inline OMML equation that remains editable with Word equation tools while retaining VisualTeX LaTeX metadata." tag="ommlInline" getImage="GetRibbonImage" onAction="OnInsertInlineOmml" />
          <button id="VisualTeX.WordVsto.DisplayOmml" label="Display OMML" size="large" screentip="Insert a native Word equation" supertip="Insert a display OMML equation that remains editable with Word equation tools while retaining VisualTeX LaTeX metadata." tag="ommlDisplay" getImage="GetRibbonImage" onAction="OnInsertDisplayOmml" />
          <button id="VisualTeX.WordVsto.Edit" label="Edit Selected Formula" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <box id="VisualTeX.WordVsto.FormatConversionBox" boxStyle="vertical">
            <menu id="VisualTeX.WordVsto.VisualTeXToMathType" label="VisualTeX → MathType" screentip="Redraw as MathType OLE" supertip="Remove the original VisualTeX host and numbering, then redraw through the standard MathType insertion path.">
              <button id="VisualTeX.WordVsto.VisualTeXToMathTypeSelection" label="Convert Selection" onAction="OnConvertVisualTeXToMathTypeSelection" />
              <button id="VisualTeX.WordVsto.VisualTeXToMathTypeDocument" label="Convert Entire Document" onAction="OnConvertVisualTeXToMathTypeDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.MathTypeToVisualTeX" label="MathType → VisualTeX" screentip="Redraw as VisualTeX OLE" supertip="Remove the original MathType host and numbering, then redraw through the standard VisualTeX OLE insertion path.">
              <button id="VisualTeX.WordVsto.MathTypeToVisualTeXSelection" label="Convert Selection" onAction="OnConvertMathTypeToVisualTeXSelection" />
              <button id="VisualTeX.WordVsto.MathTypeToVisualTeXDocument" label="Convert Entire Document" onAction="OnConvertMathTypeToVisualTeXDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.OmmlToMathType" label="OMML → MathType" screentip="Convert native Word equations to MathType OLE" supertip="Read the real MathML from Word OMath/OMathPara, remove the original OMML host and numbering, then redraw as a self-contained MathType Equation.DSMT4 object.">
              <button id="VisualTeX.WordVsto.OmmlToMathTypeSelection" label="Convert Selection" onAction="OnConvertOmmlToMathTypeSelection" />
              <button id="VisualTeX.WordVsto.OmmlToMathTypeDocument" label="Convert Entire Document" onAction="OnConvertOmmlToMathTypeDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.MathTypeToOmml" label="MathType → OMML" screentip="Convert MathType OLE to native Word equations" supertip="Read MathML directly from Equation Native without launching MathType, then redraw in place through VisualTeX's Word OMML insertion and numbering path.">
              <button id="VisualTeX.WordVsto.MathTypeToOmmlSelection" label="Convert Selection" onAction="OnConvertMathTypeToOmmlSelection" />
              <button id="VisualTeX.WordVsto.MathTypeToOmmlDocument" label="Convert Entire Document" onAction="OnConvertMathTypeToOmmlDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.VisualTeXToOmml" label="VisualTeX → OMML" screentip="Convert VisualTeX OLE to Word OMML" supertip="Convert VisualTeX OLE equations in the selection or entire document to native Word OMML while preserving numbering and equation references.">
              <button id="VisualTeX.WordVsto.VisualTeXToOmmlSelection" label="Convert Selection" onAction="OnConvertVisualTeXToOmmlSelection" />
              <button id="VisualTeX.WordVsto.VisualTeXToOmmlDocument" label="Convert Entire Document" onAction="OnConvertVisualTeXToOmmlDocument" />
            </menu>
            <menu id="VisualTeX.WordVsto.OmmlToVisualTeX" label="OMML → VisualTeX" screentip="Convert Word OMML to VisualTeX OLE" supertip="Convert native Word OMML equations in the selection or entire document to VisualTeX OLE while preserving numbering and equation references.">
              <button id="VisualTeX.WordVsto.OmmlToVisualTeXSelection" label="Convert Selection" onAction="OnConvertOmmlToVisualTeXSelection" />
              <button id="VisualTeX.WordVsto.OmmlToVisualTeXDocument" label="Convert Entire Document" onAction="OnConvertOmmlToVisualTeXDocument" />
            </menu>
          </box>
          <box id="VisualTeX.WordVsto.NumberingBox" boxStyle="vertical">
            <button id="VisualTeX.WordVsto.UpdateNumbers" label="Update Equation Numbers" screentip="Update VisualTeX and MathType equation numbers" supertip="Refresh VisualTeX numbering and native MathType MTChap/MTSec/MTEqn numbering and equation references in the current document." tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationNumbers" />
            <menu id="VisualTeX.WordVsto.NumberFormat" label="Number Format" screentip="Set the equation number format for this document" supertip="Update existing VisualTeX and native MathType MTPlaceRef numbers and use the selected format for newly inserted numbered equations.">
              <toggleButton id="VisualTeX.WordVsto.NumberFormatContinuous" label="Continuous (1)" tag="continuous" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dot" label="By chapter (1.1)" tag="heading1-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading1Dash" label="By chapter (1-1)" tag="heading1-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dot" label="By section (1.1.1)" tag="heading2-dot" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
              <toggleButton id="VisualTeX.WordVsto.NumberFormatHeading2Dash" label="By section (1.1-1)" tag="heading2-dash" getPressed="GetEquationNumberFormatPressed" onAction="OnEquationNumberFormatChanged" />
            </menu>
            <button id="VisualTeX.WordVsto.InsertReference" label="Insert Equation Reference" screentip="Reference a numbered equation" supertip="Choose a numbered VisualTeX or MathType equation in the current document. VisualTeX uses Word REF; MathType retains its native ZEqnNum/GOTOBUTTON/REF reference structure." imageMso="HyperlinkInsert" onAction="OnInsertEquationReference" />
          </box>
          <button id="VisualTeX.WordVsto.BulkImport" label="Bulk Import" size="large" screentip="Import LaTeX / Markdown" supertip="Parse Markdown or LaTeX into native Word text and independently editable inline/display equations." tag="batchImport" getImage="GetRibbonImage" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.RedrawGroup" label="LaTeX Redraw">
          <menu id="VisualTeX.WordVsto.RedrawSelection" label="Redraw Selection" size="large" screentip="Redraw selected equations or LaTeX" supertip="Redraw selected LaTeX in place as VisualTeX OLE, Word OMML or MathType, or restore any of those three equation types to LaTeX source." tag="batchImport" getImage="GetRibbonImage">
            <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="LaTeX → Word OMML" screentip="Replace in place with a native Word equation" onAction="OnRedrawSelectionToOmml" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="LaTeX → VisualTeX OLE" screentip="Replace in place with a double-click editable VisualTeX OLE" onAction="OnRedrawSelectionToOle" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="LaTeX → MathType" screentip="Replace in place with a MathType Equation.DSMT4 object" onAction="OnRedrawSelectionToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawSelectionSeparator" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOleToLatex" label="VisualTeX OLE → LaTeX" screentip="Restore the selected VisualTeX OLE to LaTeX source" onAction="OnRedrawSelectionOleToLatex" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOmmlToLatex" label="OMML → LaTeX" screentip="Restore the selected Word OMML equation to LaTeX source" onAction="OnRedrawSelectionOmmlToLatex" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathTypeToLatex" label="MathType → LaTeX" screentip="Read MathType Equation Native directly and restore LaTeX source" onAction="OnRedrawSelectionMathTypeToLatex" />
          </menu>
          <menu id="VisualTeX.WordVsto.RedrawDocument" label="Redraw Entire Document" size="large" screentip="Redraw equations or LaTeX in the entire document" supertip="Redraw all LaTeX in place as VisualTeX OLE, Word OMML or MathType, or restore all three equation types to LaTeX source. A confirmation is shown before changes begin." imageMso="RefreshAll">
            <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="All LaTeX → Word OMML" onAction="OnRedrawDocumentToOmml" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="All LaTeX → VisualTeX OLE" onAction="OnRedrawDocumentToOle" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="All LaTeX → MathType" onAction="OnRedrawDocumentToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawDocumentSeparator" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOleToLatex" label="All VisualTeX OLE → LaTeX" onAction="OnRedrawDocumentOleToLatex" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOmmlToLatex" label="All OMML → LaTeX" onAction="OnRedrawDocumentOmmlToLatex" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathTypeToLatex" label="All MathType → LaTeX" onAction="OnRedrawDocumentMathTypeToLatex" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.FontSizeGroup" label="Formula Font Size">
          <button id="VisualTeX.WordVsto.FontSizeDecrease" label="Decrease" imageMso="FontSizeDecrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnDecreaseFormulaFontSize" />
          <comboBox id="VisualTeX.WordVsto.FontSize" label="Font Size" sizeString="42 pt" getText="GetFormulaFontSizeText" getEnabled="GetFormulaFontSizeEnabled" onChange="OnFormulaFontSizeChanged">
            <item id="VisualTeX.WordVsto.FontSizeChu" label="42 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoChu" label="36 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeYi" label="26 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoYi" label="24 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeEr" label="22 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoEr" label="18 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeSan" label="16 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoSan" label="15 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeSi" label="14 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoSi" label="12 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeWu" label="10.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoWu" label="9 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeLiu" label="7.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeXiaoLiu" label="6.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeQi" label="5.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.FontSizeBa" label="5 pt (CN preset)" />
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
          <button id="VisualTeX.WordVsto.FontSizeIncrease" label="Increase" imageMso="FontSizeIncrease" getEnabled="GetFormulaFontSizeEnabled" onAction="OnIncreaseFormulaFontSize" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";
}

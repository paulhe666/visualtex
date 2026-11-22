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
          <menu id="VisualTeX.WordVsto.BatchNumbering" label="Number Equations" size="large" screentip="Number display equations in bulk" supertip="Give every unnumbered display equation (VisualTeX OLE, Word OMML, MathType) in the selection or the whole document its own native number in the document's number format. Existing numbers are kept and the following ones follow on; cross-references are updated, and one Ctrl+Z undoes it." tag="updateNumbers" getImage="GetRibbonImage">
            <button id="VisualTeX.WordVsto.BatchNumberingSelection" label="Number Display Equations in Selection" onAction="OnNumberDisplayFormulasSelection" />
            <button id="VisualTeX.WordVsto.BatchNumberingDocument" label="Number All Display Equations" onAction="OnNumberDisplayFormulasDocument" />
          </menu>
          <button id="VisualTeX.WordVsto.UpdateEquationFormat" label="Update Equation Format" size="large" screentip="Update equation format" supertip="Bring every equation number an earlier VisualTeX wrote in this document into the current format (the same as VisualTeX for Mac), make every OMML equation use the document math font (the OMML Font choice), and draw the VisualTeX equations drawn in other fonts again in the current fonts, keeping their numbers, references and sizes. Equation content and kinds do not change." tag="updateNumbers" getImage="GetRibbonImage" onAction="OnUpdateEquationFormat" />
          <button id="VisualTeX.WordVsto.BulkImport" label="Bulk Import" size="large" screentip="Import LaTeX / Markdown" supertip="Parse Markdown or LaTeX into native Word text and independently editable inline/display equations." tag="batchImport" getImage="GetRibbonImage" onAction="OnBulkImport" />
        </group>
        <group id="VisualTeX.WordVsto.FormatConversionGroup" label="Format Conversion">
          <menu id="VisualTeX.WordVsto.ConvertFormulas" label="Convert Equations" size="large" screentip="Convert equation format" supertip="Turn every equation of another format (VisualTeX OLE, Word OMML, MathType) in the range into the chosen format; equations already in it stay as they are. The same equation gives the same result whatever its source; sizes, numbers and references follow this document's settings, and it can be undone." tag="batchImport" getImage="GetRibbonImage">
            <menuSeparator id="VisualTeX.WordVsto.ConvertSelectionHeader" title="Selected equations" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToOmml" label="To Word OMML" tag="wordOmml|selection" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToMathType" label="To MathType" tag="mathTypeOle|selection" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertSelectionToVisualTeX" label="To VisualTeX OLE" tag="nativeOle|selection" onAction="OnConvertFormulasTo" />
            <menuSeparator id="VisualTeX.WordVsto.ConvertDocumentHeader" title="All equations" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToOmml" label="All to Word OMML" tag="wordOmml|document" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToMathType" label="All to MathType" tag="mathTypeOle|document" onAction="OnConvertFormulasTo" />
            <button id="VisualTeX.WordVsto.ConvertDocumentToVisualTeX" label="All to VisualTeX OLE" tag="nativeOle|document" onAction="OnConvertFormulasTo" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.RedrawGroup" label="LaTeX Redraw">
          <menu id="VisualTeX.WordVsto.RedrawSelection" label="Redraw Selection" size="large" screentip="Redraw selected equations or LaTeX" supertip="Redraw selected LaTeX in place as VisualTeX OLE, Word OMML or MathType, or restore any of those three equation types to LaTeX source." tag="batchImport" getImage="GetRibbonImage">
            <button id="VisualTeX.WordVsto.RedrawSelectionOmml" label="LaTeX → Word OMML" screentip="Replace in place with a native Word equation" onAction="OnRedrawSelectionToOmml" />
            <button id="VisualTeX.WordVsto.RedrawSelectionOle" label="LaTeX → VisualTeX OLE" screentip="Replace in place with a double-click editable VisualTeX OLE" onAction="OnRedrawSelectionToOle" />
            <button id="VisualTeX.WordVsto.RedrawSelectionMathType" label="LaTeX → MathType" screentip="Replace in place with a MathType Equation.DSMT4 object" onAction="OnRedrawSelectionToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawSelectionSeparator" />
            <button id="VisualTeX.WordVsto.FormulasToLatexSelection" label="Selected Equations to LaTeX Code" screentip="Turn the selected equations back into LaTeX code" supertip="Turn the equations in the selection (VisualTeX OLE, Word OMML or MathType) back into LaTeX code in place." tag="selection" onAction="OnFormulasToLatex" />
          </menu>
          <menu id="VisualTeX.WordVsto.RedrawDocument" label="Redraw Entire Document" size="large" screentip="Redraw equations or LaTeX in the entire document" supertip="Redraw all LaTeX in place as VisualTeX OLE, Word OMML or MathType, or restore all three equation types to LaTeX source. A confirmation is shown before changes begin." imageMso="RefreshAll">
            <button id="VisualTeX.WordVsto.RedrawDocumentOmml" label="All LaTeX → Word OMML" onAction="OnRedrawDocumentToOmml" />
            <button id="VisualTeX.WordVsto.RedrawDocumentOle" label="All LaTeX → VisualTeX OLE" onAction="OnRedrawDocumentToOle" />
            <button id="VisualTeX.WordVsto.RedrawDocumentMathType" label="All LaTeX → MathType" onAction="OnRedrawDocumentToMathType" />
            <menuSeparator id="VisualTeX.WordVsto.RedrawDocumentSeparator" />
            <button id="VisualTeX.WordVsto.FormulasToLatexDocument" label="All Equations to LaTeX Code" tag="document" onAction="OnFormulasToLatex" />
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
          <comboBox id="VisualTeX.WordVsto.InlineFontSize" label="Inline" sizeString="Follow Text" screentip="Inline formula size" supertip="Once chosen, every inline formula of the current document takes this size at once (undoable), and new inline formulas (inserted, or written by LaTeX redraw, bulk import and format conversion) get it too. Follow Text gives every inline formula the size of the text beside it, and new ones follow the text." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
            <item id="VisualTeX.WordVsto.InlineFontSizeFollow" label="Follow Text" />
            <item id="VisualTeX.WordVsto.InlineFontSizeChu" label="42 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoChu" label="36 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeYi" label="26 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoYi" label="24 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeEr" label="22 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoEr" label="18 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeSan" label="16 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoSan" label="15 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeSi" label="14 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoSi" label="12 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeWu" label="10.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoWu" label="9 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeLiu" label="7.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeXiaoLiu" label="6.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeQi" label="5.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSizeBa" label="5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.InlineFontSize8" label="8" />
            <item id="VisualTeX.WordVsto.InlineFontSize9" label="9" />
            <item id="VisualTeX.WordVsto.InlineFontSize10" label="10" />
            <item id="VisualTeX.WordVsto.InlineFontSize10_5" label="10.5" />
            <item id="VisualTeX.WordVsto.InlineFontSize11" label="11" />
            <item id="VisualTeX.WordVsto.InlineFontSize12" label="12" />
            <item id="VisualTeX.WordVsto.InlineFontSize14" label="14" />
            <item id="VisualTeX.WordVsto.InlineFontSize16" label="16" />
            <item id="VisualTeX.WordVsto.InlineFontSize18" label="18" />
            <item id="VisualTeX.WordVsto.InlineFontSize20" label="20" />
            <item id="VisualTeX.WordVsto.InlineFontSize24" label="24" />
            <item id="VisualTeX.WordVsto.InlineFontSize28" label="28" />
            <item id="VisualTeX.WordVsto.InlineFontSize36" label="36" />
            <item id="VisualTeX.WordVsto.InlineFontSize48" label="48" />
            <item id="VisualTeX.WordVsto.InlineFontSize72" label="72" />
          </comboBox>
          <comboBox id="VisualTeX.WordVsto.DisplayFontSize" label="Display" sizeString="Follow Text" screentip="Display formula size" supertip="Once chosen, every display formula of the current document (numbered ones included, their numbers with them) takes this size at once (undoable), and new display formulas (inserted, or written by LaTeX redraw, bulk import and format conversion) get it too. Follow Text gives every display formula the size of the text before it, and new ones follow the text." getText="GetFormulaKindFontSizeText" getEnabled="GetFormulaKindFontSizeEnabled" onChange="OnFormulaKindFontSizeChanged">
            <item id="VisualTeX.WordVsto.DisplayFontSizeFollow" label="Follow Text" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeChu" label="42 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoChu" label="36 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeYi" label="26 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoYi" label="24 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeEr" label="22 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoEr" label="18 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeSan" label="16 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoSan" label="15 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeSi" label="14 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoSi" label="12 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeWu" label="10.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoWu" label="9 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeLiu" label="7.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeXiaoLiu" label="6.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeQi" label="5.5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSizeBa" label="5 pt (CN preset)" />
            <item id="VisualTeX.WordVsto.DisplayFontSize8" label="8" />
            <item id="VisualTeX.WordVsto.DisplayFontSize9" label="9" />
            <item id="VisualTeX.WordVsto.DisplayFontSize10" label="10" />
            <item id="VisualTeX.WordVsto.DisplayFontSize10_5" label="10.5" />
            <item id="VisualTeX.WordVsto.DisplayFontSize11" label="11" />
            <item id="VisualTeX.WordVsto.DisplayFontSize12" label="12" />
            <item id="VisualTeX.WordVsto.DisplayFontSize14" label="14" />
            <item id="VisualTeX.WordVsto.DisplayFontSize16" label="16" />
            <item id="VisualTeX.WordVsto.DisplayFontSize18" label="18" />
            <item id="VisualTeX.WordVsto.DisplayFontSize20" label="20" />
            <item id="VisualTeX.WordVsto.DisplayFontSize24" label="24" />
            <item id="VisualTeX.WordVsto.DisplayFontSize28" label="28" />
            <item id="VisualTeX.WordVsto.DisplayFontSize36" label="36" />
            <item id="VisualTeX.WordVsto.DisplayFontSize48" label="48" />
            <item id="VisualTeX.WordVsto.DisplayFontSize72" label="72" />
          </comboBox>
          <menu id="VisualTeX.WordVsto.OmmlMathFont" label="OMML Font" imageMso="FontDialog" tag="menu" screentip="OMML equation font" supertip="The math font of native Word equations (OMML). With a font chosen, VisualTeX makes it the document math font before writing OMML and brings the existing equations of the current document in line at once; Update Equation Format does the same. WPS has no document math font: the chosen font is used for the equations VisualTeX inserts or edits." getEnabled="GetOmmlMathFontEnabled">
            <toggleButton id="VisualTeX.WordVsto.OmmlMathFontFollow" label="Follow Document (No Change)" tag="follow" screentip="Follow document" supertip="Leave the document math font unchanged (default)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
            <toggleButton id="VisualTeX.WordVsto.OmmlMathFontCambria" label="Cambria Math" tag="cambria" screentip="Cambria Math" supertip="Word's default math font." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
            <toggleButton id="VisualTeX.WordVsto.OmmlMathFontStix" getLabel="GetOmmlMathFontLabel" tag="stix" screentip="STIX Two Math (bundled)" supertip="The STIX Two Math that ships with VisualTeX (the official TrueType build, Times-style, embedded by Save as PDF)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
            <toggleButton id="VisualTeX.WordVsto.OmmlMathFontTermes" getLabel="GetOmmlMathFontLabel" tag="termes" screentip="TeX Gyre Termes Math (bundled)" supertip="The TeX Gyre Termes Math that ships with VisualTeX (Times-style, close to LaTeX's newtxmath; TrueType build, embedded by Save as PDF)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
            <toggleButton id="VisualTeX.WordVsto.OmmlMathFontLatinModern" getLabel="GetOmmlMathFontLabel" tag="latin-modern" screentip="Latin Modern Math (bundled)" supertip="The Latin Modern Math that ships with VisualTeX (TrueType build, embedded by Save as PDF)." getEnabled="GetOmmlMathFontEnabled" getPressed="GetOmmlMathFontPressed" onAction="OnOmmlMathFontChanged" />
          </menu>
        </group>
        <group id="VisualTeX.WordVsto.CompatibilityRepairGroup" label="Compatibility Repair">
          <button id="VisualTeX.WordVsto.RepairInlineBaselines" label="Repair Formula Baselines" size="large" screentip="Scan and repair inline equation offsets" supertip="Scan inline MathType / VisualTeX OLE equations across the document. VisualTeX formulas are calibrated from their own saved render baseline; MathType formulas use a stable document baseline profile. After confirmation, only vertical offsets are repaired; formula content, font size, numbering and references are not changed." tag="repairBaseline" getImage="GetRibbonImage" onAction="OnRepairInlineFormulaBaselines" />
          <button id="VisualTeX.WordVsto.RepairMathTypeDisplays" label="Repair Formula Display" size="large" screentip="Redraw equation pictures with MathType's own layout" supertip="Scan the MathType equations in the document and redraw their display pictures from their stored data with VisualTeX's built-in MathType renderer. This fixes wrong symbols drawn while MathType's fonts were missing. Only the pictures are replaced and realigned on MathType's baseline; equation data, numbering and references are not changed." tag="repairBaseline" getImage="GetRibbonImage" onAction="OnRepairMathTypeDisplays" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
""";
}

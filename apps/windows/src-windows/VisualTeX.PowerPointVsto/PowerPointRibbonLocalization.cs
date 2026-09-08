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
          <button id="VisualTeX.PowerPointVsto.New" label="New Formula" size="large" tag="insertFormula" getImage="GetRibbonImage" onAction="OnNewFormula" />
          <button id="VisualTeX.PowerPointVsto.Edit" label="Edit Selected Formula" size="large" tag="editSelected" getImage="GetRibbonImage" onAction="OnEditSelected" />
          <button id="VisualTeX.PowerPointVsto.ConvertOmml" label="Convert to OMML" screentip="Convert to a native PowerPoint equation" supertip="Convert to native Office Math stored as OMML inside the PPTX so it can be edited with PowerPoint equation tools." imageMso="EquationInsertNew" onAction="OnConvertSelectedOmml" />
          <button id="VisualTeX.PowerPointVsto.ConvertSelected" label="Convert to Native OLE" screentip="Convert to an embedded editable OLE object" supertip="Preserve the visual appearance while embedding the object in the PowerPoint file so it can be edited again with VisualTeX." tag="convertToOle" getImage="GetRibbonImage" onAction="OnConvertSelected" />
          <button id="VisualTeX.PowerPointVsto.ExportPicture" label="Convert to SVG" imageMso="PictureInsertFromFile" onAction="OnExportSelectedAsPicture" />
          <button id="VisualTeX.PowerPointVsto.Delete" label="Delete Selected Formula" imageMso="Delete" onAction="OnDeleteSelected" />
          <button id="VisualTeX.PowerPointVsto.OpenDesktop" label="Open VisualTeX" imageMso="FileOpen" onAction="OnOpenDesktop" />
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

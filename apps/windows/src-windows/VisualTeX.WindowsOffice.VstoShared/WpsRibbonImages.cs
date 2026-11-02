using System.Xml.Linq;

namespace VisualTeX.WindowsOffice.VstoShared;

/// <summary>
/// Ribbon images in WPS Office. WPS draws an add-in's own pictures (getImage) in a fixed 16 px
/// box, also on large buttons and at any display scale, and without their transparency; its
/// built-in icons (imageMso) are drawn at full size in the style of the current WPS skin, but
/// only for the Office ids WPS knows (others show a "?" box). In WPS each control below gets the
/// listed WPS-known icon instead of VisualTeX's own picture or an id WPS lacks.
/// </summary>
internal static class WpsRibbonImages
{
    // Checked in WPS Office 12.1.0.28488 (2026-10-27). Unknown there: EquationInsertNew,
    // EquationBracketGallery, FontSizeIncrease/Decrease(Word), FontDialog, NumberingRestart,
    // UpdateFields, TableOfContentsUpdate, PictureReset/Change, ShapeChangeShapeGallery.
    private static readonly Dictionary<string, string> IconByControl = new(StringComparer.Ordinal)
    {
        ["VisualTeX.WordVsto.Inline"] = "EquationScriptGallery",
        ["VisualTeX.WordVsto.Display"] = "EquationFractionGallery",
        ["VisualTeX.WordVsto.InlineOmml"] = "EquationProfessional",
        ["VisualTeX.WordVsto.DisplayOmml"] = "EquationInsertGallery",
        ["VisualTeX.WordVsto.Edit"] = "ObjectEditDialog",
        ["VisualTeX.WordVsto.UpdateNumbers"] = "RecurrenceEdit",
        ["VisualTeX.WordVsto.BatchNumbering"] = "Numbering",
        ["VisualTeX.WordVsto.UpdateEquationFormat"] = "FormatPainter",
        ["VisualTeX.WordVsto.BulkImport"] = "TextFromFileInsert",
        ["VisualTeX.WordVsto.RedrawSelection"] = "Refresh",
        ["VisualTeX.WordVsto.FontSizeDecrease"] = "ShrinkFont",
        ["VisualTeX.WordVsto.FontSizeIncrease"] = "GrowFont",
        ["VisualTeX.WordVsto.OmmlMathFont"] = "EquationOptions",
        ["VisualTeX.WordVsto.RepairInlineBaselines"] = "ObjectsAlignMiddleVertical",
        ["VisualTeX.WordVsto.RepairMathTypeDisplays"] = "ReviewDisplayForReview",
        ["VisualTeX.PowerPointVsto.New"] = "PictureInsertFromFile",
        ["VisualTeX.PowerPointVsto.NewOle"] = "OleObjectctInsert",
        ["VisualTeX.PowerPointVsto.NewMathType"] = "EquationOptions",
        ["VisualTeX.PowerPointVsto.InsertInlineOmml"] = "EquationProfessional",
        ["VisualTeX.PowerPointVsto.InsertDisplayOmml"] = "EquationInsertGallery",
        ["VisualTeX.PowerPointVsto.Edit"] = "ObjectEditDialog",
        ["VisualTeX.PowerPointVsto.UnifyFormat"] = "FormatPainter",
        ["VisualTeX.PowerPointVsto.VisualTeXTargetForm"] = "ConvertTextToTable",
        // PowerPoint shows TextBoxInsert, an id WPS was not checked for.
        ["VisualTeX.PowerPointVsto.FormulaToLatex"] = "Refresh",
        ["VisualTeX.PowerPointVsto.FontSizeDecrease"] = "ShrinkFont",
        ["VisualTeX.PowerPointVsto.FontSizeIncrease"] = "GrowFont",
    };

    /// <summary>The ribbon XML for the current host: unchanged outside WPS.</summary>
    internal static string ForHost(string ribbonXml) =>
        OfficeHostEnvironment.IsWps ? UseWpsIcons(ribbonXml) : ribbonXml;

    internal static string UseWpsIcons(string ribbonXml)
    {
        try
        {
            var document = XDocument.Parse(ribbonXml, LoadOptions.PreserveWhitespace);
            var changed = false;
            foreach (var element in document.Descendants())
            {
                var id = (string?)element.Attribute("id");
                if (id is null || !IconByControl.TryGetValue(id, out var icon)) continue;
                element.Attribute("getImage")?.Remove();
                element.Attribute("image")?.Remove();
                element.SetAttributeValue("imageMso", icon);
                changed = true;
            }
            return changed ? document.ToString(SaveOptions.DisableFormatting) : ribbonXml;
        }
        catch
        {
            return ribbonXml;
        }
    }
}

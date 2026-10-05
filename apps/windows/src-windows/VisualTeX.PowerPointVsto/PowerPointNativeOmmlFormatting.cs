using System.Globalization;
using System.Xml.Linq;

namespace VisualTeX.PowerPointVsto;

// Bridge the host's DrawingML run colors at the PPT boundary. The original
// package XML remains untouched and is retained for stale-source checks.
internal static class PowerPointNativeOmmlFormatting
{
    internal static string ForWordReverseTransform(string omml)
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var root = XElement.Parse(omml, LoadOptions.PreserveWhitespace);
        foreach (var owner in root.DescendantsAndSelf().Where(e => e.Name == m + "r" || e.Name == m + "ctrlPr"))
        {
            var drawing = owner.Element(a + "rPr");
            var rgb = drawing?.Element(a + "solidFill")?.Element(a + "srgbClr");
            var value = (string?)rgb?.Attribute("val");
            // Do not reinterpret theme colors, alpha or drawing effects as plain
            // RGB. Their rendering needs resolved presentation/theme context.
            if (rgb is null || rgb.HasElements || value?.Length != 6
                || !uint.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _))
                continue;
            var properties = owner.Element(w + "rPr");
            if (properties is null) { properties = new XElement(w + "rPr"); owner.AddFirst(properties); }
            properties.Elements(w + "color").Remove();
            properties.Add(new XElement(w + "color", new XAttribute(w + "val", value!.ToUpperInvariant())));
        }
        return root.ToString(SaveOptions.DisableFormatting);
    }
}

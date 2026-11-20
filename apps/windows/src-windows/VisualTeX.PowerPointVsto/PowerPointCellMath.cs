using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.PowerPointVsto;

// A table cell takes an equation only through PowerPoint's linear format, which keeps
// the characters of an equation but not their formatting: a letter's style must be in
// the character itself. Upright letters become normal text (quoted in linear format)
// and bold, script, fraktur, double-struck, sans-serif and monospace letters become
// Unicode's mathematical alphanumerics. Colors are not kept.
internal static class PowerPointCellMath
{
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    internal static string Prepare(string mathMl)
    {
        XDocument document;
        try { document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace); }
        catch (XmlException error) { throw new InvalidDataException("The formula's MathML is not valid XML.", error); }
        foreach (var element in document.Descendants().Where(element => element.Name.Namespace == MathMl
                     && element.Name.LocalName is "mi" or "mn" or "mtext" or "ms").ToList())
        {
            var variant = InheritedVariant(element);
            var text = element.Value;
            if (element.Name.LocalName == "mi" && IsUpright(variant, text) && text.Any(char.IsLetter))
            {
                // A function name before U+2061 stays an identifier: PowerPoint keeps
                // functions upright itself.
                if (IsFunctionName(element)) continue;
                element.Name = MathMl + "mtext";
                element.Attribute("mathvariant")?.Remove();
                continue;
            }
            if (variant is null or "normal" or "italic" || element.Name.LocalName is "mtext" or "ms") continue;
            var styled = Styled(text, variant);
            if (styled is null) continue;
            element.Value = styled;
            element.Attribute("mathvariant")?.Remove();
            // An identifier spelled in mathematical alphanumerics is never italicized again.
            if (element.Name.LocalName == "mi" && text.Length == 1) element.SetAttributeValue("mathvariant", "normal");
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    // A multi-letter identifier is upright unless styled otherwise; a single letter only
    // when marked normal.
    private static bool IsUpright(string? variant, string text) =>
        variant == "normal"
        || variant is null && new System.Globalization.StringInfo(text.Trim()).LengthInTextElements > 1;

    private static bool IsFunctionName(XElement identifier)
    {
        var next = identifier.ElementsAfterSelf().FirstOrDefault();
        return next is not null && next.Name == MathMl + "mo" && next.Value.Trim() == "⁡";
    }

    private static string? InheritedVariant(XElement element)
    {
        for (var current = element; current is not null; current = current.Parent)
        {
            var variant = (string?)current.Attribute("mathvariant");
            if (variant is not null) return variant;
            if (current.Name == MathMl + "math") break;
        }
        return null;
    }

    // Unicode's Mathematical Alphanumeric Symbols for a MathML mathvariant, null when the
    // variant has no letters there. Letters with an earlier code point (the Letterlike
    // Symbols holes) use it.
    internal static string? Styled(string text, string variant)
    {
        int? upper = null, lower = null, digit = null;
        switch (variant)
        {
            case "bold": upper = 0x1D400; lower = 0x1D41A; digit = 0x1D7CE; break;
            case "bold-italic": upper = 0x1D468; lower = 0x1D482; digit = 0x1D7CE; break;
            case "script": upper = 0x1D49C; lower = 0x1D4B6; break;
            case "bold-script": upper = 0x1D4D0; lower = 0x1D4EA; digit = 0x1D7CE; break;
            case "fraktur": upper = 0x1D504; lower = 0x1D51E; break;
            case "double-struck": upper = 0x1D538; lower = 0x1D552; digit = 0x1D7D8; break;
            case "bold-fraktur": upper = 0x1D56C; lower = 0x1D586; digit = 0x1D7CE; break;
            case "sans-serif": upper = 0x1D5A0; lower = 0x1D5BA; digit = 0x1D7E2; break;
            case "bold-sans-serif": upper = 0x1D5D4; lower = 0x1D5EE; digit = 0x1D7EC; break;
            case "sans-serif-italic": upper = 0x1D608; lower = 0x1D622; digit = 0x1D7E2; break;
            case "sans-serif-bold-italic": upper = 0x1D63C; lower = 0x1D656; digit = 0x1D7EC; break;
            case "monospace": upper = 0x1D670; lower = 0x1D68A; digit = 0x1D7F6; break;
            default: return null;
        }
        var result = new StringBuilder();
        foreach (var character in text)
        {
            int? code = character switch
            {
                >= 'A' and <= 'Z' when upper is not null => upper + (character - 'A'),
                >= 'a' and <= 'z' when lower is not null => lower + (character - 'a'),
                >= '0' and <= '9' when digit is not null => digit + (character - '0'),
                _ => null,
            };
            if (code is null) { result.Append(character); continue; }
            result.Append(Holes.TryGetValue(code.Value, out var hole) ? hole : char.ConvertFromUtf32(code.Value));
        }
        return result.ToString();
    }

    private static readonly Dictionary<int, string> Holes = new()
    {
        [0x1D49D] = "ℬ", [0x1D4A0] = "ℰ", [0x1D4A1] = "ℱ", [0x1D4A3] = "ℋ", [0x1D4A4] = "ℐ",
        [0x1D4A7] = "ℒ", [0x1D4A8] = "ℳ", [0x1D4AD] = "ℛ", [0x1D4BA] = "ℯ", [0x1D4BC] = "ℊ",
        [0x1D4C4] = "ℴ", [0x1D506] = "ℭ", [0x1D50B] = "ℌ", [0x1D50C] = "ℑ", [0x1D515] = "ℜ",
        [0x1D51D] = "ℨ", [0x1D53A] = "ℂ", [0x1D53F] = "ℍ", [0x1D545] = "ℕ", [0x1D547] = "ℙ",
        [0x1D548] = "ℚ", [0x1D549] = "ℝ", [0x1D551] = "ℤ",
    };
}

using System.Drawing;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>Opaque text colors shared by source recovery and the Office XML boundary.</summary>
internal static class MathMlColorContract
{
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";

    internal static string? Normalize(string? value)
    {
        var color = value?.Trim();
        if (string.IsNullOrEmpty(color)) return null;
        if (Regex.IsMatch(color!, "^#[0-9a-fA-F]{6}$")) return color!.ToUpperInvariant();
        if (Regex.IsMatch(color!, "^#[0-9a-fA-F]{3}$"))
            return "#" + string.Concat(color!.Substring(1).Select(c => new string(c, 2))).ToUpperInvariant();
        var named = Color.FromName(color!);
        return named.IsKnownColor && !named.IsSystemColor && named.A == 255
            ? "#" + named.R.ToString("X2") + named.G.ToString("X2") + named.B.ToString("X2") : null;
    }

    internal static string? Explicit(XElement element) => Normalize(
        (string?)element.Attribute("mathcolor") ?? (string?)element.Attribute(MathMl + "mathcolor")
        ?? (string?)element.Attribute("color") ?? (string?)element.Attribute(MathMl + "color"));

    internal static string? Effective(XElement element) =>
        element.AncestorsAndSelf().Select(Explicit).FirstOrDefault(color => color is not null);

    internal static string Wrap(string source, string? color) =>
        color is null || string.IsNullOrWhiteSpace(source) ? source : @"\textcolor{" + color + "}{" + source + "}";

    internal static string MaterializeInheritedColors(string mathMl)
    {
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        foreach (var element in document.Root!.DescendantsAndSelf())
        {
            var color = Effective(element);
            if (color is not null) element.SetAttributeValue("mathcolor", color);
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }
}

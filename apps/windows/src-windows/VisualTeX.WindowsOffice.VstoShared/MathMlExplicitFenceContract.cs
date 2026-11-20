using System;
using System.Linq;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// Recognizes explicit source fence structure independently of a native format's
/// available delimiter templates. It does not infer roles from scalar glyphs.
/// </summary>
internal static class MathMlExplicitFenceContract
{
    private static readonly XNamespace PresentationMath = "http://www.w3.org/1998/Math/MathML";

    internal static bool HasTexClass(XElement element, string expected) =>
        string.Equals(element.Attribute("data-mjx-texclass")?.Value,
            expected, StringComparison.OrdinalIgnoreCase);

    internal static bool IsProjectableEndpoint(XElement element)
    {
        if (element.Name.LocalName != "mo"
            || (element.Name.Namespace != PresentationMath && element.Name.Namespace != XNamespace.None)
            || element.HasElements)
            return false;
        // mfenced has one shared delimiter presentation, not independent token
        // styles. Consume only role and the true fence/stretch/symmetry flags.
        // Other local attributes retain the original mo nodes and their existing
        // behavior; size/style/alternate/spacing are unresolved capabilities.
        return element.Attributes().All(attribute =>
            attribute.IsNamespaceDeclaration
            || attribute.Name == "data-mjx-texclass"
            || ((attribute.Name == "fence" || attribute.Name == "stretchy"
                    || attribute.Name == "symmetric")
                && string.Equals(attribute.Value, "true", StringComparison.OrdinalIgnoreCase)));
    }

    internal static bool TryGetRow(
        XElement row, out string open, out string close, out XElement[] body)
    {
        open = string.Empty;
        close = string.Empty;
        body = Array.Empty<XElement>();
        if (row.Name.LocalName != "mrow"
            || (row.Name.Namespace != PresentationMath && row.Name.Namespace != XNamespace.None)
            || !HasTexClass(row, "INNER"))
            return false;
        var elements = row.Elements().ToArray();
        if (elements.Length < 2) return false;
        var first = elements[0];
        var last = elements[elements.Length - 1];
        if (first.Name.Namespace != row.Name.Namespace || last.Name.Namespace != row.Name.Namespace
            || !HasTexClass(first, "OPEN") || !HasTexClass(last, "CLOSE")
            || !IsProjectableEndpoint(first) || !IsProjectableEndpoint(last))
            return false;
        if (first.Value.Length == 0 && last.Value.Length == 0) return false;
        open = first.Value;
        close = last.Value;
        body = elements.Skip(1).Take(elements.Length - 2).ToArray();
        return true;
    }
}

using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// MathML token style authority before a native adapter selects fonts or merges
/// runs. A missing declaration stays distinct from the effective default.
/// </summary>
internal static class MathMlTokenStyleContract
{
    private const string MathMlNamespace = "http://www.w3.org/1998/Math/MathML";

    // Match the supported single-character contract in MathJax's MmlMi:
    // one Unicode scalar followed by its recognized combining-mark ranges.
    // Grapheme-cluster count and UTF-16 string length are not this contract.
    // XML forbids isolated surrogates; keep the predicate equally precise.
    private static readonly Regex SingleCharacter = new(
        @"\A(?:[\uD800-\uDBFF][\uDC00-\uDFFF]|[^\uD800-\uDFFF\r\n\u2028\u2029])[\u0300-\u036F\u1AB0-\u1ABE\u1DC0-\u1DFF\u20D0-\u20EF]*\z",
        RegexOptions.CultureInvariant);

    internal static string? ResolveDeclaredVariant(XElement element)
    {
        foreach (var current in element.AncestorsAndSelf())
        {
            // Native MTEF readback is unqualified; the Office boundary is
            // namespaced. Neither may inherit style from foreign XML.
            if (!IsMathMlElement(current)) break;
            var declared = ((string?)current.Attribute("mathvariant"))?.Trim();
            if (!string.IsNullOrEmpty(declared))
                return declared!.ToLowerInvariant();
            if (current.Name.LocalName == "math") break;
        }
        return null;
    }

    internal static string? ResolveEffectiveVariant(XElement token) =>
        ResolveEffectiveVariant(token, ResolveDeclaredVariant(token));

    internal static string? ResolveEffectiveVariant(XElement token, string? declaredVariant)
    {
        if (!IsMathMlElement(token)
            || token.Name.LocalName is not ("mi" or "mo" or "mn" or "mtext" or "ms"))
            return null;
        if (!string.IsNullOrWhiteSpace(declaredVariant))
            return declaredVariant!.Trim().ToLowerInvariant();
        // Embedded glyphs/foreign content are outside this text-token contract.
        // Do not invent an italic/normal result from their flattened Value.
        if (token.HasElements) return null;
        return token.Name.LocalName == "mi" && IsSingleMathMlCharacter(token.Value.Trim())
            ? "italic"
            : "normal";
    }

    internal static bool IsSingleMathMlCharacter(string text) =>
        SingleCharacter.IsMatch(text);

    private static bool IsMathMlElement(XElement element) =>
        element.Name.NamespaceName.Length == 0
        || element.Name.NamespaceName == MathMlNamespace;
}

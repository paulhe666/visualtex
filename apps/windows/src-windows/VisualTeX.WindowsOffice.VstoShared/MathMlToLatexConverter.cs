using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

internal static class MathMlToLatexConverter
{
    private static readonly IReadOnlyDictionary<string, string> TokenMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["−"] = "-",
            ["±"] = @"\pm ",
            ["∓"] = @"\mp ",
            ["×"] = @"\times ",
            ["·"] = @"\cdot ",
            ["÷"] = @"\div ",
            ["∞"] = @"\infty ",
            ["∫"] = @"\int ",
            ["∬"] = @"\iint ",
            ["∭"] = @"\iiint ",
            ["∮"] = @"\oint ",
            ["∑"] = @"\sum ",
            ["∏"] = @"\prod ",
            ["∐"] = @"\coprod ",
            ["∂"] = @"\partial ",
            ["∇"] = @"\nabla ",
            ["√"] = @"\sqrt{}",
            ["≠"] = @"\ne ",
            ["≈"] = @"\approx ",
            ["≡"] = @"\equiv ",
            ["≤"] = @"\le ",
            ["≥"] = @"\ge ",
            ["≪"] = @"\ll ",
            ["≫"] = @"\gg ",
            ["∝"] = @"\propto ",
            ["∈"] = @"\in ",
            ["∉"] = @"\notin ",
            ["∋"] = @"\ni ",
            ["⊂"] = @"\subset ",
            ["⊆"] = @"\subseteq ",
            ["⊃"] = @"\supset ",
            ["⊇"] = @"\supseteq ",
            ["∪"] = @"\cup ",
            ["∩"] = @"\cap ",
            ["∅"] = @"\varnothing ",
            ["∧"] = @"\land ",
            ["∨"] = @"\lor ",
            ["¬"] = @"\neg ",
            ["⇒"] = @"\Rightarrow ",
            ["⇔"] = @"\Leftrightarrow ",
            ["→"] = @"\to ",
            ["←"] = @"\leftarrow ",
            ["↔"] = @"\leftrightarrow ",
            ["↦"] = @"\mapsto ",
            ["⟂"] = @"\perp ",
            ["∥"] = @"\parallel ",
            ["…"] = @"\dots ",
            ["⋯"] = @"\cdots ",
            ["⋮"] = @"\vdots ",
            ["⋱"] = @"\ddots ",
            ["α"] = @"\alpha ",
            ["β"] = @"\beta ",
            ["γ"] = @"\gamma ",
            ["δ"] = @"\delta ",
            ["ε"] = @"\epsilon ",
            ["ϵ"] = @"\varepsilon ",
            ["ζ"] = @"\zeta ",
            ["η"] = @"\eta ",
            ["θ"] = @"\theta ",
            ["ϑ"] = @"\vartheta ",
            ["ι"] = @"\iota ",
            ["κ"] = @"\kappa ",
            ["λ"] = @"\lambda ",
            ["μ"] = @"\mu ",
            ["ν"] = @"\nu ",
            ["ξ"] = @"\xi ",
            ["π"] = @"\pi ",
            ["ϖ"] = @"\varpi ",
            ["ρ"] = @"\rho ",
            ["ϱ"] = @"\varrho ",
            ["σ"] = @"\sigma ",
            ["ς"] = @"\varsigma ",
            ["τ"] = @"\tau ",
            ["υ"] = @"\upsilon ",
            ["φ"] = @"\phi ",
            ["ϕ"] = @"\varphi ",
            ["χ"] = @"\chi ",
            ["ψ"] = @"\psi ",
            ["ω"] = @"\omega ",
            ["Γ"] = @"\Gamma ",
            ["Δ"] = @"\Delta ",
            ["Θ"] = @"\Theta ",
            ["Λ"] = @"\Lambda ",
            ["Ξ"] = @"\Xi ",
            ["Π"] = @"\Pi ",
            ["Σ"] = @"\Sigma ",
            ["Υ"] = @"\Upsilon ",
            ["Φ"] = @"\Phi ",
            ["Ψ"] = @"\Psi ",
            ["Ω"] = @"\Omega ",
        };

    internal static string Convert(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("Word OMML conversion produced empty MathML.");
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.None);
        var root = document.Root
            ?? throw new InvalidDataException("Word MathML has no root element.");
        return Normalize(ConvertElement(root));
    }

    private static string ConvertElement(XElement element)
    {
        var name = element.Name.LocalName;
        return name switch
        {
            "math" or "mrow" or "mstyle" or "mpadded" => ConvertChildren(element),
            "semantics" or "maction" => element.Elements().Select(ConvertElement).FirstOrDefault() ?? string.Empty,
            "annotation" or "annotation-xml" => string.Empty,
            "mi" or "mn" or "mo" => ConvertToken(element.Value),
            "mtext" => @"\text{" + EscapeText(element.Value) + "}",
            "mspace" => @"\,",
            "mfrac" => ConvertFraction(element),
            "msqrt" => @"\sqrt{" + ConvertChildren(element) + "}",
            "mroot" => ConvertRoot(element),
            "msup" => ConvertScript(element, subscript: false, superscript: true),
            "msub" => ConvertScript(element, subscript: true, superscript: false),
            "msubsup" => ConvertSubSup(element),
            "mover" => ConvertOver(element),
            "munder" => ConvertUnder(element),
            "munderover" => ConvertUnderOver(element),
            "mfenced" => ConvertFenced(element),
            "mtable" => ConvertTable(element),
            "mtr" or "mlabeledtr" => string.Join(" & ", element.Elements().Select(ConvertElement)),
            "mtd" => ConvertChildren(element),
            "menclose" => ConvertEnclose(element),
            "mphantom" => @"\phantom{" + ConvertChildren(element) + "}",
            "mmultiscripts" => ConvertMultiScripts(element),
            "none" => string.Empty,
            _ => ConvertChildren(element),
        };
    }

    private static string ConvertChildren(XElement element) =>
        string.Concat(element.Elements().Select(ConvertElement));

    private static string ConvertFraction(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        return @"\frac{" + ConvertElement(children[0]) + "}{" + ConvertElement(children[1]) + "}";
    }

    private static string ConvertRoot(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return @"\sqrt{" + ConvertChildren(element) + "}";
        return @"\sqrt[" + ConvertElement(children[1]) + "]{" + ConvertElement(children[0]) + "}";
    }

    private static string ConvertScript(XElement element, bool subscript, bool superscript)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        var result = GroupBase(ConvertElement(children[0]));
        if (subscript) result += "_{" + ConvertElement(children[1]) + "}";
        if (superscript) result += "^{" + ConvertElement(children[1]) + "}";
        return result;
    }

    private static string ConvertSubSup(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 3) return ConvertChildren(element);
        return GroupBase(ConvertElement(children[0]))
            + "_{" + ConvertElement(children[1]) + "}"
            + "^{" + ConvertElement(children[2]) + "}";
    }

    private static string ConvertOver(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        var body = ConvertElement(children[0]);
        var over = children[1].Value.Trim();
        return over switch
        {
            "¯" or "‾" => @"\overline{" + body + "}",
            "→" => @"\vec{" + body + "}",
            "^" or "ˆ" => @"\hat{" + body + "}",
            "~" or "˜" => @"\tilde{" + body + "}",
            "." or "˙" => @"\dot{" + body + "}",
            "¨" => @"\ddot{" + body + "}",
            _ => @"\overset{" + ConvertElement(children[1]) + "}{" + body + "}",
        };
    }

    private static string ConvertUnder(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        return @"\underset{" + ConvertElement(children[1]) + "}{" + ConvertElement(children[0]) + "}";
    }

    private static string ConvertUnderOver(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 3) return ConvertChildren(element);
        return GroupBase(ConvertElement(children[0]))
            + "_{" + ConvertElement(children[1]) + "}"
            + "^{" + ConvertElement(children[2]) + "}";
    }

    private static string ConvertFenced(XElement element)
    {
        var open = (string?)element.Attribute("open") ?? "(";
        var close = (string?)element.Attribute("close") ?? ")";
        var separators = (string?)element.Attribute("separators") ?? ",";
        var parts = element.Elements().Select(ConvertElement).ToList();
        var separator = separators.Length > 0 ? separators[0].ToString() : ",";
        return Delimiter(open, left: true)
            + string.Join(separator, parts)
            + Delimiter(close, left: false);
    }

    private static string ConvertTable(XElement element)
    {
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .Select(row => string.Join(" & ", row.Elements().Select(ConvertElement)))
            .ToList();
        return @"\begin{matrix}" + string.Join(@" \\ ", rows) + @"\end{matrix}";
    }

    private static string ConvertEnclose(XElement element)
    {
        var notation = ((string?)element.Attribute("notation") ?? string.Empty).ToLowerInvariant();
        var body = ConvertChildren(element);
        if (notation.Contains("box")) return @"\boxed{" + body + "}";
        if (notation.Contains("radical")) return @"\sqrt{" + body + "}";
        return body;
    }

    private static string ConvertMultiScripts(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count == 0) return string.Empty;
        var builder = new StringBuilder(GroupBase(ConvertElement(children[0])));
        var index = 1;
        while (index < children.Count && children[index].Name.LocalName != "mprescripts")
        {
            var sub = ConvertElement(children[index]);
            var sup = index + 1 < children.Count ? ConvertElement(children[index + 1]) : string.Empty;
            if (!string.IsNullOrEmpty(sub)) builder.Append("_{").Append(sub).Append('}');
            if (!string.IsNullOrEmpty(sup)) builder.Append("^{").Append(sup).Append('}');
            index += 2;
        }
        return builder.ToString();
    }

    private static string ConvertToken(string value)
    {
        var token = value.Trim();
        if (TokenMap.TryGetValue(token, out var mapped)) return mapped;
        if (token.Length == 0) return string.Empty;
        return token switch
        {
            "{" => @"\{",
            "}" => @"\}",
            "#" => @"\#",
            "%" => @"\%",
            "&" => @"\&",
            "_" => @"\_",
            _ => token,
        };
    }

    private static string Delimiter(string value, bool left)
    {
        var escaped = value switch
        {
            "{" => @"\{",
            "}" => @"\}",
            "|" => @"\lvert",
            "‖" => @"\lVert",
            "〈" or "⟨" => @"\langle",
            "〉" or "⟩" => @"\rangle",
            _ => value,
        };
        return (left ? @"\left" : @"\right") + escaped + " ";
    }

    private static string GroupBase(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.Length == 1 || trimmed.StartsWith("\\", StringComparison.Ordinal)) return trimmed;
        return "{" + trimmed + "}";
    }

    private static string EscapeText(string value) =>
        value.Replace("\\", @"\textbackslash{}")
            .Replace("{", @"\{")
            .Replace("}", @"\}")
            .Replace("#", @"\#")
            .Replace("%", @"\%")
            .Replace("&", @"\&")
            .Replace("_", @"\_");

    private static string Normalize(string value)
    {
        var result = value.Trim();
        while (result.IndexOf("  ", StringComparison.Ordinal) >= 0)
            result = result.Replace("  ", " ");
        return result;
    }
}

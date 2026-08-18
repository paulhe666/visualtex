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
            ["ℏ"] = @"\hbar ",
            ["∫"] = @"\int ",
            ["∬"] = @"\iint ",
            ["∭"] = @"\iiint ",
            ["∮"] = @"\oint ",
            ["∑"] = @"\sum ",
            ["∏"] = @"\prod ",
            ["∐"] = @"\coprod ",
            ["⋃"] = @"\bigcup ",
            ["⋂"] = @"\bigcap ",
            ["∗"] = @"\ast ",
            ["†"] = @"\dagger ",
            ["‡"] = @"\ddagger ",
            ["′"] = @"\prime ",
            ["″"] = @"\prime\prime ",
            ["∀"] = @"\forall ",
            ["∃"] = @"\exists ",
            ["∣"] = @"\mid ",
            ["〈"] = @"\langle ",
            ["〈"] = @"\langle ",
            ["⟨"] = @"\langle ",
            ["〉"] = @"\rangle ",
            ["〉"] = @"\rangle ",
            ["⟩"] = @"\rangle ",
            ["∂"] = @"\partial ",
            ["∇"] = @"\nabla ",
            ["⁡"] = string.Empty,
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
            // MathJax follows TeX glyph naming rather than Unicode's character
            // names: \epsilon is U+03F5 (ϵ) while \varepsilon is U+03B5 (ε).
            // Keep this mapping inverse-exact so MathType -> VisualTeX ->
            // MathType cannot silently swap the two epsilon forms.
            ["ε"] = @"\varepsilon ",
            ["ϵ"] = @"\epsilon ",
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
            // Likewise TeX \phi is U+03D5 (ϕ) and \varphi is U+03C6 (φ).
            ["φ"] = @"\varphi ",
            ["ϕ"] = @"\phi ",
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
            "mi" or "mn" or "mo" => ConvertTokenElement(element),
            "mtext" => ConvertTextElement(element),
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
        if (subscript) result += "_{" + ScriptArgument(children[1]) + "}";
        if (superscript) result += "^{" + ScriptArgument(children[1]) + "}";
        return result;
    }

    private static string ConvertSubSup(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 3) return ConvertChildren(element);
        return GroupBase(ConvertElement(children[0]))
            + "_{" + ScriptArgument(children[1]) + "}"
            + "^{" + ScriptArgument(children[2]) + "}";
    }

    private static string ScriptArgument(XElement element) =>
        ConvertElement(element).TrimEnd();

    private static string ConvertOver(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        if (TryConvertAnnotatedHorizontalBrace(children[0], children[1], top: true, out var brace))
            return brace;
        var body = ConvertElement(children[0]);
        var over = children[1].Value.Trim();
        return over switch
        {
            "¯" or "‾" => @"\overline{" + body + "}",
            "→" => @"\vec{" + body + "}",
            "←" => @"\overleftarrow{" + body + "}",
            "↔" => @"\overleftrightarrow{" + body + "}",
            "^" or "ˆ" => @"\hat{" + body + "}",
            "~" or "˜" => @"\tilde{" + body + "}",
            "." or "˙" => @"\dot{" + body + "}",
            "¨" => @"\ddot{" + body + "}",
            "⏞" or "\uFE37" => @"\overbrace{" + body + "}",
            _ => @"\overset{" + ConvertElement(children[1]) + "}{" + body + "}",
        };
    }

    private static string ConvertUnder(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        if (TryConvertAnnotatedHorizontalBrace(children[0], children[1], top: false, out var brace))
            return brace;
        var body = ConvertElement(children[0]);
        var under = children[1].Value.Trim();
        return under switch
        {
            "_" or "¯" or "‾" => @"\underline{" + body + "}",
            "⏟" or "\uFE38" => @"\underbrace{" + body + "}",
            _ => @"\underset{" + ConvertElement(children[1]) + "}{" + body + "}",
        };
    }

    private static bool TryConvertAnnotatedHorizontalBrace(
        XElement baseElement,
        XElement annotation,
        bool top,
        out string latex)
    {
        latex = string.Empty;
        if (baseElement.Name.LocalName != (top ? "mover" : "munder")) return false;
        var inner = baseElement.Elements().ToList();
        if (inner.Count < 2) return false;
        var expectedMark = top ? "⏞" : "⏟";
        var presentationMark = top ? "\uFE37" : "\uFE38";
        var marker = inner[1].Value.Trim();
        var replacementOnly = marker.Length > 0
            && marker.All(character => character == '\uFFFD');
        var stretchy = string.Equals(
            (string?)inner[1].Attribute("stretchy"),
            "true",
            StringComparison.OrdinalIgnoreCase);
        // MathType 7 may export horizontal braces as the Unicode vertical
        // presentation forms U+FE37/U+FE38, or as replacement glyphs on some
        // machines. The nested mover/munder shape plus stretchy operator remains
        // the stable signal for recovering an annotated brace.
        if (marker != expectedMark
            && marker != presentationMark
            && !(replacementOnly && stretchy)) return false;
        var body = ConvertElement(inner[0]);
        var note = ConvertElement(annotation).TrimEnd();
        latex = top
            ? @"\overbrace{" + body + "}^{" + note + "}"
            : @"\underbrace{" + body + "}_{" + note + "}";
        return true;
    }

    private static string ConvertUnderOver(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 3) return ConvertChildren(element);
        var baseLatex = children[0].Name.LocalName == "mo"
            ? children[0].Value.Trim() switch
            {
                "∪" => @"\bigcup ",
                "∩" => @"\bigcap ",
                _ => ConvertElement(children[0]),
            }
            : ConvertElement(children[0]);
        return GroupBase(baseLatex)
            + "_{" + ScriptArgument(children[1]) + "}"
            + "^{" + ScriptArgument(children[2]) + "}";
    }

    private static string ConvertFenced(XElement element)
    {
        var open = (string?)element.Attribute("open") ?? "(";
        var close = (string?)element.Attribute("close") ?? ")";
        if (TryConvertMathTypeBinomialPile(element, open, close, out var binomial))
            return binomial;
        var separators = (string?)element.Attribute("separators") ?? ",";
        var parts = element.Elements().Select(ConvertElement).ToList();
        var separator = separators.Length > 0 ? separators[0].ToString() : ",";
        return Delimiter(open, left: true)
            + string.Join(separator, parts)
            + Delimiter(close, left: false);
    }

    private static bool TryConvertMathTypeBinomialPile(
        XElement fenced,
        string open,
        string close,
        out string latex)
    {
        latex = string.Empty;
        if (open != "(" || close != ")") return false;
        var table = fenced.Elements().SingleOrDefault();
        if (table is not null && table.Name.LocalName == "mrow")
        {
            var nested = table.Elements().ToArray();
            if (nested.Length == 1 && nested[0].Name.LocalName == "mtable")
                table = nested[0];
        }
        if (table is null || table.Name.LocalName != "mtable") return false;
        if (!string.Equals(
                (string?)table.Attribute("data-mtef-pile"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            return false;
        var rows = table.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length != 2) return false;
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        if (cells.Any(row => row.Length != 1)) return false;
        latex = @"\binom{" + ConvertElement(cells[0][0]) + "}{"
            + ConvertElement(cells[1][0]) + "}";
        return true;
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
        var up = notation.Contains("updiagonalstrike");
        var down = notation.Contains("downdiagonalstrike");
        if (up && down) return @"\xcancel{" + body + "}";
        if (up) return @"\cancel{" + body + "}";
        if (down) return @"\bcancel{" + body + "}";
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

    private static string ConvertTextElement(XElement element)
    {
        var value = element.Value.Trim();
        // Word exports one-character upright mathematical identifiers such as
        // \mathrm{e}, \mathrm{i} and \mathrm{d} as mtext rather than mi.
        // Preserve their mathematical upright semantics while keeping genuine
        // prose and multi-character annotations as \text{...}.
        if (value.Length == 1 && value[0] <= '\u024F' && char.IsLetter(value[0]))
            return @"\mathrm{" + EscapeMathIdentifier(value) + "}";
        return @"\text{" + EscapeText(element.Value) + "}";
    }

    private static string ConvertTokenElement(XElement element)
    {
        var token = element.Value.Trim();
        var variant = ((string?)element.Attribute("mathvariant") ?? string.Empty).Trim();
        if (element.Name.LocalName is "mi" or "mo"
            && TryConvertNamedOperator(token, out var namedOperator))
            return namedOperator;
        var explicitlyUpright =
            variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
            || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
        if (element.Name.LocalName == "mi"
            && TryConvertLetterlikeIdentifier(token, out var letterlikeLatex))
            return letterlikeLatex;
        if (element.Name.LocalName == "mi" && IsLatinIdentifier(token))
        {
            var escaped = EscapeMathIdentifier(token);
            if (variant.IndexOf("double-struck", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbb{" + escaped + "}";
            if (variant.IndexOf("fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathfrak{" + escaped + "}";
            if (variant.IndexOf("script", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathcal{" + escaped + "}";
            if (variant.IndexOf("monospace", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathtt{" + escaped + "}";
            if (variant.IndexOf("sans-serif", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathsf{" + escaped + "}";
            if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbf{" + escaped + "}";
            if (explicitlyUpright)
                return @"\mathrm{" + escaped + "}";
        }
        return ConvertToken(token);
    }

    private static bool TryConvertNamedOperator(string token, out string latex)
    {
        latex = token switch
        {
            "lim" => @"\lim ",
            "max" => @"\max ",
            "min" => @"\min ",
            "sup" => @"\sup ",
            "inf" => @"\inf ",
            "sin" => @"\sin ",
            "cos" => @"\cos ",
            "tan" => @"\tan ",
            "cot" => @"\cot ",
            "sec" => @"\sec ",
            "csc" => @"\csc ",
            "sinh" => @"\sinh ",
            "cosh" => @"\cosh ",
            "tanh" => @"\tanh ",
            "log" => @"\log ",
            "ln" => @"\ln ",
            "exp" => @"\exp ",
            "det" => @"\det ",
            "gcd" => @"\gcd ",
            _ => string.Empty,
        };
        return latex.Length > 0;
    }

    private static bool TryConvertLetterlikeIdentifier(
        string token,
        out string latex)
    {
        latex = token switch
        {
            "ℂ" => @"\mathbb{C}",
            "ℍ" => @"\mathbb{H}",
            "ℕ" => @"\mathbb{N}",
            "ℙ" => @"\mathbb{P}",
            "ℚ" => @"\mathbb{Q}",
            "ℝ" => @"\mathbb{R}",
            "ℤ" => @"\mathbb{Z}",
            "ℋ" => @"\mathcal{H}",
            "ℐ" => @"\mathcal{I}",
            "ℒ" => @"\mathcal{L}",
            "ℛ" => @"\mathcal{R}",
            "ℬ" => @"\mathcal{B}",
            "ℯ" => @"\mathcal{e}",
            "ℰ" => @"\mathcal{E}",
            "ℱ" => @"\mathcal{F}",
            "ℳ" => @"\mathcal{M}",
            "ℴ" => @"\mathcal{o}",
            "ℌ" => @"\mathfrak{H}",
            "ℑ" => @"\mathfrak{I}",
            "ℜ" => @"\mathfrak{R}",
            "ℭ" => @"\mathfrak{C}",
            _ => string.Empty,
        };
        return latex.Length > 0;
    }

    private static bool IsLatinIdentifier(string token) =>
        token.Length > 0
        && token.Any(char.IsLetter)
        && token.All(character =>
            (character <= '\u024F' && char.IsLetterOrDigit(character))
            || character is '.' or ',');

    private static string EscapeMathIdentifier(string value) =>
        value.Replace("\\", @"\backslash ")
            .Replace("{", @"\{")
            .Replace("}", @"\}")
            .Replace("#", @"\#")
            .Replace("%", @"\%")
            .Replace("&", @"\&")
            .Replace("_", @"\_");

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
            "" or "." => ".",
            "{" => @"\{",
            "}" => @"\}",
            "|" => @"\lvert",
            "‖" => @"\lVert",
            "〈" or "⟨" => @"\langle",
            "〉" or "⟩" => @"\rangle",
            "⌊" => @"\lfloor",
            "⌋" => @"\rfloor",
            "⌈" => @"\lceil",
            "⌉" => @"\rceil",
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

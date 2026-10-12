using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

internal static class MathMlToLatexConverter
{
    // MathJax AMS alternate forms have the same semantic scalar as the
    // ordinary glyph. Only explicit per-token alternate evidence selects these
    // commands; bare Unicode and container attributes never imply an alternate.
    private static readonly IReadOnlyDictionary<string, string> AlternateTokenMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["ℏ"] = @"\hbar ",
            ["∥"] = @"\shortparallel ",
            ["∦"] = @"\nshortparallel ",
            ["∣"] = @"\shortmid ",
            ["∤"] = @"\nshortmid ",
            ["∖"] = @"\smallsetminus ",
            ["⊊"] = @"\varsubsetneq ",
            ["⊋"] = @"\varsupsetneq ",
            ["⫋"] = @"\varsubsetneqq ",
            ["⫌"] = @"\varsupsetneqq ",
            ["⊈"] = @"\nsubseteqq ",
            ["⊉"] = @"\nsupseteqq ",
        };

    // Defaults of the canonical named token commands below, as authored by
    // MathJax's AMS mi mapping. Preserve a normal imported token explicitly
    // rather than letting command expansion silently restore implicit italic.
    private static readonly IReadOnlyDictionary<string, string> TokenDefaultMathVariants =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["∡"] = "italic",
            ["∢"] = "italic",
        };

    private static readonly IReadOnlyDictionary<string, string> TokenMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["−"] = "-",
            ["±"] = @"\pm ",
            ["∓"] = @"\mp ",
            ["×"] = @"\times ",
            ["·"] = @"\cdot ",
            ["⋅"] = @"\cdot ",
            ["°"] = @"\circ ",
            ["∘"] = @"\circ ",
            ["•"] = @"\bullet ",
            ["⋄"] = @"\diamond ",
            ["◃"] = @"\triangleleft ",
            ["⊲"] = @"\triangleleft ",
            ["▹"] = @"\triangleright ",
            ["⊳"] = @"\triangleright ",
            ["⊕"] = @"\oplus ",
            ["⊎"] = @"\uplus ",
            ["⊓"] = @"\sqcap ",
            ["⊔"] = @"\sqcup ",
            ["≀"] = @"\wr ",
            ["⨿"] = @"\amalg ",
            ["⊖"] = @"\ominus ",
            ["⊗"] = @"\otimes ",
            ["⊘"] = @"\oslash ",
            ["⊙"] = @"\odot ",
            ["◯"] = @"\bigcirc ",
            ["÷"] = @"\div ",
            ["∞"] = @"\infty ",
            ["ℏ"] = @"\hslash ",
            ["ℵ"] = @"\aleph ",
            ["℘"] = @"\wp ",
            ["ℜ"] = @"\Re ",
            ["ℑ"] = @"\Im ",
            ["ℓ"] = @"\ell ",
            ["ı"] = @"\imath ",
            ["ȷ"] = @"\jmath ",
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
            ["⋆"] = @"\star ",
            ["†"] = @"\dagger ",
            ["‡"] = @"\ddagger ",
            ["'"] = @"\prime ",
            ["′"] = @"\prime ",
            ["″"] = @"\prime\prime ",
            ["‴"] = @"\prime\prime\prime ",
            ["⁗"] = @"\prime\prime\prime\prime ",
            ["‵"] = @"\backprime ",
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
            ["⁡"] = "⁡", // Preserve the actual application marker in its source slot.
            // A standalone radical glyph is TeX \\surd. Structured square
            // roots arrive as <msqrt> and are handled separately below.
            ["√"] = @"\surd ",
            ["≠"] = @"\ne ",
            ["≈"] = @"\approx ",
            ["≡"] = @"\equiv ",
            ["≤"] = @"\le ",
            ["≥"] = @"\ge ",
            ["≪"] = @"\ll ",
            ["≫"] = @"\gg ",
            ["≃"] = @"\simeq ",
            ["≍"] = @"\asymp ",
            ["≐"] = @"\doteq ",
            ["≺"] = @"\prec ",
            ["≻"] = @"\succ ",
            ["⪯"] = @"\preceq ",
            ["≼"] = @"\preccurlyeq ",
            ["⪰"] = @"\succeq ",
            ["≽"] = @"\succcurlyeq ",
            ["⊢"] = @"\vdash ",
            ["⊣"] = @"\dashv ",
            ["⊨"] = @"\models ",
            ["∠"] = @"\angle ",
            ["∡"] = @"\measuredangle ",
            ["∢"] = @"\sphericalangle ",
            ["∴"] = @"\therefore ",
            ["∵"] = @"\because ",
            // Geometry shapes, including the small/medium forms MathJax emits.
            ["△"] = @"\triangle ",
            ["▵"] = @"\triangle ",
            ["▲"] = @"\blacktriangle ",
            ["▴"] = @"\blacktriangle ",
            ["▽"] = @"\triangledown ",
            ["▿"] = @"\triangledown ",
            ["▼"] = @"\blacktriangledown ",
            ["▾"] = @"\blacktriangledown ",
            ["□"] = @"\square ",
            ["◻"] = @"\square ",
            ["■"] = @"\blacksquare ",
            ["◼"] = @"\blacksquare ",
            ["★"] = @"\bigstar ",
            ["∖"] = @"\setminus ",
            ["∝"] = @"\propto ",
            ["∼"] = @"\sim ",
            ["~"] = @"\sim ",
            ["≅"] = @"\cong ",
            ["∈"] = @"\in ",
            ["∉"] = @"\notin ",
            ["∋"] = @"\ni ",
            ["⊏"] = @"\sqsubset ",
            ["⊑"] = @"\sqsubseteq ",
            ["⊐"] = @"\sqsupset ",
            ["⊒"] = @"\sqsupseteq ",
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
            ["⇐"] = @"\Leftarrow ",
            ["⇒"] = @"\Rightarrow ",
            ["⇔"] = @"\Leftrightarrow ",
            ["→"] = @"\to ",
            ["←"] = @"\leftarrow ",
            ["↔"] = @"\leftrightarrow ",
            ["↑"] = @"\uparrow ",
            ["↓"] = @"\downarrow ",
            ["↕"] = @"\updownarrow ",
            ["⇑"] = @"\Uparrow ",
            ["⇓"] = @"\Downarrow ",
            ["⇕"] = @"\Updownarrow ",
            ["↼"] = @"\leftharpoonup ",
            ["↽"] = @"\leftharpoondown ",
            ["⇀"] = @"\rightharpoonup ",
            ["⇁"] = @"\rightharpoondown ",
            ["↩"] = @"\hookleftarrow ",
            ["↪"] = @"\hookrightarrow ",
            ["↦"] = @"\mapsto ",
            ["↗"] = @"\nearrow ",
            ["↘"] = @"\searrow ",
            ["↙"] = @"\swarrow ",
            ["↖"] = @"\nwarrow ",
            ["⟂"] = @"\perp ",
            ["⊤"] = @"\top ",
            ["⊥"] = @"\perp ",
            ["∥"] = @"\parallel ",
            ["…"] = @"\dots ",
            ["⋯"] = @"\cdots ",
            ["⋮"] = @"\vdots ",
            ["⋱"] = @"\ddots ",
            ["♠"] = @"\spadesuit ",
            ["♣"] = @"\clubsuit ",
            ["♡"] = @"\heartsuit ",
            ["♢"] = @"\diamondsuit ",
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
            ["ϰ"] = @"\varkappa ",
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
        var document = FormulaXml.Parse(mathMl, "MathML", LoadOptions.PreserveWhitespace);
        var root = document.Root
            ?? throw new InvalidDataException("Word MathML has no root element.");
        return Normalize(ConvertElement(root));
    }

    private static string ConvertElement(XElement element) =>
        MathMlColorContract.Wrap(ConvertElementContent(element), MathMlColorContract.Explicit(element));

    private static string ConvertElementContent(XElement element)
    {
        var name = element.Name.LocalName;
        return name switch
        {
            "math" => ConvertChildren(element),
            "mrow" or "mstyle" or "mpadded" => ConvertClassedContainer(element),
            "semantics" or "maction" => element.Elements().Select(ConvertElement).FirstOrDefault() ?? string.Empty,
            "annotation" or "annotation-xml" => string.Empty,
            "mi" or "mn" or "mo" => ConvertTokenElement(element),
            "mtext" => ConvertTextElement(element),
            "mspace" => ConvertSpace(element),
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

    private static string ConvertSpace(XElement element)
    {
        var space = MathSpaceContract.ParseWidth((string?)element.Attribute("width"));
        // Recover every exactly supported source dimension using the shared
        // contract. Unrepresented/context-dependent widths retain the legacy
        // behavior for now; that fallback is not a claim of width fidelity.
        return space.LatexSource ?? @"\,";
    }

    private static string? ExplicitAtomClassCommand(XElement element) =>
        ((string?)element.Attribute("data-mjx-texclass"))?.Trim().ToUpperInvariant() switch
        {
            "REL" => @"\mathrel",
            "BIN" => @"\mathbin",
            "OP" => @"\mathop",
            _ => null,
        };

    private static string ConvertClassedContainer(XElement element)
    {
        var content = ConvertChildren(element);
        var command = ExplicitAtomClassCommand(element);
        // Preserve explicit atom semantics, never infer a package command from
        // the glyph. ORD and delimiter wrappers keep their established paths.
        return command is null || string.IsNullOrWhiteSpace(content)
            ? content
            : command + "{" + content.Trim() + "}";
    }

    private static string ConvertChildren(XElement element)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 3)
            return string.Concat(children.Select(ConvertElement));

        var builder = new StringBuilder();
        for (var index = 0; index < children.Length;)
        {
            if (index + 2 < children.Length
                && IsSimpleOperator(children[index], "(")
                && IsSimpleOperator(children[index + 2], ")")
                && TryConvertLooseBinomialBody(children[index + 1], out var binomial))
            {
                builder.Append(binomial);
                index += 3;
                continue;
            }
            builder.Append(ConvertElement(children[index]));
            index++;
        }
        return builder.ToString();
    }

    private static bool TryConvertLooseBinomialBody(XElement body, out string latex)
    {
        latex = string.Empty;
        body = UnwrapSingleTransparentMathMlContainer(body);
        if (body.Name.LocalName == "mfrac")
        {
            var lineThickness = ((string?)body.Attribute("linethickness") ?? string.Empty).Trim();
            var fractionChildren = body.Elements().ToArray();
            if (!lineThickness.StartsWith("0", StringComparison.Ordinal)
                || fractionChildren.Length < 2)
                return false;
            latex = @"\binom{" + ConvertElement(fractionChildren[0]) + "}{"
                + ConvertElement(fractionChildren[1]) + "}";
            return true;
        }

        if (body.Name.LocalName != "mtable"
            || !string.Equals(
                (string?)body.Attribute("data-mtef-pile"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            return false;
        var rows = body.Elements()
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

    private static bool IsSimpleOperator(XElement element, string value)
    {
        element = UnwrapSingleTransparentMathMlContainer(element);
        return element.Name.LocalName == "mo"
            && string.Equals(element.Value.Trim(), value, StringComparison.Ordinal);
    }

    private static string ConvertFraction(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        // A zero-width rule is a stacked pair without a bar (\atop, an m:f of
        // type noBar). \binom is recognized with its fence elsewhere; the bare
        // form must not come back with a rule, or an edit would draw one.
        var thickness = ((string?)element.Attribute("linethickness") ?? string.Empty).Trim();
        if (thickness.StartsWith("0", StringComparison.Ordinal))
            return @"\genfrac{}{}{0pt}{}{" + ConvertElement(children[0]) + "}{" + ConvertElement(children[1]) + "}";
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
        var stretchy = MathMlAccentGeometry.IsStretchy(children[1]);
        var accentStyle = over is "¯" or "‾" or "―" or "→" or "←" or "↔"
            or "^" or "ˆ" or "~" or "˜" or "." or "˙" or "¨"
            ? MathMlTokenStyleContract.ResolveEffectiveVariant(children[1]) : null;
        var styleCommand = accentStyle switch
        {
            "bold" => @"\mathbf",
            "italic" => @"\mathit",
            "bold-italic" => @"\mathbfit",
            _ => null,
        };
        // The mark owns this scope; reset its body before emitting the body's
        // own styles so the accent font cannot restyle its operands.
        if (styleCommand is not null) body = @"\mathnormal{" + body + "}";
        var markColor = MathMlColorContract.Effective(children[1]);
        var parentColor = MathMlColorContract.Effective(element);
        var independentColor = markColor is not null && markColor != parentColor;
        if (independentColor)
            body = MathMlColorContract.Wrap(body, MathMlColorContract.Effective(children[0]) ?? "#000000");
        var result = over switch
        {
            "¯" or "‾" or "―" => @"\overline{" + body + "}",
            "→" => (stretchy ? @"\overrightarrow{" : @"\vec{") + body + "}",
            "←" => @"\overleftarrow{" + body + "}",
            "↔" => @"\overleftrightarrow{" + body + "}",
            "^" or "ˆ" => (stretchy ? @"\widehat{" : @"\hat{") + body + "}",
            "~" or "˜" => (stretchy ? @"\widetilde{" : @"\tilde{") + body + "}",
            "." or "˙" => @"\dot{" + body + "}",
            "¨" => @"\ddot{" + body + "}",
            "⏞" or "\uFE37" => @"\overbrace{" + body + "}",
            _ => @"\overset{" + ConvertElement(children[1]) + "}{" + body + "}",
        };
        if (styleCommand is not null) result = styleCommand + "{" + result + "}";
        return MathMlColorContract.Wrap(result, independentColor ? markColor : null);
    }

    private static string ConvertUnder(XElement element)
    {
        var children = element.Elements().ToList();
        if (children.Count < 2) return ConvertChildren(element);
        if (TryConvertAnnotatedHorizontalBrace(children[0], children[1], top: false, out var brace))
            return brace;
        var body = ConvertElement(children[0]);
        var under = children[1].Value.Trim();
        if (IsLimitLikeOperatorLatex(body))
            return GroupBase(body) + "_{" + ScriptArgument(children[1]) + "}";
        return under switch
        {
            "_" or "¯" or "‾" or "―" => @"\underline{" + body + "}",
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
        baseElement = UnwrapSingleTransparentMathMlContainer(baseElement);
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

    private static XElement UnwrapSingleTransparentMathMlContainer(XElement element)
    {
        var current = element;
        while (current.Name.LocalName is "mrow" or "mstyle" or "mpadded")
        {
            if (ExplicitAtomClassCommand(current) is not null) break;
            var children = current.Elements().ToArray();
            if (children.Length != 1) break;
            current = children[0];
        }
        return current;
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
        if (TryConvertFencedNoBarFractionToBinomial(element, open, close, out var noBarBinomial))
            return noBarBinomial;
        if (TryConvertMathTypeBinomialPile(element, open, close, out var binomial))
            return binomial;
        var separators = (string?)element.Attribute("separators") ?? ",";
        var parts = element.Elements().Select(ConvertElement).ToList();
        var separator = separators.Length > 0 ? separators[0].ToString() : string.Empty;
        return Delimiter(open, left: true)
            + string.Join(separator, parts)
            + Delimiter(close, left: false);
    }

    private static bool TryConvertFencedNoBarFractionToBinomial(
        XElement fenced,
        string open,
        string close,
        out string latex)
    {
        latex = string.Empty;
        if (open != "(" || close != ")") return false;
        var children = fenced.Elements().Take(2).ToArray();
        if (children.Length != 1) return false;
        var child = UnwrapSingleTransparentMathMlContainer(children[0]);
        if (child.Name.LocalName != "mfrac") return false;
        var lineThickness = ((string?)child.Attribute("linethickness") ?? string.Empty).Trim();
        if (!lineThickness.StartsWith("0", StringComparison.Ordinal)) return false;
        var fractionChildren = child.Elements().ToArray();
        if (fractionChildren.Length < 2) return false;
        latex = @"\binom{" + ConvertElement(fractionChildren[0]) + "}{"
            + ConvertElement(fractionChildren[1]) + "}";
        return true;
    }

    private static bool TryConvertMathTypeBinomialPile(
        XElement fenced,
        string open,
        string close,
        out string latex)
    {
        latex = string.Empty;
        if (open != "(" || close != ")") return false;
        var fencedChildren = fenced.Elements().Take(2).ToArray();
        if (fencedChildren.Length != 1) return false;
        var table = fencedChildren[0];
        if (table.Name.LocalName == "mrow")
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
        if (string.Equals(
                (string?)element.Attribute("data-mtef-tabs"),
                "true",
                StringComparison.OrdinalIgnoreCase))
        {
            // MathType stores native multi-point equation alignment as PILE rows
            // separated by fnMARKER tab characters. The MTEF reader reconstructs
            // those tabs as table cells and marks the private provenance here so
            // editing/conversion returns to `aligned` instead of degrading to a
            // plain matrix and losing the user's `&`/`&&` alignment semantics.
            return @"\begin{aligned}" + string.Join(@" \\ ", rows) + @"\end{aligned}";
        }
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
        var postPair = 0;
        while (index < children.Count && children[index].Name.LocalName != "mprescripts")
        {
            var sub = ScriptArgument(children[index]);
            var sup = index + 1 < children.Count && children[index + 1].Name.LocalName != "mprescripts"
                ? ScriptArgument(children[index + 1]) : string.Empty;
            // Each additional pair is an adjacent script column, not a second
            // superscript on the same TeX atom (which would be invalid TeX).
            if (postPair++ > 0 && (sub.Length > 0 || sup.Length > 0)) builder.Append("{}");
            if (sub.Length > 0) builder.Append("_{").Append(sub).Append('}');
            if (sup.Length > 0) builder.Append("^{").Append(sup).Append('}');
            index += 2;
        }
        if (index < children.Count && children[index].Name.LocalName == "mprescripts")
        {
            var prescripts = new List<string>();
            for (index++; index < children.Count; index += 2)
            {
                var sub = ScriptArgument(children[index]);
                var sup = index + 1 < children.Count ? ScriptArgument(children[index + 1]) : string.Empty;
                if (sub.Length == 0 && sup.Length == 0) continue;
                prescripts.Add("{}"
                    + (sub.Length > 0 ? "_{" + sub + "}" : string.Empty)
                    + (sup.Length > 0 ? "^{" + sup + "}" : string.Empty));
            }
            // MathML lists prescript columns from the base outward. TeX is
            // emitted from left to right, so those columns precede it in reverse.
            prescripts.Reverse();
            builder.Insert(0, string.Concat(prescripts));
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
        // A lone geometry symbol stored as text (older VisualTeX equations put
        // \triangle in MathType's Text face) is the math symbol; inside \text{}
        // the editor's text font has no glyph for it.
        if (value.Length == 1 && GeometryTextSymbols.Contains(value[0]))
            return ConvertToken(value);
        return @"\text{" + EscapeText(element.Value) + "}";
    }

    private static readonly HashSet<char> GeometryTextSymbols = new()
    {
        '△', '▵', '▲', '▴', '▽', '▿', '▼', '▾', '□', '◻', '■', '◼', '★',
        '∠', '∡', '∵', '∴', '⊥', '∥', '≅',
    };

    private static string ConvertTokenElement(XElement element)
    {
        var token = element.Value.Trim();
        var declaredVariant = MathMlTokenStyleContract.ResolveDeclaredVariant(element);
        var variant = declaredVariant ?? string.Empty;
        var effectiveVariant = MathMlTokenStyleContract.ResolveEffectiveVariant(element, declaredVariant);
        // Source spelling is an adapter decision: some named TeX symbols have
        // upright intrinsic defaults, so materialize their effective mi italic.
        // Ordinary Latin and Greek retain their established command spelling.
        if (declaredVariant is null && effectiveVariant == "italic"
            && char.GetUnicodeCategory(token, 0) == System.Globalization.UnicodeCategory.MathSymbol)
            variant = effectiveVariant;
        var nativeCharacterStyle = int.TryParse(
            (string?)element.Attribute("data-mtef-character-style"),
            out var parsedNativeCharacterStyle)
            ? parsedNativeCharacterStyle
            : 0;
        if (element.Name.LocalName == "mi" && token.Length > 1
            && token.All(c => c is >= 'A' and <= 'Z' or >= 'a' and <= 'z')
            && effectiveVariant is "normal" or "upright")
        {
            var atomClass = ((string?)element.Attribute("data-mjx-texclass"))?.Trim().ToUpperInvariant();
            var isOperator = atomClass == "OP" || (atomClass is null
                && (string?)element.Attribute("data-mjx-auto-op") != "false");
            // Preserve one upright identifier token without inventing application.
            // Explicit atom classes override MathJax normal-mi autoOP inference.
            return (isOperator ? @"\operatorname{" : @"\mathrm{") + EscapeMathIdentifier(token) + "}";
        }
        if (element.Name.LocalName == "mo"
            && TryConvertNamedOperator(token, out var namedOperator))
        {
            // A genuine MathType explicit FONT_STYLE_DEF can style an operator
            // even though its server-exported MathML omits that presentation.
            // The direct MTEF reader carries the style as mathvariant; do not
            // return the named command before applying it.
            if (variant.IndexOf("bold-italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbfit{" + namedOperator.Trim() + "}";
            if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\boldsymbol{" + namedOperator.Trim() + "}";
            if (variant.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathit{" + namedOperator.Trim() + "}";
            return namedOperator;
        }
        var explicitlyUpright =
            variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
            || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
        if (element.Name.LocalName == "mi"
            && TryConvertLetterlikeIdentifier(token, out var letterlikeLatex))
            return ApplyNativeMathTypeCharacterStyle(
                letterlikeLatex,
                nativeCharacterStyle,
                variant,
                greekIdentifier: false);
        if (variant == "italic" && nativeCharacterStyle is 0 or 2
            && TryGetIntrinsicItalicIdentifierSource(element, token, out var italicSource))
            return italicSource;
        if (element.Name.LocalName == "mi" && IsLatinIdentifier(token)
            || element.Name.LocalName == "mn")
        {
            var escaped = EscapeMathIdentifier(token);
            string? converted = null;
            if (variant.IndexOf("double-struck", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathbb{" + escaped + "}";
            else if (variant.IndexOf("bold-fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\boldsymbol{\mathfrak{" + escaped + "}}";
            else if (variant.IndexOf("fraktur", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathfrak{" + escaped + "}";
            else if (variant.IndexOf("bold-script", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\boldsymbol{\mathcal{" + escaped + "}}";
            else if (variant.IndexOf("script", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathcal{" + escaped + "}";
            else if (variant.IndexOf("monospace", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathtt{" + escaped + "}";
            else if (variant.IndexOf("bold-sans-serif", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\boldsymbol{\mathsf{" + escaped + "}}";
            else if (variant.IndexOf("sans-serif", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathsf{" + escaped + "}";
            else if (variant.IndexOf("bold-italic", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathbfit{" + escaped + "}";
            else if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                converted = @"\mathbf{" + escaped + "}";
            else if (element.Name.LocalName == "mn"
                && string.Equals(variant, "italic", StringComparison.OrdinalIgnoreCase))
                converted = @"\mathit{" + escaped + "}";
            else if (explicitlyUpright)
                converted = @"\mathrm{" + escaped + "}";
            if (converted is not null)
                return ApplyNativeMathTypeCharacterStyle(
                    converted,
                    nativeCharacterStyle,
                    variant,
                    greekIdentifier: false);
        }
        if (element.Name.LocalName == "mi" && IsGreekIdentifier(token))
            return ConvertGreekIdentifier(token, nativeCharacterStyle, declaredVariant, effectiveVariant);
        var plain = element.Name.LocalName is "mi" or "mo"
            && string.Equals(((string?)element.Attribute("data-mjx-alternate"))?.Trim(), "1", StringComparison.Ordinal)
            && AlternateTokenMap.TryGetValue(token, out var alternateCommand)
            ? alternateCommand
            : ConvertToken(token);
        if (element.Name.LocalName is "mi" or "mo")
        {
            if (variant.IndexOf("bold-italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbfit{" + plain.Trim() + "}";
            if (variant.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathbf{" + plain.Trim() + "}";
            if (variant.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0)
                return @"\mathit{" + plain.Trim() + "}";
            if ((explicitlyUpright && element.Name.LocalName == "mi")
                || ((explicitlyUpright || variant.Length == 0)
                    && TokenDefaultMathVariants.TryGetValue(token, out var intrinsicVariant)
                    && string.Equals(intrinsicVariant, "italic", StringComparison.Ordinal)))
                return @"\mathrm{" + plain.Trim() + "}";
        }
        return plain;
    }

    private static bool TryGetIntrinsicItalicIdentifierSource(
        XElement element, string token, out string source)
    {
        source = string.Empty;
        // MathML italic is the mathematical italic alphabet. TeX \mathit is a
        // different alphabet for Latin letters (MathJax TEX-MI vs TEX-I), and
        // only its actual token-level provenance can authorize that spelling.
        if (element.Name.LocalName != "mi"
            || !MathMlTokenStyleContract.IsSingleMathMlCharacter(token)
            || string.Equals((string?)element.Attribute("data-mjx-variant"),
                "-tex-mathit", StringComparison.Ordinal))
            return false;
        if (IsLatinIdentifier(token))
            source = EscapeMathIdentifier(token);
        else if (IsGreekIdentifier(token) && token.All(char.IsLower))
            source = ConvertToken(token);
        return source.Length > 0;
    }

    private static string ConvertGreekIdentifier(
        string token, int nativeCharacterStyle, string? declaredVariant, string? effectiveVariant)
    {
        // Keep the established Unicode spelling for explicitly styled Greek.
        // It participates directly in the requested mathematical alphabet,
        // including frontends with TeX fixed-family control-sequence defaults.
        // Preserve the requested style before choosing source spelling.
        var variant = declaredVariant ?? effectiveVariant ?? string.Empty;
        // Lowercase Greek commands already have an italic intrinsic default.
        // An undeclared uppercase mi does not: e.g. TeX \\Gamma is normal while
        // MathML <mi>Γ</mi> is italic. Source spelling must preserve that default.
        if (declaredVariant is null && nativeCharacterStyle == 0
            && variant == "italic" && token.All(char.IsLower))
            return ConvertToken(token);
        variant = nativeCharacterStyle switch
        {
            1 => "bold",
            2 => "italic",
            3 => "bold-italic",
            _ => variant,
        };
        if (variant is "normal" or "upright")
            return @"\mathrm{" + token + "}";
        if (variant == "italic")
            return @"\mathit{" + token + "}";
        if (variant == "bold")
            return token.All(char.IsUpper) && TokenMap.TryGetValue(token, out var upperCommand)
                ? @"\boldsymbol{" + upperCommand.Trim() + "}"
                : @"\mathbf{" + token + "}";
        if (variant == "bold-italic")
            return token.All(char.IsLower)
                ? @"\boldsymbol{" + ConvertToken(token).Trim() + "}"
                : @"\boldsymbol{\mathit{" + token + "}}";
        return ApplyNativeMathTypeCharacterStyle(
            ConvertToken(token), nativeCharacterStyle, variant, greekIdentifier: true);
    }

    private static string ApplyNativeMathTypeCharacterStyle(
        string latex,
        int characterStyle,
        string mathVariant,
        bool greekIdentifier)
    {
        if (characterStyle is < 1 or > 3)
            return latex;

        var hasBold = mathVariant.IndexOf(
            "bold",
            StringComparison.OrdinalIgnoreCase) >= 0;
        var hasItalic = mathVariant.IndexOf(
            "italic",
            StringComparison.OrdinalIgnoreCase) >= 0;

        return characterStyle switch
        {
            1 when hasBold => latex,
            1 when greekIdentifier => @"\boldsymbol{" + latex.Trim() + "}",
            1 => @"\mathbf{" + latex + "}",
            2 when hasItalic => latex,
            2 => @"\mathit{" + latex + "}",
            3 when hasBold && hasItalic => latex,
            3 when greekIdentifier => @"\boldsymbol{" + latex.Trim() + "}",
            3 => @"\mathbfit{" + latex + "}",
            _ => latex,
        };
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
            "limsup" => @"\limsup ",
            "liminf" => @"\liminf ",
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

    private static bool IsLimitLikeOperatorLatex(string value)
    {
        return value.Trim() switch
        {
            @"\lim" or @"\max" or @"\min" or @"\sup" or @"\inf"
                or @"\limsup" or @"\liminf" => true,
            _ => false,
        };
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
            "ℑ" => @"\Im ",
            "ℜ" => @"\Re ",
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

    private static bool IsGreekIdentifier(string token) =>
        token.Length > 0
        && token.All(character =>
            character is >= '\u0370' and <= '\u03FF'
            || character is >= '\u1F00' and <= '\u1FFF');

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
        // MathML token text is literal data, not a TeX fragment. MTEF can
        // legitimately group a percent with digits (one <mn>100%</mn>).
        // Escaping only when the entire token equalled "%" turned its suffix
        // into a TeX comment on the next edit. Serialize every literal character
        // before combining tokens; structural commands/alignments are emitted
        // by their own nodes and TokenMap above, never escaped a second time.
        return EscapeMathIdentifier(token);
    }

    // Canonical scalar spellings from the production TeX delimiter handlers:
    // base, AMS and mathtools (including their active priority/shadowing order).
    // The production roundtrip probe checks the complete runtime inventory.
    // GetDelimiter accepts registered spellings, not arbitrary Unicode scalars:
    // e.g. U+27EE and U+23D0 require \\lgroup and \\arrowvert. Keep this role
    // separate from TokenMap so ordinary operators/prose retain their semantics.
    private static readonly IReadOnlyDictionary<string, string> DelimiterMap =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["("] = @"(",
            [")"] = @")",
            ["["] = @"[",
            ["]"] = @"]",
            ["/"] = @"/",
            ["\\"] = @"\backslash",
            ["|"] = @"|",
            ["‖"] = @"\|",
            ["{"] = @"\{",
            ["}"] = @"\}",
            ["⟨"] = @"\langle",
            ["⟩"] = @"\rangle",
            ["⌈"] = @"\lceil",
            ["⌉"] = @"\rceil",
            ["⌊"] = @"\lfloor",
            ["⌋"] = @"\rfloor",
            ["⌜"] = @"\ulcorner",
            ["⌝"] = @"\urcorner",
            ["⌞"] = @"\llcorner",
            ["⌟"] = @"\lrcorner",
            ["⎰"] = @"\lmoustache",
            ["⎱"] = @"\rmoustache",
            ["⟮"] = @"\lgroup",
            ["⟯"] = @"\rgroup",
            ["⏐"] = @"\arrowvert",
            ["⎪"] = @"\bracevert",
            ["↑"] = @"\uparrow",
            ["↓"] = @"\downarrow",
            ["↕"] = @"\updownarrow",
            ["⇑"] = @"\Uparrow",
            ["⇓"] = @"\Downarrow",
            ["⇕"] = @"\Updownarrow",
            // Preserve existing invisible and legacy angle spellings.
            [""] = ".",
            ["."] = ".",
            ["<"] = "<",
            [">"] = ">",
            ["〈"] = @"\langle",
            ["〉"] = @"\rangle",
        };

    private static string Delimiter(string value, bool left)
    {
        // Unknown values retain the legacy source rather than being silently
        // substituted with a different glyph or claimed as supported.
        var escaped = DelimiterMap.TryGetValue(value, out var mapped) ? mapped : value;
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

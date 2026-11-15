using System.Globalization;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefCodec;
using static VisualTeX.WordVsto.MathTypeMtefAssembly;
using static VisualTeX.WordVsto.MathTypeMtefCharacters;
using static VisualTeX.WordVsto.MathTypeMtefFences;
using static VisualTeX.WordVsto.MathTypeMtefSignature;

namespace VisualTeX.WordVsto;

/// <summary>
/// Emitting MathML structures as MTEF objects: lines and templates (fractions, roots, scripts, accents, integrals,
/// big operators, limits, fences), piles and matrices.
/// </summary>
internal static class MathTypeMtefTemplates
{
    private static void EmitMathSpace(XElement element, List<byte> output)
    {
        var space = MathSpaceContract.ParseWidth((string?)element.Attribute("width")).NativeSpace;
        if (space is null)
        {
            // Preserve compatibility for unproven widths; this is not a claim
            // that ASCII space faithfully represents an arbitrary MathML gap.
            EmitCharacter(' ', TypefaceSpace, output);
            return;
        }
        // An edited width is authoritative over stale native CHAR metadata.
        // For an unchanged known space, preserve its original encoding flags.
        if ((int?)element.Attribute("data-mtef-explicit-typeface") == TypefaceSpace
            && (int?)element.Attribute("data-mtef-mtcode") == space.MtCode
            && !IsTrueMathMlBoolean((string?)element.Attribute("data-mtef-no-mtcode"))
            && TryEmitPreservedCharacter(element, TypefaceSpace, output))
            return;
        EmitScalar(space.MtCode, TypefaceSpace, output);
    }

    internal static void EmitNode(XNode node, MtefWriter output)
    {
        if (node is XElement element
            && MathMlColorContract.Explicit(element) is { } color)
        {
            using (output.EnterExplicitColor(color))
                EmitNodeCore(node, output);
            return;
        }
        EmitNodeCore(node, output);
    }

    private static void EmitNodeCore(XNode node, MtefWriter output)
    {
        if (node is XText text)
        {
            EmitText(text.Value, TypefaceText, output);
            return;
        }
        if (node is not XElement element) return;
        var local = element.Name.LocalName;
        switch (local)
        {
            case "math":
            case "mrow":
            case "semantics":
            case "mpadded":
                EmitContainerChildren(element, output, inheritedMathVariant: null);
                break;
            case "mphantom":
                EmitPhantomSpace(element, output);
                break;
            case "mstyle":
                EmitContainerChildren(
                    element,
                    output,
                    ((string?)element.Attribute("mathvariant"))?.Trim());
                break;
            case "mi":
                EmitIdentifier(element, output);
                break;
            case "mn":
                if (element.Attribute("data-mtef-explicit-typeface") is not null)
                    EmitIdentifier(element, output);
                else EmitText(element.Value, TypefaceNumber, output);
                break;
            case "mo":
            {
                var operatorValue = element.Value.Trim();
                if (TryResolveNativeIntegral(element, out var integral))
                    EmitIntegralTemplate(integral, main: null, lower: null, upper: null, output);
                else if (TryResolveNativeBigOperator(element, out var standaloneBigOperator))
                    EmitBigOperatorTemplate(standaloneBigOperator, main: null, lower: null, upper: null, output);
                else if (IsNamedLimitOperator(element)
                    || operatorValue.Length > 1 && operatorValue.All(char.IsLetter))
                    // MathJax emits alphabetic operators such as TeX `\bmod` as
                    // one upright <mo> token. Writing each ASCII letter through
                    // the Symbol typeface made the MTEF reader recover three
                    // unrelated operator tokens (`m`, `o`, `d`), tripping the
                    // strict standalone semantic gate. MathType's native
                    // function run is the matching MTEF representation for an
                    // upright alphabetic operator and round-trips as one run.
                {
                    if (MathMlTokenStyleContract.ResolveEffectiveVariant(element) is "bold" or "italic" or "bold-italic")
                        EmitIdentifier(element, output);
                    else
                        EmitFunctionRun(operatorValue, output);
                }
                else
                    EmitOperator(element, output);
                break;
            }
            case "mtext":
                EmitText(element.Value, TypefaceText, output);
                break;
            case "mspace":
                EmitMathSpace(element, output);
                break;
            case "mfrac":
                EmitFraction(element, output);
                break;
            case "msqrt":
                EmitSquareRoot(element, output);
                break;
            case "mroot":
                EmitNthRoot(element, output);
                break;
            case "msup":
                EmitScript(element, TemplateSup, output);
                break;
            case "msub":
                EmitScript(element, TemplateSub, output);
                break;
            case "msubsup":
                EmitScript(element, TemplateSubSup, output);
                break;
            case "mfenced":
                EmitFenced(element, output);
                break;
            case "mover":
                EmitOver(element, output);
                break;
            case "munder":
                EmitUnder(element, output);
                break;
            case "munderover":
                EmitUnderOver(element, output);
                break;
            case "mtable":
                EmitMatrix(element, output);
                break;
            case "mtr":
            case "mlabeledtr":
            case "mtd":
                EmitContainerChildren(element, output, inheritedMathVariant: null);
                break;
            case "menclose":
                EmitEnclose(element, output);
                break;
            case "mmultiscripts":
                EmitMultiScripts(element, output);
                break;
            case "none":
            case "annotation":
            case "annotation-xml":
                break;
            default:
                // Unknown pure presentation wrappers are flattened. Unknown leaf
                // tokens are kept as text so common MathML extensions remain
                // editable instead of producing structurally invalid MTEF.
                if (element.Elements().Any())
                    EmitContainerChildren(element, output, inheritedMathVariant: null);
                else if (!string.IsNullOrEmpty(element.Value))
                    EmitText(element.Value, TypefaceText, output);
                break;
        }
    }

    internal static IEnumerable<XNode> SignificantChildren(XElement element) =>
        element.Nodes().Where(node =>
            node is XElement
            || node is XText text && !string.IsNullOrWhiteSpace(text.Value));

    internal static bool IsEmptyScriptBase(XElement element)
    {
        // Presentation-only wrappers cannot turn a zero-content base into an
        // atom. The same predicate governs explicit TeX prescript adoption and
        // native trailing-script isolation, so mstyle/mpadded/semantics nesting
        // never changes which neighboring object owns the script.
        if (element.Name.LocalName is not ("mrow" or "mi" or "mn" or "mtext"
            or "mstyle" or "mpadded" or "semantics" or "math"))
            return false;
        if (element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            return false;
        return element.Elements().Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
            .All(IsEmptyScriptBase);
    }

    private static void EmitScriptBase(XElement element, MtefWriter output)
    {
        // MTEF trailing-script templates consume the preceding OBJECT, not a
        // MathML token. Adjacent CHAR records of the same native run can be
        // coalesced by MathType (for example mtext "%" followed by an upright
        // identifier "e" becomes one Text object "%e"). If the base is emitted
        // as a bare CHAR, the following script then binds to that merged object
        // and silently captures the preceding sibling. A LINE is MTEF's native
        // grouping object, so every explicit MathML script base is materialized
        // as one LINE. This preserves the token/object boundary independent of
        // typeface, styling, or the preceding sibling instead of special-casing
        // percent, text, numbers, or any particular formula.
        EmitLine(element, output);
    }

    private static void EmitContainerChildren(
        XElement element,
        MtefWriter output,
        string? inheritedMathVariant)
    {
        var children = new List<XNode>();
        foreach (var child in SignificantChildren(element))
        {
            if (child is XElement childElement
                && (childElement.Name.LocalName == "annotation"
                    || childElement.Name.LocalName == "annotation-xml"))
                continue;
            if (string.IsNullOrWhiteSpace(inheritedMathVariant)
                || child is not XElement styledChild)
            {
                children.Add(child);
                continue;
            }
            var clone = new XElement(styledChild);
            ApplyInheritedMathVariant(clone, inheritedMathVariant!);
            children.Add(clone);
        }

        for (var index = 0; index < children.Count; index++)
        {
            if (children[index] is XElement bigOperatorHead
                && TryResolveBigOperatorHead(bigOperatorHead, out var resolvedHead))
            {
                // The operand runs to the next sign or relation outside brackets,
                // as in MathType's own storage of \prod_{k}(1+a_k); a closing
                // bracket it did not open ends it, and an integrand ends at the
                // space before its differential (\int f\,dx).
                var mainNodes = new List<XNode>();
                var cursor = index + 1;
                var depth = 0;
                while (cursor < children.Count)
                {
                    var change = BracketDepthChange(children[cursor]);
                    if (depth == 0 && (change < 0 || IsBigOperatorMainBoundary(children[cursor])
                            || resolvedHead.IsIntegral && IsExplicitSpace(children[cursor])))
                        break;
                    depth += change;
                    mainNodes.Add(CloneNode(children[cursor]));
                    cursor++;
                }
                if (mainNodes.Count > 0)
                {
                    var main = new XElement("mrow", mainNodes);
                    EmitResolvedBigOperator(resolvedHead, main, output);
                    index = cursor - 1;
                    continue;
                }
            }
            EmitNode(children[index], output);
        }
    }

    private readonly struct ResolvedBigOperatorHead
    {
        public ResolvedBigOperatorHead(
            NativeIntegral integral,
            NativeBigOperator bigOperator,
            XElement? lower,
            XElement? upper,
            bool limitsAbove)
        {
            Integral = integral;
            BigOperator = bigOperator;
            Lower = lower;
            Upper = upper;
            LimitsAbove = limitsAbove;
        }

        public NativeIntegral Integral { get; }
        public NativeBigOperator BigOperator { get; }
        public XElement? Lower { get; }
        public XElement? Upper { get; }
        // munder/mover/munderover: limits above and below the operator.
        public bool LimitsAbove { get; }
        public bool IsIntegral => Integral.VariationKind != 0;
    }

    private static bool TryResolveBigOperatorHead(
        XElement element,
        out ResolvedBigOperatorHead resolved)
    {
        resolved = default;
        var local = element.Name.LocalName;
        var children = element.Elements().ToArray();
        XElement? baseElement = element;
        XElement? lower = null;
        XElement? upper = null;
        if (local is "msub" or "msup" or "msubsup" or "munder" or "mover" or "munderover")
        {
            if (children.Length < 2) return false;
            if (local == "mover" && IsOverDecoration(element, children[1]))
                return false;
            baseElement = children[0];
            if (local is "msub" or "munder") lower = children[1];
            else if (local is "msup" or "mover") upper = children[1];
            else
            {
                if (children.Length < 3) return false;
                lower = children[1];
                upper = children[2];
            }
        }
        else if (local != "mo")
        {
            return false;
        }

        var limitsAbove = LimitsAbove(element);
        if (TryResolveNativeIntegral(baseElement, out var integral))
        {
            resolved = new ResolvedBigOperatorHead(integral, default, lower, upper, limitsAbove);
            return true;
        }
        if (TryResolveNativeBigOperator(
                baseElement,
                out var bigOperator,
                allowAmbiguousUnionIntersection: local != "mo"))
        {
            resolved = new ResolvedBigOperatorHead(default, bigOperator, lower, upper, limitsAbove);
            return true;
        }
        return false;
    }

    private static void EmitResolvedBigOperator(
        ResolvedBigOperatorHead resolved,
        XElement main,
        MtefWriter output)
    {
        if (resolved.IsIntegral)
            EmitIntegralTemplate(
                resolved.Integral,
                main,
                resolved.Lower,
                resolved.Upper,
                output,
                resolved.LimitsAbove);
        else
            EmitBigOperatorTemplate(
                resolved.BigOperator,
                main,
                resolved.Lower,
                resolved.Upper,
                output,
                resolved.LimitsAbove);
    }

    private static bool IsBigOperatorMainBoundary(XNode node)
    {
        if (node is not XElement element || element.Name.LocalName != "mo") return false;
        var value = element.Value.Trim();
        return value is "+" or "-" or "−" or "±" or "∓"
            or "=" or "≠" or "<" or ">" or "≤" or "≥" or "≈" or "≃" or "∼"
            or "≡" or "∝" or "∈" or "∉" or "⊂" or "⊆" or "⊃" or "⊇"
            or "," or ";";
    }

    // MathJax's \, is an mspace, wrapped in an mstyle.
    private static bool IsExplicitSpace(XNode node) =>
        node is XElement element
        && (element.Name.LocalName == "mspace"
            || element.Name.LocalName == "mstyle" && element.Elements().Count() == 1
                && element.Elements().Single().Name.LocalName == "mspace");

    private static int BracketDepthChange(XNode node) =>
        node is XElement { Name.LocalName: "mo" } bracket
            ? bracket.Value.Trim() switch
            {
                "(" or "[" or "{" or "⟨" or "⌈" or "⌊" => 1,
                ")" or "]" or "}" or "⟩" or "⌉" or "⌋" => -1,
                _ => 0,
            }
            : 0;

    internal static XNode CloneNode(XNode node) => node switch
    {
        XElement element => new XElement(element),
        XText text => new XText(text.Value),
        _ => new XText(node.ToString()),
    };

    private static void ApplyInheritedMathVariant(XElement element, string mathVariant)
    {
        if (element.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
        {
            if (element.Attribute("mathvariant") is null)
                element.SetAttributeValue("mathvariant", mathVariant);
            return;
        }
        foreach (var descendant in element.Descendants())
        {
            if (descendant.Name.LocalName is not ("mi" or "mn" or "mo" or "mtext"))
                continue;
            if (descendant.Attribute("mathvariant") is null)
                descendant.SetAttributeValue("mathvariant", mathVariant);
        }
    }

    private static void EmitFraction(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
            throw new InvalidDataException("MathML mfrac requires numerator and denominator.");

        var lineThickness = ((string?)element.Attribute("linethickness") ?? string.Empty).Trim();
        if (lineThickness.Length > 0 && lineThickness.StartsWith("0", StringComparison.Ordinal))
        {
            // MathJax represents \binom and related stacked expressions as an
            // mfrac with a zero rule. MathType's native equivalent is a centered
            // PILE, not the ordinary fraction template (which always draws a rule).
            output.AddRange(new byte[]
            {
                RecordPile,
                0,
                2, // horizontal alignment: centered
                1, // vertical alignment: centered
            });
            EmitLine(children[0], output);
            EmitLine(children[1], output);
            output.Add(RecordEnd);
            return;
        }

        // \tfrac (MathJax: the fraction alone in a displaystyle="false" mstyle)
        // is MathType's small fraction: variation 1, both slots in script size.
        if (element.Parent is { } style && style.Name.LocalName == "mstyle"
            && string.Equals(((string?)style.Attribute("displaystyle"))?.Trim(), "false", StringComparison.Ordinal)
            && style.Elements().Count() == 1)
        {
            output.AddRange(new byte[] { RecordTemplate, 0, TemplateFraction, 1, 0 });
            using (output.EnterScriptSize())
            {
                EmitLine(children[0], output);
                EmitLine(children[1], output);
                output.Add(RecordEnd);
            }
            return;
        }

        output.AddRange(new byte[] { RecordTemplate, 0, TemplateFraction, 0, 0 });
        EmitLine(children[0], output);
        EmitLine(children[1], output);
        output.Add(RecordEnd);
        output.RestoreContentSize();
    }

    private static void EmitSquareRoot(XElement element, MtefWriter output)
    {
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateRoot, 0, 0 });
        EmitLineContents(SignificantChildren(element), output);
        using var scriptSize = output.EnterScriptSize();
        output.AddRange(new byte[] { RecordLine, LineNull });
        output.Add(RecordEnd);
    }

    private static void EmitNthRoot(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
            throw new InvalidDataException("MathML mroot requires radicand and index.");
        output.AddRange(new byte[] { RecordTemplate, 0, TemplateRoot, 1, 0 });
        EmitLine(children[0], output);
        using var scriptSize = output.EnterScriptSize();
        EmitLine(children[1], output);
        output.Add(RecordEnd);
    }

    private static void EmitScript(XElement element, byte selector, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        var required = selector == TemplateSubSup ? 3 : 2;
        if (children.Length < required)
            throw new InvalidDataException($"MathML {element.Name.LocalName} has too few children.");
        var limitsAbove = LimitsAbove(element);
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(
                integral,
                main: null,
                lower: selector is TemplateSub or TemplateSubSup ? children[1] : null,
                upper: selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output,
                limitsAbove);
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: selector is TemplateSub or TemplateSubSup ? children[1] : null,
                upper: selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output,
                limitsAbove);
            return;
        }
        if (IsNamedLimitOperator(children[0]))
        {
            EmitLimitTemplate(
                children[0],
                selector is TemplateSub or TemplateSubSup ? children[1] : null,
                selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output,
                limitsAbove);
            return;
        }
        if (TryEmitPrimeEmbellishments(element, output))
            return;
        if (limitsAbove)
        {
            EmitUnderOverLimitTemplate(
                children[0],
                selector is TemplateSub or TemplateSubSup ? children[1] : null,
                selector == TemplateSup ? children[1] : selector == TemplateSubSup ? children[2] : null,
                output);
            return;
        }
        // A native leading trailing-script has no base object. Give it its own
        // LINE so it cannot consume a preceding sibling. The TMPL is first in
        // that LINE, just as in genuine MathType storage; a preceding empty LINE
        // would supply a different baseline, and LINE_NULL is not a base at all.
        var emptyBase = IsEmptyScriptBase(children[0]);
        if (emptyBase) output.AddRange(new byte[] { RecordLine, 0 });
        else EmitScriptBase(children[0], output);
        output.AddRange(new byte[] { RecordTemplate, 0, selector, 0, 0 });
        using (output.EnterScriptSize())
        {
            if (selector == TemplateSub)
            {
                EmitLine(children[1], output);
                output.AddRange(new byte[] { RecordLine, LineNull });
            }
            else if (selector == TemplateSup)
            {
                output.AddRange(new byte[] { RecordLine, LineNull });
                EmitLine(children[1], output);
            }
            else
            {
                EmitLine(children[1], output);
                EmitLine(children[2], output);
            }
            output.Add(RecordEnd);
        }
        if (emptyBase) output.Add(RecordEnd);
    }

    private static void EmitOver(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
            return;
        }
        if (TryGetAnnotatedHorizontalBraceBody(children[0], top: true, out var annotatedBraceBody))
        {
            EmitHorizontalFenceTemplate(
                TemplateHorizontalBrace,
                top: true,
                annotatedBraceBody,
                children[1],
                output);
            return;
        }
        if (!IsOverDecoration(element, children[1]))
        {
            // An upper argument is not an accent, even if its glyph resembles
            // one. The common script writer owns both ordinary superscripts
            // and big-operator limits (and their native size/object scopes).
            EmitScript(element, TemplateSup, output);
            return;
        }
        var over = NormalizeAccentMark(children[1].Value.Trim());
        // On one character a fixed accent is the CHAR's embellishment, as
        // MathType stores it. MathType's TeX import also stores a wide one
        // (\widehat, \overline) there, losing the distinction; its templates
        // keep it, and MathType draws both.
        // An accent on a TeX group (\hat{\mathbf{e}}) is MathType's template.
        var embellishmentBody = UnwrapSingleTokenAccentBody(children[0]);
        if (embellishmentBody is not null
            && !MathMlAccentGeometry.IsStretchy(children[1])
            && !IsTexGroup(children[0])
            && TryEmitSingleCharacterEmbellishment(embellishmentBody, over, output)) return;
        switch (over)
        {
            case "\u00AF":
            case "\u203E":
            case "\u2015":
            case "\u02C9":
                EmitSingleSlotTemplate(TemplateOverbar, 0, children[0], output);
                return;
            case "→":
                EmitVectorTemplate(0x02, 0x20D7, children[0], output);
                return;
            case "←":
                EmitVectorTemplate(0x01, 0x20D6, children[0], output);
                return;
            case "↔":
                EmitVectorTemplate(0x03, 0x20E1, children[0], output);
                return;
            case "^":
            case "ˆ":
                EmitGlyphAccentTemplate(TemplateHat, children[0], children[1], output);
                return;
            case "~":
            case "˜":
                EmitGlyphAccentTemplate(TemplateTilde, children[0], children[1], output);
                return;
            case "⌢":
            case "⏜":
                EmitGlyphAccentTemplate(TemplateArc, children[0], children[1], output);
                return;
            case "⏞":
                EmitHorizontalFenceTemplate(
                    TemplateHorizontalBrace,
                    top: true,
                    children[0],
                    annotation: null,
                    output);
                return;
        }
        if (ResolveOverAccent(element, children[1])
            && TryEmitLimitTemplateAccent(children[0], over, output))
            return;
        // A centered accent on a group is a HatBox with a selected glyph, not
        // a trailing superscript. This includes dots, checks, breves and rings,
        // and works for a compound native atom (for example a contour integral)
        // without replacing that atom with a font-dependent fixed CHAR.
        if (ResolveOverAccent(element, children[1])
            && EnumerateBmpScalars(children[1].Value.Trim()).Count() == 1)
        {
            EmitGlyphAccentTemplate(TemplateHat, children[0], children[1], output);
            return;
        }
        EmitScriptBase(children[0], output);
        EmitTrailingScript(children[1], isSubscript: false, output);
    }

    private static void EmitUnder(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 2)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
            return;
        }
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(integral, main: null, lower: children[1], upper: null, output, LimitsAbove(element));
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: children[1],
                upper: null,
                output,
                LimitsAbove(element));
            return;
        }
        if (TryGetAnnotatedHorizontalBraceBody(children[0], top: false, out var annotatedBraceBody))
        {
            EmitHorizontalFenceTemplate(
                TemplateHorizontalBrace,
                top: false,
                annotatedBraceBody,
                children[1],
                output);
            return;
        }
        if (IsNamedLimitOperator(children[0]))
        {
            EmitLimitTemplate(children[0], children[1], upper: null, output, LimitsAbove(element));
            return;
        }
        var under = children[1].Value.Trim();
        // On one character MathType's import stores \underline as the CHAR's
        // underbar embellishment.
        if (under is "_" or "¯" or "‾" or "―" or "ˉ" or "̅" or "̲"
            && UnwrapSingleTokenAccentBody(children[0]) is { } underlined
            && TryEmitEmbellishedCharacter(underlined, EmbellUnderbar, output))
            return;
        if (under is "_" or "¯" or "‾" or "―" or "ˉ" or "̅" or "̲")
        {
            EmitSingleSlotTemplate(TemplateUnderbar, 0, children[0], output);
            return;
        }
        if (under == "⏟")
        {
            EmitHorizontalFenceTemplate(
                TemplateHorizontalBrace,
                top: false,
                children[0],
                annotation: null,
                output);
            return;
        }
        // Use the same object-boundary writer as msub. In particular, a null
        // basis must keep its entire trailing script inside one isolated LINE;
        // otherwise a subsequent read/rewrite can reinterpret it as a left
        // script on the following sibling.
        EmitScript(element, TemplateSub, output);
    }

    private static void EmitUnderOver(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length < 3)
        {
            EmitContainerChildren(element, output, inheritedMathVariant: null);
            return;
        }
        if (TryResolveNativeIntegral(children[0], out var integral))
        {
            EmitIntegralTemplate(
                integral,
                main: null,
                lower: children[1],
                upper: children[2],
                output,
                LimitsAbove(element));
            return;
        }
        if (TryResolveNativeBigOperator(
                children[0],
                out var bigOperator,
                allowAmbiguousUnionIntersection: true))
        {
            EmitBigOperatorTemplate(
                bigOperator,
                main: null,
                lower: children[1],
                upper: children[2],
                output,
                LimitsAbove(element));
            return;
        }
        if (IsNamedLimitOperator(children[0]))
        {
            EmitLimitTemplate(children[0], children[1], children[2], output, LimitsAbove(element));
            return;
        }
        EmitScript(element, TemplateSubSup, output);
    }

    internal readonly struct NativeIntegral
    {
        public NativeIntegral(int variationKind, int integralCount,
            int symbolTypeface = TypefaceSymbol, int adornmentTypeface = TypefaceMtExtra)
        {
            VariationKind = variationKind;
            IntegralCount = integralCount;
            SymbolTypeface = symbolTypeface;
            AdornmentTypeface = adornmentTypeface;
        }

        public int VariationKind { get; }
        public int IntegralCount { get; }
        public int SymbolTypeface { get; }
        public int AdornmentTypeface { get; }
    }

    internal static bool TryResolveNativeIntegral(
        XElement element,
        out NativeIntegral integral)
    {
        integral = default;
        if (element.Name.LocalName != "mo") return false;
        integral = element.Value.Trim() switch
        {
            "∫" => new NativeIntegral(0x01, 1),
            "∬" => new NativeIntegral(0x02, 2),
            "∭" => new NativeIntegral(0x03, 3),
            "∮" => new NativeIntegral(0x05, 1),
            // ∲ ∳ (clockwise contours) are no MathType integral kind: the
            // template draws a plain ∫ with limits and crashes MathType's
            // renderer without them. They are written as operator characters.
            _ => default,
        };
        if (integral.VariationKind != 0
            && int.TryParse((string?)element.Attribute("data-mtef-template-typeface"), out var styled))
            integral = new NativeIntegral(integral.VariationKind, integral.IntegralCount, styled,
                int.TryParse((string?)element.Attribute("data-mtef-adornment-typeface"), out var adornment)
                    ? adornment : TypefaceMtExtra);
        return integral.VariationKind != 0;
    }

    private static void EmitIntegralTemplate(
        NativeIntegral integral,
        XElement? main,
        XElement? lower,
        XElement? upper,
        MtefWriter output,
        bool limitsAbove = false)
    {
        // MathType 7 uses selector 15 for the integral family. Genuine native
        // equations encode lower/upper presence in 0x10/0x20, limits above and
        // below the sign (TeX \limits) in 0x40, and the integral kind in the
        // low nibble (1=single, 2=double, 3=triple, 5=contour).
        // A contour template also requires MathType Extra's loop adornment
        // character before the ordinary integral glyph. Variation 4 with only
        // the integral character makes MathPage access invalid native state.
        // MathType writes a lone integral sign as the character in symbol size,
        // and only the single integral template takes an upper limit: with the
        // upper bit a double/triple/contour template crashes MathType's renderer,
        // so MathType itself uses its integral-operator template for those.
        if (main is null && lower is null && upper is null
            && TryEmitLoneIntegralSign(integral, output))
            return;
        if (upper is not null && integral.VariationKind != 0x01)
        {
            EmitIntegralOperatorTemplate(integral, main, lower, upper, output, limitsAbove);
            return;
        }
        var variation = integral.VariationKind
            | (lower is null ? 0 : 0x10)
            | (upper is null ? 0 : 0x20)
            | (limitsAbove ? 0x40 : 0);
        EmitTemplateHeader(TemplateIntegral, variation, output);
        if (main is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(main, output);
        // Limits step down from the enclosing size; a nested integral uses
        // SUB2 limits and SUBSYM, and returns to its parent's content size.
        var symbolSize = output.SymbolSize;
        using var limitsSize = output.EnterScriptSize();
        if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(lower, output);
        if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(upper, output);
        output.Add(symbolSize);
        EmitIntegralSignCharacters(integral, output);
        output.Add(RecordEnd);
        // Disposing limitsSize restores the enclosing size after the template,
        // not necessarily FULL (this integral itself may be in a script).
    }

    private static void EmitIntegralSignCharacters(NativeIntegral integral, List<byte> output)
    {
        if (integral.VariationKind == 0x05)
        {
            // Exact bytes emitted by genuine MathType 7 for the closed-loop
            // adornment: fnMTEXTRA, MTCode U+EE11, legacy position 0xD1.
            EmitScalar(
                0xEE11,
                integral.AdornmentTypeface,
                output,
                includeEncoded8: true,
                encoded8Override: 0xD1);
        }
        for (var index = 0; index < integral.IntegralCount; index++)
        {
            EmitScalar(
                0x222B,
                integral.SymbolTypeface,
                output,
                includeEncoded8: true,
                encoded8Override: 0xF2);
        }
    }

    // MathType's own storage of a lone integral sign: the character in symbol
    // size. MathType draws the double and contour signs from its far-east face;
    // it has no glyph for the triple sign, which therefore keeps the template.
    private static bool TryEmitLoneIntegralSign(NativeIntegral integral, MtefWriter output)
    {
        if (integral.SymbolTypeface != TypefaceSymbol) return false;
        var symbolSize = output.SymbolSize;
        switch (integral.VariationKind)
        {
            case 0x01:
                output.Add(symbolSize);
                EmitScalar(0x222B, TypefaceSymbol, output, includeEncoded8: true, encoded8Override: 0xF2);
                break;
            case 0x02:
                output.Add(symbolSize);
                EmitScalar(0x222C, TypefaceFarEastText, output, includeEncoded8: false);
                break;
            case 0x05:
                output.Add(symbolSize);
                EmitScalar(0x222E, TypefaceFarEastText, output, includeEncoded8: false);
                break;
            default:
                return false;
        }
        output.RestoreContentSize();
        return true;
    }

    // MathType's integral-operator form for an integral with both limits whose
    // template cannot carry an upper limit: an empty main slot, the limits, the
    // sign in its own slot, and the integrand after the template. Limits sit
    // beside the sign (tmINTOP) unless \limits puts them above/below (tmSUMOP).
    private static void EmitIntegralOperatorTemplate(
        NativeIntegral integral,
        XElement? main,
        XElement? lower,
        XElement? upper,
        MtefWriter output,
        bool limitsAbove)
    {
        var variation = (lower is null ? 0 : 0x10)
            | (upper is null ? 0 : 0x20)
            | (limitsAbove ? 0x40 : 0);
        EmitTemplateHeader(limitsAbove ? TemplateSumOperator : TemplateIntegralOperator, variation, output);
        output.AddRange(new byte[] { RecordLine, LineNull });
        var symbolSize = output.SymbolSize;
        using (output.EnterScriptSize())
        {
            if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
            else EmitLine(lower, output);
            if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
            else EmitLine(upper, output);
            output.Add(symbolSize);
            output.AddRange(new byte[] { RecordLine, 0 });
            // MathType draws the double and contour signs from its far-east
            // face; it has no triple sign, so that one keeps MathType's own
            // template glyphs (three integral signs).
            if (integral.SymbolTypeface == TypefaceSymbol && integral.VariationKind == 0x02)
                EmitScalar(0x222C, TypefaceFarEastText, output, includeEncoded8: false);
            else if (integral.SymbolTypeface == TypefaceSymbol && integral.VariationKind == 0x05)
                EmitScalar(0x222E, TypefaceFarEastText, output, includeEncoded8: false);
            else
                EmitIntegralSignCharacters(integral, output);
            output.Add(RecordEnd); // operator line
            output.Add(RecordEnd); // template
        }
        if (main is not null) EmitNode(main, output);
    }

    internal readonly struct NativeBigOperator
    {
        public NativeBigOperator(
            byte selector,
            int mtCode,
            int typeface,
            byte encoded8)
        {
            Selector = selector;
            MtCode = mtCode;
            Typeface = typeface;
            Encoded8 = encoded8;
        }

        public byte Selector { get; }
        public int MtCode { get; }
        public int Typeface { get; }
        public byte Encoded8 { get; }
    }

    internal static bool TryResolveNativeBigOperator(
        XElement element,
        out NativeBigOperator bigOperator,
        bool allowAmbiguousUnionIntersection = false)
    {
        bigOperator = default;
        if (element.Name.LocalName != "mo") return false;
        var value = element.Value.Trim();
        var explicitBigOperator = string.Equals(
                ((string?)element.Attribute("data-mjx-texclass"))?.Trim(),
                "OP",
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                ((string?)element.Attribute("movablelimits"))?.Trim(),
                "true",
                StringComparison.OrdinalIgnoreCase);
        bigOperator = value switch
        {
            "∑" => new NativeBigOperator(TemplateSum, 0x2211, TypefaceSymbol, 0xE5),
            "∏" => new NativeBigOperator(TemplateProduct, 0x220F, TypefaceSymbol, 0xD5),
            "∐" => new NativeBigOperator(TemplateCoproduct, 0x2210, TypefaceMtExtra, 0x43),
            "⋃" => new NativeBigOperator(TemplateUnion, 0x222A, TypefaceMtExtra, 0x55),
            "⋂" => new NativeBigOperator(TemplateIntersection, 0x2229, TypefaceMtExtra, 0x49),
            "∪" when allowAmbiguousUnionIntersection || explicitBigOperator
                => new NativeBigOperator(TemplateUnion, 0x222A, TypefaceMtExtra, 0x55),
            "∩" when allowAmbiguousUnionIntersection || explicitBigOperator
                => new NativeBigOperator(TemplateIntersection, 0x2229, TypefaceMtExtra, 0x49),
            _ => default,
        };
        if (bigOperator.Selector != 0
            && int.TryParse((string?)element.Attribute("data-mtef-template-typeface"), out var styled))
            bigOperator = new NativeBigOperator(bigOperator.Selector, bigOperator.MtCode, styled, bigOperator.Encoded8);
        return bigOperator.Selector != 0;
    }

    private static void EmitBigOperatorTemplate(
        NativeBigOperator bigOperator,
        XElement? main,
        XElement? lower,
        XElement? upper,
        MtefWriter output,
        bool limitsAbove = true)
    {
        // MathType 7 native BigOp storage observed from genuine Equation.DSMT4:
        // line slots main/lower/upper followed by SYM + CHAR. The variation
        // records which limits exist (0x10 lower, 0x20 upper) and whether they
        // sit above/below the operator (0x40, display) or beside it (inline).
        // A lone operator is MathType's character in symbol size, no template;
        // MathType writes a summation without limits with 0x40, the other
        // operators without limits with 0.
        if (main is null && lower is null && upper is null
            && TryEmitLoneBigOperator(bigOperator, output))
            return;
        // MathType's renderer crashes on an upper limit without a lower one;
        // its own templates write an empty lower slot for that.
        if (lower is null && upper is not null) lower = new XElement("mrow");
        var variation = (lower is null ? 0 : 0x10)
            | (upper is null ? 0 : 0x20)
            | (lower is null && upper is null
                ? bigOperator.Selector == TemplateSum ? 0x40 : 0
                : limitsAbove ? 0x40 : 0);
        EmitTemplateHeader(bigOperator.Selector, variation, output);
        if (main is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(main, output);
        var symbolSize = output.SymbolSize;
        using var limitsSize = output.EnterScriptSize();
        if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(lower, output);
        if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(upper, output);
        output.Add(symbolSize);
        EmitScalar(
            bigOperator.MtCode,
            bigOperator.Typeface,
            output,
            includeEncoded8: true,
            encoded8Override: bigOperator.Encoded8);
        output.Add(RecordEnd);
        // limitsSize restores the parent content size, including nested SUB.
    }

    // A lone big operator as MathType stores it: the operator character in
    // symbol size. Union and intersection use the Symbol font's characters
    // there (their templates use MT Extra's larger glyphs); a styled union or
    // intersection has no such Symbol binding and keeps its template.
    private static bool TryEmitLoneBigOperator(NativeBigOperator bigOperator, MtefWriter output)
    {
        var symbolCharacter = bigOperator.Selector switch
        {
            TemplateUnion => (Scalar: 0x222A, Encoded8: (byte)0xC8),
            TemplateIntersection => (Scalar: 0x2229, Encoded8: (byte)0xC7),
            _ => (Scalar: 0, Encoded8: (byte)0),
        };
        if (symbolCharacter.Scalar != 0 && bigOperator.Typeface != TypefaceMtExtra) return false;
        output.Add(output.SymbolSize);
        if (symbolCharacter.Scalar != 0)
            EmitScalar(symbolCharacter.Scalar, TypefaceSymbol, output,
                includeEncoded8: true, encoded8Override: symbolCharacter.Encoded8);
        else
            EmitScalar(
                bigOperator.MtCode,
                bigOperator.Typeface,
                output,
                includeEncoded8: true,
                encoded8Override: bigOperator.Encoded8);
        output.RestoreContentSize();
        return true;
    }

    // The operators whose MathML dictionary entry has movable limits.
    private const string MovableLimitOperators = "∑∏∐⋀⋁⋂⋃⨀⨁⨂⨄⨆⨉";

    // Where under/over limits sit, by TeX's (and MathML's) rule: an operator
    // with movable limits (the large operators, \lim and the like) has them
    // beside the sign outside display style; integrals, \limits and \underset
    // keep them above and below.
    private static bool LimitsAbove(XElement container)
    {
        if (container.Name.LocalName is not ("munder" or "mover" or "munderover")) return false;
        var operatorElement = container.Elements().FirstOrDefault();
        if (operatorElement is null) return true;
        var movable = ((string?)operatorElement.Attribute("movablelimits"))?.Trim();
        var value = operatorElement.Value.Trim();
        var isMovable = movable is null
            ? operatorElement.Name.LocalName == "mo" && value.Length == 1
                && MovableLimitOperators.IndexOf(value[0]) >= 0
            : string.Equals(movable, "true", StringComparison.OrdinalIgnoreCase);
        return !isMovable || IsDisplayStyle(container);
    }

    // MathML display style at an element: math@display="block" starts it,
    // mstyle and mtable set it, and fraction parts, scripts, limits and root
    // indices are not in it (nor table cells, unless the table says so).
    private static bool IsDisplayStyle(XElement element)
    {
        for (XElement? child = element, parent = element.Parent; parent is not null; child = parent, parent = parent.Parent)
        {
            var displayStyle = ((string?)parent.Attribute("displaystyle"))?.Trim();
            switch (parent.Name.LocalName)
            {
                case "mstyle" when displayStyle is not null:
                case "mtable":
                    return string.Equals(displayStyle, "true", StringComparison.OrdinalIgnoreCase);
                case "mfrac":
                    return false;
                case "msub" or "msup" or "msubsup" or "munder" or "mover" or "munderover"
                    or "mmultiscripts" or "mroot" when child != parent.Elements().FirstOrDefault():
                    return false;
                case "math":
                    return string.Equals((string?)parent.Attribute("display"), "block", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(displayStyle, "true", StringComparison.OrdinalIgnoreCase);
            }
        }
        return false;
    }

    private static bool IsNamedLimitOperator(XElement element)
    {
        if (element.Name.LocalName is not ("mi" or "mo")) return false;
        var value = element.Value.Trim();
        // one or more words of letters (MathJax's \liminf is "lim inf")
        var words = value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (value.Length <= 1 || !words.All(word => word.All(char.IsLetter))) return false;
        var movableLimits = string.Equals(
            (string?)element.Attribute("movablelimits"),
            "true",
            StringComparison.OrdinalIgnoreCase);
        var texClass = ((string?)element.Attribute("data-mjx-texclass") ?? string.Empty).Trim();
        var variant = ((string?)element.Attribute("mathvariant") ?? string.Empty).Trim();
        return movableLimits
            || string.Equals(texClass, "OP", StringComparison.OrdinalIgnoreCase)
            || variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) >= 0
            || variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void EmitLimitTemplate(
        XElement operatorElement,
        XElement? lower,
        XElement? upper,
        MtefWriter output,
        bool limitsAbove = false)
    {
        // Named operators such as max/min/lim/sup/inf are function-style CHAR
        // records. With limits beside them MathType follows the run with the
        // ordinary sub/sup template; in display style it puts the run in the
        // main slot of its limit template, as it does for \lim.
        if (limitsAbove && (lower is not null || upper is not null))
        {
            EmitUnderOverLimitTemplate(operatorElement, lower, upper, output);
            return;
        }
        var name = operatorElement.Value.Trim();
        if (lower is null && upper is null)
        {
            EmitFunctionRun(name, output);
            return;
        }
        // The script binds to the preceding object: a multi-word name
        // ("lim sup") is one LINE so it does not bind to the last word.
        if (name.Any(char.IsWhiteSpace))
        {
            output.AddRange(new byte[] { RecordLine, 0 });
            EmitFunctionRun(name, output);
            output.Add(RecordEnd);
        }
        else EmitFunctionRun(name, output);

        var selector = lower is not null && upper is not null
            ? TemplateSubSup
            : lower is not null
                ? TemplateSub
                : TemplateSup;
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            selector,
            0,
            0,
        });
        using var scriptSize = output.EnterScriptSize();
        if (lower is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitLine(lower, output);
        if (upper is null)
            output.AddRange(new byte[] { RecordLine, LineNull });
        else
            EmitLine(upper, output);
        output.Add(RecordEnd);
    }

    // MathType's limit template (tmLIM): a base with limits centred below
    // and/or above it. Both of MathType's importers store \underset, \overset
    // and \stackrel this way: main LINE, then the lower and upper LINEs in
    // script size.
    private static void EmitUnderOverLimitTemplate(
        XElement baseElement,
        XElement? lower,
        XElement? upper,
        MtefWriter output)
    {
        EmitTemplateHeader(TemplateLimit, (lower is null ? 0 : 0x10) | (upper is null ? 0 : 0x20), output);
        if (IsEmptyScriptBase(baseElement)) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(baseElement, output);
        using (output.EnterScriptSize())
        {
            if (lower is null) output.AddRange(new byte[] { RecordLine, LineNull });
            else EmitLine(lower, output);
            if (upper is null) output.AddRange(new byte[] { RecordLine, LineNull });
            else EmitLine(upper, output);
            output.Add(RecordEnd);
        }
    }

    // MathType's own storage of \check, \acute, \grave and \mathring, and of
    // \breve on more than one character (both of its importers): the limit
    // template with the accent character in the upper slot, in the Function
    // face (the ring in MT Extra).
    private static bool TryEmitLimitTemplateAccent(XElement body, string mark, MtefWriter output)
    {
        // A dot or double dot on a group: one or two Function-face periods.
        var (scalar, typeface, encoded8, count) = mark switch
        {
            "ˇ" => (0x02C7, TypefaceFunction, (byte?)null, 1),
            "´" => (0x00B4, TypefaceFunction, (byte?)null, 1),
            "`" => (0x0060, TypefaceFunction, (byte?)null, 1),
            "˘" => (0x02D8, TypefaceFunction, (byte?)null, 1),
            "˚" => (0x02DA, TypefaceMtExtra, (byte?)0xFE, 1),
            "˙" => (0x002E, TypefaceFunction, (byte?)null, 1),
            "¨" => (0x002E, TypefaceFunction, (byte?)null, 2),
            _ => (0, 0, (byte?)null, 0),
        };
        if (scalar == 0) return false;
        EmitTemplateHeader(TemplateLimit, 0x20, output);
        if (IsEmptyScriptBase(body)) output.AddRange(new byte[] { RecordLine, LineNull });
        else EmitLine(body, output);
        using (output.EnterScriptSize())
        {
            output.AddRange(new byte[] { RecordLine, LineNull });
            output.AddRange(new byte[] { RecordLine, 0 });
            for (var index = 0; index < count; index++)
                EmitScalar(scalar, typeface, output, includeEncoded8: encoded8 is not null, encoded8Override: encoded8);
            output.Add(RecordEnd);
            output.Add(RecordEnd);
        }
        return true;
    }

    private static void EmitFunctionRun(string name, List<byte> output)
    {
        // A multi-word name (MathJax's \liminf is "lim inf") is one function
        // run per word with a native thin space between them: a space CHAR in
        // the function face would make MathType read the whole name as text.
        var words = name.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        for (var word = 0; word < words.Length; word++)
        {
            if (word > 0) EmitMathSpace(new XElement("mspace", new XAttribute("width", "0.167em")), output);
            var functionScalars = EnumerateBmpScalars(words[word]).ToArray();
            for (var index = 0; index < functionScalars.Length; index++)
            {
                EmitScalar(
                    functionScalars[index],
                    TypefaceFunction,
                    output,
                    functionStart: index == 0);
            }
        }
    }

    // MathType has no invisible content; a phantom keeps only its room. Zero
    // width (\vphantom, \mathstrut) writes nothing; otherwise native spaces of
    // roughly the content's width (half an em per visible character).
    private static void EmitPhantomSpace(XElement element, List<byte> output)
    {
        if (element.Parent is { } padded && padded.Name.LocalName == "mpadded"
            && padded.Attribute("width") is { } width
            && IsZeroMathMlLength(width.Value))
            return;
        var em = 0.5 * element.Descendants()
            .Where(token => !token.HasElements && token.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
            .Sum(token => token.Value.Trim().Length);
        for (; em >= 1; em -= 1)
            EmitMathSpace(new XElement("mspace", new XAttribute("width", "1em")), output);
        if (em >= 0.25)
            EmitMathSpace(new XElement("mspace", new XAttribute("width", "0.278em")), output);
    }

    private static void EmitTrailingScript(
        XElement script,
        bool isSubscript,
        MtefWriter output)
    {
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            isSubscript ? TemplateSub : TemplateSup,
            0,
            0,
        });
        using var scriptSize = output.EnterScriptSize();
        if (isSubscript)
        {
            EmitLine(script, output);
            output.AddRange(new byte[] { RecordLine, LineNull });
        }
        else
        {
            output.AddRange(new byte[] { RecordLine, LineNull });
            EmitLine(script, output);
        }
        output.Add(RecordEnd);
    }

    private static void EmitSingleSlotTemplate(
        byte selector,
        int variation,
        XElement body,
        MtefWriter output)
    {
        EmitTemplateHeader(selector, variation, output);
        EmitLine(body, output);
        output.Add(RecordEnd);
        output.RestoreContentSize();
    }

    private static void EmitGlyphAccentTemplate(
        byte selector, XElement body, XElement sourceMark, MtefWriter output)
    {
        // HatBoxClass owns a main LINE and a trailing glyph CHAR. A one-slot
        // template is incomplete native data even when our reader can infer a
        // plausible mark from its selector. The selected CHAR is authoritative:
        // tmHAT can carry a fixed dot/check/etc. as well as a stretching hat.
        EmitTemplateHeader(selector, 0, output);
        EmitLine(body, output);
        var mark = NormalizeAccentMark(sourceMark.Value.Trim());
        // On a group holding one (styled) character MathType's import also
        // uses its expanding glyph (\hat{\mathbf{e}}).
        var expanding = MathMlAccentGeometry.IsStretchy(sourceMark)
            || IsTexGroup(body) && UnwrapSingleTokenAccentBody(body) is not null;
        var combining = mark switch { "^" => 0x0302, "~" => 0x0303, "⌢" or "⏜" => 0x2322, _ => 0 };
        if (expanding && combining != 0)
            EmitScalar(combining, TypefaceFence, output);
        else
            EmitOperator(sourceMark, output);
        output.Add(RecordEnd);
        output.RestoreContentSize();
    }

    private static void EmitVectorTemplate(
        int variation,
        int combiningArrow,
        XElement body,
        MtefWriter output)
    {
        // Genuine MathType 7 tmVEC HatBox records are not a generic one-slot
        // template. After the body line MathType persists the selected combining
        // arrow as an explicit MTCode character using its expanding-glyph
        // typeface. Omitting this trailing character makes Word accept the CFB
        // initially but reject/rematerialize its OLE presentation for multi-token
        // vectors such as \overrightarrow{AB} / \overleftrightarrow{AB}.
        EmitTemplateHeader(TemplateVector, variation, output);
        EmitColor(0, output);
        EmitLine(body, output);
        EmitScalar(
            combiningArrow,
            TypefaceFence,
            output,
            includeEncoded8: false);
        output.Add(RecordEnd);
        output.RestoreContentSize();
    }

    private static bool TryGetAnnotatedHorizontalBraceBody(
        XElement candidate,
        bool top,
        out XElement body)
    {
        body = null!;
        var current = candidate;
        while (current.Name.LocalName == "mrow")
        {
            var wrapped = current.Elements().ToArray();
            if (wrapped.Length != 1) break;
            current = wrapped[0];
        }
        if (current.Name.LocalName != (top ? "mover" : "munder")) return false;
        var inner = current.Elements().ToArray();
        if (inner.Length < 2 || inner[1].Name.LocalName != "mo") return false;
        var marker = inner[1].Value.Trim();
        var expected = top ? "⏞" : "⏟";
        if (!IsHorizontalBraceMarker(marker, top)) return false;
        body = inner[0];
        return true;
    }

    private static void EmitHorizontalFenceTemplate(
        byte selector,
        bool top,
        XElement body,
        XElement? annotation,
        MtefWriter output)
    {
        EmitTemplateHeader(selector, top ? 1 : 0, output);

        // Match MathType 7's native HorizontalBrace/HorizontalBracket layout.
        // A real Equation.DSMT4 stores the main body at full size, switches to
        // SUB for the annotation, restores FULL, then writes the private fence
        // glyph using the expanding-fence typeface. Writing the public Unicode
        // U+23DE/U+23DF operators here produces an MTEF stream that VisualTeX can
        // parse itself but that MathType may reject while activating the OLE.
        EmitColor(0, output);
        EmitMathTypeColoredLine(body, output, resetColorAfter: false);
        using (output.EnterScriptSize())
        {
            EmitColor(0, output);
            // MathType keeps an unlabelled brace's label slot as an empty LINE
            // (not LINE_NULL), which also places the brace.
            if (annotation is null)
                output.AddRange(new byte[] { RecordLine, 0, RecordEnd });
            else
                EmitMathTypeColoredLine(annotation, output, resetColorAfter: false);
        }
        EmitScalar(
            top ? 0xFE37 : 0xFE38,
            TypefaceFence,
            output,
            includeEncoded8: false);
        output.Add(RecordEnd);
    }

    private static void EmitTemplateHeader(byte selector, int variation, List<byte> output)
    {
        output.Add(RecordTemplate);
        output.Add(0);
        output.Add(selector);
        if (variation < 0x80)
        {
            output.Add((byte)variation);
        }
        else
        {
            output.Add((byte)(0x80 | (variation & 0x7F)));
            output.Add((byte)(variation >> 8));
        }
        output.Add(0);
    }

    private static XElement? UnwrapSingleTokenAccentBody(XElement body)
    {
        var current = body;
        while (current.Name.LocalName == "mrow")
        {
            var children = current.Elements().ToArray();
            if (children.Length != 1) break;
            current = children[0];
        }
        return current.Name.LocalName is "mi" or "mn" or "mo" ? current : null;
    }

    internal static string NormalizeAccentMark(string mark) => mark switch
    {
        "\u0305" or "\u203E" or "\u2015" or "\u02C9" => "¯",
        "\u0302" or "ˆ" => "^",
        "\u0303" or "˜" => "~",
        "\u20D7" => "→",
        "\u20D6" => "←",
        "\u20E1" => "↔",
        "." or "\u0307" => "˙",
        "\u0308" => "¨",
        "\u030C" => "ˇ",
        "\u0306" => "˘",
        "\u0301" => "´",
        "\u0300" => "`",
        "\u030A" => "˚",
        _ => mark,
    };

    private static bool TryEmitSingleCharacterEmbellishment(
        XElement body,
        string over,
        MtefWriter output)
    {
        var code = NormalizeAccentMark(over) switch
        {
            "." or "˙" => EmbellDot,
            "¨" => EmbellDoubleDot,
            "~" or "˜" => EmbellTilde,
            "^" or "ˆ" => EmbellHat,
            "→" => EmbellRightArrow,
            "←" => EmbellLeftArrow,
            "↔" => EmbellBothArrow,
            "\u00AF" or "\u203E" or "\u2015" or "\u02C9" => EmbellOverbar,
            "⃛" => EmbellTripleDot,
            "⃜" => EmbellQuadrupleDot,
            "˘" => EmbellBreve,
            _ => (byte)0,
        };
        return TryEmitEmbellishedCharacter(body, code, output);
    }

    // MathJax's form of a TeX group ({...}, \mathbf{...}): MathType's TeX import
    // treats an accent or prime on a group as a template, not an embellishment.
    private static bool IsTexGroup(XElement element) =>
        element.Name.LocalName == "mrow"
        && string.Equals((string?)element.Attribute("data-mjx-texclass"), "ORD", StringComparison.Ordinal);

    // MathType's import stores f', f'' and f''' as the character's prime
    // embellishments (a bare prime superscript; f^{\prime} stays a script).
    private static bool TryEmitPrimeEmbellishments(XElement script, MtefWriter output)
    {
        var children = script.Elements().ToArray();
        // On a fence or operator ((f+g)') MathType keeps the superscript.
        if (script.Name.LocalName != "msup" || children.Length != 2
            || children[1].Name.LocalName != "mo" || children[0].Name.LocalName is not ("mi" or "mn"))
            return false;
        var primes = children[1].Value.Trim();
        if (primes.Length == 0) return false;
        var codes = new List<byte>();
        foreach (var prime in primes)
        {
            codes.Add(prime switch
            {
                '′' => EmbellPrime,
                '″' => EmbellDoublePrime,
                '‴' => EmbellTriplePrime,
                _ => (byte)0,
            });
        }
        return codes.All(code => code != 0) && TryEmitEmbellishedCharacter(children[0], codes, output);
    }

    private static bool TryEmitEmbellishedCharacter(
        XElement body,
        byte code,
        MtefWriter output) =>
        TryEmitEmbellishedCharacter(body, new[] { code }, output);

    private static bool TryEmitEmbellishedCharacter(
        XElement body,
        IReadOnlyList<byte> codes,
        MtefWriter output)
    {
        if (codes.Count == 0 || codes.Any(code => code == 0) || body.Name.LocalName is not ("mi" or "mn" or "mo"))
            return false;
        var scalars = EnumerateBmpScalars(body.Value).ToArray();
        if (scalars.Length != 1) return false;

        // Decoration changes only the CHAR's EMBELL list. First emit the bare
        // token through its canonical character path, including every explicit
        // font, native MTCode, encoded8 position and function-start option.
        // Re-resolving just a typeface here loses special operator glyphs as well
        // as styled identifiers; keeping their bytes makes the invariant apply
        // to all existing and future character mappings, not an accent whitelist.
        var character = output.CreateScratch();
        // Use the actual ordinary-object writer, not the fixed-glyph shortcut.
        // A large/composite operator owns a template, possibly multiple styled
        // glyphs. It must receive a group accent around that same native object.
        EmitNode(body, character);
        if (character.Count < 5 || character[0] != RecordChar)
            return false;
        var options = character[1];
        // Token emitters produce one unnudged CHAR. Do not decorate a template,
        // a run, or a pre-decorated object by guessing where its bytes end.
        if ((options & (0x08 | CharHasEmbellishment)) != 0) return false;
        var end = 2 + (character[2] == 0xFF ? 3 : 1);
        if ((options & 0x20) == 0) end += 2;
        if ((options & CharEncoded8) != 0) end++;
        if ((options & 0x10) != 0) end += 2;
        if (end != character.Count) return false;
        character[1] |= CharHasEmbellishment;
        output.AddRange(character);
        foreach (var code in codes) output.AddRange(new byte[] { RecordEmbellishment, 0, code });
        output.Add(RecordEnd);
        return true;
    }

    internal static bool MathTypeEmissionRequiresColorDefinition(
        XElement preparedMath,
        bool topLevelTableIsAlignedPile)
    {
        var topLevelElements = SignificantChildren(preparedMath)
            .OfType<XElement>()
            .Where(element => element.Name.LocalName is not ("annotation" or "annotation-xml"))
            .ToArray();
        XElement? alignedTopLevelTable =
            topLevelTableIsAlignedPile
            && topLevelElements.Length == 1
            && topLevelElements[0].Name.LocalName == "mtable"
                ? topLevelElements[0]
                : null;

        // EmitMatrix() is the only mtable path that writes COLOR 1. A top-level
        // table emitted through TryEmitAlignedPile() must be excluded; any nested
        // mtable still goes through EmitMatrix().
        if (preparedMath.DescendantsAndSelf()
            .Any(element => element.Name.LocalName == "mtable"
                && !ReferenceEquals(element, alignedTopLevelTable)))
            return true;

        // Native horizontal brace/bracket templates also use colored LINE slots.
        foreach (var element in preparedMath.Descendants())
        {
            if (element.Name.LocalName == "mover")
            {
                var children = element.Elements().ToArray();
                if (children.Length >= 2
                    && (IsHorizontalBraceMarker(children[1].Value.Trim(), top: true)
                        || TryGetAnnotatedHorizontalBraceBody(
                            children[0],
                            top: true,
                            out _)))
                    return true;
            }
            else if (element.Name.LocalName == "munder")
            {
                var children = element.Elements().ToArray();
                if (children.Length >= 2
                    && (IsHorizontalBraceMarker(children[1].Value.Trim(), top: false)
                        || TryGetAnnotatedHorizontalBraceBody(
                            children[0],
                            top: false,
                            out _)))
                    return true;
            }
        }
        return false;
    }

    internal static bool TryEmitAlignedPile(
        XElement element,
        byte[] sourceMtef,
        MtefWriter output)
    {
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length == 0) return false;

        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        var columnCount = cells.Max(row => row.Length);
        if (columnCount < 2 || (columnCount & 1) != 0) return false;

        var columnAlignment = ((string?)element.Attribute("columnalign") ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (columnAlignment.Length < columnCount) return false;
        for (var column = 0; column < columnCount; column++)
        {
            var expected = (column & 1) == 0 ? "right" : "left";
            if (!string.Equals(
                    columnAlignment[column],
                    expected,
                    StringComparison.OrdinalIgnoreCase))
                return false;
        }

        var pairCount = columnCount / 2;
        var rulerText = ((string?)element.Attribute(VisualTexMtefRulerStopsAttribute) ?? string.Empty)
            .Trim();
        var rulerStops = rulerText.Length == 0
            ? Array.Empty<int>()
            : rulerText.Split(',')
                .Select(value => int.TryParse(
                    value.Trim(),
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out var parsed)
                    ? parsed
                    : -1)
                .ToArray();
        var hasExactRuler = rulerStops.Length == pairCount
            && rulerStops.All(stop => stop > 0 && stop <= ushort.MaxValue)
            && rulerStops.Zip(rulerStops.Skip(1), (left, right) => right > left).All(valid => valid);

        // MathType's true multi-point alignment has two distinct fnMARKER
        // primitives. U+0009 starts a tab group; MTCode 0xEF00 is MathType's
        // non-printing alignment symbol inside that group. With an explicit
        // RULER, each group is positioned so its alignment symbol lands exactly
        // on the supplied stop. This is the native equivalent of one TeX
        // right/left `&` pair. A plain sequence of Ctrl+Tab markers is not enough:
        // MathType's default stops are 0.5in apart, so a wide left operand can
        // skip a stop and visibly misalign one row.
        output.AddRange(new byte[]
        {
            RecordPile,
            hasExactRuler ? (byte)0x02 : (byte)0,
            1, // horizontal alignment: left; tab groups own their reference points
            1, // vertical alignment: center line
        });
        if (hasExactRuler)
        {
            // LP_RULER is an inline field of LINE/PILE. MathType's native MTEF
            // readers consume n_stops immediately after the alignment bytes; the
            // enclosing option bit already identifies the payload as a RULER.
            // Emitting an additional standalone record tag (7) here shifts every
            // following byte and MathPage renders the equation blank.
            output.Add((byte)rulerStops.Length);
            foreach (var stop in rulerStops)
            {
                output.Add(0); // left tab; an in-group alignment symbol overrides it
                output.Add((byte)(stop & 0xff));
                output.Add((byte)((stop >> 8) & 0xff));
            }
        }

        // PILE is the root structure. Keep only root formatting state here;
        // missing font/encoding definitions are emitted globally before the
        // initial SIZE record by BuildRootStructure/AssembleRewrittenMtef.
        CopyRootLeadingFormattingRecords(sourceMtef, output);
        foreach (var row in cells)
        {
            output.AddRange(new byte[] { RecordLine, 0 });
            if (hasExactRuler)
            {
                for (var pair = 0; pair < pairCount; pair++)
                {
                    // Start the group at the next explicit ruler stop, then place
                    // the TeX pair boundary at MathType's own alignment symbol.
                    EmitScalar('\t', TypefaceMarker, output, includeEncoded8: false);
                    var leftColumn = pair * 2;
                    var rightColumn = leftColumn + 1;
                    if (leftColumn < row.Length)
                        EmitContainerChildren(row[leftColumn], output, inheritedMathVariant: null);
                    EmitScalar(
                        MathTypeAlignmentMarkerMtCode,
                        TypefaceMarker,
                        output,
                        includeEncoded8: false);
                    if (rightColumn < row.Length)
                        EmitContainerChildren(row[rightColumn], output, inheritedMathVariant: null);
                }
            }
            else
            {
                // Compatibility fallback for old MathML/third-party calls that
                // lack VisualTeX's SVG-derived ruler metadata. It preserves all
                // tab groups semantically, but production Office MathType exports
                // are expected to carry an exact ruler and use the branch above.
                for (var column = 0; column < columnCount; column++)
                {
                    if (column > 0)
                        EmitScalar('\t', TypefaceMarker, output, includeEncoded8: false);
                    if (column < row.Length)
                        EmitContainerChildren(row[column], output, inheritedMathVariant: null);
                }
            }
            output.Add(RecordEnd);
        }
        output.Add(RecordEnd);
        return true;
    }

    private static void EmitMatrix(XElement element, MtefWriter output)
    {
        var rows = element.Elements()
            .Where(row => row.Name.LocalName is "mtr" or "mlabeledtr")
            .ToArray();
        if (rows.Length == 0)
        {
            output.AddRange(new byte[] { RecordMatrix, 0, 1, 2, 1, 1, 1, 0, 0 });
            output.AddRange(new byte[] { RecordLine, LineNull, RecordEnd });
            output.RestoreContentSize();
            return;
        }
        var cells = rows.Select(row => row.Elements()
                .Where(cell => cell.Name.LocalName == "mtd")
                .ToArray())
            .ToArray();
        var columnCount = Math.Max(1, cells.Max(row => row.Length));
        if (rows.Length > byte.MaxValue || columnCount > byte.MaxValue)
            throw new InvalidDataException("MathType MTEF matrix exceeds 255 rows or columns.");

        // One horizontal justification for every cell: MathType left-justifies
        // \begin{cases} (MathML columnalign="left left"), and centres by default.
        var columnAlign = ((string?)element.Attribute("columnalign") ?? string.Empty)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var horizontalJustification = columnAlign.Length == 1
            ? columnAlign[0].ToLowerInvariant() switch { "left" => (byte)0, "right" => (byte)2, _ => (byte)1 }
            : (byte)1;
        output.AddRange(new byte[]
        {
            RecordMatrix,
            0,
            1, // vertical alignment: center line
            horizontalJustification,
            1, // vertical cell justification: baseline center
            (byte)rows.Length,
            (byte)columnCount,
        });
        for (var index = 0; index < PartitionByteCount(rows.Length + 1); index++) output.Add(0);
        for (var index = 0; index < PartitionByteCount(columnCount + 1); index++) output.Add(0);
        for (var row = 0; row < rows.Length; row++)
        {
            for (var column = 0; column < columnCount; column++)
            {
                var isLastCell = row == rows.Length - 1 && column == columnCount - 1;
                if (column < cells[row].Length)
                    EmitMathTypeColoredLine(cells[row][column], output, resetColorAfter: !isLastCell);
                else output.AddRange(new byte[] { RecordLine, LineNull });
            }
        }
        output.Add(RecordEnd);
    }

    internal static int PartitionByteCount(int partitionCount) =>
        Math.Max(1, (partitionCount + 3) / 4);

    private static void EmitEnclose(XElement element, MtefWriter output)
    {
        var notation = ((string?)element.Attribute("notation") ?? string.Empty)
            .ToLowerInvariant();
        var significant = SignificantChildren(element).ToArray();
        var body = significant.Length == 1 && significant[0] is XElement singleElement
            ? new XElement(singleElement)
            : new XElement("mrow", significant);
        if (notation.Contains("radical"))
        {
            EmitSquareRoot(new XElement("msqrt", new XElement(body)), output);
            return;
        }
        if (notation.Contains("box"))
        {
            EmitSingleSlotTemplate(TemplateBox, 0x1E, body, output);
            return;
        }
        var strikeVariation = 0;
        if (notation.Contains("horizontalstrike")) strikeVariation |= 0x01;
        if (notation.Contains("updiagonalstrike")) strikeVariation |= 0x02;
        if (notation.Contains("downdiagonalstrike")) strikeVariation |= 0x04;
        if (strikeVariation != 0)
        {
            EmitSingleSlotTemplate(TemplateStrike, strikeVariation, body, output);
            return;
        }
        EmitNode(body, output);
    }

    private static void EmitMultiScripts(XElement element, MtefWriter output)
    {
        var children = element.Elements().ToArray();
        if (children.Length == 0) return;

        var isolateEmpty = IsEmptyScriptBase(children[0]);
        if (isolateEmpty) output.AddRange(new byte[] { RecordLine, 0 });
        var prescriptsIndex = Array.FindIndex(
            children,
            child => child.Name.LocalName == "mprescripts");
        if (prescriptsIndex <= 0) prescriptsIndex = children.Length;

        // MTEF v5 tvSU_PRECEDES=1 makes a script template consume the next
        // object, rather than the preceding object. Emit prescript pairs from
        // outermost to innermost so the pair nearest the base is immediately
        // before it. MathML stores those pairs nearest-first.
        var prescriptCount = children.Length - prescriptsIndex - 1;
        var lastPrescriptPair = prescriptsIndex + 1
            + Math.Max(0, (prescriptCount - 1) / 2) * 2;
        for (var preIndex = lastPrescriptPair;
             prescriptCount > 0 && preIndex > prescriptsIndex;
             preIndex -= 2)
        {
            EmitScriptPair(
                children[preIndex],
                preIndex + 1 < children.Length ? children[preIndex + 1] : null,
                precedes: true,
                output);
        }

        EmitScriptBase(children[0], output);
        var index = 1;
        while (index < prescriptsIndex)
        {
            var sub = children[index];
            var sup = index + 1 < prescriptsIndex ? children[index + 1] : null;
            if (sub.Name.LocalName == "none" && (sup is null || sup.Name.LocalName == "none"))
            {
                index += 2;
                continue;
            }
            if (sub.Name.LocalName == "none" && sup is not null)
                EmitTrailingScript(sup, isSubscript: false, output);
            else if (sup is null || sup.Name.LocalName == "none")
                EmitTrailingScript(sub, isSubscript: true, output);
            else
            {
                output.AddRange(new byte[] { RecordTemplate, 0, TemplateSubSup, 0, 0 });
                using (output.EnterScriptSize())
                {
                    EmitLine(sub, output);
                    EmitLine(sup, output);
                    output.Add(RecordEnd);
                }
            }
            index += 2;
        }
        if (isolateEmpty) output.Add(RecordEnd);
    }

    private static void EmitScriptPair(
        XElement? sub,
        XElement? sup,
        bool precedes,
        MtefWriter output)
    {
        var hasSub = sub is not null && sub.Name.LocalName != "none";
        var hasSup = sup is not null && sup.Name.LocalName != "none";
        if (!hasSub && !hasSup) return;

        var selector = hasSub && hasSup ? TemplateSubSup : hasSub ? TemplateSub : TemplateSup;
        output.AddRange(new byte[]
        {
            RecordTemplate,
            0,
            selector,
            precedes ? (byte)1 : (byte)0,
            0,
        });
        using var scriptSize = output.EnterScriptSize();
        if (hasSub) EmitLine(sub!, output);
        else output.AddRange(new byte[] { RecordLine, LineNull });
        if (hasSup) EmitLine(sup!, output);
        else output.AddRange(new byte[] { RecordLine, LineNull });
        output.Add(RecordEnd);
    }

    private static void EmitFenced(XElement element, MtefWriter output)
    {
        var open = NormalizeFence((string?)element.Attribute("open") ?? "(");
        var close = NormalizeFence((string?)element.Attribute("close") ?? ")");
        if ((open, close) is ("⟨", "|") or ("|", "⟩"))
        {
            // tmDIRAC owns two LINE slots and one shared vertical fence.
            // Missing halves still occupy a null LINE; ordinary CHARs lose
            // both scalable geometry and the structural bra/ket boundary.
            var left = open == "⟨";
            EmitTemplateHeader(TemplateDirac, left ? 1 : 2, output);
            EmitColor(0, output);
            if (left) EmitLineContents(SignificantChildren(element), output);
            else output.AddRange(new byte[] { RecordLine, LineNull });
            if (!left) EmitLineContents(SignificantChildren(element), output);
            else output.AddRange(new byte[] { RecordLine, LineNull });
            if (left) EmitFenceCharacter("⟨", opening: true, output);
            EmitFenceCharacter("|", opening: true, output);
            if (!left) EmitFenceCharacter("⟩", opening: false, output);
            output.Add(RecordEnd);
            return;
        }
        var selector = SelectFenceTemplate(open, close);
        if (selector is null)
        {
            if (!string.IsNullOrEmpty(open)) EmitOperator(open, output);
            foreach (var child in SignificantChildren(element)) EmitNode(child, output);
            if (!string.IsNullOrEmpty(close)) EmitOperator(close, output);
            return;
        }
        // tmINTERVAL uses two independent two-bit endpoint selectors, unlike
        // the left/right presence bits of the other ParBox templates (MTEF v5).
        var variation = selector == TemplateInterval
            ? IntervalEndpoint(open) | (IntervalEndpoint(close) << 4)
            : (string.IsNullOrEmpty(open) ? 0 : 1) | (string.IsNullOrEmpty(close) ? 0 : 2);
        EmitTemplateHeader(selector.Value, variation, output);
        EmitColor(0, output);
        EmitLineContents(SignificantChildren(element), output);
        if (!string.IsNullOrEmpty(open)) EmitFenceCharacter(open, opening: true, output);
        if (!string.IsNullOrEmpty(close)) EmitFenceCharacter(close, opening: false, output);
        output.Add(RecordEnd);
    }

    private static void EmitColor(int index, MtefWriter output) =>
        output.EmitNativeColor(index);

    private static void EmitMathTypeColoredLine(
        XElement element,
        MtefWriter output,
        bool resetColorAfter)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        EmitColor(1, output);
        EmitNode(element, output);
        output.Add(RecordEnd);
        if (resetColorAfter) EmitColor(0, output);
    }

    private static void EmitFenceCharacter(
        string value,
        bool opening,
        List<byte> output)
    {
        foreach (var scalar in EnumerateBmpScalars(value))
        {
            // MathType's fnFENCE typeface is MTCode-encoded rather than a
            // generic Unicode font. These values are captured from genuine
            // MathType 7 Equation Native streams for the same expandable
            // delimiters. Some delimiters (| and ||) use the same Unicode text
            // on both sides but distinct left/right MTCode glyphs, so the side
            // of the fence is part of the mapping.
            var mtCode = scalar switch
            {
                '(' => 0x0028,
                ')' => 0x0029,
                '[' => 0x005B,
                ']' => 0x005D,
                '{' => 0x007B,
                '}' => 0x007D,
                0x27E8 => 0x2329, // \langle
                0x27E9 => 0x232A, // \rangle
                '|' => opening ? 0xEC07 : 0xEC08,
                0x2016 => opening ? 0xEC09 : 0xEC0A, // \| / norm bars
                0x230A => 0xF8F0, // \lfloor
                0x230B => 0xF8FB, // \rfloor
                0x2308 => 0xF8EE, // \lceil
                0x2309 => 0xF8F9, // \rceil
                _ => scalar,
            };
            EmitScalar(mtCode, TypefaceFence, output, includeEncoded8: false);
        }
    }

    internal static string NormalizeFence(string value) => value switch
    {
        "." => string.Empty,
        "〈" or "〈" or "⟨" => "⟨",
        "〉" or "〉" or "⟩" => "⟩",
        _ => value,
    };

    internal static byte? SelectFenceTemplate(string open, string close)
    {
        if ((open, close) is ("⟨", "|") or ("|", "⟩")) return TemplateDirac;
        var probe = !string.IsNullOrEmpty(open) ? open : close;
        if ((open is "⟨" or "〈" || string.IsNullOrEmpty(open))
            && (close is "⟩" or "〉" || string.IsNullOrEmpty(close))
            && probe is "⟨" or "〈" or "⟩" or "〉") return TemplateAngle;
        if ((open == "(" || string.IsNullOrEmpty(open))
            && (close == ")" || string.IsNullOrEmpty(close))
            && probe is "(" or ")") return TemplateParen;
        if ((open == "[" || string.IsNullOrEmpty(open))
            && (close == "]" || string.IsNullOrEmpty(close))
            && probe is "[" or "]") return TemplateBracket;
        if ((open == "{" || string.IsNullOrEmpty(open))
            && (close == "}" || string.IsNullOrEmpty(close))
            && probe is "{" or "}") return TemplateBrace;
        if ((open == "|" || string.IsNullOrEmpty(open))
            && (close == "|" || string.IsNullOrEmpty(close))
            && probe == "|") return TemplateBar;
        if ((open is "‖" or "||" || string.IsNullOrEmpty(open))
            && (close is "‖" or "||" || string.IsNullOrEmpty(close))
            && probe is "‖" or "||") return TemplateDoubleBar;
        // MathType's renderer crashes on a one-sided floor or ceiling template
        // and on every interval template but these five endpoint pairs (MTEF
        // can encode all sixteen); the others are written as plain characters.
        if (open == "⌊" && close == "⌋") return TemplateFloor;
        if (open == "⌈" && close == "⌉") return TemplateCeiling;
        if ((open, close) is ("(", "]") or ("[", ")") or ("[", "[") or ("]", "[") or ("]", "]"))
            return TemplateInterval;
        return null;
    }

    private static int IntervalEndpoint(string value) => value switch
    {
        "(" => 0, ")" => 1, "[" => 2, "]" => 3, _ => -1,
    };

    private static void EmitLine(XElement element, MtefWriter output)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        EmitNode(element, output);
        output.Add(RecordEnd);
    }

    private static void EmitLineContents(IEnumerable<XNode> nodes, MtefWriter output)
    {
        output.AddRange(new byte[] { RecordLine, 0 });
        foreach (var node in nodes) EmitNode(node, output);
        output.Add(RecordEnd);
    }
}

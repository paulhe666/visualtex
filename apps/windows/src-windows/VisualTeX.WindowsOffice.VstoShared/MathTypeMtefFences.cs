using System.Text;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefSignature;
using static VisualTeX.WordVsto.MathTypeMtefTemplates;

namespace VisualTeX.WordVsto;

/// <summary>
/// MathML reshaped the way MathType stores it before writing: inherited math variants, loose fences and binomials,
/// MathJax fence rows, split fenced scripts and over accents.
/// </summary>
internal static class MathTypeMtefFences
{
    internal static void MaterializeInheritedMathVariants(
        XElement element,
        string? inheritedMathVariant)
    {
        var ownVariant = ((string?)element.Attribute("mathvariant"))?.Trim();
        var variant = string.IsNullOrWhiteSpace(ownVariant)
            ? inheritedMathVariant
            : ownVariant;
        if (element.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
        {
            if (element.Attribute("mathvariant") is null
                && !string.IsNullOrWhiteSpace(variant))
                element.SetAttributeValue("mathvariant", variant);
            return;
        }
        foreach (var child in element.Elements())
            MaterializeInheritedMathVariants(child, variant);
    }

    internal static void NormalizeMathJaxFenceRows(XElement root)
    {
        var rows = root.DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "mrow")
            .Reverse()
            .ToArray();
        foreach (var row in rows)
        {
            if (!TryGetMathJaxFenceRow(row, out var open, out var close, out var children))
                continue;
            // Recognizing a source fence does not establish a native template.
            // Keep unsupported endpoint tokens on their original emission path,
            // including verified CHAR font/encoding bindings and local attributes.
            if (!SelectFenceTemplate(open, close).HasValue) continue;
            var fenced = new XElement(
                "mfenced",
                new XAttribute("open", NormalizeFence(open)),
                new XAttribute("close", NormalizeFence(close)),
                new XAttribute("separators", string.Empty),
                children.Select(child => new XElement(child)));
            row.ReplaceNodes(fenced);
        }
    }

    internal static void MaterializeLooseFenceGroups(XElement root)
    {
        // OMML's reverse projection can express a native group as sibling
        // fences, with a trailing script attached to the closing fence. The
        // semantic reader already recognizes that group. Encode the same group
        // as one native fence OBJECT before the trailing-script writer binds it.
        // Explicitly fixed fence CHARs are excluded by the shared predicates.
        foreach (var container in root.DescendantsAndSelf()
                     .Where(element => element.Name.LocalName is "math" or "mrow" or "mstyle")
                     .Reverse().ToArray())
        {
            var children = container.Elements().ToArray();
            for (var index = 0; index < children.Length; index++)
            {
                if (!TryGetLooseOpeningFenceToken(children[index], out var open))
                    continue;
                var stack = new List<string> { open };
                for (var cursor = index + 1; cursor < children.Length; cursor++)
                {
                    var candidate = children[cursor];
                    var scriptChildren = candidate.Elements().ToArray();
                    var isScript = candidate.Name.LocalName is "msub" or "msup" or "msubsup";
                    var requiredSlots = candidate.Name.LocalName == "msubsup" ? 3 : 2;
                    var closingCandidate = isScript && scriptChildren.Length == requiredSlots
                        ? scriptChildren[0] : candidate;
                    if (TryGetLooseClosingFenceToken(closingCandidate, out var close))
                    {
                        if (!AreMatchingFences(stack[stack.Count - 1], close))
                            break;
                        stack.RemoveAt(stack.Count - 1);
                        if (stack.Count > 0) continue;
                        var inner = new XElement(container.Name.Namespace + "mrow",
                            children.Skip(index + 1).Take(cursor - index - 1)
                                .Select(element => new XElement(element)));
                        MaterializeLooseFenceGroups(inner);
                        var fenced = new XElement(container.Name.Namespace + "mfenced",
                            new XAttribute("open", NormalizeFence(open)),
                            new XAttribute("close", NormalizeFence(close)),
                            new XAttribute("separators", string.Empty), inner);
                        XElement replacement = fenced;
                        if (isScript)
                        {
                            replacement = new XElement(candidate);
                            replacement.Elements().First().ReplaceWith(fenced);
                        }
                        children[index].ReplaceWith(replacement);
                        foreach (var removed in children.Skip(index + 1).Take(cursor - index))
                            removed.Remove();
                        children = container.Elements().ToArray();
                        break;
                    }
                    if (TryGetLooseOpeningFenceToken(candidate, out var nestedOpen))
                        stack.Add(nestedOpen);
                }
            }
        }
    }

    internal static string CanonicalizeElementSequence(
        XElement[] elements,
        string? variant)
    {
        if (elements.Length == 0)
            return string.Empty;
        if (TryCanonicalizeSplitFencedScriptSequence(
                elements,
                variant,
                out var splitFencedScript))
            return splitFencedScript;
        if (TryCanonicalizeLooseBinomialSequence(
                elements,
                variant,
                out var looseBinomialSequence))
            return looseBinomialSequence;
        if (TryCanonicalizeLooseFenceSequence(
                elements,
                variant,
                out var looseFenceSequence))
            return looseFenceSequence;
        return string.Concat(
            elements.Select(
                child => CanonicalizeMathMl(
                    child,
                    variant)));
    }

    private static bool TryCanonicalizeLooseBinomialSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 3) return false;
        var builder = new StringBuilder();
        var replaced = false;
        for (var index = 0; index < elements.Length;)
        {
            if (index + 2 < elements.Length
                && TryGetLooseFenceToken(elements[index], out var open)
                && TryGetLooseFenceToken(elements[index + 2], out var close)
                && NormalizeFence(open) == "("
                && NormalizeFence(close) == ")")
            {
                if (TryCanonicalizeLooseBinomialBody(
                        elements[index + 1],
                        variant,
                        out var binomialBody))
                {
                    builder.Append(binomialBody);
                    index += 3;
                    replaced = true;
                    continue;
                }
            }
            builder.Append(CanonicalizeMathMl(elements[index], variant));
            index++;
        }
        if (!replaced) return false;
        signature = builder.ToString();
        return true;
    }

    private static bool TryCanonicalizeLooseFenceSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 2) return false;

        var builder = new StringBuilder();
        var replaced = false;
        for (var index = 0; index < elements.Length;)
        {
            if (TryGetLooseOpeningFenceToken(
                    elements[index],
                    out var open)
                && TryFindMatchingLooseFence(
                    elements,
                    index,
                    open,
                    out var closeIndex,
                    out var close))
            {
                var innerElements =
                    elements
                        .Skip(index + 1)
                        .Take(closeIndex - index - 1)
                        .ToArray();
                innerElements =
                    StripRedundantNestedLooseFenceLayers(
                        innerElements,
                        open,
                        close);
                var innerSignature =
                    CanonicalizeElementSequence(
                        innerElements,
                        variant);

                builder.Append(
                    CanonicalFenceSignature(
                        open,
                        close,
                        innerSignature));
                index = closeIndex + 1;
                replaced = true;
                continue;
            }

            builder.Append(
                CanonicalizeMathMl(
                    elements[index],
                    variant));
            index++;
        }

        if (!replaced) return false;
        signature = builder.ToString();
        return true;
    }

    private static bool TryFindMatchingLooseFence(
        XElement[] elements,
        int openIndex,
        string open,
        out int closeIndex,
        out string close)
    {
        closeIndex = -1;
        close = string.Empty;
        var normalizedOpen = NormalizeFence(open);
        var expectedClose =
            MatchingClosingFence(
                normalizedOpen);
        if (expectedClose.Length == 0)
            return false;

        var depth = 0;
        for (var index = openIndex + 1;
             index < elements.Length;
             index++)
        {
            if (TryGetLooseOpeningFenceToken(
                    elements[index],
                    out var nestedOpen)
                && string.Equals(
                    NormalizeFence(nestedOpen),
                    normalizedOpen,
                    StringComparison.Ordinal))
            {
                depth++;
                continue;
            }

            if (!TryGetLooseClosingFenceToken(
                    elements[index],
                    out var candidateClose)
                || !string.Equals(
                    NormalizeFence(candidateClose),
                    expectedClose,
                    StringComparison.Ordinal))
                continue;

            if (depth > 0)
            {
                depth--;
                continue;
            }

            closeIndex = index;
            close = candidateClose;
            return true;
        }

        return false;
    }

    private static bool TryGetLooseOpeningFenceToken(
        XElement candidate,
        out string fence)
    {
        if (!TryGetLooseFenceToken(
                candidate,
                out fence))
            return false;

        var normalized =
            NormalizeFence(
                fence);
        if (normalized is "(" or "[" or "{" or "⟨" or "⌈" or "⌊")
            return true;

        if (normalized is "|" or "‖")
            return HasExplicitFenceRole(
                candidate,
                "OPEN");

        return false;
    }

    private static bool TryGetLooseClosingFenceToken(
        XElement candidate,
        out string fence)
    {
        if (!TryGetLooseFenceToken(
                candidate,
                out fence))
            return false;

        var normalized =
            NormalizeFence(
                fence);
        if (normalized is ")" or "]" or "}" or "⟩" or "⌉" or "⌋")
            return true;

        if (normalized is "|" or "‖")
            return HasExplicitFenceRole(
                candidate,
                "CLOSE");

        return false;
    }

    private static bool HasExplicitFenceRole(
        XElement candidate,
        string expectedRole)
    {
        XElement token = candidate;
        if (candidate.Name.LocalName == "mrow")
        {
            var children = candidate.Elements().ToArray();
            if (children.Length != 1
                || children[0].Name.LocalName != "mo")
                return false;
            token = children[0];
        }

        var candidateClass =
            ((string?)candidate.Attribute(
                "data-mjx-texclass")
             ?? string.Empty).Trim();
        var tokenClass =
            ((string?)token.Attribute(
                "data-mjx-texclass")
             ?? string.Empty).Trim();
        if (string.Equals(
                candidateClass,
                expectedRole,
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(
                tokenClass,
                expectedRole,
                StringComparison.OrdinalIgnoreCase))
            return true;

        return string.Equals(
                (string?)token.Attribute(
                    "fence"),
                "true",
                StringComparison.OrdinalIgnoreCase)
            && string.Equals(
                (string?)token.Attribute(
                    "stretchy"),
                "true",
                StringComparison.OrdinalIgnoreCase);
    }

    private static string MatchingClosingFence(
        string open) =>
        NormalizeFence(
            open) switch
        {
            "(" => ")",
            "[" => "]",
            "{" => "}",
            "⟨" => "⟩",
            "⌈" => "⌉",
            "⌊" => "⌋",
            "|" => "|",
            "‖" => "‖",
            _ => string.Empty,
        };

    private static XElement[] StripRedundantNestedLooseFenceLayers(
        XElement[] elements,
        string outerOpen,
        string outerClose)
    {
        var normalizedOuterOpen =
            NormalizeFence(
                outerOpen);
        var normalizedOuterClose =
            NormalizeFence(
                outerClose);
        var current =
            elements;
        while (current.Length >= 2
               && TryGetLooseOpeningFenceToken(
                    current[0],
                    out var nestedOpen)
               && TryFindMatchingLooseFence(
                    current,
                    0,
                    nestedOpen,
                    out var nestedCloseIndex,
                    out var nestedClose)
               && nestedCloseIndex ==
                    current.Length - 1
               && string.Equals(
                    NormalizeFence(
                        nestedOpen),
                    normalizedOuterOpen,
                    StringComparison.Ordinal)
               && string.Equals(
                    NormalizeFence(
                        nestedClose),
                    normalizedOuterClose,
                    StringComparison.Ordinal))
        {
            current =
                current
                    .Skip(1)
                    .Take(
                        current.Length - 2)
                    .ToArray();
        }

        return current;
    }

    internal static string CanonicalFenceSignature(
        string open,
        string close,
        string innerSignature)
    {
        var normalizedOpen =
            NormalizeFence(
                open);
        var normalizedClose =
            NormalizeFence(
                close);

        if (normalizedOpen.Length > 0
            && normalizedClose.Length == 0)
        {
            return "o("
                + normalizedOpen
                + ")"
                + innerSignature;
        }

        if (normalizedOpen.Length == 0
            && normalizedClose.Length > 0)
        {
            return innerSignature
                + "o("
                + normalizedClose
                + ")";
        }

        return "fence("
            + normalizedOpen
            + ","
            + normalizedClose
            + ","
            + innerSignature
            + ")";
    }

    private static bool TryCanonicalizeLooseBinomialBody(
        XElement body,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (body.Name.LocalName == "mfrac")
        {
            var thickness = ((string?)body.Attribute("linethickness") ?? string.Empty).Trim();
            var fractionChildren = body.Elements().ToArray();
            if (!thickness.StartsWith("0", StringComparison.Ordinal)
                || fractionChildren.Length != 2)
                return false;
            signature = "binom("
                + CanonicalizeMathMl(fractionChildren[0], variant)
                + ","
                + CanonicalizeMathMl(fractionChildren[1], variant)
                + ")";
            return true;
        }

        body = UnwrapSingleTransparentContainer(body);
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
        var top = string.Concat(cells[0][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        var bottom = string.Concat(cells[1][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        signature = "binom(" + top + "," + bottom + ")";
        return true;
    }

    private static XElement UnwrapSingleTransparentContainer(XElement element)
    {
        while (element.Name.LocalName is "mrow" or "mstyle" or "mpadded")
        {
            var children = element.Elements().ToArray();
            if (children.Length != 1) break;
            element = children[0];
        }
        return element;
    }

    internal static bool TryCanonicalizeMathJaxBinomialFence(
        string open,
        string close,
        XElement[] children,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (NormalizeFence(open) != "(" || NormalizeFence(close) != ")") return false;
        if (children.Length != 1 || children[0].Name.LocalName != "mfrac") return false;
        var lineThickness = (children[0].Attribute("linethickness")?.Value ?? string.Empty).Trim();
        if (lineThickness.Length == 0 || !lineThickness.StartsWith("0", StringComparison.Ordinal)) return false;
        var fractionChildren = children[0].Elements().ToArray();
        if (fractionChildren.Length != 2) return false;
        var top = CanonicalizeMathMl(fractionChildren[0], variant);
        var bottom = CanonicalizeMathMl(fractionChildren[1], variant);
        signature = "binom(" + top + "," + bottom + ")";
        return true;
    }

    internal static bool TryCanonicalizeMathTypeBinomialFence(
        XElement fenced,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        var open = NormalizeFence((string?)fenced.Attribute("open") ?? "(");
        var close = NormalizeFence((string?)fenced.Attribute("close") ?? ")");
        if (open != "(" || close != ")") return false;
        var table = fenced.Elements().SingleOrDefault();
        // Word's OMML round-trip commonly materializes a native no-bar
        // fraction as mfenced -> mrow -> mfrac(linethickness=0). Treat that
        // as the same binomial semantics before checking MathType's PILE form.
        if (table is not null
            && TryCanonicalizeLooseBinomialBody(table, variant, out signature))
            return true;
        if (table is not null)
            table = UnwrapSingleTransparentContainer(table);
        if (table?.Name.LocalName == "mfrac")
            return TryCanonicalizeMathJaxBinomialFence(
                open,
                close,
                new[] { table },
                variant,
                out signature);
        if (table is null || table.Name.LocalName != "mtable") return false;
        // A genuine MathType binomial reaches us through a PILE record. An
        // explicit 2x1 matrix inside parentheses is visually similar, but is a
        // MATRIX record and must remain a matrix. The old shape-only heuristic
        // collapsed both into binom(...), which made strict VisualTeX→MathType
        // round-trip validation reject explicit column matrices such as the
        // binomial-theorem formula from 文档1.
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
        var top = string.Concat(cells[0][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        var bottom = string.Concat(cells[1][0].Elements()
            .Select(child => CanonicalizeMathMl(child, variant)));
        signature = "binom(" + top + "," + bottom + ")";
        return true;
    }

    private static bool TryCanonicalizeSplitFencedScriptSequence(
        XElement[] elements,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (elements.Length < 3) return false;
        if (!TryGetLooseOpeningFenceToken(
                elements[0],
                out var open))
            return false;

        for (var scriptIndex = 2;
             scriptIndex < elements.Length;
             scriptIndex++)
        {
            var script = elements[scriptIndex];
            var scriptKind = script.Name.LocalName;
            if (scriptKind is not ("msub" or "msup" or "msubsup"))
                continue;

            var scriptChildren =
                script.Elements().ToArray();
            var minimumChildren =
                scriptKind == "msubsup"
                    ? 3
                    : 2;
            if (scriptChildren.Length < minimumChildren)
                continue;
            if (!TryGetLooseClosingFenceToken(
                    scriptChildren[0],
                    out var close))
                continue;
            if (!AreMatchingFences(
                    open,
                    close))
                continue;

            var innerElements =
                elements
                    .Skip(1)
                    .Take(scriptIndex - 1)
                    .ToArray();
            innerElements =
                StripRedundantNestedLooseFenceLayers(
                    innerElements,
                    open,
                    close);
            var innerSignature =
                CanonicalizeElementSequence(
                    innerElements,
                    variant);
            if (innerSignature.Length == 0)
                continue;

            var baseSignature =
                CanonicalFenceSignature(
                    open,
                    close,
                    innerSignature);
            string scriptedSignature;
            if (scriptKind == "msub")
            {
                var subSignature =
                    CanonicalizeMathMl(
                        scriptChildren[1],
                        variant);
                if (subSignature.Length == 0)
                    continue;
                scriptedSignature =
                    "sub("
                    + baseSignature
                    + ","
                    + subSignature
                    + ")";
            }
            else if (scriptKind == "msup")
            {
                var supSignature =
                    CanonicalizeMathMl(
                        scriptChildren[1],
                        variant);
                if (supSignature.Length == 0)
                    continue;
                scriptedSignature =
                    "sup("
                    + baseSignature
                    + ","
                    + supSignature
                    + ")";
            }
            else
            {
                var subSignature =
                    CanonicalizeMathMl(
                        scriptChildren[1],
                        variant);
                var supSignature =
                    CanonicalizeMathMl(
                        scriptChildren[2],
                        variant);
                if (subSignature.Length == 0
                    && supSignature.Length == 0)
                    continue;
                if (subSignature.Length == 0)
                {
                    scriptedSignature =
                        "sup("
                        + baseSignature
                        + ","
                        + supSignature
                        + ")";
                }
                else if (supSignature.Length == 0)
                {
                    scriptedSignature =
                        "sub("
                        + baseSignature
                        + ","
                        + subSignature
                        + ")";
                }
                else
                {
                    scriptedSignature =
                        "subsup("
                        + baseSignature
                        + ","
                        + subSignature
                        + ","
                        + supSignature
                        + ")";
                }
            }

            var tail =
                CanonicalizeElementSequence(
                    elements
                        .Skip(scriptIndex + 1)
                        .ToArray(),
                    variant);
            signature =
                scriptedSignature
                + tail;
            return true;
        }

        return false;
    }

    // Presentation MathML gives explicit fixed-glyph attributes precedence over
    // TeX spacing classes. OPEN/CLOSE on \\lvert, \\rvert and \\langle describe
    // spacing, not an expandable fence template. MTEF CHAR does not retain those
    // MathJax annotations, so canonicalization and emission must agree here.
    private static bool IsExplicitlyFixedFence(XElement token) =>
        IsFalseMathMlBoolean((string?)token.Attribute("fence"))
        || IsFalseMathMlBoolean((string?)token.Attribute("stretchy"));

    // The source's attachment semantics decide decoration versus an upper
    // limit before a big-operator shortcut is selected. Overline and horizontal
    // braces have native decoration templates even with accent=false.
    internal static bool IsOverDecoration(XElement mover, XElement overscript) =>
        ResolveOverAccent(mover, overscript)
        || NormalizeAccentMark(overscript.Value.Trim()) == "¯"
        || IsHorizontalBraceMarker(overscript.Value.Trim(), top: true);

    internal static bool ResolveOverAccent(XElement mover, XElement overscript)
    {
        var own = (string?)mover.Attribute("accent");
        if (IsTrueMathMlBoolean(own)) return true;
        if (IsFalseMathMlBoolean(own)) return false;

        // MathML 3, 3.4.5: an omitted accent inherits the core mo's accent,
        // including its operator-dictionary default. Absence is not false.
        var core = overscript;
        while (core.Name.LocalName is "mrow" or "mstyle" or "mpadded" or "semantics")
        {
            var contents = core.Elements()
                .Where(child => child.Name.LocalName is not ("annotation" or "annotation-xml"))
                .ToArray();
            if (contents.Length != 1) return false;
            core = contents[0];
        }
        if (core.Name.LocalName != "mo") return false;
        var mark = (string?)core.Attribute("accent");
        if (IsTrueMathMlBoolean(mark)) return true;
        if (IsFalseMathMlBoolean(mark)) return false;
        return NormalizeAccentMark(core.Value.Trim()) is "→" or "←" or "↔"
            or "^" or "~" or "˙" or "¨" or "¯"
            or "⌢" or "⏜" or "ˇ" or "˘" or "´" or "`" or "˚"
            or "⃛" or "⃜";
    }

    private static bool TryGetLooseFenceToken(XElement candidate, out string fence)
    {
        fence = string.Empty;
        XElement? token = candidate.Name.LocalName == "mo" ? candidate : null;
        if (token is null && candidate.Name.LocalName == "mrow")
        {
            var nested = candidate.Elements().ToArray();
            if (nested.Length == 1 && nested[0].Name.LocalName == "mo")
                token = nested[0];
        }
        if (token is null || IsExplicitlyFixedFence(token)) return false;
        fence = NormalizeFence(token.Value.Trim());
        return fence is "(" or ")" or "[" or "]" or "{" or "}"
            or "⟨" or "⟩" or "⌈" or "⌉" or "⌊" or "⌋" or "|" or "‖";
    }

    private static bool AreMatchingFences(string open, string close) =>
        (NormalizeFence(open), NormalizeFence(close)) switch
        {
            ("(", ")") => true,
            ("[", "]") => true,
            ("{", "}") => true,
            ("⟨", "⟩") => true,
            ("⌈", "⌉") => true,
            ("⌊", "⌋") => true,
            ("|", "|") => true,
            ("‖", "‖") => true,
            _ => false,
        };

    internal static bool TryCanonicalizeMathJaxFencedSuperscriptRow(
        XElement row,
        string? variant,
        out string signature)
    {
        signature = string.Empty;
        if (row.Name.LocalName != "mrow") return false;
        var elements = row.Elements().ToArray();
        if (elements.Length < 3) return false;
        if (!TryGetMathJaxFenceToken(elements[0], "OPEN", out var open)) return false;

        var script = elements[elements.Length - 1];
        if (script.Name.LocalName != "msup") return false;
        var scriptChildren = script.Elements().ToArray();
        if (scriptChildren.Length < 2) return false;
        if (!TryGetMathJaxFenceToken(scriptChildren[0], "CLOSE", out var close)) return false;
        if (!AreMatchingFences(open, close)) return false;

        var inner = string.Concat(elements
            .Skip(1)
            .Take(elements.Length - 2)
            .Select(child => CanonicalizeMathMl(child, variant)));
        var baseSignature = "o(" + NormalizeFence(open) + ")"
            + inner
            + "o(" + NormalizeFence(close) + ")";
        var exponentSignature = CanonicalizeMathMl(scriptChildren[1], variant);
        signature = exponentSignature.Length == 0
            ? baseSignature
            : "sup(" + baseSignature + "," + exponentSignature + ")";
        return true;
    }

    internal static bool TryGetMathJaxFenceRow(
        XElement row,
        out string open,
        out string close,
        out XElement[] children)
    {
        open = string.Empty;
        close = string.Empty;
        children = Array.Empty<XElement>();
        // Explicit source structure is independent of MTEF template support.
        // Unsupported templates still use EmitFenced\'s existing fallback;
        // semantic comparison must not erase a source fence to excuse that loss.
        if (MathMlExplicitFenceContract.TryGetRow(row, out open, out close, out children))
            return true;
        if (row.Name.LocalName != "mrow") return false;
        var elements = row.Elements().ToArray();
        if (elements.Length < 3) return false;
        if (!TryGetMathJaxFenceToken(elements[0], "OPEN", out open)) return false;
        if (!TryGetMathJaxFenceToken(elements[elements.Length - 1], "CLOSE", out close)) return false;
        // MathJax represents \left\{ ... \right. with an explicit empty
        // closing fence. It is a one-sided delimiter template, not an empty
        // operator after an ordinary brace glyph. Use the same recognition for
        // MTEF geometry and semantic comparison.
        var oneSided = (open.Length == 0) != (close.Length == 0);
        if (!AreMatchingFences(open, close)
            && !(oneSided && SelectFenceTemplate(open, close).HasValue)) return false;
        children = elements.Skip(1).Take(elements.Length - 2).ToArray();
        return children.Length > 0;
    }

    private static bool TryGetMathJaxFenceToken(
        XElement candidate,
        string expectedClass,
        out string fence)
    {
        fence = string.Empty;
        XElement? token = null;
        if (candidate.Name.LocalName == "mo")
            token = candidate;
        else if (candidate.Name.LocalName == "mrow")
        {
            var nested = candidate.Elements().ToArray();
            if (nested.Length == 1 && nested[0].Name.LocalName == "mo")
                token = nested[0];
        }
        if (token is null) return false;

        var candidateClass = ((string?)candidate.Attribute("data-mjx-texclass") ?? string.Empty).Trim();
        var tokenClass = ((string?)token.Attribute("data-mjx-texclass") ?? string.Empty).Trim();
        var markerMatches = string.Equals(candidateClass, expectedClass, StringComparison.OrdinalIgnoreCase)
            || string.Equals(tokenClass, expectedClass, StringComparison.OrdinalIgnoreCase);
        fence = NormalizeFence(token.Value.Trim());
        // This rule must run before OPEN/CLOSE spacing-class recognition and
        // must be identical to the loose-sequence canonicalizer's rule.
        if (IsExplicitlyFixedFence(token)) return false;
        if (markerMatches) return true;
        if (fence.Length == 0
            && string.Equals((string?)token.Attribute("fence"), "true", StringComparison.OrdinalIgnoreCase))
            return true;

        // Word's OMML->MathML transform emits structurally complete delimiter
        // rows without MathJax's data-mjx-texclass markers. MathType 7 otherwise
        // imports only the first opening delimiter into a big-operator main slot
        // and leaves the remaining operand as root siblings. Accept an unmarked
        // token only when its glyph unambiguously matches the requested side;
        // the caller also verifies that the two delimiters form a valid pair.
        return string.Equals(expectedClass, "OPEN", StringComparison.OrdinalIgnoreCase)
            ? fence is "(" or "[" or "{" or "⟨" or "⌈" or "⌊" or "|" or "‖"
            : string.Equals(expectedClass, "CLOSE", StringComparison.OrdinalIgnoreCase)
                && fence is ")" or "]" or "}" or "⟩" or "⌉" or "⌋" or "|" or "‖";
    }
}

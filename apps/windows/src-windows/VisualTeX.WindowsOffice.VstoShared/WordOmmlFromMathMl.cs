using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using static VisualTeX.WordVsto.WordOmmlConverter;
using static VisualTeX.WordVsto.WordOmmlFingerprints;

namespace VisualTeX.WordVsto;

/// <summary>
/// MathML made ready for Word's MathML-to-OMML transform (fences, spaces, token styles, functions, accents,
/// scripts, n-ary operands, tables), and the OMML it returns normalized and validated.
/// </summary>
internal static class WordOmmlFromMathMl
{
    private const string NaryCharacters =
        "∑∏∐∫∬∭∮∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜⋀⋁⋂⋃";
    private const string ExtendedIntegralCharacters =
        "∯∰∱∲∳⨋⨌⨍⨎⨏⨐⨑⨒⨓⨔⨕⨖⨗⨘⨙⨚⨛⨜";

    internal static void ValidateMathMlForOmml(string mathMl)
    {
        // Presentation MathML is already parsed mathematical content. Literal
        // mtext (including paths, command documentation and colored text) must
        // not be interpreted as LaTeX a second time. The TeX producer reports
        // unknown commands at parse time; this boundary rejects real error nodes.
        // Use expanded XML names, so namespace aliases cannot bypass validation.
        var document = FormulaXml.Parse(mathMl, "MathML");
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        if (document.Root?.Name != presentationMath + "math")
            throw new InvalidDataException("The OMML source must have a Presentation MathML math root.");
        if (document.Descendants(presentationMath + "merror").Any())
            throw new InvalidDataException(
                "MathML contains an error node and cannot be inserted as OMML.");
    }

    internal static string NormalizeMathTypeBinomialPiles(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("data-mtef-pile", StringComparison.OrdinalIgnoreCase) < 0)
            return mathMl;

        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        foreach (var table in document.Descendants(presentationMath + "mtable").ToList())
        {
            if (!string.Equals(
                    table.Attribute("data-mtef-pile")?.Value,
                    "true",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            var rows = table.Elements(presentationMath + "mtr").ToArray();
            if (rows.Length != 2) continue;
            var cells = rows
                .Select(row => row.Elements(presentationMath + "mtd").ToArray())
                .ToArray();
            if (cells.Any(row => row.Length != 1)) continue;

            // MTEF uses a two-row, one-column PILE inside parentheses for its
            // native binomial template. Leaving that PILE as <mtable> makes the
            // Office XSLT materialize a 2x1 matrix, which is visually similar but
            // semantically different from Word's native no-bar fraction. The
            // private decoder marker distinguishes this from a genuine MATRIX,
            // so explicit column matrices must not enter this path.
            XElement CopyCellAsRow(XElement cell) =>
                new(
                    presentationMath + "mrow",
                    new XElement(cell).Nodes());

            table.ReplaceWith(
                new XElement(
                    presentationMath + "mfrac",
                    new XAttribute("linethickness", "0"),
                    CopyCellAsRow(cells[0][0]),
                    CopyCellAsRow(cells[1][0])));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    /// <summary>
    /// MathJax renders \binom{n}{k} as an ORD row of
    /// [OPEN row(mo)] [mfrac linethickness=0] [CLOSE row(mo)], with fixed-size
    /// delimiters. The extra wrapper rows hide the delimiters from the Office
    /// XSLT, which then writes them as two ordinary characters that do not grow
    /// with the fraction (small parentheses in Word). Native Word writes a
    /// binomial as m:d around a bar-less m:f, so project exactly this shape to
    /// an mfenced around the fraction. Hand-sized delimiters (\bigl, \Bigr) and
    /// any other body are left untouched.
    /// </summary>
    internal static string NormalizeMathJaxBinomialDelimiters(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("linethickness", StringComparison.Ordinal) < 0)
            return mathMl;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        bool TryGetWrappedFence(XElement element, string role, out string fence)
        {
            fence = string.Empty;
            if (element.Name != presentationMath + "mrow"
                || !MathMlExplicitFenceContract.HasTexClass(element, role))
                return false;
            var inner = element.Elements().ToArray();
            if (inner.Length != 1
                || inner[0].Name != presentationMath + "mo"
                || inner[0].HasElements
                || string.IsNullOrWhiteSpace(inner[0].Value))
                return false;
            fence = inner[0].Value.Trim();
            return true;
        }

        foreach (var row in document.Descendants(presentationMath + "mrow").Reverse().ToList())
        {
            var children = row.Elements().ToArray();
            if (children.Length != 3) continue;
            if (!TryGetWrappedFence(children[0], "OPEN", out var open)
                || !TryGetWrappedFence(children[2], "CLOSE", out var close))
                continue;
            var fraction = children[1];
            if (fraction.Name != presentationMath + "mfrac") continue;
            var thickness = (fraction.Attribute("linethickness")?.Value ?? string.Empty).Trim();
            if (!thickness.StartsWith("0", StringComparison.Ordinal)
                || fraction.Elements().Count() != 2)
                continue;
            row.ReplaceNodes(new XElement(
                presentationMath + "mfenced",
                new XAttribute("open", open),
                new XAttribute("close", close),
                new XAttribute("separators", string.Empty),
                new XElement(fraction)));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static readonly (string Open, string Close)[] AdaptiveBracketPairs =
    {
        ("(", ")"), ("[", "]"), ("{", "}"), ("⟨", "⟩"),
        ("⌊", "⌋"), ("⌈", "⌉"),
    };

    /// <summary>
    /// Word's own equation editor builds matching brackets into m:d, which
    /// grows with its content. A bracket that arrives as a bare character
    /// (typed without \left/\right, or hand-sized with \big) is just a glyph
    /// and stays small beside tall content, and a closing bracket that carries
    /// a superscript sits alone as the base of the script. Pair brackets that
    /// match within one sequence of siblings and write them as one fenced
    /// group. Unmatched or mismatched brackets (intervals such as [0,1)),
    /// vertical bars and relation-class brackets are left alone.
    /// </summary>
    internal static string NormalizePlainBracketPairs(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        // MathML read back from a MathType equation carries data-mtef-*
        // attributes. Its literal brackets are the author's choice and a
        // conversion keeps them as they are.
        if (mathMl.IndexOf("data-mtef-", StringComparison.Ordinal) >= 0) return mathMl;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace mml = "http://www.w3.org/1998/Math/MathML";
        var containerNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "math", "mrow", "mstyle", "mtd", "msqrt", "mpadded", "mphantom",
        };
        var scriptNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "msup", "msub", "msubsup",
        };

        string? Token(XElement element, string role)
        {
            var token = element;
            if (element.Name == mml + "mrow")
            {
                var inner = element.Elements().ToArray();
                if (inner.Length != 1) return null;
                var wrapperClass = element.Attribute("data-mjx-texclass")?.Value;
                if (!string.IsNullOrEmpty(wrapperClass)
                    && !string.Equals(wrapperClass, role, StringComparison.OrdinalIgnoreCase))
                    return null;
                token = inner[0];
            }
            if (token.Name != mml + "mo" || token.HasElements) return null;
            // MathJax marks every unprefixed bracket (and \{, \lfloor, ...)
            // fence="false" stretchy="false", so those flags are the ordinary
            // case here. A bracket that Word wrote as a literal character is
            // recognizable by mathvariant and keeps its fixed size; MathType
            // sources are skipped as a whole above.
            if (token.Attribute("mathvariant") is not null) return null;
            var tokenClass = token.Attribute("data-mjx-texclass")?.Value;
            if (!string.IsNullOrEmpty(tokenClass)
                && !string.Equals(tokenClass, role, StringComparison.OrdinalIgnoreCase))
                return null;
            return token.Value.Trim();
        }

        string? ScriptBaseClose(XElement element)
        {
            if (!scriptNames.Contains(element.Name.LocalName)) return null;
            var first = element.Elements().FirstOrDefault();
            return first is null ? null : Token(first, "CLOSE");
        }

        bool Matches(string open, string close) =>
            AdaptiveBracketPairs.Any(pair => pair.Open == open && pair.Close == close);
        bool IsOpen(string? value) =>
            value is not null && AdaptiveBracketPairs.Any(pair => pair.Open == value);

        List<XElement> Rebuild(List<XElement> items)
        {
            var pairs = new Dictionary<int, (int Close, string Open, string CloseText, bool InScript)>();
            var stack = new List<(int Index, string Open)>();
            for (var index = 0; index < items.Count; index++)
            {
                var open = Token(items[index], "OPEN");
                if (IsOpen(open))
                {
                    stack.Add((index, open!));
                    continue;
                }
                var scriptClose = ScriptBaseClose(items[index]);
                var close = scriptClose ?? Token(items[index], "CLOSE");
                if (close is null || stack.Count == 0) continue;
                var top = stack[stack.Count - 1];
                if (!Matches(top.Open, close)) continue;
                stack.RemoveAt(stack.Count - 1);
                var inScript = scriptClose is not null;
                // An empty group has no operand to build; leave it as text.
                if (!inScript && index == top.Index + 1) continue;
                if (inScript && index == top.Index + 1) continue;
                pairs[top.Index] = (index, top.Open, close, inScript);
            }
            if (pairs.Count == 0) return items;

            List<XElement> Build(int start, int end)
            {
                var result = new List<XElement>();
                for (var index = start; index < end; index++)
                {
                    if (!pairs.TryGetValue(index, out var pair))
                    {
                        result.Add(items[index]);
                        continue;
                    }
                    var body = Build(index + 1, pair.Close);
                    var operand = body.Count == 1 && body[0].Name == mml + "mtable"
                        ? new XElement(body[0])
                        : new XElement(mml + "mrow", body.Select(node => new XElement(node)));
                    var fenced = new XElement(
                        mml + "mfenced",
                        new XAttribute("open", pair.Open),
                        new XAttribute("close", pair.CloseText),
                        new XAttribute("separators", string.Empty),
                        operand);
                    if (pair.InScript)
                    {
                        var script = new XElement(items[pair.Close]);
                        script.Elements().First().ReplaceWith(fenced);
                        result.Add(script);
                    }
                    else
                    {
                        result.Add(fenced);
                    }
                    index = pair.Close;
                }
                return result;
            }
            return Build(0, items.Count);
        }

        var changed = false;
        foreach (var container in document.Descendants()
                     .Where(element => containerNames.Contains(element.Name.LocalName)
                         && element.Name.Namespace == mml)
                     .Reverse()
                     .ToList())
        {
            var children = container.Elements().ToList();
            if (children.Count < 2) continue;
            // Explicit \left...\right rows are already one native group; only
            // their body can hold further bare brackets.
            var explicitRow = MathMlExplicitFenceContract.TryGetRow(container, out _, out _, out _)
                || (MathMlExplicitFenceContract.HasTexClass(container, "INNER")
                    && MathMlExplicitFenceContract.HasTexClass(children[0], "OPEN")
                    && MathMlExplicitFenceContract.HasTexClass(children[children.Count - 1], "CLOSE"));
            var head = new List<XElement>();
            var tail = new List<XElement>();
            var middle = children;
            if (explicitRow)
            {
                head.Add(children[0]);
                tail.Add(children[children.Count - 1]);
                middle = children.Skip(1).Take(children.Count - 2).ToList();
            }
            var rebuilt = Rebuild(middle);
            if (ReferenceEquals(rebuilt, middle)) continue;
            container.ReplaceNodes(head.Concat(rebuilt).Concat(tail));
            changed = true;
        }
        return changed ? document.ToString(SaveOptions.DisableFormatting) : mathMl;
    }

    internal static string NormalizeFencedMathMlTables(string mathMl)
    {
        // Both legacy fenced tables and explicit MathJax INNER/OPEN/CLOSE rows
        // need native delimiter structure. Expanded XML names remain authoritative.
        if (string.IsNullOrWhiteSpace(mathMl)
            || mathMl.IndexOf("mrow", StringComparison.Ordinal) < 0)
            return mathMl;
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        // Only the legacy table fallback uses character recognition. Ordinary
        // content requires the producer's explicit group and endpoint roles.
        var fallbackFenceCharacters = new HashSet<string>(StringComparer.Ordinal)
        {
            "(", ")", "[", "]", "{", "}", "|", "‖", "⌈", "⌉", "⌊", "⌋",
            "⟨", "⟩", "/", "\\", "↑", "↓", "↕", "⇑", "⇓", "⇕",
        };

        bool IsLegacyTableFence(XElement element, string expectedTexClass)
        {
            if (element.Name != presentationMath + "mo"
                || !MathMlExplicitFenceContract.IsProjectableEndpoint(element)) return false;
            if (MathMlExplicitFenceContract.HasTexClass(element, expectedTexClass)) return true;
            if (string.Equals(element.Attribute("fence")?.Value, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(element.Attribute("stretchy")?.Value, "true", StringComparison.OrdinalIgnoreCase))
                return true;
            return fallbackFenceCharacters.Contains(element.Value);
        }

        // Normalize nested groups before their parents. Retaining each source
        // row also preserves inherited attributes and applied-function boundaries.
        foreach (var row in document.Descendants(presentationMath + "mrow").Reverse().ToList())
        {
            var children = row.Elements().ToArray();
            if (children.Length < 2) continue;
            var first = children[0];
            var last = children[children.Length - 1];
            var explicitFence = MathMlExplicitFenceContract.TryGetRow(row, out _, out _, out _);
            var tableBody = children.Length == 3
                && children[1].Name == presentationMath + "mtable";
            var legacyTableFence = tableBody
                && IsLegacyTableFence(first, "OPEN") && IsLegacyTableFence(last, "CLOSE");
            if (!explicitFence && !legacyTableFence) continue;
            var open = first.Value;
            var close = last.Value;
            if (open.Length == 0 && close.Length == 0) continue;

            // mfenced treats each child as a separate operand. Keep the entire
            // original body in one mrow so no implicit separators/operand splits
            // are introduced. A sole mtable remains direct for matrix validation
            // and existing Office table handling.
            var bodyNodes = row.Nodes().Where(node => node != first && node != last).ToArray();
            var operand = tableBody
                ? new XElement(children[1])
                : new XElement(presentationMath + "mrow", bodyNodes);
            row.ReplaceNodes(new XElement(
                presentationMath + "mfenced",
                new XAttribute("open", open),
                new XAttribute("close", close),
                new XAttribute("separators", string.Empty),
                operand));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static bool IsMatrixLikeMathMlTable(XElement table, XNamespace presentationMath)
    {
        var parent = table.Parent;
        if (parent?.Name == presentationMath + "mfenced")
        {
            var fencedRows = table.Elements(presentationMath + "mtr").ToArray();
            if (fencedRows.Length == 0) return false;
            var fencedColumns = fencedRows[0].Elements(presentationMath + "mtd").Count();
            // Office may represent a one-column fenced stack as an equation
            // array rather than m:m. The surrounding m:d is still correct and
            // stretchable, so matrix-dimension validation is only meaningful
            // for real multi-column tables.
            return fencedColumns > 1
                && fencedRows.All(row =>
                    row.Elements(presentationMath + "mtd").Count() == fencedColumns);
        }

        // MathJax also uses mtable for aligned equations, cases and substack.
        // Those structures legitimately become m:eqArr or limit constructs in
        // OMML, so only validate table shapes that carry matrix-like alignment.
        if (string.Equals(
                table.Attribute("displaystyle")?.Value,
                "false",
                StringComparison.OrdinalIgnoreCase)
            || table.Attribute("scriptlevel") is not null)
            return false;
        var alignments = (table.Attribute("columnalign")?.Value ?? string.Empty)
            .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (alignments.Any(value =>
                string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)))
            return false;
        var rows = table.Elements(presentationMath + "mtr").ToArray();
        if (rows.Length == 0) return false;
        var columns = rows[0].Elements(presentationMath + "mtd").Count();
        // A one-column mtable is the normal MathJax representation of
        // \substack and must become an OMML limit stack rather than m:m.
        return columns > 1
            && rows.All(row =>
                row.Elements(presentationMath + "mtd").Count() == columns);
    }

    private const string OmmlSpaceTextAttribute = "data-visualtex-omml-space";

    internal static (string MathMl, IReadOnlyList<MathSpaceOmmlCapability> Capabilities)
        ProjectMathMlSpacesForOmml(string mathMl)
    {
        if (mathMl.IndexOf("mspace", StringComparison.Ordinal) < 0)
            return (mathMl, Array.Empty<MathSpaceOmmlCapability>());
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var capabilities = new List<MathSpaceOmmlCapability>();
        foreach (var space in document.Descendants(presentationMath + "mspace"))
        {
            // Never trust a caller-supplied projection attribute. Only the
            // shared exact-width contract may select native space characters.
            space.Attribute(OmmlSpaceTextAttribute)?.Remove();
            var capability = MathSpaceContract.DescribeOmmlSpace(
                (string?)space.Attribute("width"), (string?)space.Attribute("height"),
                (string?)space.Attribute("depth"), (string?)space.Attribute("linebreak"));
            capabilities.Add(capability);
            if (capability.CanProject)
                space.SetAttributeValue(OmmlSpaceTextAttribute, capability.Text);
            // Unsupported dimensions retain stock behavior. The returned audit
            // limits expose that gap; do not silently round, reject the whole
            // formula, or pretend it was an exact spacing conversion.
        }
        return (document.ToString(SaveOptions.DisableFormatting), capabilities);
    }

    internal static void RestoreOmmlExplicitSpaceTokens(XElement mathMlRoot)
    {
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        foreach (var parent in mathMlRoot.DescendantsAndSelf().ToArray())
        {
            var children = parent.Elements().ToArray();
            for (var index = 0; index < children.Length; index++)
            {
                var first = children[index];
                if (first.Name != presentationMath + "mo" || first.HasElements
                    || MathSpaceContract.FindOmmlSpaceText(first.Value) is null)
                    continue;

                var text = first.Value;
                var identity = MathSpaceContract.FindOmmlSpaceText(text)!;
                var last = index;
                // Word may coalesce math runs, and stock reverse XSL emits one
                // mo per space character. Decode only adjacent pure space
                // tokens with identical attributes, never prose or mixed text.
                for (var next = index + 1; next < children.Length; next++)
                {
                    var candidate = children[next];
                    if (candidate.Name != presentationMath + "mo" || candidate.HasElements
                        || !first.Attributes().Select(a => (a.Name, a.Value))
                            .SequenceEqual(candidate.Attributes().Select(a => (a.Name, a.Value)))
                        || MathSpaceContract.FindOmmlSpaceText(candidate.Value) is null
                        || text.Length + candidate.Value.Length > MathSpaceContract.MaximumOmmlSpaceTextLength)
                        break;
                    var combined = text + candidate.Value;
                    var combinedIdentity = MathSpaceContract.FindOmmlSpaceText(combined);
                    if (combinedIdentity is null) break;
                    text = combined;
                    identity = combinedIdentity;
                    last = next;
                }
                first.ReplaceWith(new XElement(presentationMath + "mspace",
                    first.Attributes().Where(attribute => attribute.Name != "width"),
                    new XAttribute("width", identity.MathMlWidth)));
                for (var remove = index + 1; remove <= last; remove++)
                    children[remove].Remove();
                index = last;
            }
        }
    }

    private const string OmmlRunRoleAttribute = "data-visualtex-omml-run-role";

    internal static string NormalizeMathMlTokenStyleAuthority(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        foreach (var token in document.Descendants().Where(element =>
                     element.Name.Namespace == presentationMath
                     && element.Name.LocalName is "mi" or "mo" or "mn" or "mtext" or "ms"))
        {
            // Resolve once on the source tree. Both Office run grouping and run
            // property emission consume this same effective style, before token
            // text is flattened. In particular, an ordinary mo is normal while
            // a single-character mi is italic.
            var variant = MathMlTokenStyleContract.ResolveEffectiveVariant(token);
            if (variant is null) continue;
            token.SetAttributeValue("mathvariant", variant);

            // Equal typography does not make prose, an identifier, a number,
            // or a named operator interchangeable. Preserve these boundaries
            // while allowing adjacent tokens of the same role and style to
            // share a run. This private projection attribute is consumed by the
            // app-owned stylesheet and is never copied into the OMML result.
            var role = token.Name.LocalName == "mo" && ContainsUnicodeLetter(token.Value)
                ? "named-operator"
                : token.Name.LocalName;
            token.SetAttributeValue(OmmlRunRoleAttribute, role);
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? mathMl;
    }

    internal static string NormalizeExplicitUprightRuns(string omml, string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml) || string.IsNullOrWhiteSpace(mathMl))
            return omml;
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        foreach (var run in document.Descendants(math + "r"))
        {
            var properties = run.Element(math + "rPr");
            // The source token projection now provides the style to Office.
            // Only an actual plain mathematical run permits upright cleanup.
            // Omitted style is italic, not evidence of an upright named word;
            // m:nor alone belongs to prose and must remain Normal Text.
            if ((string?)properties?.Element(math + "sty")?.Attribute(math + "val") != "p")
                continue;
            properties!.Element(math + "nor")?.Remove();
            var text = string.Concat(run.Elements(math + "t").Select(element => element.Value));
            if (!ContainsUnicodeLetter(text)) continue;
            var wordProperties = run.Element(word + "rPr");
            if (wordProperties is null)
            {
                wordProperties = new XElement(word + "rPr");
                properties.AddAfterSelf(wordProperties);
            }
            if (wordProperties.Element(word + "noProof") is null)
                wordProperties.Add(new XElement(word + "noProof"));
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string NormalizeAppliedFunctionStructures(string omml, string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml)
            || string.IsNullOrWhiteSpace(mathMl)
            || (mathMl.IndexOf('\u2061') < 0
                && mathMl.IndexOf("2061", StringComparison.OrdinalIgnoreCase) < 0))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace math = MathNamespace;
        var source = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);

        var applications = new List<(string FunctionToken, string? Open, string? Close, int ArgumentTextElements)>();
        foreach (var marker in source
                     .Descendants(presentationMath + "mo")
                     .Where(element => string.Equals(element.Value, "\u2061", StringComparison.Ordinal)))
        {
            var parent = marker.Parent;
            if (parent is null) continue;
            var siblings = parent.Elements().ToArray();
            var markerIndex = Array.IndexOf(siblings, marker);
            if (markerIndex <= 0 || markerIndex + 1 >= siblings.Length) continue;

            var functionSource = siblings[markerIndex - 1];
            var functionToken = functionSource
                .DescendantsAndSelf()
                .Where(element =>
                    element.Name == presentationMath + "mi"
                    || element.Name == presentationMath + "mo")
                .Select(element => Regex.Match(element.Value, @"\p{L}+"))
                .FirstOrDefault(match => match.Success)?
                .Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(functionToken)) continue;

            var argumentSource = siblings[markerIndex + 1];
            string? open = null;
            string? close = null;
            if (argumentSource.Name == presentationMath + "mo")
            {
                (open, close) = argumentSource.Value switch
                {
                    "(" => ("(", ")"),
                    "[" => ("[", "]"),
                    "{" => ("{", "}"),
                    _ => (null, null),
                };
            }

            var simpleArgument = argumentSource.Name == presentationMath + "mi"
                || argumentSource.Name == presentationMath + "mn"
                || argumentSource.Name == presentationMath + "mo"
                || argumentSource.Name == presentationMath + "mtext";
            var argumentTextElements = simpleArgument && open is null
                ? StringInfo.ParseCombiningCharacters(argumentSource.Value).Length
                : 0;
            applications.Add((functionToken, open, close, argumentTextElements));
        }
        if (applications.Count == 0) return omml;

        var target = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);

        static string RunText(XElement run, XNamespace mathNamespace) =>
            string.Concat(run.Elements(mathNamespace + "t").Select(text => text.Value));

        static XElement CloneRunWithText(
            XElement sourceRun,
            XNamespace mathNamespace,
            string text)
        {
            var clone = new XElement(sourceRun);
            clone.Elements(mathNamespace + "t").Remove();
            var textElement = new XElement(mathNamespace + "t", text);
            if (text.Length > 0
                && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                textElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            clone.Add(textElement);
            return clone;
        }

        static int TextElementEnd(string value, int count)
        {
            if (string.IsNullOrEmpty(value) || count <= 0) return 0;
            var starts = StringInfo.ParseCombiningCharacters(value);
            if (starts.Length <= count) return value.Length;
            return starts[count];
        }

        static int FindMatchingDelimiterEnd(
            string value,
            string open,
            string close,
            ref int depth,
            ref bool started)
        {
            for (var index = 0; index < value.Length; index++)
            {
                var character = value[index].ToString();
                if (string.Equals(character, open, StringComparison.Ordinal))
                {
                    depth++;
                    started = true;
                }
                else if (started && string.Equals(character, close, StringComparison.Ordinal))
                {
                    depth--;
                    if (depth == 0) return index + 1;
                }
            }
            return -1;
        }

        foreach (var application in applications)
        {
            var markerText = target
                .Descendants(math + "t")
                .FirstOrDefault(text => text.Value.IndexOf('\u2061') >= 0);
            var markerRun = markerText?.Parent;
            if (markerRun?.Name != math + "r") continue;
            var parent = markerRun.Parent;
            if (parent is null) continue;
            var functionElement = markerRun.ElementsBeforeSelf().LastOrDefault();
            if (functionElement is null) continue;
            var functionVisibleText = string.Concat(
                functionElement.Descendants(math + "t").Select(text => text.Value));
            if (functionVisibleText.IndexOf(
                    application.FunctionToken,
                    StringComparison.OrdinalIgnoreCase) < 0)
                continue;

            var markerRunText = RunText(markerRun, math);
            var markerIndex = markerRunText.IndexOf('\u2061');
            if (markerIndex < 0) continue;
            // NormalizeExplicitUprightRuns deliberately split the function name
            // into its own m:r, so ApplyFunction must begin the following run.
            // Anything else is an unfamiliar Office-XSL shape and is left intact.
            if (markerIndex != 0) continue;
            var afterMarker = markerRunText.Substring(1);
            var argumentElements = new List<XElement>();
            var consumed = new List<XElement> { markerRun };
            XElement? tail = null;
            var completed = false;

            if (application.Open is not null && application.Close is not null)
            {
                var depth = 0;
                var started = false;
                var end = FindMatchingDelimiterEnd(
                    afterMarker,
                    application.Open,
                    application.Close,
                    ref depth,
                    ref started);
                if (end >= 0)
                {
                    argumentElements.Add(CloneRunWithText(
                        markerRun,
                        math,
                        afterMarker.Substring(0, end)));
                    if (end < afterMarker.Length)
                        tail = CloneRunWithText(markerRun, math, afterMarker.Substring(end));
                    completed = true;
                }
                else if (started)
                {
                    if (afterMarker.Length > 0)
                        argumentElements.Add(CloneRunWithText(markerRun, math, afterMarker));
                    foreach (var sibling in markerRun.ElementsAfterSelf().ToArray())
                    {
                        if (sibling.Name == math + "r")
                        {
                            var siblingText = RunText(sibling, math);
                            end = FindMatchingDelimiterEnd(
                                siblingText,
                                application.Open,
                                application.Close,
                                ref depth,
                                ref started);
                            consumed.Add(sibling);
                            if (end >= 0)
                            {
                                if (end > 0)
                                    argumentElements.Add(CloneRunWithText(
                                        sibling,
                                        math,
                                        siblingText.Substring(0, end)));
                                if (end < siblingText.Length)
                                    tail = CloneRunWithText(sibling, math, siblingText.Substring(end));
                                completed = true;
                                break;
                            }
                            argumentElements.Add(new XElement(sibling));
                            continue;
                        }

                        var visible = string.Concat(
                            sibling.Descendants(math + "t").Select(text => text.Value));
                        _ = FindMatchingDelimiterEnd(
                            visible,
                            application.Open,
                            application.Close,
                            ref depth,
                            ref started);
                        consumed.Add(sibling);
                        argumentElements.Add(new XElement(sibling));
                        if (started && depth == 0)
                        {
                            completed = true;
                            break;
                        }
                    }
                }
            }
            else
            {
                var sourceTextElements = Math.Max(0, application.ArgumentTextElements);
                if (afterMarker.Length > 0)
                {
                    var take = sourceTextElements > 0
                        ? TextElementEnd(afterMarker, sourceTextElements)
                        : TextElementEnd(afterMarker, 1);
                    if (take > 0)
                    {
                        argumentElements.Add(CloneRunWithText(
                            markerRun,
                            math,
                            afterMarker.Substring(0, take)));
                        if (take < afterMarker.Length)
                            tail = CloneRunWithText(markerRun, math, afterMarker.Substring(take));
                        completed = true;
                    }
                }
                else
                {
                    var sibling = markerRun.ElementsAfterSelf().FirstOrDefault();
                    if (sibling is not null)
                    {
                        consumed.Add(sibling);
                        if (sibling.Name == math + "r" && sourceTextElements > 0)
                        {
                            var siblingText = RunText(sibling, math);
                            var take = TextElementEnd(siblingText, sourceTextElements);
                            if (take > 0)
                            {
                                argumentElements.Add(CloneRunWithText(
                                    sibling,
                                    math,
                                    siblingText.Substring(0, take)));
                                if (take < siblingText.Length)
                                    tail = CloneRunWithText(sibling, math, siblingText.Substring(take));
                                completed = true;
                            }
                        }
                        else
                        {
                            argumentElements.Add(new XElement(sibling));
                            completed = true;
                        }
                    }
                }
            }

            if (!completed || argumentElements.Count == 0) continue;

            var function = new XElement(
                math + "func",
                new XElement(math + "fName", new XElement(functionElement)),
                new XElement(math + "e", argumentElements));
            functionElement.ReplaceWith(function);
            foreach (var element in consumed)
                element.Remove();
            if (tail is not null) function.AddAfterSelf(tail);
        }

        return target.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string NormalizeMathMlAccents(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        var accentCharacters = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["^"] = "\u0302",
            ["~"] = "\u0303",
            ["→"] = "\u20D7",
            ["←"] = "\u20D6",
            ["↔"] = "\u20E1",
            ["¯"] = "\u0305",
            ["‾"] = "\u0305",
            ["―"] = "\u0305",
            ["ˉ"] = "\u0305",
            ["˙"] = "\u0307",
            ["¨"] = "\u0308",
            ["ˇ"] = "\u030C",
            ["˘"] = "\u0306",
            ["´"] = "\u0301",
            ["`"] = "\u0300",
            ["˚"] = "\u030A",
        };

        foreach (var mover in document.Descendants(presentationMath + "mover").ToList())
        {
            var children = mover.Elements().ToArray();
            if (children.Length != 2) continue;
            var mark = children[1];
            if (mark.Name != presentationMath + "mo") continue;
            if (string.Equals(
                    mark.Attribute("accent")?.Value,
                    "false",
                    StringComparison.OrdinalIgnoreCase)
                || string.Equals(
                    mover.Attribute("accent")?.Value,
                    "false",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            var sourceCharacter = mark.Value;
            if (!accentCharacters.TryGetValue(sourceCharacter, out var combiningCharacter))
                continue;

            // Office's MML2OMML.XSL only creates a native m:acc node when the
            // MathML mover is explicitly marked as an accent and the mark is a
            // combining accent character. MathJax emits spacing characters
            // such as ^, ˙ and → without accent=true, which Office otherwise
            // converts into m:limUpp or replacement glyphs/placeholder boxes.
            mover.SetAttributeValue("accent", "true");
            mark.SetAttributeValue("accent", "true");
            mark.SetAttributeValue("stretchy", null);
            mark.SetAttributeValue("data-mjx-pseudoscript", null);
            mark.Value = combiningCharacter;
        }

        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static string NormalizeNestedEmptyBaseScripts(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace mathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        var simpleScriptNames = new HashSet<XName>
        {
            mathMlNamespace + "msub",
            mathMlNamespace + "msup",
        };
        var allScriptNames = new HashSet<XName>(simpleScriptNames)
        {
            mathMlNamespace + "msubsup",
        };
        var transparentWrappers = new HashSet<XName>
        {
            mathMlNamespace + "mrow",
            mathMlNamespace + "mstyle",
            mathMlNamespace + "mpadded",
            mathMlNamespace + "mphantom",
            mathMlNamespace + "semantics",
        };

        bool IsEmptyMathNode(XElement element)
        {
            if (element.Name == mathMlNamespace + "mspace") return false;
            if (element.Nodes().OfType<XText>().Any(node => !string.IsNullOrWhiteSpace(node.Value)))
                return false;
            var children = element.Elements().ToArray();
            return children.Length == 0 || children.All(IsEmptyMathNode);
        }

        bool IsOnlyContentOfOuterScriptArgument(XElement candidate)
        {
            XElement current = candidate;
            while (current.Parent is XElement parent)
            {
                if (transparentWrappers.Contains(parent.Name))
                {
                    if (parent.Elements().Any(sibling =>
                            sibling != current && !IsEmptyMathNode(sibling)))
                        return false;
                    if (parent.Nodes().OfType<XText>().Any(node =>
                            !string.IsNullOrWhiteSpace(node.Value)))
                        return false;
                    current = parent;
                    continue;
                }
                if (!allScriptNames.Contains(parent.Name)) return false;
                var children = parent.Elements().ToList();
                var position = children.IndexOf(current);
                return position >= 1;
            }
            return false;
        }

        foreach (var script in document
                     .Descendants()
                     .Where(element => simpleScriptNames.Contains(element.Name))
                     .Reverse()
                     .ToList())
        {
            var children = script.Elements().ToArray();
            if (children.Length < 2
                || !IsEmptyMathNode(children[0])
                || !IsOnlyContentOfOuterScriptArgument(script))
                continue;

            // MathJax represents sources such as f_{_{\\mathrm H}} as an
            // outer subscript whose argument contains another subscript with
            // an empty base. Office faithfully renders that empty base as a
            // dotted equation placeholder. Inside an existing script slot the
            // extra empty-base level carries no useful layout information, so
            // replace it with its visible script argument. Standalone empty-
            // base scripts are intentionally preserved for prescript/tensor
            // notation.
            script.ReplaceWith(new XElement(children[1]));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static string NormalizeMathMlPrescripts(string mathMl)
    {
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        XNamespace p = "http://www.w3.org/1998/Math/MathML";
        var scriptNames = new HashSet<XName> { p + "msub", p + "msup", p + "msubsup" };
        // Only these containers infer a sequence of expressions. In a fraction,
        // root, script, etc. the next child is ANOTHER ARGUMENT, not the next atom.
        var sequenceNames = new HashSet<XName>
        {
            p + "math", p + "mrow", p + "mstyle", p + "mtd", p + "msqrt",
            p + "mphantom", p + "mpadded", p + "menclose",
        };
        var baseNames = new HashSet<XName>
        {
            p + "mi", p + "mn", p + "mrow", p + "mstyle", p + "mfrac",
            p + "msqrt", p + "mroot", p + "msub", p + "msup", p + "msubsup",
            p + "mmultiscripts", p + "mover", p + "munder", p + "munderover", p + "mfenced",
        };
        bool IsEmptyBase(XElement element) =>
            (element.Name == p + "mrow" || element.Name == p + "mi"
                || element.Name == p + "mn" || element.Name == p + "mtext")
            && element.Value.Length == 0
            && element.Elements().All(IsEmptyBase);

        foreach (var script in document.Descendants()
                     .Where(element => scriptNames.Contains(element.Name)).Reverse().ToArray())
        {
            if (script.Parent is null || !sequenceNames.Contains(script.Parent.Name)) continue;
            var arguments = script.Elements().ToArray();
            if (arguments.Length != (script.Name == p + "msubsup" ? 3 : 2)
                || !IsEmptyBase(arguments[0])) continue;
            var basis = script.ElementsAfterSelf().FirstOrDefault();
            if (basis is null || !baseNames.Contains(basis.Name)
                || string.IsNullOrWhiteSpace(basis.Value)) continue;
            var sub = script.Name == p + "msup" ? new XElement(p + "none") : new XElement(arguments[1]);
            var sup = script.Name == p + "msub" ? new XElement(p + "none")
                : new XElement(arguments[script.Name == p + "msubsup" ? 2 : 1]);
            if (string.IsNullOrWhiteSpace(sub.Value) && string.IsNullOrWhiteSpace(sup.Value)) continue;

            // {}^a X and {}_b^a X explicitly encode left-side scripts. Convert
            // their source structure to native MathML prescripts so Office emits
            // sPre, rather than permitting a dotted, empty sSup/sSub base. No
            // filler glyphs or changes to the source LaTeX/adjacent operators.
            var native = new XElement(p + "mmultiscripts",
                script.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                    .Select(attribute => new XAttribute(attribute)),
                new XElement(basis), new XElement(p + "mprescripts"), sub, sup);
            script.ReplaceWith(native);
            basis.Remove();
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static readonly HashSet<string> NaryRowParents = new(StringComparer.Ordinal)
    {
        "math", "mrow", "mstyle", "msqrt", "mtd", "mpadded", "menclose", "mphantom", "merror",
    };

    private const string NaryOperandOpeners = "([{⟨⌈⌊";
    private const string NaryOperandClosers = ")]}⟩⌉⌋";
    private const string NaryOperandStops = "=<>≤≥≠≈≡∼≃≅∝→←↔⇒⇐⇔↦∈∉⊂⊆⊃⊇≪≫≺≻⩽⩾,;";
    private const string NaryOperandSigns = "+-−±∓";

    // The run of siblings a big operator applies to when the producer left
    // them loose: up to a relation, a separator or a fence closing something
    // opened before the operator. A sum also ends at a top-level sign; an
    // integral only at one that comes after its differential.
    private static List<XElement> NaryLooseOperand(IEnumerable<XElement> following, string operatorText)
    {
        var isIntegral = operatorText.All(character =>
            character is >= '∫' and <= '∳' or >= '⨋' and <= '⨜');
        var siblings = following.ToArray();
        var operand = new List<XElement>();
        var depth = 0;
        var sawDifferential = false;
        for (var index = 0; index < siblings.Length; index++)
        {
            var element = siblings[index];
            var text = element.Value.Trim();
            if (element.Name.LocalName == "mo" && text.Length == 1)
            {
                if (NaryOperandOpeners.IndexOf(text[0]) >= 0) depth++;
                else if (NaryOperandClosers.IndexOf(text[0]) >= 0)
                {
                    if (depth == 0) break;
                    depth--;
                }
                else if (depth == 0)
                {
                    if (NaryOperandStops.IndexOf(text[0]) >= 0) break;
                    if (text[0] == '.' && index == siblings.Length - 1) break;
                    if (NaryOperandSigns.IndexOf(text[0]) >= 0
                        && operand.Count > 0
                        && (!isIntegral || sawDifferential))
                        break;
                }
            }
            if (depth == 0
                && text is "d" or "ⅆ"
                && !element.HasElements
                && index + 1 < siblings.Length
                && siblings[index + 1].Name.LocalName == "mi")
                sawDifferential = true;
            operand.Add(element);
        }
        return operand;
    }

    internal static string NormalizeNaryArguments(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl)) return mathMl;
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace mathMlNamespace = "http://www.w3.org/1998/Math/MathML";
        var limitNames = new HashSet<XName>
        {
            mathMlNamespace + "munder",
            mathMlNamespace + "mover",
            mathMlNamespace + "munderover",
            mathMlNamespace + "msub",
            mathMlNamespace + "msup",
            mathMlNamespace + "msubsup",
        };
        foreach (var op in document
                     .Descendants(mathMlNamespace + "mo")
                     .Where(element =>
                         !string.IsNullOrEmpty(element.Value)
                         && element.Value.All(character => NaryCharacters.IndexOf(character) >= 0)
                         && (element.Parent is null || !limitNames.Contains(element.Parent.Name)))
                     .ToList())
        {
            // Office only creates a native m:nary for a bare operator when it
            // is carried by a limit structure. Apply that structure to inline
            // formulas too; later normalization hides the synthetic empty
            // limit (NormalizeDisplayNaryOmml).
            var hasArgument = op.ElementsAfterSelf().Any() || LooseOperandGroup(op) is not null;
            var syntheticLimit = new XElement(
                mathMlNamespace + "msub",
                new XElement(op),
                new XElement(mathMlNamespace + "mrow"));
            op.ReplaceWith(syntheticLimit);
            if (!hasArgument)
            {
                syntheticLimit.AddAfterSelf(
                    new XElement(
                        mathMlNamespace + "mrow",
                        new XElement(mathMlNamespace + "mspace", new XAttribute("width", "0em"))));
            }
        }

        // A group that holds the operator alone, in a row with more after it:
        // MathType writes this when the template's operand slot was left empty
        // and the operand typed after the template.
        XElement? LooseOperandGroup(XElement owner) =>
            owner.Parent is { } group
            && group.Name == mathMlNamespace + "mrow"
            && group.Elements().Count() == 1
            && group.Parent is { } outer
            && NaryRowParents.Contains(outer.Name.LocalName)
            && outer.Name.Namespace == mathMlNamespace
            && group.ElementsAfterSelf().Any()
                ? group
                : null;

        bool IsNaryLimit(XElement element)
        {
            if (!limitNames.Contains(element.Name)) return false;
            var op = element.Elements().FirstOrDefault();
            return op?.Name == mathMlNamespace + "mo"
                && !string.IsNullOrEmpty(op.Value)
                && op.Value.All(character => NaryCharacters.IndexOf(character) >= 0);
        }

        // Presentation MathML expresses consecutive operators as a flat sequence:
        // sum_n sum_m <body>. Office's stylesheet consumes only the immediately
        // following mrow as an n-ary operand. Wrapping just the second sum therefore
        // creates an inner m:nary with an empty <m:e/> and leaves <body> outside it.
        // Fold every consecutive n-ary chain from right to left so the innermost
        // operator owns the real body and each outer operator owns that complete
        // nested expression.
        foreach (var parent in document
                     .Descendants()
                     .Where(element => element.Elements().Any())
                     .ToList())
        {
            var children = parent.Elements().ToArray();
            var chains = new List<(int Start, int End, XElement? Operand)>();
            for (var start = 0; start < children.Length; start++)
            {
                if (!IsNaryLimit(children[start])) continue;
                var end = start;
                while (end + 1 < children.Length
                       && IsNaryLimit(children[end + 1]))
                    end++;
                if (end > start)
                {
                    chains.Add((
                        start,
                        end,
                        end + 1 < children.Length ? children[end + 1] : null));
                }
                start = end;
            }

            foreach (var chain in chains.OrderByDescending(item => item.Start))
            {
                XElement operand;
                if (chain.Operand is null)
                {
                    operand = new XElement(
                        mathMlNamespace + "mrow",
                        new XElement(
                            mathMlNamespace + "mspace",
                            new XAttribute("width", "0em")));
                    children[chain.End].AddAfterSelf(operand);
                }
                else if (chain.Operand.Name == mathMlNamespace + "mrow"
                         || chain.Operand.Name == mathMlNamespace + "mstyle")
                {
                    operand = chain.Operand;
                }
                else
                {
                    operand = new XElement(mathMlNamespace + "mrow");
                    chain.Operand.ReplaceWith(operand);
                    operand.Add(chain.Operand);
                }

                for (var index = chain.End; index > chain.Start; index--)
                {
                    var inner = children[index];
                    var outer = children[index - 1];
                    inner.Remove();
                    operand.Remove();
                    var nestedOperand = new XElement(
                        mathMlNamespace + "mrow",
                        inner,
                        operand);
                    outer.AddAfterSelf(nestedOperand);
                    operand = nestedOperand;
                }
            }
        }

        foreach (var limit in document.Descendants().Where(element => limitNames.Contains(element.Name)).ToList())
        {
            var op = limit.Elements().FirstOrDefault();
            if (op?.Name != mathMlNamespace + "mo"
                || string.IsNullOrEmpty(op.Value)
                || op.Value.Any(character => NaryCharacters.IndexOf(character) < 0))
                continue;
            var argument = limit.ElementsAfterSelf().FirstOrDefault();
            if (argument is null && LooseOperandGroup(limit) is { } group)
            {
                var loose = NaryLooseOperand(group.ElementsAfterSelf(), op.Value);
                foreach (var element in loose) element.Remove();
                group.Add(loose.Count > 0
                    ? new XElement(mathMlNamespace + "mrow", loose)
                    : new XElement(
                        mathMlNamespace + "mrow",
                        new XElement(mathMlNamespace + "mspace", new XAttribute("width", "0em"))));
                continue;
            }
            if (argument is null
                || argument.Name == mathMlNamespace + "mrow"
                || argument.Name == mathMlNamespace + "mstyle")
                continue;

            // Office's MML2OMML.XSL recognizes an n-ary operand only when the
            // immediately following sibling is mrow or mstyle. MathJax emits a
            // valid flat Presentation MathML sequence (for example
            // munderover + mi), which Office otherwise converts to <m:e/> and
            // Word displays as a dotted placeholder box.
            argument.ReplaceWith(new XElement(mathMlNamespace + "mrow", argument));
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    internal static (string MathMl, IReadOnlyList<string> NaryCharacters)
        ReplaceExtendedIntegralsWithOfficePlaceholders(string mathMl)
    {
        if (string.IsNullOrWhiteSpace(mathMl))
            return (mathMl, Array.Empty<string>());

        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = false,
            MaxCharactersInDocument = 4_000_000,
        };
        using var text = new StringReader(mathMl);
        using var reader = XmlReader.Create(text, settings);
        var document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var operators = document
            .Descendants(presentationMath + "mo")
            .Where(element =>
                element.Value.Length == 1
                && NaryCharacters.IndexOf(element.Value[0]) >= 0)
            .ToArray();
        var characters = operators.Select(element => element.Value).ToArray();

        foreach (var op in operators)
        {
            if (ExtendedIntegralCharacters.IndexOf(op.Value[0]) >= 0)
            {
                // Office's MML2OMML transform knows how to attach limits and
                // the following operand to a standard integral. Use it only as
                // a structural placeholder; the exact extended character is
                // restored in OMML immediately after the transform.
                op.Value = "∫";
            }
        }

        return (document.ToString(SaveOptions.DisableFormatting), characters);
    }

    internal static string RestoreExtendedIntegralCharacters(
        string omml,
        IReadOnlyList<string> sourceNaryCharacters)
    {
        if (sourceNaryCharacters.Count == 0
            || !sourceNaryCharacters.Any(character =>
                character.Length == 1
                && ExtendedIntegralCharacters.IndexOf(character[0]) >= 0))
            return omml;

        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        var naries = document.Descendants(math + "nary").ToArray();
        if (naries.Length != sourceNaryCharacters.Count)
        {
            throw new InvalidDataException(
                "Office changed the number of n-ary operators while converting extended integrals. "
                + $"MathML={sourceNaryCharacters.Count}; OMML={naries.Length}.");
        }

        for (var index = 0; index < sourceNaryCharacters.Count; index++)
        {
            var sourceCharacter = sourceNaryCharacters[index];
            if (sourceCharacter.Length != 1
                || ExtendedIntegralCharacters.IndexOf(sourceCharacter[0]) < 0)
                continue;

            var properties = naries[index].Element(math + "naryPr");
            if (properties is null)
            {
                properties = new XElement(math + "naryPr");
                naries[index].AddFirst(properties);
            }
            var character = properties.Element(math + "chr");
            if (character is null)
            {
                character = new XElement(math + "chr");
                properties.AddFirst(character);
            }
            character.SetAttributeValue(math + "val", sourceCharacter);
        }

        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static bool IsBlockMathMl(string mathMl)
    {
        try
        {
            using var text = new StringReader(mathMl);
            using var reader = XmlReader.Create(text, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                MaxCharactersInDocument = 4_000_000,
            });
            var document = XDocument.Load(reader, LoadOptions.None);
            return string.Equals(
                document.Root?.Attribute("display")?.Value,
                "block",
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    internal static string NormalizeExplicitTableColumnAlignment(
        string omml,
        string mathMl)
    {
        if (string.IsNullOrWhiteSpace(omml) || string.IsNullOrWhiteSpace(mathMl))
            return omml;

        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XNamespace officeMath = MathNamespace;
        var sourceDocument = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        var targetDocument = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceTables = sourceDocument
            .Descendants(presentationMath + "mtable")
            .Where(table => !string.IsNullOrWhiteSpace(table.Attribute("columnalign")?.Value))
            .Select(table =>
            {
                var rows = table.Elements(presentationMath + "mtr").ToArray();
                var columnCount = rows.Length == 0
                    ? 0
                    : rows.Max(row => row.Elements(presentationMath + "mtd").Count());
                var raw = (table.Attribute("columnalign")?.Value ?? string.Empty)
                    .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (columnCount <= 0 || raw.Length == 0) return null;
                var alignments = Enumerable.Range(0, columnCount)
                    .Select(index => NormalizeMathMlColumnAlignment(
                        raw[Math.Min(index, raw.Length - 1)]))
                    .ToArray();
                return new
                {
                    RowCount = rows.Length,
                    ColumnCount = columnCount,
                    Alignments = alignments,
                };
            })
            .Where(table => table is not null)
            .ToArray();
        if (sourceTables.Length == 0) return omml;

        var matrices = targetDocument.Descendants(officeMath + "m").ToList();
        var used = new HashSet<XElement>();
        foreach (var sourceTable in sourceTables)
        {
            var target = matrices.FirstOrDefault(matrix =>
            {
                if (used.Contains(matrix)) return false;
                var rows = matrix.Elements(officeMath + "mr").ToArray();
                if (rows.Length != sourceTable!.RowCount) return false;
                return rows.All(row =>
                    row.Elements(officeMath + "e").Count() == sourceTable.ColumnCount);
            });
            if (target is null) continue;
            used.Add(target);

            var properties = target.Element(officeMath + "mPr");
            if (properties is null)
            {
                properties = new XElement(officeMath + "mPr");
                target.AddFirst(properties);
            }
            var columns = new XElement(officeMath + "mcs");
            foreach (var alignment in sourceTable!.Alignments)
            {
                columns.Add(
                    new XElement(
                        officeMath + "mc",
                        new XElement(
                            officeMath + "mcPr",
                            new XElement(
                                officeMath + "count",
                                new XAttribute(officeMath + "val", "1")),
                            new XElement(
                                officeMath + "mcJc",
                                new XAttribute(officeMath + "val", alignment)))));
            }
            var existing = properties.Element(officeMath + "mcs");
            if (existing is null) properties.Add(columns);
            else existing.ReplaceWith(columns);
        }

        return targetDocument.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static string NormalizeOptionalPrescriptSlots(string omml)
    {
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var script in document.Descendants(math + "sPre"))
        {
            if (!HasVisibleMathContent(script.Element(math + "e"))) continue;
            foreach (var name in new[] { "sub", "sup" })
            {
                var slot = script.Element(math + name);
                var other = script.Element(math + (name == "sub" ? "sup" : "sub"));
                if (slot is null || !HasVisibleMathContent(other) || HasVisibleMathContent(slot)) continue;
                // Preserve unfamiliar controls/extensions rather than replacing them.
                if (slot.Descendants().Any(element => element.Name != math + "r" && element.Name != math + "t"
                    && element.Name != math + "rPr")) continue;
                // CT_SPre has a pair of slots but no subHide/supHide. Represent the
                // explicitly absent half as an empty native phantom with zero extent.
                // There is no filler character, duplicated operand, font change or
                // visible placeholder in this structural representation of <none/>.
                slot.ReplaceNodes(new XElement(math + "phant",
                    new XElement(math + "phantPr",
                        new XElement(math + "show", new XAttribute(math + "val", "0")),
                        new XElement(math + "zeroWid", new XAttribute(math + "val", "1")),
                        new XElement(math + "zeroAsc", new XAttribute(math + "val", "1")),
                        new XElement(math + "zeroDesc", new XAttribute(math + "val", "1"))),
                    new XElement(math + "e")));
            }
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static bool IsHiddenPhantom(XElement element)
    {
        XNamespace math = MathNamespace;
        var show = element.Element(math + "phantPr")?.Element(math + "show");
        return element.Name == math + "phant" && show is not null && !ReadMathBooleanValue(show);
    }

    private static string[] RequiredOmmlArgumentNames(string ownerName) => ownerName switch
    {
        "f" => new[] { "num", "den" },
        "rad" => new[] { "deg", "e" },
        "nary" => new[] { "sub", "sup", "e" },
        "sSub" => new[] { "e", "sub" },
        "sSup" => new[] { "e", "sup" },
        "sSubSup" => new[] { "e", "sub", "sup" },
        "sPre" => new[] { "sub", "sup", "e" },
        _ => Array.Empty<string>(),
    };

    internal static string NormalizeGeneratedOmmlArgumentOrder(string omml)
    {
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var owner in document.Descendants().Where(element => element.Name.Namespace == math).Reverse().ToArray())
        {
            var roles = RequiredOmmlArgumentNames(owner.Name.LocalName);
            if (roles.Length == 0) continue;
            var propertyName = math + (owner.Name.LocalName + "Pr");
            var children = owner.Elements().ToArray();
            // Normalize the producer's known argument grammar, never reorder
            // operands within an argument or guess through extensions/duplicates.
            if (roles.Any(role => children.Count(child => child.Name == math + role) != 1)
                || children.Count(child => child.Name == propertyName) > 1
                || children.Any(child => child.Name != propertyName && !roles.Any(role => child.Name == math + role))
                || owner.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                continue;
            var ordered = children.Where(child => child.Name == propertyName)
                .Concat(roles.Select(role => children.Single(child => child.Name == math + role))).ToArray();
            if (children.SequenceEqual(ordered)) continue;
            // Office MML2OMML emits sPre as e/sub/sup, but CT_SPre requires
            // sub/sup/e. Word repairs this on import. Produce the proper grammar
            // before capture/hash rather than relaxing the materialization check.
            owner.ReplaceNodes(ordered.Select(child => new XElement(child)));
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    internal static void ValidateNoVisibleEmptyOmmlSlots(XDocument document)
    {
        XNamespace math = MathNamespace;
        var structuralSlotNames = new HashSet<XName>
        {
            math + "e",
            math + "sub",
            math + "sup",
            math + "num",
            math + "den",
            math + "lim",
            math + "fName",
            math + "deg",
        };
        foreach (var owner in document.Descendants().Where(element => element.Name.Namespace == math))
        {
            // An absent required argument is not an intentionally empty argument.
            // Keep this in the shared contract rather than a second post-insert list.
            var required = RequiredOmmlArgumentNames(owner.Name.LocalName);
            foreach (var name in required)
                if (owner.Elements(math + name).Count() != 1)
                    throw new InvalidDataException(
                        $"Office OMML requires exactly one {name} argument inside {owner.Name.LocalName}.");
        }
        foreach (var slot in document.Descendants()
                     .Where(element => structuralSlotNames.Contains(element.Name)))
        {
            if (HasVisibleMathContent(slot)) continue;
            // A phantom explicitly suppresses its inner editor placeholder as well
            // as glyphs. Required-argument cardinality is still checked above.
            if (slot.Ancestors().Any(IsHiddenPhantom)) continue;

            var parent = slot.Parent;
            var hidden = slot.Name == math + "e"
                && parent?.Name == math + "mr"
                && parent.Parent?.Name == math + "m"
                && IsMathBooleanEnabled(
                    parent.Parent.Element(math + "mPr"),
                    math + "plcHide")
                || slot.Name == math + "deg"
                && parent?.Name == math + "rad"
                && IsMathBooleanEnabled(parent.Element(math + "radPr"), math + "degHide")
                || slot.Name == math + "sub"
                && parent?.Name == math + "nary"
                && IsMathBooleanEnabled(parent.Element(math + "naryPr"), math + "subHide")
                || slot.Name == math + "sup"
                && parent?.Name == math + "nary"
                && IsMathBooleanEnabled(parent.Element(math + "naryPr"), math + "supHide")
                // Native prescripts carry a pair of arguments; either half may
                // be absent in the source. The base and the other half must exist.
                || (slot.Name == math + "sub" || slot.Name == math + "sup")
                && parent?.Name == math + "sPre"
                && HasVisibleMathContent(parent.Element(math + "e"))
                && HasVisibleMathContent(parent.Element(slot.Name == math + "sub" ? math + "sup" : math + "sub"));
            if (hidden) continue;

            throw new InvalidDataException(
                $"Office OMML contains a visible empty {slot.Name.LocalName} slot "
                + $"inside {parent?.Name.LocalName ?? "unknown"}.");
        }
    }

    internal static string NormalizeOmmlPlaceholderVisibility(string omml)
    {
        if (string.IsNullOrWhiteSpace(omml)) return omml;
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var matrix in document.Descendants(math + "m"))
        {
            var properties = matrix.Element(math + "mPr");
            if (properties is null)
            {
                properties = new XElement(math + "mPr");
                matrix.AddFirst(properties);
            }
            var placeholderVisibility = properties.Element(math + "plcHide");
            if (placeholderVisibility is null)
            {
                placeholderVisibility = new XElement(math + "plcHide");
                properties.AddFirst(placeholderVisibility);
            }
            placeholderVisibility.SetAttributeValue(math + "val", "1");
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static string NormalizeMathMlColumnAlignment(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "left" => "left",
            "right" => "right",
            _ => "center",
        };
    }

    internal static string NormalizeDisplayNaryOmml(string omml, bool display)
    {
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        XNamespace math = MathNamespace;
        foreach (var nary in document.Descendants(math + "nary"))
        {
            var properties = nary.Element(math + "naryPr");
            if (properties is null)
            {
                properties = new XElement(math + "naryPr");
                nary.AddFirst(properties);
            }

            // A sum or integral keeps its size, as in LaTeX: Word already draws
            // it large in a display equation. With m:grow it stretched to the
            // height of its operand, so beside a fraction, brackets or a matrix
            // the sign became huge while its limits stayed script size. An
            // inline operator follows the paragraph's font size the same way.
            properties.Element(math + "grow")?.Remove();
            // Empty limits must be hidden for both inline and display OMML, or a
            // bare integral exposes Word's dotted sub/sup slots.

            SetNaryLimitVisibility(
                properties,
                math + "subHide",
                !HasNaryLimitContent(nary.Element(math + "sub")));
            SetNaryLimitVisibility(
                properties,
                math + "supHide",
                !HasNaryLimitContent(nary.Element(math + "sup")));
        }
        return document.Root?.ToString(SaveOptions.DisableFormatting) ?? omml;
    }

    private static bool HasNaryLimitContent(XElement? limit)
    {
        if (limit is null) return false;
        return limit
            .DescendantsAndSelf()
            .Where(element => element.Name.LocalName == "t")
            .Any(element => !string.IsNullOrWhiteSpace(element.Value));
    }

    private static void SetNaryLimitVisibility(
        XElement properties,
        XName propertyName,
        bool hidden)
    {
        var property = properties.Element(propertyName);
        if (!hidden)
        {
            property?.Remove();
            return;
        }
        if (property is null)
        {
            property = new XElement(propertyName);
            properties.Add(property);
        }
        property.SetAttributeValue(XName.Get("val", MathNamespace), "1");
    }

    internal static bool HasVisibleMathContent(XElement? slot)
    {
        if (slot is null) return false;
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        return slot.DescendantsAndSelf()
            .Any(element => (element.Name == math + "t" || element.Name == word + "t")
                && element.Value.Length > 0);
    }

    internal static bool IsMathBooleanEnabled(XElement? properties, XName propertyName)
    {
        var property = properties?.Element(propertyName);
        if (property is null) return false;
        return ReadMathBooleanValue(property);
    }

    internal static bool ReadMathBooleanValue(XElement property)
    {
        // Element omission is handled by the caller. A present switch is on
        // unless its value is explicitly false, read case-insensitively as
        // Office does: empty or unknown values never erase content as false,
        // and never make a formula Word displays unreadable.
        var value = property.Attribute(XName.Get("val", MathNamespace))?.Value;
        return value?.ToLowerInvariant() is not ("0" or "false" or "off");
    }
}

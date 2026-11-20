using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// Projects presentation MathML to Office Math (OMML) without Office's
/// MML2OMML.XSL. Office updates replace that stylesheet, so the projection is
/// owned here. It reproduces the structure, run grouping and run properties
/// that VisualTeX's insertion pipeline was built against: Office's classic
/// stylesheet together with VisualTeX's adaptations (literal text, explicit
/// token-style and role boundaries, italic and monospace alphabets, native
/// accent control formatting, projected spaces and text colors).
///
/// Every well-formed MathML document projects. Elements without an OMML
/// counterpart contribute the projection of their children, exactly as the
/// stylesheet's catch-all did.
/// </summary>
internal static class MathMlOmmlProjection
{
    private static readonly XNamespace Mml = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string RunRoleAttribute = "data-visualtex-omml-run-role";
    private const string SpaceTextAttribute = "data-visualtex-omml-space";

    private static readonly string[] NaryOperators =
    {
        "\u222B", "\u222C", "\u222D", "\u222E", "\u222F", "\u2230", "\u2232", "\u2233", "\u2231",
        "\u2229", "\u222A", "\u220F", "\u2210", "\u2211", "\u22C0", "\u22C1", "\u22C2", "\u22C3",
    };

    private static readonly string[] GrowingNaryOperators =
    {
        "\u222B", "\u222E", "\u222F", "\u2232", "\u2233", "\u2229", "\u222A", "\u220F", "\u2211",
        "\u22C0", "\u22C1", "\u22C2", "\u22C3",
    };

    /// <summary>
    /// Projects one MathML document. The result is serialized as the
    /// stylesheet's output was: one m:oMath document element.
    /// </summary>
    internal static string Project(string mathMl)
    {
        XDocument source;
        using (var text = new StringReader(mathMl))
        using (var reader = XmlReader.Create(text, OmmlMathMlProjection.ReaderSettings()))
            source = XDocument.Load(reader, LoadOptions.None);

        // The stylesheet's literal m:oMath carried every namespace in its scope.
        var equation = new XElement(M + "oMath",
            new XAttribute(XNamespace.Xmlns + "m", M.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "mml", Mml.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "w", W.NamespaceName));
        if (source.Root is not null) equation.Add(Apply(source.Root));

        var settings = new XmlWriterSettings { OmitXmlDeclaration = true, Encoding = new UTF8Encoding(false) };
        var output = new StringBuilder();
        using (var writer = XmlWriter.Create(output, settings)) equation.WriteTo(writer);
        return output.ToString();
    }

    // ---- XPath helpers -------------------------------------------------

    private static bool Is(XElement? element, string name) => element is not null && element.Name == Mml + name;

    private static bool IsAny(XElement? element, params string[] names) =>
        element is not null && element.Name.Namespace == Mml && names.Contains(element.Name.LocalName);

    private static bool IsToken(XElement? element) => IsAny(element, "mi", "mn", "mo", "ms", "mtext");

    private static bool HasGlyph(XElement element) => element.Elements(Mml + "mglyph").Any();

    // Office's "choose when @x, value-of @x, otherwise value-of @mml:x".
    private static string Value(XElement? element, string name) =>
        element is null ? string.Empty
            : (string?)element.Attribute(name) ?? (string?)element.Attribute(Mml + name) ?? string.Empty;

    private static string? Plain(XElement? element, string name) => (string?)element?.Attribute(name);
    private static string? Qualified(XElement? element, string name) => (string?)element?.Attribute(Mml + name);

    // @x='v' or @mml:x='v'
    private static bool AttributeIs(XElement element, string name, string value) =>
        Plain(element, name) == value || Qualified(element, name) == value;

    // (not(@x) or @x='') for one attribute
    private static bool Blank(string? value) => string.IsNullOrEmpty(value);

    private static bool BothBlank(XElement element, string name) =>
        Blank(Plain(element, name)) && Blank(Qualified(element, name));

    private static XElement? Child(XElement? element, int position) =>
        element?.Elements().Skip(position - 1).FirstOrDefault();

    private static XElement? PrecedingSibling(XElement element) => element.ElementsBeforeSelf().LastOrDefault();

    private static string StringValue(XElement? element) => element?.Value ?? string.Empty;

    private static string NormalizeSpace(string value)
    {
        var builder = new StringBuilder(value.Length);
        var pendingSpace = false;
        foreach (var character in value)
        {
            if (character is ' ' or '\t' or '\r' or '\n')
            {
                pendingSpace = builder.Length > 0;
                continue;
            }
            if (pendingSpace) builder.Append(' ');
            pendingSpace = false;
            builder.Append(character);
        }
        return builder.ToString();
    }

    private static string LowerAscii(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
            builder.Append(character is >= 'A' and <= 'Z' ? (char)(character + 32) : character);
        return builder.ToString();
    }

    // The first text node of an element, as XPath sees it (adjacent text and
    // CDATA are one node).
    private static string? FirstText(XElement element)
    {
        var builder = new StringBuilder();
        var found = false;
        foreach (var node in element.Nodes())
        {
            if (node is XText text)
            {
                builder.Append(text.Value);
                found = true;
            }
            else if (found) break;
        }
        return found ? builder.ToString() : null;
    }

    // string(number(text())) = 'NaN' as .NET's XSLT number() decides it:
    // optional sign, decimal digits, optional exponent, finite result.
    private static readonly System.Text.RegularExpressions.Regex XPathNumber = new(
        @"^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?$",
        System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static bool IsNotANumber(string? value)
    {
        if (value is null) return true;
        var trimmed = value.Trim(' ', '\t', '\r', '\n');
        if (!XPathNumber.IsMatch(trimmed)) return true;
        var number = double.Parse(trimmed, System.Globalization.NumberStyles.Float,
            System.Globalization.CultureInfo.InvariantCulture);
        return double.IsInfinity(number);
    }

    private static string FirstScalar(string value) =>
        value.Length >= 2 && char.IsHighSurrogate(value[0]) && char.IsLowSurrogate(value[1])
            ? value.Substring(0, 2)
            : value.Length == 0 ? value : value.Substring(0, 1);

    // ---- Template dispatch ---------------------------------------------

    private static List<XNode> Apply(XElement element)
    {
        if (element.Name.Namespace != Mml) return ApplyChildren(element);
        if (IsNativeAccent(element)) return One(NativeAccent(element));
        switch (element.Name.LocalName)
        {
            case "mspace":
                return element.Attribute(SpaceTextAttribute) is not null ? Space(element) : ApplyChildren(element);
            case "mrow":
            case "mstyle":
                return Row(element);
            case "mi":
            case "mn":
            case "mo":
            case "ms":
            case "mtext":
                return HasGlyph(element) ? GlyphToken(element) : Token(element);
            case "mfrac": return One(Fraction(element));
            case "menclose":
            case "msqrt": return Enclosure(element);
            case "mroot": return One(Root(element));
            case "munder": return One(Under(element));
            case "mover": return One(Over(element));
            case "munderover": return One(UnderOver(element));
            case "mfenced": return One(Fenced(element));
            case "msub": return One(Script(element, "msub"));
            case "msup": return One(Script(element, "msup"));
            case "msubsup": return One(Script(element, "msubsup"));
            case "mmultiscripts": return MultiScripts(element);
            case "mtable": return One(Table(element));
            case "mglyph": return One(StandaloneGlyph(element));
            case "mpadded": return Padded(element);
            case "mphantom": return One(Phantom(element));
            default: return ApplyChildren(element);
        }
    }

    private static List<XNode> ApplyChildren(XElement element)
    {
        var nodes = new List<XNode>();
        foreach (var child in element.Elements()) nodes.AddRange(Apply(child));
        return nodes;
    }

    private static List<XNode> ApplyOptional(XElement? element) =>
        element is null ? new List<XNode>() : Apply(element);

    private static List<XNode> ApplyAll(IEnumerable<XElement> elements)
    {
        var nodes = new List<XNode>();
        foreach (var element in elements) nodes.AddRange(Apply(element));
        return nodes;
    }

    private static List<XNode> One(XNode node) => new() { node };

    private static XElement Om(string name, params object?[] content) => new(M + name, content);

    private static XAttribute Val(string value) => new(M + "val", value);

    // ---- Rows, functions, linear fractions ------------------------------

    private static List<XNode> Row(XElement row)
    {
        if (IsNaryArgument(row)) return new List<XNode>();
        if (IsLinearFraction(row)) return One(LinearFraction(row, row));
        if (IsFunction(row)) return One(Function(row));
        return ApplyChildren(row);
    }

    private static bool IsNaryArgument(XElement element)
    {
        var previous = PrecedingSibling(element);
        return IsAny(previous, "munder", "mover", "munderover", "msub", "msup", "msubsup")
            && IsNary(Child(previous, 1));
    }

    private static bool IsLinearFraction(XElement? element) =>
        Is(element, "mrow")
        && element!.Elements().Count() == 3
        && Is(Child(element, 2), "mo")
        && NormalizeSpace(StringValue(Child(element, 2))) == "/";

    private static bool IsFunction(XElement? element) =>
        element is not null
        && element.Elements().Count() == 3
        && Is(element, "mrow")
        && Is(Child(element, 2), "mo")
        && NormalizeSpace(StringValue(Child(element, 2))) == "\u2061";

    private static XElement LinearFraction(XElement fraction, XElement context) => Om("f",
        Om("fPr", Om("type", Val("lin")), ControlColor(context)),
        Om("num", ArgumentProperties(context), ApplyOptional(Child(fraction, 1))),
        Om("den", ArgumentProperties(context), ApplyOptional(Child(fraction, 3))));

    private static XElement Function(XElement function) => Om("func",
        Om("fName", ApplyOptional(Child(function, 1))),
        Om("e", ApplyOptional(Child(function, 3))));

    // m:argPr for an explicit mstyle script level 0, 1 or 2 on the context.
    private static XElement? ArgumentProperties(XElement context)
    {
        var styles = context.AncestorsAndSelf(Mml + "mstyle").ToArray();
        bool Level(string? value) => value is "0" or "1" or "2";
        if (!styles.Any(style => Level(Plain(style, "scriptlevel")))
            && !styles.Any(style => Level(Qualified(style, "scriptlevel"))))
            return null;
        var nearest = styles.FirstOrDefault(style => style.Attribute("scriptlevel") is not null);
        return Om("argPr", Om("scrLvl", Val(Plain(nearest, "scriptlevel") ?? Qualified(nearest, "scriptlevel") ?? string.Empty)));
    }

    // VisualTeX's text color for a fraction or radical control character.
    private static XElement? ControlColor(XElement context)
    {
        var color = Plain(context, "mathcolor") ?? string.Empty;
        return color.Length == 0 ? null : Om("ctrlPr", WordColor(color));
    }

    private static XElement WordColor(string mathColor) =>
        new(W + "rPr", new XElement(W + "color", new XAttribute(W + "val", mathColor.Substring(1))));

    // ---- Tokens and runs -----------------------------------------------

    private static bool IsNonGlyphToken(XElement? element) => IsToken(element) && !HasGlyph(element!);

    private static List<XNode> Token(XElement token)
    {
        var startOfRun = !token.ElementsBeforeSelf().Any() || !IsNonGlyphToken(PrecedingSibling(token));
        var parent = token.Parent;
        var collect = !IsLinearFraction(parent) && !IsFunction(parent)
            && IsAny(parent, "mrow", "mstyle", "msqrt", "menclose", "math", "mphantom", "mtd", "maction");
        if (collect)
            return startOfRun ? RunsFrom(token) : new List<XNode>();

        // A token outside a run is always whitespace-normalized; literal text
        // only skips Office's operator rewrites.
        var text = NormalizeSpace(StringValue(token));
        return One(Om("r",
            RunProperties(Value(token, "mathvariant"), Value(token, "fontstyle"), Value(token, "fontweight"),
                token, Is(token, "mtext")),
            Om("t", OutputText(text, IsAny(token, "mtext", "ms")))));
    }

    // Office's CreateRunWithSameProp: consecutive tokens with matching run
    // properties share one m:r.
    private static List<XNode> RunsFrom(XElement first)
    {
        var runs = new List<XNode>();
        var current = first;
        while (true)
        {
            var siblings = current.ElementsAfterSelf().ToList();
            var limit = siblings.FindIndex(sibling => EndsRun(current, sibling));
            var members = limit < 0
                ? siblings.Where(IsToken).ToList()
                : siblings.Take(limit).ToList();

            var literal = IsAny(current, "mtext", "ms");
            var text = new StringBuilder(TokenRunText(current, literal));
            foreach (var member in members)
                text.Append(Is(member, "ms") ? QuotedLiteral(member)
                    : IsAny(member, "mtext", "ms") ? StringValue(member) : NormalizeSpace(StringValue(member)));
            runs.Add(Om("r",
                RunProperties(Value(current, "mathvariant"), Value(current, "fontstyle"), Value(current, "fontweight"),
                    current, Is(current, "mtext")),
                Om("t", OutputText(text.ToString(), literal))));

            if (limit < 0) break;
            var next = siblings[limit];
            if (!IsToken(next) || next.Elements(Mml + "mglyph").Any()) break;
            current = next;
        }
        return runs;
    }

    private static string TokenRunText(XElement token, bool literal) =>
        Is(token, "ms") ? QuotedLiteral(token)
            : literal ? StringValue(token) : NormalizeSpace(StringValue(token));

    private static string QuotedLiteral(XElement ms) =>
        LeftQuote(ms) + NormalizeSpace(StringValue(ms)) + RightQuote(ms);

    private static string LeftQuote(XElement ms) =>
        BothBlank(ms, "lquote") ? "\"" : Value(ms, "lquote");

    private static string RightQuote(XElement ms) =>
        BothBlank(ms, "rquote") ? "\"" : Value(ms, "rquote");

    // The stylesheet's grouping predicate: true when sibling must not join
    // the run that starts at first.
    private static bool EndsRun(XElement first, XElement sibling)
    {
        if ((Plain(sibling, "mathcolor") ?? string.Empty) != (Plain(first, "mathcolor") ?? string.Empty)) return true;
        if ((Plain(sibling, "mathvariant") ?? string.Empty) != (Plain(first, "mathvariant") ?? string.Empty)) return true;
        if ((Plain(sibling, RunRoleAttribute) ?? string.Empty) != (Plain(first, RunRoleAttribute) ?? string.Empty)) return true;
        if (Is(sibling, "mo") && NormalizeSpace(StringValue(sibling)) == "\u2061") return true;
        if (!IsToken(sibling)) return true;
        if (HasGlyph(sibling)) return true;
        var firstIsText = Is(first, "mtext");
        if (firstIsText != Is(sibling, "mtext")) return true;

        var font = FontOf(Value(first, "mathvariant"), Value(first, "fontstyle"), Value(first, "fontweight"), first);
        return !(FontMatches(font, sibling) && FamilyMatches(Value(first, "font-family"), sibling));
    }

    private static bool FontMatches(string font, XElement s)
    {
        var noVariant = BothBlank(s, "mathvariant");
        var notBold = Plain(s, "fontweight") != "bold" && Qualified(s, "fontweight") != "bold";
        var boldWeight = AttributeIs(s, "fontweight", "bold");
        var italicStyle = AttributeIs(s, "fontstyle", "italic");
        var length = NormalizeSpace(StringValue(s)).Length;
        var multiLetter = Is(s, "mi") && length > 1;
        var singleLetter = Is(s, "mi") && length <= 1;
        var numeral = Is(s, "mn") && !IsNotANumber(FirstText(s));
        var nonNumeral = Is(s, "mn") && IsNotANumber(FirstText(s));
        var noWeight = BothBlank(s, "fontweight");

        if (Plain(s, "mathvariant") is { } plainVariant && plainVariant == font) return true;
        if (Qualified(s, "mathvariant") is { } qualifiedVariant && qualifiedVariant == font) return true;
        if (font == "normal" && (AttributeIs(s, "mathvariant", "normal")
                || noVariant && ((AttributeIs(s, "fontstyle", "normal") && notBold) || multiLetter || nonNumeral)))
            return true;
        if (font == "italic" && (AttributeIs(s, "mathvariant", "italic")
                || noVariant && ((italicStyle && notBold) || numeral || Is(s, "mo") || singleLetter)))
            return true;
        if (font == "bold" && (AttributeIs(s, "mathvariant", "bold")
                || noVariant && boldWeight && (AttributeIs(s, "fontstyle", "normal") || singleLetter)))
            return true;
        if (font is "bi" or "bold-italic" && (AttributeIs(s, "mathvariant", "bold-italic")
                || noVariant && ((boldWeight && italicStyle)
                    || (boldWeight && (Is(s, "mn") || Is(s, "mo") || singleLetter)))))
            return true;
        if (font.Length == 0 && ((noVariant && BothBlank(s, "fontstyle") && noWeight)
                || AttributeIs(s, "mathvariant", "italic")
                || noVariant && ((AttributeIs(s, "fontweight", "normal") && italicStyle)
                    || (noWeight && italicStyle)
                    || (noWeight && BothBlank(s, "fontstyle")))))
            return true;
        if (font == "normal"
            && ((Is(s, "mi") && Blank(Plain(s, "mathvariant")) && BothBlank(s, "fontstyle")
                    && noWeight && length > 1)
                || (IsAny(s, "ms", "mtext") && Blank(Plain(s, "mathvariant")) && Blank(Plain(s, "fontstyle"))
                    && Blank(Qualified(s, "fontweight")))))
            return true;
        return false;
    }

    private static bool FamilyMatches(string family, XElement s) =>
        (Plain(s, "font-family") is { } plain && plain == family)
        || (Qualified(s, "font-family") is { } qualified && qualified == family)
        || (family.Length == 0 && BothBlank(s, "font-family"));

    // Office's GetFontCur.
    private static string FontOf(string variant, string style, string weight, XObject? node)
    {
        if (variant.Length > 0) return variant;
        if (node is null) return "italic";
        var element = node as XElement;
        if ((Is(element, "mi") && NormalizeSpace(StringValue(element)).Length <= 1)
            || (Is(element, "mn") && !IsNotANumber(FirstText(element!)))
            || Is(element, "mo"))
        {
            if (style == "normal" && weight == "bold") return "bold";
            if (style == "normal") return "normal";
            if (weight == "bold") return "bi";
            return "italic";
        }
        if (style == "italic" && weight == "bold") return "bi";
        if (style == "italic") return "italic";
        if (weight == "bold") return "bold";
        return "normal";
    }

    // Office's CreateRunProp with VisualTeX's run color.
    private static IEnumerable<XElement> RunProperties(
        string variant, string style, string weight, XObject node, bool normalText)
    {
        var font = FontOf(variant, style, weight, node);
        if (normalText || font.Length > 0)
            yield return Om("rPr", normalText ? Om("nor") : null, ScriptAndStyle(font, normalText));
        var color = (node as XElement)?.Attribute("mathcolor")?.Value ?? string.Empty;
        if (color.Length > 0) yield return WordColor(color);
    }

    // Office's CreateMathScrStyProp with VisualTeX's monospace and italic.
    private static IEnumerable<XElement> ScriptAndStyle(string font, bool normalText)
    {
        XElement Scr(string value) => Om("scr", Val(value));
        XElement Sty(string value) => Om("sty", Val(value));
        switch (font)
        {
            case "normal" when !normalText: return new[] { Sty("p") };
            case "bold": return new[] { Sty("b") };
            case "monospace": return new[] { Scr("monospace"), Sty("p") };
            case "italic": return new[] { Sty("i") };
            case "script": return new[] { Scr("script") };
            case "bold-script": return new[] { Scr("script"), Sty("b") };
            case "double-struck": return new[] { Scr("double-struck"), Sty("p") };
            case "fraktur": return new[] { Scr("fraktur"), Sty("p") };
            case "bold-fraktur": return new[] { Scr("fraktur"), Sty("b") };
            case "sans-serif": return new[] { Scr("sans-serif"), Sty("p") };
            case "bold-sans-serif": return new[] { Scr("sans-serif"), Sty("b") };
            case "sans-serif-italic": return new[] { Scr("sans-serif") };
            case "sans-serif-bold-italic": return new[] { Scr("sans-serif"), Sty("bi") };
            case "bi":
            case "bold-italic": return new[] { Sty("bi") };
            default: return Array.Empty<XElement>();
        }
    }

    // Office's OutputText; literal text (mtext, ms) is written as authored.
    private static string OutputText(string input, bool literal)
    {
        if (literal) return input;
        var builder = new StringBuilder(input.Length);
        foreach (var character in input)
        {
            if (character is '\u2062' or '\u200B') continue;
            builder.Append(character == '\u00A0' ? ' ' : character);
        }
        return builder.ToString().Replace("\u2A75", "==");
    }

    private static List<XNode> GlyphToken(XElement token)
    {
        var variant = Value(token, "mathvariant");
        var style = Value(token, "fontstyle");
        var weight = Value(token, "fontweight");
        var normalText = Is(token, "mtext");
        var literal = IsAny(token, "mtext", "ms");
        var nodes = new List<XNode>();
        if (Is(token, "ms"))
            nodes.Add(Om("r", RunProperties(variant, style, weight, token, normalText),
                Om("t", OutputText(LeftQuote(token), literal))));
        foreach (var item in token.Nodes())
        {
            string text;
            bool forceNormal;
            if (item is XElement glyph && glyph.Name == Mml + "mglyph")
            {
                text = NormalizeSpace(Value(glyph, "alt"));
                forceNormal = true;
            }
            else if (item is XText textNode)
            {
                text = NormalizeSpace(textNode.Value);
                forceNormal = false;
            }
            else continue;
            if (text.Length == 0) continue;
            nodes.Add(Om("r", RunProperties(variant, style, weight, item, forceNormal || normalText),
                Om("t", OutputText(text, false))));
        }
        if (Is(token, "ms"))
            nodes.Add(Om("r", RunProperties(variant, style, weight, token, normalText),
                Om("t", OutputText(RightQuote(token), literal))));
        return nodes;
    }

    private static XElement StandaloneGlyph(XElement glyph) => Om("r",
        Om("rPr", Om("nor")),
        Om("t", OutputText(NormalizeSpace(Value(glyph, "alt")), false)));

    private static List<XNode> Space(XElement space)
    {
        var text = (string?)space.Attribute(SpaceTextAttribute) ?? string.Empty;
        if (text.Length == 0) return new List<XNode>();
        return One(Om("r", Om("t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
    }

    // ---- Fractions, radicals, enclosures ---------------------------------

    private static XElement Fraction(XElement fraction)
    {
        var thickness = LowerAscii(Value(fraction, "linethickness"));
        var bar = thickness.Length == 0 || thickness is "thin" or "medium" or "thick"
            || thickness.IndexOfAny("123456789".ToCharArray()) >= 0;
        var type = !bar ? "noBar"
            : AttributeIs(fraction, "bevelled", "true") ? "skw" : "bar";
        return Om("f",
            Om("fPr", Om("type", Val(type)), ControlColor(fraction)),
            Om("num", ArgumentProperties(fraction), ApplyOptional(Child(fraction, 1))),
            Om("den", ArgumentProperties(fraction), ApplyOptional(Child(fraction, 2))));
    }

    private static List<XNode> Enclosure(XElement enclosure)
    {
        var notation = LowerAscii(Value(enclosure, "notation"));
        if (notation == "radical" || notation.Length == 0 || Is(enclosure, "msqrt"))
            return One(Om("rad",
                Om("radPr", Om("degHide", Val("on")), ControlColor(enclosure)),
                Om("deg", ArgumentProperties(enclosure)),
                Om("e", ArgumentProperties(enclosure), ApplyChildren(enclosure))));
        if (notation is "actuarial" or "longdiv") return new List<XNode>();

        bool Has(string part) => notation.IndexOf(part, StringComparison.Ordinal) >= 0;
        var box = Has("box") || Has("circle") || Has("roundedbox");
        var top = Has("top");
        var bottom = Has("bottom");
        var left = Has("left");
        var right = Has("right");
        var strikeH = Has("horizontalstrike");
        var strikeV = Has("verticalstrike");
        var strikeBltr = Has("updiagonalstrike");
        var strikeTlbr = Has("downdiagonalstrike");
        XElement? properties = null;
        if (strikeH || strikeV || strikeBltr || strikeTlbr || (!box && !(top && bottom && left && right)))
        {
            properties = Om("borderBoxPr");
            if (!box)
            {
                if (!top) properties.Add(Om("hideTop", Val("on")));
                if (!bottom) properties.Add(Om("hideBot", Val("on")));
                if (!left) properties.Add(Om("hideLeft", Val("on")));
                if (!right) properties.Add(Om("hideRight", Val("on")));
            }
            if (strikeH) properties.Add(Om("strikeH", Val("on")));
            if (strikeV) properties.Add(Om("strikeV", Val("on")));
            if (strikeBltr) properties.Add(Om("strikeBLTR", Val("on")));
            if (strikeTlbr) properties.Add(Om("strikeTLBR", Val("on")));
        }
        return One(Om("borderBox", properties,
            Om("e", ArgumentProperties(enclosure), ApplyChildren(enclosure))));
    }

    private static XElement Root(XElement root) => Om("rad",
        Om("radPr", Om("degHide", Val("off")), ControlColor(root)),
        Om("deg", ArgumentProperties(root), ApplyOptional(Child(root, 2))),
        Om("e", ArgumentProperties(root), ApplyOptional(Child(root, 1))));

    // ---- Under, over, scripts and n-ary operators -------------------------

    private static string AccentOf(XElement? element, bool under) =>
        element is null ? string.Empty
            : LowerAscii(Plain(element, under ? "accentunder" : "accent")
                ?? Qualified(element, under ? "accentunder" : "accent") ?? string.Empty);

    private static bool IsNary(XElement? operand)
    {
        if (operand is null) return false;
        var text = NormalizeSpace(StringValue(operand));
        if (!NaryOperators.Contains(text)) return false;
        var parent = operand.Parent;
        var accent = Is(parent, "munder") ? AccentOf(parent, under: true) : AccentOf(parent, under: false);
        if (accent == "true") return false;
        var subtree = operand.DescendantsAndSelf().ToArray();
        return Is(subtree[subtree.Length - 1], "mo")
            && subtree.All(element => IsAny(element, "mo", "mstyle", "mrow"));
    }

    private static XElement Nary(XElement owner, string type, XElement? sub, XElement? sup)
    {
        var symbol = NormalizeSpace(StringValue(Child(owner, 1)));
        var growValue = LowerAscii(Value(Child(owner, 1), "stretchy"));
        var grow = growValue == "true" ? "1" : growValue == "false" ? "0"
            : GrowingNaryOperators.Contains(symbol) ? "1" : "0";
        var underOver = type is "munder" or "mover" or "munderover";
        return Om("nary",
            Om("naryPr",
                Om("chr", Val(symbol)),
                Om("limLoc", Val(underOver ? "undOvr" : "subSup")),
                Om("grow", Val(grow)),
                Om("subHide", Val(type is "mover" or "msup" ? "on" : "off")),
                Om("supHide", Val(type is "munder" or "msub" ? "on" : "off"))),
            Om("sub", ArgumentProperties(owner), ApplyOptional(sub)),
            Om("sup", ArgumentProperties(owner), ApplyOptional(sup)),
            Om("e", ArgumentProperties(owner), NaryArgument(owner.ElementsAfterSelf().FirstOrDefault(), owner)));
    }

    // Office's NaryHandleMrowMstyle: the following row is the n-ary's argument.
    private static List<XNode> NaryArgument(XElement? argument, XElement owner)
    {
        if (Is(argument, "mrow"))
            return IsLinearFraction(argument) ? One(LinearFraction(argument!, owner)) : ApplyChildren(argument!);
        if (Is(argument, "mstyle")) return ApplyChildren(argument!);
        return new List<XNode>();
    }

    private static bool IsBar(XElement element, bool under)
    {
        if (AccentOf(element, under) == "true") return false;
        var mark = Child(element, 2);
        if (!Is(mark, "mo")) return false;
        var text = StringValue(mark);
        return under ? text is "\u0332" or "_" : text is "\u0305" or "\u00AF";
    }

    // child::*[2] = mml:mo compares string values with every mo child.
    private static bool SecondEqualsSomeOperator(XElement element)
    {
        var second = Child(element, 2);
        if (second is null) return false;
        var value = StringValue(second);
        return element.Elements(Mml + "mo").Any(mo => StringValue(mo) == value);
    }

    private static bool IsAccent(XElement over)
    {
        var accent = AccentOf(over, under: false);
        var second = Child(over, 2);
        var markAccent = string.Empty;
        if (SecondEqualsSomeOperator(over))
        {
            if (second!.Attribute("accent") is { } plain) markAccent = LowerAscii(plain.Value);
            else if (second.Attribute(Mml + "accent") is { } qualified) markAccent = LowerAscii(qualified.Value);
        }
        var accented = markAccent == "true" || (markAccent.Length == 0 && accent == "true");
        return accented && SecondEqualsSomeOperator(over) && StringValue(second).Length <= 1;
    }

    private static bool IsGroupCharacter(XElement element, bool under)
    {
        if (AccentOf(element, under) != "false") return false;
        var children = element.Elements().ToArray();
        if (children.Length != 2) return false;
        if (!((Is(children[0], "mrow") && Is(children[1], "mo")) || (Is(children[0], "mo") && Is(children[1], "mrow"))))
            return false;
        return StringValue(element.Element(Mml + "mo")).Length <= 1;
    }

    private static XElement GroupCharacter(XElement element, string position, string justification) => Om("groupChr",
        Om("groupChrPr",
            Om("chr", Val(StringValue(element.Element(Mml + "mo")))),
            Om("pos", Val(position)),
            Om("vertJc", Val(justification))),
        Om("e", ApplyAll(element.Elements(Mml + "mrow"))));

    private static XElement Under(XElement under)
    {
        if (IsNary(Child(under, 1))) return Nary(under, "munder", Child(under, 2), null);
        if (IsBar(under, under: true))
            return Om("bar", Om("barPr", Om("pos", Val("bot"))),
                Om("e", ArgumentProperties(under), ApplyOptional(Child(under, 1))));
        if (IsGroupCharacter(under, under: true))
            return GroupCharacter(under, Is(Child(under, 1), "mrow") ? "bot" : "top", "top");
        return Om("limLow",
            Om("e", ArgumentProperties(under), ApplyOptional(Child(under, 1))),
            Om("lim", ArgumentProperties(under), ApplyOptional(Child(under, 2))));
    }

    // Office's ToUpperCombining.
    private static string ToUpperCombining(string character) => character switch
    {
        "\u02D8" => "\u0306",
        "\u00B8" => "\u0312",
        "`" => "\u0300",
        "-" or "\u2212" => "\u0305",
        "." or "\u02D9" => "\u0307",
        "\u02DD" => "\u030B",
        "\u00B4" => "\u0301",
        "~" or "\u02DC" => "\u0303",
        "\u00A8" => "\u0308",
        "\u02C7" => "\u030C",
        "^" => "\u0302",
        "\u00AF" => "\u0305",
        "\u2192" or "\u27F6" => "\u20D7",
        "\u2190" => "\u20D6",
        _ => character,
    };

    private static XElement Over(XElement over)
    {
        if (IsNary(Child(over, 1))) return Nary(over, "mover", null, Child(over, 2));
        if (IsBar(over, under: false))
            return Om("bar", Om("barPr", Om("pos", Val("top"))),
                Om("e", ArgumentProperties(over), ApplyOptional(Child(over, 1))));
        if (IsAccent(over))
            return Om("acc",
                Om("accPr", Om("chr", Val(ToUpperCombining(StringValue(Child(over, 2)))))),
                Om("e", ArgumentProperties(over), ApplyOptional(Child(over, 1))));
        if (IsGroupCharacter(over, under: false))
            return GroupCharacter(over, Is(Child(over, 1), "mrow") ? "top" : "bot", "bot");
        return Om("limUpp",
            Om("e", ArgumentProperties(over), ApplyOptional(Child(over, 1))),
            Om("lim", ArgumentProperties(over), ApplyOptional(Child(over, 2))));
    }

    // VisualTeX's native accent: a one-character styled mark whose control
    // formatting carries the mark's own bold, italic and color.
    private static bool IsNativeAccent(XElement element)
    {
        if (!Is(element, "mover") || Plain(element, "accent") != "true") return false;
        var children = element.Elements().ToArray();
        return children.Length == 2
            && StringValue(children[1]).Length == 1
            && Is(children[1], "mo")
            && Plain(children[1], "mathvariant") is "normal" or "bold" or "italic" or "bold-italic";
    }

    private static XElement NativeAccent(XElement over)
    {
        var mark = Child(over, 2)!;
        var variant = Plain(mark, "mathvariant");
        XElement Flag(string name, bool on) => new(W + name, new XAttribute(W + "val", on ? "1" : "0"));
        var color = Plain(mark, "mathcolor") ?? string.Empty;
        return Om("acc",
            Om("accPr",
                Om("chr", Val(StringValue(mark))),
                Om("ctrlPr", new XElement(W + "rPr",
                    Flag("b", variant is "bold" or "bold-italic"),
                    Flag("i", variant is "italic" or "bold-italic"),
                    color.Length > 0 ? new XElement(W + "color", new XAttribute(W + "val", color.Substring(1))) : null))),
            Om("e", ApplyOptional(Child(over, 1))));
    }

    private static XElement UnderOver(XElement element)
    {
        if (IsNary(Child(element, 1))) return Nary(element, "munderover", Child(element, 2), Child(element, 3));
        return Om("limUpp",
            Om("e", ArgumentProperties(element),
                Om("limLow",
                    Om("e", ArgumentProperties(element), ApplyOptional(Child(element, 1))),
                    Om("lim", ArgumentProperties(element), ApplyOptional(Child(element, 2))))),
            Om("lim", ArgumentProperties(element), ApplyOptional(Child(element, 3))));
    }

    private static XElement Script(XElement script, string type)
    {
        var baseElement = Child(script, 1);
        if (IsNary(baseElement))
            return type switch
            {
                "msub" => Nary(script, type, Child(script, 2), null),
                "msup" => Nary(script, type, null, Child(script, 2)),
                _ => Nary(script, type, Child(script, 2), Child(script, 3)),
            };
        XElement Slot(string name, int position) =>
            Om(name, ArgumentProperties(script), ApplyOptional(Child(script, position)));
        return type switch
        {
            "msub" => Om("sSub", Slot("e", 1), Slot("sub", 2)),
            "msup" => Om("sSup", Slot("e", 1), Slot("sup", 2)),
            _ => Om("sSubSup", Slot("e", 1), Slot("sub", 2), Slot("sup", 3)),
        };
    }

    private static List<XNode> MultiScripts(XElement multi)
    {
        var children = multi.Elements().ToArray();
        var prescriptsIndex = Array.FindIndex(children, child => Is(child, "mprescripts"));
        var prescripts = prescriptsIndex < 0 ? Array.Empty<XElement>() : children.Skip(prescriptsIndex + 1).ToArray();
        var prescriptCount = prescripts.Count(child => !Is(child, "none"));
        bool BeforePrescripts(int index) => prescriptsIndex < 0 || index < prescriptsIndex;
        var superscripts = Enumerable.Range(0, children.Length).Count(index =>
            BeforePrescripts(index) && (index + 1) % 2 == 1 && !Is(children[index], "none")) - 1;
        var subscripts = Enumerable.Range(0, children.Length).Count(index =>
            BeforePrescripts(index) && (index + 1) % 2 == 0 && !Is(children[index], "none"));
        var scriptCount = superscripts + subscripts;
        var scripts = Enumerable.Range(0, children.Length).Count(BeforePrescripts) - 1;

        XElement BaseSlot() => Om("e", ArgumentProperties(multi), ApplyOptional(Child(multi, 1)));
        IEnumerable<XElement> Split(IList<XElement> nodes) => new[]
        {
            Om("sub", ArgumentProperties(multi), ApplyAll(nodes.Where((_, index) => (index + 1) % 2 == 1))),
            Om("sup", ArgumentProperties(multi), ApplyAll(nodes.Where((_, index) => (index + 1) % 2 == 0))),
        };

        if (prescriptCount <= 0 && scriptCount <= 0) return ApplyOptional(Child(multi, 1));
        if (prescriptCount <= 0)
        {
            var rest = children.Skip(1).ToList();
            if (superscripts > 0 && subscripts > 0) return One(Om("sSubSup", BaseSlot(), Split(rest)));
            if (subscripts > 0) return One(Om("sSub", BaseSlot(), Om("sub", ApplyAll(rest))));
            return One(Om("sSup", BaseSlot(), Om("sup", ApplyAll(rest))));
        }
        if (scriptCount <= 0) return One(Om("sPre", BaseSlot(), Split(prescripts)));

        var postscripts = children.Skip(1).Take(scripts).ToList();
        XElement inner;
        if (superscripts > 0 && subscripts > 0) inner = Om("sSubSup", BaseSlot(), Split(postscripts));
        else if (subscripts > 0) inner = Om("sSub", BaseSlot(), Om("sub", ApplyAll(postscripts)));
        else inner = Om("sSup", BaseSlot(), Om("sup", ApplyAll(postscripts)));
        return One(Om("sPre", Om("e", inner), Split(prescripts)));
    }

    // ---- Fences ---------------------------------------------------------

    private static XElement Fenced(XElement fence)
    {
        bool Valid(string name) => fence.Attribute(name) is not null || fence.Attribute(Mml + name) is not null;
        var openValid = Valid("open");
        var open = Value(fence, "open");
        var separatorsValid = Valid("separators");
        var separator = FirstScalar(Value(fence, "separators"));
        var closeValid = Valid("close");
        var close = Value(fence, "close");

        var delimiter = Om("d");
        if ((openValid && open != "(") || (closeValid && close != ")") || separator != "|")
        {
            var properties = Om("dPr");
            if (openValid && open != "(") properties.Add(Om("begChr", Val(open)));
            if (separator != "|") properties.Add(Om("sepChr", Val(separatorsValid ? separator : ",")));
            if (closeValid && close != ")") properties.Add(Om("endChr", Val(close)));
            delimiter.Add(properties);
        }
        foreach (var child in fence.Elements())
            delimiter.Add(Om("e", ArgumentProperties(child), Apply(child)));
        return delimiter;
    }

    // ---- Tables and equation arrays --------------------------------------

    private static bool IsEquationArray(XElement table)
    {
        bool NoLines(string name)
        {
            var plain = Plain(table, name);
            var qualified = Qualified(table, name);
            return (Blank(plain) || plain == "none") && (Blank(qualified) || qualified == "none");
        }
        return NoLines("frame") && NoLines("columnlines") && NoLines("rowlines")
            && table.Elements(Mml + "mtr").All(row => row.Elements(Mml + "mtd").Count() == 1)
            && !table.Elements(Mml + "mlabeledtr").Any();
    }

    private static XElement Table(XElement table)
    {
        if (IsEquationArray(table))
            return Om("eqArr", table.Elements(Mml + "mtr")
                .Select(row => Om("e", EquationArrayRow(row.Elements(Mml + "mtd").SelectMany(cell => cell.Elements())))));

        var columns = 0;
        foreach (var child in table.Elements())
        {
            var count = Is(child, "mlabeledtr") ? child.Elements().Count() - 1
                : Is(child, "mtr") ? child.Elements().Count() : 1;
            if (count > columns) columns = count;
        }
        var matrix = Om("m", Om("mPr",
            Om("baseJc", Val("center")),
            Om("plcHide", Val("on")),
            Om("mcs", Om("mc", Om("mcPr", Om("count", Val(columns.ToString(System.Globalization.CultureInfo.InvariantCulture))), Om("mcJc", Val("center")))))));
        foreach (var child in table.Elements())
        {
            var row = Om("mr");
            IEnumerable<XElement> cells = Is(child, "mtr") ? child.Elements()
                : Is(child, "mlabeledtr") ? child.Elements().Skip(1)
                : new[] { child };
            var used = 0;
            foreach (var cell in cells)
            {
                row.Add(Om("e", Apply(cell)));
                used++;
            }
            for (var index = used; index < columns; index++) row.Add(Om("e"));
            matrix.Add(row);
        }
        return matrix;
    }

    private static bool HasAlignmentDescendant(XElement element) =>
        element.Descendants().Any(descendant => IsAny(descendant, "maligngroup", "malignmark"));

    private static bool FirstAlignmentAlreadyFound(XElement? element)
    {
        while (element is not null)
        {
            if (element.ElementsBeforeSelf().Any(sibling =>
                    IsAny(sibling, "maligngroup", "malignmark") || HasAlignmentDescendant(sibling)))
                return true;
            if (Is(element.Parent, "mtd")) return false;
            element = element.Parent;
        }
        return false;
    }

    // Office's ProcessEqArrayRow: alignment points become '&'.
    private static List<XNode> EquationArrayRow(IEnumerable<XElement> items)
    {
        var nodes = new List<XNode>();
        foreach (var item in items)
        {
            var special = Is(item, "mrow") && (IsLinearFraction(item) || IsFunction(item));
            if (IsAny(item, "maligngroup", "malignmark"))
            {
                if (Is(item, "malignmark") || FirstAlignmentAlreadyFound(item))
                    nodes.Add(Om("r", Om("t", "&")));
            }
            else if (!IsNaryArgument(item) && ((Is(item, "mrow") && !special) || Is(item, "mstyle")))
            {
                nodes.AddRange(EquationArrayRow(item.Elements()));
            }
            else if (HasAlignmentDescendant(item))
            {
                var count = item.Descendants().Count(descendant => IsAny(descendant, "maligngroup", "malignmark"));
                nodes.Add(Om("r", Om("t", OutputText(new string('&', count), IsAny(item, "mtext", "ms")))));
                nodes.AddRange(Apply(item));
            }
            else
            {
                nodes.AddRange(Apply(item));
            }
        }
        return nodes;
    }

    // ---- Phantoms -------------------------------------------------------

    private static int Full(string value)
    {
        if (value.IndexOfAny("123456789".ToCharArray()) >= 0) return 1;
        if (value.IndexOf('0') >= 0) return 0;
        return 1;
    }

    private static XElement? PhantomProperties(int show, int width, int height, int depth)
    {
        if (show != 0 && width != 0 && height != 0 && depth != 0) return null;
        return Om("phantPr",
            show == 0 ? Om("show", Val("off")) : null,
            width == 0 ? Om("zeroWid", Val("on")) : null,
            height == 0 ? Om("zeroAsc", Val("on")) : null,
            depth == 0 ? Om("zeroDesc", Val("on")) : null);
    }

    private static List<XNode> Padded(XElement padded)
    {
        if (Is(padded.Parent, "mphantom") && !padded.ElementsBeforeSelf().Any() && !padded.ElementsAfterSelf().Any())
            return ApplyChildren(padded);
        return One(Om("phant",
            PhantomProperties(1, Full(Value(padded, "width")), Full(Value(padded, "height")), Full(Value(padded, "depth"))),
            Om("e", ApplyChildren(padded))));
    }

    private static XElement Phantom(XElement phantom)
    {
        var children = phantom.Elements().ToArray();
        XElement? properties;
        if (children.Length == 1 && Is(children[0], "mpadded"))
        {
            var padded = children[0];
            properties = PhantomProperties(0,
                Full(LowerAscii(Value(padded, "width"))),
                Full(LowerAscii(Value(padded, "height"))),
                Full(LowerAscii(Value(padded, "depth"))));
        }
        else properties = PhantomProperties(0, 1, 1, 1);
        return Om("phant", properties, Om("e", ApplyChildren(phantom)));
    }
}

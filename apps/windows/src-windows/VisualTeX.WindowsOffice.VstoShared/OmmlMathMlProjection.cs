using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// Projects Office Math (OMML) to presentation MathML without Office's
/// OMML2MML.XSL. Office updates replace that stylesheet and change its output
/// shape, so the projection is owned here. It reproduces the token, structure
/// and attribute contract that VisualTeX's recovery passes and the LaTeX and
/// MathType consumers were built against: Office's classic stylesheet together
/// with VisualTeX's adaptations (styled operators and numbers, run colors,
/// fixed literal operators, Unicode scalar tokens).
///
/// Every well-formed OMML document projects. Unknown elements contribute the
/// projection of their children, exactly as the stylesheet's catch-all did,
/// so content written by a newer Word is never rejected.
/// </summary>
internal static class OmmlMathMlProjection
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";
    private const string NoBreakSpace = "\u00A0";

    // Every character OMML recognizes as an operator (Office's sOperators).
    private static readonly HashSet<char> Operators = new(
        "\u00A8\u0021\u0022\u0023\u0026\u0028\u0029\u002B\u002C\u002D\u002E\u002F\u003A"
        + "\u003B\u003C\u003D\u003E\u003F\u0040\u005B\u005C\u005D\u005E\u005F\u0060\u007B"
        + "\u007C\u007D\u007E\u00A1\u00A6\u00AC\u00AF\u00B0\u00B1\u00B2\u00B3\u00B4\u00B7\u00B9\u00BF"
        + "\u00D7\u00F7\u02C7\u02D8\u02D9\u02DC\u02DD"
        + Range(0x0300, 0x0338) + "\u033F"
        + "\u2000\u2001\u2002\u2003\u2004\u2005\u2006\u2009\u200A\u2010\u2012\u2013"
        + "\u2014\u2016\u2020\u2021\u2022\u2024\u2025\u2026\u2032\u2033\u2034\u203C"
        + "\u2040\u2044\u204E\u204F\u2050\u2057\u2061\u2062\u2063\u2070\u2074\u2075"
        + "\u2076\u2077\u2078\u2079\u207A\u207B\u207C\u207D\u207E" + Range(0x2080, 0x208E)
        + Range(0x20D0, 0x20E1) + Range(0x20E4, 0x20EA)
        + "\u2140\u2146" + Range(0x2190, 0x21B3) + "\u21B6\u21B7" + Range(0x21BA, 0x21E9)
        + Range(0x21F3, 0x2204) + Range(0x2206, 0x220D) + Range(0x220F, 0x221D)
        + Range(0x2223, 0x223E) + Range(0x2240, 0x22A3) + Range(0x22A5, 0x22BD)
        + Range(0x22C0, 0x22FF)
        + "\u2305\u2306\u2308\u2309\u230A\u230B\u231C\u231D\u231E\u231F\u2322\u2323\u2329\u232A"
        + "\u233D\u233F\u23B0\u23B1\u23DC\u23DD\u23DE\u23DF\u23E0\u2502\u251C\u2524"
        + "\u252C\u2534\u2581\u2588\u2592\u25A0\u25A1\u25AD\u25B2\u25B3\u25B4\u25B5"
        + "\u25B6\u25B7\u25B8\u25B9\u25BC\u25BD\u25BE\u25BF\u25C0\u25C1\u25C2\u25C3"
        + "\u25C4\u25C5\u25CA\u25CB\u25E6\u25EB\u25EC\u25F8\u25F9\u25FA\u25FB\u25FC"
        + "\u25FD\u25FE\u25FF\u2605\u2606\u2772\u2773" + Range(0x27D1, 0x27EB)
        + Range(0x27F0, 0x2980) + Range(0x2982, 0x299A)
        + "\u29B6\u29B7\u29B8\u29B9\u29C0\u29C1\u29C4\u29C5\u29C6\u29C7\u29C8"
        + Range(0x29CE, 0x29DB) + "\u29DF" + Range(0x29E1, 0x29E6) + "\u29EB"
        + Range(0x29F4, 0x2AE0) + Range(0x2AE2, 0x2AF0) + Range(0x2AF2, 0x2AFF)
        + "\u2B04\u2B06\u2B07\u2B0C\u2B0D\u3014\u3015\u3016\u3017\u3018\u3019\uFF01\uFF06"
        + "\uFF08\uFF09\uFF0B\uFF0C\uFF0D\uFF0E\uFF0F\uFF1A\uFF1B\uFF1C\uFF1D\uFF1E"
        + "\uFF1F\uFF20\uFF3B\uFF3C\uFF3D\uFF3E\uFF3F\uFF5B\uFF5C\uFF5D");

    private static string Range(int first, int last)
    {
        var builder = new StringBuilder();
        for (var code = first; code <= last; code++) builder.Append((char)code);
        return builder.ToString();
    }

    // The reader settings of the transform input: insignificant whitespace,
    // comments and DTDs never reach the projection.
    internal static XmlReaderSettings ReaderSettings() => new()
    {
        DtdProcessing = DtdProcessing.Prohibit,
        XmlResolver = null,
        IgnoreComments = true,
        IgnoreWhitespace = true,
        MaxCharactersInDocument = 4_000_000,
    };

    /// <summary>Projects one OMML document. The result is a MathML math element.</summary>
    internal static XElement Project(string omml)
    {
        XDocument source;
        using (var text = new StringReader(omml))
        using (var reader = XmlReader.Create(text, ReaderSettings()))
            source = XDocument.Load(reader, LoadOptions.None);

        var math = new XElement(MathMl + "math");
        if (source.Root is not null) math.Add(Apply(source.Root));
        return Settle(math);
    }

    // The projection was historically serialized and parsed again with the
    // same reader settings. Whitespace-only tokens did not survive that, and
    // line breaks were normalized; keep that observable result.
    private static XElement Settle(XElement math)
    {
        var settings = new XmlWriterSettings { OmitXmlDeclaration = true, Encoding = new UTF8Encoding(false) };
        var output = new StringBuilder();
        using (var writer = XmlWriter.Create(output, settings)) math.WriteTo(writer);
        using var text = new StringReader(output.ToString());
        using var reader = XmlReader.Create(text, ReaderSettings());
        return XDocument.Load(reader, LoadOptions.None).Root!;
    }

    private static List<XNode> Apply(XElement element)
    {
        if (element.Name.Namespace != M) return ApplyChildren(element);
        return element.Name.LocalName switch
        {
            "borderBox" => One(BorderBox(element)),
            "acc" => One(Accent(element)),
            "sPre" => One(Prescripts(element)),
            "m" => One(Matrix(element)),
            "phant" => One(Phantom(element)),
            "rad" => One(Radical(element)),
            "nary" => One(Nary(element)),
            "limLow" => One(El("munder", Row(First(element, "e")), Row(First(element, "lim")))),
            "limUpp" => One(El("mover", Row(First(element, "e")), Row(First(element, "lim")))),
            "sSub" => One(El("msub", Row(First(element, "e")), Row(First(element, "sub")))),
            "sSup" => One(El("msup", Row(First(element, "e")), Row(First(element, "sup")))),
            "sSubSup" => One(El("msubsup",
                Row(First(element, "e")), Row(First(element, "sub")), Row(First(element, "sup")))),
            "groupChr" => One(GroupCharacter(element)),
            "func" => One(Function(element)),
            "f" => One(Fraction(element)),
            "e" or "den" or "num" or "lim" or "sup" or "sub" => Argument(element),
            "bar" => One(Bar(element)),
            "d" => One(Delimiter(element)),
            "r" => Run(element),
            "eqArr" => One(EquationArray(element)),
            _ => ApplyChildren(element),
        };
    }

    private static List<XNode> ApplyChildren(XElement element)
    {
        var nodes = new List<XNode>();
        foreach (var child in element.Elements()) nodes.AddRange(Apply(child));
        return nodes;
    }

    private static List<XNode> ApplyOptional(XElement? element) =>
        element is null ? new List<XNode>() : Apply(element);

    private static List<XNode> One(XNode node) => new() { node };

    private static XElement El(string name, params object?[] content) => new(MathMl + name, content);

    private static XElement Row(XElement? element) => El("mrow", ApplyOptional(element));

    private static XElement? First(XElement element, string name) => element.Element(M + name);

    private static XElement? Last(XElement? element, string name) =>
        element?.Elements(M + name).LastOrDefault();

    // An XPath attribute node-set: the m:val of every named child, in document order.
    private static string[] Values(IEnumerable<XElement> owners, string name) => owners
        .SelectMany(owner => owner.Elements(M + name))
        .Select(child => (string?)child.Attribute(M + "val"))
        .Where(value => value is not null)
        .Select(value => value!)
        .ToArray();

    private static string[] Values(XElement? owner, string name) =>
        owner is null ? Array.Empty<string>() : Values(new[] { owner }, name);

    // XPath 1.0 comparisons of a node-set with a string are existential.
    private static bool Eq(string[] set, string value) => set.Any(item => item == value);
    private static bool Ne(string[] set, string value) => set.Any(item => item != value);
    private static string Str(string[] set) => set.Length > 0 ? set[0] : string.Empty;

    private static string LowerAscii(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
            builder.Append(character is >= 'A' and <= 'Z' ? (char)(character + 32) : character);
        return builder.ToString();
    }

    // Office's TFromStrVal: on/1/true, off/0/false, otherwise neither.
    private static int Tristate(string value) => value switch
    {
        "on" or "1" or "true" => 1,
        "off" or "0" or "false" => 0,
        _ => -1,
    };

    private static int ForceFalse(string lowerValue) => Tristate(lowerValue) == 0 ? 0 : 1;
    private static int ForceTrue(string lowerValue) => Tristate(lowerValue) == 1 ? 1 : 0;

    // An on/off property: absent is off, present is on unless explicitly false.
    private static int Switch(XElement? properties, string name)
    {
        if (properties?.Element(M + name) is null) return 0;
        return ForceFalse(LowerAscii(Str(Values(properties, name))));
    }

    private static int Normal(XElement run) => Switch(Last(run, "rPr"), "nor");

    private static string[] Colors(IEnumerable<XElement?> runProperties) => runProperties
        .Where(properties => properties is not null)
        .SelectMany(properties => properties!.Elements(W + "color"))
        .Select(color => (string?)color.Attribute(W + "val"))
        .Where(value => value is not null)
        .Select(value => value!)
        .ToArray();

    private static string[] RunColors(XElement? run) =>
        run is null ? Array.Empty<string>() : Colors(new[] { run.Elements(W + "rPr").LastOrDefault() });

    // The nearest m:r at or above the context node.
    private static string[] ContextRunColors(XElement context) =>
        RunColors(context.AncestorsAndSelf(M + "r").FirstOrDefault());

    private static string[] ControlColors(XElement owner) => Colors(owner.Elements()
        .Where(element => element.Name == M + "fPr" || element.Name == M + "radPr")
        .SelectMany(properties => properties.Elements(M + "ctrlPr"))
        .Select(control => control.Elements(W + "rPr").LastOrDefault()));

    private static XAttribute? MathColor(string[] colors) =>
        Str(colors).Length > 0 && Ne(colors, "auto")
            ? new XAttribute("mathcolor", "#" + Str(colors))
            : null;

    private static XElement BorderBox(XElement box)
    {
        var properties = Last(box, "borderBoxPr");
        int Flag(string name) =>
            ForceTrue(LowerAscii((string?)Last(properties, name)?.Attribute(M + "val") ?? string.Empty));
        var hideTop = Flag("hideTop");
        var hideBottom = Flag("hideBot");
        var hideLeft = Flag("hideLeft");
        var hideRight = Flag("hideRight");
        var strikeH = Flag("strikeH");
        var strikeV = Flag("strikeV");
        var strikeBltr = Flag("strikeBLTR");
        var strikeTlbr = Flag("strikeTLBR");
        if (hideTop == 1 && hideBottom == 1 && hideLeft == 1 && hideRight == 1
            && strikeH == 0 && strikeV == 0 && strikeBltr == 0 && strikeTlbr == 0)
            return El("mrow", ApplyOptional(First(box, "e")));

        string notation;
        if (hideTop == 0 && hideBottom == 0 && hideLeft == 0 && hideRight == 0)
            notation = "box";
        else
        {
            notation = "left right top bottom";
            if (hideTop == 1) notation = notation.Replace("top", string.Empty);
            if (hideBottom == 1) notation = notation.Replace("bottom", string.Empty);
            if (hideLeft == 1) notation = notation.Replace("left", string.Empty);
            if (hideRight == 1) notation = notation.Replace("right", string.Empty);
        }
        if (strikeH == 1) notation += " horizontalstrike";
        if (strikeV == 1) notation += " verticalstrike";
        if (strikeBltr == 1) notation += " updiagonalstrike";
        if (strikeTlbr == 1) notation += " downdiagonalstrike";
        return El("menclose", new XAttribute("notation", NormalizeSpace(notation)),
            ApplyOptional(First(box, "e")));
    }

    private static string NormalizeSpace(string value) => string.Join(" ",
        value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));

    // Office's ToNonCombining; the first matching row of its table wins.
    private static string ToNonCombining(string character) => character switch
    {
        "\u0306" or "\u032E" => "\u02D8",
        "\u0312" or "\u0327" => "\u00B8",
        "\u0300" or "\u0316" => "`",
        "\u0305" or "\u0332" => "-",
        "\u0323" => ".",
        "\u0307" => "\u02D9",
        "\u030B" or "\u02DD" => "\u02DD",
        "\u0301" or "\u0317" => "\u00B4",
        "\u0303" or "\u0330" => "~",
        "\u0308" or "\u0324" => "\u00A8",
        "\u030C" or "\u032C" => "\u02C7",
        "\u0302" or "\u032D" => "^",
        "\u20D7" or "\u20EF" => "\u2192",
        "\u20D6" or "\u20EE" => "\u2190",
        _ => character,
    };

    private static XElement Accent(XElement accent)
    {
        var mark = Last(accent, "accPr")?.Element(M + "chr") is null
            ? "\u0302"
            : FirstScalar(Str(Values(accent.Elements(M + "accPr"), "chr")));
        var baseSlot = First(accent, "e");
        var result = El("mover", new XAttribute("accent", "true"), Row(baseSlot));
        if (mark.Length == 0)
        {
            result.Add(El("mo"));
            return result;
        }

        // The mark takes the style of the base's runs, as Office's stylesheet did.
        var baseRunProperties = (baseSlot?.Elements() ?? Enumerable.Empty<XElement>())
            .Select(child => Last(child, "rPr"))
            .Where(properties => properties is not null)
            .Select(properties => properties!)
            .ToArray();
        var normal = baseRunProperties.Any(properties => properties.Element(M + "nor") is not null)
            ? ForceFalse(LowerAscii(Str(Values(baseRunProperties, "nor"))))
            : 0;
        result.Add(ParseText(
            ToNonCombining(mark),
            Values(baseRunProperties, "scr"),
            Values(baseRunProperties, "sty"),
            normal,
            accent));
        return result;
    }

    private static XElement Prescripts(XElement script)
    {
        XElement Slot(XElement? slot) =>
            slot is not null && slot.HasElements ? El("mrow", Apply(slot)) : El("none");
        return El("mmultiscripts",
            Row(First(script, "e")),
            El("mprescripts"),
            Slot(First(script, "sub")),
            Slot(First(script, "sup")));
    }

    private static XElement Matrix(XElement matrix)
    {
        var alignment = LowerAscii(Str(Values(
            (Last(matrix, "mPr")?.Elements(M + "mcs") ?? Enumerable.Empty<XElement>())
                .SelectMany(columns => columns.Elements(M + "mc"))
                .Select(column => Last(column, "mcPr"))
                .Where(properties => properties is not null)
                .Select(properties => properties!),
            "mcJc")));
        var table = El("mtable");
        if (alignment is "left" or "right") table.Add(new XAttribute("columnalign", alignment));
        foreach (var row in matrix.Elements(M + "mr"))
            table.Add(El("mtr", row.Elements(M + "e").Select(cell => El("mtd", Apply(cell)))));
        return table;
    }

    private static XElement Phantom(XElement phantom)
    {
        var properties = Last(phantom, "phantPr");
        int Flag(string name)
        {
            var flag = Last(properties, name);
            return flag is null
                ? 0
                : ForceFalse(LowerAscii((string?)flag.Attribute(M + "val") ?? string.Empty));
        }
        var zeroWidth = Flag("zeroWid");
        var zeroAscent = Flag("zeroAsc");
        var zeroDescent = Flag("zeroDesc");
        var show = ForceFalse(LowerAscii((string?)Last(properties, "show")?.Attribute(M + "val") ?? string.Empty));

        XElement Content() => El("mrow", phantom.Elements(M + "e").SelectMany(Apply));
        XElement Padded() => El("mpadded",
            zeroWidth == 1 ? new XAttribute("width", "0in") : null,
            zeroAscent == 1 ? new XAttribute("height", "0in") : null,
            zeroDescent == 1 ? new XAttribute("depth", "0in") : null,
            Content());
        if (show == 1) return Padded();
        if (zeroWidth == 0 && zeroAscent == 0 && zeroDescent == 0) return El("mphantom", Content());
        return El("mphantom", Padded());
    }

    private static XElement Radical(XElement radical)
    {
        var hideDegree = Switch(Last(radical, "radPr"), "degHide");
        var color = MathColor(ControlColors(radical));
        if (hideDegree == 1)
            return El("msqrt", color, ApplyOptional(First(radical, "e")));
        return El("mroot", color, Row(First(radical, "e")), Row(First(radical, "deg")));
    }

    private static XElement Nary(XElement nary)
    {
        var properties = Last(nary, "naryPr");
        string Lower(string name, string absent) =>
            properties?.Element(M + name) is null ? absent : LowerAscii(Str(Values(properties, name)));
        var hideSub = ForceFalse(Lower("subHide", "off"));
        var hideSup = ForceFalse(Lower("supHide", "off"));
        var grow = ForceFalse(Lower("grow", "off"));
        var limitsAsScripts = properties?.Element(M + "limLoc") is null
            || LowerAscii(Str(Values(properties, "limLoc"))) == "subsup";

        var characters = Values(properties, "chr");
        var symbol = characters.Length == 0 || Eq(characters, string.Empty) ? "\u222B" : Str(characters);
        XElement Operator() => El("mo", new XAttribute("stretchy", grow == 1 ? "true" : "false"), symbol);
        XElement Sub() => Row(First(nary, "sub"));
        XElement Sup() => Row(First(nary, "sup"));

        XElement head;
        if (hideSup == 1 && hideSub == 1) head = Operator();
        else if (hideSub == 1) head = El(limitsAsScripts ? "msup" : "mover", Operator(), Sup());
        else if (hideSup == 1) head = El(limitsAsScripts ? "msub" : "munder", Operator(), Sub());
        else head = El(limitsAsScripts ? "msubsup" : "munderover", Operator(), Sub(), Sup());
        return El("mrow", head, Row(First(nary, "e")));
    }

    private static XElement GroupCharacter(XElement group)
    {
        var properties = Last(group, "groupChrPr");
        var hasCharacter = properties?.Element(M + "chr") is not null;
        var characterValue = Str(Values(properties, "chr"));
        var character = hasCharacter && characterValue.Length == 0
            ? string.Empty
            : characterValue.Length > 0 ? FirstScalar(characterValue) : "\u23DF";
        var position = LowerAscii(Str(Values(properties, "pos")));
        var justification = LowerAscii(Str(Values(properties, "vertJc")));
        XElement Mark() => El("mo", character.Length > 0 ? character : null);
        var content = Row(First(group, "e"));
        if (position == "top")
            return justification == "bot"
                ? El("mover", new XAttribute("accent", "false"), content, Mark())
                : El("munder", new XAttribute("accentunder", "false"), Mark(), content);
        return justification == "bot"
            ? El("mover", new XAttribute("accent", "false"), Mark(), content)
            : El("munder", new XAttribute("accentunder", "false"), content, Mark());
    }

    private static XElement Function(XElement function) => El("mrow",
        El("mrow", function.Elements(M + "fName").SelectMany(name => name.Elements()).SelectMany(Apply)),
        El("mo", "\u2061"),
        El("mrow", function.Elements(M + "e").SelectMany(Apply)));

    private static XElement Fraction(XElement fraction)
    {
        var type = LowerAscii(Str(Values(Last(fraction, "fPr"), "type")));
        if (type == "lin")
            return El("mrow", Row(First(fraction, "num")), El("mo", "/"), Row(First(fraction, "den")));
        return El("mfrac",
            MathColor(ControlColors(fraction)),
            type == "skw" ? new XAttribute("bevelled", "true") : null,
            type == "nobar" ? new XAttribute("linethickness", "0pt") : null,
            Row(First(fraction, "num")),
            Row(First(fraction, "den")));
    }

    private static List<XNode> Argument(XElement argument)
    {
        var level = Values(Last(argument, "argPr"), "scrLvl");
        if (level.Length == 0) return ApplyChildren(argument);
        return One(El("mstyle", new XAttribute("scriptlevel", Str(level)), ApplyChildren(argument)));
    }

    private static XElement Bar(XElement bar)
    {
        var position = LowerAscii(Str(Values(bar.Elements(M + "barPr"), "pos")));
        return position == "top"
            ? El("mover", new XAttribute("accent", "false"), Row(First(bar, "e")), El("mo", "\u00AF"))
            : El("munder", new XAttribute("underaccent", "false"), Row(First(bar, "e")), El("mo", "_"));
    }

    private static XElement Delimiter(XElement delimiter)
    {
        var properties = First(delimiter, "dPr");
        var fence = El("mfenced");
        var open = Values(properties, "begChr");
        if (open.Length > 0 && !Eq(open, "(")) fence.Add(new XAttribute("open", Str(open)));
        var close = Values(properties, "endChr");
        if (close.Length > 0 && !Eq(close, ")")) fence.Add(new XAttribute("close", Str(close)));
        var separator = Values(properties, "sepChr");
        if (!Eq(separator, ","))
            fence.Add(new XAttribute("separators", separator.Length == 0 ? "|" : Str(separator)));
        foreach (var item in delimiter.Elements(M + "e")) fence.Add(El("mrow", Apply(item)));
        return fence;
    }

    private static List<XNode> Run(XElement run)
    {
        if (Normal(run) == 1)
        {
            var text = run.Descendants(M + "t").FirstOrDefault()?.Value ?? string.Empty;
            return One(El("mtext", MathColor(RunColors(run)), Text(text.Replace(" ", NoBreakSpace))));
        }
        var nodes = new List<XNode>();
        foreach (var text in run.Descendants(M + "t"))
        {
            var properties = Last(text.Parent, "rPr");
            nodes.AddRange(ParseText(text.Value, Values(properties, "scr"), Values(properties, "sty"), 0, text));
        }
        return nodes;
    }

    private static XText? Text(string value) => value.Length == 0 ? null : new XText(value);

    private static XElement EquationArray(XElement array)
    {
        var table = El("mtable");
        foreach (var equation in array.Elements(M + "e"))
        {
            var level = Values(Last(equation, "argPr"), "scrLvl");
            var row = new List<XNode> { El("maligngroup") };
            row.AddRange(EquationArrayRow(array, equation));
            var cell = Ne(level, "0") || level.Length == 0 || Eq(level, string.Empty)
                ? El("mrow", row)
                : El("mstyle", new XAttribute("scriptlevel", Str(level)), row);
            table.Add(El("mtr", El("mtd", cell)));
        }
        return table;
    }

    // Office's CreateEqArrRow: '&' alternates between an alignment mark and an
    // alignment group, counted across the runs of one equation. Office reads
    // the run style relative to the equation, so it is the array's own m:rPr.
    private static List<XNode> EquationArrayRow(XElement array, XElement equation)
    {
        var nodes = new List<XNode>();
        var align = 1;
        var arrayProperties = Last(array, "rPr");
        var script = Values(arrayProperties, "scr");
        var style = Values(arrayProperties, "sty");
        foreach (var child in equation.Elements())
        {
            var allText = string.Concat(child.Elements(M + "t").Select(text => text.Value));
            if (child.Name == M + "r")
                nodes.AddRange(ParseEquationArrayText(allText, script, style, Normal(child), align, RunColors(child)));
            else
                nodes.AddRange(Apply(child));
            var ampersands = allText.Count(character => character == '&');
            align = (align + ampersands % 2) % 2;
        }
        return nodes;
    }

    private static List<XNode> ParseEquationArrayText(
        string source, string[] script, string[] style, int normal, int align, string[] colors)
    {
        var nodes = new List<XNode>();
        var scalars = Scalars(source);
        var pattern = TokenPattern(scalars);
        var none = Array.Empty<string>();
        for (var index = 0; index < scalars.Count;)
        {
            var scalar = scalars[index];
            if (scalar == "&")
            {
                nodes.Add(El(align == 0 ? "maligngroup" : "malignmark"));
                align = align == 1 ? 0 : 1;
                index++;
                continue;
            }
            var color = MathColor(colors);
            if (pattern[index] == '0')
            {
                var length = NumberLength(pattern, index);
                var number = string.Concat(scalars.Skip(index).Take(length));
                nodes.Add(normal == 0
                    ? El("mn", color, TokenAttributes(script, style, normal, null, "mn", none), number)
                    : El("mtext", color, number));
                index += length;
            }
            else if (pattern[index] == '-')
            {
                nodes.Add(normal == 0
                    ? El("mo", color, new XAttribute("stretchy", "false"),
                        TokenAttributes(script, style, normal, null, "mo", none), scalar)
                    : El("mtext", color, scalar));
                index++;
            }
            else
            {
                var text = scalar.Replace(" ", NoBreakSpace);
                nodes.Add(normal == 0
                    ? El("mi", color, TokenAttributes(script, style, normal, 1, "mi", none), text)
                    : El("mtext", color, text));
                index++;
            }
        }
        return nodes;
    }

    // Office's ParseMt over Unicode scalars: a run splits into identifiers,
    // numbers and single-character operators. Inside a function name the
    // identifier extends to the first number or operator.
    private static List<XNode> ParseText(string source, string[] script, string[] style, int normal, XElement context)
    {
        var nodes = new List<XNode>();
        var scalars = Scalars(source);
        var pattern = TokenPattern(scalars);
        var inFunctionName = context.Ancestors(M + "fName").Any();
        var literalRunOperator = context.Name == M + "t" && context.Parent?.Name == M + "r";
        var colors = ContextRunColors(context);
        for (var index = 0; index < scalars.Count;)
        {
            if (pattern[index] == '-')
            {
                nodes.Add(El("mo",
                    literalRunOperator ? new XAttribute("stretchy", "false") : null,
                    TokenAttributes(script, style, normal, null, "mo", colors),
                    scalars[index]));
                index++;
            }
            else if (pattern[index] == '0')
            {
                var length = NumberLength(pattern, index);
                nodes.Add(El("mn",
                    TokenAttributes(script, style, normal, null, "mn", colors),
                    string.Concat(scalars.Skip(index).Take(length))));
                index += length;
            }
            else
            {
                var length = inFunctionName ? FunctionNameLength(pattern, index) : 1;
                nodes.Add(El("mi",
                    TokenAttributes(script, style, normal, length, "mi", colors),
                    string.Concat(scalars.Skip(index).Take(length)).Replace(" ", NoBreakSpace)));
                index += length;
            }
        }
        return nodes;
    }

    private static int NumberLength(char[] pattern, int start)
    {
        var end = start;
        while (end < pattern.Length && pattern[end] == '0') end++;
        return end - start;
    }

    // Office's nCharToPrint inside m:fName, for the text that remains.
    private static int FunctionNameLength(char[] pattern, int start)
    {
        var length = pattern.Length - start;
        var firstOperator = length;
        var firstNumber = length;
        for (var index = start; index < pattern.Length; index++)
        {
            if (pattern[index] == '-') { firstOperator = index - start + 1; break; }
        }
        for (var index = start; index < pattern.Length; index++)
        {
            if (pattern[index] == '0') { firstNumber = index - start + 1; break; }
        }
        var last = pattern[pattern.Length - 1];
        if (firstOperator == firstNumber && firstOperator == length && last != '0' && last != '-')
            return length;
        return firstOperator < firstNumber ? firstOperator - 1 : firstNumber - 1;
    }

    private static List<string> Scalars(string value)
    {
        var scalars = new List<string>(value.Length);
        for (var index = 0; index < value.Length;)
        {
            var length = char.IsHighSurrogate(value[index])
                && index + 1 < value.Length
                && char.IsLowSurrogate(value[index + 1]) ? 2 : 1;
            scalars.Add(value.Substring(index, length));
            index += length;
        }
        return scalars;
    }

    private static string FirstScalar(string value) => value.Length == 0 ? value : Scalars(value)[0];

    // One pattern character per scalar: '0' for number text, '-' for an
    // operator, otherwise the scalar's first code unit. A '.' or ',' between
    // two digits belongs to the number (Office's SReplaceNumWithZero, which
    // replaces non-overlapping "0.0", then "0,0", from the left).
    private static char[] TokenPattern(List<string> scalars)
    {
        var pattern = scalars
            .Select(scalar => scalar.Length == 1 && scalar[0] is >= '0' and <= '9' ? '0' : scalar[0])
            .ToArray();
        ReplaceSeparators(pattern, '.');
        ReplaceSeparators(pattern, ',');
        for (var index = 0; index < pattern.Length; index++)
        {
            if (scalars[index].Length == 1 && pattern[index] != '0' && Operators.Contains(scalars[index][0]))
                pattern[index] = '-';
        }
        return pattern;
    }

    private static void ReplaceSeparators(char[] pattern, char separator)
    {
        for (var index = 0; index + 2 < pattern.Length;)
        {
            if (pattern[index] == '0' && pattern[index + 1] == separator && pattern[index + 2] == '0')
            {
                pattern[index + 1] = '0';
                index += 3;
            }
            else index++;
        }
    }

    // Office's CreateTokenAttributes with VisualTeX's adaptations: styled
    // operators and numbers keep their variant, upright Roman digits stay
    // unmarked, and the run color follows.
    private static IEnumerable<XAttribute> TokenAttributes(
        string[] script, string[] style, int normal, int? characters, string tokenType, string[] colors)
    {
        var attributes = new List<XAttribute>();
        if (normal == 1)
        {
            attributes.Add(new XAttribute("mathvariant", "normal"));
        }
        else
        {
            var variant = Variant(script, style, tokenType);
            var identifier = tokenType == "mi";
            if (tokenType is "mo" or "mn")
            {
                if (variant.Length > 0) attributes.Add(new XAttribute("mathvariant", variant));
            }
            else if (identifier && characters == 1 && (variant.Length == 0 || variant == "italic"))
            {
                // A single-character identifier is italic by default.
            }
            else if (identifier && characters > 1 && (variant.Length == 0 || variant == "italic"))
            {
                attributes.Add(new XAttribute("mathvariant", "italic"));
            }
            else if (variant != "italic" && variant.Length > 0)
            {
                attributes.Add(new XAttribute("mathvariant", variant));
            }
            else
            {
                if (!(identifier && characters == 1) && !(Eq(style, "p") || Eq(style, "b")))
                    attributes.Add(new XAttribute("fontstyle", "italic"));
                if (Eq(style, "b") || Eq(style, "bi"))
                    attributes.Add(new XAttribute("fontweight", "bold"));
            }
        }
        var color = MathColor(colors);
        if (color is not null) attributes.Add(color);
        return attributes;
    }

    private static string Variant(string[] script, string[] style, string tokenType)
    {
        bool Script(string value) => Eq(script, value);
        bool Style(string value) => Eq(style, value);
        if (tokenType == "mn" && (Str(script).Length == 0 || Script("roman")) && (Style("p") || Style(string.Empty)))
            return string.Empty;
        if (Script("monospace")) return "monospace";
        if (Script("sans-serif") && Style("i")) return "sans-serif-italic";
        if (Script("sans-serif") && Style("b")) return "bold-sans-serif";
        if (Script("sans-serif") && Style("bi")) return "sans-serif-bold-italic";
        if (Script("sans-serif")) return "sans-serif";
        if (Script("fraktur") && (Style("b") || Style("bi"))) return "bold-fraktur";
        if (Script("fraktur")) return "fraktur";
        if (Script("double-struck")) return "double-struck";
        if (Script("script") && (Style("b") || Style("bi"))) return "bold-script";
        if (Script("script")) return "script";
        var roman = Script("roman") || script.Length == 0 || Script(string.Empty);
        if (roman && Style("b")) return "bold";
        if (roman && Style("i")) return "italic";
        if (roman && Style("p")) return "normal";
        if (roman && Style("bi")) return "bold-italic";
        return string.Empty;
    }
}

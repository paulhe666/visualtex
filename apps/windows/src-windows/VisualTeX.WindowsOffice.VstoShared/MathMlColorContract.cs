using System.Drawing;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>Opaque text colors shared by source recovery and the Office XML boundary.</summary>
internal static class MathMlColorContract
{
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace Math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Xsl = "http://www.w3.org/1999/XSL/Transform";

    internal static string? Normalize(string? value)
    {
        var color = value?.Trim();
        if (string.IsNullOrEmpty(color)) return null;
        if (Regex.IsMatch(color!, "^#[0-9a-fA-F]{6}$")) return color!.ToUpperInvariant();
        if (Regex.IsMatch(color!, "^#[0-9a-fA-F]{3}$"))
            return "#" + string.Concat(color!.Substring(1).Select(c => new string(c, 2))).ToUpperInvariant();
        var named = Color.FromName(color!);
        return named.IsKnownColor && !named.IsSystemColor && named.A == 255
            ? "#" + named.R.ToString("X2") + named.G.ToString("X2") + named.B.ToString("X2") : null;
    }

    internal static string? Explicit(XElement element) => Normalize(
        (string?)element.Attribute("mathcolor") ?? (string?)element.Attribute(MathMl + "mathcolor")
        ?? (string?)element.Attribute("color") ?? (string?)element.Attribute(MathMl + "color"));

    internal static string? Effective(XElement element) =>
        element.AncestorsAndSelf().Select(Explicit).FirstOrDefault(color => color is not null);

    internal static string Wrap(string source, string? color) =>
        color is null || string.IsNullOrWhiteSpace(source) ? source : @"\textcolor{" + color + "}{" + source + "}";

    internal static string MaterializeInheritedColors(string mathMl)
    {
        var document = XDocument.Parse(mathMl, LoadOptions.PreserveWhitespace);
        foreach (var element in document.Root!.DescendantsAndSelf())
        {
            var color = Effective(element);
            if (color is not null) element.SetAttributeValue("mathcolor", color);
        }
        return document.ToString(SaveOptions.DisableFormatting);
    }

    private static XElement WordColor(string select) =>
        new(Word + "color", new XElement(Xsl + "attribute", new XAttribute("name", "w:val"),
            new XElement(Xsl + "value-of", new XAttribute("select", "substring(" + select + ",2)"))));

    private static XElement ColoredRunProperties(string select) =>
        new(Xsl + "if", new XAttribute("test", "string(" + select + ")!=''"),
            new XElement(Word + "rPr", WordColor(select)));

    internal static void AdaptForwardTransform(XDocument stylesheet)
    {
        stylesheet.Root!.SetAttributeValue(XNamespace.Xmlns + "w", Word.NamespaceName);
        // Stock Office computes color parameters but never emits w:rPr.
        stylesheet.Descendants(Xsl + "template")
            .Single(e => (string?)e.Attribute("name") == "CreateRunProp")
            .Add(ColoredRunProperties("$ndCur/@mathcolor"));

        // Color is a run boundary even when mathematical font and role match.
        var grouping = stylesheet.Descendants(Xsl + "template")
            .Single(e => (string?)e.Attribute("name") == "CreateRunWithSameProp")
            .Elements(Xsl + "variable").Single(e => (string?)e.Attribute("name") == "nndBeforeLim");
        var select = (string)grouping.Attribute("select")!;
        const string sequence = "$ndTokenFirst/following-sibling::*";
        var start = select.IndexOf(sequence, StringComparison.Ordinal);
        var predicate = select.IndexOf('[', start + sequence.Length);
        grouping.SetAttributeValue("select", select.Insert(predicate + 1,
            "string(@mathcolor) != string($ndTokenFirst/@mathcolor) or "));

        // Fraction rules and radical signs own native control formatting.
        // The current stylesheet node is their source MathML structure.
        foreach (var properties in stylesheet.Descendants(Xsl + "element")
                     .Where(e => (string?)e.Attribute("name") is "m:fPr" or "m:radPr").ToArray())
            properties.Add(new XElement(Xsl + "if", new XAttribute("test", "string(@mathcolor)!=''"),
                new XElement(Math + "ctrlPr", new XElement(Word + "rPr", WordColor("@mathcolor")))));

        // The product's accent template already owns b/i. Add color to that
        // same control run, from the accent token rather than its base.
        foreach (var properties in stylesheet.Descendants(Math + "accPr"))
        {
            var run = properties.Element(Math + "ctrlPr")?.Element(Word + "rPr");
            if (run is not null)
                run.Add(new XElement(Xsl + "if", new XAttribute("test", "string(*[2]/@mathcolor)!=''"),
                    WordColor("*[2]/@mathcolor")));
        }
    }

    private static XElement MathColor(string select) =>
        new(Xsl + "if", new XAttribute("test", "string(" + select + ")!='' and " + select + "!='auto'"),
            new XElement(Xsl + "attribute", new XAttribute("name", "mathcolor"),
                new XElement(Xsl + "value-of", new XAttribute("select", "concat('#'," + select + ")"))));

    internal static void AdaptReverseTransform(XDocument stylesheet)
    {
        stylesheet.Root!.SetAttributeValue(XNamespace.Xmlns + "w", Word.NamespaceName);
        const string runColor = "ancestor-or-self::m:r[1]/w:rPr[last()]/w:color/@w:val";
        // ParseMt runs with m:t provenance, so token attributes can read the
        // exact native run, including tokens inside a function or script.
        stylesheet.Descendants(Xsl + "template")
            .Single(e => (string?)e.Attribute("name") == "CreateTokenAttributes")
            .Add(MathColor(runColor));
        var runTemplate = stylesheet.Descendants(Xsl + "template")
            .Single(e => (string?)e.Attribute("match") == "m:r");
        foreach (var text in runTemplate.Descendants(MathMl + "mtext"))
            text.AddFirst(MathColor("w:rPr[last()]/w:color/@w:val"));

        // Equation arrays use a separate tokenizer with a passed run node.
        var parser = stylesheet.Descendants(Xsl + "template")
            .Single(e => (string?)e.Attribute("name") == "ParseEqArrMr");
        parser.AddFirst(new XElement(Xsl + "param", new XAttribute("name", "runColor")));
        foreach (var token in parser.Descendants().Where(e => e.Name.Namespace == MathMl
                     && e.Name.LocalName is "mi" or "mn" or "mo" or "mtext"))
            token.AddFirst(MathColor("$runColor"));
        foreach (var call in stylesheet.Descendants(Xsl + "call-template")
                     .Where(e => (string?)e.Attribute("name") == "ParseEqArrMr"))
            call.Add(new XElement(Xsl + "with-param", new XAttribute("name", "runColor"),
                new XAttribute("select", call.Ancestors(Xsl + "template").First() == parser
                    ? "$runColor" : "$ndCur/w:rPr[last()]/w:color/@w:val")));

        foreach (var element in stylesheet.Descendants().Where(e => e.Name.Namespace == MathMl
                     && e.Name.LocalName is "mfrac" or "msqrt" or "mroot"))
            element.AddFirst(MathColor("(m:fPr|m:radPr)/m:ctrlPr/w:rPr[last()]/w:color/@w:val"));
    }
}

using System.Xml.Linq;
using System.Xml.Xsl;

namespace VisualTeX.WindowsOffice.Tests;

// Office's classic MML2OMML.XSL with the adaptations VisualTeX applied while
// it still projected MathML through that stylesheet. Tests use it only as the
// reference the owned projection must reproduce.
internal static class ClassicOfficeForwardStylesheet
{
    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string OmmlSpaceTextAttribute = "data-visualtex-omml-space";
    private const string OmmlRunRoleAttribute = "data-visualtex-omml-run-role";
    private static readonly XNamespace Xsl = "http://www.w3.org/1999/XSL/Transform";
    private static readonly XNamespace Math = MathNamespace;
    private static readonly XNamespace Word = WordNamespace;

    internal static string ResolvePath() => ClassicOfficeStylesheetFiles.ResolveTransformPath("MML2OMML.XSL");

    internal static XslCompiledTransform Load(string path, bool tolerant = false)
    {
        // Adapt the installed transform in memory. The Office installation is
        // unchanged. Its legacy grouping treats an implicit mo and an explicit
        // italic token as equivalent, then omits their italic run properties.
        // Require the materialized source style/role before that lossy merge.
        XNamespace xsl = "http://www.w3.org/1999/XSL/Transform";
        XNamespace math = MathNamespace;
        XElement Template(XDocument stylesheet, string name) => stylesheet.Descendants(xsl + "template")
            .Single(element => (string?)element.Attribute("name") == name);
        var source = XDocument.Load(path);

        // Native accent marks are control characters, independent of the base.
        // The normalized mark already carries the shared effective token style.
        // Emit real ctrlPr formatting rather than inheriting the first base run.
        XNamespace word = WordNamespace;
        XElement ControlFlag(string name, string condition) => new XElement(word + name,
            new XElement(xsl + "attribute", new XAttribute("name", "w:val"),
                new XElement(xsl + "choose",
                    new XElement(xsl + "when", new XAttribute("test", condition), "1"),
                    new XElement(xsl + "otherwise", "0"))));
        source.Root!.SetAttributeValue(XNamespace.Xmlns + "w", word.NamespaceName);
        source.Root.Add(new XElement(xsl + "template",
            new XAttribute("match", "mml:mover[@accent='true' and count(*)=2 and string-length(*[2])=1 and *[2][self::mml:mo and (@mathvariant='normal' or @mathvariant='bold' or @mathvariant='italic' or @mathvariant='bold-italic')]]"),
            new XAttribute("priority", "10"),
            new XElement(math + "acc",
                new XElement(math + "accPr",
                    new XElement(math + "chr", new XElement(xsl + "attribute",
                        new XAttribute("name", "m:val"), new XElement(xsl + "value-of", new XAttribute("select", "*[2]")))),
                    new XElement(math + "ctrlPr", new XElement(word + "rPr",
                        ControlFlag("b", "*[2]/@mathvariant='bold' or *[2]/@mathvariant='bold-italic'"),
                        ControlFlag("i", "*[2]/@mathvariant='italic' or *[2]/@mathvariant='bold-italic'")))),
                new XElement(math + "e", new XElement(xsl + "apply-templates", new XAttribute("select", "*[1]"))))));

        // Stock XSL has no mspace rule: its catch-all discards empty elements.
        // Emit only exact native text selected by the shared source projection.
        source.Root!.Add(new XElement(xsl + "template",
            new XAttribute("match", "mml:mspace[@" + OmmlSpaceTextAttribute + "]"),
            new XElement(xsl + "if",
                new XAttribute("test", "string-length(@" + OmmlSpaceTextAttribute + ") > 0"),
                new XElement(math + "r",
                    new XElement(math + "t",
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        new XElement(xsl + "value-of",
                            new XAttribute("select", "@" + OmmlSpaceTextAttribute)))))));

        var stylesheet = ClassicOfficeStylesheetFiles.AdaptStylesheet(source, tolerant,
            ("OutputText literal text", AdaptLiteralTextOutput),
            ("CreateRunWithSameProp grouping", stylesheet =>
            {
                var grouping = Template(stylesheet, "CreateRunWithSameProp")
                    .Elements(xsl + "variable")
                    .Single(element => (string?)element.Attribute("name") == "nndBeforeLim");
                var selection = (string?)grouping.Attribute("select")
                    ?? throw new InvalidDataException("Office run grouping has no selection expression.");
                const string sequence = "$ndTokenFirst/following-sibling::*";
                var sequenceIndex = selection.IndexOf(sequence, StringComparison.Ordinal);
                var predicateIndex = sequenceIndex < 0
                    ? -1
                    : selection.IndexOf('[', sequenceIndex + sequence.Length);
                if (predicateIndex < 0)
                    throw new InvalidDataException("Office run grouping has an unsupported selection expression.");
                var boundaries =
                    "string(@mathvariant) != string($ndTokenFirst/@mathvariant) or "
                    + "string(@" + OmmlRunRoleAttribute + ") != string($ndTokenFirst/@"
                    + OmmlRunRoleAttribute + ") or "
                    // ApplyFunction starts a structural boundary even though it shares
                    // the normal operator style. Do not merge it into a function name;
                    // it may still share a run with its following opening delimiter.
                    + "(self::mml:mo and normalize-space(.)='⁡') or ";
                grouping.SetAttributeValue("select", selection.Insert(predicateIndex + 1, boundaries));
            }),
            ("CreateMathRPR", stylesheet =>
                Template(stylesheet, "CreateMathRPR").Elements(xsl + "if").Single()
                    .SetAttributeValue("test", "$fNor=1 or $sFontCur!=''")),
            ("CreateMathScrStyProp italic", stylesheet =>
            {
                var italic = Template(stylesheet, "CreateMathScrStyProp")
                    .Elements(xsl + "choose")
                    .Elements(xsl + "when")
                    .Single(element => (string?)element.Attribute("test") == "$font='italic'");
                italic.ReplaceNodes(new XElement(math + "sty", new XAttribute(math + "val", "i")));
                // Monospace is a mathematical alphabet, represented by native scr
                // rather than a Word text font or Normal Text. The source projection
                // has already resolved inherited variants and kept token roles apart.
                italic.AddBeforeSelf(new XElement(xsl + "when",
                    new XAttribute("test", "$font='monospace'"),
                    new XElement(math + "scr", new XAttribute(math + "val", "monospace")),
                    new XElement(math + "sty", new XAttribute(math + "val", "p"))));
            }),
            ("text color", AdaptForwardColor));
        var transform = new XslCompiledTransform(enableDebug: false);
        using var reader = stylesheet.CreateReader();
        transform.Load(reader, new XsltSettings(enableDocumentFunction: false, enableScript: false), null);
        return transform;
    }

    private static void AdaptLiteralTextOutput(XDocument stylesheet)
    {
        XNamespace xsl = "http://www.w3.org/1999/XSL/Transform";
        const string literal = "self::mml:mtext or self::mml:ms";
        var output = stylesheet.Descendants(xsl + "template")
            .Single(element => (string?)element.Attribute("name") == "OutputText");
        var body = output.Elements().Where(element => element.Name != xsl + "param").ToArray();
        foreach (var node in body) node.Remove();
        output.Add(new XElement(xsl + "param", new XAttribute("name", "fLiteral"), new XAttribute("select", literal)),
            new XElement(xsl + "choose",
                new XElement(xsl + "when", new XAttribute("test", "$fLiteral"),
                    new XElement(xsl + "value-of", new XAttribute("select", "$sInput"))),
                new XElement(xsl + "otherwise", body)));

        // Literal text is not a stream of mathematical operators. Preserve its
        // characters and whitespace instead of applying Office's operator-only
        // OutputText rewrites (invisible multiplication, spaces, ligatures).
        // Grouped runs have an explicit first token; never infer their role from
        // the caller's current node. Existing grouping already separates mtext.
        foreach (var call in stylesheet.Descendants(xsl + "call-template")
                     .Where(element => (string?)element.Attribute("name") == "OutputText").ToArray())
        {
            var template = call.Ancestors(xsl + "template").First();
            var grouped = (string?)template.Attribute("name") == "CreateRunWithSameProp";
            call.Add(new XElement(xsl + "with-param", new XAttribute("name", "fLiteral"),
                new XAttribute("select", grouped ? "$ndTokenFirst/self::mml:mtext or $ndTokenFirst/self::mml:ms" : literal)));
            foreach (var value in call.Descendants(xsl + "value-of").ToArray())
            {
                var expression = (string?)value.Attribute("select");
                var token = expression == "normalize-space($ndTokenFirst)" ? "$ndTokenFirst"
                    : expression == "normalize-space(.)" ? "." : null;
                if (token is null) continue;
                value.ReplaceWith(new XElement(xsl + "choose",
                    new XElement(xsl + "when", new XAttribute("test", token == "." ? literal
                        : "$ndTokenFirst/self::mml:mtext or $ndTokenFirst/self::mml:ms"),
                        new XElement(xsl + "value-of", new XAttribute("select", token))),
                    new XElement(xsl + "otherwise", new XElement(value))));
            }
        }
    }

    private static XElement WordColor(string select) =>
        new(Word + "color", new XElement(Xsl + "attribute", new XAttribute("name", "w:val"),
            new XElement(Xsl + "value-of", new XAttribute("select", "substring(" + select + ",2)"))));

    private static XElement ColoredRunProperties(string select) =>
        new(Xsl + "if", new XAttribute("test", "string(" + select + ")!=''"),
            new XElement(Word + "rPr", WordColor(select)));

    private static void AdaptForwardColor(XDocument stylesheet)
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
}

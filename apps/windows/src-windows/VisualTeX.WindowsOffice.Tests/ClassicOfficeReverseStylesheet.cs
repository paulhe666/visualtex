using System.Text;
using System.Xml.Linq;
using System.Xml.Xsl;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

// Office's classic OMML2MML.XSL with the adaptations VisualTeX applied while
// it still projected OMML through that stylesheet. Tests use it only as the
// reference the owned projection must reproduce; product code never reads it.
internal static class ClassicOfficeReverseStylesheet
{
    private static readonly XNamespace Xsl = "http://www.w3.org/1999/XSL/Transform";
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";
    private static readonly XNamespace Word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static string ResolvePath() => ClassicOfficeStylesheetFiles.ResolveTransformPath("OMML2MML.XSL");

    internal const string UnicodeScalarXPathNamespace = "urn:visualtex:unicode-scalar-xpath";

    internal sealed class UnicodeScalarXPathFunctions
    {
        public double StringLength(string value)
        {
            var count = 0;
            for (var index = 0; index < value.Length; count++)
                index += char.IsSurrogatePair(value, index) ? 2 : 1;
            return count;
        }

        public string Substring(string value, double start)
            => Slice(value, RoundXPath(start), double.PositiveInfinity);

        public string Substring(string value, double start, double length)
        {
            var first = RoundXPath(start);
            return Slice(value, first, first + RoundXPath(length));
        }

        private static double RoundXPath(double value)
        {
            var rounded = Math.Round(value);
            return value - rounded == 0.5 ? rounded + 1 : rounded;
        }

        private static string Slice(string value, double first, double end)
        {
            var result = new StringBuilder();
            var position = 1;
            for (var index = 0; index < value.Length; position++)
            {
                var count = char.IsSurrogatePair(value, index) ? 2 : 1;
                if (position >= first && position < end) result.Append(value, index, count);
                index += count;
            }
            return result.ToString();
        }
    }

    internal static string RewriteUnicodeScalarXPath(string expression)
    {
        var result = new StringBuilder();
        var quote = '\0';
        for (var index = 0; index < expression.Length;)
        {
            var character = expression[index];
            if (quote != '\0')
            {
                result.Append(character); index++;
                if (character == quote) quote = '\0';
            }
            else if (character is '\'' or '"')
            {
                quote = character; result.Append(character); index++;
            }
            else if (char.IsLetter(character) || character == '_')
            {
                var start = index++;
                while (index < expression.Length &&
                    (char.IsLetterOrDigit(expression[index]) || "_-.:".IndexOf(expression[index]) >= 0)) index++;
                var name = expression.Substring(start, index - start);
                var next = index;
                while (next < expression.Length && char.IsWhiteSpace(expression[next])) next++;
                if (next < expression.Length && expression[next] == '(' &&
                    (name == "substring" || name == "string-length"))
                    result.Append("vtUnicode:").Append(name == "substring" ? "Substring" : "StringLength");
                else result.Append(name);
            }
            else { result.Append(character); index++; }
        }
        return result.ToString();
    }

    internal static XslCompiledTransform Load(string path, bool tolerant = false)
    {
        // Office's legacy MathML projection deliberately discards styled
        // operators and numbers even when OMML carries their exact m:sty.
        // Correct that projection before compilation, retaining the original
        // token/structure transform and leaving the Office installation intact.
        XNamespace xsl = "http://www.w3.org/1999/XSL/Transform";
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        XElement Template(XDocument stylesheet, string name) => stylesheet.Descendants(xsl + "template")
            .Single(element => (string?)element.Attribute("name") == name);
        var source = XDocument.Load(path);
        source.Root!.SetAttributeValue(XNamespace.Xmlns + "vtUnicode", UnicodeScalarXPathNamespace);
        var stylesheet = ClassicOfficeStylesheetFiles.AdaptStylesheet(source, tolerant,
            ("ParseMt Unicode scalars", stylesheet =>
            {
                // The legacy token parser counts UTF-16 units on .NET. Its traversal
                // must consume complete Unicode scalars, never half a surrogate pair.
                foreach (var expression in Template(stylesheet, "ParseMt").DescendantsAndSelf().Attributes()
                             .Where(attribute => attribute.Name.LocalName is "select" or "test"))
                    expression.Value = RewriteUnicodeScalarXPath(expression.Value);
            }),
            // A literal operator in an OMML run has fixed glyph geometry. Preserve
            // that at the tokenizer, where native provenance is still available,
            // rather than letting MathML's operator dictionary invent stretching.
            // ParseMt also creates accent marks: only its m:r/m:t context is a run.
            // ParseEqArrMr is called exclusively for CreateEqArrRow's m:r branch;
            // its current node is the containing equation row, not the run itself.
            ("ParseMt operator", stylesheet =>
                Template(stylesheet, "ParseMt").Descendants(presentationMath + "mo").Single().AddFirst(
                    new XElement(xsl + "if", new XAttribute("test", "self::m:t[parent::m:r]"),
                        new XElement(xsl + "attribute", new XAttribute("name", "stretchy"), "false")))),
            ("ParseEqArrMr operator", stylesheet =>
                Template(stylesheet, "ParseEqArrMr").Descendants(presentationMath + "mo").Single().AddFirst(
                    new XElement(xsl + "attribute", new XAttribute("name", "stretchy"), "false"))),
            ("CreateTokenAttributes operator style", stylesheet =>
            {
                var operatorStyle = Template(stylesheet, "CreateTokenAttributes").Descendants(xsl + "when")
                    .Single(element => (string?)element.Attribute("test") == "$sTokenType='mo' and $mathvariant!='normal'");
                operatorStyle.ReplaceWith(new XElement(xsl + "when",
                    new XAttribute("test", "$sTokenType='mo' or $sTokenType='mn'"),
                    new XElement(xsl + "if", new XAttribute("test", "$mathvariant!=''"),
                        new XElement(xsl + "attribute", new XAttribute("name", "mathvariant"),
                            new XElement(xsl + "value-of", new XAttribute("select", "$mathvariant"))))));
                foreach (var call in stylesheet.Descendants(xsl + "call-template")
                             .Where(element => (string?)element.Attribute("name") == "CreateTokenAttributes"))
                {
                    var token = call.Elements(xsl + "with-param")
                        .FirstOrDefault(element => (string?)element.Attribute("name") == "sTokenType");
                    if ((string?)token?.Attribute("select") is not ("\'mo\'" or "\'mn\'")) continue;
                    foreach (var name in new[] { "scr", "sty" })
                        call.Elements(xsl + "with-param")
                            .Single(element => (string?)element.Attribute("name") == name)
                            .SetAttributeValue("select", "$" + name);
                }
            }),
            ("CreateTokenAttributes number variant", stylesheet =>
                Template(stylesheet, "CreateTokenAttributes").Descendants(xsl + "variable")
                    .Single(element => (string?)element.Attribute("name") == "mathvariant")
                    .Descendants(xsl + "when")
                    .Single(element => (string?)element.Attribute("test") == "$sTokenType='mn'")
                    // Upright digits use the default only in the ordinary Roman
                    // alphabet. Preserve explicit mathematical alphabets, including
                    // monospace, through the existing script-to-variant mapping.
                    // string() also covers an omitted scr supplied as an empty node-set.
                    .SetAttributeValue("test", "$sTokenType='mn' and (string($scr)='' or $scr='roman') and ($sty='p' or $sty='')")),
            ("text color", AdaptColor));
        var transform = new XslCompiledTransform(enableDebug: false);
        using var reader = stylesheet.CreateReader();
        transform.Load(reader, new XsltSettings(enableDocumentFunction: false, enableScript: false), null);
        return transform;
    }

    private static XElement MathColor(string select) =>
        new(Xsl + "if", new XAttribute("test", "string(" + select + ")!='' and " + select + "!='auto'"),
            new XElement(Xsl + "attribute", new XAttribute("name", "mathcolor"),
                new XElement(Xsl + "value-of", new XAttribute("select", "concat('#'," + select + ")"))));

    private static void AdaptColor(XDocument stylesheet)
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

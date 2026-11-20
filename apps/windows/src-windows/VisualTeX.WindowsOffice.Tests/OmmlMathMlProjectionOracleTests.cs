using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

// Proves the owned OMML projection reproduces Office's classic OMML2MML.XSL
// with VisualTeX's adaptations, on generated documents that cover every OMML
// structure and property value shape.
public sealed class OmmlMathMlProjectionOracleTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace MathMl = "http://www.w3.org/1998/Math/MathML";
    private readonly ITestOutputHelper output;

    public OmmlMathMlProjectionOracleTests(ITestOutputHelper output) => this.output = output;

    // Every generated document compares: no character beyond the Basic
    // Multilingual Plane, which the classic stylesheet cannot always write.
    [Fact]
    public void GeneratedDocumentsProjectExactlyAsTheClassicOfficeStylesheet() => Compare(astral: false);

    // With supplementary-plane characters the stylesheet fails wherever it
    // splits one in half; the projection must still match everywhere else.
    [Fact]
    public void GeneratedDocumentsWithAstralCharactersMatchWhereverTheStylesheetWrites() => Compare(astral: true);

    private void Compare(bool astral)
    {
        // VT_ORACLE_XSL compares against another Office copy of the stylesheet.
        var oraclePath = Environment.GetEnvironmentVariable("VT_ORACLE_XSL");
        var oracle = string.IsNullOrEmpty(oraclePath)
            ? ClassicOfficeReverseStylesheet.Load(ClassicOfficeReverseStylesheet.ResolvePath())
            : ClassicOfficeStylesheetFiles.LoadFirstUsableTransform("OMML2MML.XSL", new[] { oraclePath! }, ClassicOfficeReverseStylesheet.Load);
        var count = int.TryParse(Environment.GetEnvironmentVariable("VT_PROJECTION_CASES"), out var n) ? n : 3000;
        var failures = new List<string>();
        var oracleErrors = new List<string>();
        var oracleErrorCount = 0;
        for (var seed = 1; seed <= count && failures.Count < 5; seed++)
        {
            var omml = new Generator(seed, astral).Document();
            var actual = Canonical(OmmlMathMlProjection.Project(omml));
            string expected;
            try { expected = Canonical(OracleProject(oracle, omml)); }
            catch (Exception error)
            {
                oracleErrorCount++;
                if (oracleErrors.Count < 3) oracleErrors.Add($"seed {seed}: {error.GetType().Name}: {error.Message}\nOMML: {omml}");
                continue;
            }
            if (expected != actual)
                failures.Add($"seed {seed}\nOMML: {omml}\nXSL:  {expected}\nOWN:  {actual}");
        }
        output.WriteLine($"cases={count} stylesheetErrors={oracleErrorCount} failures={failures.Count}");
        foreach (var error in oracleErrors) output.WriteLine("STYLESHEET ERROR " + error);
        Assert.True(failures.Count == 0, string.Join("\n\n", failures));
        if (!astral) Assert.Equal(0, oracleErrorCount);
    }

    internal static XElement OracleProject(XslCompiledTransform transform, string omml)
    {
        var outputSettings = transform.OutputSettings?.Clone() ?? new XmlWriterSettings();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(omml);
        using var source = XmlReader.Create(sourceText, OmmlMathMlProjection.ReaderSettings());
        using var outputText = new StringWriter();
        var arguments = new XsltArgumentList();
        arguments.AddExtensionObject(ClassicOfficeReverseStylesheet.UnicodeScalarXPathNamespace, new ClassicOfficeReverseStylesheet.UnicodeScalarXPathFunctions());
        using (var writer = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, arguments, writer);
        using var transformedText = new StringReader(outputText.ToString());
        using var reader = XmlReader.Create(transformedText, OmmlMathMlProjection.ReaderSettings());
        var document = XDocument.Load(reader, LoadOptions.None);
        return document.Root!.Name.LocalName == "math"
            ? document.Root
            : document.Descendants().First(element => element.Name.LocalName == "math");
    }

    // Element names (MathML namespace), attributes in a stable order, text exactly.
    internal static string Canonical(XElement element)
    {
        var builder = new StringBuilder();
        Write(element, builder);
        return builder.ToString();
    }

    private static void Write(XElement element, StringBuilder builder)
    {
        builder.Append('<').Append(element.Name.Namespace == MathMl ? element.Name.LocalName : element.Name.ToString());
        foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration)
                     .OrderBy(a => a.Name.ToString(), StringComparer.Ordinal))
            builder.Append(' ').Append(attribute.Name).Append("=\"").Append(Visible(attribute.Value)).Append('"');
        builder.Append('>');
        foreach (var node in element.Nodes())
        {
            if (node is XElement child) Write(child, builder);
            else if (node is XText text) builder.Append('[').Append(Visible(text.Value)).Append(']');
        }
        builder.Append("</").Append(element.Name.LocalName).Append('>');
    }

    private static string Visible(string value)
    {
        var builder = new StringBuilder();
        foreach (var character in value)
        {
            if (character is >= ' ' and <= '~') builder.Append(character);
            else builder.Append("\\u").Append(((int)character).ToString("X4"));
        }
        return builder.ToString();
    }

    // Random OMML covering the structures, property shapes and run text that
    // Word writes, plus well-formed variants a newer Word or another producer
    // may write.
    internal sealed class Generator
    {
        private readonly Random random;
        private readonly bool astral;
        private readonly StringBuilder xml = new();

        private readonly int? equations;

        public Generator(int seed, bool astral, int? equations = null)
        {
            random = new Random(seed);
            this.astral = astral;
            this.equations = equations;
        }

        private bool Chance(int percent) => random.Next(100) < percent;
        private T Pick<T>(params T[] items) => items[random.Next(items.Length)];

        public string Document()
        {
            xml.Append("<m:oMathPara xmlns:m=\"").Append(M).Append("\" xmlns:w=\"").Append(W).Append("\">");
            var maths = equations ?? 1 + random.Next(2);
            for (var index = 0; index < maths; index++)
            {
                xml.Append("<m:oMath>");
                Content(3);
                xml.Append("</m:oMath>");
            }
            xml.Append("</m:oMathPara>");
            return xml.ToString();
        }

        private void Content(int depth)
        {
            var items = 1 + random.Next(4);
            for (var index = 0; index < items; index++)
            {
                if (depth <= 0 || Chance(45)) Run();
                else Structure(depth - 1);
            }
        }

        private void Slot(string name, int depth)
        {
            xml.Append("<m:").Append(name).Append('>');
            if (Chance(15))
                xml.Append("<m:argPr>")
                    .Append(Chance(70) ? "<m:scrLvl m:val=\"" + Pick("0", "1", "2", "") + "\"/>" : "<m:scrLvl/>")
                    .Append("</m:argPr>");
            if (!Chance(10)) Content(depth);
            if (Chance(15)) xml.Append("<m:ctrlPr>").Append(WordRunProperties()).Append("</m:ctrlPr>");
            xml.Append("</m:").Append(name).Append('>');
        }

        private string OnOff(string name) => Chance(30)
            ? "<m:" + name + "/>"
            : "<m:" + name + " m:val=\"" + Pick("0", "1", "on", "off", "true", "false", "On", "x", "") + "\"/>";

        private string WordRunProperties()
        {
            var builder = new StringBuilder("<w:rPr>");
            if (Chance(40)) builder.Append("<w:color w:val=\"").Append(Pick("FF0000", "00aa00", "auto", "", "2563EB")).Append("\"/>");
            if (Chance(15)) builder.Append("<w:b/>");
            builder.Append("</w:rPr>");
            return builder.ToString();
        }

        private void Run()
        {
            xml.Append("<m:r>");
            var propertyBlocks = Chance(70) ? 1 : random.Next(3);
            for (var block = 0; block < propertyBlocks; block++)
            {
                xml.Append("<m:rPr>");
                if (Chance(30)) xml.Append("<m:scr m:val=\"").Append(Pick("roman", "script", "fraktur", "double-struck", "sans-serif", "monospace", "")).Append("\"/>");
                if (Chance(5)) xml.Append("<m:scr/>");
                if (Chance(45)) xml.Append("<m:sty m:val=\"").Append(Pick("p", "b", "i", "bi", "")).Append("\"/>");
                if (Chance(5)) xml.Append("<m:sty m:val=\"b\"/>");
                if (Chance(12)) xml.Append(OnOff("nor"));
                if (Chance(5)) xml.Append("<m:lit/>");
                xml.Append("</m:rPr>");
            }
            if (Chance(35)) xml.Append(WordRunProperties());
            var texts = Chance(85) ? 1 : random.Next(3);
            for (var index = 0; index < texts; index++)
            {
                xml.Append(Chance(30) ? "<m:t xml:space=\"preserve\">" : "<m:t>");
                xml.Append(Escape(RunText()));
                xml.Append("</m:t>");
            }
            xml.Append("</m:r>");
        }

        private string RunText()
        {
            var pool = new[]
            {
                "x", "y", "a", "b", "A", "R", "sin", "lim", "0", "1", "2", "9", "12", "3.5", "1,000", "1.2.3", ".", ",",
                "+", "-", "=", "(", ")", "[", "]", "{", "}", "|", "/", "*", "!", "'", "<", ">", "&", "&&", "^", "_",
                "\u2212", "\u00D7", "\u00B7", "\u2061", "\u2062", "\u2211", "\u222B", "\u2192", "\u2264", "\u2208", "\u221E",
                "\u03B1", "\u03B2", "\u0394", "\u03C0", "\u2202", "\u2207", "\u211D", "\u2102", "\u2113",
                " ", "  ", "\u00A0", "\t", "\u2003", "\u0302", "\u0305", "x\u0302",
                "\uFF10", "\uFF0B", "\u4E2D", "\u00B2", "\u2032",
            };
            if (astral) pool = pool.Concat(new[] { "\uD835\uDC65", "\uD835\uDF0B", "x\uD835\uDC00" }).ToArray();
            var parts = 1 + random.Next(4);
            var builder = new StringBuilder();
            for (var index = 0; index < parts; index++) builder.Append(Pick(pool));
            return builder.ToString();
        }

        private static string Escape(string value) =>
            value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("\t", "&#9;");

        private string Properties(string name, params string[] parts) =>
            "<m:" + name + ">" + string.Concat(parts) + "</m:" + name + ">";

        private void Structure(int depth)
        {
            switch (random.Next(22))
            {
                case 0:
                    xml.Append("<m:acc>");
                    if (Chance(80))
                    {
                        xml.Append("<m:accPr>");
                        if (Chance(80))
                            xml.Append(Chance(85)
                                ? "<m:chr m:val=\"" + Escape(Pick("\u0302", "\u0303", "\u0305", "\u0307", "\u0308", "\u20D7", "\u20D6", "\u030C", "\u0306", "\u0301", "\u0300", "\u20E1", "^", "~", "", "ab")) + "\"/>"
                                : "<m:chr/>");
                        if (Chance(20)) xml.Append("<m:ctrlPr>").Append(WordRunProperties()).Append("</m:ctrlPr>");
                        xml.Append("</m:accPr>");
                    }
                    Slot("e", depth);
                    xml.Append("</m:acc>");
                    break;
                case 1:
                    xml.Append("<m:bar>");
                    if (Chance(80)) xml.Append("<m:barPr><m:pos m:val=\"").Append(Pick("top", "bot", "Top", "")).Append("\"/></m:barPr>");
                    Slot("e", depth);
                    xml.Append("</m:bar>");
                    break;
                case 2:
                    xml.Append("<m:borderBox>");
                    if (Chance(80))
                    {
                        xml.Append("<m:borderBoxPr>");
                        foreach (var flag in new[] { "hideTop", "hideBot", "hideLeft", "hideRight", "strikeH", "strikeV", "strikeBLTR", "strikeTLBR" })
                            if (Chance(30)) xml.Append(OnOff(flag));
                        xml.Append("</m:borderBoxPr>");
                    }
                    Slot("e", depth);
                    xml.Append("</m:borderBox>");
                    break;
                case 3:
                    xml.Append("<m:box>");
                    if (Chance(30)) xml.Append(Properties("boxPr", OnOff("opEmu")));
                    Slot("e", depth);
                    xml.Append("</m:box>");
                    break;
                case 4:
                    xml.Append("<m:d>");
                    var delimiterProperties = Chance(80) ? 1 : random.Next(3);
                    for (var index = 0; index < delimiterProperties; index++)
                    {
                        xml.Append("<m:dPr>");
                        if (Chance(50))
                            xml.Append(Chance(90)
                                ? "<m:begChr m:val=\"" + Escape(Pick("(", "[", "{", "|", "\u2016", "\u27E8", "\u230A", "", "((")) + "\"/>"
                                : "<m:begChr/>");
                        if (Chance(20)) xml.Append("<m:sepChr m:val=\"").Append(Escape(Pick(",", "|", ";", ""))).Append("\"/>");
                        if (Chance(50))
                            xml.Append(Chance(90)
                                ? "<m:endChr m:val=\"" + Escape(Pick(")", "]", "}", "|", "\u2016", "\u27E9", "\u230B", "")) + "\"/>"
                                : "<m:endChr/>");
                        if (Chance(15)) xml.Append("<m:grow/>");
                        if (Chance(15)) xml.Append("<m:ctrlPr>").Append(WordRunProperties()).Append("</m:ctrlPr>");
                        xml.Append("</m:dPr>");
                    }
                    var items = 1 + random.Next(3);
                    for (var index = 0; index < items; index++) Slot("e", depth);
                    xml.Append("</m:d>");
                    break;
                case 5:
                    xml.Append("<m:eqArr>");
                    if (Chance(30)) xml.Append("<m:eqArrPr><m:baseJc m:val=\"center\"/></m:eqArrPr>");
                    var rows = 1 + random.Next(3);
                    for (var index = 0; index < rows; index++) Slot("e", depth);
                    xml.Append("</m:eqArr>");
                    break;
                case 6:
                    xml.Append("<m:f>");
                    if (Chance(60))
                    {
                        xml.Append("<m:fPr>");
                        if (Chance(70)) xml.Append("<m:type m:val=\"").Append(Pick("bar", "skw", "lin", "noBar", "Lin", "")).Append("\"/>");
                        if (Chance(25)) xml.Append("<m:ctrlPr>").Append(WordRunProperties()).Append("</m:ctrlPr>");
                        xml.Append("</m:fPr>");
                    }
                    Slot("num", depth);
                    Slot("den", depth);
                    xml.Append("</m:f>");
                    break;
                case 7:
                    xml.Append("<m:func><m:funcPr/>");
                    Slot("fName", depth);
                    Slot("e", depth);
                    xml.Append("</m:func>");
                    break;
                case 8:
                    xml.Append("<m:groupChr>");
                    if (Chance(80))
                    {
                        xml.Append("<m:groupChrPr>");
                        if (Chance(60))
                            xml.Append(Chance(85)
                                ? "<m:chr m:val=\"" + Escape(Pick("\u23DF", "\u23DE", "\u2190", "\u2192", "", "ab")) + "\"/>"
                                : "<m:chr/>");
                        if (Chance(50)) xml.Append("<m:pos m:val=\"").Append(Pick("top", "bot", "")).Append("\"/>");
                        if (Chance(50)) xml.Append("<m:vertJc m:val=\"").Append(Pick("top", "bot", "")).Append("\"/>");
                        xml.Append("</m:groupChrPr>");
                    }
                    Slot("e", depth);
                    xml.Append("</m:groupChr>");
                    break;
                case 9:
                case 10:
                    var limit = Pick("limLow", "limUpp");
                    xml.Append("<m:").Append(limit).Append('>');
                    Slot("e", depth);
                    Slot("lim", depth);
                    xml.Append("</m:").Append(limit).Append('>');
                    break;
                case 11:
                    xml.Append("<m:m>");
                    if (Chance(50))
                    {
                        xml.Append("<m:mPr><m:mcs>");
                        var columns = 1 + random.Next(2);
                        for (var index = 0; index < columns; index++)
                            xml.Append("<m:mc><m:mcPr><m:count m:val=\"1\"/><m:mcJc m:val=\"")
                                .Append(Pick("left", "right", "center", "Left")).Append("\"/></m:mcPr></m:mc>");
                        xml.Append("</m:mcs></m:mPr>");
                    }
                    var matrixRows = 1 + random.Next(2);
                    var matrixColumns = 1 + random.Next(2);
                    for (var row = 0; row < matrixRows; row++)
                    {
                        xml.Append("<m:mr>");
                        for (var column = 0; column < matrixColumns; column++) Slot("e", depth);
                        xml.Append("</m:mr>");
                    }
                    xml.Append("</m:m>");
                    break;
                case 12:
                case 13:
                    xml.Append("<m:nary>");
                    if (Chance(85))
                    {
                        xml.Append("<m:naryPr>");
                        if (Chance(70)) xml.Append("<m:chr m:val=\"").Append(Escape(Pick("\u2211", "\u220F", "\u222B", "\u222E", "\u22C3", ""))).Append("\"/>");
                        if (Chance(50)) xml.Append("<m:limLoc m:val=\"").Append(Pick("undOvr", "subSup", "")).Append("\"/>");
                        if (Chance(25)) xml.Append(OnOff("grow"));
                        if (Chance(30)) xml.Append(OnOff("subHide"));
                        if (Chance(30)) xml.Append(OnOff("supHide"));
                        xml.Append("</m:naryPr>");
                    }
                    Slot("sub", depth);
                    Slot("sup", depth);
                    Slot("e", depth);
                    xml.Append("</m:nary>");
                    break;
                case 14:
                    xml.Append("<m:phant>");
                    if (Chance(75))
                    {
                        xml.Append("<m:phantPr>");
                        foreach (var flag in new[] { "show", "zeroWid", "zeroAsc", "zeroDesc", "transp" })
                            if (Chance(40)) xml.Append(OnOff(flag));
                        xml.Append("</m:phantPr>");
                    }
                    Slot("e", depth);
                    xml.Append("</m:phant>");
                    break;
                case 15:
                    xml.Append("<m:rad>");
                    if (Chance(70))
                    {
                        xml.Append("<m:radPr>");
                        if (Chance(60)) xml.Append(OnOff("degHide"));
                        if (Chance(25)) xml.Append("<m:ctrlPr>").Append(WordRunProperties()).Append("</m:ctrlPr>");
                        xml.Append("</m:radPr>");
                    }
                    Slot("deg", depth);
                    Slot("e", depth);
                    xml.Append("</m:rad>");
                    break;
                case 16:
                    xml.Append("<m:sPre>");
                    Slot("sub", depth);
                    Slot("sup", depth);
                    Slot("e", depth);
                    xml.Append("</m:sPre>");
                    break;
                case 17:
                    xml.Append("<m:sSub>");
                    Slot("e", depth);
                    Slot("sub", depth);
                    xml.Append("</m:sSub>");
                    break;
                case 18:
                    xml.Append("<m:sSup>");
                    Slot("e", depth);
                    Slot("sup", depth);
                    xml.Append("</m:sSup>");
                    break;
                case 19:
                    xml.Append("<m:sSubSup>");
                    Slot("e", depth);
                    Slot("sub", depth);
                    Slot("sup", depth);
                    xml.Append("</m:sSubSup>");
                    break;
                case 20:
                    // Content a newer Word may write: unknown OMML and Word wrappers.
                    var wrapper = Pick("m:futureObject", "w:ins", "w:customXml", "m:box");
                    xml.Append('<').Append(wrapper).Append('>');
                    Content(depth);
                    xml.Append("</").Append(wrapper).Append('>');
                    break;
                default:
                    xml.Append("<w:r><w:t>prose</w:t></w:r>");
                    break;
            }
        }
    }
}

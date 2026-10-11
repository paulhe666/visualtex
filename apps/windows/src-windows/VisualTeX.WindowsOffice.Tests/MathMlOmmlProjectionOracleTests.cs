using System.Text;
using System.Xml;
using System.Xml.Xsl;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

// Proves the owned MathML projection reproduces Office's classic MML2OMML.XSL
// with VisualTeX's adaptations: the OMML the insertion pipeline receives is
// identical, character for character, on generated documents covering every
// MathML element and attribute shape the stylesheet distinguishes.
public sealed class MathMlOmmlProjectionOracleTests
{
    private const string MathMl = "http://www.w3.org/1998/Math/MathML";
    private readonly ITestOutputHelper output;

    public MathMlOmmlProjectionOracleTests(ITestOutputHelper output) => this.output = output;

    [Fact]
    public void GeneratedDocumentsProjectExactlyAsTheClassicOfficeStylesheet() => Compare(astral: false);

    [Fact]
    public void GeneratedDocumentsWithAstralCharactersMatchWhereverTheStylesheetWrites() => Compare(astral: true);

    private void Compare(bool astral)
    {
        var oraclePath = Environment.GetEnvironmentVariable("VT_FORWARD_ORACLE_XSL");
        var oracle = string.IsNullOrEmpty(oraclePath)
            ? ClassicOfficeForwardStylesheet.Load(ClassicOfficeForwardStylesheet.ResolvePath())
            : ClassicOfficeStylesheetFiles.LoadFirstUsableTransform("MML2OMML.XSL", new[] { oraclePath! }, ClassicOfficeForwardStylesheet.Load);
        var count = int.TryParse(Environment.GetEnvironmentVariable("VT_PROJECTION_CASES"), out var n) ? n : 3000;
        var failures = new List<string>();
        var oracleErrors = new List<string>();
        var oracleErrorCount = 0;
        for (var seed = 1; seed <= count && failures.Count < 5; seed++)
        {
            var mathMl = new Generator(seed, astral).Document();
            var actual = WordOmmlConverter.ExtractSingleOMath(MathMlOmmlProjection.Project(mathMl));
            string expected;
            try { expected = WordOmmlConverter.ExtractSingleOMath(OracleProject(oracle, mathMl)); }
            catch (Exception error)
            {
                oracleErrorCount++;
                if (oracleErrors.Count < 3) oracleErrors.Add($"seed {seed}: {error.GetType().Name}: {error.Message}");
                continue;
            }
            if (expected != actual)
                failures.Add($"seed {seed}\nMATHML: {mathMl}\nXSL: {expected}\nOWN: {actual}");
        }
        output.WriteLine($"cases={count} stylesheetErrors={oracleErrorCount} failures={failures.Count}");
        foreach (var error in oracleErrors) output.WriteLine("STYLESHEET ERROR " + error);
        Assert.True(failures.Count == 0, string.Join("\n\n", failures));
        if (!astral) Assert.Equal(0, oracleErrorCount);
    }

    internal static string OracleProject(XslCompiledTransform transform, string mathMl)
    {
        var outputSettings = transform.OutputSettings?.Clone() ?? new XmlWriterSettings();
        outputSettings.OmitXmlDeclaration = true;
        outputSettings.Encoding = new UTF8Encoding(false);
        using var sourceText = new StringReader(mathMl);
        using var source = XmlReader.Create(sourceText, OmmlMathMlProjection.ReaderSettings());
        using var outputText = new StringWriter();
        using (var writer = XmlWriter.Create(outputText, outputSettings))
            transform.Transform(source, writer);
        return outputText.ToString();
    }

    // Random presentation MathML: every element and attribute value shape
    // the classic stylesheet branches on, plus content it does not know.
    internal sealed class Generator
    {
        private readonly Random random;
        private readonly bool astral;
        private readonly StringBuilder xml = new();

        public Generator(int seed, bool astral)
        {
            random = new Random(seed);
            this.astral = astral;
        }

        private bool Chance(int percent) => random.Next(100) < percent;
        private T Pick<T>(params T[] items) => items[random.Next(items.Length)];

        public string Document()
        {
            xml.Append("<math xmlns=\"").Append(MathMl).Append("\" xmlns:mml=\"").Append(MathMl).Append('"');
            if (Chance(30)) xml.Append(" display=\"").Append(Pick("block", "inline")).Append('"');
            xml.Append('>');
            Content(3);
            xml.Append("</math>");
            return xml.ToString();
        }

        private static string Escape(string value) =>
            value.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("\t", "&#9;");

        private void Attribute(string name, params string[] values)
        {
            var qualified = Chance(4);
            xml.Append(' ').Append(qualified ? "mml:" + name : name).Append("=\"").Append(Escape(Pick(values))).Append('"');
        }

        private void Content(int depth)
        {
            var items = 1 + random.Next(5);
            for (var index = 0; index < items; index++)
            {
                if (depth <= 0 || Chance(50)) Token();
                else Structure(depth - 1);
            }
        }

        private void Row(int depth, string name = "mrow")
        {
            xml.Append('<').Append(name);
            if (name == "mstyle")
            {
                if (Chance(30)) Attribute("scriptlevel", "0", "1", "2", "+1", "-1", "");
                if (Chance(20)) Attribute("color", "red", "#00AA00", "");
                if (Chance(20)) Attribute("mathvariant", "bold", "normal", "italic");
            }
            if (Chance(10)) Attribute("mathcolor", "#FF0000", "#2563EB");
            xml.Append('>');
            Content(depth);
            xml.Append("</").Append(name).Append('>');
        }

        private string Text(string kind)
        {
            var letters = new[] { "x", "y", "a", "ab", "sin", "lim", "R", "\u03B1", "\u0394", "\u03C0", " x ", "x y", "\u00A0" };
            var numbers = new[] { "1", "12", "3.5", ".5", "5.", "-2", "+1", "1,000", "1e3", "1E-3", "1e400", "Infinity", ".", " 7 ", "\u0661", "\uFF11" };
            var operators = new[] { "+", "-", "=", "(", ")", "/", "\u2061", "\u2062", "\u200B", "\u2211", "\u222B", "\u220F", "\u22C3", "|", "_", "\u00AF", "\u0332", "\u0305", "^", "~", "\u2192", "\u2A75", "&", "<", ",", "" };
            var prose = new[] { "if ", " and ", "x\u2062y", "a\u00A0b", "  two  spaces ", "\u2A75", "text" };
            var pool = kind switch
            {
                "mi" => letters,
                "mn" => numbers,
                "mo" => operators,
                _ => prose,
            };
            var value = Pick(pool);
            if (astral && Chance(15)) value += Pick("\uD835\uDC65", "\uD835\uDF0B", "\uD83D\uDE00");
            return value;
        }

        private void Token(string? forced = null)
        {
            var kind = forced ?? Pick("mi", "mi", "mi", "mn", "mn", "mo", "mo", "mo", "mtext", "ms");
            xml.Append('<').Append(kind);
            if (Chance(35)) Attribute("mathvariant", "normal", "bold", "italic", "bold-italic", "double-struck", "script", "bold-script", "fraktur", "bold-fraktur", "sans-serif", "bold-sans-serif", "sans-serif-italic", "sans-serif-bold-italic", "monospace", "initial", "Bold", "");
            if (Chance(15)) Attribute("fontstyle", "normal", "italic", "");
            if (Chance(15)) Attribute("fontweight", "normal", "bold", "");
            if (Chance(12)) Attribute("mathcolor", "#FF0000", "#00aa00", "#2563EB");
            if (Chance(6)) Attribute("color", "#123456", "blue");
            if (Chance(8)) Attribute("font-family", "Arial", "Cambria Math", "");
            if (Chance(5)) Attribute("mathsize", "2em", "big");
            if (Chance(15)) Attribute("data-visualtex-omml-run-role", "operator", "identifier", "number", "");
            if (kind == "mo" && Chance(20)) Attribute("stretchy", "true", "false", "True");
            if (kind == "mo" && Chance(10)) Attribute("accent", "true", "false");
            if (kind == "ms" && Chance(30)) Attribute("lquote", "\u00AB", "'", "");
            if (kind == "ms" && Chance(30)) Attribute("rquote", "\u00BB", "'", "");
            xml.Append('>');
            if (Chance(4))
            {
                xml.Append("<mglyph alt=\"").Append(Escape(Pick("glyph", " g ", ""))).Append("\" src=\"g.png\"");
                if (Chance(30)) xml.Append(" mathcolor=\"#FF0000\"");
                xml.Append("/>");
                if (Chance(50)) xml.Append(Escape(Text(kind)));
            }
            else xml.Append(Escape(Text(kind)));
            xml.Append("</").Append(kind).Append('>');
        }

        private void Operand(int depth)
        {
            if (Chance(55)) Token();
            else if (Chance(60)) Row(depth);
            else Structure(depth);
        }

        private void NaryOperator()
        {
            if (Chance(70))
            {
                xml.Append("<mo");
                if (Chance(30)) Attribute("stretchy", "true", "false", "TRUE");
                xml.Append('>').Append(Pick("\u2211", "\u222B", "\u220F", "\u222E", "\u22C3", "\u2210", "\u222C", " \u2211 ", "+")).Append("</mo>");
            }
            else xml.Append("<mrow><mo>").Append(Pick("\u2211", "\u222B")).Append("</mo></mrow>");
        }

        private void Structure(int depth)
        {
            switch (random.Next(24))
            {
                case 0:
                    xml.Append("<mfrac");
                    if (Chance(30)) Attribute("linethickness", "0", "0pt", "0.0em", "thin", "THICK", "2px", "", "medium", "0%");
                    if (Chance(20)) Attribute("bevelled", "true", "false");
                    if (Chance(15)) Attribute("mathcolor", "#FF0000");
                    xml.Append('>');
                    Operand(depth); Operand(depth);
                    xml.Append("</mfrac>");
                    break;
                case 1:
                    Row(depth, Pick("msqrt", "msqrt"));
                    break;
                case 2:
                    xml.Append("<menclose");
                    if (Chance(80)) Attribute("notation", "radical", "box", "roundedbox", "circle", "top", "bottom left", "left right top bottom", "horizontalstrike", "updiagonalstrike downdiagonalstrike", "verticalstrike box", "actuarial", "longdiv", "", "TOP");
                    if (Chance(15)) Attribute("mathcolor", "#FF0000");
                    xml.Append('>');
                    Content(depth);
                    xml.Append("</menclose>");
                    break;
                case 3:
                    xml.Append("<mroot");
                    if (Chance(15)) Attribute("mathcolor", "#00AA00");
                    xml.Append('>');
                    Operand(depth); Operand(depth);
                    xml.Append("</mroot>");
                    break;
                case 4:
                case 5:
                {
                    var name = Pick("msub", "msup", "msubsup");
                    xml.Append('<').Append(name).Append('>');
                    if (Chance(35)) NaryOperator(); else Operand(depth);
                    Operand(depth);
                    if (name == "msubsup") Operand(depth);
                    xml.Append("</").Append(name).Append('>');
                    if (Chance(50)) Row(depth, Pick("mrow", "mstyle"));
                    break;
                }
                case 6:
                case 7:
                {
                    var name = Pick("munder", "mover", "munderover");
                    xml.Append('<').Append(name);
                    if (name != "munder" && Chance(40)) Attribute("accent", "true", "false", "TRUE", "");
                    if (name != "mover" && Chance(40)) Attribute("accentunder", "true", "false", "");
                    xml.Append('>');
                    var shape = random.Next(5);
                    if (shape == 0) NaryOperator(); else if (shape == 1) Row(depth); else Operand(depth);
                    if (Chance(55))
                    {
                        xml.Append("<mo");
                        if (Chance(30)) Attribute("accent", "true", "false");
                        if (Chance(40)) Attribute("mathvariant", "normal", "bold", "italic", "bold-italic", "script");
                        if (Chance(20)) Attribute("mathcolor", "#FF0000");
                        xml.Append('>').Append(Escape(Pick("\u00AF", "\u0332", "_", "\u0305", "^", "~", "\u02D9", "\u2192", "\u23DF", "\u23DE", ".", "-", "\u2212", "\u00A8", "ab", "\u02C7"))).Append("</mo>");
                    }
                    else Operand(depth);
                    if (name == "munderover") Operand(depth);
                    xml.Append("</").Append(name).Append('>');
                    if (Chance(40)) Row(depth, Pick("mrow", "mstyle"));
                    break;
                }
                case 8:
                    xml.Append("<mfenced");
                    if (Chance(50)) Attribute("open", "(", "[", "{", "|", "\u27E8", "", "((");
                    if (Chance(50)) Attribute("close", ")", "]", "}", "|", "\u27E9", "");
                    if (Chance(40)) Attribute("separators", ",", "|", ";,", "", " ");
                    xml.Append('>');
                    var items = random.Next(4);
                    for (var index = 0; index < items; index++) Operand(depth);
                    xml.Append("</mfenced>");
                    break;
                case 9:
                case 10:
                    Table(depth);
                    break;
                case 11:
                    xml.Append("<mmultiscripts>");
                    Operand(depth);
                    var post = random.Next(5);
                    for (var index = 0; index < post; index++) { if (Chance(30)) xml.Append("<none/>"); else Token(); }
                    if (Chance(70))
                    {
                        xml.Append("<mprescripts/>");
                        var pre = random.Next(5);
                        for (var index = 0; index < pre; index++) { if (Chance(30)) xml.Append("<none/>"); else Token(); }
                    }
                    xml.Append("</mmultiscripts>");
                    break;
                case 12:
                    xml.Append("<mpadded");
                    if (Chance(50)) Attribute("width", "0", "0em", "1em", "+0.5em", "height", "");
                    if (Chance(50)) Attribute("height", "0", "2ex", "");
                    if (Chance(50)) Attribute("depth", "0", "0.0", "1");
                    xml.Append('>');
                    Content(depth);
                    xml.Append("</mpadded>");
                    break;
                case 13:
                    xml.Append("<mphantom>");
                    if (Chance(40))
                    {
                        xml.Append("<mpadded");
                        if (Chance(60)) Attribute("width", "0", "1em", "0EM");
                        if (Chance(60)) Attribute("height", "0", "1ex");
                        if (Chance(60)) Attribute("depth", "0", "2pt");
                        xml.Append('>');
                        Content(depth);
                        xml.Append("</mpadded>");
                    }
                    else Content(depth);
                    xml.Append("</mphantom>");
                    break;
                case 14:
                    xml.Append("<mspace");
                    if (Chance(70)) xml.Append(" data-visualtex-omml-space=\"").Append(Escape(Pick(" ", "\u2003", "\u2009\u2009", ""))).Append('"');
                    if (Chance(50)) xml.Append(" width=\"1em\"");
                    xml.Append("/>");
                    break;
                case 15:
                    // Linear fraction shape.
                    xml.Append("<mrow>");
                    Row(depth);
                    xml.Append("<mo>").Append(Pick("/", " / ", "\u2215")).Append("</mo>");
                    Row(depth);
                    xml.Append("</mrow>");
                    break;
                case 16:
                    // Function application shape.
                    xml.Append("<mrow>");
                    if (Chance(70)) Row(depth); else Token("mi");
                    xml.Append("<mo>\u2061</mo>");
                    if (Chance(70)) Row(depth); else Token();
                    xml.Append("</mrow>");
                    break;
                case 17:
                    Row(depth, "mstyle");
                    break;
                case 18:
                    xml.Append("<mglyph alt=\"").Append(Pick("alpha", " a b ", "")).Append("\" src=\"x.png\"/>");
                    break;
                case 19:
                    var wrapper = Pick("semantics", "maction", "merror2", "annotation-holder");
                    xml.Append('<').Append(wrapper).Append('>');
                    Content(depth);
                    xml.Append("</").Append(wrapper).Append('>');
                    break;
                case 20:
                    xml.Append("<mover accent=\"true\">");
                    Operand(depth);
                    xml.Append("<mo mathvariant=\"").Append(Pick("normal", "bold", "italic", "bold-italic")).Append('"');
                    if (Chance(40)) xml.Append(" mathcolor=\"#FF0000\"");
                    xml.Append('>').Append(Pick("\u02C6", "\u02DC", "\u2192", "\u00AF", "^", "ab")).Append("</mo></mover>");
                    break;
                case 21:
                    xml.Append("<foreign:node xmlns:foreign=\"urn:other\">");
                    Content(depth);
                    xml.Append("</foreign:node>");
                    break;
                default:
                    Row(depth);
                    break;
            }
        }

        private void Table(int depth)
        {
            xml.Append("<mtable");
            if (Chance(20)) Attribute("frame", "none", "solid", "");
            if (Chance(15)) Attribute("columnlines", "none", "solid");
            if (Chance(15)) Attribute("rowlines", "none", "dashed");
            xml.Append('>');
            var rows = 1 + random.Next(3);
            var single = Chance(55);
            for (var row = 0; row < rows; row++)
            {
                if (Chance(8))
                {
                    xml.Append("<mlabeledtr><mtd><mtext>(1)</mtext></mtd><mtd>");
                    Content(depth);
                    xml.Append("</mtd></mlabeledtr>");
                    continue;
                }
                if (Chance(5)) { Token(); continue; }
                xml.Append("<mtr>");
                var cells = single ? 1 : random.Next(4);
                for (var cell = 0; cell < cells; cell++)
                {
                    xml.Append("<mtd>");
                    var parts = 1 + random.Next(4);
                    for (var part = 0; part < parts; part++)
                    {
                        var choice = random.Next(10);
                        if (choice == 0) xml.Append("<maligngroup/>");
                        else if (choice == 1) xml.Append("<malignmark/>");
                        else if (choice == 2)
                        {
                            xml.Append("<mrow>");
                            Token();
                            xml.Append(Pick("<maligngroup/>", "<malignmark/>"));
                            Token();
                            xml.Append("</mrow>");
                        }
                        else if (choice == 3)
                        {
                            xml.Append("<mfrac><mrow><maligngroup/>");
                            Token();
                            xml.Append("</mrow>");
                            Token();
                            xml.Append("</mfrac>");
                        }
                        else Operand(depth);
                    }
                    xml.Append("</mtd>");
                }
                xml.Append("</mtr>");
            }
            xml.Append("</mtable>");
        }
    }
}

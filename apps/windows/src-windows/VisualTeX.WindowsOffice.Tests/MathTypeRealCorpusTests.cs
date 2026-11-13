using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Equations saved by MathType itself (test-fixtures/mathtype-real). The other MathType tests
// read back what VisualTeX wrote, which cannot show what the reader does not know about
// MathType's own output: every item below was wrong or failed before this corpus was read.
public sealed class MathTypeRealCorpusTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static string Root => Path.Combine(AppContext.BaseDirectory, "mathtype-real");

    public static IEnumerable<object[]> Equations() =>
        System.IO.Directory.GetFiles(Root, "*.native")
            .Select(path => new object[] { Path.GetFileNameWithoutExtension(path) });

    private static string MathMl(string name) =>
        MathTypeMtefCodec.ReadEquationNativeMathMl(File.ReadAllBytes(Path.Combine(Root, name + ".native")));

    private static string Latex(string name) => MathMlToLatexConverter.Convert(MathMl(name));

    private static XElement Omml(string name) => XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(MathMl(name)));

    private static string Text(XElement? element) =>
        element is null ? string.Empty : string.Concat(element.Descendants(M + "t").Select(t => t.Value));

    [Fact]
    public void TheCorpusIsPresent() => Assert.True(Equations().Count() >= 26);

    [Theory]
    [MemberData(nameof(Equations))]
    public void ConvertsToLatexWithoutRawGlyphs(string name)
    {
        var latex = Latex(name);
        Assert.False(string.IsNullOrWhiteSpace(latex));
        // A character that has no LaTeX name here would be pasted as a raw glyph.
        // The contour integrals past \oint are written as their characters on purpose.
        Assert.DoesNotContain(latex, character =>
            character > 0x7F && character != 0xA0 && !char.IsLetter(character)
            && character is not (>= (char)0x222F and <= (char)0x2233));
    }

    [Theory]
    [MemberData(nameof(Equations))]
    public void ConvertsToOmmlWithEveryOperatorAndFunctionOwningItsOperand(string name)
    {
        var omml = Omml(name);
        foreach (var owner in omml.Descendants().Where(e => e.Name == M + "nary" || e.Name == M + "func"))
        {
            var operand = owner.Element(M + "e");
            Assert.NotNull(operand);
            Assert.Contains(operand!.Elements(), child => child.Name.LocalName is not ("argPr" or "ctrlPr"));
        }
    }

    [Fact]
    public void SummationStyleOperatorKeepsItsOwnSign() =>
        Assert.Contains(@"C\sum_{i=1}^{l}\xi_{i}", Latex("299"));

    [Fact]
    public void OperandTypedAfterTheTemplateIsStillTheSumBody()
    {
        var sum = Assert.Single(Omml("299").Descendants(M + "nary"));
        Assert.Equal("ξi", Text(sum.Element(M + "e")));
    }

    [Fact]
    public void FunctionStyleNameIsAnAppliedFunction()
    {
        Assert.Contains(@"\ln a_{", Latex("long_uint_comment"));
        var functions = Omml("long_uint_comment").Descendants(M + "func").ToArray();
        Assert.Equal(2, functions.Length);
        Assert.All(functions, function => Assert.Equal("ln", Text(function.Element(M + "fName"))));
    }

    [Fact]
    public void EveryFenceTemplateIsRead()
    {
        var latex = Latex("fences");
        Assert.Contains(@"\left\llbracket p\right\rrbracket", latex);
        Assert.Contains(@"\left\langle a|b\right\rangle", latex);
        Assert.Contains(@"\overbracket{a}^{b}+\underbracket{a}_{b}", latex);
    }

    [Fact]
    public void EveryIntegralKindKeepsItsSign()
    {
        var signs = string.Concat(XElement.Parse(MathMl("integrals")).Descendants()
            .Where(e => e.Name.LocalName == "mo").Select(e => e.Value)
            .Where(value => value.Length == 1 && value[0] is >= '∫' and <= '∳').Distinct());
        Assert.Equal("∫∬∭∮∯∰∳∲", signs);
    }

    [Fact]
    public void WholeOperandIsTheSumBody()
    {
        var sum = Assert.Single(Omml("equation1").Descendants(M + "nary"));
        Assert.Single(sum.Element(M + "e")!.Elements(M + "f"));
        Assert.Empty(sum.ElementsAfterSelf());
    }
}

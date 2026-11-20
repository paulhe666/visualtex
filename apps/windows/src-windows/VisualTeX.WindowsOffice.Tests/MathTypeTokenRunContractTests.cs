using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeTokenRunContractTests
{
    [Theory]
    [InlineData("bold")][InlineData("italic")][InlineData("bold-italic")]
    [InlineData("sans-serif")][InlineData("monospace")]
    public void ExplicitNumberFontsKeepDecimalRolesInEveryContainer(string variant)
    {
        foreach (var value in new[] { "123.45", ".125", "12.", "١٢٫٥" })
        foreach (var container in new[] { "root", "fraction", "subscript" })
        {
            var number = new XElement("mn", new XAttribute("mathvariant", variant), value);
            var body = container == "fraction" ? new XElement("mfrac", number, new XElement("mn", "7"))
                : container == "subscript" ? new XElement("msub", new XElement("mi", "x"), number) : number;
            var source = new XElement("math", body).ToString();
            var native = MathTypeMtefCodec.CreateEquationNative(source, false);
            var actual = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
            var restored = Assert.Single(actual.Descendants().Where(e => e.Value == value && !e.HasElements));
            Assert.Equal("mn", restored.Name.LocalName);
            Assert.Equal(variant, (string?)restored.Attribute("mathvariant"));
            Assert.Equal(value, restored.Value);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual.ToString()));
            var edited = MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, actual.ToString(), false);
            Assert.Equal(native.Mtef, edited.Mtef);
        }
    }

    [Theory]
    [InlineData("mass = 2 kg")][InlineData("a+b≤c")][InlineData("price: 2.5%")]
    [InlineData("ℝ≠ℂ")][InlineData("[a,b]")]
    public void NativeTextRoleKeepsPunctuationAndSymbolsLiteral(string text)
    {
        var source = new XElement("math", new XElement("mtext", text)).ToString();
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        var token = Assert.Single(actual.Root!.Elements());
        Assert.Equal("mtext", token.Name.LocalName);
        Assert.Equal(text, token.Value);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual.ToString()));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(source.Replace("mtext", "mo")));
    }

    [Theory]
    [InlineData("+-")][InlineData(":=")][InlineData("≤→")][InlineData("=≥-")]
    public void OperatorRunPreservesEveryGlyphAndStyle(string operators)
    {
        foreach (var variant in new[] { "normal", "italic", "bold", "bold-italic" })
        {
            var source = new XElement("math", new XElement("mo", new XAttribute("mathvariant", variant), operators)).ToString();
            var native = MathTypeMtefCodec.CreateEquationNative(source, false);
            var tokens = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative)).Root!.Elements().ToArray();
            Assert.All(tokens, token => Assert.Equal("mo", token.Name.LocalName));
            Assert.Equal(operators.Replace('-', '−'), string.Concat(tokens.Select(t => t.Value)).Replace('-', '−'));
            if (variant != "normal") Assert.All(tokens, token => Assert.Equal(variant, (string?)token.Attribute("mathvariant")));
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(new XElement("math", tokens).ToString()));
            if (variant != "normal") Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(source.Replace(variant, "normal")));
        }
    }

    [Fact]
    public void DecimalReconstructionDoesNotCrossStyleSpaceOrStructureBoundaries()
    {
        const string source = "<math><mn mathvariant='bold'>12</mn><mi mathvariant='italic'>.</mi><mn mathvariant='bold'>5</mn><mspace width='1em'/><mi mathvariant='bold'>.</mi></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        Assert.Equal(new[] { "mn", "mi", "mn", "mspace", "mi" }, actual.Root!.Elements().Select(e => e.Name.LocalName));
        Assert.Equal(new[] { "12", "5" }, actual.Descendants().Where(e => e.Name.LocalName == "mn").Select(e => e.Value));
    }
}

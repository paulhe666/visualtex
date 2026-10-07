using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeTextStyleContractTests
{
    [Theory]
    [InlineData("bold", "Times New Roman")]
    [InlineData("italic", "Times New Roman")]
    [InlineData("bold-italic", "Times New Roman")]
    [InlineData("sans-serif", "Arial")]
    [InlineData("bold-sans-serif", "Arial")]
    [InlineData("sans-serif-italic", "Arial")]
    [InlineData("sans-serif-bold-italic", "Arial")]
    [InlineData("monospace", "Courier New")]
    public void LiteralTextKeepsNativeRoleFontAndStyleThroughCreationAndEditing(string variant, string font)
    {
        foreach (var value in new[] { "rate", "speed = 3", "hello world", "a+b=2" })
        foreach (var wrap in new[] { false, true })
        {
            var text = new XElement("mtext", new XAttribute("mathvariant", variant), value);
            var source = new XElement("math", new XElement("mi", new XAttribute("mathvariant", "bold-italic"), "x"),
                wrap ? new XElement("mfrac", text, new XElement("mn", "7")) : text).ToString();
            var native = MathTypeMtefCodec.CreateEquationNative(source, false);
            var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
            var actual = XDocument.Parse(xml);
            var restored = Assert.Single(actual.Descendants(), e => e.Name.LocalName == "mtext");
            Assert.Equal(value, restored.Value);
            Assert.Equal(variant, (string?)restored.Attribute("mathvariant"));
            Assert.Contains(font + "\0", Encoding.ASCII.GetString(native.Mtef, 0, native.StructureOffset));
            // Literal punctuation still uses fnTEXT, not a Symbol CHAR whose
            // math spacing would turn the text into a different expression.
            foreach (var c in value.Where(c => !char.IsWhiteSpace(c)))
                Assert.True(native.Mtef.AsSpan(native.StructureOffset).IndexOf(new byte[] { 2, 0, 129, (byte)c, 0 }) >= 0);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
            restored.Attribute("mathvariant")!.Remove();
            Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual.ToString()));
            Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
        }
    }

    [Theory]
    [InlineData("bold")][InlineData("italic")][InlineData("bold-italic")]
    public void TextStyleLossIsRejectedEvenWhenCharactersAreIdentical(string variant)
    {
        foreach (var value in new[] { "rate", "mass = 2 kg", "hello world", "3" })
        {
            var source = new XElement("math", new XElement("mtext", new XAttribute("mathvariant", variant), value));
            var lost = new XElement(source);
            lost.Descendants("mtext").Single().Attribute("mathvariant")!.Remove();
            Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source.ToString()), MathTypeMtefCodec.SemanticSignature(lost.ToString()));
        }
    }

    [Fact]
    public void MixedNativeTextBindingsCannotMasqueradeAsCompleteFidelity()
    {
        // This is a rejection contract, not a successful mixed-style conversion.
        const string source = "<math><mtext mathvariant='bold'>bold = 2</mtext><mtext mathvariant='italic'>italic = 3</mtext></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var read = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(read));
    }

    [Fact]
    public void ATextSpaceMustNotHideDifferentWesternAndFarEastStyles()
    {
        // The separate Far-East logical face has not acquired bold merely
        // because the preceding western text has. This is a loss rejection.
        const string source = "<math><mtext mathvariant='bold'>a 中</mtext></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var read = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var actual = XDocument.Parse(read);
        Assert.Equal(2, actual.Descendants().Count(e => e.Name.LocalName == "mtext"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(read));
    }
}

using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNativeAngleAndParallelTests
{
    public static IEnumerable<object[]> Angles()
    {
        foreach (var glyph in new[] { "∠", "∡", "∢" })
        foreach (var variant in new[] { "", "normal", "italic", "bold-italic" })
            yield return new object[] { glyph, variant };
    }

    [Theory]
    [MemberData(nameof(Angles))]
    public void AngleGlyphAndEffectiveStyleRoundTrip(string glyph, string variant)
    {
        var source = $"<math><mi{(variant.Length == 0 ? "" : " mathvariant='" + variant + "'")}>{glyph}</mi></math>";
        var (native, xml, token) = Decode(source);
        Assert.Equal("mi", token.Name.LocalName);
        Assert.Equal(variant.Length == 0 ? "italic" : variant, (string?)token.Attribute("mathvariant"));
        Assert.Equal(glyph == "∠" ? 0xD0 : glyph == "∡" ? 0x52 : 0x53, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal((int)glyph[0], (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Theory]
    [InlineData(false, "normal")][InlineData(false, "italic")][InlineData(false, "bold-italic")]
    [InlineData(true, "normal")][InlineData(true, "italic")][InlineData(true, "bold-italic")]
    public void NegatedParallelPreservesItsOrdinaryOrShortGlyph(bool alternate, string variant)
    {
        var source = $"<math><mo mathvariant='{variant}'{(alternate ? " data-mjx-alternate='1'" : "")}>∦</mo></math>";
        var (native, xml, token) = Decode(source);
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal("∦", token.Value);
        Assert.Equal(alternate ? "1" : null, (string?)token.Attribute("data-mjx-alternate"));
        Assert.Equal(alternate ? 0xEA2F : 0x2226, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(alternate ? 0x8B : 0x8D, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Fact]
    public void ImplicitSingleSymbolMiMeansItalicButOrdinaryMoDoesNot()
    {
        Assert.Equal(MathTypeMtefCodec.SemanticSignature("<math><mi>∡</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='italic'>∡</mi></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mo>∥</mo></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi>∥</mi></math>"));
    }

    [Fact]
    public void InheritedNormalAndNearestExplicitOverrideRemainAuthoritative()
    {
        const string source = "<math><mstyle mathvariant='normal'><mi>∡</mi><mi mathvariant='italic'>∢</mi></mstyle></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(new[] { "normal", "italic" }, XDocument.Parse(xml).Descendants().Where(e => e.Name.LocalName == "mi")
            .Select(e => (string?)e.Attribute("mathvariant")));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    [Fact]
    public void PlainOperatorAngleCannotAcquireImplicitIdentifierItalic()
    {
        var (_, xml, token) = Decode("<math><mo>∡</mo></math>");
        Assert.Equal("normal", (string?)token.Attribute("mathvariant"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature("<math><mo>∡</mo></math>"), MathTypeMtefCodec.SemanticSignature(xml));
    }

    [Fact]
    public void OrdinaryAngleRetainsItsRoleInStyledScriptedSource()
    {
        const string source = "<math><mrow><msub><mi mathvariant='bold-italic'>∠</mi><mi>A</mi></msub><mi>B</mi><mi>C</mi></mrow></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var angle = XDocument.Parse(xml).Descendants().Single(e => e.Value == "∠" && !e.HasElements);
        Assert.Equal("mi", angle.Name.LocalName);
        Assert.Equal("bold-italic", (string?)angle.Attribute("mathvariant"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(source.Replace("bold-italic", "bold")));
    }

    [Fact]
    public void MultiCharacterMiAndProseDefaultsAreUnaffected()
    {
        Assert.Equal(MathTypeMtefCodec.SemanticSignature("<math><mi>∡∢</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='normal'>∡∢</mi></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mtext>∡</mtext></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi>∡</mi></math>"));
    }

    private static (MathTypeMtefCodec.RewriteResult Native, string Xml, XElement Token) Decode(string source)
    {
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        return (native, xml, XDocument.Parse(xml).Descendants().Single(e => !e.HasElements));
    }
}

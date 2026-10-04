using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeGreekExplicitCharacterStyleTests
{
    public static IEnumerable<object[]> StyledGreek()
    {
        foreach (var glyph in new[] { "Γ", "Δ", "α", "θ" })
        foreach (var variant in new[] { "normal", "italic", "bold-italic" })
            yield return new object[] { glyph, variant };
    }

    [Theory]
    [MemberData(nameof(StyledGreek))]
    public void ExplicitGreekStyleIsACharacterStyleNotABuiltinPreference(string glyph, string variant)
    {
        var source = $"<math><mi mathvariant='{variant}'>{glyph}</mi></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName == "mi");
        Assert.Equal(glyph, token.Value);
        Assert.Equal(variant, (string?)token.Attribute("mathvariant"));
        Assert.True((int?)token.Attribute("data-mtef-explicit-typeface") < 0);
        Assert.Equal(variant == "normal" ? 0 : variant == "italic" ? 2 : 3,
            (int?)token.Attribute("data-mtef-character-style") ?? 0);
        Assert.Contains("Times New Roman", Encoding.ASCII.GetString(native.Mtef));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Fact]
    public void GreekDefaultsUseMathMlStyleWhileCompatibleNativePreferenceStaysUnchanged()
    {
        var lower = MathTypeMtefCodec.CreateEquationNative("<math><mi>α</mi></math>", false);
        var upper = MathTypeMtefCodec.CreateEquationNative("<math><mi>Γ</mi></math>", false);
        Assert.Equal(new byte[] { 1, 0, 2, 4, 0x84, 0xB1, 3, 0x61, 0, 0 }, lower.Mtef.Skip(lower.StructureOffset));
        var token = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(upper.EquationNative))
            .Descendants().Single(e => e.Name.LocalName == "mi");
        Assert.Equal("italic", (string?)token.Attribute("mathvariant"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature("<math><mi>Γ</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='italic'>Γ</mi></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='normal'>Γ</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='italic'>Γ</mi></math>"));
    }
}

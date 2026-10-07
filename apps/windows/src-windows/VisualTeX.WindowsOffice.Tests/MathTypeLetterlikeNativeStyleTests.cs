using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeLetterlikeNativeStyleTests
{
    [Theory]
    [InlineData("", "italic", 2)]
    [InlineData("normal", "normal", 0)]
    [InlineData("italic", "italic", 2)]
    [InlineData("bold", "bold", 1)]
    [InlineData("bold-italic", "bold-italic", 3)]
    public void NativeLetterlikeGlyphKeepsItsEncodingRoleVariantAndCharacterStyle(string declared, string expected, int bits)
    {
        var source = "<math><msubsup><mi data-mjx-alternate='1'"
            + (declared.Length == 0 ? "" : " mathvariant='" + declared + "'")
            + ">ℏ</mi><mi>i</mi><mn>2</mn></msubsup></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = Assert.Single(XDocument.Parse(actual).Descendants(), e => e.Name.LocalName == "mi" && e.Value == "ℏ");
        Assert.Equal(expected, (string?)token.Attribute("mathvariant"));
        Assert.Equal("1", (string?)token.Attribute("data-mjx-alternate"));
        Assert.Equal(0x210F, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(0x68, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Contains("MT Extra\0", System.Text.Encoding.ASCII.GetString(native.Mtef));
        Assert.Contains("MTExtra\0", System.Text.Encoding.ASCII.GetString(native.Mtef));
        Assert.True((int?)token.Attribute("data-mtef-explicit-typeface") < 0);
        Assert.Equal(bits, (int?)token.Attribute("data-mtef-character-style") ?? 0);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, actual, false).Mtef);
    }
}

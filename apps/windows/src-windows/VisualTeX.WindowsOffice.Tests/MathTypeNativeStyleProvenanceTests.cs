using System.Reflection;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNativeStyleProvenanceTests
{
    [Theory]
    [InlineData(0x0393, 0x47, 5)][InlineData(0x0394, 0x44, 5)][InlineData(0x03B1, 0x61, 4)]
    public void NativeNormalItalicNormalKeepsGlyphIdentityAndRebindsEncodingTogether(int scalar, byte position, byte face)
    {
        var original = RawNativeGreek(scalar, position, face, 0);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(original);
        var token = Token(xml);
        Assert.Equal("normal", (string?)token.Attribute("mathvariant"));
        Assert.Equal((int)face, (int?)token.Attribute("data-mtef-explicit-typeface"));
        Assert.Equal((int)position, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(original.Skip(28), MathTypeMtefCodec.RewriteEquationNative(original, xml, false).Mtef);

        var edited = XDocument.Parse(xml);
        edited.Descendants().Single(e => e.Name.LocalName == "mi").SetAttributeValue("mathvariant", "italic");
        var italic = MathTypeMtefCodec.RewriteEquationNative(original, edited.ToString(), false);
        var italicXml = MathTypeMtefCodec.ReadEquationNativeMathMl(italic.EquationNative);
        token = Token(italicXml);
        Assert.Equal(((char)scalar).ToString(), token.Value);
        Assert.Equal("italic", (string?)token.Attribute("mathvariant"));
        Assert.True((int?)token.Attribute("data-mtef-explicit-typeface") < 0);
        Assert.Null(token.Attribute("data-mtef-encoded8"));
        Assert.Equal(scalar, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(italic.Mtef, MathTypeMtefCodec.RewriteEquationNative(italic.EquationNative, italicXml, false).Mtef);

        edited = XDocument.Parse(italicXml);
        edited.Descendants().Single(e => e.Name.LocalName == "mi").SetAttributeValue("mathvariant", "normal");
        var normal = MathTypeMtefCodec.RewriteEquationNative(italic.EquationNative, edited.ToString(), false);
        var normalXml = MathTypeMtefCodec.ReadEquationNativeMathMl(normal.EquationNative);
        token = Token(normalXml);
        Assert.Equal(((char)scalar).ToString(), token.Value);
        Assert.Equal("normal", (string?)token.Attribute("mathvariant"));
        Assert.Null(token.Attribute("data-mtef-encoded8"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(normalXml), MathTypeMtefCodec.SemanticSignature(italicXml));
        Assert.Equal(normal.Mtef, MathTypeMtefCodec.RewriteEquationNative(normal.EquationNative, normalXml, false).Mtef);
    }

    [Theory]
    [InlineData(0, "normal")][InlineData(1, "bold")][InlineData(2, "italic")][InlineData(3, "bold-italic")]
    public void BuiltInGreekReadbackMaterializesActualPreferenceAndPreservesUnchangedBytes(byte bits, string expected)
    {
        var original = RawNativeGreek(0x0393, 0x47, 5, bits);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(original);
        Assert.Equal(expected, (string?)Token(xml).Attribute("mathvariant"));
        Assert.Equal(original.Skip(28), MathTypeMtefCodec.RewriteEquationNative(original, xml, false).Mtef);
    }

    [Fact]
    public void LiteralSingleMiGammaDefaultCannotMasqueradeAsUprightNativeGamma()
    {
        // Actual production latexToMathMl: literal Unicode Γ versus command \\Gamma.
        const string literal = "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>&#x393;</mi></math>";
        const string upright = "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi mathvariant='normal'>&#x393;</mi></math>";
        const string italic = "<math><mi mathvariant='italic'>Γ</mi></math>";
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(literal), MathTypeMtefCodec.SemanticSignature(italic));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(literal), MathTypeMtefCodec.SemanticSignature(upright));
        foreach (var source in new[] { literal, upright, italic })
        {
            var generated = MathTypeMtefCodec.CreateEquationNative(source, false);
            var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(actual));
        }
    }

    private static XElement Token(string xml) => XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName == "mi");

    private static byte[] RawNativeGreek(int scalar, byte encoded8, byte face, byte bits)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi></math>", false);
        var bytes = seed.Mtef.Take(seed.StructureOffset).Concat(new byte[] {
            1, 0, 2, 4, (byte)(face + 128), (byte)scalar, (byte)(scalar >> 8), encoded8, 0, 0 }).ToArray();
        const BindingFlags flags = BindingFlags.Static | BindingFlags.NonPublic;
        var codec = typeof(MathTypeMtefCodec);
        var layout = codec.GetMethod("ReadPrefixLayout", flags)!.Invoke(null, new object[] { bytes })!;
        var records = (IReadOnlyList<(byte Type, int Start, int End)>)layout.GetType()
            .GetProperty("Records", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(layout)!;
        var cursor = records.Single(r => r.Type == 18).Start + 2;
        var skip = codec.GetMethod("SkipDimensionArray", flags)!;
        cursor = (int)skip.Invoke(null, new object[] { bytes, cursor })!;
        cursor = (int)skip.Invoke(null, new object[] { bytes, cursor })!;
        var count = bytes[cursor++];
        for (var typeface = 1; typeface <= count; typeface++)
        {
            var font = bytes[cursor++];
            Assert.NotEqual(255, font);
            if (font == 0) continue;
            if (typeface == face) bytes[cursor] = bits;
            cursor++;
        }
        var native = seed.EquationNative.Take(28).Concat(bytes).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)bytes.Length), 0, native, 8, 4);
        return native;
    }
}

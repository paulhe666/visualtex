using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNativeGlyphAuthorityTests
{
    public static IEnumerable<object[]> Corners()
    {
        foreach (var item in new[] { (0x231C, 0x84), (0x231D, 0x86), (0x231E, 0x85), (0x231F, 0x87) })
        foreach (var variant in new[] { "normal", "italic", "bold-italic" })
            yield return new object[] { item.Item1, item.Item2, variant };
    }

    [Theory]
    [MemberData(nameof(Corners))]
    public void CornersKeepTheirOperatorIdentityAndNativeFont(int scalar, int position, string variant)
    {
        var source = $"<math><mo mathvariant='{variant}'>{(char)scalar}</mo></math>";
        var (native, xml, token) = RoundTrip(source);
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal(((char)scalar).ToString(), token.Value);
        Assert.Equal(scalar, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(position, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(variant == "normal" ? null : variant, (string?)token.Attribute("mathvariant"));
        Assert.Contains("EuclidMath1\0", Encoding.ASCII.GetString(native.Mtef));
        Assert.Contains("Euclid Math One\0", Encoding.ASCII.GetString(native.Mtef));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("italic")]
    [InlineData("bold-italic")]
    public void ShortParallelKeepsTheVendorCodeAlternateIdentityAndStyle(string variant)
    {
        var source = $"<math><mo data-mjx-alternate='1' mathvariant='{variant}'>∥</mo></math>";
        var (_, _, token) = RoundTrip(source);
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal("∥", token.Value);
        Assert.Equal("1", (string?)token.Attribute("data-mjx-alternate"));
        Assert.Equal(0xE92F, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(0x89, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(variant == "normal" ? null : variant, (string?)token.Attribute("mathvariant"));
    }

    [Fact]
    public void OrdinaryAndShortParallelDoNotBecomeTheSameNativeGlyph()
    {
        var plain = MathTypeMtefCodec.CreateEquationNative("<math><mo>∥</mo></math>", false);
        var shortForm = MathTypeMtefCodec.CreateEquationNative("<math><mo data-mjx-alternate='1'>∥</mo></math>", false);
        Assert.NotEqual(plain.Mtef, shortForm.Mtef);
        Assert.DoesNotContain("EuclidMath1", Encoding.ASCII.GetString(plain.Mtef));
        Assert.Contains("EuclidMath1", Encoding.ASCII.GetString(shortForm.Mtef));
    }

    [Theory]
    [InlineData(0x10, 0)]
    [InlineData(0x30, 0)]
    [InlineData(0x10, 2)]
    [InlineData(0x30, 3)]
    public void KnownGlyphRetainsEncoded16AndMtCodePresence(int options, int style)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi mathvariant='script'>A</mi></math>", false);
        var defs = seed.Mtef.Skip(seed.StructureOffset + 2).Take(seed.Mtef.Length - seed.StructureOffset - 10).ToArray();
        Assert.Equal(new byte[] { 8, 6, 0 }, defs.Skip(defs.Length - 3));
        defs[defs.Length - 1] = (byte)style;
        var character = new List<byte> { 2, (byte)options, 0x7F };
        if ((options & 0x20) == 0) character.AddRange(new byte[] { 0x2F, 0xE9 });
        character.AddRange(new byte[] { 0x89, 0 });
        var mtef = seed.Mtef.Take(seed.StructureOffset).Concat(defs).Concat(new byte[] { 1, 0 })
            .Concat(character).Concat(new byte[] { 0, 0 }).ToArray();
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)mtef.Length), 0, native, 8, 4);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName == "mo");
        Assert.Equal("∥", token.Value);
        Assert.Equal("1", (string?)token.Attribute("data-mjx-alternate"));
        Assert.Equal(0x89, (int?)token.Attribute("data-mtef-encoded16"));
        Assert.Null(token.Attribute("data-mtef-encoded8"));
        Assert.Equal((options & 0x20) != 0, token.Attribute("data-mtef-no-mtcode") != null);
        Assert.Equal(mtef, MathTypeMtefCodec.RewriteEquationNative(native, xml, false).Mtef);
    }

    [Fact]
    public void UnverifiedDoubleSolidusDoesNotReuseTheParallelAuthority()
    {
        var native = MathTypeMtefCodec.CreateEquationNative("<math><mo>⫽</mo></math>", false);
        Assert.DoesNotContain("EuclidMath1", Encoding.ASCII.GetString(native.Mtef));
    }

    private static (MathTypeMtefCodec.RewriteResult Native, string Xml, XElement Token) RoundTrip(string source)
    {
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
        return (native, xml, token);
    }
}

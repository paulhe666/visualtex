using System.Reflection;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeBuiltInGlyphAuthorityTests
{
    public static IEnumerable<object[]> NativeAngles()
    {
        foreach (var scalar in new[] { 0x2220, 0x2221, 0x2222 })
        foreach (var style in new[] { 0, 1, 2, 3 })
            yield return new object[] { scalar, style };
    }

    [Theory]
    [MemberData(nameof(NativeAngles))]
    public void BuiltInAngleUsesTheSameGlyphAndActualStyleAuthority(int scalar, int style)
    {
        var native = RawNative(scalar, style, 0x04);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var token = XDocument.Parse(xml).Descendants().Single(e => !e.HasElements);
        Assert.Equal("mi", token.Name.LocalName);
        Assert.Equal(((char)scalar).ToString(), token.Value);
        Assert.Equal(new[] { "normal", "bold", "italic", "bold-italic" }[style], (string?)token.Attribute("mathvariant"));
        Assert.Equal(scalar == 0x2220 ? 0xD0 : scalar == 0x2221 ? 0x52 : 0x53, (int?)token.Attribute("data-mtef-encoded8"));
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native, xml, false);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(xml),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
    }

    [Theory]
    [InlineData(0x10)][InlineData(0x30)]
    public void BuiltInKnownGlyphDoesNotLoseWordEncodingOrAbsentMtCode(int options)
    {
        var native = RawNative(0x2221, 0, options);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native, xml, false);
        foreach (var text in new[] { xml, MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative) })
        {
            var token = XDocument.Parse(text).Descendants().Single(e => !e.HasElements);
            Assert.Equal("∡", token.Value);
            Assert.Equal("normal", (string?)token.Attribute("mathvariant"));
            Assert.Equal(0x52, (int?)token.Attribute("data-mtef-encoded16"));
            Assert.Null(token.Attribute("data-mtef-encoded8"));
            Assert.Equal(options == 0x30, token.Attribute("data-mtef-no-mtcode") != null);
        }
    }

    [Fact]
    public void ATypefaceNumberDoesNotOverrideTheActualFontAssignment()
    {
        var native = RawNative(0x2221, 0, 0x30, fontOverride: 1);
        var token = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native))
            .Descendants().Single(e => !e.HasElements);
        // Without MTCode, position52 is R in Times, not measuredangle.
        Assert.Equal("R", token.Value);
        Assert.Equal("normal", (string?)token.Attribute("mathvariant"));
        Assert.NotEqual("∡", token.Value);
    }

    private static byte[] RawNative(int scalar, int style, int options, byte? fontOverride = null)
    {
        // Reuse the verified Symbol D0 and M0001/M0002 MT Extra CHAR contracts
        // with their existing production preferences instead of inventing a font.
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mo>∥</mo></math>", false);
        var mtef = seed.Mtef.Take(seed.StructureOffset).ToList();
        var face = scalar == 0x2220 ? 6 : 11;
        var character = new List<byte> { 2, (byte)options, (byte)(face + 128) };
        if ((options & 0x20) == 0) character.AddRange(new[] { (byte)scalar, (byte)(scalar >> 8) });
        character.Add((byte)(scalar == 0x2220 ? 0xD0 : scalar == 0x2221 ? 0x52 : 0x53));
        if ((options & 0x10) != 0) character.Add(0);
        mtef.AddRange(new byte[] { 1, 0 }); mtef.AddRange(character); mtef.AddRange(new byte[] { 0, 0 });
        var bytes = mtef.ToArray();
        const BindingFlags flags = BindingFlags.Static | BindingFlags.NonPublic;
        var codec = typeof(MathTypeMtefCodec);
        var layout = codec.GetMethod("ReadPrefixLayout", flags)!.Invoke(null, new object[] { bytes })!;
        var records = (IReadOnlyList<(byte Type, int Start, int End)>)layout.GetType()
            .GetProperty("Records", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(layout)!;
        var cursor = records.Single(r => r.Type == 18).Start + 2;
        var skip = codec.GetMethod("SkipDimensionArray", flags)!;
        cursor = (int)skip.Invoke(null, new object[] { bytes, cursor })!;
        cursor = (int)skip.Invoke(null, new object[] { bytes, cursor })!;
        Assert.True(bytes[cursor++] >= 11);
        for (var typeface = 1; typeface <= 11; typeface++)
        {
            var fontOffset = cursor;
            var font = bytes[cursor++];
            Assert.NotEqual(255, font);
            if (font == 0) continue;
            if (typeface == face)
            {
                if (fontOverride.HasValue) bytes[fontOffset] = fontOverride.Value;
                bytes[cursor] = (byte)style;
            }
            cursor++;
        }
        var native = seed.EquationNative.Take(28).Concat(bytes).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)bytes.Length), 0, native, 8, 4);
        return native;
    }
}

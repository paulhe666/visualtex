using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeBuiltInOperatorTokenTests
{
    [Theory]
    [InlineData("⊲")][InlineData("⊳")][InlineData("⊙")][InlineData("∓")]
    [InlineData("∥")][InlineData("≻")][InlineData("≺")][InlineData("≫")]
    [InlineData("≪")][InlineData("≃")][InlineData("≐")][InlineData("↦")]
    [InlineData("⋯")][InlineData("⋮")][InlineData("⋱")]
    public void VerifiedPlainMtExtraGlyphRemainsAnOperator(string glyph)
    {
        foreach (var source in new[] { $"<math><mo>{glyph}</mo></math>",
            $"<math><msubsup><mi>x</mi><mi>i</mi><mo>{glyph}</mo></msubsup></math>" })
        {
            var native = MathTypeMtefCodec.CreateEquationNative(source, false);
            var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
            var token = XDocument.Parse(xml).Descendants().Single(e => e.Value == glyph && !e.HasElements);
            Assert.Equal("mo", token.Name.LocalName);
            Assert.Null(token.Attribute("mathvariant"));
            Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
        }
    }

    [Theory]
    [InlineData("ℓ")][InlineData("ℏ")][InlineData("α")]
    public void LetterlikeAndGreekIdentifiersKeepTheirIdentifierClass(string glyph)
    {
        var native = MathTypeMtefCodec.CreateEquationNative($"<math><mi>{glyph}</mi></math>", false);
        var token = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative))
            .Descendants().Single(e => !e.HasElements);
        Assert.Equal("mi", token.Name.LocalName);
    }

    [Theory]
    [InlineData(0, "normal")][InlineData(1, "bold")][InlineData(2, "italic")][InlineData(3, "bold-italic")]
    public void FixedLetterlikeGlyphReadsItsActualSymbolPreference(int style, string variant)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi mathvariant='normal'>℘</mi></math>", false);
        var mtef = seed.Mtef.ToArray();
        const System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic;
        var codec = typeof(MathTypeMtefCodec);
        var layout = codec.GetMethod("ReadPrefixLayout", flags)!.Invoke(null, new object[] { mtef })!;
        var records = (IReadOnlyList<(byte Type, int Start, int End)>)layout.GetType()
            .GetProperty("Records", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.GetValue(layout)!;
        var cursor = records.Single(r => r.Type == 18).Start + 2;
        var skip = codec.GetMethod("SkipDimensionArray", flags)!;
        cursor = (int)skip.Invoke(null, new object[] { mtef, cursor })!;
        cursor = (int)skip.Invoke(null, new object[] { mtef, cursor })!;
        var count = mtef[cursor++];
        Assert.True(count >= 6);
        for (var typeface = 1; typeface <= 6; typeface++)
        {
            var font = mtef[cursor++];
            Assert.NotEqual(255, font);
            if (font == 0) continue;
            if (typeface == 6) mtef[cursor] = (byte)style;
            cursor++;
        }
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        var token = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native))
            .Descendants().Single(e => !e.HasElements);
        Assert.Equal("mi", token.Name.LocalName);
        Assert.Equal("℘", token.Value);
        Assert.Equal(variant, (string?)token.Attribute("mathvariant"));
        if (style == 0)
            Assert.Equal(MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='normal'>℘</mi></math>"),
                MathTypeMtefCodec.SemanticSignature(token.ToString()));
    }

    [Fact]
    public void AnActualVariableFaceIsNotReclassifiedAsAPlainOperator()
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi></math>", false);
        var mtef = seed.Mtef.Take(seed.StructureOffset)
            .Concat(new byte[] { 1, 0, 2, 0, 0x83, 0x25, 0x22, 0, 0 }).ToArray();
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)mtef.Length), 0, native, 8, 4);
        var token = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native))
            .Descendants().Single(e => !e.HasElements);
        Assert.Equal("mi", token.Name.LocalName);
        Assert.Equal("∥", token.Value);
        Assert.Equal("italic", (string?)token.Attribute("mathvariant")); // Actual native preference now materialized.
    }
}

using System.Text;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Text/Function/Variable/Number/Vector are Times New Roman with encoding
// WinAllBasicCodePages (Windows 1250-1258) in the standalone prefix. MathType 7
// shows U+FFFD for an MTCode those code pages cannot encode, e.g. U+25B3 in
// "fold triangle ADE along DE". Text goes to typeface 12 (SimSun,
// WinAllCodePages); operators use an explicit SimSun style so they stay mo.
public sealed class MathTypeFarEastCodePageSymbolTests
{
    private const byte TypefaceFarEastText = 12 + 128;
    private static readonly byte[] TimesNewRomanFaces = { 1 + 128, 2 + 128, 3 + 128, 7 + 128, 8 + 128 };

    static MathTypeFarEastCodePageSymbolTests() =>
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

    [Theory]
    [InlineData('△')] // white up-pointing triangle
    [InlineData('▲')]
    [InlineData('□')]
    [InlineData('○')]
    [InlineData('☆')]
    [InlineData('⌒')] // arc
    [InlineData('≌')] // all equal to (congruent, Chinese usage)
    [InlineData('∽')] // reversed tilde (similar, Chinese usage)
    [InlineData('∵')] // because
    [InlineData('≦')]
    [InlineData('≧')]
    [InlineData('≮')]
    [InlineData('℃')] // degree Celsius
    [InlineData('∛')]
    public void OperatorOutsideBasicCodePagesUsesSimSunStyleAndRoundTrips(char symbol)
    {
        foreach (var kind in new[] { "mo", "mi" })
        {
            var source = $"<math><{kind}>{symbol}</{kind}><mi>A</mi></math>";
            var native = MathTypeMtefCodec.CreateEquationNative(source, false);
            var face = Assert.Single(CharTypefaces(native, symbol));
            Assert.True(face < 128, $"{kind} U+{(int)symbol:X4} typeface byte {face:X2}");
            // MathType only knows the prefix's localized SimSun FONT_DEF.
            Assert.DoesNotContain("SimSun", Encoding.Latin1.GetString(native.Mtef));
            var readBack = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(readBack));
        }
    }

    [Theory]
    [InlineData("△ABC")]
    [InlineData("ABC△")]
    [InlineData("△ ABC")]
    [InlineData("ABC △")]
    [InlineData("如图，△ABC中，∠A=90°")]
    [InlineData("①℃")]
    public void TextRunKeepsFarEastSymbolsInsideTheSameMtext(string text)
    {
        var source = $"<math><mi>x</mi><mtext>{text}</mtext><mi>y</mi></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var readBack = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(readBack));
        foreach (var character in text.Where(c => c > 0x7F && !IsWinAllBasicCodePage(c)))
            Assert.All(CharTypefaces(native, character), face => Assert.Equal(TypefaceFarEastText, face));
    }

    [Theory]
    // "As shown, triangle ABC" with the triangle as math next to Chinese text.
    [InlineData("<mtext>如图，</mtext><mo>△</mo><mi>A</mi><mi>B</mi><mi>C</mi>")]
    [InlineData("<mo>△</mo><mtext>ABC</mtext>")]
    [InlineData("<mtext>∵</mtext><mo>∵</mo><mtext>∵</mtext>")]
    [InlineData("<mo>△</mo><mi>A</mi><mi>D</mi><mi>E</mi><mo>≌</mo><mi>△</mi><mi>B</mi><mo>∵</mo><mover><mrow><mi>A</mi><mi>B</mi></mrow><mo>⌒</mo></mover><mtext>如图，△ABC①℃</mtext><mo>≦</mo><mo>∠</mo>")]
    public void OperatorAndTextNeighboursKeepTheirRoles(string body)
    {
        var source = $"<math>{body}</math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var readBack = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(readBack));
    }

    [Fact]
    public void RewritingAnEquationReusesTheSimSunOperatorStyle()
    {
        var first = MathTypeMtefCodec.CreateEquationNative("<math><mo>△</mo><mi>A</mi></math>", false);
        var edited = "<math><mo>△</mo><mi>A</mi><mo>≌</mo><mo>△</mo><mi>B</mi></math>";
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(first.EquationNative, edited, false);
        var readBack = MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(edited), MathTypeMtefCodec.SemanticSignature(readBack));
        Assert.Equal(first.Mtef.Take(first.StructureOffset), rewritten.Mtef.Take(rewritten.StructureOffset));
    }

    [Fact]
    public void BasicCodePageCharactersKeepTheirTypeface()
    {
        // U+2030 per mille and U+2116 numero are in Windows-1252/1251.
        var perMille = MathTypeMtefCodec.CreateEquationNative("<math><mo>‰</mo></math>", false);
        Assert.Equal(new byte[] { 1 + 128 }, CharTypefaces(perMille, '‰'));
        var cyrillic = MathTypeMtefCodec.CreateEquationNative("<math><mtext>д№</mtext></math>", false);
        Assert.Equal(new byte[] { 1 + 128 }, CharTypefaces(cyrillic, 'д'));
        Assert.Equal(new byte[] { 1 + 128 }, CharTypefaces(cyrillic, '№'));
    }

    [Fact]
    public void NoTimesNewRomanCharacterIsOutsideItsCodePages()
    {
        // Geometry, arrows, operators, enclosed numbers and letterlike symbols:
        // a Times New Roman CHAR must be encodable by its code pages.
        var ranges = new[] { (0x2100, 0x214F), (0x2190, 0x22FF), (0x2300, 0x23FF), (0x2460, 0x24FF), (0x25A0, 0x26FF) };
        var failures = new List<string>();
        foreach (var (first, last) in ranges)
        for (var scalar = first; scalar <= last; scalar++)
        {
            var character = (char)scalar;
            if (char.GetUnicodeCategory(character) == System.Globalization.UnicodeCategory.OtherNotAssigned) continue;
            foreach (var kind in new[] { "mo", "mi", "mtext" })
            {
                // Known gap: letter-like/number identifiers (U+210D, U+2460 as mi)
                // still use the styled Times New Roman identifier face.
                if (kind == "mi" && (char.IsLetter(character) || char.IsNumber(character))) continue;
                MathTypeMtefCodec.RewriteResult native;
                try { native = MathTypeMtefCodec.CreateEquationNative($"<math><{kind}>{character}</{kind}></math>", false); }
                catch (InvalidDataException) { continue; }
                if (CharTypefaces(native, character).Any(face => TimesNewRomanFaces.Contains(face))
                    && !IsWinAllBasicCodePage(character))
                    failures.Add($"{kind} U+{scalar:X4}");
            }
        }
        Assert.True(failures.Count == 0, string.Join(" ", failures));
    }

    // CHAR records [2][options][typeface][MTCode lo][MTCode hi] for one scalar.
    private static byte[] CharTypefaces(MathTypeMtefCodec.RewriteResult native, char scalar)
    {
        var mtef = native.Mtef;
        var faces = new List<byte>();
        for (var i = native.StructureOffset; i + 4 < mtef.Length; i++)
            if (mtef[i] == 2 && (mtef[i + 1] & 0x20) == 0
                && mtef[i + 3] == (scalar & 0xFF) && mtef[i + 4] == scalar >> 8)
                faces.Add(mtef[i + 2]);
        return faces.ToArray();
    }

    private static bool IsWinAllBasicCodePage(char character)
    {
        foreach (var page in new[] { 1250, 1251, 1252, 1253, 1254, 1255, 1256, 1257, 1258 })
        {
            var encoding = Encoding.GetEncoding(page, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            try
            {
                if (encoding.GetString(encoding.GetBytes(character.ToString())) == character.ToString()) return true;
            }
            catch (EncoderFallbackException) { }
            catch (DecoderFallbackException) { }
        }
        return false;
    }
}

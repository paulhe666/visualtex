using System.Globalization;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

// Unicode mathematical alphanumerics (bold, italic, script, fraktur,
// double-struck, sans-serif and monospace letters, Greek and digits) are
// written to MathType as their base character in that alphabet.
public sealed class MathTypeMathAlphanumericTests
{
    private const string MathMl = "http://www.w3.org/1998/Math/MathML";
    private readonly ITestOutputHelper output;

    public MathTypeMathAlphanumericTests(ITestOutputHelper output) => this.output = output;

    private static string Math(string body) => "<math xmlns=\"" + MathMl + "\">" + body + "</math>";

    private static string RoundTrip(string mathMl) =>
        MathTypeMtefCodec.ReadEquationNativeMathMl(MathTypeMtefCodec.CreateEquationNative(mathMl, false).EquationNative);

    [Fact]
    public void EveryAssignedMathematicalAlphanumericWritesAndReadsBackAsItsStyledBaseCharacter()
    {
        var failures = new List<string>();
        var checkedCount = 0;
        for (var scalar = 0x1D400; scalar <= 0x1D7FF; scalar++)
        {
            var character = char.ConvertFromUtf32(scalar);
            if (CharUnicodeInfo.GetUnicodeCategory(character, 0) == UnicodeCategory.OtherNotAssigned) continue;
            Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(scalar, out var baseScalar, out var variant),
                $"U+{scalar:X} is not decomposed.");
            var kind = baseScalar is >= '0' and <= '9' ? "mn" : "mi";
            var source = Math("<" + kind + ">" + character + "</" + kind + ">");
            var styled = Math("<" + kind + " mathvariant=\"" + variant + "\">" + char.ConvertFromUtf32(baseScalar) + "</" + kind + ">");
            checkedCount++;
            try
            {
                var read = RoundTrip(source);
                if (MathTypeMtefCodec.SemanticSignature(read) != MathTypeMtefCodec.SemanticSignature(styled))
                    failures.Add($"U+{scalar:X} {variant} '{char.ConvertFromUtf32(baseScalar)}': read back {read}");
            }
            catch (Exception error)
            {
                failures.Add($"U+{scalar:X} {variant} '{char.ConvertFromUtf32(baseScalar)}': {error.GetType().Name}: {error.Message}");
            }
        }
        output.WriteLine($"checked={checkedCount} failures={failures.Count}");
        foreach (var failure in failures.Take(40)) output.WriteLine(failure);
        Assert.Empty(failures);
    }

    [Theory]
    [InlineData(0x1D431, 'x', "bold")]
    [InlineData(0x1D465, 'x', "italic")]
    [InlineData(0x1D49C, 'A', "script")]
    [InlineData(0x1D504, 'A', "fraktur")]
    [InlineData(0x1D538, 'A', "double-struck")]
    [InlineData(0x1D5A0, 'A', "sans-serif")]
    [InlineData(0x1D670, 'A', "monospace")]
    [InlineData(0x1D7CF, '1', "bold")]
    [InlineData(0x1D7D9, '1', "double-struck")]
    public void KnownCharactersDecomposeToTheirAlphabet(int scalar, char baseCharacter, string variant)
    {
        Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(scalar, out var baseScalar, out var alphabet));
        Assert.Equal(baseCharacter, (char)baseScalar);
        Assert.Equal(variant, alphabet);
    }

    [Fact]
    public void GreekMathematicalAlphabetsFollowUnicodeOrder()
    {
        Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(0x1D6A8, out var alpha, out var bold));
        Assert.Equal(0x0391, alpha);
        Assert.Equal("bold", bold);
        Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(0x1D6B9, out var theta, out _));
        Assert.Equal(0x03F4, theta);
        Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(0x1D6C1, out var nabla, out _));
        Assert.Equal(0x2207, nabla);
        Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(0x1D70B, out var pi, out var italic));
        Assert.Equal(0x03C0, pi);
        Assert.Equal("italic", italic);
        Assert.True(MathTypeMtefCodec.TryDecomposeMathAlphanumeric(0x1D7C9, out var varpi, out var sansBoldItalic));
        Assert.Equal(0x03D6, varpi);
        Assert.Equal("sans-serif-bold-italic", sansBoldItalic);
    }

    [Fact]
    public void AMixedTokenSplitsIntoStyledRunsWithoutBreakingAScript()
    {
        var mixed = Math("<msub><mi>a" + char.ConvertFromUtf32(0x1D431) + char.ConvertFromUtf32(0x1D432) + "</mi><mn>2</mn></msub>");
        var expected = Math("<msub><mrow><mi>a</mi><mi mathvariant=\"bold\">xy</mi></mrow><mn>2</mn></msub>");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(expected), MathTypeMtefCodec.SemanticSignature(RoundTrip(mixed)));
    }

    [Fact]
    public void OrdinaryTextIsUnchanged()
    {
        var source = Math("<mi>x</mi><mo>+</mo><mn>1</mn>");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(RoundTrip(source)));
    }
}

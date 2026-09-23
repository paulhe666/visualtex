using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeRedrawCanonicalizationTests
{
    private static string RoundTrip(string source)
    {
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, true, 12);
        return MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
    }

    private static void AssertRoundTrip(string source)
    {
        var actual = RoundTrip(source);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(actual));
        var original = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, true, 12);
        foreach (var editSource in new[] { source, actual })
        {
            var rewritten = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(
                original.EquationNative, editSource, true, 12);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),
                MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
        }
    }

    [Theory]
    [InlineData("|", "|")]
    [InlineData("‖", "‖")]
    [InlineData("(", ")")]
    [InlineData("[", "]")]
    [InlineData("⟨", "⟩")]
    [InlineData("⌈", "⌉")]
    [InlineData("⌊", "⌋")]
    public void FixedDelimiterRolesDoNotDependOnMathJaxAnnotationsOrRowGrouping(string open, string close)
    {
        var first = $"<mo data-mjx-texclass='OPEN' fence='false' stretchy='false'>{open}</mo>";
        var last = $"<mo data-mjx-texclass='CLOSE' fence='false' stretchy='false'>{close}</mo>";
        const string body = "<msub><mi>x</mi><mn>8</mn></msub><mo>−</mo><mi>μ</mi>";
        foreach (var script in new[] { "", "msub", "msup", "msubsup" })
        {
            var closeWithScript = script.Length == 0 ? last
                : $"<{script}>{last}<mn>2</mn>" + (script == "msubsup" ? "<mn>3</mn>" : "") + $"</{script}>";
            var tokens = first + body + closeWithScript;
            foreach (var content in new[] { tokens, "<mrow>" + tokens + "</mrow>",
                         "<mi>a</mi><mo>+</mo>" + tokens + "<mo>&lt;</mo><mi>ε</mi>",
                         "<mtable><mtr><mtd>" + tokens + "</mtd></mtr></mtable>" })
            {
                var source = "<math>" + content + "</math>";
                AssertRoundTrip(source);
                Assert.DoesNotContain(XElement.Parse(RoundTrip(source)).Descendants(),
                    e => e.Name.LocalName == "mfenced");
            }
        }
    }

    [Fact]
    public void FixedParenthesesAfterBigOperatorSurviveNativeMainSlotGrouping()
    {
        AssertRoundTrip("<math><munderover><mo>∏</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow>"
            + "<mn>14</mn></munderover><mo stretchy='false'>(</mo><mn>1</mn><mo>+</mo>"
            + "<msub><mi>a</mi><mi>k</mi></msub><mo stretchy='false'>)</mo></math>");
    }

    [Fact]
    public void FixedParenthesesInsideScalableFenceKeepTheSameSignature()
    {
        AssertRoundTrip("<math><msubsup><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>[</mo>"
            + "<mi>f</mi><mo stretchy='false'>(</mo><mi>x</mi><mo stretchy='false'>)</mo>"
            + "<mi>g</mi><mo stretchy='false'>(</mo><mi>x</mi><mo stretchy='false'>)</mo>"
            + "<mo data-mjx-texclass='CLOSE'>]</mo></mrow><mi>a</mi><mi>b</mi></msubsup></math>");
    }

    [Theory]
    [InlineData("^")]
    [InlineData("→")]
    [InlineData("←")]
    [InlineData("↔")]
    [InlineData("~")]
    [InlineData("˙")]
    [InlineData("¨")]
    public void OmittedAccentUsesTheCoreOperatorDefault(string mark)
    {
        var implicitAccent = $"<math><msub><mrow><mover><mi>x</mi><mo stretchy='false'>{mark}</mo></mover>"
            + "</mrow><mn>23</mn></msub></math>";
        var explicitAccent = implicitAccent.Replace("<mover>", "<mover accent='true'>");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(explicitAccent),
            MathTypeMtefCodec.SemanticSignature(implicitAccent));
        AssertRoundTrip(implicitAccent);
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(implicitAccent),
            MathTypeMtefCodec.SemanticSignature(implicitAccent.Replace("<mover>", "<mover accent='false'>")));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(implicitAccent),
            MathTypeMtefCodec.SemanticSignature(implicitAccent.Replace("<mover>", "<msup>").Replace("</mover>", "</msup>")));
    }

    [Fact]
    public void ExplicitAccentFalseOverridesCoreOperatorTrue()
    {
        const string limit = "<math><mover accent='false'><mi>x</mi><mo accent='true'>^</mo></mover></math>";
        const string accent = "<math><mover accent='true'><mi>x</mi><mo>^</mo></mover></math>";
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(limit), MathTypeMtefCodec.SemanticSignature(accent));
    }

    [Fact]
    public void FixedNormUsesTheNativeMtExtraGlyphRatherThanUnicodeTextFallback()
    {
        const string source = "<math><mo fence='false' stretchy='false'>‖</mo><mi>x</mi>"
            + "<mo fence='false' stretchy='false'>‖</mo></math>";
        var generated = MathTypeMtefCodec.CreateEquationNative(source, true);
        var records = generated.Mtef.Skip(generated.StructureOffset).ToArray();
        // Independently checked against MathType 7's actual rendered output:
        // CHAR + encoded8 flag + fnMTEXTRA + U+2016 + MT Extra position 0x50.
        byte[] glyph = { 2, 4, 139, 0x16, 0x20, 0x50 };
        Assert.Equal(2, Enumerable.Range(0, records.Length - glyph.Length + 1)
            .Count(i => records.Skip(i).Take(glyph.Length).SequenceEqual(glyph)));
        AssertRoundTrip(source);
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature("<math><mtext>‖</mtext><mi>x</mi><mtext>‖</mtext></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(source.Replace("‖", "|")));
    }

    [Theory]
    [InlineData("<mfrac><mi>a</mi><mi>b</mi></mfrac>", "<mfrac><mi>b</mi><mi>a</mi></mfrac>")]
    [InlineData("<msub><mi>x</mi><mn>2</mn></msub>", "<msup><mi>x</mi><mn>2</mn></msup>")]
    [InlineData("<mi>x</mi><mo>+</mo><mi>y</mi>", "<mi>x</mi><mo>−</mo><mi>y</mi>")]
    [InlineData("<msup><mi>x</mi><mn>2</mn></msup>", "<msup><mi>x</mi><mn>3</mn></msup>")]
    public void GenuineMathematicalChangesRemainRejected(string expected, string corrupted)
    {
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math>" + expected + "</math>"),
            MathTypeMtefCodec.SemanticSignature("<math>" + corrupted + "</math>"));
    }

    [Fact]
    public void MissingOrDuplicatedFixedDelimitersCannotPassValidation()
    {
        const string source = "<math><mo data-mjx-texclass='OPEN' fence='false' stretchy='false'>|</mo><mi>x</mi>"
            + "<mo data-mjx-texclass='CLOSE' fence='false' stretchy='false'>|</mo></math>";
        var expected = MathTypeMtefCodec.SemanticSignature(source);
        var actual = XElement.Parse(RoundTrip(source));
        actual.Elements().Last().Remove();
        Assert.NotEqual(expected, MathTypeMtefCodec.SemanticSignature(actual.ToString()));
        actual = XElement.Parse(source);
        actual.AddFirst(new XElement(actual.Elements().First()));
        actual.Add(new XElement(actual.Elements().Last()));
        Assert.NotEqual(expected, MathTypeMtefCodec.SemanticSignature(actual.ToString()));
        Assert.NotEqual(expected, MathTypeMtefCodec.SemanticSignature(source.Replace(">x<", ">y<")));
    }
}

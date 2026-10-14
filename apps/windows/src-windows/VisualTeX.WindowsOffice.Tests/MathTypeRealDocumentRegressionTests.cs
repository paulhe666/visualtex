using System.Xml.Linq;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeRealDocumentRegressionTests
{
    private static string RoundTrip(string source)
    {
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, true, 12);
        return MathTypeMtefCodec.ReadEquationNativeMathMl(generated.EquationNative);
    }

    private static void AssertRoundTrip(string source)
    {
        var expected = MathTypeMtefCodec.SemanticSignature(source);
        var actual = RoundTrip(source);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(actual));
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, true, 12);
        var rewritten = MathTypeMtefCodec.RewriteEquationNativeAtFontSize(
            generated.EquationNative, actual, true, 18);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(
            MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
    }

    public static IEnumerable<object[]> EmptyPrescriptBases()
    {
        foreach (var basis in new[] { "<mi/>", "<mrow/>", "<mrow><mi/></mrow>", "<mstyle><mrow/></mstyle>" })
        foreach (var script in new[] { "msub", "msup", "msubsup" })
        foreach (var prefix in new[] { "", "<mi>x</mi><mo>+</mo>" })
        {
            var pair = script == "msubsup";
            yield return new object[] { "<math>" + prefix + "<" + script + ">" + basis
                + "<mn>2</mn>" + (pair ? "<mn>4</mn>" : "")
                + "</" + script + "><mi>H</mi><mi>e</mi></math>" };
        }
    }

    public static IEnumerable<object[]> EmptyNonPrescriptBases()
    {
        foreach (var basis in new[] { "<mi/>", "<mrow/>", "<mrow><mi/></mrow>", "<mstyle><mrow/></mstyle>" })
        foreach (var script in new[] { "munder", "mover", "munderover", "mmultiscripts" })
        foreach (var prefix in new[] { "", "<mi>x</mi><mo>+</mo>" })
        {
            var pair = script is "munderover" or "mmultiscripts";
            yield return new object[] { "<math>" + prefix + "<" + script + ">" + basis
                + "<mn>2</mn>" + (pair ? "<mn>4</mn>" : "")
                + "</" + script + "><mi>H</mi><mi>e</mi></math>" };
        }
    }

    [Theory]
    [MemberData(nameof(EmptyPrescriptBases))]
    public void EmptyScriptBaseFollowedByAtomUsesNativePrescript(string source)
    {
        AssertRoundTrip(source);
        var actual = XElement.Parse(RoundTrip(source));
        var script = Assert.Single(actual.Descendants()
            .Where(element => element.Name.LocalName == "mmultiscripts"));
        Assert.Equal("H", script.Elements().First().Value);
        Assert.Contains(script.Elements(), element => element.Name.LocalName == "mprescripts");
        Assert.DoesNotContain(actual.Descendants(), element =>
            element.Name.LocalName is "msub" or "msup" or "msubsup"
            && string.IsNullOrWhiteSpace(element.Elements().FirstOrDefault()?.Value));
    }

    [Theory]
    [MemberData(nameof(EmptyNonPrescriptBases))]
    public void EmptyNonPrescriptBaseIsStillIndependentOfThePreviousSibling(string source)
    {
        AssertRoundTrip(source);
        var actual = XElement.Parse(RoundTrip(source));
        // Under/over limits are MathType's limit template, whose empty main
        // slot cannot take a neighbour either.
        var script = actual.Descendants().First(element =>
            element.Name.LocalName is "msub" or "msup" or "msubsup" or "munder" or "mover" or "munderover");
        Assert.Empty(script.Elements().First().Value);
    }

    [Fact]
    public void EmptyScriptWithoutFollowingAtomRemainsIndependentOfPreviousSibling()
    {
        const string source = "<math><mi>x</mi><mo>+</mo><msubsup><mrow/><mn>2</mn><mn>4</mn></msubsup><mo>+</mo></math>";
        AssertRoundTrip(source);
        var actual = XElement.Parse(RoundTrip(source));
        var script = Assert.Single(actual.Descendants()
            .Where(element => element.Name.LocalName == "msubsup"));
        Assert.Empty(script.Elements().First().Value);
        Assert.Empty(actual.Descendants()
            .Where(element => element.Name.LocalName == "mmultiscripts"));
    }

    [Fact]
    public void NuclearReactionUsesNativePrescriptWithoutChangingExistingPostScripts()
    {
        const string source = "<math><msubsup><mi/><mrow><mn>27</mn></mrow><mrow><mn>59</mn></mrow></msubsup>"
            + "<mi>C</mi><mi>o</mi><msubsup><mo>+</mo><mrow><mn>0</mn></mrow><mrow><mn>1</mn></mrow></msubsup>"
            + "<mi>n</mi><msubsup><mo stretchy='false'>→</mo><mrow><mn>27</mn></mrow><mrow><mn>60</mn></mrow></msubsup>"
            + "<mi>C</mi><mi>o</mi><mo>+</mo><mi>γ</mi></math>";
        AssertRoundTrip(source);
        var actual = XElement.Parse(RoundTrip(source));
        var prescript = Assert.Single(actual.Descendants()
            .Where(element => element.Name.LocalName == "mmultiscripts")
            .ToArray());
        Assert.Equal("C", prescript.Elements().First().Value);
        Assert.Equal("27/59", string.Join("/", prescript.Elements().Skip(2).Select(slot => slot.Value)));

        var postScripts = actual.Descendants()
            .Where(element => element.Name.LocalName == "msubsup")
            .ToArray();
        Assert.Equal(new[] { "+", "→" },
            postScripts.Select(script => script.Elements().First().Value));
        Assert.Equal(new[] { "0/1", "27/60" },
            postScripts.Select(script =>
            {
                var slots = script.Elements().Skip(1).ToArray();
                return slots[0].Value + "/" + slots[1].Value;
            }));
    }

    [Theory]
    [InlineData("中文氢燃烧")]
    [InlineData("かなカナ")]
    [InlineData("한글")]
    public void EastAsianTextUsesMathTypeFarEastTypefaceAndRoundTrips(string literal)
    {
        var source = "<math><mi>E</mi><mo>=</mo><mtext>" + literal
            + "</mtext><mo>+</mo><mi>x</mi></math>";
        AssertRoundTrip(source);

        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
            source, inline: true, 12);
        var structure = generated.Mtef.AsSpan(generated.StructureOffset);
        foreach (var character in literal)
        {
            var farEastRecord = new byte[]
            {
                2, 0, 12 + 128,
                (byte)(character & 0xFF),
                (byte)(character >> 8),
            };
            var westernTextRecord = new byte[]
            {
                2, 0, 1 + 128,
                (byte)(character & 0xFF),
                (byte)(character >> 8),
            };
            Assert.True(structure.IndexOf(farEastRecord) >= 0,
                $"U+{(int)character:X4} was not emitted with MTEF fnTEXT_FE.");
            Assert.True(structure.IndexOf(westernTextRecord) < 0,
                $"U+{(int)character:X4} was incorrectly emitted with the western Text typeface.");
        }
    }

    [Fact]
    public void WesternTextRemainsInTheWesternTextTypeface()
    {
        const string source = "<math><mtext>A</mtext></math>";
        var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
            source, inline: true, 12);
        var structure = generated.Mtef.AsSpan(generated.StructureOffset);
        Assert.True(structure.IndexOf(new byte[] { 2, 0, 1 + 128, (byte)'A', 0 }) >= 0);
        Assert.True(structure.IndexOf(new byte[] { 2, 0, 12 + 128, (byte)'A', 0 }) < 0);
    }

    [Theory]
    [InlineData("<math><mo>∘</mo></math>")]
    [InlineData("<math><mi>f</mi><mo>∘</mo><mi>g</mi></math>")]
    [InlineData("<math><msup><mn>90</mn><mo>∘</mo></msup></math>")]
    public void CircleMtCodeRemainsCircleIndependentOfLegacyFontPosition(string source)
    {
        AssertRoundTrip(source);
        Assert.Contains("∘", RoundTrip(source));
        Assert.DoesNotContain("°", RoundTrip(source));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(source.Replace("∘", "°")));
    }

    [Theory]
    [InlineData("%")]
    [InlineData("#")]
    [InlineData("!")]
    [InlineData("@")]
    [InlineData("123")]
    [InlineData("a.b")]
    public void UprightOrdinaryIdentifierCanUseNativeTextTypeface(string literal)
    {
        var escaped = new XElement("mi", new XAttribute("mathvariant", "normal"), literal).ToString();
        AssertRoundTrip("<math>" + escaped + "</math>");
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math>" + escaped + "</math>"),
            MathTypeMtefCodec.SemanticSignature("<math>" + escaped.Replace("normal", "bold") + "</math>"));
    }

    [Theory]
    [InlineData("\u200B")]
    [InlineData("\u2060")]
    [InlineData("\u200B\u2060")]
    [InlineData("\u200D\u200C")]
    public void UnrenderableFormatControlsCannotBeAcceptedByRelaxingTokenComparison(string controls)
    {
        // MathPage renders these unsupported CHARs as question-mark diamonds.
        // Word boundary artifacts must be sanitized before MathJax, not made to
        // pass merely by declaring their mtext and mo representations equivalent.
        var separated = "<math><mi>α</mi>" + string.Concat(controls.Select(c => "<mo>" + c + "</mo>")) + "</math>";
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(separated),
            MathTypeMtefCodec.SemanticSignature(RoundTrip(separated)));
    }

    [Theory]
    [InlineData("\\alpha \u200B\u2060", "\\alpha ")]
    [InlineData("x\u200Cy\uFEFF", "xy")]
    [InlineData("x\u200Bx", "x")]
    [InlineData("x\u2061y\u200D", "x\u2061y\u200D")]
    public void ConverterSourceUsesTheSameBoundaryCleanupAsMetadata(string source, string expected)
    {
        var cleaned = OfficeSessionDocument.SanitizeFormulaBoundaryArtifacts(source);
        Assert.Equal(expected, cleaned);
        var session = new OfficeSessionDocument { Lines = new List<FormulaLine> { new() { Id = "line", Latex = source } } };
        Assert.Equal(session.ToMetadata().Latex, cleaned);
    }

    [Fact]
    public void ActualSymbolAndStyleLossMustRemainDetectable()
    {
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mo>+</mo></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mtext>+</mtext></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='bold'>x</mi></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mi mathvariant='normal'>x</mi></math>"));
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mo>\u200B</mo></math>"),
            MathTypeMtefCodec.SemanticSignature("<math><mo>\u2060</mo></math>"));
    }
}

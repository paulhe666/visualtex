using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeGreekAndLooseFenceRoundTripTests
{
    private static string RoundTrip(string xml) =>
        MathTypeMtefCodec.ReadEquationNativeMathMl(
            MathTypeMtefCodec.CreateEquationNativeAtFontSize(xml, true, 12).EquationNative);

    public static IEnumerable<object[]> GreekLetters()
    {
        foreach (var letter in "αβγδεζηθικλμνξοπρςστυφχψωϑϕϖϱϵϰΓΔΘΛΞΠΣΥΦΨΩ")
            yield return new object[] { letter.ToString() };
    }

    [Theory]
    [MemberData(nameof(GreekLetters))]
    public void ExplicitUprightGreekOwnsNormalFontStyleAndUnicodeCode(string letter)
    {
        var xml = $"<math><mi mathvariant='normal'>{letter}</mi></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(xml, true, 12);
        var output = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(output).Descendants().Single(e => e.Name.LocalName == "mi");
        Assert.Equal(letter, token.Value);
        Assert.Equal("normal", (string?)token.Attribute("mathvariant"));
        Assert.True((int?)token.Attribute("data-mtef-explicit-typeface") < 0);
        Assert.Equal((int)letter[0], (int?)token.Attribute("data-mtef-mtcode"));
        // Upright capitals take MathType's own Symbol position (its TeX
        // importer's choice); other letters stay Unicode text-font characters
        // and must not reuse a Symbol/Euclid encoded8 slot.
        if (letter[0] is >= (char)0x0391 and <= (char)0x03A9)
            Assert.NotNull(token.Attribute("data-mtef-encoded8"));
        else
            Assert.Null(token.Attribute("data-mtef-encoded8"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(xml), MathTypeMtefCodec.SemanticSignature(output));
        var implicitNative = MathTypeMtefCodec.CreateEquationNativeAtFontSize(
            $"<math><mi>{letter}</mi></math>", true, 12);
        Assert.False(native.EquationNative.SequenceEqual(implicitNative.EquationNative));
    }

    [Fact]
    public void InheritedUprightGreekInsideNestedScriptsKeepsEachOwner()
    {
        const string xml = "<math><msubsup><mi>x</mi><mn>1</mn><mstyle mathvariant='normal'><msup><mi>α</mi><mi>β</mi></msup></mstyle></msubsup></math>";
        var output = RoundTrip(xml);
        var root = XDocument.Parse(output);
        Assert.Equal(2, root.Descendants().Count(e => e.Name.LocalName == "mi" && (string?)e.Attribute("mathvariant") == "normal"));
        Assert.Single(root.Descendants(), e => e.Name.LocalName == "msup");
        Assert.Single(root.Descendants(), e => e.Name.LocalName == "msubsup");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(xml), MathTypeMtefCodec.SemanticSignature(output));
    }

    public static IEnumerable<object[]> GroupedScripts()
    {
        foreach (var kind in new[] { "msup", "msub", "msubsup" })
        {
            var slots = kind == "msubsup" ? "<mi>i</mi><mn>2</mn>" : "<mn>2</mn>";
            yield return new object[] {
                $"<math><mo>(</mo><mi>x</mi><mo>+</mo><mi>y</mi><{kind}><mrow><mo>)</mo></mrow>{slots}</{kind}></math>", kind, 1 };
            yield return new object[] {
                $"<math><mi>a</mi><mo>+</mo><mo>(</mo><mi>x</mi><{kind}><mo>)</mo>{slots}</{kind}><mo>+</mo><mi>b</mi></math>", kind, 1 };
        }
        yield return new object[] {
            "<math><mo>(</mo><mo>[</mo><mi>x</mi><msup><mo>]</mo><mn>2</mn></msup><msub><mo>)</mo><mi>i</mi></msub></math>", "msub", 2 };
        yield return new object[] {
            "<math><mo>(</mo><mi>x</mi><msup><mo>)</mo><mn>2</mn></msup><mo>+</mo><mo>[</mo><mi>y</mi><msub><mo>]</mo><mi>j</mi></msub></math>", "msup", 2 };
    }

    [Theory]
    [MemberData(nameof(GroupedScripts))]
    public void LooseClosingFenceScriptBecomesActualGroupedNativeBase(string xml, string kind, int fences)
    {
        var output = RoundTrip(xml);
        var document = XDocument.Parse(output);
        Assert.Equal(fences, document.Descendants().Count(e => e.Name.LocalName == "mfenced"));
        var scripts = document.Descendants().Where(e => e.Name.LocalName == kind).ToArray();
        Assert.NotEmpty(scripts);
        Assert.Contains(scripts, e => e.Elements().First().DescendantsAndSelf().Any(b => b.Name.LocalName == "mfenced"));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(xml), MathTypeMtefCodec.SemanticSignature(output));
    }

    [Theory]
    [InlineData("<math><mo>(</mo><mi>x</mi><msup><mo>]</mo><mn>2</mn></msup></math>")]
    [InlineData("<math><mo>(</mo><mi>x</mi><msup><mi>y</mi><mn>2</mn></msup></math>")]
    [InlineData("<math><mo stretchy='false'>(</mo><mi>x</mi><msup><mo stretchy='false'>)</mo><mn>2</mn></msup></math>")]
    [InlineData("<math><mo>|</mo><mi>x</mi><msup><mo>|</mo><mn>2</mn></msup></math>")]
    public void UnmatchedFixedOrAmbiguousBarsAreNotPromotedToNativeGroups(string xml)
    {
        var prepared = MathTypeMtefCodec.PrepareMathMlForMathTypeInterop(xml);
        Assert.DoesNotContain(XDocument.Parse(prepared).Descendants(), e => e.Name.LocalName == "mfenced");
    }

    [Theory]
    [InlineData("<math><mi>a</mi><mo>+</mo><msup><mrow/><mn>2</mn></msup><mi>x</mi></math>")]
    [InlineData("<math><msubsup><mrow/><mn>6</mn><mn>14</mn></msubsup><mi mathvariant='normal'>C</mi></math>")]
    public void EmptyScriptBasesKeepExistingPrescriptOwnership(string xml)
    {
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(xml), MathTypeMtefCodec.SemanticSignature(RoundTrip(xml)));
    }
}

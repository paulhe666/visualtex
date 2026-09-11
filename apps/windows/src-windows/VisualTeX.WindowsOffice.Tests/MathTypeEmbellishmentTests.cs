using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeEmbellishmentTests
{
    // A raw CHAR + EMBELL list, independent of the writer being tested.
    // WIRIS MTEF v5: CHAR=2, EMBELL=6, emb1PRIME=5, emb2PRIME=6,
    // embBPRIME=7 and emb3PRIME=18. Signed typeface 2 is encoded as 130.
    private static byte[] Character(char value, params byte[] embellishments) =>
        new byte[] { 2, (byte)(embellishments.Length == 0 ? 0 : 1), 130, (byte)value, (byte)(value >> 8) }
        .Concat(embellishments.SelectMany(code => new byte[] { 6, 0, code }))
        .Concat(embellishments.Length == 0 ? Array.Empty<byte>() : new byte[] { 0 }).ToArray();

    private static byte[] Native(params byte[][] objects)
    {
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi></math>", false);
        var mtef = seed.Mtef.Take(seed.StructureOffset)
            .Concat(new byte[] { 1, 0 }).Concat(objects.SelectMany(o => o))
            .Concat(new byte[] { 0, 0 }).ToArray();
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)mtef.Length), 0, native, 8, 4);
        return native;
    }

    private static XElement Read(params byte[][] objects) =>
        XElement.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(Native(objects)));

    [Theory]
    [InlineData(5, "′")]
    [InlineData(6, "″")]
    [InlineData(18, "‴")]
    public void PrimeEmbellishmentsBecomeActualSuperscripts(int code, string mark)
    {
        var math = Read(Character('y', (byte)code));
        var script = Assert.Single(math.Descendants().Where(e => e.Name.LocalName == "msup"));
        Assert.Equal("y", script.Elements().First().Value);
        Assert.Equal(mark, script.Elements().Last().Value);
        Assert.DoesNotContain(math.Descendants(), e => e.Name.LocalName == "mover");
        var latex = MathMlToLatexConverter.Convert(math.ToString());
        Assert.Contains("prime", latex);
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature("<math><mi>y</mi></math>"),
            MathTypeMtefCodec.SemanticSignature(math.ToString()));
    }

    [Fact]
    public void BackPrimeIsAPrescriptNotARightSuperscript()
    {
        var math = Read(Character('x', 7));
        Assert.Contains(math.Descendants(), e => e.Name.LocalName == "mprescripts");
        Assert.Contains("‵", math.Value);
        Assert.Equal(@"{}^{\backprime}\mathrm{x}", MathMlToLatexConverter.Convert(math.ToString()));
    }

    [Fact]
    public void MultiscriptsPreserveBothSidesAndAllScriptColumns()
    {
        const string source = "<math><mmultiscripts><mi>T</mi><mi>i</mi><mi>j</mi>"
            + "<mi>k</mi><mi>l</mi><mprescripts/><mi>a</mi><mi>b</mi>"
            + "<mi>c</mi><mi>d</mi></mmultiscripts></math>";
        Assert.Equal(@"{}_{c}^{d}{}_{a}^{b}T_{i}^{j}{}_{k}^{l}", MathMlToLatexConverter.Convert(source));
    }

    [Fact]
    public void BackAndForwardPrimeCannotDisappearAtLatexBoundary()
    {
        const string source = "<math><mmultiscripts><mi>x</mi><none/><mo>&#x2033;</mo>"
            + "<mprescripts/><none/><mo>&#x2035;</mo></mmultiscripts></math>";
        Assert.Equal(@"{}^{\backprime}x^{\prime\prime}", MathMlToLatexConverter.Convert(source));
    }

    [Fact]
    public void MultiplePrimesAndAccentBelongToTheSameCharacter()
    {
        var math = Read(Character('z', 5, 9, 6));
        var sup = Assert.Single(math.Descendants().Where(e => e.Name.LocalName == "msup"));
        Assert.Equal("′″", sup.Elements().Last().Value);
        Assert.Equal("mover", sup.Elements().First().Name.LocalName);
        Assert.Equal("z", sup.Descendants().Single(e => e.Name.LocalName == "mi").Value);
    }

    [Fact]
    public void NestedSubscriptKeepsItsOwnDoublePrime()
    {
        // E with subscript y-double-prime and superscript II, as in the report.
        var math = Read(Character('E'), new byte[] { 3, 0, 29, 0, 0, 1, 0 },
            Character('y', 6), new byte[] { 0, 1, 0 }, Character('I'), Character('I'), new byte[] { 0, 0 });
        var scripts = Assert.Single(math.Descendants().Where(e => e.Name.LocalName == "msubsup"));
        Assert.Equal("y″", scripts.Elements().ElementAt(1).Value);
        Assert.Equal("II", scripts.Elements().ElementAt(2).Value);
        Assert.Contains(scripts.Elements().ElementAt(1).DescendantsAndSelf(), e => e.Name.LocalName == "msup");
    }

    [Fact]
    public void UnknownEmbellishmentMustNotSilentlyLoseContent()
        => Assert.Throws<InvalidDataException>(() => Read(Character('x', 99)));
}

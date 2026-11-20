using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeIntervalAndSeparatorTests
{
    public static IEnumerable<object[]> Intervals()
    {
        const string sides = "()[]";
        for (var left = 0; left < sides.Length; left++)
        for (var right = 0; right < sides.Length; right++)
            yield return new object[] { sides[left].ToString(), sides[right].ToString(), left | (right << 4) };
    }

    [Theory]
    [MemberData(nameof(Intervals))]
    public void DocumentedNativeIntervalRecordRetainsEndpoints(string open, string close, int variation)
    {
        // Independent MTEF v5 fixture: selector 9, four possible characters on
        // each side, body LINE then both fnFENCE CHAR slots. No encoder roundtrip
        // can hide a matching mistake in writer and reader.
        var seed = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi></math>", false);
        var root = new byte[] { 1, 0, 3, 0, 9, (byte)variation, 0,
            1, 0, 2, 0, 131, 120, 0, 0,
            2, 0, 150, (byte)open[0], 0, 2, 0, 150, (byte)close[0], 0, 0, 0, 0 };
        var mtef = seed.Mtef.Take(seed.StructureOffset).Concat(root).ToArray();
        var equationNative = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        BitConverter.GetBytes(mtef.Length).CopyTo(equationNative, 8);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(equationNative);
        AssertFence(xml, open, close, "x");
    }

    // The endpoint pairs MathType's renderer draws (the other interval
    // templates crash it), plus the parenthesis and bracket templates.
    private static readonly string[] RenderableIntervals = { "(]", "[)", "[[", "][", "]]", "()", "[]" };

    public static IEnumerable<object[]> WrittenIntervals() =>
        Intervals().Where(row => RenderableIntervals.Contains((string)row[0] + (string)row[1]));

    public static IEnumerable<object[]> UnrenderableIntervals() =>
        Intervals().Where(row => !RenderableIntervals.Contains((string)row[0] + (string)row[1]));

    [Theory]
    [MemberData(nameof(UnrenderableIntervals))]
    public void IntervalTheRendererCannotDrawIsNotWrittenAsAnIntervalTemplate(string open, string close, int variation)
    {
        var source = "<math><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>"
            + open + "</mo><mi>x</mi><mo data-mjx-texclass='CLOSE'>" + close + "</mo></mrow></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var records = native.Mtef.Skip(native.StructureOffset).ToArray();
        Assert.False(records[2] == 3 && records[4] == 9 && records[5] == variation);
        // the fence role is lost, so the strict comparison reports it
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),
            MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative)));
    }

    [Theory]
    [MemberData(nameof(WrittenIntervals))]
    public void ExplicitIntervalRowWritesNativeFence(string open, string close, int variation)
    {
        var source = "<math><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>"
            + open + "</mo><mi>x</mi><mo data-mjx-texclass='CLOSE'>" + close + "</mo></mrow></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var records = native.Mtef.Skip(native.StructureOffset).ToArray();
        Assert.Equal(3, records[2]); // TMPL, never ordinary delimiter CHARs.
        if ((open, close) is not ("(", ")") and not ("[", "]"))
        {
            Assert.Equal(9, records[4]);
            Assert.Equal(variation, records[5]);
        }
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        AssertFence(xml, open, close, "x");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    [Theory]
    [InlineData("", "a,b,c,d")]
    [InlineData(" separators=''", "abcd")]
    [InlineData(" separators=' ; | '", "a;b|c|d")]
    [InlineData(" separators='&#x3B;;'", "a;b;c;d")]
    public void MfencedSeparatorsAreActualNativeCharacters(string attributes, string expected)
    {
        var source = "<math><mfenced" + attributes + "><mi>a</mi><mi>b</mi><mi>c</mi><mi>d</mi></mfenced></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        AssertFence(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative), "(", ")", expected);
    }

    [Fact]
    public void ExplicitRowDoesNotInventDefaultCommas()
    {
        const string source = "<math><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>(</mo><mi>a</mi><mo>+</mo><mi>b</mi><mo data-mjx-texclass='CLOSE'>)</mo></mrow></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        AssertFence(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative), "(", ")", "a+b");
    }

    [Fact]
    public void ImplicitSeparatorsInheritStyleBeforeNativeFontSelection()
    {
        const string source = "<math><mstyle mathvariant='bold'><mfenced separators=';'><mi>x</mi><mi>y</mi></mfenced></mstyle></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        var separator = Assert.Single(xml.Descendants(), e => e.Name.LocalName == "mo" && e.Value == ";");
        Assert.Equal("bold", (string?)separator.Attribute("mathvariant"));
    }

    [Fact]
    public void NestedFenceOperandsAndScriptsRetainTheirSlotsOnRewrite()
    {
        const string source = "<math><msup><mfenced open='[' close=')'><mfenced><mi>x</mi><mi>y</mi></mfenced><mfrac><mi>a</mi><mi>b</mi></mfrac></mfenced><mn>2</mn></msup></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, source, false);
        var xml = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative));
        var script = Assert.Single(xml.Descendants(), e => e.Name.LocalName == "msup");
        Assert.Equal("2", script.Elements().Last().Value);
        var fences = script.Elements().First().DescendantsAndSelf().Where(e => e.Name.LocalName == "mfenced").ToArray();
        Assert.Equal(2, fences.Length);
        Assert.Equal("x,y,ab", fences[0].Value);
        Assert.Equal("x,y", fences[1].Value);
        Assert.Single(fences[0].Descendants(), e => e.Name.LocalName == "mfrac");
    }

    private static void AssertFence(string xml, string open, string close, string body)
    {
        var fence = Assert.Single(XDocument.Parse(xml).Descendants(), e => e.Name.LocalName == "mfenced");
        Assert.Equal(open, (string?)fence.Attribute("open"));
        Assert.Equal(close, (string?)fence.Attribute("close"));
        Assert.Equal(body, fence.Value);
    }
}

using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNativeSpaceContractTests
{
    [Theory]
    [InlineData("1em", "1em", 0xEF05)]
    [InlineData("1.00em", "1em", 0xEF05)]
    [InlineData(" 2em ", "2em", 0xEF06)]
    [InlineData("2.000em", "2em", 0xEF06)]
    public void KnownNativeSpaceRetainsWidthAndIdentity(string width, string expected, int code)
    {
        var native = MathTypeMtefCodec.CreateEquationNative($"<math><mi>x</mi><mspace width='{width}'/><mi>y</mi></math>", false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var gap = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName == "mspace");
        Assert.Equal(expected, (string?)gap.Attribute("width"));
        Assert.Equal(code, (int?)gap.Attribute("data-mtef-mtcode"));
        Assert.Equal(24, (int?)gap.Attribute("data-mtef-explicit-typeface"));
        Assert.Contains(Convert.ToHexString(new byte[] { 2, 0, 0x98, (byte)code, (byte)(code >> 8) }), Convert.ToHexString(native.Mtef));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Theory]
    [InlineData("0.167em")]
    [InlineData("-0.167em")]
    [InlineData("0.2778em")]
    [InlineData("0.333em")]
    [InlineData("1.01em")]
    [InlineData("1.000000000000000000000000000000000001em")]
    [InlineData("12pt")]
    [InlineData("thinmathspace")]
    [InlineData("")]
    public void UncharacterizedWidthsRetainExistingFallbackWithoutClaimingIdentity(string width)
    {
        var native = MathTypeMtefCodec.CreateEquationNative($"<math><mi>x</mi><mspace width='{width}'/><mi>y</mi></math>", false);
        var gap = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative))
            .Descendants().Single(e => e.Name.LocalName == "mspace");
        Assert.Equal("0.2em", (string?)gap.Attribute("width"));
        Assert.Null(gap.Attribute("data-mtef-mtcode"));
        Assert.Contains("0200982000", Convert.ToHexString(native.Mtef));
    }

    [Fact]
    public void EditedWidthOverridesStaleNativeIdentity()
    {
        var native = MathTypeMtefCodec.CreateEquationNative("<math><mi>x</mi><mspace width='2em'/><mi>y</mi></math>", false);
        var doc = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        doc.Descendants().Single(e => e.Name.LocalName == "mspace").SetAttributeValue("width", "1em");
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, doc.ToString(), false);
        var gap = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative))
            .Descendants().Single(e => e.Name.LocalName == "mspace");
        Assert.Equal("1em", (string?)gap.Attribute("width"));
        Assert.Equal(0xEF05, (int?)gap.Attribute("data-mtef-mtcode"));
    }

    [Fact]
    public void QuadBetweenProseRunsIsNotCoalescedToAnAsciiSpace()
    {
        const string source = "<math><mtext>two words</mtext><mspace width='1em'/><mtext>more words</mtext></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var elements = XDocument.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative)).Root!.Elements().ToArray();
        Assert.Equal(new[] { "mtext", "mspace", "mtext" }, elements.Select(e => e.Name.LocalName));
        Assert.Equal("two words", elements[0].Value);
        Assert.Equal("1em", (string?)elements[1].Attribute("width"));
        Assert.Equal("more words", elements[2].Value);
    }

    [Theory]
    [InlineData(0x04, 0x05)]
    [InlineData(0x10, 0xEF05)]
    public void KnownNativeSpacePreservesItsOriginalEncodedPositionWidth(int options, int position)
    {
        var native = MathTypeMtefCodec.CreateEquationNative("<math><mspace width='1em'/></math>", false);
        var bytes = native.Mtef.Take(native.StructureOffset).Concat(new byte[] { 1, 0, 2, (byte)options, 0x98, 0x05, 0xEF, (byte)position })
            .Concat(options == 0x10 ? new[] { (byte)(position >> 8) } : Array.Empty<byte>()).Concat(new byte[] { 0, 0 }).ToArray();
        var ole = native.EquationNative.Take(28).Concat(bytes).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)bytes.Length), 0, ole, 8, 4);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(ole);
        var gap = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName == "mspace");
        Assert.Equal("1em", (string?)gap.Attribute("width"));
        Assert.Equal(position, (int?)gap.Attribute(options == 0x10 ? "data-mtef-encoded16" : "data-mtef-encoded8"));
        Assert.Equal(bytes, MathTypeMtefCodec.RewriteEquationNative(ole, xml, false).Mtef);
    }

    [Fact]
    public void AdjacentAndScriptedSpacesKeepTheirIndividualWidths()
    {
        const string source = "<math><mspace width='1em'/><mspace width='2em'/><msub><mi>x</mi><mrow><mi>a</mi><mspace width='1em'/><mi>b</mi></mrow></msub></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(new[] { "1em", "2em", "1em" }, XDocument.Parse(xml).Descendants().Where(e => e.Name.LocalName == "mspace").Select(e => (string?)e.Attribute("width")));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }
}

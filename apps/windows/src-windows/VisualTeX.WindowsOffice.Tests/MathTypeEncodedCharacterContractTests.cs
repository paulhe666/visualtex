using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeEncodedCharacterContractTests
{
    [Fact]
    public void Encoded16AndSemanticMtCodeRetainTheirIndependentValues()
    {
        var raw = new byte[] { 2, 0x10, 0x7F, 0x1C, 0x23, 0x1D, 0x23 };
        var (native, xml, token) = Decode(raw);
        Assert.Equal(0x231C, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(0x231D, (int?)token.Attribute("data-mtef-encoded16"));
        Assert.Null(token.Attribute("data-mtef-encoded8"));
        AssertRoundTrip(native, xml, raw);
        // The same contract must survive the operator emission branch too.
        token.Name = token.Name.Namespace + "mo";
        AssertRoundTrip(native, token.Document!.ToString(), raw);
    }

    [Fact]
    public void Encoded16WithoutMtCodeDoesNotInventAnMtCodeOnRewrite()
    {
        var raw = new byte[] { 2, 0x30, 0x7F, 0x1C, 0x23 };
        var (native, xml, token) = Decode(raw);
        Assert.Null(token.Attribute("data-mtef-mtcode"));
        Assert.Equal("true", (string?)token.Attribute("data-mtef-no-mtcode"));
        Assert.Equal(0x231C, (int?)token.Attribute("data-mtef-encoded16"));
        AssertRoundTrip(native, xml, raw);
        token.Name = token.Name.Namespace + "mo";
        AssertRoundTrip(native, token.Document!.ToString(), raw);
    }

    [Fact]
    public void SmallEncoded16PositionRetainsItsDeclaredWidth()
    {
        var raw = new byte[] { 2, 0x10, 0x7F, 0x41, 0, 0x41, 0 };
        var (native, xml, token) = Decode(raw);
        Assert.Equal(65, (int?)token.Attribute("data-mtef-encoded16"));
        Assert.Null(token.Attribute("data-mtef-encoded8"));
        AssertRoundTrip(native, xml, raw);
    }

    [Fact]
    public void Encoded8AndGreekMappingsRemainUnchanged()
    {
        foreach (var raw in new[] {
            new byte[] { 2, 0x04, 0x7F, 0x41, 0, 0x41 },
            new byte[] { 2, 0x24, 0x7F, 0x41 } })
        {
            var (native, xml, token) = Decode(raw);
            Assert.Equal(65, (int?)token.Attribute("data-mtef-encoded8"));
            Assert.Null(token.Attribute("data-mtef-encoded16"));
            AssertRoundTrip(native, xml, raw);
        }
        var greek = MathTypeMtefCodec.CreateEquationNative("<math><mi>α</mi></math>", false);
        Assert.Equal(new byte[] { 1, 0, 2, 4, 0x84, 0xB1, 3, 0x61, 0, 0 },
            greek.Mtef.Skip(greek.StructureOffset).ToArray());
        var upright = MathTypeMtefCodec.CreateEquationNative("<math><mi mathvariant='normal'>α</mi></math>", false);
        var xmlUpright = MathTypeMtefCodec.ReadEquationNativeMathMl(upright.EquationNative);
        var alpha = XDocument.Parse(xmlUpright).Descendants().Single(e => e.Name.LocalName == "mi");
        Assert.Null(alpha.Attribute("data-mtef-encoded8"));
        Assert.Null(alpha.Attribute("data-mtef-encoded16"));
        Assert.Equal(0x03B1, (int?)alpha.Attribute("data-mtef-mtcode"));
    }

    [Fact]
    public void ConflictingFontPositionWidthsCannotBeWritten()
    {
        var raw = new byte[] { 2, 0x10, 0x7F, 0x41, 0, 0x41, 0 };
        var (native, _, token) = Decode(raw);
        token.SetAttributeValue("data-mtef-encoded8", 65);
        Assert.Throws<InvalidDataException>(() => MathTypeMtefCodec.RewriteEquationNative(
            native, token.Document!.ToString(), false));
    }

    [Fact]
    public void Encoded16KeepsFunctionStartAndEmbellishmentFlags()
    {
        var raw = new byte[] { 2, 0x13, 0x7F, 0x41, 0, 0x41, 0, 6, 0, 2, 0 };
        var (native, xml, token) = Decode(raw);
        Assert.Equal("true", (string?)token.Attribute("data-mtef-function-start"));
        Assert.Equal(65, (int?)token.Attribute("data-mtef-encoded16"));
        AssertRoundTrip(native, xml, raw);
    }

    private static (byte[] Native, string Xml, XElement Token) Decode(byte[] character)
    {
        // Use a production-generated explicit font prefix, then inject one raw
        // CHAR whose MTCode and encoded position are independently specified.
        var seed = MathTypeMtefCodec.CreateEquationNative(
            "<math><mi mathvariant='normal'>α</mi></math>", false);
        Assert.Equal(new byte[] { 2, 0, 0x7F, 0xB1, 3, 0, 0 }, seed.Mtef.Skip(seed.Mtef.Length - 7).ToArray());
        // Definitions may be serialized at the beginning of a LINE. Move those
        // exact generated definitions into the global prefix so replacement of
        // the root CHAR cannot drop its font/style contract.
        var definitions = seed.Mtef.Skip(seed.StructureOffset + 2)
            .Take(seed.Mtef.Length - seed.StructureOffset - 2 - 7);
        var mtef = seed.Mtef.Take(seed.StructureOffset).Concat(definitions)
            .Concat(new byte[] { 1, 0 }).Concat(character).Concat(new byte[] { 0, 0 }).ToArray();
        var native = seed.EquationNative.Take(28).Concat(mtef).ToArray();
        Buffer.BlockCopy(BitConverter.GetBytes((uint)mtef.Length), 0, native, 8, 4);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Attribute("data-mtef-explicit-typeface") != null);
        return (native, xml, token);
    }

    private static void AssertRoundTrip(byte[] native, string xml, byte[] character)
    {
        var rewritten = MathTypeMtefCodec.RewriteEquationNative(native, xml, false);
        Assert.Equal(new byte[] { 1, 0 }.Concat(character).Concat(new byte[] { 0, 0 }).ToArray(),
            rewritten.Mtef.Skip(rewritten.StructureOffset).ToArray());
        Assert.Equal(MathTypeMtefCodec.ExtractNativeMtef(native).Take(rewritten.StructureOffset).ToArray(),
            rewritten.Mtef.Take(rewritten.StructureOffset).ToArray());
    }
}

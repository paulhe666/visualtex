using System.Reflection;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeAlternateGlyphAuthorityTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var item in new[] {
            (0x2223, 0xE92E, 0x88, "Euclid Math One", "EuclidMath1"),
            (0x2224, 0xEA2E, 0x8A, "Euclid Math One", "EuclidMath1"),
            (0x228A, 0xEA42, 0xD8, "Euclid Math Two", "EuclidMath2"),
            (0x228B, 0xEA43, 0xD9, "Euclid Math Two", "EuclidMath2")
        })
        foreach (var variant in new[] { "normal", "italic", "bold", "bold-italic" })
            yield return new object[] { item.Item1, item.Item2, item.Item3, item.Item4, item.Item5, variant };
    }

    public static IEnumerable<object[]> StyleCases() =>
        Cases().Select(item => new object[] { item[0], item[5] });

    public static IEnumerable<object[]> OrdinaryAuthorityCases()
    {
        foreach (var item in new[] {
            (0x2224, 0x8C, "Euclid Math One", "EuclidMath1"),
            (0x228A, 0xD6, "Euclid Math Two", "EuclidMath2"),
            (0x228B, 0xD7, "Euclid Math Two", "EuclidMath2")
        })
        foreach (var variant in new[] { "normal", "italic", "bold", "bold-italic" })
            yield return new object[] { item.Item1, item.Item2, item.Item3, item.Item4, variant };
    }

    public static IEnumerable<object[]> IdentifierCases()
    {
        foreach (var scalar in new[] { 0x78, 0x03B1, 0x2220 })
        foreach (var variant in new[] { "normal", "italic", "bold", "bold-italic" })
            yield return new object[] { scalar, variant };
        foreach (var variant in new[] { "script", "bold-script", "double-struck" })
            yield return new object[] { 0x41, variant };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void AlternateKeepsExactAuthorityAndStyle(int scalar, int mtCode, int slot, string font, string encoding, string variant)
    {
        var source = Source(scalar, variant, true);
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = Token(xml);
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Equal("1", (string?)token.Attribute("data-mjx-alternate"));
        Assert.Equal(mtCode, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(slot, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(variant, (string?)token.Attribute("mathvariant") ?? "normal");
        Assert.Contains(font + "\0", Encoding.ASCII.GetString(native.Mtef));
        Assert.Contains(encoding + "\0", Encoding.ASCII.GetString(native.Mtef));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Theory]
    [MemberData(nameof(StyleCases))]
    public void OrdinaryCounterpartDoesNotAcquireAlternate(int scalar, string variant)
    {
        var source = Source(scalar, variant, false);
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = Token(xml);
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Null(token.Attribute("data-mjx-alternate"));
        Assert.Equal(variant, (string?)token.Attribute("mathvariant") ?? "normal");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Theory]
    [MemberData(nameof(StyleCases))]
    public void OrdinaryAndAlternateNativeStreamsRemainDistinct(int scalar, string variant)
    {
        var ordinary = MathTypeMtefCodec.CreateEquationNative(Source(scalar, variant, false), false);
        var alternate = MathTypeMtefCodec.CreateEquationNative(Source(scalar, variant, true), false);
        Assert.False(ordinary.Mtef.SequenceEqual(alternate.Mtef));
    }

    [Theory]
    [MemberData(nameof(OrdinaryAuthorityCases))]
    public void OrdinaryRelationKeepsExactNativeAuthorityAndStyle(int scalar, int slot, string font, string encoding, string variant)
    {
        var source = Source(scalar, variant, false);
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = Token(xml);
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Null(token.Attribute("data-mjx-alternate"));
        Assert.Equal(scalar, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(slot, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(variant, (string?)token.Attribute("mathvariant") ?? "normal");
        var bits = (byte)(variant == "bold" ? 1 : variant == "italic" ? 2 : variant == "bold-italic" ? 3 : 0);
        Assert.Equal((font, encoding, bits), NativeBinding(native.Mtef, token));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);

        var alternate = MathTypeMtefCodec.CreateEquationNative(Source(scalar, variant, true), false);
        var alternateToken = Token(MathTypeMtefCodec.ReadEquationNativeMathMl(alternate.EquationNative));
        Assert.Equal("mo", alternateToken.Name.LocalName);
        Assert.Equal("1", (string?)alternateToken.Attribute("data-mjx-alternate"));
        Assert.NotEqual((int?)token.Attribute("data-mtef-mtcode"), (int?)alternateToken.Attribute("data-mtef-mtcode"));
        Assert.NotEqual((int?)token.Attribute("data-mtef-encoded8"), (int?)alternateToken.Attribute("data-mtef-encoded8"));
        Assert.Equal((font, encoding, bits), NativeBinding(alternate.Mtef, alternateToken));
    }

    [Theory]
    [MemberData(nameof(IdentifierCases))]
    public void OtherIdentifiersKeepTheirTokenClassAndAlphabetStyle(int scalar, string variant)
    {
        var source = $"<math><mi mathvariant='{variant}'>&#x{scalar:X};</mi></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = Token(xml);
        Assert.Equal("mi", token.Name.LocalName);
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Null(token.Attribute("data-mjx-alternate"));
        Assert.Equal(variant, (string?)token.Attribute("mathvariant") ?? "italic");
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    [Theory]
    [InlineData("normal")]
    [InlineData("italic")]
    [InlineData("bold")]
    [InlineData("bold-italic")]
    public void OrdinaryModelsDoesNotAcquireAlternateFromTheAmbiguousVDashTuple(string variant)
    {
        // Vendor input maps models and vDash to the same font/MTCode/slot.
        // Do not infer alternate identity for the existing ordinary tuple.
        // A future distinct vDash tuple can be added without freezing this loss.
        var ordinary = MathTypeMtefCodec.CreateEquationNative(Source(0x22A8, variant, false), false);
        Assert.Null(Token(MathTypeMtefCodec.ReadEquationNativeMathMl(ordinary.EquationNative)).Attribute("data-mjx-alternate"));
    }

    private static (string Font, string Encoding, byte Style) NativeBinding(byte[] mtef, XElement token)
    {
        // Follow CHAR -> FONT_STYLE_DEF -> FONT_DEF -> ENCODING_DEF. Merely
        // finding an expected font name anywhere in the stream is insufficient.
        const BindingFlags flags = BindingFlags.Static | BindingFlags.NonPublic;
        var codec = typeof(MathTypeMtefCodec);
        var bindings = (Dictionary<int, (int FontDefinition, byte CharacterStyle)>)codec
            .GetMethod("ReadNativeCharacterStyleBindings", flags)!.Invoke(null, new object[] { mtef })!;
        var face = (int)token.Attribute("data-mtef-explicit-typeface")!;
        var binding = bindings[face];
        var layout = codec.GetMethod("ReadPrefixLayout", flags)!.Invoke(null, new object[] { mtef })!;
        var prefix = (IReadOnlyList<(byte Type, int Start, int End)>)layout.GetType()
            .GetProperty("Records", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(layout)!;
        var root = (IReadOnlyList<(byte Type, int Start, int End)>)codec
            .GetMethod("ReadRootLeadingFormattingRecords", flags)!.Invoke(null, new object[] { mtef })!;
        var records = prefix.Concat(root).ToArray();
        var definition = records.Where(r => r.Type == 17).ElementAt(binding.FontDefinition - 1);
        var args = new object[] { mtef, definition.Start + 1 };
        var encodingIndex = (int)codec.GetMethod("ReadUnsigned", flags)!.Invoke(null, args)!;
        var nameStart = (int)args[1];
        var font = Encoding.ASCII.GetString(mtef, nameStart, definition.End - nameStart - 1);
        // MTEF reserves encoding indices 1-4; stream definitions start at 5.
        var encoding = records.Where(r => r.Type == 19).ElementAt(encodingIndex - 5);
        var encodingName = Encoding.ASCII.GetString(mtef, encoding.Start + 1, encoding.End - encoding.Start - 2);
        return (font, encodingName, binding.CharacterStyle);
    }

    private static string Source(int scalar, string variant, bool alternate) =>
        $"<math><mo mathvariant='{variant}'{(alternate ? " data-mjx-alternate='1'" : "")}>&#x{scalar:X};</mo></math>";
    private static XElement Token(string xml) =>
        XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
}

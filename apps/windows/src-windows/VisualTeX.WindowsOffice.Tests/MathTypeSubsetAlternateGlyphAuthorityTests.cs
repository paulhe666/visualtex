using System.Reflection;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeSubsetAlternateGlyphAuthorityTests
{
    public static IEnumerable<object[]> Families()
    {
        foreach (var family in new[] {
            (0x2288, 0x2288, 0xDA, 0xEA0C, 0xE0),
            (0x2289, 0x2289, 0xDB, 0xEA0B, 0xE1),
            (0x2ACB, 0xEA44, 0xDC, 0xEA46, 0xDE),
            (0x2ACC, 0xEA45, 0xDD, 0xEA47, 0xDF)
        })
        foreach (var style in new[] { "normal", "italic", "bold", "bold-italic" })
            yield return new object[] { family.Item1, family.Item2, family.Item3, family.Item4, family.Item5, style };
    }

    [Theory]
    [MemberData(nameof(Families))]
    public void OrdinaryAndAlternateKeepExactDistinctAuthorities(int scalar, int ordinaryMtCode, int ordinarySlot,
        int alternateMtCode, int alternateSlot, string style)
    {
        var ordinary = Verify(scalar, ordinaryMtCode, ordinarySlot, style, false);
        var alternate = Verify(scalar, alternateMtCode, alternateSlot, style, true);
        Assert.False(ordinary.SequenceEqual(alternate));
    }

    private static byte[] Verify(int scalar, int mtCode, int slot, string style, bool alternate)
    {
        var source = $"<math><mo mathvariant='{style}'{(alternate ? " data-mjx-alternate='1'" : "")}>&#x{scalar:X};</mo></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12);
        var decoded = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(decoded).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Equal(alternate ? "1" : null, (string?)token.Attribute("data-mjx-alternate"));
        Assert.Equal(style, MathMlTokenStyleContract.ResolveEffectiveVariant(token));
        Assert.Equal(mtCode, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(slot, (int?)token.Attribute("data-mtef-encoded8"));
        var bits = (byte)(style == "bold" ? 1 : style == "italic" ? 2 : style == "bold-italic" ? 3 : 0);
        Assert.Equal(("Euclid Math Two", "EuclidMath2", bits), NativeBinding(native.Mtef, token));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(decoded));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, decoded, false).Mtef);
        return native.Mtef;
    }

    private static (string Font, string Encoding, byte Style) NativeBinding(byte[] mtef, XElement token)
    {
        // Follow this token's actual CHAR -> FONT_STYLE_DEF -> FONT_DEF ->
        // ENCODING_DEF binding; a matching unrelated stream string is not proof.
        const BindingFlags flags = BindingFlags.Static | BindingFlags.NonPublic;
        var codec = typeof(MathTypeMtefCodec);
        var bindings = (Dictionary<int, (int FontDefinition, byte CharacterStyle)>)codec
            .GetMethod("ReadNativeCharacterStyleBindings", flags)!.Invoke(null, new object[] { mtef })!;
        var binding = bindings[(int)token.Attribute("data-mtef-explicit-typeface")!];
        var layout = codec.GetMethod("ReadPrefixLayout", flags)!.Invoke(null, new object[] { mtef })!;
        var prefix = (IReadOnlyList<(byte Type, int Start, int End)>)layout.GetType()
            .GetProperty("Records", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(layout)!;
        var leading = (IReadOnlyList<(byte Type, int Start, int End)>)codec
            .GetMethod("ReadRootLeadingFormattingRecords", flags)!.Invoke(null, new object[] { mtef })!;
        var records = prefix.Concat(leading).ToArray();
        var definition = records.Where(r => r.Type == 17).ElementAt(binding.FontDefinition - 1);
        var arguments = new object[] { mtef, definition.Start + 1 };
        var encodingIndex = (int)codec.GetMethod("ReadUnsigned", flags)!.Invoke(null, arguments)!;
        var nameStart = (int)arguments[1];
        var font = Encoding.ASCII.GetString(mtef, nameStart, definition.End - nameStart - 1);
        var encoding = records.Where(r => r.Type == 19).ElementAt(encodingIndex - 5);
        var encodingName = Encoding.ASCII.GetString(mtef, encoding.Start + 1, encoding.End - encoding.Start - 2);
        return (font, encodingName, binding.CharacterStyle);
    }
}

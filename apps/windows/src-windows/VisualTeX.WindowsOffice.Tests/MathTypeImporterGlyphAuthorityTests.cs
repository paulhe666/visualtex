using System.Reflection;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Font positions recorded from MathType 7.9's own importer
// (MTSetEqnFromLangStr with MathJax MathML, read back as MTEF).
public sealed class MathTypeImporterGlyphAuthorityTests
{
    public static IEnumerable<object[]> Cases()
    {
        // MathJax's token form for each command: element and alternate marker.
        foreach (var item in new[] {
            (0x25B3, 0x56, "MT Extra", "mo", false),        // \bigtriangleup
            (0x2235, 0x51, "MT Extra", "mo", false),        // \because
            (0x21CC, 0x83, "MT Extra", "mo", true),         // \rightleftharpoons
            (0x2218, 0x6F, "MT Extra", "mo", false),        // \circ
            (0x223C, 0x7E, "Symbol", "mo", false),          // \sim
            (0x22A4, 0x95, "Euclid Math One", "mi", false), // \top
            (0x22A8, 0x91, "Euclid Math One", "mo", false), // \models
            (0x2216, 0x82, "Euclid Math One", "mo", false), // \setminus
            (0x2293, 0xF3, "Euclid Math Two", "mo", false), // \sqcap
            (0x2A7D, 0x84, "Euclid Math Two", "mo", false), // \leqslant
        })
        foreach (var variant in new[] { "normal", "italic", "bold", "bold-italic" })
            yield return new object[] { item.Item1, item.Item2, item.Item3, item.Item4, item.Item5, variant };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void SymbolUsesMathTypesOwnFontPosition(int scalar, int slot, string font, string element, bool alternate, string variant)
    {
        var source = $"<math><{element} mathvariant='{variant}'{(alternate ? " data-mjx-alternate='1'" : "")}>&#x{scalar:X};</{element}></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
        Assert.Equal(element, token.Name.LocalName);
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Equal(slot, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(variant, (string?)token.Attribute("mathvariant") ?? "normal");
        var bits = (byte)(variant == "bold" ? 1 : variant == "italic" ? 2 : variant == "bold-italic" ? 3 : 0);
        Assert.Equal((font, bits), NativeBinding(native.Mtef, token));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
        Assert.Equal(native.Mtef, MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, xml, false).Mtef);
    }

    // TeX-importer positions: upright capitals, Hebrew letters and MathType's
    // private codes for negated relations (MTCode differs from the character).
    [Theory]
    [InlineData("mi", 0x0393, 0x0393, 0x47, "Symbol")]           // \Gamma
    [InlineData("mi", 0x03A9, 0x03A9, 0x57, "Symbol")]           // \Omega
    [InlineData("mi", 0x2136, 0x2136, 0xFB, "Euclid Math One")]  // \beth
    [InlineData("mo", 0x2A85, 0xE932, 0x8A, "Euclid Math Two")]  // \lessapprox
    [InlineData("mo", 0x2AC5, 0xE90C, 0xD4, "Euclid Math Two")]  // \subseteqq
    public void TexImporterPositionsRoundTrip(string element, int scalar, int mtCode, int slot, string font)
    {
        var source = $"<math><{element} mathvariant='normal'>&#x{scalar:X};</{element}></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
        Assert.Equal(char.ConvertFromUtf32(scalar), token.Value);
        Assert.Equal(mtCode, (int?)token.Attribute("data-mtef-mtcode"));
        Assert.Equal(slot, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal((font, (byte)0), NativeBinding(native.Mtef, token));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    // MathType stores these variants as another command's glyph, so they
    // are written, and read back, as that command.
    [Theory]
    [InlineData(0x25B3, 0x25B3, false)] // \vartriangle -> \triangle
    [InlineData(0x223C, 0x223C, false)] // \thicksim -> \sim
    [InlineData(0x2216, 0x2216, false)] // \smallsetminus -> \setminus
    [InlineData(0x2A87, 0x2270, true)]  // \nleqslant -> \nleqq
    public void MergedVariantIsWrittenAsMathTypesCommand(int scalar, int merged, bool alternate)
    {
        var source = $"<math><mo data-mjx-alternate='1'>&#x{scalar:X};</mo></math>";
        var expected = $"<math><mo{(alternate ? " data-mjx-alternate='1'" : "")}>&#x{merged:X};</mo></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        Assert.Equal(MathTypeMtefCodec.CreateEquationNative(expected, false).Mtef, native.Mtef);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    // MathType keeps ASCII punctuation in its Function face (Times New Roman)
    // and \mid in its built-in typeface 12; both read back as operators.
    [Theory]
    [InlineData("(", 2)][InlineData(")", 2)][InlineData(",", 2)][InlineData(":", 2)]
    [InlineData("|", 2)][InlineData("!", 2)][InlineData("&#x2223;", 12)]
    [InlineData("+", 6)][InlineData("=", 6)]
    public void OperatorUsesMathTypesFace(string value, int face)
    {
        var source = $"<math><mo>{value}</mo></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName is "mi" or "mo" or "mtext");
        Assert.Equal("mo", token.Name.LocalName);
        Assert.Equal(face, NativeFace(native.Mtef));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    private static int NativeFace(byte[] mtef)
    {
        // The single CHAR record of the root line: [2][options][typeface+128]...
        var position = Array.IndexOf(mtef, (byte)2, MathTypeMtefCodec.FindRootStructureOffset(mtef));
        return mtef[position + 2] - 128;
    }

    [Fact]
    public void MathJaxTriangleIdentifierUsesMtExtra()
    {
        // MathJax writes \triangle as an upright identifier; MathType stores it
        // as MT Extra position 0x56, the same glyph as \bigtriangleup.
        const string source = "<math><mi mathvariant='normal'>&#x25B3;</mi></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var xml = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var token = XDocument.Parse(xml).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
        Assert.Equal(0x56, (int?)token.Attribute("data-mtef-encoded8"));
        Assert.Equal(("MT Extra", (byte)0), NativeBinding(native.Mtef, token));
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(source), MathTypeMtefCodec.SemanticSignature(xml));
    }

    private static (string Font, byte Style) NativeBinding(byte[] mtef, XElement token)
    {
        const BindingFlags flags = BindingFlags.Static | BindingFlags.NonPublic;
        var codec = typeof(MathTypeMtefCodec);
        var bindings = (Dictionary<int, (int FontDefinition, byte CharacterStyle)>)codec
            .GetMethod("ReadNativeCharacterStyleBindings", flags)!.Invoke(null, new object[] { mtef })!;
        var binding = bindings[(int)token.Attribute("data-mtef-explicit-typeface")!];
        var layout = codec.GetMethod("ReadPrefixLayout", flags)!.Invoke(null, new object[] { mtef })!;
        var prefix = (IReadOnlyList<(byte Type, int Start, int End)>)layout.GetType()
            .GetProperty("Records", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(layout)!;
        var root = (IReadOnlyList<(byte Type, int Start, int End)>)codec
            .GetMethod("ReadRootLeadingFormattingRecords", flags)!.Invoke(null, new object[] { mtef })!;
        var definition = prefix.Concat(root).Where(r => r.Type == 17).ElementAt(binding.FontDefinition - 1);
        var args = new object[] { mtef, definition.Start + 1 };
        codec.GetMethod("ReadUnsigned", flags)!.Invoke(null, args);
        var nameStart = (int)args[1];
        return (Encoding.ASCII.GetString(mtef, nameStart, definition.End - nameStart - 1), binding.CharacterStyle);
    }
}

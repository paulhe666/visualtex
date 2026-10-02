using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeGroupAccentContractTests
{
    public static IEnumerable<object[]> Groups()
    {
        foreach (var body in new[] {
            "<mrow><mi>a</mi><mi>b</mi></mrow>",
            "<mrow><mi>′</mi><mi>′</mi></mrow>",
            "<mrow><mi>′</mi><mi>′</mi><mi>′</mi><mi>′</mi></mrow>",
            "<mfrac><mi>a</mi><mi>b</mi></mfrac>",
            "<mo mathvariant='italic'>∮</mo>",
            "<mo mathvariant='bold-italic'>∑</mo>" })
        foreach (var mark in new[] { "˙", "¨", "^", "~", "→", "←", "↔", "¯", "ˇ", "˘", "´", "`", "˚" })
        foreach (var nested in new[] { false, true })
            yield return new object[] { body, mark, nested };
    }

    [Theory]
    [MemberData(nameof(Groups))]
    public void EveryGroupAccentPreservesBothNativeObjectAndSelectedMark(string body, string mark, bool nested)
    {
        var accent = new XElement("mover", new XAttribute("accent", "true"), XElement.Parse(body), new XElement("mo", mark));
        var expression = nested ? new XElement("msub", new XElement("mover", new XAttribute("accent", "true"), accent, new XElement("mo", "^")), new XElement("mn", "2")) : accent;
        var source = new XElement("math", expression).ToString();
        var expected = MathTypeMtefCodec.SemanticSignature(source);
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12);
        var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(actual));
        var rewrite = MathTypeMtefCodec.RewriteEquationNative(native.EquationNative, actual, false);
        Assert.Equal(expected, MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewrite.EquationNative)));
        Assert.NotEqual(expected, MathTypeMtefCodec.SemanticSignature(new XElement("math", XElement.Parse(body)).ToString()));
    }

    [Theory]
    [InlineData("^", 0x0302)]
    [InlineData("~", 0x0303)]
    public void StretchingHatBoxIncludesItsRequiredNativeGlyph(string mark, int expectedGlyph)
    {
        var source = $"<math><mover accent='true'><mrow><mi>a</mi><mi>b</mi></mrow><mo stretchy='true'>{mark}</mo></mover></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12);
        var bytes = native.Mtef.Skip(native.StructureOffset).ToArray();
        // A real HatBox contains the main LINE followed by the expanding CHAR.
        // The reader and writer cannot make a missing glyph pass by agreeing on
        // the selector alone. Check the native record independently.
        var required = new byte[] { 2, 0, 150, (byte)expectedGlyph, (byte)(expectedGlyph >> 8) };
        Assert.True(Enumerable.Range(0, Math.Max(0, bytes.Length - required.Length + 1)).Any(i => bytes.Skip(i).Take(required.Length).SequenceEqual(required)), "HatBox omitted its native expanding glyph CHAR.");
    }
}

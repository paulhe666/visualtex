using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlVerifiedAlternateSourceTests
{
    public static IEnumerable<object[]> Families()
    {
        foreach (var pair in new[]
        {
            ("∣", @"\shortmid"), ("∤", @"\nshortmid"), ("∖", @"\smallsetminus"),
            ("⊊", @"\varsubsetneq"), ("⊋", @"\varsupsetneq"),
            ("⫋", @"\varsubsetneqq"), ("⫌", @"\varsupsetneqq"),
            ("⊈", @"\nsubseteqq"), ("⊉", @"\nsupseteqq"),
        })
        foreach (var style in new[] { "normal", "italic", "bold", "bold-italic" })
            yield return new object[] { pair.Item1, pair.Item2, style };
    }

    [Theory]
    [MemberData(nameof(Families))]
    public void VerifiedNativeAlternateRetainsSourceCommandRoleAndStyle(string token, string command, string style)
    {
        var source = $"<math><mo data-mjx-alternate='1' mathvariant='{style}'>{token}</mo></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12);
        var decoded = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var element = XDocument.Parse(decoded).Descendants().Single(e => e.Name.LocalName == "mo");
        Assert.Equal(token, element.Value);
        Assert.Equal("1", (string?)element.Attribute("data-mjx-alternate"));
        Assert.Equal(style, MathMlTokenStyleContract.ResolveEffectiveVariant(element));
        var expected = style switch
        {
            "italic" => @"\mathit{" + command + "}",
            "bold" => @"\mathbf{" + command + "}",
            "bold-italic" => @"\mathbfit{" + command + "}",
            _ => command,
        };
        Assert.Equal(expected, MathMlToLatexConverter.Convert(decoded));
    }

    [Theory]
    [MemberData(nameof(Families))]
    public void OrdinaryNativeCounterpartDoesNotAcquireAlternateSource(string token, string command, string style)
    {
        var source = $"<math><mo mathvariant='{style}'>{token}</mo></math>";
        var native = MathTypeMtefCodec.CreateEquationNativeAtFontSize(source, false, 12);
        var decoded = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        var element = XDocument.Parse(decoded).Descendants().Single(e => e.Name.LocalName is "mi" or "mo");
        // Both ordinary and alternate relation glyphs now use their verified
        // native authority; typography must not change the operator role.
        Assert.Equal("mo", element.Name.LocalName);
        Assert.Equal(token, element.Value);
        Assert.Null(element.Attribute("data-mjx-alternate"));
        Assert.Equal(style, MathMlTokenStyleContract.ResolveEffectiveVariant(element));
        Assert.DoesNotContain(command, MathMlToLatexConverter.Convert(decoded));
        Assert.DoesNotContain(command, MathMlToLatexConverter.Convert(
            $"<math><mrow data-mjx-alternate='1'><mo mathvariant='{style}'>{token}</mo></mrow></math>"));
    }
}

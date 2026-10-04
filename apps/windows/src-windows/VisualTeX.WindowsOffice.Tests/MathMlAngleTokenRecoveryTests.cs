using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlAngleTokenRecoveryTests
{
    public static IEnumerable<object[]> Cases()
    {
        foreach (var (token, command) in new[] { ("∡", @"\measuredangle"), ("∢", @"\sphericalangle") })
        foreach (var (tag, variant, wrapper) in new[]
        {
            ("mo", "", @"\mathrm"), ("mi", "", @"\mathit"),
            ("mi", "normal", @"\mathrm"), ("mo", "normal", @"\mathrm"),
            ("mi", "italic", @"\mathit"), ("mo", "italic", @"\mathit"),
            ("mi", "bold", @"\mathbf"), ("mo", "bold-italic", @"\mathbfit"),
        })
        {
            var attribute = variant.Length == 0 ? "" : $" mathvariant='{variant}'";
            yield return new object[] { $"<{tag}{attribute}>{token}</{tag}>", wrapper + "{" + command + "}" };
        }
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void CanonicalAngleCommandPreservesEffectiveStyle(string tokenMathMl, string expected)
        => Assert.Equal(expected, MathMlToLatexConverter.Convert("<math>" + tokenMathMl + "</math>"));

    [Theory]
    [MemberData(nameof(Cases))]
    public void CanonicalAngleCommandPreservesScripts(string tokenMathMl, string expected)
        => Assert.Equal(expected + "_{i}^{2}", MathMlToLatexConverter.Convert(
            "<math><msubsup>" + tokenMathMl + "<mi>i</mi><mn>2</mn></msubsup></math>"));

    [Theory]
    [InlineData("∡", @"\measuredangle")]
    [InlineData("∢", @"\sphericalangle")]
    public void NearestInheritedVariantWins(string token, string command)
    {
        Assert.Equal(@"\mathrm{" + command + "}", MathMlToLatexConverter.Convert(
            $"<math><mstyle mathvariant='italic'><mstyle mathvariant='normal'><mi>{token}</mi></mstyle></mstyle></math>"));
        Assert.Equal(@"\mathit{" + command + "}", MathMlToLatexConverter.Convert(
            $"<math><mstyle mathvariant='normal'><mi mathvariant='italic'>{token}</mi></mstyle></math>"));
        Assert.Equal(@"\mathrm{" + command + "}", MathMlToLatexConverter.Convert(
            $"<math><mstyle mathvariant='bold-italic'><mo mathvariant='normal'>{token}</mo></mstyle></math>"));
    }

    [Fact]
    public void ExistingNormalAngleCommandKeepsItsDefault()
    {
        Assert.Equal(@"\mathrm{\angle}", MathMlToLatexConverter.Convert("<math><mi mathvariant='normal'>∠</mi></math>"));
        Assert.Equal(@"\angle", MathMlToLatexConverter.Convert("<math><mo>∠</mo></math>"));
        Assert.Equal(@"\mathit{\angle}", MathMlToLatexConverter.Convert("<math><mi>∠</mi></math>"));
    }
}

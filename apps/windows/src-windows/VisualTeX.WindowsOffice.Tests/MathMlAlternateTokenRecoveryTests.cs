using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlAlternateTokenRecoveryTests
{
    [Theory]
    [InlineData("∥", @"\shortparallel")]
    [InlineData("∦", @"\nshortparallel")]
    public void ExplicitAlternateRecoversItsSupportedCommand(string token, string command)
    {
        Assert.Equal(command, MathMlToLatexConverter.Convert(
            $"<math><mo data-mjx-alternate='1'>{token}</mo></math>"));
        Assert.Equal("a" + command + " b", MathMlToLatexConverter.Convert(
            $"<math><mi>a</mi><mo data-mjx-alternate='1'>{token}</mo><mi>b</mi></math>"));
    }

    [Theory]
    [InlineData("∥", @"\shortparallel", "italic", @"\mathit")]
    [InlineData("∥", @"\shortparallel", "bold", @"\mathbf")]
    [InlineData("∥", @"\shortparallel", "bold-italic", @"\mathbfit")]
    [InlineData("∦", @"\nshortparallel", "italic", @"\mathit")]
    [InlineData("∦", @"\nshortparallel", "bold", @"\mathbf")]
    [InlineData("∦", @"\nshortparallel", "bold-italic", @"\mathbfit")]
    public void AlternateSelectionPrecedesStyleAndScriptRecovery(
        string token, string command, string variant, string wrapper)
        => Assert.Equal(wrapper + "{" + command + "}_{i}^{2}",
            MathMlToLatexConverter.Convert(
                $"<math><msubsup><mstyle mathvariant='{variant}'>"
                + $"<mo data-mjx-alternate='1'>{token}</mo></mstyle><mi>i</mi><mn>2</mn>"
                + "</msubsup></math>"));

    [Theory]
    [InlineData("", @"\parallel")]
    [InlineData("0", @"\parallel")]
    [InlineData("2", @"\parallel")]
    [InlineData("true", @"\parallel")]
    public void OnlyKnownExplicitAlternateEvidenceChangesTheCommand(string alternate, string command)
        => Assert.Equal(command, MathMlToLatexConverter.Convert(
            $"<math><mo data-mjx-alternate='{alternate}'>∥</mo></math>"));

    [Fact]
    public void ContainerAlternateDoesNotLeakToChildren()
        => Assert.Equal(@"\parallel", MathMlToLatexConverter.Convert(
            "<math><mrow data-mjx-alternate='1'><mo>∥</mo></mrow></math>"));


    public static IEnumerable<object[]> PlanckGlyphSourceCases()
    {
        foreach (var role in new[] { "mi", "mo" })
        foreach (var style in new[] { "normal", "italic", "bold", "bold-italic" })
        foreach (var alternate in new[] { false, true })
            yield return new object[] { role, style, alternate };
    }

    [Theory]
    [MemberData(nameof(PlanckGlyphSourceCases))]
    public void PlanckGlyphSelectionPreservesDeclaredInheritedStyleAndScriptSlots(
        string role, string style, bool alternate)
    {
        var command = alternate ? @"\hbar" : @"\hslash";
        var expected = style switch
        {
            "italic" => @"\mathit{" + command + "}",
            "bold" => @"\mathbf{" + command + "}",
            "bold-italic" => @"\mathbfit{" + command + "}",
            "normal" when role == "mi" => @"\mathrm{" + command + "}",
            _ => command,
        };
        var flag = alternate ? " data-mjx-alternate='1'" : string.Empty;
        var token = $"<{role}{flag}>ℏ</{role}>";
        Assert.Equal(expected, MathMlToLatexConverter.Convert(
            $"<math><{role} mathvariant='{style}'{flag}>ℏ</{role}></math>"));
        Assert.Equal(expected, MathMlToLatexConverter.Convert(
            $"<math><mstyle mathvariant='{style}'>{token}</mstyle></math>"));
        Assert.Equal(expected + "_{i}^{2}", MathMlToLatexConverter.Convert(
            $"<math><msubsup><mstyle mathvariant='{style}'>{token}</mstyle>"
            + "<mi>i</mi><mn>2</mn></msubsup></math>"));
    }

    [Theory]
    [InlineData("mi", false)]
    [InlineData("mo", false)]
    [InlineData("mi", true)]
    [InlineData("mo", true)]
    public void PlanckGlyphWithoutDeclaredStyleUsesOnlyItsOwnAlternate(string role, bool alternate)
    {
        var flag = alternate ? " data-mjx-alternate='1'" : string.Empty;
        Assert.Equal(alternate ? @"\hbar" : @"\hslash", MathMlToLatexConverter.Convert(
            $"<math><{role}{flag}>ℏ</{role}></math>"));
    }

    [Theory]
    [InlineData("")]
    [InlineData("0")]
    [InlineData("2")]
    [InlineData("true")]
    public void PlanckGlyphInvalidOrContainerAlternateCannotInventHbar(string flag)
    {
        foreach (var role in new[] { "mi", "mo" })
        {
            Assert.Equal(@"\hslash", MathMlToLatexConverter.Convert(
                $"<math><{role} data-mjx-alternate='{flag}'>ℏ</{role}></math>"));
            Assert.Equal(@"\hslash", MathMlToLatexConverter.Convert(
                $"<math><mrow data-mjx-alternate='1'><{role}>ℏ</{role}></mrow></math>"));
        }
    }

    [Fact]
    public void PlanckGlyphAlternateDoesNotChangeProseRecovery()
        => Assert.Equal(@"\text{ℏ}", MathMlToLatexConverter.Convert(
            "<math><mtext data-mjx-alternate='1'>ℏ</mtext></math>"));

    [Fact]
    public void OtherScalarsRemainOnTheirExistingRecoveryPath()
        => Assert.Equal("⫽", MathMlToLatexConverter.Convert(
            "<math><mo data-mjx-alternate='1'>⫽</mo></math>"));
}

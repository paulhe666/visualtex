using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlDelimiterSourceRecoveryTests
{
    public static IEnumerable<object[]> ActiveDelimiterCases()
    {
        yield return new object[] { "(", @"(" };
        yield return new object[] { ")", @")" };
        yield return new object[] { "[", @"[" };
        yield return new object[] { "]", @"]" };
        yield return new object[] { "/", @"/" };
        yield return new object[] { "\\", @"\backslash" };
        yield return new object[] { "|", @"|" };
        yield return new object[] { "‖", @"\|" };
        yield return new object[] { "{", @"\{" };
        yield return new object[] { "}", @"\}" };
        yield return new object[] { "⟨", @"\langle" };
        yield return new object[] { "⟩", @"\rangle" };
        yield return new object[] { "⌈", @"\lceil" };
        yield return new object[] { "⌉", @"\rceil" };
        yield return new object[] { "⌊", @"\lfloor" };
        yield return new object[] { "⌋", @"\rfloor" };
        yield return new object[] { "⌜", @"\ulcorner" };
        yield return new object[] { "⌝", @"\urcorner" };
        yield return new object[] { "⌞", @"\llcorner" };
        yield return new object[] { "⌟", @"\lrcorner" };
        yield return new object[] { "⎰", @"\lmoustache" };
        yield return new object[] { "⎱", @"\rmoustache" };
        yield return new object[] { "⟮", @"\lgroup" };
        yield return new object[] { "⟯", @"\rgroup" };
        yield return new object[] { "⏐", @"\arrowvert" };
        yield return new object[] { "⎪", @"\bracevert" };
        yield return new object[] { "↑", @"\uparrow" };
        yield return new object[] { "↓", @"\downarrow" };
        yield return new object[] { "↕", @"\updownarrow" };
        yield return new object[] { "⇑", @"\Uparrow" };
        yield return new object[] { "⇓", @"\Downarrow" };
        yield return new object[] { "⇕", @"\Updownarrow" };
        yield return new object[] { "", @"." };
        yield return new object[] { ".", @"." };
        yield return new object[] { "<", @"<" };
        yield return new object[] { ">", @">" };
        yield return new object[] { "〈", @"\langle" };
        yield return new object[] { "〉", @"\rangle" };
    }

    [Theory]
    [MemberData(nameof(ActiveDelimiterCases))]
    public void DelimiterRoleUsesRegisteredCanonicalSourceOnEitherSide(string scalar, string source)
    {
        Assert.Equal(@"\left" + source + @" x\right.",
            Recover(Fence(scalar, "", new XElement("mi", "x"))));
        Assert.Equal(@"\left. x\right" + source,
            Recover(Fence("", scalar, new XElement("mi", "x"))));
    }

    [Theory]
    [InlineData("⚑")]
    [InlineData("∥")]
    public void UnsupportedScalarRetainsLegacySourceWithoutSubstitution(string scalar)
        => Assert.Equal(@"\left" + scalar + @" x\right.",
            Recover(Fence(scalar, "", new XElement("mi", "x"))));

    [Theory]
    [InlineData("(", ")", @"\left( xyz\right)")]
    [InlineData("⟮", "⟯", @"\left\lgroup xyz\right\rgroup")]
    [InlineData("⏐", "", @"\left\arrowvert xyz\right.")]
    [InlineData("", "", @"\left. xyz\right.")]
    public void ExplicitEmptySeparatorsDoNotInventCommas(string open, string close, string expected)
        => Assert.Equal(expected, Recover(Fence(open, close,
            new XElement("mi", "x"), new XElement("mi", "y"), new XElement("mi", "z"))));

    [Fact]
    public void MissingSeparatorsStillDefaultsToComma()
        => Assert.Equal(@"\left( x,y,z\right)", MathMlToLatexConverter.Convert(
            "<math><mfenced><mi>x</mi><mi>y</mi><mi>z</mi></mfenced></math>"));

    [Fact]
    public void ExplicitCommaSeparatorsArePreserved()
        => Assert.Equal(@"\left( x,y,z\right)", MathMlToLatexConverter.Convert(
            "<math><mfenced separators=','><mi>x</mi><mi>y</mi><mi>z</mi></mfenced></math>"));

    [Fact]
    public void F07CornersRecoverNamedSource()
        => Assert.Equal(@"\left\ulcorner x\right\urcorner +\left\llcorner y\right\lrcorner",
            Recover(Fence("⌜", "⌝", new XElement("mi", "x")),
                new XElement("mo", "+"), Fence("⌞", "⌟", new XElement("mi", "y"))));

    [Fact]
    public void F09GroupAndOneSidedArrowRecoverNamedSource()
        => Assert.Equal(@"\left\lgroup x\right\rgroup +\left\arrowvert y\right.",
            Recover(Fence("⟮", "⟯", new XElement("mi", "x")),
                new XElement("mo", "+"), Fence("⏐", "", new XElement("mi", "y"))));

    [Fact]
    public void NestedFencesKeepIdentityAndEmptySeparator()
        => Assert.Equal(@"\left\lgroup x\left\arrowvert y\right. z\right\rgroup",
            Recover(Fence("⟮", "⟯", new XElement("mi", "x"),
                Fence("⏐", "", new XElement("mi", "y")), new XElement("mi", "z"))));

    [Fact]
    public void EmptyBodyKeepsBothFenceRoles()
        => Assert.Equal(@"\left\lmoustache \right\rmoustache",
            Recover(Fence("⎰", "⎱")));

    [Theory]
    [InlineData("<mo>↑</mo>", @"\uparrow")]
    [InlineData("<mo>∥</mo>", @"\parallel")]
    [InlineData("<mo>⟮</mo><mo>⏐</mo>", "⟮⏐")]
    [InlineData("<mtext>⟮|⏐</mtext>", @"\text{⟮|⏐}")]
    public void OrdinaryOperatorsAndProseDoNotAcquireDelimiterRole(string body, string expected)
        => Assert.Equal(expected, MathMlToLatexConverter.Convert("<math>" + body + "</math>"));

    private static XElement Fence(string open, string close, params XElement[] children)
        => new("mfenced", new XAttribute("open", open), new XAttribute("close", close),
            new XAttribute("separators", ""), children);

    private static string Recover(params XElement[] children)
        => MathMlToLatexConverter.Convert(new XElement("math", children).ToString(SaveOptions.DisableFormatting));
}

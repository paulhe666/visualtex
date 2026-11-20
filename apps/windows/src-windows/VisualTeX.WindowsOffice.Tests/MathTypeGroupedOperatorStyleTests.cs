using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeGroupedOperatorStyleTests
{
    [Theory]
    [InlineData("normal")][InlineData("bold")][InlineData("italic")][InlineData("bold-italic")]
    public void GroupedUnicodeOperatorsEqualSameStyledScalarTokens(string style)
    {
        var grouped = $"<math><mo mathvariant='{style}'>↦∥∼</mo></math>";
        var split = $"<math><mstyle mathvariant='{style}'><mo>↦</mo><mo>∥</mo><mo>∼</mo></mstyle></math>";
        Assert.Equal(Signature(grouped), Signature(split));
    }

    [Theory]
    [InlineData("bold", "normal")][InlineData("italic", "normal")]
    [InlineData("bold-italic", "bold")][InlineData("bold-italic", "italic")]
    public void GroupedUnicodeOperatorsCannotLoseEffectiveStyle(string expected, string actual)
    {
        Assert.NotEqual(Signature($"<math><mo mathvariant='{expected}'>↦∥∼</mo></math>"),
            Signature($"<math><mo mathvariant='{actual}'>↦∥∼</mo></math>"));
    }

    [Theory]
    [InlineData("normal")][InlineData("bold")][InlineData("italic")][InlineData("bold-italic")]
    public void DefinitionOperatorSplitPreservesStyle(string style)
    {
        Assert.Equal(Signature($"<math><mo mathvariant='{style}'>:=</mo></math>"),
            Signature($"<math><mstyle mathvariant='{style}'><mo>:</mo><mo>=</mo></mstyle></math>"));
        if (style != "normal")
            Assert.NotEqual(Signature($"<math><mo mathvariant='{style}'>:=</mo></math>"), Signature("<math><mo>:=</mo></math>"));
    }

    [Theory]
    [InlineData("normal")][InlineData("bold")][InlineData("italic")][InlineData("bold-italic")]
    public void AlphabeticOperatorKeepsExistingNameEquivalenceAndActualStyle(string style)
    {
        Assert.Equal(Signature($"<math><mo mathvariant='{style}'>sin</mo></math>"),
            Signature($"<math><mi mathvariant='{style}'>sin</mi></math>"));
        if (style != "normal")
            Assert.NotEqual(Signature($"<math><mo mathvariant='{style}'>sin</mo></math>"), Signature("<math><mo>sin</mo></math>"));
    }

    [Theory]
    [InlineData("∪", "⋃")][InlineData("∩", "⋂")]
    public void ScriptBaseAliasKeepsEffectiveStyleAndNearestOverride(string alias, string canonical)
    {
        string Formula(string glyph, string own) => $"<math><mstyle mathvariant='bold'><msub><mo mathvariant='{own}'>{glyph}</mo><mi>i</mi></msub></mstyle></math>";
        Assert.Equal(Signature(Formula(alias, "bold-italic")), Signature(Formula(canonical, "bold-italic")));
        Assert.NotEqual(Signature(Formula(alias, "bold-italic")), Signature(Formula(canonical, "bold")));
        Assert.Equal(Signature(Formula(alias, "normal")), Signature(Formula(canonical, "normal")));
        Assert.NotEqual(Signature(Formula(alias, "normal")), Signature(Formula(canonical, "")));
    }

    [Fact]
    public void PlainNameRoleAndNearestStyleOverrideRemainStable()
    {
        Assert.Equal(Signature("<math><mo>sin</mo></math>"), Signature("<math><mi mathvariant='normal'>sin</mi></math>"));
        Assert.Equal(Signature("<math><mstyle mathvariant='bold'><mo mathvariant='normal'>∥∼</mo></mstyle></math>"),
            Signature("<math><mo>∥</mo><mo>∼</mo></math>"));
        Assert.NotEqual(Signature("<math><mo mathvariant='bold-italic'>sin</mo></math>"),
            Signature("<math><mo mathvariant='bold'>sin</mo></math>"));
    }

    [Fact]
    public void NonBmpOperatorIsNotSplitIntoSurrogateTokens()
    {
        const string symbols = "\U0001EEF0\U0001EEF1";
        var signature = Signature($"<math><mo mathvariant='bold'>{symbols}</mo></math>");
        Assert.Equal("o[bold](" + symbols + ")", signature);
        Assert.NotEqual(signature, Signature($"<math><mo>{symbols}</mo></math>"));
    }

    [Theory]
    [InlineData("∥∼")][InlineData(":=")][InlineData("sin")]
    public void GroupedWriterPreservesEffectiveStyleThroughTheSharedScalarPath(string value)
    {
        // Every emitted scalar must retain the effective group style.
        var source = $"<math><mo mathvariant='bold-italic'>{value}</mo></math>";
        var native = MathTypeMtefCodec.CreateEquationNative(source, false);
        var actual = MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
        Assert.Equal(Signature(source), Signature(actual));
        Assert.NotEqual(Signature(source.Replace("bold-italic", "bold")), Signature(actual));
    }

    private static string Signature(string mathMl) => MathTypeMtefCodec.SemanticSignature(mathMl);
}

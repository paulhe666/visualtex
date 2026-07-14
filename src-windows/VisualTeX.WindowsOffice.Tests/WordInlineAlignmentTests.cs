using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordInlineAlignmentTests
{
    [Fact]
    public void AlignsExportedFormulaBaselineWithWordTextBaseline()
    {
        Assert.Equal(-4, WordInlineAlignment.CalculateFontPosition(15, 20, 15));
    }

    [Fact]
    public void PreservesAlignmentWhenTheFormulaIsScaled()
    {
        Assert.Equal(-8, WordInlineAlignment.CalculateFontPosition(30, 40, 30));
    }

    [Theory]
    [InlineData(20d, 20d, null)]
    [InlineData(20d, 0d, 0d)]
    [InlineData(20d, 20d, -1d)]
    [InlineData(20d, 20d, 21d)]
    [InlineData(20d, 20d, 20d)]
    public void InvalidOrBottomEdgeBaselinesDoNotMoveTheFormula(
        double actualHeight,
        double exportedHeight,
        double? baseline)
    {
        Assert.Equal(0, WordInlineAlignment.CalculateFontPosition(
            (float)actualHeight,
            (float)exportedHeight,
            baseline is null ? null : (float)baseline.Value));
    }
}

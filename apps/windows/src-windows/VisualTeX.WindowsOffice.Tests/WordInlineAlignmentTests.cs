using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordInlineAlignmentTests
{
    [Fact]
    public void AlignsExportedFormulaBaselineWithWordTextBaseline()
    {
        Assert.Equal(-3, WordInlineAlignment.CalculateFontPosition(15, 20, 15));
    }

    [Fact]
    public void RawGeometryAlignmentScalesWithTheObjectHeight()
    {
        Assert.Equal(-7, WordInlineAlignment.CalculateFontPosition(30, 40, 30));
    }

    [Fact]
    public void CompleteMetadataAppliesLargeFontOpticalLiftAt42Pt()
    {
        Assert.Equal(
            -8,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                43,
                18.9867f,
                14.16f,
                existingFontPosition: -4,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42));
    }

    [Theory]
    [InlineData(12d, -2)]
    [InlineData(18d, -4)]
    [InlineData(24d, -5)]
    public void CommonTextSizesDoNotReceiveLargeFontLift(
        double targetFontSize,
        int expectedPosition)
    {
        Assert.Equal(
            expectedPosition,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                (float)targetFontSize,
                exportedHeight: 20,
                exportedBaseline: 15,
                existingFontPosition: null,
                sourceSemanticFontSizePoints: targetFontSize,
                targetSemanticFontSizePoints: targetFontSize));
    }

    [Fact]
    public void LegacyFormulaScalesItsExistingCorrectBaseline()
    {
        Assert.Equal(
            -11,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                43,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: -4,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42));
    }

    [Fact]
    public void LegacyBaselineScalingRoundTripsAcrossCommonSizes()
    {
        var enlarged = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
            43,
            exportedHeight: 0,
            exportedBaseline: null,
            existingFontPosition: -4,
            sourceSemanticFontSizePoints: 14,
            targetSemanticFontSizePoints: 42);
        Assert.Equal(-11, enlarged);
        Assert.Equal(
            -4,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                14.333f,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: enlarged,
                sourceSemanticFontSizePoints: 42,
                targetSemanticFontSizePoints: 14));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0d)]
    [InlineData(9999999d)]
    public void LegacyFormulaWithoutUsablePositionUsesDescentFallback(
        double? existingPosition)
    {
        Assert.Equal(
            -11,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                43,
                exportedHeight: 0,
                exportedBaseline: null,
                existingFontPosition: existingPosition is null
                    ? null
                    : (float)existingPosition.Value,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: 42));
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

    [Theory]
    [InlineData(14.20395d, 18.9386d, 13.1086d, -3)] // \frac{1}{2}
    [InlineData(10.25595d, 13.6746d, 12.6746d, 0)]  // L^2
    [InlineData(10.3284d, 13.7712d, 10.562d, -1)]   // L_z
    [InlineData(11.14005d, 14.8534d, 12.7054d, -1)] // e^(i*pi) + 1 = 0
    public void AlignsDifferentBoxAspectRatiosByTheirExportedMathBaseline(
        double actualHeight,
        double exportedHeight,
        double exportedBaseline,
        int expectedPosition)
    {
        Assert.Equal(
            expectedPosition,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                (float)actualHeight,
                (float)exportedHeight,
                (float)exportedBaseline,
                existingFontPosition: null,
                sourceSemanticFontSizePoints: 10.5,
                targetSemanticFontSizePoints: 10.5));
    }

    [Theory]
    [InlineData(10.5d, 14d, 11.284d, -1)]           // L_z
    [InlineData(10.79295d, 14.3906d, 11.6746d, -1)] // L^2 and L_zL^2
    [InlineData(10.81605d, 14.4214d, 11.7054d, -1)] // e^(i*pi) + 1 = 0
    public void AlignsStableInlineFramesAtOneSharedTextBaseline(
        double actualHeight,
        double exportedHeight,
        double exportedBaseline,
        int expectedPosition)
    {
        Assert.Equal(
            expectedPosition,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                (float)actualHeight,
                (float)exportedHeight,
                (float)exportedBaseline,
                existingFontPosition: null,
                sourceSemanticFontSizePoints: 10.5,
                targetSemanticFontSizePoints: 10.5));
    }
}

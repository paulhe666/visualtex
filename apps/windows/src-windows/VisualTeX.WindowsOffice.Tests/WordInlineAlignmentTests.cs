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

    [Theory]
    [InlineData(10d, 10d, 9.02d, 0)]
    [InlineData(10d, 10d, 9.01d, -1)]
    [InlineData(10d, 10d, 9.00d, -1)]
    public void UsesCompleteWholePointDescentWithOnlyNearIntegerFloatSnapping(
        double actualHeight,
        double exportedHeight,
        double exportedBaseline,
        int expectedPosition)
    {
        Assert.Equal(
            expectedPosition,
            WordInlineAlignment.CalculateFontPosition(
                (float)actualHeight,
                (float)exportedHeight,
                (float)exportedBaseline));
    }

    [Theory]
    [InlineData(10.25595d, 13.6746d, 12.6746d, 0)]  // L^2
    [InlineData(10.3284d, 13.7712d, 10.562d, -2)]   // L_z
    [InlineData(11.14005d, 14.8534d, 12.7054d, -1)] // e^(i*pi) + 1 = 0
    [InlineData(14.20395d, 18.9386d, 13.1086d, -4)] // \frac{1}{2}
    [InlineData(11.91285d, 15.8838d, 12.6746d, -2)] // L_zL^2 with normal padding
    [InlineData(15.8976d, 21.1968d, 15.3262d, -4)]  // \int_a^b c\,\mathrm{d}e
    public void UsesOneRenderedBaselineRuleForDifferentFormulaStructures(
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
    [InlineData(12d, -3)]
    [InlineData(18d, -4)]
    [InlineData(24d, -6)]
    [InlineData(42d, -10)]
    public void CompleteMetadataHasNoFixedOrLargeFontOpticalLift(
        double actualHeight,
        int expectedPosition)
    {
        Assert.Equal(
            expectedPosition,
            WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                (float)actualHeight,
                exportedHeight: 20,
                exportedBaseline: 15,
                existingFontPosition: -99,
                sourceSemanticFontSizePoints: 14,
                targetSemanticFontSizePoints: actualHeight));
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
        Assert.Equal(
            0,
            WordInlineAlignment.CalculateFontPosition(
                (float)actualHeight,
                (float)exportedHeight,
                baseline is null ? null : (float)baseline.Value));
    }
}

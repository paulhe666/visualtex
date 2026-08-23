using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class FormulaFontSizeTests
{
    [Theory]
    [InlineData(13.24, 13.24)]
    [InlineData(13.255, 13.26)]
    [InlineData(4.0, 5.0)]
    [InlineData(250.0, 200.0)]
    public void NormalizesToCentipointRangeWithoutPresetQuantization(double input, double expected)
    {
        Assert.Equal((float)expected, FormulaFontSize.Normalize(input));
    }

    [Fact]
    public void PresetNavigationIncludesChineseStandardSizes()
    {
        Assert.Equal(14f, FormulaFontSize.NextPreset(13.5));
        Assert.Equal(12f, FormulaFontSize.PreviousPreset(13.5));
        Assert.Equal(10.5f, FormulaFontSize.NextPreset(10));
        Assert.Equal(10f, FormulaFontSize.PreviousPreset(10.5));
        Assert.Equal(15f, FormulaFontSize.NextPreset(14));
        Assert.Equal(42f, FormulaFontSize.NextPreset(36));
        Assert.Equal(6.5f, FormulaFontSize.PreviousPreset(7.5));
    }

    [Theory]
    [InlineData("初号", 42)]
    [InlineData("小初", 36)]
    [InlineData("一号", 26)]
    [InlineData("小一", 24)]
    [InlineData("二号", 22)]
    [InlineData("小二", 18)]
    [InlineData("三号", 16)]
    [InlineData("小三", 15)]
    [InlineData("四号", 14)]
    [InlineData("小四", 12)]
    [InlineData("五号", 10.5)]
    [InlineData("小五", 9)]
    [InlineData("六号", 7.5)]
    [InlineData("小六", 6.5)]
    [InlineData("七号", 5.5)]
    [InlineData("八号", 5)]
    public void ParsesChineseFontSizeNames(string value, double expected)
    {
        Assert.Equal((float)expected, FormulaFontSize.Parse(value));
    }

    [Fact]
    public void FormatsStandardChineseSizesAndNumericCustomSizes()
    {
        Assert.Equal("四号", FormulaFontSize.FormatDisplay(14));
        Assert.Equal("小初", FormulaFontSize.FormatDisplay(36));
        Assert.Equal("13.5", FormulaFontSize.FormatDisplay(13.5));
        Assert.Equal("13.25", FormulaFontSize.FormatDisplay(13.25));
        Assert.Equal("小四（12 磅）", FormulaFontSize.Describe(12));
    }

    [Fact]
    public void OleSizeUsesSemanticPointRatioWithoutDistortion()
    {
        var metadata = Metadata();
        var size = FormulaFontSize.OleSizeAt(metadata, 21);

        Assert.Equal(450f, size.Width, 3);
        Assert.Equal(112.5f, size.Height, 3);
        Assert.Equal(4f, size.Width / size.Height, 3);
    }

    [Fact]
    public void InfersOleFontSizeFromCurrentPhysicalHeight()
    {
        var metadata = Metadata();

        Assert.Equal(21f, FormulaFontSize.InferOleFontSize(450f, 112.5f, metadata));
    }

    [Fact]
    public void SmallInlineOleDoesNotLoseFontSizeAtTwelvePointObjectFloor()
    {
        var metadata = new FormulaMetadata
        {
            FontSizePt = 11,
            RenderFontSizePt = 11,
            RenderWidthPx = 25.171866666666663,
            RenderHeightPx = 10.926133333333333,
        };
        var actualWidthPoints = (float)metadata.RenderWidthPx.Value * 0.75f;
        var actualHeightPoints = (float)metadata.RenderHeightPx.Value * 0.75f;

        Assert.Equal(
            11f,
            FormulaFontSize.InferOleFontSize(
                actualWidthPoints,
                actualHeightPoints,
                metadata));
    }

    [Fact]
    public void StoredInlineOleGeometryKeepsExactSemanticSizeAcrossReopenQuantization()
    {
        var metadata = new FormulaMetadata
        {
            FontSizePt = 14,
            RenderFontSizePt = 14,
            RenderWidthPx = 400,
            RenderHeightPx = 100,
            DisplayMode = "inline",
            WordInlineOleWidthPt = 52.25,
            WordInlineOleHeightPt = 13.75,
        };

        Assert.Equal(
            14f,
            FormulaFontSize.InferOleFontSize(
                currentWidthPoints: 52.5f,
                currentHeightPoints: 13.5f,
                metadata));
    }

    [Fact]
    public void StoredInlineOleGeometryStillDetectsRealUserResize()
    {
        var metadata = new FormulaMetadata
        {
            FontSizePt = 14,
            RenderFontSizePt = 14,
            RenderWidthPx = 400,
            RenderHeightPx = 100,
            DisplayMode = "inline",
            WordInlineOleWidthPt = 40,
            WordInlineOleHeightPt = 10,
        };

        Assert.Equal(
            21f,
            FormulaFontSize.InferOleFontSize(
                currentWidthPoints: 60f,
                currentHeightPoints: 15f,
                metadata));
    }

    [Fact]
    public void SmallInlineOleKeepsSemanticSizeAcrossWordGeometryRounding()
    {
        var metadata = new FormulaMetadata
        {
            FontSizePt = 11,
            RenderFontSizePt = 11,
            RenderWidthPx = 25,
            RenderHeightPx = 11,
        };

        Assert.Equal(
            11f,
            FormulaFontSize.InferOleFontSize(
                currentWidthPoints: 19f,
                currentHeightPoints: 8.5f,
                metadata));
    }

    private static FormulaMetadata Metadata() => new()
    {
        FontSizePt = 14,
        RenderFontSizePt = 14,
        RenderWidthPx = 400,
        RenderHeightPx = 100,
    };
}

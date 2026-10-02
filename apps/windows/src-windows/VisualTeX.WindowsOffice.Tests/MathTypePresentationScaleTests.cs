using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypePresentationScaleTests
{
    [Theory]
    [InlineData(121.5f, 95f, 97f)]
    [InlineData(24.3f, 27f, 27f)]
    [InlineData(188.75f, 160f, 160f)]
    [InlineData(60.6f, 73f, 73f)]
    [InlineData(51.95f, 52f, 64f)]
    public void ExistingMathTypeEditPreservesIntrinsicToDisplayedExtentRatio(
        float sourceOriginalExtent,
        float sourceDisplayedExtent,
        float targetDisplayedExtent)
    {
        var actual = WordFormulaService.CalculateMathTypePreservedOriginalExtent(
            sourceOriginalExtent,
            sourceDisplayedExtent,
            targetDisplayedExtent);

        Assert.Equal(
            targetDisplayedExtent * sourceOriginalExtent / sourceDisplayedExtent,
            actual,
            4);
    }

    [Theory]
    [InlineData(0f, 95f, 97f)]
    [InlineData(121.5f, 0f, 97f)]
    [InlineData(121.5f, 95f, 0f)]
    public void InvalidOriginalExtentRelationshipFallsBackToTargetDisplayExtent(
        float sourceOriginalExtent,
        float sourceDisplayedExtent,
        float targetDisplayedExtent)
    {
        Assert.Equal(
            targetDisplayedExtent,
            WordFormulaService.CalculateMathTypePreservedOriginalExtent(
                sourceOriginalExtent,
                sourceDisplayedExtent,
                targetDisplayedExtent));
    }

    [Fact]
    public void PreservesLegacyInlineMathTypeHorizontalAndVerticalScale()
    {
        var widthScale = WordFormulaService.CalculateMathTypeNativePresentationScale(28.95f, 33f);
        var heightScale = WordFormulaService.CalculateMathTypeNativePresentationScale(12.4f, 14f);

        Assert.Equal(28.95f / 33f, widthScale, 4);
        Assert.Equal(12.4f / 14f, heightScale, 4);
        Assert.Equal(69.30f, 79f * widthScale, 2);
        Assert.Equal(14.17f, 16f * heightScale, 2);
    }

    [Fact]
    public void KeepsWordOriginalExtentRelationshipWhileNativePresentationRescales()
    {
        const float nativeWidth = 79f;
        const float nativeHeight = 16f;
        const float sourceDisplayedWidth = 28.95f;
        const float sourceDisplayedHeight = 12.4f;
        const float sourceOriginalWidth = 36f;
        const float sourceOriginalHeight = 11.25f;
        var widthScale = WordFormulaService.CalculateMathTypeNativePresentationScale(
            sourceDisplayedWidth,
            33f);
        var heightScale = WordFormulaService.CalculateMathTypeNativePresentationScale(
            sourceDisplayedHeight,
            14f);

        var displayedWidth = nativeWidth * widthScale;
        var displayedHeight = nativeHeight * heightScale;
        var originalWidth = WordFormulaService.CalculateMathTypePreservedOriginalExtent(
            sourceOriginalWidth,
            sourceDisplayedWidth,
            displayedWidth);
        var originalHeight = WordFormulaService.CalculateMathTypePreservedOriginalExtent(
            sourceOriginalHeight,
            sourceDisplayedHeight,
            displayedHeight);
        var geometry = WordFormulaService.CalculateMathTypeWordPresentationGeometry(
            "width:28.95pt;height:12.4pt",
            displayedWidth,
            displayedHeight,
            originalWidth,
            originalHeight);

        Assert.Equal(69.30f, displayedWidth, 2);
        Assert.Equal(14.17f, displayedHeight, 2);
        Assert.Equal((int)Math.Round(originalWidth * 20f), geometry.OriginalWidthTwips);
        Assert.Equal((int)Math.Round(originalHeight * 20f), geometry.OriginalHeightTwips);
    }

    [Fact]
    public void PreservesFractionalPointPrecisionInWordOriginalExtent()
    {
        var geometry = WordFormulaService.CalculateMathTypeWordPresentationGeometry(
            "width:28.95pt;height:12.4pt",
            widthPt: 63.3f,
            heightPt: 13.85f,
            originalWidthPt: 63.3f,
            originalHeightPt: 13.85f);

        Assert.Equal("width:63.3pt;height:13.85pt", geometry.ShapeStyle);
        Assert.Equal(1266, geometry.OriginalWidthTwips);
        Assert.Equal(277, geometry.OriginalHeightTwips);
    }

    [Fact]
    public void AddsNativeMathTypeWindowRecordsToMathPageWmf()
    {
        var source = CreateMinimalPlaceableWmf();

        var normalized = WordFormulaService.EnsureMathTypePlaceableWmfWindowRecords(source);

        Assert.Equal(source.Length + 20, normalized.Length);
        Assert.Equal(22u, BitConverter.ToUInt32(normalized, 28));
        Assert.Equal(0x020b, BitConverter.ToUInt16(normalized, 44));
        Assert.Equal(0, BitConverter.ToInt16(normalized, 46));
        Assert.Equal(0, BitConverter.ToInt16(normalized, 48));
        Assert.Equal(0x020c, BitConverter.ToUInt16(normalized, 54));
        Assert.Equal(512, BitConverter.ToInt16(normalized, 56));
        Assert.Equal(2528, BitConverter.ToInt16(normalized, 58));
        Assert.Equal(0, BitConverter.ToUInt16(normalized, 64));
        Assert.Same(
            normalized,
            WordFormulaService.EnsureMathTypePlaceableWmfWindowRecords(normalized));
    }

    [Theory]
    [InlineData(0f, 33f)]
    [InlineData(28.95f, 0f)]
    [InlineData(28.95f, -1f)]
    public void FallsBackToUnitScaleForInvalidSourceGeometry(float wordExtent, float nativeExtent)
    {
        Assert.Equal(1f, WordFormulaService.CalculateMathTypeNativePresentationScale(wordExtent, nativeExtent));
    }

    private static byte[] CreateMinimalPlaceableWmf()
    {
        var bytes = new byte[46];
        WriteUInt32(bytes, 0, 0x9ac6cdd7);
        WriteInt16(bytes, 10, 2528);
        WriteInt16(bytes, 12, 512);
        WriteUInt16(bytes, 14, 2304);
        WriteUInt16(bytes, 22, 1);
        WriteUInt16(bytes, 24, 9);
        WriteUInt16(bytes, 26, 0x0300);
        WriteUInt32(bytes, 28, 12);
        WriteUInt32(bytes, 34, 3);
        WriteUInt32(bytes, 40, 3);
        return bytes;
    }

    private static void WriteInt16(byte[] target, int offset, short value) =>
        WriteUInt16(target, offset, unchecked((ushort)value));

    private static void WriteUInt16(byte[] target, int offset, ushort value)
    {
        target[offset] = (byte)value;
        target[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] target, int offset, uint value)
    {
        target[offset] = (byte)value;
        target[offset + 1] = (byte)(value >> 8);
        target[offset + 2] = (byte)(value >> 16);
        target[offset + 3] = (byte)(value >> 24);
    }
}

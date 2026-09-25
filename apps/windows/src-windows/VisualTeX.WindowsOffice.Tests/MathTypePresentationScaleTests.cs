using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypePresentationScaleTests
{
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
    public void KeepsWordOriginalExtentAlignedWithScaledInlinePresentation()
    {
        const float nativeWidth = 79f;
        const float nativeHeight = 16f;
        var widthScale = WordFormulaService.CalculateMathTypeNativePresentationScale(28.95f, 33f);
        var heightScale = WordFormulaService.CalculateMathTypeNativePresentationScale(12.4f, 14f);

        var displayedWidth = nativeWidth * widthScale;
        var displayedHeight = nativeHeight * heightScale;
        var geometry = WordFormulaService.CalculateMathTypeWordPresentationGeometry(
            "width:28.95pt;height:12.4pt",
            displayedWidth,
            displayedHeight,
            displayedWidth,
            displayedHeight);

        Assert.Equal(69.30f, displayedWidth, 2);
        Assert.Equal(14.17f, displayedHeight, 2);
        Assert.Equal(1386, geometry.OriginalWidthTwips);
        Assert.Equal(283, geometry.OriginalHeightTwips);
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

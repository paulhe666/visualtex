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

    [Fact]
    public void PictureDrawingIgnoresMathTypeUninitializedPadding()
    {
        // Two renders of one equation by MathType differ only in bytes that
        // never draw: a string's pad byte, the rest of a face name, pen width.y.
        var first = CreateMathTypeTextWmf(padding: 0x6D, faceTail: 0xA3, penTail: 0x11);
        var second = CreateMathTypeTextWmf(padding: 0xFF, faceTail: 0x06, penTail: 0x77);

        Assert.NotEqual(first, second);
        Assert.Equal(
            WordFormulaService.MathTypePictureDrawing(first),
            WordFormulaService.MathTypePictureDrawing(second));
        Assert.Equal(
            WordFormulaService.MathTypePictureDrawing(first),
            WordFormulaService.MathTypePictureDrawing(
                WordFormulaService.EnsureMathTypePlaceableWmfWindowRecords(first)));
    }

    [Fact]
    public void PictureDrawingSeesADifferentGlyphOrFont()
    {
        var expected = WordFormulaService.MathTypePictureDrawing(
            CreateMathTypeTextWmf(padding: 0, faceTail: 0, penTail: 0));

        Assert.NotEqual(expected, WordFormulaService.MathTypePictureDrawing(
            CreateMathTypeTextWmf(padding: 0, faceTail: 0, penTail: 0, glyph: 0x6E)));
        Assert.NotEqual(expected, WordFormulaService.MathTypePictureDrawing(
            CreateMathTypeTextWmf(padding: 0, faceTail: 0, penTail: 0, face: "Arial")));
    }

    // Font, pen, select, a three-glyph EXTTEXTOUT and EOF, as MathType writes them.
    private static byte[] CreateMathTypeTextWmf(
        byte padding, byte faceTail, byte penTail, byte glyph = 0x56, string face = "MT Extra")
    {
        var font = new byte[6 + 18 + 32];
        WriteUInt32(font, 0, (uint)(font.Length / 2));
        WriteUInt16(font, 4, 0x02FB);
        WriteInt16(font, 6, -384);
        WriteUInt16(font, 14, 400);
        font[19] = 2;
        var name = System.Text.Encoding.ASCII.GetBytes(face);
        Buffer.BlockCopy(name, 0, font, 24, name.Length);
        for (var index = 24 + name.Length + 1; index < font.Length; index++) font[index] = faceTail;

        var pen = new byte[6 + 10];
        WriteUInt32(pen, 0, (uint)(pen.Length / 2));
        WriteUInt16(pen, 4, 0x02FA);
        WriteInt16(pen, 8, 10);
        pen[10] = pen[11] = penTail;

        var select = new byte[8];
        WriteUInt32(select, 0, 4);
        WriteUInt16(select, 4, 0x012D);

        var text = new byte[6 + 8 + 4 + 6];
        WriteUInt32(text, 0, (uint)(text.Length / 2));
        WriteUInt16(text, 4, 0x0A32);
        WriteInt16(text, 6, 100);
        WriteInt16(text, 8, 50);
        WriteInt16(text, 10, 3);
        text[14] = glyph;
        text[15] = 0x56;
        text[16] = 0x6F;
        text[17] = padding;
        WriteInt16(text, 18, 120);
        WriteInt16(text, 20, 120);
        WriteInt16(text, 22, 80);

        var end = new byte[6];
        WriteUInt32(end, 0, 3);

        var body = font.Concat(pen).Concat(select).Concat(text).Concat(end).ToArray();
        var bytes = new byte[40 + body.Length];
        WriteUInt32(bytes, 0, 0x9ac6cdd7);
        WriteInt16(bytes, 10, 400);
        WriteInt16(bytes, 12, 200);
        WriteUInt16(bytes, 14, 1440);
        WriteUInt16(bytes, 22, 1);
        WriteUInt16(bytes, 24, 9);
        WriteUInt16(bytes, 26, 0x0300);
        WriteUInt32(bytes, 28, (uint)((18 + body.Length) / 2));
        WriteUInt16(bytes, 32, 2);
        WriteUInt32(bytes, 34, (uint)(font.Length / 2));
        Buffer.BlockCopy(body, 0, bytes, 40, body.Length);
        return bytes;
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

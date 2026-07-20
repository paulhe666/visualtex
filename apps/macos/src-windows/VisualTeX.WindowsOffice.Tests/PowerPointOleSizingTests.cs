using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class PowerPointOleSizingTests
{
    [Fact]
    public void LongerReplacementKeepsTheExistingVisualScale()
    {
        var size = PowerPointOleService.CalculateReplacementSize(
            replacementPixelWidth: 800,
            replacementPixelHeight: 100,
            oldWidth: 200f,
            oldHeight: 50f,
            originalRenderWidthPx: 400,
            originalRenderHeightPx: 100,
            replacementRenderWidthPx: 800,
            replacementRenderHeightPx: 100);

        Assert.Equal(400f, size.Width, 3);
        Assert.Equal(50f, size.Height, 3);
    }

    [Fact]
    public void TallerReplacementGrowsInsteadOfShrinkingIntoTheOldBox()
    {
        var size = PowerPointOleService.CalculateReplacementSize(
            replacementPixelWidth: 800,
            replacementPixelHeight: 200,
            oldWidth: 200f,
            oldHeight: 50f,
            originalRenderWidthPx: 400,
            originalRenderHeightPx: 100,
            replacementRenderWidthPx: 800,
            replacementRenderHeightPx: 200);

        Assert.Equal(400f, size.Width, 3);
        Assert.Equal(100f, size.Height, 3);
    }

    [Fact]
    public void LegacyFormulaUsesItsPhysicalHeightAsTheFontSizeReference()
    {
        var size = PowerPointOleService.CalculateReplacementSize(
            replacementPixelWidth: 800,
            replacementPixelHeight: 100,
            oldWidth: 200f,
            oldHeight: 50f,
            originalRenderWidthPx: null,
            originalRenderHeightPx: null,
            replacementRenderWidthPx: 800,
            replacementRenderHeightPx: 100);

        Assert.Equal(400f, size.Width, 3);
        Assert.Equal(50f, size.Height, 3);
    }
}

using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOleSizingTests
{
    [Fact]
    public void WiderEditUsesCurrentExportBoundsInsteadOfPreviousShapeBounds()
    {
        var initial = WordOleService.FitImage(480, 120, 240, 72);
        var wider = WordOleService.FitImage(720, 90, 240, 72);
        var restored = WordOleService.FitImage(480, 120, 240, 72);

        AssertClose(240, initial.Width);
        AssertClose(60, initial.Height);
        AssertClose(240, wider.Width);
        AssertClose(30, wider.Height);
        AssertClose(initial.Width, restored.Width);
        AssertClose(initial.Height, restored.Height);
    }

    [Fact]
    public void TallerEditPreservesAspectRatioWithinCurrentExportBounds()
    {
        var matrix = WordOleService.FitImage(240, 240, 240, 72);

        AssertClose(72, matrix.Width);
        AssertClose(72, matrix.Height);
    }

    private static void AssertClose(float expected, float actual) =>
        Assert.InRange(actual, expected - 0.01f, expected + 0.01f);
}

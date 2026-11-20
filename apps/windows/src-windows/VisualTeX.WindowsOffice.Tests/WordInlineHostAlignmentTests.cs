using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordInlineHostAlignmentTests
{
    [Theory]
    [InlineData(11.75f, 18.015625f, 13f)]
    [InlineData(16.3f, 24.952f, 18.432f)]
    [InlineData(60f, 80f, 3f)]
    [InlineData(6f, 8f, 7f)]
    public void ExplicitCharacterCenterDoesNotApplyPictureDescentTwice(
        float heightPt, float renderHeight, float baseline)
    {
        Assert.Equal(0, WordInlineAlignment.CalculateFontPositionForHost(
            WordInlineHostAlignment.Center, heightPt, renderHeight, baseline));
    }

    [Theory]
    [InlineData(WordInlineHostAlignment.Automatic, 11.75f, 18.015625f, 13f, -3)]
    [InlineData(WordInlineHostAlignment.Baseline, 11.75f, 18.015625f, 13f, -3)]
    [InlineData(WordInlineHostAlignment.Automatic, 16.3f, 24.952f, 18.432f, -4)]
    [InlineData(WordInlineHostAlignment.Baseline, 16.3f, 24.952f, 18.432f, -4)]
    public void BaselineAnchoredHostsRetainTheEstablishedNativeCalculation(
        WordInlineHostAlignment host, float heightPt, float renderHeight,
        float baseline, int expected)
    {
        Assert.Equal(expected, WordInlineAlignment.CalculateFontPositionForHost(
            host, heightPt, renderHeight, baseline));
    }

    [Theory]
    [InlineData(11.75f, 12d, 0)]
    [InlineData(13.5f, 12d, -1)]
    [InlineData(18.5f, 12d, -3)]
    [InlineData(24f, 12d, -6)]
    public void FarEast50UsesHalfOfHeightBeyondTheSemanticEmBox(
        float heightPt, double semanticSizePt, int expected)
    {
        Assert.Equal(expected, WordInlineAlignment.CalculateFontPositionForHost(
            WordInlineHostAlignment.FarEast50,
            heightPt,
            exportedHeight: 999f,
            exportedBaseline: 1f,
            existingFontPosition: -99f,
            sourceSemanticFontSizePoints: 48d,
            targetSemanticFontSizePoints: semanticSizePt));
    }

    [Theory]
    [InlineData(WordInlineHostAlignment.Center)]
    [InlineData(WordInlineHostAlignment.Top)]
    public void ExplicitParagraphBoxAnchoringIsNotOverriddenBySourcePosition(
        WordInlineHostAlignment host)
    {
        foreach (var oldPosition in new float[] { -15, -3, 0, 5, 25 })
            Assert.Equal(0, WordInlineAlignment.CalculateFontPositionForHost(
                host, 30, 40, 30, oldPosition, 12, 24));
    }

    [Theory]
    [InlineData(-3f)]
    [InlineData(0f)]
    [InlineData(5f)]
    public void ValidTargetRenderGeometryDoesNotInheritAnEarlierFormatBaseline(float oldPosition)
    {
        Assert.Equal(-7, WordInlineAlignment.CalculateFontPositionForHost(
            WordInlineHostAlignment.Automatic, 30, 40, 30, oldPosition, 12, 24));
        Assert.Equal(0, WordInlineAlignment.CalculateFontPositionForHost(
            WordInlineHostAlignment.Center, 30, 40, 30, oldPosition, 12, 24));
        Assert.Equal(-3, WordInlineAlignment.CalculateFontPositionForHost(
            WordInlineHostAlignment.FarEast50, 30, 40, 30, oldPosition, 12, 24));
    }

    [Fact]
    public void AnUndefinedHostIsNotSilentlyTreatedAsBaselineAligned()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            WordInlineAlignment.CalculateFontPositionForHost(
                (WordInlineHostAlignment)9999999, 15, 20, 15));
    }
}

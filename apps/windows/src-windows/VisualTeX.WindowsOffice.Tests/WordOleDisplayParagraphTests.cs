using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOleDisplayParagraphTests
{
    [Fact]
    public void CollapsedCaretInEmptyParagraphReusesThatParagraph()
    {
        Assert.True(
            WordOleService.ShouldReuseCurrentParagraphForDisplayFormula(12, 12, "\r"));
    }

    [Fact]
    public void EmptyTableCellParagraphCanAlsoBeReused()
    {
        Assert.True(
            WordOleService.ShouldReuseCurrentParagraphForDisplayFormula(12, 12, "\r\a"));
    }

    [Theory]
    [InlineData(12, 13, "\r")]
    [InlineData(12, 12, "text\r")]
    [InlineData(12, 12, " \r")]
    [InlineData(12, 12, "")]
    public void SelectionOrParagraphWithContentCreatesASeparateDisplayParagraph(
        int rangeStart,
        int rangeEnd,
        string paragraphText)
    {
        Assert.False(
            WordOleService.ShouldReuseCurrentParagraphForDisplayFormula(
                rangeStart,
                rangeEnd,
                paragraphText));
    }
}

using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOleBridge;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordEquationNumberingTests
{
    [Fact]
    public void EquationBookmarkRoundTripsPersistentFormulaId()
    {
        var formulaId = Guid.NewGuid().ToString();
        var bookmark = WordOleService.EquationBookmarkName(formulaId);

        Assert.StartsWith("VTEq_", bookmark, StringComparison.Ordinal);
        Assert.True(bookmark.Length <= 40);
        Assert.True(
            WordOleService.TryFormulaIdFromEquationBookmark(
                bookmark,
                out var decoded));
        Assert.Equal(formulaId, decoded, ignoreCase: true);
    }

    [Fact]
    public void OnlyVisualTeXSequenceFieldsAreRecognized()
    {
        Assert.True(
            WordOleService.IsVisualTeXSequenceFieldCode(
                " SEQ VisualTeXEquation \\* ARABIC "));
        Assert.False(WordOleService.IsVisualTeXSequenceFieldCode("SEQ Figure \\* ARABIC"));
        Assert.False(WordOleService.IsVisualTeXSequenceFieldCode(null));
    }

    [Fact]
    public void EquationTabsCenterFormulaAndRightAlignNumberWithinTextArea()
    {
        var positions = WordOleService.CalculateEquationTabStops(
            pageWidth: 612,
            leftMargin: 72,
            rightMargin: 72,
            leftIndent: 0,
            rightIndent: 0);

        Assert.Equal(234, positions.Center);
        Assert.Equal(468, positions.Right);
    }

    [Fact]
    public void EquationFieldIsInsertedBeforeAnExistingClosingParenthesis()
    {
        var scaffold = WordOleService.EquationNumberScaffold();

        Assert.Equal("\t()", scaffold.Text);
        Assert.Equal('(', scaffold.Text[scaffold.FieldOffset - 1]);
        Assert.Equal(')', scaffold.Text[scaffold.FieldOffset]);
    }

    [Theory]
    [InlineData(72, 12, 30)]
    [InlineData(32, 11, 11)]
    [InlineData(12, 12, 0)]
    [InlineData(8, 11, 0)]
    public void EquationNumberIsRaisedToTheFormulaVisualCenter(
        float formulaHeight,
        float fontSize,
        int expectedPosition)
    {
        Assert.Equal(
            expectedPosition,
            WordOleService.CalculateEquationNumberFontPosition(
                formulaHeight,
                fontSize));
    }

    [Fact]
    public void InvalidNumberFontSizeFallsBackToNormalWordBodySize()
    {
        Assert.Equal(
            31,
            WordOleService.CalculateEquationNumberFontPosition(72, float.NaN));
    }

    [Fact]
    public void NumberingIsRejectedForInlineMetadata()
    {
        var metadata = new FormulaMetadata
        {
            FormulaId = Guid.NewGuid().ToString(),
            DisplayMode = "inline",
            Numbered = true,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = "x" },
            },
        };

        Assert.Throws<InvalidOperationException>(metadata.Validate);
    }
}

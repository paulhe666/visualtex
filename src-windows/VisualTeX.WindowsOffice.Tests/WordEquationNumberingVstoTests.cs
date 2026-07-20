using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordEquationNumberingVstoTests
{
    [Fact]
    public void TabStopsUseTheAvailableTextWidth()
    {
        var positions = WordEquationNumbering.CalculateEquationTabStops(
            pageWidth: 612f,
            leftMargin: 72f,
            rightMargin: 72f,
            leftIndent: 0f,
            rightIndent: 0f);

        Assert.Equal(234f, positions.Center);
        Assert.Equal(468f, positions.Right);
    }

    [Theory]
    [InlineData(24f, 11f, 7)]
    [InlineData(11f, 11f, 0)]
    [InlineData(float.NaN, 11f, 0)]
    public void NumberPositionCentersTheTextBesideTheFormula(
        float formulaHeight,
        float fontSize,
        int expected)
    {
        Assert.Equal(
            expected,
            WordEquationNumbering.CalculateEquationNumberFontPosition(
                formulaHeight,
                fontSize));
    }

    [Fact]
    public void NumberScaffoldKeepsClosingParenthesisOutsideTheField()
    {
        var scaffold = WordEquationNumbering.EquationNumberScaffold();

        Assert.Equal("\t()", scaffold.Text);
        Assert.Equal(2, scaffold.FieldOffset);
    }

    [Fact]
    public void EquationBookmarkRoundTripsThePersistentFormulaId()
    {
        var formulaId = Guid.NewGuid().ToString();
        var bookmark = WordEquationNumbering.EquationBookmarkName(formulaId);

        Assert.StartsWith("VTEq_", bookmark, StringComparison.Ordinal);
        Assert.True(WordEquationNumbering.TryFormulaIdFromEquationBookmark(
            bookmark,
            out var roundTrip));
        Assert.Equal(Guid.Parse(formulaId), Guid.Parse(roundTrip));
    }

    [Theory]
    [InlineData(" SEQ VisualTeXEquation \\* ARABIC ", true)]
    [InlineData("SEQ OtherEquation \\* ARABIC", false)]
    [InlineData(null, false)]
    public void OnlyVisualTeXSequenceFieldsAreUpdated(string? code, bool expected)
    {
        Assert.Equal(expected, WordEquationNumbering.IsVisualTeXSequenceFieldCode(code));
    }

    [Theory]
    [InlineData(1f, false)]
    [InlineData(2f, false)]
    [InlineData(10.5f, true)]
    [InlineData(72f, true)]
    [InlineData(float.NaN, false)]
    public void VisibleReferenceFontSizeRejectsHiddenCaptionFormatting(
        float size,
        bool expected)
    {
        Assert.Equal(expected, WordEquationNumbering.IsNormalTextSize(size));
    }

    [Theory]
    [InlineData(" SEQ Equation \\* ARABIC ", "Equation", true)]
    [InlineData(" SEQ 公式 \\* ARABIC ", "公式", true)]
    [InlineData(" SEQ OtherEquation \\* ARABIC ", "Equation", false)]
    [InlineData(null, "Equation", false)]
    public void NativeEquationSequenceUsesTheCurrentWordCaptionLabel(
        string? code,
        string label,
        bool expected)
    {
        Assert.Equal(
            expected,
            WordEquationNumbering.IsNativeEquationSequenceFieldCode(code, label));
    }
}

using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.Tests;

// Update Equation Format's font step: which VisualTeX formulas were drawn in other fonts than the main app draws in
// now (the rule the macOS app and the shared TypeScript share: formulaFontsDiffer), and the fonts as the companion's
// preferences carry them.
public sealed class FormulaFontSettingsTests
{
    private static readonly FormulaFontSettings TimesFaces = new()
    {
        LetterFont = "katex",
        ChineseFont = "kaiti",
        Faces = new FormulaFaces { Letter = "Times New Roman", Greek = "Times New Roman", Digit = "Times New Roman" },
    };

    private static FormulaMetadata Formula(string latex, string? letterFont, string? chineseFont, FormulaFaces? faces) => new()
    {
        Latex = latex,
        FormulaLetterFont = letterFont,
        FormulaChineseFont = chineseFont,
        FormulaFaces = faces,
    };

    [Fact]
    public void AFormulaDrawnInTheSameFontsIsCurrent()
    {
        var faces = new FormulaFaces { Letter = " times new roman", Greek = "Times New Roman ", Digit = "TIMES NEW ROMAN" };
        Assert.False(TimesFaces.DifferFrom(Formula("x+y", "KaTeX", "kaiti", faces)));
    }

    [Fact]
    public void AFormulaWithoutAFacesRecordWasDrawnWithoutFaces()
    {
        Assert.True(TimesFaces.DifferFrom(Formula("x+y", "katex", "kaiti", null)));
        Assert.False(new FormulaFontSettings().DifferFrom(Formula("x+y", null, null, null)));
    }

    [Fact]
    public void AnotherMathFontOrFaceDiffers()
    {
        var faces = TimesFaces.Faces;
        Assert.True(TimesFaces.DifferFrom(Formula("x+y", "times", "kaiti", faces)));
        Assert.True(TimesFaces.DifferFrom(Formula("x+y", "katex", "kaiti",
            new FormulaFaces { Letter = "Times New Roman", Greek = "Cambria", Digit = "Times New Roman" })));
    }

    [Fact]
    public void TheChineseFontCountsOnlyForChineseText()
    {
        var faces = TimesFaces.Faces;
        Assert.False(TimesFaces.DifferFrom(Formula("x+y", "katex", "songti", faces)));
        Assert.True(TimesFaces.DifferFrom(Formula("S_{\\text{圆}}=\\pi r^2", "katex", "songti", faces)));
        Assert.False(TimesFaces.DifferFrom(Formula("S_{\\text{圆}}=\\pi r^2", "katex", "kaiti", faces)));
    }

    [Fact]
    public void PreferencesGiveTheFonts()
    {
        var fonts = FormulaFontSettings.FromPreferencesJson(
            "{\"editorPreferences\":{\"settings\":{\"formulaLetterFont\":\"times\",\"formulaChineseFont\":\"kaiti\"," +
            "\"formulaFaces\":{\"letter\":\" Times New Roman \",\"greek\":\"\",\"digit\":\"Cambria\"}}}}");
        Assert.Equal("times", fonts.LetterFont);
        Assert.Equal("kaiti", fonts.ChineseFont);
        Assert.Equal("Times New Roman", fonts.Faces.Letter);
        Assert.Equal(string.Empty, fonts.Faces.Greek);
        Assert.Equal("Cambria", fonts.Faces.Digit);
        Assert.Equal("newtxmath；字母：Times New Roman，数字：Cambria", fonts.Describe(false));
        Assert.Equal("newtxmath; letters: Times New Roman, digits: Cambria", fonts.Describe(true));
    }

    [Fact]
    public void MissingOrEarlierPreferencesKeepTheDefaults()
    {
        var fonts = FormulaFontSettings.FromPreferencesJson("{\"editorPreferences\":{\"settings\":{\"formulaLetterFont\":\"cm\"}}}");
        Assert.Equal("katex", fonts.LetterFont);
        Assert.Equal("system", fonts.ChineseFont);
        Assert.Equal("KaTeX", fonts.Describe(true));
        Assert.Equal("katex", FormulaFontSettings.FromPreferencesJson("{}").LetterFont);
    }
}

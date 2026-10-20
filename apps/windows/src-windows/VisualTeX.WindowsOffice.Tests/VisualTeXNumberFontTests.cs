using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// A VisualTeX display number takes the Word font of its formula's letters.
public sealed class VisualTeXNumberFontTests
{
    [Theory]
    [InlineData("katex", "Cambria")]
    [InlineData("cambria", "Cambria")]
    [InlineData("times", "Times New Roman")]
    public void FixedLetterFontsMapToOneWordFont(string letterFont, string expected) =>
        Assert.Equal(expected, WordVisualTeXParagraphNumbering.NumberFontName(letterFont));

    [Theory]
    [InlineData("stix", new[] { "STIX Two Text", "STIX Two Math", "Times New Roman" })]
    [InlineData("palatino", new[] { "Palatino", "Palatino Linotype", "Book Antiqua" })]
    [InlineData("helvetica", new[] { "Helvetica Neue", "Helvetica", "Arial" })]
    public void OtherLetterFontsUseAnInstalledFaceOfTheirStack(string letterFont, string[] stack) =>
        Assert.Contains(WordVisualTeXParagraphNumbering.NumberFontName(letterFont), stack);

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("comic")]
    public void UnknownLetterFontKeepsTheBodyFont(string? letterFont) =>
        Assert.Null(WordVisualTeXParagraphNumbering.NumberFontName(letterFont));
}

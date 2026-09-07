using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlImportSignatureTests
{
    private static string Math(string content) =>
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" + content + "</m:oMath>";
    private static string Signature(string content) => WordOmmlConverter.ComputeImportedOmmlContentSignature(Math(content));
    private const string Sup = "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>";

    [Fact]
    public void WordControlFormattingAndEmptyPropertiesPreserveImportedContent()
    {
        var word = Sup.Replace("<m:sSup>", "<m:sSup><m:sSupPr><m:ctrlPr><w:rPr><w:szCs w:val=\"21\"/></w:rPr></m:ctrlPr></m:sSupPr>");
        Assert.Equal(Signature(Sup), Signature(word));
        Assert.Equal(Signature(Sup), WordOmmlConverter.ComputeImportedOmmlContentSignature(
            Math(word).Replace("xmlns:m=", "xmlns:q=").Replace("m:", "q:")));
    }

    [Fact]
    public void MaterializationStoresTheActualWordFingerprintAfterProvingContent()
    {
        var prepared = Math(Sup);
        var actual = Math(Sup.Replace("<m:sSup>", "<m:sSup><m:sSupPr/>"));
        Assert.NotEqual(WordOmmlConverter.ComputeOmmlFingerprint(prepared),
            WordOmmlConverter.ComputeOmmlFingerprint(actual));
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(actual),
            WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(prepared, actual));
    }

    [Theory]
    [InlineData("x", "y")]
    [InlineData("2", "3")]
    public void MaterializationCannotAdoptChangedContent(string from, string to)
        => Assert.Throws<System.IO.InvalidDataException>(() =>
            WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(Math(Sup), Math(Sup.Replace(from, to))));

    [Theory]
    [InlineData("\u2032", "'")]
    [InlineData("\u2033", "''")]
    [InlineData("\u2034", "'''")]
    [InlineData("\u2057", "''''")]
    public void WordPrimeSpellingsHaveTheSameImportedContentSignature(
        string unicodePrime,
        string asciiPrime)
    {
        var unicode = $"<m:sSup><m:e><m:r><m:t>F</m:t></m:r></m:e><m:sup><m:r><m:t>{unicodePrime}</m:t></m:r></m:sup></m:sSup>";
        var ascii = $"<m:sSup><m:e><m:r><m:t>F</m:t></m:r></m:e><m:sup><m:r><m:t>{asciiPrime}</m:t></m:r></m:sup></m:sSup>";
        Assert.Equal(Signature(unicode), Signature(ascii));
    }

    [Fact]
    public void DefaultHatAccentMatchesWordsOmittedAccentCharacter()
    {
        const string explicitHat =
            "<m:acc><m:accPr><m:chr m:val=\"\u0302\"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>";
        const string omittedHat =
            "<m:acc><m:accPr/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>";
        const string tilde =
            "<m:acc><m:accPr><m:chr m:val=\"\u0303\"/></m:accPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>";

        Assert.Equal(Signature(explicitHat), Signature(omittedHat));
        Assert.NotEqual(Signature(explicitHat), Signature(tilde));
    }

    [Fact]
    public void InsertedMathLocatorPrefersNearestEquationBeforeLargestSpan()
    {
        Assert.True(WordOmmlConverter.ShouldPreferInsertedMathCandidate(
            distance: 1, span: 12, bestDistance: 4, bestSpan: 80));
        Assert.False(WordOmmlConverter.ShouldPreferInsertedMathCandidate(
            distance: 4, span: 80, bestDistance: 1, bestSpan: 12));
        Assert.True(WordOmmlConverter.ShouldPreferInsertedMathCandidate(
            distance: 1, span: 80, bestDistance: 1, bestSpan: 12));
        Assert.False(WordOmmlConverter.ShouldPreferInsertedMathCandidate(
            distance: 17, span: 800, bestDistance: int.MaxValue, bestSpan: -1));
    }

    private static string Integral(string glyph, string lower = "−1") =>
        "<m:nary><m:naryPr>" + glyph + "<m:limLoc m:val=\"subSup\"/></m:naryPr><m:sub><m:r><m:t>" + lower
        + "</m:t></m:r></m:sub><m:sup><m:r><m:t>1</m:t></m:r></m:sup><m:e>" + Sup + "</m:e></m:nary>";

    [Fact]
    public void DefaultIntegralAndMathematicalMinusHaveEquivalentWordEncodings()
        => Assert.Equal(Signature(Integral("<m:chr m:val=\"∫\"/>")), Signature(Integral("", "-1")));

    [Fact]
    public void ChangedOperatorOrLimitCannotClaimTheSourceIdentity()
    {
        Assert.NotEqual(Signature(Integral("")), Signature(Integral("<m:chr m:val=\"∑\"/>")));
        Assert.NotEqual(Signature(Integral("")), Signature(Integral("", "1")));
    }

    [Theory]
    [InlineData("x", "y")]
    [InlineData("2", "3")]
    [InlineData("sSup", "sSub")]
    public void FormulaContentChangesRemainDistinct(string from, string to)
        => Assert.NotEqual(Signature(Sup), Signature(Sup.Replace(from, to)));

    [Fact]
    public void OnOffLexicalFormsAreEquivalentButOppositeValuesRemainDistinct()
    {
        var root = "<m:rad><m:radPr><m:degHide m:val=\"on\"/></m:radPr><m:deg/><m:e>" + Sup + "</m:e></m:rad>";
        Assert.Equal(Signature(root), Signature(root.Replace("\"on\"", "\"1\"")));
        Assert.NotEqual(Signature(root), Signature(root.Replace("\"on\"", "\"0\"")));
    }

    [Fact]
    public void LiteralTextDoesNotFoldUnicodeMinusIntoHyphen()
    {
        var text = "<m:r><m:rPr><m:nor/></m:rPr><m:t>−</m:t></m:r>";
        Assert.NotEqual(Signature(text), Signature(text.Replace("−", "-")));
    }

    [Theory]
    [InlineData("")]
    [InlineData("<m:fPr/>")]
    [InlineData("<m:fPr><m:type/></m:fPr>")]
    public void WordOmissionOfDefaultFractionTypePreservesContent(string properties)
    {
        var operands = "<m:num>" + Sup + "</m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>";
        Assert.Equal(Signature("<m:f><m:fPr><m:type m:val=\"bar\"/></m:fPr>" + operands),
            Signature("<m:f>" + properties + operands));
    }

    [Theory]
    [InlineData("noBar")]
    [InlineData("lin")]
    [InlineData("skw")]
    public void OtherFractionTypesCannotClaimBarIdentity(string type)
    {
        var fraction = "<m:f><m:fPr><m:type m:val=\"bar\"/></m:fPr><m:num>" + Sup
            + "</m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>";
        Assert.NotEqual(Signature(fraction), Signature(fraction.Replace("\"bar\"", "\"" + type + "\"")));
        Assert.NotEqual(Signature(fraction), Signature(fraction.Replace("<m:t>2</m:t></m:r></m:den>", "<m:t>3</m:t></m:r></m:den>")));
    }

    [Fact]
    public void UnusedSingleArgumentSeparatorDoesNotChangeContent()
    {
        var delimiter = "<m:d><m:dPr><m:begChr m:val=\"{\"/><m:sepChr m:val=\",\"/><m:endChr m:val=\"\"/></m:dPr><m:e>" + Sup + "</m:e></m:d>";
        Assert.Equal(Signature(delimiter), Signature(delimiter.Replace("<m:sepChr m:val=\",\"/>", "")));
        Assert.NotEqual(Signature(delimiter), Signature(delimiter.Replace("<m:endChr m:val=\"\"/>", "<m:endChr m:val=\"}\"/>")));
        var multiple = delimiter.Replace("</m:d>", "<m:e>" + Sup + "</m:e></m:d>");
        Assert.NotEqual(Signature(multiple), Signature(multiple.Replace("<m:sepChr m:val=\",\"/>", "")));
    }

    [Theory]
    [InlineData("")]
    [InlineData("<m:mPr/>")]
    [InlineData("<m:mPr><m:baseJc/></m:mPr>")]
    public void MatrixDefaultCenterOmissionIsEquivalent(string actualProperties)
    {
        const string row = "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>";
        var expected = "<m:m><m:mPr><m:baseJc m:val=\"center\"/></m:mPr>" + row + "</m:m>";
        var actual = "<m:m>" + actualProperties + row + "</m:m>";
        Assert.Equal(Signature(expected), Signature(actual));
        Assert.NotEqual(Signature(expected), Signature(actual.Replace("<m:t>b</m:t>", "<m:t>c</m:t>")));
    }

    [Theory]
    [InlineData("top")]
    [InlineData("bot")]
    public void MatrixNondefaultVerticalAlignmentRemainsDistinct(string alignment)
    {
        var matrix = "<m:m><m:mPr><m:baseJc m:val=\"center\"/></m:mPr><m:mr><m:e>" + Sup + "</m:e></m:mr></m:m>";
        Assert.NotEqual(Signature(matrix), Signature(matrix.Replace("\"center\"", "\"" + alignment + "\"")));
    }

    [Fact]
    public void WordAsteriskSpellingIsEquivalentOnlyInMathematicalRuns()
    {
        var star = "<m:sSup><m:e><m:r><m:t>ψ</m:t></m:r></m:e><m:sup><m:r><m:t>∗</m:t></m:r></m:sup></m:sSup>";
        Assert.Equal(Signature(star), Signature(star.Replace("∗", "*")));
        Assert.NotEqual(Signature(star), Signature(star.Replace("∗", "+")));
        var literal = "<m:r><m:rPr><m:lit/></m:rPr><m:t>∗</m:t></m:r>";
        Assert.NotEqual(Signature(literal), Signature(literal.Replace("∗", "*")));
    }
}

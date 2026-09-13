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
    public void WordExplicitPlainStyleOnOperatorsAndDigitsMatchesImplicitMathDefaults()
    {
        const string prepared =
            "<m:r><m:t>E=m</m:t></m:r>"
            + "<m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e>"
            + "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>";
        const string materialized =
            "<m:r><m:t>E</m:t></m:r>"
            + "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>=</m:t></m:r>"
            + "<m:r><m:t>m</m:t></m:r>"
            + "<m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e>"
            + "<m:sup><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>2</m:t></m:r></m:sup></m:sSup>";

        Assert.Equal(Signature(prepared), Signature(materialized));
    }

    [Fact]
    public void ExplicitPlainStyleOnLettersRemainsSemantic()
    {
        const string implicitItalic = "<m:r><m:t>x</m:t></m:r>";
        const string explicitPlain =
            "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>x</m:t></m:r>";

        Assert.NotEqual(Signature(implicitItalic), Signature(explicitPlain));
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

    [Theory]
    [InlineData("ℜ", "R", "fraktur")]
    [InlineData("ℑ", "I", "fraktur")]
    [InlineData("ℒ", "L", "script")]
    [InlineData("ℝ", "R", "double-struck")]
    public void LegacyMathAlphabetGlyphMatchesWordsScriptProperty(
        string glyph,
        string baseLetter,
        string script)
    {
        var unicode = $"<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>{glyph}</m:t></m:r>";
        var word = $"<m:r><m:rPr><m:scr m:val=\"{script}\"/><m:sty m:val=\"p\"/></m:rPr><m:t>{baseLetter}</m:t></m:r>";
        Assert.Equal(Signature(unicode), Signature(word));
        Assert.NotEqual(Signature(unicode), Signature(word.Replace(script, "roman")));
    }

    [Fact]
    public void WordScriptRunCanAbsorbAdjacentPunctuationWithoutChangingContent()
    {
        const string prepared =
            "<m:r><m:rPr><m:scr m:val=\"script\"/></m:rPr><m:t>L</m:t></m:r>"
            + "<m:r><m:t>{</m:t></m:r>";
        const string materialized =
            "<m:r><m:rPr><m:scr m:val=\"script\"/></m:rPr><m:t>L{</m:t></m:r>";
        const string changed =
            "<m:r><m:rPr><m:scr m:val=\"script\"/></m:rPr><m:t>M{</m:t></m:r>";

        Assert.Equal(Signature(prepared), Signature(materialized));
        Assert.NotEqual(Signature(prepared), Signature(changed));
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

    [Fact]
    public void DisplayFractionZeroScriptLevelHintMatchesWordImport()
    {
        const string prepared = "<m:r><m:t>y=</m:t></m:r><m:f><m:fPr><m:type m:val=\"bar\"/></m:fPr>"
            + "<m:num><m:argPr><m:scrLvl m:val=\"0\"/></m:argPr><m:r><m:t>1</m:t></m:r></m:num>"
            + "<m:den><m:argPr><m:scrLvl m:val=\"0\"/></m:argPr><m:r><m:t>x−2</m:t></m:r></m:den></m:f>";
        var actual = prepared.Replace("<m:argPr><m:scrLvl m:val=\"0\"/></m:argPr>", "")
            .Replace("<m:fPr><m:type m:val=\"bar\"/></m:fPr>", "<m:fPr/>").Replace("x−2", "x-2");
        Assert.Equal(Signature(prepared), Signature(actual));
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(Math(actual)),
            WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(Math(prepared), Math(actual)));
        Assert.NotEqual(Signature(prepared), Signature(actual.Replace("x-2", "x-3")));
        Assert.NotEqual(Signature(prepared), Signature(actual.Replace("<m:t>1</m:t>", "<m:t>2</m:t>")));
    }

    [Theory]
    [InlineData("scrLvl", "1")]
    [InlineData("scrLvl", "-1")]
    [InlineData("argSz", "1")]
    [InlineData("argSz", "-2")]
    public void NonNeutralArgumentHintsCannotBeDiscarded(string name, string value)
    {
        var argument = $"<m:argPr><m:{name} m:val=\"{value}\"/></m:argPr>";
        var body = "<m:f><m:num>" + argument + "<m:r><m:t>1</m:t></m:r></m:num>"
            + "<m:den><m:r><m:t>x</m:t></m:r></m:den></m:f>";
        Assert.NotEqual(Signature(body), Signature(body.Replace(argument, "")));
    }

    [Fact]
    public void EmptyNormalTextRunCanDisappearBetweenMathOperands()
    {
        const string empty = "<m:r><m:rPr><m:nor/></m:rPr><m:t/></m:r>";
        const string prefix = "<m:r><m:t>f(x)=|x|,</m:t></m:r>";
        const string suffix = "<m:r><m:t>g(x)=</m:t></m:r>";
        var prepared = prefix + empty + suffix + Sup;
        var actual = "<m:r><m:t>f(x)=|x|,g(x)=</m:t></m:r>" + Sup;
        Assert.Equal(Signature(prepared), Signature(actual));
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(Math(actual)),
            WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(Math(prepared), Math(actual)));
        Assert.NotEqual(Signature(prepared), Signature(actual.Replace("g(x)", "h(x)")));
        Assert.NotEqual(Signature(prepared), Signature(prepared.Replace("<m:t/>", "<m:t xml:space=\"preserve\"> </m:t>")));
        Assert.NotEqual(Signature(prepared), Signature(prepared.Replace("<m:nor/>", "<m:nor/><m:aln m:val=\"1\"/>")));
    }

    [Fact]
    public void EmptyRunsNeverEraseFractionArgumentStructure()
    {
        var fraction = "<m:f><m:num><m:r><m:t/></m:r></m:num><m:den><m:r><m:t>x</m:t></m:r></m:den></m:f>";
        var emptyNumerator = fraction.Replace("<m:r><m:t/></m:r>", "");
        Assert.Equal(Signature(fraction), Signature(emptyNumerator));
        Assert.NotEqual(Signature(fraction), Signature(emptyNumerator.Replace("<m:num></m:num>", "")));
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

        var explicitVisible = "<m:rad><m:radPr><m:degHide m:val=\"off\"/></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e>" + Sup + "</m:e></m:rad>";
        Assert.Equal(
            Signature(explicitVisible),
            Signature(explicitVisible.Replace("<m:degHide m:val=\"off\"/>", "")));

        var differentialOff = "<m:box><m:boxPr><m:diff m:val=\"false\"/></m:boxPr><m:e>" + Sup + "</m:e></m:box>";
        Assert.Equal(Signature(differentialOff), Signature(differentialOff.Replace("<m:diff m:val=\"false\"/>", "")));
        Assert.NotEqual(Signature(differentialOff), Signature(differentialOff.Replace("\"false\"", "\"true\"")));

        var phantomShown = "<m:phant><m:phantPr><m:show m:val=\"true\"/></m:phantPr><m:e>" + Sup + "</m:e></m:phant>";
        Assert.Equal(Signature(phantomShown), Signature(phantomShown.Replace("<m:show m:val=\"true\"/>", "")));
        Assert.NotEqual(Signature(phantomShown), Signature(phantomShown.Replace("\"true\"", "\"false\"")));

        var naryGrowOff = "<m:nary><m:naryPr><m:grow m:val=\"false\"/></m:naryPr><m:sub/><m:sup/><m:e>" + Sup + "</m:e></m:nary>";
        Assert.Equal(Signature(naryGrowOff), Signature(naryGrowOff.Replace("<m:grow m:val=\"false\"/>", "")));
        Assert.NotEqual(Signature(naryGrowOff), Signature(naryGrowOff.Replace("\"false\"", "\"true\"")));

        var delimiterGrowOn = "<m:d><m:dPr><m:grow m:val=\"true\"/></m:dPr><m:e>" + Sup + "</m:e></m:d>";
        Assert.Equal(Signature(delimiterGrowOn), Signature(delimiterGrowOn.Replace("<m:grow m:val=\"true\"/>", "")));
        Assert.NotEqual(Signature(delimiterGrowOn), Signature(delimiterGrowOn.Replace("\"true\"", "\"false\"")));
    }

    private static string MatrixWithColumns(string columnMarkup) =>
        "<m:m><m:mPr><m:mcs>" + columnMarkup + "</m:mcs></m:mPr>"
        + "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
        + "<m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m>";

    private static string MatrixColumn(int count, string justification) =>
        $"<m:mc><m:mcPr><m:count m:val=\"{count}\"/><m:mcJc m:val=\"{justification}\"/></m:mcPr></m:mc>";

    [Fact]
    public void WordCoalescedIdenticalMatrixColumnGroupsPreserveImportedContent()
    {
        var prepared = MatrixWithColumns(MatrixColumn(1, "left") + MatrixColumn(1, "left"));
        var materialized = MatrixWithColumns(MatrixColumn(2, "left"));

        Assert.Equal(Signature(prepared), Signature(materialized));
        Assert.Equal(
            WordOmmlConverter.ComputeOmmlFingerprint(Math(materialized)),
            WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(Math(prepared), Math(materialized)));
    }

    [Fact]
    public void MatrixColumnAlignmentAndCountChangesRemainSemantic()
    {
        var twoLeft = MatrixWithColumns(MatrixColumn(2, "left"));
        var leftRight = MatrixWithColumns(MatrixColumn(1, "left") + MatrixColumn(1, "right"));
        var oneLeft = MatrixWithColumns(MatrixColumn(1, "left"));

        Assert.NotEqual(Signature(twoLeft), Signature(leftRight));
        Assert.NotEqual(Signature(twoLeft), Signature(oneLeft));
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

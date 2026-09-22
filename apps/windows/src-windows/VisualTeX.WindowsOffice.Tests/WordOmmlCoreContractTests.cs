using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlCoreContractTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string P = "http://www.w3.org/1998/Math/MathML";
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static string Math(string body) => $"<m:oMath xmlns:m=\"{M}\">{body}</m:oMath>";
    private static string Mml(string body) => $"<math xmlns=\"{P}\" display=\"block\">{body}</math>";
    private static string Run(string text) => $"<m:r><m:t>{text}</m:t></m:r>";
    private static string Flag(string name, string? value) =>
        value is null ? $"<m:{name}/>" : $"<m:{name} m:val=\"{value}\"/>";

    [Theory]
    [InlineData("1")]
    [InlineData("on")]
    [InlineData("true")]
    [InlineData(null)]
    public void BothValidationStagesAgreeOnHiddenOptionalArguments(string? value)
    {
        var xml = Math("<m:rad><m:radPr>" + Flag("degHide", value) + "</m:radPr><m:deg/><m:e>" + Run("x") + "</m:e></m:rad>"
            + "<m:nary><m:naryPr>" + Flag("subHide", value) + Flag("supHide", value)
            + "</m:naryPr><m:sub/><m:sup/><m:e>" + Run("y") + "</m:e></m:nary>"
            + "<m:m><m:mPr>" + Flag("plcHide", value) + "</m:mPr><m:mr><m:e/><m:e>" + Run("z") + "</m:e></m:mr></m:m>");
        WordOmmlConverter.ValidateNoVisibleEmptyOmmlSlots(XDocument.Parse(xml));
        WordOmmlConverter.ValidateMaterializedOmml(xml);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("off")]
    [InlineData("false")]
    public void BothValidationStagesRejectVisibleEmptyRadicalDegree(string value)
    {
        var xml = Math("<m:rad><m:radPr>" + Flag("degHide", value) + "</m:radPr><m:deg/><m:e>" + Run("x") + "</m:e></m:rad>");
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateNoVisibleEmptyOmmlSlots(XDocument.Parse(xml)));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateMaterializedOmml(xml));
    }

    [Theory]
    [InlineData("rad", "deg", "degHide")]
    [InlineData("nary", "sub", "subHide")]
    [InlineData("nary", "sup", "supHide")]
    public void MissingHidePropertyDoesNotEnableHiding(string owner, string slot, string property)
    {
        var xml = Math($"<m:{owner}><m:{owner}Pr/><m:{slot}/><m:e>" + Run("x") + $"</m:e></m:{owner}>");
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateNoVisibleEmptyOmmlSlots(XDocument.Parse(xml)));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateMaterializedOmml(xml));
        Assert.DoesNotContain(property, xml, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("<msqrt><mi>x</mi></msqrt>")]
    [InlineData("<mfrac><msqrt><mi>x</mi></msqrt><mi>y</mi></mfrac>")]
    [InlineData("<msqrt><msqrt><mi>x</mi></msqrt></msqrt>")]
    [InlineData("<mroot><mi>x</mi><mn>3</mn></mroot>")]
    public void RootsUseSameCoreContractInsideAndOutsideMatrices(string root)
    {
        var matrix = "<mtable columnalign=\"center center\"><mtr><mtd>" + root
            + "</mtd><mtd><mi>y</mi></mtd></mtr><mtr><mtd><mn>0</mn></mtd><mtd><mn>1</mn></mtd></mtr></mtable>";
        foreach (var body in new[] { root, matrix, "<mfenced open=\"(\" close=\")\">" + matrix + "</mfenced>",
                     "<mfenced open=\"{\" close=\"\">" + matrix + "</mfenced>" })
        {
            var omml = WordOmmlConverter.TransformMathMlToOmml(Mml(body));
            WordOmmlConverter.ValidateMaterializedOmml(omml);
        }
    }

    [Theory]
    [InlineData("<m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den/></m:f>")]
    [InlineData("<m:f><m:num><m:r><m:t>x</m:t></m:r></m:num></m:f>")]
    [InlineData("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup/></m:sSup>")]
    [InlineData("<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e></m:sSup>")]
    public void EmptyOrMissingMandatoryArgumentsRemainRejected(string body)
    {
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateMaterializedOmml(Math(body)));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateOmmlResult(Math(body), Mml("<mi>x</mi>")));
    }

    [Theory]
    [InlineData("", "merror")]
    [InlineData("p:", "p:merror")]
    public void MathMlErrorNodesCannotBypassValidationWithNamespaceAliases(string prefix, string errorName)
    {
        var declaration = prefix.Length == 0 ? $"xmlns=\"{P}\"" : $"xmlns:p=\"{P}\"";
        var xml = $"<{prefix}math {declaration}><{errorName}><{prefix}mtext>bad</{prefix}mtext></{errorName}></{prefix}math>";
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateMathMlForOmml(xml));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.TransformMathMlToOmml(xml));
    }

    [Theory]
    [InlineData("<mtext>C:\\Users</mtext>", "C:\\Users")]
    [InlineData("<mtext mathcolor=\"red\">important</mtext>", "important")]
    [InlineData("<mtext>\\alpha</mtext>", "\\alpha")]
    public void PresentationMathMlLiteralTextIsNotReparsedAsLatex(string body, string expected)
    {
        var xml = WordOmmlConverter.TransformMathMlToOmml(Mml(body));
        Assert.Contains(expected, string.Concat(XDocument.Parse(xml).Descendants(XName.Get("t", M)).Select(e => e.Value)), StringComparison.Ordinal);
        WordOmmlConverter.ValidateMaterializedOmml(xml);
    }

    [Fact]
    public void ExtractSingleEquationRejectsMultipleOrNestedEquations()
    {
        var first = Math(Run("x"));
        var second = Math(Run("y"));
        var multiple = $"<m:oMathPara xmlns:m=\"{M}\">{first}{second}</m:oMathPara>";
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ExtractSingleOMath(multiple));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ExtractSingleOMath(Math(first)));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(first, multiple));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateMaterializedOmml(multiple));
    }

    [Theory]
    [InlineData("<m:begChr m:val=\"(\"/><m:endChr m:val=\")\"/>")]
    [InlineData("<m:begChr m:val=\"(\"/>")]
    [InlineData("<m:endChr m:val=\")\"/>")]
    public void ExplicitDefaultDelimitersMatchOmittedProperties(string properties)
    {
        var a = Math("<m:d><m:dPr>" + properties + "</m:dPr><m:e>" + Run("x") + "</m:e></m:d>");
        var b = Math("<m:d><m:e>" + Run("x") + "</m:e></m:d>");
        Assert.Equal(WordOmmlConverter.ComputeImportedOmmlContentSignature(a), WordOmmlConverter.ComputeImportedOmmlContentSignature(b));
    }

    [Theory]
    [InlineData("begChr")]
    [InlineData("endChr")]
    public void EmptyDelimiterIsNotTheDefaultParenthesis(string name)
    {
        var omitted = Math("<m:d><m:e>" + Run("x") + "</m:e></m:d>");
        var noValue = Math("<m:d><m:dPr><m:" + name + "/></m:dPr><m:e>" + Run("x") + "</m:e></m:d>");
        var emptyValue = noValue.Replace("<m:" + name + "/>", "<m:" + name + " m:val=\"\"/>");
        Assert.NotEqual(WordOmmlConverter.ComputeImportedOmmlContentSignature(omitted), WordOmmlConverter.ComputeImportedOmmlContentSignature(noValue));
        Assert.Equal(WordOmmlConverter.ComputeImportedOmmlContentSignature(noValue), WordOmmlConverter.ComputeImportedOmmlContentSignature(emptyValue));
    }

    [Theory]
    [InlineData("<m:sty m:val=\"i\"/>")]
    [InlineData("<m:sty/>")]
    public void DefaultItalicStyleIsAnEquivalentEncoding(string style)
    {
        Assert.Equal(WordOmmlConverter.ComputeImportedOmmlContentSignature(Math(Run("x"))),
            WordOmmlConverter.ComputeImportedOmmlContentSignature(Math("<m:r><m:rPr>" + style + "</m:rPr><m:t>x</m:t></m:r>")));
    }

    [Theory]
    [InlineData("p")]
    [InlineData("b")]
    [InlineData("bi")]
    public void NonDefaultLetterStylesStillChangeContent(string style)
    {
        Assert.NotEqual(WordOmmlConverter.ComputeImportedOmmlContentSignature(Math(Run("x"))),
            WordOmmlConverter.ComputeImportedOmmlContentSignature(Math("<m:r><m:rPr><m:sty m:val=\"" + style + "\"/></m:rPr><m:t>x</m:t></m:r>")));
    }

    [Fact]
    public void RedundantXmlSpaceMayChangeButTextWhitespaceMayNot()
    {
        var plain = Math(Run("x"));
        var annotated = Math("<m:r><m:t xml:space=\"preserve\">x</m:t></m:r>");
        var spaced = Math("<m:r><m:t xml:space=\"preserve\"> x</m:t></m:r>");
        Assert.Equal(WordOmmlConverter.ComputeImportedOmmlContentSignature(plain), WordOmmlConverter.ComputeImportedOmmlContentSignature(annotated));
        Assert.NotEqual(WordOmmlConverter.ComputeImportedOmmlContentSignature(plain), WordOmmlConverter.ComputeImportedOmmlContentSignature(spaced));
    }

    [Fact]
    public void WrongEquationIsStillRejectedAfterNormalization()
    {
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(Math(Run("x")), Math(Run("y"))));
    }

    [Theory]
    [InlineData("SEQ VisualTeXEquation")]
    [InlineData("seq  VisualTeXEquation \\* ARABIC")]
    [InlineData("SEQ &quot;VisualTeXEquation&quot;")]
    public void SimpleNumberFieldsAreNotProvedAbsent(string instruction)
    {
        var xml = $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:fldSimple w:instr=\"{instruction}\">"
            + "<w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:body></w:document>";
        Assert.False(WordDocumentXml.CanProveNoVisualTeXEquationNumberFields(xml));
    }

    [Fact]
    public void AdjacentComplexFieldInstructionsCannotHideVisualTeXSequence()
    {
        var xml = $"<w:document xmlns:w=\"{W}\"><w:body><w:p>"
            + "<w:r><w:instrText>REF Other</w:instrText></w:r>"
            + "<w:r><w:instrText>SEQ Visual</w:instrText></w:r>"
            + "<w:r><w:instrText>TeXEquation</w:instrText></w:r>"
            + "<w:r><w:instrText>SEQ Other</w:instrText></w:r>"
            + "</w:p></w:body></w:document>";
        Assert.False(WordDocumentXml.CanProveNoVisualTeXEquationNumberFields(xml));
    }

    [Fact]
    public void MalformedSimpleFieldCannotProveAbsence()
    {
        var xml = $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:fldSimple>"
            + "<w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:body></w:document>";
        Assert.False(WordDocumentXml.CanProveNoVisualTeXEquationNumberFields(xml));
        Assert.False(WordDocumentXml.CanProveNoMathTypePlaceRefFields(xml));
    }

    [Fact]
    public void NamespaceAliasesUseTheSameFencedTableNormalization()
    {
        var source = Mml("<mrow><mo data-mjx-texclass=\"OPEN\">{</mo><mtable>"
            + "<mtr><mtd><msqrt><mi>x</mi></msqrt></mtd><mtd><mn>1</mn></mtd></mtr>"
            + "<mtr><mtd><mi>y</mi></mtd><mtd><mn>0</mn></mtd></mtr>"
            + "</mtable><mo data-mjx-texclass=\"CLOSE\"></mo></mrow>");
        var prefixed = source.Replace("xmlns=", "xmlns:p=");
        prefixed = System.Text.RegularExpressions.Regex.Replace(prefixed, @"<(\/?)([A-Za-z][A-Za-z0-9]*)(?=[\s/>])", "<$1p:$2");
        var a = WordOmmlConverter.TransformMathMlToOmml(source);
        var b = WordOmmlConverter.TransformMathMlToOmml(prefixed);
        Assert.Equal(WordOmmlConverter.ComputeImportedOmmlContentSignature(a), WordOmmlConverter.ComputeImportedOmmlContentSignature(b));
        Assert.Single(XDocument.Parse(b).Descendants(XName.Get("d", M)));
        WordOmmlConverter.ValidateMaterializedOmml(b);
    }

    [Fact]
    public void SimpleMathTypeReferenceFieldsAreNotProvedAbsent()
    {
        var xml = $"<w:document xmlns:w=\"{W}\"><w:body><w:p><w:fldSimple w:instr=\"MTPlaceRef 123\">"
            + "<w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:body></w:document>";
        Assert.False(WordDocumentXml.CanProveNoMathTypePlaceRefFields(xml));
    }

    [Theory]
    [InlineData("<msup><mrow/><mn>14</mn></msup><mi>C</mi>", "", "14", "C")]
    [InlineData("<msub><mi/><mi>i</mi></msub><mi>A</mi>", "i", "", "A")]
    [InlineData("<msubsup><mrow/><mn>6</mn><mn>14</mn></msubsup><mi>C</mi>", "6", "14", "C")]
    public void EmptyBaseFollowedByAnAtomBecomesNativePrescript(string source, string sub, string sup, string basis)
    {
        foreach (var body in new[] { source, "<mtable columnalign=\"center center\"><mtr><mtd>" + source
                     + "</mtd><mtd><mn>1</mn></mtd></mtr></mtable>" })
        {
            var xml = WordOmmlConverter.TransformMathMlToOmml(Mml(body));
            var tree = XDocument.Parse(xml);
            XNamespace m = M;
            var prescript = Assert.Single(tree.Descendants(m + "sPre"));
            Assert.Equal(sub, prescript.Element(m + "sub")?.Value);
            Assert.Equal(sup, prescript.Element(m + "sup")?.Value);
            Assert.Equal(basis, prescript.Element(m + "e")?.Value);
            Assert.DoesNotContain(tree.Descendants(m + "t"), text => text.Value.Contains('\u200B'));
            WordOmmlConverter.ValidateMaterializedOmml(xml);
        }
    }

    [Fact]
    public void PrescriptRetainsFollowingAtomsOwnPostscript()
    {
        var source = "<msup><mrow/><mn>14</mn></msup><msub><mi>C</mi><mi>n</mi></msub>";
        var xml = WordOmmlConverter.TransformMathMlToOmml(Mml(source));
        XNamespace m = M;
        var prescript = Assert.Single(XDocument.Parse(xml).Descendants(m + "sPre"));
        var postscript = Assert.Single(prescript.Element(m + "e")!.Descendants(m + "sSub"));
        Assert.Equal("C", postscript.Element(m + "e")?.Value);
        Assert.Equal("n", postscript.Element(m + "sub")?.Value);
    }

    [Theory]
    [InlineData("<msup><mrow/><mn>2</mn></msup><mo>+</mo><mi>x</mi>")]
    [InlineData("<mfrac><msup><mrow/><mn>2</mn></msup><mi>x</mi></mfrac>")]
    [InlineData("<mtable><mtr><mtd><msup><mrow/><mn>2</mn></msup></mtd><mtd><mi>x</mi></mtd></mtr></mtable>")]
    public void PrescriptConversionNeverStealsAnOperatorOrCrossesArgumentBoundary(string source)
    {
        // An incomplete empty-base script is still rejected. A following operator,
        // fraction denominator or neighbouring table cell cannot become its base.
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.TransformMathMlToOmml(Mml(source)));
    }

    [Theory]
    [InlineData("<m:degHide m:val=\"bogus\"/>")]
    [InlineData("<m:degHide m:val=\"\"/>")]
    public void InvalidBooleanCannotBeSilentlyNormalizedToFalse(string property)
    {
        var xml = Math("<m:rad><m:radPr>" + property + "</m:radPr><m:deg/><m:e>" + Run("x") + "</m:e></m:rad>");
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ComputeImportedOmmlContentSignature(xml));
    }
}

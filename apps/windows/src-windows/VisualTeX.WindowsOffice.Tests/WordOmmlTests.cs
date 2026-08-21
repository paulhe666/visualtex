using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlTests
{
    private const string MathNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void OmmlMetadataPartRoundTripsOriginalLatexAndUuid()
    {
        var metadata = Metadata();

        var xml = WordOmmlFormulaStore.BuildPartXml(metadata);

        Assert.Contains(WordOmmlFormulaStore.NamespaceUri, xml, StringComparison.Ordinal);
        Assert.True(WordOmmlFormulaStore.TryDecodePartXml(xml, out var decoded));
        Assert.Equal(metadata.FormulaId, decoded.FormulaId);
        Assert.Equal(metadata.Latex, decoded.Latex);
        Assert.Equal(metadata.Lines[0].Latex, decoded.Lines[0].Latex);
        Assert.True(decoded.Numbered);
    }

    [Fact]
    public void OmmlMetadataPartRejectsMismatchedFormulaId()
    {
        var metadata = Metadata();
        var xml = WordOmmlFormulaStore.BuildPartXml(metadata)
            .Replace(metadata.FormulaId, Guid.NewGuid().ToString(), StringComparison.Ordinal);

        Assert.False(WordOmmlFormulaStore.TryDecodePartXml(xml, out _));
    }

    [Fact]
    public void ExtractSingleOMathAcceptsMathParagraphAndBuildsMinimalDocxXml()
    {
        var wrapped =
            $"<m:oMathPara xmlns:m=\"{MathNamespace}\"><m:oMath>"
            + "<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num>"
            + "<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>"
            + "</m:oMath></m:oMathPara>";

        var equation = WordOmmlConverter.ExtractSingleOMath(wrapped);
        var documentXml = WordOmmlConverter.BuildDocumentXml(equation);

        Assert.StartsWith("<m:oMath", equation, StringComparison.Ordinal);
        Assert.Contains("<m:f>", equation, StringComparison.Ordinal);
        Assert.Contains("<w:document", documentXml, StringComparison.Ordinal);
        Assert.Contains(equation, documentXml, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeMathMlTransformRejectsLiteralUnresolvedLatexCommands()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mtext mathcolor=\"red\">\\bm</mtext><mi>v</mi></math>";

        var error = Assert.Throws<InvalidDataException>(() =>
            WordOmmlConverter.TransformMathMlToOmml(mathMl));

        Assert.Contains("unresolved LaTeX command", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OmmlRoundTripKeepsOrdinaryEulerLettersAsMathIdentifiers()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup>"
            + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var identifiers = document.Descendants(presentationMath + "mi")
            .Select(element => element.Value)
            .ToArray();

        Assert.Contains("e", identifiers);
        Assert.Contains("i", identifiers);
        Assert.Contains("π", identifiers);
        Assert.DoesNotContain(document.Descendants(presentationMath + "mtext"),
            element => element.Value is "e" or "i");
    }

    [Theory]
    [InlineData("lim")]
    [InlineData("max")]
    [InlineData("min")]
    [InlineData("sup")]
    [InlineData("inf")]
    [InlineData("limsup")]
    [InlineData("liminf")]
    public void WordMaterializedLimitOperatorKeepsUprightBaseGrouped(string functionName)
    {
        // This fixture mirrors the actual Word-saved structure produced after
        // OMath.BuildUp(): Word stores a standard operator in m:limLow but may
        // discard the original m:nor flag from its base run.
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:limLow><m:limLowPr/><m:e><m:r>"
            + $"<m:t>{functionName}</m:t>"
            + "</m:r></m:e><m:lim><m:r><m:t>x→0</m:t></m:r></m:lim></m:limLow>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: true);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var script = document
            .Descendants()
            .First(element => element.Name == presentationMath + "msub"
                || element.Name == presentationMath + "munder");
        var baseToken = script.Elements().First();

        Assert.Equal(presentationMath + "mi", baseToken.Name);
        Assert.Equal(functionName, baseToken.Value);
        Assert.Equal("normal", baseToken.Attribute("mathvariant")?.Value);
    }

    [Theory]
    [InlineData("sin")]
    [InlineData("cos")]
    [InlineData("log")]
    [InlineData("abc")]
    public void WordMaterializedPlainScriptBaseKeepsSingleUprightToken(string baseText)
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:sSup><m:e><m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr>"
            + $"<m:t>{baseText}</m:t></m:r></m:e>"
            + "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var baseToken = document.Descendants(presentationMath + "msup").First().Elements().First();

        Assert.Equal(presentationMath + "mi", baseToken.Name);
        Assert.Equal(baseText, baseToken.Value);
        Assert.Equal("normal", baseToken.Attribute("mathvariant")?.Value);
    }

    [Theory]
    [InlineData("log")]
    [InlineData("ln")]
    [InlineData("exp")]
    [InlineData("sin")]
    public void WordMaterializedFunctionApplicationKeepsPlainRunGrouped(string functionName)
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr>"
            + $"<m:t>{functionName}</m:t></m:r>"
            + "<m:r><m:t>⁡x</m:t></m:r>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        Assert.Contains(
            document.Descendants(presentationMath + "mi"),
            token => token.Value == functionName
                && token.Attribute("mathvariant")?.Value == "normal");
    }

    [Fact]
    public void WordMaterializedPlainRunWithoutFunctionApplicationIsNotForcedIntoFunctionToken()
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>abc</m:t></m:r>"
            + "<m:r><m:t>+x</m:t></m:r>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        Assert.DoesNotContain(
            document.Descendants(presentationMath + "mi"),
            token => token.Value == "abc"
                && token.Attribute("mathvariant")?.Value == "normal");
    }

    [Theory]
    [InlineData("∀")]
    [InlineData("∃")]
    [InlineData("¬")]
    public void OmmlReverseRestoresPureMathSymbolTextAsOperator(string symbol)
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:r><m:rPr><m:nor/></m:rPr>"
            + $"<m:t>{symbol}</m:t></m:r>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        Assert.Contains(
            document.Descendants(presentationMath + "mo"),
            token => token.Value == symbol);
        Assert.DoesNotContain(
            document.Descendants(presentationMath + "mtext"),
            token => token.Value == symbol);
    }

    [Fact]
    public void OmmlReverseDoesNotReclassifyOrdinaryTextAsOperator()
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:r><m:rPr><m:nor/></m:rPr><m:t>if</m:t></m:r>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        Assert.DoesNotContain(
            document.Descendants(presentationMath + "mo"),
            token => token.Value == "if");
    }

    [Fact]
    public void MathTypePileBinomialTransformsToNativeNoBarFractionOmml()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mfenced open=\"(\" close=\")\" separators=\"\"><mrow>"
            + "<mtable data-mtef-pile=\"true\">"
            + "<mtr><mtd><mi>n</mi></mtd></mtr>"
            + "<mtr><mtd><mi>k</mi></mtd></mtr>"
            + "</mtable></mrow></mfenced></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace officeMath = MathNamespace;
        var fraction = document.Descendants(officeMath + "f").Single();

        Assert.Equal(
            "noBar",
            fraction.Element(officeMath + "fPr")?
                .Element(officeMath + "type")?
                .Attribute(officeMath + "val")?
                .Value);
        Assert.Empty(document.Descendants(officeMath + "m"));
    }

    [Fact]
    public void ExplicitColumnMatrixDoesNotUseMathTypePileBinomialNormalization()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mfenced open=\"(\" close=\")\" separators=\"\"><mrow>"
            + "<mtable><mtr><mtd><mi>n</mi></mtd></mtr>"
            + "<mtr><mtd><mi>k</mi></mtd></mtr></mtable>"
            + "</mrow></mfenced></math>";

        var normalized = WordOmmlConverter.NormalizeMathTypeBinomialPiles(mathMl);
        var document = XDocument.Parse(normalized);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        Assert.Single(document.Descendants(presentationMath + "mtable"));
        Assert.Empty(document.Descendants(presentationMath + "mfrac"));
    }

    [Fact]
    public void OmmlReverseRestoresNoBarFractionThicknessForBinomialSemantics()
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:r><m:t>(</m:t></m:r>"
            + "<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr>"
            + "<m:num><m:r><m:t>n</m:t></m:r></m:num>"
            + "<m:den><m:r><m:t>k</m:t></m:r></m:den></m:f>"
            + "<m:r><m:t>)</m:t></m:r>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var fraction = document.Descendants(presentationMath + "mfrac").Single();
        var children = document.Root!.Elements().ToArray();

        Assert.Equal("0", fraction.Attribute("linethickness")?.Value);
        Assert.Equal(3, children.Length);
        Assert.Equal("mo", children[0].Name.LocalName);
        Assert.Equal("(", children[0].Value);
        Assert.Same(fraction, children[1]);
        Assert.Equal("mo", children[2].Name.LocalName);
        Assert.Equal(")", children[2].Value);
    }

    [Fact]
    public void OmmlReverseDoesNotZeroOrdinaryFractionThickness()
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:f><m:fPr/>"
            + "<m:num><m:r><m:t>a</m:t></m:r></m:num>"
            + "<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var fraction = document.Descendants(presentationMath + "mfrac").Single();

        Assert.Null(fraction.Attribute("linethickness"));
    }

    [Fact]
    public void WordMaterializedLimitDoesNotJoinArbitraryVariableName()
    {
        var omml =
            $"<m:oMath xmlns:m=\"{MathNamespace}\">"
            + "<m:limLow><m:limLowPr/><m:e><m:r><m:t>abc</m:t></m:r></m:e>"
            + "<m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limLow>"
            + "</m:oMath>";

        var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
        var document = XDocument.Parse(roundTrip);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        Assert.DoesNotContain(
            document.Descendants(presentationMath + "mi"),
            token => token.Value == "abc"
                && token.Attribute("mathvariant")?.Value == "normal");
    }

    [Fact]
    public void OfficeMathMlTransformProducesNativeFractionAndSuperscript()
    {
        var transformPath = WordOmmlConverter.ResolveTransformPath();
        Assert.True(File.Exists(transformPath));
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mfrac><mi>a</mi><mi>b</mi></mfrac><mo>+</mo>"
            + "<msup><mi>x</mi><mn>2</mn></msup></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);

        Assert.Contains("<m:f>", omml, StringComparison.Ordinal);
        Assert.Contains("<m:sSup>", omml, StringComparison.Ordinal);
        Assert.DoesNotContain("<math", omml, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeMathMlTransformRemovesVisualTeXTypingAnchorButKeepsFollowingDigit()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<msup><mi>c</mi><mn>2</mn></msup>"
            + "<mrow data-mjx-texclass=\"ORD\"><mo>&#x200C;</mo></mrow>"
            + "<mn>1</mn></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var visibleText = string.Concat(document.Descendants(math + "t").Select(node => node.Value));

        Assert.DoesNotContain("\u200C", omml, StringComparison.Ordinal);
        Assert.DoesNotContain("200C", omml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1", visibleText, StringComparison.Ordinal);
    }

    [Fact]
    public void AlignedMathMlPreservesRightLeftAlignmentAroundAmpersandColumn()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mtable displaystyle=\"true\" columnalign=\"right left\" columnspacing=\"0em\">"
            + "<mtr><mtd><msub><mi>R</mi><mi>o</mi></msub></mtd>"
            + "<mtd><mi></mi><mo>=</mo><mfrac><mi>v</mi><mi>i</mi></mfrac></mtd></mtr>"
            + "<mtr><mtd><msub><mi>A</mi><mi>v</mi></msub></mtd>"
            + "<mtd><mi></mi><mo>=</mo><mi>g</mi></mtd></mtr>"
            + "</mtable></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var matrix = document.Descendants(math + "m").Single();
        var columns = matrix
            .Element(math + "mPr")?
            .Element(math + "mcs")?
            .Elements(math + "mc")
            .ToArray();

        Assert.NotNull(columns);
        Assert.Equal(2, columns!.Length);
        Assert.Equal(
            new[] { "right", "left" },
            columns.Select(column =>
                column.Element(math + "mcPr")?
                    .Element(math + "mcJc")?
                    .Attribute(math + "val")?
                    .Value));
        Assert.All(columns, column => Assert.Equal(
            "1",
            column.Element(math + "mcPr")?
                .Element(math + "count")?
                .Attribute(math + "val")?
                .Value));
        Assert.All(
            matrix.Elements(math + "mr"),
            row => Assert.Equal(2, row.Elements(math + "e").Count()));
    }

    [Fact]
    public void BoldItalicMathMlIdentifiersUseNativeBoldItalicOmmlRuns()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>A</mi><mi mathvariant=\"bold-italic\">v</mi><mo>=</mo>"
            + "<mi mathvariant=\"bold-italic\">&#x3BB;</mi></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var boldItalicRuns = document.Descendants(math + "r")
            .Where(run =>
                run.Element(math + "rPr")?
                    .Element(math + "sty")?
                    .Attribute(math + "val")?
                    .Value == "bi")
            .Select(run => run.Element(math + "t")?.Value)
            .ToArray();

        Assert.Contains("v", boldItalicRuns);
        Assert.Contains("λ", boldItalicRuns);
    }

    [Fact]
    public void VmatrixMathMlBecomesOneNativeDelimiterContainingAThreeByThreeMatrix()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi mathvariant=\"normal\">&#x2207;</mi><mo>&#xD7;</mo>"
            + "<mi mathvariant=\"bold-italic\">F</mi><mo>=</mo>"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">|</mo>"
            + "<mtable columnspacing=\"1em\" rowspacing=\"4pt\">"
            + "<mtr><mtd><msub><mi mathvariant=\"bold-italic\">e</mi><mi>x</mi></msub></mtd>"
            + "<mtd><msub><mi mathvariant=\"bold-italic\">e</mi><mi>y</mi></msub></mtd>"
            + "<mtd><msub><mi mathvariant=\"bold-italic\">e</mi><mi>z</mi></msub></mtd></mtr>"
            + "<mtr><mtd><msub><mi>&#x2202;</mi><mi>x</mi></msub></mtd>"
            + "<mtd><msub><mi>&#x2202;</mi><mi>y</mi></msub></mtd>"
            + "<mtd><msub><mi>&#x2202;</mi><mi>z</mi></msub></mtd></mtr>"
            + "<mtr><mtd><msub><mi>F</mi><mi>x</mi></msub></mtd>"
            + "<mtd><msub><mi>F</mi><mi>y</mi></msub></mtd>"
            + "<mtd><msub><mi>F</mi><mi>z</mi></msub></mtd></mtr>"
            + "</mtable><mo data-mjx-texclass=\"CLOSE\">|</mo></mrow><mo>.</mo></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var delimiter = document.Descendants(math + "d").SingleOrDefault();
        Assert.NotNull(delimiter);
        Assert.Equal(
            "|",
            delimiter!.Element(math + "dPr")?
                .Element(math + "begChr")?
                .Attribute(math + "val")?
                .Value);
        Assert.Equal(
            "|",
            delimiter.Element(math + "dPr")?
                .Element(math + "endChr")?
                .Attribute(math + "val")?
                .Value);
        var matrix = delimiter.Descendants(math + "m").SingleOrDefault();
        Assert.NotNull(matrix);
        var rows = matrix!.Elements(math + "mr").ToArray();
        Assert.Equal(3, rows.Length);
        Assert.All(rows, row => Assert.Equal(3, row.Elements(math + "e").Count()));
        Assert.DoesNotContain(
            rows.SelectMany(row => row.Elements(math + "e")),
            cell => !cell.Descendants(math + "t").Any(text =>
                !string.IsNullOrWhiteSpace(text.Value)));
    }

    [Fact]
    public void CasesMathMlBecomesOneSidedNativeDelimiterThatGrowsWithItsTable()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>f</mi><mo stretchy=\"false\">(</mo><mi>x</mi><mo stretchy=\"false\">)</mo><mo>=</mo>"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">{</mo>"
            + "<mtable columnalign=\"left left\" columnspacing=\"1em\" rowspacing=\".2em\">"
            + "<mtr><mtd><msup><mi>x</mi><mn>2</mn></msup><mo>,</mo></mtd>"
            + "<mtd><mi>x</mi><mo>&gt;</mo><mn>0</mn></mtd></mtr>"
            + "<mtr><mtd><mn>0</mn><mo>,</mo></mtd>"
            + "<mtd><mi>x</mi><mo>=</mo><mn>0</mn></mtd></mtr>"
            + "</mtable><mo data-mjx-texclass=\"CLOSE\" fence=\"true\" stretchy=\"true\" symmetric=\"true\"></mo>"
            + "</mrow></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var delimiter = document.Descendants(math + "d").SingleOrDefault();

        Assert.NotNull(delimiter);
        Assert.Equal(
            "{",
            delimiter!.Element(math + "dPr")?
                .Element(math + "begChr")?
                .Attribute(math + "val")?
                .Value);
        Assert.Equal(
            string.Empty,
            delimiter.Element(math + "dPr")?
                .Element(math + "endChr")?
                .Attribute(math + "val")?
                .Value ?? string.Empty);
        Assert.NotEqual(
            "0",
            delimiter.Element(math + "dPr")?
                .Element(math + "grow")?
                .Attribute(math + "val")?
                .Value);
        var matrix = delimiter.Descendants(math + "m").SingleOrDefault();
        Assert.NotNull(matrix);
        Assert.Equal(2, matrix!.Elements(math + "mr").Count());
        Assert.All(matrix.Elements(math + "mr"), row =>
            Assert.Equal(2, row.Elements(math + "e").Count()));
        var alignments = matrix
            .Element(math + "mPr")?
            .Element(math + "mcs")?
            .Elements(math + "mc")
            .Select(column => column
                .Element(math + "mcPr")?
                .Element(math + "mcJc")?
                .Attribute(math + "val")?
                .Value)
            .ToArray();
        Assert.Equal(new[] { "left", "left" }, alignments);
    }

    [Fact]
    public void LeftOnlyBraceAroundOneColumnTableBecomesNativeGrowingDelimiter()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">{</mo>"
            + "<mtable columnalign=\"left\"><mtr><mtd><mi>a</mi></mtd></mtr>"
            + "<mtr><mtd><mi>b</mi></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\" fence=\"true\" stretchy=\"true\" symmetric=\"true\"></mo>"
            + "</mrow></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var delimiter = document.Descendants(math + "d").SingleOrDefault();

        Assert.NotNull(delimiter);
        Assert.Equal(
            "{",
            delimiter!.Element(math + "dPr")?
                .Element(math + "begChr")?
                .Attribute(math + "val")?
                .Value);
        var visibleText = string.Concat(
            delimiter.Descendants(math + "t").Select(text => text.Value));
        Assert.Contains("a", visibleText, StringComparison.Ordinal);
        Assert.Contains("b", visibleText, StringComparison.Ordinal);
    }

    [Fact]
    public void RightOnlyBraceAroundTableBecomesNativeGrowingDelimiter()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mrow data-mjx-texclass=\"INNER\">"
            + "<mo data-mjx-texclass=\"OPEN\" fence=\"true\" stretchy=\"true\" symmetric=\"true\"></mo>"
            + "<mtable columnalign=\"left\"><mtr><mtd><mi>a</mi></mtd></mtr>"
            + "<mtr><mtd><mi>b</mi></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\">}</mo></mrow></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var delimiter = document.Descendants(math + "d").SingleOrDefault();

        Assert.NotNull(delimiter);
        Assert.Equal(
            "}",
            delimiter!.Element(math + "dPr")?
                .Element(math + "endChr")?
                .Attribute(math + "val")?
                .Value);
        var visibleText = string.Concat(
            delimiter.Descendants(math + "t").Select(text => text.Value));
        Assert.Contains("a", visibleText, StringComparison.Ordinal);
        Assert.Contains("b", visibleText, StringComparison.Ordinal);
    }

    [Fact]
    public void AngleFencedTableBecomesNativeGrowingDelimiter()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">&#x27E8;</mo>"
            + "<mtable columnalign=\"center center\"><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
            + "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\">&#x27E9;</mo></mrow></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var delimiter = document.Descendants(math + "d").SingleOrDefault();

        Assert.NotNull(delimiter);
        Assert.Equal("⟨", delimiter!.Element(math + "dPr")?.Element(math + "begChr")?.Attribute(math + "val")?.Value);
        Assert.Equal("⟩", delimiter.Element(math + "dPr")?.Element(math + "endChr")?.Attribute(math + "val")?.Value);
        Assert.Equal(2, delimiter.Descendants(math + "m").Single().Elements(math + "mr").Count());
    }

    [Theory]
    [InlineData("(", ")")]
    [InlineData("[", "]")]
    [InlineData("{", "}")]
    [InlineData("|", "|")]
    [InlineData("‖", "‖")]
    [InlineData("⌈", "⌉")]
    [InlineData("⌊", "⌋")]
    [InlineData("⟨", "⟩")]
    public void CommonFencedTablesUseNativeGrowingDelimiters(string open, string close)
    {
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">"
            + open + "</mo><mtable columnalign=\"center center\">"
            + "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
            + "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\">" + close + "</mo></mrow></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var delimiter = document.Descendants(math + "d").SingleOrDefault();

        Assert.NotNull(delimiter);
        var actualOpen = delimiter!.Element(math + "dPr")?
            .Element(math + "begChr")?
            .Attribute(math + "val")?
            .Value ?? "(";
        var actualClose = delimiter.Element(math + "dPr")?
            .Element(math + "endChr")?
            .Attribute(math + "val")?
            .Value ?? ")";
        Assert.Equal(open, actualOpen);
        Assert.Equal(close, actualClose);
        Assert.NotEqual(
            "0",
            delimiter.Element(math + "dPr")?
                .Element(math + "grow")?
                .Attribute(math + "val")?
                .Value);
        var matrix = delimiter.Descendants(math + "m").SingleOrDefault();
        Assert.NotNull(matrix);
        Assert.Equal(2, matrix!.Elements(math + "mr").Count());
        Assert.All(matrix.Elements(math + "mr"), row => Assert.Equal(2, row.Elements(math + "e").Count()));
    }

    [Fact]
    public void MatrixPlaceholderVisibilityIsForcedOnBeforeWordInsertion()
    {
        const string omml =
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:m><m:mPr><m:plcHide m:val=\"0\" /></m:mPr>"
            + "<m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e>"
            + "<m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr>"
            + "</m:m></m:oMath>";

        var normalized = WordOmmlConverter.NormalizeOmmlPlaceholderVisibility(omml);
        var document = XDocument.Parse(normalized);
        XNamespace math = MathNamespace;

        Assert.Equal(
            "1",
            document.Descendants(math + "m").Single()
                .Element(math + "mPr")?
                .Element(math + "plcHide")?
                .Attribute(math + "val")?
                .Value);
    }

    [Fact]
    public void InlineNaryOmmlHidesEmptyLimitsWithoutForcingDisplayGrowth()
    {
        const string omml =
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:nary><m:naryPr /><m:sub /><m:sup />"
            + "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:nary>"
            + "</m:oMath>";

        var normalized = WordOmmlConverter.NormalizeDisplayNaryOmml(
            omml,
            display: false);
        var document = XDocument.Parse(normalized);
        XNamespace math = MathNamespace;
        var properties = document.Descendants(math + "naryPr").Single();

        Assert.Equal("1", properties.Element(math + "subHide")?.Attribute(math + "val")?.Value);
        Assert.Equal("1", properties.Element(math + "supHide")?.Attribute(math + "val")?.Value);
        Assert.Null(properties.Element(math + "grow"));
    }

    [Fact]
    public void InlineBareIntegralTransformsToNativeNaryWithHiddenLimits()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\">"
            + "<mo largeop=\"true\" movablelimits=\"true\">&#x222B;</mo>"
            + "<mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo><mi>d</mi><mi>x</mi>"
            + "</math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var nary = document.Descendants(math + "nary").Single();
        var properties = nary.Element(math + "naryPr");

        Assert.Equal("1", properties?.Element(math + "subHide")?.Attribute(math + "val")?.Value);
        Assert.Equal("1", properties?.Element(math + "supHide")?.Attribute(math + "val")?.Value);
        Assert.Null(properties?.Element(math + "grow"));
        Assert.True(HasVisibleMathText(nary.Element(math + "e")));
        WordOmmlConverter.ValidateMaterializedOmml(omml);
    }

    [Fact]
    public void MaterializedOmmlValidationAcceptsProperlyHiddenEmptySlots()
    {
        const string omml =
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:m><m:mPr><m:plcHide m:val=\"1\" /></m:mPr>"
            + "<m:mr><m:e /><m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m>"
            + "<m:rad><m:radPr><m:degHide m:val=\"1\" /></m:radPr>"
            + "<m:deg /><m:e><m:r><m:t>y</m:t></m:r></m:e></m:rad>"
            + "<m:nary><m:naryPr><m:subHide m:val=\"1\" />"
            + "<m:supHide m:val=\"1\" /></m:naryPr><m:sub /><m:sup />"
            + "<m:e><m:r><m:t>z</m:t></m:r></m:e></m:nary>"
            + "</m:oMath>";

        WordOmmlConverter.ValidateMaterializedOmml(omml);
    }

    [Fact]
    public void MaterializedOmmlValidationRejectsUnhiddenEmptySlots()
    {
        const string unhiddenMatrix =
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:m><m:mPr /><m:mr><m:e />"
            + "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:mr></m:m>"
            + "</m:oMath>";
        const string emptyScript =
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:sSub><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub /></m:sSub>"
            + "</m:oMath>";

        Assert.Throws<InvalidDataException>(() =>
            WordOmmlConverter.ValidateMaterializedOmml(unhiddenMatrix));
        Assert.Throws<InvalidDataException>(() =>
            WordOmmlConverter.ValidateMaterializedOmml(emptyScript));
    }

    [Fact]
    public void HiddenOptionalOmmlSlotsAreAllowed()
    {
        var document = XDocument.Parse(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:rad><m:radPr><m:degHide m:val=\"1\" /></m:radPr>"
            + "<m:deg /><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>"
            + "<m:nary><m:naryPr><m:supHide m:val=\"1\" /></m:naryPr>"
            + "<m:sub><m:r><m:t>V</m:t></m:r></m:sub><m:sup />"
            + "<m:e><m:r><m:t>F</m:t></m:r></m:e></m:nary>"
            + "</m:oMath>");

        WordOmmlConverter.ValidateNoVisibleEmptyOmmlSlots(document);
    }

    [Fact]
    public void VisibleEmptyOmmlSlotsAreRejected()
    {
        var document = XDocument.Parse(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:m><m:mPr><m:plcHide m:val=\"1\" /></m:mPr>"
            + "<m:mr><m:e /></m:mr></m:m>"
            + "</m:oMath>");

        var error = Assert.Throws<InvalidDataException>(() =>
            WordOmmlConverter.ValidateNoVisibleEmptyOmmlSlots(document));
        Assert.Contains("visible empty e slot", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UprightMathMlIdentifiersUseExplicitNormalOmmlRuns()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi mathvariant=\"normal\">d</mi><mi>x</mi><mo>+</mo>"
            + "<mi>d</mi><mi>f</mi><mo>+</mo>"
            + "<mi mathvariant=\"normal\">e</mi>"
            + "<msup><mi mathvariant=\"normal\">i</mi><mi>x</mi></msup>"
            + "</math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var runs = document.Descendants(math + "r").ToArray();

        static bool IsNormalRun(XElement run, XNamespace math) =>
            run.Element(math + "rPr")?.Element(math + "nor") is not null;

        foreach (var token in new[] { "d", "e", "i" })
        {
            Assert.Contains(
                runs,
                run =>
                    run.Element(math + "t")?.Value == token
                    && IsNormalRun(run, math)
                    && run.Element(word + "rPr")?.Element(word + "noProof") is not null);
        }

        Assert.Contains(
            runs,
            run =>
                (run.Element(math + "t")?.Value.Contains("df", StringComparison.Ordinal) ?? false)
                && !IsNormalRun(run, math));
    }

    [Theory]
    [InlineData("^", "\u0302")]
    [InlineData("~", "\u0303")]
    [InlineData("→", "\u20D7")]
    [InlineData("←", "\u20D6")]
    [InlineData("↔", "\u20E1")]
    [InlineData("¯", "\u0305")]
    [InlineData("‾", "\u0305")]
    [InlineData("―", "\u0305")]
    [InlineData("ˉ", "\u0305")]
    [InlineData("˙", "\u0307")]
    [InlineData("¨", "\u0308")]
    [InlineData("ˇ", "\u030C")]
    [InlineData("˘", "\u0306")]
    [InlineData("´", "\u0301")]
    [InlineData("`", "\u0300")]
    [InlineData("˚", "\u030A")]
    public void MathMlAccentsBecomeNativeOmmlAccentsWithoutPlaceholderGlyphs(
        string sourceMark,
        string expectedCombiningMark)
    {
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mrow><mover><mi>x</mi><mo stretchy=\"false\">"
            + sourceMark
            + "</mo></mover></mrow></math>";

        var normalized = WordOmmlConverter.NormalizeMathMlAccents(mathMl);
        var normalizedDocument = XDocument.Parse(normalized);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var mover = normalizedDocument.Descendants(presentationMath + "mover").Single();
        var mark = mover.Elements().Skip(1).Single();
        Assert.Equal("true", mover.Attribute("accent")?.Value);
        Assert.Equal("true", mark.Attribute("accent")?.Value);
        Assert.Equal(expectedCombiningMark, mark.Value);

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var ommlDocument = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var accent = ommlDocument.Descendants(math + "acc").SingleOrDefault();
        Assert.NotNull(accent);
        Assert.Equal(
            expectedCombiningMark,
            accent!
                .Element(math + "accPr")?
                .Element(math + "chr")?
                .Attribute(math + "val")?
                .Value);
        Assert.Equal("x", accent.Element(math + "e")?.Value);
        Assert.DoesNotContain("\uFFFD", omml, StringComparison.Ordinal);
        Assert.DoesNotContain("□", omml, StringComparison.Ordinal);
    }

    [Fact]
    public void ExplicitOversetIsNotMistakenForAnAccent()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mover><mi>x</mi><mo accent=\"false\" stretchy=\"false\">→</mo></mover>"
            + "</math>";

        var normalized = WordOmmlConverter.NormalizeMathMlAccents(mathMl);
        var document = XDocument.Parse(normalized);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var mover = document.Descendants(presentationMath + "mover").Single();

        Assert.Null(mover.Attribute("accent"));
        Assert.Equal("false", mover.Elements().Skip(1).Single().Attribute("accent")?.Value);
        Assert.Equal("→", mover.Elements().Skip(1).Single().Value);
    }

    [Fact]
    public void NestedEmptyBaseSubscriptsAreFlattenedBeforeOmmlConversion()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msub><mi>f</mi><mrow><msub><mi></mi><mrow><mi mathvariant=\"normal\">H</mi></mrow></msub></mrow></msub>"
            + "<mo>=</mo><mfrac><mn>1</mn><mrow><mn>2</mn><mi>&#x3C0;</mi>"
            + "<msub><mi>&#x3C4;</mi><mrow><msub><mi></mi><mrow><mi mathvariant=\"normal\">H</mi></mrow></msub></mrow></msub>"
            + "</mrow></mfrac></math>";

        var normalized = WordOmmlConverter.NormalizeNestedEmptyBaseScripts(mathMl);
        var normalizedDocument = XDocument.Parse(normalized);
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";
        var normalizedScripts = normalizedDocument.Descendants(presentationMath + "msub").ToArray();

        Assert.Equal(2, normalizedScripts.Length);
        Assert.All(normalizedScripts, script =>
        {
            var elements = script.Elements().ToArray();
            Assert.True(elements.Length >= 2);
            Assert.False(string.IsNullOrWhiteSpace(elements[0].Value));
            Assert.Equal("H", elements[1].Value.Trim());
        });

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var ommlDocument = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var scripts = ommlDocument.Descendants(math + "sSub").ToArray();
        Assert.Equal(2, scripts.Length);
        Assert.All(scripts, script =>
        {
            Assert.True(HasVisibleMathText(script.Element(math + "e")));
            Assert.True(HasVisibleMathText(script.Element(math + "sub")));
        });
        Assert.DoesNotContain(scripts, script =>
            script.Element(math + "e")?.Descendants(math + "t").All(text =>
                string.IsNullOrWhiteSpace(text.Value)) != false);
    }

    [Fact]
    public void NestedEmptyBaseSuperscriptIsFlattenedButStandalonePrescriptIsPreserved()
    {
        const string nested =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<msup><mi>x</mi><mrow><msup><mi></mi><mn>2</mn></msup></mrow></msup>"
            + "</math>";
        const string standalone =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<msub><mi></mi><mi>i</mi></msub><mi>A</mi>"
            + "</math>";
        XNamespace presentationMath = "http://www.w3.org/1998/Math/MathML";

        var nestedDocument = XDocument.Parse(
            WordOmmlConverter.NormalizeNestedEmptyBaseScripts(nested));
        var nestedScripts = nestedDocument.Descendants(presentationMath + "msup").ToArray();
        Assert.Single(nestedScripts);
        Assert.Equal("x", nestedScripts[0].Elements().First().Value);
        Assert.Equal("2", nestedScripts[0].Elements().Skip(1).First().Value);

        var standaloneDocument = XDocument.Parse(
            WordOmmlConverter.NormalizeNestedEmptyBaseScripts(standalone));
        var standaloneScript = standaloneDocument.Descendants(presentationMath + "msub").Single();
        Assert.True(string.IsNullOrWhiteSpace(standaloneScript.Elements().First().Value));
        Assert.Equal("i", standaloneScript.Elements().Skip(1).First().Value);
    }

    [Theory]
    [MemberData(nameof(NaryMathMlCases))]
    public void OfficeMathMlTransformNeverLeavesAnEmptyNaryOperand(
        string latex,
        string mathMl)
    {
        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var naries = document.Descendants(math + "nary").ToArray();

        Assert.NotEmpty(naries);
        Assert.All(naries, nary =>
        {
            var operand = nary.Element(math + "e");
            Assert.NotNull(operand);
            Assert.True(
                operand!.Elements().Any(),
                $"{latex} produced an empty m:nary/m:e operand: {omml}");
        });
    }

    [Fact]
    public void OfficeTransformSupportsSyntheticEmptyLimitForBareIntegral()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msub><mo>&#x222B;</mo><mrow /></msub><mrow><mi>a</mi></mrow>"
            + "<mi>d</mi><mi>b</mi></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var nary = document.Descendants(math + "nary").SingleOrDefault();

        Assert.True(nary is not null, $"Synthetic bare integral did not become m:nary: {omml}");
        Assert.True(nary!.Element(math + "e")?.Elements().Any() == true);
    }

    [Fact]
    public void BareDisplayIntegralUsesGrowingNativeNaryLayoutWithoutPlaceholderLimit()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mo>&#x222B;</mo><mi>a</mi><mi>d</mi><mi>b</mi></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var nary = document.Descendants(math + "nary").SingleOrDefault();

        Assert.True(nary is not null, $"Bare display integral was not converted to m:nary: {omml}");
        var properties = nary!.Element(math + "naryPr");
        Assert.Equal("1", properties?.Element(math + "grow")?.Attribute(math + "val")?.Value);
        Assert.Equal("1", properties?.Element(math + "subHide")?.Attribute(math + "val")?.Value);
        Assert.Equal("1", properties?.Element(math + "supHide")?.Attribute(math + "val")?.Value);
        Assert.True(nary.Element(math + "e")?.Elements().Any() == true);
    }

    [Fact]
    public void DefiniteIntegralKeepsBothVisibleLimits()
    {
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msubsup><mo>&#x222B;</mo><mn>0</mn><mn>1</mn></msubsup>"
            + "<mrow><mi>x</mi></mrow><mi>d</mi><mi>x</mi></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var nary = document.Descendants(math + "nary").Single();
        var properties = nary.Element(math + "naryPr");

        Assert.Null(properties?.Element(math + "subHide"));
        Assert.Null(properties?.Element(math + "supHide"));
    }

    [Theory]
    [InlineData("∯")]
    [InlineData("∰")]
    [InlineData("∱")]
    [InlineData("∲")]
    [InlineData("∳")]
    [InlineData("⨋")]
    [InlineData("⨌")]
    [InlineData("⨍")]
    [InlineData("⨎")]
    [InlineData("⨏")]
    [InlineData("⨐")]
    [InlineData("⨑")]
    [InlineData("⨒")]
    [InlineData("⨓")]
    [InlineData("⨔")]
    [InlineData("⨕")]
    [InlineData("⨖")]
    [InlineData("⨗")]
    [InlineData("⨘")]
    [InlineData("⨙")]
    [InlineData("⨚")]
    [InlineData("⨛")]
    [InlineData("⨜")]
    public void ExtendedIntegralOperatorsBecomeNativeOmmlNaries(string symbol)
    {
        var mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + $"<msub><mo>{symbol}</mo><mi>S</mi></msub>"
            + "<mrow><mi>f</mi><mi>d</mi><mi>S</mi></mrow></math>";

        var omml = WordOmmlConverter.TransformMathMlToOmml(mathMl);
        var document = XDocument.Parse(omml);
        XNamespace math = MathNamespace;
        var nary = document.Descendants(math + "nary").SingleOrDefault();

        Assert.True(nary is not null, $"{symbol} did not become native OMML n-ary content: {omml}");
        Assert.DoesNotContain(
            nary!.Ancestors(),
            ancestor =>
                ancestor.Name == math + "sSub"
                || ancestor.Name == math + "sSup"
                || ancestor.Name == math + "sSubSup");
        Assert.Equal(symbol, nary.Element(math + "naryPr")?.Element(math + "chr")?.Attribute(math + "val")?.Value);
        Assert.Contains(
            nary.Element(math + "sub")?.Descendants(math + "t") ?? Enumerable.Empty<XElement>(),
            text => text.Value == "S");
        Assert.Contains(
            nary.Element(math + "e")?.Descendants(math + "t") ?? Enumerable.Empty<XElement>(),
            text => text.Value.Contains("f", StringComparison.Ordinal));
    }

    [Fact]
    public void OmmlContentFingerprintIgnoresRunFontSize()
    {
        const string word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var first =
            $"<m:oMath xmlns:m=\"{MathNamespace}\" xmlns:w=\"{word}\">"
            + "<m:r><m:rPr><w:rPr><w:sz w:val=\"28\"/><w:szCs w:val=\"28\"/></w:rPr></m:rPr><m:t>x</m:t></m:r>"
            + "</m:oMath>";
        var second = first.Replace("28", "36", StringComparison.Ordinal);

        Assert.Equal(
            WordOmmlConverter.ComputeOmmlFingerprint(first),
            WordOmmlConverter.ComputeOmmlFingerprint(second));
    }

    [Fact]
    public void OmmlContentFingerprintStillTracksFormulaStructure()
    {
        var first =
            $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r><m:t>x</m:t></m:r></m:oMath>";
        var second =
            $"<m:oMath xmlns:m=\"{MathNamespace}\"><m:r><m:t>y</m:t></m:r></m:oMath>";

        Assert.NotEqual(
            WordOmmlConverter.ComputeOmmlFingerprint(first),
            WordOmmlConverter.ComputeOmmlFingerprint(second));
    }

    [Fact]
    public void OmmlBookmarkNameRoundTripsPersistentFormulaId()
    {
        var formulaId = Guid.NewGuid().ToString();
        var name = WordOmmlFormulaStore.BookmarkName(formulaId);

        Assert.StartsWith(WordOmmlFormulaStore.BookmarkPrefix, name, StringComparison.Ordinal);
        Assert.True(Guid.TryParseExact(
            name.Substring(WordOmmlFormulaStore.BookmarkPrefix.Length),
            "N",
            out var roundTrip));
        Assert.Equal(Guid.Parse(formulaId), roundTrip);
    }

    private static bool HasVisibleMathText(XElement? element)
    {
        if (element is null) return false;
        XNamespace math = MathNamespace;
        return element
            .Descendants(math + "t")
            .Any(text => !string.IsNullOrWhiteSpace(text.Value));
    }

    private static FormulaMetadata Metadata()
    {
        var now = DateTimeOffset.UtcNow.ToString("O");
        var latex = "\\frac{a}{b}+x^2";
        return new FormulaMetadata
        {
            FormulaId = Guid.NewGuid().ToString(),
            Title = "Native OMML",
            Latex = latex,
            CodeFormat = "latex",
            DisplayMode = "block",
            Numbered = true,
            RenderWidthPx = 180,
            RenderHeightPx = 48,
            Baseline = 36,
            CreatedWithVersion = "1.0.18",
            UpdatedWithVersion = "1.0.18",
            CreatedAt = now,
            UpdatedAt = now,
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString(), Latex = latex },
            },
        };
    }

    public static IEnumerable<object[]> NaryMathMlCases()
    {
        const string ns = "http://www.w3.org/1998/Math/MathML";
        yield return new object[]
        {
            @"\sum_b^z c",
            $"<math xmlns=\"{ns}\" display=\"block\"><munderover><mo>&#x2211;</mo><mi>b</mi><mi>z</mi></munderover><mi>c</mi></math>",
        };
        yield return new object[]
        {
            @"\sum_{b}^{z} c",
            $"<math xmlns=\"{ns}\" display=\"block\"><munderover><mo>&#x2211;</mo><mrow><mi>b</mi></mrow><mrow><mi>z</mi></mrow></munderover><mi>c</mi></math>",
        };
        yield return new object[]
        {
            @"\oint_l^u x\,dy",
            $"<math xmlns=\"{ns}\" display=\"block\"><msubsup><mo>&#x222E;</mo><mi>l</mi><mi>u</mi></msubsup><mi>x</mi><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>y</mi></math>",
        };
        yield return new object[]
        {
            @"\oint_l x\,dy",
            $"<math xmlns=\"{ns}\" display=\"block\"><msub><mo>&#x222E;</mo><mi>l</mi></msub><mi>x</mi><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>y</mi></math>",
        };
        yield return new object[]
        {
            @"\int_0^1 x^2\,dx",
            $"<math xmlns=\"{ns}\" display=\"block\"><msubsup><mo>&#x222B;</mo><mn>0</mn><mn>1</mn></msubsup><msup><mi>x</mi><mn>2</mn></msup><mstyle><mspace width=\"0.167em\"/></mstyle><mi>d</mi><mi>x</mi></math>",
        };
        yield return new object[]
        {
            @"\prod_{i=1}^{n} a_i",
            $"<math xmlns=\"{ns}\" display=\"block\"><munderover><mo>&#x220F;</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mrow><mi>n</mi></mrow></munderover><msub><mi>a</mi><mi>i</mi></msub></math>",
        };
    }
}

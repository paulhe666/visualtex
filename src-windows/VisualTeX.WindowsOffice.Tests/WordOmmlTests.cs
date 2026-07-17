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

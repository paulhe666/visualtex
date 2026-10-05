using System.Xml;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class FormulaXmlBoundaryTests
{
    private static string Math(string body) => "<math xmlns='http://www.w3.org/1998/Math/MathML'>" + body + "</math>";
    private static string DeepBody(int count) => string.Concat(Enumerable.Repeat("<mrow>", count))
        + "<mi>x</mi>" + string.Concat(Enumerable.Repeat("</mrow>", count));

    [Theory]
    [InlineData("validate")]
    [InlineData("omml")]
    [InlineData("latex")]
    [InlineData("mtef-create")]
    [InlineData("mtef-rewrite")]
    [InlineData("mtef-signature")]
    [InlineData("mtef-interop")]
    public void PublicMathMlEntrypointsRejectExcessDepthBeforeRecursion(string route)
    {
        var input = Math(DeepBody(260));
        Assert.Throws<InvalidDataException>(() => Invoke(route, input));
    }

    [Theory]
    [InlineData("validate")]
    [InlineData("omml")]
    [InlineData("latex")]
    [InlineData("mtef-create")]
    [InlineData("mtef-signature")]
    [InlineData("mtef-interop")]
    public void PublicEntrypointsRejectDtdAndExternalEntities(string route)
    {
        var input = "<!DOCTYPE math [<!ENTITY external SYSTEM 'file:///does-not-exist'>]>" + Math("<mtext>&external;</mtext>");
        Assert.Throws<XmlException>(() => Invoke(route, input));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OmmlEntrypointsRejectExcessDepthBeforeRecursiveNumberHandling(bool reverse)
    {
        var input = "<m:oMath xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'>"
            + string.Concat(Enumerable.Repeat("<m:e>", 260)) + "<m:r><m:t>x</m:t></m:r>"
            + string.Concat(Enumerable.Repeat("</m:e>", 260)) + "</m:oMath>";
        Assert.Throws<InvalidDataException>(() =>
        {
            if (reverse) WordOmmlConverter.TransformOmmlToMathMl(input, false);
            else WordOmmlConverter.ExtractSingleOMath(input);
        });
    }

    [Fact]
    public void DepthAndElementLimitsApplyDuringLoadIncludingEmptyElements()
    {
        FormulaXml.Parse("<root>" + string.Concat(Enumerable.Repeat("<e>", 255))
            + string.Concat(Enumerable.Repeat("</e>", 255)) + "</root>", "fixture");
        Assert.Throws<InvalidDataException>(() => FormulaXml.Parse("<root>" + string.Concat(Enumerable.Repeat("<e>", 256))
            + string.Concat(Enumerable.Repeat("</e>", 256)) + "</root>", "fixture"));
        FormulaXml.Parse("<root>" + string.Concat(Enumerable.Repeat("<e/>", 249_999)) + "</root>", "fixture");
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ValidateMathMlForOmml(
            Math(string.Concat(Enumerable.Repeat("<mi/>", 250_000)))));
        Assert.Throws<InvalidDataException>(() => MathMlToLatexConverter.Convert(Math("<mtext>" + new string('x', 4_000_000) + "</mtext>")));
    }

    [Theory]
    [InlineData(0x200B)]
    [InlineData(0x200C)]
    [InlineData(0x2060)]
    [InlineData(0xFEFF)]
    public void LegalTokenContentSurvivesLiteralDecimalAndHexXmlForms(int scalar)
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var expected = "ع" + char.ConvertFromUtf32(scalar) + "ب";
        foreach (var spelling in new[] { char.ConvertFromUtf32(scalar), "&#" + scalar + ";", "&#x" + scalar.ToString("X") + ";" })
        {
            var source = Math("<mtext>ع" + spelling + "ب</mtext>");
            var native = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(source));
            Assert.Equal(expected, string.Concat(native.Descendants(m + "t").Select(node => node.Value)));
        }
    }

    [Fact]
    public void ZeroWidthOnlyTokenIsNotTreatedAsOwnedHostBoundary()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var native = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(Math("<mtext>\u200C</mtext><mi>x</mi>")));
        Assert.Equal("\u200Cx", string.Concat(native.Descendants(m + "t").Select(node => node.Value)));
    }

    [Fact]
    public void LiteralTextPreservesSpacingAndOperatorLikeCharactersAcrossGroups()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string text = " a  \u00A0\u2062\u2A75 b ";
        var native = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(Math(
            "<mi>x</mi><mtext>" + text + "</mtext><mtext> next </mtext><mo>+</mo><mi>y</mi>")));
        Assert.Equal("x" + text + " next +y", string.Concat(native.Descendants(m + "t").Select(node => node.Value)));
    }

    [Fact]
    public void StringLiteralsKeepTextAndQuoteSemanticsNextToMathTokens()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string text = "a\u200B\u00A0\u2A75b";
        var native = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(Math(
            "<mi>x</mi><ms lquote='[' rquote=']'>" + text + "</ms><mi>y</mi>")));
        Assert.Equal("x[" + text + "]y", string.Concat(native.Descendants(m + "t").Select(node => node.Value)));
    }

    private static void Invoke(string route, string input)
    {
        switch (route)
        {
            case "validate": WordOmmlConverter.ValidateMathMlForOmml(input); break;
            case "omml": WordOmmlConverter.TransformMathMlToOmml(input); break;
            case "latex": MathMlToLatexConverter.Convert(input); break;
            case "mtef-create": MathTypeMtefCodec.CreateEquationNative(input, true); break;
            case "mtef-rewrite":
                var seed = MathTypeMtefCodec.CreateEquationNative(Math("<mi>x</mi>"), true);
                MathTypeMtefCodec.RewriteEquationNative(seed.EquationNative, input, true); break;
            case "mtef-signature": MathTypeMtefCodec.SemanticSignature(input); break;
            case "mtef-interop": MathTypeMtefCodec.PrepareMathMlForMathTypeInterop(input); break;
            default: throw new ArgumentException(route);
        }
    }
}

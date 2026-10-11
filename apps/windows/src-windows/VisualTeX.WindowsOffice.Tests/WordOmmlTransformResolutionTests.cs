using System.Text;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlTransformResolutionTests
{
    // A valid stylesheet with the Office file name but none of the Microsoft
    // templates VisualTeX adapts, as a third-party or foreign copy found first
    // on the search path would be.
    private const string ForeignStylesheet =
        "<xsl:stylesheet version=\"1.0\" xmlns:xsl=\"http://www.w3.org/1999/XSL/Transform\">"
        + "<xsl:template match=\"/\"><math xmlns=\"http://www.w3.org/1998/Math/MathML\"/></xsl:template>"
        + "</xsl:stylesheet>";

    [Fact]
    public void AForeignStylesheetFirstOnThePathFallsBackToTheOfficeCopy()
    {
        using var temp = new TemporaryDirectory();
        var foreign = Path.Combine(temp.Path, "OMML2MML.XSL");
        File.WriteAllText(foreign, ForeignStylesheet, new UTF8Encoding(false));
        var office = WordOmmlConverter.ResolveReverseTransformPath();

        var transform = WordOmmlConverter.LoadFirstUsableTransform(
            "OMML2MML.XSL", new[] { foreign, office }, WordOmmlConverter.LoadReverseTransform);

        Assert.NotNull(transform);
    }

    [Fact]
    public void OnlyForeignStylesheetsReportEachRejectedFile()
    {
        using var temp = new TemporaryDirectory();
        var foreign = Path.Combine(temp.Path, "MML2OMML.XSL");
        File.WriteAllText(foreign, ForeignStylesheet, new UTF8Encoding(false));

        var error = Assert.Throws<InvalidDataException>(() =>
            WordOmmlConverter.LoadFirstUsableTransform(
                "MML2OMML.XSL", new[] { foreign }, WordOmmlConverter.LoadTransform));

        Assert.Contains("MML2OMML.XSL", error.Message);
        Assert.Contains(foreign, error.Message);
    }

    // Office builds word their templates differently. A copy where one adapted
    // template has another shape is still a Microsoft stylesheet and must be
    // used, with only that adaptation left out.
    [Fact]
    public void AnOfficeCopyMissingOneReverseAdaptationStillConverts()
    {
        using var temp = new TemporaryDirectory();
        var variant = WriteVariant(temp, WordOmmlConverter.ResolveReverseTransformPath(),
            "$sTokenType='mo' and $mathvariant!='normal'",
            "$sTokenType='mo' and not($mathvariant='normal')");

        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.LoadReverseTransform(variant));
        var transform = WordOmmlConverter.LoadFirstUsableTransform(
            "OMML2MML.XSL", new[] { variant }, WordOmmlConverter.LoadReverseTransform);

        var output = Transform(transform,
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<m:r><m:t>x+1</m:t></m:r></m:oMath>");
        Assert.Contains(">x<", output);
        Assert.Contains(">1<", output);
    }

    [Fact]
    public void AnOfficeCopyMissingOneForwardAdaptationStillLoads()
    {
        using var temp = new TemporaryDirectory();
        var variant = WriteVariant(temp, WordOmmlConverter.ResolveTransformPath(),
            "$font='italic'", "$font = 'italic'");

        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.LoadTransform(variant));
        var transform = WordOmmlConverter.LoadFirstUsableTransform(
            "MML2OMML.XSL", new[] { variant }, WordOmmlConverter.LoadTransform);

        var output = Transform(transform,
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>");
        Assert.Contains("oMath", output);
    }

    private static string WriteVariant(TemporaryDirectory temp, string office, string anchor, string replacement)
    {
        var text = File.ReadAllText(office);
        Assert.Contains(anchor, text);
        var variant = Path.Combine(temp.Path, Path.GetFileName(office));
        File.WriteAllText(variant, text.Replace(anchor, replacement), new UTF8Encoding(false));
        return variant;
    }

    private static string Transform(System.Xml.Xsl.XslCompiledTransform transform, string xml)
    {
        using var input = System.Xml.XmlReader.Create(new StringReader(xml));
        using var output = new StringWriter();
        var arguments = new System.Xml.Xsl.XsltArgumentList();
        arguments.AddExtensionObject(
            "urn:visualtex:unicode-scalar-xpath", new WordOmmlConverter.UnicodeScalarXPathFunctions());
        transform.Transform(input, arguments, output);
        return output.ToString();
    }

    private sealed class TemporaryDirectory : IDisposable
    {
        public TemporaryDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                $"visualtex-xsl-resolution-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose()
        {
            try { Directory.Delete(Path, true); } catch { }
        }
    }
}

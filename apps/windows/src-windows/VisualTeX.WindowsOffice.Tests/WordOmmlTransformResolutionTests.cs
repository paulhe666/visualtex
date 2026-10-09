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

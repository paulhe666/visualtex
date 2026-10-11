using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Writes the real formulas of Office documents as an acceptance corpus for
// the owned-Word conversion runs. VT_ACCEPTANCE_CORPUS_SOURCES lists .docx
// files (';'), VT_ACCEPTANCE_CORPUS_OUTPUT is the JSON written.
public sealed class AcceptanceCorpusWriter
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    [Fact]
    public void WriteCorpus()
    {
        var sources = (Environment.GetEnvironmentVariable("VT_ACCEPTANCE_CORPUS_SOURCES") ?? string.Empty)
            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        var outputPath = Environment.GetEnvironmentVariable("VT_ACCEPTANCE_CORPUS_OUTPUT");
        if (sources.Length == 0 || string.IsNullOrEmpty(outputPath)) return;

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var formulas = new List<object>();
        foreach (var source in sources)
        {
            using var archive = ZipFile.OpenRead(source);
            var entry = archive.GetEntry("word/document.xml")!;
            XDocument document;
            using (var stream = entry.Open()) document = XDocument.Load(stream);
            foreach (var equation in document.Descendants(M + "oMath"))
            {
                var display = equation.Parent?.Name == M + "oMathPara";
                var copy = new XElement(equation);
                foreach (var scope in equation.AncestorsAndSelf())
                foreach (var declaration in scope.Attributes().Where(attribute => attribute.IsNamespaceDeclaration))
                    if (copy.Attribute(declaration.Name) is null) copy.SetAttributeValue(declaration.Name, declaration.Value);
                var mathMl = WordOmmlConverter.TransformOmmlToMathMl(copy.ToString(SaveOptions.DisableFormatting), display);
                var latex = MathMlToLatexConverter.Convert(mathMl);
                if (string.IsNullOrWhiteSpace(latex) || !seen.Add(latex)) continue;
                formulas.Add(new { latex, mathMl, displayMode = display ? "block" : "inline" });
            }
        }
        File.WriteAllText(outputPath!, JsonSerializer.Serialize(new { acceptanceCorpus = true, formulas },
            new JsonSerializerOptions { WriteIndented = true }));
        Assert.NotEmpty(formulas);
    }
}

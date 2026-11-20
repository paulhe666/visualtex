using System.IO.Compression;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
using Xunit.Abstractions;

namespace VisualTeX.WindowsOffice.Tests;

// Every formula in real Word and PowerPoint documents projects exactly as the
// classic Office stylesheets did, in both directions. VT_OFFICE_CORPUS lists
// directories (separated by ';') searched for .docx and .pptx files.
public sealed class OfficeDocumentCorpusOracleTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private readonly ITestOutputHelper output;

    public OfficeDocumentCorpusOracleTests(ITestOutputHelper output) => this.output = output;

    [Fact]
    public void EveryDocumentFormulaProjectsAsTheClassicStylesheets()
    {
        var roots = (Environment.GetEnvironmentVariable("VT_OFFICE_CORPUS") ?? string.Empty)
            .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        if (roots.Length == 0) return;

        var reverse = ClassicOfficeReverseStylesheet.Load(ClassicOfficeReverseStylesheet.ResolvePath());
        var forward = ClassicOfficeForwardStylesheet.Load(ClassicOfficeForwardStylesheet.ResolvePath());
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var files = 0;
        var formulas = 0;
        var failures = new List<string>();
        foreach (var path in roots.SelectMany(root =>
                     Directory.EnumerateFiles(root, "*.*", SearchOption.AllDirectories)
                         .Where(file => file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)
                             || file.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase))))
        {
            IEnumerable<string> equations;
            try { equations = ReadEquations(path).ToList(); }
            catch (Exception error)
            {
                output.WriteLine($"SKIPPED {path}: {error.GetType().Name}: {error.Message}");
                continue;
            }
            files++;
            foreach (var omml in equations)
            {
                if (!seen.Add(omml)) continue;
                formulas++;
                var expectedMathMl = OmmlMathMlProjectionOracleTests.Canonical(
                    OmmlMathMlProjectionOracleTests.OracleProject(reverse, omml));
                var actualMathMl = OmmlMathMlProjectionOracleTests.Canonical(OmmlMathMlProjection.Project(omml));
                if (expectedMathMl != actualMathMl)
                    failures.Add($"READ {path}\nOMML: {omml}\nXSL: {expectedMathMl}\nOWN: {actualMathMl}");

                // The insertion direction on the MathML the read pipeline produces.
                string mathMl;
                try { mathMl = WordOmmlConverter.TransformOmmlToMathMl(omml, false); }
                catch (Exception error)
                {
                    failures.Add($"READ PIPELINE {path}: {error.GetType().Name}: {error.Message}\nOMML: {omml}");
                    continue;
                }
                var expectedOmml = WordOmmlConverter.ExtractSingleOMath(
                    MathMlOmmlProjectionOracleTests.OracleProject(forward, mathMl));
                var actualOmml = WordOmmlConverter.ExtractSingleOMath(MathMlOmmlProjection.Project(mathMl));
                if (expectedOmml != actualOmml)
                    failures.Add($"INSERT {path}\nMATHML: {mathMl}\nXSL: {expectedOmml}\nOWN: {actualOmml}");
                try { WordOmmlConverter.TransformMathMlToOmml(mathMl); }
                catch (Exception error)
                {
                    failures.Add($"INSERT PIPELINE {path}: {error.GetType().Name}: {error.Message}\nMATHML: {mathMl}");
                }
            }
        }
        output.WriteLine($"files={files} distinctFormulas={formulas} failures={failures.Count}");
        foreach (var failure in failures.Take(10)) output.WriteLine(failure + "\n");
        Assert.True(formulas > 0, "The corpus contained no formulas.");
        Assert.Empty(failures);
    }

    private static IEnumerable<string> ReadEquations(string path)
    {
        using var archive = ZipFile.OpenRead(path);
        foreach (var entry in archive.Entries.Where(entry =>
                     entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                     && (entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase)
                         || entry.FullName.StartsWith("ppt/slides/", StringComparison.OrdinalIgnoreCase))))
        {
            XDocument document;
            using (var stream = entry.Open()) document = XDocument.Load(stream);
            foreach (var equation in document.Descendants(M + "oMath"))
            {
                // The equation as Word hands it over: its own namespaces declared.
                var copy = new XElement(equation);
                foreach (var scope in equation.AncestorsAndSelf())
                foreach (var declaration in scope.Attributes().Where(attribute => attribute.IsNamespaceDeclaration))
                    if (copy.Attribute(declaration.Name) is null) copy.SetAttributeValue(declaration.Name, declaration.Value);
                yield return copy.ToString(SaveOptions.DisableFormatting);
            }
        }
    }
}

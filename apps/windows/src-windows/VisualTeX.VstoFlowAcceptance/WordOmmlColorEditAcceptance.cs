using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class OmmlColorState
    {
        public string Text { get; set; } = "";
        public string?[] Colors { get; set; } = Array.Empty<string?>();
        public string?[] Styles { get; set; } = Array.Empty<string?>();
        public string?[] ControlColors { get; set; } = Array.Empty<string?>();
        public string RecoveredLatex { get; set; } = "";
    }

    private static void RunWordOmmlColorEditAcceptance(string artifactRoot)
    {
        using var input = JsonDocument.Parse(File.ReadAllText(
            Environment.GetEnvironmentVariable("VISUALTEX_OMML_COLOR_CASES")!));
        var rows = new List<object>();
        foreach (var item in input.RootElement.GetProperty("cases").EnumerateArray())
        {
            var name = item.GetProperty("name").GetString()!;
            var targetMathMl = item.GetProperty("mathMl").GetString()!;
            var path = Path.Combine(artifactRoot, name + ".docx");
            var expectedText = item.GetProperty("expectedText").GetString();
            var expectedColors = item.GetProperty("expectedColors").EnumerateArray().Select(e => e.ValueKind == JsonValueKind.Null ? null : e.GetString()).ToArray();
            var expectedControls = item.GetProperty("expectedControlColors").EnumerateArray().Select(e => e.GetString()).ToArray();
            void Check(OmmlColorState state)
            {
                AssertEqual(expectedText, state.Text, name + " text");
                AssertEqual(JsonSerializer.Serialize(expectedColors), JsonSerializer.Serialize(state.Colors), name + " exact colors");
                AssertEqual(JsonSerializer.Serialize(expectedControls), JsonSerializer.Serialize(state.ControlColors), name + " control colors");
            }
            using (var host = new WordPerformanceHost(null))
            {
                Word.Range? insertion = null; Word.Range? inserted = null;
                try
                {
                    host.Document.Content.Text = "LEFT RIGHT\r";
                    insertion = host.Document.Range(5, 5);
                    inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(host.Document, insertion,
                        WordOmmlConverter.TransformMathMlToOmml(item.GetProperty("sourceMathMl").GetString()!),
                        display: false, mathFontName: host.Document.OMathFontName);
                    var before = CaptureOmmlColorState(host.Document, artifactRoot, name, "before");
                    var surrounding = host.Document.Content.Text;
                    var source = WordFormulaHostLookup.ResolveLocal(host.Document, inserted, WordFormulaHostKind.Omml)
                        ?? throw new InvalidDataException("Owned color-edit source was not resolved.");
                    var request = new WordFormulaHostWriteRequest { Kind = WordFormulaHostKind.Omml,
                        DisplayMode = source.DisplayMode, Numbered = false, MathMl = targetMathMl, FontSizePoints = 12 };
                    WordFormulaHostMutationKernel.Replace(host.Application, host.Document, source, request);
                    var applied = CaptureOmmlColorState(host.Document, artifactRoot, name, "apply"); Check(applied);
                    AssertEqual(surrounding, host.Document.Content.Text, name + " surrounding text");
                    AssertEqual(JsonSerializer.Serialize(before.Styles), JsonSerializer.Serialize(applied.Styles), name + " existing math styles");
                    host.Document.Undo(1);
                    var undone = CaptureOmmlColorState(host.Document, artifactRoot, name, "undo");
                    AssertEqual(JsonSerializer.Serialize(before.Colors), JsonSerializer.Serialize(undone.Colors), name + " one-step Undo colors");
                    host.Document.Redo(1);
                    var redone = CaptureOmmlColorState(host.Document, artifactRoot, name, "redo"); Check(redone);
                    host.Save(path);
                    rows.Add(new { name, before, applied, undone, redone });
                }
                finally { Release(inserted); Release(insertion); }
            }
            using (var reopened = new WordPerformanceHost(path))
            {
                var saved = CaptureOmmlColorState(reopened.Document, artifactRoot, name, "reopen"); Check(saved);
                Word.OMaths? maths = null; Word.OMath? equation = null; Word.Range? range = null;
                try
                {
                    maths = reopened.Document.OMaths; equation = maths[1]; range = equation.Range;
                    var source = WordFormulaHostLookup.ResolveLocal(reopened.Document, range, WordFormulaHostKind.Omml)
                        ?? throw new InvalidDataException("Reopened color-edit source was not resolved.");
                    var native = WordFormulaHostSemanticReader.Read(reopened.Document, source);
                    AssertTrue(native.Latex.Contains("textcolor"), name + " native recovery has color");
                    WordFormulaHostMutationKernel.Replace(reopened.Application, reopened.Document, source,
                        new WordFormulaHostWriteRequest { Kind = WordFormulaHostKind.Omml, DisplayMode = source.DisplayMode,
                            Numbered = false, MathMl = native.MathMl, FontSizePoints = 12 });
                    var unchanged = CaptureOmmlColorState(reopened.Document, artifactRoot, name, "unchanged"); Check(unchanged);
                    rows.Add(new { name, saved, unchanged, native.Latex });
                }
                finally { Release(range); Release(equation); Release(maths); }
            }
        }
        File.WriteAllText(Path.Combine(artifactRoot, "color-lifecycle.json"),
            JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("PASS owned OMML color edit, Undo/Redo, separate-process reopen and unchanged native reapply.");
    }

    private static OmmlColorState CaptureOmmlColorState(Word.Document document, string root, string name, string stage)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var xml = content.WordOpenXML;
            File.WriteAllText(Path.Combine(root, name + "-" + stage + ".xml"), xml);
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var math = XDocument.Parse(xml).Descendants(m + "oMath").Single();
            var colors = new List<string?>(); var styles = new List<string?>();
            var text = new System.Text.StringBuilder();
            foreach (var run in math.Descendants(m + "r"))
            {
                var value = (string?)run.Element(m + "t") ?? "";
                var color = (string?)run.Element(w + "rPr")?.Element(w + "color")?.Attribute(w + "val");
                if (color == "auto") color = null;
                var style = (string?)run.Element(m + "rPr")?.Element(m + "sty")?.Attribute(m + "val") ?? "i";
                foreach (var character in value) { text.Append(character); colors.Add(color?.ToUpperInvariant()); styles.Add(style); }
            }
            var controlColors = math.Descendants(m + "ctrlPr").Select(e =>
                (string?)e.Element(w + "rPr")?.Element(w + "color")?.Attribute(w + "val")).Where(c => c is not null && c != "auto").ToArray();
            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(math.ToString(), false);
            File.WriteAllText(Path.Combine(root, name + "-" + stage + ".mathml"), mathMl);
            return new OmmlColorState { Text = text.ToString(), Colors = colors.ToArray(), Styles = styles.ToArray(),
                ControlColors = controlColors, RecoveredLatex = MathMlToLatexConverter.Convert(mathMl) };
        }
        finally { Release(content); }
    }
}

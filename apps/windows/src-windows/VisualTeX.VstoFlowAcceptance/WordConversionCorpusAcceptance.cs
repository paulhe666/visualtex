using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void VerifySavedExternalConversion(string artifactRoot)
    {
        var source = Environment.GetEnvironmentVariable("VISUALTEX_VERIFY_CONVERSION_SOURCE")!;
        var output = Environment.GetEnvironmentVariable("VISUALTEX_VERIFY_CONVERSION_TARGET")!;
        var targetMode = Environment.GetEnvironmentVariable("VISUALTEX_VERIFY_CONVERSION_TARGET_MODE")!;
        WordFormulaFormatConversionPlan before;
        string topology; int paragraphs;
        using (var host = new WordPerformanceHost(source))
        {
            before = new WordFormulaService(host.Application).CaptureFormulaFormatConversionPlan(true, FormulaOleContract.WordOmmlMode, targetMode);
            topology = LargeRedrawTableTopology(host.Document);
            paragraphs = host.Document.Paragraphs.Count;
        }
        using (var host = new WordPerformanceHost(output))
        {
            var service = new WordFormulaService(host.Application);
            AssertExternalConversionStructure(host, service, before, targetMode, topology, paragraphs, "Saved output");
            var after = service.CaptureFormulaFormatConversionPlan(true, targetMode, FormulaOleContract.WordOmmlMode);
            AssertEqual(before.Targets.Count, after.Targets.Count, "Saved output count");
            for (var index = 0; index < before.Targets.Count; index++)
                AssertExternalConversionFormula(before.Targets[index], after.Targets[index], "Saved output", index);
        }
        File.WriteAllText(Path.Combine(artifactRoot, "verification.json"), JsonSerializer.Serialize(new { Status = "Passed", Count = before.Targets.Count, Source = source, Target = output }));
    }
    private static void AssertExternalConversionStructure(WordPerformanceHost host, WordFormulaService service,
        WordFormulaFormatConversionPlan before, string mode, string topology, int paragraphs, string label)
    {
        AssertEqual(topology, LargeRedrawTableTopology(host.Document), label + " table topology");
        AssertEqual(paragraphs, host.Document.Paragraphs.Count, label + " paragraph boundaries");
        AssertConversionNumberSequence(host, service, mode, before.Targets.Count(item => item.Numbered), label);
        for (var index = 1; index <= before.Targets.Count; index++)
        {
            AssertTrue(host.Document.Content.Text.Contains($"BEGIN[{index:0000}] 前文"), label + " leading prose " + index);
            AssertTrue(host.Document.Content.Text.Contains($"后文 END[{index:0000}]"), label + " trailing prose " + index);
        }
    }
    private static void AssertConversionNumberSequence(WordPerformanceHost host, WordFormulaService service,
        string mode, int expectedCount, string label)
    {
        var numberSource = mode == FormulaOleContract.NativeOleMode ? EquationReferenceSource.VisualTeX
            : mode == FormulaOleContract.MathTypeOleMode ? EquationReferenceSource.MathType : EquationReferenceSource.WordOmml;
        var references = mode == FormulaOleContract.MathTypeOleMode
            ? MathTypeEquationReferences.GetTargets(host.Document) : service.GetCanonicalEquationReferenceTargets(host.Document);
        var numbers = references.Where(item => item.Source == numberSource)
            .OrderBy(item => item.Position).Select(item => item.NumberText.Trim().Trim('(', ')')).ToArray();
        AssertTrue(numbers.SequenceEqual(Enumerable.Range(1, expectedCount).Select(value => value.ToString())),
            label + " visible numbers: " + string.Join(",", numbers));
    }
    private static void AssertExternalConversionFormula(WordFormulaFormatConversionTarget before,
        WordFormulaFormatConversionTarget after, string label, int index)
    {
        AssertEqual(before.DisplayMode, after.DisplayMode, label + " display " + index);
        if (after.SourceMathMl is null) { AssertEqual(before.Latex, after.Latex, label + " durable LaTeX " + index); return; }
        var expected = before.SourceMathMl;
        if (expected is null)
        {
            var journal = Environment.GetEnvironmentVariable("VISUALTEX_RENDER_MATRIX_EXPECTED_JOURNAL")!;
            foreach (var line in File.ReadLines(journal))
            {
                using var entry = JsonDocument.Parse(line);
                if (entry.RootElement.GetProperty("latex").GetString() == before.Latex)
                    expected = entry.RootElement.GetProperty("mathMl").GetString();
            }
        }
        AssertTrue(expected is not null, label + " missing expected MathML " + index);
        AssertEqual(WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(expected!),
            WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(after.SourceMathMl), label + " semantics " + index);
    }
    private sealed class ConversionCorpusItem
    {
        internal string Latex = "";
        internal string MathMl = "";
    }
    private static List<ConversionCorpusItem> LoadConversionCorpus(string path)
    {
        using var json = JsonDocument.Parse(File.ReadAllText(path));
        var wanted = new HashSet<string>(new[] {
            @"a\parallel b", @"a\shortparallel b", @"a\nparallel b", @"a\nshortparallel b",
            @"a\mid b", @"a\shortmid b", @"a\emptyset b", @"a\varnothing b",
            @"x^2", @"x_2", @"x_i^j", @"x^{a_b}", @"x_{a^b}", @"x^{a^{b_c}}",
            @"{}^{14}_{6}\mathrm{C}", @"a+{}^{2}x", @"(x+y)^2", @"{\frac{a}{b}}^{n+1}",
            @"\sqrt{x_i^2+y_i^2}", @"\sum_{i=1}^{n}x_i^2",
            @"\angle ABC+\measuredangle DEF+\sphericalangle GHI",
            @"\boldsymbol{\angle}+\mathcal{\angle}+\mathbb{R}+\mathfrak{g}"
        }, StringComparer.Ordinal);
        var result = new List<ConversionCorpusItem>();
        foreach (var item in json.RootElement.GetProperty("formulas").EnumerateArray())
        {
            var latex = item.GetProperty("latex").GetString()!;
            if (wanted.Remove(latex))
                result.Add(new ConversionCorpusItem { Latex = latex, MathMl = item.GetProperty("mathMl").GetString()! });
        }
        AssertEqual(0, wanted.Count, "Focused corpus missing formulas: " + string.Join(";", wanted));
        return result;
    }
    private static void CreateConversionCorpusFixture(string root, string path, List<ConversionCorpusItem> corpus)
    {
        var application = CreateFreshFormatScalingWord(root, "conversion-symbols");
        Word.Document? document = null;
        try
        {
            document = application.Documents.Add();
            document.Activate();
            document.Content.Text = "VisualTeX format conversion scaling source\r";
            WordEquationNumbering.SetEquationNumberFormatPreference(document, EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            foreach (var formula in corpus)
            {
                SelectDocumentEnd(document);
                var session = CreateOmmlMathTypeAcceptanceSession(formula.MathMl, "block", true, FormulaOleContract.WordOmmlMode);
                service.InsertOmml(session, formula.MathMl);
                SelectDocumentEnd(document);
                application.Selection.TypeParagraph();
            }
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
        }
        finally
        {
            if (document is not null) { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(document); }
            QuitWordApplicationIfOwned(application); Release(application);
            ForceComCleanup();
        }
    }
    private static void CreateVisualTeXConversionCorpusFixture(string path, List<ConversionCorpusItem> corpus)
    {
        var oldNumber = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS");
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS", "1");
            using var host = new WordPerformanceHost(null);
            host.Document.Content.Text = "VisualTeX format conversion scaling source\r"
                + string.Join("\r", corpus.Select(item => "$$" + item.Latex + "$$")) + "\r";
            host.AddIn.OnRedrawDocumentToOle(new object());
            WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(20));
            var service = new WordFormulaService(host.Application);
            AssertEqual(corpus.Count, service.CountFormulaObjectsForLatex(true, FormulaOleContract.NativeOleMode), "Corpus direct VisualTeX seed");
            host.Save(path);
        }
        finally { Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_NUMBER_DISPLAY_FORMULAS", oldNumber); }
    }
    private static void AssertConversionCorpusFormula(ConversionCorpusItem item,
        WordFormulaFormatConversionTarget before, WordFormulaFormatConversionTarget after,
        string root, string label, int index)
    {
        if (after.SourceMathMl is null)
            AssertEqual(before.Latex, after.Latex, label + " durable LaTeX " + index);
        else
        {
            var expected = before.SourceMathMl ?? item.MathMl;
            var expectedSignature = WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(expected);
            var actualSignature = WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(after.SourceMathMl);
            File.WriteAllText(Path.Combine(root, label + "-semantic-" + index + ".json"),
                JsonSerializer.Serialize(new { item.Latex, expected, actual = after.SourceMathMl, expectedSignature, actualSignature }));
            AssertEqual(expectedSignature, actualSignature, label + " structural semantics " + index);
        }
    }
    private static class ConversionCaptureNative
    {
        [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] internal static extern bool SetForegroundWindow(IntPtr window);
        [DllImport("user32.dll")] internal static extern bool ShowWindow(IntPtr window, int command);
    }
    private static void CaptureConversionCorpusScreenshots(WordPerformanceHost host,
        IList<WordFormulaFormatConversionTarget> formulas, string root, string label)
    {
        host.Application.Visible = true;
        host.Document.Activate();
        var window = host.Document.ActiveWindow;
        try
        {
            window.View.Zoom.Percentage = 150;
            var handle = ConversionCaptureNative.GetAncestor(new IntPtr(window.Hwnd), 2);
            ConversionCaptureNative.ShowWindow(handle, 9);
            ConversionCaptureNative.SetForegroundWindow(handle);
            for (var index = 0; index < formulas.Count; index++)
            {
                var formula = formulas[index];
                var range = WordFormulaOperationLocator.ResolveCapturedRange(host.Document, formula.SourceObjectId);
                try
                {
                    window.ScrollIntoView(range, true);
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(120);
                    window.GetPoint(out var left, out var top, out var width, out var height, range);
                    var screen = System.Windows.Forms.Screen.FromHandle(handle).Bounds;
                    left = Math.Max(screen.Left, left - 30); top = Math.Max(screen.Top, top - 25);
                    width = Math.Min(width + 60, screen.Right - left); height = Math.Min(height + 50, screen.Bottom - top);
                    using var bitmap = new Bitmap(width, height);
                    using var graphics = Graphics.FromImage(bitmap);
                    graphics.CopyFromScreen(left, top, 0, 0, bitmap.Size);
                    bitmap.Save(Path.Combine(root, label + "-formula-" + (index + 1).ToString("000") + ".png"), ImageFormat.Png);
                }
                finally { Release(range); }
            }
        }
        finally { Release(window); host.Application.Visible = false; }
    }
}

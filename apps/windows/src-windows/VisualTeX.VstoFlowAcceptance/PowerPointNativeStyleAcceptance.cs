using System.Diagnostics;
using System.Drawing;
using System.IO.Compression;
using System.Reflection;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Office.Core;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointNativeStyleAcceptance(string artifactRoot)
    {
        if (Process.GetProcessesByName("POWERPNT").Length != 0)
            throw new InvalidOperationException("Native style acceptance requires no existing PowerPoint process; user processes are never closed.");
        var startedUtc = DateTime.UtcNow;
        PowerPoint.Application? app = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPointFormulaService? service = null;
        var checks = new List<object>();
        var observations = new List<object>();
        var ownedPid = 0; DateTime ownedStart = default; var failed = 0;
        var m = (XNamespace)"http://schemas.openxmlformats.org/officeDocument/2006/math";
        var a = (XNamespace)"http://schemas.openxmlformats.org/drawingml/2006/main";
        void Check(string name, Action check)
        {
            try { check(); checks.Add(new { name, passed = true, error = (string?)null }); Console.WriteLine("PASS " + name); }
            catch (Exception error) { failed++; checks.Add(new { name, passed = false, error = error.ToString() }); Console.WriteLine("FAIL " + name + ": " + error.Message); }
        }
        string Capture(string name)
        {
            var path = Path.Combine(artifactRoot, name + ".pptx");
            presentation!.SaveCopyAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            string omml;
            using (var zip = ZipFile.OpenRead(path))
            using (var stream = zip.GetEntry("ppt/slides/slide1.xml")!.Open())
            {
                var doc = XDocument.Load(stream);
                omml = doc.Descendants(m + "oMath").Single().ToString(SaveOptions.DisableFormatting);
            }
            File.WriteAllText(Path.Combine(artifactRoot, name + ".omml.xml"), omml);
            var png = Path.Combine(artifactRoot, name + ".png");
            slide!.Export(png, "PNG", 1280, 720);
            var red = 0; var blue = 0;
            using (var bitmap = new Bitmap(png))
                for (var y = 0; y < bitmap.Height; y++) for (var x = 0; x < bitmap.Width; x++)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (pixel.R > 150 && pixel.G < 100 && pixel.B < 100) red++;
                    if (pixel.B > 150 && pixel.R < 100 && pixel.G < 150) blue++;
                }
            var runs = XElement.Parse(omml).Descendants(m + "r").Select(r => new
            {
                text = (string?)r.Element(m + "t"),
                style = r.Element(a + "rPr")?.ToString(SaveOptions.DisableFormatting),
                mathStyle = r.Element(m + "rPr")?.ToString(SaveOptions.DisableFormatting),
            }).ToArray();
            observations.Add(new { name, shapeId = shape!.Id, text = shape.TextFrame2.TextRange.Text, runs, redPixels = red, bluePixels = blue, pptx = path, png });
            Console.WriteLine(name + ": " + JsonSerializer.Serialize(new { runs, redPixels = red, bluePixels = blue }));
            return omml;
        }
        string Reverse(string omml)
        {
            var kernel = typeof(PowerPointFormulaService).Assembly.GetType("VisualTeX.PowerPointVsto.PowerPointNativeOmmlKernel", true)!;
            return (string)kernel.GetMethod("ReadNativeMathMl", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, new object[] { omml })!;
        }
        OfficeSessionDocument Create(string mathMl)
        {
            var session = CreatePowerPointAcceptanceSession("create", "wordOmml", Guid.NewGuid().ToString(), service!.CaptureSlideInsertionTarget().ObjectId, null);
            session.Lines = new List<FormulaLine> { new() { Id = Guid.NewGuid().ToString(), Latex = "x+y" } };
            session.ExportResult!.MathMl = mathMl;
            return session;
        }
        try
        {
            app = new PowerPoint.Application();
            GetWindowThreadProcessId(new IntPtr(app.HWND), out var pid);
            var process = Process.GetProcessById((int)pid);
            if (process.ProcessName != "POWERPNT" || process.StartTime.ToUniversalTime() < startedUtc.AddSeconds(-2) || app.Presentations.Count != 0)
                throw new InvalidOperationException("Cannot establish exclusive PowerPoint ownership.");
            ownedPid = process.Id; ownedStart = process.StartTime.ToUniversalTime();
            app.Visible = MsoTriState.msoTrue;
            presentation = app.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            app.ActiveWindow.View.GotoSlide(1);
            service = new PowerPointFormulaService(app);
            var colored = "<math xmlns='http://www.w3.org/1998/Math/MathML'><mrow><mi mathcolor='#FF0000'>x</mi><mo>+</mo><mi mathcolor='#0000FF'>y</mi></mrow></math>";
            var result = service.InsertOmml(Create(colored));
            shape = slide.Shapes[result.ObjectId];
            var forward = Capture("01-colored-insert");
            Check("Forward red/blue MathML survives native insertion", () =>
            {
                AssertContains(forward, "FF0000", "Red MathML color missing from actual PPTX.");
                AssertContains(forward, "0000FF", "Blue MathML color missing from actual PPTX.");
            });
            // Author formatting through the native host on this owned simple equation.
            var range = shape.TextFrame2.TextRange;
            var text = range.Text;
            var elements = System.Globalization.StringInfo.GetTextElementEnumerator(text);
            var xIndex = -1; var yIndex = -1; var xLength = 0; var yLength = 0;
            while (elements.MoveNext())
            {
                var element = elements.GetTextElement();
                var normalized = element.Normalize(System.Text.NormalizationForm.FormKC);
                if (normalized == "x") { xIndex = elements.ElementIndex; xLength = element.Length; }
                if (normalized == "y") { yIndex = elements.ElementIndex; yLength = element.Length; }
            }
            AssertTrue(xIndex >= 0 && yIndex >= 0, "Native fixture does not contain x and y.");
            var xRange = range.Characters[xIndex + 1, xLength];
            var yRange = range.Characters[yIndex + 1, yLength];
            xRange.Font.Name = "Arial"; xRange.Font.Size = 18; xRange.Font.Bold = MsoTriState.msoFalse; xRange.Font.Italic = MsoTriState.msoFalse;
            xRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
            yRange.Font.Name = "Georgia"; yRange.Font.Size = 32; yRange.Font.Bold = MsoTriState.msoTrue; yRange.Font.Italic = MsoTriState.msoTrue;
            yRange.Font.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.Blue);
            Release(xRange); Release(yRange); Release(range);
            var native = Capture("02-native-authored-styles");
            shape.Select(MsoTriState.msoTrue);
            var selected = service.LoadNativeOmmlSelection(service.ReadSelection());
            var recovered = Reverse(native);
            File.WriteAllText(Path.Combine(artifactRoot, "02-recovered.mathml.xml"), recovered);
            File.WriteAllText(Path.Combine(artifactRoot, "02-recovered.latex.txt"), selected.Metadata!.Latex);
            observations.Add(new { name = "recovered-source", mathMl = recovered, latex = selected.Metadata.Latex, selected.Metadata.FontSizePt, selected.Metadata.FormulaLetterFont });
            Check("Native red/blue color source recovery", () =>
            {
                AssertContains(recovered.ToUpperInvariant(), "#FF0000", "Native red missing from recovered source.");
                AssertContains(recovered.ToUpperInvariant(), "#0000FF", "Native blue missing from recovered source.");
            });
            var nativeRuns = XElement.Parse(native).Descendants(m + "r").ToArray();
            var actualFaces = nativeRuns.Select(r => (string?)r.Element(a + "rPr")?.Element(a + "latin")?.Attribute("typeface")).Distinct().ToArray();
            var actualSizes = nativeRuns.Select(r => (string?)r.Element(a + "rPr")?.Attribute("sz")).Distinct().ToArray();
            observations.Add(new { name = "native-format-preconditions", requestedFaces = new[] { "Arial", "Georgia" }, actualFaces,
                requestedSizes = new[] { 1800, 3200 }, actualSizes,
                customFaceFixtureEstablished = native.Contains("Arial") && native.Contains("Georgia"),
                mixedSizeFixtureEstablished = native.Contains("sz=\"1800\"") && native.Contains("sz=\"3200\"") });
            // The baseline retained the original failed expectations. The host
            // did not establish those requested formats, so do not misreport
            // their absence as a plugin recovery failure or an accepted case.
            checks.Add(new { name = "Arbitrary per-run fonts and mixed sizes", passed = (bool?)null,
                status = "not established by native COM authoring", actualFaces, actualSizes });
            var edit = CreatePowerPointAcceptanceSession("edit", "wordOmml", selected.FormulaId!, selected.ObjectId, selected.Metadata);
            edit.SourceDocumentId = selected.DocumentId;
            edit.Lines = selected.Metadata.Lines;
            edit.ExportResult!.MathMl = recovered;
            edit.FontSizePt = selected.Metadata.FontSizePt ?? 20;
            var replaced = service.ReplaceOmml(edit);
            Release(shape); shape = slide.Shapes[replaced.ObjectId];
            var after = Capture("03-recovered-source-reinsert");
            Check("Native style roundtrip colors", () =>
            {
                AssertContains(after, "FF0000", "Roundtrip dropped red.");
                AssertContains(after, "0000FF", "Roundtrip dropped blue.");
            });
            Check("Native style roundtrip actual host face and uniform size", () =>
            {
                foreach (var face in actualFaces.Where(f => !string.IsNullOrEmpty(f))) AssertContains(after, face!, "Actual native face lost.");
                foreach (var size in actualSizes.Where(s => !string.IsNullOrEmpty(s))) AssertContains(after, "sz=\"" + size + "\"", "Actual native size lost.");
            });
            Check("Native bold-italic Unicode semantic roundtrip", () =>
            {
                var styledY = char.ConvertFromUtf32(0x1D49A);
                AssertContains(native, styledY, "Native bold-italic Y fixture missing.");
                AssertContains(recovered, styledY, "Source lost bold-italic mathematical scalar.");
                AssertContains(after, styledY, "Output lost bold-italic mathematical scalar.");
            });
            Check("Native upright x semantic roundtrip", () =>
            {
                AssertContains(recovered, "mathvariant=\"normal\"", "Source lost upright x.");
                var xRun = XElement.Parse(after).Descendants(m + "r").Single(r =>
                    ((string?)r.Element(m + "t"))?.Normalize(System.Text.NormalizationForm.FormKC) == "x");
                AssertEqual("x", (string?)xRun.Element(m + "t"), "Output changed upright x into a styled mathematical scalar.");
                AssertTrue((string?)xRun.Element(m + "rPr")?.Element(m + "sty")?.Attribute(m + "val") == "p"
                    || xRun.Element(m + "rPr")?.Element(m + "nor") is not null,
                    "Output lacks native upright semantics.");
            });
            var reopenPath = Path.Combine(artifactRoot, "03-recovered-source-reinsert.pptx");
            Release(shape); shape = null; Release(slide); slide = null;
            presentation.Close(); Release(presentation);
            presentation = app.Presentations.Open(reopenPath, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            slide = presentation.Slides[1]; shape = slide.Shapes[1];
            var reopened = Capture("04-reopened-styles");
            Check("Native styled save and reopen", () =>
            {
                AssertContains(reopened, "FF0000", "Reopen lost red.");
                AssertContains(reopened, "0000FF", "Reopen lost blue.");
                AssertContains(reopened, char.ConvertFromUtf32(0x1D49A), "Reopen lost bold-italic y.");
                var xRun = XElement.Parse(reopened).Descendants(m + "r").Single(r =>
                    ((string?)r.Element(m + "t"))?.Normalize(System.Text.NormalizationForm.FormKC) == "x");
                AssertEqual("x", (string?)xRun.Element(m + "t"), "Reopen lost upright x.");
                AssertEqual("p", (string?)xRun.Element(m + "rPr")?.Element(m + "sty")?.Attribute(m + "val"), "Reopen lost native upright property.");
            });
            // Isolated API probe, not a Ribbon click and not product acceptance.
            shape.Delete(); Release(shape);
            shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 220, 72);
            shape.TextFrame.TextRange.Text = "x+y";
            shape.TextFrame.TextRange.Select();
            app.CommandBars.ExecuteMso(app.CommandBars.GetEnabledMso("EquationInsertNew")
                ? "EquationInsertNew" : "InsertBuildingBlocksEquationsGallery");
            var probeText = shape.TextFrame2.TextRange.Text;
            var probeElements = System.Globalization.StringInfo.GetTextElementEnumerator(probeText);
            AssertTrue(probeElements.MoveNext() && probeElements.GetTextElement().Normalize(System.Text.NormalizationForm.FormKC) == "x",
                "Upright probe cannot identify its exact own first token.");
            var probeRange = shape.TextFrame2.TextRange.Characters[1, probeElements.GetTextElement().Length];
            probeRange.Font.Italic = MsoTriState.msoFalse;
            Release(probeRange);
            shape.TextFrame2.TextRange.Select();
            if (app.CommandBars.GetEnabledMso("EquationProfessional")) app.CommandBars.ExecuteMso("EquationProfessional");
            var probeOmml = Capture("05-upright-after-zone-probe");
            var probeX = XElement.Parse(probeOmml).Descendants(m + "r").First();
            observations.Add(new { name = "after-zone-upright-probe", before = probeText,
                uprightPreserved = (string?)probeX.Element(m + "t") == "x",
                scope = "owned fixture API probe; zero physical Ribbon clicks" });
        }
        finally
        {
            try { service?.Dispose(); } catch (Exception error) { Console.WriteLine("Kernel cleanup: " + error); }
            Release(shape); Release(slide);
            if (presentation is not null) { try { presentation.Saved = MsoTriState.msoTrue; presentation.Close(); } catch { } Release(presentation); }
            if (app is not null && ownedPid != 0)
            {
                try { if (Process.GetProcessById(ownedPid).StartTime.ToUniversalTime() == ownedStart && app.Presentations.Count == 0) app.Quit(); }
                catch (ArgumentException) { }
            }
            Release(app);
            File.WriteAllText(Path.Combine(artifactRoot, "native-styles.json"), JsonSerializer.Serialize(new { startedUtc, ownedPid, ownedStart, failed, checks, observations, actualRibbonClicks = 0 }, new JsonSerializerOptions { WriteIndented = true }));
        }
        if (failed != 0) throw new InvalidOperationException(failed + " native style checks failed; see native-styles.json.");
    }
}

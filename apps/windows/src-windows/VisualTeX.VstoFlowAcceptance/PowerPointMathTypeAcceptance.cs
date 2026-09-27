using System.Reflection;
using Microsoft.Office.Core;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointMathTypeAcceptance(string artifactRoot)
    {
        PowerPoint.Application? app = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        var assembly = typeof(PowerPointFormulaService).Assembly;
        try
        {
            app = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
            presentation = app.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            app.ActiveWindow.View.GotoSlide(1);
            using var service = new PowerPointFormulaService(app);
            var selection = service.ReadSelection();
            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString(), Mode = "create", Host = "powerpoint",
                FormulaId = Guid.NewGuid().ToString(), SourceDocumentId = selection.DocumentId,
                SourceObjectId = selection.ObjectId, ObjectMode = "mathTypeOle", DisplayMode = "block",
                Title = "MathType acceptance", CodeFormat = "raw", FontSizePt = 20,
                Lines = new() { new() { Id = Guid.NewGuid().ToString(), Latex = "x=1" } },
                ExportResult = new() { MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mn>1</mn></math>", Width = 80, Height = 32 },
            };
            using (var prepared = service.PrepareMathType(session))
            {
                var previewPath = (string)prepared.GetType().GetField("EmfPath", BindingFlags.Instance | BindingFlags.NonPublic)!.GetValue(prepared)!;
                File.Copy(previewPath, Path.Combine(artifactRoot, "native-preview.emf"), true);
                using var metafile = new System.Drawing.Imaging.Metafile(previewPath);
                using var bitmap = new System.Drawing.Bitmap(640, 240);
                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                {
                    graphics.Clear(System.Drawing.Color.White);
                    graphics.DrawImage(metafile, 0, 0, 640, 240);
                }
                bitmap.Save(Path.Combine(artifactRoot, "native-preview.png"));
                service.ApplyMathType(session, prepared);
            }
            shape = slide.Shapes[1];
            if (shape.OLEFormat.ProgID != "Equation.DSMT4") throw new Exception("Inserted object is not native MathType.");
            shape.Left = 113; shape.Top = 87;
            shape.LockAspectRatio = MsoTriState.msoFalse;
            shape.Width *= 1.3f; shape.Height *= 1.2f;
            var originalWidth = shape.Width;
            var originalHeight = shape.Height;
            shape.Tags.Add("AcceptanceTag", "preserved");
            var action = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
            action.Action = PowerPoint.PpActionType.ppActionHyperlink;
            action.Hyperlink.Address = "https://example.invalid/formula";
            Release(action);
            var effect = slide.TimeLine.MainSequence.AddEffect(shape, PowerPoint.MsoAnimEffect.msoAnimEffectFade);
            effect.Timing.Duration = 1.75f;
            shape.Select(MsoTriState.msoTrue);
            var source = service.LoadMathTypeSelection(service.ReadSelection());
            if (source.Metadata is null || !source.Metadata.Latex.Contains("x")) throw new Exception("MathType source not read.");
            Console.WriteLine("PASS create and read native MathType: " + source.Metadata.Latex);
            session.Mode = "edit";
            session.FormulaId = source.FormulaId!;
            session.SourceObjectId = source.ObjectId;
            session.OriginalMetadata = source.Metadata;
            session.Lines[0].Latex = @"\frac{a}{b}";
            session.ExportResult!.MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>";
            using (var prepared = service.PrepareMathType(session))
                service.ApplyMathType(session, prepared);
            Release(shape);
            shape = slide.Shapes[1];
            if (slide.Shapes.Count != 1 || Math.Abs(shape.Left - 113) > 0.05 || Math.Abs(shape.Top - 87) > 0.05
                || Math.Abs(shape.Rotation) > 0.05 || shape.Tags["AcceptanceTag"] != "preserved")
                throw new Exception("Edit changed PowerPoint geometry, tags or object count.");
            if (effect.Shape.Id != shape.Id || Math.Abs(effect.Timing.Duration - 1.75f) > 0.001)
                throw new Exception("Edit lost animation target or timing.");
            action = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
            if (action.Hyperlink.Address != "https://example.invalid/formula") throw new Exception("Edit lost its hyperlink.");
            Release(action);
            Console.WriteLine("PASS replace keeps position, native OLE orientation, tags and animation");
            shape.Select(MsoTriState.msoTrue);
            var edited = service.LoadMathTypeSelection(service.ReadSelection());
            if (!edited.Metadata!.Latex.Contains("frac")) throw new Exception("Edited source did not round-trip.");
            var path = Path.Combine(artifactRoot, "PowerPoint-MathType-Acceptance.pptx");
            presentation.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            Release(effect); Release(shape); shape = null; Release(slide); slide = null;
            presentation.Close(); Release(presentation);
            presentation = app.Presentations.Open(path, WithWindow: MsoTriState.msoTrue);
            slide = presentation.Slides[1]; shape = slide.Shapes[1];
            shape.Select(MsoTriState.msoTrue);
            var reopened = service.LoadMathTypeSelection(service.ReadSelection());
            if (!reopened.Metadata!.Latex.Contains("frac")) throw new Exception("Saved MathType object did not reopen.");
            slide.Export(Path.Combine(artifactRoot, "PowerPoint-MathType-Acceptance.png"), "PNG", 1280, 720);
            using (var bitmap = new System.Drawing.Bitmap(Path.Combine(artifactRoot, "PowerPoint-MathType-Acceptance.png")))
            {
                var ink = 0;
                for (var y = 0; y < bitmap.Height; y++)
                for (var x = 0; x < bitmap.Width; x++)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (pixel.R < 220 || pixel.G < 220 || pixel.B < 220) ink++;
                }
                if (ink < 50) throw new Exception("PowerPoint's saved MathType preview is blank.");
            }
            Console.WriteLine("PASS save/reopen and slide render");

            session.FormulaId = reopened.FormulaId!;
            session.SourceDocumentId = reopened.DocumentId;
            session.SourceObjectId = reopened.ObjectId;
            session.OriginalMetadata = reopened.Metadata;
            session.Lines[0].Latex = "x=1";
            session.ExportResult!.MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mn>1</mn></math>";
            using (var prepared = service.PrepareMathType(session))
                service.ApplyMathType(session, prepared);
            Release(shape); shape = slide.Shapes[1];
            if (Math.Abs(shape.Width - originalWidth) > 0.1 || Math.Abs(shape.Height - originalHeight) > 0.1)
                throw new Exception($"Repeated edits drifted scale: {originalWidth}x{originalHeight} -> {shape.Width}x{shape.Height}.");
            Console.WriteLine("PASS independent X/Y scale survives edit-save-reopen-edit roundtrip");

            shape.Select(MsoTriState.msoTrue);
            var unchanged = service.LoadMathTypeSelection(service.ReadSelection());
            session.FormulaId = unchanged.FormulaId!;
            session.SourceObjectId = unchanged.ObjectId;
            session.OriginalMetadata = unchanged.Metadata;
            var unchangedId = shape.Id;
            using (var prepared = service.PrepareMathType(session))
            {
                session.SourceDocumentId = "a-different-presentation";
                var rejected = false;
                try { service.ApplyMathType(session, prepared); }
                catch (InvalidOperationException) { rejected = true; }
                finally { session.SourceDocumentId = unchanged.DocumentId; }
                if (!rejected || slide.Shapes.Count != 1 || slide.Shapes[1].Id != unchangedId)
                    throw new Exception("A stale document target was not rejected safely.");
            }
            Console.WriteLine("PASS stale document rejection leaves the original object intact");

            // Read an existing real-world MathType storage without modifying
            // the fixture or manufacturing its MTEF with VisualTeX's encoder.
            var nativeFixture = Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_SOURCE_DOCX");
            if (string.IsNullOrWhiteSpace(nativeFixture))
            {
                Console.WriteLine("SKIP third-party storage: set VISUALTEX_ACCEPTANCE_MATHTYPE_SOURCE_DOCX.");
                return;
            }
            var storageType = assembly.GetType("VisualTeX.WordVsto.MathTypeOleStorage", true)!;
            byte[]? nativeStorage = null;
            using (var archive = System.IO.Compression.ZipFile.OpenRead(nativeFixture))
            {
                foreach (var item in archive.Entries.Where(item => item.FullName.StartsWith("word/embeddings/", StringComparison.Ordinal)))
                {
                    using var entry = item.Open();
                    using var buffer = new MemoryStream();
                    entry.CopyTo(buffer);
                    var candidate = buffer.ToArray();
                    if (!(bool)storageType.GetMethod("LooksLikeMathTypeCompoundFile", BindingFlags.Static | BindingFlags.NonPublic)!
                            .Invoke(null, new object[] { candidate })!) continue;
                    nativeStorage = candidate;
                    break;
                }
            }
            if (nativeStorage is null) throw new Exception("The supplied fixture has no native MathType object.");
            using (var clipboard = (IDisposable)storageType.GetMethod("BeginStandaloneClipboardTransaction", BindingFlags.Static | BindingFlags.NonPublic)!
                       .Invoke(null, new object[] { nativeStorage, Path.Combine(artifactRoot, "native-preview.emf"), 40f, 19f })!)
            {
                var pasted = slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteOLEObject);
                Release(shape); shape = pasted[1]; Release(pasted);
            }
            shape.Select(MsoTriState.msoTrue);
            var nativeSelection = service.LoadMathTypeSelection(service.ReadSelection());
            if (string.IsNullOrWhiteSpace(nativeSelection.Metadata?.Latex)) throw new Exception("Could not read third-party MathType formula.");
            Console.WriteLine("Native fixture source: " + nativeSelection.Metadata!.Latex);
            session.FormulaId = nativeSelection.FormulaId!;
            session.SourceObjectId = nativeSelection.ObjectId;
            session.OriginalMetadata = nativeSelection.Metadata;
            session.Lines[0].Latex = @"\sqrt{y}";
            session.ExportResult.MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msqrt><mi>y</mi></msqrt></math>";
            using (var prepared = service.PrepareMathType(session))
                service.ApplyMathType(session, prepared);
            var nativeEdited = service.LoadMathTypeSelection(service.ReadSelection());
            if (!nativeEdited.Metadata!.Latex.Contains("sqrt")) throw new Exception("Installed-MathType-authored formula did not edit.");
            Console.WriteLine("PASS read/edit existing third-party MathType storage");
            Release(shape); shape = null;
            Release(slide);
            slide = presentation.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            app.ActiveWindow.View.GotoSlide(2);
            var rootTarget = service.CaptureSlideInsertionTarget();
            session.Mode = "create";
            session.FormulaId = Guid.NewGuid().ToString();
            session.SourceObjectId = rootTarget.ObjectId;
            session.OriginalMetadata = null;
            session.Lines[0].Latex = @"\sqrt[3]{x}";
            session.ExportResult.MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mroot><mi>x</mi><mn>3</mn></mroot></math>";
            using (var prepared = service.PrepareMathType(session))
                service.ApplyMathType(session, prepared);
            slide.Export(Path.Combine(artifactRoot, "PowerPoint-MathType-CubeRoot.png"), "PNG", 1280, 720);
            Console.WriteLine("PASS native cube-root creation (visual evidence exported)");
            presentation.SaveAs(Path.Combine(artifactRoot, "PowerPoint-MathType-Native-Edited.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
        }
        catch
        {
            try { presentation?.SaveCopyAs(Path.Combine(artifactRoot, "failure.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation); } catch { }
            throw;
        }
        finally
        {
            Release(shape); Release(slide);
            if (presentation is not null) { try { presentation.Close(); } catch { } }
            Release(presentation);
            // PowerPoint may reuse the user's process; close only this fixture.
            Release(app);
        }
    }
}

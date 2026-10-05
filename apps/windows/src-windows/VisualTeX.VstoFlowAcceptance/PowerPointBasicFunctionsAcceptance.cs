using System.Diagnostics;
using System.Text.Json;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointBasicFunctionsAcceptance(string artifactRoot)
    {
        // This small host test never changes add-in registry/settings. Refuse
        // to attach to an existing PowerPoint, even if COM would share it.
        var before = Process.GetProcessesByName("POWERPNT")
            .Select(p => new { p.Id, StartedUtc = p.StartTime.ToUniversalTime() }).ToArray();
        if (before.Length != 0)
            throw new InvalidOperationException("Close no user presentation: basic acceptance requires no existing PowerPoint process.");
        var startedUtc = DateTime.UtcNow;
        var checks = new List<object>();
        PowerPoint.Application? application = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPointFormulaService? service = null;
        int ownedPid = 0;
        DateTime ownedStart = default;
        var failed = 0;
        void Check(string name, Action action)
        {
            try
            {
                action();
                checks.Add(new { name, passed = true, error = (string?)null });
                Console.WriteLine("PASS " + name);
            }
            catch (Exception error)
            {
                failed++;
                checks.Add(new { name, passed = false, error = error.ToString() });
                Console.WriteLine("FAIL " + name + ": " + error);
            }
        }
        try
        {
            application = new PowerPoint.Application();
            GetWindowThreadProcessId(new IntPtr(application.HWND), out var pid);
            var process = Process.GetProcessById((int)pid);
            if (process.ProcessName != "POWERPNT" || process.StartTime.ToUniversalTime() < startedUtc.AddSeconds(-2)
                || application.Presentations.Count != 0)
                throw new InvalidOperationException("PowerPoint COM ownership could not be established; no document was created.");
            ownedPid = process.Id;
            ownedStart = process.StartTime.ToUniversalTime();
            application.Visible = MsoTriState.msoTrue;
            presentation = application.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            application.ActiveWindow.View.GotoSlide(1);
            service = new PowerPointFormulaService(application);
            var svgPath = Path.Combine(artifactRoot, "fixture.svg");
            File.WriteAllText(svgPath, CreateSvg(220, 72));
            Check("SVG insert, select and edit", () =>
            {
                var session = CreatePowerPointAcceptanceSession("create", "crossPlatformPicture",
                    Guid.NewGuid().ToString(), service.CaptureSlideInsertionTarget().ObjectId, null);
                var result = service.Insert(session, svgPath);
                shape = slide.Shapes[result.ObjectId];
                shape.Select(MsoTriState.msoTrue);
                var selected = service.ReadSelection();
                AssertTrue(selected.Metadata is not null, "Inserted SVG has no edit metadata.");
                var edit = CreatePowerPointAcceptanceSession("edit", "crossPlatformPicture",
                    selected.FormulaId!, selected.ObjectId, selected.Metadata);
                var edited = service.Replace(edit, svgPath);
                Release(shape);
                shape = slide.Shapes[edited.ObjectId];
                shape.Select(MsoTriState.msoTrue);
                AssertEqual(session.FormulaId, service.ReadSelection().FormulaId, "SVG edit lost formula identity.");
            });
            Check("SVG font size and delete", () =>
            {
                AssertTrue(shape is not null, "SVG fixture missing.");
                shape!.Select(MsoTriState.msoTrue);
                AssertNear(24, service.SetSelectedFormulaFontSize(24), 0.01f, "SVG font size did not update.");
                var count = slide.Shapes.Count;
                service.DeleteSelectedFormula();
                Release(shape); shape = null;
                AssertEqual(count - 1, slide.Shapes.Count, "Delete did not remove exactly one SVG.");
            });
            Check("OLE insert, edit and SVG conversion", () =>
            {
                var pngFixturePath = Path.Combine(artifactRoot, "fixture.png");
                WriteAcceptancePng(pngFixturePath, "a/b + sqrt(x)", 440, 144);
                var session = CreatePowerPointAcceptanceSession("create", "nativeOle",
                    Guid.NewGuid().ToString(), service.CaptureSlideInsertionTarget().ObjectId, null);
                session.ExportResult!.PngBase64 = Convert.ToBase64String(File.ReadAllBytes(pngFixturePath));
                using var materializer = new VisualTeXSessionClient();
                var pngPath = materializer.MaterializePng(session);
                var controlledSvgPath = materializer.MaterializeSvg(session);
                var emfPath = CreatePowerPointAcceptanceEmf(controlledSvgPath, 220, 72);
                try
                {
                    var inserted = service.InsertOle(session, pngPath, emfPath);
                    shape = slide.Shapes[inserted.ObjectId];
                    AssertPowerPointOle(shape, "OLE insert failed.");
                    shape.Select(MsoTriState.msoTrue);
                    var selected = service.ReadSelection();
                    var edit = CreatePowerPointAcceptanceSession("edit", "nativeOle",
                        selected.FormulaId!, selected.ObjectId, selected.Metadata);
                    var updated = service.ReplaceOle(edit, pngPath, emfPath);
                    Release(shape); shape = slide.Shapes[updated.ObjectId];
                    AssertPowerPointOle(shape, "OLE edit failed.");
                    shape.Select(MsoTriState.msoTrue);
                    selected = service.ReadSelection();
                    var toSvg = CreatePowerPointAcceptanceSession("edit", "crossPlatformPicture",
                        selected.FormulaId!, selected.ObjectId, selected.Metadata);
                    var converted = service.Replace(toSvg, svgPath);
                    Release(shape); shape = slide.Shapes[converted.ObjectId];
                    shape.Select(MsoTriState.msoTrue);
                    AssertTrue(service.ReadSelection().Metadata is not null, "OLE to SVG lost editable source.");
                    service.DeleteSelectedFormula();
                    Release(shape); shape = null;
                }
                finally { try { File.Delete(emfPath); File.Delete(pngPath); File.Delete(controlledSvgPath); } catch { } }
            });
            Check("Native OMML insert", () =>
            {
                var session = CreatePowerPointAcceptanceSession("create", "wordOmml",
                    Guid.NewGuid().ToString(), service.CaptureSlideInsertionTarget().ObjectId, null);
                var result = service.InsertOmml(session);
                shape = slide.Shapes[result.ObjectId];
                AssertTrue(IsPowerPointNativeEquation(shape), "Insert produced no native equation.");
                AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, session.FormulaId);
                shape.AlternativeText = "user-authored equation description";
                shape.Title = "User equation title";
            });
            Check("Native OMML shape selection and transient source", () =>
            {
                AssertTrue(shape is not null, "Native fixture missing.");
                shape!.Select(MsoTriState.msoTrue);
                var selected = service.ReadSelection();
                AssertTrue(PowerPointFormulaService.IsNativeOmmlReference(selected.ObjectId), "Native source candidate has no exact identity.");
                object nativeZones = ((dynamic)shape.TextFrame2.TextRange).MathZones();
                Console.WriteLine($"Native zone inventory: count={((dynamic)nativeZones).Count}, length={((dynamic)nativeZones).Length}");
                Release(nativeZones);
                selected = service.LoadNativeOmmlSelection(selected);
                AssertTrue(selected.Metadata is not null && selected.ObjectMode == "wordOmml", "Explicit native source was not recognized.");
                AssertContains(selected.Metadata?.Latex, @"\frac", "Native fraction source missing.");
                AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, selected.FormulaId!);
            });
            Check("Native OMML font, edit and SVG round trip", () =>
            {
                AssertTrue(shape is not null, "Native fixture missing.");
                shape!.Select(MsoTriState.msoTrue);
                AssertNear(24, service.SetSelectedFormulaFontSize(24), 0.01f, "Native font size failed.");
                var selected = service.LoadNativeOmmlSelection(service.ReadSelection());
                var edit = CreatePowerPointAcceptanceSession("edit", "wordOmml",
                    selected.FormulaId!, selected.ObjectId, selected.Metadata);
                edit.SourceDocumentId = selected.DocumentId;
                var edited = service.ReplaceOmml(edit);
                Release(shape); shape = slide.Shapes[edited.ObjectId];
                shape.Select(MsoTriState.msoTrue);
                selected = service.LoadNativeOmmlSelection(service.ReadSelection());
                var toPicture = CreatePowerPointAcceptanceSession("edit", "crossPlatformPicture",
                    selected.FormulaId!, selected.ObjectId, selected.Metadata);
                toPicture.SourceDocumentId = selected.DocumentId;
                var converted = service.Replace(toPicture, svgPath);
                Release(shape); shape = slide.Shapes[converted.ObjectId];
                shape.Select(MsoTriState.msoTrue);
                var picture = service.ReadSelection();
                AssertTrue(picture.Metadata is not null, "Native to SVG lost edit metadata.");
                var back = CreatePowerPointAcceptanceSession("edit", "wordOmml",
                    picture.FormulaId!, picture.ObjectId, picture.Metadata);
                back.SourceDocumentId = picture.DocumentId;
                var restored = service.ReplaceOmml(back);
                Release(shape); shape = slide.Shapes[restored.ObjectId];
                AssertTrue(IsPowerPointNativeEquation(shape), "SVG to native failed.");
                AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, picture.FormulaId!);
                AssertEqual("user-authored equation description", shape.AlternativeText, "Round trip lost user alt text.");
                AssertEqual("User equation title", shape.Title, "Round trip lost user title.");
            });
            Check("Explicit prose never reuses the old formula", () =>
            {
                var fixture = CreatePowerPointAcceptanceSession("create", "crossPlatformPicture",
                    Guid.NewGuid().ToString(), service.CaptureSlideInsertionTarget().ObjectId, null);
                var inserted = service.Insert(fixture, svgPath);
                var old = slide.Shapes[inserted.ObjectId];
                PowerPoint.Shape? prose = null;
                try
                {
                    old.Select(MsoTriState.msoTrue);
                    AssertTrue(service.ReadSelection().Metadata is not null, "Formula selection precondition failed.");
                    prose = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 250, 40);
                    prose.TextFrame.TextRange.Text = "ordinary text";
                    prose.TextFrame.TextRange.Select();
                    object proseZones = ((dynamic)prose.TextFrame2.TextRange).MathZones();
                    Console.WriteLine($"Prose zone inventory: count={((dynamic)proseZones).Count}, length={((dynamic)proseZones).Length}");
                    Release(proseZones);
                    var addIn = new ThisAddIn();
                    var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                    var resolved = (OfficeSelection)typeof(ThisAddIn).GetMethod("ResolveFormulaSelection", flags)!
                        .Invoke(addIn, new object[] { service })!;
                    AssertTrue(resolved.Metadata is null && resolved.FormulaId is null,
                        "Explicit prose resolved to the previously selected formula.");
                    AssertEqual("ordinary text", prose.TextFrame.TextRange.Text, "Selection check changed prose.");
                    application.ActiveWindow.Selection.Unselect();
                    resolved = (OfficeSelection)typeof(ThisAddIn).GetMethod("ResolveFormulaSelection", flags)!
                        .Invoke(addIn, new object[] { service })!;
                    AssertTrue(resolved.Metadata is null && resolved.FormulaId is null,
                        "Blank-slide deselection reused the old formula.");
                }
                finally { if (prose is not null) prose.Delete(); old.Delete(); Release(prose); Release(old); }
            });
            Check("Changed native source is retained and repeated Edit keeps its owner", () =>
            {
                shape!.Select(MsoTriState.msoTrue);
                var selected = service.LoadNativeOmmlSelection(service.ReadSelection());
                var flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                var sourceField = typeof(PowerPointFormulaService).GetField("_nativeOmmlSource", flags)!;
                var source = sourceField.GetValue(service);
                var addIn = new ThisAddIn();
                using var lifetime = new CancellationTokenSource();
                typeof(ThisAddIn).GetField("_formulaService", flags)!.SetValue(addIn, service);
                typeof(ThisAddIn).GetField("_lifetime", flags)!.SetValue(addIn, lifetime);
                var gate = (SemaphoreSlim)typeof(ThisAddIn).GetField("_operationGate", flags)!.GetValue(addIn)!;
                gate.Wait();
                typeof(ThisAddIn).GetMethod("BeginSelectedSession", flags)!.Invoke(addIn, new object?[] { null, false });
                Thread.Sleep(2200);
                AssertTrue(ReferenceEquals(source, sourceField.GetValue(service)), "Repeated Edit replaced the active source snapshot.");
                gate.Release();
                shape.TextFrame2.TextRange.Font.Size = 31;
                var count = slide.Shapes.Count;
                var oldId = shape.Id;
                var stale = CreatePowerPointAcceptanceSession("edit", "crossPlatformPicture",
                    selected.FormulaId!, selected.ObjectId, selected.Metadata);
                stale.SourceDocumentId = selected.DocumentId;
                var refused = false;
                try { service.Replace(stale, svgPath); } catch (InvalidOperationException) { refused = true; }
                AssertTrue(refused, "Changed native source was replaced by a stale edit.");
                AssertEqual(count, slide.Shapes.Count, "Stale native edit changed shape count.");
                AssertEqual(oldId, shape.Id, "Stale native edit replaced the source.");
                AssertNear(31, shape.TextFrame2.TextRange.Font.Size, 0.01f, "Stale native edit lost current native formatting.");
            });
            Check("Mixed native textbox is never replaced as a whole", () =>
            {
                AssertTrue(shape is not null, "Native fixture missing.");
                shape!.TextFrame.TextRange.InsertAfter(" neighbor");
                shape.Select(MsoTriState.msoTrue);
                var beforeText = shape.TextFrame.TextRange.Text;
                var neighbor = shape.TextFrame.TextRange.Characters(beforeText.Length - 8, 9);
                neighbor.Font.Size = 18;
                object mixedZones = ((dynamic)shape.TextFrame2.TextRange).MathZones();
                Console.WriteLine($"Mixed zone inventory: count={((dynamic)mixedZones).Count}, length={((dynamic)mixedZones).Length}, whole={shape.TextFrame2.TextRange.Length}");
                ((dynamic)mixedZones).Select();
                AssertNear(30, service.SetSelectedFormulaFontSize(30), 0.01f, "Selected math-zone font failed.");
                AssertNear(18, neighbor.Font.Size, 0.01f, "Native font change modified adjacent prose.");
                Release(mixedZones);
                Release(neighbor);
                shape.Select(MsoTriState.msoTrue);
                var beforeCount = slide.Shapes.Count;
                var refused = false;
                try { service.LoadNativeOmmlSelection(service.ReadSelection()); }
                catch (InvalidOperationException) { refused = true; }
                AssertTrue(refused, "Mixed text was accepted for whole-shape replacement.");
                AssertEqual(beforeText, shape.TextFrame.TextRange.Text, "Native source capture changed mixed text.");
                AssertEqual(beforeCount, slide.Shapes.Count, "Native source capture changed shapes.");
                shape.TextFrame.TextRange.Characters(beforeText.Length - 8, 9).Delete();
            });
            Check("Native OMML save and reopen", () =>
            {
                AssertTrue(shape is not null && IsPowerPointNativeEquation(shape), "Native fixture missing.");
                var path = Path.Combine(artifactRoot, "native-reopen.pptx");
                presentation.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoTrue);
                Release(shape); shape = null;
                Release(slide); slide = null;
                presentation.Close();
                Release(presentation); presentation = null;
                presentation = application.Presentations.Open(path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                slide = presentation.Slides[1];
                AssertEqual(1, slide.Shapes.Count, "Unexpected shape count after reopen.");
                shape = slide.Shapes[1];
                AssertTrue(IsPowerPointNativeEquation(shape), "Saved formula is no longer native OMML.");
                shape.Select(MsoTriState.msoTrue);
                var reopenedSource = service.LoadNativeOmmlSelection(service.ReadSelection());
                AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, reopenedSource.FormulaId!);
                AssertEqual("user-authored equation description", shape.AlternativeText, "Reopen lost user alt text.");
            });
        }
        finally
        {
            try { service?.Dispose(); } catch (Exception error) { Console.WriteLine("Kernel cleanup: " + error); }
            Release(shape);
            Release(slide);
            if (presentation is not null)
            {
                try { presentation.Saved = MsoTriState.msoTrue; presentation.Close(); } catch { }
                Release(presentation);
            }
            if (application is not null && ownedPid != 0)
            {
                try
                {
                    var current = Process.GetProcessById(ownedPid);
                    if (current.StartTime.ToUniversalTime() == ownedStart && application.Presentations.Count == 0)
                        application.Quit();
                    else Console.WriteLine("Owned PowerPoint acquired another presentation; left running.");
                }
                catch (ArgumentException) { }
            }
            Release(application);
            File.WriteAllText(Path.Combine(artifactRoot, "basic-functions.json"),
                JsonSerializer.Serialize(new { startedUtc, ownedPid, ownedStart, failed, checks },
                    new JsonSerializerOptions { WriteIndented = true }));
        }
        if (failed != 0)
            throw new InvalidOperationException(failed + " bounded PowerPoint basic checks failed; see basic-functions.json.");
    }
}

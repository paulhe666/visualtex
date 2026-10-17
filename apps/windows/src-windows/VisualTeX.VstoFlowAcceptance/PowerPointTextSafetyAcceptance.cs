using System.Diagnostics;
using System.Text.Json;
using System.Reflection;
using System.Xml.Linq;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // No Program.Main registry-suppression scope and no existing presentation.
    private static void RunPowerPointTextSafetyAcceptance(string artifactRoot)
    {
        if (Process.GetProcessesByName("POWERPNT").Length != 0)
            throw new InvalidOperationException("Text safety acceptance requires no existing PowerPoint; leave user processes untouched.");
        var startedUtc = DateTime.UtcNow;
        PowerPoint.Application? app = null;
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        PowerPointFormulaService? service = null;
        int ownedPid = 0;
        DateTime ownedStart = default;
        var checks = new List<object>();
        Exception? failure = null;
        string[] ReadHistory()
        {
            CommandBars? bars = null;
            CommandBar? standard = null;
            CommandBarControl? control = null;
            try
            {
                bars = app!.CommandBars;
                standard = bars["Standard"];
                control = standard.FindControl(Id: 128);
                var history = control as CommandBarComboBox
                    ?? throw new InvalidOperationException("Native PowerPoint Undo history is not a readable combo box.");
                if (!bars.GetEnabledMso("Undo")) return Array.Empty<string>();
                return Enumerable.Range(1, history.ListCount).Select(i => history.get_List(i)).ToArray();
            }
            finally { Release(control); Release(standard); Release(bars); }
        }
        try
        {
            app = new PowerPoint.Application();
            GetWindowThreadProcessId(new IntPtr(app.HWND), out var pid);
            using var process = Process.GetProcessById((int)pid);
            if (process.ProcessName != "POWERPNT" || process.StartTime.ToUniversalTime() < startedUtc.AddSeconds(-2)
                || app.Presentations.Count != 0)
                throw new InvalidOperationException("PowerPoint process ownership was not established; no document created.");
            ownedPid = process.Id;
            ownedStart = process.StartTime.ToUniversalTime();
            app.Visible = MsoTriState.msoTrue;
            presentation = app.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            app.ActiveWindow.View.GotoSlide(1);
            shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 30, 30, 500, 150);
            shape.TextFrame.TextRange.Text = "before SELECT after";
            shape.TextFrame.TextRange.Characters(1, 6).Font.Bold = MsoTriState.msoTrue;
            shape.TextFrame.TextRange.Characters(15, 5).Font.Italic = MsoTriState.msoTrue;
            app.StartNewUndoEntry();
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "before.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            var before = ReadHistory();
            app.StartNewUndoEntry();
            shape.TextFrame.TextRange.Characters(8, 6).Text = "changed";
            shape.TextFrame.TextRange.Characters(8, 7).Font.Size = 31;
            app.StartNewUndoEntry();
            var after = ReadHistory();
            checks.Add(new { name = "native Undo prefix", before, after });
            AssertTrue(after.Length > before.Length && after.Skip(after.Length - before.Length).SequenceEqual(before),
                "PowerPoint history does not retain the original boundary.");
            for (var i = 0; i < after.Length - before.Length; i++) app.CommandBars.ExecuteMso("Undo");
            AssertEqual("before SELECT after", shape.TextFrame.TextRange.Text, "Native Undo changed earlier text.");
            AssertTrue(shape.TextFrame.TextRange.Characters(1, 6).Font.Bold == MsoTriState.msoTrue,
                "Native Undo failed to retain leading formatting.");
            AssertTrue(shape.TextFrame.TextRange.Characters(15, 5).Font.Italic == MsoTriState.msoTrue,
                "Native Undo failed to retain trailing formatting.");
            var restored = ReadHistory();
            AssertTrue(restored.SequenceEqual(before), "Native Undo crossed the original boundary.");
            presentation.SaveCopyAs(Path.Combine(artifactRoot, "restored.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            checks.Add(new { name = "native Undo capability", passed = true, restored });

            service = new PowerPointFormulaService(app);
            var flags = BindingFlags.Static | BindingFlags.NonPublic;
            var assembly = typeof(PowerPointFormulaService).Assembly;
            var transaction = assembly.GetType("VisualTeX.PowerPointVsto.PowerPointTextEditTransaction", true)!;
            var package = assembly.GetType("VisualTeX.PowerPointVsto.PowerPointNativeOmmlPackage", true)!;
            string ShapeXml() => (string)package.GetMethod("ReadTextShapeXml", flags)!.Invoke(null,
                new object[] { presentation, slide.SlideID, shape.Id })!;
            string Hash(string xml) => (string)transaction.GetMethod("Fingerprint", flags)!.Invoke(null, new object[] { xml })!;
            // Independent test oracle retains all text, geometry and authored
            // properties; only documented native proofing-cache attributes vary.
            string AuthoredXml(string xml)
            {
                var node = XElement.Parse(xml, LoadOptions.PreserveWhitespace);
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                foreach (var name in new[] { "rPr", "endParaRPr", "defRPr" })
                    foreach (var properties in node.Descendants(a + name))
                        foreach (var attribute in new[] { "dirty", "err", "smtClean" }) properties.Attribute(attribute)?.Remove();
                return node.ToString(SaveOptions.DisableFormatting);
            }
            void Execute(string source, Action mutation) => transaction.GetMethod("Execute", flags)!.Invoke(null,
                new object[] { app, presentation, slide.SlideID, shape.Id, Hash(source), mutation });
            void Check(string name, Action action)
            {
                action();
                checks.Add(new { name, passed = true });
                Console.WriteLine("PASS " + name);
            }
            Check("compatibility branch preserves exact native shape identity", () =>
            {
                const string shapeXml = "<p:sp><p:nvSpPr><p:cNvPr id='2'/></p:nvSpPr><p:txBody>{0}</p:txBody></p:sp>";
                const string wrapper = "<p:sld xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:a14='http://schemas.microsoft.com/office/drawing/2010/main'>{0}</p:sld>";
                var native = string.Format(shapeXml, "native");
                var fallback = string.Format(shapeXml, "fallback");
                var choice = "<mc:AlternateContent><mc:Choice Requires='a14'>" + native + "</mc:Choice><mc:Fallback>" + fallback + "</mc:Fallback></mc:AlternateContent>";
                XElement Select(string content) => (XElement)package.GetMethod("SelectTextShape", flags)!.Invoke(null,
                    new object[] { XDocument.Parse(string.Format(wrapper, content)), 2 })!;
                AssertEqual("native", Select(choice).Value, "Fallback was mistaken for live native text.");
                foreach (var invalid in new[] { native + native, choice.Replace("Requires='a14'", "Requires='p'"),
                    choice.Replace("<mc:Choice Requires='a14'>", "<mc:Choice Requires='p'>" + native + "</mc:Choice><mc:Choice Requires='a14'>") })
                {
                    Exception? rejected = null;
                    try { Select(invalid); } catch (TargetInvocationException error) { rejected = error.InnerException; }
                    AssertTrue(rejected is InvalidOperationException or InvalidDataException, "Ambiguous/unsupported shape was accepted.");
                }
            });
            Check("source proof retains every authored property", () =>
            {
                const string root = "<p:sp xmlns:p='http://schemas.openxmlformats.org/presentationml/2006/main' xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'><p:txBody><a:p><a:r><a:rPr {0}/><a:t> x </a:t></a:r></a:p></p:txBody></p:sp>";
                var baseline = Hash(string.Format(root, ""));
                foreach (var attribute in new[] { "b='1'", "i='1'", "sz='2200'", "spc='25'", "baseline='1200'", "noProof='1'", "lang='zh-CN'" })
                    AssertTrue(baseline != Hash(string.Format(root, attribute)), "Source proof lost authoring attribute " + attribute);
                AssertTrue(baseline != Hash(string.Format(root, "").Replace("> x <", ">x<")), "Source proof lost whitespace.");
                AssertTrue(baseline != Hash(string.Format(root, "").Replace("> x <", "> y <")), "Source proof lost same-length content changes.");
                AssertEqual(baseline, Hash(string.Format(root, "dirty='0' err='0' smtClean='0'")), "Native proofing caches became authoring content.");
            });
            void ExpectRejected(Action action, Type? expected = null)
            {
                Exception? rejected = null;
                try { action(); }
                catch (TargetInvocationException error) { rejected = error.InnerException; }
                catch (InvalidOperationException error) { rejected = error; }
                catch (InvalidDataException error) { rejected = error; }
                AssertTrue(rejected?.GetType() == (expected ?? typeof(InvalidOperationException)), "Expected a pre-mutation rejection, actual=" + rejected);
            }
            OfficeSelection Capture()
            {
                shape.TextFrame.TextRange.Characters(8, 6).Select();
                return service.CaptureTextInsertionTarget();
            }
            OfficeSessionDocument Session(OfficeSelection selected, string display = "inline")
            {
                var value = CreatePowerPointAcceptanceSession("create", "wordOmml", Guid.NewGuid().ToString(),
                    selected.ObjectId, null, displayMode: display);
                value.SourceDocumentId = selected.DocumentId;
                value.ExportResult!.MathMl = "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi><mo>+</mo><mi>y</mi></math>";
                return value;
            }
            Check("deselection cannot reuse text target", () =>
            {
                var addIn = new ThisAddIn();
                var resolve = typeof(ThisAddIn).GetMethod("ResolveOmmlTextInsertionTarget", BindingFlags.Instance | BindingFlags.NonPublic)!;
                Capture();
                AssertTrue(resolve.Invoke(addIn, new object[] { service }) is OfficeSelection, "Live text target unavailable.");
                app.ActiveWindow.Selection.Unselect();
                AssertTrue(resolve.Invoke(addIn, new object[] { service }) is null, "Deselected old text target reused.");
            });
            Check("legacy or documentless text targets cannot fall back to standalone insertion", () =>
            {
                var selected = Capture();
                var source = ShapeXml();
                var count = slide.Shapes.Count;
                var missingDocument = Session(selected);
                missingDocument.SourceDocumentId = null;
                ExpectRejected(() => service.InsertOmml(missingDocument));
                var legacy = Session(selected);
                legacy.SourceObjectId = selected.ObjectId!.Substring(0, selected.ObjectId.LastIndexOf(':'));
                ExpectRejected(() => service.InsertOmml(legacy), typeof(InvalidDataException));
                AssertEqual(count, slide.Shapes.Count, "An invalid captured target created another object.");
                AssertEqual(AuthoredXml(source), AuthoredXml(ShapeXml()), "An invalid captured target altered the text.");
            });
            Check("passive selection is read-only and copied formulas keep exact target identity", () =>
            {
                var svgPath = Path.Combine(artifactRoot, "identity-fixture.svg");
                File.WriteAllText(svgPath, CreateSvg(220, 72));
                var initial = CreatePowerPointAcceptanceSession("create", "crossPlatformPicture",
                    Guid.NewGuid().ToString(), service.CaptureSlideInsertionTarget().ObjectId, null);
                var inserted = service.Insert(initial, svgPath);
                PowerPoint.Shape? original = null, copy = null, deletedCopy = null;
                PowerPoint.ShapeRange? copies = null;
                string Snapshot(PowerPoint.Shape item)
                {
                    var tags = new List<string>();
                    for (var i = 1; i <= item.Tags.Count; i++) tags.Add(item.Tags.Name(i) + "=" + item.Tags.Value(i));
                    return JsonSerializer.Serialize(new { item.Name, item.AlternativeText, item.LockAspectRatio, tags });
                }
                try
                {
                    original = slide.Shapes[inserted.ObjectId];
                    var originalState = Snapshot(original);
                    copies = original.Duplicate(); copy = copies[1]; Release(copies); copies = null;
                    copy.Name = original.Name;
                    copy.LockAspectRatio = MsoTriState.msoFalse;
                    copy.Select(MsoTriState.msoTrue);
                    app.StartNewUndoEntry();
                    var history = ReadHistory();
                    var state = Snapshot(copy);
                    presentation.Saved = MsoTriState.msoTrue;
                    var selectedCopy = service.ReadSelection();
                    service.ReadSelection(); service.ReadSelection();
                    AssertEqual(state, Snapshot(copy), "Passive selection changed copied shape metadata/formatting.");
                    AssertTrue(presentation.Saved == MsoTriState.msoTrue, "Passive selection dirtied the presentation.");
                    AssertTrue(history.SequenceEqual(ReadHistory()), "Passive selection changed Undo history.");
                    var prepared = service.PrepareManagedSelectionForEdit(selectedCopy);
                    AssertTrue(prepared.FormulaId != selectedCopy.FormulaId, "Explicit edit did not rekey copied identity.");
                    AssertEqual(originalState, Snapshot(original), "Copy repair touched the original formula.");
                    AssertEqual(prepared.FormulaId, service.ReadSelection().FormulaId, "Repaired identity did not persist.");

                    copies = original.Duplicate(); deletedCopy = copies[1]; Release(copies); copies = null;
                    deletedCopy.Name = original.Name;
                    deletedCopy.Select(MsoTriState.msoTrue);
                    var captured = service.ReadSelection();
                    deletedCopy.Delete(); Release(deletedCopy); deletedCopy = null;
                    var edit = CreatePowerPointAcceptanceSession("edit", "crossPlatformPicture", captured.FormulaId!,
                        captured.ObjectId, captured.Metadata);
                    edit.SourceDocumentId = captured.DocumentId;
                    ExpectRejected(() => service.Replace(edit, svgPath));
                    AssertEqual(originalState, Snapshot(original), "Missing copied target fell back to another formula.");
                }
                finally
                {
                    copy?.Delete(); original?.Delete();
                    Release(copies); Release(deletedCopy); Release(copy); Release(original);
                }
            });
            foreach (var stage in new[] { "before-write", "after-text", "after-format" })
                Check("exact native recovery " + stage, () =>
                {
                    var source = ShapeXml();
                    File.WriteAllText(Path.Combine(artifactRoot, stage + "-source.xml"), source);
                    File.WriteAllText(Path.Combine(artifactRoot, stage + "-source-second-read.xml"), ShapeXml());
                    var undo = ReadHistory();
                    const string injected = "injected-text-transaction-failure";
                    Exception? observed = null;
                    try
                    {
                        Execute(source, () =>
                        {
                            if (stage != "before-write") shape.TextFrame.TextRange.Characters(8, 6).Text = "changed";
                            if (stage == "after-format") shape.TextFrame.TextRange.Characters(8, 7).Font.Size = 37;
                            throw new InvalidOperationException(injected);
                        });
                    }
                    catch (TargetInvocationException error) { observed = error.InnerException; }
                    AssertTrue(observed is InvalidOperationException && observed.Message == injected,
                        "Recovery must preserve the primary failure, not conceal a rollback failure: " + observed);
                    var recovered = ShapeXml();
                    File.WriteAllText(Path.Combine(artifactRoot, stage + "-restored.xml"), recovered);
                    AssertEqual(AuthoredXml(source), AuthoredXml(recovered), "Native recovery lost authored shape XML.");
                    AssertTrue(undo.SequenceEqual(ReadHistory()), "Recovery changed the earlier user Undo boundary.");
                });
            Check("same-length source edit rejects write", () =>
            {
                var selected = Capture();
                shape.TextFrame.TextRange.Characters(8, 6).Text = "ALTERD";
                var edited = ShapeXml();
                ExpectRejected(() => service.InsertOmml(Session(selected)));
                AssertEqual(AuthoredXml(edited), AuthoredXml(ShapeXml()), "Stale text save altered new user content.");
                shape.TextFrame.TextRange.Characters(8, 6).Text = "SELECT";
            });
            Check("format-only source edit rejects write", () =>
            {
                var selected = Capture();
                shape.TextFrame.TextRange.Characters(8, 6).Font.Color.RGB = 255;
                var edited = ShapeXml();
                ExpectRejected(() => service.InsertOmml(Session(selected)));
                AssertEqual(AuthoredXml(edited), AuthoredXml(ShapeXml()), "Stale text save altered new formatting.");
            });
            Check("invalid MathML display preparation leaves source intact", () =>
            {
                var selected = Capture();
                var source = ShapeXml();
                var invalid = Session(selected, "block");
                invalid.ExportResult!.MathMl = "<math><mfrac>";
                Exception? observed = null;
                try { service.InsertOmml(invalid); } catch (Exception error) { observed = error; }
                AssertTrue(observed is not null, "Invalid MathML was accepted.");
                AssertEqual(AuthoredXml(source), AuthoredXml(ShapeXml()), "Failed display preparation deleted selected prose.");
            });
            foreach (var display in new[] { "inline", "block" })
                Check("exact recovery after native conversion " + display, () =>
                {
                    var kernel = assembly.GetType("VisualTeX.PowerPointVsto.PowerPointNativeOmmlKernel", true)!;
                    var prepare = kernel.GetMethod("PrepareTextInsertion", flags)!;
                    var source = ShapeXml();
                    var undo = ReadHistory();
                    var write = (Action)prepare.Invoke(null, new object[] { app, shape, 8, 6,
                        "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi><mo>+</mo><mi>y</mi></math>", display, null })!;
                    Exception? observed = null;
                    try { Execute(source, () => { write(); throw new InvalidOperationException("after-native-write"); }); }
                    catch (TargetInvocationException error) { observed = error.InnerException; }
                    AssertTrue(observed is InvalidOperationException && observed.Message == "after-native-write", "Native recovery failed: " + observed);
                    AssertEqual(AuthoredXml(source), AuthoredXml(ShapeXml()), "Native conversion rollback lost authored shape XML.");
                    AssertTrue(undo.SequenceEqual(ReadHistory()), "Native conversion recovery crossed the earlier history.");
                });
            var successFailures = new List<Exception>();
            foreach (var display in new[] { "inline", "block" })
                {
                    var selected = Capture();
                    var source = ShapeXml();
                    var originalHistory = ReadHistory();
                    var sourceId = shape.Id;
                    var count = slide.Shapes.Count;
                    service.InsertOmml(Session(selected, display));
                    AssertEqual(sourceId, shape.Id, "Text insertion replaced the original shape.");
                    AssertEqual(count, slide.Shapes.Count, "Text insertion changed shape inventory.");
                    var target = ShapeXml();
                    File.WriteAllText(Path.Combine(artifactRoot, display + "-source.xml"), source);
                    File.WriteAllText(Path.Combine(artifactRoot, display + "-target.xml"), target);
                    AssertTrue(target.Contains("oMath"), "No native equation was persisted.");
                    AssertTrue(shape.TextFrame.TextRange.Text.StartsWith("before") && shape.TextFrame.TextRange.Text.EndsWith("after"),
                        "Text insertion damaged surrounding prose.");
                    presentation.SaveCopyAs(Path.Combine(artifactRoot, display + ".pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                    Check("native text insertion " + display, () => AssertTrue(target.Contains("oMath"), "No native equation."));
                    try
                    {
                        app.CommandBars.ExecuteMso("Undo");
                        var undone = ShapeXml();
                        File.WriteAllText(Path.Combine(artifactRoot, display + "-one-undo.xml"), undone);
                        AssertEqual(AuthoredXml(source), AuthoredXml(undone), "One Undo did not restore original source/formatting.");
                        app.CommandBars.ExecuteMso("Redo");
                        AssertEqual(AuthoredXml(target), AuthoredXml(ShapeXml()), "One Redo changed inserted native math.");
                        checks.Add(new { name = "single Undo Redo " + display, passed = true });
                    }
                    catch (Exception error)
                    {
                        successFailures.Add(error);
                        checks.Add(new { name = "single Undo Redo " + display, passed = false, error = error.ToString() });
                        Console.WriteLine("FAIL single Undo Redo " + display);
                    }
                    // Restore only the test-owned remaining prefix so the other
                    // display mode can be examined; this is NOT a one-Undo pass.
                    var history = ReadHistory();
                    AssertTrue(history.Length >= originalHistory.Length && history.Skip(history.Length - originalHistory.Length).SequenceEqual(originalHistory),
                        "Cannot prove fixture-owned remaining history.");
                    for (var i = 0; i < history.Length - originalHistory.Length; i++) app.CommandBars.ExecuteMso("Undo");
                    AssertEqual(AuthoredXml(source), AuthoredXml(ShapeXml()), "Fixture restoration lost original source.");
                }
            if (successFailures.Count > 0) throw new AggregateException("Native insertion succeeded, but single Undo/Redo acceptance remains failed.", successFailures);
        }
        catch (Exception error) { failure = error; throw; }
        finally
        {
            try { service?.Dispose(); } catch (Exception error) { checks.Add(new { name = "private Word cleanup", error = error.ToString() }); }
            Release(shape); Release(slide);
            if (presentation is not null)
            {
                try { presentation.Saved = MsoTriState.msoTrue; presentation.Close(); }
                catch (Exception error) { checks.Add(new { name = "owned document cleanup", error = error.ToString() }); }
                Release(presentation);
            }
            if (app is not null && ownedPid != 0)
            {
                try
                {
                    using var current = Process.GetProcessById(ownedPid);
                    if (current.StartTime.ToUniversalTime() == ownedStart && app.Presentations.Count == 0) app.Quit();
                }
                catch (ArgumentException) { }
            }
            Release(app);
            File.WriteAllText(Path.Combine(artifactRoot, "text-safety.json"), JsonSerializer.Serialize(
                new { startedUtc, ownedPid, ownedStart, scope = "bounded source freshness/native rollback/text insertion; no physical Ribbon input", registryWrites = false,
                    passed = failure is null, error = failure?.ToString(), checks }, new JsonSerializerOptions { WriteIndented = true }));
        }
    }
}

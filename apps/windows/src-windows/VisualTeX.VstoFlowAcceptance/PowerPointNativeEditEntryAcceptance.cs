using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunPowerPointNativeEditEntryAcceptance(string root)
    {
        if (Process.GetProcessesByName("POWERPNT").Length != 0
            || Process.GetProcessesByName("WINWORD").Length != 0)
            throw new InvalidOperationException("Entry probe needs no existing PowerPoint or Word; user instances are not touched.");
        var started = DateTime.UtcNow;
        var app = new PowerPoint.Application();
        GetWindowThreadProcessId(new IntPtr(app.HWND), out var pid);
        var process = Process.GetProcessById((int)pid);
        if (process.StartTime.ToUniversalTime() < started.AddSeconds(-2) || app.Presentations.Count != 0)
            throw new InvalidOperationException("PPT ownership not established.");
        var ownedStart = process.StartTime.ToUniversalTime();
        PowerPoint.Presentation? presentation = null;
        PowerPoint.Slide? slide = null;
        PowerPoint.Shape? shape = null;
        var flags = BindingFlags.NonPublic | BindingFlags.Instance;
        var type = typeof(VisualTeX.PowerPointVsto.ThisAddIn);
        var addIn = new VisualTeX.PowerPointVsto.ThisAddIn();
        using var client = new VisualTeXSessionClient();
        using var lifetime = new CancellationTokenSource();
        var dispatcher = (IDisposable)Activator.CreateInstance(type.Assembly.GetType("VisualTeX.PowerPointVsto.OfficeUiDispatcher", true)!, true)!;
        using var service = new PowerPointFormulaService(app);
        string? sessionId = null;
        IDisposable? mouseHook = null;
        var rows = new List<object>();
        void Set(string name, object value) => type.GetField(name, flags)!.SetValue(addIn, value);
        void Pump(int milliseconds)
        {
            var until = Stopwatch.StartNew();
            while (until.ElapsedMilliseconds < milliseconds) { WinForms.Application.DoEvents(); Thread.Sleep(10); }
        }
        try
        {
            app.Visible = MsoTriState.msoTrue;
            presentation = app.Presentations.Add(MsoTriState.msoTrue);
            slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            app.ActiveWindow.View.GotoSlide(1);
            var source = CreatePowerPointAcceptanceSession("create", "wordOmml", Guid.NewGuid().ToString(), service.CaptureSlideInsertionTarget().ObjectId, null);
            var mixed = Environment.GetEnvironmentVariable("VISUALTEX_PPT_NATIVE_ENTRY_FIXTURE") == "mixed";
            if (mixed)
            {
                shape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 100, 100, 600, 120);
                shape.TextFrame.TextRange.Text = "before PLACEHOLDER after";
                shape.TextFrame.TextRange.Font.Size = 24;
                shape.TextFrame.TextRange.Characters(1, 6).Font.Color.RGB = 0x332211;
                shape.TextFrame.TextRange.Characters(8, 11).Select();
                var target = service.CaptureTextInsertionTarget();
                source.SourceDocumentId = target.DocumentId;
                source.SourceObjectId = target.ObjectId;
                source.DisplayMode = "inline";
                Release(shape); shape = null;
            }
            var result = service.InsertOmml(source);
            shape = slide.Shapes[result.ObjectId];
            shape.Select(MsoTriState.msoTrue);
            AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, source.FormulaId!);
            Set("_application", app); Set("_dispatcher", dispatcher); Set("_formulaService", service);
            Set("_sessionClient", client); Set("_lifetime", lifetime);
            var selected = service.ReadSelection();
            rows.Add(new { stage="selection", selected.ObjectId, selected.ObjectMode, hasMetadata=selected.Metadata is not null });
            var timer = Stopwatch.StartNew();
            addIn.OnEditSelected(new object());
            while (timer.Elapsed < TimeSpan.FromSeconds(25))
            {
                Pump(40);
                sessionId = (string?)type.GetField("_activeSessionId", flags)!.GetValue(addIn);
                if (!string.IsNullOrEmpty(sessionId) || !string.IsNullOrEmpty(addIn.DiagnosticLastError)) break;
            }
            rows.Add(new { stage="Ribbon callback", sessionId, error=addIn.DiagnosticLastError, elapsedMs=timer.ElapsedMilliseconds });
            if (string.IsNullOrEmpty(sessionId)) throw new InvalidOperationException("Native Edit did not create a session: " + addIn.DiagnosticLastError);
            Pump(1000);
            var session = client.GetSessionAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            rows.Add(new { stage="editor-session", session.Status, session.ObjectMode, session.SourceObjectId, session.OriginalMetadata, session.Lines });
            AssertEqual("wordOmml", session.ObjectMode, "Native edit target changed format");
            AssertContains(session.Lines[0].Latex, "frac", "Native edit source missing");
            var originalShapeId = shape.Id;
            var originalCount = slide.Shapes.Count;
            Commit(client, session, mixed ? "inline" : "block", "wordOmml", "z^3+9",
                mathMl: "<math xmlns='http://www.w3.org/1998/Math/MathML'><msup><mi>z</mi><mn>3</mn></msup><mo>+</mo><mn>9</mn></math>");
            timer.Restart();
            while (timer.Elapsed < TimeSpan.FromSeconds(25)
                && !string.IsNullOrEmpty((string?)type.GetField("_activeSessionId", flags)!.GetValue(addIn))) Pump(40);
            AssertEqual("", addIn.DiagnosticLastError, "Native editor commit failed");
            AssertEqual(originalCount, slide.Shapes.Count, "Native edit added a shape");
            AssertEqual(originalShapeId, shape.Id, "Native edit replaced the text box");
            if (mixed)
            {
                AssertTrue(shape.TextFrame.TextRange.Text.StartsWith("before ") && shape.TextFrame.TextRange.Text.EndsWith(" after"), "Neighboring prose changed");
                AssertEqual(0x332211, shape.TextFrame.TextRange.Characters(1,6).Font.Color.RGB, "Prose color changed");
            }
            shape.Select(MsoTriState.msoTrue);
            var updated=service.LoadNativeOmmlSelection(service.ReadSelection());
            AssertContains(updated.Metadata!.Latex,"9","Native edit content did not change");
            rows.Add(new {stage="commit",shapeId=shape.Id,shapeCount=slide.Shapes.Count,updated.Metadata.Latex,mixed});
            client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            sessionId=null;
            object zone=((dynamic)shape.TextFrame2.TextRange).MathZones(1,1);
            var x=app.ActiveWindow.PointsToScreenPixelsX(Convert.ToSingle(((dynamic)zone).BoundLeft)+Convert.ToSingle(((dynamic)zone).BoundWidth)/2);
            var y=app.ActiveWindow.PointsToScreenPixelsY(Convert.ToSingle(((dynamic)zone).BoundTop)+Convert.ToSingle(((dynamic)zone).BoundHeight)/2);
            Release(zone);
            var hit=service.ReadFormulaAtScreenPoint(x,y);
            AssertTrue(hit is not null && PowerPointFormulaService.IsNativeOmmlReference(hit.ObjectId), "Native double-click hit missed");
            var physical = Environment.GetEnvironmentVariable("VISUALTEX_PPT_NATIVE_PHYSICAL_DOUBLE_CLICK") == "1";
            if (physical)
            {
                var hookType=type.Assembly.GetType("VisualTeX.PowerPointVsto.PowerPointDoubleClickHook",true)!;
                mouseHook=(IDisposable)Activator.CreateInstance(hookType, new object[]{
                    new Action<int,int>((px,py)=>type.GetMethod("OnNativeDoubleClick",flags)!.Invoke(addIn,new object[]{px,py}))})!;
                hookType.GetMethod("Start")!.Invoke(mouseHook,null);
                SetForegroundWindow(new IntPtr(app.HWND)); Pump(250);
                for(var click=0;click<2;click++)
                {
                    GetWindowThreadProcessId(GetForegroundWindowForAcceptance(),out var foregroundPid);
                    AssertEqual(pid,foregroundPid,"Only the owned PowerPoint may receive clicks");
                    AssertTrue(SetCursorPos(x,y),"Could not position cursor over owned formula");
                    mouse_event(MouseLeftDown,0,0,0,UIntPtr.Zero);
                    mouse_event(MouseLeftUp,0,0,0,UIntPtr.Zero);
                    Pump(80);
                }
            }
            else type.GetMethod("OnNativeDoubleClick",flags)!.Invoke(addIn,new object[]{x,y});
            timer.Restart();
            while(timer.Elapsed<TimeSpan.FromSeconds(25))
            {
                Pump(40); sessionId=(string?)type.GetField("_activeSessionId",flags)!.GetValue(addIn);
                if(!string.IsNullOrEmpty(sessionId)||!string.IsNullOrEmpty(addIn.DiagnosticLastError)) break;
            }
            AssertTrue(!string.IsNullOrEmpty(sessionId), "Native double-click routing did not open editor: "+addIn.DiagnosticLastError);
            rows.Add(new {stage=physical?"physical double-click":"double-click routing",sessionId,x,y});
            client.PatchAsync(sessionId!,new {status="cancelled",explicitCancel=true},CancellationToken.None).GetAwaiter().GetResult();
            Pump(300);
            AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape, source.FormulaId!);
            presentation.SaveAs(Path.Combine(root, "native-edit-entry.pptx"), PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoTrue);
            var path=Path.Combine(root,"native-edit-entry.pptx");
            Release(shape); shape=null; Release(slide); slide=null;
            presentation.Close(); Release(presentation); presentation=null;
            presentation=app.Presentations.Open(path,WithWindow:MsoTriState.msoTrue);
            slide=presentation.Slides[1]; shape=slide.Shapes[1];
            shape.Select(MsoTriState.msoTrue);
            var reopened=service.LoadNativeOmmlSelection(service.ReadSelection());
            AssertContains(reopened.Metadata!.Latex,"9","Reopened native equation differs");
            AssertPowerPointNativeOmmlHasNoVisualTeXIdentity(shape,reopened.FormulaId!);
            rows.Add(new {stage="reopened",reopened.Metadata.Latex});
        }
        finally
        {
            mouseHook?.Dispose();
            lifetime.Cancel(); Pump(100);
            if (sessionId is not null) client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
            dispatcher.Dispose();
            Release(shape); Release(slide);
            if (presentation is not null) { presentation.Saved=MsoTriState.msoTrue; presentation.Close(); Release(presentation); }
            if (!process.HasExited && process.StartTime.ToUniversalTime()==ownedStart && app.Presentations.Count==0) app.Quit();
            Release(app);
            File.WriteAllText(Path.Combine(root,"native-entry.json"),JsonSerializer.Serialize(new {started,pid,ownedStart,rows},new JsonSerializerOptions{WriteIndented=true}));
        }
    }
}

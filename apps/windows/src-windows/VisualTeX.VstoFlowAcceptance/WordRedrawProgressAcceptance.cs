using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.Contracts;
using WinForms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    [DllImport("user32.dll", EntryPoint = "GetWindow")]
    private static extern IntPtr RedrawProgressWindowOwner(IntPtr window, uint command);

    private static void RunOwnedRedrawProgressAcceptance(string root)
    {
        var failures = new List<string>();
        var observations = new List<object>();
        var input = Environment.GetEnvironmentVariable("VISUALTEX_REDRAW_PROGRESS_DOCUMENT");
        var confirm = Environment.GetEnvironmentVariable("VISUALTEX_REDRAW_PROGRESS_CONFIRM") == "1";
        var expectedCount = string.IsNullOrWhiteSpace(input) ? 20 : 200;
        string? ownedInput = null;
        if (!string.IsNullOrWhiteSpace(input))
        {
            ownedInput = Path.Combine(root, "owned-source-copy.docx");
            File.Copy(input, ownedInput);
        }
        using var host = new WordPerformanceHost(ownedInput);
        host.Application.Visible = true;
        if (ownedInput is null)
        {
            var content = host.Document.Content;
            var source = string.Join("\r", Enumerable.Range(1, 20).Select(i => $"Example {i}: $x_{{{i}}}+1$")) + "\r";
            try { content.Text = source; content.Font.Size = 12; }
            finally { Release(content); }
        }
        var initialContent = host.Document.Content;
        string capturedSource;
        try { capturedSource = initialContent.Text; }
        finally { Release(initialContent); }
        File.WriteAllText(Path.Combine(root, "source-before.txt"), capturedSource);
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", Path.Combine(root, "progress-trace.log"));
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", Path.Combine(root, "redraw.log"));
        var window = host.Document.ActiveWindow;
        IntPtr wordHandle;
        try { wordHandle = new IntPtr(window.Hwnd); }
        finally { Release(window); }

        var service = new WordFormulaService(host.Application);
        var watch = Stopwatch.StartNew();
        var stages = new List<(string Stage, int Completed, int Total, long Milliseconds)>();
        service.OperationProgress = (stage, completed, total) => stages.Add((stage, completed, total, watch.ElapsedMilliseconds));
        var plan = service.CaptureLatexRedrawPlan(true);
        service.OperationProgress = null;
        if (plan.Targets.Count != expectedCount) failures.Add("Captured formula count differs from independent seed.");
        if (!stages.Any(s => s.Stage == "正在读取公式字号…" && s.Completed == expectedCount && s.Total == expectedCount))
            failures.Add("No actual font preparation completion was reported.");
        observations.Add(new { phase = "capture", count = plan.Targets.Count, timings = plan.CaptureTimings,
            stages = stages.Select(s => new { s.Stage, s.Completed, s.Total, s.Milliseconds }).ToArray() });

        var oldAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        // Only this runner's branch changes. The already-created owned Word
        // still has its inherited acceptance environment suppressing installation.
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
        var sawDialog = false;
        Stopwatch? dialogWatch = null;
        using var timer = new WinForms.Timer { Interval = 40 };
        timer.Tick += (_, _) =>
        {
            var dialog = WinForms.Application.OpenForms.Cast<WinForms.Form>()
                .FirstOrDefault(f => f.GetType().FullName == "VisualTeX.WordVsto.LatexRedrawDialog");
            if (dialog is null) return;
            dialogWatch ??= Stopwatch.StartNew();
            if (dialogWatch.ElapsedMilliseconds < 800) return;
            timer.Stop();
            sawDialog = true;
            var visibleProgress = new List<object>();
            var labels = new List<string>();
            var owner = RedrawProgressWindowOwner(dialog.Handle, 4);
            try
            {
                OwnedWordNativeWindows.EnumWindows((candidate, _) =>
                {
                    GetWindowThreadProcessId(candidate, out var pid);
                    if (pid != (uint)Process.GetCurrentProcess().Id || candidate == dialog.Handle || !IsWindowVisible(candidate)) return true;
                    var text = new StringBuilder(512); GetWindowText(candidate, text, text.Capacity);
                    if (text.ToString() != dialog.Text) return true;
                    OwnedWordNativeWindows.EnumChildWindows(candidate, (child, unused) =>
                    {
                        var label = new StringBuilder(512); GetWindowText(child, label, label.Capacity);
                        if (label.Length > 0) labels.Add(label.ToString());
                        return true;
                    }, IntPtr.Zero);
                    visibleProgress.Add(new { handle = candidate.ToInt64(), title = text.ToString() });
                    return true;
                }, IntPtr.Zero);
                if (visibleProgress.Count != 0) failures.Add("Progress remains visible while redraw waits for confirmation.");
                if (owner != wordHandle) failures.Add("Redraw confirmation has no explicit Word owner.");
                observations.Add(new { phase = "normal-confirmation", dialog = dialog.Handle.ToInt64(),
                    owner = owner.ToInt64(), word = wordHandle.ToInt64(), visibleProgress, labels,
                    waitedMilliseconds = dialogWatch.ElapsedMilliseconds });
            }
            finally { dialog.DialogResult = confirm ? WinForms.DialogResult.OK : WinForms.DialogResult.Cancel; }
        };
        try
        {
            timer.Start();
            var task = (Task)typeof(ThisAddIn).GetMethod("RedrawLatexAsync", BindingFlags.NonPublic | BindingFlags.Instance)!
                .Invoke(host.AddIn, new object[] { true, FormulaOleContract.MathTypeOleMode })!;
            var deadline = Stopwatch.StartNew();
            while (!task.IsCompleted && deadline.Elapsed < TimeSpan.FromMinutes(4))
            { WinForms.Application.DoEvents(); Thread.Sleep(20); }
            if (!task.IsCompleted) throw new TimeoutException("Owned redraw did not finish after confirmation cancellation.");
            task.GetAwaiter().GetResult();
            if (!sawDialog) failures.Add("Normal redraw confirmation was not observed.");
            var after = host.Document.Content;
            try
            {
                File.WriteAllText(Path.Combine(root, "source-after.txt"), after.Text);
                if (!confirm && after.Text != capturedSource) failures.Add("Cancelled redraw changed source content.");
            }
            finally { Release(after); }
            if (!confirm && (host.Document.InlineShapes.Count != 0 || host.Document.OMaths.Count != 0))
                failures.Add("Cancelled redraw created formula hosts.");
            if (confirm)
            {
                if (ownedInput is not null) throw new InvalidOperationException("Positive confirmation uses only the independent20 corpus.");
                var converted = service.CaptureFormulaFormatConversionPlan(true, FormulaOleContract.MathTypeOleMode, FormulaOleContract.NativeOleMode);
                if (converted.Targets.Count != 20) failures.Add("Confirmed redraw did not create20 MathType objects.");
                for (var index = 0; index < converted.Targets.Count; index++)
                {
                    var expected = $"<math><msub><mi>x</mi><mn>{index+1}</mn></msub><mo>+</mo><mn>1</mn></math>";
                    if (MathTypeMtefCodec.SemanticSignature(converted.Targets[index].SourceMathMl!) != MathTypeMtefCodec.SemanticSignature(expected))
                        failures.Add("Confirmed redraw changed formula " + (index+1));
                }
                var trace = File.ReadAllText(Path.Combine(root, "progress-trace.log"));
                if (!trace.Contains("stage=正在原位写入重绘公式… completed=20 total=20"))
                    failures.Add("Confirmed redraw never reported write completion20/20.");
                if (!trace.Contains("stage=正在准备重绘… completed=0 total=0 "))
                    failures.Add("Confirmed redraw did not start a new execution progress scope.");
                host.Save(Path.Combine(root, "redraw-result.docx"));
                observations.Add(new { phase = "confirmed-redraw", formulas = converted.Targets.Count, elapsedMs = deadline.ElapsedMilliseconds });
            }
            var gate = (SemaphoreSlim)typeof(ThisAddIn).GetField("_operationGate", BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(host.AddIn)!;
            if (gate.CurrentCount != 1) failures.Add("Cancellation did not release the operation gate.");
        }
        finally
        {
            timer.Stop();
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", oldAcceptance);
            File.WriteAllText(Path.Combine(root, "redraw-progress.json"), JsonSerializer.Serialize(new {
                passed = failures.Count == 0, confirm, input, observations, failures, dll = typeof(ThisAddIn).Assembly.Location,
                mvid = typeof(ThisAddIn).Assembly.ManifestModule.ModuleVersionId }, new JsonSerializerOptions { WriteIndented = true }));
        }
        if (failures.Count != 0) throw new InvalidDataException(string.Join("; ", failures));
        Console.WriteLine($"[REDRAW PROGRESS] Capture{expectedCount}, real confirmation (accept={confirm}) and content/gate checks passed.");
    }
}

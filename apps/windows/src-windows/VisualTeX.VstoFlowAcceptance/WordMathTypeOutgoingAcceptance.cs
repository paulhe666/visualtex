using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunOwnedMathTypeOutgoing(string root)
    {
        var input = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_DOCUMENT")
            ?? throw new InvalidOperationException("Explicit owned source copy required.");
        var targetMode = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_TARGET") ?? FormulaOleContract.WordOmmlMode;
        var output = Path.Combine(root, "converted.docx");
        File.Copy(input, output);
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", Path.Combine(root, "trace.log"));
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", Path.Combine(root, "redraw.log"));
        using var host = new WordPerformanceHost(output);
        var normalUi = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_NORMAL_UI") == "1";
        if (normalUi) host.Application.Visible = true;
        var service = new WordFormulaService(host.Application);
        var wholeDocument = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_LOCAL") != "1";
        if (!wholeDocument)
        {
            var shapes = host.Document.InlineShapes;
            var shape = shapes[1];
            var range = shape.Range;
            var selection = host.Application.Selection;
            try { selection.SetRange(range.Start, range.Start); }
            finally { Release(selection); Release(range); Release(shape); Release(shapes); }
        }
        void Run(string method, params object[] args)
        {
            var watch = Stopwatch.StartNew();
            var savedAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
            using var confirmations = new System.Windows.Forms.Timer { Interval = 100 };
            var seen = new HashSet<IntPtr>();
            confirmations.Tick += (_, _) =>
            {
                foreach (System.Windows.Forms.Form form in System.Windows.Forms.Application.OpenForms)
                {
                    if (form.GetType().Name == "LatexRedrawDialog" && seen.Add(form.Handle))
                    {
                        Console.WriteLine("[OUTGOING] Normal redraw confirmation accepted with default options.");
                        form.DialogResult = System.Windows.Forms.DialogResult.OK;
                        return;
                    }
                }
                EnumWindows((window, _) =>
                {
                    GetWindowThreadProcessId(window, out var pid);
                    if (pid != (uint)Process.GetCurrentProcess().Id || !IsWindowVisible(window) || seen.Contains(window)) return true;
                    var className = new System.Text.StringBuilder(64);
                    OwnedWordNativeWindows.GetClassName(window, className, className.Capacity);
                    if (className.ToString() != "#32770") return true;
                    var title = new System.Text.StringBuilder(256); GetWindowText(window, title, title.Capacity);
                    if (title.ToString() != "VisualTeX 公式格式转换") return true;
                    var yes = OutgoingDialogItem(window, 6);
                    var ok = OutgoingDialogItem(window, 1);
                    if (yes == IntPtr.Zero && ok == IntPtr.Zero) return true; // status window, not a dialog
                    seen.Add(window);
                    Console.WriteLine("[OUTGOING] Normal conversion dialog action=" + (yes != IntPtr.Zero ? "Yes" : "OK"));
                    OutgoingPostMessage(window, 0x0111, new IntPtr(yes != IntPtr.Zero ? 6 : 1), IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
            };
            if (normalUi) { Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null); confirmations.Start(); }
            try
            {
            var task = (Task)typeof(ThisAddIn).GetMethod(method, BindingFlags.NonPublic | BindingFlags.Instance)!
                .Invoke(host.AddIn, args)!;
            while (!task.IsCompleted && watch.Elapsed < TimeSpan.FromMinutes(15))
            { System.Windows.Forms.Application.DoEvents(); Thread.Sleep(20); }
            if (!task.IsCompleted) throw new TimeoutException(method);
            task.GetAwaiter().GetResult();
            Console.WriteLine($"[OUTGOING] {method}: {watch.Elapsed.TotalSeconds:F2}s");
            }
            finally { confirmations.Stop(); Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", savedAcceptance); }
        }
        if (Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_REDRAW") == "1")
        {
            Run("RedrawLatexAsync", true, FormulaOleContract.MathTypeOleMode);
        }
        // The user converts immediately after redraw. Saving or inspecting the
        // sources first can materialize Word's OLE/field state and mask the bug.
        if (Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_REDRAW") == "1"
            || Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_DIRECT") == "1")
        {
            Run("ConvertFormulaFormatAsync", true, FormulaOleContract.MathTypeOleMode, targetMode);
            host.Document.Save();
            AssertEqual(0, CountMathTypeOleShapes(host.Document), "Immediate redraw -> conversion failed; see trace.log.");
            return;
        }
        var originalCount = CountMathTypeOleShapes(host.Document);
        var plan = service.CaptureFormulaFormatConversionPlan(wholeDocument, FormulaOleContract.MathTypeOleMode, targetMode);
        AssertTrue(plan.Targets.Count > 0, "No MathType sources.");
        var evidence = new List<object>();
        foreach (var target in plan.Targets)
        {
            var range = WordFormulaOperationLocator.ResolveCapturedRange(host.Document, target.SourceObjectId);
            try
            {
                using var binding = WordMathTypeHostAdapter.BindCapturedHost(host.Application, host.Document, range);
                var source = binding.Host;
                var font = FormulaFontSize.ResolveSemanticFontSize(source.Metadata!);
                var visibleFont = source.Display ? font : (float?)typeof(WordFormulaService)
                    .GetMethod("ReadMathTypeInlinePresentationFontSize", BindingFlags.NonPublic | BindingFlags.Static)!
                    .Invoke(null, new object[] { binding.MathTypeShape! }) ?? font;
                evidence.Add(new { target.SourceStart, target.Latex,
                    mathMlEqual = binding.MathMl == target.SourceMathMl,
                    displayBefore = target.DisplayMode, displayAfter = source.DisplayMode,
                    numberedBefore = target.Numbered, numberedAfter = source.Numbering.Numbered,
                    semanticFontBefore = FormulaFontSize.ResolveSemanticFontSize(target.Metadata), semanticFontAfter = font,
                    presentationFontBefore = target.FontSizePt, presentationFontAfter = visibleFont });
            }
            finally { Release(range); }
        }
        File.WriteAllText(Path.Combine(root, "source-audit.json"), JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
        Run("ConvertFormulaFormatAsync", wholeDocument, FormulaOleContract.MathTypeOleMode, targetMode);
        var remaining = CountMathTypeOleShapes(host.Document);
        var equations = host.Document.OMaths;
        int converted;
        try { converted = targetMode == FormulaOleContract.WordOmmlMode ? equations.Count : CountVisualTeXNativeOleShapes(host.Document); }
        finally { Release(equations); }
        File.WriteAllText(Path.Combine(root, "result.json"), JsonSerializer.Serialize(new { input, targetMode,
            expected = plan.Targets.Count, remaining, converted, dll = typeof(ThisAddIn).Assembly.Location,
            mvid = typeof(ThisAddIn).Assembly.ManifestModule.ModuleVersionId }, new JsonSerializerOptions { WriteIndented = true }));
        host.Document.Save();
        AssertEqual(originalCount - plan.Targets.Count, remaining, "MathType outgoing left source objects; see trace.log and source-audit.json.");
        AssertEqual(plan.Targets.Count, converted, "MathType outgoing target count.");
    }

    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "GetDlgItem")]
    private static extern IntPtr OutgoingDialogItem(IntPtr window, int id);
    [System.Runtime.InteropServices.DllImport("user32.dll", EntryPoint = "PostMessageW")]
    private static extern bool OutgoingPostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
}

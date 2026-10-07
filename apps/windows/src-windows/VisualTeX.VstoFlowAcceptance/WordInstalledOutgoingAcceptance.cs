using System.Diagnostics;
using System.Text;
using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunInstalledMathTypeOutgoing(string root)
    {
        var input = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_DOCUMENT")!;
        var target = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_TARGET")!;
        var output = Path.Combine(root, "installed-result.docx");
        File.Copy(input, output);
        var trace = Path.Combine(root, "installed-trace.log");
        string ReadTrace()
        {
            using var stream = new FileStream(trace, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", trace);
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", Path.Combine(root, "redraw.log"));
        Environment.SetEnvironmentVariable("VISUALTEX_REPLACE_INSTALLED", "1");
        Word.Application? app = null;
        Word.Document? doc = null;
        object? automation = null;
        Microsoft.Office.Core.COMAddIns? addins = null;
        Microsoft.Office.Core.COMAddIn? addin = null;
        using var stop = new ManualResetEventSlim();
        var dialogs = new List<object>();
        Thread? responder = null;
        try
        {
            app = CreateWordApplication(visible: false);
            doc = app.Documents.Open(output, ReadOnly: false, AddToRecentFiles: false, Visible: true);
            app.Visible = true; doc.Activate();
            Console.WriteLine($"[INSTALLED OUTGOING] UserControl={app.UserControl} CompatibilityMode={doc.CompatibilityMode}");
            var wordWindow = doc.ActiveWindow;
            uint wordPid;
            try { GetWindowThreadProcessId(new IntPtr(wordWindow.Hwnd), out wordPid); }
            finally { Release(wordWindow); }
            addins = app.COMAddIns;
            addin = addins.Item("VisualTeX.WordVsto");
            AssertTrue(addin.Connect, "Installed plugin not already loaded in the owned Word; do not change registration.");
            automation = addin.Object;
            // Across processes the registered IDispatch proxy is available;
            // the managed Ribbon interface has no system-wide proxy/type library.
            dynamic callbacks = automation;
            AssertTrue(ReadTrace().Contains("word-addin-loaded pid=" + wordPid), "No in-process plugin identity evidence.");
            responder = new Thread(() =>
            {
                var seen = new HashSet<IntPtr>();
                while (!stop.Wait(100))
                {
                    EnumWindows((window, _) =>
                    {
                        GetWindowThreadProcessId(window, out var pid);
                        if (pid != wordPid || !IsWindowVisible(window) || seen.Contains(window)) return true;
                        var title = new StringBuilder(256); GetWindowText(window, title, title.Capacity);
                        if (title.ToString() != "VisualTeX LaTeX 重绘" && title.ToString() != "VisualTeX 公式格式转换") return true;
                        var children = new List<string>();
                        IntPtr redraw = IntPtr.Zero;
                        OwnedWordNativeWindows.EnumChildWindows(window, (child, unused) =>
                        {
                            var text = new StringBuilder(2048); GetWindowText(child, text, text.Capacity);
                            if (text.Length > 0) children.Add(text.ToString());
                            if (text.ToString() == "开始重绘") redraw = child;
                            return true;
                        }, IntPtr.Zero);
                        var className = new StringBuilder(128); OwnedWordNativeWindows.GetClassName(window, className, className.Capacity);
                        if (redraw == IntPtr.Zero && className.ToString() != "#32770") return true;
                        var button = OutgoingDialogItem(window, 6) != IntPtr.Zero ? 6 : 1;
                        lock (dialogs) dialogs.Add(new { title = title.ToString(), children, button, window = window.ToInt64() });
                        seen.Add(window);
                        if (redraw != IntPtr.Zero) OutgoingPostMessage(redraw, 0x00F5, IntPtr.Zero, IntPtr.Zero);
                        else OutgoingPostMessage(window, 0x0111, new IntPtr(button), IntPtr.Zero);
                        return true;
                    }, IntPtr.Zero);
                }
            }) { IsBackground = true, Name = "Owned Word dialog acceptance" };
            responder.Start();
            void WaitFor(string success, string failure)
            {
                var watch = Stopwatch.StartNew();
                while (watch.Elapsed < TimeSpan.FromMinutes(12))
                {
                    System.Windows.Forms.Application.DoEvents();
                    var log = ReadTrace();
                    if (log.Contains(failure)) throw new InvalidDataException("Installed command failed; see installed-trace.log and dialogs.json.");
                    if (log.Contains(success)) { Thread.Sleep(300); return; }
                    Thread.Sleep(100);
                }
                throw new TimeoutException("Installed command did not finish: " + success);
            }
            callbacks.OnRedrawDocumentToMathType(null!);
            WaitFor("operation-progress stage=LaTeX 重绘完成", "redraw-failed");
            Console.WriteLine("[INSTALLED OUTGOING] redraw finished in Word PID=" + wordPid);
            var selectAll = Environment.GetEnvironmentVariable("VISUALTEX_OUTGOING_SELECTION") == "1";
            if (selectAll)
            {
                var content = doc.Content;
                try { content.Select(); }
                finally { Release(content); }
            }
            if (target == "wordOmml")
            {
                if (selectAll) callbacks.OnConvertMathTypeToOmmlSelection(null!);
                else callbacks.OnConvertMathTypeToOmmlDocument(null!);
            }
            else
            {
                if (selectAll) callbacks.OnConvertMathTypeToVisualTeXSelection(null!);
                else callbacks.OnConvertMathTypeToVisualTeXDocument(null!);
            }
            WaitFor("format-conversion-complete source=MathType", "format-conversion-failed");
            AssertEqual(0, CountMathTypeOleShapes(doc), "Installed outgoing source count.");
            doc.Save();
            Console.WriteLine("[INSTALLED OUTGOING] converted successfully.");
        }
        finally
        {
            stop.Set(); responder?.Join(2000);
            File.WriteAllText(Path.Combine(root, "dialogs.json"), JsonSerializer.Serialize(dialogs, new JsonSerializerOptions { WriteIndented = true }));
            Release(automation); Release(addin); Release(addins);
            if (doc is not null) { doc.Save(); doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); }
            if (app is not null) { QuitWordApplicationIfOwned(app); Release(app); }
            ForceComCleanup();
        }
    }
}

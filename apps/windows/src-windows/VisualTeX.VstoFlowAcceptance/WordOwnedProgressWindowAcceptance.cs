using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunOwnedProgressWindowAcceptance(string root)
    {
        using var host = new WordPerformanceHost(null);
        host.Application.Visible = true;
        var window = host.Document.ActiveWindow;
        IntPtr owner;
        try { owner = new IntPtr(window.Hwnd); }
        finally { Release(window); }
        var title = "VisualTeX 进度窗口验收 " + Guid.NewGuid().ToString("N");
        var type = typeof(ThisAddIn).Assembly.GetType("VisualTeX.WordVsto.WordOperationProgress", true)!;
        var disposed = false;
        var progress = (IDisposable)Activator.CreateInstance(type, BindingFlags.NonPublic | BindingFlags.Instance,
            null, new object[] { title, new Action(() => disposed = true), owner }, null)!;
        var report = type.GetMethod("Report", BindingFlags.NonPublic | BindingFlags.Instance)!;
        IntPtr handle = IntPtr.Zero;
        try
        {
            var wait = Stopwatch.StartNew();
            while (handle == IntPtr.Zero && wait.Elapsed < TimeSpan.FromSeconds(5))
            {
                OwnedWordNativeWindows.EnumWindows((candidate, _) =>
                {
                    GetWindowThreadProcessId(candidate, out var pid);
                    if (pid != (uint)Process.GetCurrentProcess().Id) return true;
                    var text = new StringBuilder(256); GetWindowText(candidate, text, text.Capacity);
                    if (text.ToString() == title) handle = candidate;
                    return true;
                }, IntPtr.Zero);
                Thread.Sleep(40);
            }
            AssertTrue(handle != IntPtr.Zero, "Progress window did not appear");
            report.Invoke(progress, new object[] { "正在写入目标公式…", 37, 200 });
            Thread.Sleep(1300); // Deliberately do not pump the host/Office dispatcher.
            var labels = new List<string>();
            OwnedWordNativeWindows.EnumChildWindows(handle, (child, _) =>
            {
                var text = new StringBuilder(512); GetWindowText(child, text, text.Capacity);
                if (text.Length != 0) labels.Add(text.ToString());
                return true;
            }, IntPtr.Zero);
            AssertTrue(labels.Any(x => x.Contains("37/200")), "Visible count failed while Office dispatcher was blocked");
            AssertTrue(labels.Any(x => x.Contains("00:01")), "Progress clock failed while Office dispatcher was blocked");
            AssertTrue(labels.Contains("正在写入目标公式…"), "Visible stage missing");
            File.WriteAllText(Path.Combine(root, "progress-window.json"), JsonSerializer.Serialize(new {
                handle=handle.ToInt64(), owner=owner.ToInt64(), blockedOfficeStaMilliseconds=1300,
                labels, passed=true }, new JsonSerializerOptions { WriteIndented=true }));
        }
        finally { progress.Dispose(); }
        AssertTrue(disposed, "Progress did not release its callback");
    }
}

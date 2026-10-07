using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using VisualTeX.WordVsto;
using Forms = System.Windows.Forms;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    // Owned windows only: the second Form represents a different application.
    // No saved foreground handle is restored (it may belong to the user).
    private static void RunProgressFocusAcceptance(string root)
    {
        using var owner = new Forms.Form { Text = "CP318 owned host", Width = 400, Height = 200 };
        using var other = new Forms.Form { Text = "CP318 owned other app", Width = 400, Height = 200 };
        owner.Show(); other.Show();
        var type = typeof(ThisAddIn).Assembly.GetType("VisualTeX.WordVsto.WordOperationProgress", true)!;
        var rows = new List<object>();
        var failures = 0;
        foreach (var foreground in new[] { owner, other })
        {
            foreground.Activate();
            SetForegroundWindow(foreground.Handle);
            Forms.Application.DoEvents();
            var expected = GetForegroundWindowForAcceptance();
            // A background test process may be denied foreground activation by
            // Windows. Record that missing prerequisite and still exercise the
            // passive progress lifecycle without forcing focus from the user.
            var prerequisite = expected == foreground.Handle;
            var title = "CP318 progress " + Guid.NewGuid().ToString("N");
            var progress = (IDisposable)Activator.CreateInstance(type, BindingFlags.NonPublic | BindingFlags.Instance,
                null, new object[] { title, new Action(() => { }), owner.Handle }, null)!;
            var transitions = new List<long>();
            var watch = Stopwatch.StartNew();
            IntPtr handle = IntPtr.Zero;
            while (watch.ElapsedMilliseconds < 900)
            {
                Forms.Application.DoEvents();
                var current = GetForegroundWindowForAcceptance().ToInt64();
                if (transitions.Count == 0 || transitions[transitions.Count - 1] != current) transitions.Add(current);
                EnumWindows((candidate, _) =>
                {
                    GetWindowThreadProcessId(candidate, out var pid);
                    if (pid != (uint)Process.GetCurrentProcess().Id) return true;
                    var text = new StringBuilder(128); GetWindowText(candidate, text, text.Capacity);
                    if (text.ToString() == title) handle = candidate;
                    return true;
                }, IntPtr.Zero);
                Thread.Sleep(10);
            }
            AssertTrue(handle != IntPtr.Zero, "Focus probe progress was never visible.");
            progress.Dispose();
            watch.Restart();
            while (watch.ElapsedMilliseconds < 400)
            {
                Forms.Application.DoEvents();
                var current = GetForegroundWindowForAcceptance().ToInt64();
                if (transitions.Count == 0 || transitions[transitions.Count - 1] != current) transitions.Add(current);
                Thread.Sleep(10);
            }
            var passed = transitions.All(h => h == expected.ToInt64());
            if (!passed) failures++;
            rows.Add(new { requested = foreground.Text, prerequisite, expected = expected.ToInt64(), progress = handle.ToInt64(), transitions, passed });
        }
        File.WriteAllText(Path.Combine(root, "focus.json"), JsonSerializer.Serialize(new { rows, failures }, new JsonSerializerOptions { WriteIndented = true }));
        AssertEqual(0, failures, "Progress must not activate itself or another window when shown/closed.");
    }
}

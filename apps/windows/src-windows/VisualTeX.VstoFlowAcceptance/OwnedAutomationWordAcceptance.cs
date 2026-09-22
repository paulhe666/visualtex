using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Never use CoCreateInstance for this test: Word may reuse an already running
    // /Automation class factory. Create a process, then acquire native OM ONLY
    // through that process's own document HWND.
    private static Word.Application CreateFreshRollbackAutomationWord(string artifactRoot)
    {
        var executable = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "Microsoft Office", "root", "Office16", "WINWORD.EXE");
        if (!File.Exists(executable)) throw new FileNotFoundException("Word executable is missing.", executable);
        var seed = Path.Combine(artifactRoot, "owned-empty-startup.docx");
        CreateEmptyWordPerformanceSeed(seed);
        var protectedIds = Process.GetProcessesByName("WINWORD").Select(process => process.Id).ToArray();
        var process = Process.Start(new ProcessStartInfo(executable, "/x /Automation \"" + seed + "\"")
        {
            UseShellExecute = false,
            WorkingDirectory = artifactRoot,
        }) ?? throw new InvalidOperationException("Could not start an independent /Automation Word.");
        var started = process.StartTime.ToUniversalTime();
        if (protectedIds.Contains(process.Id)) throw new InvalidOperationException("Word returned a pre-existing process ID.");
        File.WriteAllText(Path.Combine(artifactRoot, "owned-word-process.json"),
            System.Text.Json.JsonSerializer.Serialize(new { pid = process.Id, started, protectedIds,
                command = executable + " /x /Automation " + seed, seed }));
        try
        {
            var deadline = DateTime.UtcNow.AddSeconds(35);
            while (DateTime.UtcNow < deadline && !process.HasExited)
            {
                var children = new List<IntPtr>();
                OwnedWordNativeWindows.EnumWindows((window, _) =>
                {
                    OwnedWordNativeWindows.GetWindowThreadProcessId(window, out var pid);
                    if (pid != (uint)process.Id) return true;
                    OwnedWordNativeWindows.EnumChildWindows(window, (child, unused) =>
                    {
                        var name = new StringBuilder(128);
                        OwnedWordNativeWindows.GetClassName(child, name, name.Capacity);
                        if (name.ToString() == "_WwG") children.Add(child);
                        return true;
                    }, IntPtr.Zero);
                    return true;
                }, IntPtr.Zero);
                foreach (var child in children)
                {
                    object? native = null;
                    Word.Application? application = null;
                    Word.Document? document = null;
                    var transfer = false;
                    try
                    {
                        var dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                        var result = OwnedWordNativeWindows.AccessibleObjectFromWindow(child, 0xFFFFFFF0,
                            ref dispatch, out native);
                        if (result != 0 || native is not Word.Window window) continue;
                        application = window.Application;
                        document = window.Document;
                        OwnedWordNativeWindows.GetWindowThreadProcessId(new IntPtr(window.Hwnd), out var owner);
                        if (owner != (uint)process.Id || !string.Equals(document.FullName, seed, StringComparison.OrdinalIgnoreCase)
                            || application.Documents.Count != 1 || document.OMaths.Count != 0 || document.InlineShapes.Count != 0)
                            throw new InvalidDataException("Native Word object does not match the newly owned empty seed.");
                        application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                        application.Visible = false;
                        Console.WriteLine($"[EXPLICIT OWNED WORD] pid={process.Id}; seed={seed}; coCreateInstance=False");
                        transfer = true;
                        return application;
                    }
                    catch (COMException) { }
                    finally
                    {
                        Release(document);
                        if (!transfer) Release(application);
                        Release(native);
                    }
                }
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(80);
            }
            throw new InvalidOperationException("The newly created Word process did not expose its own native OM window.");
        }
        catch
        {
            // Only the exact process created above can be stopped on startup failure.
            if (!process.HasExited && process.StartTime.ToUniversalTime() == started
                && !protectedIds.Contains(process.Id)) process.Kill();
            throw;
        }
        finally { process.Dispose(); }
    }

    private static Word.Application CreateFreshAcceptanceAutomationWord(string artifactRoot)
    {
        var application = CreateFreshRollbackAutomationWord(artifactRoot);
        // Keep the owned startup seed open for the lifetime of the acceptance.
        // Closing Word's only /Automation document can let Office enter its
        // automatic shutdown path before the test document is fully established.
        // The acceptance creates/activates its own second document immediately;
        // QuitWordApplicationIfOwned closes both at the end.
        if (application.Documents.Count != 1 || application.ActiveDocument is null)
            throw new InvalidDataException("Owned /Automation Word did not retain exactly one startup seed.");
        return application;
    }

    private static class OwnedWordNativeWindows
    {
        internal delegate bool EnumerateWindow(IntPtr window, IntPtr context);
        [DllImport("user32.dll")] internal static extern bool EnumWindows(EnumerateWindow callback, IntPtr context);
        [DllImport("user32.dll")] internal static extern bool EnumChildWindows(IntPtr parent, EnumerateWindow callback, IntPtr context);
        [DllImport("user32.dll")] internal static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] internal static extern int GetClassName(IntPtr window, StringBuilder name, int maxCount);
        [DllImport("oleacc.dll")] internal static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId,
            ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object? result);
    }
}

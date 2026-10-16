using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using VisualTeX.WindowsOffice.Contracts;
using WinForms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // Reported: quick clicks while setting paragraph styles opened a formula in
    // VisualTeX. Word's Paragraph dialog lay over the formula and the double-click
    // hook only checked that Word was in front, not that the presses reached the
    // page. Real mouse input in an owned, visible Word, for an OMML and a selected
    // VisualTeX OLE formula:
    //   dialog   - double-click on the Paragraph dialog lying over the formula;
    //   close    - the first press closes the dialog (its X lies over the formula),
    //              the second lands on the formula; Windows sees no double-click;
    //   page     - with no dialog, a double-click on the formula opens VisualTeX.
    private static void RunWordOwnedDoubleClickOverlayAcceptance(
        VisualTeXSessionClient client,
        string root)
    {
        Environment.SetEnvironmentVariable("VISUALTEX_PERF_VISIBLE", "1");
        Environment.SetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH",
            Path.Combine(root, "double-click-trace.log"));
        var consoleWindow = GetConsoleWindow();
        if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 0);
        var results = new List<object>();
        var failures = new List<string>();
        try
        {
            using var host = new WordPerformanceHost(null);
            var application = host.Application;
            var document = host.Document;
            application.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            document.Content.Text = "Overlay double-click\r$$\\frac{a+b}{c}$$\r";
            host.AddIn.OnRedrawDocumentToOle(new object());
            WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(3));
            document.Content.InsertAfter("$$m+n=k$$\r");
            host.AddIn.OnRedrawDocumentToOmml(new object());
            WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(3));
            AssertEqual(1, document.InlineShapes.Count, "VisualTeX OLE seed count");
            AssertEqual(1, document.OMaths.Count, "OMML seed count");

            var window = application.ActiveWindow;
            var frame = new IntPtr(window.Hwnd);
            GetWindowThreadProcessId(frame, out var wordProcessId);
            const uint noMoveNoSizeShow = 0x0001 | 0x0002 | 0x0040;
            SetWindowPos(frame, new IntPtr(-1), 0, 0, 0, 0, noMoveNoSizeShow);
            SetForegroundWindow(frame);
            if (GetWindowRect(frame, out var frameRect))
            {
                // A click on the title bar makes the owned Word the foreground window.
                ClickAt(frameRect.Left + Math.Max(40, (frameRect.Right - frameRect.Left) / 2), frameRect.Top + 18);
                Thread.Sleep(700);
            }

            var cases = new (string Name, Func<Word.Range> Range, string Mode)[]
            {
                ("omml", () => document.OMaths[1].Range, FormulaOleContract.WordOmmlMode),
                ("ole", () => document.InlineShapes[1].Range, FormulaOleContract.NativeOleMode),
            };
            foreach (var (name, rangeOf, mode) in cases)
            {
                var range = rangeOf();
                try
                {
                    var (x, y) = SelectAndLocate(application, window, range);
                    Console.WriteLine($"[OVERLAY] {name}: formula center {x},{y}");

                    var dialogOutcome = DoubleClickThroughParagraphDialog(
                        application, wordProcessId, x, y, OverlayNative.HtCaption);
                    var dialogSession = TakeNewSession(client, host.AddIn, dialogOutcome.SessionsBefore);
                    if (dialogSession is not null)
                        failures.Add($"{name}: a double-click on the Paragraph dialog over the formula opened VisualTeX ({dialogSession}).");
                    if (!dialogOutcome.Covered)
                        failures.Add($"{name}: the Paragraph dialog could not be placed over the formula ({dialogOutcome.Note}).");

                    (x, y) = SelectAndLocate(application, window, range);
                    var closeOutcome = DoubleClickThroughParagraphDialog(
                        application, wordProcessId, x, y, OverlayNative.HtClose);
                    var closeSession = TakeNewSession(client, host.AddIn, closeOutcome.SessionsBefore);
                    if (closeSession is not null)
                        failures.Add($"{name}: closing the dialog and then pressing on the formula opened VisualTeX ({closeSession}).");
                    if (!closeOutcome.Covered)
                        failures.Add($"{name}: the dialog's close button could not be placed over the formula ({closeOutcome.Note}).");

                    (x, y) = SelectAndLocate(application, window, range);
                    var before = SnapshotSessionIds();
                    DoubleClickAt(x, y);
                    var pageSession = TryWaitForNewSession(before, TimeSpan.FromSeconds(20));
                    string? pageMode = null;
                    if (pageSession is null)
                    {
                        failures.Add($"{name}: a double-click on the formula itself did not open VisualTeX.");
                    }
                    else
                    {
                        var ready = WaitForUnchangedEditorReady(client, pageSession, TimeSpan.FromSeconds(20));
                        pageMode = ready.ObjectMode;
                        if (!string.Equals(ready.ObjectMode, mode, StringComparison.Ordinal))
                            failures.Add($"{name}: the formula double-click opened {ready.ObjectMode}, expected {mode}.");
                        CancelSession(client, host.AddIn, pageSession);
                    }
                    results.Add(new
                    {
                        formula = name,
                        dialogCovered = dialogOutcome.Covered,
                        dialogDoubleClickOpened = dialogSession is not null,
                        closeCovered = closeOutcome.Covered,
                        closeThenPageOpened = closeSession is not null,
                        pageDoubleClickOpened = pageSession is not null,
                        pageMode,
                    });
                    Console.WriteLine(
                        $"[OVERLAY] {name}: dialog={(dialogSession is null ? "no editor" : "EDITOR")} "
                        + $"close-then-page={(closeSession is null ? "no editor" : "EDITOR")} "
                        + $"page={(pageSession is null ? "NO EDITOR" : "editor")}");
                }
                finally { Release(range); }
            }
            Release(window);
        }
        finally
        {
            if (consoleWindow != IntPtr.Zero) ShowWindow(consoleWindow, 5);
            File.WriteAllText(
                Path.Combine(root, "double-click-overlay-result.json"),
                JsonSerializer.Serialize(
                    new { status = failures.Count == 0 ? "Passed" : "Failed", failures, results },
                    new JsonSerializerOptions { WriteIndented = true }));
        }
        if (failures.Count > 0)
            throw new InvalidOperationException(string.Join(Environment.NewLine, failures));
        Console.WriteLine("PASS Word double-click overlay acceptance");
    }

    private sealed class OverlayOutcome
    {
        internal HashSet<string> SessionsBefore { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        internal bool Covered { get; set; }
        internal string Note { get; set; } = string.Empty;
    }

    private static (int X, int Y) SelectAndLocate(Word.Application application, Word.Window window, Word.Range range)
    {
        range.Select();
        WinForms.Application.DoEvents();
        Thread.Sleep(250);
        window.GetPoint(out var left, out var top, out var width, out var height, range);
        AssertTrue(width > 0 && height > 0, "Word returned no screen rectangle for the formula.");
        // Leave time for a previous gesture to expire before the next one.
        Thread.Sleep(Math.Max(600, (int)OverlayNative.GetDoubleClickTime() + 100));
        return (left + width / 2, top + height / 2);
    }

    // Opens Word's Paragraph dialog (modal, so Display blocks this thread), moves
    // it so the requested part of its title bar lies over the formula center and
    // double-clicks there from a worker thread, then closes the dialog.
    private static OverlayOutcome DoubleClickThroughParagraphDialog(
        Word.Application application,
        uint wordProcessId,
        int x,
        int y,
        int hitTestPart)
    {
        var outcome = new OverlayOutcome { SessionsBefore = SnapshotSessionIds() };
        Exception? workerError = null;
        var worker = new Thread(() =>
        {
            try
            {
                var dialog = WaitForWordDialog(wordProcessId, TimeSpan.FromSeconds(15));
                if (dialog == IntPtr.Zero)
                {
                    outcome.Note = "Paragraph dialog did not appear";
                    return;
                }
                outcome.Covered = PlaceDialogPartAt(dialog, hitTestPart, x, y, out var note);
                outcome.Note = note;
                if (outcome.Covered)
                {
                    var gap = hitTestPart == OverlayNative.HtClose
                        ? Math.Min(250, (int)OverlayNative.GetDoubleClickTime() / 2)
                        : 90;
                    DoubleClickAt(x, y, gap);
                }
                Thread.Sleep(1500);
                for (var attempt = 0; attempt < 6 && OverlayNative.IsWindow(dialog) && IsWindowVisible(dialog); attempt++)
                {
                    OverlayNative.PostMessage(dialog, OverlayNative.WmSysCommand, new IntPtr(OverlayNative.ScClose), IntPtr.Zero);
                    Thread.Sleep(500);
                }
            }
            catch (Exception error) { workerError = error; }
        })
        {
            IsBackground = true,
            Name = "Overlay double-click driver",
        };
        worker.Start();
        Word.Dialogs? dialogs = null;
        Word.Dialog? paragraph = null;
        try
        {
            dialogs = application.Dialogs;
            paragraph = dialogs[Word.WdWordDialog.wdDialogFormatParagraph];
            object timeout = 30000;
            paragraph.Display(ref timeout);
        }
        finally
        {
            Release(paragraph);
            Release(dialogs);
        }
        AssertTrue(worker.Join(TimeSpan.FromSeconds(40)), "The overlay double-click driver did not finish.");
        if (workerError is not null) throw new InvalidOperationException("Overlay driver failed.", workerError);
        return outcome;
    }

    private static IntPtr WaitForWordDialog(uint wordProcessId, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            var found = IntPtr.Zero;
            EnumWindows((window, _) =>
            {
                GetWindowThreadProcessId(window, out var processId);
                if (processId != wordProcessId || !IsWindowVisible(window)) return true;
                var name = new StringBuilder(64);
                OverlayNative.GetClassName(window, name, name.Capacity);
                if (name.ToString() != "bosa_sdm_msword") return true;
                found = window;
                return false;
            }, IntPtr.Zero);
            if (found != IntPtr.Zero) return found;
            Thread.Sleep(100);
        }
        return IntPtr.Zero;
    }

    // Moves the dialog so a point of the requested title-bar part (caption or
    // close button) lands exactly on the formula center, and proves that the
    // dialog, not the page, is what a press there reaches.
    private static bool PlaceDialogPartAt(IntPtr dialog, int hitTestPart, int x, int y, out string note)
    {
        note = string.Empty;
        if (!GetWindowRect(dialog, out var rect))
        {
            note = "no dialog rectangle";
            return false;
        }
        var width = rect.Right - rect.Left;
        var candidates = new List<(int Dx, int Dy)>();
        for (var dy = 6; dy <= 40; dy += 4)
        {
            if (hitTestPart == OverlayNative.HtClose)
            {
                for (var dx = 6; dx <= 120; dx += 6) candidates.Add((width - dx, dy));
            }
            else
            {
                for (var dx = width / 3; dx <= width / 2; dx += 20) candidates.Add((dx, dy));
            }
        }
        foreach (var (dx, dy) in candidates)
        {
            OverlayNative.SetWindowPos(dialog, IntPtr.Zero, x - dx, y - dy, 0, 0,
                OverlayNative.SwpNoSize | OverlayNative.SwpNoZOrder | OverlayNative.SwpNoActivate);
            Thread.Sleep(60);
            var lParam = new IntPtr((y << 16) | (x & 0xFFFF));
            if (OverlayNative.SendMessageTimeout(dialog, OverlayNative.WmNcHitTest, IntPtr.Zero, lParam,
                    OverlayNative.SmtoAbortIfHung, 1000, out var code) == IntPtr.Zero)
                continue;
            if (code.ToInt64() != hitTestPart) continue;
            var reached = OverlayNative.GetAncestor(OverlayNative.WindowFromPoint(new OverlayNative.Point { X = x, Y = y }), 2);
            if (reached != dialog) continue;
            note = $"offset {dx},{dy}";
            return true;
        }
        note = "no title-bar point with hit-test " + hitTestPart;
        return false;
    }

    private static void ClickAt(int x, int y)
    {
        SetCursorPos(x, y);
        Thread.Sleep(60);
        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
    }

    private static void DoubleClickAt(int x, int y, int gapMilliseconds = 90)
    {
        SetCursorPos(x, y);
        Thread.Sleep(120);
        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
        Thread.Sleep(gapMilliseconds);
        mouse_event(MouseLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseLeftUp, 0, 0, 0, UIntPtr.Zero);
    }

    private static string? TryWaitForNewSession(HashSet<string> before, TimeSpan timeout)
    {
        try { return WaitForNewSession(before, "word", timeout); }
        catch (TimeoutException) { return null; }
    }

    // Waits a while for a session that must not appear; cancels it if it did.
    private static string? TakeNewSession(VisualTeXSessionClient client, object addIn, HashSet<string> before)
    {
        var opened = TryWaitForNewSession(before, TimeSpan.FromSeconds(4));
        if (opened is not null) CancelSession(client, addIn, opened);
        return opened;
    }

    private static void CancelSession(VisualTeXSessionClient client, object addIn, string sessionId)
    {
        try
        {
            client.PatchAsync(sessionId, new { status = "cancelled", explicitCancel = true }, CancellationToken.None)
                .GetAwaiter().GetResult();
        }
        catch { }
        try { client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult(); }
        catch { }
        WaitForAddInIdleFast(addIn, TimeSpan.FromSeconds(20));
    }

    private static class OverlayNative
    {
        internal const int HtCaption = 2;
        internal const int HtClose = 20;
        internal const uint WmNcHitTest = 0x0084;
        internal const uint WmSysCommand = 0x0112;
        internal const int ScClose = 0xF060;
        internal const uint SwpNoSize = 0x0001;
        internal const uint SwpNoZOrder = 0x0004;
        internal const uint SwpNoActivate = 0x0010;
        internal const uint SmtoAbortIfHung = 0x0002;

        [StructLayout(LayoutKind.Sequential)]
        internal struct Point { public int X; public int Y; }

        [DllImport("user32.dll")] internal static extern IntPtr WindowFromPoint(Point point);
        [DllImport("user32.dll")] internal static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] internal static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] internal static extern uint GetDoubleClickTime();
        [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
        internal static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
        [DllImport("user32.dll", EntryPoint = "PostMessageW")]
        internal static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW")]
        internal static extern IntPtr SendMessageTimeout(IntPtr window, uint message, IntPtr wParam, IntPtr lParam,
            uint flags, uint timeout, out IntPtr result);
        [DllImport("user32.dll")]
        internal static extern bool SetWindowPos(IntPtr window, IntPtr insertAfter, int x, int y, int width, int height, uint flags);
    }
}

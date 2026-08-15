using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

/// <summary>
/// MathType 7 fallback writer for existing Word-owned OLE equations.
///
/// Reading uses MathTypeOleInterop/IDataObject directly. MathType 7.8.2 on the
/// current acceptance machine exposes MathML through GetData but rejects MathML
/// SetData. For that version we edit the *existing Word-owned OLE equation* via
/// MathType's normal OLE editor, close only the current equation with Ctrl+F4,
/// and let Word/MathType commit through their native OLE relationship.
///
/// Never send Alt+F4/WM_CLOSE to the MathType application: doing so asks MathType
/// to exit while it is servicing Word and triggers the destructive forced-exit
/// warning reported by MathType itself.
/// </summary>
internal static class MathTypeOleBridge
{
    private const string MathTypeProcessName = "MathType";
    private const string EquationWindowClass = "EQNWINCLASS";
    private const string DialogWindowClass = "#32770";
    private const uint BmClick = 0x00F5;
    private const uint GwOwner = 4;
    private const int IdOk = 1;
    private const int IdYes = 6;
    private const int IdNo = 7;
    private const int IdCancel = 2;
    private static readonly object Gate = new();

    internal static string UpdateExistingOleFromLatex(
        Word.Application application,
        Word.InlineShape shape,
        string latex,
        string expectedMathMl)
    {
        if (application is null) throw new ArgumentNullException(nameof(application));
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException("VisualTeX did not provide LaTeX for the MathType equation.");
        if (string.IsNullOrWhiteSpace(expectedMathMl))
            throw new InvalidDataException("VisualTeX did not provide MathML for MathType verification.");

        lock (Gate)
        {
            // Capture the original equation through the non-UI IDataObject path so
            // a failed native-editor commit can be rolled back through the exact
            // same MathType OLE editor without synthesizing MTEF ourselves.
            var originalMathMl = MathTypeOleInterop.ReadMathMl(shape);
            var originalLatex = MathMlToLatexConverter.Convert(originalMathMl).Trim();
            if (string.IsNullOrWhiteSpace(originalLatex))
                throw new InvalidDataException(
                    "The source MathType equation could not be converted to a rollback-safe LaTeX representation.");

            EditWordOwnedOle(application, shape, latex, saveChanges: true);

            try
            {
                var actualMathMl = MathTypeOleInterop.ReadMathMl(shape);
                var expectedLatex = MathMlToLatexConverter.Convert(expectedMathMl).Trim();
                var actualLatex = MathMlToLatexConverter.Convert(actualMathMl).Trim();
                if (string.IsNullOrWhiteSpace(actualLatex))
                    throw new InvalidDataException(
                        "MathType saved the OLE equation but VisualTeX could not read it back.");
                if (!AreEquivalentLatex(expectedLatex, actualLatex))
                {
                    throw new InvalidDataException(
                        "MathType did not round-trip the VisualTeX equation faithfully. "
                        + $"Expected='{expectedLatex}', actual='{actualLatex}'.");
                }
                return actualMathMl;
            }
            catch
            {
                // The first edit has already been committed to the Word OLE. Put
                // the original formula back using MathType itself, then rethrow.
                try
                {
                    EditWordOwnedOle(application, shape, originalLatex, saveChanges: true);
                    _ = MathTypeOleInterop.ReadMathMl(shape);
                }
                catch
                {
                    // Preserve the original verification exception. A second
                    // failure here is intentionally not hidden by deleting or
                    // replacing the third-party OLE object.
                }
                throw;
            }
        }
    }

    private static void EditWordOwnedOle(
        Word.Application application,
        Word.InlineShape shape,
        string latex,
        bool saveChanges)
    {
        Word.OLEFormat? format = null;
        Word.Window? wordWindow = null;
        var clipboardBackup = TryCaptureClipboard();
        var previousForeground = GetForegroundWindow();
        Exception? driverError = null;
        using var driverFinished = new ManualResetEventSlim(false);

        try
        {
            var existingMathTypeWindows = EnumerateMathTypeEquationWindows();
            if (existingMathTypeWindows.Count > 0)
            {
                throw new InvalidOperationException(
                    "MathType currently has an interactive equation window open. "
                    + "Finish that MathType edit before VisualTeX writes this OLE equation.");
            }

            wordWindow = application.ActiveWindow
                ?? throw new InvalidOperationException("Word has no active window for MathType OLE editing.");
            var wordHwnd = new IntPtr(wordWindow.Hwnd);
            SetForegroundWindow(wordHwnd);
            Thread.Sleep(120);

            var driver = new Thread(() =>
            {
                IntPtr editorWindow = IntPtr.Zero;
                uint mathTypePid = 0;
                try
                {
                    (editorWindow, mathTypePid) = WaitForWordOwnedMathTypeEditor(
                        wordHwnd,
                        TimeSpan.FromSeconds(15));

                    SetForegroundWindow(editorWindow);
                    Thread.Sleep(250);
                    System.Windows.Forms.SendKeys.SendWait("^a");
                    Thread.Sleep(60);
                    System.Windows.Forms.SendKeys.SendWait("{DELETE}");
                    Thread.Sleep(80);
                    System.Windows.Forms.Clipboard.SetText(
                        latex,
                        System.Windows.Forms.TextDataFormat.UnicodeText);
                    System.Windows.Forms.SendKeys.SendWait("^v");
                    Thread.Sleep(350);

                    ThrowForUnexpectedMathTypeDialogs(mathTypePid, editorWindow);

                    // Close only the current MathType equation/document. Ctrl+F4
                    // is materially different from Alt+F4: the latter exits the
                    // MathType application while it is acting as Word's OLE server.
                    System.Windows.Forms.SendKeys.SendWait("^{F4}");
                    WaitForEquationSessionToClose(
                        editorWindow,
                        mathTypePid,
                        saveChanges,
                        TimeSpan.FromSeconds(10));
                }
                catch (Exception error)
                {
                    driverError = error;
                    if (editorWindow != IntPtr.Zero && IsWindow(editorWindow))
                        TryAbortEquationSession(editorWindow, mathTypePid);
                }
                finally
                {
                    driverFinished.Set();
                }
            })
            {
                IsBackground = true,
                Name = "VisualTeX MathType OLE writer",
            };
            driver.SetApartmentState(ApartmentState.STA);
            driver.Start();

            format = shape.OLEFormat;
            object openVerb = (int)Word.WdOLEVerb.wdOLEVerbOpen;
            format.DoVerb(ref openVerb);

            if (!driverFinished.Wait(TimeSpan.FromSeconds(20)))
                throw new TimeoutException(
                    "MathType OLE editor did not finish the VisualTeX update within the safety timeout.");
            if (driverError is not null)
                throw new InvalidOperationException(
                    "MathType could not safely update the Word-owned OLE equation.",
                    driverError);
        }
        finally
        {
            TryRestoreClipboard(clipboardBackup);
            RestoreForeground(previousForeground, application);
            Release(format);
            Release(wordWindow);
        }
    }

    private static (IntPtr Handle, uint ProcessId) WaitForWordOwnedMathTypeEditor(
        IntPtr wordWindow,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            var foreground = GetForegroundWindow();
            if (foreground != IntPtr.Zero && foreground != wordWindow
                && IsMathTypeEquationWindow(foreground, out var foregroundPid))
                return (foreground, foregroundPid);

            foreach (var candidate in EnumerateMathTypeEquationWindows())
                return (candidate.Handle, candidate.ProcessId);

            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(60);
        }
        throw new TimeoutException(
            "MathType did not expose the Word-owned equation editor window.");
    }

    private static void WaitForEquationSessionToClose(
        IntPtr editorWindow,
        uint processId,
        bool saveChanges,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            HandleCloseDialogs(processId, editorWindow, saveChanges);
            if (!IsWindow(editorWindow)) return;
            Thread.Sleep(70);
        }
        throw new TimeoutException(
            "MathType did not close the current Word-owned equation session.");
    }

    private static void HandleCloseDialogs(
        uint processId,
        IntPtr editorWindow,
        bool saveChanges)
    {
        string? unexpected = null;
        EnumWindows((dialog, _) =>
        {
            GetWindowThreadProcessId(dialog, out var pid);
            if (pid != processId || !IsDialog(dialog)) return true;
            var owner = GetWindow(dialog, GwOwner);
            if (owner != IntPtr.Zero && owner != editorWindow) return true;

            var text = GetDialogText(dialog);
            var title = GetWindowTitle(dialog);
            if (IsForcedExitWarning(text))
            {
                // Never confirm an attempt to terminate MathType while it services
                // Word. Cancel it if possible and fail the update.
                var cancel = GetDlgItem(dialog, IdCancel);
                if (cancel != IntPtr.Zero)
                    SendMessage(cancel, BmClick, IntPtr.Zero, IntPtr.Zero);
                unexpected =
                    $"MathType reported a forced-exit warning while servicing Word. Title='{title}', Text='{text}'.";
                return false;
            }

            if (IsSavePrompt(text))
            {
                var button = GetDlgItem(dialog, saveChanges ? IdYes : IdNo);
                if (button == IntPtr.Zero)
                {
                    unexpected =
                        $"MathType save prompt did not expose the expected button. Title='{title}', Text='{text}'.";
                    return false;
                }
                SendMessage(button, BmClick, IntPtr.Zero, IntPtr.Zero);
                return true;
            }

            unexpected = $"Unexpected MathType dialog. Title='{title}', Text='{text}'.";
            return false;
        }, IntPtr.Zero);

        if (!string.IsNullOrWhiteSpace(unexpected))
            throw new InvalidOperationException(unexpected);
    }

    private static void ThrowForUnexpectedMathTypeDialogs(uint processId, IntPtr editorWindow)
    {
        string? unexpected = null;
        EnumWindows((dialog, _) =>
        {
            GetWindowThreadProcessId(dialog, out var pid);
            if (pid != processId || !IsDialog(dialog)) return true;
            var owner = GetWindow(dialog, GwOwner);
            if (owner != IntPtr.Zero && owner != editorWindow) return true;
            var text = GetDialogText(dialog);
            if (IsForcedExitWarning(text))
            {
                var cancel = GetDlgItem(dialog, IdCancel);
                if (cancel != IntPtr.Zero)
                    SendMessage(cancel, BmClick, IntPtr.Zero, IntPtr.Zero);
            }
            unexpected =
                $"MathType opened a dialog before the equation could be committed. Title='{GetWindowTitle(dialog)}', Text='{text}'.";
            return false;
        }, IntPtr.Zero);
        if (!string.IsNullOrWhiteSpace(unexpected))
            throw new InvalidOperationException(unexpected);
    }

    private static void TryAbortEquationSession(IntPtr editorWindow, uint processId)
    {
        try
        {
            if (!IsWindow(editorWindow)) return;
            SetForegroundWindow(editorWindow);
            System.Windows.Forms.SendKeys.SendWait("^{F4}");
            var deadline = DateTime.UtcNow + TimeSpan.FromSeconds(3);
            while (DateTime.UtcNow < deadline && IsWindow(editorWindow))
            {
                try { HandleCloseDialogs(processId, editorWindow, saveChanges: false); }
                catch { }
                Thread.Sleep(70);
            }
        }
        catch { }
    }

    private static bool AreEquivalentLatex(string expected, string actual)
    {
        var left = CanonicalizeLatex(expected);
        var right = CanonicalizeLatex(actual);
        return string.Equals(left, right, StringComparison.Ordinal);
    }

    private static string CanonicalizeLatex(string value)
    {
        var normalized = Regex.Replace(value ?? string.Empty, @"\s+", string.Empty)
            .Replace(@"\left", string.Empty)
            .Replace(@"\right", string.Empty);
        normalized = Regex.Replace(normalized, @"([_^])\{([A-Za-z0-9])\}", "$1$2");
        return normalized;
    }

    private static List<(IntPtr Handle, uint ProcessId)> EnumerateMathTypeEquationWindows()
    {
        var result = new List<(IntPtr, uint)>();
        EnumWindows((window, _) =>
        {
            if (!IsWindowVisible(window)
                || !string.Equals(GetClassName(window), EquationWindowClass, StringComparison.Ordinal))
                return true;
            if (IsMathTypeEquationWindow(window, out var pid))
                result.Add((window, pid));
            return true;
        }, IntPtr.Zero);
        return result;
    }

    private static bool IsMathTypeEquationWindow(IntPtr window, out uint processId)
    {
        processId = 0;
        if (!string.Equals(GetClassName(window), EquationWindowClass, StringComparison.Ordinal))
            return false;
        GetWindowThreadProcessId(window, out processId);
        if (processId == 0) return false;
        try
        {
            using var process = Process.GetProcessById((int)processId);
            return string.Equals(process.ProcessName, MathTypeProcessName, StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static bool IsDialog(IntPtr window) =>
        string.Equals(GetClassName(window), DialogWindowClass, StringComparison.Ordinal);

    private static bool IsSavePrompt(string text) =>
        text.IndexOf("保存", StringComparison.OrdinalIgnoreCase) >= 0
        || text.IndexOf("save", StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool IsForcedExitWarning(string text) =>
        (text.IndexOf("另一个应用", StringComparison.OrdinalIgnoreCase) >= 0
            || text.IndexOf("another application", StringComparison.OrdinalIgnoreCase) >= 0)
        && (text.IndexOf("退出", StringComparison.OrdinalIgnoreCase) >= 0
            || text.IndexOf("exit", StringComparison.OrdinalIgnoreCase) >= 0);

    private static string GetDialogText(IntPtr dialog)
    {
        var parts = new List<string>();
        EnumChildWindows(dialog, (child, _) =>
        {
            if (string.Equals(GetClassName(child), "Static", StringComparison.Ordinal))
            {
                var text = GetWindowTitle(child);
                if (!string.IsNullOrWhiteSpace(text)) parts.Add(text);
            }
            return true;
        }, IntPtr.Zero);
        return string.Join(" ", parts);
    }

    private static string GetWindowTitle(IntPtr window)
    {
        var length = GetWindowTextLength(window);
        if (length <= 0) return string.Empty;
        var builder = new StringBuilder(length + 1);
        GetWindowText(window, builder, builder.Capacity);
        return builder.ToString();
    }

    private static string GetClassName(IntPtr window)
    {
        var builder = new StringBuilder(128);
        return GetClassNameNative(window, builder, builder.Capacity) > 0
            ? builder.ToString()
            : string.Empty;
    }

    private static System.Windows.Forms.IDataObject? TryCaptureClipboard()
    {
        try { return System.Windows.Forms.Clipboard.GetDataObject(); }
        catch { return null; }
    }

    private static void TryRestoreClipboard(System.Windows.Forms.IDataObject? backup)
    {
        if (backup is null) return;
        try { System.Windows.Forms.Clipboard.SetDataObject(backup, true); }
        catch { }
    }

    private static void RestoreForeground(IntPtr previousForeground, Word.Application application)
    {
        try
        {
            if (previousForeground != IntPtr.Zero && IsWindow(previousForeground))
            {
                SetForegroundWindow(previousForeground);
                return;
            }
            Word.Window? activeWindow = null;
            try
            {
                activeWindow = application.ActiveWindow;
                if (activeWindow is not null)
                    SetForegroundWindow(new IntPtr(activeWindow.Hwnd));
            }
            finally { Release(activeWindow); }
        }
        catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }

    private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(
        IntPtr parent,
        EnumWindowsProc callback,
        IntPtr parameter);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindow(IntPtr window);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr window);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr window, StringBuilder text, int maxCount);

    [DllImport("user32.dll")]
    private static extern int GetWindowTextLength(IntPtr window);

    [DllImport("user32.dll", EntryPoint = "GetClassNameW", CharSet = CharSet.Unicode)]
    private static extern int GetClassNameNative(IntPtr window, StringBuilder className, int maxCount);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr window, uint command);

    [DllImport("user32.dll")]
    private static extern IntPtr GetDlgItem(IntPtr dialog, int controlId);

    [DllImport("user32.dll", EntryPoint = "SendMessageW")]
    private static extern IntPtr SendMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();
}

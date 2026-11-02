using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;
using WordApplication = Microsoft.Office.Interop.Word.Application;
using WordView = Microsoft.Office.Interop.Word.View;
using WordWindow = Microsoft.Office.Interop.Word.Window;
using WordWindows = Microsoft.Office.Interop.Word.Windows;

namespace VisualTeX.WordVsto;

/// <summary>
/// WPS forgets the drawn pictures of the OLE formulas already in a document (VisualTeX and MathType alike)
/// whenever Range.InsertXML or Range.InsertFile imports into it: the formulas that are off screen then show blank
/// until their paragraph is laid out again (a click sometimes does that); the saved document reopens intact.
/// Measured on WPS 12.1.0.29162 with a package of one text run; a Range.FormattedText copy does not do it, and
/// Word is not affected. Showing the window's field codes and hiding them again lays every EMBED field out anew
/// and brings all of them back (75 ms for 160 formulas, scroll position kept); ScreenRefresh, the view type, zoom,
/// Repaginate, picture placeholders and formatting marks do not.
/// Every import into a WPS document marks that document. Once the add-in's work on the UI thread is over (no
/// custom undo record open, screen updating back on), the windows of each marked document show their field codes
/// and hide them again: a view setting, so there is no undo step and the document's saved state is unchanged.
/// </summary>
internal static class WpsOleDisplayRefresh
{
    private const int BusyRetryMilliseconds = 200;
    // A batch that never turns screen updating back on must not postpone the refresh forever (30 s).
    private const int MaximumBusyRetries = 150;
    private static readonly HashSet<string> Marked = new(StringComparer.OrdinalIgnoreCase);
    private static WordApplication? application;
    private static Action<Action>? post;
    private static bool scheduled;
    private static int busyRetries;

    /// <summary>The add-in's Word-compatible application and its UI dispatcher's Post; nulls at shutdown.</summary>
    internal static void Attach(WordApplication? host, Action<Action>? uiPost)
    {
        application = host;
        post = uiPost;
        Marked.Clear();
        scheduled = false;
        busyRetries = 0;
    }

    /// <summary>An import into the document of <paramref name="range"/> (before or after it: the refresh comes later).</summary>
    internal static void Mark(Range range)
    {
        if (!OfficeHostEnvironment.IsWps || post is null) return;
        Document? document = null;
        try
        {
            document = range.Document;
            Mark(document);
        }
        catch (COMException) { }
        finally { Release(document); }
    }

    internal static void Mark(Document document)
    {
        if (!OfficeHostEnvironment.IsWps || post is null) return;
        string name;
        try { name = document.FullName; }
        catch (COMException) { return; }
        Marked.Add(name);
        if (!scheduled) Schedule(0);
    }

    private static void Schedule(int delayMilliseconds)
    {
        var uiPost = post;
        if (uiPost is null) return;
        scheduled = true;
        if (delayMilliseconds <= 0)
            uiPost(Run);
        else
            _ = System.Threading.Tasks.Task.Delay(delayMilliseconds).ContinueWith(
                _ => uiPost(Run),
                System.Threading.Tasks.TaskScheduler.Default);
    }

    private static void Run()
    {
        scheduled = false;
        var host = application;
        if (host is null || Marked.Count == 0) return;
        if (Busy(host) && busyRetries < MaximumBusyRetries)
        {
            busyRetries++;
            Schedule(BusyRetryMilliseconds);
            return;
        }
        busyRetries = 0;
        var names = new HashSet<string>(Marked, StringComparer.OrdinalIgnoreCase);
        Marked.Clear();
        Documents? documents = null;
        try
        {
            documents = host.Documents;
            for (var index = 1; index <= documents.Count; index++)
            {
                Document? document = null;
                try
                {
                    document = documents[index];
                    if (names.Contains(document.FullName)) LayOutFieldsAgain(document);
                }
                catch (COMException error)
                {
                    WordDoubleClickHook.TraceMessage($"wps-ole-display-refresh-failed error={error.Message}");
                }
                finally { Release(document); }
            }
        }
        catch (COMException error)
        {
            WordDoubleClickHook.TraceMessage($"wps-ole-display-refresh-failed error={error.Message}");
        }
        finally { Release(documents); }
    }

    // The add-in is still writing: a custom undo record is open (in WPS a member call that fails inside it rolls
    // the whole record back), or a batch has screen updating suspended.
    private static bool Busy(WordApplication host)
    {
        UndoRecord? undo = null;
        try
        {
            if (!host.ScreenUpdating) return true;
            undo = host.UndoRecord;
            return undo.IsRecordingCustomRecord;
        }
        catch (COMException) { return false; }
        finally { Release(undo); }
    }

    private static void LayOutFieldsAgain(Document document)
    {
        WordWindows? windows = null;
        try
        {
            windows = document.Windows;
            var count = windows.Count;
            for (var index = 1; index <= count; index++)
            {
                WordWindow? window = null;
                WordView? view = null;
                try
                {
                    window = windows[index];
                    view = window.View;
                    var shown = view.ShowFieldCodes;
                    view.ShowFieldCodes = !shown;
                    view.ShowFieldCodes = shown;
                }
                finally
                {
                    Release(view);
                    Release(window);
                }
            }
            WordDoubleClickHook.TraceMessage($"wps-ole-display-refreshed windows={count}");
        }
        finally { Release(windows); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}

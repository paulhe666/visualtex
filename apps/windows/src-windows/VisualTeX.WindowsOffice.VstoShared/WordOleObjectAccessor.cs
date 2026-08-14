using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal static class WordOleObjectAccessor
{
    public static object GetRunningObject(OLEFormat format)
    {
        if (format is null) throw new ArgumentNullException(nameof(format));
        var runningObject = TryGetObject(format);
        if (runningObject is not null) return runningObject;

        // A freshly pasted embedded OLE object is not guaranteed to expose its
        // in-proc automation object synchronously with DoVerb(). Office 2019/2021
        // and machines with slower COM activation can return from DoVerb while
        // the server is still transitioning to the running state. The old code
        // performed exactly one immediate format.Object read, which made pasted
        // formulas intermittently look like ordinary non-VisualTeX objects.
        Exception? activationError = null;
        object showVerb = (int)WdOLEVerb.wdOLEVerbShow;
        try { format.DoVerb(ref showVerb); }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            activationError = error;
        }

        // Keep the UI-thread retry bounded. In the normal case the first read
        // succeeds; the backoff exists only for a dormant copied/pasted object.
        var delayMilliseconds = 15;
        for (var attempt = 0; attempt < 8; attempt++)
        {
            runningObject = TryGetObject(format);
            if (runningObject is not null) return runningObject;
            if (attempt == 7) break;
            Thread.Sleep(delayMilliseconds);
            delayMilliseconds = Math.Min(delayMilliseconds * 2, 60);
        }

        throw new COMException(
            "Word activated the VisualTeX OLE object but did not expose its running COM object."
            + (activationError is null ? string.Empty : $" {activationError.Message}"));
    }

    private static object? TryGetObject(OLEFormat format)
    {
        try { return format.Object; }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            return null;
        }
    }
}

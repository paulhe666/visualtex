using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal static class WordOleObjectAccessor
{
    public static object GetRunningObject(OLEFormat format)
    {
        if (format is null) throw new ArgumentNullException(nameof(format));
        try
        {
            return format.Object;
        }
        catch (Exception error) when (error is COMException or InvalidCastException)
        {
            object showVerb = (int)WdOLEVerb.wdOLEVerbShow;
            format.DoVerb(ref showVerb);
            return format.Object;
        }
    }
}

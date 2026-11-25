using System.Diagnostics;

namespace VisualTeX.WordVsto;

/// <summary>
/// Whether a MathType program (MathType's editor, or Equation Editor 3.0) runs in the user's session now. While one
/// does it may be building or editing an object in the document, and exporting that object's WordOpenXML (which makes
/// Word save the OLE data) would re-enter it; with none running, a MathType object is plain stored data. VisualTeX's
/// own MathType renderer runs isolated in another session (VisualTeX.NativeRenderWorker) and does not count. Checked
/// at most every two seconds.
/// </summary>
internal static class MathTypeServerProcess
{
    private static readonly string[] Names = { "MathType", "EQNEDT32" };
    private static readonly long Interval = Stopwatch.Frequency * 2;
    private static long _checkedAt = long.MinValue;
    private static bool _running;

    internal static bool IsRunning()
    {
        var now = Stopwatch.GetTimestamp();
        if (_checkedAt != long.MinValue && now - _checkedAt < Interval) return _running;
        var running = false;
        try
        {
            int session;
            using (var self = Process.GetCurrentProcess()) session = self.SessionId;
            foreach (var name in Names)
            {
                foreach (var process in Process.GetProcessesByName(name))
                {
                    try { if (process.SessionId == session) running = true; }
                    catch { running = true; }   // cannot tell: treat as running
                    finally { process.Dispose(); }
                }
                if (running) break;
            }
        }
        catch { running = true; }
        _running = running;
        _checkedAt = now;
        return running;
    }
}

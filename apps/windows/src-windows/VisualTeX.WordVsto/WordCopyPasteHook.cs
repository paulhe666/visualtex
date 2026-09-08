using System.Diagnostics;
using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal enum WordClipboardGesture
{
    Copy,
}

internal readonly struct WordClipboardGestureEvent
{
    internal WordClipboardGestureEvent(
        WordClipboardGesture gesture,
        uint clipboardSequence,
        long observedTimestamp)
    {
        Gesture = gesture;
        ClipboardSequence = clipboardSequence;
        ObservedTimestamp = observedTimestamp;
    }

    internal WordClipboardGesture Gesture { get; }
    internal uint ClipboardSequence { get; }
    internal long ObservedTimestamp { get; }
}

/// <summary>
/// Observes Windows clipboard sequence changes while this add-in's Word process
/// owns the foreground window. This covers keyboard, Ribbon and context-menu Copy
/// without intercepting Word's native command or touching COM off the Office STA.
/// Paste is detected from Word's own WindowSelectionChange event: the clipboard
/// sequence does not change on Paste, while the OLE/OMath inventory does.
/// </summary>
internal sealed class WordCopyPasteHook : IDisposable
{
    private readonly Action<WordClipboardGestureEvent> _callback;
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private readonly ManualResetEventSlim _stop = new(false);
    private int _ownerProcessId = Process.GetCurrentProcess().Id;
    private uint _lastClipboardSequence;

    internal WordCopyPasteHook(Action<WordClipboardGestureEvent> callback)
    {
        _callback = callback;
        _thread = new Thread(Run)
        {
            IsBackground = true,
            Name = "VisualTeX Word Clipboard Observer",
        };
    }

    internal void BindOwnerWindow(int windowHandle)
    {
        if (windowHandle == 0) return;
        GetWindowThreadProcessId(new IntPtr(windowHandle), out var processId);
        if (processId == 0) return;
        Interlocked.Exchange(ref _ownerProcessId, unchecked((int)processId));
    }

    internal void Start()
    {
        _lastClipboardSequence = GetClipboardSequenceNumber();
        _thread.Start();
        if (!_ready.Wait(TimeSpan.FromSeconds(5)))
            throw new TimeoutException("Word clipboard observer did not start.");
        WordDoubleClickHook.TraceMessage(
            $"copy-paste-observer-started clipboardSequence={_lastClipboardSequence}");
    }

    internal static uint CurrentClipboardSequence => GetClipboardSequenceNumber();

    private void Run()
    {
        _ready.Set();
        while (!_stop.Wait(45))
        {
            var current = GetClipboardSequenceNumber();
            if (current == _lastClipboardSequence) continue;
            _lastClipboardSequence = current;
            if (!IsWordForeground()) continue;

            // Word can update a rich clipboard payload in several closely spaced
            // native writes. Wait briefly for the final sequence before capturing
            // the still-selected source formula on the Office STA.
            Thread.Sleep(65);
            current = GetClipboardSequenceNumber();
            _lastClipboardSequence = current;
            if (!IsWordForeground()) continue;
            try
            {
                WordDoubleClickHook.TraceMessage(
                    $"copy-paste-clipboard-change clipboardSequence={current}");
                _callback(new WordClipboardGestureEvent(
                    WordClipboardGesture.Copy,
                    current,
                    Stopwatch.GetTimestamp()));
            }
            catch (Exception error)
            {
                WordDoubleClickHook.TraceMessage(
                    $"copy-paste-callback-error {error.GetType().Name}: {error.Message}");
            }
        }
    }

    private bool IsWordForeground()
    {
        var window = GetForegroundWindow();
        if (window == IntPtr.Zero) return false;
        GetWindowThreadProcessId(window, out var processId);
        return WordDoubleClickRouting.ForegroundProcessBelongsToOwner(
            processId,
            Volatile.Read(ref _ownerProcessId));
    }

    public void Dispose()
    {
        _stop.Set();
        _thread.Join(TimeSpan.FromSeconds(2));
        _stop.Dispose();
        _ready.Dispose();
    }

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll")]
    private static extern uint GetClipboardSequenceNumber();
}

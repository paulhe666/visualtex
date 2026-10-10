using System.Runtime.InteropServices;
using System.Text;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// PowerPoint can only move an OLE shape between presentations through the
/// clipboard. These helpers keep that transport from disturbing the user: the
/// user's clipboard is saved before VisualTeX copies and put back afterwards, and
/// waits serve other programs' clipboard requests instead of blocking them.
/// All members run on the Office UI thread.
/// </summary>
internal static class PowerPointClipboard
{
    private const uint CfText = 1, CfBitmap = 2, CfMetafilePict = 3, CfOemText = 7, CfDib = 8,
        CfPalette = 9, CfUnicodeText = 13, CfEnhMetafile = 14, CfLocale = 16, CfDibV5 = 17;
    private const long MaximumSnapshotBytes = 64L * 1024 * 1024;
    private static int _preserveDepth;
    private static Snapshot? _outerSnapshot;
    private static OwnerWindow? _owner;

    /// <summary>
    /// Saves the user's clipboard now and restores it when disposed. Nested scopes
    /// share the outermost snapshot.
    /// </summary>
    internal static IDisposable Preserve()
    {
        if (_preserveDepth++ == 0) _outerSnapshot = Snapshot.TryCapture();
        return new Scope();
    }

    // Serve only sent messages (WM_RENDERFORMAT from a program reading data
    // PowerPoint copied lazily); posted work such as Ribbon callbacks never runs.
    internal static void ServeSentMessages(int milliseconds)
    {
        const uint QsSendMessage = 0x0040;
        const uint PmQsSendMessage = QsSendMessage << 16;
        var deadline = Environment.TickCount + milliseconds;
        for (int remaining; (remaining = deadline - Environment.TickCount) > 0;)
        {
            MsgWaitForMultipleObjectsEx(0, IntPtr.Zero, (uint)remaining, QsSendMessage, 0);
            PeekMessage(out _, IntPtr.Zero, 0, 0, PmQsSendMessage);
        }
    }

    internal static string Holder()
    {
        try
        {
            var window = GetOpenClipboardWindow();
            if (window == IntPtr.Zero) return "nobody";
            GetWindowThreadProcessId(window, out var processId);
            using var process = System.Diagnostics.Process.GetProcessById((int)processId);
            return process.ProcessName;
        }
        catch { return "unknown"; }
    }

    private static bool TryOpen(IntPtr owner)
    {
        for (var attempt = 0; attempt < 10; attempt++)
        {
            if (OpenClipboard(owner)) return true;
            ServeSentMessages(30 + 20 * attempt);
        }
        return false;
    }

    private sealed class Scope : IDisposable
    {
        private bool _disposed;
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            if (--_preserveDepth != 0) return;
            var snapshot = _outerSnapshot;
            _outerSnapshot = null;
            snapshot?.Restore();
        }
    }

    private sealed class Snapshot
    {
        private readonly List<(uint Format, byte[]? Bytes, IntPtr Metafile)> _items;

        private Snapshot(List<(uint, byte[]?, IntPtr)> items) => _items = items;

        internal static Snapshot? TryCapture()
        {
            try
            {
                if (!TryOpen(IntPtr.Zero)) return null;
                var items = new List<(uint, byte[]?, IntPtr)>();
                try
                {
                    var formats = new List<uint>();
                    for (var format = EnumClipboardFormats(0); format != 0; format = EnumClipboardFormats(format))
                        formats.Add(format);
                    long total = 0;
                    foreach (var format in formats)
                    {
                        if (IsSkipped(format, formats)) continue;
                        var handle = GetClipboardData(format);
                        if (handle == IntPtr.Zero) continue;
                        if (format == CfEnhMetafile)
                        {
                            var copy = CopyEnhMetaFile(handle, null);
                            if (copy != IntPtr.Zero) items.Add((format, null, copy));
                            continue;
                        }
                        var size = (long)GlobalSize(handle).ToUInt64();
                        if (size <= 0) continue;
                        total += size;
                        // Too large to hold safely: leave the transport's content
                        // rather than restore a partial clipboard.
                        if (total > MaximumSnapshotBytes) { Release(items); return null; }
                        var pointer = GlobalLock(handle);
                        if (pointer == IntPtr.Zero) continue;
                        try
                        {
                            var bytes = new byte[size];
                            Marshal.Copy(pointer, bytes, 0, bytes.Length);
                            items.Add((format, bytes, IntPtr.Zero));
                        }
                        finally { GlobalUnlock(handle); }
                    }
                }
                finally { CloseClipboard(); }
                return new Snapshot(items);
            }
            catch { return null; }
        }

        internal void Restore()
        {
            try
            {
                _owner ??= new OwnerWindow();
                if (!TryOpen(_owner.Handle)) { Release(_items); return; }
                try
                {
                    EmptyClipboard();
                    foreach (var (format, bytes, metafile) in _items)
                    {
                        if (metafile != IntPtr.Zero)
                        {
                            if (SetClipboardData(format, metafile) == IntPtr.Zero) DeleteEnhMetaFile(metafile);
                            continue;
                        }
                        var memory = GlobalAlloc(0x0002 /* GMEM_MOVEABLE */, (UIntPtr)bytes!.Length);
                        if (memory == IntPtr.Zero) continue;
                        var pointer = GlobalLock(memory);
                        if (pointer == IntPtr.Zero) { GlobalFree(memory); continue; }
                        Marshal.Copy(bytes, 0, pointer, bytes.Length);
                        GlobalUnlock(memory);
                        if (SetClipboardData(format, memory) == IntPtr.Zero) GlobalFree(memory);
                    }
                }
                finally { CloseClipboard(); }
            }
            catch { }
        }

        private static void Release(List<(uint Format, byte[]? Bytes, IntPtr Metafile)> items)
        {
            foreach (var item in items)
                if (item.Metafile != IntPtr.Zero) DeleteEnhMetaFile(item.Metafile);
            items.Clear();
        }

        // GDI-handle formats cannot be copied as memory, and Windows synthesizes
        // the text/bitmap variants again from the formats that are restored. OLE's
        // own markers point at the previous owner's data object, which is gone.
        private static bool IsSkipped(uint format, List<uint> formats)
        {
            if (format is CfBitmap or CfMetafilePict or CfPalette) return true;
            if (format is CfText or CfOemText or CfLocale && formats.Contains(CfUnicodeText)) return true;
            if (format == CfDibV5 && formats.Contains(CfDib)) return true;
            if (format is >= 0x0080 and <= 0x008F) return true; // display/owner-display formats
            if (format is >= 0x0200 and <= 0x03FF) return true; // private and GDI-object ranges
            if (format < 0xC000) return false;
            var name = new StringBuilder(128);
            if (GetClipboardFormatName(format, name, name.Capacity) == 0) return false;
            var text = name.ToString();
            return text is "DataObject" or "Ole Private Data" or "OLE Clipboard Persist On Flush";
        }
    }

    // A clipboard owner is required: SetClipboardData fails after EmptyClipboard
    // when the clipboard was opened without a window.
    private sealed class OwnerWindow : System.Windows.Forms.NativeWindow
    {
        internal OwnerWindow()
        {
            CreateHandle(new System.Windows.Forms.CreateParams
            {
                Caption = "VisualTeX PowerPoint Clipboard",
                Parent = new IntPtr(-3), // HWND_MESSAGE
            });
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMessage
    {
        public IntPtr Hwnd;
        public uint Message;
        public UIntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    private static extern uint MsgWaitForMultipleObjectsEx(uint count, IntPtr handles, uint milliseconds, uint wakeMask, uint flags);
    [DllImport("user32.dll")]
    private static extern bool PeekMessage(out NativeMessage message, IntPtr window, uint filterMin, uint filterMax, uint remove);
    [DllImport("user32.dll")]
    private static extern IntPtr GetOpenClipboardWindow();
    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool OpenClipboard(IntPtr owner);
    [DllImport("user32.dll")]
    private static extern bool CloseClipboard();
    [DllImport("user32.dll")]
    private static extern bool EmptyClipboard();
    [DllImport("user32.dll")]
    private static extern uint EnumClipboardFormats(uint format);
    [DllImport("user32.dll")]
    private static extern IntPtr GetClipboardData(uint format);
    [DllImport("user32.dll")]
    private static extern IntPtr SetClipboardData(uint format, IntPtr memory);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClipboardFormatName(uint format, StringBuilder name, int capacity);
    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalAlloc(uint flags, UIntPtr bytes);
    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalFree(IntPtr memory);
    [DllImport("kernel32.dll")]
    private static extern IntPtr GlobalLock(IntPtr memory);
    [DllImport("kernel32.dll")]
    private static extern bool GlobalUnlock(IntPtr memory);
    [DllImport("kernel32.dll")]
    private static extern UIntPtr GlobalSize(IntPtr memory);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CopyEnhMetaFile(IntPtr metafile, string? fileName);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteEnhMetaFile(IntPtr metafile);
}

using System.Runtime.InteropServices;
using VisualTeX.WordVsto;

namespace VisualTeX.PowerPointVsto;

internal static class PowerPointMathTypePreview
{
    internal static string FromNativeWmf(string wmfPath, float widthPt, float heightPt)
    {
        var bytes = File.ReadAllBytes(wmfPath);
        // MathPage stores its logical coordinate system only in the Aldus
        // placeable header, not in SETWINDOWORG/SETWINDOWEXT records. GDI's
        // SetWinMetaFileBits ignores that header; preserve the mapping explicitly
        // or glyphs are rendered outside the OLE frame.
        if (bytes.Length > 22 && BitConverter.ToUInt32(bytes, 0) == 0x9AC6CDD7)
        {
            var left = BitConverter.ToInt16(bytes, 6);
            var top = BitConverter.ToInt16(bytes, 8);
            var width = BitConverter.ToInt16(bytes, 10) - left;
            var height = BitConverter.ToInt16(bytes, 12) - top;
            if (width <= 0 || height <= 0 || width > short.MaxValue || height > short.MaxValue)
                throw new InvalidDataException("Invalid MathType WMF coordinate bounds.");
            using var mapped = new MemoryStream();
            using (var writer = new BinaryWriter(mapped, System.Text.Encoding.UTF8, leaveOpen: true))
            {
                writer.Write(bytes, 22, 18);
                writer.Write(4); writer.Write((ushort)0x0103); writer.Write((short)8);
                writer.Write(5); writer.Write((ushort)0x020B); writer.Write(top); writer.Write(left);
                writer.Write(5); writer.Write((ushort)0x020C); writer.Write((short)height); writer.Write((short)width);
                writer.Write(bytes, 40, bytes.Length - 40);
                writer.Seek(6, SeekOrigin.Begin);
                writer.Write(checked((uint)mapped.Length / 2));
            }
            bytes = mapped.ToArray();
        }
        var picture = new MetafilePicture
        {
            MappingMode = 8,
            Width = (int)Math.Round(widthPt * 2540 / 72),
            Height = (int)Math.Round(heightPt * 2540 / 72),
        };
        var referenceDc = GetDC(IntPtr.Zero);
        IntPtr handle;
        try { handle = SetWinMetaFileBits((uint)bytes.Length, bytes, referenceDc, ref picture); }
        finally { if (referenceDc != IntPtr.Zero) ReleaseDC(IntPtr.Zero, referenceDc); }
        if (handle == IntPtr.Zero) throw new InvalidDataException("MathType's vector preview could not be loaded.");
        var path = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-mathtype-{Guid.NewGuid():N}.emf");
        try
        {
            var copy = CopyEnhMetaFile(handle, path);
            if (copy == IntPtr.Zero) throw new InvalidDataException("MathType's vector preview could not be saved.");
            DeleteEnhMetaFile(copy);
            return path;
        }
        finally { DeleteEnhMetaFile(handle); }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MetafilePicture
    {
        internal int MappingMode;
        internal int Width;
        internal int Height;
        internal IntPtr Metafile;
    }

    [DllImport("gdi32.dll")]
    private static extern IntPtr SetWinMetaFileBits(uint size, byte[] bytes, IntPtr referenceDc, ref MetafilePicture picture);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CopyEnhMetaFile(IntPtr handle, string path);
    [DllImport("gdi32.dll")]
    private static extern bool DeleteEnhMetaFile(IntPtr handle);
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr window);
    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr window, IntPtr dc);
}

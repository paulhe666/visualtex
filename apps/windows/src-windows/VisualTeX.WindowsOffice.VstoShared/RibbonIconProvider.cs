using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class RibbonIconProvider
{
    private static readonly object Gate = new();
    private static readonly Dictionary<string, IconResource> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    internal static object? GetImage(string? key)
    {
        if (key is null) return null;
        var requiredKey = key.Trim();
        if (requiredKey.Length == 0) return null;
        if (!RibbonVectorIconRenderer.Keys.Contains(requiredKey)) return null;

        lock (Gate)
        {
            if (Cache.TryGetValue(requiredKey, out var cached))
                return cached.PictureDisp;

            var bitmap = RibbonVectorIconRenderer.Create(requiredKey);
            var pictureDisp = (OfficeHostEnvironment.IsWps ? IconPicture.TryCreate(bitmap) : null)
                ?? PictureDispHost.ToPictureDisp(bitmap);
            Cache[requiredKey] = new IconResource(bitmap, pictureDisp);
            return pictureDisp;
        }
    }

    /// <summary>
    /// WPS draws a ribbon picture's bitmap without its alpha channel, so the transparent area
    /// of a GetHbitmap bitmap (composited onto light grey) shows as a grey square. An icon
    /// picture keeps its alpha in WPS.
    /// </summary>
    private static class IconPicture
    {
        private const int PictureTypeIcon = 3;
        private static Guid pictureDispId = new("7BF80981-BF32-101A-8BBB-00AA00300CAB");

        // PICTDESC: the union's largest member (the metafile's handle and extents) sets its size.
        [StructLayout(LayoutKind.Sequential)]
        private struct PictureDescription
        {
            internal int Size;
            internal int Type;
            internal IntPtr Handle;
            internal int ExtentX;
            internal int ExtentY;
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void OleCreatePictureIndirect(ref PictureDescription description,
            ref Guid interfaceId, bool own, [MarshalAs(UnmanagedType.IUnknown)] out object picture);

        [DllImport("user32.dll")]
        private static extern bool DestroyIcon(IntPtr icon);

        internal static object? TryCreate(Bitmap bitmap)
        {
            var icon = IntPtr.Zero;
            try
            {
                icon = bitmap.GetHicon();
                var description = new PictureDescription
                {
                    Size = Marshal.SizeOf<PictureDescription>(),
                    Type = PictureTypeIcon,
                    Handle = icon,
                };
                // The picture owns the icon from here on.
                OleCreatePictureIndirect(ref description, ref pictureDispId, true, out var picture);
                return picture;
            }
            catch
            {
                if (icon != IntPtr.Zero) DestroyIcon(icon);
                return null;
            }
        }
    }

    private sealed class IconResource
    {
        internal IconResource(Bitmap bitmap, object pictureDisp)
        {
            Bitmap = bitmap;
            PictureDisp = pictureDisp;
        }

        // Office's IPictureDisp wrapper continues to reference the GDI image.
        internal Bitmap Bitmap { get; }
        internal object PictureDisp { get; }
    }

    private sealed class PictureDispHost : AxHost
    {
        private PictureDispHost() : base(string.Empty) { }

        internal static object ToPictureDisp(Image image) =>
            GetIPictureDispFromPicture(image);
    }
}

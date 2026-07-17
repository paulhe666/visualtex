using System.Drawing;
using System.Windows.Forms;

namespace VisualTeX.WindowsOffice.VstoShared;

internal static class RibbonIconProvider
{
    private static readonly object Gate = new();
    private static readonly Dictionary<string, IconResource> Cache =
        new(StringComparer.OrdinalIgnoreCase);

    private static readonly IReadOnlyDictionary<string, string> IconBase64 =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["oleDisplay"] = RibbonIconData.OleDisplay,
            ["ommlDisplay"] = RibbonIconData.OmmlDisplay,
            ["oleInline"] = RibbonIconData.OleInline,
            ["ommlInline"] = RibbonIconData.OmmlInline,
            ["insertFormula"] = RibbonIconData.InsertFormula,
            ["updateNumbers"] = RibbonIconData.UpdateNumbers,
            ["editSelected"] = RibbonIconData.EditSelected,
            ["convertToOmml"] = RibbonIconData.ConvertToOmml,
            ["convertToOle"] = RibbonIconData.ConvertToOle,
        };

    internal static object? GetImage(string? key)
    {
        if (key is null) return null;
        var requiredKey = key.Trim();
        if (requiredKey.Length == 0) return null;
        if (!IconBase64.TryGetValue(requiredKey, out var encoded)) return null;

        lock (Gate)
        {
            if (Cache.TryGetValue(requiredKey, out var cached))
                return cached.PictureDisp;

            var bytes = Convert.FromBase64String(encoded);
            using var stream = new MemoryStream(bytes, writable: false);
            using var source = Image.FromStream(stream, false, true);
            var bitmap = new Bitmap(source);
            var pictureDisp = PictureDispHost.ToPictureDisp(bitmap);
            Cache[requiredKey] = new IconResource(bitmap, pictureDisp);
            return pictureDisp;
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

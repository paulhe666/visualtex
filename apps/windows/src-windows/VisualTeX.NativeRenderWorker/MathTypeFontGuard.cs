using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;

namespace VisualTeX.NativeRenderWorker;

/// <summary>
/// MathType draws many symbols from its own fonts (MT Extra and the Euclid
/// fonts). When this session cannot resolve them, or resolves "MT Extra" to
/// Office's 1992 copy (46 glyphs, installed by Office under the same name),
/// MathType still renders but writes placeholders and substitute glyphs into
/// the picture, and that picture is then saved in the document.
///
/// Before MathType starts, make this session resolve every MathType font to
/// the exact file bundled with the private runtime. Fonts are added only to
/// this service's own session (session 0), never installed system-wide; a
/// different copy of a MathType family is unloaded only from this session.
/// If a font still cannot be resolved, rendering is refused.
/// </summary>
internal static class MathTypeFontGuard
{
    internal const string FontsUnavailable = "MathType fonts are unavailable to the private renderer.";

    internal sealed class BundledFont
    {
        public string Path = string.Empty;
        public string Family = string.Empty;
        public int Weight;
        public bool Italic;
        public byte[] Head = Array.Empty<byte>();
    }

    internal static bool Ensure(string runtimeRoot, out string error)
    {
        error = string.Empty;
        var directory = System.IO.Path.Combine(runtimeRoot, "Fonts", "TrueType");
        if (!Directory.Exists(directory))
        {
            error = FontsUnavailable + " The bundled font directory is missing.";
            return false;
        }
        var bundled = new List<BundledFont>();
        foreach (var file in Directory.GetFiles(directory, "*.ttf").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            if (!TryReadFont(file, out var font))
            {
                error = FontsUnavailable + $" {System.IO.Path.GetFileName(file)} is not a readable TrueType font.";
                return false;
            }
            bundled.Add(font!);
        }
        if (bundled.Count == 0)
        {
            error = FontsUnavailable + " No bundled fonts were found.";
            return false;
        }

        foreach (var font in bundled)
        {
            if (Resolves(font)) continue;
            // Another copy of this family is loaded first (for example Office's
            // 1992 MT Extra) and would keep winning: unload it from this session.
            foreach (var other in LoadedCopies(font))
            {
                for (var i = 0; i < 16 && RemoveFontResourceW(other); i++) { }
            }
            if (!Resolves(font) && AddFontResourceW(font.Path) == 0 && IsRequired(font))
            {
                error = FontsUnavailable + $" Windows did not load {System.IO.Path.GetFileName(font.Path)}.";
                return false;
            }
        }
        foreach (var font in bundled.Where(IsRequired))
        {
            if (!Resolves(font))
            {
                error = FontsUnavailable + $" {font.Family} ({System.IO.Path.GetFileName(font.Path)}) does not resolve to MathType's font.";
                return false;
            }
        }
        return true;
    }

    // Fonts MathType's default styles draw from (and the only MathType fonts in
    // VisualTeX's MTEF). Fences and the Tiger fonts serve optional MathType
    // preference sets: they are loaded when possible but never block rendering.
    internal static bool IsRequired(BundledFont font) =>
        string.Equals(font.Family, "MT Extra", StringComparison.OrdinalIgnoreCase)
        || font.Family.StartsWith("Euclid", StringComparison.OrdinalIgnoreCase);

    // True when GDI selects exactly this file for its family and style.
    internal static bool Resolves(BundledFont font)
    {
        var dc = CreateCompatibleDC(IntPtr.Zero);
        if (dc == IntPtr.Zero) return false;
        var handle = CreateFontW(-1000, 0, 0, 0, font.Weight, font.Italic ? 1u : 0u, 0, 0,
            1 /* DEFAULT_CHARSET, as MathType's pictures request */, 7 /* OUT_TT_ONLY_PRECIS */, 0, 0, 0, font.Family);
        if (handle == IntPtr.Zero) { DeleteDC(dc); return false; }
        var old = SelectObject(dc, handle);
        try
        {
            var face = new char[64];
            var length = GetTextFaceW(dc, face.Length, face);
            var actual = length > 0 ? new string(face, 0, length - 1).TrimEnd('\0') : string.Empty;
            if (!string.Equals(actual, font.Family, StringComparison.OrdinalIgnoreCase)) return false;
            var size = GetFontData(dc, HeadTag, 0, null, 0);
            if (size == GdiError || size != font.Head.Length) return false;
            var head = new byte[size];
            if (GetFontData(dc, HeadTag, 0, head, size) != size) return false;
            return head.AsSpan().SequenceEqual(font.Head);
        }
        finally
        {
            SelectObject(dc, old);
            DeleteObject(handle);
            DeleteDC(dc);
        }
    }

    // Other copies of this family and style that this session may have loaded:
    // registered files, plus files in the Fonts folder that look like this font.
    // A copy whose registration the installer replaced (for example Office's
    // MTEXTRA.TTF) stays on disk and stays loaded here until Windows restarts.
    private static IEnumerable<string> LoadedCopies(BundledFont font)
    {
        var fontsDirectory = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        var stem = NormalizedStem(font.Path);
        var candidates = new List<string>();
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
        {
            try
            {
                using var root = RegistryKey.OpenBaseKey(hive, RegistryView.Registry64);
                using var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts");
                if (key is null) continue;
                foreach (var name in key.GetValueNames())
                {
                    if (key.GetValue(name) is not string value || value.Length == 0) continue;
                    var path = System.IO.Path.IsPathRooted(value) ? value : System.IO.Path.Combine(fontsDirectory, value);
                    // Only entries that can plausibly be this family; never
                    // touch unrelated fonts.
                    if (name.IndexOf(font.Family, StringComparison.OrdinalIgnoreCase) >= 0
                        || string.Equals(NormalizedStem(path), stem, StringComparison.OrdinalIgnoreCase))
                        candidates.Add(path);
                }
            }
            catch (System.Security.SecurityException) { }
            catch (UnauthorizedAccessException) { }
        }
        try
        {
            foreach (var path in Directory.EnumerateFiles(fontsDirectory))
            {
                if (string.Equals(NormalizedStem(path), stem, StringComparison.OrdinalIgnoreCase)) candidates.Add(path);
            }
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in candidates)
        {
            if (!seen.Add(path) || string.Equals(System.IO.Path.GetFullPath(path), System.IO.Path.GetFullPath(font.Path), StringComparison.OrdinalIgnoreCase))
                continue;
            if (!TryReadFont(path, out var other)) continue;
            if (!string.Equals(other!.Family, font.Family, StringComparison.OrdinalIgnoreCase)
                || other.Weight != font.Weight || other.Italic != font.Italic) continue;
            if (other.Head.AsSpan().SequenceEqual(font.Head)) continue;
            yield return path;
        }
    }

    // "mtextra_02.ttf", "MTExtra.ttf" and "mtextra_visualtex.ttf" are copies of
    // "mtextra.ttf".
    private static string NormalizedStem(string path)
    {
        var stem = System.IO.Path.GetFileNameWithoutExtension(path).Replace(" ", string.Empty).ToLowerInvariant();
        var underscore = stem.LastIndexOf('_');
        if (underscore > 0 && stem.Substring(underscore + 1).All(char.IsDigit)) stem = stem.Substring(0, underscore);
        if (stem.EndsWith("_visualtex", StringComparison.Ordinal)) stem = stem.Substring(0, stem.Length - "_visualtex".Length);
        return stem;
    }

    // Family (name ID 1), weight and italic from the font file itself, plus its
    // 'head' table, which identifies the exact font version.
    internal static bool TryReadFont(string path, out BundledFont? font)
    {
        font = null;
        try
        {
            var bytes = File.ReadAllBytes(path);
            if (bytes.Length < 12) return false;
            int tables = ReadUInt16(bytes, 4);
            int name = -1, nameLength = 0, head = -1, headLength = 0, os2 = -1;
            for (var i = 0; i < tables; i++)
            {
                var record = 12 + i * 16;
                if (record + 16 > bytes.Length) return false;
                var tag = Encoding.ASCII.GetString(bytes, record, 4);
                var offset = (int)ReadUInt32(bytes, record + 8);
                var length = (int)ReadUInt32(bytes, record + 12);
                if (offset < 0 || length < 0 || offset + length > bytes.Length) return false;
                if (tag == "name") { name = offset; nameLength = length; }
                else if (tag == "head") { head = offset; headLength = length; }
                else if (tag == "OS/2") os2 = offset;
            }
            if (name < 0 || head < 0 || headLength < 54) return false;
            var family = ReadFamily(bytes, name, nameLength);
            if (string.IsNullOrEmpty(family)) return false;
            var macStyle = ReadUInt16(bytes, head + 44);
            var weight = os2 >= 0 && os2 + 6 <= bytes.Length ? ReadUInt16(bytes, os2 + 4) : 0;
            if (weight is < 100 or > 1000) weight = (macStyle & 1) != 0 ? 700 : 400;
            font = new BundledFont
            {
                Path = System.IO.Path.GetFullPath(path),
                Family = family!,
                Weight = (macStyle & 1) != 0 ? Math.Max(weight, 600) : Math.Min(weight, 500),
                Italic = (macStyle & 2) != 0,
                Head = bytes.AsSpan(head, headLength).ToArray(),
            };
            return true;
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    private static string? ReadFamily(byte[] bytes, int table, int length)
    {
        if (length < 6) return null;
        int count = ReadUInt16(bytes, table + 2);
        int strings = table + ReadUInt16(bytes, table + 4);
        string? windows = null, mac = null;
        for (var i = 0; i < count; i++)
        {
            var record = table + 6 + i * 12;
            if (record + 12 > table + length) break;
            int platform = ReadUInt16(bytes, record), encoding = ReadUInt16(bytes, record + 2);
            int language = ReadUInt16(bytes, record + 4), nameId = ReadUInt16(bytes, record + 6);
            int size = ReadUInt16(bytes, record + 8), offset = ReadUInt16(bytes, record + 10);
            if (nameId != 1 || strings + offset + size > bytes.Length) continue;
            if (platform == 3 && (encoding == 0 || encoding == 1))
            {
                var value = Encoding.BigEndianUnicode.GetString(bytes, strings + offset, size);
                if (windows is null || language == 0x409) windows = value;
            }
            else if (platform == 1 && encoding == 0 && mac is null)
            {
                mac = Encoding.Latin1.GetString(bytes, strings + offset, size);
            }
        }
        return (windows ?? mac)?.Trim();
    }

    private static int ReadUInt16(byte[] bytes, int offset) => bytes[offset] << 8 | bytes[offset + 1];
    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3]);

    private const uint HeadTag = 0x64616568; // 'head'
    private const uint GdiError = 0xFFFFFFFF;

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern int AddFontResourceW(string path);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern bool RemoveFontResourceW(string path);
    [DllImport("gdi32.dll")] private static extern IntPtr CreateCompatibleDC(IntPtr dc);
    [DllImport("gdi32.dll")] private static extern bool DeleteDC(IntPtr dc);
    [DllImport("gdi32.dll")] private static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
    [DllImport("gdi32.dll")] private static extern bool DeleteObject(IntPtr value);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateFontW(int height, int width, int escapement, int orientation, int weight,
        uint italic, uint underline, uint strikeOut, uint charSet, uint outPrecision, uint clipPrecision,
        uint quality, uint pitchAndFamily, string face);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] private static extern int GetTextFaceW(IntPtr dc, int count, [Out] char[] face);
    [DllImport("gdi32.dll")] private static extern uint GetFontData(IntPtr dc, uint table, uint offset, byte[]? buffer, uint length);
}

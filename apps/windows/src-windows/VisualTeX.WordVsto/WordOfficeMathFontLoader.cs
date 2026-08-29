using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace VisualTeX.WordVsto;

internal static class WordOfficeMathFontLoader
{
    internal const string LatinModernMathFamily = "Latin Modern Math";
    internal const string LatinModernMathFileName = "latinmodern-math.otf";
    internal const string LatinModernMathSha256 =
        "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F";

    private const uint WmFontChange = 0x001D;
    private static readonly IntPtr HwndBroadcast = new(0xffff);
    private static readonly object Gate = new();
    private static bool _verified;
    private static bool _sessionRegistrationUsed;
    private static string _verifiedPath = string.Empty;

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int AddFontResourceEx(
        string fileName,
        uint flags,
        IntPtr reserved);

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RemoveFontResourceEx(
        string fileName,
        uint flags,
        IntPtr reserved);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(
        IntPtr window,
        uint message,
        IntPtr wParam,
        IntPtr lParam,
        uint flags,
        uint timeoutMilliseconds,
        out IntPtr result);

    internal static bool IsLoaded
    {
        get
        {
            lock (Gate) return _verified;
        }
    }

    internal static bool SessionRegistrationUsed
    {
        get
        {
            lock (Gate) return _sessionRegistrationUsed;
        }
    }

    internal static string LoadedPath
    {
        get
        {
            lock (Gate) return _verifiedPath;
        }
    }

    internal static void EnsureLoaded()
    {
        lock (Gate)
        {
            if (_verified) return;

            var fontPath = ResolveInstalledFontPath();
            if (string.IsNullOrWhiteSpace(fontPath))
            {
                throw new FileNotFoundException(
                    "Latin Modern Math is not installed by the VisualTeX Office integration. Reinstall the latest Office integration and restart Word.",
                    LatinModernMathFileName);
            }

            var actualHash = ComputeSha256(fontPath);
            if (!string.Equals(
                    actualHash,
                    LatinModernMathSha256,
                    StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException(
                    "The installed Latin Modern Math font failed VisualTeX's SHA-256 integrity check. Reinstall the latest Office integration.");
            }

            var globallyResolvable = CanResolveLatinModernMath();
            if (!globallyResolvable)
            {
                // The MSI installs Latin Modern Math globally, but acceptance runs
                // and an in-place Office upgrade can load the new VSTO payload before
                // Windows refreshes the process font cache. Register the exact,
                // hash-verified file in the current Windows session as a deterministic
                // fallback. Do not use FR_PRIVATE: Word's Office Math/PDF subsystem
                // does not consistently consult process-private fonts even though
                // ordinary GDI text can use them. System.Drawing also keeps a stale
                // process-wide family cache, so AddFontResourceEx's return value is
                // the authoritative result for this session-registration path.
                var added = AddFontResourceEx(fontPath, 0, IntPtr.Zero);
                if (added <= 0)
                {
                    throw new InvalidOperationException(
                        "Word could not load VisualTeX's verified Latin Modern Math font into the current process.");
                }
                _sessionRegistrationUsed = true;

                try
                {
                    _ = SendMessageTimeout(
                        HwndBroadcast,
                        WmFontChange,
                        IntPtr.Zero,
                        IntPtr.Zero,
                        flags: 0x0002,
                        timeoutMilliseconds: 500,
                        out _);
                }
                catch
                {
                    // AddFontResourceEx already made the family available to this
                    // process. A blocked desktop broadcast is not a fatal error.
                }
            }

            _verified = true;
            _verifiedPath = fontPath;
        }
    }

    internal static bool TryEnsureLoaded(out string error)
    {
        try
        {
            EnsureLoaded();
            error = string.Empty;
            return true;
        }
        catch (Exception exception)
        {
            error = exception.Message;
            return false;
        }
    }

    internal static void UnloadSessionRegistration()
    {
        lock (Gate)
        {
            if (!_sessionRegistrationUsed || string.IsNullOrWhiteSpace(_verifiedPath))
                return;
            try { _ = RemoveFontResourceEx(_verifiedPath, 0, IntPtr.Zero); }
            catch { }
            _sessionRegistrationUsed = false;
            _verified = false;
            _verifiedPath = string.Empty;
            try
            {
                _ = SendMessageTimeout(
                    HwndBroadcast,
                    WmFontChange,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    flags: 0x0002,
                    timeoutMilliseconds: 250,
                    out _);
            }
            catch { }
        }
    }

    private static string ResolveInstalledFontPath()
    {
        var assemblyDirectory = Path.GetDirectoryName(
            typeof(WordOfficeMathFontLoader).Assembly.Location)
            ?? string.Empty;
        var candidates = new[]
        {
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                LatinModernMathFileName),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft",
                "Windows",
                "Fonts",
                LatinModernMathFileName),
            Path.Combine(assemblyDirectory, LatinModernMathFileName),
        };
        foreach (var candidate in candidates)
        {
            if (!string.IsNullOrWhiteSpace(candidate)
                && File.Exists(candidate))
                return candidate;
        }
        return string.Empty;
    }

    private static bool CanResolveLatinModernMath()
    {
        try
        {
            using var font = new Font(
                LatinModernMathFamily,
                12f,
                FontStyle.Regular,
                GraphicsUnit.Point);
            return string.Equals(
                font.Name,
                LatinModernMathFamily,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string ComputeSha256(string path)
    {
        using var stream = File.OpenRead(path);
        using var sha256 = SHA256.Create();
        return BitConverter.ToString(sha256.ComputeHash(stream)).Replace("-", string.Empty);
    }
}

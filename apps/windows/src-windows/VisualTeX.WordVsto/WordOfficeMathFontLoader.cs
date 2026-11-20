using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace VisualTeX.WordVsto;

/// <summary>
/// A math font the VisualTeX Office integration installs for Word equations
/// (office-fonts/README-*.txt). All of them are TrueType builds: Word's PDF export
/// embeds TrueType math; CFF math it rasterizes, dropping the parentheses of native
/// equation numbers.
/// </summary>
internal sealed class BundledMathFont
{
    internal BundledMathFont(string family, string fileName, string sha256, bool acceptsInstalledFamily)
    {
        Family = family;
        FileName = fileName;
        Sha256 = sha256;
        AcceptsInstalledFamily = acceptsInstalledFamily;
    }

    internal string Family { get; }
    internal string FileName { get; }
    internal string Sha256 { get; }

    /// <summary>
    /// The font keeps its published name (STIX Two Math), so a copy the user installed
    /// stands in for the bundled file; the renamed VisualTeX builds have no other source.
    /// </summary>
    internal bool AcceptsInstalledFamily { get; }
}

internal static class WordOfficeMathFontLoader
{
    // VisualTeX Latin Modern Math: Latin Modern Math 1.959 with its outlines
    // converted to TrueType (office-fonts/README-VisualTeX-Latin-Modern-Math.txt).
    // The original CFF font stays installed for documents that still name
    // "Latin Modern Math".
    internal const string LatinModernMathFamily = "VisualTeX Latin Modern Math";
    internal const string LatinModernMathFileName = "VisualTeXLatinModernMath.ttf";
    internal const string LatinModernMathSha256 =
        "BB3646F9296C10C9B5CCD95BF5680297A4E35B3BC25EA5F4F7593A9C405E0C93";
    // The CFF original that VisualTeX before 1.2.8 set as the document math font.
    internal const string LegacyLatinModernMathFamily = "Latin Modern Math";

    // VisualTeX TeX Gyre Termes Math: TeX Gyre Termes Math 1.543 (Times-style) with its
    // outlines converted to TrueType (office-fonts/README-VisualTeX-TeX-Gyre-Termes-Math.txt).
    internal const string TexGyreTermesMathFamily = "VisualTeX TeX Gyre Termes Math";
    internal const string TexGyreTermesMathFileName = "VisualTeXTeXGyreTermesMath.ttf";
    internal const string TexGyreTermesMathSha256 =
        "9F58AD9C3C92D18BB439FF31535FBA7A8038624E5E5E7605E2681FFDB9C8B1EE";

    // STIX Two Math 2.13 b171: the TrueType build the STIX Fonts project publishes,
    // unmodified (office-fonts/README-STIX-Two-Math.txt).
    internal const string StixTwoMathFamily = "STIX Two Math";
    internal const string StixTwoMathFileName = "STIXTwoMath-Regular.ttf";
    internal const string StixTwoMathSha256 =
        "49B22767985171BCE579307B21D98F4D35012FA39E316ACC4FFCAB8644DFB69B";

    internal static readonly BundledMathFont LatinModernMath =
        new(LatinModernMathFamily, LatinModernMathFileName, LatinModernMathSha256, acceptsInstalledFamily: false);
    internal static readonly BundledMathFont TexGyreTermesMath =
        new(TexGyreTermesMathFamily, TexGyreTermesMathFileName, TexGyreTermesMathSha256, acceptsInstalledFamily: false);
    internal static readonly BundledMathFont StixTwoMath =
        new(StixTwoMathFamily, StixTwoMathFileName, StixTwoMathSha256, acceptsInstalledFamily: true);

    internal static IReadOnlyList<BundledMathFont> All { get; } =
        new[] { LatinModernMath, TexGyreTermesMath, StixTwoMath };

    /// <summary>The bundled font of a family, or null for any other font.</summary>
    internal static BundledMathFont? Find(string? family)
    {
        var name = family?.Trim();
        return string.IsNullOrEmpty(name)
            ? null
            : All.FirstOrDefault(font => string.Equals(font.Family, name, StringComparison.OrdinalIgnoreCase));
    }

    private sealed class LoadState
    {
        internal bool Verified { get; set; }
        internal bool SessionRegistrationUsed { get; set; }
        internal string Path { get; set; } = string.Empty;
    }

    private const uint WmFontChange = 0x001D;
    private static readonly IntPtr HwndBroadcast = new(0xffff);
    private static readonly object Gate = new();
    private static readonly Dictionary<BundledMathFont, LoadState> States = new();

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

    private static LoadState State(BundledMathFont font)
    {
        if (!States.TryGetValue(font, out var state))
            States[font] = state = new LoadState();
        return state;
    }

    // Latin Modern Math, the font the acceptance runs and older callers check.
    internal static bool IsLoaded
    {
        get
        {
            lock (Gate) return State(LatinModernMath).Verified;
        }
    }

    internal static bool SessionRegistrationUsed
    {
        get
        {
            lock (Gate) return State(LatinModernMath).SessionRegistrationUsed;
        }
    }

    internal static string LoadedPath
    {
        get
        {
            lock (Gate) return State(LatinModernMath).Path;
        }
    }

    internal static void EnsureLoaded() => EnsureLoaded(LatinModernMath);

    internal static bool TryEnsureLoaded(out string error) => TryEnsureLoaded(LatinModernMath, out error);

    internal static void EnsureLoaded(BundledMathFont font)
    {
        lock (Gate)
        {
            var state = State(font);
            if (state.Verified) return;

            var fontPath = ResolveInstalledFontPath(font.FileName);
            if (string.IsNullOrWhiteSpace(fontPath))
            {
                if (font.AcceptsInstalledFamily && CanResolve(font.Family))
                {
                    state.Verified = true;
                    return;
                }
                throw new FileNotFoundException(
                    $"{font.Family} is not installed by the VisualTeX Office integration. Reinstall the latest Office integration and restart Word.",
                    font.FileName);
            }

            var actualHash = ComputeSha256(fontPath);
            if (!string.Equals(actualHash, font.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                // Another release of a published font under the same file name is still that font.
                if (font.AcceptsInstalledFamily && CanResolve(font.Family))
                {
                    state.Verified = true;
                    return;
                }
                throw new InvalidDataException(
                    $"The installed {font.Family} font failed VisualTeX's SHA-256 integrity check. Reinstall the latest Office integration.");
            }

            if (!CanResolve(font.Family))
            {
                // The MSI installs the bundled fonts globally, but acceptance runs
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
                        $"Word could not load VisualTeX's verified {font.Family} font into the current process.");
                }
                state.SessionRegistrationUsed = true;
                BroadcastFontChange(500);
            }

            state.Verified = true;
            state.Path = fontPath;
        }
    }

    internal static bool TryEnsureLoaded(BundledMathFont font, out string error)
    {
        try
        {
            EnsureLoaded(font);
            error = string.Empty;
            return true;
        }
        catch (Exception exception)
        {
            error = exception.Message;
            return false;
        }
    }

    /// <summary>Undoes the session registrations this process made (add-in shutdown).</summary>
    internal static void UnloadSessionRegistration()
    {
        lock (Gate)
        {
            var removed = false;
            foreach (var state in States.Values)
            {
                if (!state.SessionRegistrationUsed || string.IsNullOrWhiteSpace(state.Path)) continue;
                try { _ = RemoveFontResourceEx(state.Path, 0, IntPtr.Zero); }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                state.SessionRegistrationUsed = false;
                state.Verified = false;
                state.Path = string.Empty;
                removed = true;
            }
            if (removed) BroadcastFontChange(250);
        }
    }

    private static void BroadcastFontChange(uint timeoutMilliseconds)
    {
        try
        {
            _ = SendMessageTimeout(
                HwndBroadcast,
                WmFontChange,
                IntPtr.Zero,
                IntPtr.Zero,
                flags: 0x0002,
                timeoutMilliseconds: timeoutMilliseconds,
                out _);
        }
        catch (Exception ignored)
        {
            // AddFontResourceEx already made the family available to this
            // process. A blocked desktop broadcast is not a fatal error.
            IgnoredErrors.Ignore(ignored);
        }
    }

    private static string ResolveInstalledFontPath(string fileName)
    {
        var assemblyDirectory = Path.GetDirectoryName(
            typeof(WordOfficeMathFontLoader).Assembly.Location)
            ?? string.Empty;
        var candidates = new[]
        {
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
                fileName),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Microsoft",
                "Windows",
                "Fonts",
                fileName),
            Path.Combine(assemblyDirectory, fileName),
        };
        foreach (var candidate in candidates)
        {
            if (!string.IsNullOrWhiteSpace(candidate)
                && File.Exists(candidate))
                return candidate;
        }
        return string.Empty;
    }

    private static bool CanResolve(string family)
    {
        try
        {
            using var font = new Font(
                family,
                12f,
                FontStyle.Regular,
                GraphicsUnit.Point);
            return string.Equals(
                font.Name,
                family,
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

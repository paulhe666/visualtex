using Microsoft.Win32;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

/// <summary>
/// The per-user OMML Font choice (ribbon menu, HKCU) and whether its font is usable on
/// this PC. Installed into <see cref="WordOmmlMathFontPolicy"/> at add-in startup.
/// </summary>
internal static class WordOmmlMathFontSetting
{
    private const string RegistryPath = @"Software\VisualTeX\WordVsto";
    private const string ValueName = "OmmlMathFont";

    private static readonly object Gate = new();
    private static readonly Dictionary<string, bool> Availability = new(StringComparer.OrdinalIgnoreCase);
    private static OmmlMathFontChoice? _choice;

    internal static OmmlMathFontChoice Choice
    {
        get
        {
            lock (Gate)
            {
                if (_choice.HasValue) return _choice.Value;
                var choice = OmmlMathFontChoice.FollowDocument;
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(RegistryPath, writable: false);
                    if (key?.GetValue(ValueName) is string stored)
                        choice = WordOmmlMathFontPolicy.Parse(stored);
                }
                catch { }
                _choice = choice;
                return choice;
            }
        }
    }

    internal static void Save(OmmlMathFontChoice choice)
    {
        lock (Gate)
        {
            _choice = choice;
            Availability.Clear();
        }
        try
        {
            using var key = Registry.CurrentUser.CreateSubKey(RegistryPath, writable: true);
            key?.SetValue(ValueName, WordOmmlMathFontPolicy.Id(choice), RegistryValueKind.String);
        }
        catch
        {
            // Keep the in-process choice even if the profile rejects persistence.
        }
    }

    /// <summary>Forget cached font availability (the user may have installed a font).</summary>
    internal static void RefreshAvailability()
    {
        lock (Gate) Availability.Clear();
    }

    internal static bool IsAvailable(OmmlMathFontChoice choice) =>
        WordOmmlMathFontPolicy.FontFamily(choice) is not { } family || IsAvailable(family);

    internal static bool IsAvailable(string family)
    {
        lock (Gate)
        {
            if (Availability.TryGetValue(family, out var known)) return known;
        }
        bool available;
        if (WordOfficeMathFontLoader.Find(family) is { } bundled)
        {
            // Bundled with the Office integration; loaded (and hash-checked) on demand.
            available = WordOfficeMathFontLoader.TryEnsureLoaded(bundled, out _);
        }
        else
        {
            try
            {
                // GDI+ substitutes a fallback family for a font that is not installed.
                using var font = new System.Drawing.Font(family, 12f);
                available = string.Equals(font.Name, family, StringComparison.OrdinalIgnoreCase);
            }
            catch { available = false; }
        }
        lock (Gate) Availability[family] = available;
        return available;
    }

    /// <summary>The choice the ribbon shows as selected: follow for a missing font.</summary>
    internal static OmmlMathFontChoice EffectiveChoice =>
        WordOmmlMathFontPolicy.EffectiveFontFamily(Choice, IsAvailable) is null
            ? OmmlMathFontChoice.FollowDocument
            : Choice;

    internal static string? EffectiveFont =>
        WordOmmlMathFontPolicy.EffectiveFontFamily(Choice, IsAvailable);

    internal static void Install() =>
        WordOmmlMathFontPolicy.Configure(
            () => EffectiveFont,
            (document, font) => WordFormulaService.ApplyDocumentOmmlMathFont(
                (Microsoft.Office.Interop.Word.Document)document, font));

    internal static void Uninstall() => WordOmmlMathFontPolicy.Configure(null, null);
}

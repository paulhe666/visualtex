using System;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// The fonts chosen for letters, Greek letters and digits over the math font (VisualTeX Settings); an empty name
/// follows the math font. A formula records the ones it was drawn in (<see cref="FormulaMetadata.FormulaFaces"/>);
/// a formula without the record was drawn without them.
/// </summary>
public sealed class FormulaFaces
{
    [JsonPropertyName("letter")]
    public string Letter { get; set; } = string.Empty;

    [JsonPropertyName("greek")]
    public string Greek { get; set; } = string.Empty;

    [JsonPropertyName("digit")]
    public string Digit { get; set; } = string.Empty;

    internal void Validate()
    {
        foreach (var name in new[] { Letter, Greek, Digit })
            if ((name ?? string.Empty).Length > 128)
                throw new InvalidOperationException("A VisualTeX formula face name is too long.");
    }
}

/// <summary>
/// The fonts the main app draws formulas in now - the math font (KaTeX or newtxmath, stored as "times"), the Chinese
/// font and the faces - as its published editor preferences carry them, and whether a formula was drawn in others.
/// </summary>
public sealed class FormulaFontSettings
{
    public string LetterFont { get; set; } = "katex";
    public string ChineseFont { get; set; } = "system";
    public FormulaFaces Faces { get; set; } = new();

    /// <summary>
    /// Whether the formula's recorded fonts differ from these. A record a formula lacks counts as what VisualTeX
    /// drew without it: KaTeX, the system Chinese font, no faces. The Chinese font counts only for a formula with
    /// Chinese text, which is all it changes.
    /// </summary>
    public bool DifferFrom(FormulaMetadata metadata)
    {
        if (metadata is null) return false;
        var letter = string.IsNullOrWhiteSpace(metadata.FormulaLetterFont) ? "katex" : metadata.FormulaLetterFont!.Trim();
        if (!string.Equals(letter, LetterFont, StringComparison.OrdinalIgnoreCase)) return true;
        var faces = metadata.FormulaFaces ?? new FormulaFaces();
        if (!SameFace(faces.Letter, Faces.Letter) || !SameFace(faces.Greek, Faces.Greek) || !SameFace(faces.Digit, Faces.Digit))
            return true;
        var chinese = string.IsNullOrWhiteSpace(metadata.FormulaChineseFont) ? "system" : metadata.FormulaChineseFont!.Trim();
        return !string.Equals(chinese, ChineseFont, StringComparison.OrdinalIgnoreCase)
            && ContainsChinese(metadata.Latex);
    }

    /// <summary>The fonts in a few words, for confirmations: "KaTeX; letters, Greek letters, digits: Times New Roman".</summary>
    public string Describe(bool english)
    {
        var math = string.Equals(LetterFont, "times", StringComparison.OrdinalIgnoreCase) ? "newtxmath" : "KaTeX";
        var parts = new System.Collections.Generic.List<string>();
        void Add(string name, string zh, string en)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            parts.Add((english ? en : zh) + (english ? ": " : "：") + name.Trim());
        }
        Add(Faces.Letter, "字母", "letters");
        Add(Faces.Greek, "希腊字母", "Greek letters");
        Add(Faces.Digit, "数字", "digits");
        return parts.Count == 0 ? math : math + (english ? "; " : "；") + string.Join(english ? ", " : "，", parts);
    }

    /// <summary>
    /// Reads the fonts from the companion's preferences answer (<c>GET /api/v1/preferences</c>:
    /// <c>editorPreferences.settings.formulaLetterFont / formulaChineseFont / formulaFaces</c>); a font an earlier
    /// version does not send keeps the default. Null when the answer has no editor settings at all: the main window
    /// publishes them, so a VisualTeX started without it (by an Office add-in) does not know its fonts yet.
    /// </summary>
    public static FormulaFontSettings? FromPreferencesJson(string json)
    {
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 64 });
        if (!document.RootElement.TryGetProperty("editorPreferences", out var preferences)
            || preferences.ValueKind != JsonValueKind.Object
            || !preferences.TryGetProperty("settings", out var settings)
            || settings.ValueKind != JsonValueKind.Object)
            return null;
        var result = new FormulaFontSettings();
        if (settings.TryGetProperty("formulaLetterFont", out var letter) && letter.ValueKind == JsonValueKind.String)
            // fonts of earlier versions are drawn in KaTeX now (normalizeFormulaLetterFont)
            result.LetterFont = string.Equals(letter.GetString(), "times", StringComparison.OrdinalIgnoreCase) ? "times" : "katex";
        if (settings.TryGetProperty("formulaChineseFont", out var chinese) && chinese.ValueKind == JsonValueKind.String
            && !string.IsNullOrWhiteSpace(chinese.GetString()))
            result.ChineseFont = chinese.GetString()!.Trim();
        if (settings.TryGetProperty("formulaFaces", out var faces) && faces.ValueKind == JsonValueKind.Object)
        {
            string Name(string key) =>
                faces.TryGetProperty(key, out var value) && value.ValueKind == JsonValueKind.String
                    ? (value.GetString() ?? string.Empty).Trim()
                    : string.Empty;
            result.Faces = new FormulaFaces { Letter = Name("letter"), Greek = Name("greek"), Digit = Name("digit") };
        }
        return result;
    }

    private static bool SameFace(string? recorded, string? current) =>
        string.Equals((recorded ?? string.Empty).Trim(), (current ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase);

    private static bool ContainsChinese(string? latex)
    {
        if (string.IsNullOrEmpty(latex)) return false;
        foreach (var character in latex!)
            // CJK ideographs, compatibility ideographs, CJK punctuation and full-width forms
            if (character is >= '\u3400' and <= '\u9FFF' or >= '\uF900' and <= '\uFAFF'
                or >= '\u3000' and <= '\u303F' or >= '\uFF00' and <= '\uFFEF')
                return true;
        return false;
    }
}

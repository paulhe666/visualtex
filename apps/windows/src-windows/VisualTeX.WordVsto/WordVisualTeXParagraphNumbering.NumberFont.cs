namespace VisualTeX.WordVsto;

internal static partial class WordVisualTeXParagraphNumbering
{
    // The Word font for a number whose formula uses this letter font: the
    // first installed face of the editor's font stack for it. KaTeX and Cambria
    // formulas use Cambria, whose digits and brackets are Cambria Math's;
    // Cambria Math itself gives a text line its very tall math line height.
    // A formula from before letter fonts were recorded keeps the body font.
    internal static string? NumberFontName(string? letterFont) =>
        letterFont switch
        {
            "katex" or "cambria" => "Cambria",
            "times" => "Times New Roman",
            "stix" => FirstInstalledFont("STIX Two Text", "STIX Two Math") ?? "Times New Roman",
            "palatino" => FirstInstalledFont("Palatino", "Palatino Linotype", "Book Antiqua") ?? "Palatino Linotype",
            "helvetica" => FirstInstalledFont("Helvetica Neue", "Helvetica") ?? "Arial",
            _ => null,
        };

    private static HashSet<string>? installedFontNames;

    private static string? FirstInstalledFont(params string[] names)
    {
        if (installedFontNames is null)
        {
            using var fonts = new System.Drawing.Text.InstalledFontCollection();
            installedFontNames = new HashSet<string>(
                fonts.Families.Select(family => family.Name),
                StringComparer.OrdinalIgnoreCase);
        }
        return names.FirstOrDefault(installedFontNames.Contains);
    }
}

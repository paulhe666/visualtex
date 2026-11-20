namespace VisualTeX.WordVsto;

internal static partial class WordVisualTeXParagraphNumbering
{
    private const string PlaceRefMarker = "MACROBUTTON VisualTeXPlaceRef";

    internal static bool IsPlaceRefCode(
        string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            PlaceRefMarker,
            StringComparison.OrdinalIgnoreCase) >= 0;
}

using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

internal static partial class WordEquationNumberFormat
{
    static partial void InferDocumentEquationNumberFormat(
        Document document,
        ref string? formatId)
    {
        formatId = WordEquationNumberFormatAuthority.TryInferFromNumbers(document, out var inferred)
            ? inferred
            : null;
    }
}

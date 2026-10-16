using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

/// <summary>
/// Resolves one document-wide equation-number format for every formula family.
///
/// The format is never stored in the document. The one chosen in this Word
/// session wins; otherwise the first proven equation-number series in document
/// order (Word-native or MathType) shows it, and a document without numbers
/// uses the user's default. This keeps OMML, MathType and VisualTeX counters
/// independent while preventing each host family from inventing its own
/// display format.
/// </summary>
internal static class WordEquationNumberFormatAuthority
{
    internal static string Resolve(Document document)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));

        var stored =
            WordEquationNumbering.CaptureEquationNumberFormatVariable(
                document);
        if (stored.Exists)
            return EquationNumberFormat.Resolve(stored.Value).Id;

        if (TryInferFromNumbers(document, out var inferred))
            return inferred;

        return EquationNumberFormat.Resolve(
            WordEquationNumbering
                .GetDefaultEquationNumberFormatId()).Id;
    }

    internal static bool TryInferFromNumbers(Document document, out string formatId)
    {
        formatId = string.Empty;
        var candidates =
            new List<(int Start, string FormatId)>(2);
        if (WordNativeNumberingAdoptionPlan.TryInferDocumentFormat(
                document,
                out var nativeStart,
                out var nativeFormat))
        {
            candidates.Add(
                (nativeStart, nativeFormat));
        }

        if (MathTypeEquationNumbering.TryInferDocumentFormat(
                document,
                out var mathTypeStart,
                out var mathTypeFormat))
        {
            candidates.Add(
                (mathTypeStart, mathTypeFormat));
        }

        if (candidates.Count == 0)
            return false;
        var first =
            candidates
                .OrderBy(candidate => candidate.Start)
                .First();
        formatId = EquationNumberFormat.Resolve(
            first.FormatId).Id;
        return true;
    }
}

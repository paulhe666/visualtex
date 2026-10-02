using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

/// <summary>
/// Resolves one document-wide equation-number format for every formula family.
///
/// New VisualTeX documents persist the selected format in the Word document
/// variable. Older/native documents predate that variable, so the first proven
/// native equation-number series in document order becomes the migration
/// authority. This keeps OMML, MathType and VisualTeX counters independent while
/// preventing each host family from inventing its own display format.
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

        if (candidates.Count != 0)
        {
            var first =
                candidates
                    .OrderBy(candidate => candidate.Start)
                    .First();
            return EquationNumberFormat.Resolve(
                first.FormatId).Id;
        }

        return EquationNumberFormat.Resolve(
            WordEquationNumbering
                .GetDefaultEquationNumberFormatId()).Id;
    }
}

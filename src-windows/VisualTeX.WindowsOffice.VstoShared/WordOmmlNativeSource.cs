using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordOmmlNativeSource
{
    internal static FormulaMetadata RefreshForVisualTeX(
        Document document,
        Bookmark bookmark,
        FormulaMetadata stored)
    {
        Range? equationRange = null;
        try
        {
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var fingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                equationRange.WordOpenXML);
            if (string.Equals(
                    stored.NativeOmmlFingerprint,
                    fingerprint,
                    StringComparison.OrdinalIgnoreCase))
                return stored;

            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
                equationRange.WordOpenXML,
                display: string.Equals(
                    stored.DisplayMode,
                    "block",
                    StringComparison.Ordinal));
            var latex = MathMlToLatexConverter.Convert(mathMl);
            if (string.IsNullOrWhiteSpace(latex))
                throw new InvalidDataException(
                    "The Word-native OMML equation could not be converted back to editable LaTeX.");

            var refreshed = Clone(stored);
            var lineId = refreshed.Lines.FirstOrDefault()?.Id;
            if (string.IsNullOrWhiteSpace(lineId)) lineId = Guid.NewGuid().ToString();
            refreshed.Latex = latex;
            refreshed.Lines = new List<FormulaLine>
            {
                new() { Id = lineId!, Latex = latex },
            };
            refreshed.CodeFormat = "raw";
            refreshed.NativeOmmlFingerprint = fingerprint;
            refreshed.Validate();
            return refreshed;
        }
        finally
        {
            Release(equationRange);
        }
    }

    internal static void StampFingerprint(FormulaMetadata metadata, Range equationRange)
    {
        metadata.NativeOmmlFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
            equationRange.WordOpenXML);
    }

    private static FormulaMetadata Clone(FormulaMetadata metadata)
    {
        var clone = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata));
        return clone
            ?? throw new InvalidDataException("Unable to clone VisualTeX formula metadata.");
    }

    private static void Release(object? value)
    {
        if (value is null || !System.Runtime.InteropServices.Marshal.IsComObject(value)) return;
        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(value); } catch { }
    }
}

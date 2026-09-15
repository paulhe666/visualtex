using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private static WordInlineHostAlignment ReadInlineHostAlignment(Range range)
    {
        ParagraphFormat? paragraphFormat = null;
        try
        {
            paragraphFormat = range.ParagraphFormat;
            var value = (int)paragraphFormat.BaseLineAlignment;
            if (value < 0 || value > 4)
                throw new InvalidOperationException(
                    "The inline equation does not have one defined paragraph character-alignment mode.");
            return (WordInlineHostAlignment)value;
        }
        finally { Release(paragraphFormat); }
    }

    private static int CalculateVisualTeXInlinePosition(
        Range host,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition = null,
        double sourceSemanticFontSizePoints = 14,
        double targetSemanticFontSizePoints = 14) =>
        WordInlineAlignment.CalculateFontPositionForHost(
            ReadInlineHostAlignment(host), actualHeightPoints, exportedHeight,
            exportedBaseline, existingFontPosition, sourceSemanticFontSizePoints,
            targetSemanticFontSizePoints);

    private static int CalculateVisualTeXInlinePosition(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline,
        float? existingFontPosition = null,
        double sourceSemanticFontSizePoints = 14,
        double targetSemanticFontSizePoints = 14)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            return CalculateVisualTeXInlinePosition(
                range, actualHeightPoints, exportedHeight, exportedBaseline,
                existingFontPosition, sourceSemanticFontSizePoints,
                targetSemanticFontSizePoints);
        }
        finally { Release(range); }
    }

    private static void ApplyVisualTeXInlineHostAlignment(
        InlineShape shape,
        FormulaMetadata metadata)
    {
        var semanticSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
        ApplyInlineBaseline(
            shape, shape.Height, (float)(metadata.RenderHeightPx ?? 0),
            metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
            semanticSize);
        CacheFinalWordInlineOleGeometry(shape, metadata, inline: true);
    }
}

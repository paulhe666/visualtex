using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordFormulaHostSetup;
using static VisualTeX.WordVsto.WordInlineBaseline;

namespace VisualTeX.WordVsto;

/// <summary>
/// The vertical position of inline OLE formulas: Word's font position for VisualTeX and MathType objects, measured
/// previews, and the presentation size after an edit.
/// </summary>
internal static class WordInlineOlePosition
{
    internal readonly struct InlineOlePreviewMetrics
    {
        internal InlineOlePreviewMetrics(float inkHeightRatio, float bottomWhitespaceRatio)
        {
            InkHeightRatio = inkHeightRatio;
            BottomWhitespaceRatio = bottomWhitespaceRatio;
        }

        internal float InkHeightRatio { get; }
        internal float BottomWhitespaceRatio { get; }
    }

    internal static float? ReadDefinedShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        Range? probe = null;
        Document? document = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            range = shape.Range;
            document = range.Document;

            // An embedded OLE is an EMBED field, not a single character range.
            // Current VisualTeX and genuine MathType objects keep the field
            // instruction on the prose baseline and position only the U+0001
            // object/result character. Read that character first so resizing a
            // legacy formula scales its real baseline instead of treating the
            // mixed field range as wdUndefined.
            for (var position = range.Start; position < range.End; position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                var objectPosition = font.Position;
                if (objectPosition != (int)WdConstants.wdUndefined
                    && objectPosition >= -256
                    && objectPosition <= 256)
                    return objectPosition;
            }

            Release(font);
            font = null;
            font = range.Font;
            var fallbackPosition = font.Position;
            return fallbackPosition == (int)WdConstants.wdUndefined
                || fallbackPosition < -256
                || fallbackPosition > 256
                    ? null
                    : fallbackPosition;
        }
        catch { return null; }
        finally
        {
            Release(font);
            Release(probe);
            Release(range);
            Release(document);
        }
    }

    private static int ResolveVisualTeXInlineEditWordPosition(InlineShape shape)
    {
        var current = ReadInlineOleWordPosition(shape);
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? peers = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return current;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            peers = paragraphRange.InlineShapes;
            var positions = new List<int>();
            for (var index = 1; index <= peers.Count; index++)
            {
                InlineShape? peer = null;
                Range? peerRange = null;
                try
                {
                    peer = peers[index];
                    peerRange = peer.Range;
                    if (peerRange.Start == shapeRange.Start
                        && peerRange.End == shapeRange.End)
                        continue;
                    if (!MathTypeOleInterop.IsMathTypeOle(peer)) continue;
                    positions.Add(ReadInlineOleWordPosition(peer));
                }
                catch (COMException) { }
                finally
                {
                    Release(peerRange);
                    Release(peer);
                }
            }
            if (positions.Count < 2) return current;
            var profile = positions
                .GroupBy(position => position)
                .Select(group => new { Position = group.Key, Count = group.Count() })
                .OrderByDescending(item => item.Count)
                .ThenBy(item => Math.Abs(item.Position))
                .First();
            if (profile.Count < 2 || profile.Count * 2 <= positions.Count)
                return current;
            return profile.Position;
        }
        catch { return current; }
        finally
        {
            Release(peers);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static int ReadInlineOleWordPosition(InlineShape shape)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            // Word exposes an embedded MathType object as an EMBED field.  The
            // InlineShape.Range therefore spans the field instruction plus the
            // single U+0001 object/result character.  Genuine MathType keeps the
            // field instruction at Position=0 and applies the vertical offset only
            // to U+0001.  Reading range.Start therefore returns the wrong baseline.
            for (var position = shapeRange.Start; position < shapeRange.End; position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(position, position + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                var wordPosition = font.Position;
                if (wordPosition != (int)WdConstants.wdUndefined
                    && wordPosition >= -256
                    && wordPosition <= 256)
                    return wordPosition;
            }
        }
        catch { }
        finally
        {
            Release(font);
            Release(probe);
            Release(shapeRange);
            Release(document);
        }
        return (int)Math.Round(ReadDefinedShapeFontPosition(shape) ?? 0f);
    }

    private static int CalculateMathTypeOleWordPosition(
        float actualHeightPoints,
        float exportedHeight,
        float? exportedBaseline)
    {
        if (!(actualHeightPoints > 0)
            || !(exportedHeight > 0)
            || !exportedBaseline.HasValue
            || float.IsNaN(exportedBaseline.Value)
            || float.IsInfinity(exportedBaseline.Value)
            || exportedBaseline.Value < 0
            || exportedBaseline.Value >= exportedHeight)
            return 0;

        // MathType's Word integration stores the OLE object's character position
        // from the equation baseline toward the bottom of the picture. Keep this
        // calculation local to MathType instead of sharing VisualTeX's ordinary
        // inline-OLE optical alignment code; changes to VisualTeX inline layout must
        // never alter Equation.DSMT4 placement again.
        var baselineFromBottomPoints =
            actualHeightPoints * (exportedHeight - exportedBaseline.Value) / exportedHeight;
        var rounded = Math.Max(
            0,
            (int)Math.Round(baselineFromBottomPoints, MidpointRounding.AwayFromZero));

        // The offline WMF presentation includes Word's one-point character-box
        // allowance. Native MathType samples (ordinary fractions, radicals and
        // mixed inline equations) place the U+0001 object one point above the raw
        // picture descent, while the WMF/dxaOrig dimensions retain the full bound.
        return -Math.Max(0, rounded - 1);
    }

    internal static void SetInlineOleWordPosition(
        InlineShape shape,
        int position,
        bool hostInstructionAlreadyBaseline = false)
    {
        Range? shapeRange = null;
        try
        {
            shapeRange = shape.Range;
            SetInlineOleWordPosition(
                shapeRange,
                position,
                hostInstructionAlreadyBaseline);
        }
        finally { Release(shapeRange); }
    }

    internal static void SetInlineOleWordPosition(
        Range hostRange,
        int position,
        bool hostInstructionAlreadyBaseline = false)
    {
        Range? shapeRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = hostRange.Duplicate;
            // Match genuine MathType field formatting exactly: EMBED instruction
            // characters remain on the paragraph baseline; only the U+0001 object
            // result receives the equation's vertical offset. Applying Position to
            // shape.Range shifts the whole field and can leak into following prose.
            if (!hostInstructionAlreadyBaseline)
            {
                font = shapeRange.Font;
                font.Position = 0;
                Release(font);
                font = null;
            }

            var clamped = Math.Max(-256, Math.Min(256, position));
            // InlineShape.Range is already the exact local OLE host. Resolve its
            // one EMBED Result directly instead of creating a COM Range for every
            // character in the field code/result span. Keep the old walk only for
            // unusual Word field topologies that do not expose the canonical host.
            try
            {
                fields = shapeRange.Fields;
                if (fields.Count == 1)
                {
                    field = fields[1];
                    if (field.Type == WdFieldType.wdFieldEmbed)
                    {
                        result = field.Result;
                        var text = result.Text ?? string.Empty;
                        if ((string.Equals(text, "\u0001", StringComparison.Ordinal)
                                || string.Equals(text, "\uFFFC", StringComparison.Ordinal))
                            && result.Start >= shapeRange.Start
                            && result.End <= shapeRange.End)
                        {
                            font = result.Font;
                            font.Position = clamped;
                            return;
                        }
                    }
                }
            }
            catch { }
            finally
            {
                Release(font);
                font = null;
                Release(result);
                result = null;
                Release(field);
                field = null;
                Release(fields);
                fields = null;
            }

            document = shapeRange.Document;
            for (var index = shapeRange.Start; index < shapeRange.End; index++)
            {
                Release(probe);
                probe = document.Range(index, index + 1);
                if (!string.Equals(probe.Text, "\u0001", StringComparison.Ordinal))
                    continue;
                font = probe.Font;
                font.Position = clamped;
                return;
            }
        }
        finally
        {
            Release(font);
            Release(probe);
            Release(result);
            Release(field);
            Release(fields);
            Release(shapeRange);
            Release(document);
        }
    }

    internal static InlineOlePreviewMetrics? TryMeasureInlineOlePreview(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            var bytes = range.EnhMetaFileBits as byte[];
            if (bytes is null || bytes.Length == 0) return null;
            using var stream = new System.IO.MemoryStream(bytes, writable: false);
            using var metafile = new System.Drawing.Imaging.Metafile(stream);
            return MeasureMetafilePreview(metafile);
        }
        catch { return null; }
        finally { Release(range); }
    }

    private static InlineOlePreviewMetrics? TryMeasureMetafilePreview(string emfPath)
    {
        try
        {
            using var metafile = new System.Drawing.Imaging.Metafile(emfPath);
            return MeasureMetafilePreview(metafile);
        }
        catch { return null; }
    }

    private static InlineOlePreviewMetrics? MeasureMetafilePreview(
        System.Drawing.Imaging.Metafile metafile)
    {
        const int width = 640;
        const int height = 240;
        using var bitmap = new Bitmap(width, height);
        using (var graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            graphics.DrawImage(metafile, 0, 0, width, height);
        }

        var minY = height;
        var maxY = -1;
        for (var y = 0; y < height; y++)
        {
            for (var x = 0; x < width; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.R >= 245 && pixel.G >= 245 && pixel.B >= 245) continue;
                minY = Math.Min(minY, y);
                maxY = Math.Max(maxY, y);
            }
        }
        if (maxY < minY) return null;

        var inkHeightRatio = (maxY - minY + 1f) / height;
        var bottomWhitespaceRatio = (height - 1f - maxY) / height;
        if (!(inkHeightRatio > 0.01f)
            || float.IsNaN(inkHeightRatio)
            || float.IsInfinity(inkHeightRatio))
            return null;
        return new InlineOlePreviewMetrics(inkHeightRatio, bottomWhitespaceRatio);
    }

    private static (float Width, float Height) CalculateMathTypeEditedPresentationSize(
        float oldWidth,
        float oldHeight,
        InlineOlePreviewMetrics? sourcePreview,
        InlineOlePreviewMetrics? editedPreview,
        float? newRenderWidth,
        float? newRenderHeight,
        double? originalRenderWidth,
        double? originalRenderHeight,
        double? originalFontSizePt,
        double? originalRenderFontSizePt)
    {
        var fallback = OfficeFormulaSizing.EditedSize(
            oldWidth,
            oldHeight,
            originalRenderWidth,
            originalRenderHeight,
            newRenderWidth ?? oldWidth / 0.75f,
            newRenderHeight ?? oldHeight / 0.75f,
            originalFontSizePt: originalFontSizePt,
            originalRenderFontSizePt: originalRenderFontSizePt);
        if (!sourcePreview.HasValue || !editedPreview.HasValue
            || !(oldHeight > 0)
            || !(sourcePreview.Value.InkHeightRatio > 0.01f)
            || !(editedPreview.Value.InkHeightRatio > 0.01f))
            return fallback;

        var height = oldHeight
            * sourcePreview.Value.InkHeightRatio
            / editedPreview.Value.InkHeightRatio;
        if (!(height > 0) || float.IsNaN(height) || float.IsInfinity(height))
            return fallback;

        var aspect = newRenderWidth is > 0 && newRenderHeight is > 0
            ? newRenderWidth.Value / newRenderHeight.Value
            : fallback.Width / Math.Max(0.01f, fallback.Height);
        if (!(aspect > 0) || float.IsNaN(aspect) || float.IsInfinity(aspect))
            return fallback;

        // Preserve the native MathType glyph scale, not the outer OLE box. A
        // MathType preview typically contains appreciable ascent/descent padding,
        // whereas VisualTeX's EMF is tightly cropped. Keeping the same outer
        // height would therefore enlarge the visible glyphs by 20–40 percent.
        height = Math.Max(1f, Math.Min(oldHeight * 4f, Math.Max(oldHeight * 0.25f, height)));
        return (Math.Max(1f, height * aspect), height);
    }

    private static void ResetShapeFontPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetRangeFontPosition(range);
        }
        finally { Release(range); }
    }

    internal static void ResetDisplayFormulaPosition(InlineShape shape)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            ResetDisplayFormulaPosition(range);
        }
        finally { Release(range); }
    }

    internal static void ResetDisplayFormulaPosition(Range formulaRange)
    {
        ResetRangeFontPosition(formulaRange);
        ResetParagraphTypingPosition(formulaRange);
    }
    internal static WordInlineHostAlignment ReadInlineHostAlignment(Range range)
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

    internal static int CalculateVisualTeXInlinePosition(
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

    internal static int CalculateVisualTeXInlinePosition(
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

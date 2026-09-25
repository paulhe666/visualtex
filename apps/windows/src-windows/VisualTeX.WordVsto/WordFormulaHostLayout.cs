using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// Word presentation rules for OMML and VisualTeX hosts.
/// Structural insertion/deletion and numbering are intentionally not handled
/// here.
/// </summary>
internal static class WordFormulaHostLayout
{
    internal static void ApplyVisualTeXGeometry(
        InlineShape shape,
        FormulaMetadata metadata,
        float widthPoints,
        float heightPoints,
        float exportedHeightPixels,
        float? exportedBaselinePixels)
    {
        if (shape is null) throw new ArgumentNullException(nameof(shape));
        if (metadata is null) throw new ArgumentNullException(nameof(metadata));

        var width = Math.Max(1f, widthPoints);
        var height = Math.Max(1f, heightPoints);
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
        shape.Width = width;
        shape.Height = height;
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

        var inline = string.Equals(
            metadata.DisplayMode,
            "inline",
            StringComparison.OrdinalIgnoreCase);
        if (!inline)
        {
            metadata.WordInlineOleWidthPt = null;
            metadata.WordInlineOleHeightPt = null;
            metadata.WordInlineOlePositionPt = null;
            SetInlineOleWordPosition(shape, 0);
            WordFormulaMetadataReader.CacheMetadata(shape, metadata);
            return;
        }

        metadata.WordInlineOleWidthPt = shape.Width;
        metadata.WordInlineOleHeightPt = shape.Height;
        var semanticSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
        var position = CalculateVisualTeXInlinePosition(
            shape,
            shape.Height,
            exportedHeightPixels,
            exportedBaselinePixels,
            semanticSize);
        metadata.WordInlineOlePositionPt = position;
        metadata.WordInlineSourceBottomWhitespacePt = null;
        metadata.Validate();

        SetInlineOleWordPosition(shape, position);
        WordFormulaMetadataReader.CacheMetadata(shape, metadata);
    }

    internal static void ConfigureDisplayParagraph(
        Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A display formula must belong to exactly one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraphRange.ParagraphFormat;
            format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
        }
        finally
        {
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void ConfigureVisualTeXDisplayParagraph(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata)
    {
        WordEquationNumbering.ConfigureVisualTeXDisplayHostLayout(
            document,
            formulaRange,
            formulaHeightPoints,
            metadata);
    }

    internal static void ApplyOmmlLocalTypography(
        Range equationRange,
        double fontSizePoints)
    {
        if (equationRange is null)
            throw new ArgumentNullException(nameof(equationRange));

        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            var normalized = FormulaFontSize.NormalizeWordOmmlSize(fontSizePoints);
            font = equationRange.Font;
            font.Position = 0;
            font.Size = normalized;
            try { font.SizeBi = normalized; } catch { }
        }
        finally { Release(font); }
    }

    /// <summary>
    /// Restores an ordinary Word insertion point after an inline host without
    /// creating VTBL/U+200C or any other persistent document character.
    /// </summary>
    internal static void RestoreInlineCaret(
        WordApplication application,
        Range formulaRange)
    {
        Selection? selection = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        Range? source = null;
        Microsoft.Office.Interop.Word.Font? sourceFont = null;
        try
        {
            selection = application.Selection;
            selection.SetRange(formulaRange.End, formulaRange.End);
            selectionFont = selection.Font;

            source = FindNearestOrdinaryTextCharacter(formulaRange);
            if (source is not null)
            {
                sourceFont = source.Font;
                TryCopyTypingFont(sourceFont, selectionFont);
            }

            // Position is the only property that must never inherit from the OLE
            // result character.
            selectionFont.Position = 0;
        }
        catch
        {
            try
            {
                selection?.SetRange(formulaRange.End, formulaRange.End);
                selectionFont ??= selection?.Font;
                if (selectionFont is not null) selectionFont.Position = 0;
            }
            catch { }
        }
        finally
        {
            Release(sourceFont);
            Release(source);
            Release(selectionFont);
            Release(selection);
        }
    }

    private static int CalculateVisualTeXInlinePosition(
        InlineShape shape,
        float actualHeightPoints,
        float exportedHeightPixels,
        float? exportedBaselinePixels,
        double semanticFontSizePoints)
    {
        Range? range = null;
        ParagraphFormat? paragraph = null;
        try
        {
            range = shape.Range;
            paragraph = range.ParagraphFormat;
            var raw = (int)paragraph.BaseLineAlignment;
            if (raw < 0 || raw > 4) raw = (int)WordInlineHostAlignment.Baseline;
            return WordInlineAlignment.CalculateFontPositionForHost(
                (WordInlineHostAlignment)raw,
                actualHeightPoints,
                exportedHeightPixels,
                exportedBaselinePixels,
                existingFontPosition: null,
                sourceSemanticFontSizePoints: semanticFontSizePoints,
                targetSemanticFontSizePoints: semanticFontSizePoints);
        }
        finally
        {
            Release(paragraph);
            Release(range);
        }
    }

    private static void SetInlineOleWordPosition(
        InlineShape shape,
        int position)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Document? document = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;

            // Keep the EMBED instruction on the prose baseline.
            font = shapeRange.Font;
            font.Position = 0;
            Release(font);
            font = null;

            var clamped = Math.Max(-256, Math.Min(256, position));
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
            Release(document);
            Release(shapeRange);
        }
    }

    private static Range? FindNearestOrdinaryTextCharacter(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? candidate = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count < 1) return null;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var bodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);

            for (var position = formulaRange.End; position < bodyEnd; position++)
            {
                Release(candidate);
                candidate = paragraphRange.Document.Range(position, position + 1);
                if (IsOrdinaryTypingCharacter(candidate.Text))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }

            for (var position = formulaRange.Start - 1;
                 position >= paragraphRange.Start;
                 position--)
            {
                Release(candidate);
                candidate = paragraphRange.Document.Range(position, position + 1);
                if (IsOrdinaryTypingCharacter(candidate.Text))
                {
                    var result = candidate;
                    candidate = null;
                    return result;
                }
            }
            return null;
        }
        finally
        {
            Release(candidate);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool IsOrdinaryTypingCharacter(string? text)
    {
        if (string.IsNullOrEmpty(text) || text.Length != 1) return false;
        var value = text[0];
        return value is not '\r' and not '\n' and not '\t'
            and not '\v' and not '\a' and not '\u0001'
            and not '\u200B' and not '\u200C';
    }

    private static void TryCopyTypingFont(
        Microsoft.Office.Interop.Word.Font source,
        Microsoft.Office.Interop.Word.Font target)
    {
        try { target.Name = source.Name; } catch { }
        try
        {
            var size = source.Size;
            if (size > 0 && size < 1000) target.Size = size;
        }
        catch { }
        try
        {
            var sizeBi = source.SizeBi;
            if (sizeBi > 0 && sizeBi < 1000) target.SizeBi = sizeBi;
        }
        catch { }
        try { target.Bold = source.Bold; } catch { }
        try { target.Italic = source.Italic; } catch { }
        try { target.Color = source.Color; } catch { }
        try { target.Underline = source.Underline; } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

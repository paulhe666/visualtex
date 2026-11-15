using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordInlineBaseline;
using static VisualTeX.WordVsto.WordInlineOlePosition;

namespace VisualTeX.WordVsto;

/// <summary>
/// Setting up a formula host once it is written: its size and metadata, inline alignment or display position, its
/// number, the inline OLE size kept in its metadata, and the edit hosts that can be reused.
/// </summary>
internal static class WordFormulaHostSetup
{
    internal static void TryReconcileShape(
        Document document,
        InlineShape shape,
        FormulaMetadata metadata,
        bool numberingOrderMayHaveChanged = true,
        bool reuseExistingNumberedTableFormatting = false,
        Table? knownNumberedTable = null,
        bool numberingScaffoldOnly = false)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
            {
                RemoveInlineBaselineSentinel(document, metadata.FormulaId);
                ResetDisplayFormulaPosition(range);
            }
            if (numberingScaffoldOnly && metadata.Numbered)
            {
                WordNumberingReconcile.BuildFormulaNumberingScaffoldForConversion(
                    document,
                    range,
                    shape.Height,
                    metadata,
                    knownNumberedTable);
            }
            else
            {
                WordNumberingReconcile.ReconcileFormula(
                    document,
                    range,
                    shape.Height,
                    metadata,
                    numberingOrderMayHaveChanged,
                    reuseExistingNumberedTableFormatting,
                    knownNumberedTable);
            }
        }
        finally { Release(range); }
    }

    private static bool NumberingOrderMayHaveChanged(
        FormulaMetadata? originalMetadata,
        FormulaMetadata metadata)
    {
        if (originalMetadata is null) return true;
        return originalMetadata.Numbered != metadata.Numbered
            || !string.Equals(
                originalMetadata.DisplayMode,
                metadata.DisplayMode,
                StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryReadStoredWordInlineOleSize(
        FormulaMetadata? metadata,
        out float width,
        out float height)
    {
        width = 0;
        height = 0;
        if (metadata?.WordInlineOleWidthPt is not > 0
            || metadata.WordInlineOleHeightPt is not > 0)
            return false;
        width = (float)metadata.WordInlineOleWidthPt.Value;
        height = (float)metadata.WordInlineOleHeightPt.Value;
        return width > 0
            && height > 0
            && !float.IsNaN(width)
            && !float.IsInfinity(width)
            && !float.IsNaN(height)
            && !float.IsInfinity(height);
    }

    private static void StoreWordInlineOleSize(
        FormulaMetadata metadata,
        float width,
        float height,
        bool inline)
    {
        if (!inline
            || width <= 0
            || height <= 0
            || float.IsNaN(width)
            || float.IsInfinity(width)
            || float.IsNaN(height)
            || float.IsInfinity(height))
        {
            metadata.WordInlineOleWidthPt = null;
            metadata.WordInlineOleHeightPt = null;
            metadata.WordInlineOlePositionPt = null;
            return;
        }
        metadata.WordInlineOleWidthPt = width;
        metadata.WordInlineOleHeightPt = height;
    }

    internal static void CacheFinalWordInlineOleGeometry(
        InlineShape shape,
        FormulaMetadata metadata,
        bool inline)
    {
        if (!inline) return;
        StoreWordInlineOleSize(metadata, shape.Width, shape.Height, inline: true);
        metadata.WordInlineOlePositionPt = ReadInlineOleWordPosition(shape);
        metadata.Validate();
        // Host geometry is Word-specific and can differ slightly from the OLE
        // server's requested extent after Font.Position materializes the EMBED
        // result. Keep that final Word state in the Word-side metadata cache so a
        // later baseline repair can restore the exact last-known-good extent.
        WordFormulaMetadataReader.CacheMetadata(shape, metadata);
    }

    internal static bool IsNumberedBlockOmml(FormulaMetadata metadata) =>
        metadata.Numbered
        && string.Equals(
            metadata.DisplayMode,
            "block",
            StringComparison.OrdinalIgnoreCase);

    internal static bool TryNormalizeDefinedWordFontSize(
        float value,
        out float fontSizePt)
    {
        fontSizePt = FormulaFontSize.DefaultPt;
        // Word uses 9,999,999 for mixed/undefined Font.Size. Fall back to the
        // persisted semantic size instead of treating that sentinel as a real value.
        if (value <= 0f
            || value >= 1000f
            || float.IsNaN(value)
            || float.IsInfinity(value))
            return false;
        fontSizePt = FormulaFontSize.Normalize(value);
        return true;
    }

    private static bool HasReusableInlineOmmlEditHost(
        Range formulaRange)
    {
        Tables? tables = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Fields? fields = null;
        InlineShapes? shapes = null;
        try
        {
            tables = formulaRange.Tables;
            if (tables.Count > 0) return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathInline) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            return fields.Count == 0 && shapes.Count == 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(shapes);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(tables);
        }
    }

    private static bool HasReusableContentOnlyDisplayOmmlEditHost(
        Range formulaRange)
    {
        Tables? tables = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Fields? fields = null;
        InlineShapes? shapes = null;
        try
        {
            // Same-state unnumbered Display edits replace only this OMath. A
            // user's cell is also a valid unchanged host: routing it through
            // numbering reconciliation needlessly replaces paragraph runs and
            // then compares the new content against a provisional fingerprint.
            // Keep the content-only placeholder transaction and final full-tree
            // validation shared with body Display edits. Only the table case
            // needs a local physical-cell check; the large body fast path keeps
            // avoiding the layout-triggering Information(wdWithInTable) query.
            tables = formulaRange.Tables;
            if (tables.Count > 0 && !WordFormulaHost.IsWithinSingleCell(formulaRange))
                return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            return fields.Count == 0 && shapes.Count == 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(shapes);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(tables);
        }
    }

    private static bool HasReusableStandaloneDisplayOmmlHost(
        Document document,
        Range formulaRange)
    {
        Tables? tables = null;
        OMaths? maths = null;
        OMath? math = null;
        Range? mathRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        InlineShapes? shapes = null;
        Fields? fields = null;
        try
        {
            tables = formulaRange.Tables;
            if (tables.Count > 0) return false;
            maths = formulaRange.OMaths;
            if (maths.Count != 1) return false;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay) return false;
            mathRange = math.Range.Duplicate;
            if (mathRange.Start != formulaRange.Start
                || mathRange.End != formulaRange.End)
                return false;
            fields = formulaRange.Fields;
            shapes = formulaRange.InlineShapes;
            if (fields.Count != 0 || shapes.Count != 0) return false;

            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (formulaRange.Start < paragraphRange.Start
                || formulaRange.End > paragraphRange.End)
                return false;
            prefix = document.Range(paragraphRange.Start, formulaRange.Start);
            suffix = document.Range(formulaRange.End, paragraphRange.End);
            return !ContainsVisibleBodyText(prefix.Text)
                && !ContainsVisibleBodyText(suffix.Text);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(fields);
            Release(shapes);
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(tables);
        }
    }

    internal static void Configure(
        InlineShape shape,
        FormulaMetadata metadata,
        float maxWidth,
        float maxHeight,
        string imagePath,
        float exportedHeight,
        float? exportedBaseline,
        bool alignInline,
        bool nativeOleKnown = false,
        bool trustExportDimensions = false)
    {
        // maxWidth/maxHeight are already the SVG's physical size after the
        // 96 dpi CSS-pixel to 72 dpi Word-point conversion. A 12 pt minimum
        // width scales narrow inline formulas (notably x) far above their
        // semantic font size. Keep only a one-point safety floor.
        var width = Math.Max(1f, maxWidth);
        var height = Math.Max(1f, maxHeight);
        if (!trustExportDimensions)
        {
            using var image = Image.FromFile(imagePath);
            var ratio = image.Width / (float)Math.Max(1, image.Height);
            height = width / ratio;
            if (maxHeight > 0 && height > maxHeight)
            {
                height = maxHeight;
                width = height * ratio;
            }
        }
        // An OLE object is initially created with the placeholder preview's 4:1
        // aspect ratio. Setting only Width while aspect-ratio locking is enabled
        // therefore distorts the real formula. Apply both natural dimensions
        // explicitly, then lock the resolved ratio for later user resizing.
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
        shape.Width = width;
        shape.Height = height;
        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
        if (!nativeOleKnown && !WordFormulaMetadataReader.IsNativeOle(shape))
        {
            var encoded = FormulaMetadataCodec.Encode(metadata);
            shape.Title = encoded;
            shape.AlternativeText = encoded;
        }
        if (alignInline)
            ApplyInlineBaseline(
                shape,
                shape.Height,
                exportedHeight,
                exportedBaseline,
                FormulaFontSize.ResolveSemanticFontSize(metadata));
        else
            ResetDisplayFormulaPosition(shape);
    }

    private static bool ShouldAlignInline(InlineShape shape, FormulaMetadata metadata)
    {
        Range? range = null;
        try
        {
            range = shape.Range;
            return ShouldAlignInline(range, metadata);
        }
        finally { Release(range); }
    }

    private static bool ShouldAlignInline(Range formulaRange, FormulaMetadata metadata)
    {
        if (string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
            return true;
        return HasVisibleSurroundingText(formulaRange);
    }

    private static bool HasVisibleFollowingInlineText(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? after = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var paragraphBodyEnd = Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            if (formulaRange.End >= paragraphBodyEnd) return false;
            after = paragraphRange.Duplicate;
            after.SetRange(formulaRange.End, paragraphBodyEnd);
            return ContainsVisibleBodyText(after.Text);
        }
        catch { return false; }
        finally
        {
            Release(after);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static bool HasVisibleSurroundingText(Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count == 0) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (formulaRange.Start > paragraphRange.Start)
            {
                before = paragraphRange.Duplicate;
                before.SetRange(paragraphRange.Start, formulaRange.Start);
                if (ContainsVisibleBodyText(before.Text)) return true;
            }
            if (formulaRange.End < paragraphRange.End)
            {
                after = paragraphRange.Duplicate;
                after.SetRange(formulaRange.End, paragraphRange.End);
                if (ContainsVisibleBodyText(after.Text)) return true;
            }
            return false;
        }
        catch { return false; }
        finally
        {
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static bool ContainsVisibleBodyText(string? value)
    {
        if (string.IsNullOrEmpty(value)) return false;
        foreach (var character in value!)
        {
            if (character is '\r' or '\n' or '\t' or '\v' or '\a' or '\u0001' or '\u200B' or '\u200C')
                continue;
            if (!char.IsWhiteSpace(character)) return true;
        }
        return false;
    }
}

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordEquationNumbering;
using static VisualTeX.WordVsto.WordNumberedDisplayLayout;
using static VisualTeX.WordVsto.WordNumberingFields;
using static VisualTeX.WordVsto.WordNumberingReconcile;
using static VisualTeX.WordVsto.WordVisibleEquationNumbers;

namespace VisualTeX.WordVsto;

/// <summary>
/// The equation paragraph: the VisualTeX display host, tab stops, the measured formula width and its physical
/// centre.
/// </summary>
internal static class WordEquationParagraphLayout
{
    private const int WdTabAlignmentCenter = 1;
    private const int WdTabAlignmentRight = 2;
    private const int WdTabLeaderSpaces = 0;

    internal static void NormalizePreservedVisualTeXNumberedTabParagraph(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata)
    {
        Range? leading = null;
        OMaths? maths = null;
        InlineShapes? shapes = null;
        try
        {
            if (RangeIsWhollyWithinTable(formulaRange))
                throw new InvalidDataException(
                    "A preserved VisualTeX body tab paragraph cannot be inside a Word table.");
            if (formulaRange.Start <= document.Content.Start)
                throw new InvalidDataException(
                    "A preserved VisualTeX numbered OLE lost its leading center-tab slot.");

            leading = document.Range(
                formulaRange.Start - 1,
                formulaRange.Start);
            if (!string.Equals(
                    leading.Text,
                    "\t",
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "A preserved VisualTeX numbered OLE lost the leading center-tab character.");

            maths = formulaRange.OMaths;
            shapes = formulaRange.InlineShapes;
            if (maths.Count != 0
                || shapes.Count != 1)
                throw new InvalidDataException(
                    "The preserved VisualTeX tab host no longer contains exactly one OLE formula.");

            // Replacing an InlineShape can make Word reapply the ordinary block
            // writer's centered paragraph format even though the surrounding
            // center/right-tab numbering scaffold stayed intact. Restore only the
            // canonical paragraph geometry here; SEQ/REF fields and text ranges are
            // deliberately untouched.
            ConfigureEquationParagraph(
                formulaRange,
                numbered: true);
            AlignDisplayOleObjectToParagraphBaseline(
                document,
                formulaRange,
                formulaHeightPoints,
                metadata);

            // Replacing an InlineShape at the left edge of a preserved tab host
            // makes Word grow the existing VTEq_* bookmark backwards so it owns
            // the replacement OLE as well as the visible TAB/(REF) label. The
            // SEQ target and local REF remain healthy; repair only that convenience
            // bookmark from the proven local canonical label after the new OLE is
            // fully materialized.
            if (!TryRebindHealthyBodyTabVisibleNumber(
                    document,
                    formulaRange,
                    metadata.FormulaId,
                    NativeNumberBookmarkName(metadata.FormulaId)))
                throw new InvalidDataException(
                    "The preserved VisualTeX tab host lost its exact local equation-number label after OLE replacement.");
        }
        finally
        {
            Release(shapes);
            Release(maths);
            Release(leading);
        }
    }

    internal static void ConfigureVisualTeXDisplayHostLayout(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata,
        float? knownCenterTabPoints = null,
        float? knownRightTabPoints = null)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (formulaRange is null)
            throw new ArgumentNullException(nameof(formulaRange));
        if (metadata is null)
            throw new ArgumentNullException(nameof(metadata));
        if (!string.Equals(
                metadata.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "VisualTeX display-host layout requires block metadata.");
        using (WordOperationMetrics.Measure(
                   "FormatConversion.VisualTeXDisplay.ContainsNativeOmml"))
        {
            if (ContainsNativeOmml(formulaRange))
                throw new InvalidDataException(
                    "VisualTeX OLE display-host layout cannot be applied to native OMML.");
        }

        PrepareVisualTeXDisplayHostLayout(
            document,
            formulaRange,
            knownCenterTabPoints,
            knownRightTabPoints);
        AlignVisualTeXDisplayHostBaseline(
            document,
            formulaRange,
            formulaHeightPoints,
            metadata);
    }

    /// <summary>
    /// No indents on a paragraph laid out by tab stops (a display formula's line). An indent given in character
    /// units (「首行缩进 2 字符」, the usual Chinese body text) wins over one given in points: FirstLineIndent = 0
    /// alone leaves such a paragraph indented, and on a tab-laid line a first-line indent moves the formula and
    /// its number off their stops (measured in Word: both 21 pt to the left with 2 characters at 10.5 pt). The
    /// character units go first. <paramref name="firstLineOnly"/>: the left and right indents stay.
    /// </summary>
    internal static void ZeroIndents(ParagraphFormat format, bool firstLineOnly = false)
    {
        ZeroCharacterUnits(() => format.CharacterUnitFirstLineIndent, value => format.CharacterUnitFirstLineIndent = value);
        format.FirstLineIndent = 0f;
        if (firstLineOnly) return;
        ZeroCharacterUnits(() => format.CharacterUnitLeftIndent, value => format.CharacterUnitLeftIndent = value);
        format.LeftIndent = 0f;
        ZeroCharacterUnits(() => format.CharacterUnitRightIndent, value => format.CharacterUnitRightIndent = value);
        format.RightIndent = 0f;
    }

    private static void ZeroCharacterUnits(Func<float> read, Action<float> write)
    {
        try
        {
            if (Math.Abs(read()) > 0.001f) write(0f);
        }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    internal static void PrepareVisualTeXDisplayHostLayout(
        Document document,
        Range formulaRange,
        float? knownCenterTabPoints = null,
        float? knownRightTabPoints = null)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (formulaRange is null)
            throw new ArgumentNullException(nameof(formulaRange));

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        try
        {
            using (WordOperationMetrics.Measure(
                       "FormatConversion.VisualTeXDisplay.ParagraphFormat"))
            {
                paragraphs = formulaRange.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidDataException(
                        "A VisualTeX display host must belong to exactly one Word paragraph.");
                paragraph = paragraphs[1];
                var current = paragraph.Format;
                try { format = current.Duplicate; }
                finally { Release(current); }
                // Prepare the same six values on a detached format, then apply
                // them once. Live setters individually invalidate Word layout.
                format.Alignment =
                    WdParagraphAlignment.wdAlignParagraphJustify;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                ZeroIndents(format);
                paragraph.Format = format;
            }
        }
        finally
        {
            Release(format);
            Release(paragraph);
            Release(paragraphs);
        }

        using (WordOperationMetrics.Measure(
                   "FormatConversion.VisualTeXDisplay.TabStops"))
        {
            ConfigureEquationParagraphTabStops(
                formulaRange,
                numbered: true,
                knownCenterTabPoints,
                knownRightTabPoints);
        }
        using (WordOperationMetrics.Measure(
                   "FormatConversion.VisualTeXDisplay.LeadingTab"))
        {
            EnsureLeadingEquationTab(
                document,
                formulaRange);
        }
    }

    internal static void AlignVisualTeXDisplayHostBaseline(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        FormulaMetadata metadata)
    {
        using (WordOperationMetrics.Measure(
                   "FormatConversion.VisualTeXDisplay.AlignBaseline"))
        {
            AlignDisplayOleObjectToParagraphBaseline(
                document,
                formulaRange,
                formulaHeightPoints,
                metadata);
        }
    }

    internal static void ConfigureEquationParagraph(
        Range formulaRange,
        bool numbered,
        float? nativeOmmlWidthPoints = null)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Document? document = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        ListFormat? listFormat = null;
        OMaths? maths = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;
            var nativeOmmlParagraph = false;
            try
            {
                maths = formulaRange.OMaths;
                nativeOmmlParagraph = maths.Count > 0
                    && !(bool)formulaRange.get_Information(WdInformation.wdWithInTable);
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            ZeroIndents(format);
            if (!nativeOmmlParagraph)
            {
                // OLE/picture formulas carry their own transparent preview
                // margins, so their host paragraph must stay compact. Native
                // OMML formulas have no image padding and should retain the
                // document paragraph style exactly as a formula inserted by
                // Word itself would.
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                format.BaseLineAlignment =
                    (WdBaselineAlignment)(int)WordInlineHostAlignment.Automatic;
            }
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
            format.PageBreakBefore = 0;
            format.WidowControl = 0;
            try
            {
                listFormat = paragraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch (Exception ignored)
            {
                // Protected/custom stories can reject list normalization. The
                // page-break flags above still remove Word's black paragraph
                // marker when formatting marks are shown.
                IgnoredErrors.Ignore(ignored);
            }
            tabStops = format.TabStops;
            tabStops.ClearAll();

            // Word measures tab positions from the left edge of the column (or
            // table cell) the line is in: the stops come from that width, not the
            // page's (WordDisplayColumns). The indents were just set to zero.
            document = formulaRange.Document;
            var positions = WordDisplayColumns.TabStops(document, paragraphRange);
            if (!numbered && nativeOmmlParagraph)
            {
                // Standalone native OMML remains Word-native display math. The
                // tab-driven host below is specifically the VisualTeX OLE layout
                // contract and must not be imposed on Word's own OMath paragraph.
                format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                return;
            }
            // Use the same semantic center tab as MathType/VisualTeX OLE. Once
            // Latin Modern runs are represented as native normal-text OMath runs,
            // Word correctly centers the painted equation box and, on save/reopen,
            // may internally materialize an equivalent measured left tab from the
            // final glyph width. Persisting our own pre-save width estimate caused
            // a visible drift after Word normalized the OpenType glyph metrics.
            format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            tabStops.Add(
                positions.Center,
                (WdTabAlignment)WdTabAlignmentCenter,
                (WdTabLeader)WdTabLeaderSpaces);
            tabStops.Add(
                positions.Right,
                (WdTabAlignment)WdTabAlignmentRight,
                (WdTabLeader)WdTabLeaderSpaces);
        }
        finally
        {
            Release(maths);
            Release(listFormat);
            Release(tabStops);
            Release(format);
            Release(document);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void ConfigureEquationParagraphTabStops(
        Range formulaRange,
        bool numbered,
        float? knownCenterTabPoints = null,
        float? knownRightTabPoints = null)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Document? document = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "VisualTeX equation tab geometry requires one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            format = paragraph.Format;
            tabStops = format.TabStops;
            using (WordOperationMetrics.Measure("Display.TabStops.ClearAll"))
                tabStops.ClearAll();

            (float Center, float Right) positions;
            if (knownCenterTabPoints.HasValue
                && knownRightTabPoints.HasValue
                && knownCenterTabPoints.Value > 0f
                && knownRightTabPoints.Value >= knownCenterTabPoints.Value)
            {
                positions = (
                    knownCenterTabPoints.Value,
                    knownRightTabPoints.Value);
            }
            else
            {
                // the column (or table cell) the paragraph is in (WordDisplayColumns)
                using (WordOperationMetrics.Measure("Display.TabStops.ColumnWidth"))
                {
                    document = formulaRange.Document;
                    positions = WordDisplayColumns.TabStops(
                        document,
                        paragraphRange,
                        format.LeftIndent,
                        format.FirstLineIndent,
                        format.RightIndent);
                }
            }
            tabStops.Add(
                positions.Center,
                (WdTabAlignment)WdTabAlignmentCenter,
                (WdTabLeader)WdTabLeaderSpaces);
            if (numbered)
            {
                tabStops.Add(
                    positions.Right,
                    (WdTabAlignment)WdTabAlignmentRight,
                    (WdTabLeader)WdTabLeaderSpaces);
            }
        }
        finally
        {
            Release(tabStops);
            Release(format);
            Release(document);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static float? MeasureNativeOmmlWidthPoints(Range formulaRange)
    {
        Document? document = null;
        Range? visualRange = null;
        Range? start = null;
        Range? end = null;
        Window? window = null;
        Microsoft.Office.Interop.Word.View? view = null;
        Microsoft.Office.Interop.Word.Zoom? zoom = null;
        try
        {
            document = formulaRange.Document;
            var formulaText = formulaRange.Text ?? string.Empty;
            var visibleStart = formulaRange.Start
                + (formulaText.StartsWith("\t", StringComparison.Ordinal) ? 1 : 0);
            visibleStart = Math.Min(visibleStart, formulaRange.End);
            visualRange = document.Range(visibleStart, formulaRange.End);

            // Prefer Word's range geometry. It is stable across repeated bulk
            // insertions and does not force the hidden Word window to allocate a
            // screen rectangle for every OMath. Repeated Window.GetPoint calls can
            // terminate WINWORD after a few dozen formulas on Office 2021. Any
            // transient sub-point discrepancy is closed by
            // CorrectNumberedOmmlPhysicalCenter immediately after the tab is set.
            try
            {
                start = document.Range(visibleStart, visibleStart);
                end = formulaRange.Duplicate;
                end.Collapse(WdCollapseDirection.wdCollapseEnd);
                var startX = Convert.ToSingle(
                    start.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var endX = Convert.ToSingle(
                    end.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var rangeWidth = endX - startX;
                if (rangeWidth > 0f
                    && !float.IsNaN(rangeWidth)
                    && !float.IsInfinity(rangeWidth))
                {
                    if (string.Equals(
                            Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                            "1",
                            StringComparison.Ordinal))
                    {
                        Console.WriteLine(
                            $"    [layout] native OMML width: range={startX:0.###}-{endX:0.###}, points={rangeWidth:0.###}");
                    }
                    return rangeWidth;
                }
            }
            catch (COMException ignored)
            {
                // Protected/custom stories can reject positional information.
                // Fall through to the screen-box fallback only for that case.
                IgnoredErrors.Ignore(ignored);
            }

            // Exceptional fallback for stories without usable Range.Information.
            // Keep this out of the normal bulk path because Word's GetPoint API is
            // not stable under many consecutive hidden-window measurements.
            try
            {
                window = document.ActiveWindow;
                window.GetPoint(
                    out _,
                    out _,
                    out var screenWidth,
                    out _,
                    visualRange);
                view = window.View;
                zoom = view.Zoom;
                var zoomPercentage = zoom.Percentage;
                var dpi = 96u;
                try
                {
                    var detectedDpi = GetDpiForWindow(new IntPtr(window.Hwnd));
                    if (detectedDpi > 0) dpi = detectedDpi;
                }
                catch (EntryPointNotFoundException ignored)
                {
                    // Windows versions without per-window DPI APIs use 96 DPI.
                    IgnoredErrors.Ignore(ignored);
                }
                if (screenWidth <= 0 || zoomPercentage <= 0 || dpi == 0)
                    return null;
                var screenWidthPoints = screenWidth
                    * 72f
                    * 100f
                    / dpi
                    / zoomPercentage;
                return screenWidthPoints > 0f
                    && !float.IsNaN(screenWidthPoints)
                    && !float.IsInfinity(screenWidthPoints)
                        ? screenWidthPoints
                        : null;
            }
            catch (COMException)
            {
                return null;
            }
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(zoom);
            Release(view);
            Release(window);
            Release(end);
            Release(start);
            Release(visualRange);
            Release(document);
        }
    }

    private static void CorrectNumberedOmmlPhysicalCenter(Range formulaRange)
    {
        Document? document = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? formulaTab = null;
        Range? start = null;
        Range? end = null;
        try
        {
            if (!ContainsNativeOmml(formulaRange)
                || IsNumberedEquationTable(formulaRange))
                return;

            document = formulaRange.Document;
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;

            // the column the paragraph is in, as ConfigureEquationParagraph set it
            var expected = WordDisplayColumns.TabStops(document, paragraphRange);

            var formulaText = formulaRange.Text ?? string.Empty;
            var visibleStart = formulaRange.Start
                + (formulaText.StartsWith("\t", StringComparison.Ordinal) ? 1 : 0);
            visibleStart = Math.Min(visibleStart, formulaRange.End);

            // The width returned by Window.GetPoint is quantized to physical pixels.
            // After save/reopen Word can resolve the same OpenType glyphs a fraction
            // of a point wider or narrower than that first estimate. Close the loop
            // against Word's actual post-layout Range geometry and nudge only the
            // left formula tab; the right number tab and every field/bookmark remain
            // untouched. The correction is linear with the left tab position, so
            // one post-layout adjustment is sufficient. A second adjustment before
            // Word has repaginated can read the same stale box and apply the delta
            // twice. The hard guard avoids reacting to undefined COM geometry in
            // background/protected stories.
            {
                Release(end);
                end = null;
                Release(start);
                start = document.Range(visibleStart, visibleStart);
                end = formulaRange.Duplicate;
                end.Collapse(WdCollapseDirection.wdCollapseEnd);
                var startX = Convert.ToSingle(
                    start.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var endX = Convert.ToSingle(
                    end.get_Information(WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                if (float.IsNaN(startX)
                    || float.IsInfinity(startX)
                    || float.IsNaN(endX)
                    || float.IsInfinity(endX)
                    || startX < -10000f
                    || endX < -10000f
                    || endX <= startX)
                    return;

                var actualCenter = (startX + endX) / 2f;
                var correction = expected.Center - actualCenter;
                if (Math.Abs(correction) <= 0.1f) return;
                if (Math.Abs(correction) > 12f) return;

                Release(formulaTab);
                formulaTab = null;
                Release(tabStops);
                tabStops = format.TabStops;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    TabStop? candidate = null;
                    try
                    {
                        candidate = tabStops[index];
                        if (candidate.Alignment != WdTabAlignment.wdAlignTabLeft
                            || candidate.Position >= expected.Right - 0.1f)
                            continue;
                        formulaTab = candidate;
                        candidate = null;
                        break;
                    }
                    finally { Release(candidate); }
                }
                if (formulaTab is null) return;

                var correctedPosition = Math.Max(
                    0.1f,
                    Math.Min(expected.Right - 0.1f, formulaTab.Position + correction));
                // Moving the existing custom tab preserves the right-aligned number
                // stop. Clearing and re-adding the formula stop can make Word discard
                // the sibling custom tab in some builds, which leaves the paragraph
                // without a usable numbering layout.
                formulaTab.Position = correctedPosition;
                if (string.Equals(
                        Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE"),
                        "1",
                        StringComparison.Ordinal))
                {
                    Console.WriteLine(
                        $"    [layout] corrected OMML center: actual={actualCenter:0.###}, target={expected.Center:0.###}, tab={correctedPosition:0.###}");
                }
            }
        }
        catch (Exception ignored)
        {
            // Physical centering is a layout refinement. If Word refuses geometry
            // queries in a protected/background story, retain the width-based tab
            // computed by ConfigureEquationParagraph rather than breaking numbering.
            IgnoredErrors.Ignore(ignored);
        }
        finally
        {
            Release(end);
            Release(start);
            Release(formulaTab);
            Release(tabStops);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(document);
        }
    }

    internal static (float Center, float Right) CalculateEquationTabStops(
        float pageWidth,
        float leftMargin,
        float rightMargin,
        float leftIndent,
        float rightIndent)
    {
        var availableWidth = Math.Max(
            72f,
            pageWidth
                - Math.Max(0f, leftMargin)
                - Math.Max(0f, rightMargin)
                - Math.Max(0f, leftIndent)
                - Math.Max(0f, rightIndent));
        return (availableWidth / 2f, availableWidth);
    }

    internal static int CalculateEquationNumberFontPosition(
        float formulaHeightPoints,
        float numberFontSizePoints)
    {
        if (float.IsNaN(formulaHeightPoints)
            || float.IsInfinity(formulaHeightPoints)
            || formulaHeightPoints <= 0)
            return 0;
        if (float.IsNaN(numberFontSizePoints)
            || float.IsInfinity(numberFontSizePoints)
            || numberFontSizePoints <= 0
            || numberFontSizePoints > 256)
            numberFontSizePoints = 11f;
        return Math.Max(
            0,
            (int)Math.Round(
                (formulaHeightPoints - numberFontSizePoints) / 2f,
                MidpointRounding.AwayFromZero));
    }

    internal static (string Text, int FieldOffset) EquationNumberScaffold() => ("\t()", 2);

    internal static void EnsureLeadingEquationTab(Document document, Range formulaRange)
    {
        Range? preceding = null;
        Range? insertion = null;
        try
        {
            var formulaText = formulaRange.Text ?? string.Empty;
            if (formulaText.StartsWith("\t", StringComparison.Ordinal)) return;
            var insertionPosition = formulaRange.Start;
            if (formulaRange.Start > 0)
            {
                object precedingStart = formulaRange.Start - 1;
                object precedingEnd = formulaRange.Start;
                preceding = document.Range(ref precedingStart, ref precedingEnd);
                if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return;

                // A display OMath is preceded by Word's vertical-tab math
                // separator (0x0B). Its layout tab therefore sits one character
                // earlier and must be inspected/inserted outside the OMath edge.
                if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal))
                {
                    insertionPosition = formulaRange.Start - 1;
                    if (formulaRange.Start > 1)
                    {
                        preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
                        if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal)) return;
                    }
                }
            }

            object start = insertionPosition;
            object end = insertionPosition;
            insertion = document.Range(ref start, ref end);
            insertion.Text = "\t";
        }
        finally
        {
            Release(insertion);
            Release(preceding);
        }
    }

    private static void RemoveTrailingOmmlDisplaySeparators(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? visibleRange = null;
        Range? character = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            visibleRange = FindVisibleEquationNumberRange(document, formulaId);
            var scanEnd = visibleRange is not null
                ? visibleRange.Start
                : Math.Max(paragraphRange.Start, paragraphRange.End - 1);
            // Word can expose one or more U+000B display-math separators after a
            // native OMath. Inside a table these were previously removed by the
            // cell normalizer; in the tab paragraph they would act as a manual
            // line break and place the right-tab number on another visual line.
            // Delete only separators strictly outside OMath.Range, scanning
            // backwards so the live equation and visible-number ranges stay valid.
            for (var position = scanEnd - 1;
                 position >= formulaRange.End;
                 position--)
            {
                Release(character);
                character = document.Range(position, position + 1);
                if (string.Equals(character.Text, "\v", StringComparison.Ordinal))
                    character.Delete();
            }
        }
        finally
        {
            Release(character);
            Release(visibleRange);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void RemoveLeadingEquationTab(Document document, Range formulaRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? preceding = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var formulaText = formulaRange.Text ?? string.Empty;
            if (formulaText.StartsWith("\t", StringComparison.Ordinal)
                && formulaRange.Start < formulaRange.End)
            {
                preceding = document.Range(formulaRange.Start, formulaRange.Start + 1);
                preceding.Delete();
                return;
            }
            if (formulaRange.Start <= paragraphRange.Start) return;
            object start = formulaRange.Start - 1;
            object end = formulaRange.Start;
            preceding = document.Range(ref start, ref end);
            if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal))
            {
                preceding.Delete();
                return;
            }
            if (string.Equals(preceding.Text, "\v", StringComparison.Ordinal)
                && formulaRange.Start - 2 >= paragraphRange.Start)
            {
                preceding.SetRange(formulaRange.Start - 2, formulaRange.Start - 1);
                if (string.Equals(preceding.Text, "\t", StringComparison.Ordinal))
                    preceding.Delete();
            }
        }
        finally
        {
            Release(preceding);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }
}

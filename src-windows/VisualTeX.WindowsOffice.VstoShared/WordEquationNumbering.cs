using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed class EquationReferenceTarget
{
    public EquationReferenceTarget(
        string formulaId,
        int nativeReferenceItem,
        string numberText,
        string latexPreview,
        int position)
    {
        FormulaId = formulaId;
        NativeReferenceItem = nativeReferenceItem;
        NumberText = numberText;
        LatexPreview = latexPreview;
        Position = position;
    }

    public string FormulaId { get; }
    public int NativeReferenceItem { get; }
    public string NumberText { get; }
    public string LatexPreview { get; }
    public int Position { get; }

    public override string ToString() => $"({NumberText})    {LatexPreview}";
}

internal enum EquationReferenceStyle
{
    Parenthesized,
    EquationPrefix,
    NumberOnly,
}

internal static class WordEquationNumbering
{
    private const int WdTabAlignmentCenter = 1;
    private const int WdTabAlignmentRight = 2;
    private const int WdTabLeaderSpaces = 0;
    private const int WdFieldEmpty = -1;
    private const string LegacyEquationSequenceName = "VisualTeXEquation";
    private const string EquationBookmarkPrefix = "VTEq_";
    private const string NativeCaptionBookmarkPrefix = "VTEqCap_";
    private const string NativeNumberBookmarkPrefix = "VTEqNum_";
    private const float NativeCaptionFrameSizePoints = 1f;
    private const float NativeCaptionOffscreenPositionPoints = -1000f;

    public static void TryReconcile(Document document)
    {
        try { Reconcile(document); }
        catch
        {
            // Formula insertion/update is already durable. The user can retry
            // only the numbering command without duplicating or losing it.
        }
    }

    internal static void RemoveFormulaNumberingArtifacts(
        Document document,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(document, formulaId);
    }

    public static int Reconcile(Document document)
    {
        var numberedFormulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var ommlFormulaIds = WordOmmlFormulaStore.FormulaIds(document);
        InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            var inlineCount = inlineShapes.Count;
            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                Range? formulaRange = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange = shape.Range;
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            shape.Height,
                            metadata.FormulaId);
                        numberedFormulaIds.Add(metadata.FormulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(
                            document,
                            formulaRange,
                            metadata.FormulaId);
                    }
                }
                finally
                {
                    Release(formulaRange);
                    Release(shape);
                }
            }

            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                Range? formulaRange = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata is null || metadata.DisplayMode != "block") continue;
                    formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                    if (metadata.Numbered)
                    {
                        ConfigureNumberedDisplayFormula(
                            document,
                            formulaRange,
                            WordOmmlFormulaStore.EstimateHeightPoints(bookmark),
                            formulaId);
                        numberedFormulaIds.Add(formulaId);
                    }
                    else
                    {
                        ConfigureUnnumberedDisplayFormula(document, formulaRange, formulaId);
                    }
                }
                finally
                {
                    Release(formulaRange);
                    Release(bookmark);
                }
            }

            RemoveOrphanEquationArtifacts(document, numberedFormulaIds);

            // Word caches SEQ results independently from REF results. After a
            // numbered formula is deleted, refresh every native Equation SEQ
            // field in document order before updating any visible or body REF
            // field. Otherwise a REF can continue displaying the removed
            // formula's old ordinal until Word performs a later global update.
            UpdateNativeEquationSequenceFields(document);

            for (var index = 1; index <= inlineCount; index++)
            {
                InlineShape? shape = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(document, shape.Height, metadata.FormulaId);
                }
                finally { Release(shape); }
            }
            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata?.DisplayMode == "block" && metadata.Numbered)
                        UpdateEquationNumberFields(
                            document,
                            WordOmmlFormulaStore.EstimateHeightPoints(bookmark),
                            formulaId);
                }
                finally { Release(bookmark); }
            }
            UpdateNativeCrossReferences(document);
        }
        finally { Release(inlineShapes); }

        return numberedFormulaIds.Count;
    }

    private static FormulaMetadata? ReadMetadata(InlineShape shape) =>
        WordFormulaMetadataReader.TryRead(shape);

    private static void ConfigureNumberedDisplayFormula(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        string formulaId)
    {
        EnsureNumberedOmmlIsDisplay(formulaRange);
        var tableLayout = IsNumberedEquationTable(formulaRange);
        ConfigureEquationParagraph(formulaRange, numbered: !tableLayout);
        if (tableLayout)
            ConfigureNumberedEquationTable(formulaRange);
        else
            EnsureLeadingEquationTab(document, formulaRange);
        var sequenceName = GetNativeEquationSequenceName(document);
        EnsureNativeCaption(document, formulaRange, formulaId, sequenceName);
        EnsureVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaId);
    }

    private static void EnsureNumberedOmmlIsDisplay(Range formulaRange)
    {
        OMaths? maths = null;
        OMath? math = null;
        Range? refreshed = null;
        try
        {
            maths = formulaRange.OMaths;
            if (maths.Count == 0) return;
            math = maths[1];
            if (math.Type != WdOMathType.wdOMathDisplay)
            {
                math.Type = WdOMathType.wdOMathDisplay;
                math.BuildUp();
            }
            refreshed = math.Range;
            formulaRange.SetRange(refreshed.Start, refreshed.End);
        }
        finally
        {
            Release(refreshed);
            Release(math);
            Release(maths);
        }
    }

    private static bool IsNumberedEquationTable(Range formulaRange)
    {
        try
        {
            return (bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                && formulaRange.Tables.Count > 0
                && formulaRange.Tables[1].Columns.Count >= 3;
        }
        catch { return false; }
    }

    private static void ConfigureNumberedEquationTable(Range formulaRange)
    {
        Table? table = null;
        Columns? columns = null;
        Column? leftColumn = null;
        Column? centerColumn = null;
        Column? rightColumn = null;
        Cell? centerCell = null;
        Cell? numberCell = null;
        Range? centerRange = null;
        Range? numberRange = null;
        ParagraphFormat? centerFormat = null;
        ParagraphFormat? numberFormat = null;
        Borders? borders = null;
        try
        {
            table = formulaRange.Tables[1];
            table.AllowAutoFit = false;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            table.PreferredWidth = 100f;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            columns = table.Columns;
            leftColumn = columns[1];
            centerColumn = columns[2];
            rightColumn = columns[3];
            leftColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            centerColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            rightColumn.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            leftColumn.PreferredWidth = 20f;
            centerColumn.PreferredWidth = 60f;
            rightColumn.PreferredWidth = 20f;
            borders = table.Borders;
            borders.Enable = 0;
            centerCell = table.Cell(1, 2);
            numberCell = table.Cell(1, 3);
            centerCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            numberCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            centerRange = centerCell.Range;
            numberRange = numberCell.Range;
            centerFormat = centerRange.ParagraphFormat;
            numberFormat = numberRange.ParagraphFormat;
            centerFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            numberFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            centerFormat.LeftIndent = centerFormat.RightIndent = 0f;
            centerFormat.FirstLineIndent = 0f;
            numberFormat.LeftIndent = numberFormat.RightIndent = 0f;
            numberFormat.FirstLineIndent = 0f;
            centerFormat.SpaceBefore = centerFormat.SpaceAfter = 0f;
            numberFormat.SpaceBefore = numberFormat.SpaceAfter = 0f;
        }
        finally
        {
            Release(rightColumn);
            Release(centerColumn);
            Release(leftColumn);
            Release(columns);
            Release(borders);
            Release(numberFormat);
            Release(centerFormat);
            Release(numberRange);
            Release(centerRange);
            Release(numberCell);
            Release(centerCell);
            Release(table);
        }
    }

    private static void ConfigureUnnumberedDisplayFormula(
        Document document,
        Range formulaRange,
        string formulaId)
    {
        RemoveVisibleEquationNumber(document, formulaId);
        RemoveNativeCaption(document, formulaId);
        RemoveLeadingEquationTab(document, formulaRange);
        ConfigureEquationParagraph(formulaRange, numbered: false);
    }

    private static void ConfigureEquationParagraph(Range formulaRange, bool numbered)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        ListFormat? listFormat = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            format.KeepTogether = 0;
            format.KeepWithNext = 0;
            format.PageBreakBefore = 0;
            format.WidowControl = 0;
            try
            {
                listFormat = paragraphRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch
            {
                // Protected/custom stories can reject list normalization. The
                // page-break flags above still remove Word's black paragraph
                // marker when formatting marks are shown.
            }
            tabStops = format.TabStops;
            tabStops.ClearAll();

            if (!numbered)
            {
                format.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                return;
            }

            var pageWidth = 612f;
            var leftMargin = 72f;
            var rightMargin = 72f;
            try
            {
                sections = paragraphRange.Sections;
                if (sections.Count > 0)
                {
                    section = sections[1];
                    pageSetup = section.PageSetup;
                    pageWidth = pageSetup.PageWidth;
                    leftMargin = pageSetup.LeftMargin;
                    rightMargin = pageSetup.RightMargin;
                }
            }
            catch
            {
                // Standard US Letter and one-inch margins are a safe fallback
                // for protected/custom stories without an exposed PageSetup.
            }

            var positions = CalculateEquationTabStops(pageWidth, leftMargin, rightMargin, 0, 0);
            format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
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
            Release(listFormat);
            Release(tabStops);
            Release(format);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
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

    private static void EnsureLeadingEquationTab(Document document, Range formulaRange)
    {
        Range? preceding = null;
        Range? insertion = null;
        try
        {
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

    private static void RemoveLeadingEquationTab(Document document, Range formulaRange)
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

    private static void EnsureNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName)
    {
        if (TryGetNativeCaptionRanges(
                document,
                formulaId,
                nativeSequenceName,
                out var captionRange,
                out var numberRange))
        {
            try { StyleNativeCaption(captionRange, numberRange); }
            finally
            {
                Release(numberRange);
                Release(captionRange);
            }
            return;
        }
        Release(numberRange);
        Release(captionRange);

        RemoveNativeCaption(document, formulaId);
        CreateNativeCaption(document, formulaRange, formulaId, nativeSequenceName);
    }

    private static void CreateNativeCaption(
        Document document,
        Range formulaRange,
        string formulaId,
        string nativeSequenceName)
    {
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Table? formulaTable = null;
        Range? formulaTableRange = null;
        Range? fieldInsertion = null;
        Fields? fields = null;
        Field? captionField = null;
        Range? numberRange = null;
        Range? captionRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Bookmarks? bookmarks = null;
        try
        {
            // Word's Range.InsertCaption mutates the equation paragraph. For a
            // trailing inline OMath it moves the ordinary run after the equation
            // into a new paragraph, so the visible REF field is subsequently
            // absorbed into m:oMath. Build the native SEQ caption in a dedicated
            // hidden paragraph instead and leave the equation paragraph intact.
            int captionStart;
            if (IsNumberedEquationTable(formulaRange))
            {
                formulaTable = formulaRange.Tables[1];
                formulaTableRange = formulaTable.Range;
                formulaTableRange.InsertParagraphAfter();
                // InsertParagraphAfter expands the table range so its new End
                // is outside the table. The original End (and new End - 1) are
                // structural cell boundaries that reject Fields.Add.
                captionStart = formulaTableRange.End;
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                captionStart = formulaParagraphRange.End;
                formulaParagraphRange.InsertParagraphAfter();
            }

            object insertionStart = captionStart;
            object insertionEnd = captionStart;
            fieldInsertion = document.Range(ref insertionStart, ref insertionEnd);
            fields = document.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"SEQ {nativeSequenceName} \\* ARABIC";
            object preserveFormatting = true;
            captionField = fields.Add(
                fieldInsertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            captionField.Update();
            numberRange = captionField.Result;
            paragraphs = numberRange.Paragraphs;
            paragraph = paragraphs[1];
            captionRange = paragraph.Range;
            try
            {
                object captionStyle = WdBuiltinStyle.wdStyleCaption;
                captionRange.set_Style(ref captionStyle);
            }
            catch
            {
                // Some locked/custom documents reject assigning the built-in
                // caption style. The SEQ field and bookmarks remain valid.
            }

            bookmarks = document.Bookmarks;
            bookmarks.Add(NativeNumberBookmarkName(formulaId), numberRange);
            bookmarks.Add(NativeCaptionBookmarkName(formulaId), captionRange);
            StyleNativeCaption(captionRange, numberRange);
        }
        finally
        {
            Release(bookmarks);
            Release(paragraph);
            Release(paragraphs);
            Release(captionRange);
            Release(numberRange);
            Release(captionField);
            Release(fields);
            Release(fieldInsertion);
            Release(formulaTableRange);
            Release(formulaTable);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
        }
    }

    private static Field? FindNewNativeEquationField(
        Document document,
        string nativeSequenceName,
        ISet<int> existingPositions,
        int formulaPosition)
    {
        Fields? fields = null;
        Field? result = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? fieldResult = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    fieldResult = field.Result;
                    if (existingPositions.Contains(fieldResult.Start)) continue;
                    var distance = Math.Abs(fieldResult.Start - formulaPosition);
                    if (distance >= bestDistance) continue;
                    Release(result);
                    result = field;
                    field = null;
                    bestDistance = distance;
                }
                finally
                {
                    Release(fieldResult);
                    Release(code);
                    Release(field);
                }
            }
            return result;
        }
        finally
        {
            Release(fields);
        }
    }

    private static void StyleNativeCaption(Range captionRange, Range numberRange)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        ParagraphFormat? paragraph = null;
        ListFormat? listFormat = null;
        Frames? frames = null;
        Frame? frame = null;
        Borders? borders = null;
        try
        {
            // Keep the native SEQ field at a real body-text size. Word copies
            // source-field formatting into REF results during several update
            // paths, so shrinking the SEQ glyph itself to one point makes the
            // visible cross-reference one point as well.
            var fontSize = ResolveNormalStyleFontSize(captionRange);
            font = captionRange.Font;
            font.Hidden = 0;
            font.Size = fontSize;
            font.Color = WdColor.wdColorAutomatic;
            font.Position = 0;
            numberFont = numberRange.Font;
            numberFont.Hidden = 0;
            numberFont.Size = fontSize;
            numberFont.Color = WdColor.wdColorAutomatic;
            numberFont.Position = 0;

            paragraph = captionRange.ParagraphFormat;
            paragraph.SpaceBefore = 0f;
            paragraph.SpaceAfter = 0f;
            paragraph.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            paragraph.LineSpacing = NativeCaptionFrameSizePoints;
            paragraph.KeepTogether = 0;
            paragraph.KeepWithNext = 0;
            paragraph.PageBreakBefore = 0;
            paragraph.WidowControl = 0;
            try
            {
                listFormat = captionRange.ListFormat;
                listFormat.RemoveNumbers(WdNumberType.wdNumberParagraph);
            }
            catch { }

            // A Word Frame remains in the main document story, so native
            // caption enumeration and InsertCrossReference can still find the
            // SEQ field. The one-point frame is moved far outside the page;
            // only the container is tiny and invisible, not the SEQ formatting.
            frames = captionRange.Frames;
            frame = frames.Count > 0 ? frames[1] : frames.Add(captionRange);
            frame.WidthRule = WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = WdFrameSizeRule.wdFrameExact;
            frame.Width = NativeCaptionFrameSizePoints;
            frame.Height = NativeCaptionFrameSizePoints;
            frame.RelativeHorizontalPosition =
                WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
            frame.RelativeVerticalPosition =
                WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
            frame.HorizontalPosition = NativeCaptionOffscreenPositionPoints;
            frame.VerticalPosition = NativeCaptionOffscreenPositionPoints;
            frame.HorizontalDistanceFromText = 0f;
            frame.VerticalDistanceFromText = 0f;
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = frame.Borders;
            borders.Enable = 0;
        }
        finally
        {
            Release(borders);
            Release(frame);
            Release(frames);
            Release(listFormat);
            Release(paragraph);
            Release(numberFont);
            Release(font);
        }
    }

    private static float ResolveNormalStyleFontSize(Range range)
    {
        Styles? styles = null;
        Style? normalStyle = null;
        Microsoft.Office.Interop.Word.Font? normalFont = null;
        try
        {
            styles = range.Document.Styles;
            object normalStyleIndex = WdBuiltinStyle.wdStyleNormal;
            normalStyle = styles[ref normalStyleIndex];
            normalFont = normalStyle.Font;
            var size = normalFont.Size;
            return IsNormalTextSize(size) ? size : 11f;
        }
        catch
        {
            return 11f;
        }
        finally
        {
            Release(normalFont);
            Release(normalStyle);
            Release(styles);
        }
    }

    internal static void RestoreTypingAfterDisplayFormula(
        Document document,
        Range formulaRange,
        string formulaId,
        Selection selection,
        float preferredFontSize)
    {
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Table? table = null;
        Range? anchor = null;
        Range? documentRange = null;
        Range? continuation = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Microsoft.Office.Interop.Word.Font? paragraphFont = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        ParagraphFormat? paragraphFormat = null;
        ParagraphFormat? selectionParagraphFormat = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            if (bookmarks.Exists(captionName))
            {
                captionBookmark = bookmarks[captionName];
                anchor = captionBookmark.Range;
            }
            else if ((bool)formulaRange.get_Information(WdInformation.wdWithInTable)
                     && formulaRange.Tables.Count > 0)
            {
                table = formulaRange.Tables[1];
                anchor = table.Range;
            }
            else
            {
                paragraphs = formulaRange.Paragraphs;
                paragraph = paragraphs[1];
                anchor = paragraph.Range;
            }

            documentRange = document.Content;
            var continuationStart = Math.Min(anchor.End, documentRange.End);
            if (continuationStart >= documentRange.End - 1)
            {
                var insertionStart = anchor.End;
                anchor.InsertParagraphAfter();
                continuationStart = insertionStart;
            }

            object start = continuationStart;
            object end = continuationStart;
            continuation = document.Range(ref start, ref end);
            Release(paragraphs);
            paragraphs = continuation.Paragraphs;
            Release(paragraph);
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var visibleText = (paragraphRange.Text ?? string.Empty)
                .Trim('\r', '\a', '\v', '\t', ' ');
            var fontSize = IsNormalTextSize(preferredFontSize)
                ? preferredFontSize
                : 11f;

            if (visibleText.Length == 0)
            {
                try
                {
                    object normalStyle = WdBuiltinStyle.wdStyleNormal;
                    paragraphRange.set_Style(ref normalStyle);
                }
                catch { }
                paragraphFormat = paragraphRange.ParagraphFormat;
                paragraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                paragraphFont = paragraphRange.Font;
                paragraphFont.Hidden = 0;
                paragraphFont.Color = WdColor.wdColorAutomatic;
                paragraphFont.Size = fontSize;
                paragraphFont.Position = 0;
            }

            selection.SetRange(continuationStart, continuationStart);
            selectionFont = selection.Font;
            selectionFont.Hidden = 0;
            selectionFont.Color = WdColor.wdColorAutomatic;
            selectionFont.Size = fontSize;
            selectionFont.Position = 0;
            if (visibleText.Length == 0)
            {
                selectionParagraphFormat = selection.ParagraphFormat;
                selectionParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }
        finally
        {
            Release(selectionParagraphFormat);
            Release(selectionFont);
            Release(paragraphFont);
            Release(paragraphFormat);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(continuation);
            Release(documentRange);
            Release(anchor);
            Release(table);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static bool TryGetNativeCaptionRanges(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out Range? captionRange,
        out Range? numberRange)
    {
        captionRange = null;
        numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? captionBookmark = null;
        Bookmark? numberBookmark = null;
        Field? nativeField = null;
        try
        {
            bookmarks = document.Bookmarks;
            var captionName = NativeCaptionBookmarkName(formulaId);
            var numberName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(captionName) || !bookmarks.Exists(numberName)) return false;
            captionBookmark = bookmarks[captionName];
            numberBookmark = bookmarks[numberName];
            captionRange = captionBookmark.Range;
            numberRange = numberBookmark.Range;
            nativeField = FindNativeEquationFieldAtRange(
                document,
                numberRange,
                nativeSequenceName);
            if (nativeField is not null) return true;
            Release(numberRange);
            numberRange = null;
            Release(captionRange);
            captionRange = null;
            return false;
        }
        finally
        {
            Release(nativeField);
            Release(numberBookmark);
            Release(captionBookmark);
            Release(bookmarks);
        }
    }

    private static void EnsureVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        string formulaId)
    {
        var targetBookmarkName = NativeNumberBookmarkName(formulaId);
        if (HasVisibleEquationNumber(
                document,
                formulaRange,
                formulaId,
                targetBookmarkName)) return;
        RemoveVisibleEquationNumber(document, formulaId);
        InsertVisibleEquationNumber(
            document,
            formulaRange,
            formulaHeightPoints,
            formulaId,
            targetBookmarkName);
    }

    private static bool HasVisibleEquationNumber(
        Document document,
        Range formulaRange,
        string formulaId,
        string targetBookmarkName)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        OMaths? maths = null;
        Paragraphs? formulaParagraphs = null;
        Paragraph? formulaParagraph = null;
        Range? formulaParagraphRange = null;
        Paragraphs? numberParagraphs = null;
        Paragraph? numberParagraph = null;
        Range? numberParagraphRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return false;
            bookmark = bookmarks[name];
            range = bookmark.Range;

            // Older builds inserted the visible REF field at OMath.Range.End.
            // Word still considers that position part of m:oMath, so the tab,
            // parentheses and number became equation content. A valid number is
            // an ordinary Word run after the formula, never a child of OMML.
            if (range.Start < formulaRange.End) return false;
            maths = range.OMaths;
            if (maths.Count > 0) return false;
            var tableLayout = IsNumberedEquationTable(formulaRange);
            var visibleText = range.Text ?? string.Empty;
            var expectedPrefix = tableLayout ? "(" : "\t(";
            if (!visibleText.StartsWith(expectedPrefix, StringComparison.Ordinal)
                || !visibleText.EndsWith(")", StringComparison.Ordinal))
                return false;

            if (tableLayout)
            {
                if (!(bool)range.get_Information(WdInformation.wdWithInTable)
                    || range.Tables.Count == 0
                    || range.Tables[1].Range.Start != formulaRange.Tables[1].Range.Start)
                    return false;
            }
            else
            {
                formulaParagraphs = formulaRange.Paragraphs;
                formulaParagraph = formulaParagraphs[1];
                formulaParagraphRange = formulaParagraph.Range;
                numberParagraphs = range.Paragraphs;
                numberParagraph = numberParagraphs[1];
                numberParagraphRange = numberParagraph.Range;
                if (formulaParagraphRange.Start != numberParagraphRange.Start) return false;
            }

            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (IsReferenceToBookmark(code.Text, targetBookmarkName)) return true;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return false;
        }
        finally
        {
            Release(numberParagraphRange);
            Release(numberParagraph);
            Release(numberParagraphs);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(maths);
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void InsertVisibleEquationNumber(
        Document document,
        Range formulaRange,
        float formulaHeightPoints,
        string formulaId,
        string targetBookmarkName)
    {
        Range? scaffoldRange = null;
        Range? fieldRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? fieldResult = null;
        Range? bookmarkRange = null;
        Bookmarks? bookmarks = null;
        try
        {
            var tableLayout = IsNumberedEquationTable(formulaRange);
            var suffixStart = PrepareEquationNumberInsertionPosition(formulaRange);
            var scaffold = tableLayout ? (Text: "()", FieldOffset: 1) : EquationNumberScaffold();
            object suffixStartObject = suffixStart;
            object suffixEndObject = suffixStart;
            scaffoldRange = document.Range(ref suffixStartObject, ref suffixEndObject);
            scaffoldRange.Text = scaffold.Text;

            var fieldStart = suffixStart + scaffold.FieldOffset;
            object fieldStartObject = fieldStart;
            object fieldEndObject = fieldStart;
            fieldRange = document.Range(ref fieldStartObject, ref fieldEndObject);
            fields = document.Fields;
            object fieldType = WdFieldEmpty;
            object fieldCode = $"REF {targetBookmarkName} \\h";
            object preserveFormatting = true;
            field = fields.Add(
                fieldRange,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            field.Update();
            fieldResult = field.Result;

            // Inserting a field inside scaffoldRange expands that Range to
            // include the complete field and the closing parenthesis. Using
            // field.Result.End misses the closing parenthesis because Word's
            // hidden field-end character sits between them.
            object labelStartObject = suffixStart;
            object labelEndObject = scaffoldRange.End;
            bookmarkRange = document.Range(ref labelStartObject, ref labelEndObject);
            bookmarks = document.Bookmarks;
            bookmarks.Add(EquationBookmarkName(formulaId), bookmarkRange);
            if (tableLayout)
            {
                ParagraphFormat? format = null;
                try
                {
                    format = bookmarkRange.ParagraphFormat;
                    format.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                }
                finally { Release(format); }
            }
            AlignEquationNumberVertically(
                bookmarkRange,
                tableLayout ? 0f : formulaHeightPoints);
        }
        finally
        {
            Release(bookmarks);
            Release(bookmarkRange);
            Release(fieldResult);
            Release(field);
            Release(fields);
            Release(fieldRange);
            Release(scaffoldRange);
        }
    }

    private static int PrepareEquationNumberInsertionPosition(Range formulaRange)
    {
        if (IsNumberedEquationTable(formulaRange))
        {
            Table? table = null;
            Cell? cell = null;
            Range? cellRange = null;
            Range? editableRange = null;
            try
            {
                table = formulaRange.Tables[1];
                cell = table.Cell(1, 3);
                cellRange = cell.Range;
                // This cell is reserved exclusively for the generated number.
                // Word can leave an empty paragraph behind when a REF field is
                // removed/reconciled. Centering that empty paragraph together
                // with the new number pushes the visible number downward.
                // Clear everything except the structural cell mark, which
                // normalizes the cell to exactly one paragraph, then insert at
                // the beginning of that paragraph.
                editableRange = cellRange.Duplicate;
                editableRange.End = Math.Max(
                    editableRange.Start,
                    editableRange.End - 1);
                editableRange.Text = string.Empty;
                return cellRange.Start;
            }
            finally
            {
                Release(editableRange);
                Release(cellRange);
                Release(cell);
                Release(table);
            }
        }
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return Math.Max(paragraphRange.Start, paragraphRange.End - 1);
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RemoveVisibleEquationNumber(Document document, string formulaId)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Range? trailing = null;
        OMaths? maths = null;
        OMath? containingMath = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            var start = range.Start;
            var text = range.Text ?? string.Empty;
            try
            {
                maths = range.OMaths;
                if (maths.Count > 0) containingMath = maths[1];
            }
            catch { }
            range.Delete();

            // Legacy OMML numbering bookmarks stopped immediately before the
            // closing parenthesis. Remove that orphan as part of migration.
            if (!text.EndsWith(")", StringComparison.Ordinal)
                && start < document.Content.End)
            {
                object trailingStart = start;
                object trailingEnd = Math.Min(document.Content.End, start + 1);
                trailing = document.Range(ref trailingStart, ref trailingEnd);
                if (string.Equals(trailing.Text, ")", StringComparison.Ordinal))
                    trailing.Delete();
            }
            try { containingMath?.BuildUp(); } catch { }
        }
        finally
        {
            Release(containingMath);
            Release(maths);
            Release(trailing);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void RemoveNativeCaption(Document document, string formulaId)
    {
        DeleteBookmarkOnly(document, NativeNumberBookmarkName(formulaId));
        DeleteBookmarkedRange(document, NativeCaptionBookmarkName(formulaId));
    }

    private static void RemoveOrphanEquationArtifacts(
        Document document,
        ISet<string> numberedFormulaIds)
    {
        RemoveOrphanBookmarks(
            document,
            EquationBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeCaptionBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: true);
        RemoveOrphanBookmarks(
            document,
            NativeNumberBookmarkPrefix,
            numberedFormulaIds,
            deleteRange: false);
    }

    private static void RemoveOrphanBookmarks(
        Document document,
        string prefix,
        ISet<string> activeFormulaIds,
        bool deleteRange)
    {
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = bookmarks.Count; index >= 1; index--)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (!TryFormulaIdFromBookmark(bookmark.Name, prefix, out var formulaId)
                        || activeFormulaIds.Contains(formulaId))
                        continue;
                    if (deleteRange)
                    {
                        range = bookmark.Range;
                        range.Delete();
                    }
                    else
                    {
                        bookmark.Delete();
                    }
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(bookmarks); }
    }

    private static void UpdateEquationNumberFields(
        Document document,
        float formulaHeightPoints,
        string formulaId)
    {
        UpdateFieldInBookmark(
            document,
            EquationBookmarkName(formulaId),
            code => IsReferenceToBookmark(code, NativeNumberBookmarkName(formulaId)));

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var visibleName = EquationBookmarkName(formulaId);
            if (!bookmarks.Exists(visibleName)) return;
            bookmark = bookmarks[visibleName];
            range = bookmark.Range;
            // A numbered table centers both cells vertically. Applying the
            // legacy height-derived baseline shift as well makes OLE and OMML
            // numbers disagree because their measured heights differ.
            AlignEquationNumberVertically(
                range,
                IsNumberedEquationTable(range) ? 0f : formulaHeightPoints);
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void UpdateNativeEquationSequenceFields(Document document)
    {
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        Fields? fields = null;
        var sequenceFields = new List<(int Index, int Position)>();
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName))
                        continue;
                    result = field.Result;
                    sequenceFields.Add((index, result.Start));
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }

            var orderedFields = sequenceFields
                .OrderBy(item => item.Position)
                .ToList();
            for (var ordinal = 0; ordinal < orderedFields.Count; ordinal++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[orderedFields[ordinal].Index];
                    code = field.Code;
                    // Force each surviving caption to its current document-order
                    // ordinal. Plain Field.Update can retain a cached SEQ value
                    // after a formula was manually deleted, especially when the
                    // deleted OMML bookmark survived as a collapsed anchor.
                    code.Text =
                        $" SEQ {nativeSequenceName} \\r {ordinal + 1} \\* ARABIC ";
                    field.Update();
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
    }

    private static void UpdateFieldInBookmark(
        Document document,
        string bookmarkName,
        Func<string?, bool> predicate)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Fields? fields = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName)) return;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            fields = range.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (predicate(code.Text)) field.Update();
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally
        {
            Release(fields);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void AlignEquationNumberVertically(
        Range numberRange,
        float formulaHeightPoints)
    {
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            font = numberRange.Font;
            var numberFontSize = 11f;
            try { numberFontSize = font.Size; } catch { }
            if (float.IsNaN(numberFontSize)
                || float.IsInfinity(numberFontSize)
                || numberFontSize <= 2f
                || numberFontSize > 256f)
                numberFontSize = 11f;

            // The native caption target is deliberately white and one point.
            // Word propagates that appearance into REF results unless the
            // visible range is normalized after every field update.
            font.Hidden = 0;
            font.Color = WdColor.wdColorAutomatic;
            font.Size = numberFontSize;
            font.Position = CalculateEquationNumberFontPosition(
                formulaHeightPoints,
                numberFontSize);
        }
        finally { Release(font); }
    }

    internal static IReadOnlyList<EquationReferenceTarget> GetEquationReferenceTargets(
        Document document)
    {
        Reconcile(document);
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        var nativeFieldPositions = GetNativeEquationFieldPositions(document, nativeSequenceName);
        var nativeItems = document.GetCrossReferenceItems(WdCaptionLabelID.wdCaptionEquation) as Array;
        if (nativeItems is null || nativeItems.Length == 0)
            return Array.Empty<EquationReferenceTarget>();

        var targets = new List<EquationReferenceTarget>();
        void AddTarget(FormulaMetadata metadata, int position)
        {
            if (metadata.DisplayMode != "block" || !metadata.Numbered) return;
            if (!TryGetNativeCaptionInfo(
                    document,
                    metadata.FormulaId,
                    nativeSequenceName,
                    out var fieldPosition,
                    out var numberText))
                return;
            var nativeOrdinal = nativeFieldPositions.IndexOf(fieldPosition);
            if (nativeOrdinal < 0 || nativeOrdinal >= nativeItems.Length) return;
            var latex = string.Join(" ", metadata.Lines.Select(line => line.Latex))
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            if (latex.Length > 90) latex = latex.Substring(0, 87) + "…";
            targets.Add(new EquationReferenceTarget(
                metadata.FormulaId,
                nativeOrdinal + 1,
                numberText,
                latex,
                position));
        }

        var ommlFormulaIds = WordOmmlFormulaStore.FormulaIds(document);
        InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = document.InlineShapes;
            for (var index = 1; index <= inlineShapes.Count; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    shape = inlineShapes[index];
                    var metadata = ReadMetadata(shape);
                    if (metadata is null) continue;
                    range = shape.Range;
                    AddTarget(metadata, range.Start);
                }
                finally
                {
                    Release(range);
                    Release(shape);
                }
            }

            foreach (var formulaId in ommlFormulaIds)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                    if (bookmark is null) continue;
                    var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                    if (metadata is null) continue;
                    range = bookmark.Range;
                    AddTarget(metadata, range.Start);
                }
                finally
                {
                    Release(range);
                    Release(bookmark);
                }
            }
        }
        finally { Release(inlineShapes); }
        return targets.OrderBy(target => target.Position).ToArray();
    }

    internal static void InsertEquationReference(
        Document document,
        Selection selection,
        EquationReferenceTarget target,
        EquationReferenceStyle style)
    {
        var prefix = style switch
        {
            EquationReferenceStyle.EquationPrefix => "式（",
            EquationReferenceStyle.Parenthesized => "(",
            _ => string.Empty,
        };
        var suffix = style switch
        {
            EquationReferenceStyle.EquationPrefix => "）",
            EquationReferenceStyle.Parenthesized => ")",
            _ => string.Empty,
        };

        if (!string.IsNullOrEmpty(prefix)) selection.TypeText(prefix);
        selection.InsertCrossReference(
            ReferenceType: WdCaptionLabelID.wdCaptionEquation,
            ReferenceKind: WdReferenceKind.wdEntireCaption,
            ReferenceItem: target.NativeReferenceItem,
            InsertAsHyperlink: true,
            IncludePosition: false);
        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        if (!string.IsNullOrEmpty(suffix)) selection.TypeText(suffix);
        UpdateNativeCrossReferences(document);
    }

    internal static int UpdateNativeCrossReferences(Document document)
    {
        // Migrate captions created by older builds before updating any REF.
        // Those captions stored a one-point SEQ directly in the body. Keeping
        // this migration on the refresh path repairs existing documents even
        // when the user does not edit or reinsert the original formula.
        NormalizeNativeCaptionFrames(document);

        Fields? fields = null;
        var updated = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    field.Update();
                    NormalizeReferenceResult(field);
                    updated++;
                }
                finally { Release(field); }
            }
        }
        finally { Release(fields); }
        return updated;
    }

    private static void NormalizeNativeCaptionFrames(Document document)
    {
        var formulaIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Bookmarks? bookmarks = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                try
                {
                    bookmark = bookmarks[index];
                    if (TryFormulaIdFromBookmark(
                            bookmark.Name,
                            NativeCaptionBookmarkPrefix,
                            out var formulaId))
                        formulaIds.Add(formulaId);
                }
                finally { Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }

        if (formulaIds.Count == 0) return;
        var nativeSequenceName = GetNativeEquationSequenceName(document);
        foreach (var formulaId in formulaIds)
        {
            Range? captionRange = null;
            Range? numberRange = null;
            try
            {
                if (!TryGetNativeCaptionRanges(
                        document,
                        formulaId,
                        nativeSequenceName,
                        out captionRange,
                        out numberRange))
                    continue;
                StyleNativeCaption(captionRange, numberRange);
            }
            catch
            {
                // A damaged or protected caption must not prevent unrelated
                // cross-references elsewhere in the document from refreshing.
            }
            finally
            {
                Release(numberRange);
                Release(captionRange);
            }
        }
    }

    private static void NormalizeReferenceResult(Field field)
    {
        Range? code = null;
        Range? result = null;
        Microsoft.Office.Interop.Word.Font? codeFont = null;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
        try
        {
            result = field.Result;
            var size = ResolveReferenceFontSize(result);
            code = field.Code;
            var codeText = code.Text ?? string.Empty;
            if (codeText.IndexOf("CHARFORMAT", StringComparison.OrdinalIgnoreCase) < 0)
                code.Text = codeText.TrimEnd() + " \\* CHARFORMAT ";

            // CHARFORMAT makes Word reuse the field-code formatting whenever
            // the REF is refreshed. This prevents the hidden one-point SEQ
            // caption from leaking back into the visible reference result.
            codeFont = code.Font;
            codeFont.Hidden = 0;
            codeFont.Color = WdColor.wdColorAutomatic;
            codeFont.Size = size;
            field.Update();

            Release(result);
            result = field.Result;
            resultFont = result.Font;
            resultFont.Hidden = 0;
            resultFont.Color = WdColor.wdColorAutomatic;
            resultFont.Size = size;
            resultFont.Position = 0;
        }
        finally
        {
            Release(resultFont);
            Release(codeFont);
            Release(result);
            Release(code);
        }
    }

    private static float ResolveReferenceFontSize(Range result)
    {
        Range? neighbor = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            foreach (var bounds in new[]
            {
                (Start: Math.Max(0, result.Start - 1), End: result.Start),
                (Start: result.End, End: result.End + 1),
            })
            {
                if (bounds.End <= bounds.Start) continue;
                try
                {
                    neighbor = result.Duplicate;
                    neighbor.SetRange(bounds.Start, bounds.End);
                    var text = neighbor.Text ?? string.Empty;
                    if (text is "\r" or "\a" or "\v") continue;
                    font = neighbor.Font;
                    var neighborSize = font.Size;
                    if (IsNormalTextSize(neighborSize)) return neighborSize;
                }
                catch { }
                finally
                {
                    Release(font);
                    font = null;
                    Release(neighbor);
                    neighbor = null;
                }
            }

            font = result.Font;
            var resultSize = font.Size;
            return IsNormalTextSize(resultSize) ? resultSize : 11f;
        }
        catch
        {
            return 11f;
        }
        finally
        {
            Release(font);
            Release(neighbor);
        }
    }

    internal static bool IsNormalTextSize(float size) =>
        !float.IsNaN(size)
        && !float.IsInfinity(size)
        && size > 2f
        && size <= 256f;

    private static List<int> GetNativeEquationFieldPositions(
        Document document,
        string nativeSequenceName)
    {
        var positions = new List<int>();
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    positions.Add(result.Start);
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }
        positions.Sort();
        return positions;
    }

    private static bool TryGetNativeCaptionInfo(
        Document document,
        string formulaId,
        string nativeSequenceName,
        out int position,
        out string numberText)
    {
        position = -1;
        numberText = string.Empty;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        Field? field = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            var bookmarkName = NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            range = bookmark.Range;
            field = FindNativeEquationFieldAtRange(document, range, nativeSequenceName);
            if (field is null) return false;
            field.Update();
            result = field.Result;
            position = result.Start;
            numberText = (result.Text ?? string.Empty).Trim();
            return !string.IsNullOrWhiteSpace(numberText);
        }
        finally
        {
            Release(result);
            Release(field);
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static Field? FindNativeEquationFieldAtRange(
        Document document,
        Range targetRange,
        string nativeSequenceName)
    {
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    code = field.Code;
                    if (!IsNativeEquationSequenceFieldCode(code.Text, nativeSequenceName)) continue;
                    result = field.Result;
                    var overlaps = result.Start < targetRange.End
                        && result.End > targetRange.Start;
                    var sameCollapsedPosition = result.Start == targetRange.Start
                        && result.End == targetRange.End;
                    if (!overlaps && !sameCollapsedPosition) continue;
                    var found = field;
                    field = null;
                    return found;
                }
                finally
                {
                    Release(result);
                    Release(code);
                    Release(field);
                }
            }
            return null;
        }
        finally { Release(fields); }
    }

    private static string GetNativeEquationSequenceName(Document document)
    {
        Microsoft.Office.Interop.Word.Application? application = null;
        CaptionLabels? labels = null;
        CaptionLabel? label = null;
        try
        {
            application = document.Application;
            labels = application.CaptionLabels;
            label = labels[WdCaptionLabelID.wdCaptionEquation];
            var name = label.Name;
            if (string.IsNullOrWhiteSpace(name))
                throw new InvalidOperationException("Word built-in Equation caption label is unavailable.");
            return name;
        }
        finally
        {
            Release(label);
            Release(labels);
            Release(application);
        }
    }

    internal static string EquationBookmarkName(string formulaId) =>
        BookmarkName(EquationBookmarkPrefix, formulaId);

    internal static string NativeCaptionBookmarkName(string formulaId) =>
        BookmarkName(NativeCaptionBookmarkPrefix, formulaId);

    internal static string NativeNumberBookmarkName(string formulaId) =>
        BookmarkName(NativeNumberBookmarkPrefix, formulaId);

    private static string BookmarkName(string prefix, string formulaId)
    {
        if (!Guid.TryParse(formulaId, out var value))
            throw new InvalidOperationException("VisualTeX formulaId must be a UUID.");
        return $"{prefix}{value:N}";
    }

    internal static bool TryFormulaIdFromEquationBookmark(
        string? bookmarkName,
        out string formulaId) =>
        TryFormulaIdFromBookmark(bookmarkName, EquationBookmarkPrefix, out formulaId);

    private static bool TryFormulaIdFromBookmark(
        string? bookmarkName,
        string prefix,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(bookmarkName)) return false;
        var name = bookmarkName!;
        if (!name.StartsWith(prefix, StringComparison.Ordinal)
            || !Guid.TryParseExact(name.Substring(prefix.Length), "N", out var value))
            return false;
        formulaId = value.ToString();
        return true;
    }

    internal static bool IsVisualTeXSequenceFieldCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"SEQ {LegacyEquationSequenceName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    internal static bool IsNativeEquationSequenceFieldCode(
        string? code,
        string nativeSequenceName)
    {
        if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(nativeSequenceName))
            return false;
        return code!.IndexOf(
                   $"SEQ {nativeSequenceName}",
                   StringComparison.OrdinalIgnoreCase) >= 0
            || code.IndexOf(
                   $"SEQ \"{nativeSequenceName}\"",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool IsReferenceToBookmark(string? code, string bookmarkName) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            $"REF {bookmarkName}",
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static void DeleteBookmarkedRange(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            range = bookmark.Range;
            range.Delete();
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void DeleteBookmarkOnly(Document document, string name)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(name)) return;
            bookmark = bookmarks[name];
            bookmark.Delete();
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

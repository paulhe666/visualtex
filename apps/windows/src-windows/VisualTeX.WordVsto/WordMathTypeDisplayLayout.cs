using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordFormulaLatexSource;

namespace VisualTeX.WordVsto;

/// <summary>
/// Laying out MathType display formulas: MathType's native styles, display tab geometry and paragraph format, the
/// column width, number fields, and repairs after InsertXML.
/// </summary>
internal static class WordMathTypeDisplayLayout
{
    internal static void EnsureMathTypeNativeStyles(
        Document document,
        WordCharacterFormatting? displayBodyFormatting = null)
    {
        Styles? styles = null;
        Style? displayStyle = null;
        Style? sectionStyle = null;
        ParagraphFormat? displayFormat = null;
        TabStops? displayTabs = null;
        TabStop? tab = null;
        PageSetup? pageSetup = null;
        Microsoft.Office.Interop.Word.Font? sectionFont = null;
        Microsoft.Office.Interop.Word.Font? displayFont = null;
        try
        {
            styles = document.Styles;
            // Section-state initialization must not pre-create an incomplete
            // display style. Native MathType seeds that style from the first
            // insertion's body formatting and reuses it for subsequent equations.
            if (displayBodyFormatting is not null)
            {
                object displayName = "MTDisplayEquation";
                try { displayStyle = styles.get_Item(ref displayName); }
                catch
                {
                    object paragraphType = WdStyleType.wdStyleTypeParagraph;
                    displayStyle = styles.Add("MTDisplayEquation", ref paragraphType);
                    object normalStyle = WdBuiltinStyle.wdStyleNormal;
                    displayStyle.set_BaseStyle(ref normalStyle);
                    displayStyle.set_NextParagraphStyle(ref normalStyle);
                    displayFont = displayStyle.Font;
                    displayBodyFormatting.ApplyToParagraphStyle(displayFont);
                    // Match MathType's own behavior: create the style only when the
                    // document does not already have it. Never overwrite an existing
                    // native MathType style, because its tab geometry may have been
                    // adapted by MathType for the current document/template.
                    displayFormat = displayStyle.ParagraphFormat;
                    displayFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    displayFormat.SpaceBefore = 0;
                    displayFormat.SpaceAfter = 0;
                    displayFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    displayTabs = displayFormat.TabStops;
                    displayTabs.ClearAll();
                    pageSetup = document.PageSetup;
                    var usableWidth = Math.Max(
                        72f,
                        pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
                    tab = displayTabs.Add(
                        usableWidth / 2f,
                        WdTabAlignment.wdAlignTabCenter,
                        WdTabLeader.wdTabLeaderSpaces);
                    Release(tab);
                    tab = displayTabs.Add(
                        usableWidth,
                        WdTabAlignment.wdAlignTabRight,
                        WdTabLeader.wdTabLeaderSpaces);
                    Release(tab);
                    tab = null;
                }
            }

            object sectionName = "MTEquationSection";
            try { sectionStyle = styles.get_Item(ref sectionName); }
            catch
            {
                object characterType = WdStyleType.wdStyleTypeCharacter;
                sectionStyle = styles.Add("MTEquationSection", ref characterType);
            }
            sectionFont = sectionStyle.Font;
            sectionFont.Hidden = -1;
            sectionFont.Color = WdColor.wdColorRed;
        }
        finally
        {
            Release(displayFont);
            Release(sectionFont);
            Release(pageSetup);
            Release(tab);
            Release(displayTabs);
            Release(displayFormat);
            Release(sectionStyle);
            Release(displayStyle);
            Release(styles);
        }
    }

    private static int FinalizeMathTypeRedrawDisplayLayouts(
        Document document,
        IReadOnlyList<(string FormulaId, Range Range, string Signature)> identities,
        ISet<string> displayFormulaIds)
    {
        var normalized = 0;
        foreach (var identity in identities)
        {
            if (!displayFormulaIds.Contains(identity.FormulaId)) continue;
            InlineShapes? shapes = null;
            InlineShape? shape = null;
            try
            {
                shapes = identity.Range.InlineShapes;
                if (shapes.Count != 1)
                    throw new InvalidDataException(
                        $"MathType redraw display {identity.FormulaId} no longer owns exactly one OLE object.");
                shape = shapes[1];
                if (!MathTypeOleInterop.IsMathTypeOle(shape))
                    throw new InvalidDataException(
                        $"MathType redraw display {identity.FormulaId} no longer owns Equation.DSMT4.");
                NormalizeMathTypeDisplayTabGeometry(document, shape);
                normalized++;
            }
            finally
            {
                Release(shape);
                Release(shapes);
            }
        }
        return normalized;
    }

    private static void NormalizeMathTypeDisplayTabGeometry(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A completed MathType display equation no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            format = paragraph.Format;

            var columnWidth = ResolveMathTypeDisplayColumnWidth(document, paragraphRange);
            var firstLineLeft = Math.Max(
                0f,
                format.LeftIndent + format.FirstLineIndent);
            var rightEdge = Math.Max(
                firstLineLeft + 1f,
                columnWidth - Math.Max(0f, format.RightIndent));
            rightEdge = Math.Min(columnWidth, rightEdge);
            var center = firstLineLeft + ((rightEdge - firstLineLeft) / 2f);

            tabs = format.TabStops;
            tabs.ClearAll();
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(
                center,
                WdTabAlignment.wdAlignTabCenter,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = null;
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(
                rightEdge,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = null;

            // Re-read the effective collection after both mutations. This catches
            // the Word case that originally caused the bug: Add() can return
            // successfully while a later style import makes the inherited full-page
            // tab pair effective again.
            Release(tabs);
            tabs = format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1; index <= tabs.Count; index++)
            {
                Release(tab);
                tab = tabs[index];
                if (tab.Alignment == WdTabAlignment.wdAlignTabCenter
                    && Math.Abs(tab.Position - center) <= 0.5f)
                    hasCenter = true;
                if (tab.Alignment == WdTabAlignment.wdAlignTabRight
                    && Math.Abs(tab.Position - rightEdge) <= 0.5f)
                    hasRight = true;
            }
            if (!hasCenter || !hasRight)
                throw new InvalidDataException(
                    "Word did not retain the completed MathType display equation's column-local tab geometry.");
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static void NormalizeMathTypeDisplayParagraphFormat(
        Document document,
        Paragraph paragraph,
        Range paragraphRange,
        MathTypeDisplayParagraphLayout? preservedLayout = null,
        float? knownColumnWidth = null,
        WordParagraphSnapshot? sourceParagraph = null,
        bool keepUserParagraph = false)
    {
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            format = paragraph.Format;
            if (preservedLayout is null && keepUserParagraph)
            {
                // The user's style owns indents and spacing; MathType's tab
                // scaffold needs only a non-centered paragraph.
                sourceParagraph!.Restore(paragraphRange);
                if (format.Alignment != WdParagraphAlignment.wdAlignParagraphJustify)
                    format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            }
            else if (preservedLayout is null)
            {
                format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                format.LeftIndent = 0;
                format.RightIndent = 0;
                format.FirstLineIndent = 0;
                format.SpaceBefore = 0;
                format.SpaceAfter = 0;
                format.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                // MTDisplayEquation owns spacing and the first line; the
                // user's own left/right indent of the paragraph stays.
                sourceParagraph?.Restore(paragraphRange, keepStyle: false);
            }
            else
            {
                // Redraw replaces source LaTeX in an already-owned paragraph. Keep
                // that paragraph's page-flow/spacing contract instead of silently
                // replacing it with MTDisplayEquation's document-wide defaults.
                // The native MathType center/right tab scaffold is rebuilt below,
                // so source custom tabs are intentionally not restored here.
                // MathType display equations are positioned by the native center/right
                // tab scaffold below. Preserving a centered LaTeX source paragraph here
                // applies paragraph centering on top of that tab geometry and visibly
                // shifts the OLE in both single- and multi-column layouts. Keep the
                // source paragraph's flow/spacing metrics, but restore MathType's
                // tab-driven paragraph alignment contract.
                format.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                format.LeftIndent = preservedLayout.LeftIndent;
                format.RightIndent = preservedLayout.RightIndent;
                format.FirstLineIndent = preservedLayout.FirstLineIndent;
                format.SpaceBefore = preservedLayout.SpaceBefore;
                format.SpaceAfter = preservedLayout.SpaceAfter;
                format.LineSpacingRule = preservedLayout.LineSpacingRule;
                try { format.LineSpacing = preservedLayout.LineSpacing; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                format.KeepTogether = preservedLayout.KeepTogether;
                format.KeepWithNext = preservedLayout.KeepWithNext;
                format.WidowControl = preservedLayout.WidowControl;
                format.PageBreakBefore = preservedLayout.PageBreakBefore;
            }

            tabs = format.TabStops;
            tabs.ClearAll();
            // Word invalidates the COM TabStops collection after ClearAll when the
            // cleared stops were inherited from a paragraph style. Reusing that
            // stale collection makes Add() return normally but silently leaves the
            // inherited MTDisplayEquation tabs in force. Reacquire after every
            // structural mutation so the direct paragraph overrides are real.
            Release(tabs);
            tabs = null;

            // MTDisplayEquation is a document-global style. MathType normally seeds
            // its tabs from the full page text width, which is wrong as soon as a
            // section uses newspaper columns: a two-column page can inherit a center
            // tab at half of the *whole page* width, already beyond the first column.
            // Direct-format every live display row from the width of the column that
            // actually owns this paragraph. This also overrides an old/full-page
            // MTDisplayEquation style without mutating the global style itself.
            // Large redraws resolve the section/column geometry once while the
            // source document is still intact and pass it here. Avoiding repeated
            // Sections/PageSetup/TextColumns COM calls is essential: on a 1000-formula
            // corpus those queries alone can add minutes of Word repagination work.
            float columnWidth;
            using (WordOperationMetrics.Measure("MathType.Display.ResolveColumnWidth"))
                columnWidth = knownColumnWidth
                    ?? ResolveMathTypeDisplayColumnWidth(document, paragraphRange);
            var firstLineLeft = Math.Min(
                Math.Max(0f, columnWidth - 1f),
                Math.Max(
                    0f,
                    format.LeftIndent + format.FirstLineIndent));
            var rightEdge = Math.Max(
                firstLineLeft + 1f,
                columnWidth - Math.Max(0f, format.RightIndent));
            rightEdge = Math.Min(columnWidth, rightEdge);
            var center = firstLineLeft + ((rightEdge - firstLineLeft) / 2f);

            tabs = format.TabStops;
            tab = tabs.Add(
                center,
                WdTabAlignment.wdAlignTabCenter,
                WdTabLeader.wdTabLeaderSpaces);
            Release(tab);
            tab = null;
            Release(tabs);
            tabs = format.TabStops;
            tab = tabs.Add(
                rightEdge,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
        }
    }

    private static float ResolveMathTypeDisplayColumnWidth(
        Document document,
        Range paragraphRange)
    {
        // A range with no intersecting tables cannot own a cell. This structural
        // negative test avoids Word's layout-backed wdWithInTable information
        // query for ordinary body paragraphs, particularly expensive immediately
        // after an OLE/style mutation in revision-rich documents. A positive or
        // unavailable table result still uses the exact existing cell proof.
        var mayOwnCell = true;
        Tables? localTables = null;
        using (WordOperationMetrics.Measure("MathType.Geometry.LocalTables"))
        {
            try { localTables = paragraphRange.Tables; mayOwnCell = localTables.Count != 0; }
            catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
            finally { Release(localTables); }
        }
        if (mayOwnCell)
        {
            using (WordOperationMetrics.Measure("MathType.Geometry.CellWidth"))
                if (TryResolveOwningTableCellContentWidth(paragraphRange, out var cellWidth))
                    return cellWidth;
        }

        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        TextColumns? textColumns = null;
        TextColumn? textColumn = null;
        try
        {
            try
            {
                sections = paragraphRange.Sections;
                if (sections.Count > 0)
                {
                    section = sections[1];
                    pageSetup = section.PageSetup;
                }
            }
            catch
            {
                Release(pageSetup);
                pageSetup = null;
                Release(section);
                section = null;
                Release(sections);
                sections = null;
            }

            pageSetup ??= document.PageSetup;
            var pageTextWidth = Math.Max(
                72f,
                pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
            textColumns = pageSetup.TextColumns;
            if (textColumns is null || textColumns.Count <= 1)
                return pageTextWidth;

            var widths = new List<float>(textColumns.Count);
            var spaces = new List<float>(textColumns.Count);
            for (var index = 1; index <= textColumns.Count; index++)
            {
                Release(textColumn);
                textColumn = textColumns[index];
                var width = textColumn.Width;
                if (!(width > 1f) || float.IsNaN(width) || float.IsInfinity(width))
                    continue;

                // Word exposes SpaceAfter only for columns that are followed by
                // another column. On the final TextColumn some Office builds return
                // a missing COM value (and others throw). Letting that escape made
                // the whole resolver fall back to full-page width, exactly the
                // geometry that breaks MathType display equations in newspaper
                // columns. A final column has no following gap by definition.
                var spaceAfter = 0f;
                if (index < textColumns.Count)
                {
                    try
                    {
                        var candidateSpace = textColumn.SpaceAfter;
                        if (candidateSpace > 0f
                            && !float.IsNaN(candidateSpace)
                            && !float.IsInfinity(candidateSpace))
                            spaceAfter = candidateSpace;
                    }
                    catch (Exception ignored)
                    {
                        // The width is still authoritative. Missing spacing must
                        // never downgrade a valid multi-column section to page width.
                        IgnoredErrors.Ignore(ignored);
                    }
                }
                widths.Add(width);
                spaces.Add(spaceAfter);
            }
            if (widths.Count == 0) return pageTextWidth;
            if (widths.Count == 1 || widths.All(width => width == widths[0]))
                return widths[0];

            // Unequal-width newspaper columns are uncommon but legal. Word reports
            // a live paragraph's horizontal page coordinate, so use it to select
            // the owning column when pagination is available. If background layout
            // cannot provide a coordinate, the narrowest column is the conservative
            // fallback: it cannot place the equation outside any valid column.
            try
            {
                var pageX = Convert.ToSingle(paragraphRange.get_Information(
                    WdInformation.wdHorizontalPositionRelativeToPage));
                if (!float.IsNaN(pageX)
                    && !float.IsInfinity(pageX)
                    && pageX >= 0f)
                {
                    var relativeX = pageX - pageSetup.LeftMargin;
                    var cursor = 0f;
                    for (var index = 0; index < widths.Count; index++)
                    {
                        var next = cursor + widths[index] + spaces[index];
                        if (relativeX < next || index == widths.Count - 1)
                            return widths[index];
                        cursor = next;
                    }
                }
            }
            catch (Exception ignored)
            {
                // Background/protected stories can reject page-position queries.
                IgnoredErrors.Ignore(ignored);
            }

            return widths.Min();
        }
        catch
        {
            // Preserve the mature single-column behavior when section/column COM
            // metadata is unavailable in a custom story.
            try
            {
                pageSetup ??= document.PageSetup;
                return Math.Max(
                    72f,
                    pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin);
            }
            catch
            {
                return 468f;
            }
        }
        finally
        {
            Release(textColumn);
            Release(textColumns);
            Release(pageSetup);
            Release(section);
            Release(sections);
        }
    }

    internal static string ReadMathTypeNumberPositionPreference(Document document)
    {
        object? propertiesObject = null;
        object? propertyObject = null;
        try
        {
            propertiesObject = document.CustomDocumentProperties;
            if (propertiesObject is null) return "right";
            dynamic properties = propertiesObject;
            try
            {
                propertyObject = properties["MTEqnNumsOnRight"];
                dynamic property = propertyObject;
                var value = property.Value;
                if (value is bool right) return right ? "right" : "left";
                if (value is int integer) return integer != 0 ? "right" : "left";
                var text = Convert.ToString((object)value);
                if (bool.TryParse(text, out bool parsed))
                    return parsed ? "right" : "left";
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            return "right";
        }
        finally
        {
            Release(propertyObject);
            Release(propertiesObject);
        }
    }

    internal static void UpdateNestedMathTypeNumberFields(Field outer)
    {
        Range? code = null;
        Fields? fields = null;
        Field? field = null;
        try
        {
            code = outer.Code;
            fields = code.Fields;
            // WPS: never update the MACROBUTTON or a hidden SEQ (see
            // WordFieldCode.IsHiddenSequence); the visible SEQ counts without them.
            var wps = OfficeHostEnvironment.IsWps;
            for (var index = 1 + WordFieldCode.OwnerEntries(fields, code); index <= fields.Count; index++)
            {
                Release(field);
                field = fields[index];
                if (wps)
                {
                    Range? fieldCode = null;
                    try
                    {
                        fieldCode = field.Code;
                        if (WordFieldCode.IsHiddenSequence(fieldCode.Text)) continue;
                    }
                    finally { Release(fieldCode); }
                }
                try { field.Update(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            if (!wps)
                try { outer.Update(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(code);
        }
    }

    internal static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        Range range)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        try
        {
            paragraphs = range.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            return CaptureMathTypeDisplayParagraphLayout(paragraph);
        }
        catch { return null; }
        finally
        {
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return null;
            paragraph = paragraphs[1];
            return CaptureMathTypeDisplayParagraphLayout(paragraph);
        }
        catch { return null; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static MathTypeDisplayParagraphLayout? CaptureMathTypeDisplayParagraphLayout(
        Paragraph paragraph)
    {
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            format = paragraph.Format;
            var layout = new MathTypeDisplayParagraphLayout
            {
                Alignment = format.Alignment,
                LeftIndent = format.LeftIndent,
                RightIndent = format.RightIndent,
                FirstLineIndent = format.FirstLineIndent,
                SpaceBefore = format.SpaceBefore,
                SpaceAfter = format.SpaceAfter,
                LineSpacingRule = format.LineSpacingRule,
                LineSpacing = format.LineSpacing,
                BaseLineAlignment = format.BaseLineAlignment,
                KeepTogether = format.KeepTogether,
                KeepWithNext = format.KeepWithNext,
                WidowControl = format.WidowControl,
                PageBreakBefore = format.PageBreakBefore,
            };
            tabStops = format.TabStops;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop);
                tabStop = tabStops[index];
                if (tabStop.Alignment == WdTabAlignment.wdAlignTabLeft) continue;
                layout.SpecialTabStops.Add((tabStop.Position, tabStop.Alignment, tabStop.Leader));
            }
            return layout;
        }
        catch { return null; }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
        }
    }

    internal static int ReadDocumentParagraphCount(Document document)
    {
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            return paragraphs.Count;
        }
        finally { Release(paragraphs); }
    }

    internal static string CaptureLeadingHorizontalWhitespace(
        Document document,
        int position)
    {
        Range? probe = null;
        Range? content = null;
        try
        {
            content = document.Content;
            var start = Math.Max(
                content.Start,
                Math.Min(position, content.End));
            var end = Math.Min(
                content.End,
                start + 64);
            if (end <= start)
                return string.Empty;

            probe = document.Range(
                start,
                end);
            var text = probe.Text
                ?? string.Empty;
            var length = 0;
            while (length < text.Length)
            {
                var character = text[length];
                if (character is not (' ' or '\t' or '\u00A0'))
                    break;
                length++;
            }
            return length == 0
                ? string.Empty
                : text.Substring(0, length);
        }
        finally
        {
            Release(content);
            Release(probe);
        }
    }

    internal static bool RepairMathTypeInsertXmlParagraphSplit(
        Document document,
        InlineShape shape,
        int sourceParagraphCount,
        string? preservedRightBoundaryWhitespace = null,
        MathTypeLocalEditScope? localScope = null)
    {
        Range? shapeRange = null;
        try
        {
            shapeRange = shape.Range;
            return RepairMathTypeInsertXmlParagraphSplit(
                document,
                shapeRange,
                sourceParagraphCount,
                preservedRightBoundaryWhitespace,
                localScope);
        }
        finally { Release(shapeRange); }
    }

    internal static bool RepairMathTypeInsertXmlParagraphSplit(
        Document document,
        Range hostRange,
        int sourceParagraphCount,
        string? preservedRightBoundaryWhitespace = null,
        MathTypeLocalEditScope? localScope = null)
    {
        var currentParagraphCount = localScope?.ReadParagraphCount() ?? ReadDocumentParagraphCount(document);
        if (currentParagraphCount == sourceParagraphCount) return false;
        if (currentParagraphCount != sourceParagraphCount + 1)
            throw new InvalidOperationException(
                $"Word changed the paragraph count unexpectedly while materializing a MathType OLE: before={sourceParagraphCount}, after={currentParagraphCount}.");

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? paragraphMark = null;
        Range? followingProbe = null;
        Range? mergeRange = null;
        try
        {
            shapeRange = hostRange.Duplicate;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidOperationException(
                    "The materialized MathType OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.End <= paragraphRange.Start)
                throw new InvalidOperationException(
                    "The materialized MathType OLE paragraph has no paragraph mark to repair.");
            paragraphMark = document.Range(paragraphRange.End - 1, paragraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "Word inserted a MathType OLE paragraph boundary that VisualTeX could not identify safely.");

            // InsertXML materializes an inline MathType OLE as a complete Word
            // paragraph and can split the user's original paragraph at the
            // insertion point. Deleting only that paragraph mark lets Word
            // normalize away leading spaces from the second paragraph. Merge the
            // exact structural boundary and its existing leading whitespace in
            // one replacement instead: the paragraph mark disappears while the
            // user's whitespace is recreated byte-for-byte as ordinary text.
            var probeStart = paragraphRange.End;
            var probeEnd = Math.Min(
                document.Content.End,
                probeStart + 64);
            if (probeEnd > probeStart)
            {
                followingProbe = document.Range(
                    probeStart,
                    probeEnd);
            }

            var followingText =
                followingProbe?.Text
                ?? string.Empty;
            var actualLeadingLength = 0;
            while (actualLeadingLength < followingText.Length)
            {
                var character =
                    followingText[actualLeadingLength];
                if (character is not (' ' or '\t' or '\u00A0'))
                    break;
                actualLeadingLength++;
            }

            var desiredLeadingWhitespace =
                preservedRightBoundaryWhitespace
                ?? (actualLeadingLength == 0
                    ? string.Empty
                    : followingText.Substring(
                        0,
                        actualLeadingLength));

            if (actualLeadingLength == 0
                && desiredLeadingWhitespace.Length == 0)
            {
                paragraphMark.Delete();
            }
            else
            {
                mergeRange = document.Range(
                    paragraphMark.Start,
                    paragraphRange.End
                    + actualLeadingLength);
                mergeRange.Text =
                    desiredLeadingWhitespace;
            }
        }
        finally
        {
            Release(mergeRange);
            Release(followingProbe);
            Release(paragraphMark);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }

        var repairedParagraphCount = localScope?.ReadParagraphCount() ?? ReadDocumentParagraphCount(document);
        if (repairedParagraphCount != sourceParagraphCount)
            throw new InvalidOperationException(
                $"VisualTeX could not restore the original MathType OLE paragraph structure: expected={sourceParagraphCount}, actual={repairedParagraphCount}.");
        return true;
    }

    internal static MathTypeDisplayParagraphLayout? RepairDetachedMathTypeNumberParagraph(
        Document document,
        InlineShape shape)
    {
        Range? shapeRange = null;
        Paragraphs? shapeParagraphs = null;
        Paragraph? shapeParagraph = null;
        Range? shapeParagraphRange = null;
        Range? nextProbe = null;
        Paragraphs? nextParagraphs = null;
        Paragraph? nextParagraph = null;
        Range? nextRange = null;
        Range? paragraphMark = null;
        try
        {
            shapeRange = shape.Range;
            shapeParagraphs = shapeRange.Paragraphs;
            if (shapeParagraphs.Count != 1) return null;
            shapeParagraph = shapeParagraphs[1];
            shapeParagraphRange = shapeParagraph.Range;
            if (shapeParagraphRange.End >= document.Content.End) return null;

            // Probe the first character after the formula paragraph mark.  A
            // non-empty one-character range reliably belongs to the following
            // paragraph, unlike a collapsed boundary range which Word can resolve
            // to either side depending on field state.
            nextProbe = document.Range(
                shapeParagraphRange.End,
                Math.Min(document.Content.End, shapeParagraphRange.End + 1));
            nextParagraphs = nextProbe.Paragraphs;
            if (nextParagraphs.Count < 1) return null;
            nextParagraph = nextParagraphs[1];
            nextRange = nextParagraph.Range;
            if (!IsDetachedMathTypeNumberParagraph(nextRange)) return null;

            var numberingLayout = CaptureMathTypeDisplayParagraphLayout(nextParagraph);
            paragraphMark = document.Range(
                shapeParagraphRange.End - 1,
                shapeParagraphRange.End);
            if (!string.Equals(paragraphMark.Text, "\r", StringComparison.Ordinal))
                return null;
            paragraphMark.Delete();
            return numberingLayout;
        }
        catch
        {
            // This is compatibility recovery for documents already damaged by an
            // older VisualTeX build.  Never make an otherwise valid MathType edit
            // fail merely because an adjacent paragraph only resembles MTPlaceRef.
            return null;
        }
        finally
        {
            Release(paragraphMark);
            Release(nextRange);
            Release(nextParagraph);
            Release(nextParagraphs);
            Release(nextProbe);
            Release(shapeParagraphRange);
            Release(shapeParagraph);
            Release(shapeParagraphs);
            Release(shapeRange);
        }
    }

    private static bool IsDetachedMathTypeNumberParagraph(Range range)
    {
        InlineShapes? inlineShapes = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            inlineShapes = range.InlineShapes;
            if (inlineShapes.Count != 0) return false;

            fields = WordFormulaHost.GetLocalFields(range);
            var hasPlaceRef = false;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    hasPlaceRef = true;
            }
            if (!hasPlaceRef) return false;

            var text = (range.Text ?? string.Empty)
                .Replace("\r", string.Empty)
                .Replace("\a", string.Empty)
                .Replace("\t", string.Empty)
                .Replace("\u0013", string.Empty)
                .Replace("\u0014", string.Empty)
                .Replace("\u0015", string.Empty)
                .Trim();
            if (text.Length == 0) return true;

            var sawNumber = false;
            foreach (var character in text)
            {
                if (char.IsWhiteSpace(character)) continue;
                if (char.IsDigit(character))
                {
                    sawNumber = true;
                    continue;
                }
                if ("()[]{}.,:;-–—/\\".IndexOf(character) >= 0) continue;
                return false;
            }
            return sawNumber;
        }
        catch { return false; }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(inlineShapes);
        }
    }

    internal static void RestoreMathTypeDisplayParagraphLayout(
        InlineShape shape,
        MathTypeDisplayParagraphLayout layout)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        ParagraphFormat? format = null;
        TabStops? tabStops = null;
        TabStop? tabStop = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return;
            paragraph = paragraphs[1];
            format = paragraph.Format;
            format.Alignment = layout.Alignment;
            format.LeftIndent = layout.LeftIndent;
            format.RightIndent = layout.RightIndent;
            format.FirstLineIndent = layout.FirstLineIndent;
            format.SpaceBefore = layout.SpaceBefore;
            format.SpaceAfter = layout.SpaceAfter;
            format.LineSpacingRule = layout.LineSpacingRule;
            try { format.LineSpacing = layout.LineSpacing; } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            format.BaseLineAlignment = layout.BaseLineAlignment;
            format.KeepTogether = layout.KeepTogether;
            format.KeepWithNext = layout.KeepWithNext;
            format.WidowControl = layout.WidowControl;
            format.PageBreakBefore = layout.PageBreakBefore;

            if (layout.SpecialTabStops.Count == 0) return;
            tabStops = format.TabStops;
            foreach (var special in layout.SpecialTabStops)
            {
                var exists = false;
                for (var index = 1; index <= tabStops.Count; index++)
                {
                    Release(tabStop);
                    tabStop = tabStops[index];
                    if (Math.Abs(tabStop.Position - special.Position) <= 0.5f
                        && tabStop.Alignment == special.Alignment
                        && tabStop.Leader == special.Leader)
                    {
                        exists = true;
                        break;
                    }
                }
                if (exists) continue;
                Release(tabStop);
                tabStop = tabStops.Add(special.Position, special.Alignment, special.Leader);
            }
        }
        finally
        {
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }
}

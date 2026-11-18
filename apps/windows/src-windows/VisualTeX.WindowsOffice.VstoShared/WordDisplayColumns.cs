using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// The width a display formula's line is laid out in, for its centre and right tab stops (DisplayColumnGeometry):
/// the owning table cell; else the section's single column (the page's text width); else its columns' common width;
/// else, for columns of different widths, the column the paragraph's position falls in (the narrowest when Word
/// cannot tell). Every VisualTeX writer (OLE, picture, MathType, native tables) uses this one answer.
/// </summary>
internal static class WordDisplayColumns
{
    /// <summary>A section's text geometry: left margin, the page's text width and its columns.</summary>
    internal sealed class SectionLayout
    {
        internal SectionLayout(float leftMargin, float pageTextWidth, float[] widths, float[] spaces)
        {
            LeftMargin = leftMargin;
            PageTextWidth = pageTextWidth;
            Widths = widths;
            Spaces = spaces;
        }

        internal float LeftMargin { get; }
        internal float PageTextWidth { get; }
        internal float[] Widths { get; }
        internal float[] Spaces { get; }

        /// <summary>The width every line of the section is laid out in, or null when its columns differ.</summary>
        internal float? UniformWidth =>
            Widths.Length <= 1 ? PageTextWidth
            : DisplayColumnGeometry.AreEqual(Widths) ? Widths[0]
            : null;

        internal static SectionLayout Letter { get; } = new(72f, 468f, Array.Empty<float>(), Array.Empty<float>());
    }

    internal static SectionLayout Read(PageSetup pageSetup)
    {
        var leftMargin = pageSetup.LeftMargin;
        var pageTextWidth = Math.Max(72f, pageSetup.PageWidth - leftMargin - pageSetup.RightMargin);
        TextColumns? textColumns = null;
        TextColumn? textColumn = null;
        try
        {
            textColumns = pageSetup.TextColumns;
            if (textColumns is null || textColumns.Count <= 1)
                return new SectionLayout(leftMargin, pageTextWidth, Array.Empty<float>(), Array.Empty<float>());
            var widths = new List<float>(textColumns.Count);
            var spaces = new List<float>(textColumns.Count);
            for (var index = 1; index <= textColumns.Count; index++)
            {
                Release(textColumn);
                textColumn = textColumns[index];
                var width = textColumn.Width;
                if (!(width > 1f) || float.IsNaN(width) || float.IsInfinity(width))
                    continue;
                // Word exposes SpaceAfter only for a column another one follows; on the last column some Office
                // builds return a missing value and others throw. A missing gap never turns a multi-column section
                // into a page-wide one.
                var spaceAfter = 0f;
                if (index < textColumns.Count)
                {
                    try
                    {
                        var candidate = textColumn.SpaceAfter;
                        if (candidate > 0f && !float.IsNaN(candidate) && !float.IsInfinity(candidate))
                            spaceAfter = candidate;
                    }
                    catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                }
                widths.Add(width);
                spaces.Add(spaceAfter);
            }
            return widths.Count <= 1
                ? new SectionLayout(leftMargin, widths.Count == 1 ? widths[0] : pageTextWidth, Array.Empty<float>(), Array.Empty<float>())
                : new SectionLayout(leftMargin, pageTextWidth, widths.ToArray(), spaces.ToArray());
        }
        catch (Exception ignored)
        {
            // Custom stories can hide their columns: the page's text width, as before columns were read.
            IgnoredErrors.Ignore(ignored);
            return new SectionLayout(leftMargin, pageTextWidth, Array.Empty<float>(), Array.Empty<float>());
        }
        finally
        {
            Release(textColumn);
            Release(textColumns);
        }
    }

    /// <summary>The geometry of the section a range is in (the document's first section's when Word hides it).</summary>
    internal static SectionLayout ReadFor(Document document, Range range)
    {
        Sections? sections = null;
        Section? section = null;
        PageSetup? pageSetup = null;
        try
        {
            try
            {
                sections = range.Sections;
                if (sections.Count > 0)
                {
                    section = sections[1];
                    pageSetup = section.PageSetup;
                }
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            pageSetup ??= document.PageSetup;
            return Read(pageSetup);
        }
        catch (Exception ignored)
        {
            IgnoredErrors.Ignore(ignored);
            return SectionLayout.Letter;
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
        }
    }

    /// <summary>
    /// The width the line at <paramref name="range"/> is laid out in. <paramref name="ownCell"/> false skips the
    /// owning table cell (sizing a table itself, which sits in the column).
    /// </summary>
    internal static float ColumnWidth(Document document, Range range, bool ownCell = true, SectionLayout? section = null)
    {
        if (ownCell && TryOwningCellWidth(range, out var cellWidth))
            return cellWidth;
        var layout = section ?? ReadFor(document, range);
        return layout.UniformWidth ?? DisplayColumnGeometry.ColumnAt(layout.Widths, layout.Spaces, OffsetFromLeftMargin(range, layout));
    }

    /// <summary>The centre and right stops for a display paragraph at <paramref name="range"/> with these indents.</summary>
    internal static (float Center, float Right) TabStops(
        Document document,
        Range range,
        float leftIndent = 0f,
        float firstLineIndent = 0f,
        float rightIndent = 0f,
        SectionLayout? section = null) =>
        DisplayColumnGeometry.TabStops(ColumnWidth(document, range, ownCell: true, section), leftIndent, firstLineIndent, rightIndent);

    /// <summary>
    /// Section geometry read once per batch: single-column and equal-column sections give every formula the same
    /// width; only a section with columns of different widths asks Word where each paragraph is. The section is found
    /// live (wdActiveEndSectionNumber, about 30 ms), except in the usual one-section document, so positions that move
    /// while the batch writes do not matter.
    /// </summary>
    internal sealed class BatchCache
    {
        private readonly Dictionary<int, SectionLayout> _bySection = new();
        private int _sectionCount = -1;
        private List<(int Start, int End, SectionLayout Layout)>? _byPosition;

        /// <summary>
        /// A cache that finds sections by position: every section's extent is read now, for a batch whose callers
        /// pass the positions the formulas had when it started (no live section lookup per formula).
        /// </summary>
        internal static BatchCache BySectionPosition(Document document)
        {
            var cache = new BatchCache { _byPosition = new List<(int Start, int End, SectionLayout Layout)>() };
            Sections? sections = null;
            try
            {
                sections = document.Sections;
                for (var index = 1; index <= sections.Count; index++)
                {
                    Section? section = null;
                    Range? sectionRange = null;
                    PageSetup? setup = null;
                    try
                    {
                        section = sections[index];
                        sectionRange = section.Range;
                        setup = section.PageSetup;
                        cache._byPosition.Add((sectionRange.Start, sectionRange.End, Read(setup)));
                    }
                    finally { Release(setup); Release(sectionRange); Release(section); }
                }
            }
            catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            finally { Release(sections); }
            return cache;
        }

        /// <summary>
        /// The width the line at <paramref name="range"/> is laid out in, or null when Word cannot tell.
        /// <paramref name="sectionPosition"/>: for a cache made by <see cref="BySectionPosition"/>, the position the
        /// formula had when the batch started (default: the range's start).
        /// </summary>
        internal float? ColumnWidthFor(Document document, Range range, int? sectionPosition = null)
        {
            try
            {
                if (TryOwningCellWidth(range, out var cellWidth))
                    return cellWidth;
                var layout = _byPosition is null ? LayoutFor(document, range) : LayoutAt(sectionPosition ?? range.Start);
                if (layout is null) return null;
                return layout.UniformWidth
                    ?? DisplayColumnGeometry.ColumnAt(layout.Widths, layout.Spaces, OffsetFromLeftMargin(range, layout));
            }
            catch (Exception ignored)
            {
                IgnoredErrors.Ignore(ignored);
                return null;
            }
        }

        /// <summary>The stops for a display formula written at <paramref name="range"/> (indents zero), or null.</summary>
        internal (float Center, float Right)? TabStopsFor(Document document, Range range, int? sectionPosition = null) =>
            ColumnWidthFor(document, range, sectionPosition) is { } width ? DisplayColumnGeometry.TabStops(width) : null;

        private SectionLayout? LayoutAt(int position)
        {
            var sections = _byPosition!;
            foreach (var (start, end, layout) in sections)
                if (start <= position && position < end)
                    return layout;
            return sections.Count > 0 && position >= sections[sections.Count - 1].End ? sections[sections.Count - 1].Layout : null;
        }

        private SectionLayout? LayoutFor(Document document, Range range)
        {
            if (_sectionCount < 0)
            {
                Sections? all = null;
                try { all = document.Sections; _sectionCount = all.Count; }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); _sectionCount = 0; }
                finally { Release(all); }
            }
            int number;
            if (_sectionCount == 1)
                number = 1;
            else
            {
                try { number = Convert.ToInt32(range.get_Information(WdInformation.wdActiveEndSectionNumber)); }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); return null; }
                if (number <= 0) return null;
            }
            if (_bySection.TryGetValue(number, out var cached))
                return cached;
            Sections? sections = null;
            Section? section = null;
            PageSetup? setup = null;
            try
            {
                sections = document.Sections;
                if (number > sections.Count) return null;
                section = sections[number];
                setup = section.PageSetup;
                var layout = Read(setup);
                _bySection[number] = layout;
                return layout;
            }
            catch (Exception ignored)
            {
                IgnoredErrors.Ignore(ignored);
                return null;
            }
            finally
            {
                Release(setup);
                Release(section);
                Release(sections);
            }
        }
    }

    private static float? OffsetFromLeftMargin(Range range, SectionLayout layout)
    {
        try
        {
            var pageX = Convert.ToSingle(range.get_Information(WdInformation.wdHorizontalPositionRelativeToPage));
            if (float.IsNaN(pageX) || float.IsInfinity(pageX) || pageX < 0f) return null;
            return pageX - layout.LeftMargin;
        }
        catch (Exception ignored)
        {
            // Background and protected stories can refuse page positions: the narrowest column then.
            IgnoredErrors.Ignore(ignored);
            return null;
        }
    }

    // A range with no table in it owns no cell: that cheap structural answer avoids Word's layout-backed
    // wdWithInTable query for ordinary body paragraphs.
    private static bool TryOwningCellWidth(Range range, out float width)
    {
        width = 0f;
        Tables? tables = null;
        try
        {
            tables = range.Tables;
            if (tables.Count == 0) return false;
        }
        catch (COMException ignored) { IgnoredErrors.Ignore(ignored); }
        finally { Release(tables); }
        return WordFormulaHost.TryGetCellContentWidth(range, out width);
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}

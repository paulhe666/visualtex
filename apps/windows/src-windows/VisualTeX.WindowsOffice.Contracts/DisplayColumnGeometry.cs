using System;
using System.Collections.Generic;
using System.Linq;

namespace VisualTeX.WindowsOffice.Contracts;

/// <summary>
/// Where a display formula's tab stops go: Word measures tab positions from the left edge of the column the line is
/// in, so the centre and right stops come from that column's width, not from the page's text width (a two-column
/// section otherwise puts the formula on the column edge and the number in the other column or off the page). The
/// same rule as VisualTeX for Mac (VTEquationLayoutWidth) and WPS (textColumnWidth).
/// </summary>
public static class DisplayColumnGeometry
{
    /// <summary>Whether the columns all have one width (one column counts as equal).</summary>
    public static bool AreEqual(IReadOnlyList<float> widths) =>
        widths.Count <= 1 || widths.All(width => Math.Abs(width - widths[0]) <= 0.05f);

    /// <summary>
    /// The width of the column a line starts in. <paramref name="offsetFromLeftMargin"/> is the line's horizontal
    /// distance from the left margin; the columns are walked left to right, each taking its width and the space
    /// after it. Without a usable position (Word could not lay the paragraph out) the narrowest column is the safe
    /// answer: it cannot put the formula outside any column.
    /// </summary>
    public static float ColumnAt(
        IReadOnlyList<float> widths,
        IReadOnlyList<float> spacesAfter,
        float? offsetFromLeftMargin)
    {
        if (widths.Count == 0) throw new ArgumentException("A section has at least one column.", nameof(widths));
        if (AreEqual(widths)) return widths[0];
        if (offsetFromLeftMargin is { } offset && !float.IsNaN(offset) && !float.IsInfinity(offset) && offset >= -1f)
        {
            var cursor = 0f;
            for (var index = 0; index < widths.Count; index++)
            {
                var space = index < spacesAfter.Count ? Math.Max(0f, spacesAfter[index]) : 0f;
                var next = cursor + widths[index] + space;
                if (offset < next || index == widths.Count - 1)
                    return widths[index];
                cursor = next;
            }
        }
        return widths.Min();
    }

    /// <summary>
    /// The centre and right stops for a display paragraph in a column (or table cell) this wide: the right stop at
    /// the column's right edge less the right indent, the centre stop halfway between the first line's left edge
    /// and the right stop.
    /// </summary>
    public static (float Center, float Right) TabStops(
        float columnWidth,
        float leftIndent = 0f,
        float firstLineIndent = 0f,
        float rightIndent = 0f)
    {
        var width = columnWidth > 1f && !float.IsNaN(columnWidth) && !float.IsInfinity(columnWidth) ? columnWidth : 468f;
        var firstLineLeft = Math.Min(Math.Max(0f, width - 1f), Math.Max(0f, Finite(leftIndent) + Finite(firstLineIndent)));
        var rightEdge = Math.Max(firstLineLeft + 1f, width - Math.Max(0f, Finite(rightIndent)));
        rightEdge = Math.Min(width, rightEdge);
        return (firstLineLeft + (rightEdge - firstLineLeft) / 2f, rightEdge);
    }

    private static float Finite(float value) => float.IsNaN(value) || float.IsInfinity(value) ? 0f : value;
}

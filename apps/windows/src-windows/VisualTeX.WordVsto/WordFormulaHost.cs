using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>Read-only physical ownership shared by redraw, conversion and layout.
/// A document containing tables does not make its body paragraphs table cells.
/// No caller may query Cells until the actual source position is proven in-table.</summary>
internal static class WordFormulaHost
{
    internal static bool IsContainedByCell(int sourceStart, int sourceEnd, int cellStart, int cellEnd)
        => sourceStart >= cellStart && sourceStart < cellEnd
            && sourceEnd >= sourceStart && sourceEnd <= cellEnd;

    // The caller owns the returned COM reference. Always resolves locally, using
    // the leading source character (not selection.End and not a document scan).
    internal static Cell? TryGetOwningCell(Range source)
    {
        Range? probe = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? owner = null;
        try
        {
            var start = source.Start;
            var end = source.End;
            probe = source.Duplicate;
            probe.SetRange(start, end > start ? start + 1 : start);
            if (!(bool)probe.get_Information(WdInformation.wdWithInTable)) return null;
            cells = probe.Cells;
            if (cells.Count != 1) return null;
            cell = cells[1];
            owner = cell.Range;
            if (source.StoryType != owner.StoryType
                || !IsContainedByCell(start, end, owner.Start, owner.End)) return null;
            var result = cell;
            cell = null;
            return result;
        }
        catch (COMException error) when (error.HResult == unchecked((int)0x800A1713))
        {
            // Word may temporarily retain in-table affinity at a collapsed
            // boundary after a structural replacement. This specific absence is
            // not an error; other COM failures are not silently reclassified.
            return null;
        }
        finally { Release(owner); Release(cell); Release(cells); Release(probe); }
    }

    internal static bool IsWhollyInsideMath(Range source)
    {
        // Word collection membership is not geometric ownership: a caption or
        // cell boundary can report neighbouring OMath objects. Only containment
        // in the actual mathematical interval makes an alias mathematical.
        OMaths? maths = null;
        OMath? math = null;
        Range? owner = null;
        try
        {
            maths = source.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Release(owner); owner = null;
                Release(math); math = maths[index];
                owner = math.Range;
                if (source.StoryType == owner.StoryType
                    && ContainsPhysicalRange(owner.Start, owner.End, source.Start, source.End))
                    return true;
            }
            return false;
        }
        finally { Release(owner); Release(math); Release(maths); }
    }

    internal static bool ContainsPhysicalRange(int ownerStart, int ownerEnd, int start, int end)
        => ownerEnd > ownerStart && start >= ownerStart && start < ownerEnd
            && end >= start && end <= ownerEnd;

    internal static bool IsWithinSingleCell(Range source)
    {
        var cell = TryGetOwningCell(source);
        try { return cell is not null; }
        finally { Release(cell); }
    }

    // Range.Fields in Word 2021 can include every preceding body field when
    // a range ends at the first table cell's CR+BEL marker. Query the editable
    // portion, never the structural cell marker. This also bounds the query's
    // cost to the paragraph rather than the number of earlier document fields.
    // The returned collection is owned by the caller, independently of the range.
    internal static Fields GetLocalFields(Range source)
    {
        Range? local = null;
        Range? terminal = null;
        try
        {
            local = source.Duplicate;
            if (local.End > local.Start)
            {
                terminal = local.Duplicate;
                terminal.SetRange(local.End - 1, local.End);
                if (IsCellStructuralMarker(terminal.Text))
                    local.End--;
            }
            return local.Fields;
        }
        finally { Release(terminal); Release(local); }
    }

    internal static int ParagraphTerminatorStoryLength(string? onePositionText)
        => IsCellStructuralMarker(onePositionText) || onePositionText == "\r" ? 1 : 0;

    internal static bool IsCellStructuralMarker(string? text)
        => string.Equals(text, "\r\a", StringComparison.Ordinal)
            || string.Equals(text, "\a", StringComparison.Ordinal);

    internal static bool TryGetCellContentWidth(Range source, out float width)
    {
        width = 0;
        Cell? cell = null;
        try
        {
            cell = TryGetOwningCell(source);
            if (cell is null) return false;
            var total = cell.Width;
            if (!ValidWidth(total)) return false;
            var left = cell.LeftPadding;
            var right = cell.RightPadding;
            if (float.IsNaN(left) || float.IsInfinity(left) || left < 0 || left >= total) left = 0;
            if (float.IsNaN(right) || float.IsInfinity(right) || right < 0 || right >= total) right = 0;
            var usable = total - left - right;
            width = ValidWidth(usable) ? usable : total;
            return true;
        }
        finally { Release(cell); }
    }

    private static bool ValidWidth(float value) => value > 1 && !float.IsNaN(value) && !float.IsInfinity(value)
        && value != (float)WdConstants.wdUndefined;
    private static void Release(object? value) { if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
}

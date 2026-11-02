using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace VisualTeX.PowerPointVsto;

// The text a native equation is written into: a shape's own text, or one table
// cell's (PowerPoint keeps equations in cells as text; pictures and OLE objects
// cannot go into a cell). "<shape id>" or "<shape id>.<row>.<column>" is the
// shape part of the add-in's text-insertion and native-equation references.
internal readonly struct PowerPointTextHost
{
    internal PowerPointTextHost(int shapeId, int row = 0, int column = 0)
    {
        ShapeId = shapeId;
        Row = row;
        Column = column;
    }

    internal int ShapeId { get; }
    internal int Row { get; }
    internal int Column { get; }
    internal bool IsCell => Row > 0;

    public override string ToString() => IsCell
        ? string.Format(CultureInfo.InvariantCulture, "{0}.{1}.{2}", ShapeId, Row, Column)
        : ShapeId.ToString(CultureInfo.InvariantCulture);

    internal static bool TryParse(string? value, out PowerPointTextHost host)
    {
        host = default;
        if (string.IsNullOrEmpty(value)) return false;
        var parts = value!.Split('.');
        if (parts.Length is not (1 or 3)) return false;
        var numbers = new int[parts.Length];
        for (var index = 0; index < parts.Length; index++)
            if (!int.TryParse(parts[index], NumberStyles.None, CultureInfo.InvariantCulture, out numbers[index])
                || numbers[index] <= 0)
                return false;
        host = parts.Length == 1
            ? new PowerPointTextHost(numbers[0])
            : new PowerPointTextHost(numbers[0], numbers[1], numbers[2]);
        return true;
    }

    // The cell of a table shape that holds the text caret or text selection:
    // PowerPoint marks exactly that cell Selected (for a merged cell, each of its
    // grid positions, all with the merged cell's frame; the first is the one
    // whose text it is). False when no cell or several cells are selected.
    internal static bool TryFindSelectedCell(Shape tableShape, out int row, out int column)
    {
        row = column = 0;
        if (tableShape.HasTable != MsoTriState.msoTrue) return false;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        (float Left, float Top, float Width, float Height)? frame = null;
        try
        {
            table = tableShape.Table;
            rows = table.Rows;
            columns = table.Columns;
            int rowCount = rows.Count, columnCount = columns.Count;
            for (var r = 1; r <= rowCount; r++)
            {
                for (var c = 1; c <= columnCount; c++)
                {
                    Cell? cell = null;
                    Shape? shape = null;
                    try
                    {
                        cell = table.Cell(r, c);
                        if (!cell.Selected) continue;
                        shape = cell.Shape;
                        var bounds = (shape.Left, shape.Top, shape.Width, shape.Height);
                        if (frame is null)
                        {
                            frame = bounds;
                            row = r;
                            column = c;
                        }
                        else if (!SameFrame(frame.Value, bounds))
                        {
                            row = column = 0;
                            return false;
                        }
                    }
                    finally { Release(shape); Release(cell); }
                }
            }
            return frame is not null;
        }
        catch (COMException)
        {
            row = column = 0;
            return false;
        }
        finally { Release(columns); Release(rows); Release(table); }
    }

    private static bool SameFrame((float Left, float Top, float Width, float Height) a, (float Left, float Top, float Width, float Height) b) =>
        Math.Abs(a.Left - b.Left) < 0.5f && Math.Abs(a.Top - b.Top) < 0.5f
        && Math.Abs(a.Width - b.Width) < 0.5f && Math.Abs(a.Height - b.Height) < 0.5f;

    // A cell's text shape (Table.Cell(row, column).Shape).
    internal static Shape CellShape(Shape tableShape, int row, int column)
    {
        if (tableShape.HasTable != MsoTriState.msoTrue)
            throw new InvalidOperationException("The PowerPoint table holding the equation no longer exists.");
        Table? table = null;
        Cell? cell = null;
        try
        {
            table = tableShape.Table;
            cell = table.Cell(row, column);
            return cell.Shape;
        }
        catch (COMException error)
        {
            throw new InvalidOperationException("The PowerPoint table cell holding the equation no longer exists.", error);
        }
        finally { Release(cell); Release(table); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}

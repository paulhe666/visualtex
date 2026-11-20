using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using PptTextRange = Microsoft.Office.Interop.PowerPoint.TextRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace VisualTeX.PowerPointVsto;

// PowerPoint turns an equation pasted into a table cell into linear text (its own copy
// of an equation too); only typing one there builds it up. A cell therefore gets its
// equation the way typing does: PowerPoint's own linear format of the equation, made in
// a scratch text box off the slide (where pasting keeps math), written into the cell and
// built up again by PowerPoint's Insert Equation and Professional commands. Linear
// format keeps no formatting: PowerPointCellMath spells letter styles into characters,
// and colors are not kept.
internal static class PowerPointCellEquations
{
    internal static Action PrepareInsertion(Application application, Shape cell, int start, int length,
        string mathMl, string displayMode)
    {
        if (cell.HasTextFrame != MsoTriState.msoTrue)
            throw new InvalidOperationException("The PowerPoint table cell no longer has editable text.");
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("PowerPoint table-cell OMML insertion requires MathML export data.");
        var display = string.Equals(displayMode, "block", StringComparison.Ordinal);
        if (!display && !string.Equals(displayMode, "inline", StringComparison.Ordinal))
            throw new InvalidDataException("PowerPoint table-cell OMML insertion requires inline or block display mode.");
        var prepared = PowerPointCellMath.Prepare(mathMl);
        var text = ReadText(cell);
        if (start < 1 || length < 0 || start - 1 + length > text.Length)
            throw new InvalidOperationException(
                "The PowerPoint text changed while the VisualTeX editor was open; the captured insertion range is no longer valid.");
        return () =>
        {
            var linear = Linearize(application, cell, prepared);
            Write(application, cell, start, length, linear, display);
        };
    }

    private static string Linearize(Application application, Shape cell, string mathMl)
    {
        Slide? slide = null;
        Presentation? presentation = null;
        PageSetup? page = null;
        Shapes? shapes = null;
        Shape? box = null;
        try
        {
            slide = cell.Parent as Slide
                ?? throw new InvalidOperationException("The PowerPoint table cell is no longer on a slide.");
            presentation = (Presentation)slide.Parent;
            page = presentation.PageSetup;
            shapes = slide.Shapes;
            // Right of the slide: never part of what the slide shows while it exists.
            box = shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, page.SlideWidth + 100f, 0f, 800f, 100f);
            PowerPointNativeOmmlKernel.InsertStandalone(application, box, mathMl, false, 20f);
            var zone = ZoneAt(box, 1, ReadText(box).Length)
                ?? throw new InvalidOperationException("PowerPoint did not create the equation to write into the table cell.");
            Select(box, zone.Start, zone.Length);
            Execute(application, "EquationLinearFormat");
            var text = ReadText(box);
            zone = ZoneAt(box, 1, text.Length)
                ?? throw new InvalidOperationException("PowerPoint lost the equation while writing its linear format.");
            // Positions count UTF-16 units, as .NET strings do.
            var linear = zone.Start - 1 + zone.Length <= text.Length ? text.Substring(zone.Start - 1, zone.Length) : string.Empty;
            if (linear.Trim().Length == 0)
                throw new InvalidOperationException("PowerPoint gave the equation no linear format.");
            return linear;
        }
        finally
        {
            if (box is not null)
            {
                try { box.Delete(); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            Release(box);
            Release(shapes);
            Release(page);
            Release(presentation);
            Release(slide);
        }
    }

    // The linear text replaces [start, start + length) (a display equation in its own
    // paragraph), PowerPoint makes an equation of it and builds it up, and the caret
    // goes after it.
    private static void Write(Application application, Shape cell, int start, int length, string linear, bool display)
    {
        var original = ReadText(cell);
        var leading = display && start > 1 && original[start - 2] != '\r';
        var trailing = display && start - 1 + length < original.Length && original[start - 1 + length] != '\r';
        PptTextRange? text = null, target = null;
        try
        {
            text = cell.TextFrame.TextRange;
            target = text.Characters(start, length);
            // Assign, never Delete: Delete also removes a space next to the text.
            target.Text = (leading ? "\r" : string.Empty) + linear + (trailing ? "\r" : string.Empty);
        }
        finally { Release(target); Release(text); }
        var at = start + (leading ? 1 : 0);
        // The linear text as PowerPoint holds it (an assignment can normalize characters).
        var written = ReadText(cell).Length - (original.Length - length) - (leading ? 1 : 0) - (trailing ? 1 : 0);
        if (written <= 0)
            throw new InvalidOperationException("PowerPoint did not write the equation's linear text into the table cell.");

        PowerPointNativeOmmlKernel.ActivateShapeView(application, cell);
        var zone = ZoneAt(cell, at, written);
        if (zone is null)
        {
            Select(cell, at, written);
            Execute(application, "EquationInsertNew");
            zone = ZoneAt(cell, at, written)
                ?? throw new InvalidOperationException("PowerPoint did not make an equation of the table cell's linear text.");
        }
        Select(cell, zone.Value.Start, zone.Value.Length);
        Execute(application, "EquationProfessional");
        zone = ZoneAt(cell, at, 1)
            ?? throw new InvalidOperationException("PowerPoint did not build up the table cell's equation.");

        PptTextRange? equation = null;
        try
        {
            text = cell.TextFrame.TextRange;
            equation = text.Characters(zone.Value.Start, zone.Value.Length);
            if (display)
            {
                var paragraph = equation.ParagraphFormat;
                try
                {
                    paragraph.Alignment = PpParagraphAlignment.ppAlignCenter;
                    paragraph.Bullet.Visible = MsoTriState.msoFalse;
                }
                catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
                finally { Release(paragraph); }
            }
        }
        finally { Release(equation); Release(text); }
        // A caret failure after a successful write is not a failed insertion.
        try { Select(cell, zone.Value.Start + zone.Value.Length, 0); }
        catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
    }

    // The math zone of a shape's (or cell's) text overlapping [start, start + length).
    private static (int Start, int Length)? ZoneAt(Shape shape, int start, int length)
    {
        object? frame = null, text = null, zones = null;
        try
        {
            frame = shape.TextFrame2;
            text = ((dynamic)frame).TextRange;
            zones = ((dynamic)text).MathZones();
            var count = Convert.ToInt32(((dynamic)zones).Count);
            for (var index = 1; index <= count; index++)
            {
                object? zone = null;
                try
                {
                    zone = ((dynamic)text).MathZones(index, 1);
                    int zoneStart = Convert.ToInt32(((dynamic)zone).Start), zoneLength = Convert.ToInt32(((dynamic)zone).Length);
                    if (zoneStart < start + Math.Max(1, length) && zoneStart + zoneLength > start)
                        return (zoneStart, zoneLength);
                }
                finally { Release(zone); }
            }
            return null;
        }
        catch (COMException) { return null; }
        finally { Release(zones); Release(text); Release(frame); }
    }

    // PowerPoint's ribbon commands act on the selection once PowerPoint has taken it in.
    private static void Select(Shape shape, int start, int length)
    {
        PptTextRange? text = null, range = null;
        try
        {
            text = shape.TextFrame.TextRange;
            range = text.Characters(start, length);
            range.Select();
        }
        finally { Release(range); Release(text); }
        System.Windows.Forms.Application.DoEvents();
    }

    private static void Execute(Application application, string idMso)
    {
        CommandBars? bars = null;
        try
        {
            bars = application.CommandBars;
            for (var attempt = 0; !bars.GetEnabledMso(idMso); attempt++)
            {
                if (attempt == 20)
                    throw new InvalidOperationException($"PowerPoint's {idMso} command is not available for the table cell's equation.");
                Thread.Sleep(25);
                System.Windows.Forms.Application.DoEvents();
            }
            bars.ExecuteMso(idMso);
        }
        finally { Release(bars); }
        System.Windows.Forms.Application.DoEvents();
    }

    private static string ReadText(Shape shape)
    {
        PptTextRange? text = null;
        try
        {
            text = shape.TextFrame.TextRange;
            return text.Text;
        }
        finally { Release(text); }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
    }
}

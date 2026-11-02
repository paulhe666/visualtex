using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using ColorFormat = Microsoft.Office.Interop.PowerPoint.ColorFormat;
using Font = Microsoft.Office.Interop.PowerPoint.Font;
using PptTextFrame = Microsoft.Office.Interop.PowerPoint.TextFrame;
using PptTextRange = Microsoft.Office.Interop.PowerPoint.TextRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// Native equations in WPS Presentation, which has neither MathZones nor a way
/// to paste PowerPoint's equation clipboard package: an equation is found by
/// matching WPS's saved XML of the shape to its live text (WpsMathText), and
/// written by pasting WPS's own copy of it from a windowless scratch
/// presentation that imported a one-slide file (WpsEquationSlide). The user's
/// presentation never gets a temporary slide, so Undo only sees the paste.
/// </summary>
internal static class WpsNativeEquations
{
    internal static (XElement Shape, IReadOnlyList<(int Start, int Length)> Zones) Locate(
        Presentation presentation, int slideId, Shape shape, string? savedShapeXml = null)
    {
        var xml = XElement.Parse(
            savedShapeXml ?? PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slideId, shape.Id),
            LoadOptions.PreserveWhitespace);
        return (xml, WpsMathText.Locate(xml, LiveText(shape)) ?? throw Unlocatable());
    }

    /// <summary>The equations of a shape, read through its own slide and presentation.</summary>
    internal static IReadOnlyList<(int Start, int Length)> Locate(Shape shape)
    {
        Slide? slide = null;
        Presentation? presentation = null;
        try
        {
            slide = shape.Parent as Slide
                ?? throw new InvalidOperationException("The WPS equation shape is no longer on a slide.");
            presentation = slide.Parent as Presentation
                ?? throw new InvalidOperationException("The WPS equation slide is no longer in a presentation.");
            return Locate(presentation, slide.SlideID, shape).Zones;
        }
        finally { Release(presentation); Release(slide); }
    }

    internal static IReadOnlyList<(int Start, int Length)> Locate(XElement savedShape, Shape shape) =>
        WpsMathText.Locate(savedShape, LiveText(shape)) ?? throw Unlocatable();

    private static Exception Unlocatable() =>
        new InvalidOperationException(OfficePluginLanguage.Text(
            "无法确定这个文本框中原生公式的位置（WPS 的文本与其保存内容不一致）。",
            "The native equations in this text box could not be located (WPS's text does not match its saved content)."));

    internal static string LiveText(Shape shape)
    {
        PptTextFrame? frame = null;
        PptTextRange? text = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return string.Empty;
            frame = shape.TextFrame;
            text = frame.TextRange;
            return text.Text ?? string.Empty;
        }
        finally { Release(text); Release(frame); }
    }

    /// <summary>The character range of a shape's text (WPS has no math-zone ranges).</summary>
    internal static PptTextRange Characters(Shape shape, int start, int length)
    {
        PptTextFrame? frame = null;
        PptTextRange? text = null;
        try
        {
            frame = shape.TextFrame;
            text = frame.TextRange;
            return text.Characters(start, length);
        }
        finally { Release(text); Release(frame); }
    }

    internal static List<(float Left, float Top, float Width, float Height, float FontSize)> Geometry(
        Shape shape, IReadOnlyList<(int Start, int Length)> zones)
    {
        var result = new List<(float Left, float Top, float Width, float Height, float FontSize)>();
        foreach (var (start, length) in zones)
        {
            PptTextRange? range = null;
            Font? font = null;
            try
            {
                range = Characters(shape, start, length);
                font = range.Font;
                float size;
                try { size = font.Size; }
                catch { size = 0; }
                // A zone with mixed run sizes reports an undefined size.
                if (!(size > 0) || size > 4000) size = 0;
                result.Add((range.BoundLeft, range.BoundTop, range.BoundWidth, range.BoundHeight, size));
            }
            finally { Release(font); Release(range); }
        }
        return result;
    }

    // Per equation, the color and length of each run as WPS draws it.
    internal static List<List<(string Color, int Length)>> RunColors(Shape shape, IReadOnlyList<(int Start, int Length)> zones)
    {
        var result = new List<List<(string Color, int Length)>>();
        foreach (var (start, length) in zones)
        {
            var colors = new List<(string Color, int Length)>();
            PptTextRange? range = null;
            PptTextRange? runs = null;
            try
            {
                range = Characters(shape, start, length);
                runs = range.Runs();
                for (var position = 1; position <= runs.Count; position++)
                {
                    PptTextRange? run = null;
                    Font? font = null;
                    ColorFormat? color = null;
                    try
                    {
                        run = range.Runs(position, 1);
                        font = run.Font;
                        color = font.Color;
                        colors.Add((PowerPointConversionColors.FromOfficeRgb(color.RGB), run.Length));
                    }
                    catch { }
                    finally { Release(color); Release(font); Release(run); }
                }
            }
            catch { }
            finally { Release(runs); Release(range); }
            result.Add(colors);
        }
        return result;
    }

    /// <summary>
    /// Pastes the equation <paramref name="paragraphs"/> over [start, start +
    /// length) of the shape's text and returns the pasted range. The shape's slide
    /// must be in the active window; the caller preserves the user's clipboard.
    /// <paramref name="beforePaste"/> runs once everything that could fail
    /// without touching the text has succeeded.
    /// </summary>
    internal static (int Start, int Length) Paste(Application application, Shape shape, int start, int length,
        string paragraphs, Action? beforePaste)
    {
        var path = Path.Combine(Path.GetTempPath(), $"visualtex-wps-equation-{Guid.NewGuid():N}.pptx");
        Presentations? presentations = null;
        Presentation? scratch = null;
        try
        {
            File.WriteAllBytes(path, WpsEquationSlide.Build(paragraphs));
            presentations = application.Presentations;
            scratch = presentations.Add(MsoTriState.msoFalse);
            var copied = CopyEquation(scratch, path);
            beforePaste?.Invoke();
            return PasteCopied(application, shape, start, length, copied);
        }
        finally
        {
            if (scratch is not null)
            {
                try { scratch.Saved = MsoTriState.msoTrue; } catch { }
                try { scratch.Close(); } catch { }
            }
            Release(scratch);
            Release(presentations);
            try { File.Delete(path); } catch { }
        }
    }

    // Returns the copied equation as WPS spells it in text.
    private static string CopyEquation(Presentation scratch, string path)
    {
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        Shape? box = null;
        PptTextFrame? frame = null;
        PptTextRange? text = null;
        try
        {
            slides = scratch.Slides;
            slides.InsertFromFile(path, 0, 1, 1);
            if (slides.Count != 1)
                throw new InvalidOperationException("WPS did not import the native equation slide.");
            slide = slides[1];
            shapes = slide.Shapes;
            if (shapes.Count != 1)
                throw new InvalidOperationException("WPS imported the native equation slide without its text box.");
            box = shapes[1];
            frame = box.TextFrame;
            text = frame.TextRange;
            if (text.Length == 0)
                throw new InvalidOperationException("WPS imported the native equation slide without its equation.");
            var spelled = text.Text ?? string.Empty;
            // A clipboard listener (input method, clipboard history) holding the
            // clipboard makes the copy fail: serve it and retry.
            for (var attempt = 1; ; attempt++)
            {
                try
                {
                    text.Copy();
                    PowerPointClipboard.ServeSentMessages(20);
                    return spelled;
                }
                catch (COMException) when (attempt < 5)
                {
                    PowerPointClipboard.ServeSentMessages(100 * attempt);
                }
                catch (COMException error)
                {
                    throw new InvalidOperationException(
                        $"{error.Message} (clipboard held by: {PowerPointClipboard.Holder()})", error);
                }
            }
        }
        finally
        {
            Release(text);
            Release(frame);
            Release(box);
            Release(shapes);
            Release(slide);
            Release(slides);
        }
    }

    // Selects the target range and pastes there with the window's own Paste (a
    // TextRange.Paste in WPS always appends at the end of the text). Nothing is
    // pasted unless WPS selected exactly the target, and the text around the
    // pasted range must be exactly the text that was around the target. An edit
    // that only restyled its equation pastes the same spelling over itself, so
    // the text stays as it was: that paste is not repeated.
    private static (int Start, int Length) PasteCopied(Application application, Shape shape, int start, int length,
        string copied)
    {
        var original = LiveText(shape);
        if (start < 1 || length < 0 || start - 1 + length > original.Length)
            throw new InvalidOperationException("The WPS text insertion position is no longer valid.");
        PptTextFrame? frame = null;
        PptTextRange? text = null;
        PptTextRange? target = null;
        try
        {
            frame = shape.TextFrame;
            text = frame.TextRange;
            target = original.Length == 0 ? text : text.Characters(start, length);
            target.Select();
        }
        finally
        {
            if (!ReferenceEquals(target, text)) Release(target);
            Release(text);
            Release(frame);
        }
        RequireSelection(application, shape, start, length);
        var same = string.Equals(copied, original.Substring(start - 1, length), StringComparison.Ordinal);

        for (var attempt = 1; ; attempt++)
        {
            DocumentWindow? window = null;
            View? view = null;
            try
            {
                window = application.ActiveWindow;
                view = window.View;
                view.Paste();
            }
            catch (COMException) when (attempt < 3) { }
            finally { Release(view); Release(window); }

            var after = LiveText(shape);
            if (same || !string.Equals(after, original, StringComparison.Ordinal))
            {
                var pasted = after.Length - (original.Length - length);
                if (pasted <= 0
                    || !after.StartsWith(original.Substring(0, start - 1), StringComparison.Ordinal)
                    || !after.EndsWith(original.Substring(start - 1 + length), StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        "WPS did not paste the native equation at the expected text position.");
                return (start, pasted);
            }
            if (attempt >= 3)
                throw new InvalidOperationException(
                    $"WPS did not paste the native equation (clipboard held by: {PowerPointClipboard.Holder()}).");
            PowerPointClipboard.ServeSentMessages(100 * attempt);
        }
    }

    private static void RequireSelection(Application application, Shape shape, int start, int length)
    {
        DocumentWindow? window = null;
        Selection? selection = null;
        ShapeRange? shapes = null;
        Shape? selectedShape = null;
        PptTextRange? selected = null;
        try
        {
            window = application.ActiveWindow;
            selection = window.Selection;
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                shapes = selection.ShapeRange;
                selectedShape = shapes.Count == 1 ? shapes[1] : null;
                selected = selection.TextRange;
                if (selectedShape is not null && selectedShape.Id == shape.Id
                    && selected.Start == start && selected.Length == length)
                    return;
            }
            throw new InvalidOperationException(
                "WPS did not select the exact text the native equation replaces; nothing was pasted.");
        }
        finally
        {
            Release(selected);
            Release(selectedShape);
            Release(shapes);
            Release(selection);
            Release(window);
        }
    }

    private static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            try { Marshal.ReleaseComObject(value); } catch { }
        }
    }
}

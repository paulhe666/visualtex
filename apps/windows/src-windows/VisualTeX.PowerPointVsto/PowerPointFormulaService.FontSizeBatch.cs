using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using PptFont = Microsoft.Office.Interop.PowerPoint.Font;
using PptTextRange = Microsoft.Office.Interop.PowerPoint.TextRange;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace VisualTeX.PowerPointVsto;

// One size for every inline or every display formula of the presentation (Ribbon 行内 / 行间, as in Word and on the
// Mac), in one Undo step: VisualTeX objects and pictures are scaled as the single-formula command scales them, native
// equations take the size as their text does. MathType objects are left as they are (their size is MathType's own).
public sealed partial class PowerPointFormulaService
{
    internal sealed class FormulaFontSizeBatchResult
    {
        internal int Resized { get; set; }
        internal int MathTypeLeft { get; set; }
        internal List<string> Failures { get; } = new();
    }

    internal FormulaFontSizeBatchResult ApplyFontSizeToPresentationFormulas(bool display, double requestedFontSizePt)
    {
        var target = FormulaFontSize.Normalize(requestedFontSizePt);
        var result = new FormulaFontSizeBatchResult();
        Presentation? presentation = null;
        Slides? slides = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException(Text("没有打开的演示文稿。", "No presentation is open."));
            EnsureWritable(presentation);
            var formulas = ScanAutomationFormulas(presentation, new HashSet<string>(), 0, slideId: null)
                .Where(formula => string.Equals(formula.Display, "block", StringComparison.Ordinal) == display)
                .ToList();
            result.MathTypeLeft = formulas.Count(formula => formula.Kind == AutomationMathTypeKind);
            formulas.RemoveAll(formula => formula.Kind == AutomationMathTypeKind || formula.Error is not null);
            if (formulas.Count == 0) return result;
            slides = presentation.Slides;
            StartNewUndoEntry();
            foreach (var group in formulas.GroupBy(formula => (formula.SlideId, formula.ShapeId)))
            {
                Slide? slide = null;
                Shapes? shapes = null;
                Shape? shape = null;
                try
                {
                    slide = FindSlideById(slides, group.Key.SlideId);
                    shapes = slide.Shapes;
                    shape = ShapeById(shapes, group.Key.ShapeId)
                        ?? throw new InvalidOperationException(Text("公式所在的对象已不存在。", "The formula's shape no longer exists."));
                    if (group.First().Kind == AutomationOmmlKind)
                    {
                        result.Resized += ResizeEquations(
                            presentation, group.Key.SlideId, shape, group.Select(formula => formula.Ordinal).ToList(), target);
                    }
                    else
                    {
                        var metadata = ReadPictureMetadata(shape)
                            ?? throw new InvalidOperationException(Text("无法读取公式数据。", "The formula's data could not be read."));
                        ResizeFormulaShape(shape, metadata, target);
                        result.Resized++;
                    }
                }
                catch (Exception error)
                {
                    result.Failures.Add(OfficeErrorText.Describe(error));
                }
                finally
                {
                    Release(shape);
                    Release(shapes);
                    Release(slide);
                }
            }
            return result;
        }
        finally
        {
            StartNewUndoEntry();
            Release(slides);
            Release(presentation);
        }
    }

    private static Shape? ShapeById(Shapes shapes, int shapeId)
    {
        for (var index = 1; index <= shapes.Count; index++)
        {
            Shape? shape = null;
            try
            {
                shape = shapes[index];
                if (shape.Id != shapeId) continue;
                var found = shape;
                shape = null;
                return found;
            }
            finally { Release(shape); }
        }
        return null;
    }

    // A text box's native equations (1-based, text order) at the size: MathZones in PowerPoint; in WPS, which has
    // none, the characters its saved XML places them at.
    private static int ResizeEquations(Presentation presentation, int slideId, Shape shape, List<int> ordinals, float size)
    {
        var resized = 0;
        if (OfficeHostEnvironment.IsWps)
        {
            var zones = WpsNativeEquations.Locate(presentation, slideId, shape).Zones;
            foreach (var ordinal in ordinals)
            {
                if (ordinal < 1 || ordinal > zones.Count) continue;
                PptTextRange? range = null;
                PptFont? font = null;
                try
                {
                    range = WpsNativeEquations.Characters(shape, zones[ordinal - 1].Start, zones[ordinal - 1].Length);
                    font = range.Font;
                    font.Size = size;
                    resized++;
                }
                finally
                {
                    Release(font);
                    Release(range);
                }
            }
            return resized;
        }
        object? frame = null, text = null;
        try
        {
            frame = shape.TextFrame2;
            text = ((dynamic)frame).TextRange;
            foreach (var ordinal in ordinals)
            {
                object? zone = null, font = null;
                try
                {
                    zone = ((dynamic)text).MathZones(ordinal, 1);
                    font = ((dynamic)zone).Font;
                    ((dynamic)font).Size = size;
                    resized++;
                }
                finally
                {
                    Release(font);
                    Release(zone);
                }
            }
        }
        finally
        {
            Release(text);
            Release(frame);
        }
        return resized;
    }
}

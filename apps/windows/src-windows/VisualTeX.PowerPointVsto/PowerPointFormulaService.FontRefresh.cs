using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace VisualTeX.PowerPointVsto;

// Update Equation Format in PowerPoint: the VisualTeX formulas (SVG pictures and OLE objects) drawn in other fonts
// than the ones the main app draws in now are drawn again from their LaTeX and put where they were - the same id,
// form, centre, rotation and layer, at the size they show at, the width following the new glyphs - in one Undo step.
public sealed partial class PowerPointFormulaService
{
    /// <summary>A VisualTeX formula drawn in other fonts, as the scan read it.</summary>
    internal sealed class FontRefreshTarget
    {
        internal int SlideId { get; set; }
        internal int SlideIndex { get; set; }
        internal int ShapeId { get; set; }
        internal string DocumentId { get; set; } = string.Empty;
        internal string ObjectMode { get; set; } = FormulaOleContract.CrossPlatformPictureMode;
        internal FormulaMetadata Metadata { get; set; } = new();
        /// <summary>The size the formula shows at (its shape's), which it is drawn again at.</summary>
        internal float EffectiveFontSizePt { get; set; }
    }

    /// <summary>A formula drawn again: its session (export result, metadata) and the files to write.</summary>
    internal sealed class FontRefreshWrite
    {
        internal FontRefreshTarget Target { get; set; } = new();
        internal OfficeSessionDocument Session { get; set; } = new();
        internal float PlacementScale { get; set; } = 1f;
        internal string? PngPath { get; set; }
        internal string? EmfPath { get; set; }
        internal string? SvgPath { get; set; }
    }

    /// <summary>The presentation's VisualTeX pictures and OLE objects whose recorded fonts differ from <paramref name="current"/>.</summary>
    internal IReadOnlyList<FontRefreshTarget> CaptureFontRefreshTargets(FormulaFontSettings current)
    {
        Presentation? presentation = null;
        Slides? slides = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException(Text("没有打开的演示文稿。", "No presentation is open."));
            var documentId = DocumentIdentity(presentation);
            slides = presentation.Slides;
            var targets = new List<FontRefreshTarget>();
            for (var slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    shapes = slide.Shapes;
                    for (var index = 1; index <= shapes.Count; index++)
                    {
                        Shape? shape = null;
                        try
                        {
                            shape = shapes[index];
                            // VisualTeX's own forms only: its OLE objects and its pictures (an SVG picture is msoGraphic, 28)
                            var ole = IsNativeOle(shape);
                            var type = shape.Type;
                            if (!ole && type != MsoShapeType.msoPicture && (int)type != 28) continue;
                            // the metadata tag every VisualTeX object carries (no OLE activation)
                            var metadata = ReadPictureMetadata(shape);
                            if (metadata is null
                                || string.IsNullOrWhiteSpace(metadata.Latex)
                                || !Guid.TryParse(metadata.FormulaId, out _)
                                || !current.DifferFrom(metadata))
                                continue;
                            targets.Add(new FontRefreshTarget
                            {
                                SlideId = slide.SlideID,
                                SlideIndex = slideIndex,
                                ShapeId = shape.Id,
                                DocumentId = documentId,
                                ObjectMode = ole ? FormulaOleContract.NativeOleMode : FormulaOleContract.CrossPlatformPictureMode,
                                Metadata = metadata,
                                EffectiveFontSizePt = InferPowerPointFormulaFontSize(shape.Width, shape.Height, metadata),
                            });
                        }
                        finally { Release(shape); }
                    }
                }
                finally
                {
                    Release(shapes);
                    Release(slide);
                }
            }
            return targets;
        }
        finally
        {
            Release(slides);
            Release(presentation);
        }
    }

    /// <summary>
    /// Puts the formulas drawn again in place of the old ones, found again by slide and shape (and checked by id and
    /// LaTeX: one that changed since the scan stays as it is and is reported), in one Undo step.
    /// </summary>
    internal int ApplyFontRefresh(IReadOnlyList<FontRefreshWrite> writes, List<string> failures)
    {
        Presentation? presentation = null;
        Slides? slides = null;
        var written = 0;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException(Text("没有打开的演示文稿。", "No presentation is open."));
            EnsureWritable(presentation);
            slides = presentation.Slides;
            StartNewUndoEntry();
            foreach (var write in writes)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                Shape? shape = null;
                Shape? replacement = null;
                try
                {
                    slide = FindSlideById(slides, write.Target.SlideId);
                    shapes = slide.Shapes;
                    shape = ShapeById(shapes, write.Target.ShapeId);
                    var metadata = shape is null ? null : ReadPictureMetadata(shape);
                    if (shape is null
                        || metadata is null
                        || !string.Equals(metadata.FormulaId, write.Target.Metadata.FormulaId, StringComparison.OrdinalIgnoreCase)
                        || !string.Equals(metadata.Latex, write.Target.Metadata.Latex, StringComparison.Ordinal))
                        throw new InvalidOperationException(Text(
                            "公式在准备期间被修改或删除，保持原样。",
                            "The equation changed or was removed while it was being prepared and was left as it is."));
                    replacement = ReplaceInCurrentFonts(slide, shapes, shape, write);
                    written++;
                }
                catch (Exception error)
                {
                    failures.Add(Text($"第 {write.Target.SlideIndex} 张幻灯片：", $"Slide {write.Target.SlideIndex}: ")
                        + OfficeErrorText.Describe(error));
                }
                finally
                {
                    Release(replacement);
                    Release(shape);
                    Release(shapes);
                    Release(slide);
                }
            }
            return written;
        }
        finally
        {
            StartNewUndoEntry();
            Release(slides);
            Release(presentation);
        }
    }

    // As an edit puts a formula back (Replace / ReplaceOle), but at the size the formula shows at: the conversion's
    // placement (render size x placement scale, centred on the old object). The edit's rule for the same LaTeX keeps
    // the old box, which would squeeze glyphs of other widths into it.
    private Shape ReplaceInCurrentFonts(Slide slide, Shapes shapes, Shape oldShape, FontRefreshWrite write)
    {
        var session = write.Session;
        var metadata = session.ToMetadata();
        metadata.Validate();
        var export = session.ExportResult
            ?? throw new InvalidOperationException(Text("VisualTeX 没有返回重绘结果。", "VisualTeX returned no drawing."));
        var width = Math.Max(1f, export.Width * 0.75f) * write.PlacementScale;
        var height = Math.Max(1f, export.Height * 0.75f) * write.PlacementScale;
        var centerX = oldShape.Left + oldShape.Width / 2f;
        var centerY = oldShape.Top + oldShape.Height / 2f;
        var (left, top) = PowerPointConversionRules.CenteredAt(centerX, centerY, width, height);
        var rotation = oldShape.Rotation;
        var zOrder = oldShape.ZOrderPosition;
        Shape? replacement = null;
        try
        {
            if (string.Equals(write.Target.ObjectMode, FormulaOleContract.NativeOleMode, StringComparison.Ordinal))
            {
                if (write.PngPath is null || write.EmfPath is null)
                    throw new InvalidOperationException(Text("缺少 VisualTeX 原生 OLE 预览图。", "VisualTeX native OLE preview is unavailable."));
                replacement = AddOleObjectOffscreen(slide, width, height);
                replacement.Visible = MsoTriState.msoFalse;
                InitializeOle(replacement, metadata, write.EmfPath, write.PngPath);
                TryApplyRotation(replacement, rotation);
                Configure(replacement, metadata);
                MoveToZOrder(replacement, zOrder + 1);
                ApplyOleSizeAndRefresh(replacement, width, height);
                RestoreOlePosition(replacement, left, top);
                replacement.Visible = MsoTriState.msoTrue;
                DeleteReplacedShape(slide, oldShape, replacement);
                ScheduleOleGeometryRestore(
                    write.Target.DocumentId,
                    metadata.FormulaId,
                    ManagedReference(slide, replacement),
                    width,
                    height,
                    left,
                    top);
            }
            else
            {
                if (write.SvgPath is null)
                    throw new InvalidOperationException(Text("缺少 VisualTeX 图片导出结果。", "VisualTeX picture export is unavailable."));
                replacement = shapes.AddPicture(write.SvgPath, MsoTriState.msoFalse, MsoTriState.msoTrue, left, top, width, height);
                ApplyPictureSize(replacement, width, height);
                replacement.Left = left;
                replacement.Top = top;
                TryApplyRotation(replacement, rotation);
                Configure(replacement, metadata);
                MoveToZOrder(replacement, zOrder + 1);
                DeleteReplacedShape(slide, oldShape, replacement);
            }
            var result = replacement;
            replacement = null;
            return result;
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
    }
}

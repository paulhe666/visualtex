using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

internal sealed class PowerPointFormulaService
{
    private const string FormulaIdTag = "VisualTeXFormulaId";
    private const string MetadataTag = "VisualTeXMetadata";
    private const string SlideReferencePrefix = "visualtex-ppt-vsto-slide:";
    private readonly Application _application;

    public PowerPointFormulaService(Application application)
    {
        _application = application;
    }

    public OfficeSelection ReadSelection()
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Selection? selection = null;
        ShapeRange? range = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = (Slide)view.Slide;
            selection = window.Selection;
            FormulaMetadata? metadata = null;
            string? objectId = SlideReference(slide);
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                range = selection.ShapeRange;
                if (range.Count == 1)
                {
                    shape = range[1];
                    objectId = shape.Name;
                    metadata = ReadMetadata(shape);
                }
            }
            return new OfficeSelection
            {
                Host = "powerpoint",
                DocumentId = DocumentIdentity(presentation),
                ObjectId = objectId,
                ReadOnly = presentation.ReadOnly == MsoTriState.msoTrue,
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
            };
        }
        finally
        {
            Release(shape);
            Release(range);
            Release(selection);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeObjectResult Insert(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = ResolveTargetSlide(presentation, session.SourceObjectId, view);
            var width = Math.Max(12f, (session.ExportResult?.Width ?? 240) * 0.75f);
            var height = Math.Max(12f, (session.ExportResult?.Height ?? 80) * 0.75f);
            var scale = Math.Min(1f, Math.Min(600f / width, 400f / height));
            width *= scale;
            height *= scale;
            var left = Math.Max(0f, (presentation.PageSetup.SlideWidth - width) / 2f);
            var top = Math.Max(0f, (presentation.PageSetup.SlideHeight - height) / 2f);
            shape = slide.Shapes.AddPicture(
                imagePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                left,
                top,
                width,
                height);
            Configure(shape, metadata);
            return Result(session, presentation, shape.Name);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeObjectResult Replace(OfficeSessionDocument session, string imagePath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? oldShape = null;
        Shape? replacement = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            (slide, oldShape) = FindFormula(
                presentation,
                session.FormulaId,
                session.SourceObjectId);
            if (slide is null || oldShape is null)
                throw new InvalidOperationException("The target PowerPoint formula no longer exists.");

            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            var exportWidth = Math.Max(1f, session.ExportResult?.Width ?? oldWidth);
            var exportHeight = Math.Max(1f, session.ExportResult?.Height ?? oldHeight);
            var ratio = exportWidth / exportHeight;
            var width = oldWidth;
            var height = width / ratio;
            if (height > oldHeight)
            {
                height = oldHeight;
                width = height * ratio;
            }
            var newLeft = left + (oldWidth - width) / 2f;
            var newTop = top + (oldHeight - height) / 2f;

            replacement = slide.Shapes.AddPicture(
                imagePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                newLeft,
                newTop,
                width,
                height);
            replacement.Rotation = rotation;
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            oldShape.Delete();
            return Result(session, presentation, replacement.Name);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            StartNewUndoEntry();
            Release(replacement);
            Release(oldShape);
            Release(slide);
            Release(presentation);
        }
    }

    private void StartNewUndoEntry()
    {
        try { _application.StartNewUndoEntry(); } catch { }
    }

    private static (Slide? Slide, Shape? Shape) FindFormula(
        Presentation presentation,
        string formulaId,
        string? preferredObjectId)
    {
        Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (var slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    shapes = slide.Shapes;
                    for (var shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                    {
                        Shape? shape = null;
                        try
                        {
                            shape = shapes[shapeIndex];
                            var metadata = ReadMetadata(shape);
                            var preferredMatch = !string.IsNullOrEmpty(preferredObjectId)
                                && shape.Name == preferredObjectId;
                            var metadataMatch = metadata?.FormulaId == formulaId
                                || shape.Name == $"VisualTeX_{formulaId}";
                            if (metadataMatch || (preferredMatch && metadata?.FormulaId == formulaId))
                            {
                                var foundSlide = slide;
                                var foundShape = shape;
                                slide = null;
                                shape = null;
                                return (foundSlide, foundShape);
                            }
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
            return (null, null);
        }
        finally { Release(slides); }
    }

    private static FormulaMetadata? ReadMetadata(Shape shape)
    {
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            string? encoded = null;
            try { encoded = tags[MetadataTag]; } catch { }
            return FormulaMetadataCodec.Decode(encoded)
                ?? FormulaMetadataCodec.Decode(shape.AlternativeText);
        }
        finally { Release(tags); }
    }

    private static void Configure(Shape shape, FormulaMetadata metadata)
    {
        var encoded = FormulaMetadataCodec.Encode(metadata);
        shape.LockAspectRatio = MsoTriState.msoTrue;
        shape.Name = $"VisualTeX_{metadata.FormulaId}";
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            tags.Add(FormulaIdTag, metadata.FormulaId);
            tags.Add(MetadataTag, encoded);
        }
        finally { Release(tags); }
        shape.AlternativeText = encoded;
    }

    private static void MoveToZOrder(Shape shape, int target)
    {
        for (var attempts = 0; attempts < 512 && shape.ZOrderPosition > target; attempts++)
            shape.ZOrder(MsoZOrderCmd.msoSendBackward);
    }

    private static OfficeObjectResult Result(
        OfficeSessionDocument session,
        Presentation presentation,
        string objectId) =>
        new()
        {
            FormulaId = session.FormulaId,
            DocumentId = DocumentIdentity(presentation),
            ObjectId = objectId,
        };

    private static string SlideReference(Slide slide) =>
        $"{SlideReferencePrefix}{slide.SlideID}:{slide.SlideIndex}";

    private static Slide ResolveTargetSlide(
        Presentation presentation,
        string? sourceObjectId,
        View view)
    {
        if (TryParseSlideReference(sourceObjectId, out var slideId))
        {
            Slides? slides = null;
            try
            {
                slides = presentation.Slides;
                for (var index = 1; index <= slides.Count; index++)
                {
                    Slide? candidate = null;
                    try
                    {
                        candidate = slides[index];
                        if (candidate.SlideID != slideId) continue;
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { Release(candidate); }
                }
            }
            finally { Release(slides); }
            throw new InvalidOperationException(
                "The PowerPoint slide selected when the formula editor opened no longer exists.");
        }
        return (Slide)view.Slide;
    }

    private static bool TryParseSlideReference(string? value, out int slideId)
    {
        slideId = 0;
        if (string.IsNullOrWhiteSpace(value)
            || !value.StartsWith(SlideReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = value.Substring(SlideReferencePrefix.Length);
        var separator = payload.IndexOf(':');
        if (separator >= 0) payload = payload.Substring(0, separator);
        return int.TryParse(payload, out slideId) && slideId > 0;
    }

    private static void EnsureSourceDocument(
        Presentation presentation,
        string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(presentation);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active PowerPoint presentation changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(Presentation presentation)
    {
        try
        {
            if (!string.IsNullOrWhiteSpace(presentation.FullName)) return presentation.FullName;
        }
        catch { }
        return presentation.Name;
    }

    private static void EnsureWritable(Presentation presentation)
    {
        if (presentation.ReadOnly == MsoTriState.msoTrue)
            throw new UnauthorizedAccessException("The active PowerPoint presentation is read-only.");
    }

    private void EnsureNotSlideShow()
    {
        SlideShowWindows? windows = null;
        try
        {
            windows = _application.SlideShowWindows;
            if (windows.Count > 0)
                throw new InvalidOperationException("PowerPoint slide show mode does not allow formula editing.");
        }
        finally { Release(windows); }
    }

    private static void TryDelete(Shape? shape)
    {
        if (shape is null) return;
        try { shape.Delete(); } catch { }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); } catch { }
    }
}

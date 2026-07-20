using System.Drawing;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOleBridge;

internal sealed class PowerPointOleService : IPowerPointFormulaService
{
    private const int MsoFalse = 0;
    private const int MsoTrue = -1;
    private const int MsoSendBackward = 1;
    private const string FormulaIdTag = "VisualTeXFormulaId";
    private const string MetadataTag = "VisualTeXMetadata";
    private const string SlideReferencePrefix = "visualtex-ppt-ole-slide:";

    public OfficeSelection GetSelection()
    {
        object? app = null;
        object? presentation = null;
        object? window = null;
        object? view = null;
        object? slide = null;
        object? selection = null;
        object? shapeRange = null;
        object? shape = null;
        try
        {
            app = RunningOfficeLocator.GetPowerPointApplication();
            dynamic ppt = app;
            EnsureNotSlideShow(ppt);
            presentation = ppt.ActivePresentation;
            window = ppt.ActiveWindow;
            dynamic deck = presentation;
            dynamic activeWindow = window;
            view = activeWindow.View;
            slide = ((dynamic)view).Slide;
            selection = activeWindow.Selection;
            dynamic selected = selection;
            FormulaMetadata? metadata = null;
            string? objectId = SlideReference(slide);
            if ((int)selected.Type == 2 && (int)selected.ShapeRange.Count == 1)
            {
                shapeRange = selected.ShapeRange;
                shape = ((dynamic)shapeRange).Item(1);
                objectId = ((dynamic)shape).Name as string;
                metadata = ReadMetadata(shape);
            }
            return new OfficeSelection
            {
                Host = "powerpoint",
                DocumentId = DocumentIdentity(deck),
                ObjectId = objectId,
                ReadOnly = IsReadOnly(deck),
                FormulaId = metadata?.FormulaId,
                Metadata = metadata,
            };
        }
        finally
        {
            ComRelease.Final(shape);
            ComRelease.Final(shapeRange);
            ComRelease.Final(selection);
            ComRelease.Final(slide);
            ComRelease.Final(view);
            ComRelease.Final(window);
            ComRelease.Final(presentation);
            ComRelease.Final(app);
        }
    }

    public OfficeObjectResult InsertFormula(SessionInfo session)
    {
        session.Metadata.Validate();
        object? app = null;
        object? presentation = null;
        object? window = null;
        object? view = null;
        object? slide = null;
        object? shape = null;
        try
        {
            app = RunningOfficeLocator.GetPowerPointApplication();
            StartNewUndoEntry(app);
            dynamic ppt = app;
            EnsureNotSlideShow(ppt);
            presentation = ppt.ActivePresentation;
            dynamic deck = presentation;
            EnsureWritable(deck);
            EnsureSourceDocument(deck, session.SourceDocumentId);
            window = ppt.ActiveWindow;
            view = ((dynamic)window).View;
            slide = ResolveTargetSlide(deck, session.SourceObjectId, view);
            var size = FitImage(session.ImagePath, session.Width, session.Height);
            var slideWidth = Convert.ToSingle(deck.PageSetup.SlideWidth);
            var slideHeight = Convert.ToSingle(deck.PageSetup.SlideHeight);
            var left = Math.Max(0, (slideWidth - size.Width) / 2f);
            var top = Math.Max(0, (slideHeight - size.Height) / 2f);
            shape = ((dynamic)slide).Shapes.AddPicture(
                session.ImagePath,
                MsoFalse,
                MsoTrue,
                left,
                top,
                size.Width,
                size.Height);
            ConfigureShape(shape, session.Metadata);
            return new OfficeObjectResult
            {
                FormulaId = session.FormulaId,
                DocumentId = DocumentIdentity(deck),
                ObjectId = ((dynamic)shape).Name as string,
            };
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            StartNewUndoEntry(app);
            ComRelease.Final(shape);
            ComRelease.Final(slide);
            ComRelease.Final(view);
            ComRelease.Final(window);
            ComRelease.Final(presentation);
            ComRelease.Final(app);
        }
    }

    public OfficeObjectResult ReplaceFormula(SessionInfo session)
    {
        session.Metadata.Validate();
        object? app = null;
        object? presentation = null;
        object? slide = null;
        object? oldShape = null;
        object? newShape = null;
        try
        {
            app = RunningOfficeLocator.GetPowerPointApplication();
            StartNewUndoEntry(app);
            dynamic ppt = app;
            EnsureNotSlideShow(ppt);
            presentation = ppt.ActivePresentation;
            dynamic deck = presentation;
            EnsureWritable(deck);
            EnsureSourceDocument(deck, session.SourceDocumentId);
            var locatedFormula = FindFormula((object)deck, session.FormulaId, session.SourceObjectId);
            slide = locatedFormula.Slide;
            oldShape = locatedFormula.Shape;
            if (oldShape is null || slide is null)
                throw new InvalidOperationException("The target PowerPoint formula no longer exists.");

            dynamic original = oldShape;
            var left = Convert.ToSingle(original.Left);
            var top = Convert.ToSingle(original.Top);
            var width = Convert.ToSingle(original.Width);
            var height = Convert.ToSingle(original.Height);
            var rotation = Convert.ToSingle(original.Rotation);
            var zOrder = Convert.ToInt32(original.ZOrderPosition);
            var originalMetadata = ReadMetadata(oldShape);
            var size = ReplacementSize(
                session.ImagePath,
                width,
                height,
                originalMetadata,
                session.Metadata);
            var replacementLeft = left + (width - size.Width) / 2f;
            var replacementTop = top + (height - size.Height) / 2f;

            newShape = ReplacementTransaction.Execute<object>(
                () => ((dynamic)slide).Shapes.AddPicture(
                    session.ImagePath,
                    MsoFalse,
                    MsoTrue,
                    replacementLeft,
                    replacementTop,
                    size.Width,
                    size.Height),
                candidate =>
                {
                    ((dynamic)candidate).Rotation = rotation;
                    ConfigureShape(candidate, session.Metadata);
                    MoveToZOrder(candidate, zOrder + 1);
                },
                () => original.Delete(),
                TryDelete);
            return new OfficeObjectResult
            {
                FormulaId = session.FormulaId,
                DocumentId = DocumentIdentity(deck),
                ObjectId = ((dynamic)newShape).Name as string,
            };
        }
        catch
        {
            TryDelete(newShape);
            throw;
        }
        finally
        {
            StartNewUndoEntry(app);
            ComRelease.Final(newShape);
            ComRelease.Final(oldShape);
            ComRelease.Final(slide);
            ComRelease.Final(presentation);
            ComRelease.Final(app);
        }
    }

    public OfficeObjectResult MarkFormula(string formulaId)
    {
        var selection = GetSelection();
        if (selection.ObjectId is null)
            throw new InvalidOperationException("Select one PowerPoint shape before marking it.");
        object? app = null;
        object? presentation = null;
        object? slide = null;
        object? shape = null;
        try
        {
            app = RunningOfficeLocator.GetPowerPointApplication();
            StartNewUndoEntry(app);
            presentation = ((dynamic)app).ActivePresentation;
            var locatedFormula = FindFormula(presentation, formulaId, selection.ObjectId, allowNameOnly: true);
            slide = locatedFormula.Slide;
            shape = locatedFormula.Shape;
            if (shape is null) throw new InvalidOperationException("Selected shape was not found.");
            dynamic target = shape;
            target.Name = ShapeName(formulaId);
            object? tags = null;
            try
            {
                tags = target.Tags;
                ((dynamic)tags).Add(FormulaIdTag, formulaId);
            }
            finally { ComRelease.Final(tags); }
            return new OfficeObjectResult
            {
                FormulaId = formulaId,
                DocumentId = DocumentIdentity((dynamic)presentation),
                ObjectId = target.Name as string,
            };
        }
        finally
        {
            StartNewUndoEntry(app);
            ComRelease.Final(shape);
            ComRelease.Final(slide);
            ComRelease.Final(presentation);
            ComRelease.Final(app);
        }
    }

    public void DeleteFormula(string formulaId)
    {
        object? app = null;
        object? presentation = null;
        object? slide = null;
        object? shape = null;
        try
        {
            app = RunningOfficeLocator.GetPowerPointApplication();
            StartNewUndoEntry(app);
            presentation = ((dynamic)app).ActivePresentation;
            EnsureWritable((dynamic)presentation);
            var locatedFormula = FindFormula(presentation, formulaId, null);
            slide = locatedFormula.Slide;
            shape = locatedFormula.Shape;
            if (shape is null) throw new InvalidOperationException("The target formula no longer exists.");
            ((dynamic)shape).Delete();
        }
        finally
        {
            StartNewUndoEntry(app);
            ComRelease.Final(shape);
            ComRelease.Final(slide);
            ComRelease.Final(presentation);
            ComRelease.Final(app);
        }
    }

    private static void StartNewUndoEntry(object? application)
    {
        if (application is null) return;
        try { ((dynamic)application).StartNewUndoEntry(); } catch { }
    }

    private static (object? Slide, object? Shape) FindFormula(
        dynamic presentation,
        string formulaId,
        string? preferredObjectId,
        bool allowNameOnly = false)
    {
        object? slides = null;
        try
        {
            slides = presentation.Slides;
            var slideCount = Convert.ToInt32(((dynamic)slides).Count);
            for (var slideIndex = 1; slideIndex <= slideCount; slideIndex++)
            {
                object? slide = null;
                object? shapes = null;
                try
                {
                    slide = ((dynamic)slides).Item(slideIndex);
                    shapes = ((dynamic)slide).Shapes;
                    var shapeCount = Convert.ToInt32(((dynamic)shapes).Count);
                    for (var shapeIndex = 1; shapeIndex <= shapeCount; shapeIndex++)
                    {
                        object? shape = null;
                        try
                        {
                            shape = ((dynamic)shapes).Item(shapeIndex);
                            dynamic candidate = shape;
                            var name = candidate.Name as string;
                            var preferredMatch = !string.IsNullOrEmpty(preferredObjectId)
                                && string.Equals(name, preferredObjectId, StringComparison.Ordinal);
                            var metadata = ReadMetadata(shape);
                            var metadataMatch = metadata?.FormulaId == formulaId
                                || string.Equals(name, ShapeName(formulaId), StringComparison.Ordinal);
                            if (metadataMatch
                                || (preferredMatch && metadata?.FormulaId == formulaId)
                                || (allowNameOnly && preferredMatch))
                            {
                                var foundSlide = slide;
                                var foundShape = shape;
                                slide = null;
                                shape = null;
                                return (foundSlide, foundShape);
                            }
                        }
                        finally { ComRelease.Final(shape); }
                    }
                }
                finally
                {
                    ComRelease.Final(shapes);
                    ComRelease.Final(slide);
                }
            }
            return (null, null);
        }
        finally { ComRelease.Final(slides); }
    }

    private static FormulaMetadata? ReadMetadata(object shape)
    {
        object? tags = null;
        try
        {
            dynamic candidate = shape;
            tags = candidate.Tags;
            string? encoded = null;
            try { encoded = ((dynamic)tags).Item(MetadataTag) as string; } catch { }
            encoded ??= candidate.AlternativeText as string;
            return MetadataCodec.Decode(encoded);
        }
        finally { ComRelease.Final(tags); }
    }

    private static void ConfigureShape(object shape, FormulaMetadata metadata)
    {
        dynamic target = shape;
        var encoded = MetadataCodec.Encode(metadata);
        target.LockAspectRatio = MsoTrue;
        target.Name = ShapeName(metadata.FormulaId);
        object? tags = null;
        try
        {
            tags = target.Tags;
            ((dynamic)tags).Add(FormulaIdTag, metadata.FormulaId);
            ((dynamic)tags).Add(MetadataTag, encoded);
        }
        finally { ComRelease.Final(tags); }
        target.AlternativeText = encoded;
    }

    private static string ShapeName(string formulaId) => $"VisualTeX_{formulaId}";

    private static SizeF FitImage(string path, float maxWidth, float maxHeight)
    {
        using var image = Image.FromFile(path);
        var ratio = image.Width / (float)Math.Max(1, image.Height);
        var width = Math.Max(12f, maxWidth);
        var height = width / ratio;
        if (maxHeight > 0 && height > maxHeight)
        {
            height = maxHeight;
            width = height * ratio;
        }
        return new SizeF(width, height);
    }

    private static SizeF ReplacementSize(
        string path,
        float oldWidth,
        float oldHeight,
        FormulaMetadata? originalMetadata,
        FormulaMetadata replacementMetadata)
    {
        using var image = Image.FromFile(path);
        return CalculateReplacementSize(
            image.Width,
            image.Height,
            oldWidth,
            oldHeight,
            originalMetadata?.RenderWidthPx,
            originalMetadata?.RenderHeightPx,
            replacementMetadata.RenderWidthPx,
            replacementMetadata.RenderHeightPx);
    }

    internal static SizeF CalculateReplacementSize(
        int replacementPixelWidth,
        int replacementPixelHeight,
        float oldWidth,
        float oldHeight,
        double? originalRenderWidthPx,
        double? originalRenderHeightPx,
        double? replacementRenderWidthPx,
        double? replacementRenderHeightPx)
    {
        var replacementRatio = replacementPixelWidth /
            (float)Math.Max(1, replacementPixelHeight);
        var scale = PositiveFinite(originalRenderHeightPx)
            && PositiveFinite(replacementRenderWidthPx)
            && PositiveFinite(replacementRenderHeightPx)
            && oldHeight > 0
                ? oldHeight / (float)originalRenderHeightPx!.Value
                : PositiveFinite(originalRenderWidthPx)
                    && PositiveFinite(replacementRenderWidthPx)
                    && PositiveFinite(replacementRenderHeightPx)
                    && oldWidth > 0
                        ? oldWidth / (float)originalRenderWidthPx!.Value
                        : 0f;

        if (scale > 0 && float.IsFinite(scale))
        {
            return new SizeF(
                Math.Max(12f, (float)replacementRenderWidthPx!.Value * scale),
                Math.Max(1f, (float)replacementRenderHeightPx!.Value * scale));
        }

        // Formulas created by older VisualTeX builds do not have natural render
        // bounds in their metadata. Keep the existing physical height as the
        // font-size reference and let the replacement grow horizontally instead
        // of forcing it back into the old bounding box.
        var fallbackHeight = Math.Max(1f, oldHeight);
        return new SizeF(
            Math.Max(12f, fallbackHeight * replacementRatio),
            fallbackHeight);
    }

    private static bool PositiveFinite(double? value) =>
        value is > 0 && double.IsFinite(value.Value);

    private static void MoveToZOrder(object shape, int target)
    {
        dynamic candidate = shape;
        for (var attempts = 0; attempts < 512; attempts++)
        {
            var current = Convert.ToInt32(candidate.ZOrderPosition);
            if (current <= target) break;
            candidate.ZOrder(MsoSendBackward);
        }
    }

    private static string SlideReference(object slide)
    {
        dynamic value = slide;
        return $"{SlideReferencePrefix}{Convert.ToInt32(value.SlideID)}:{Convert.ToInt32(value.SlideIndex)}";
    }

    private static object ResolveTargetSlide(
        dynamic presentation,
        string? sourceObjectId,
        object view)
    {
        if (TryParseSlideReference(sourceObjectId, out var slideId))
        {
            object? slides = null;
            try
            {
                slides = presentation.Slides;
                var count = Convert.ToInt32(((dynamic)slides).Count);
                for (var index = 1; index <= count; index++)
                {
                    object? candidate = null;
                    try
                    {
                        candidate = ((dynamic)slides).Item(index);
                        if (Convert.ToInt32(((dynamic)candidate).SlideID) != slideId)
                            continue;
                        var result = candidate;
                        candidate = null;
                        return result;
                    }
                    finally { ComRelease.Final(candidate); }
                }
            }
            finally { ComRelease.Final(slides); }
            throw new InvalidOperationException(
                "The PowerPoint slide selected when the formula editor opened no longer exists.");
        }
        return ((dynamic)view).Slide;
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

    private static void EnsureSourceDocument(dynamic presentation, string? expectedIdentity)
    {
        if (string.IsNullOrWhiteSpace(expectedIdentity)) return;
        var actual = DocumentIdentity(presentation);
        if (!string.Equals(actual, expectedIdentity, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "The active PowerPoint presentation changed while the VisualTeX editor was open.");
    }

    private static string DocumentIdentity(dynamic presentation)
    {
        try
        {
            var fullName = presentation.FullName as string;
            if (!string.IsNullOrWhiteSpace(fullName)) return fullName;
        }
        catch { }
        return presentation.Name as string ?? "PowerPoint";
    }

    private static bool IsReadOnly(dynamic presentation)
    {
        try { return Convert.ToBoolean(presentation.ReadOnly); }
        catch { return false; }
    }

    private static void EnsureWritable(dynamic presentation)
    {
        if (IsReadOnly(presentation))
            throw new UnauthorizedAccessException("The active PowerPoint presentation is read-only.");
    }

    private static void EnsureNotSlideShow(dynamic app)
    {
        object? windows = null;
        try
        {
            windows = app.SlideShowWindows;
            if (Convert.ToInt32(((dynamic)windows).Count) > 0)
                throw new InvalidOperationException("PowerPoint slide show mode does not allow formula editing.");
        }
        finally { ComRelease.Final(windows); }
    }

    private static void TryDelete(object? shape)
    {
        if (shape is null) return;
        try { ((dynamic)shape).Delete(); } catch { }
    }
}

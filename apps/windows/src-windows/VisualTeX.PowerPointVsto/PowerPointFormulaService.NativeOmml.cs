using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace VisualTeX.PowerPointVsto;

public sealed partial class PowerPointFormulaService
{
    private const string NativeOmmlReferencePrefix = "visualtex-ppt-native-omml:";
    private NativeOmmlSource? _nativeOmmlSource;
    private sealed class NativeOmmlSource
    {
        internal string DocumentId = string.Empty;
        internal string ObjectId = string.Empty;
        internal string FormulaId = string.Empty;
        internal string Omml = string.Empty;
    }

    public static bool IsNativeOmmlReference(string? reference) =>
        TryParseNativeOmmlReference(reference, out _, out _);

    private static string NativeOmmlReference(Slide slide, Shape shape) =>
        $"{NativeOmmlReferencePrefix}{slide.SlideID}:{shape.Id}";

    private static bool TryParseNativeOmmlReference(string? reference, out int slideId, out int shapeId)
    {
        slideId = shapeId = 0;
        if (reference is null || !reference.StartsWith(NativeOmmlReferencePrefix, StringComparison.Ordinal)) return false;
        var parts = reference.Substring(NativeOmmlReferencePrefix.Length).Split(':');
        return parts.Length == 2 && int.TryParse(parts[0], out slideId) && slideId > 0
            && int.TryParse(parts[1], out shapeId) && shapeId > 0;
    }

    private static (Slide? Slide, Shape? Shape) FindNativeOmml(Presentation presentation, int slideId, int shapeId)
    {
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            slides = presentation.Slides;
            try { slide = slides.FindBySlideID(slideId); } catch { return (null, null); }
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                if (shape.Id == shapeId && shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var result = (slide, shape);
                    slide = null;
                    return result;
                }
                Release(shape);
            }
            return (null, null);
        }
        finally { Release(shapes); Release(slide); Release(slides); }
    }

    public OfficeSelection LoadNativeOmmlSelection(OfficeSelection selection)
    {
        if (!IsNativeOmmlReference(selection.ObjectId)) return selection;
        if (!TryParseNativeOmmlReference(selection.ObjectId, out var slideId, out var shapeId))
            throw new InvalidOperationException("The native equation selection has no exact slide and shape identity.");
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, selection.DocumentId);
            (slide, shape) = FindNativeOmml(presentation, slideId, shapeId);
            if (shape is null) throw new InvalidOperationException("The selected native equation no longer exists.");
            var omml = PowerPointNativeOmmlPackage.ReadStandalone(presentation, slideId, shapeId);
            var mathMl = PowerPointNativeOmmlKernel.ReadNativeMathMl(omml);
            var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
            if (string.IsNullOrWhiteSpace(latex)) throw new InvalidDataException("The selected native equation has no recoverable source.");
            var size = PowerPointOmmlBridge.ReadFontSize(shape) ?? 20f;
            var now = DateTimeOffset.UtcNow.ToString("O");
            var metadata = new FormulaMetadata
            {
                FormulaId = Guid.NewGuid().ToString("D"), Title = "PowerPoint Native Formula",
                Latex = latex,
                Lines = new List<FormulaLine> { new() { Id = Guid.NewGuid().ToString("D"), Latex = latex } },
                CodeFormat = "raw", DisplayMode = "block", Numbered = false,
                FontSizePt = size, RenderFontSizePt = size,
                RenderWidthPx = shape.Width / 0.75, RenderHeightPx = shape.Height / 0.75,
                CreatedWithVersion = "1.2.8", UpdatedWithVersion = "1.2.8",
                CreatedAt = now, UpdatedAt = now,
            };
            metadata.Validate();
            _nativeOmmlSource = new NativeOmmlSource
            {
                DocumentId = selection.DocumentId!, ObjectId = selection.ObjectId!,
                FormulaId = metadata.FormulaId, Omml = omml,
            };
            return new OfficeSelection
            {
                Host = "powerpoint", DocumentId = selection.DocumentId, ObjectId = selection.ObjectId,
                ReadOnly = selection.ReadOnly, ObjectMode = FormulaOleContract.WordOmmlMode,
                FormulaId = metadata.FormulaId, Metadata = metadata,
            };
        }
        finally { Release(shape); Release(slide); Release(presentation); }
    }

    private void EnsureNativeOmmlSourceUnchanged(OfficeSessionDocument session, Presentation presentation)
    {
        if (!TryParseNativeOmmlReference(session.SourceObjectId, out var slideId, out var shapeId)) return;
        var source = _nativeOmmlSource;
        if (source is null || source.DocumentId != session.SourceDocumentId
            || source.ObjectId != session.SourceObjectId || source.FormulaId != session.FormulaId)
            throw new InvalidOperationException("Reopen the native equation before saving it.");
        var current = PowerPointNativeOmmlPackage.ReadStandalone(presentation, slideId, shapeId);
        if (!string.Equals(source.Omml, current, StringComparison.Ordinal))
            throw new InvalidOperationException("The native equation changed while the editor was open. Reopen it before saving.");
    }

    private float? NativeShapeFontSize(OfficeSelection selection, float? requested = null)
    {
        if (!TryParseNativeOmmlReference(selection.ObjectId, out var slideId, out var shapeId)) return null;
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, selection.DocumentId);
            (slide, shape) = FindNativeOmml(presentation, slideId, shapeId);
            if (shape is null) return null;
            if (requested.HasValue)
            {
                EnsureWritable(presentation);
                // A selected text range is handled separately; changing a whole
                // object must never apply to adjacent prose or multiple equations.
                _ = PowerPointNativeOmmlPackage.ReadStandalone(presentation, slideId, shapeId);
                StartNewUndoEntry();
                PowerPointOmmlBridge.SetFontSize(shape, requested.Value);
                StartNewUndoEntry();
            }
            return PowerPointOmmlBridge.ReadFontSize(shape);
        }
        finally { Release(shape); Release(slide); Release(presentation); }
    }

}

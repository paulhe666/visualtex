using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.PowerPointVsto;

public sealed partial class PowerPointFormulaService
{
    private const string ManagedReferencePrefix = "visualtex-ppt-managed:";

    private static string ManagedReference(Slide slide, Shape shape) =>
        $"{ManagedReferencePrefix}{slide.SlideID}:{shape.Id}";

    private static (Slide? Slide, Shape? Shape) FindManagedFormula(
        Presentation presentation, string reference, string formulaId)
    {
        var parts = reference.Substring(ManagedReferencePrefix.Length).Split(':');
        if (parts.Length != 2 || !int.TryParse(parts[0], out var slideId) || slideId <= 0
            || !int.TryParse(parts[1], out var shapeId) || shapeId <= 0)
            throw new InvalidDataException("Invalid captured PowerPoint formula identity.");
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
                Shape? shape = shapes[index];
                try
                {
                    if (shape.Id != shapeId) continue;
                    if (!string.Equals(ReadMetadata(shape)?.FormulaId, formulaId, StringComparison.OrdinalIgnoreCase))
                        return (null, null);
                    var result = (slide, shape);
                    slide = null;
                    shape = null;
                    return result;
                }
                finally { Release(shape); }
            }
            return (null, null);
        }
        finally { Release(shapes); Release(slide); Release(slides); }
    }

    // Called only after the explicit edit command acquires the operation gate.
    // Passive reads never repair tags, rename shapes or scan for duplicate IDs.
    // Exact slide/shape identity keeps copied formulas distinct before repair.
    public OfficeSelection PrepareManagedSelectionForEdit(OfficeSelection selection)
    {
        if (selection.ObjectId?.StartsWith(ManagedReferencePrefix, StringComparison.Ordinal) != true)
            return selection;
        if (string.IsNullOrWhiteSpace(selection.DocumentId) || string.IsNullOrWhiteSpace(selection.FormulaId))
            throw new InvalidOperationException("The captured formula has no source identity.");
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation;
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, selection.DocumentId);
            (slide, shape) = FindManagedFormula(presentation, selection.ObjectId!, selection.FormulaId!);
            if (slide is null || shape is null)
                throw new InvalidOperationException("The captured formula no longer exists; another copy will not be substituted.");
            var metadata = ReadMetadata(shape)
                ?? throw new InvalidOperationException("The formula metadata is no longer available.");
            StartNewUndoEntry();
            metadata = EnsureUniqueFormulaIdentity(presentation, slide, shape, metadata);
            metadata.FontSizePt = InferPowerPointFormulaFontSize(shape.Width, shape.Height, metadata);
            return new OfficeSelection
            {
                Host = selection.Host, DocumentId = selection.DocumentId, ObjectId = selection.ObjectId,
                ReadOnly = selection.ReadOnly, FormulaId = metadata.FormulaId, Metadata = metadata,
                ObjectMode = selection.ObjectMode,
            };
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape); Release(slide); Release(presentation);
        }
    }
}

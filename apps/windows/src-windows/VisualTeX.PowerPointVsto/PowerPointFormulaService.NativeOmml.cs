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
        internal string Zone = string.Empty;
        internal int Start, Length;
        internal string? ShapeFingerprint;
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
        return (parts.Length == 2 || parts.Length == 4
                && int.TryParse(parts[2], out var start) && start > 0
                && int.TryParse(parts[3], out var length) && length >= 0)
            && int.TryParse(parts[0], out slideId) && slideId > 0
            && int.TryParse(parts[1], out shapeId) && shapeId > 0;
    }

    private static string NativeOmmlTextReference(Slide slide, Shape shape, object range) =>
        NativeOmmlReference(slide, shape) + ":" + Convert.ToInt32(((dynamic)range).Start)
            + ":" + Convert.ToInt32(((dynamic)range).Length);

    private static string? HitNativeMathZone(DocumentWindow window, Slide slide, Shape shape, int x, int y)
    {
        object? frame = null, range = null, zones = null;
        try
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue) return null;
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            zones = ((dynamic)range).MathZones();
            var count = Convert.ToInt32(((dynamic)zones).Count);
            for (var index = 1; index <= count; index++)
            {
                object? zone = null;
                try
                {
                    zone = ((dynamic)range).MathZones(index, 1);
                    float left = Convert.ToSingle(((dynamic)zone).BoundLeft), top = Convert.ToSingle(((dynamic)zone).BoundTop);
                    float right = left + Convert.ToSingle(((dynamic)zone).BoundWidth), bottom = top + Convert.ToSingle(((dynamic)zone).BoundHeight);
                    if (x >= window.PointsToScreenPixelsX(left) && x <= window.PointsToScreenPixelsX(right)
                        && y >= window.PointsToScreenPixelsY(top) && y <= window.PointsToScreenPixelsY(bottom))
                        return NativeOmmlTextReference(slide, shape, zone);
                }
                finally { Release(zone); }
            }
            return null;
        }
        finally { Release(zones); Release(range); Release(frame); }
    }

    // Math-zone offsets are transient coordinates, never persisted metadata.
    // They select one native equation while leaving all neighboring text intact.
    private static (int Ordinal, int Start, int Length, int Count) ResolveNativeMathZone(Shape shape, string reference)
    {
        var parts = reference.Substring(NativeOmmlReferencePrefix.Length).Split(':');
        var selectedStart = parts.Length == 4 ? int.Parse(parts[2]) : 0;
        var selectedLength = parts.Length == 4 ? int.Parse(parts[3]) : 0;
        object? frame = null, range = null, zones = null;
        try
        {
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            zones = ((dynamic)range).MathZones();
            var count = Convert.ToInt32(((dynamic)zones).Count);
            var matches = new List<(int Ordinal, int Start, int Length, int Count)>();
            for (var index = 1; index <= count; index++)
            {
                object? zone = null;
                try
                {
                    zone = ((dynamic)range).MathZones(index, 1);
                    int start = Convert.ToInt32(((dynamic)zone).Start), length = Convert.ToInt32(((dynamic)zone).Length);
                    if (selectedStart == 0 || (selectedLength == 0
                        ? selectedStart >= start && selectedStart < start + length
                        : selectedStart < start + length && selectedStart + selectedLength > start))
                        matches.Add((index, start, length, count));
                }
                finally { Release(zone); }
            }
            if (matches.Count != 1)
                throw new InvalidOperationException(Text(
                    "请把光标放入要编辑的一个原生公式中，或只选择该公式。",
                    "Place the caret inside the one native equation to edit, or select only that equation."));
            return matches[0];
        }
        finally { Release(zones); Release(range); Release(frame); }
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
            var zone = ResolveNativeMathZone(shape, selection.ObjectId!);
            var shapeXml = PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slideId, shapeId);
            var omml = PowerPointNativeOmmlPackage.ReadEquation(shapeXml, zone.Ordinal, zone.Count);
            var nativeZone = PowerPointNativeOmmlPackage.ReadEquationZone(shapeXml, zone.Ordinal, zone.Count);
            var mathMl = PowerPointNativeOmmlKernel.ReadNativeMathMl(omml);
            var latex = MathMlToLatexConverter.ConvertWordEquation(mathMl).Trim();
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
                FormulaId = metadata.FormulaId, Omml = omml, Zone = nativeZone,
                Start = zone.Start, Length = zone.Length,
                ShapeFingerprint = PowerPointTextEditTransaction.Fingerprint(shapeXml),
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

    private OfficeObjectResult ReplaceNativeOmmlText(OfficeSessionDocument session, Presentation presentation,
        Slide slide, Shape shape, string mathMl)
    {
        var source = _nativeOmmlSource;
        if (source is null || source.DocumentId != session.SourceDocumentId
            || source.ObjectId != session.SourceObjectId || source.FormulaId != session.FormulaId
            || source.ShapeFingerprint is null)
            throw new InvalidOperationException("Reopen the native equation before saving it.");
        // Replace exactly the existing math zone. Its paragraph and neighboring
        // content already exist; inserting a new display paragraph would split them.
        var write = PowerPointNativeOmmlKernel.PrepareTextInsertion(
            _application, shape, source.Start, source.Length, mathMl, "inline", source.Zone);
        PowerPointTextEditTransaction.Execute(_application, presentation, slide.SlideID, shape.Id,
            source.ShapeFingerprint, write);
        return Result(session, presentation, shape.Name);
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

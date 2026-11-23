using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace VisualTeX.PowerPointVsto;

public sealed partial class PowerPointFormulaService
{
    private const string NativeOmmlReferencePrefix = "visualtex-ppt-native-omml:";
    private static string TableCellHoldsOnlyTextMessage => Text(
        "PowerPoint 表格单元格只能容纳文字和原生公式（OMML），不能放入图片或 OLE 对象；单元格中的公式只能以原生公式形式编辑。",
        "A PowerPoint table cell holds only text and native (OMML) equations, not pictures or OLE objects: an equation in a cell can only be edited as a native equation.");
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
        NativeOmmlReference(slide.SlideID, new PowerPointTextHost(shape.Id));

    private static string NativeOmmlReference(int slideId, PowerPointTextHost host) =>
        $"{NativeOmmlReferencePrefix}{slideId}:{host}";

    private static bool TryParseNativeOmmlReference(string? reference, out int slideId, out PowerPointTextHost host)
    {
        slideId = 0;
        host = default;
        if (reference is null || !reference.StartsWith(NativeOmmlReferencePrefix, StringComparison.Ordinal)) return false;
        var parts = reference.Substring(NativeOmmlReferencePrefix.Length).Split(':');
        return (parts.Length == 2 || parts.Length == 4
                && int.TryParse(parts[2], out var start) && start > 0
                && int.TryParse(parts[3], out var length) && length >= 0)
            && int.TryParse(parts[0], out slideId) && slideId > 0
            && PowerPointTextHost.TryParse(parts[1], out host);
    }

    private static string NativeOmmlTextReference(Slide slide, Shape shape, object range) =>
        NativeOmmlTextReference(slide.SlideID, new PowerPointTextHost(shape.Id), range);

    private static string NativeOmmlTextReference(int slideId, PowerPointTextHost host, object range) =>
        NativeOmmlReference(slideId, host) + ":" + Convert.ToInt32(((dynamic)range).Start)
            + ":" + Convert.ToInt32(((dynamic)range).Length);

    // Equations in table cells are written through PowerPoint's own math text;
    // WPS Presentation's tables are not supported.
    private static void EnsureTextHostSupported(PowerPointTextHost host)
    {
        if (host.IsCell && OfficeHostEnvironment.IsWps)
            throw new InvalidOperationException(Text(
                "WPS 演示的表格单元格中暂不支持 VisualTeX 公式。",
                "VisualTeX formulas are not supported in WPS Presentation table cells yet."));
    }

    private static string? HitNativeMathZone(DocumentWindow window, Slide slide, Shape shape, int x, int y)
    {
        if (OfficeHostEnvironment.IsWps) return HitNativeMathZoneInWps(window, slide, shape, x, y);
        if (shape.HasTextFrame != MsoTriState.msoTrue) return null;
        var zone = HitMathZone(window, shape, x, y);
        return zone is null ? null
            : $"{NativeOmmlReference(slide.SlideID, new PowerPointTextHost(shape.Id))}:{zone.Value.Start}:{zone.Value.Length}";
    }

    // RangeFromPoint finds nothing over a table: the cell under the point is
    // searched for an equation instead.
    private static string? HitNativeMathZoneInTables(DocumentWindow window, Slide slide, int x, int y)
    {
        if (OfficeHostEnvironment.IsWps) return null;
        Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Shape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (shape.HasTable != MsoTriState.msoTrue
                        || !ContainsScreenPoint(window, shape.Left, shape.Top, shape.Width, shape.Height, x, y))
                        continue;
                    var hit = HitMathZoneInTable(window, shape, x, y);
                    if (hit is not null)
                        return $"{NativeOmmlReference(slide.SlideID, hit.Value.Host)}:{hit.Value.Start}:{hit.Value.Length}";
                }
                finally { Release(shape); }
            }
            return null;
        }
        finally { Release(shapes); }
    }

    private static (PowerPointTextHost Host, int Start, int Length)? HitMathZoneInTable(
        DocumentWindow window, Shape tableShape, int x, int y)
    {
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        try
        {
            table = tableShape.Table;
            rows = table.Rows;
            columns = table.Columns;
            int rowCount = rows.Count, columnCount = columns.Count;
            for (var row = 1; row <= rowCount; row++)
            {
                for (var column = 1; column <= columnCount; column++)
                {
                    Cell? cell = null;
                    Shape? cellShape = null;
                    try
                    {
                        cell = table.Cell(row, column);
                        cellShape = cell.Shape;
                        // The first grid position of a merged cell is the one holding its text.
                        if (!ContainsScreenPoint(window, cellShape.Left, cellShape.Top, cellShape.Width, cellShape.Height, x, y))
                            continue;
                        // A cell's text reports its bounds from the table's corner.
                        var zone = HitMathZone(window, cellShape, x, y, tableShape.Left, tableShape.Top);
                        return zone is null ? null
                            : (new PowerPointTextHost(tableShape.Id, row, column), zone.Value.Start, zone.Value.Length);
                    }
                    finally { Release(cellShape); Release(cell); }
                }
            }
            return null;
        }
        catch (System.Runtime.InteropServices.COMException) { return null; }
        finally { Release(columns); Release(rows); Release(table); }
    }

    private static bool ContainsScreenPoint(DocumentWindow window, float left, float top, float width, float height, int x, int y) =>
        x >= window.PointsToScreenPixelsX(left) && x <= window.PointsToScreenPixelsX(left + width)
        && y >= window.PointsToScreenPixelsY(top) && y <= window.PointsToScreenPixelsY(top + height);

    // The math zone of a text shape (or cell) under a screen point; offset: the origin
    // of the text's bounds on the slide.
    private static (int Start, int Length)? HitMathZone(DocumentWindow window, Shape shape, int x, int y,
        float offsetX = 0, float offsetY = 0)
    {
        object? frame = null, range = null, zones = null;
        try
        {
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
                    if (ContainsScreenPoint(window, offsetX + Convert.ToSingle(((dynamic)zone).BoundLeft), offsetY + Convert.ToSingle(((dynamic)zone).BoundTop),
                            Convert.ToSingle(((dynamic)zone).BoundWidth), Convert.ToSingle(((dynamic)zone).BoundHeight), x, y))
                        return (Convert.ToInt32(((dynamic)zone).Start), Convert.ToInt32(((dynamic)zone).Length));
                }
                finally { Release(zone); }
            }
            return null;
        }
        finally { Release(zones); Release(range); Release(frame); }
    }

    // A double-click on text without any structure or math letter reads no saved
    // copy of the presentation.
    private static string? HitNativeMathZoneInWps(DocumentWindow window, Slide slide, Shape shape, int x, int y)
    {
        if (shape.HasTextFrame != MsoTriState.msoTrue
            || !WpsMathText.MayContainMath(WpsNativeEquations.LiveText(shape))) return null;
        Presentation? presentation = null;
        try
        {
            presentation = (Presentation)slide.Parent;
            var zones = WpsNativeEquations.Locate(presentation, slide.SlideID, shape).Zones;
            var bounds = WpsNativeEquations.Geometry(shape, zones);
            for (var index = 0; index < zones.Count; index++)
            {
                var (left, top, width, height, _) = bounds[index];
                if (x >= window.PointsToScreenPixelsX(left) && x <= window.PointsToScreenPixelsX(left + width)
                    && y >= window.PointsToScreenPixelsY(top) && y <= window.PointsToScreenPixelsY(top + height))
                    return $"{NativeOmmlReference(slide, shape)}:{zones[index].Start}:{zones[index].Length}";
            }
            return null;
        }
        finally { Release(presentation); }
    }

    private static (int Ordinal, int Start, int Length, int Count) ResolveNativeMathZoneInWps(
        Presentation presentation, int slideId, Shape shape, string reference, string shapeXml)
    {
        var parts = reference.Substring(NativeOmmlReferencePrefix.Length).Split(':');
        var selectedStart = parts.Length == 4 ? int.Parse(parts[2]) : 0;
        var selectedLength = parts.Length == 4 ? int.Parse(parts[3]) : 0;
        var zones = WpsNativeEquations.Locate(presentation, slideId, shape, shapeXml).Zones;
        var index = selectedStart == 0
            ? zones.Count == 1 ? 0 : -1
            : WpsMathText.ZoneAt(zones, selectedStart, selectedLength);
        if (index < 0)
            throw new InvalidOperationException(Text(
                "请把光标放入要编辑的一个原生公式中，或只选择该公式。",
                "Place the caret inside the one native equation to edit, or select only that equation."));
        return (index + 1, zones[index].Start, zones[index].Length, zones.Count);
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

    // The text shape holding a native equation: the shape itself, or for an
    // equation in a table cell that cell's text shape.
    private static (Slide? Slide, Shape? Shape) FindNativeOmml(Presentation presentation, int slideId, PowerPointTextHost host)
    {
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            slides = presentation.Slides;
            try { slide = FindSlideById(slides, slideId); } catch { return (null, null); }
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                if (shape.Id == host.ShapeId && !host.IsCell && shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    var result = (slide, shape);
                    slide = null;
                    return result;
                }
                if (shape.Id == host.ShapeId && host.IsCell && shape.HasTable == MsoTriState.msoTrue)
                {
                    try
                    {
                        var cell = PowerPointTextHost.CellShape(shape, host.Row, host.Column);
                        var result = (slide, cell);
                        slide = null;
                        return result;
                    }
                    catch (InvalidOperationException) { return (null, null); }
                    finally { Release(shape); }
                }
                Release(shape);
            }
            return (null, null);
        }
        finally { Release(shapes); Release(slide); Release(slides); }
    }

    private static string ShapeName(Slide slide, int shapeId)
    {
        Shapes? shapes = null;
        try
        {
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Shape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Id == shapeId) return shape.Name;
                }
                finally { Release(shape); }
            }
            throw new InvalidOperationException("The PowerPoint shape holding the equation no longer exists.");
        }
        finally { Release(shapes); }
    }

    public OfficeSelection LoadNativeOmmlSelection(OfficeSelection selection)
    {
        if (!IsNativeOmmlReference(selection.ObjectId)) return selection;
        if (!TryParseNativeOmmlReference(selection.ObjectId, out var slideId, out var host))
            throw new InvalidOperationException("The native equation selection has no exact slide and shape identity.");
        EnsureTextHostSupported(host);
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, selection.DocumentId);
            (slide, shape) = FindNativeOmml(presentation, slideId, host);
            if (shape is null) throw new InvalidOperationException("The selected native equation no longer exists.");
            string shapeXml;
            (int Ordinal, int Start, int Length, int Count) zone;
            if (OfficeHostEnvironment.IsWps)
            {
                shapeXml = PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slideId, host);
                zone = ResolveNativeMathZoneInWps(presentation, slideId, shape, selection.ObjectId!, shapeXml);
            }
            else
            {
                zone = ResolveNativeMathZone(shape, selection.ObjectId!);
                shapeXml = PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slideId, host);
            }
            var omml = PowerPointNativeOmmlPackage.ReadEquation(shapeXml, zone.Ordinal, zone.Count);
            var nativeZone = PowerPointNativeOmmlPackage.ReadEquationZone(shapeXml, zone.Ordinal, zone.Count);
            var mathMl = PowerPointNativeOmmlKernel.ReadNativeMathMl(omml);
            var latex = MathMlToLatexConverter.ConvertWordEquation(mathMl).Trim();
            if (string.IsNullOrWhiteSpace(latex)) throw new InvalidDataException("The selected native equation has no recoverable source.");
            var size = (OfficeHostEnvironment.IsWps
                ? PowerPointOmmlBridge.ReadFontSizeInWps(shape, (zone.Start, zone.Length))
                : PowerPointOmmlBridge.ReadFontSize(shape)) ?? 20f;
            var now = DateTimeOffset.UtcNow.ToString("O");
            var metadata = new FormulaMetadata
            {
                FormulaId = Guid.NewGuid().ToString("D"), Title = "PowerPoint Native Formula",
                Latex = latex,
                Lines = new List<FormulaLine> { new() { Id = Guid.NewGuid().ToString("D"), Latex = latex } },
                CodeFormat = "raw", DisplayMode = "block", Numbered = false,
                FontSizePt = size, RenderFontSizePt = size,
                RenderWidthPx = shape.Width / 0.75, RenderHeightPx = shape.Height / 0.75,
                CreatedWithVersion = "1.3.1", UpdatedWithVersion = "1.3.1",
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
        if (!TryParseNativeOmmlReference(session.SourceObjectId, out var slideId, out var host)
            || slideId != slide.SlideID)
            throw new InvalidOperationException("The native equation selection has no exact slide and shape identity.");
        // Replace exactly the existing math zone. Its paragraph and neighboring
        // content already exist; inserting a new display paragraph would split them.
        var write = host.IsCell
            ? PowerPointCellEquations.PrepareInsertion(_application, shape, source.Start, source.Length, mathMl, "inline")
            : PowerPointNativeOmmlKernel.PrepareTextInsertion(
                _application, shape, source.Start, source.Length, mathMl, "inline", source.Zone);
        PowerPointTextEditTransaction.Execute(_application, presentation, slideId, host,
            source.ShapeFingerprint, write);
        return Result(session, presentation, host.IsCell ? ShapeName(slide, host.ShapeId) : shape.Name);
    }

    private void EnsureNativeOmmlSourceUnchanged(OfficeSessionDocument session, Presentation presentation)
    {
        if (!TryParseNativeOmmlReference(session.SourceObjectId, out var slideId, out var host)) return;
        var source = _nativeOmmlSource;
        if (source is null || source.DocumentId != session.SourceDocumentId
            || source.ObjectId != session.SourceObjectId || source.FormulaId != session.FormulaId)
            throw new InvalidOperationException("Reopen the native equation before saving it.");
        if (host.IsCell)
            throw new InvalidOperationException(TableCellHoldsOnlyTextMessage);
        var current = PowerPointNativeOmmlPackage.ReadStandalone(presentation, slideId, host.ShapeId);
        if (!string.Equals(source.Omml, current, StringComparison.Ordinal))
            throw new InvalidOperationException("The native equation changed while the editor was open. Reopen it before saving.");
    }

    private float? NativeShapeFontSize(OfficeSelection selection, float? requested = null)
    {
        if (!TryParseNativeOmmlReference(selection.ObjectId, out var slideId, out var host)) return null;
        // An equation in a table cell is sized as text, with the caret in it.
        if (host.IsCell && (requested.HasValue || OfficeHostEnvironment.IsWps)) return null;
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, selection.DocumentId);
            (slide, shape) = FindNativeOmml(presentation, slideId, host);
            if (shape is null) return null;
            if (requested.HasValue)
            {
                EnsureWritable(presentation);
                // A selected text range is handled separately; changing a whole
                // object must never apply to adjacent prose or multiple equations.
                _ = PowerPointNativeOmmlPackage.ReadStandalone(presentation, slideId, host.ShapeId);
                StartNewUndoEntry();
                PowerPointOmmlBridge.SetFontSize(shape, requested.Value);
                StartNewUndoEntry();
            }
            return PowerPointOmmlBridge.ReadFontSize(shape);
        }
        finally { Release(shape); Release(slide); Release(presentation); }
    }

}

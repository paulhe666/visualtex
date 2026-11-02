using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// Read operations for agents (docs/office-mcp/DESIGN.md §5, §6) in PowerPoint and WPS
/// Presentation. One saved copy of the presentation (as format conversion takes it)
/// gives native equations and MathType objects without the clipboard or activating
/// anything; VisualTeX objects are read from their tags. Nothing is written.
/// </summary>
public sealed partial class PowerPointFormulaService
{
    private const string AutomationVisualTeXKind = "visualtex_ole";
    private const string AutomationPictureKind = "picture";
    private const string AutomationOmmlKind = "office_equation";
    private const string AutomationMathTypeKind = "mathtype";

    private sealed class AutomationSlideFormula
    {
        internal string Id { get; set; } = string.Empty;
        internal string Kind { get; set; } = string.Empty;
        internal int SlideId { get; set; }
        internal int SlideIndex { get; set; }
        internal int ShapeId { get; set; }
        internal string ShapeName { get; set; } = string.Empty;
        internal int Ordinal { get; set; }
        internal string? FormulaId { get; set; }
        internal string? Latex { get; set; }
        internal IReadOnlyList<string>? Lines { get; set; }
        internal string Display { get; set; } = "block";
        internal double? FontSizePt { get; set; }
        internal string ContextBefore { get; set; } = string.Empty;
        internal string ContextAfter { get; set; } = string.Empty;
        internal string? Error { get; set; }
    }

    internal object AutomationListDocuments()
    {
        var foreground = OfficeAutomationText.ProcessIsForeground();
        var activeIdentity = TryActivePresentationIdentity();
        var result = new List<object>();
        Presentations? presentations = null;
        try
        {
            presentations = _application.Presentations;
            for (var index = 1; index <= presentations.Count; index++)
            {
                Presentation? presentation = null;
                try
                {
                    presentation = presentations[index];
                    if (!HasWindow(presentation)) continue;
                    var identity = DocumentIdentity(presentation);
                    var active = string.Equals(identity, activeIdentity, StringComparison.OrdinalIgnoreCase);
                    string path;
                    try { path = presentation.Path ?? string.Empty; } catch { path = string.Empty; }
                    result.Add(new
                    {
                        key = OfficeAutomationText.DocumentKey(identity),
                        name = SafeRead(() => presentation.Name) ?? identity,
                        path = path.Length == 0 ? null : identity,
                        active,
                        foreground = foreground && active,
                        saved = SafeRead(() => (bool?)(presentation.Saved == MsoTriState.msoTrue)),
                        read_only = SafeRead(() => (bool?)(presentation.ReadOnly == MsoTriState.msoTrue)),
                    });
                }
                finally { Release(presentation); }
            }
        }
        finally { Release(presentations); }
        return new { documents = result };
    }

    internal object AutomationListFormulas(string key, int offset, int limit, int contextChars, HashSet<string> kinds)
    {
        Presentation? presentation = null;
        try
        {
            presentation = FindAutomationPresentation(key);
            var formulas = ScanAutomationFormulas(presentation, kinds, contextChars, slideId: null);
            return new
            {
                revision = OfficeAutomationText.Hash8(string.Join(";", formulas.Select(formula => formula.Id + "=" + formula.Latex))),
                total = formulas.Count,
                offset,
                formulas = formulas.Skip(offset).Take(limit).Select(formula => DescribeAutomationFormula(formula, full: false)).ToList(),
            };
        }
        finally { Release(presentation); }
    }

    internal object AutomationReadFormula(string key, string formulaId)
    {
        Presentation? presentation = null;
        try
        {
            presentation = FindAutomationPresentation(key);
            int? slideId = null;
            var parts = formulaId.Split(':');
            if (parts[0] == "ppt")
            {
                if (parts.Length is not (3 or 4) || !int.TryParse(parts[1], out var parsed))
                    throw new OfficeAutomationException("formula_not_found", "This is not a PowerPoint formula id from list_formulas.");
                slideId = parsed;
            }
            else if (parts[0] != "vtx")
            {
                throw new OfficeAutomationException("formula_not_found", "This is not a PowerPoint formula id from list_formulas.");
            }
            var formula = ScanAutomationFormulas(presentation, new HashSet<string>(), 120, slideId)
                .FirstOrDefault(candidate => candidate.Id == formulaId)
                ?? throw new OfficeAutomationException(
                    parts[0] == "ppt" && parts.Length == 4 ? "stale_formula_id" : "formula_not_found",
                    "No formula in this presentation has that id; call list_formulas again.");
            return DescribeAutomationFormula(formula, full: true);
        }
        finally { Release(presentation); }
    }

    internal object AutomationGetSelection(string? key)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        Selection? selection = null;
        View? view = null;
        Slide? slide = null;
        ShapeRange? range = null;
        Shape? shape = null;
        try
        {
            presentation = FindAutomationPresentation(key);
            var identity = DocumentIdentity(presentation);
            var active = string.Equals(identity, TryActivePresentationIdentity(), StringComparison.OrdinalIgnoreCase);
            window = active ? _application.ActiveWindow : FirstWindow(presentation);
            if (window is null)
                throw new OfficeAutomationException("document_not_found", "The presentation has no window.");
            selection = window.Selection;
            int? slideIndex = null;
            int? slideId = null;
            try
            {
                view = window.View;
                slide = (Slide)view.Slide;
                slideIndex = slide.SlideIndex;
                slideId = slide.SlideID;
            }
            catch { }
            var type = selection.Type;
            string kind = type switch
            {
                PpSelectionType.ppSelectionSlides => "slide",
                PpSelectionType.ppSelectionShapes => "shape",
                PpSelectionType.ppSelectionText => "text",
                _ => "none",
            };
            object? formula = null;
            string? textPreview = null;
            int? shapeId = null;
            if (type is PpSelectionType.ppSelectionShapes or PpSelectionType.ppSelectionText && slideId is not null)
            {
                try { range = selection.ShapeRange; } catch { range = null; }
                if (range?.Count == 1)
                {
                    shape = range[1];
                    shapeId = shape.Id;
                    var onShape = ScanAutomationFormulas(presentation, new HashSet<string>(), 60, slideId)
                        .Where(candidate => candidate.ShapeId == shapeId)
                        .ToList();
                    var chosen = onShape.Count == 1 ? onShape[0]
                        : onShape.Count > 1 && type == PpSelectionType.ppSelectionText
                            ? onShape.FirstOrDefault(candidate => candidate.Ordinal == SelectedMathZoneOrdinal(selection, shape))
                            : null;
                    if (chosen is not null)
                    {
                        kind = "formula";
                        formula = DescribeAutomationFormula(chosen, full: true);
                    }
                    if (type == PpSelectionType.ppSelectionText && chosen is null)
                    {
                        try { textPreview = OfficeAutomationText.Clean(selection.TextRange.Text); }
                        catch { }
                        if (textPreview is { Length: > 200 }) textPreview = textPreview.Substring(0, 200) + "…";
                    }
                }
            }
            return new
            {
                document = OfficeAutomationText.DocumentKey(identity),
                foreground = active && OfficeAutomationText.ProcessIsForeground(),
                kind,
                formula,
                text_preview = textPreview,
                location = new { slide = slideIndex, slide_id = slideId, shape_id = shapeId },
                read_only = presentation.ReadOnly == MsoTriState.msoTrue,
            };
        }
        finally
        {
            Release(shape);
            Release(range);
            Release(slide);
            Release(view);
            Release(selection);
            Release(window);
            Release(presentation);
        }
    }

    // Writes: placing the target, then the add-in's own command does the work.

    /// Brings the presentation's window to the front: the add-in's commands work
    /// on the active window and its selection.
    internal DocumentWindow AutomationActivate(string? key)
    {
        Presentation? presentation = null;
        try
        {
            presentation = FindAutomationPresentation(key);
            if (presentation.ReadOnly == MsoTriState.msoTrue)
                throw new OfficeAutomationException("document_read_only", "The presentation is read-only.");
            var window = FirstWindow(presentation)
                ?? throw new OfficeAutomationException("document_not_found", "The presentation has no window.");
            window.Activate();
            return window;
        }
        finally { Release(presentation); }
    }

    /// Shows a formula's slide and selects the formula (its text for an equation in a text box).
    internal object AutomationSelectFormula(string key, string formulaId)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            window = AutomationActivate(key);
            presentation = FindAutomationPresentation(key);
            int? slideFilter = formulaId.StartsWith("ppt:", StringComparison.Ordinal)
                && formulaId.Split(':') is { Length: >= 3 } parts && int.TryParse(parts[1], out var parsed)
                ? parsed : null;
            var formula = ScanAutomationFormulas(presentation, new HashSet<string>(), 40, slideFilter)
                .FirstOrDefault(candidate => candidate.Id == formulaId)
                ?? throw new OfficeAutomationException("formula_not_found", "No formula in this presentation has that id; call list_formulas again.");
            view = window.View;
            view.GotoSlide(formula.SlideIndex);
            slides = presentation.Slides;
            slide = FindSlideById(slides, formula.SlideId);
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Shape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (shape.Id != formula.ShapeId) continue;
                    if (formula.Ordinal > 0 && !OfficeHostEnvironment.IsWps)
                    {
                        object? frame = null, range = null, zone = null;
                        try
                        {
                            frame = shape.TextFrame2;
                            range = ((dynamic)frame).TextRange;
                            zone = ((dynamic)range).MathZones(formula.Ordinal, 1);
                            ((dynamic)zone).Select();
                        }
                        finally { Release(zone); Release(range); Release(frame); }
                    }
                    else shape.Select();
                    return DescribeAutomationFormula(formula, full: false);
                }
                finally { Release(shape); }
            }
            throw new OfficeAutomationException("formula_not_found", "The formula's shape no longer exists.");
        }
        finally
        {
            Release(shapes);
            Release(slide);
            Release(slides);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    internal sealed class AutomationInsertion
    {
        internal bool AtCursor { get; set; }
        internal int SlideId { get; set; }
        internal HashSet<int> ShapeIds { get; } = new();
        internal float? Left { get; set; }
        internal float? Top { get; set; }
    }

    /// Shows the slide a new formula goes on ({"slide": n, "left_pt": x, "top_pt": y});
    /// "cursor" keeps the user's place (a text caret receives a native equation).
    internal AutomationInsertion AutomationPlaceInsertion(string? key, JsonElement location)
    {
        DocumentWindow? window = null;
        View? view = null;
        Selection? selection = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            window = AutomationActivate(key);
            var insertion = new AutomationInsertion();
            var slideNumber = location.ValueKind == JsonValueKind.Object
                ? OfficeAutomationText.Number(location, "slide")
                : null;
            if (location.ValueKind == JsonValueKind.Object)
            {
                insertion.Left = (float?)OfficeAutomationText.Number(location, "left_pt");
                insertion.Top = (float?)OfficeAutomationText.Number(location, "top_pt");
            }
            view = window.View;
            if (slideNumber is null)
            {
                if (location.ValueKind == JsonValueKind.String && location.GetString() is not (null or "" or "cursor"))
                    throw new OfficeAutomationException("invalid_arguments", "location is \"cursor\" or {\"slide\": n, \"left_pt\": x, \"top_pt\": y}.");
                insertion.AtCursor = true;
            }
            else
            {
                view.GotoSlide((int)slideNumber.Value);
                selection = window.Selection;
                try { selection.Unselect(); } catch { }
            }
            slide = (Slide)view.Slide;
            insertion.SlideId = slide.SlideID;
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Shape? shape = null;
                try
                {
                    shape = shapes[index];
                    insertion.ShapeIds.Add(shape.Id);
                }
                finally { Release(shape); }
            }
            return insertion;
        }
        catch (System.Runtime.InteropServices.COMException error)
        {
            throw new OfficeAutomationException("invalid_arguments", "That slide does not exist: " + error.Message);
        }
        finally
        {
            Release(shapes);
            Release(slide);
            Release(selection);
            Release(view);
            Release(window);
        }
    }

    /// The formula an insertion added (a shape that was not on the slide before),
    /// moved to the requested position.
    internal object? AutomationDescribeInserted(string? key, AutomationInsertion insertion)
    {
        Presentation? presentation = null;
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            presentation = FindAutomationPresentation(key);
            slides = presentation.Slides;
            slide = FindSlideById(slides, insertion.SlideId);
            shapes = slide.Shapes;
            var added = new List<int>();
            for (var index = 1; index <= shapes.Count; index++)
            {
                Shape? shape = null;
                try
                {
                    shape = shapes[index];
                    if (insertion.ShapeIds.Contains(shape.Id)) continue;
                    added.Add(shape.Id);
                    if (insertion.Left is { } left) shape.Left = left;
                    if (insertion.Top is { } top) shape.Top = top;
                }
                finally { Release(shape); }
            }
            var formulas = ScanAutomationFormulas(presentation, new HashSet<string>(), 40, insertion.SlideId)
                .Where(formula => added.Contains(formula.ShapeId))
                .Select(formula => DescribeAutomationFormula(formula, full: true))
                .ToList();
            return formulas.Count == 1 ? formulas[0] : formulas.Count == 0 ? null : formulas;
        }
        catch { return null; }
        finally
        {
            Release(shapes);
            Release(slide);
            Release(slides);
            Release(presentation);
        }
    }

    // Presentations

    private string? TryActivePresentationIdentity()
    {
        Presentation? active = null;
        try
        {
            if (_application.Presentations.Count == 0) return null;
            active = _application.ActivePresentation;
            return DocumentIdentity(active);
        }
        catch { return null; }
        finally { Release(active); }
    }

    private static bool HasWindow(Presentation presentation)
    {
        DocumentWindows? windows = null;
        try
        {
            windows = presentation.Windows;
            return windows.Count > 0;
        }
        catch { return false; }
        finally { Release(windows); }
    }

    private static DocumentWindow? FirstWindow(Presentation presentation)
    {
        DocumentWindows? windows = null;
        try
        {
            windows = presentation.Windows;
            return windows.Count > 0 ? windows[1] : null;
        }
        catch { return null; }
        finally { Release(windows); }
    }

    private Presentation FindAutomationPresentation(string? key)
    {
        if (string.IsNullOrEmpty(key))
        {
            try
            {
                if (_application.Presentations.Count > 0) return _application.ActivePresentation;
            }
            catch { }
            throw new OfficeAutomationException("document_not_found", "No presentation is open.");
        }
        Presentations? presentations = null;
        try
        {
            presentations = _application.Presentations;
            for (var index = 1; index <= presentations.Count; index++)
            {
                var presentation = presentations[index];
                bool match;
                try { match = OfficeAutomationText.DocumentKey(DocumentIdentity(presentation)) == key; }
                catch { match = false; }
                if (match) return presentation;
                Release(presentation);
            }
        }
        finally { Release(presentations); }
        throw OfficeAutomationText.DocumentNotFound();
    }

    // Formulas

    private List<AutomationSlideFormula> ScanAutomationFormulas(
        Presentation presentation,
        HashSet<string> kinds,
        int contextChars,
        int? slideId)
    {
        bool Wanted(string kind) => kinds.Count == 0 || kinds.Contains(kind);
        var result = new List<AutomationSlideFormula>();
        var snapshot = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-agent-{Guid.NewGuid():N}.pptx");
        Slides? slides = null;
        try
        {
            presentation.SaveCopyAs(snapshot, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using var stream = File.OpenRead(snapshot);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            slides = presentation.Slides;
            for (var index = 1; index <= slides.Count; index++)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                try
                {
                    slide = slides[index];
                    var id = slide.SlideID;
                    if (slideId is not null && id != slideId) continue;
                    XDocument slideXml;
                    try { slideXml = PowerPointNativeOmmlPackage.ReadSlide(archive, id); }
                    catch { continue; }
                    var shapesXml = PowerPointSlideXml.IndexTopLevelShapes(slideXml);
                    if (shapesXml.Count == 0) continue;
                    shapes = slide.Shapes;
                    for (var position = 1; position <= shapes.Count; position++)
                    {
                        Shape? shape = null;
                        try
                        {
                            shape = shapes[position];
                            if (!shapesXml.TryGetValue(shape.Id, out var info)) continue;
                            ScanAutomationShape(archive, shape, info, id, index, Wanted, contextChars, result);
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
        }
        finally
        {
            Release(slides);
            try { File.Delete(snapshot); } catch { }
        }
        AssignAutomationIds(result);
        return result;
    }

    private static void ScanAutomationShape(
        ZipArchive archive,
        Shape shape,
        PowerPointSlideShapeInfo info,
        int slideId,
        int slideIndex,
        Func<string, bool> wanted,
        int contextChars,
        List<AutomationSlideFormula> result)
    {
        AutomationSlideFormula New(string kind, int ordinal = 0) => new()
        {
            Kind = kind,
            SlideId = slideId,
            SlideIndex = slideIndex,
            ShapeId = info.ShapeId,
            ShapeName = SafeShapeName(shape),
            Ordinal = ordinal,
        };

        switch (info.Kind)
        {
            case PowerPointSlideShapeKind.Ole
                when string.Equals(info.ProgId, FormulaOleContract.ProgId, StringComparison.OrdinalIgnoreCase):
            case PowerPointSlideShapeKind.Picture:
            {
                var kind = info.Kind == PowerPointSlideShapeKind.Ole ? AutomationVisualTeXKind : AutomationPictureKind;
                if (!wanted(kind)) return;
                var metadata = ReadPictureMetadata(shape);
                if (metadata is null) return;
                var formula = New(kind);
                formula.FormulaId = metadata.FormulaId;
                formula.Latex = metadata.Latex;
                var lines = metadata.Lines?.Select(line => line.Latex).Where(line => line is not null).ToList();
                formula.Lines = lines is { Count: > 0 } ? lines : new List<string> { metadata.Latex };
                formula.Display = string.Equals(metadata.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase) ? "inline" : "block";
                formula.FontSizePt = metadata.FontSizePt > 0 ? metadata.FontSizePt : null;
                result.Add(formula);
                return;
            }
            case PowerPointSlideShapeKind.Ole
                when info.ProgId?.StartsWith("Equation.DSMT", StringComparison.OrdinalIgnoreCase) == true:
            {
                if (!wanted(AutomationMathTypeKind) || info.Linked) return;
                var formula = New(AutomationMathTypeKind);
                try
                {
                    var equation = PowerPointMathTypePackage.Read(archive, slideId, info.ShapeId);
                    var native = MathTypeOleStorage.ReadEquationNative(equation.CompoundFile);
                    var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
                    formula.Latex = MathMlToLatexConverter.Convert(mathMl).Trim();
                    formula.Lines = new List<string> { formula.Latex };
                    formula.FontSizePt = MathTypeMtefCodec.ReadEquationNativeFullFontSize(native);
                }
                catch (Exception error) { formula.Error = OfficeErrorText.Describe(error); }
                result.Add(formula);
                return;
            }
            case PowerPointSlideShapeKind.TextWithMath when info.Math is not null:
            {
                if (!wanted(AutomationOmmlKind)) return;
                var (text, marks) = AutomationShapeText(info.Element);
                for (var ordinal = 0; ordinal < info.Math.Equations.Count; ordinal++)
                {
                    var equation = info.Math.Equations[ordinal];
                    var formula = New(AutomationOmmlKind, ordinal + 1);
                    formula.Display = equation.Display ? "block" : "inline";
                    try
                    {
                        var mathMl = PowerPointNativeOmmlKernel.ReadNativeMathMl(equation.Omml);
                        formula.Latex = MathMlToLatexConverter.ConvertWordEquation(mathMl).Trim();
                        formula.Lines = new List<string> { formula.Latex };
                    }
                    catch (Exception error) { formula.Error = OfficeErrorText.Describe(error); }
                    if (ordinal < marks.Count && contextChars > 0)
                    {
                        var mark = marks[ordinal];
                        var start = Math.Max(0, mark - contextChars);
                        formula.ContextBefore = text.Substring(start, mark - start);
                        formula.ContextAfter = text.Substring(mark, Math.Min(contextChars, text.Length - mark));
                    }
                    result.Add(formula);
                }
                return;
            }
        }
    }

    /// A text shape's visible text outside its equations, and where each equation sits in it.
    private static (string Text, List<int> Marks) AutomationShapeText(XElement shape)
    {
        var builder = new StringBuilder();
        var marks = new List<int>();
        var paragraphs = shape.Element(PowerPointSlideXml.P + "txBody")?.Elements(PowerPointSlideXml.A + "p").ToArray()
            ?? Array.Empty<XElement>();
        for (var index = 0; index < paragraphs.Length; index++)
        {
            if (index > 0) builder.Append('\n');
            foreach (var element in PowerPointSlideXml.LiveDescendants(paragraphs[index]))
            {
                if (element.Name == PowerPointSlideXml.M + "oMath")
                {
                    if (!element.Ancestors(PowerPointSlideXml.M + "oMath").Any()) marks.Add(builder.Length);
                }
                else if (element.Name == PowerPointSlideXml.A + "t" && !element.Ancestors(PowerPointSlideXml.M + "oMath").Any())
                {
                    builder.Append(element.Value);
                }
                else if (element.Name == PowerPointSlideXml.A + "br")
                {
                    builder.Append('\n');
                }
            }
        }
        return (builder.ToString(), marks);
    }

    /// VisualTeX objects keep their own formula id (`vtx:`); the others are named by slide and shape.
    private static void AssignAutomationIds(List<AutomationSlideFormula> formulas)
    {
        var unique = formulas
            .Where(formula => !string.IsNullOrEmpty(formula.FormulaId))
            .GroupBy(formula => formula.FormulaId!, StringComparer.OrdinalIgnoreCase)
            .Where(group => group.Count() == 1)
            .Select(group => group.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var formula in formulas)
        {
            formula.Id = formula.FormulaId is { } id && unique.Contains(id)
                ? "vtx:" + id
                : formula.Ordinal > 0
                    ? $"ppt:{formula.SlideId}:{formula.ShapeId}:{formula.Ordinal}"
                    : $"ppt:{formula.SlideId}:{formula.ShapeId}";
        }
    }

    private static object DescribeAutomationFormula(AutomationSlideFormula formula, bool full) => new
    {
        formula_id = formula.Id,
        kind = formula.Kind,
        latex = formula.Latex,
        lines = full ? formula.Lines : null,
        display = formula.Display,
        numbered = false,
        font_size_pt = full ? formula.FontSizePt : null,
        location = new
        {
            slide = formula.SlideIndex,
            slide_id = formula.SlideId,
            shape_id = formula.ShapeId,
            shape_name = formula.ShapeName,
            equation = formula.Ordinal > 0 ? formula.Ordinal : (int?)null,
        },
        context_before = formula.ContextBefore,
        context_after = formula.ContextAfter,
        error = formula.Error,
    };

    /// The 1-based equation of a text shape that holds the text caret (PowerPoint only).
    private static int SelectedMathZoneOrdinal(Selection selection, Shape shape)
    {
        if (OfficeHostEnvironment.IsWps) return 0;
        object? text = null, frame = null, range = null, zones = null;
        try
        {
            text = ((dynamic)selection).TextRange2;
            var caret = Convert.ToInt32(((dynamic)text).Start);
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
                    int start = Convert.ToInt32(((dynamic)zone).Start), length = Convert.ToInt32(((dynamic)zone).Length);
                    if (caret >= start && caret <= start + length) return index;
                }
                finally { Release(zone); }
            }
            return 0;
        }
        catch { return 0; }
        finally { Release(zones); Release(range); Release(frame); Release(text); }
    }

    private static T? SafeRead<T>(Func<T> read)
    {
        try { return read(); }
        catch { return default; }
    }
}

using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using GroupShapes = Microsoft.Office.Interop.PowerPoint.GroupShapes;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;

namespace VisualTeX.PowerPointVsto;

// A whole batch: what was found in scope, what will be converted and what was
// left alone (with the reason), plus the view to restore afterwards.
internal sealed class PowerPointFormulaConversionPlan
{
    internal PowerPointFormulaConversionPlan(PowerPointConversionRequest request) => Request = request;

    internal PowerPointConversionRequest Request { get; }
    internal string DocumentId { get; set; } = string.Empty;
    internal List<PowerPointConversionUnit> Units { get; } = new();
    internal List<PowerPointConversionSkip> Skips { get; } = new();
    internal int ScopeSlideCount { get; set; }
    internal int OriginalSlideId { get; set; }
    internal PpViewType OriginalViewType { get; set; } = PpViewType.ppViewNormal;
    internal int ItemCount => Units.Sum(unit => unit.Items.Count);
}

// One source shape. It normally holds one formula; an equation-only text box
// with several native equations becomes several objects.
internal sealed class PowerPointConversionUnit
{
    internal int SlideId { get; set; }
    internal int SlideIndex { get; set; }
    internal int ShapeId { get; set; }
    internal string ShapeName { get; set; } = string.Empty;
    internal PowerPointFormulaKind SourceKind { get; set; }
    // Source proofs, compared again right before writing.
    internal string? TextFingerprint { get; set; }
    internal byte[]? EquationNative { get; set; }
    internal string? MetadataFingerprint { get; set; }
    internal float Rotation { get; set; }
    internal bool FlipHorizontal { get; set; }
    internal bool FlipVertical { get; set; }
    internal bool Visible { get; set; } = true;
    internal List<PowerPointConversionItem> Items { get; } = new();
}

internal sealed class PowerPointConversionItem
{
    internal string Id { get; } = Guid.NewGuid().ToString("N");
    internal string Latex { get; set; } = string.Empty;
    internal List<FormulaLine>? Lines { get; set; }
    internal string CodeFormat { get; set; } = "raw";
    internal string? SourceMathMl { get; set; }
    internal string DisplayMode { get; set; } = "block";
    // The size the audience sees, including any PowerPoint scaling.
    internal double EffectiveFontSizePt { get; set; }
    internal FormulaMetadata? SourceMetadata { get; set; }
    // The color the source shows for its uncolored parts (#RRGGBB): the inherited
    // text color for a native equation, black for MathType and VisualTeX.
    internal string? BaseColor { get; set; }
    // Visual box of the formula itself (for native equations, the math zone,
    // not the possibly much larger text box), in slide points.
    internal float CenterX { get; set; }
    internal float CenterY { get; set; }
    internal float VisualWidth { get; set; }
    internal float VisualHeight { get; set; }
}

// Everything a target writer needs, prepared off the Office thread.
internal sealed class PreparedPowerPointConversion
{
    internal string ItemId { get; set; } = string.Empty;
    internal float RenderFontSizePt { get; set; }
    internal float PlacementScale { get; set; } = 1f;
    internal string? Failure { get; set; }
    internal OfficeSessionDocument? Session { get; set; }
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? EmfPath { get; set; }
    internal string? SvgPath { get; set; }
    internal float WidthPt { get; set; }
    internal float HeightPt { get; set; }
    internal byte[]? MathTypeCompoundFile { get; set; }
    internal byte[]? MathTypeEquationNative { get; set; }
    internal byte[]? MathTypePreviewEmf { get; set; }
    internal float MathTypeNaturalWidthPt { get; set; }
    internal float MathTypeNaturalHeightPt { get; set; }
}

internal sealed class PowerPointConversionUnitResult
{
    internal bool Converted { get; set; }
    internal int FormulaCount { get; set; }
    internal PowerPointConversionSkipReason? Reason { get; set; }
    internal string? Detail { get; set; }
    internal int SlideId { get; set; }
    internal List<int> ReplacementShapeIds { get; } = new();
}

// The hidden presentation that carries every MathType equation of a batch. It
// is loaded once; each equation is copied from it onto its target slide.
internal sealed class PowerPointMathTypeStaging
{
    internal Presentation? Presentation { get; set; }
    internal string Root { get; set; } = string.Empty;
    internal Dictionary<string, string> FrameNames { get; } = new(StringComparer.Ordinal);
}

public sealed partial class PowerPointFormulaService
{
    private const string BatchStagingPlaceholderPrefix = "VisualTeX.MathType.Batch.";
    private const string BatchStagingFramePrefix = "VisualTeX.MathType.BatchFrame.";

    // ---- Capture ------------------------------------------------------------

    internal PowerPointFormulaConversionPlan CaptureFormulaConversionPlan(PowerPointConversionRequest request)
    {
        var plan = new PowerPointFormulaConversionPlan(request);
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? current = null;
        string? snapshot = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException(Text("没有打开的演示文稿。", "No PowerPoint presentation is open."));
            EnsureWritable(presentation);
            plan.DocumentId = DocumentIdentity(presentation);
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException(Text("没有活动的 PowerPoint 窗口。", "No active PowerPoint window."));
            try { plan.OriginalViewType = window.ViewType; } catch { }
            view = window.View;
            try
            {
                current = (Slide)view.Slide;
                plan.OriginalSlideId = current.SlideID;
            }
            catch { plan.OriginalSlideId = 0; }

            var scope = ResolveConversionScope(presentation, window, plan);
            plan.ScopeSlideCount = scope.Count;
            if (scope.Count == 0) return plan;

            snapshot = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-convert-{Guid.NewGuid():N}.pptx");
            presentation.SaveCopyAs(snapshot, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using var stream = File.OpenRead(snapshot);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            foreach (var entry in scope)
                CaptureConversionSlide(presentation, archive, plan, entry.SlideId, entry.SlideIndex, entry.ShapeIds);
            return plan;
        }
        finally
        {
            if (snapshot is not null) { try { File.Delete(snapshot); } catch { } }
            Release(current);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    private List<(int SlideId, int SlideIndex, HashSet<int>? ShapeIds)> ResolveConversionScope(
        Presentation presentation,
        DocumentWindow window,
        PowerPointFormulaConversionPlan plan)
    {
        var result = new List<(int SlideId, int SlideIndex, HashSet<int>? ShapeIds)>();
        Slides? slides = null;
        Selection? selection = null;
        SlideRange? slideRange = null;
        ShapeRange? shapeRange = null;
        try
        {
            slides = presentation.Slides;
            switch (plan.Request.Scope)
            {
                case PowerPointConversionScope.Presentation:
                    for (var index = 1; index <= slides.Count; index++)
                    {
                        Slide? slide = null;
                        try
                        {
                            slide = slides[index];
                            result.Add((slide.SlideID, slide.SlideIndex, null));
                        }
                        finally { Release(slide); }
                    }
                    return result;

                case PowerPointConversionScope.Slide:
                    if (plan.OriginalSlideId != 0)
                    {
                        result.Add((plan.OriginalSlideId, SlideIndexOf(slides, plan.OriginalSlideId), null));
                        return result;
                    }
                    selection = window.Selection;
                    if (selection.Type == PpSelectionType.ppSelectionSlides)
                    {
                        slideRange = selection.SlideRange;
                        if (slideRange.Count >= 1)
                        {
                            var first = slideRange[1];
                            try { result.Add((first.SlideID, first.SlideIndex, null)); }
                            finally { Release(first); }
                            return result;
                        }
                    }
                    throw new InvalidOperationException(Text(
                        "请先在普通视图中打开要转换的幻灯片。",
                        "Open the slide to convert in Normal view first."));

                default:
                    selection = window.Selection;
                    if (selection.Type == PpSelectionType.ppSelectionSlides)
                    {
                        slideRange = selection.SlideRange;
                        for (var index = 1; index <= slideRange.Count; index++)
                        {
                            Slide? slide = null;
                            try
                            {
                                slide = slideRange[index];
                                result.Add((slide.SlideID, slide.SlideIndex, null));
                            }
                            finally { Release(slide); }
                        }
                        result.Sort((left, right) => left.SlideIndex.CompareTo(right.SlideIndex));
                        return result;
                    }
                    if (selection.Type is PpSelectionType.ppSelectionShapes or PpSelectionType.ppSelectionText
                        && plan.OriginalSlideId != 0)
                    {
                        shapeRange = selection.ShapeRange;
                        var ids = new HashSet<int>();
                        for (var index = 1; index <= shapeRange.Count; index++)
                        {
                            Shape? shape = null;
                            Shape? parent = null;
                            try
                            {
                                shape = shapeRange[index];
                                // A shape picked inside a group is converted with
                                // its group's rules (reported, not split out).
                                try { parent = shape.ParentGroup; } catch { parent = null; }
                                ids.Add(parent?.Id ?? shape.Id);
                            }
                            finally { Release(parent); Release(shape); }
                        }
                        result.Add((plan.OriginalSlideId, SlideIndexOf(slides, plan.OriginalSlideId), ids));
                        return result;
                    }
                    throw new InvalidOperationException(Text(
                        "请先选择要转换的公式，或在左侧缩略图中选择幻灯片。",
                        "Select the equations to convert, or select slides in the thumbnail pane."));
            }
        }
        finally
        {
            Release(shapeRange);
            Release(slideRange);
            Release(selection);
            Release(slides);
        }
    }

    private static int SlideIndexOf(Slides slides, int slideId)
    {
        Slide? slide = null;
        try
        {
            slide = slides.FindBySlideID(slideId);
            return slide.SlideIndex;
        }
        catch { return 0; }
        finally { Release(slide); }
    }

    private void CaptureConversionSlide(
        Presentation presentation,
        ZipArchive archive,
        PowerPointFormulaConversionPlan plan,
        int slideId,
        int slideIndex,
        HashSet<int>? shapeIds)
    {
        var request = plan.Request;
        XDocument slideXml;
        try { slideXml = PowerPointNativeOmmlPackage.ReadSlide(archive, slideId); }
        catch (Exception error)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, 1,
                Text("无法读取该幻灯片的已保存内容：", "The slide's saved content could not be read: ") + Describe(error)));
            return;
        }
        var index = PowerPointSlideXml.IndexTopLevelShapes(slideXml);
        if (index.Count == 0) return;

        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        HashSet<int>? animated = null;
        try
        {
            slides = presentation.Slides;
            slide = slides.FindBySlideID(slideId);
            shapes = slide.Shapes;
            for (var position = 1; position <= shapes.Count; position++)
            {
                Shape? shape = null;
                try
                {
                    shape = shapes[position];
                    var id = shape.Id;
                    if (shapeIds is not null && !shapeIds.Contains(id)) continue;
                    if (!index.TryGetValue(id, out var info)) continue;
                    switch (info.Kind)
                    {
                        case PowerPointSlideShapeKind.Group:
                        {
                            var nested = CountGroupedFormulas(shape, request);
                            if (nested > 0)
                                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.InGroup, slideIndex, nested));
                            break;
                        }
                        case PowerPointSlideShapeKind.TableWithMath:
                            if (request.Accepts(PowerPointFormulaKind.Omml))
                                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.InTable, slideIndex, info.NestedFormulaHints));
                            break;
                        case PowerPointSlideShapeKind.Ole:
                            if (string.Equals(info.ProgId, FormulaOleContract.ProgId, StringComparison.OrdinalIgnoreCase))
                            {
                                if (request.Accepts(PowerPointFormulaKind.VisualTeXOle))
                                    CaptureVisualTeXUnit(plan, slide, shape, slideIndex, PowerPointFormulaKind.VisualTeXOle);
                            }
                            else if (request.Accepts(PowerPointFormulaKind.MathType)
                                     && info.ProgId?.StartsWith("Equation.DSMT", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                // A linked object is a file reference, not an embedded
                                // equation PowerPoint could hand over for conversion.
                                if (info.Linked)
                                    plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Linked, slideIndex));
                                else if (IsMathTypeShape(shape))
                                    CaptureMathTypeUnit(plan, archive, slide, shape, slideIndex);
                            }
                            break;
                        case PowerPointSlideShapeKind.Picture:
                            if (request.Accepts(PowerPointFormulaKind.VisualTeXPicture) && ReadPictureMetadata(shape) is not null)
                                CaptureVisualTeXUnit(plan, slide, shape, slideIndex, PowerPointFormulaKind.VisualTeXPicture);
                            break;
                        case PowerPointSlideShapeKind.TextWithMath:
                            if (request.Accepts(PowerPointFormulaKind.Omml) && info.Math is not null)
                            {
                                animated ??= CollectAnimatedShapeIds(slide);
                                CaptureOmmlUnit(plan, slideXml, slide, shape, slideIndex, info.Math, animated);
                            }
                            break;
                    }
                }
                finally { Release(shape); }
            }
        }
        finally
        {
            Release(shapes);
            Release(slide);
            Release(slides);
        }
    }

    private PowerPointConversionUnit NewConversionUnit(
        Slide slide,
        Shape shape,
        int slideIndex,
        PowerPointFormulaKind kind)
    {
        var unit = new PowerPointConversionUnit
        {
            SlideId = slide.SlideID,
            SlideIndex = slideIndex,
            ShapeId = shape.Id,
            ShapeName = SafeShapeName(shape),
            SourceKind = kind,
        };
        try { unit.Rotation = shape.Rotation; } catch { }
        try { unit.FlipHorizontal = shape.HorizontalFlip == MsoTriState.msoTrue; } catch { }
        try { unit.FlipVertical = shape.VerticalFlip == MsoTriState.msoTrue; } catch { }
        try { unit.Visible = shape.Visible != MsoTriState.msoFalse; } catch { }
        return unit;
    }

    // Rotation/flip that the target cannot represent is reported, never dropped.
    private static bool TargetKeepsTransform(PowerPointFormulaConversionPlan plan, PowerPointConversionUnit unit)
    {
        var rotated = Math.Abs(unit.Rotation) > 0.01f;
        var flipped = unit.FlipHorizontal || unit.FlipVertical;
        if (!rotated && !flipped) return true;
        var target = plan.Request.TargetMode;
        if (rotated && !PowerPointConversionRules.TargetSupportsRotation(target)) return false;
        if (flipped && !PowerPointConversionRules.TargetSupportsFlip(target)) return false;
        return true;
    }

    private void CaptureVisualTeXUnit(
        PowerPointFormulaConversionPlan plan,
        Slide slide,
        Shape shape,
        int slideIndex,
        PowerPointFormulaKind kind)
    {
        var encoded = ReadEncodedMetadata(shape);
        var metadata = FormulaMetadataCodec.Decode(encoded);
        if (metadata is null || string.IsNullOrWhiteSpace(metadata.Latex))
        {
            plan.Skips.Add(new PowerPointConversionSkip(
                PowerPointConversionSkipReason.Unreadable, slideIndex, 1,
                Text("VisualTeX 公式缺少可读取的源码。", "The VisualTeX formula has no readable source.")));
            return;
        }
        var unit = NewConversionUnit(slide, shape, slideIndex, kind);
        unit.MetadataFingerprint = encoded;
        if (!TargetKeepsTransform(plan, unit))
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.RotatedOrFlipped, slideIndex));
            return;
        }
        var width = shape.Width;
        var height = shape.Height;
        unit.Items.Add(new PowerPointConversionItem
        {
            Latex = metadata.Latex,
            Lines = metadata.Lines?.Count > 0 ? metadata.Lines : null,
            CodeFormat = string.IsNullOrWhiteSpace(metadata.CodeFormat) ? "raw" : metadata.CodeFormat,
            DisplayMode = string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal) ? "inline" : "block",
            EffectiveFontSizePt = InferPowerPointFormulaFontSize(width, height, metadata),
            SourceMetadata = metadata,
            BaseColor = PowerPointConversionColors.Black,
            CenterX = shape.Left + width / 2f,
            CenterY = shape.Top + height / 2f,
            VisualWidth = width,
            VisualHeight = height,
        });
        plan.Units.Add(unit);
    }

    private void CaptureMathTypeUnit(
        PowerPointFormulaConversionPlan plan,
        ZipArchive archive,
        Slide slide,
        Shape shape,
        int slideIndex)
    {
        try
        {
            var unit = NewConversionUnit(slide, shape, slideIndex, PowerPointFormulaKind.MathType);
            if (!TargetKeepsTransform(plan, unit))
            {
                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.RotatedOrFlipped, slideIndex));
                return;
            }
            var equation = PowerPointMathTypePackage.Read(archive, slide.SlideID, shape.Id);
            var native = MathTypeOleStorage.ReadEquationNative(equation.CompoundFile);
            var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
            var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
            if (string.IsNullOrWhiteSpace(latex))
                throw new InvalidDataException(Text("MathType 公式为空或无法识别。", "The MathType equation is empty or unsupported."));
            var width = shape.Width;
            var height = shape.Height;
            var nativeSize = MathTypeMtefCodec.ReadEquationNativeFullFontSize(native);
            unit.EquationNative = native;
            unit.Items.Add(new PowerPointConversionItem
            {
                Latex = latex,
                SourceMathMl = mathMl,
                DisplayMode = "block",
                BaseColor = PowerPointConversionColors.Black,
                EffectiveFontSizePt = PowerPointConversionRules.ScaledFontSize(
                    nativeSize, width, height, equation.WidthPt, equation.HeightPt),
                CenterX = shape.Left + width / 2f,
                CenterY = shape.Top + height / 2f,
                VisualWidth = width,
                VisualHeight = height,
            });
            plan.Units.Add(unit);
        }
        catch (Exception error)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, 1,
                Text("读取 MathType 公式失败：", "Reading the MathType equation failed: ") + Describe(error)));
        }
    }

    private void CaptureOmmlUnit(
        PowerPointFormulaConversionPlan plan,
        XDocument slideXml,
        Slide slide,
        Shape shape,
        int slideIndex,
        PowerPointOmmlShapeAnalysis analysis,
        HashSet<int> animated)
    {
        var count = analysis.Equations.Count;
        if (count == 0) return;
        if (!analysis.EquationOnly)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.OmmlMixedWithText, slideIndex, count));
            return;
        }
        if (count > 1 && animated.Contains(shape.Id))
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.OmmlSplitWithAnimation, slideIndex, count));
            return;
        }
        try
        {
            var unit = NewConversionUnit(slide, shape, slideIndex, PowerPointFormulaKind.Omml);
            var transformed = Math.Abs(unit.Rotation) > 0.01f || unit.FlipHorizontal || unit.FlipVertical;
            if (!TargetKeepsTransform(plan, unit) || transformed && count > 1)
            {
                plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.RotatedOrFlipped, slideIndex, count));
                return;
            }
            var live = PowerPointNativeOmmlPackage.SelectTextShape(slideXml, shape.Id);
            unit.TextFingerprint = PowerPointTextEditTransaction.Fingerprint(live.ToString(SaveOptions.DisableFormatting));
            var zones = ReadMathZoneGeometry(shape);
            if (zones.Count != count)
                throw new InvalidDataException(Text(
                    "文本框中的公式数量与已保存内容不一致。",
                    "The text box's equations do not match its saved content."));
            var fallbackSize = PowerPointOmmlBridge.ReadFontSize(shape) ?? 20f;
            var zoneColors = ReadMathZoneRunColors(shape);
            for (var ordinal = 0; ordinal < count; ordinal++)
            {
                var equation = analysis.Equations[ordinal];
                var mathMl = PowerPointNativeOmmlKernel.ReadNativeMathMl(equation.Omml);
                // Uncolored runs show the inherited text color (white on a dark
                // theme); MathType and VisualTeX would draw them black.
                var baseColor = ordinal < zoneColors.Count
                    ? PowerPointConversionColors.InheritedColor(zoneColors[ordinal], PowerPointConversionColors.ExplicitOmmlColors(equation.Omml))
                    : null;
                if (plan.Request.TargetMode != FormulaOleContract.WordOmmlMode
                    && PowerPointConversionColors.NeedsExplicit(baseColor, PowerPointConversionColors.Black))
                    mathMl = PowerPointConversionColors.ApplyBaseColor(mathMl, baseColor);
                var latex = MathMlToLatexConverter.ConvertWordEquation(mathMl).Trim();
                if (string.IsNullOrWhiteSpace(latex))
                    throw new InvalidDataException(Text("原生公式为空或无法识别。", "The native equation is empty or unsupported."));
                var zone = zones[ordinal];
                // A rotated single-equation box keeps its own center: the zone's
                // bounds are reported in the unrotated text frame.
                var useShapeBox = transformed || zone.Width <= 0 || zone.Height <= 0;
                unit.Items.Add(new PowerPointConversionItem
                {
                    Latex = latex,
                    SourceMathMl = mathMl,
                    DisplayMode = equation.Display ? "block" : "inline",
                    BaseColor = baseColor,
                    EffectiveFontSizePt = zone.FontSize > 0 ? zone.FontSize : fallbackSize,
                    CenterX = useShapeBox ? shape.Left + shape.Width / 2f : zone.Left + zone.Width / 2f,
                    CenterY = useShapeBox ? shape.Top + shape.Height / 2f : zone.Top + zone.Height / 2f,
                    VisualWidth = useShapeBox ? shape.Width : zone.Width,
                    VisualHeight = useShapeBox ? shape.Height : zone.Height,
                });
            }
            plan.Units.Add(unit);
        }
        catch (Exception error)
        {
            plan.Skips.Add(new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, slideIndex, count,
                Text("读取原生公式失败：", "Reading the native equation failed: ") + Describe(error)));
        }
    }

    private static List<(float Left, float Top, float Width, float Height, float FontSize)> ReadMathZoneGeometry(Shape shape)
    {
        var result = new List<(float Left, float Top, float Width, float Height, float FontSize)>();
        object? frame = null, range = null, zones = null;
        try
        {
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            try { zones = ((dynamic)range).MathZones(); }
            catch { zones = ((dynamic)range).MathZones(Type.Missing, Type.Missing); }
            var count = Convert.ToInt32(((dynamic)zones).Count);
            for (var index = 1; index <= count; index++)
            {
                object? zone = null, font = null;
                try
                {
                    zone = ((dynamic)range).MathZones(index, 1);
                    font = ((dynamic)zone).Font;
                    float size;
                    try { size = Convert.ToSingle(((dynamic)font).Size); }
                    catch { size = 0; }
                    // A zone with mixed run sizes reports an undefined size.
                    if (!(size > 0) || size > 4000) size = 0;
                    result.Add((
                        Convert.ToSingle(((dynamic)zone).BoundLeft),
                        Convert.ToSingle(((dynamic)zone).BoundTop),
                        Convert.ToSingle(((dynamic)zone).BoundWidth),
                        Convert.ToSingle(((dynamic)zone).BoundHeight),
                        size));
                }
                finally { Release(font); Release(zone); }
            }
            return result;
        }
        finally { Release(zones); Release(range); Release(frame); }
    }

    // Per math zone, the color and length of each run as PowerPoint draws it
    // (theme colors resolved). Unreadable runs are left out.
    private static List<List<(string Color, int Length)>> ReadMathZoneRunColors(Shape shape)
    {
        var result = new List<List<(string Color, int Length)>>();
        object? frame = null, range = null, zones = null;
        try
        {
            frame = shape.TextFrame2;
            range = ((dynamic)frame).TextRange;
            try { zones = ((dynamic)range).MathZones(); }
            catch { zones = ((dynamic)range).MathZones(Type.Missing, Type.Missing); }
            var count = Convert.ToInt32(((dynamic)zones).Count);
            for (var index = 1; index <= count; index++)
            {
                var colors = new List<(string Color, int Length)>();
                object? zone = null, runs = null;
                try
                {
                    zone = ((dynamic)range).MathZones(index, 1);
                    runs = ((dynamic)zone).Runs();
                    var runCount = Convert.ToInt32(((dynamic)runs).Count);
                    for (var position = 1; position <= runCount; position++)
                    {
                        object? run = null, font = null, fill = null, fore = null;
                        try
                        {
                            run = ((dynamic)zone).Runs(position, 1);
                            var length = Convert.ToInt32(((dynamic)run).Length);
                            font = ((dynamic)run).Font;
                            fill = ((dynamic)font).Fill;
                            fore = ((dynamic)fill).ForeColor;
                            colors.Add((PowerPointConversionColors.FromOfficeRgb(Convert.ToInt32(((dynamic)fore).RGB)), length));
                        }
                        catch { }
                        finally { Release(fore); Release(fill); Release(font); Release(run); }
                    }
                }
                catch { }
                finally { Release(runs); Release(zone); }
                result.Add(colors);
            }
        }
        catch { }
        finally { Release(zones); Release(range); Release(frame); }
        return result;
    }

    private static HashSet<int> CollectAnimatedShapeIds(Slide slide)
    {
        var ids = new HashSet<int>();
        TimeLine? timeline = null;
        Sequences? interactive = null;
        void Collect(Sequence? sequence)
        {
            if (sequence is null) return;
            try
            {
                for (var index = 1; index <= sequence.Count; index++)
                {
                    Effect? effect = null;
                    Shape? animatedShape = null;
                    try
                    {
                        effect = sequence[index];
                        animatedShape = effect.Shape;
                        ids.Add(animatedShape.Id);
                    }
                    catch { }
                    finally { Release(animatedShape); Release(effect); }
                }
            }
            finally { Release(sequence); }
        }
        try
        {
            timeline = slide.TimeLine;
            Collect(timeline.MainSequence);
            interactive = timeline.InteractiveSequences;
            for (var index = 1; index <= interactive.Count; index++) Collect(interactive[index]);
        }
        catch { }
        finally { Release(interactive); Release(timeline); }
        return ids;
    }

    private int CountGroupedFormulas(Shape group, PowerPointConversionRequest request)
    {
        GroupShapes? items = null;
        var count = 0;
        try
        {
            items = group.GroupItems;
            for (var index = 1; index <= items.Count; index++)
            {
                Shape? item = null;
                try
                {
                    item = items[index];
                    if (IsMathTypeShape(item))
                    {
                        if (request.Accepts(PowerPointFormulaKind.MathType)) count++;
                    }
                    else if (ReadPictureMetadata(item) is not null)
                    {
                        if (request.Accepts(IsNativeOle(item) ? PowerPointFormulaKind.VisualTeXOle : PowerPointFormulaKind.VisualTeXPicture))
                            count++;
                    }
                    else if (request.Accepts(PowerPointFormulaKind.Omml) && item.HasTextFrame == MsoTriState.msoTrue)
                    {
                        count += ReadMathZoneGeometry(item).Count;
                    }
                }
                catch { }
                finally { Release(item); }
            }
        }
        catch { }
        finally { Release(items); }
        return count;
    }

    private static string? ReadEncodedMetadata(Shape shape)
    {
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            string? encoded = null;
            try { encoded = tags[MetadataTag]; } catch { }
            if (FormulaMetadataCodec.Decode(encoded) is not null) return encoded;
            try { encoded = shape.AlternativeText; } catch { encoded = null; }
            return FormulaMetadataCodec.Decode(encoded) is not null ? encoded : null;
        }
        finally { Release(tags); }
    }

    private static string SafeShapeName(Shape shape)
    {
        try { return shape.Name; } catch { return string.Empty; }
    }

    // ---- Verification right before writing ----------------------------------

    // Rendering can take a while. Re-read MathType and native-equation sources
    // from one fresh snapshot; anything that changed is left alone and reported.
    internal void VerifyFormulaConversionSources(PowerPointFormulaConversionPlan plan)
    {
        if (!plan.Units.Any(unit => unit.SourceKind is PowerPointFormulaKind.MathType or PowerPointFormulaKind.Omml))
            return;
        Presentation? presentation = null;
        string? snapshot = null;
        try
        {
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, plan.DocumentId);
            snapshot = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-verify-{Guid.NewGuid():N}.pptx");
            presentation.SaveCopyAs(snapshot, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using var stream = File.OpenRead(snapshot);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            // A slide that cannot be read keeps its exception, so every formula
            // on it reports why instead of looking "changed".
            var slides = new Dictionary<int, (XDocument? Xml, Exception? Error)>();
            (XDocument? Xml, Exception? Error) SlideXml(int slideId)
            {
                if (slides.TryGetValue(slideId, out var cached)) return cached;
                (XDocument? Xml, Exception? Error) loaded;
                try { loaded = (PowerPointNativeOmmlPackage.ReadSlide(archive, slideId), null); }
                catch (Exception error) { loaded = (null, error); }
                slides[slideId] = loaded;
                return loaded;
            }
            foreach (var unit in plan.Units.ToArray())
            {
                var unchanged = true;
                string? readFailure = null;
                try
                {
                    switch (unit.SourceKind)
                    {
                        case PowerPointFormulaKind.MathType:
                            unchanged = MathTypeMtefCodec.EquationNativeContentEquals(
                                MathTypeOleStorage.ReadEquationNative(
                                    PowerPointMathTypePackage.Read(archive, unit.SlideId, unit.ShapeId).CompoundFile),
                                unit.EquationNative!);
                            break;
                        case PowerPointFormulaKind.Omml:
                        {
                            var slideXml = SlideXml(unit.SlideId);
                            if (slideXml.Xml is null)
                            {
                                unchanged = false;
                                readFailure = Text("无法重新读取该幻灯片的已保存内容：", "The slide's saved content could not be read again: ")
                                    + Describe(slideXml.Error);
                                break;
                            }
                            unchanged = string.Equals(
                                PowerPointTextEditTransaction.Fingerprint(
                                    PowerPointNativeOmmlPackage.SelectTextShape(slideXml.Xml, unit.ShapeId).ToString(SaveOptions.DisableFormatting)),
                                unit.TextFingerprint,
                                StringComparison.Ordinal);
                            break;
                        }
                    }
                }
                catch (Exception error)
                {
                    unchanged = false;
                    readFailure = Text($"转换前重新读取{ShapeLabel(unit)}失败：", $"Re-reading {ShapeLabel(unit)} before writing failed: ")
                        + Describe(error);
                }
                if (unchanged) continue;
                plan.Units.Remove(unit);
                plan.Skips.Add(readFailure is not null
                    ? new PowerPointConversionSkip(PowerPointConversionSkipReason.Unreadable, unit.SlideIndex, unit.Items.Count,
                        readFailure, UnitLatex(unit))
                    : new PowerPointConversionSkip(PowerPointConversionSkipReason.Changed, unit.SlideIndex, unit.Items.Count,
                        unit.SourceKind == PowerPointFormulaKind.MathType
                            ? Text($"扫描后{ShapeLabel(unit)}中的 MathType 公式内容已改变。",
                                $"The MathType equation in {ShapeLabel(unit)} changed after it was scanned.")
                            : Text($"扫描后{ShapeLabel(unit)}的文本内容或格式已改变。",
                                $"The text or formatting of {ShapeLabel(unit)} changed after it was scanned."),
                        UnitLatex(unit)));
            }
        }
        finally
        {
            if (snapshot is not null) { try { File.Delete(snapshot); } catch { } }
            Release(presentation);
        }
    }

    // ---- Writing: VisualTeX OLE, SVG picture and native equation targets -----

    internal PowerPointConversionUnitResult ApplyConversionUnit(
        PowerPointFormulaConversionPlan plan,
        PowerPointConversionUnit unit,
        IReadOnlyDictionary<string, PreparedPowerPointConversion> prepared)
    {
        var result = new PowerPointConversionUnitResult { SlideId = unit.SlideId, FormulaCount = unit.Items.Count };
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? source = null;
        var created = new List<Shape>();
        var target = plan.Request.TargetMode;
        (IntPtr Hwnd, long Generation)? redrawFreeze = null;
        var deferRedraw = false;
        var sourceRetired = false;
        var oleRestores = new List<(string FormulaId, int ShapeId, float Width, float Height, float Left, float Top)>();
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, plan.DocumentId);
            string? missing;
            (slide, source, missing) = FindConversionSource(presentation, unit);
            if (slide is null || source is null)
            {
                result.Reason = PowerPointConversionSkipReason.Changed;
                result.Detail = missing;
                return result;
            }
            StartNewUndoEntry();
            if (target == FormulaOleContract.NativeOleMode)
                redrawFreeze = FreezePowerPointWindowRedraw();
            foreach (var item in unit.Items)
            {
                var formula = prepared[item.Id];
                var replacement = target switch
                {
                    FormulaOleContract.NativeOleMode => CreateOleReplacement(slide, item, formula, oleRestores),
                    FormulaOleContract.CrossPlatformPictureMode => CreatePictureReplacement(slide, unit, item, formula),
                    FormulaOleContract.WordOmmlMode => CreateOmmlReplacement(slide, unit, item, formula),
                    _ => throw new NotSupportedException(target),
                };
                created.Add(replacement);
                MoveToZOrder(replacement, source.ZOrderPosition + 1);
                replacement.Visible = unit.Visible ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            }
            if (target == FormulaOleContract.WordOmmlMode)
                NameConversionReplacements(created, unit.ShapeName, fallbackPrefix: null);
            FinishConversionReplacement(slide, source, created);
            // From here on the source is gone: the replacements are the formula.
            sourceRetired = true;
            result.Converted = true;
            foreach (var replacement in created)
            {
                try { result.ReplacementShapeIds.Add(replacement.Id); } catch { }
            }
            deferRedraw = oleRestores.Count > 0;
            foreach (var restore in oleRestores)
                ScheduleOleGeometryRestore(plan.DocumentId, restore.FormulaId,
                    $"{ManagedReferencePrefix}{unit.SlideId}:{restore.ShapeId}",
                    restore.Width, restore.Height, restore.Left, restore.Top);
            return result;
        }
        catch (Exception error)
        {
            if (sourceRetired) return result;
            foreach (var replacement in created) TryDelete(replacement);
            result.Converted = false;
            result.Reason = PowerPointConversionSkipReason.Failed;
            result.Detail = ApplyFailure(unit, error, Text("写入替换对象", "writing the replacement"));
            return result;
        }
        finally
        {
            FinishPowerPointWindowRedrawFreeze(redrawFreeze, deferRedraw);
            StartNewUndoEntry();
            foreach (var replacement in created) Release(replacement);
            Release(source);
            Release(slide);
            Release(presentation);
        }
    }

    private Shape CreateOleReplacement(
        Slide slide,
        PowerPointConversionItem item,
        PreparedPowerPointConversion formula,
        List<(string FormulaId, int ShapeId, float Width, float Height, float Left, float Top)> restores)
    {
        var metadata = (formula.Session ?? throw new InvalidDataException(Text("缺少 VisualTeX 渲染结果。", "Missing VisualTeX render result."))).ToMetadata();
        metadata.Validate();
        var (left, top) = PowerPointConversionRules.CenteredAt(item.CenterX, item.CenterY, formula.WidthPt, formula.HeightPt);
        var shape = AddOleObjectOffscreen(slide, formula.WidthPt, formula.HeightPt);
        try
        {
            shape.Visible = MsoTriState.msoFalse;
            InitializeOle(shape, metadata,
                formula.EmfPath ?? throw new InvalidDataException(Text("缺少 OLE 对象的预览图。", "Missing OLE preview.")),
                formula.PngPath ?? throw new InvalidDataException(Text("缺少 OLE 对象的图片。", "Missing OLE picture.")));
            Configure(shape, metadata);
            // Host geometry is the very last write: PowerPoint can rebuild the
            // OLE presentation while metadata and container state settle.
            ApplyOleSizeAndRefresh(shape, formula.WidthPt, formula.HeightPt);
            RestoreOlePosition(shape, left, top);
            restores.Add((metadata.FormulaId, shape.Id, formula.WidthPt, formula.HeightPt, left, top));
            return shape;
        }
        catch
        {
            TryDelete(shape);
            Release(shape);
            throw;
        }
    }

    private static Shape CreatePictureReplacement(
        Slide slide,
        PowerPointConversionUnit unit,
        PowerPointConversionItem item,
        PreparedPowerPointConversion formula)
    {
        var metadata = (formula.Session ?? throw new InvalidDataException(Text("缺少 VisualTeX 渲染结果。", "Missing VisualTeX render result."))).ToMetadata();
        metadata.Validate();
        var (left, top) = PowerPointConversionRules.CenteredAt(item.CenterX, item.CenterY, formula.WidthPt, formula.HeightPt);
        var shape = slide.Shapes.AddPicture(
            formula.SvgPath ?? throw new InvalidDataException(Text("缺少 SVG 图片。", "Missing SVG picture.")),
            MsoTriState.msoFalse,
            MsoTriState.msoTrue,
            left,
            top,
            formula.WidthPt,
            formula.HeightPt);
        try
        {
            // PowerPoint can restore an SVG's intrinsic aspect ratio on insert;
            // reapply the exact box (same rule as single conversions).
            ApplyPictureSize(shape, formula.WidthPt, formula.HeightPt);
            shape.Left = left;
            shape.Top = top;
            TryApplyRotation(shape, unit.Rotation);
            if (unit.FlipHorizontal) shape.Flip(MsoFlipCmd.msoFlipHorizontal);
            if (unit.FlipVertical) shape.Flip(MsoFlipCmd.msoFlipVertical);
            Configure(shape, metadata);
            return shape;
        }
        catch
        {
            TryDelete(shape);
            Release(shape);
            throw;
        }
    }

    private Shape CreateOmmlReplacement(
        Slide slide,
        PowerPointConversionUnit unit,
        PowerPointConversionItem item,
        PreparedPowerPointConversion formula)
    {
        var mathMl = formula.MathMl ?? throw new InvalidDataException(Text("缺少用于原生公式的 MathML。", "Missing MathML for the native equation."));
        var fontSize = FormulaFontSize.Normalize(item.EffectiveFontSizePt);
        var metadata = new FormulaMetadata
        {
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "PowerPoint Native Formula",
            Latex = item.Latex,
            Lines = new List<FormulaLine> { new() { Id = Guid.NewGuid().ToString("D"), Latex = item.Latex } },
            CodeFormat = "raw",
            DisplayMode = item.DisplayMode,
            FontSizePt = fontSize,
            RenderFontSizePt = fontSize,
        };
        var width = Math.Max(1f, item.VisualWidth);
        var height = Math.Max(1f, item.VisualHeight);
        var (left, top) = PowerPointConversionRules.CenteredAt(item.CenterX, item.CenterY, width, height);
        var shape = PowerPointNativeOmmlKernel.CreateStandalone(_application, slide, mathMl, metadata, left, top, width, height, item.BaseColor);
        try
        {
            FitTextBoxAroundCenter(shape, item.CenterX, item.CenterY);
            TryApplyRotation(shape, unit.Rotation);
            return shape;
        }
        catch
        {
            TryDelete(shape);
            Release(shape);
            throw;
        }
    }

    // A native equation box hugs its equation, and the equation itself (not
    // just its box: line spacing and alignment offset it inside the box) is
    // centered where the source formula was, so the layout does not shift and
    // no large invisible click target is left behind.
    private static void FitTextBoxAroundCenter(Shape shape, float centerX, float centerY)
    {
        object? frame = null;
        try
        {
            frame = shape.TextFrame2;
            ((dynamic)frame).WordWrap = MsoTriState.msoFalse;
            ((dynamic)frame).AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
        }
        catch { }
        finally { Release(frame); }
        try
        {
            shape.Left = centerX - shape.Width / 2f;
            shape.Top = centerY - shape.Height / 2f;
            var zones = ReadMathZoneGeometry(shape);
            if (zones.Count == 1 && zones[0].Width > 0 && zones[0].Height > 0)
            {
                var zone = zones[0];
                shape.Left += centerX - (zone.Left + zone.Width / 2f);
                shape.Top += centerY - (zone.Top + zone.Height / 2f);
            }
        }
        catch { }
    }

    private const string MathTypeNamePrefix = "MathType_";

    // The first object keeps the source's name, so the Selection and Animation
    // panes (and macros) still find it; further pieces of a split text box get
    // numbered variants. VisualTeX's generated names (VisualTeX_<id>) and
    // internal staging names are never carried to another format.
    private static void NameConversionReplacements(IReadOnlyList<Shape> created, string sourceName, string? fallbackPrefix)
    {
        var keep = !string.IsNullOrWhiteSpace(sourceName)
            && !sourceName.StartsWith("VisualTeX_", StringComparison.Ordinal);
        for (var index = 0; index < created.Count; index++)
        {
            var name = keep
                ? (index == 0 ? sourceName : $"{sourceName} ({index + 1})")
                : fallbackPrefix is null ? null : fallbackPrefix + Guid.NewGuid().ToString("N");
            if (name is null) continue;
            try { created[index].Name = name; } catch { }
        }
    }

    private static void FinishConversionReplacement(Slide slide, Shape source, IReadOnlyList<Shape> created)
    {
        if (created.Count == 1)
        {
            DeleteReplacedShape(slide, source, created[0]);
            return;
        }
        // An equation-only text box split into several objects: connectors and
        // the box's own tags/actions follow the first object. (Animated boxes
        // are never split; capture reports them instead.)
        var first = created[0];
        try { first.Title = source.Title; } catch { }
        CopyMathTypeShapeTags(source, first);
        CopyMathTypeShapeActions(source, first);
        var placeholder = PlaceholderCleanup.Capture(slide, source);
        using var references = TransferMathTypeShapeReferences(slide, source, first);
        source.Delete();
        references.Commit();
        placeholder?.RemoveRecreatedEmptyPlaceholder(slide);
    }

    // The source shape, or (null, null, why it can no longer be converted).
    private (Slide? Slide, Shape? Shape, string? Missing) FindConversionSource(Presentation presentation, PowerPointConversionUnit unit)
    {
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            slides = presentation.Slides;
            try { slide = slides.FindBySlideID(unit.SlideId); }
            catch (Exception error)
            {
                return (null, null, Text("该公式所在的幻灯片已被删除或无法访问：", "The slide holding this equation was deleted or cannot be accessed: ")
                    + Describe(error));
            }
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Shape? shape = shapes[index];
                try
                {
                    if (shape.Id != unit.ShapeId) continue;
                    var mismatch = ConversionSourceMismatch(shape, unit);
                    if (mismatch is not null) return (null, null, mismatch);
                    var found = (slide, shape, (string?)null);
                    slide = null;
                    shape = null;
                    return found;
                }
                finally { Release(shape); }
            }
            return (null, null, Text($"{ShapeLabel(unit)}已不在该幻灯片上（扫描后被删除）。",
                $"{CapitalizedShapeLabel(unit)} is no longer on the slide (it was deleted after the scan)."));
        }
        finally
        {
            Release(shapes);
            Release(slide);
            Release(slides);
        }
    }

    // Null when the shape is still the scanned formula, otherwise why not.
    private static string? ConversionSourceMismatch(Shape shape, PowerPointConversionUnit unit)
    {
        try
        {
            var matches = unit.SourceKind switch
            {
                PowerPointFormulaKind.MathType => IsMathTypeShape(shape),
                PowerPointFormulaKind.Omml => shape.HasTextFrame == MsoTriState.msoTrue,
                PowerPointFormulaKind.VisualTeXOle => IsNativeOle(shape)
                    && string.Equals(ReadEncodedMetadata(shape), unit.MetadataFingerprint, StringComparison.Ordinal),
                _ => !IsNativeOle(shape)
                    && string.Equals(ReadEncodedMetadata(shape), unit.MetadataFingerprint, StringComparison.Ordinal),
            };
            return matches
                ? null
                : Text($"扫描后{ShapeLabel(unit)}已被修改，不再是原来的公式。",
                    $"{CapitalizedShapeLabel(unit)} was modified after the scan and is no longer the same equation.");
        }
        catch (Exception error)
        {
            return Text($"检查{ShapeLabel(unit)}是否仍是原公式时出错：", $"Checking whether {ShapeLabel(unit)} is still the same equation failed: ")
                + Describe(error);
        }
    }

    // ---- Writing: MathType targets -------------------------------------------

    internal PowerPointMathTypeStaging StageMathTypeConversions(
        IReadOnlyList<(PowerPointConversionItem Item, PreparedPowerPointConversion Prepared)> formulas)
    {
        var staging = new PowerPointMathTypeStaging
        {
            Root = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-mathtype-batch-{Guid.NewGuid():N}"),
        };
        Directory.CreateDirectory(staging.Root);
        Presentations? presentations = null;
        Presentation? draft = null;
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            var package = Path.Combine(staging.Root, "equations.pptx");
            presentations = _application.Presentations;
            draft = presentations.Add(MsoTriState.msoFalse);
            slides = draft.Slides;
            slide = slides.Add(1, PpSlideLayout.ppLayoutBlank);
            shapes = slide.Shapes;
            var staged = new List<PowerPointMathTypePackage.StagedEquation>(formulas.Count);
            for (var index = 0; index < formulas.Count; index++)
            {
                var (item, formula) = formulas[index];
                Shape? placeholder = null;
                try
                {
                    placeholder = shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0,
                        Math.Max(1f, formula.WidthPt), Math.Max(1f, formula.HeightPt));
                    placeholder.Name = BatchStagingPlaceholderPrefix + index;
                }
                finally { Release(placeholder); }
                var frameName = BatchStagingFramePrefix + index;
                staging.FrameNames[item.Id] = frameName;
                staged.Add(new PowerPointMathTypePackage.StagedEquation(
                    BatchStagingPlaceholderPrefix + index,
                    frameName,
                    formula.MathTypeCompoundFile ?? throw new InvalidDataException(Text("缺少 MathType 公式数据。", "Missing MathType storage.")),
                    formula.MathTypePreviewEmf ?? throw new InvalidDataException(Text("缺少 MathType 预览图。", "Missing MathType preview.")),
                    formula.MathTypeNaturalWidthPt,
                    formula.MathTypeNaturalHeightPt));
            }
            draft.SaveCopyAs(package, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            Release(shapes); shapes = null;
            Release(slide); slide = null;
            Release(slides); slides = null;
            draft.Close();
            Release(draft); draft = null;

            PowerPointMathTypePackage.WriteEquations(package, staged);
            staging.Presentation = presentations.Open(package, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            return staging;
        }
        catch
        {
            CloseMathTypeStaging(staging);
            throw;
        }
        finally
        {
            Release(shapes);
            Release(slide);
            Release(slides);
            if (draft is not null) { try { draft.Close(); } catch { } }
            Release(draft);
            Release(presentations);
        }
    }

    internal void CloseMathTypeStaging(PowerPointMathTypeStaging? staging)
    {
        if (staging is null) return;
        if (staging.Presentation is not null)
        {
            try { staging.Presentation.Close(); } catch { }
            Release(staging.Presentation);
            staging.Presentation = null;
        }
        if (!string.IsNullOrWhiteSpace(staging.Root))
        {
            try { Directory.Delete(staging.Root, true); } catch { }
        }
    }

    // Pastes a chunk of MathType replacements, proves each pasted storage with
    // one snapshot, then retires the sources of the proven ones. A replacement
    // that cannot be proven is removed and its source is kept.
    internal List<PowerPointConversionUnitResult> ApplyMathTypeConversionChunk(
        PowerPointFormulaConversionPlan plan,
        IReadOnlyList<PowerPointConversionUnit> units,
        IReadOnlyDictionary<string, PreparedPowerPointConversion> prepared,
        PowerPointMathTypeStaging staging)
    {
        var results = new List<PowerPointConversionUnitResult>();
        var pending = new List<(PowerPointConversionUnit Unit, Slide Slide, Shape Source, List<(Shape Shape, PowerPointConversionItem Item)> Created, PowerPointConversionUnitResult Result)>();
        Presentation? presentation = null;
        Slides? stagingSlides = null;
        Slide? stagingSlide = null;
        Shapes? stagingShapes = null;
        string? snapshot = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, plan.DocumentId);
            stagingSlides = staging.Presentation!.Slides;
            stagingSlide = stagingSlides[1];
            stagingShapes = stagingSlide.Shapes;

            using (PowerPointClipboard.Preserve())
            {
                foreach (var unit in units)
                {
                    var result = new PowerPointConversionUnitResult { SlideId = unit.SlideId, FormulaCount = unit.Items.Count };
                    results.Add(result);
                    var (slide, source, missing) = FindConversionSource(presentation, unit);
                    if (slide is null || source is null)
                    {
                        result.Reason = PowerPointConversionSkipReason.Changed;
                        result.Detail = missing;
                        continue;
                    }
                    var created = new List<(Shape Shape, PowerPointConversionItem Item)>();
                    try
                    {
                        StartNewUndoEntry();
                        Shapes? targetShapes = null;
                        try
                        {
                            targetShapes = slide.Shapes;
                            foreach (var item in unit.Items)
                            {
                                var formula = prepared[item.Id];
                                Shape? staged = null;
                                ShapeRange? pasted = null;
                                try
                                {
                                    staged = stagingShapes[staging.FrameNames[item.Id]];
                                    pasted = CopyShapeInto(staged, targetShapes);
                                    if (pasted.Count != 1)
                                        throw new InvalidDataException(Text("PowerPoint 没有恰好粘贴出一个 MathType 对象。", "PowerPoint did not paste exactly one MathType object."));
                                    var replacement = pasted[1];
                                    created.Add((replacement, item));
                                    if (!IsMathTypeShape(replacement))
                                        throw new InvalidDataException(Text("PowerPoint 没有创建 MathType OLE 对象。", "PowerPoint did not create a MathType OLE object."));
                                    var (left, top) = PowerPointConversionRules.CenteredAt(item.CenterX, item.CenterY, formula.WidthPt, formula.HeightPt);
                                    replacement.LockAspectRatio = MsoTriState.msoFalse;
                                    replacement.Width = formula.WidthPt;
                                    replacement.Height = formula.HeightPt;
                                    replacement.Left = left;
                                    replacement.Top = top;
                                    replacement.LockAspectRatio = MsoTriState.msoTrue;
                                    MoveToZOrder(replacement, source.ZOrderPosition + 1);
                                    replacement.Visible = unit.Visible ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                                }
                                finally { Release(pasted); Release(staged); }
                            }
                        }
                        finally { Release(targetShapes); }
                        pending.Add((unit, slide, source, created, result));
                    }
                    catch (Exception error)
                    {
                        foreach (var (shape, _) in created) { TryDelete(shape); Release(shape); }
                        Release(source);
                        Release(slide);
                        result.Reason = PowerPointConversionSkipReason.Failed;
                        result.Detail = ApplyFailure(unit, error, Text("粘贴 MathType 对象", "pasting the MathType object"));
                    }
                }
            }

            if (pending.Count == 0) return results;
            snapshot = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-mathtype-verify-{Guid.NewGuid():N}.pptx");
            presentation.SaveCopyAs(snapshot, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using (var stream = File.OpenRead(snapshot))
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                foreach (var entry in pending)
                {
                    var sourceRetired = false;
                    try
                    {
                        foreach (var (shape, item) in entry.Created)
                        {
                            var saved = PowerPointMathTypePackage.Read(archive, entry.Unit.SlideId, shape.Id);
                            if (!MathTypeMtefCodec.EquationNativeContentEquals(
                                    MathTypeOleStorage.ReadEquationNative(saved.CompoundFile),
                                    prepared[item.Id].MathTypeEquationNative!))
                                throw new InvalidDataException(Text("PowerPoint 保存后 MathType 公式数据与写入的不一致。", "PowerPoint did not preserve the MathType equation storage."));
                        }
                        var shapes = entry.Created.Select(created => created.Shape).ToList();
                        if (shapes.Count == 1)
                        {
                            var replacement = shapes[0];
                            replacement.AlternativeText = AlternativeTextFor(entry.Source);
                            replacement.Title = entry.Source.Title;
                            CopyMathTypeShapeTags(entry.Source, replacement);
                            CopyMathTypeShapeActions(entry.Source, replacement);
                            var placeholder = PlaceholderCleanup.Capture(entry.Slide, entry.Source);
                            using var references = TransferMathTypeShapeReferences(entry.Slide, entry.Source, replacement);
                            NameConversionReplacements(shapes, entry.Unit.ShapeName, MathTypeNamePrefix);
                            entry.Source.Delete();
                            sourceRetired = true;
                            references.Commit();
                            placeholder?.RemoveRecreatedEmptyPlaceholder(entry.Slide);
                        }
                        else
                        {
                            NameConversionReplacements(shapes, entry.Unit.ShapeName, MathTypeNamePrefix);
                            FinishConversionReplacement(entry.Slide, entry.Source, shapes);
                            sourceRetired = true;
                        }
                        entry.Result.Converted = true;
                        foreach (var shape in shapes)
                        {
                            try { entry.Result.ReplacementShapeIds.Add(shape.Id); } catch { }
                        }
                    }
                    catch (Exception error)
                    {
                        if (sourceRetired) { entry.Result.Converted = true; continue; }
                        foreach (var (shape, _) in entry.Created) TryDelete(shape);
                        entry.Result.Reason = PowerPointConversionSkipReason.Failed;
                        entry.Result.Detail = ApplyFailure(entry.Unit, error, Text("校验并完成 MathType 替换", "verifying and finishing the MathType replacement"));
                    }
                    finally
                    {
                        StartNewUndoEntry();
                    }
                }
            }
            return results;
        }
        finally
        {
            foreach (var entry in pending)
            {
                foreach (var (shape, _) in entry.Created) Release(shape);
                Release(entry.Source);
                Release(entry.Slide);
            }
            if (snapshot is not null) { try { File.Delete(snapshot); } catch { } }
            Release(stagingShapes);
            Release(stagingSlide);
            Release(stagingSlides);
            Release(presentation);
        }
    }

    private static string AlternativeTextFor(Shape source)
    {
        try
        {
            var text = source.AlternativeText;
            return !string.IsNullOrWhiteSpace(text) && FormulaMetadataCodec.Decode(text) is null ? text : string.Empty;
        }
        catch { return string.Empty; }
    }

    // ---- Afterwards ------------------------------------------------------------

    // Back to the slide the user started from; a selection conversion leaves the
    // new objects selected so the user can see exactly what changed.
    internal void FinishFormulaConversion(PowerPointFormulaConversionPlan plan, IReadOnlyList<(int SlideId, int ShapeId)> replacements)
    {
        DocumentWindow? window = null;
        View? view = null;
        Presentation? presentation = null;
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        ShapeRange? range = null;
        try
        {
            window = _application.ActiveWindow;
            presentation = _application.ActivePresentation;
            if (window is null || presentation is null) return;
            try
            {
                if (window.ViewType != plan.OriginalViewType) window.ViewType = plan.OriginalViewType;
            }
            catch { }
            if (plan.OriginalSlideId == 0) return;
            slides = presentation.Slides;
            try { slide = slides.FindBySlideID(plan.OriginalSlideId); }
            catch { return; }
            view = window.View;
            try
            {
                Slide? currentSlide = null;
                try { currentSlide = (Slide)view.Slide; } catch { }
                if (currentSlide is null || currentSlide.SlideID != slide.SlideID) view.GotoSlide(slide.SlideIndex);
                Release(currentSlide);
            }
            catch { }
            var selectable = replacements.Where(entry => entry.SlideId == plan.OriginalSlideId).Select(entry => entry.ShapeId).ToHashSet();
            try
            {
                if (plan.Request.Scope != PowerPointConversionScope.Selection || selectable.Count == 0)
                {
                    Selection? selection = null;
                    try
                    {
                        selection = window.Selection;
                        selection.Unselect();
                    }
                    finally { Release(selection); }
                    return;
                }
                shapes = slide.Shapes;
                var indexes = new List<object>();
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Shape? shape = null;
                    try
                    {
                        shape = shapes[index];
                        if (selectable.Contains(shape.Id) && shape.Visible == MsoTriState.msoTrue) indexes.Add(index);
                    }
                    finally { Release(shape); }
                }
                if (indexes.Count == 0) return;
                range = shapes.Range(indexes.ToArray());
                range.Select(MsoTriState.msoTrue);
            }
            catch { }
        }
        catch { }
        finally
        {
            Release(range);
            Release(shapes);
            Release(slide);
            Release(slides);
            Release(view);
            Release(presentation);
            Release(window);
        }
    }

    private static string Text(string chinese, string english) =>
        VisualTeX.WindowsOffice.VstoShared.OfficePluginLanguage.Text(chinese, english);

    // The specific, localized reason for an exception (COM code, every cause).
    private static string Describe(Exception? error) =>
        VisualTeX.WindowsOffice.VstoShared.OfficeErrorText.Describe(error);

    // "形状“Equation 3”（ID 12）" / "shape \"Equation 3\" (ID 12)": which object failed.
    private static string ShapeLabel(PowerPointConversionUnit unit) =>
        string.IsNullOrWhiteSpace(unit.ShapeName)
            ? Text($"形状（ID {unit.ShapeId}）", $"shape ID {unit.ShapeId}")
            : Text($"形状“{unit.ShapeName}”（ID {unit.ShapeId}）", $"shape \"{unit.ShapeName}\" (ID {unit.ShapeId})");

    private static string CapitalizedShapeLabel(PowerPointConversionUnit unit)
    {
        var label = ShapeLabel(unit);
        return label.Length > 0 && char.IsLower(label[0]) ? char.ToUpperInvariant(label[0]) + label.Substring(1) : label;
    }

    private static string ApplyFailure(PowerPointConversionUnit unit, Exception error, string stage) =>
        Text($"{ShapeLabel(unit)}：{stage}失败：", $"{CapitalizedShapeLabel(unit)}: {stage} failed: ") + Describe(error);

    // The LaTeX shown for a unit's report line (one formula only; a multi-
    // formula text box is identified by its shape and count instead).
    internal static string? UnitLatex(PowerPointConversionUnit unit) =>
        unit.Items.Count == 1 ? unit.Items[0].Latex : null;

    // Deleting a filled content placeholder makes PowerPoint put an empty
    // "Click to add text" placeholder back. After a conversion the formula
    // lives in its own object, so that leftover prompt is removed too. Title
    // placeholders are kept: they carry the slide's title for navigation.
    internal sealed class PlaceholderCleanup
    {
        private readonly HashSet<int> _existing;
        private PlaceholderCleanup(HashSet<int> existing) => _existing = existing;

        internal static PlaceholderCleanup? Capture(Slide slide, Shape source)
        {
            try
            {
                if (source.Type != MsoShapeType.msoPlaceholder) return null;
                PlaceholderFormat? format = null;
                try
                {
                    format = source.PlaceholderFormat;
                    if (format.Type is PpPlaceholderType.ppPlaceholderTitle
                        or PpPlaceholderType.ppPlaceholderCenterTitle
                        or PpPlaceholderType.ppPlaceholderVerticalTitle)
                        return null;
                }
                finally { Release(format); }
                return new PlaceholderCleanup(PlaceholderIds(slide));
            }
            catch { return null; }
        }

        internal void RemoveRecreatedEmptyPlaceholder(Slide slide)
        {
            Shapes? shapes = null;
            Placeholders? placeholders = null;
            try
            {
                shapes = slide.Shapes;
                placeholders = shapes.Placeholders;
                for (var index = placeholders.Count; index >= 1; index--)
                {
                    Shape? placeholder = null;
                    try
                    {
                        placeholder = placeholders[index];
                        if (_existing.Contains(placeholder.Id) || !IsEmptyPlaceholder(placeholder)) continue;
                        placeholder.Delete();
                    }
                    catch { }
                    finally { Release(placeholder); }
                }
            }
            catch { }
            finally { Release(placeholders); Release(shapes); }
        }

        private static HashSet<int> PlaceholderIds(Slide slide)
        {
            var ids = new HashSet<int>();
            Shapes? shapes = null;
            Placeholders? placeholders = null;
            try
            {
                shapes = slide.Shapes;
                placeholders = shapes.Placeholders;
                for (var index = 1; index <= placeholders.Count; index++)
                {
                    Shape? placeholder = null;
                    try
                    {
                        placeholder = placeholders[index];
                        ids.Add(placeholder.Id);
                    }
                    finally { Release(placeholder); }
                }
            }
            finally { Release(placeholders); Release(shapes); }
            return ids;
        }

        private static bool IsEmptyPlaceholder(Shape placeholder)
        {
            PlaceholderFormat? format = null;
            try
            {
                format = placeholder.PlaceholderFormat;
                if (format.ContainedType != MsoShapeType.msoAutoShape
                    && format.ContainedType != MsoShapeType.msoTextBox)
                    return false;
                if (placeholder.HasTextFrame != MsoTriState.msoTrue) return true;
                var frame = placeholder.TextFrame2;
                try { return frame.HasText != MsoTriState.msoTrue; }
                finally { Release(frame); }
            }
            catch { return false; }
            finally { Release(format); }
        }
    }
}

// Blocks user input to PowerPoint while a batch writes, the way a modal dialog
// does: every PowerPoint frame window of this process is disabled and later
// re-enabled. Programmatic edits, painting and the progress window keep working.
internal sealed class PowerPointModalGuard : IDisposable
{
    private readonly List<IntPtr> _disabled = new();

    private PowerPointModalGuard() { }

    internal static PowerPointModalGuard Enter()
    {
        var guard = new PowerPointModalGuard();
        var processId = (uint)Process.GetCurrentProcess().Id;
        EnumWindows((window, _) =>
        {
            GetWindowThreadProcessId(window, out var owner);
            if (owner != processId || !IsWindowVisible(window) || !IsWindowEnabled(window)) return true;
            var name = new StringBuilder(64);
            GetClassName(window, name, name.Capacity);
            if (!string.Equals(name.ToString(), "PPTFrameClass", StringComparison.Ordinal)) return true;
            EnableWindow(window, false);
            guard._disabled.Add(window);
            return true;
        }, IntPtr.Zero);
        return guard;
    }

    public void Dispose()
    {
        foreach (var window in _disabled)
        {
            try
            {
                if (IsWindow(window)) EnableWindow(window, true);
            }
            catch { }
        }
        _disabled.Clear();
    }

    private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool IsWindowEnabled(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool EnableWindow(IntPtr window, bool enable);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder name, int capacity);
}

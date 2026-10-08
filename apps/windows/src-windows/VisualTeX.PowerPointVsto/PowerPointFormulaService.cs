using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using View = Microsoft.Office.Interop.PowerPoint.View;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.PowerPointVsto;

public sealed partial class PowerPointFormulaService : IDisposable
{
    private const string FormulaIdTag = "VisualTeXFormulaId";
    private const string MetadataTag = "VisualTeXMetadata";
    private const string IdentityOwnerTag = "VisualTeXIdentityOwner";
    private const string SlideReferencePrefix = "visualtex-ppt-vsto-slide:";
    private const string TextInsertionReferencePrefix = "visualtex-ppt-vsto-text:";
    private const uint WmSetRedraw = 0x000B;
    private const uint RdwInvalidate = 0x0001;
    private const uint RdwErase = 0x0004;
    private const uint RdwAllChildren = 0x0080;
    private const uint RdwUpdateNow = 0x0100;
    private const uint RdwFrame = 0x0400;
    private readonly Application _application;
    private readonly Action<Action>? _postToOfficeUi;
    private readonly Action<Action, int>? _postDelayedToOfficeUi;
    private readonly Action<string, Shape>? _oleStageProbe;
    private readonly Action<string, IntPtr>? _windowRedrawProbe;
    private readonly Dictionary<string, long> _oleGeometryRestoreGenerations = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<IntPtr, long> _windowRedrawFreezeGenerations = new();
    private long _nextWindowRedrawFreezeGeneration;

    public PowerPointFormulaService(
        Application application,
        Action<Action>? postToOfficeUi = null,
        Action<Action, int>? postDelayedToOfficeUi = null,
        Action<string, Shape>? oleStageProbe = null,
        Action<string, IntPtr>? windowRedrawProbe = null)
    {
        _application = application;
        _postToOfficeUi = postToOfficeUi;
        _postDelayedToOfficeUi = postDelayedToOfficeUi;
        _oleStageProbe = oleStageProbe;
        _windowRedrawProbe = windowRedrawProbe;
    }

    public void Dispose() =>
        PowerPointNativeOmmlKernel.Shutdown();

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern IntPtr SendMessage(
        IntPtr hWnd,
        uint message,
        IntPtr wParam,
        IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool RedrawWindow(
        IntPtr hWnd,
        IntPtr updateRect,
        IntPtr updateRegion,
        uint flags);

    public OfficeSelection ReadSelection() => ReadSelection(null);

    public OfficeSelection CaptureSlideInsertionTarget()
    {
        var selection = ReadSelection();
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        try
        {
            window = _application.ActiveWindow;
            view = window.View;
            slide = (Slide)view.Slide;
            selection.ObjectId = SlideReference(slide);
            selection.Metadata = null;
            selection.FormulaId = null;
            selection.ObjectMode = null;
            return selection;
        }
        finally { Release(slide); Release(view); Release(window); }
    }

    public OfficeSelection ReadSelection(Selection? providedSelection)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Selection? selection = null;
        ShapeRange? range = null;
        Shape? shape = null;
        var ownsSelection = providedSelection is null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = (Slide)view.Slide;
            selection = providedSelection ?? window.Selection;
            FormulaMetadata? metadata = null;
            string? objectMode = null;
            string? objectId = SlideReference(slide);
            if (selection.Type is PpSelectionType.ppSelectionShapes or PpSelectionType.ppSelectionText)
            {
                try { range = selection.ShapeRange; } catch { range = null; }
                if (range?.Count == 1)
                {
                    shape = range[1];
                    objectId = shape.Name;
                    var isMathType = IsMathTypeShape(shape);
                    if (isMathType)
                    {
                        objectMode = FormulaOleContract.MathTypeOleMode;
                        objectId = MathTypeReference(slide, shape);
                    }
                    var storedMetadata = isMathType ? null : ReadMetadata(shape);
                    if (storedMetadata is not null)
                    {
                        metadata = storedMetadata;
                        objectId = ManagedReference(slide, shape);
                        metadata.FontSizePt = InferPowerPointFormulaFontSize(
                            shape.Width,
                            shape.Height,
                            metadata);
                        objectMode = IsNativeOle(shape)
                            ? FormulaOleContract.NativeOleMode
                            : FormulaOleContract.CrossPlatformPictureMode;
                    }
                    else if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        // A descriptor is only a candidate. Native OMML is proven
                        // by exact package source capture on the explicit command;
                        // passive selection does not call MathZones or copy data.
                        objectId = NativeOmmlReference(slide, shape);
                        if (selection.Type == PpSelectionType.ppSelectionText)
                        {
                            object? text = null;
                            try
                            {
                                text = ((dynamic)selection).TextRange2;
                                objectId = NativeOmmlTextReference(slide, shape, text);
                            }
                            finally { Release(text); }
                        }
                    }
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
                ObjectMode = objectMode,
            };
        }
        finally
        {
            Release(shape);
            Release(range);
            if (ownsSelection) Release(selection);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeSelection CaptureTextInsertionTarget() =>
        CaptureTextInsertionTargetCore(requireTextSelection: true)
        ?? throw new InvalidOperationException(
            "请先把文本光标或文本选区放到一个现有文本框中。");

    public OfficeSelection? TryCaptureTextInsertionTarget() =>
        CaptureTextInsertionTargetCore(requireTextSelection: false);

    private OfficeSelection? CaptureTextInsertionTargetCore(
        bool requireTextSelection)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Selection? selection = null;
        ShapeRange? range = null;
        Shape? shape = null;
        object? textRange = null;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = (Slide)view.Slide;
            selection = window.Selection;
            if (selection.Type != PpSelectionType.ppSelectionText)
            {
                if (!requireTextSelection) return null;
                throw new InvalidOperationException(
                    "请先把文本光标或文本选区放到一个现有文本框中。");
            }

            range = selection.ShapeRange;
            if (range.Count != 1)
                throw new InvalidOperationException("请只在一个文本框中插入 OMML 公式。");
            shape = range[1];
            if (shape.HasTextFrame != MsoTriState.msoTrue)
                throw new InvalidOperationException("当前选择不是可编辑文本框。");

            dynamic selected = selection;
            textRange = selected.TextRange2;
            var start = Convert.ToInt32(((dynamic)textRange).Start);
            var length = Convert.ToInt32(((dynamic)textRange).Length);
            if (start < 1 || length < 0)
                throw new InvalidOperationException("PowerPoint 没有返回有效的文本插入位置。");

            return new OfficeSelection
            {
                Host = "powerpoint",
                DocumentId = DocumentIdentity(presentation),
                ObjectId = TextInsertionReference(
                    slide.SlideID,
                    shape.Id,
                    start,
                    length,
                    PowerPointTextEditTransaction.Fingerprint(
                        PowerPointNativeOmmlPackage.ReadTextShapeXml(presentation, slide.SlideID, shape.Id))),
                ReadOnly = false,
                FormulaId = null,
                Metadata = null,
                ObjectMode = FormulaOleContract.WordOmmlMode,
            };
        }
        finally
        {
            Release(textRange);
            Release(shape);
            Release(range);
            Release(selection);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeSelection? ReadFormulaAtScreenPoint(int screenX, int screenY)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        object? hit = null;
        object? hitParent = null;
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
            try { hit = window.RangeFromPoint(screenX, screenY); }
            catch { return null; }
            shape = hit as Shape;
            if (shape is null && (hit is TextRange || hit is Microsoft.Office.Core.TextRange2))
            {
                hitParent = ((dynamic)hit).Parent;
                shape = ((dynamic)hitParent).Parent as Shape;
            }
            if (shape is null || IsMathTypeShape(shape)) return null;

            var metadata = ReadMetadata(shape);
            if (metadata is null)
            {
                var nativeReference = HitNativeMathZone(window, slide, shape, screenX, screenY);
                return nativeReference is null ? null : new OfficeSelection
                {
                    Host = "powerpoint", DocumentId = DocumentIdentity(presentation),
                    ObjectId = nativeReference, ReadOnly = presentation.ReadOnly == MsoTriState.msoTrue,
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                };
            }
            metadata.FontSizePt = InferPowerPointFormulaFontSize(
                shape.Width,
                shape.Height,
                metadata);
            var objectMode = IsNativeOle(shape)
                ? FormulaOleContract.NativeOleMode
                : FormulaOleContract.CrossPlatformPictureMode;

            return new OfficeSelection
            {
                Host = "powerpoint",
                DocumentId = DocumentIdentity(presentation),
                ObjectId = ManagedReference(slide, shape),
                ReadOnly = presentation.ReadOnly == MsoTriState.msoTrue,
                FormulaId = metadata.FormulaId,
                Metadata = metadata,
                ObjectMode = objectMode,
            };
        }
        finally
        {
            if (shape is not null)
            {
                Release(shape);
                if (hitParent is null) hit = null;
            }
            Release(hitParent);
            Release(hit);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public float ReadCurrentTypingFontSize()
    {
        object? mathRange = null;
        object? font = null;
        try
        {
            if (TryGetSelectedMathRange(out mathRange))
            {
                dynamic range = mathRange!;
                font = range.Font;
                var size = Convert.ToDouble(((dynamic)font).Size);
                return FormulaFontSize.Normalize(size, 20f);
            }

            var window = _application.ActiveWindow;
            if (window is null) return 20f;
            Selection? selection = null;
            try
            {
                selection = window.Selection;
                dynamic selected = selection;
                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    object textRange = selected.TextRange2;
                    try
                    {
                        font = ((dynamic)textRange).Font;
                        var size = Convert.ToDouble(((dynamic)font).Size);
                        return FormulaFontSize.Normalize(size, 20f);
                    }
                    finally { Release(textRange); }
                }
            }
            finally { Release(selection); Release(window); }
            return 20f;
        }
        catch { return 20f; }
        finally
        {
            Release(font);
            Release(mathRange);
        }
    }

    public float? GetSelectedFormulaFontSize()
    {
        try
        {
            var selected = ReadSelection();
            if (selected.Metadata is not null)
                return FormulaFontSize.ResolveSemanticFontSize(selected.Metadata);
            if (selected.ObjectMode == FormulaOleContract.WordOmmlMode)
                return NativeShapeFontSize(selected);
        }
        catch { }

        object? mathRange = null;
        object? font = null;
        try
        {
            if (!TryGetSelectedMathRange(out mathRange)) return null;
            dynamic range = mathRange!;
            font = range.Font;
            var size = Convert.ToDouble(((dynamic)font).Size);
            return FormulaFontSize.Normalize(size, 18f);
        }
        catch { return null; }
        finally
        {
            Release(font);
            Release(mathRange);
        }
    }

    public float SetSelectedFormulaFontSize(double requestedFontSizePt)
    {
        var target = FormulaFontSize.Normalize(requestedFontSizePt);
        var selected = ReadSelection();
        if (selected.ReadOnly)
            throw new UnauthorizedAccessException("The active PowerPoint presentation is read-only.");
        if (selected.Metadata is null || string.IsNullOrWhiteSpace(selected.FormulaId))
        {
            if (SetSelectedMathZoneFontSize(target)) return target;
            if (IsNativeOmmlReference(selected.ObjectId)
                && NativeShapeFontSize(selected, target).HasValue) return target;
            throw new InvalidOperationException("请先选择一个 VisualTeX 公式或 PowerPoint 原生公式。");
        }

        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            (slide, shape) = FindFormula(presentation, selected.FormulaId!, selected.ObjectId);
            if (slide is null || shape is null)
                throw new InvalidOperationException("The selected PowerPoint formula no longer exists.");

            var metadata = selected.Metadata;
            var currentFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
            var fontScale = target / Math.Max(0.5f, currentFontSize);
            var size = ScaleCurrentShapeSize(
                shape.Width,
                shape.Height,
                fontScale,
                600f,
                400f);
            metadata.FontSizePt = target;
            // Persist the semantic size before changing PowerPoint geometry.
            // Width/Height mutations can synchronously dispatch selection events;
            // those event handlers must not observe and re-save the old font size.
            Configure(shape, metadata);
            var centerX = shape.Left + shape.Width / 2f;
            var centerY = shape.Top + shape.Height / 2f;
            if (IsNativeOle(shape))
                ApplyOleSizeAndRefresh(shape, size.Width, size.Height);
            else
            {
                shape.LockAspectRatio = MsoTriState.msoFalse;
                shape.Width = size.Width;
                shape.Height = size.Height;
                shape.LockAspectRatio = MsoTriState.msoTrue;
            }
            shape.Left = centerX - size.Width / 2f;
            shape.Top = centerY - size.Height / 2f;
            Configure(shape, metadata);
            return target;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(presentation);
        }
    }

    private bool SetSelectedMathZoneFontSize(float target)
    {
        object? mathRange = null;
        object? font = null;
        try
        {
            if (!TryGetSelectedMathRange(out mathRange)) return false;
            StartNewUndoEntry();
            dynamic range = mathRange!;
            font = range.Font;
            ((dynamic)font).Size = target;
            return true;
        }
        finally
        {
            StartNewUndoEntry();
            Release(font);
            Release(mathRange);
        }
    }

    private bool IsCurrentSelectionNativeEquation()
    {
        DocumentWindow? window = null;
        Selection? selection = null;
        try
        {
            window = _application.ActiveWindow;
            if (window is null) return false;
            selection = window.Selection;
            if (selection.Type != PpSelectionType.ppSelectionText)
                return false;

            return _application.CommandBars.GetEnabledMso(
                       "EquationProfessional")
                || _application.CommandBars.GetEnabledMso(
                    "EquationLinearFormat");
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(selection);
            Release(window);
        }
    }

    private bool TryGetSelectedMathRange(out object? mathRange)
    {
        mathRange = null;
        DocumentWindow? window = null;
        Selection? selection = null;
        object? textRange = null;
        try
        {
            window = _application.ActiveWindow;
            if (window is null) return false;
            selection = window.Selection;
            if (selection.Type != PpSelectionType.ppSelectionText)
                return false;

            var isNativeEquation =
                _application.CommandBars.GetEnabledMso(
                    "EquationProfessional")
                || _application.CommandBars.GetEnabledMso(
                    "EquationLinearFormat");
            if (!isNativeEquation)
                return false;

            dynamic selected = selection;
            textRange = selected.TextRange2;
            // Command state already established an explicit native math target.
            // Resolve its single indexed zone, never the full collection span.
            mathRange = PowerPointOmmlBridge.ReadSingleSelectedMathZone(textRange);
            return mathRange is not null;
        }
        catch
        {
            Release(mathRange);
            mathRange = null;
            return false;
        }
        finally
        {
            Release(textRange);
            Release(selection);
            Release(window);
        }
    }

    public string DeleteSelectedFormula()
    {
        var selected = ReadSelection();
        if (IsNativeOmmlReference(selected.ObjectId))
            selected = LoadNativeOmmlSelection(selected);
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            (slide, shape) = FindFormula(presentation, requiredFormulaId, selected.ObjectId);
            if (slide is null || shape is null)
                throw new InvalidOperationException("The selected PowerPoint formula no longer exists.");
            shape.Delete();
            return requiredFormulaId;
        }
        finally
        {
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(presentation);
        }
    }

    public string ExportSelectedOleAsPicture()
    {
        var selected = ReadSelection();
        var formulaId = selected.FormulaId;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new InvalidOperationException("Please select one VisualTeX formula first.");
        var requiredFormulaId = formulaId!;

        Presentation? presentation = null;
        Slide? slide = null;
        Shape? oldShape = null;
        Shape? replacement = null;
        OLEFormat? format = null;
        object? oleObject = null;
        string? pngPath = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            (slide, oldShape) = FindFormula(presentation, requiredFormulaId, selected.ObjectId);
            if (slide is null || oldShape is null)
                throw new InvalidOperationException("The selected PowerPoint formula no longer exists.");
            var metadata = ReadMetadata(oldShape)
                ?? throw new InvalidDataException("The selected formula metadata is invalid.");
            format = oldShape.OLEFormat;
            if (!string.Equals(
                    format.ProgID,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("The selected formula is already a picture.");
            oleObject = format.Object;
            pngPath = OlePngPreviewExtractor.MaterializePng(oleObject, requiredFormulaId);

            var left = oldShape.Left;
            var top = oldShape.Top;
            var width = oldShape.Width;
            var height = oldShape.Height;
            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            replacement = slide.Shapes.AddPicture(
                pngPath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                left,
                top,
                width,
                height);
            TryApplyRotation(replacement, rotation);
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            oldShape.Delete();
            return requiredFormulaId;
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            if (pngPath is not null)
            {
                try { File.Delete(pngPath); } catch { }
            }
            StartNewUndoEntry();
            Release(oleObject);
            Release(format);
            Release(replacement);
            Release(oldShape);
            Release(slide);
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
            var (left, top) = NewFormulaPosition(presentation, slide, width, height);
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

    public OfficeObjectResult InsertOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Shape? shape = null;
        (IntPtr Hwnd, long Generation)? redrawFreeze = null;
        var deferRedrawRestore = false;
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
            var (left, top) = NewFormulaPosition(presentation, slide, width, height);
            redrawFreeze = FreezePowerPointWindowRedraw(window);
            shape = AddOleObjectOffscreen(slide, width, height);
            ProbeOleStage("allocated", shape);
            shape.Visible = MsoTriState.msoFalse;
            ProbeOleStage("created", shape);
            InitializeOle(shape, metadata, emfPath, pngPath);
            Configure(shape, metadata);
            // PowerPoint can rebuild the OLE presentation while metadata and
            // container state are being finalized. Host geometry must therefore
            // be the very last mutation in the write path.
            ApplyOleSizeAndRefresh(shape, width, height);
            RestoreOlePosition(shape, left, top);
            shape.Visible = MsoTriState.msoTrue;
            ProbeOleStage("finalized", shape);
            ScheduleOleGeometryRestore(
                DocumentIdentity(presentation),
                metadata.FormulaId,
                ManagedReference(slide, shape),
                width,
                height,
                left,
                top);
            deferRedrawRestore = true;
            return Result(session, presentation, shape.Name);
        }
        catch
        {
            TryDelete(shape);
            throw;
        }
        finally
        {
            FinishPowerPointWindowRedrawFreeze(redrawFreeze, deferRedrawRestore);
            StartNewUndoEntry();
            Release(shape);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeObjectResult InsertOmml(OfficeSessionDocument session)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        var mathMl = session.ExportResult?.MathMl;
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("VisualTeX Session has no MathML export for PowerPoint OMML.");
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Shape? shape = null;
        Shape? targetTextShape = null;
        try
        {
            EnsureNotSlideShow();
            StartNewUndoEntry();
            presentation = _application.ActivePresentation
                ?? throw new InvalidOperationException("No active PowerPoint presentation.");
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);

            if (TryParseTextInsertionReference(session.SourceObjectId, out var textTarget))
            {
                if (string.IsNullOrWhiteSpace(session.SourceDocumentId))
                    throw new InvalidOperationException("The captured text target has no presentation identity; reopen the editor.");
                (slide, targetTextShape) = ResolveTextInsertionTarget(
                    presentation,
                    textTarget);
                var writeText = PowerPointNativeOmmlKernel.PrepareTextInsertion(
                    _application,
                    targetTextShape,
                    textTarget.Start,
                    textTarget.Length,
                    mathMl!,
                    session.DisplayMode);
                PowerPointTextEditTransaction.Execute(_application, presentation,
                    textTarget.SlideId, textTarget.ShapeId, textTarget.SourceFingerprint, writeText);
                return Result(session, presentation, targetTextShape.Name);
            }

            window = _application.ActiveWindow
                ?? throw new InvalidOperationException("No active PowerPoint window.");
            view = window.View;
            slide = ResolveTargetSlide(presentation, session.SourceObjectId, view);
            var width = Math.Max(12f, (session.ExportResult?.Width ?? 240) * 0.75f);
            var height = Math.Max(12f, (session.ExportResult?.Height ?? 80) * 0.75f);
            var scale = Math.Min(1f, Math.Min(600f / width, 400f / height));
            width *= scale;
            height *= scale;
            var (left, top) = NewFormulaPosition(presentation, slide, width, height);
            shape = PowerPointNativeOmmlKernel.CreateStandalone(
                _application,
                slide,
                mathMl!,
                metadata,
                left,
                top,
                width,
                height);
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
            Release(targetTextShape);
            Release(shape);
            Release(slide);
            Release(view);
            Release(window);
            Release(presentation);
        }
    }

    public OfficeObjectResult ReplaceOmml(OfficeSessionDocument session)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        var mathMl = session.ExportResult?.MathMl;
        if (string.IsNullOrWhiteSpace(mathMl))
            throw new InvalidDataException("VisualTeX Session has no MathML export for PowerPoint OMML.");
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
            if (IsNativeOmmlReference(session.SourceObjectId))
                return ReplaceNativeOmmlText(session, presentation, slide, oldShape, mathMl!);
            EnsureMathTypeSourceUnchanged(session, presentation, slide, oldShape);

            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            var originalMetadata = ReadMetadata(oldShape) ?? session.OriginalMetadata;
            var convertingToOmml = true;
            var editedSize = convertingToOmml
                && FormulaContentEquivalent(originalMetadata, metadata)
                    ? (Width: oldWidth, Height: oldHeight)
                    : OfficeFormulaSizing.EditedSize(
                        oldWidth,
                        oldHeight,
                        originalMetadata?.RenderWidthPx,
                        originalMetadata?.RenderHeightPx,
                        session.ExportResult?.Width ?? oldWidth / 0.75f,
                        session.ExportResult?.Height ?? oldHeight / 0.75f,
                        600f,
                        400f,
                        originalMetadata?.FontSizePt,
                        originalMetadata?.RenderFontSizePt);
            var newLeft = left + (oldWidth - editedSize.Width) / 2f;
            var newTop = top + (oldHeight - editedSize.Height) / 2f;

            replacement = PowerPointNativeOmmlKernel.CreateStandalone(
                _application,
                slide,
                mathMl!,
                metadata,
                newLeft,
                newTop,
                editedSize.Width,
                editedSize.Height);
            TryApplyRotation(replacement, rotation);
            MoveToZOrder(replacement, zOrder + 1);
            DeleteReplacedShape(slide, oldShape, replacement);
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

    public OfficeObjectResult ReplaceOle(
        OfficeSessionDocument session,
        string pngPath,
        string emfPath)
    {
        var metadata = session.ToMetadata();
        metadata.Validate();
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? oldShape = null;
        Shape? replacement = null;
        (IntPtr Hwnd, long Generation)? redrawFreeze = null;
        var deferRedrawRestore = false;
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
            EnsureMathTypeSourceUnchanged(session, presentation, slide, oldShape);

            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var originalMetadata = ReadMetadata(oldShape) ?? session.OriginalMetadata;
            var editedSize = ResolvePowerPointEditedSize(
                oldWidth,
                oldHeight,
                originalMetadata,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                FormulaContentEquivalent(originalMetadata, metadata));
            var newLeft = left + (oldWidth - editedSize.Width) / 2f;
            var newTop = top + (oldHeight - editedSize.Height) / 2f;
            redrawFreeze = FreezePowerPointWindowRedraw();

            if (!FormulaFontPreferencesChanged(originalMetadata, metadata)
                && TryUpdateOle(oldShape, metadata, emfPath, pngPath))
            {
                Configure(oldShape, metadata);
                // Updating the embedded presentation may cause PowerPoint to
                // reinterpret the cached preview. Restore the host box only after
                // every metadata/container operation has completed.
                ApplyOleSizeAndRefresh(oldShape, editedSize.Width, editedSize.Height);
                RestoreOlePosition(oldShape, newLeft, newTop);
                ScheduleOleGeometryRestore(
                    DocumentIdentity(presentation),
                    metadata.FormulaId,
                    ManagedReference(slide, oldShape),
                    editedSize.Width,
                    editedSize.Height,
                    newLeft,
                    newTop);
                deferRedrawRestore = true;
                return Result(session, presentation, oldShape.Name);
            }

            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            replacement = AddOleObjectOffscreen(
                slide,
                editedSize.Width,
                editedSize.Height);
            ProbeOleStage("allocated", replacement);
            replacement.Visible = MsoTriState.msoFalse;
            ProbeOleStage("created", replacement);
            InitializeOle(replacement, metadata, emfPath, pngPath);
            TryApplyRotation(replacement, rotation);
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            // The window is still redraw-suspended here, so deleting the source
            // does not expose a blank frame. Finalize the OLE completely before
            // the deferred redraw is allowed to paint the slide again.
            ApplyOleSizeAndRefresh(replacement, editedSize.Width, editedSize.Height);
            RestoreOlePosition(replacement, newLeft, newTop);
            replacement.Visible = MsoTriState.msoTrue;
            ProbeOleStage("finalized", replacement);
            DeleteReplacedShape(slide, oldShape, replacement);
            try { replacement.Select(MsoTriState.msoTrue); } catch { }
            ScheduleOleGeometryRestore(
                DocumentIdentity(presentation),
                metadata.FormulaId,
                ManagedReference(slide, replacement),
                editedSize.Width,
                editedSize.Height,
                newLeft,
                newTop);
            deferRedrawRestore = true;
            return Result(session, presentation, replacement.Name);
        }
        catch
        {
            TryDelete(replacement);
            throw;
        }
        finally
        {
            FinishPowerPointWindowRedrawFreeze(redrawFreeze, deferRedrawRestore);
            StartNewUndoEntry();
            Release(replacement);
            Release(oldShape);
            Release(slide);
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
            EnsureMathTypeSourceUnchanged(session, presentation, slide, oldShape);

            var left = oldShape.Left;
            var top = oldShape.Top;
            var oldWidth = oldShape.Width;
            var oldHeight = oldShape.Height;
            var rotation = oldShape.Rotation;
            var zOrder = oldShape.ZOrderPosition;
            var originalMetadata = ReadMetadata(oldShape) ?? session.OriginalMetadata;
            var editedSize = ResolvePowerPointEditedSize(
                oldWidth,
                oldHeight,
                originalMetadata,
                session.ExportResult?.Width ?? oldWidth / 0.75f,
                session.ExportResult?.Height ?? oldHeight / 0.75f,
                FormulaContentEquivalent(originalMetadata, metadata));
            var newLeft = left + (oldWidth - editedSize.Width) / 2f;
            var newTop = top + (oldHeight - editedSize.Height) / 2f;

            replacement = slide.Shapes.AddPicture(
                imagePath,
                MsoTriState.msoFalse,
                MsoTriState.msoTrue,
                newLeft,
                newTop,
                editedSize.Width,
                editedSize.Height);
            // PowerPoint 2021 can ignore one of the AddPicture dimensions for
            // SVG and immediately restore the file's intrinsic aspect ratio.
            // Reapply the exact formula box after the SVG shape exists so a
            // lossless OLE→picture conversion keeps the user's physical size.
            ApplyPictureSize(replacement, editedSize.Width, editedSize.Height);
            replacement.Left = newLeft;
            replacement.Top = newTop;
            TryApplyRotation(replacement, rotation);
            Configure(replacement, metadata);
            MoveToZOrder(replacement, zOrder + 1);
            DeleteReplacedShape(slide, oldShape, replacement);
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

    private static Shape AddOleObjectOffscreen(
        Slide slide,
        float width,
        float height)
    {
        // AddOLEObject creates a visible Shape synchronously. On a warm OLE
        // server PowerPoint can paint that placeholder before the next managed
        // statement has a chance to set Visible=false. Stage the object fully
        // outside the slide so even that pre-hide frame cannot flash onscreen.
        var stagingLeft = -Math.Max(2048f, width + 256f);
        var stagingTop = -Math.Max(2048f, height + 256f);
        return slide.Shapes.AddOLEObject(
            stagingLeft,
            stagingTop,
            width,
            height,
            FormulaOleContract.ProgId,
            string.Empty,
            MsoTriState.msoFalse,
            string.Empty,
            0,
            string.Empty,
            MsoTriState.msoFalse);
    }

    private void InitializeOle(
        Shape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            format = shape.OLEFormat;
            oleObject = format.Object;
            if (oleObject is not IVisualTeXFormulaObject formula)
                throw new InvalidOperationException(
                    "The inserted PowerPoint object does not expose the VisualTeX native OLE interface.");
            FormulaOleInterop.Initialize(formula, metadata, emfPath, pngPath);
            ProbeOleStage("initialized", shape);
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static bool TryUpdateOle(
        Shape shape,
        FormulaMetadata metadata,
        string emfPath,
        string pngPath)
    {
        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            try { format = shape.OLEFormat; }
            catch { return false; }
            try { oleObject = format.Object; }
            catch { return false; }
            if (oleObject is not IVisualTeXFormulaObject formula) return false;
            FormulaOleInterop.Update(formula, metadata, emfPath, pngPath);
            return true;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static (Slide? Slide, Shape? Shape) FindFormula(
        Presentation presentation,
        string formulaId,
        string? preferredObjectId)
    {
        if (preferredObjectId?.StartsWith(ManagedReferencePrefix, StringComparison.Ordinal) == true)
            return FindManagedFormula(presentation, preferredObjectId, formulaId);
        if (TryParseMathTypeReference(preferredObjectId, out var mathTypeSlide, out var mathTypeShape))
            return FindMathType(presentation, mathTypeSlide, mathTypeShape);
        if (TryParseNativeOmmlReference(preferredObjectId, out var nativeSlide, out var nativeShape))
            return FindNativeOmml(presentation, nativeSlide, nativeShape);
        Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            var expectedName = $"VisualTeX_{formulaId}";

            // Shape names are indexed directly by PowerPoint. Resolve the exact
            // object name captured when the editor opened, then the canonical
            // VisualTeX_<FormulaId> name, before considering any shape inventory.
            // This makes ordinary edit/replace and each delayed OLE geometry repair
            // O(number of slides), rather than O(all shapes in the presentation).
            // The full metadata scan remains only for legacy/renamed shapes.
            for (var slideIndex = 1; slideIndex <= slides.Count; slideIndex++)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    shapes = slide.Shapes;
                    foreach (var name in new[] { preferredObjectId, expectedName })
                    {
                        if (string.IsNullOrWhiteSpace(name)) continue;
                        var isPreferredLookup = string.Equals(
                            name,
                            preferredObjectId,
                            StringComparison.Ordinal);
                        Shape? direct = null;
                        try
                        {
                            try { direct = shapes[name!]; }
                            catch { direct = null; }
                            if (direct is null) continue;
                            var metadata = ReadMetadata(direct);
                            if (!string.Equals(
                                    metadata?.FormulaId,
                                    formulaId,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            var foundSlide = slide;
                            var foundShape = direct;
                            slide = null;
                            direct = null;
                            return (foundSlide, foundShape);
                        }
                        finally { Release(direct); }
                    }
                }
                finally
                {
                    Release(shapes);
                    Release(slide);
                }
            }

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
                            if (!string.Equals(
                                    metadata?.FormulaId,
                                    formulaId,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            var foundSlide = slide;
                            var foundShape = shape;
                            slide = null;
                            shape = null;
                            return (foundSlide, foundShape);
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

    private static FormulaMetadata EnsureUniqueFormulaIdentity(
        Presentation presentation,
        Slide selectedSlide,
        Shape selectedShape,
        FormulaMetadata metadata)
    {
        var currentOwner = ShapeIdentityToken(selectedShape);
        var storedOwner = ReadIdentityOwner(selectedShape);
        if (!string.IsNullOrWhiteSpace(storedOwner))
        {
            if (string.Equals(storedOwner, currentOwner, StringComparison.OrdinalIgnoreCase))
            {
                return metadata;
            }

            var copied = CloneWithFormulaId(metadata, Guid.NewGuid().ToString());
            Configure(selectedShape, copied);
            return copied;
        }

        var expectedName = $"VisualTeX_{metadata.FormulaId}";
        var duplicateExists = false;
        Slides? slides = null;
        try
        {
            slides = presentation.Slides;
            for (var slideIndex = 1; slideIndex <= slides.Count && !duplicateExists; slideIndex++)
            {
                Slide? slide = null;
                Shapes? shapes = null;
                try
                {
                    slide = slides[slideIndex];
                    shapes = slide.Shapes;
                    for (var shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                    {
                        Shape? candidate = null;
                        try
                        {
                            candidate = shapes[shapeIndex];
                            if (slide.SlideID == selectedSlide.SlideID
                                && candidate.Id == selectedShape.Id)
                                continue;
                            var candidateMetadata = ReadMetadata(candidate);
                            if (!string.Equals(
                                    candidateMetadata?.FormulaId,
                                    metadata.FormulaId,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            duplicateExists = true;
                            break;
                        }
                        finally { Release(candidate); }
                    }
                }
                finally
                {
                    Release(shapes);
                    Release(slide);
                }
            }
        }
        finally { Release(slides); }

        if (!duplicateExists || string.Equals(selectedShape.Name, expectedName, StringComparison.Ordinal))
        {
            Configure(selectedShape, metadata);
            return metadata;
        }

        var rekeyed = CloneWithFormulaId(metadata, Guid.NewGuid().ToString());
        Configure(selectedShape, rekeyed);
        return rekeyed;
    }

    private static FormulaMetadata CloneWithFormulaId(FormulaMetadata metadata, string formulaId)
    {
        var clone = FormulaMetadataCodec.Decode(FormulaMetadataCodec.Encode(metadata))
            ?? throw new InvalidDataException("VisualTeX PowerPoint metadata could not be cloned.");
        clone.FormulaId = formulaId;
        clone.UpdatedAt = DateTimeOffset.UtcNow.ToString("O");
        return clone;
    }

    private static FormulaMetadata? ReadMetadata(Shape shape)
    {
        var overlay = ReadPictureMetadata(shape);
        if (overlay is not null) return overlay;
        if (shape.Type is not MsoShapeType.msoEmbeddedOLEObject
            and not MsoShapeType.msoLinkedOLEObject)
            return null;

        OLEFormat? format = null;
        object? oleObject = null;
        try
        {
            try { format = shape.OLEFormat; }
            catch { return null; }
            string? progId;
            try { progId = format.ProgID; }
            catch { return null; }
            if (!string.Equals(
                    progId,
                    FormulaOleContract.ProgId,
                    StringComparison.OrdinalIgnoreCase))
                return null;
            oleObject = GetRunningOleObject(format);
            return oleObject is IVisualTeXFormulaObject formula
                ? FormulaOleInterop.ReadMetadata(formula)
                : null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(oleObject);
            Release(format);
        }
    }

    private static object? GetRunningOleObject(OLEFormat format)
    {
        object? value = null;
        try { value = format.Object; } catch { }
        if (value is not null) return value;
        try { format.DoVerb(); } catch { }
        try { value = format.Object; } catch { value = null; }
        return value;
    }

    private static string? ReadIdentityOwner(Shape shape)
    {
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            try { return tags[IdentityOwnerTag]; }
            catch { return null; }
        }
        finally { Release(tags); }
    }

    private static string ShapeIdentityToken(Shape shape)
    {
        object? parent = null;
        try
        {
            parent = shape.Parent;
            var slideId = Convert.ToInt32(((dynamic)parent).SlideID);
            return $"{slideId}:{shape.Id}";
        }
        catch
        {
            return $"shape:{shape.Id}";
        }
        finally { Release(parent); }
    }

    private static FormulaMetadata? ReadPictureMetadata(Shape shape)
    {
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            string? encoded = null;
            try { encoded = tags[MetadataTag]; } catch { }
            FormulaMetadata? metadata = FormulaMetadataCodec.Decode(encoded);
            if (metadata is not null) return metadata;
            try { encoded = shape.AlternativeText; } catch { encoded = null; }
            return FormulaMetadataCodec.Decode(encoded);
        }
        finally { Release(tags); }
    }

    private static bool FormulaContentEquivalent(
        FormulaMetadata? original,
        FormulaMetadata current)
    {
        if (original is null) return false;
        return string.Equals(
                NormalizeFormulaText(original.Latex),
                NormalizeFormulaText(current.Latex),
                StringComparison.Ordinal)
            && string.Equals(
                original.DisplayMode,
                current.DisplayMode,
                StringComparison.Ordinal);
    }

    private static bool FormulaFontPreferencesChanged(
        FormulaMetadata? original,
        FormulaMetadata current)
    {
        static string Letter(string? value) =>
            string.IsNullOrWhiteSpace(value) ? "katex" : value!.Trim();
        static string Chinese(string? value) =>
            string.IsNullOrWhiteSpace(value) ? "system" : value!.Trim();
        return !string.Equals(
                Letter(original?.FormulaLetterFont),
                Letter(current.FormulaLetterFont),
                StringComparison.OrdinalIgnoreCase)
            || !string.Equals(
                Chinese(original?.FormulaChineseFont),
                Chinese(current.FormulaChineseFont),
                StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeFormulaText(string? value) =>
        (value ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace("\r", "\n")
            .Trim();

    private static bool IsNativeOle(Shape shape)
    {
        if (shape.Type is not MsoShapeType.msoEmbeddedOLEObject
            and not MsoShapeType.msoLinkedOLEObject)
            return false;

        OLEFormat? format = null;
        try
        {
            format = shape.OLEFormat;
            return string.Equals(
                format.ProgID,
                FormulaOleContract.ProgId,
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
        finally { Release(format); }
    }

    private static float InferPowerPointFormulaFontSize(
        float currentWidth,
        float currentHeight,
        FormulaMetadata? metadata)
    {
        if (!IsPowerPointGeometryTrusted(currentWidth, currentHeight, metadata))
            return FormulaFontSize.ResolveSemanticFontSize(metadata);
        return FormulaFontSize.InferOleFontSize(currentWidth, currentHeight, metadata);
    }

    private static bool IsPowerPointGeometryTrusted(
        float currentWidth,
        float currentHeight,
        FormulaMetadata? metadata)
    {
        if (metadata?.RenderWidthPx is not > 0 || metadata.RenderHeightPx is not > 0)
            return true;
        if (currentWidth <= 0 || currentHeight <= 0
            || float.IsNaN(currentWidth) || float.IsInfinity(currentWidth)
            || float.IsNaN(currentHeight) || float.IsInfinity(currentHeight))
            return false;

        var naturalWidth = Math.Max(0.01f, (float)metadata.RenderWidthPx.Value * 0.75f);
        var naturalHeight = Math.Max(0.01f, (float)metadata.RenderHeightPx.Value * 0.75f);
        var horizontalScale = currentWidth / naturalWidth;
        var verticalScale = currentHeight / naturalHeight;
        if (horizontalScale <= 0 || verticalScale <= 0
            || float.IsNaN(horizontalScale) || float.IsInfinity(horizontalScale)
            || float.IsNaN(verticalScale) || float.IsInfinity(verticalScale))
            return false;

        // PowerPoint formula objects are aspect-locked. A host box whose X/Y
        // scale differs materially from the metadata's natural geometry is not
        // a legitimate user resize; it is an OLE-container geometry glitch and
        // must never be propagated into the next SVG/OLE conversion or font size.
        var scaleRatio = horizontalScale / verticalScale;
        return scaleRatio >= 0.90f && scaleRatio <= 1.10f;
    }

    private static (float Width, float Height) ResolvePowerPointEditedSize(
        float currentWidth,
        float currentHeight,
        FormulaMetadata? originalMetadata,
        float newRenderWidth,
        float newRenderHeight,
        bool preserveCurrentGeometry)
    {
        const float maximumWidth = 600f;
        const float maximumHeight = 400f;
        if (preserveCurrentGeometry
            && currentWidth > 0
            && currentHeight > 0
            && !float.IsNaN(currentWidth)
            && !float.IsInfinity(currentWidth)
            && !float.IsNaN(currentHeight)
            && !float.IsInfinity(currentHeight))
        {
            if (IsPowerPointGeometryTrusted(currentWidth, currentHeight, originalMetadata))
                return (currentWidth, currentHeight);

            var semanticScale = originalMetadata is null
                ? 1f
                : FormulaFontSize.ResolveSemanticFontSize(originalMetadata)
                    / Math.Max(0.5f, FormulaFontSize.ResolveRenderFontSize(originalMetadata));
            if (float.IsNaN(semanticScale) || float.IsInfinity(semanticScale) || semanticScale <= 0)
                semanticScale = 1f;
            var width = Math.Max(1f, newRenderWidth * 0.75f * semanticScale);
            var height = Math.Max(1f, newRenderHeight * 0.75f * semanticScale);
            var fit = Math.Min(
                1f,
                Math.Min(
                    maximumWidth > 0 ? maximumWidth / width : 1f,
                    maximumHeight > 0 ? maximumHeight / height : 1f));
            if (!float.IsNaN(fit) && !float.IsInfinity(fit) && fit > 0 && fit < 1f)
            {
                width *= fit;
                height *= fit;
            }
            return (Math.Max(1f, width), Math.Max(1f, height));
        }

        return OfficeFormulaSizing.EditedSize(
            currentWidth,
            currentHeight,
            originalMetadata?.RenderWidthPx,
            originalMetadata?.RenderHeightPx,
            newRenderWidth,
            newRenderHeight,
            maximumWidth,
            maximumHeight,
            originalMetadata?.FontSizePt,
            originalMetadata?.RenderFontSizePt);
    }

    private static (float Width, float Height) ScaleCurrentShapeSize(
        float currentWidth,
        float currentHeight,
        float requestedScale,
        float maximumWidth,
        float maximumHeight)
    {
        var scale = requestedScale;
        if (float.IsNaN(scale) || float.IsInfinity(scale) || scale <= 0)
            scale = 1f;
        scale = Math.Max(0.1f, Math.Min(10f, scale));
        var width = Math.Max(1f, currentWidth) * scale;
        var height = Math.Max(1f, currentHeight) * scale;
        var fit = Math.Min(
            1f,
            Math.Min(
                maximumWidth > 0 ? maximumWidth / width : 1f,
                maximumHeight > 0 ? maximumHeight / height : 1f));
        if (!float.IsNaN(fit) && !float.IsInfinity(fit) && fit > 0 && fit < 1f)
        {
            width *= fit;
            height *= fit;
        }
        return (Math.Max(1f, width), Math.Max(1f, height));
    }

    private void ProbeOleStage(string stage, Shape shape)
    {
        _oleStageProbe?.Invoke(stage, shape);
    }

    private static void ApplyPictureSize(Shape shape, float width, float height)
    {
        shape.LockAspectRatio = MsoTriState.msoFalse;
        shape.Width = Math.Max(1f, width);
        shape.Height = Math.Max(1f, height);
        shape.LockAspectRatio = MsoTriState.msoTrue;
    }

    private static void ApplyOleSizeAndRefresh(Shape shape, float width, float height)
    {
        // Keep PowerPoint's outer Shape as the single geometry authority.
        // Calling IOleObject.SetExtent from inside POWERPNT.EXE causes the
        // VisualTeX LocalServer to send OnViewChange/OnDataChange back into the
        // same PowerPoint OLE container. PowerPoint can then reinterpret the
        // cached EMF extent as a new host extent and feed that enlarged value
        // back through the next SVG/OLE conversion, producing cumulative growth.
        // The server already derives a correct 96-DPI natural extent from the
        // EMF during Initialize/Update, so only size the host box here.
        ApplyPictureSize(shape, width, height);
    }

    private static void RestoreOlePosition(Shape shape, float left, float top)
    {
        // PowerPoint can reset a newly initialized OLE object's position while
        // it synchronizes the server extent and cached presentation. Position
        // is therefore the final geometry operation, after width and height.
        shape.Left = left;
        shape.Top = top;
    }

    private (IntPtr Hwnd, long Generation)? FreezePowerPointWindowRedraw(
        DocumentWindow? knownWindow = null)
    {
        DocumentWindow? acquiredWindow = null;
        try
        {
            var window = knownWindow;
            if (window is null)
            {
                acquiredWindow = _application.ActiveWindow;
                window = acquiredWindow;
            }
            if (window is null) return null;

            var hwnd = new IntPtr(window.HWND);
            if (hwnd == IntPtr.Zero) return null;
            var generation = ++_nextWindowRedrawFreezeGeneration;
            _windowRedrawFreezeGenerations[hwnd] = generation;
            SendMessage(hwnd, WmSetRedraw, IntPtr.Zero, IntPtr.Zero);
            _windowRedrawProbe?.Invoke("suspended", hwnd);
            return (hwnd, generation);
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(acquiredWindow);
        }
    }

    private void FinishPowerPointWindowRedrawFreeze(
        (IntPtr Hwnd, long Generation)? redrawFreeze,
        bool deferUntilOleSettled)
    {
        if (redrawFreeze is null) return;
        var freeze = redrawFreeze.Value;

        void RestoreRedraw()
        {
            if (!_windowRedrawFreezeGenerations.TryGetValue(freeze.Hwnd, out var currentGeneration)
                || currentGeneration != freeze.Generation)
                return;
            _windowRedrawFreezeGenerations.Remove(freeze.Hwnd);
            try
            {
                SendMessage(freeze.Hwnd, WmSetRedraw, new IntPtr(1), IntPtr.Zero);
                RedrawWindow(
                    freeze.Hwnd,
                    IntPtr.Zero,
                    IntPtr.Zero,
                    RdwInvalidate | RdwErase | RdwFrame | RdwAllChildren | RdwUpdateNow);
            }
            finally
            {
                _windowRedrawProbe?.Invoke("restored", freeze.Hwnd);
            }
        }

        if (deferUntilOleSettled && _postDelayedToOfficeUi is not null)
        {
            // Keep redraw suppressed only through the final short OLE geometry
            // correction window. PowerPoint's late cache replay happens during
            // the first few UI turns; a roughly 300 ms window masks it without
            // making OLE conversion feel almost a second slower than SVG/OMML.
            _postDelayedToOfficeUi(RestoreRedraw, 300);
            return;
        }
        if (deferUntilOleSettled && _postToOfficeUi is not null)
        {
            _postToOfficeUi(() => _postToOfficeUi(RestoreRedraw));
            return;
        }
        RestoreRedraw();
    }

    private void ScheduleOleGeometryRestore(
        string documentId,
        string formulaId,
        string objectId,
        float width,
        float height,
        float left,
        float top)
    {
        var post = _postToOfficeUi;
        if (post is null) return;
        var delayedPost = _postDelayedToOfficeUi;
        var generationKey = documentId + "\n" + objectId;
        var generation = _oleGeometryRestoreGenerations.TryGetValue(generationKey, out var previousGeneration)
            ? previousGeneration + 1
            : 1;
        _oleGeometryRestoreGenerations[generationKey] = generation;

        bool IsCurrentGeneration() =>
            _oleGeometryRestoreGenerations.TryGetValue(generationKey, out var currentGeneration)
            && currentGeneration == generation;

        void CompleteGeneration()
        {
            if (IsCurrentGeneration())
                _oleGeometryRestoreGenerations.Remove(generationKey);
        }

        void Restore()
        {
            if (!IsCurrentGeneration()) return;
            Presentation? presentation = null;
            Slide? slide = null;
            Shape? shape = null;
            try
            {
                presentation = _application.ActivePresentation;
                if (presentation is null
                    || !string.Equals(
                        DocumentIdentity(presentation),
                        documentId,
                        StringComparison.OrdinalIgnoreCase))
                    return;
                (slide, shape) = FindFormula(presentation, formulaId, objectId);
                if (shape is null || !IsNativeOle(shape)) return;
                ApplyOleSizeAndRefresh(shape, width, height);
                RestoreOlePosition(shape, left, top);
            }
            catch
            {
                // Geometry repair is best-effort and must never destabilize the
                // PowerPoint UI after the original conversion has completed.
            }
            finally
            {
                Release(shape);
                Release(slide);
                Release(presentation);
            }
        }

        // PowerPoint performs one or more OLE presentation/layout passes only
        // after the COM/Ribbon write callback unwinds. Two immediate BeginInvoke
        // callbacks are not sufficient on high-DPI Office: the container can
        // replay the cached presentation hundreds of milliseconds later and
        // replace the host box with a DPI-distorted extent. Repair on the next
        // two UI messages and then re-check across the short post-callback window.
        post(() =>
        {
            Restore();
            post(() =>
            {
                Restore();
                if (delayedPost is null)
                    CompleteGeneration();
            });
        });
        if (delayedPost is not null)
        {
            delayedPost(Restore, 60);
            delayedPost(Restore, 140);
            delayedPost(() =>
            {
                try { Restore(); }
                finally { CompleteGeneration(); }
            }, 320);
        }
    }

    private static void Configure(Shape shape, FormulaMetadata metadata)
    {
        try { shape.LockAspectRatio = MsoTriState.msoTrue; } catch { }
        shape.Name = $"VisualTeX_{metadata.FormulaId}";
        var encoded = FormulaMetadataCodec.Encode(metadata);
        Tags? tags = null;
        try
        {
            tags = shape.Tags;
            tags.Add(FormulaIdTag, metadata.FormulaId);
            tags.Add(MetadataTag, encoded);
            tags.Add(IdentityOwnerTag, ShapeIdentityToken(shape));
        }
        finally { Release(tags); }
        try
        {
            var alternativeText = shape.AlternativeText;
            if (string.IsNullOrWhiteSpace(alternativeText)
                || FormulaMetadataCodec.Decode(alternativeText) is not null)
                shape.AlternativeText = encoded;
        }
        catch { }
    }

    private static void TryApplyRotation(Shape shape, float rotation)
    {
        if (Math.Abs(rotation) < 0.01f) return;
        try { shape.Rotation = rotation; } catch { }
    }

    private static void MoveToZOrder(Shape shape, int target)
    {
        var current = shape.ZOrderPosition;
        var requiredMoves = Math.Max(0, current - target);
        // A fixed 512-attempt ceiling silently failed on dense technical slides.
        // Bound work by the actual number of required moves instead, with a small
        // allowance for PowerPoint regrouping/reindexing shapes during replacement.
        var maxAttempts = requiredMoves + 8;
        for (var attempts = 0;
             attempts < maxAttempts && shape.ZOrderPosition > target;
             attempts++)
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

    private readonly struct TextInsertionTarget
    {
        internal TextInsertionTarget(int slideId, int shapeId, int start, int length, string sourceFingerprint)
        {
            SlideId = slideId;
            ShapeId = shapeId;
            Start = start;
            Length = length;
            SourceFingerprint = sourceFingerprint;
        }

        internal int SlideId { get; }
        internal int ShapeId { get; }
        internal int Start { get; }
        internal int Length { get; }
        internal string SourceFingerprint { get; }
    }

    private static string TextInsertionReference(
        int slideId,
        int shapeId,
        int start,
        int length,
        string sourceFingerprint) =>
        $"{TextInsertionReferencePrefix}{slideId}:{shapeId}:{start}:{length}:{sourceFingerprint}";

    internal static bool IsTextInsertionReference(string? value) =>
        !string.IsNullOrWhiteSpace(value)
        && value!.StartsWith(TextInsertionReferencePrefix, StringComparison.Ordinal);

    private static bool TryParseTextInsertionReference(
        string? value,
        out TextInsertionTarget target)
    {
        target = default;
        if (!IsTextInsertionReference(value)) return false;
        var payload = value!.Substring(TextInsertionReferencePrefix.Length);
        var parts = payload.Split(':');
        if (parts.Length != 5
            || !int.TryParse(parts[0], out var slideId)
            || !int.TryParse(parts[1], out var shapeId)
            || !int.TryParse(parts[2], out var start)
            || !int.TryParse(parts[3], out var length)
            || slideId <= 0
            || shapeId <= 0
            || start < 1
            || length < 0
            || parts[4].Length != 64
            || parts[4].Any(character => !(character >= '0' && character <= '9' || character >= 'A' && character <= 'F')))
            throw new InvalidDataException("The captured PowerPoint text target has no valid source proof; reopen the editor from the current text selection.");
        target = new TextInsertionTarget(slideId, shapeId, start, length, parts[4]);
        return true;
    }

    private static (Slide Slide, Shape Shape) ResolveTextInsertionTarget(
        Presentation presentation,
        TextInsertionTarget target)
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
                    if (slide.SlideID != target.SlideId) continue;
                    shapes = slide.Shapes;
                    for (var shapeIndex = 1; shapeIndex <= shapes.Count; shapeIndex++)
                    {
                        Shape? shape = null;
                        try
                        {
                            shape = shapes[shapeIndex];
                            if (shape.Id != target.ShapeId) continue;
                            if (shape.HasTextFrame != MsoTriState.msoTrue)
                                throw new InvalidOperationException(
                                    "The PowerPoint text box captured for OMML insertion no longer contains an editable text frame.");
                            var foundSlide = slide;
                            var foundShape = shape;
                            slide = null;
                            shape = null;
                            return (foundSlide, foundShape);
                        }
                        finally { Release(shape); }
                    }
                    throw new InvalidOperationException(
                        "The PowerPoint text box captured for OMML insertion no longer exists.");
                }
                finally
                {
                    Release(shapes);
                    Release(slide);
                }
            }
            throw new InvalidOperationException(
                "The PowerPoint slide captured for OMML insertion no longer exists.");
        }
        finally { Release(slides); }
    }

    // New formulas go to the slide center. When a shape is already centered
    // there (usually the previous new formula), step down and right like
    // PowerPoint's own repeated paste, so formulas do not pile up.
    internal static (float Left, float Top) NewFormulaPosition(
        Presentation presentation,
        Slide slide,
        float width,
        float height)
    {
        const float step = 24f;
        const float tolerance = 4f;
        var slideWidth = presentation.PageSetup.SlideWidth;
        var slideHeight = presentation.PageSetup.SlideHeight;
        var centerX = slideWidth / 2f;
        var centerY = slideHeight / 2f;
        var centers = new List<(float X, float Y)>();
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
                    centers.Add((shape.Left + shape.Width / 2f, shape.Top + shape.Height / 2f));
                }
                catch { }
                finally { Release(shape); }
            }
        }
        catch { }
        finally { Release(shapes); }

        for (var attempt = 0; attempt < 20; attempt++)
        {
            var x = centerX + attempt * step;
            var y = centerY + attempt * step;
            if (x + width / 2f > slideWidth || y + height / 2f > slideHeight) break;
            if (centers.Any(c => Math.Abs(c.X - x) <= tolerance && Math.Abs(c.Y - y) <= tolerance))
                continue;
            centerX = x;
            centerY = y;
            break;
        }
        return (
            Math.Max(0f, centerX - width / 2f),
            Math.Max(0f, centerY - height / 2f));
    }

    private static Slide ResolveTargetSlide(
        Presentation presentation,
        string? sourceObjectId,
        View view)
    {
        var hasSlide = TryParseSlideReference(sourceObjectId, out var slideId);
        if (!hasSlide && TryParseTextInsertionReference(sourceObjectId, out var textTarget))
        {
            slideId = textTarget.SlideId;
            hasSlide = true;
        }
        if (hasSlide)
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
        if (string.IsNullOrWhiteSpace(value)) return false;
        var reference = value!;
        if (!reference.StartsWith(SlideReferencePrefix, StringComparison.Ordinal))
            return false;
        var payload = reference.Substring(SlideReferencePrefix.Length);
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
        // Office may return the same RCW to the host and to this service.
        // FinalReleaseComObject would invalidate every shared reference in the
        // add-in AppDomain, so release only the reference acquired here.
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

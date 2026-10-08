using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using View = Microsoft.Office.Interop.PowerPoint.View;

namespace VisualTeX.PowerPointVsto;

public sealed partial class PowerPointFormulaService
{
    private const string MathTypeReferencePrefix = "visualtex-ppt-mathtype:";
    private MathTypeSource? _mathTypeSource;

    private sealed class MathTypeSource
    {
        internal string DocumentId = string.Empty;
        internal string ObjectId = string.Empty;
        internal string FormulaId = string.Empty;
        internal byte[] CompoundFile = Array.Empty<byte>();
        internal float NaturalWidth;
        internal float NaturalHeight;
    }

    public sealed class PreparedMathType : IDisposable
    {
        internal byte[] CompoundFile = Array.Empty<byte>();
        internal byte[]? SourceCompoundFile;
        internal string EmfPath = string.Empty;
        internal bool OwnsEmf;
        internal float Width;
        internal float Height;
        internal float SourceWidth;
        internal float SourceHeight;
        public void Dispose()
        {
            if (OwnsEmf) { try { File.Delete(EmfPath); } catch { } }
        }
    }

    private static bool IsMathTypeShape(Shape shape)
    {
        if (shape.Type != MsoShapeType.msoEmbeddedOLEObject) return false;
        OLEFormat? format = null;
        try
        {
            format = shape.OLEFormat;
            return string.Equals(format.ProgID, MathTypeOleInterop.CanonicalProgId, StringComparison.OrdinalIgnoreCase)
                || MathTypeOleInterop.TryResolveCapabilities(format.ProgID, out _);
        }
        catch { return false; }
        finally { Release(format); }
    }

    private static string MathTypeReference(Slide slide, Shape shape) =>
        $"{MathTypeReferencePrefix}{slide.SlideID}:{shape.Id}";

    private static bool TryParseMathTypeReference(string? reference, out int slideId, out int shapeId)
    {
        slideId = shapeId = 0;
        if (reference is null || !reference.StartsWith(MathTypeReferencePrefix, StringComparison.Ordinal)) return false;
        var parts = reference.Substring(MathTypeReferencePrefix.Length).Split(':');
        return parts.Length == 2 && int.TryParse(parts[0], out slideId) && slideId > 0
            && int.TryParse(parts[1], out shapeId) && shapeId > 0;
    }

    private static (Slide? Slide, Shape? Shape) FindMathType(Presentation presentation, int slideId, int shapeId)
    {
        Slides? slides = null;
        Slide? slide = null;
        Shapes? shapes = null;
        try
        {
            slides = presentation.Slides;
            slide = slides.FindBySlideID(slideId);
            shapes = slide.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                if (shape.Id == shapeId && IsMathTypeShape(shape))
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

    public OfficeSelection LoadMathTypeSelection(OfficeSelection selection)
    {
        if (selection.ObjectMode != FormulaOleContract.MathTypeOleMode) return selection;
        if (!TryParseMathTypeReference(selection.ObjectId, out var slideId, out var shapeId))
            throw new InvalidOperationException("The MathType selection has no valid slide and shape identity.");
        Presentation? presentation = null;
        Slide? slide = null;
        Shape? shape = null;
        try
        {
            presentation = _application.ActivePresentation;
            EnsureSourceDocument(presentation, selection.DocumentId);
            (slide, shape) = FindMathType(presentation, slideId, shapeId);
            if (shape is null) throw new InvalidOperationException("The selected MathType equation no longer exists.");
            var equation = PowerPointMathTypePackage.Read(presentation, slideId, shapeId);
            var native = MathTypeOleStorage.ReadEquationNative(equation.CompoundFile);
            var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
            var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
            if (string.IsNullOrWhiteSpace(latex)) throw new InvalidDataException("The MathType equation is empty or unsupported.");
            var size = MathTypeMtefCodec.ReadEquationNativeFullFontSize(native);
            var now = DateTimeOffset.UtcNow.ToString("O");
            var metadata = new FormulaMetadata
            {
                FormulaId = Guid.NewGuid().ToString("D"),
                Title = "MathType Formula",
                Latex = latex,
                Lines = new List<FormulaLine> { new() { Id = Guid.NewGuid().ToString("D"), Latex = latex } },
                CodeFormat = "raw",
                DisplayMode = "block",
                Numbered = false,
                FontSizePt = size,
                RenderFontSizePt = size,
                RenderWidthPx = equation.WidthPt > 0 ? equation.WidthPt / 0.75 : shape.Width / 0.75,
                RenderHeightPx = equation.HeightPt > 0 ? equation.HeightPt / 0.75 : shape.Height / 0.75,
                CreatedWithVersion = "1.2.8",
                UpdatedWithVersion = "1.2.8",
                CreatedAt = now,
                UpdatedAt = now,
            };
            metadata.Validate();
            _mathTypeSource = new MathTypeSource
            {
                DocumentId = selection.DocumentId!, ObjectId = selection.ObjectId!, FormulaId = metadata.FormulaId,
                CompoundFile = equation.CompoundFile, NaturalWidth = equation.WidthPt, NaturalHeight = equation.HeightPt,
            };
            return new OfficeSelection
            {
                Host = "powerpoint", DocumentId = selection.DocumentId, ObjectId = selection.ObjectId,
                ReadOnly = selection.ReadOnly, ObjectMode = FormulaOleContract.MathTypeOleMode,
                FormulaId = metadata.FormulaId, Metadata = metadata,
            };
        }
        finally { Release(shape); Release(slide); Release(presentation); }
    }

    // No Office COM calls: native preview runs in the existing isolated renderer.
    public PreparedMathType PrepareMathType(OfficeSessionDocument session)
    {
        var mathMl = session.ExportResult?.MathMl;
        if (string.IsNullOrWhiteSpace(mathMl)) throw new InvalidDataException("MathType requires MathML source.");
        var result = new PreparedMathType();
        try
        {
            byte[] mtef;
            if (session.Mode == "edit" && TryParseMathTypeReference(session.SourceObjectId, out _, out _))
            {
                var source = _mathTypeSource;
                if (source is null || source.DocumentId != session.SourceDocumentId
                    || source.ObjectId != session.SourceObjectId || source.FormulaId != session.FormulaId)
                    throw new InvalidOperationException("Reopen the MathType equation before saving it.");
                var rewritten = MathTypeOleStorage.RewriteMathTypeCompoundFile(
                    source.CompoundFile, mathMl!, inline: false, fullFontSizePt: session.FontSizePt);
                result.CompoundFile = rewritten.CompoundFile;
                result.SourceCompoundFile = source.CompoundFile;
                result.SourceWidth = source.NaturalWidth;
                result.SourceHeight = source.NaturalHeight;
                mtef = rewritten.Mtef;
            }
            else
            {
                var generated = MathTypeMtefCodec.CreateEquationNativeAtFontSize(mathMl!, inline: false, session.FontSizePt);
                result.CompoundFile = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
                mtef = generated.Mtef;
            }
            if (MathTypeMtefCodec.SemanticSignature(mathMl!) !=
                MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(result.CompoundFile)))
                throw new InvalidDataException("The equation cannot be saved to MathType without changing its content.");
            var previewsToRender = new Dictionary<string, byte[]> { ["target"] = mtef };
            if (result.SourceCompoundFile is not null)
                previewsToRender["source"] = MathTypeMtefCodec.ExtractNativeMtef(
                    MathTypeOleStorage.ReadEquationNative(result.SourceCompoundFile));
            MathTypeNativePreviewRenderer.TryRenderBatch(previewsToRender, Path.GetTempPath(), out var previews);
            try
            {
                if (!previews.ContainsKey("target")
                    || (result.SourceCompoundFile is not null && !previews.ContainsKey("source")))
                    throw new InvalidOperationException("MathType's native preview could not be rendered. The original equation has not been changed.");
                if (previews.TryGetValue("source", out var sourcePreview))
                {
                    result.SourceWidth = sourcePreview.WidthPt;
                    result.SourceHeight = sourcePreview.HeightPt;
                }
                if (previews.TryGetValue("target", out var preview))
                {
                    result.Width = preview.WidthPt;
                    result.Height = preview.HeightPt;
                    result.EmfPath = PowerPointMathTypePreview.FromNativeWmf(preview.WmfPath, result.Width, result.Height);
                    result.OwnsEmf = true;
                }
            }
            finally { foreach (var preview in previews.Values) preview.Dispose(); }
            result.CompoundFile = MathTypeOleStorage.AddEnhancedMetafilePresentationCache(
                result.CompoundFile, result.EmfPath, replaceExisting: true);
            return result;
        }
        catch { result.Dispose(); throw; }
    }

    public OfficeObjectResult ApplyMathType(OfficeSessionDocument session, PreparedMathType prepared)
    {
        Presentation? presentation = null;
        DocumentWindow? window = null;
        View? view = null;
        Slide? slide = null;
        Shape? source = null;
        Shape? replacement = null;
        Shapes? shapes = null;
        ShapeRange? pasted = null;
        (IntPtr Hwnd, long Generation)? redraw = null;
        ShapeReferenceTransfer? references = null;
        var committed = false;
        try
        {
            EnsureNotSlideShow();
            presentation = _application.ActivePresentation;
            EnsureWritable(presentation);
            EnsureSourceDocument(presentation, session.SourceDocumentId);
            window = _application.ActiveWindow;
            view = window.View;
            if (session.Mode == "edit")
            {
                (slide, source) = FindFormula(presentation, session.FormulaId, session.SourceObjectId);
                if (slide is null || source is null) throw new InvalidOperationException("The equation no longer exists.");
                if (Math.Abs(source.Rotation) > 0.01f
                    || source.HorizontalFlip == MsoTriState.msoTrue || source.VerticalFlip == MsoTriState.msoTrue)
                    throw new InvalidOperationException("PowerPoint does not support rotated or flipped MathType OLE objects. Reset the rotation and flip before converting.");
                if (prepared.SourceCompoundFile is not null)
                {
                    var current = PowerPointMathTypePackage.Read(presentation, slide.SlideID, source.Id);
                    if (!MathTypeOleStorage.ReadEquationNative(current.CompoundFile).SequenceEqual(
                            MathTypeOleStorage.ReadEquationNative(prepared.SourceCompoundFile)))
                        throw new InvalidOperationException("MathType changed this equation while VisualTeX was editing it. Reopen it before saving.");
                }
            }
            else slide = ResolveTargetSlide(presentation, session.SourceObjectId, view);

            var width = prepared.Width;
            var height = prepared.Height;
            if (source is not null && prepared.SourceWidth > 0 && prepared.SourceHeight > 0)
            {
                width *= source.Width / prepared.SourceWidth;
                height *= source.Height / prepared.SourceHeight;
            }
            else if (source is not null)
            {
                var metadata = session.OriginalMetadata;
                var size = OfficeFormulaSizing.EditedSize(source.Width, source.Height,
                    metadata?.RenderWidthPx, metadata?.RenderHeightPx, width / 0.75f, height / 0.75f,
                    preserveIndependentAxisScale: true);
                width = size.Width;
                height = size.Height;
            }
            var (left, top) = source is not null
                ? (source.Left, source.Top)
                : NewFormulaPosition(presentation, slide, width, height);
            StartNewUndoEntry();
            redraw = FreezePowerPointWindowRedraw(window);
            shapes = slide.Shapes;
            using (var clipboard = MathTypeOleStorage.BeginStandaloneClipboardTransaction(
                       prepared.CompoundFile, prepared.EmfPath, prepared.Width, prepared.Height))
            {
                pasted = shapes.PasteSpecial(PpPasteDataType.ppPasteOLEObject);
                if (pasted.Count == 0) throw new InvalidDataException("PowerPoint did not paste a MathType object.");
                replacement = pasted[1];
                // PowerPoint can return repeated references to the same OLE
                // shape for the source and presentation clipboard formats.
                // Validate object identity, not the number of range entries.
                for (var i = 2; i <= pasted.Count; i++)
                {
                    var item = pasted[i];
                    try
                    {
                        if (item.Id != replacement.Id)
                            throw new InvalidDataException("PowerPoint pasted more than one distinct object.");
                    }
                    finally { Release(item); }
                }
                if (!IsMathTypeShape(replacement)) throw new InvalidDataException("PowerPoint did not create a MathType OLE object.");
            }
            replacement.LockAspectRatio = MsoTriState.msoFalse;
            replacement.Width = width;
            replacement.Height = height;
            replacement.Left = left;
            replacement.Top = top;
            if (source is not null)
            {
                replacement.LockAspectRatio = source.LockAspectRatio;
                replacement.AlternativeText = source.AlternativeText;
                replacement.Title = source.Title;
                replacement.Visible = source.Visible;
                MoveToZOrder(replacement, source.ZOrderPosition + 1);
                CopyMathTypeShapeTags(source, replacement);
                CopyMathTypeShapeActions(source, replacement);
            }
            else replacement.LockAspectRatio = MsoTriState.msoTrue;

            // Verify serialized content before deleting the user's original.
            var saved = PowerPointMathTypePackage.Read(presentation, slide.SlideID, replacement.Id);
            if (!MathTypeOleStorage.ReadEquationNative(saved.CompoundFile).SequenceEqual(
                    MathTypeOleStorage.ReadEquationNative(prepared.CompoundFile)))
                throw new InvalidDataException("PowerPoint did not preserve the MathType equation storage.");
            if (source is not null)
            {
                references = TransferMathTypeShapeReferences(slide, source, replacement);
                replacement.Name = source.Name;
                source.Delete();
                references.Commit();
            }
            else replacement.Name = $"MathType_{session.FormulaId}";
            committed = true;
            try { replacement.Select(MsoTriState.msoTrue); } catch { }
            return Result(session, presentation, MathTypeReference(slide, replacement));
        }
        catch
        {
            references?.Dispose();
            references = null;
            if (!committed)
            {
                try { pasted?.Delete(); } catch { TryDelete(replacement); }
            }
            throw;
        }
        finally
        {
            references?.Dispose();
            FinishPowerPointWindowRedrawFreeze(redraw, false);
            StartNewUndoEntry();
            Release(pasted); Release(shapes); Release(replacement); Release(source);
            Release(slide); Release(view); Release(window); Release(presentation);
        }
    }

    private static void CopyMathTypeShapeTags(Shape source, Shape target)
    {
        Tags? sourceTags = null;
        Tags? targetTags = null;
        try
        {
            sourceTags = source.Tags;
            targetTags = target.Tags;
            for (var i = 1; i <= sourceTags.Count; i++)
            {
                var name = sourceTags.Name(i);
                if (name.Equals(MetadataTag, StringComparison.OrdinalIgnoreCase)
                    || name.Equals(FormulaIdTag, StringComparison.OrdinalIgnoreCase)
                    || name.Equals(IdentityOwnerTag, StringComparison.OrdinalIgnoreCase)) continue;
                targetTags.Add(name, sourceTags.Value(i));
            }
        }
        finally { Release(targetTags); Release(sourceTags); }
    }

    private void EnsureMathTypeSourceUnchanged(OfficeSessionDocument session, Presentation presentation, Slide slide, Shape shape)
    {
        // All non-VisualTeX sources are validated at this shared replacement boundary.
        EnsureNativeOmmlSourceUnchanged(session, presentation);
        if (!TryParseMathTypeReference(session.SourceObjectId, out _, out _)) return;
        var source = _mathTypeSource;
        if (source is null || source.DocumentId != session.SourceDocumentId
            || source.ObjectId != session.SourceObjectId || source.FormulaId != session.FormulaId)
            throw new InvalidOperationException("Reopen the MathType equation before saving it.");
        var current = PowerPointMathTypePackage.Read(presentation, slide.SlideID, shape.Id);
        if (!MathTypeOleStorage.ReadEquationNative(current.CompoundFile).SequenceEqual(
                MathTypeOleStorage.ReadEquationNative(source.CompoundFile)))
            throw new InvalidOperationException("MathType changed this equation while VisualTeX was editing it. Reopen it before saving.");
    }

    private static void DeleteReplacedShape(Slide slide, Shape source, Shape target)
    {
        target.Title = source.Title;
        var alternativeText = source.AlternativeText;
        if (!string.IsNullOrWhiteSpace(alternativeText)
            && FormulaMetadataCodec.Decode(alternativeText) is null)
            target.AlternativeText = alternativeText;
        CopyMathTypeShapeTags(source, target);
        CopyMathTypeShapeActions(source, target);
        using var references = TransferMathTypeShapeReferences(slide, source, target);
        source.Delete();
        references.Commit();
    }

    private static void CopyMathTypeShapeActions(Shape source, Shape target)
    {
        ActionSettings? sourceActions = null;
        ActionSettings? targetActions = null;
        try
        {
            sourceActions = source.ActionSettings;
            targetActions = target.ActionSettings;
            foreach (var kind in new[] { PpMouseActivation.ppMouseClick, PpMouseActivation.ppMouseOver })
            {
                ActionSetting? from = null;
                ActionSetting? to = null;
                Hyperlink? sourceLink = null;
                Hyperlink? targetLink = null;
                try
                {
                    from = sourceActions[kind];
                    if (from.Action == PpActionType.ppActionNone) continue;
                    to = targetActions[kind];
                    to.Action = from.Action;
                    to.AnimateAction = from.AnimateAction;
                    if (from.Action is PpActionType.ppActionRunMacro or PpActionType.ppActionRunProgram)
                        to.Run = from.Run;
                    else if (from.Action == PpActionType.ppActionNamedSlideShow)
                        to.SlideShowName = from.SlideShowName;
                    else if (from.Action == PpActionType.ppActionHyperlink)
                    {
                        sourceLink = from.Hyperlink;
                        targetLink = to.Hyperlink;
                        targetLink.Address = sourceLink.Address;
                        targetLink.SubAddress = sourceLink.SubAddress;
                    }
                }
                finally { Release(targetLink); Release(sourceLink); Release(to); Release(from); }
            }
        }
        finally { Release(targetActions); Release(sourceActions); }
    }

    private sealed class ShapeReferenceTransfer : IDisposable
    {
        internal readonly Stack<Action> Rollback = new();
        internal readonly List<object> Owned = new();
        internal void Commit() => Rollback.Clear();
        public void Dispose()
        {
            while (Rollback.Count > 0) { try { Rollback.Pop()(); } catch { } }
            for (var i = Owned.Count - 1; i >= 0; i--) Release(Owned[i]);
            Owned.Clear();
        }
    }

    private static ShapeReferenceTransfer TransferMathTypeShapeReferences(Slide slide, Shape source, Shape target)
    {
        // Retarget existing effects in place, preserving timeline order/timing.
        TimeLine? timeline = null;
        Sequences? interactive = null;
        var transfer = new ShapeReferenceTransfer();
        try
        {
            timeline = slide.TimeLine;
            void Retarget(Sequence sequence)
            {
                try
                {
                    for (var i = 1; i <= sequence.Count; i++)
                    {
                        var effect = sequence[i];
                        transfer.Owned.Add(effect);
                        var animated = effect.Shape;
                        if (animated.Id == source.Id)
                        {
                            effect.Shape = target;
                            transfer.Rollback.Push(() => effect.Shape = source);
                        }
                        var timing = effect.Timing;
                        transfer.Owned.Add(timing);
                        Shape? trigger = null;
                        try { trigger = timing.TriggerShape; } catch { }
                        if (trigger?.Id == source.Id)
                        {
                            timing.TriggerShape = target;
                            transfer.Rollback.Push(() => timing.TriggerShape = source);
                        }
                        Release(trigger); Release(animated);
                        // Keep changed RCWs alive until all retargets succeed.
                    }
                }
                finally { Release(sequence); }
            }
            Retarget(timeline.MainSequence);
            interactive = timeline.InteractiveSequences;
            for (var i = 1; i <= interactive.Count; i++) Retarget(interactive[i]);
            Shapes? shapes = null;
            try
            {
                shapes = slide.Shapes;
                for (var i = 1; i <= shapes.Count; i++)
                {
                    var connectorShape = shapes[i];
                    try
                    {
                        if (connectorShape.Connector != MsoTriState.msoTrue) continue;
                        var connector = connectorShape.ConnectorFormat;
                        transfer.Owned.Add(connector);
                        if (connector.BeginConnected == MsoTriState.msoTrue)
                        {
                            var connected = connector.BeginConnectedShape;
                            try
                            {
                                if (connected.Id == source.Id)
                                {
                                    var site = connector.BeginConnectionSite;
                                    connector.BeginConnect(target, site);
                                    transfer.Rollback.Push(() => connector.BeginConnect(source, site));
                                }
                            }
                            finally { Release(connected); }
                        }
                        if (connector.EndConnected == MsoTriState.msoTrue)
                        {
                            var connected = connector.EndConnectedShape;
                            try
                            {
                                if (connected.Id == source.Id)
                                {
                                    var site = connector.EndConnectionSite;
                                    connector.EndConnect(target, site);
                                    transfer.Rollback.Push(() => connector.EndConnect(source, site));
                                }
                            }
                            finally { Release(connected); }
                        }
                    }
                    finally { Release(connectorShape); }
                }
            }
            finally { Release(shapes); }
            return transfer;
        }
        catch
        {
            transfer.Dispose();
            throw;
        }
        finally { Release(interactive); Release(timeline); }
    }
}

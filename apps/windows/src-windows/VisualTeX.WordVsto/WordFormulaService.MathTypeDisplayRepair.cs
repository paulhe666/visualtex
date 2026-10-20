using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed class WordMathTypeDisplayRepairTarget
{
    internal int Start { get; set; }
    internal int End { get; set; }
    internal string CompoundFileHash { get; set; } = string.Empty;
    // MathType's own rendering of the unchanged Equation Native.
    internal byte[] PreviewWmf { get; set; } = Array.Empty<byte>();
    internal float NativeWidthPt { get; set; }
    internal float NativeHeightPt { get; set; }
    internal int NativeWordPosition { get; set; }
}

internal sealed class WordMathTypeDisplayRepairReport
{
    internal string DocumentId { get; set; } = string.Empty;
    internal int MathTypeCount { get; set; }
    internal int UnreadableCount { get; set; }
    // Equations MathType's own renderer rejects; they are left as they are.
    internal int UnrenderableCount { get; set; }
    internal List<WordMathTypeDisplayRepairTarget> Targets { get; } = new();
}

internal sealed partial class WordFormulaService
{
    // A MathType equation is shown through the picture stored next to its
    // Equation Native data. Pictures drawn while MathType's fonts were missing
    // (older VisualTeX on a PC without MathType) keep placeholder and substitute
    // glyphs. The repair has MathType's own renderer draw the unchanged Equation
    // Native again and replaces only that picture: the Equation.DSMT4 storage stays
    // byte-identical, as do numbering, references and the equation's place in the
    // text. The new picture sits on MathType's baseline like a newly inserted one.
    internal WordMathTypeDisplayRepairReport ScanMathTypeDisplayRepairs()
    {
        var report = new WordMathTypeDisplayRepairReport();
        var candidates = new List<(WordMathTypeDisplayRepairTarget Target, byte[] Mtef, byte[] Picture)>();
        Document? document = null;
        InlineShapes? shapes = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            report.DocumentId = DocumentIdentity(document);
            shapes = document.InlineShapes;
            var shapeCount = shapes.Count;
            for (var index = 1; index <= shapeCount; index++)
            {
                InlineShape? shape = null;
                Range? range = null;
                try
                {
                    if (index % 10 == 1)
                        ReportOperationProgress(
                            OfficePluginLanguage.Text("正在读取 MathType 公式…", "Reading MathType equations…"),
                            index - 1, shapeCount);
                    shape = shapes[index];
                    if (shape.Type != WdInlineShapeType.wdInlineShapeEmbeddedOLEObject
                        || !MathTypeOleInterop.IsMathTypeOle(shape))
                        continue;
                    report.MathTypeCount++;
                    MathTypeWordOpenXml.Fragment fragment;
                    byte[] mtef;
                    try
                    {
                        fragment = MathTypeWordOpenXml.Read(shape);
                        mtef = ReadMathTypeMtefFromCompoundFile(fragment.CompoundFile);
                        EnsureMathTypeDisplayRepairReadable(fragment.CompoundFile);
                    }
                    catch (Exception error) when (error is InvalidDataException or OverflowException)
                    {
                        // A damaged Equation Native (e.g. an impossible MTEF extent)
                        // is left as it is; it never stops the scan.
                        report.UnreadableCount++;
                        continue;
                    }
                    range = shape.Range;
                    candidates.Add((new WordMathTypeDisplayRepairTarget
                    {
                        Start = range.Start,
                        End = range.End,
                        CompoundFileHash = Sha256(fragment.CompoundFile),
                    }, mtef, fragment.PreviewWmf));
                }
                catch (COMException)
                {
                    report.UnreadableCount++;
                }
                finally { Release(range); Release(shape); }
            }
        }
        finally { Release(shapes); Release(document); }
        if (candidates.Count == 0) return report;
        if (!MathTypeNativePreviewRenderer.IsAvailable)
            throw new InvalidOperationException(MathTypeNativePreviewRenderer.UnavailableMessage);

        ReportOperationProgress(
            OfficePluginLanguage.Text("正在用 MathType 排版核对公式图像…", "Checking equation pictures with MathType's layout…"),
            0, 0);
        var rendered = new Dictionary<string, MathTypeNativePreviewRenderer.Result>(StringComparer.Ordinal);
        try
        {
            RenderMathTypeDisplayPreviews(
                candidates.Select((item, index) => new KeyValuePair<string, byte[]>(
                    index.ToString(CultureInfo.InvariantCulture), item.Mtef)).ToList(),
                rendered);
            if (rendered.Count == 0)
                throw new InvalidOperationException(MathTypeNativePreviewRenderer.FailureMessage);
            for (var index = 0; index < candidates.Count; index++)
            {
                if (!rendered.TryGetValue(index.ToString(CultureInfo.InvariantCulture), out var preview))
                {
                    report.UnrenderableCount++;
                    continue;
                }
                var wmf = File.ReadAllBytes(preview.WmfPath);
                // Word stores MathType's pictures with window records added;
                // compare what the pictures draw, not their container bytes.
                if (MathTypePictureDrawing(candidates[index].Picture) == MathTypePictureDrawing(wmf))
                    continue;
                var target = candidates[index].Target;
                target.PreviewWmf = wmf;
                target.NativeWidthPt = preview.WidthPt;
                target.NativeHeightPt = preview.HeightPt;
                target.NativeWordPosition = preview.WordPosition;
                report.Targets.Add(target);
            }
        }
        finally
        {
            foreach (var preview in rendered.Values) preview.Dispose();
        }
        return report;
    }

    // Redraw only equations whose data VisualTeX reads in full. MathType also
    // draws damaged data (an MTEF cut short draws as a fragment); the picture
    // stored with such an equation is then the better record of it.
    private static void EnsureMathTypeDisplayRepairReadable(byte[] compoundFile)
    {
        try
        {
            _ = MathTypeMtefCodec.ReadEquationNativeMathMl(
                MathTypeOleStorage.ReadEquationNative(compoundFile));
        }
        catch (Exception error) when (error is not OutOfMemoryException)
        {
            throw new InvalidDataException("VisualTeX cannot read this MathType equation in full.", error);
        }
    }

    // One equation MathType's renderer rejects must not block the others: a
    // failed batch is split until each rejected equation stands alone. Missing
    // fonts stop the scan, since every picture would then be drawn wrong.
    private static void RenderMathTypeDisplayPreviews(
        IReadOnlyList<KeyValuePair<string, byte[]>> items,
        Dictionary<string, MathTypeNativePreviewRenderer.Result> rendered)
    {
        var batch = items.ToDictionary(item => item.Key, item => item.Value, StringComparer.Ordinal);
        var complete = MathTypeNativePreviewRenderer.TryRenderBatchWithRetry(batch, Path.GetTempPath(), out var results);
        foreach (var result in results) rendered[result.Key] = result.Value;
        if (complete) return;
        if (MathTypeNativePreviewRenderer.LastFailureWasFonts)
            throw new InvalidOperationException(MathTypeNativePreviewRenderer.FontsUnavailableMessage);
        var missing = items.Where(item => !rendered.ContainsKey(item.Key)).ToList();
        if (missing.Count <= 1) return;
        var half = missing.Count / 2;
        RenderMathTypeDisplayPreviews(missing.Take(half).ToList(), rendered);
        RenderMathTypeDisplayPreviews(missing.Skip(half).ToList(), rendered);
    }

    internal int RepairMathTypeDisplays(WordMathTypeDisplayRepairReport report)
    {
        if (report is null) throw new ArgumentNullException(nameof(report));
        var repaired = 0;
        // Last to first: each repair is its own verified Word transaction and
        // leaves every earlier target's position untouched.
        foreach (var target in report.Targets.OrderByDescending(item => item.Start))
        {
            ReportOperationProgress(
                OfficePluginLanguage.Text("正在重新生成公式图像…", "Redrawing equation pictures…"),
                repaired, report.Targets.Count);
            try
            {
                RepairMathTypeDisplay(report.DocumentId, target);
            }
            catch (Exception error)
            {
                throw new InvalidOperationException(
                    OfficePluginLanguage.Text(
                        $"已修复 {repaired} 个公式，其余公式未修改。{error.Message}",
                        $"{repaired} equations were repaired before this error; the others were not changed. {error.Message}"),
                    error);
            }
            repaired++;
        }
        WordDoubleClickHook.TraceMessage(
            $"mathtype-display-repair-complete mathType={report.MathTypeCount} "
            + $"unreadable={report.UnreadableCount} unrenderable={report.UnrenderableCount} repaired={repaired}");
        return repaired;
    }

    private void RepairMathTypeDisplay(string documentId, WordMathTypeDisplayRepairTarget target)
    {
        Document? document = null;
        InlineShape? oldShape = null;
        InlineShape? replacement = null;
        Range? oldRange = null;
        Range? insertion = null;
        UndoRecord? undoRecord = null;
        WordViewState? viewState = null;
        MathTypeLocalEditScope? localEdit = null;
        WordUndoHistorySnapshot? undoHistory = null;
        Options? editOptions = null;
        var mutationStarted = false;
        var previousTrackRevisions = false;
        var trackingSuspended = false;
        var previousSmartCutPaste = false;
        var smartCutPasteSuspended = false;
        var previousScreenUpdating = true;
        var screenUpdatingSuspended = false;
        var undoEnded = false;
        var viewRestored = false;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(document, documentId);
            viewState = CaptureViewState();
            try
            {
                previousScreenUpdating = _application.ScreenUpdating;
                _application.ScreenUpdating = false;
                screenUpdatingSuspended = true;
            }
            catch { }

            oldShape = ResolveMathTypeDisplayRepairTarget(document, target)
                ?? throw new InvalidOperationException("A MathType equation changed after the display scan.");
            oldRange = oldShape.Range.Duplicate;
            var inline = MathTypeOleInterop.IsInlineEquation(oldShape);
            // The old Word position came from the damaged picture's own baseline.
            // Place the new picture the way MathType insertion does: on MathType's
            // baseline, except inline in a paragraph aligned by character center,
            // top or Far East 50%, which already places the picture box on that axis.
            var characterAxis = inline && IsCharacterAxisAligned(oldRange);
            var paragraphLayout = CaptureMathTypeDisplayParagraphLayout(oldShape);
            localEdit = MathTypeLocalEditScope.Capture(document, oldRange);
            var fragment = MathTypeWordOpenXml.Read(oldShape);
            if (!string.Equals(Sha256(fragment.CompoundFile), target.CompoundFileHash, StringComparison.Ordinal))
                throw new InvalidOperationException("A MathType equation changed after the display scan.");

            // Keep any scaling the document applied to MathType's intrinsic
            // extent; only MathType's drawing and that extent are replaced.
            var widthScale = CalculateMathTypeNativePresentationScale(fragment.WidthPt, fragment.OriginalWidthPt);
            var heightScale = CalculateMathTypeNativePresentationScale(fragment.HeightPt, fragment.OriginalHeightPt);
            var wordPosition = characterAxis
                ? 0
                : (int)Math.Round(target.NativeWordPosition * heightScale, MidpointRounding.AwayFromZero);
            var replacementXml = MathTypeWordOpenXml.RewriteWithPlaceableWmf(
                fragment.WordOpenXml,
                fragment.CompoundFile,
                target.PreviewWmf,
                Math.Max(1f, target.NativeWidthPt * widthScale),
                Math.Max(1f, target.NativeHeightPt * heightScale),
                target.NativeWidthPt,
                target.NativeHeightPt);
            replacementXml = MathTypeLocalFragment.Prepare(
                replacementXml,
                localEdit.ParagraphXml,
                localEdit.LocatorName,
                localEdit.SourceOleOrdinal,
                localEdit.SourceOleCount);
            undoHistory = new WordUndoHistorySnapshot(document);

            // Every WordOpenXML export happened above: one undo item from here.
            undoRecord = BeginUndoRecord("VisualTeX Repair MathType Display")
                ?? throw new InvalidOperationException("Word could not establish a MathType display repair transaction.");
            previousTrackRevisions = document.TrackRevisions;
            if (previousTrackRevisions)
            {
                document.TrackRevisions = false;
                trackingSuspended = true;
            }
            editOptions = _application.Options;
            previousSmartCutPaste = editOptions.SmartCutPaste;
            if (previousSmartCutPaste)
            {
                editOptions.SmartCutPaste = false;
                smartCutPasteSuspended = true;
            }
            mutationStarted = true;
            oldRange.InsertXML(replacementXml);

            insertion = localEdit.ResolveInsertedRange();
            localEdit.RepairAndVerifyParagraph(insertion);
            Release(insertion);
            insertion = localEdit.ResolveInsertedRange();
            InlineShapes? localShapes = null;
            try
            {
                localShapes = insertion.InlineShapes;
                replacement = localShapes[1];
            }
            finally { Release(localShapes); }
            if (!MathTypeOleInterop.IsMathTypeOle(replacement))
                throw new InvalidOperationException("Word did not materialize the repaired MathType OLE.");
            if (paragraphLayout is not null)
                RestoreMathTypeDisplayParagraphLayout(replacement, paragraphLayout);
            SetInlineOleWordPosition(replacement, wordPosition);
            if (inline) RestoreTypingBaselineAfter(replacement);

            Range? replacementRange = null;
            try
            {
                replacementRange = replacement.Range;
                localEdit.RestoreContainedBookmarks(replacementRange);
            }
            finally { Release(replacementRange); }
            localEdit.RemoveLocator();
            if (trackingSuspended)
            {
                document.TrackRevisions = previousTrackRevisions;
                trackingSuspended = false;
            }
            // A display repair is not an edit: give the user back their own
            // selection and scroll position.
            RestoreViewState(document, viewState, null);
            viewRestored = true;
            if (screenUpdatingSuspended)
            {
                _application.ScreenUpdating = previousScreenUpdating;
                screenUpdatingSuspended = false;
            }
            EndUndoRecord(undoRecord);
            undoEnded = true;

            var repaired = MathTypeWordOpenXml.Read(replacement);
            if (!repaired.CompoundFile.SequenceEqual(fragment.CompoundFile))
                throw new InvalidOperationException("Word changed the MathType equation data during display repair.");
        }
        catch (Exception primary)
        {
            viewRestored = false;
            try
            {
                if (!undoEnded && undoRecord is not null)
                {
                    EndUndoRecord(undoRecord);
                    undoEnded = true;
                }
                if (mutationStarted && document is not null && undoHistory is not null)
                {
                    undoHistory.UndoChanges(document);
                    localEdit!.VerifyRollback();
                }
            }
            catch (Exception recovery)
            {
                throw new AggregateException(
                    "The MathType display repair failed and its local Word undo recovery could not be verified.",
                    primary, recovery);
            }
            throw;
        }
        finally
        {
            if (trackingSuspended && document is not null)
                document.TrackRevisions = previousTrackRevisions;
            localEdit?.Dispose();
            if (smartCutPasteSuspended && editOptions is not null)
            {
                try { editOptions.SmartCutPaste = previousSmartCutPaste; }
                catch (Exception optionError)
                {
                    WordDoubleClickHook.TraceMessage(
                        $"mathtype-display-repair-option-restore-failed error={optionError.Message}");
                }
            }
            Release(editOptions);
            if (screenUpdatingSuspended)
            {
                try { _application.ScreenUpdating = previousScreenUpdating; } catch { }
            }
            if (!viewRestored) RestoreViewState(document, viewState, null);
            if (!undoEnded) EndUndoRecord(undoRecord);
            Release(undoRecord);
            Release(insertion);
            Release(oldRange);
            Release(replacement);
            Release(oldShape);
            Release(document);
        }
    }

    private static InlineShape? ResolveMathTypeDisplayRepairTarget(
        Document document, WordMathTypeDisplayRepairTarget target)
    {
        Range? window = null;
        InlineShapes? shapes = null;
        try
        {
            window = document.Range(Math.Max(0, target.Start - 1), target.End + 1);
            shapes = window.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                Range? range = null;
                var keep = false;
                try
                {
                    range = shape.Range;
                    if (range.Start != target.Start || range.End != target.End) continue;
                    if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                    keep = true;
                    return shape;
                }
                finally { Release(range); if (!keep) Release(shape); }
            }
            return null;
        }
        finally { Release(shapes); Release(window); }
    }

    private static bool IsCharacterAxisAligned(Range range)
    {
        ParagraphFormat? format = null;
        try
        {
            format = range.ParagraphFormat;
            return format.BaseLineAlignment is WdBaselineAlignment.wdBaselineAlignCenter
                or WdBaselineAlignment.wdBaselineAlignTop
                or WdBaselineAlignment.wdBaselineAlignFarEast50;
        }
        finally { Release(format); }
    }

    private static string Sha256(byte[] bytes)
    {
        using var hash = SHA256.Create();
        return Convert.ToBase64String(hash.ComputeHash(bytes));
    }
}

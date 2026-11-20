using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// One size for every inline or every display formula of the document (Ribbon 行内 / 行间, as on the Mac):
// VisualTeX objects and Word equations in one Undo step, then MathType objects one by one (a MathType
// rewrite exports the object's XML first, which Word does not allow inside an open Undo record).
// A VisualTeX object is resized in Word by its cached metadata, without reading its payload
// (WordVisualTeXHostWriter.PrepareCachedFontSize and the three steps after it); one whose cache cannot
// be used takes the payload path.
internal sealed partial class WordFormulaService
{
    internal sealed class FormulaFontSizeBatchResult
    {
        internal int Resized { get; set; }
        internal List<string> Failures { get; } = new();
        // Stage timings, collected only when the acceptance log is on.
        internal string? MetricDetail { get; set; }
    }

    internal FormulaFontSizeBatchResult ApplyFontSizeToDocumentFormulas(bool display, double requestedFontSizePt)
    {
        var size = FormulaFontSize.Normalize(requestedFontSizePt);
        var result = new FormulaFontSizeBatchResult();
        Document? document = null;
        Selection? selection = null;
        var cached = new Dictionary<WordFormulaHostDescriptor, WordVisualTeXHostWriter.CachedFontSizeTarget>();
        var metrics = string.IsNullOrWhiteSpace(
                Environment.GetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG"))
            ? null
            : new WordOperationMetrics();
        var watch = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(OfficePluginLanguage.Text(
                    "没有打开的 Word 文档。", "No Word document is open."));
            EnsureWritable(document);
            var target = document;

            // A numbered equation is a display formula, whatever its OMath type. A display formula's
            // numbering is needed only for a VisualTeX object's number (ApplyNumberSize): it is not
            // looked up for the others (10 ms a formula).
            var hosts = new List<WordFormulaHostDescriptor>();
            WordFormulaDocumentIndex index;
            using (WordOperationMetrics.Measure("FontSize.Index"))
                index = WordFormulaHostResolver.CaptureDocumentIndex(target);
            using (WordOperationMetrics.Measure("FontSize.Numbering"))
            foreach (var host in index.Omml.Concat(index.VisualTeX))
            {
                try
                {
                    if (host.Display)
                    {
                        if (!display) continue;
                        if (host.Kind == WordFormulaHostKind.VisualTeX)
                            host.Numbering = WordFormulaNumberingResolver.ResolveLocal(target, host);
                        hosts.Add(host);
                        continue;
                    }
                    host.Numbering = WordFormulaNumberingResolver.ResolveLocal(target, host);
                    if (host.Numbering.Numbered == display) hosts.Add(host);
                }
                catch (Exception error)
                {
                    result.Failures.Add(error.Message);
                }
            }
            hosts.Sort((left, right) => right.Range.Start.CompareTo(left.Range.Start));
            using (WordOperationMetrics.Measure("FontSize.Ole.Prepare"))
                foreach (var host in hosts.Where(item => item.Kind == WordFormulaHostKind.VisualTeX))
                {
                    var prepared = WordVisualTeXHostWriter.PrepareCachedFontSize(target, host);
                    if (prepared is not null) cached[host] = prepared;
                }
            if (hosts.Count > 0)
            {
                using (WordOperationMetrics.Measure("FontSize.Transaction"))
                result.Resized = WordFormulaMutationTransaction.Execute(
                    _application,
                    target,
                    "VisualTeX Formula Font Size",
                    () =>
                    {
                        // Objects resized by their cached metadata: every position and number, then
                        // every size, then every cache (WordVisualTeXHostWriter.PositionCached).
                        using (WordOperationMetrics.Measure("FontSize.Ole.Position"))
                            foreach (var pair in cached)
                                ApplyNumberSize(
                                    target,
                                    WordVisualTeXHostWriter.PositionCached(target, pair.Key, pair.Value, size));
                        using (WordOperationMetrics.Measure("FontSize.Ole.Resize"))
                            foreach (var prepared in cached.Values)
                                WordVisualTeXHostWriter.ResizeCached(prepared);
                        using (WordOperationMetrics.Measure("FontSize.Ole.Cache"))
                            foreach (var prepared in cached.Values)
                                WordVisualTeXHostWriter.CacheResized(prepared);
                        foreach (var host in hosts.Where(item => !cached.ContainsKey(item)))
                            ResizeHostInTransaction(target, host, size);
                        return hosts.Count;
                    });
            }

            // The MathType pass runs only when the index met an inline OLE object that is not
            // VisualTeX's (as GetEquationReferenceTargets does); it reads where each object stands,
            // not its equation (the single-formula command below reads the ones it changes).
            var mathType = new List<WordFormulaHostDescriptor>();
            if (index.HasForeignOle)
                using (WordOperationMetrics.Measure("FontSize.MathType.Index"))
                    mathType = WordMathTypeHostAdapter.CaptureLayoutIndex(target)
                        .Where(host => (host.Display || host.Numbering.Numbered) == display)
                        .OrderByDescending(host => host.Range.Start)
                        .ToList();
            if (mathType.Count > 0)
            {
                // The single-formula command resizes the selected MathType object; the user's
                // selection comes back afterwards.
                selection = _application.Selection;
                var start = selection.Start;
                var end = selection.End;
                foreach (var host in mathType)
                {
                    Range? range = null;
                    try
                    {
                        range = WordFormulaHostSemanticReader.CreateRange(target, host.Range);
                        range.Select();
                        using (WordOperationMetrics.Measure("FontSize.MathType.Each"))
                            SetSelectedFormulaFontSizeCore(size);
                        result.Resized++;
                    }
                    catch (Exception error) when (MathTypeOleInterop.IsBlankEquationError(error))
                    {
                        // An empty MathType object has nothing to size.
                    }
                    catch (Exception error)
                    {
                        result.Failures.Add(error.Message);
                    }
                    finally { Release(range); }
                }
                try { selection.SetRange(start, end); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            return result;
        }
        finally
        {
            foreach (var prepared in cached.Values) prepared.Dispose();
            if (metrics is not null)
            {
                metrics.Dispose();
                result.MetricDetail = $"totalMs={watch.ElapsedMilliseconds} cached={cached.Count} " + string.Join(
                    " | ",
                    metrics.Entries
                        .Where(entry => entry.Key.StartsWith("FontSize.", StringComparison.Ordinal))
                        .OrderBy(entry => entry.Key)
                        .Select(entry => $"{entry.Key}={entry.Value.Milliseconds:0.0}ms/{entry.Value.Calls}"));
            }
            Release(selection);
            Release(document);
        }
    }

    // As the single-formula command does (SetSelectedFormulaFontSizeCore), inside the batch's transaction.
    private static void ResizeHostInTransaction(Document document, WordFormulaHostDescriptor host, float size)
    {
        if (host.Kind == WordFormulaHostKind.Omml)
        {
            // A Word equation has nothing to own (AdoptUnownedOmmlForMutation writes nothing into
            // the document): its range is all the size needs.
            Range? range = null;
            try
            {
                range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
                using (WordOperationMetrics.Measure("FontSize.Omml.Typography"))
                    WordFormulaHostLayout.ApplyOmmlLocalTypography(range, size);
            }
            finally { Release(range); }
            return;
        }
        using (WordOperationMetrics.Measure("FontSize.Ole.Payload"))
            ApplyNumberSize(document, WordVisualTeXHostWriter.UpdateFontSize(document, host, size));
    }

    // The number of a numbered VisualTeX formula in the formula's size.
    private static void ApplyNumberSize(Document document, WordFormulaHostDescriptor updated)
    {
        if (updated.Numbering.ContainerKind == WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
            && WordVisualTeXParagraphNumbering.IsSelfContainedHost(document, updated)
            && updated.Metadata is not null)
            WordVisualTeXParagraphNumbering.ApplyNumberFormat(document, updated, updated.Metadata);
    }
}

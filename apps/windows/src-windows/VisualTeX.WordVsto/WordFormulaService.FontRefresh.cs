using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

// Update Equation Format's font step: the VisualTeX formulas drawn in other fonts than the ones the main app draws
// in now (math font, Chinese font, fonts for letters, Greek letters and digits) are drawn again from their LaTeX and
// written where they stand, each keeping its id, number, references, size and display mode, in one Undo step.
internal sealed partial class WordFormulaService
{
    /// <summary>A VisualTeX formula drawn in other fonts, as the scan read it.</summary>
    internal sealed class FontRefreshTarget
    {
        internal string FormulaId { get; set; } = string.Empty;
        internal string? DocumentId { get; set; }
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string DisplayMode { get; set; } = "inline";
        internal bool Numbered { get; set; }
        internal double FontSizePt { get; set; }
    }

    /// <summary>The document's VisualTeX formulas whose recorded fonts differ from <paramref name="current"/>.</summary>
    internal IReadOnlyList<FontRefreshTarget> CaptureFontRefreshTargets(FormulaFontSettings current)
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(NoActiveDocumentMessage);
            var targets = new List<FontRefreshTarget>();
            var documentId = DocumentIdentity(document);
            var index = WordFormulaHostResolver.CaptureDocumentIndex(document);
            foreach (var host in index.VisualTeX.OrderBy(item => item.Range.Start))
            {
                var metadata = host.Metadata;
                if (metadata is null
                    || string.IsNullOrWhiteSpace(metadata.Latex)
                    || !Guid.TryParse(host.FormulaId ?? metadata.FormulaId, out var formulaId)
                    || !current.DifferFrom(metadata))
                    continue;
                var display = !string.Equals(metadata.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase);
                var numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
                targets.Add(new FontRefreshTarget
                {
                    FormulaId = formulaId.ToString("D"),
                    DocumentId = documentId,
                    Metadata = metadata,
                    DisplayMode = display ? "block" : "inline",
                    Numbered = display && (metadata.Numbered || numbering.Numbered),
                    FontSizePt = FormulaFontSize.ResolveSemanticFontSize(metadata),
                });
            }
            return targets;
        }
        finally { Release(document); }
    }

    /// <summary>
    /// Writes the formulas drawn again in the current fonts in place of the old ones, found again by their ids (the
    /// scan's ranges do not survive other changes), in one transaction. A formula that changed since the scan is
    /// left as it is and reported.
    /// </summary>
    internal int ApplyFontRefresh(
        IReadOnlyList<(FontRefreshTarget Target, PreparedWordBulkFormula Prepared)> items,
        List<string> failures)
    {
        using var numberFormatScope = WordEquationNumberFormatAuthority.BeginOperationScope();
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(NoActiveDocumentMessage);
            EnsureWritable(document);
            var index = WordFormulaHostResolver.CaptureDocumentIndex(document);
            var hosts = new Dictionary<string, WordFormulaHostDescriptor>(StringComparer.OrdinalIgnoreCase);
            foreach (var host in index.VisualTeX)
                if (Guid.TryParse(host.FormulaId ?? host.Metadata?.FormulaId, out var id))
                    hosts[id.ToString("D")] = host;

            var writes = new List<(WordFormulaHostDescriptor Host, WordFormulaHostWriteRequest Request)>();
            foreach (var (target, prepared) in items)
            {
                if (!hosts.TryGetValue(target.FormulaId, out var host)
                    || host.Metadata is null
                    || !string.Equals(host.Metadata.Latex, target.Metadata.Latex, StringComparison.Ordinal))
                {
                    failures.Add(OfficePluginLanguage.Text(
                        $"公式在准备期间被修改或删除，保持原样：{WordFormulaRenderException.Excerpt(target.Metadata.Latex, 60)}",
                        $"The equation changed or was removed while it was being prepared and was left as it is: {WordFormulaRenderException.Excerpt(target.Metadata.Latex, 60)}"));
                    continue;
                }
                host.Numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
                writes.Add((host, BuildHostWriteRequest(prepared.Session, prepared.MathMl, prepared.PngPath, prepared.EmfPath)));
            }
            if (writes.Count == 0) return 0;
            // from the end of the document: a replacement never moves the formulas before it
            writes.Sort((left, right) => right.Host.Range.Start.CompareTo(left.Host.Range.Start));
            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Update Formula Fonts",
                () =>
                {
                    var written = 0;
                    foreach (var (host, request) in writes)
                    {
                        ReportOperationProgress(
                            OfficePluginLanguage.Text("正在写入用当前字体重绘的公式…", "Writing equations drawn in the current fonts…"),
                            written,
                            writes.Count);
                        WordFormulaHostMutationKernel.ReplaceInActiveTransaction(
                            _application,
                            document,
                            host,
                            request,
                            sourceAlreadyResolved: true);
                        written++;
                    }
                    return written;
                });
        }
        finally { Release(document); }
    }
}

using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class PreparedWordBulkFormula
{
    internal WordBulkRun Run { get; set; } = new();
    internal OfficeSessionDocument Session { get; set; } = new();
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? SvgPath { get; set; }
    internal float? RawRenderWidthPx { get; set; }
    internal float? RawRenderHeightPx { get; set; }
    internal float? RawBaselinePx { get; set; }
    internal string? EmfPath { get; set; }
    internal bool MathTypeNativePreviewAttempted { get; set; }
    internal MathTypeNativePreviewRenderer.Result? MathTypeNativePreview { get; set; }
    // Batch MathType preparation may already have produced and semantically
    // validated the exact standalone Equation Native/CFB payload while building
    // the native preview. Reuse that immutable payload during Word insertion
    // instead of running the same codec + storage round-trip a second time.
    internal MathTypeMtefCodec.RewriteResult? MathTypeStandaloneGenerated { get; set; }
    internal byte[]? MathTypeStandaloneCompoundFile { get; set; }
    internal string? MathTypeStandaloneSemanticSignature { get; set; }
}

internal sealed class RenderedWordBulkFormulaTemplate
{
    internal OfficeSessionDocument Session { get; set; } = new();
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? SvgPath { get; set; }
    internal float? RawRenderWidthPx { get; set; }
    internal float? RawRenderHeightPx { get; set; }
    internal float? RawBaselinePx { get; set; }
    internal string? EmfPath { get; set; }
}

internal sealed class WordBulkInsertResult
{
    internal int BlockCount { get; set; }
    internal int FormulaCount { get; set; }
    internal List<string> FormulaIds { get; set; } = new();
}

internal sealed class WordLatexRedrawTarget
{
    internal string Id { get; set; } = Guid.NewGuid().ToString("D");
    internal int RelativeStart { get; set; }
    internal int SourceLength { get; set; }
    internal int AbsoluteStart { get; set; } = -1;
    internal int AbsoluteEnd { get; set; } = -1;
    internal string Latex { get; set; } = string.Empty;
    internal string DisplayMode { get; set; } = "inline";
    internal bool PreserveDisplayParagraphBoundary { get; set; }
    internal bool DedicatedDisplayParagraph { get; set; }
    internal double FontSizePt { get; set; }
    // Set only when the LaTeX source was produced from an inline native object and
    // carries a VisualTeX provenance bookmark for its established Word baseline.
    internal int? SourceInlineWordPosition { get; set; }
    internal float? SourceInlineBottomWhitespacePoints { get; set; }
}

internal sealed class WordLatexRedrawPlan
{
    internal string DocumentId { get; set; } = string.Empty;
    internal int ScopeStart { get; set; }
    internal int ScopeEnd { get; set; }
    internal string SourceText { get; set; } = string.Empty;
    internal bool NumberDisplayFormulas { get; set; }
    internal List<WordLatexRedrawTarget> Targets { get; set; } = new();
    internal Dictionary<string, long> CaptureTimings { get; } =
        new(StringComparer.Ordinal);
    internal List<DeletedNumberedOmmlResidue> DeletedNumberedOmmlResidues { get; set; } = new();
}

internal sealed class DeletedNumberedOmmlResidue
{
    internal string FormulaId { get; set; } = string.Empty;
    internal int OwnerStart { get; set; }
    internal int OwnerEnd { get; set; }
}

internal sealed class WordLatexRedrawResult
{
    internal int FormulaCount { get; set; }
    internal long TotalSourceMutationMilliseconds { get; set; }
    internal long TotalInsertMilliseconds { get; set; }
    internal long MaxInsertMilliseconds { get; set; }
    // Timings are kept in actual document-order mutation order. The acceptance
    // logger uses them to detect index-dependent growth without adding file I/O
    // to the hot path.
    internal List<long> SourceMutationMilliseconds { get; } = new();
    internal List<long> InsertMilliseconds { get; } = new();
    internal List<string> InsertMetricDetails { get; } = new();
    internal Dictionary<string, long> BatchTimings { get; } =
        new(StringComparer.Ordinal);
    internal List<string> FormulaIds { get; set; } = new();
    internal List<(string TargetId, string Latex, string Error)> Failures { get; } =
        new();
    internal List<string> Warnings { get; } = new();
}

internal sealed class WordFormulaToLatexResult
{
    internal int FormulaCount { get; set; }
    internal List<string> FormulaIds { get; set; } = new();
}

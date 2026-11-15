using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal enum EquationReferenceSource
{
    VisualTeX,
    WordOmml,
    MathType,
}

internal sealed class EquationReferenceTarget
{
    public EquationReferenceTarget(
        string formulaId,
        int nativeReferenceItem,
        string numberText,
        string latexPreview,
        int position,
        EquationReferenceSource source = EquationReferenceSource.VisualTeX,
        int endPosition = -1)
    {
        FormulaId = formulaId;
        NativeReferenceItem = nativeReferenceItem;
        NumberText = numberText;
        _latexPreview = latexPreview;
        Position = position;
        EndPosition = endPosition;
        Source = source;
    }

    /// A target whose LaTeX preview is read when it is first needed (reading it is a Word OOXML or OLE
    /// read per formula: in a 400-formula document most of the time of listing the targets).
    public EquationReferenceTarget(
        string formulaId,
        int nativeReferenceItem,
        string numberText,
        Func<string> latexPreview,
        int position,
        EquationReferenceSource source,
        int endPosition)
        : this(formulaId, nativeReferenceItem, numberText, string.Empty, position, source, endPosition)
    {
        _readLatexPreview = latexPreview;
    }

    private string _latexPreview;
    private Func<string>? _readLatexPreview;

    public string FormulaId { get; }
    public int NativeReferenceItem { get; }
    public string NumberText { get; }
    public string LatexPreview
    {
        get
        {
            var read = _readLatexPreview;
            if (read is not null)
            {
                _readLatexPreview = null;
                try { _latexPreview = read() ?? string.Empty; }
                catch { _latexPreview = string.Empty; }
            }
            return _latexPreview;
        }
    }
    /// Whether the preview has been read (ToString shows it only then: a list of targets does not
    /// read every formula to draw itself).
    public bool LatexPreviewReady => _readLatexPreview is null;
    public int Position { get; }
    public int EndPosition { get; }
    public EquationReferenceSource Source { get; }

    public override string ToString()
    {
        var preview = LatexPreviewReady ? LatexPreview : string.Empty;
        return Source == EquationReferenceSource.MathType
            ? $"{NumberText}    {preview}"
            : $"({NumberText})    {preview}";
    }
}

internal enum EquationReferenceStyle
{
    Parenthesized,
    EquationPrefix,
    NumberOnly,
}

internal sealed class EquationNumberFormat
{
    public const string ContinuousId = "continuous";
    public const string Heading1DotId = "heading1-dot";
    public const string Heading1DashId = "heading1-dash";
    public const string Heading2DotId = "heading2-dot";
    public const string Heading2DashId = "heading2-dash";

    private EquationNumberFormat(
        string id,
        string displayName,
        int headingLevel,
        string separator)
    {
        Id = id;
        DisplayName = displayName;
        HeadingLevel = headingLevel;
        Separator = separator;
    }

    public string Id { get; }
    public string DisplayName { get; }
    public int HeadingLevel { get; }
    public string Separator { get; }
    public bool UsesHeading => HeadingLevel > 0;

    public static EquationNumberFormat Resolve(string? id) => id switch
    {
        Heading1DotId => new EquationNumberFormat(Heading1DotId, OfficePluginLanguage.Text("按章编号（1.1）", "By chapter (1.1)"), 1, "."),
        Heading1DashId => new EquationNumberFormat(Heading1DashId, OfficePluginLanguage.Text("按章编号（1-1）", "By chapter (1-1)"), 1, "-"),
        Heading2DotId => new EquationNumberFormat(Heading2DotId, OfficePluginLanguage.Text("按节编号（1.1.1）", "By section (1.1.1)"), 2, "."),
        Heading2DashId => new EquationNumberFormat(Heading2DashId, OfficePluginLanguage.Text("按节编号（1.1-1）", "By section (1.1-1)"), 2, "-"),
        _ => new EquationNumberFormat(ContinuousId, OfficePluginLanguage.Text("全文连续编号（1）", "Continuous (1)"), 0, string.Empty),
    };
}

internal sealed class ResolvedEquationHeadingScope
{
    internal ResolvedEquationHeadingScope(
        int scopeStart,
        int scopeEnd,
        string numberText)
    {
        ScopeStart = scopeStart;
        ScopeEnd = scopeEnd;
        NumberText = numberText;
    }

    internal int ScopeStart { get; }
    internal int ScopeEnd { get; }
    internal string NumberText { get; }
}

/// <summary>
/// What the classes of equation numbering share: the names and prefixes of the bookmarks, fields and shapes
/// VisualTeX owns, the compact typing tail, the window DPI and the numbering performance trace. Each area of
/// numbering is a static class of its own (WordEquationNumberFormat, WordNativeCaptions,
/// WordVisibleEquationNumbers, ...).
/// </summary>
internal static class WordEquationNumbering
{
    internal const int WdFieldEmpty = -1;
    internal const string LegacyEquationSequenceName = "VisualTeXEquation";
    internal const string EquationBookmarkPrefix = "VTEq_";
    internal const string NativeCaptionBookmarkPrefix = "VTEqCap_";
    internal const string NativeNumberBookmarkPrefix = "VTEqNum_";
    // Word bookmark names are limited to 40 characters. Frames.Add can create one
    // extra body-flow paragraph after an external caption immediately before a
    // user table; this 7-character prefix + 32-character UUID owns only that exact
    // generated empty paragraph so reverse conversion can reclaim it safely.
    internal const string NativeCaptionFlowTailBookmarkPrefix = "VTEqFT_";
    internal const string NativeDisplayAnchorBookmarkPrefix = "VTEqAnc_";
    // Word bookmark names are limited to 40 characters; this 7-character prefix
    // plus a 32-character UUID marks an anchor that has survived one Office turn.
    internal const string NativeDisplayAnchorCommitBookmarkPrefix = "VTAncR_";
    internal const string NativeDisplayNumberShapeNamePrefix = "VTEqShape_";
    internal const string NativeDisplayNumberShapeAlternativeTextPrefix =
        "VisualTeX numbered OMML ";
    internal const string CompactTypingTailBookmarkName = "VTNumberedTypingTail";
    internal const float CompactTypingTailFontSizePoints = 1f;
    internal const float CompactTypingTailLineSpacingPoints = 1f;
    private static readonly object NumberingPerformanceTraceSync = new();

    [DllImport("user32.dll")]
    internal static extern uint GetDpiForWindow(IntPtr windowHandle);

    internal static void TraceNumberingPerformance(string message)
    {
        Console.WriteLine(message);
        var tracePath = Environment.GetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE_PATH");
        if (string.IsNullOrWhiteSpace(tracePath)) return;
        try
        {
            lock (NumberingPerformanceTraceSync)
            {
                System.IO.File.AppendAllText(
                    tracePath,
                    $"{DateTimeOffset.Now:O} pid={System.Diagnostics.Process.GetCurrentProcess().Id} {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Diagnostic tracing must never affect numbering behavior.
        }
    }
}

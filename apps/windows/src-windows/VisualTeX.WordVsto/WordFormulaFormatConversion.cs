using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

internal sealed class WordFormulaFormatConversionPlan
{
    internal string DocumentId { get; set; } = string.Empty;
    internal bool WritableValidated { get; set; }
    internal string SourceMode { get; set; } = string.Empty;
    internal string TargetMode { get; set; } = string.Empty;
    internal bool WholeDocument { get; set; }
    internal string NumberFormatId { get; set; } = EquationNumberFormat.ContinuousId;
    // Whole-document OLE capture already visits every InlineShape. Cache exact
    // source/target counts there so Apply does not enumerate 1000 OLEFormat RCWs
    // again merely to establish its pre-transaction cardinality baseline.
    internal int? InitialSourceObjectCount { get; set; }
    internal int? InitialTargetObjectCount { get; set; }
    internal List<WordFormulaFormatConversionTarget> Targets { get; set; } = new();
}

internal sealed class WordFormulaFormatConversionTarget
{
    internal string Id { get; set; } = Guid.NewGuid().ToString("D");
    internal string SourceFormulaId { get; set; } = string.Empty;
    internal string SourceObjectId { get; set; } = string.Empty;
    internal int SourceStart { get; set; }
    internal string Latex { get; set; } = string.Empty;
    internal string? SourceMathMl { get; set; }
    internal bool SourceIsManagedOmml { get; set; }
    // Captured before any source object is deleted. Table-contained formulas use
    // cell-safe replacement paths and never participate in whole-paragraph/group
    // optimizations that can consume Word's terminal \r\a cell markers.
    internal bool SourceWithinTable { get; set; }
    // A foreign section-state field is document-owned, not part of the equation
    // being replaced. Preflight records it; the edit transaction isolates it.
    internal bool SourceHasMathTypeSectionPrefix { get; set; }
    // For table-contained inline formulas, Word can consume the first ordinary
    // whitespace character after the formula while replacing a native OMath with
    // an OLE field. Capture a short user-text suffix plus the first character's
    // formatting before any mutation so the write transaction can prove/repair
    // exactly that one-character boundary loss without guessing at surrounding text.
    internal string? FollowingInlineTableText { get; set; }
    internal WordCharacterFormatting? FollowingInlineTableFormatting { get; set; }
    internal string DisplayMode { get; set; } = "inline";
    internal bool Numbered { get; set; }
    internal int PrecedingPlainBlankParagraphCount { get; set; }
    internal string MathTypeNumberPosition { get; set; } = "right";
    internal double FontSizePt { get; set; } = FormulaFontSize.DefaultPt;
    internal float? MathTypeDisplayColumnWidth { get; set; }
    internal FormulaMetadata Metadata { get; set; } = new();
}

internal sealed class WordFormulaFormatConversionResult
{
    internal int FormulaCount { get; set; }
    internal int FailedFormulaCount { get; set; }
    internal List<string> Failures { get; set; } = new();
}

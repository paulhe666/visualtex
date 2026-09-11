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

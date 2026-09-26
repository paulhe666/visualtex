using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WordVsto;

internal enum WordFormulaHostKind
{
    Omml,
    VisualTeX,
    MathType,
}

internal sealed class WordFormulaRangeAddress
{
    internal WdStoryType StoryType { get; set; }
    internal int Start { get; set; }
    internal int End { get; set; }

    internal int Length => Math.Max(0, End - Start);

    internal bool ContainsCaret(int position) =>
        position >= Start && position <= End;

    internal bool Overlaps(int start, int end) =>
        start < End && end > Start;
}

internal enum WordFormulaNumberingContainerKind
{
    None,
    CanonicalNativeOmml,
    CanonicalBodyTabParagraph,
    CanonicalBodyTable,
    CanonicalUserTableCell,
    Legacy,
}

internal sealed class WordFormulaNumberingDescriptor
{
    internal bool Numbered { get; set; }
    internal string Position { get; set; } = "right";
    internal string? FormulaId { get; set; }
    internal WordFormulaNumberingContainerKind ContainerKind { get; set; }
    // True only when the canonical body-tab paragraph was proved by the current
    // self-contained VisualTeXPlaceRef resolver. Older compatible body-tab hosts
    // intentionally leave this false so migration/detach keeps its legacy checks.
    internal bool SelfContainedVisualTeXParagraph { get; set; }
    internal WordFormulaRangeAddress? ContainerRange { get; set; }
    internal WordFormulaRangeAddress? NumberRange { get; set; }
}

internal sealed class WordFormulaHostDescriptor
{
    internal WordFormulaHostKind Kind { get; set; }
    internal WordFormulaRangeAddress Range { get; set; } = new();
    internal string DisplayMode { get; set; } = "inline";
    internal string? FormulaId { get; set; }
    internal FormulaMetadata? Metadata { get; set; }
    // Batch indexes may use Word's cached preview metadata only as a scheduling
    // hint. A destructive mutation must re-resolve the host locally and require
    // authoritative embedded VisualTeX metadata before deleting anything.
    internal bool MetadataAuthoritative { get; set; }
    internal WordFormulaNumberingDescriptor Numbering { get; set; } = new();
    internal bool WithinTable { get; set; }
    internal string? SourceMathMl { get; set; }
    internal string? Latex { get; set; }

    internal bool Display =>
        string.Equals(DisplayMode, "block", StringComparison.OrdinalIgnoreCase);
}

internal sealed class WordFormulaSemanticPayload
{
    internal string Latex { get; set; } = string.Empty;
    internal string? MathMl { get; set; }
    internal string? WordOpenXml { get; set; }
    internal FormulaMetadata? Metadata { get; set; }
}

internal sealed class WordFormulaHostWriteRequest
{
    internal WordFormulaHostKind Kind { get; set; }
    internal string DisplayMode { get; set; } = "inline";
    internal bool Numbered { get; set; }
    internal string? FormulaId { get; set; }
    internal FormulaMetadata? Metadata { get; set; }
    internal string? MathMl { get; set; }
    internal string? PngPath { get; set; }
    internal string? EmfPath { get; set; }
    internal float WidthPoints { get; set; }
    internal float HeightPoints { get; set; }
    internal float ExportedHeightPixels { get; set; }
    internal float? ExportedBaselinePixels { get; set; }
    internal double FontSizePoints { get; set; } = FormulaFontSize.DefaultPt;
    // Optional batch-resolved display tab geometry. These values are local to the
    // exact source section/cell captured by the caller and avoid re-reading the
    // same PageSetup for every display formula in a batch.
    internal float? DisplayCenterTabPoints { get; set; }
    internal float? DisplayRightTabPoints { get; set; }

    // Set only by a whole-document/selection batch mutation such as
    // format conversion or LaTeX redraw. The batch owns one outer transaction
    // plus one final numbering pass, so the mutation kernel may avoid post-write
    // checks that would immediately be repeated by the same batch. Interactive
    // create/edit leaves this false.
    internal bool BatchMutation { get; set; }
}

internal sealed class WordFormulaHostWriteResult
{
    internal WordFormulaHostDescriptor Host { get; set; } = new();
}

internal sealed class WordFormulaDocumentIndex
{
    internal WordFormulaDocumentIndex(
        IReadOnlyList<WordFormulaHostDescriptor> omml,
        IReadOnlyList<WordFormulaHostDescriptor> visualTeX)
    {
        Omml = omml;
        VisualTeX = visualTeX;
    }

    internal IReadOnlyList<WordFormulaHostDescriptor> Omml { get; }
    internal IReadOnlyList<WordFormulaHostDescriptor> VisualTeX { get; }

    internal IReadOnlyList<WordFormulaHostDescriptor> ForKind(
        WordFormulaHostKind kind) =>
        kind switch
        {
            WordFormulaHostKind.Omml => Omml,
            WordFormulaHostKind.VisualTeX => VisualTeX,
            _ => Array.Empty<WordFormulaHostDescriptor>(),
        };
}

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
    // OMML capture reads one complete WordOpenXML snapshot. Keep only its SHA-256
    // digest so Apply can prove that no Word content changed while converter
    // sessions were being prepared, without resolving every OMML identity again.
    internal string? SourceDocumentXmlHash { get; set; }
    internal List<WordFormulaFormatConversionTarget> Targets { get; set; } = new();
    internal List<DeletedNumberedOmmlResidue> DeletedNumberedOmmlResidues { get; set; } = new();
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
    // Capture's single XML identity index can prove a numbered managed OMML owns
    // either its direct native number OMath or the standard 1x3 center/number row.
    // Apply may reuse that proof only while the normalized source-document
    // signature remains identical; legacy/ambiguous hosts keep the COM checks.
    internal bool SourceOmmlNumberHostVerifiedBySnapshot { get; set; }
    // The same immutable WordOpenXML snapshot records whether this managed OMML
    // owns VisualTeX's generated post-table separator bookmark. Apply uses this
    // only after the document signature is revalidated; an absent marker then
    // needs no per-formula Bookmarks.Exists/Delete round trip.
    internal bool SourceOmmlOwnedSeparatorPresentBySnapshot { get; set; }
    // Set only when one batch metadata inventory proved exactly one CustomXMLPart
    // owns this FormulaId. Apply revalidates the inventory before mutation and can
    // then delete that exact part by ID instead of rescanning all metadata.
    internal string? SourceOmmlMetadataPartId { get; set; }
    // A Word-native edit can move the collapsed VTOMML anchor inside or to the
    // end of the selected OMath while also changing its content fingerprint.
    // Only an explicitly selected complete, locally unique OMath may carry this
    // flag. Apply repairs that identity inside the conversion UndoRecord before
    // any structural resolver or source deletion runs.
    internal bool SourceRequiresManagedOmmlIdentityRebind { get; set; }
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
    // Preserve the source Word-layer inline baseline during MathType -> VisualTeX
    // conversion. MathType already presents the formula at this object-character
    // Position; recalculating a fresh VisualTeX descent here can visibly push an
    // otherwise correctly aligned formula down by several pixels.
    internal int? SourceInlineWordPosition { get; set; }
    // Optical provenance from the source MathType preview. Word Position is whole
    // points, so sub-point residuals are repaired later inside the VisualTeX EMF
    // while keeping the source-position conversion geometry unchanged.
    internal float? SourceInlineBottomWhitespacePoints { get; set; }
    internal float? MathTypeDisplayColumnWidth { get; set; }
    internal FormulaMetadata Metadata { get; set; } = new();
}

internal sealed class WordFormulaFormatConversionResult
{
    internal int FormulaCount { get; set; }
    internal int FailedFormulaCount { get; set; }
    internal List<string> Failures { get; set; } = new();
}

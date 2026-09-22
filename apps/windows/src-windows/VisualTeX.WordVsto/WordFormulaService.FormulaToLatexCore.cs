using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private sealed class CoreFormulaToLatexTarget
    {
        internal WordFormulaHostDescriptor Host { get; set; } = new();
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string LatexSource { get; set; } = string.Empty;
        internal int? InlineWordPosition { get; set; }
        internal float? InlineBottomWhitespacePoints { get; set; }
    }

    internal int CountFormulaObjectsForLatexCore(
        bool wholeDocument,
        string objectMode)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;

            return CaptureFormulaToLatexCoreTargets(
                    document,
                    scope,
                    wholeDocument,
                    objectMode,
                    readSemanticPayload: false)
                .Count;
        }
        finally
        {
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    internal WordFormulaToLatexResult ConvertFormulaObjectsToLatexCore(
        bool wholeDocument,
        string objectMode)
    {
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        WordViewState? viewState = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);

            selection = _application.Selection;
            scope = wholeDocument
                ? document.Content.Duplicate
                : selection.Range.Duplicate;

            var targets =
                CaptureFormulaToLatexCoreTargets(
                    document,
                    scope,
                    wholeDocument,
                    objectMode,
                    readSemanticPayload: true);
            if (targets.Count == 0)
                throw new InvalidDataException(
                    wholeDocument
                        ? "当前 Word 文档中没有找到可转换的公式。"
                        : "所选内容中没有找到可转换的公式。");

            var sourceKind =
                ObjectModeToHostKind(
                    objectMode);

            viewState = CaptureViewState();

            var result =
                WordFormulaMutationTransaction.Execute(
                    _application,
                    document,
                    "VisualTeX Formula To LaTeX",
                    () =>
                    {
                        // Removing a numbered formula must also remove its
                        // VTEqNum_<FormulaId> target. Existing REF/GOTOBUTTON
                        // fields remain as real Word references and therefore
                        // correctly become "Error! Reference source not found."
                        // once their target no longer exists. Never freeze them
                        // into plain text here.
                        var converted =
                            new WordFormulaToLatexResult();
                        foreach (var target in targets
                                     .OrderByDescending(
                                         item => item.Host.Range.Start))
                        {
                            var source =
                                WordFormulaOperationLocator
                                    .ResolveCapturedHost(
                                        _application,
                                        document,
                                        RangeReferenceFromAddress(
                                            target.Host.Range),
                                        sourceKind,
                                        target.Host.FormulaId);

                            _ = WordFormulaHostMutationKernel
                                .ReplaceSourceWithExternalTargetInActiveTransaction(
                                    _application,
                                    document,
                                    source,
                                    insertion =>
                                    {
                                        var start = insertion.Start;
                                        insertion.Text =
                                            target.LatexSource;
                                        Range? inserted = null;
                                        try
                                        {
                                            inserted = document.Range(
                                                start,
                                                start
                                                + target.LatexSource.Length);
                                            VerifyLatexSourceRange(
                                                inserted,
                                                target.LatexSource,
                                                target.Metadata.FormulaId);
                                            NormalizeLatexSourceRange(
                                                inserted,
                                                target.Metadata);
                                            if (string.Equals(
                                                    target.Metadata.DisplayMode,
                                                    "inline",
                                                    StringComparison.OrdinalIgnoreCase))
                                            {
                                                BindInlineLatexBaselineProvenance(
                                                    document,
                                                    inserted,
                                                    target.InlineWordPosition,
                                                    target.InlineBottomWhitespacePoints);
                                            }
                                            return true;
                                        }
                                        finally
                                        {
                                            Release(inserted);
                                        }
                                    });

                            converted.FormulaCount++;
                            converted.FormulaIds.Add(
                                target.Metadata.FormulaId);
                        }

                        return converted;
                    });

            try
            {
                RestoreViewState(
                    document,
                    viewState,
                    preferredSelection: null);
            }
            catch
            {
                // UI state is not document correctness.
            }

            return result;
        }
        finally
        {
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    private List<CoreFormulaToLatexTarget>
        CaptureFormulaToLatexCoreTargets(
            Document document,
            Range scope,
            bool wholeDocument,
            string objectMode,
            bool readSemanticPayload)
    {
        var sourceKind =
            ObjectModeToHostKind(
                objectMode);
        if (sourceKind == WordFormulaHostKind.MathType)
            throw new NotSupportedException(
                "MathType formula-to-LaTeX remains owned by the MathType subsystem.");

        IReadOnlyList<WordFormulaHostDescriptor> hosts;
        if (!wholeDocument
            && scope.Start == scope.End)
        {
            var local =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    scope,
                    sourceKind);
            hosts = local is null
                ? Array.Empty<WordFormulaHostDescriptor>()
                : new[] { local };
        }
        else
        {
            var index =
                WordFormulaHostResolver.CaptureScopeIndex(
                    document,
                    scope);
            hosts = index.ForKind(
                sourceKind);
        }

        var result =
            new List<CoreFormulaToLatexTarget>();
        foreach (var captured in hosts
                     .OrderBy(item => item.Range.Start))
        {
            if (!wholeDocument
                && scope.Start != scope.End
                && (captured.Range.Start < scope.Start
                    || captured.Range.End > scope.End))
                continue;

            if (!readSemanticPayload)
            {
                result.Add(
                    new CoreFormulaToLatexTarget
                    {
                        Host = captured,
                    });
                continue;
            }

            var host = captured;
            host.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (host.Numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.Legacy)
                throw new InvalidDataException(
                    "所选公式仍使用旧版编号宿主，必须先迁移为标准编号结构后再转为 LaTeX。");

            FormulaMetadata metadata;
            if (sourceKind ==
                WordFormulaHostKind.VisualTeX)
            {
                var payload =
                    WordFormulaHostSemanticReader.Read(
                        document,
                        host);
                metadata =
                    CloneCoreMetadata(
                        payload.Metadata
                        ?? throw new InvalidDataException(
                            "VisualTeX OLE 没有可读取的内嵌公式元数据。"));
                metadata.Numbered =
                    host.Numbering.Numbered;
                metadata.DisplayMode =
                    host.DisplayMode;
            }
            else
            {
                var payload =
                    WordFormulaHostSemanticReader.Read(
                        document,
                        host);
                metadata =
                    BuildReadOnlyOmmlSessionMetadata(
                        document,
                        host,
                        payload);
            }

            if (!Guid.TryParse(
                    host.FormulaId,
                    out var stableId))
            {
                stableId = Guid.TryParse(
                        metadata.FormulaId,
                        out var metadataId)
                    ? metadataId
                    : Guid.NewGuid();
            }
            host.FormulaId =
                stableId.ToString("D");
            metadata.FormulaId =
                host.FormulaId;
            metadata.Numbered =
                host.Numbering.Numbered;
            metadata.DisplayMode =
                host.DisplayMode;
            metadata.Validate();

            var inlineWordPosition =
                sourceKind == WordFormulaHostKind.VisualTeX
                && string.Equals(
                    host.DisplayMode,
                    "inline",
                    StringComparison.OrdinalIgnoreCase)
                && metadata.WordInlineOlePositionPt.HasValue
                    ? (int?)Math.Round(
                        (double)metadata.WordInlineOlePositionPt.Value,
                        MidpointRounding.AwayFromZero)
                    : null;

            result.Add(
                new CoreFormulaToLatexTarget
                {
                    Host = host,
                    Metadata = metadata,
                    LatexSource =
                        BuildFormulaLatexSource(
                            metadata),
                    InlineWordPosition =
                        inlineWordPosition,
                    InlineBottomWhitespacePoints =
                        metadata.WordInlineSourceBottomWhitespacePt.HasValue
                            ? (float?)metadata
                                .WordInlineSourceBottomWhitespacePt
                                .Value
                            : null,
                });
        }

        return result;
    }
}

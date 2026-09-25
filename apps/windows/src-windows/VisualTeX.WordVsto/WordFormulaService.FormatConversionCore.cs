using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    public WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlan(
        bool wholeDocument,
        string sourceMode,
        string targetMode) =>
        CaptureFormulaFormatConversionPlanCore(
            wholeDocument,
            sourceMode,
            targetMode);

    public WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlan(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared) =>
        ApplyFormulaFormatConversionPlanCore(
            plan,
            prepared);

    internal WordFormulaFormatConversionPlan CaptureFormulaFormatConversionPlanCore(
        bool wholeDocument,
        string sourceMode,
        string targetMode)
    {
        ValidateCoreConversionPair(sourceMode, targetMode);

        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);

            scope = wholeDocument
                ? document.Content.Duplicate
                : (selection = _application.Selection).Range.Duplicate;

            var sourceKind = ObjectModeToHostKind(sourceMode);
            var targetKind = ObjectModeToHostKind(targetMode);
            IReadOnlyList<WordFormulaHostDescriptor> hosts;

            if (!wholeDocument
                && scope.Start == scope.End)
            {
                WordFormulaHostDescriptor? local =
                    sourceKind == WordFormulaHostKind.MathType
                        ? WordMathTypeHostAdapter.ResolveLocal(
                            _application,
                            document,
                            scope)
                        : WordFormulaHostResolver.ResolveLocal(
                            document,
                            scope,
                            sourceKind);
                hosts = local is null
                    ? Array.Empty<WordFormulaHostDescriptor>()
                    : new[] { local };
            }
            else if (sourceKind == WordFormulaHostKind.MathType)
            {
                hosts = WordMathTypeHostAdapter.CaptureScopeIndex(
                    _application,
                    document,
                    scope);
            }
            else
            {
                var index = WordFormulaHostResolver.CaptureScopeIndex(
                    document,
                    scope);
                hosts = index.ForKind(sourceKind);
            }

            var targetMathTypeNumberPosition =
                string.Equals(
                    targetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                    ? GetMathTypeNumberPositionPreference()
                    : "right";

            var plan = new WordFormulaFormatConversionPlan
            {
                DocumentId = DocumentIdentity(document),
                WritableValidated = true,
                SourceMode = sourceMode,
                TargetMode = targetMode,
                WholeDocument = wholeDocument,
                NumberFormatId =
                    WordEquationNumbering.GetEquationNumberFormatId(
                        document),
            };

            foreach (var host in hosts
                         .OrderBy(item => item.Range.Start))
            {
                var target = CaptureCoreConversionTarget(
                    document,
                    host,
                    sourceKind,
                    targetKind,
                    targetMathTypeNumberPosition);
                plan.Targets.Add(target);
            }

            return plan;
        }
        finally
        {
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    internal WordFormulaFormatConversionResult ApplyFormulaFormatConversionPlanCore(
        WordFormulaFormatConversionPlan plan,
        IReadOnlyDictionary<string, PreparedWordBulkFormula> prepared)
    {
        if (plan is null)
            throw new ArgumentNullException(nameof(plan));
        if (prepared is null)
            throw new ArgumentNullException(nameof(prepared));

        ValidateCoreConversionPair(
            plan.SourceMode,
            plan.TargetMode);

        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(
                    "No active Word document.");
            EnsureWritable(document);
            EnsureSourceDocument(
                document,
                plan.DocumentId);

            if (plan.Targets.Count == 0)
                return new WordFormulaFormatConversionResult();

            var missing = plan.Targets
                .Where(target => !prepared.ContainsKey(target.Id))
                .Select(target => target.Id)
                .ToArray();
            if (missing.Length > 0)
                throw new InvalidDataException(
                    "Prepared conversion output is missing for one or more formulas.");

            var sourceKind =
                ObjectModeToHostKind(plan.SourceMode);
            var targetKind =
                ObjectModeToHostKind(plan.TargetMode);

            // VisualTeX's numbered OLE uses a private VTEqNum_<FormulaId>
            // bookmark only while that host family exists. Pure OMML must remove
            // it, but existing user REF fields must not be stranded on the old
            // name. Capture external-reference counts before the transaction,
            // while the source bookmarks are still healthy. Generated visible
            // number REF fields are excluded by CaptureReferenceCounts.
            IReadOnlyDictionary<string, int> sourceReferenceCounts =
                sourceKind == WordFormulaHostKind.VisualTeX
                && targetKind == WordFormulaHostKind.Omml
                    ? WordEquationReferenceFields
                        .CaptureReferenceCounts(
                            document)
                    : new Dictionary<string, int>(
                        StringComparer.OrdinalIgnoreCase);

            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Convert Formula Format",
                () =>
                {
                    var result =
                        new WordFormulaFormatConversionResult();
                    var referenceTargetReplacements =
                        new Dictionary<string, string>(
                            StringComparer.OrdinalIgnoreCase);

                    foreach (var target in plan.Targets
                                 .OrderByDescending(
                                     item => item.SourceStart))
                    {
                        var preparedFormula =
                            prepared[target.Id];
                        var source =
                            WordFormulaOperationLocator.ResolveCapturedHost(
                                _application,
                                document,
                                target.SourceObjectId,
                                sourceKind,
                                target.SourceFormulaId);

                        WordFormulaHostDescriptor? replacedHost =
                            null;
                        if (targetKind is
                            WordFormulaHostKind.Omml
                            or WordFormulaHostKind.VisualTeX)
                        {
                            PrepareCoreTargetSession(
                                preparedFormula.Session,
                                plan,
                                target);
                            var request = BuildHostWriteRequest(
                                preparedFormula.Session,
                                preparedFormula.MathMl,
                                preparedFormula.PngPath,
                                preparedFormula.EmfPath);
                            replacedHost =
                                WordFormulaHostMutationKernel
                                    .ReplaceInActiveTransaction(
                                        _application,
                                        document,
                                        source,
                                        request);

                            if (sourceKind ==
                                    WordFormulaHostKind.VisualTeX
                                && targetKind ==
                                    WordFormulaHostKind.Omml
                                && target.Numbered
                                && Guid.TryParse(
                                    target.SourceFormulaId,
                                    out var sourceFormulaId))
                            {
                                var oldReferenceTarget =
                                    "VTEqNum_"
                                    + sourceFormulaId.ToString("N");
                                if (sourceReferenceCounts.TryGetValue(
                                        oldReferenceTarget,
                                        out var referenceCount)
                                    && referenceCount > 0)
                                {
                                    var nativeReferenceTarget =
                                        WordEquationReferenceFields
                                            .CreateNativeOmmlNumberReferenceBookmark(
                                                document,
                                                replacedHost);
                                    referenceTargetReplacements.Add(
                                        oldReferenceTarget,
                                        nativeReferenceTarget);
                                }
                            }
                        }
                        else if (targetKind ==
                                 WordFormulaHostKind.MathType)
                        {
                            if (sourceKind ==
                                WordFormulaHostKind.MathType)
                                throw new InvalidOperationException(
                                    "MathType-to-MathType is not a format conversion.");

                            var mathMl =
                                preparedFormula.MathMl
                                ?? target.SourceMathMl
                                ?? throw new InvalidDataException(
                                    $"MathType target '{target.Latex}' has no MathML.");

                            PrepareCoreTargetSession(
                                preparedFormula.Session,
                                plan,
                                target);

                            _ =
                                WordFormulaHostMutationKernel
                                    .ReplaceSourceWithExternalTargetInActiveTransaction(
                                        _application,
                                        document,
                                        source,
                                        insertion =>
                                        {
                                            var session =
                                                preparedFormula.Session;
                                            var oldSourceObjectId =
                                                session.SourceObjectId;
                                            try
                                            {
                                                session.SourceObjectId =
                                                    RangeReference(
                                                        insertion);
                                                var native =
                                                    preparedFormula
                                                        .MathTypeNativePreview;
                                                return InsertMathTypeOle(
                                                    session,
                                                    mathMl,
                                                    preparedFormula.EmfPath,
                                                    isolatedNativePreviewWmfPath:
                                                        native?.WmfPath,
                                                    isolatedNativePreviewWidthPt:
                                                        native?.WidthPt ?? 0,
                                                    isolatedNativePreviewHeightPt:
                                                        native?.HeightPt ?? 0,
                                                    isolatedNativePreviewWordPosition:
                                                        native?.WordPosition ?? 0,
                                                    isolatedNativePreviewAttempted:
                                                        preparedFormula
                                                            .MathTypeNativePreviewAttempted,
                                                    preserveExistingDisplayParagraphBoundary:
                                                        string.Equals(
                                                            target.DisplayMode,
                                                            "block",
                                                            StringComparison.OrdinalIgnoreCase),
                                                    preserveCapturedInsertion:
                                                        true);
                                            }
                                            finally
                                            {
                                                session.SourceObjectId =
                                                    oldSourceObjectId;
                                            }
                                        });
                        }
                        else
                        {
                            throw new NotSupportedException(
                                $"Unsupported conversion target {targetKind}.");
                        }

                        result.FormulaCount++;
                    }

                    // Format conversion is intentionally applied from the end of
                    // the document toward the start so captured source ranges stay
                    // stable. Numbered OMML/VisualTeX targets therefore cannot
                    // finalize their visible ordinals one host at a time: formulas
                    // that belong earlier in the document may not exist in the
                    // target family yet. Recompute the canonical Word order once,
                    // after the complete batch has materialized. This is the same
                    // numbering kernel used by the explicit refresh command; no
                    // conversion-specific numbering structure is introduced.
                    if (targetKind is
                            WordFormulaHostKind.Omml
                            or WordFormulaHostKind.VisualTeX
                        && plan.Targets.Any(target =>
                            target.Numbered))
                    {
                        _ =
                            WordFormulaNumberingKernel
                                .RefreshCanonicalNumbers(
                                    document);
                    }

                    if (referenceTargetReplacements.Count > 0)
                    {
                        _ =
                            WordEquationReferenceFields
                                .MigrateReferenceTargets(
                                    document,
                                    referenceTargetReplacements,
                                    sourceReferenceCounts);
                    }

                    return result;
                });
        }
        finally
        {
            Release(document);
        }
    }

    private WordFormulaFormatConversionTarget CaptureCoreConversionTarget(
        Document document,
        WordFormulaHostDescriptor host,
        WordFormulaHostKind sourceKind,
        WordFormulaHostKind targetKind,
        string targetMathTypeNumberPosition)
    {
        FormulaMetadata metadata;
        string latex;
        string? mathMl = null;
        WordFormulaNumberingDescriptor numbering;

        if (sourceKind == WordFormulaHostKind.Omml)
        {
            numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            host.Numbering = numbering;
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            metadata =
                BuildReadOnlyOmmlSessionMetadata(
                    document,
                    host,
                    payload);
            latex = payload.Latex;
            mathMl = payload.MathMl;
        }
        else if (sourceKind ==
                 WordFormulaHostKind.VisualTeX)
        {
            numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            host.Numbering = numbering;
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            metadata = CloneCoreMetadata(
                payload.Metadata
                ?? throw new InvalidDataException(
                    "A VisualTeX source has no embedded metadata."));
            metadata.DisplayMode =
                host.DisplayMode;

            // Embedded VisualTeX metadata is the durable user intent. Word can
            // damage/move the visible tab-number scaffold through bookmark
            // insertion gravity while the OLE itself still correctly says
            // Numbered=true. Preserve numbering when either the canonical
            // structure is healthy or the authoritative OLE metadata says the
            // display formula is numbered. Apply will reconcile a metadata-only
            // numbered source inside the same conversion transaction before
            // detaching/replacing it.
            metadata.Numbered =
                numbering.Numbered
                || (host.Display
                    && metadata.Numbered);
            metadata.Validate();
            latex = metadata.Latex;
        }
        else if (sourceKind ==
                 WordFormulaHostKind.MathType)
        {
            var semantic =
                WordMathTypeHostAdapter.ReadSemantic(
                    _application,
                    document,
                    host);
            metadata =
                CloneCoreMetadata(
                    semantic.Metadata);
            latex = metadata.Latex;
            mathMl = semantic.MathMl;
            numbering = host.Numbering;
        }
        else
        {
            throw new NotSupportedException(
                $"Unsupported conversion source {sourceKind}.");
        }

        if (string.IsNullOrWhiteSpace(latex))
            throw new InvalidDataException(
                "The source formula has no LaTeX representation.");

        var formulaId =
            Guid.TryParse(host.FormulaId, out var owned)
                ? owned.ToString("D")
                : Guid.TryParse(metadata.FormulaId, out var semanticId)
                    ? semanticId.ToString("D")
                    : Guid.NewGuid().ToString("D");
        metadata.FormulaId = formulaId;
        metadata.DisplayMode =
            host.DisplayMode;
        metadata.Numbered =
            numbering.Numbered
            || (sourceKind ==
                    WordFormulaHostKind.VisualTeX
                && host.Display
                && metadata.Numbered);
        metadata.Validate();

        var semanticFontSize =
            FormulaFontSize.ResolveSemanticFontSize(
                metadata);
        var fontSize =
            semanticFontSize;

        if (sourceKind == WordFormulaHostKind.MathType
            && string.Equals(
                host.DisplayMode,
                "inline",
                StringComparison.OrdinalIgnoreCase)
            && targetKind is WordFormulaHostKind.Omml
                or WordFormulaHostKind.VisualTeX)
        {
            InlineShape? sourceShape = null;
            try
            {
                sourceShape = FindMathTypeOleByRange(
                    document,
                    RangeReferenceFromAddress(host.Range),
                    allowGlobalFallback: false);
                if (sourceShape is not null)
                {
                    var wordPresentationFontSize =
                        ReadMathTypeInlinePresentationFontSize(
                            sourceShape);
                    if (wordPresentationFontSize is > 0)
                    {
                        // For both pure OMML and VisualTeX, preserve the font
                        // size the user actually sees in Word. MathType's MTEF
                        // Full size can differ from its Word OLE presentation
                        // size (for example 12 pt internally but 10.5 pt in the
                        // document). Persisting the hidden MTEF size into
                        // VisualTeX metadata makes the next edit/apply jump back
                        // to 12 pt. Store the Word presentation size as the
                        // target's durable font size instead.
                        fontSize =
                            wordPresentationFontSize.Value;
                    }
                }
            }
            finally
            {
                Release(sourceShape);
            }
        }

        return new WordFormulaFormatConversionTarget
        {
            Id = Guid.NewGuid().ToString("D"),
            SourceFormulaId = formulaId,
            SourceObjectId =
                RangeReferenceFromAddress(host.Range),
            SourceStart = host.Range.Start,
            Latex = latex,
            SourceMathMl = mathMl,
            SourceIsManagedOmml =
                sourceKind == WordFormulaHostKind.Omml
                && !string.IsNullOrWhiteSpace(host.FormulaId),
            SourceWithinTable = host.WithinTable,
            DisplayMode = host.DisplayMode,
            Numbered =
                numbering.Numbered
                || (sourceKind == WordFormulaHostKind.VisualTeX
                    && host.Display
                    && metadata.Numbered),
            MathTypeNumberPosition =
                sourceKind == WordFormulaHostKind.MathType
                    ? GetMathTypeNumberPositionForRange(
                        RangeReferenceFromAddress(host.Range))
                    : targetMathTypeNumberPosition,
            FontSizePt = fontSize,
            Metadata = metadata,
        };
    }

    private static float? ReadMathTypeInlinePresentationFontSize(
        InlineShape shape)
    {
        Range? shapeRange = null;
        Range? probe = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Document? document = null;
        try
        {
            shapeRange = shape.Range;
            document = shapeRange.Document;
            for (var position = shapeRange.Start;
                 position < shapeRange.End;
                 position++)
            {
                Release(font);
                font = null;
                Release(probe);
                probe = document.Range(
                    position,
                    position + 1);
                if (!string.Equals(
                        probe.Text,
                        "\u0001",
                        StringComparison.Ordinal))
                    continue;

                font = probe.Font;
                return TryNormalizeDefinedWordFontSize(
                        font.Size,
                        out var size)
                    ? size
                    : null;
            }

            return null;
        }
        catch
        {
            return null;
        }
        finally
        {
            Release(font);
            Release(probe);
            Release(document);
            Release(shapeRange);
        }
    }

    private static void PrepareCoreTargetSession(
        OfficeSessionDocument session,
        WordFormulaFormatConversionPlan plan,
        WordFormulaFormatConversionTarget target)
    {
        session.Mode = "create";
        session.Host = "word";
        session.FormulaId =
            target.SourceFormulaId;
        session.SourceDocumentId =
            plan.DocumentId;
        session.SourceObjectId =
            target.SourceObjectId;
        session.DisplayMode =
            target.DisplayMode;
        session.ObjectMode =
            plan.TargetMode;
        session.Numbered =
            target.Numbered;
        session.MathTypeNumberPosition =
            target.MathTypeNumberPosition;
        session.FontSizePt =
            target.FontSizePt;
        session.OriginalMetadata =
            CloneCoreMetadata(
                target.Metadata);
    }

    private static FormulaMetadata CloneCoreMetadata(
        FormulaMetadata metadata) =>
        FormulaMetadataCodec.DeserializeJson(
            FormulaMetadataCodec.SerializeJson(
                metadata))
        ?? throw new InvalidDataException(
            "Unable to clone formula metadata for conversion.");

    private static void ValidateCoreConversionPair(
        string sourceMode,
        string targetMode)
    {
        _ = ObjectModeToHostKind(sourceMode);
        _ = ObjectModeToHostKind(targetMode);
        if (string.Equals(
                sourceMode,
                targetMode,
                StringComparison.Ordinal))
            throw new InvalidOperationException(
                "The source and target formula formats are identical.");
    }
}

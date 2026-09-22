using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Canonical numbering writer for OMML/VisualTeX hosts.
///
/// OMML numbering is pure Word state: one wdOMathDisplay using Word's native
/// m:eqArr + # + SEQ/STYLEREF structure. It never receives a VisualTeX identity
/// or numbering bookmark.
///
/// VisualTeX OLE cannot live inside Word's native OMath numbering host. In the
/// document body its canonical host is therefore one ordinary Word paragraph
/// with a center tab before the OLE and a right tab before the visible number;
/// VTEqNum_<FormulaId> remains the stable reference target. A 1x3 table is
/// compatibility input only, never the body producer for a new VisualTeX host.
/// </summary>
internal static class WordFormulaNumberingKernel
{
    private const string SequenceName = "VisualTeXEquation";
    private const string NumberBookmarkPrefix = "VTEqNum_";
    private const float SideColumnWidthPoints = 60f;

    private sealed class NumberPlan
    {
        internal int Ordinal { get; set; }
        internal string Prefix { get; set; } = string.Empty;
    }

    internal static WordFormulaHostDescriptor RekeyPastedNumbering(
        Document document,
        WordFormulaHostDescriptor host,
        string? copiedVisibleNumber)
    {
        if (!host.Display)
            return host;

        if (host.Kind == WordFormulaHostKind.Omml)
        {
            // Native Word paste already copied the correct m:eqArr/#(SEQ) host.
            // Pure OMML has nothing to re-key; validate and keep Word's structure.
            return WordNativeOmmlNumbering.RekeyPastedNativeHost(
                document,
                host);
        }

        if (!Guid.TryParse(host.FormulaId, out _))
            throw new InvalidDataException(
                "A pasted numbered VisualTeX formula requires a fresh UUID FormulaId.");

        var (ordinal, prefix) = ParseVisibleNumberPlan(
            copiedVisibleNumber);

        var current = WordFormulaNumberingResolver.ResolveLocal(
            document,
            host);
        if (current.ContainerKind is
            WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
            or WordFormulaNumberingContainerKind.CanonicalBodyTable
            or WordFormulaNumberingContainerKind.CanonicalUserTableCell)
        {
            RewriteCanonicalNumberLabel(
                document,
                host,
                current,
                ordinal,
                prefix);
            var refreshed = WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
            return WithNumbering(host, refreshed);
        }

        if (TryRekeyPastedBodyTable(
                document,
                host,
                ordinal,
                prefix,
                out var bodyHost))
            return bodyHost;

        if (host.WithinTable
            && TryRekeyPastedUserCell(
                document,
                host,
                ordinal,
                prefix,
                out var cellHost))
            return cellHost;

        // The user copied only the formula host, not its number scaffold. Build
        // one canonical local container, then restore the copied visible number
        // without renumbering unrelated formulas.
        var attached = AttachCanonical(
            document,
            host);
        var attachedNumbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                attached);
        RewriteCanonicalNumberLabel(
            document,
            attached,
            attachedNumbering,
            ordinal,
            prefix);
        var finalNumbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                attached);
        return WithNumbering(
            attached,
            finalNumbering);
    }

    internal static int RefreshCanonicalNumbers(
        Document document,
        string? requestedFormatId = null)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));

        var changeNumberFormat =
            !string.IsNullOrWhiteSpace(requestedFormatId);
        var formatId = EquationNumberFormat.Resolve(
            requestedFormatId
            ?? WordEquationNumbering.GetEquationNumberFormatId(document)).Id;
        var format = EquationNumberFormat.Resolve(formatId);

        // If the document contains no Equation/VisualTeX SEQ field at all,
        // there cannot be any canonical numbered OMML/VisualTeX host to refresh.
        // Avoid building a full host index and re-proving 1000 unnumbered
        // formulas one-by-one.
        if (!HasAnyCanonicalEquationSequenceField(
                document))
        {
            WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                document,
                formatId);
            return 0;
        }

        var index =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var hosts = index.Omml
            .Concat(index.VisualTeX)
            .OrderBy(host => host.Range.Start)
            .ToArray();

        var numbered = new List<(
            WordFormulaHostDescriptor Host,
            WordFormulaNumberingDescriptor Numbering)>();
        foreach (var host in hosts)
        {
            var descriptor =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (descriptor.ContainerKind ==
                WordFormulaNumberingContainerKind.Legacy)
                throw new InvalidDataException(
                    "文档中仍存在旧版 VisualTeX 编号结构；请先执行编号结构迁移，再刷新编号。");
            if (!descriptor.Numbered)
                continue;
            numbered.Add((host, descriptor));
        }

        if (numbered.Count == 0)
        {
            WordEquationNumbering.RestoreEquationNumberFormatForConversion(
                document,
                formatId);
            return 0;
        }

        var positions = numbered
            .Select(item => item.Host.Range.Start)
            .ToArray();
        var scopes = format.UsesHeading
            ? WordEquationNumbering.CaptureHeadingScopesAtPositions(
                document,
                formatId,
                positions)
            : new Dictionary<int, ResolvedEquationHeadingScope>();

        var scopeOrdinals =
            new Dictionary<int, int>();
        var plans = new List<(
            WordFormulaHostDescriptor Host,
            WordFormulaNumberingDescriptor Numbering,
            int Ordinal,
            string Prefix)>(numbered.Count);

        foreach (var item in numbered)
        {
            var scopeStart = int.MinValue;
            var prefix = string.Empty;
            if (format.UsesHeading
                && scopes.TryGetValue(
                    item.Host.Range.Start,
                    out var scope))
            {
                scopeStart = scope.ScopeStart;
                prefix =
                    scope.NumberText
                    + format.Separator;
            }

            var ordinal = scopeOrdinals.TryGetValue(
                    scopeStart,
                    out var previous)
                ? previous + 1
                : 1;
            scopeOrdinals[scopeStart] = ordinal;
            plans.Add((
                item.Host,
                item.Numbering,
                ordinal,
                prefix));
        }

        WordEquationNumbering.RestoreEquationNumberFormatForConversion(
            document,
            formatId);

        var targetBookmarks =
            new HashSet<string>(
                StringComparer.OrdinalIgnoreCase);

        foreach (var plan in plans
                     .OrderByDescending(
                         item => item.Host.Range.Start))
        {
            if (plan.Numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.CanonicalNativeOmml)
            {
                var nativeHost =
                    plan.Host;
                if (!changeNumberFormat)
                {
                    // Ordinary "update equation numbers" is a Word-field refresh.
                    // Never rebuild or adopt an anonymous pasted native OMath just
                    // to let SEQ/STYLEREF calculate their current results.
                    WordNativeOmmlNumbering.UpdateNativeFields(
                        document,
                        nativeHost);
                }
                else
                {
                    // Number-format changes rebuild only Word's native
                    // SEQ/STYLEREF topology. OMML has no durable VisualTeX ID.
                    _ = WordNativeOmmlNumbering.RebuildNumbered(
                        document,
                        nativeHost);
                }
            }
            else if (plan.Numbering.ContainerKind ==
                         WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                     && plan.Host.Kind == WordFormulaHostKind.VisualTeX
                     && WordVisualTeXParagraphNumbering.IsSelfContainedHost(
                         document,
                         plan.Host))
            {
                WordVisualTeXParagraphNumbering.Refresh(
                    document,
                    plan.Host,
                    rebuildForCurrentFormat:
                        changeNumberFormat);
            }
            else
            {
                RewriteCanonicalNumberLabel(
                    document,
                    plan.Host,
                    plan.Numbering,
                    plan.Ordinal,
                    plan.Prefix);
            }

            if (Guid.TryParse(
                    plan.Host.FormulaId,
                    out var formulaId))
                targetBookmarks.Add(
                    NumberBookmarkPrefix
                    + formulaId.ToString("N"));
        }

        if (targetBookmarks.Count > 0)
        {
            _ = WordEquationReferenceFields.UpdateNavigableReferences(
                document,
                targetBookmarks);
        }

        return plans.Count;
    }

    internal static int SetCanonicalNumberFormat(
        Document document,
        string formatId)
    {
        var resolved =
            EquationNumberFormat.Resolve(
                formatId).Id;
        var count =
            RefreshCanonicalNumbers(
                document,
                resolved);
        WordEquationNumbering.SetDefaultEquationNumberFormatPreference(
            resolved);
        return count;
    }

    internal static IReadOnlyList<EquationReferenceTarget>
        GetCanonicalReferenceTargets(
            Document document)
    {
        var index =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var targets =
            new List<EquationReferenceTarget>();
        var nativeEquationItems =
            CaptureNativeEquationReferenceItems(
                document);

        foreach (var host in index.Omml
                     .Concat(index.VisualTeX)
                     .OrderBy(item => item.Range.Start))
        {
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (!numbering.Numbered
                || numbering.ContainerKind is not
                    (WordFormulaNumberingContainerKind.CanonicalNativeOmml
                     or WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                     or WordFormulaNumberingContainerKind.CanonicalBodyTable
                     or WordFormulaNumberingContainerKind.CanonicalUserTableCell)
                || numbering.NumberRange is null)
                continue;

            Range? numberRange = null;
            try
            {
                numberRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        numbering.NumberRange);
                var numberText =
                    host.Kind == WordFormulaHostKind.VisualTeX
                    && numbering.ContainerKind ==
                        WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                        ? WordVisualTeXParagraphNumbering
                            .ReadVisibleNumberText(
                                document,
                                host)
                        : (numberRange.Text ?? string.Empty)
                            .Trim()
                            .Trim('(', ')');
                var latexPreview =
                    ReadLatexPreview(
                        document,
                        host);

                if (host.Kind == WordFormulaHostKind.Omml)
                {
                    if (numbering.ContainerKind !=
                        WordFormulaNumberingContainerKind.CanonicalNativeOmml)
                        continue;

                    var nativeReferenceItem =
                        ResolveNativeEquationReferenceItem(
                            document,
                            host,
                            nativeEquationItems);
                    if (nativeReferenceItem <= 0)
                        continue;

                    targets.Add(
                        new EquationReferenceTarget(
                            formulaId: string.Empty,
                            nativeReferenceItem,
                            numberText,
                            latexPreview,
                            host.Range.Start,
                            EquationReferenceSource.WordOmml,
                            host.Range.End));
                    continue;
                }

                if (!Guid.TryParse(
                        host.FormulaId,
                        out var formulaId))
                    continue;

                targets.Add(
                    new EquationReferenceTarget(
                        formulaId.ToString("D"),
                        nativeReferenceItem: 0,
                        numberText,
                        latexPreview,
                        host.Range.Start,
                        EquationReferenceSource.VisualTeX,
                        host.Range.End));
            }
            finally
            {
                Release(numberRange);
            }
        }

        return targets;
    }

    private static IReadOnlyDictionary<int, int>
        CaptureNativeEquationReferenceItems(
            Document document)
    {
        var result =
            new Dictionary<int, int>();
        var sequenceName =
            ResolveNativeEquationSequenceName(
                document);
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? fieldResult = null;
        try
        {
            fields = document.Fields;
            var nativeItem = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(fieldResult);
                fieldResult = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code.Duplicate;
                if (!IsNativeEquationSequenceField(
                        code.Text,
                        sequenceName))
                    continue;

                fieldResult =
                    field.Result.Duplicate;
                if (fieldResult.StoryType !=
                    WdStoryType.wdMainTextStory)
                    continue;

                nativeItem++;
                if (result.ContainsKey(
                        fieldResult.Start))
                    throw new InvalidDataException(
                        "Two native Equation SEQ fields share the same Word result boundary.");
                result.Add(
                    fieldResult.Start,
                    nativeItem);
            }
            return result;
        }
        finally
        {
            Release(fieldResult);
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static int ResolveNativeEquationReferenceItem(
        Document document,
        WordFormulaHostDescriptor host,
        IReadOnlyDictionary<int, int> nativeEquationItems)
    {
        Range? hostRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            fields = hostRange.Fields;
            var sequenceName =
                ResolveNativeEquationSequenceName(
                    document);
            var matched = 0;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code.Duplicate;
                if (!IsNativeEquationSequenceField(
                        code.Text,
                        sequenceName))
                    continue;

                result =
                    field.Result.Duplicate;
                if (!nativeEquationItems.TryGetValue(
                        result.Start,
                        out var candidate))
                    return 0;
                if (matched != 0
                    && matched != candidate)
                    throw new InvalidDataException(
                        "One native OMML equation resolves to multiple Word Equation cross-reference items.");
                matched = candidate;
            }
            return matched;
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(hostRange);
        }
    }

    private static string ResolveNativeEquationSequenceName(
        Document document)
    {
        CaptionLabels? labels = null;
        CaptionLabel? label = null;
        try
        {
            labels =
                document.Application.CaptionLabels;
            label =
                labels[WdCaptionLabelID.wdCaptionEquation];
            var name =
                (label.Name ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(name)
                ? "Equation"
                : name;
        }
        catch
        {
            return "Equation";
        }
        finally
        {
            Release(label);
            Release(labels);
        }
    }

    private static bool IsNativeEquationSequenceField(
        string? instruction,
        string sequenceName)
    {
        if (string.IsNullOrWhiteSpace(instruction)
            || string.IsNullOrWhiteSpace(sequenceName))
            return false;

        var match =
            Regex.Match(
                instruction!.Trim(),
                @"^SEQ\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))(?=\s|\\|$)",
                RegexOptions.IgnoreCase
                | RegexOptions.CultureInvariant);
        if (!match.Success)
            return false;
        var identifier =
            match.Groups["quoted"].Success
                ? match.Groups["quoted"].Value
                : match.Groups["plain"].Value;
        return string.Equals(
            identifier,
            sequenceName,
            StringComparison.OrdinalIgnoreCase);
    }

    internal static string ReferenceBookmarkName(
        string formulaId)
    {
        if (!Guid.TryParse(
                formulaId,
                out var parsed))
            throw new InvalidDataException(
                "Equation reference requires a UUID FormulaId.");
        return NumberBookmarkPrefix
            + parsed.ToString("N");
    }

    internal static WordFormulaHostDescriptor RecoverVisualTeXNumberingForConversion(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (document is null)
            throw new ArgumentNullException(
                nameof(document));
        if (host is null
            || host.Kind !=
                WordFormulaHostKind.VisualTeX
            || !host.Display)
            throw new ArgumentException(
                "VisualTeX numbering recovery requires one display VisualTeX host.",
                nameof(host));

        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            shape =
                GetExactVisualTeXNumberingShape(
                    document,
                    host);
            var metadata =
                WordFormulaMetadataReader
                    .TryReadEmbeddedNativeOle(
                        shape)
                ?? throw new InvalidDataException(
                    "The VisualTeX host has no authoritative metadata for numbering recovery.");
            if (!metadata.Numbered)
                throw new InvalidDataException(
                    "VisualTeX numbering recovery was requested for metadata that is explicitly unnumbered.");

            metadata.DisplayMode = "block";
            metadata.Numbered = true;
            metadata.Validate();

            // Legacy numbered VisualTeX documents can have a VTEqCap bookmark
            // expanded by Word bookmark gravity over later body text/OMML while
            // VTEqNum remains attached to the actual hidden SEQ result. Repair
            // only that ownership alias before invoking the mature legacy
            // ReconcileFormula path; this never deletes or rewrites user content.
            WordEquationNumbering
                .RebindExternalCaptionToNativeNumberParagraph(
                    document,
                    metadata.FormulaId);

            shapeRange =
                shape.Range.Duplicate;
            WordEquationNumbering.ReconcileFormula(
                document,
                shapeRange,
                shape.Height,
                metadata,
                numberingOrderMayHaveChanged: false);

            Release(shapeRange);
            shapeRange =
                shape.Range.Duplicate;
            var refreshed =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX)
                ?? throw new InvalidDataException(
                    "The VisualTeX host disappeared while recovering its numbered layout.");
            refreshed.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    refreshed);
            if (!refreshed.Numbering.Numbered
                || refreshed.Numbering.ContainerKind is not
                    (WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
                     or WordFormulaNumberingContainerKind.CanonicalBodyTable
                     or WordFormulaNumberingContainerKind.CanonicalUserTableCell))
                throw new InvalidDataException(
                    "The damaged VisualTeX numbered host did not recover to a canonical numbering container.");

            return refreshed;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    internal static WordFormulaHostDescriptor AttachCanonical(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (!host.Display)
            throw new InvalidOperationException(
                "Only display equations can receive an equation number.");
        if (host.Kind == WordFormulaHostKind.Omml)
            return WordNativeOmmlNumbering.Attach(
                document,
                host);

        if (!Guid.TryParse(host.FormulaId, out _))
            throw new InvalidDataException(
                "A numbered VisualTeX OLE requires one stable FormulaId.");

        var current = WordFormulaNumberingResolver.ResolveLocal(document, host);
        if (current.ContainerKind ==
            WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph)
            return WithNumbering(host, current);
        if (current.ContainerKind ==
            WordFormulaNumberingContainerKind.CanonicalBodyTable)
            return WithNumbering(host, current);
        if (current.ContainerKind ==
            WordFormulaNumberingContainerKind.CanonicalUserTableCell)
            return WithNumbering(host, current);
        if (current.ContainerKind ==
            WordFormulaNumberingContainerKind.Legacy)
            throw new InvalidDataException(
                "The formula uses a retired numbering topology and must be canonicalized before numbering is changed.");

        return host.WithinTable
            ? AttachInsideUserTableCell(document, host)
            : WordVisualTeXParagraphNumbering.Attach(
                document,
                host);
    }

    internal static WordFormulaHostDescriptor DetachCanonical(
        Document document,
        WordFormulaHostDescriptor host)
    {
        var numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
        return numbering.ContainerKind switch
        {
            WordFormulaNumberingContainerKind.None => WithNumbering(
                host,
                numbering),
            WordFormulaNumberingContainerKind.CanonicalNativeOmml =>
                WordNativeOmmlNumbering.Detach(
                    document,
                    host),
            WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph =>
                WordVisualTeXParagraphNumbering.IsSelfContainedHost(
                    document,
                    host)
                    ? WordVisualTeXParagraphNumbering.Detach(
                        document,
                        host)
                    : DetachBodyTabParagraph(
                        document,
                        host),
            WordFormulaNumberingContainerKind.CanonicalBodyTable =>
                DetachBodyTable(document, host, numbering),
            WordFormulaNumberingContainerKind.CanonicalUserTableCell =>
                DetachUserTableCell(document, host, numbering),
            _ => throw new InvalidDataException(
                "A retired numbering topology cannot be detached by the canonical writer."),
        };
    }

    internal static void ValidatePreservedContainer(
        Document document,
        WordFormulaHostDescriptor host,
        bool shouldBeNumbered)
    {
        var numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
        if (!shouldBeNumbered)
        {
            if (numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.None)
                throw new InvalidDataException(
                    "The unnumbered target retained a numbering container.");
            return;
        }

        if (numbering.ContainerKind is not
            (WordFormulaNumberingContainerKind.CanonicalNativeOmml
             or WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
             or WordFormulaNumberingContainerKind.CanonicalBodyTable
             or WordFormulaNumberingContainerKind.CanonicalUserTableCell))
            throw new InvalidDataException(
                "The numbered target is not inside a canonical numbering container.");
    }

    private static WordFormulaHostDescriptor AttachBodyTabParagraph(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX)
            throw new InvalidDataException(
                "Only VisualTeX OLE uses the canonical body tab paragraph.");

        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            shape = GetExactVisualTeXNumberingShape(
                document,
                host);
            var metadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                    shape)
                ?? throw new InvalidDataException(
                    "The numbered VisualTeX OLE has no authoritative metadata.");
            metadata.DisplayMode = "block";
            metadata.Numbered = true;
            metadata.Validate();

            shapeRange = shape.Range.Duplicate;
            WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                document,
                shapeRange,
                shape.Height,
                metadata);

            Release(shapeRange);
            shapeRange = shape.Range.Duplicate;
            var refreshed =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX)
                ?? throw new InvalidDataException(
                    "The VisualTeX OLE disappeared while its tab numbering host was created.");
            refreshed.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    refreshed);
            if (refreshed.Numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph)
                throw new InvalidDataException(
                    "VisualTeX body numbering did not converge to the canonical center/right-tab paragraph.");
            return refreshed;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    private static WordFormulaHostDescriptor DetachBodyTabParagraph(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX)
            throw new InvalidDataException(
                "Only VisualTeX OLE can detach a body tab numbering paragraph.");

        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            shape = GetExactVisualTeXNumberingShape(
                document,
                host);
            var metadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                    shape)
                ?? throw new InvalidDataException(
                    "The numbered VisualTeX OLE has no authoritative metadata while detaching numbering.");
            metadata.DisplayMode = "block";
            metadata.Numbered = false;
            metadata.Validate();

            shapeRange = shape.Range.Duplicate;
            WordEquationNumbering.ReconcileFormula(
                document,
                shapeRange,
                shape.Height,
                metadata,
                numberingOrderMayHaveChanged: false);

            Release(shapeRange);
            shapeRange = shape.Range.Duplicate;
            var refreshed =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    shapeRange,
                    WordFormulaHostKind.VisualTeX)
                ?? throw new InvalidDataException(
                    "The VisualTeX OLE disappeared while detaching its tab numbering host.");
            refreshed.Numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    refreshed);
            if (refreshed.Numbering.Numbered
                || refreshed.Numbering.ContainerKind !=
                    WordFormulaNumberingContainerKind.None)
                throw new InvalidDataException(
                    "VisualTeX tab numbering artifacts survived canonical detach.");
            return refreshed;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    private static WordFormulaHostDescriptor AttachBodyTable(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? sourceRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        Range? suffix = null;
        Range? formattedSource = null;
        Range? documentContent = null;
        Range? nextProbe = null;
        Range? boundarySeed = null;
        Range? tableAnchor = null;
        Table? table = null;
        Cell? centerCell = null;
        Range? centerRange = null;
        Range? centerInsertion = null;
        Range? sourceDeleteRange = null;
        var sourceMathCount = -1;
        var sourceShapeCount = -1;
        var sourceTextDiagnostic = string.Empty;
        try
        {
            sourceRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            paragraphs = sourceRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A numbered display formula must occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (IsWithinTable(paragraphRange))
                throw new InvalidOperationException(
                    "A body numbering table cannot be nested inside a user table.");

            prefix = document.Range(paragraphRange.Start, sourceRange.Start);
            var bodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            suffix = document.Range(sourceRange.End, bodyEnd);
            if (!ContainsOnlyStructuralText(prefix.Text)
                || !ContainsOnlyStructuralText(suffix.Text))
                throw new InvalidDataException(
                    "A numbered display formula cannot consume user prose from its paragraph.");

            var plan = ResolveNumberPlan(
                document,
                sourceRange.Start,
                host.FormulaId!);
            sourceMathCount = sourceRange.OMaths.Count;
            sourceShapeCount = sourceRange.InlineShapes.Count;
            if (sourceMathCount > 0)
            {
                OMath? sourceMath = null;
                Range? sourceMathRange = null;
                try
                {
                    sourceMath = sourceRange.OMaths[1];
                    sourceMathRange = sourceMath.Range.Duplicate;
                    sourceTextDiagnostic =
                        $"math={sourceMath.Type}:{sourceMathRange.Start}:{sourceMathRange.End};";
                }
                finally
                {
                    Release(sourceMathRange);
                    Release(sourceMath);
                }
            }
            sourceTextDiagnostic +=
                (sourceRange.Text ?? string.Empty)
                .Replace("\r", "\\r")
                .Replace("\a", "\\a")
                .Replace("\t", "\\t");
            formattedSource = sourceRange.FormattedText;

            var sourceParagraphStart = paragraphRange.Start;
            var tablePosition = paragraphRange.End;

            documentContent = document.Content;
            var followedByTable = false;
            if (tablePosition < documentContent.End)
            {
                nextProbe = document.Range(
                    tablePosition,
                    Math.Min(documentContent.End, tablePosition + 1));
                followedByTable = IsWithinTable(nextProbe);
            }
            Release(nextProbe); nextProbe = null;
            Release(documentContent); documentContent = null;

            if (followedByTable)
            {
                boundarySeed = document.Range(bodyEnd, bodyEnd);
                boundarySeed.InsertBefore("\r\r");
            }
            else
            {
                paragraphRange.InsertParagraphAfter();
            }

            Release(documentContent);
            documentContent = document.Content;
            tablePosition = Math.Max(
                documentContent.Start,
                Math.Min(tablePosition, documentContent.End - 1));
            tableAnchor = document.Range(
                tablePosition,
                Math.Min(documentContent.End, tablePosition + 1));

            if (IsWithinTable(tableAnchor))
                throw new InvalidDataException(
                    "Word placed the VisualTeX number table anchor inside a user table.");

            NormalizeEmptyTableAnchor(tableAnchor);
            table = document.Tables.Add(tableAnchor, 1, 3);
            ConfigureCanonicalTable(document, table);

            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range.Duplicate;
            centerInsertion = centerRange.Duplicate;
            centerInsertion.End = Math.Max(
                centerInsertion.Start,
                centerInsertion.End - 1);
            centerInsertion.Collapse(WdCollapseDirection.wdCollapseStart);
            centerInsertion.FormattedText = formattedSource.FormattedText;

            var copied = ResolveOneHostInScope(
                document,
                centerRange,
                host.Kind,
                host.FormulaId);
            EnsureSemanticIdentityMatches(host, copied);

            Release(documentContent);
            documentContent = document.Content;
            var liveTableRange = table.Range;
            try
            {
                sourceDeleteRange = document.Range(
                    sourceParagraphStart,
                    Math.Max(
                        sourceParagraphStart,
                        liveTableRange.Start - 1));
            }
            finally { Release(liveTableRange); }
            sourceDeleteRange.Delete();

            Release(centerRange);
            centerRange = centerCell.Range.Duplicate;
            copied = ResolveOneHostInScope(
                document,
                centerRange,
                host.Kind,
                host.FormulaId);
            BindHostIdentity(document, copied);

            CreateNumberCell(
                document,
                table,
                copied,
                plan.Ordinal,
                plan.Prefix);
            ConfigureCanonicalTable(document, table);

            var numbering = WordFormulaNumberingResolver.ResolveLocal(
                document,
                copied);
            if (numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalBodyTable)
                throw new InvalidDataException(
                    "The new numbered formula did not produce the canonical 1x3 container. "
                    + $"sourceMaths={sourceMathCount} sourceShapes={sourceShapeCount} sourceText=[{sourceTextDiagnostic}] "
                    + DescribeBodyTableState(
                        table,
                        copied));

            return WithNumbering(copied, numbering);
        }
        finally
        {
            Release(sourceDeleteRange);
            Release(centerInsertion);
            Release(centerRange);
            Release(centerCell);
            Release(table);
            Release(tableAnchor);
            Release(boundarySeed);
            Release(nextProbe);
            Release(documentContent);
            Release(formattedSource);
            Release(suffix);
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(sourceRange);
        }
    }

    private static WordFormulaHostDescriptor DetachBodyTable(
        Document document,
        WordFormulaHostDescriptor host,
        WordFormulaNumberingDescriptor numbering)
    {
        if (numbering.ContainerRange is null)
            throw new InvalidDataException(
                "The canonical number table has no captured container range.");

        Range? hostRange = null;
        Range? formattedHost = null;
        Range? tableRange = null;
        Tables? tables = null;
        Table? table = null;
        Range? insertion = null;
        Paragraphs? insertedParagraphs = null;
        Paragraph? insertedParagraph = null;
        Range? insertedParagraphRange = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            formattedHost = hostRange.FormattedText;
            tableRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                numbering.ContainerRange);
            tables = tableRange.Tables;
            if (tables.Count != 1)
                throw new InvalidDataException(
                    "The canonical number table moved before detaching.");
            table = tables[1];

            var liveTable = table.Range;
            try { liveTable.InsertParagraphBefore(); }
            finally { Release(liveTable); }

            var tableStartRange = table.Range;
            var tableStart = tableStartRange.Start;
            Release(tableStartRange);
            insertion = document.Range(
                Math.Max(0, tableStart - 1),
                tableStart);
            insertedParagraphs = insertion.Paragraphs;
            if (insertedParagraphs.Count < 1)
                throw new InvalidDataException(
                    "Word did not create a body paragraph before the numbered table.");
            insertedParagraph = insertedParagraphs[1];
            insertedParagraphRange = insertedParagraph.Range.Duplicate;
            if (IsWithinTable(insertedParagraphRange))
                throw new InvalidDataException(
                    "The detached formula paragraph remained inside the number table.");

            var bodyEnd = Math.Max(
                insertedParagraphRange.Start,
                insertedParagraphRange.End - 1);
            var body = document.Range(
                insertedParagraphRange.Start,
                bodyEnd);
            try { body.FormattedText = formattedHost.FormattedText; }
            finally { Release(body); }

            var copied = ResolveOneHostInScope(
                document,
                insertedParagraphRange,
                host.Kind,
                host.FormulaId);
            EnsureSemanticIdentityMatches(host, copied);

            // Range.Delete can clear table contents while leaving Word's
            // structural table object behind. The formula has already been
            // materialized in the body paragraph above, so delete the canonical
            // VisualTeX numbering container with Word's native Table.Delete().
            table.Delete();

            Release(insertedParagraphRange);
            insertedParagraphRange = insertedParagraph.Range.Duplicate;
            copied = ResolveOneHostInScope(
                document,
                insertedParagraphRange,
                host.Kind,
                host.FormulaId);
            BindHostIdentity(document, copied);
            ConfigureUnnumberedDisplayParagraph(insertedParagraphRange);

            return WithNumbering(
                copied,
                new WordFormulaNumberingDescriptor
                {
                    Numbered = false,
                    FormulaId = copied.FormulaId,
                    ContainerKind =
                        WordFormulaNumberingContainerKind.None,
                });
        }
        finally
        {
            Release(insertedParagraphRange);
            Release(insertedParagraph);
            Release(insertedParagraphs);
            Release(insertion);
            Release(table);
            Release(tables);
            Release(tableRange);
            Release(formattedHost);
            Release(hostRange);
        }
    }

    private static WordFormulaHostDescriptor AttachInsideUserTableCell(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? hostRange = null;
        Cells? cells = null;
        Cell? cell = null;
        Range? cellRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? before = null;
        Range? after = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            cells = hostRange.Cells;
            if (cells.Count != 1)
                throw new InvalidDataException(
                    "A numbered formula inside a user table must belong to one cell.");
            cell = cells[1];
            cellRange = cell.Range.Duplicate;
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A numbered table-cell formula must occupy one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;

            before = document.Range(paragraphRange.Start, hostRange.Start);
            // Word exposes the cell-ending CR + cell marker as two
            // text characters, but they occupy one terminal Range position.
            // Therefore the editable paragraph/cell body ends at End - 1, not
            // End - 2. Using End - 2 truncates a valid display OMath by one
            // logical position on Office 2021.
            var bodyEnd = Math.Min(
                Math.Max(paragraphRange.Start, paragraphRange.End - 1),
                Math.Max(cellRange.Start, cellRange.End - 1));
            if (bodyEnd < hostRange.End)
                throw new InvalidDataException(
                    "The table-cell host extends beyond the editable cell paragraph body.");
            after = document.Range(hostRange.End, bodyEnd);
            if (!ContainsOnlyStructuralText(before.Text)
                || !ContainsOnlyStructuralText(after.Text))
                throw new InvalidDataException(
                    "A numbered table-cell formula cannot consume user prose.");

            var plan = ResolveNumberPlan(
                document,
                hostRange.Start,
                host.FormulaId!);

            ConfigureCellParagraphTabs(
                paragraphRange);

            before.Text = "\t";
            var movedHost = ResolveOneHostInScope(
                document,
                paragraphRange,
                host.Kind,
                host.FormulaId);
            BindHostIdentity(document, movedHost);

            hostRange.SetRange(
                movedHost.Range.Start,
                movedHost.Range.End);
            CreateUserCellNumberLabel(
                document,
                movedHost,
                plan.Ordinal,
                plan.Prefix);

            var refreshed = ResolveOneHostInScope(
                document,
                paragraphRange,
                host.Kind,
                host.FormulaId);
            var descriptor = ResolveCanonicalUserCellNumbering(
                document,
                refreshed,
                paragraphRange);
            return WithNumbering(refreshed, descriptor);
        }
        finally
        {
            Release(after);
            Release(before);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(cellRange);
            Release(cell);
            Release(cells);
            Release(hostRange);
        }
    }

    private static WordFormulaHostDescriptor DetachUserTableCell(
        Document document,
        WordFormulaHostDescriptor host,
        WordFormulaNumberingDescriptor numbering)
    {
        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? prefix = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The numbered table-cell host no longer belongs to one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;

            var formulaEnd = hostRange.End;
            var bodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            var suffix = document.Range(formulaEnd, bodyEnd);
            try { suffix.Delete(); }
            finally { Release(suffix); }

            prefix = document.Range(
                paragraphRange.Start,
                hostRange.Start);
            if (ContainsOnlyStructuralText(prefix.Text))
                prefix.Delete();

            var refreshed = ResolveOneHostInScope(
                document,
                paragraphRange,
                host.Kind,
                host.FormulaId);
            BindHostIdentity(document, refreshed);
            ConfigureUnnumberedDisplayParagraph(paragraphRange);
            return WithNumbering(
                refreshed,
                new WordFormulaNumberingDescriptor
                {
                    Numbered = false,
                    FormulaId = refreshed.FormulaId,
                    ContainerKind =
                        WordFormulaNumberingContainerKind.None,
                });
        }
        finally
        {
            Release(prefix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
        }
    }

    private static WordFormulaNumberingDescriptor
        ResolveCanonicalUserCellNumbering(
            Document document,
            WordFormulaHostDescriptor host,
            Range paragraphRange)
    {
        if (!Guid.TryParse(host.FormulaId, out var parsed))
            throw new InvalidDataException(
                "The numbered table-cell host lost its FormulaId.");

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        Range? hostRange = null;
        Fields? fields = null;
        Field? field = null;
        Field? sequenceField = null;
        Range? code = null;
        Range? sequenceResult = null;
        try
        {
            bookmarks = paragraphRange.Bookmarks;
            var expected =
                NumberBookmarkPrefix + parsed.ToString("N");
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                bookmark = bookmarks[index];
                if (!string.Equals(
                        bookmark.Name,
                        expected,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Release(bookmark);
                    bookmark = null;
                    continue;
                }
                bookmarkRange = bookmark.Range.Duplicate;
                break;
            }
            if (bookmarkRange is null)
                throw new InvalidDataException(
                    "The table-cell number target bookmark was not created.");

            fields = paragraphRange.Fields;
            var count = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                code = field.Code;
                if (IsVisualTeXSequenceField(code.Text))
                {
                    count++;
                    if (sequenceField is null)
                    {
                        sequenceField = field;
                        field = null;
                    }
                }
                Release(code); code = null;
                Release(field); field = null;
            }
            if (count != 1 || sequenceField is null)
                throw new InvalidDataException(
                    "The canonical table-cell number paragraph does not contain exactly one VisualTeX SEQ field.");

            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            sequenceResult =
                sequenceField.Result.Duplicate;
            if (bookmarkRange.Start <= hostRange.End
                || sequenceResult.Start < bookmarkRange.Start
                || sequenceResult.End > bookmarkRange.End)
                throw new InvalidDataException(
                    "The canonical table-cell number target does not own the SEQ result after the formula. "
                    + $"host={hostRange.Start}:{hostRange.End} "
                    + $"bookmark={bookmarkRange.Start}:{bookmarkRange.End} "
                    + $"fieldResult={sequenceResult.Start}:{sequenceResult.End}.");

            return new WordFormulaNumberingDescriptor
            {
                Numbered = true,
                FormulaId = host.FormulaId,
                Position = "right",
                ContainerKind =
                    WordFormulaNumberingContainerKind.CanonicalUserTableCell,
                ContainerRange = Address(paragraphRange),
                NumberRange = Address(bookmarkRange),
            };
        }
        finally
        {
            Release(sequenceResult);
            Release(code);
            Release(sequenceField);
            Release(field);
            Release(fields);
            Release(hostRange);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static bool TryRekeyPastedBodyTable(
        Document document,
        WordFormulaHostDescriptor host,
        int ordinal,
        string prefix,
        out WordFormulaHostDescriptor result)
    {
        result = host;
        Range? hostRange = null;
        Tables? tables = null;
        Table? table = null;
        Rows? rows = null;
        Columns? columns = null;
        Cell? centerCell = null;
        Cell? rightCell = null;
        Range? centerRange = null;
        Range? rightRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            tables = hostRange.Tables;
            if (tables.Count != 1)
                return false;
            table = tables[1];
            rows = table.Rows;
            columns = table.Columns;
            if (rows.Count != 1 || columns.Count != 3)
                return false;

            centerCell = table.Cell(1, 2);
            centerRange = centerCell.Range.Duplicate;
            if (hostRange.Start < centerRange.Start
                || hostRange.End > centerRange.End)
                return false;

            rightCell = table.Cell(1, 3);
            rightRange = rightCell.Range.Duplicate;
            fields = rightRange.Fields;
            var ownedSequenceCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                code = field.Code;
                if (IsVisualTeXSequenceField(code.Text))
                    ownedSequenceCount++;
                Release(code); code = null;
                Release(field); field = null;
            }
            if (ownedSequenceCount != 1)
                return false;

            // Delete only copied VisualTeX-number aliases inside this local table.
            bookmarks = table.Range.Bookmarks;
            for (var index = bookmarks.Count; index >= 1; index--)
            {
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (name.StartsWith(
                        NumberBookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(
                        "VTEq_",
                        StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(
                        "VTEqCap_",
                        StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(
                        "VTEqAnc_",
                        StringComparison.OrdinalIgnoreCase))
                    bookmark.Delete();
                Release(bookmark); bookmark = null;
            }

            var editable = rightRange.Duplicate;
            try
            {
                editable.End = Math.Max(
                    editable.Start,
                    editable.End - 1);
                editable.Text = string.Empty;
            }
            finally { Release(editable); }

            CreateNumberCell(
                document,
                table,
                host,
                ordinal,
                prefix);
            ConfigureCanonicalTable(
                document,
                table);

            var refreshed =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (refreshed.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalBodyTable)
                throw new InvalidDataException(
                    "The pasted body number table could not be re-keyed canonically.");

            result = WithNumbering(
                host,
                refreshed);
            return true;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
            Release(rightRange);
            Release(centerRange);
            Release(rightCell);
            Release(centerCell);
            Release(columns);
            Release(rows);
            Release(table);
            Release(tables);
            Release(hostRange);
        }
    }

    private static bool TryRekeyPastedUserCell(
        Document document,
        WordFormulaHostDescriptor host,
        int ordinal,
        string prefix,
        out WordFormulaHostDescriptor result)
    {
        result = host;
        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? suffix = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            hostRange = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (!IsWithinTable(paragraphRange))
                return false;

            fields = paragraphRange.Fields;
            var sequenceCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                code = field.Code;
                if (IsVisualTeXSequenceField(code.Text))
                    sequenceCount++;
                Release(code); code = null;
                Release(field); field = null;
            }
            if (sequenceCount != 1)
                return false;

            bookmarks = paragraphRange.Bookmarks;
            for (var index = bookmarks.Count; index >= 1; index--)
            {
                bookmark = bookmarks[index];
                var name = bookmark.Name ?? string.Empty;
                if (name.StartsWith(
                        NumberBookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(
                        "VTEq_",
                        StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(
                        "VTEqCap_",
                        StringComparison.OrdinalIgnoreCase)
                    || name.StartsWith(
                        "VTEqAnc_",
                        StringComparison.OrdinalIgnoreCase))
                    bookmark.Delete();
                Release(bookmark); bookmark = null;
            }

            var bodyEnd = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            if (hostRange.End > bodyEnd)
                return false;
            suffix = document.Range(
                hostRange.End,
                bodyEnd);
            suffix.Text = string.Empty;
            ConfigureCellParagraphTabs(
                paragraphRange);
            CreateUserCellNumberLabel(
                document,
                host,
                ordinal,
                prefix);

            var refreshed =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (refreshed.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalUserTableCell)
                throw new InvalidDataException(
                    "The pasted table-cell number could not be re-keyed canonically.");

            result = WithNumbering(
                host,
                refreshed);
            return true;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(suffix);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
        }
    }

    private static (int Ordinal, string Prefix) ParseVisibleNumberPlan(
        string? visibleNumber)
    {
        var text = (visibleNumber ?? string.Empty)
            .Trim()
            .Trim('(', ')')
            .Trim();
        if (text.Length == 0)
            return (1, string.Empty);

        var end = text.Length - 1;
        while (end >= 0
            && char.IsWhiteSpace(text[end]))
            end--;
        var start = end;
        while (start >= 0
            && char.IsDigit(text[start]))
            start--;

        if (end < 0
            || start == end
            || !int.TryParse(
                text.Substring(
                    start + 1,
                    end - start),
                out var ordinal)
            || ordinal <= 0)
            return (1, string.Empty);

        return (
            ordinal,
            text.Substring(0, start + 1));
    }

    private static void RewriteCanonicalNumberLabel(
        Document document,
        WordFormulaHostDescriptor host,
        WordFormulaNumberingDescriptor numbering,
        int ordinal,
        string prefix)
    {
        if (numbering.ContainerKind ==
            WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph)
        {
            if (WordVisualTeXParagraphNumbering.IsSelfContainedHost(
                    document,
                    host))
            {
                WordVisualTeXParagraphNumbering.Refresh(
                    document,
                    host,
                    rebuildForCurrentFormat: true);
            }
            else
            {
                // Legacy body-tab hosts remain migration input. Preserve their
                // old VTEqCap/VTEqNum path until they are explicitly converted
                // or rewritten into the new self-contained paragraph form.
                RewriteBodyTabNumberLabel(
                    document,
                    host,
                    ordinal,
                    prefix);
            }
        }
        else if (numbering.ContainerKind ==
            WordFormulaNumberingContainerKind.CanonicalBodyTable)
        {
            if (numbering.ContainerRange is null)
                throw new InvalidDataException(
                    "The canonical number table has no container range.");

            Range? tableRange = null;
            Tables? tables = null;
            Table? table = null;
            Cell? rightCell = null;
            Range? rightRange = null;
            Range? editable = null;
            try
            {
                tableRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        numbering.ContainerRange);
                tables = tableRange.Tables;
                if (tables.Count != 1)
                    throw new InvalidDataException(
                        "The canonical number table moved before refresh.");
                table = tables[1];
                if (table.Rows.Count != 1
                    || table.Columns.Count != 3)
                    throw new InvalidDataException(
                        "The canonical number table is no longer 1x3.");

                rightCell = table.Cell(1, 3);
                rightRange = rightCell.Range.Duplicate;
                editable = rightRange.Duplicate;
                editable.End = Math.Max(
                    editable.Start,
                    editable.End - 1);
                editable.Text = string.Empty;

                CreateNumberCell(
                    document,
                    table,
                    host,
                    ordinal,
                    prefix);
            }
            finally
            {
                Release(editable);
                Release(rightRange);
                Release(rightCell);
                Release(table);
                Release(tables);
                Release(tableRange);
            }
        }
        else if (numbering.ContainerKind ==
                 WordFormulaNumberingContainerKind.CanonicalUserTableCell)
        {
            if (numbering.ContainerRange is null)
                throw new InvalidDataException(
                    "The canonical table-cell number has no paragraph range.");

            Range? paragraphRange = null;
            Range? suffix = null;
            try
            {
                paragraphRange =
                    WordFormulaHostSemanticReader.CreateRange(
                        document,
                        numbering.ContainerRange);
                var bodyEnd = Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
                if (host.Range.End > bodyEnd)
                    throw new InvalidDataException(
                        "The table-cell formula escaped its numbered paragraph.");

                suffix = document.Range(
                    host.Range.End,
                    bodyEnd);
                suffix.Text = string.Empty;
                CreateUserCellNumberLabel(
                    document,
                    host,
                    ordinal,
                    prefix);
            }
            finally
            {
                Release(suffix);
                Release(paragraphRange);
            }
        }
        else
        {
            throw new InvalidDataException(
                "Only canonical numbering containers can be refreshed.");
        }

        Range? hostRange = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            var refreshedHost =
                WordFormulaHostResolver.ResolveLocal(
                    document,
                    hostRange,
                    host.Kind)
                ?? throw new InvalidDataException(
                    "The formula host disappeared while refreshing its number.");
            if (refreshedHost.FormulaId is null
                && !string.IsNullOrWhiteSpace(host.FormulaId))
                refreshedHost.FormulaId =
                    host.FormulaId;
            var refreshed =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    refreshedHost);
            if (refreshed.ContainerKind
                != numbering.ContainerKind
                || !refreshed.Numbered)
                throw new InvalidDataException(
                    "The equation number refresh changed its canonical container.");
        }
        finally
        {
            Release(hostRange);
        }
    }

    private static void RewriteBodyTabNumberLabel(
        Document document,
        WordFormulaHostDescriptor host,
        int ordinal,
        string prefix)
    {
        if (host.Kind != WordFormulaHostKind.VisualTeX
            || !Guid.TryParse(host.FormulaId, out var parsed))
            throw new InvalidDataException(
                "A canonical body tab number requires a VisualTeX UUID host.");

        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            shape = GetExactVisualTeXNumberingShape(
                document,
                host);
            var metadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                    shape)
                ?? throw new InvalidDataException(
                    "The VisualTeX tab host lost its authoritative metadata before number refresh.");
            metadata.DisplayMode = "block";
            metadata.Numbered = true;
            metadata.Validate();

            WordEquationNumbering.RemoveFormulaNumberingArtifacts(
                document,
                parsed.ToString("D"));
            shapeRange = shape.Range.Duplicate;
            WordEquationNumbering.BuildFormulaNumberingScaffoldForConversion(
                document,
                shapeRange,
                shape.Height,
                metadata,
                plannedOrdinal: Math.Max(1, ordinal),
                plannedPrefix: prefix,
                deferFieldUpdate: false);
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
        }
    }

    private static string ReadLatexPreview(
        Document document,
        WordFormulaHostDescriptor host)
    {
        try
        {
            var payload =
                WordFormulaHostSemanticReader.Read(
                    document,
                    host);
            var latex = payload.Metadata?.Latex
                ?? payload.Latex;
            latex = (latex ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            return latex.Length <= 120
                ? latex
                : latex.Substring(0, 117) + "...";
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void CreateNumberCell(
        Document document,
        Table table,
        WordFormulaHostDescriptor host,
        int ordinal,
        string prefix)
    {
        Cell? cell = null;
        Range? cellRange = null;
        Range? editable = null;
        Range? placeholder = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        Range? numberTarget = null;
        try
        {
            cell = table.Cell(1, 3);
            cellRange = cell.Range.Duplicate;
            editable = cellRange.Duplicate;
            editable.End = Math.Max(
                editable.Start,
                editable.End - 1);
            editable.Text = "(" + prefix + "0)";

            var labelStart = editable.Start;
            var fieldStart =
                labelStart + 1 + prefix.Length;
            placeholder = document.Range(
                fieldStart,
                fieldStart + 1);
            fields = placeholder.Fields;
            object type = WdFieldType.wdFieldEmpty;
            object fieldCode =
                $"SEQ {SequenceName} \\r {Math.Max(1, ordinal)} \\* ARABIC";
            object preserve = false;
            field = fields.Add(
                placeholder,
                ref type,
                ref fieldCode,
                ref preserve);
            field.Update();
            result = field.Result;

            numberTarget = document.Range(
                labelStart + 1,
                result.End);
            BindNumberBookmark(
                document,
                host.FormulaId!,
                numberTarget);
        }
        finally
        {
            Release(numberTarget);
            Release(result);
            Release(field);
            Release(fields);
            Release(placeholder);
            Release(editable);
            Release(cellRange);
            Release(cell);
        }
    }

    private static void CreateUserCellNumberLabel(
        Document document,
        WordFormulaHostDescriptor host,
        int ordinal,
        string prefix)
    {
        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? terminal = null;
        Range? placeholder = null;
        Fields? fields = null;
        Field? field = null;
        Range? result = null;
        Range? numberTarget = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The table-cell formula no longer occupies one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            if (!IsWithinTable(paragraphRange))
                throw new InvalidDataException(
                    "The table-cell formula left its user table.");

            // A collapsed range at OMath.End still has mathematical affinity in
            // Word and InsertAfter can be absorbed into the display OMath. Anchor
            // the number to the paragraph terminator instead: InsertBefore on the
            // terminal structural range creates ordinary Word text immediately
            // before the cell ending, outside the math zone.
            terminal =
                paragraphRange.Duplicate;
            terminal.Start = Math.Max(
                paragraphRange.Start,
                paragraphRange.End - 1);
            var start = terminal.Start;
            terminal.InsertBefore(
                "\t(" + prefix + "0)");

            var fieldStart =
                start + 2 + prefix.Length;
            placeholder = document.Range(
                fieldStart,
                fieldStart + 1);
            fields = placeholder.Fields;
            object type = WdFieldType.wdFieldEmpty;
            object fieldCode =
                $"SEQ {SequenceName} \\r {Math.Max(1, ordinal)} \\* ARABIC";
            object preserve = false;
            field = fields.Add(
                placeholder,
                ref type,
                ref fieldCode,
                ref preserve);
            field.Update();
            result = field.Result;

            numberTarget = document.Range(
                start + 2,
                result.End);
            BindNumberBookmark(
                document,
                host.FormulaId
                ?? throw new InvalidDataException(
                    "The table-cell formula lost its FormulaId."),
                numberTarget);
        }
        finally
        {
            Release(numberTarget);
            Release(result);
            Release(field);
            Release(fields);
            Release(placeholder);
            Release(terminal);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
        }
    }

    private static void BindNumberBookmark(
        Document document,
        string formulaId,
        Range numberTarget)
    {
        if (!Guid.TryParse(formulaId, out var parsed))
            throw new InvalidDataException(
                "Equation number target requires a UUID FormulaId.");

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? verified = null;
        try
        {
            bookmarks = document.Bookmarks;
            var name =
                NumberBookmarkPrefix + parsed.ToString("N");
            bookmark = bookmarks.Add(name, numberTarget);
            verified = bookmark.Range;
            var requestedText =
                numberTarget.Text ?? string.Empty;
            var verifiedText =
                verified.Text ?? string.Empty;

            // A bookmark that ends exactly at a field result is normalized by
            // Word to include the invisible field-end boundary. That can advance
            // Range.End by one while leaving the visible target unchanged. The
            // required invariant is therefore the local visible target itself:
            // same story/start, complete coverage of the requested target, and
            // byte-for-byte identical visible text. Internal field-boundary
            // coordinates are Word serialization detail, not formula truth.
            if (verified.StoryType != numberTarget.StoryType
                || verified.Start != numberTarget.Start
                || verified.End < numberTarget.End
                || !string.Equals(
                    verifiedText,
                    requestedText,
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Word changed the visible equation-number reference target. "
                    + $"requested={numberTarget.StoryType}:{numberTarget.Start}:{numberTarget.End} "
                    + $"actual={verified.StoryType}:{verified.Start}:{verified.End} "
                    + $"requestedText=[{requestedText}] actualText=[{verifiedText}]");
        }
        finally
        {
            Release(verified);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static NumberPlan ResolveNumberPlan(
        Document document,
        int insertionPosition,
        string formulaId)
    {
        var formatId =
            WordEquationNumbering.GetEquationNumberFormatId(document);
        var format = EquationNumberFormat.Resolve(formatId);
        var scopeStart = int.MinValue;
        var prefix = string.Empty;

        if (format.UsesHeading)
        {
            var scopes =
                WordEquationNumbering.CaptureHeadingScopesAtPositions(
                    document,
                    formatId,
                    new[] { insertionPosition });
            if (scopes.TryGetValue(
                    insertionPosition,
                    out var scope))
            {
                scopeStart = scope.ScopeStart;
                prefix = scope.NumberText + format.Separator;
            }
        }

        var ordinal = 1;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? range = null;
        try
        {
            bookmarks = document.Bookmarks;
            var entries = new List<int>();
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                bookmark = bookmarks[index];
                if (!TryParseNumberBookmark(
                        bookmark.Name,
                        out var id)
                    || string.Equals(
                        id,
                        formulaId,
                        StringComparison.OrdinalIgnoreCase))
                {
                    Release(bookmark);
                    bookmark = null;
                    continue;
                }

                range = bookmark.Range;
                var start = range.Start;
                if (start < insertionPosition
                    && (!format.UsesHeading
                        || scopeStart == int.MinValue
                        || start > scopeStart))
                    entries.Add(start);

                Release(range); range = null;
                Release(bookmark); bookmark = null;
            }
            ordinal = entries.Count + 1;
        }
        finally
        {
            Release(range);
            Release(bookmark);
            Release(bookmarks);
        }

        return new NumberPlan
        {
            Ordinal = Math.Max(1, ordinal),
            Prefix = prefix,
        };
    }

    private static bool TryParseNumberBookmark(
        string? name,
        out string formulaId)
    {
        formulaId = string.Empty;
        if (string.IsNullOrWhiteSpace(name)
            || !name!.StartsWith(
                NumberBookmarkPrefix,
                StringComparison.OrdinalIgnoreCase)
            || !Guid.TryParseExact(
                name.Substring(NumberBookmarkPrefix.Length),
                "N",
                out var parsed))
            return false;
        formulaId = parsed.ToString("D");
        return true;
    }

    private static WordFormulaHostDescriptor ResolveOneHostInScope(
        Document document,
        Range scope,
        WordFormulaHostKind kind,
        string? expectedFormulaId)
    {
        var resolved = WordFormulaHostResolver.ResolveLocal(
                document,
                scope,
                kind)
            ?? throw new InvalidDataException(
                "Word did not preserve the formula host in its new numbering container.");
        if (!string.IsNullOrWhiteSpace(expectedFormulaId)
            && !string.IsNullOrWhiteSpace(resolved.FormulaId)
            && !string.Equals(
                resolved.FormulaId,
                expectedFormulaId,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "The moved formula host returned a different FormulaId.");
        if (resolved.FormulaId is null
            && Guid.TryParse(expectedFormulaId, out _))
            resolved.FormulaId = expectedFormulaId;
        return resolved;
    }

    private static void EnsureSemanticIdentityMatches(
        WordFormulaHostDescriptor original,
        WordFormulaHostDescriptor copied)
    {
        if (original.Kind != copied.Kind)
            throw new InvalidDataException(
                "Numbering moved the formula into a different host family.");
        if (!string.IsNullOrWhiteSpace(original.FormulaId)
            && !string.IsNullOrWhiteSpace(copied.FormulaId)
            && !string.Equals(
                original.FormulaId,
                copied.FormulaId,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                "Numbering duplicated the formula with a different identity.");
    }

    private static InlineShape GetExactVisualTeXNumberingShape(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            range =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            shapes = range.InlineShapes;
            if (shapes.Count != 1)
                throw new InvalidDataException(
                    "The VisualTeX numbering range does not contain exactly one InlineShape.");
            shape = shapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(
                    shape))
                throw new InvalidDataException(
                    "The VisualTeX numbering range contains a non-VisualTeX OLE object.");
            var metadata =
                WordFormulaMetadataReader.TryReadEmbeddedNativeOle(
                    shape)
                ?? throw new InvalidDataException(
                    "The VisualTeX numbering host has no authoritative embedded metadata.");
            if (!string.IsNullOrWhiteSpace(host.FormulaId)
                && !string.Equals(
                    metadata.FormulaId,
                    host.FormulaId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "The VisualTeX numbering host identity changed before mutation.");

            var result = shape;
            shape = null;
            return result;
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }

    private static void BindHostIdentity(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (!Guid.TryParse(host.FormulaId, out _))
            return;

        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            if (host.Kind == WordFormulaHostKind.Omml)
            {
                // Pure OMML never receives a VisualTeX identity bookmark.
                return;
            }

            if (host.Kind == WordFormulaHostKind.VisualTeX)
            {
                shapes = range.InlineShapes;
                if (shapes.Count != 1)
                    throw new InvalidDataException(
                        "The moved VisualTeX host is not one unique InlineShape.");
                shape = shapes[1];
                if (!WordFormulaMetadataReader.IsNativeOle(shape))
                    throw new InvalidDataException(
                        "The moved numbered object is not a VisualTeX OLE.");
                WordFormulaIdentityStore.BindVisualTeX(
                    shape,
                    host.FormulaId!);
            }
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }

    private static string DescribeBodyTableState(
        Table table,
        WordFormulaHostDescriptor host)
    {
        Rows? rows = null;
        Columns? columns = null;
        Range? tableRange = null;
        Cell? leftCell = null;
        Cell? centerCell = null;
        Cell? rightCell = null;
        Range? leftRange = null;
        Range? centerRange = null;
        Range? rightRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            rows = table.Rows;
            columns = table.Columns;
            tableRange = table.Range.Duplicate;
            var fieldCodes = new List<string>();
            var fieldResults = new List<string>();
            var bookmarkNames = new List<string>();
            if (rows.Count >= 1
                && columns.Count >= 3)
            {
                leftCell =
                    table.Cell(1, 1);
                centerCell =
                    table.Cell(1, 2);
                rightCell =
                    table.Cell(1, 3);
                leftRange =
                    leftCell.Range.Duplicate;
                centerRange =
                    centerCell.Range.Duplicate;
                rightRange =
                    rightCell.Range.Duplicate;
                fields =
                    rightRange.Fields;
                for (var index = 1;
                     index <= fields.Count;
                     index++)
                {
                    field = fields[index];
                    code = field.Code;
                    fieldCodes.Add(
                        (code.Text ?? string.Empty)
                        .Trim());
                    Range? fieldResult = null;
                    try
                    {
                        fieldResult =
                            field.Result.Duplicate;
                        fieldResults.Add(
                            $"{fieldResult.Start}:{fieldResult.End}:[{fieldResult.Text ?? string.Empty}]");
                    }
                    finally
                    {
                        Release(fieldResult);
                    }
                    Release(code);
                    code = null;
                    Release(field);
                    field = null;
                }
                bookmarks =
                    rightRange.Bookmarks;
                for (var index = 1;
                     index <= bookmarks.Count;
                     index++)
                {
                    bookmark =
                        bookmarks[index];
                    Range? bookmarkRange = null;
                    try
                    {
                        bookmarkRange =
                            bookmark.Range.Duplicate;
                        bookmarkNames.Add(
                            (bookmark.Name ?? string.Empty)
                            + $"@{bookmarkRange.Start}:{bookmarkRange.End}:[{bookmarkRange.Text ?? string.Empty}]");
                    }
                    finally
                    {
                        Release(bookmarkRange);
                    }
                    Release(bookmark);
                    bookmark = null;
                }
            }

            var centerMathCount = centerRange?.OMaths.Count ?? -1;
            var centerShapeCount = centerRange?.InlineShapes.Count ?? -1;
            var leftText =
                (leftRange?.Text ?? string.Empty)
                .Replace("\r", "\\r")
                .Replace("\a", "\\a")
                .Replace("\t", "\\t");
            return
                $"rows={rows.Count} cols={columns.Count} "
                + $"table={tableRange.Start}:{tableRange.End} "
                + $"host={host.Range.Start}:{host.Range.End} "
                + $"left={(leftRange is null ? "none" : leftRange.Start + ":" + leftRange.End + ":[" + leftText + "]")} "
                + $"center={(centerRange is null ? "none" : centerRange.Start + ":" + centerRange.End)} "
                + $"centerMaths={centerMathCount} centerShapes={centerShapeCount} "
                + $"right={(rightRange is null ? "none" : rightRange.Start + ":" + rightRange.End)} "
                + $"fields=[{string.Join("|", fieldCodes)}] "
                + $"fieldResults=[{string.Join("|", fieldResults)}] "
                + $"bookmarks=[{string.Join("|", bookmarkNames)}]";
        }
        catch (Exception error)
        {
            return "diagnostic-failed="
                + error.GetType().Name
                + ":"
                + error.Message;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(code);
            Release(field);
            Release(fields);
            Release(rightRange);
            Release(centerRange);
            Release(leftRange);
            Release(rightCell);
            Release(centerCell);
            Release(leftCell);
            Release(tableRange);
            Release(columns);
            Release(rows);
        }
    }

    private static void ConfigureCanonicalTable(
        Document document,
        Table table)
    {
        Rows? rows = null;
        Columns? columns = null;
        Column? left = null;
        Column? center = null;
        Column? right = null;
        Cell? leftCell = null;
        Cell? centerCell = null;
        Cell? rightCell = null;
        Range? centerRange = null;
        Range? rightRange = null;
        ParagraphFormat? centerFormat = null;
        ParagraphFormat? rightFormat = null;
        Borders? borders = null;
        Section? section = null;
        PageSetup? page = null;
        Range? tableRange = null;
        Sections? sections = null;
        try
        {
            rows = table.Rows;
            columns = table.Columns;
            if (rows.Count != 1 || columns.Count != 3)
                throw new InvalidDataException(
                    "The canonical equation-number table must be 1x3.");

            tableRange = table.Range;
            sections = tableRange.Sections;
            section = sections[1];
            page = section.PageSetup;
            var writable = Math.Max(
                180f,
                page.PageWidth - page.LeftMargin - page.RightMargin);
            var side = Math.Min(
                SideColumnWidthPoints,
                writable * 0.2f);
            var middle = Math.Max(
                72f,
                writable - (2f * side));

            table.AllowAutoFit = false;
            table.PreferredWidthType =
                WdPreferredWidthType.wdPreferredWidthPoints;
            table.PreferredWidth = writable;
            table.LeftPadding = 0f;
            table.RightPadding = 0f;
            table.TopPadding = 0f;
            table.BottomPadding = 0f;
            table.Spacing = 0f;

            left = columns[1];
            center = columns[2];
            right = columns[3];
            left.Width = side;
            center.Width = middle;
            right.Width = side;

            borders = table.Borders;
            borders.Enable = 0;

            leftCell = table.Cell(1, 1);
            centerCell = table.Cell(1, 2);
            rightCell = table.Cell(1, 3);
            leftCell.VerticalAlignment =
                WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            centerCell.VerticalAlignment =
                WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            rightCell.VerticalAlignment =
                WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            centerRange = centerCell.Range;
            centerFormat = centerRange.ParagraphFormat;
            centerFormat.Alignment =
                WdParagraphAlignment.wdAlignParagraphCenter;
            centerFormat.SpaceBefore = 0f;
            centerFormat.SpaceAfter = 0f;
            centerFormat.LeftIndent = 0f;
            centerFormat.RightIndent = 0f;
            centerFormat.FirstLineIndent = 0f;

            rightRange = rightCell.Range;
            rightFormat = rightRange.ParagraphFormat;
            rightFormat.Alignment =
                WdParagraphAlignment.wdAlignParagraphRight;
            rightFormat.SpaceBefore = 0f;
            rightFormat.SpaceAfter = 0f;
            rightFormat.LeftIndent = 0f;
            rightFormat.RightIndent = 0f;
            rightFormat.FirstLineIndent = 0f;
        }
        finally
        {
            Release(sections);
            Release(tableRange);
            Release(page);
            Release(section);
            Release(borders);
            Release(rightFormat);
            Release(centerFormat);
            Release(rightRange);
            Release(centerRange);
            Release(rightCell);
            Release(centerCell);
            Release(leftCell);
            Release(right);
            Release(center);
            Release(left);
            Release(columns);
            Release(rows);
        }
    }

    private static void NormalizeEmptyTableAnchor(Range anchor)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? range = null;
        ParagraphFormat? format = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        TabStops? tabs = null;
        try
        {
            paragraphs = anchor.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "The number-table anchor is not one body paragraph.");
            paragraph = paragraphs[1];
            range = paragraph.Range;
            if (IsWithinTable(range))
                throw new InvalidDataException(
                    "The number-table anchor belongs to a user table.");

            object normal = WdBuiltinStyle.wdStyleNormal;
            try { range.set_Style(ref normal); } catch { }
            format = range.ParagraphFormat;
            try { format.Reset(); } catch { }
            format.Alignment =
                WdParagraphAlignment.wdAlignParagraphLeft;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            tabs = format.TabStops;
            tabs.ClearAll();

            font = range.Font;
            try { font.Reset(); } catch { }
            font.Hidden = 0;
            font.Position = 0;
        }
        finally
        {
            Release(tabs);
            Release(font);
            Release(format);
            Release(range);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void ConfigureCellParagraphTabs(
        Range paragraphRange)
    {
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        try
        {
            format = paragraphRange.ParagraphFormat;
            format.Alignment =
                WdParagraphAlignment.wdAlignParagraphLeft;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
            tabs = format.TabStops;
            tabs.ClearAll();

            const float width = 240f;
            tabs.Add(
                width * 0.5f,
                WdTabAlignment.wdAlignTabCenter,
                WdTabLeader.wdTabLeaderSpaces);
            tabs.Add(
                width,
                WdTabAlignment.wdAlignTabRight,
                WdTabLeader.wdTabLeaderSpaces);
        }
        finally
        {
            Release(tabs);
            Release(format);
        }
    }

    private static void ConfigureUnnumberedDisplayParagraph(
        Range paragraphRange)
    {
        ParagraphFormat? format = null;
        try
        {
            format = paragraphRange.ParagraphFormat;
            format.Alignment =
                WdParagraphAlignment.wdAlignParagraphCenter;
            format.SpaceBefore = 0f;
            format.SpaceAfter = 0f;
            format.LeftIndent = 0f;
            format.RightIndent = 0f;
            format.FirstLineIndent = 0f;
        }
        finally { Release(format); }
    }

    private static bool HasAnyCanonicalEquationSequenceField(
        Document document)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code.Duplicate;
                var instruction =
                    code.Text
                    ?? string.Empty;
                if (WordNativeOmmlNumbering
                        .IsEquationSequenceFieldCode(
                            document,
                            instruction)
                    || IsVisualTeXSequenceField(
                        instruction))
                    return true;
            }
            return false;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool IsVisualTeXSequenceField(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf("SEQ", StringComparison.OrdinalIgnoreCase) >= 0
        && code.IndexOf(
            SequenceName,
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static bool ContainsOnlyStructuralText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        foreach (var value in text!)
        {
            if (value is '\r' or '\n' or '\t' or '\v'
                or '\a' or '\u0001')
                continue;
            if (!char.IsWhiteSpace(value)) return false;
        }
        return true;
    }

    private static bool IsWithinTable(Range range)
    {
        try
        {
            return Convert.ToBoolean(
                range.get_Information(WdInformation.wdWithInTable));
        }
        catch { return false; }
    }

    private static WordFormulaHostDescriptor WithNumbering(
        WordFormulaHostDescriptor host,
        WordFormulaNumberingDescriptor numbering) =>
        new()
        {
            Kind = host.Kind,
            Range = host.Range,
            DisplayMode = host.DisplayMode,
            FormulaId = host.FormulaId,
            Metadata = host.Metadata,
            MetadataAuthoritative = host.MetadataAuthoritative,
            Numbering = numbering,
            WithinTable = host.WithinTable,
            SourceMathMl = host.SourceMathMl,
            Latex = host.Latex,
        };

    private static WordFormulaRangeAddress Address(Range range) => new()
    {
        StoryType = range.StoryType,
        Start = range.Start,
        End = range.End,
    };

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

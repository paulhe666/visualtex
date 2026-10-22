using System.Diagnostics;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// The unnumbered display formulas one batch-numbering command will number,
/// captured in one read-only pass, plus what it leaves alone and why.
/// </summary>
internal sealed class WordDisplayNumberingPlan
{
    internal string DocumentId { get; set; } = string.Empty;
    internal int ContentEnd { get; set; }
    internal bool WholeDocument { get; set; }
    internal List<WordFormulaHostDescriptor> Omml { get; } = new();
    internal List<WordFormulaHostDescriptor> VisualTeX { get; } = new();
    internal List<WordFormulaHostDescriptor> MathType { get; } = new();
    // VisualTeX OLE paragraphs that still need the numbered tab layout (older
    // writers), keyed by the OLE start.
    internal HashSet<int> VisualTeXNeedsLayout { get; } = new();
    internal int AlreadyNumbered { get; set; }
    internal int SkippedInTable { get; set; }
    internal int SkippedOtherContent { get; set; }
    internal int SkippedOldNumbering { get; set; }
    internal int SkippedTypedNumber { get; set; }
    internal int Count => Omml.Count + VisualTeX.Count + MathType.Count;
    internal int Skipped =>
        SkippedInTable + SkippedOtherContent + SkippedOldNumbering + SkippedTypedNumber;
}

internal sealed class WordDisplayNumberingResult
{
    internal int Omml { get; set; }
    internal int VisualTeX { get; set; }
    internal int MathType { get; set; }
    internal int Count => Omml + VisualTeX + MathType;
    internal double ElapsedSeconds { get; set; }
}

internal sealed partial class WordFormulaService
{
    // Performance contract: every document-wide step runs once per command,
    // never once per formula. One read-only scan finds the formulas; one
    // preflight (before Word's undo record) reads each OMML equation once and
    // prepares each family's numbering plan once; one undo record writes each
    // formula locally (ScreenUpdating off, window minimized from 10 formulas);
    // then each family is renumbered once in document order. The per-formula
    // single-edit paths (adoption capture, nearest-template search, whole
    // document number update) must never be called from this loop.
    internal WordDisplayNumberingPlan PlanDisplayFormulaNumberingCore(bool wholeDocument)
    {
        using var numberFormatScope = WordEquationNumberFormatAuthority.BeginOperationScope();
        using var metric = WordOperationMetrics.Measure("DisplayNumbering.Plan");
        Document? document = null;
        Selection? selection = null;
        Range? scope = null;
        Range? content = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            content = document.Content;
            scope = wholeDocument
                ? content.Duplicate
                : (selection = _application.Selection).Range.Duplicate;

            var plan = new WordDisplayNumberingPlan
            {
                DocumentId = DocumentIdentity(document),
                ContentEnd = content.End,
                WholeDocument = wholeDocument,
            };

            IReadOnlyList<WordFormulaHostDescriptor> omml;
            IReadOnlyList<WordFormulaHostDescriptor> visualTeX;
            if (!wholeDocument && scope.Start == scope.End)
            {
                // A caret inside one formula numbers that formula.
                var localOmml = WordFormulaHostLookup.ResolveLocal(document, scope, WordFormulaHostKind.Omml);
                var localVisualTeX = localOmml is null
                    ? WordFormulaHostLookup.ResolveLocal(document, scope, WordFormulaHostKind.VisualTeX)
                    : null;
                omml = localOmml is null ? Array.Empty<WordFormulaHostDescriptor>() : new[] { localOmml };
                visualTeX = localVisualTeX is null ? Array.Empty<WordFormulaHostDescriptor>() : new[] { localVisualTeX };
            }
            else
            {
                var index = WordFormulaHostResolver.CaptureScopeIndex(document, scope);
                omml = index.Omml;
                visualTeX = index.VisualTeX;
            }

            using (WordOperationMetrics.Measure("DisplayNumbering.Plan.Omml"))
                foreach (var host in omml.Where(item => item.Display))
                    ClassifyOmmlForNumbering(document, host, plan);
            using (WordOperationMetrics.Measure("DisplayNumbering.Plan.VisualTeX"))
                foreach (var host in visualTeX.Where(item => item.Display))
                    ClassifyVisualTeXForNumbering(document, host, plan);
            using (WordOperationMetrics.Measure("DisplayNumbering.Plan.MathType"))
                ClassifyMathTypeForNumbering(document, scope, plan);
            return plan;
        }
        finally
        {
            Release(content);
            Release(scope);
            Release(selection);
            Release(document);
        }
    }

    private static void ClassifyOmmlForNumbering(
        Document document,
        WordFormulaHostDescriptor host,
        WordDisplayNumberingPlan plan)
    {
        var numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
        if (numbering.ContainerKind == WordFormulaNumberingContainerKind.Legacy)
        {
            plan.SkippedOldNumbering++;
            return;
        }
        // A typed Word #(n) label is a number too; the batch adopts it into
        // the same SEQ series instead of numbering the equation again.
        if (numbering.Numbered
            || WordOmmlHostWriter.HasFieldFreeWordNativeNumberHost(document, host))
        {
            plan.AlreadyNumbered++;
            return;
        }
        if (host.WithinTable)
        {
            plan.SkippedInTable++;
            return;
        }
        host.Numbering = numbering;
        plan.Omml.Add(host);
    }

    private static void ClassifyVisualTeXForNumbering(
        Document document,
        WordFormulaHostDescriptor host,
        WordDisplayNumberingPlan plan)
    {
        var numbering = WordFormulaNumberingResolver.ResolveLocal(document, host);
        if (numbering.ContainerKind == WordFormulaNumberingContainerKind.Legacy)
        {
            plan.SkippedOldNumbering++;
            return;
        }
        if (numbering.Numbered)
        {
            plan.AlreadyNumbered++;
            return;
        }
        if (host.WithinTable)
        {
            plan.SkippedInTable++;
            return;
        }
        if (!Guid.TryParse(host.FormulaId, out _))
        {
            // The index identifies VisualTeX OLEs by their identity bookmark;
            // fall back to the object's own metadata once.
            Range? range = null;
            try
            {
                range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
                var local = WordFormulaHostLookup.ResolveLocal(document, range, WordFormulaHostKind.VisualTeX);
                if (local is null || !Guid.TryParse(local.FormulaId, out _))
                {
                    plan.SkippedOtherContent++;
                    return;
                }
                host.FormulaId = local.FormulaId;
            }
            finally { Release(range); }
        }
        switch (WordVisualTeXParagraphNumbering.CheckBatchAttachReadiness(document, host))
        {
            case WordVisualTeXParagraphNumbering.BatchAttachReadiness.OtherContent:
                plan.SkippedOtherContent++;
                return;
            case WordVisualTeXParagraphNumbering.BatchAttachReadiness.NeedsLayout:
                plan.VisualTeXNeedsLayout.Add(host.Range.Start);
                break;
        }
        host.Numbering = numbering;
        plan.VisualTeX.Add(host);
    }

    private static void ClassifyMathTypeForNumbering(
        Document document,
        Range scope,
        WordDisplayNumberingPlan plan)
    {
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        Range? probe = null;
        try
        {
            if (scope.Start == scope.End)
            {
                // A caret beside or inside one MathType object.
                probe = scope.Duplicate;
                probe.SetRange(Math.Max(0, scope.Start - 1), scope.End + 1);
                shapes = probe.InlineShapes;
            }
            else
            {
                shapes = scope.InlineShapes;
            }
            // One enumerator pass: InlineShapes.Count and Item(index) both walk
            // the collection (O(n^2) over a long document).
            foreach (var current in WordComEnumeration.Each<InlineShape>(shapes))
            {
                Release(range); range = null;
                Release(shape); shape = current;
                if (!MathTypeOleInterop.IsMathTypeOle(shape))
                    continue;
                var (display, numbered, typedNumber) = MathTypeOleInterop.ClassifyDisplayNumbering(shape);
                if (!display)
                    continue;
                range = shape.Range;
                if (numbered)
                {
                    plan.AlreadyNumbered++;
                    continue;
                }
                if (WordEquationNumbering.RangeIsWhollyWithinTable(range))
                {
                    plan.SkippedInTable++;
                    continue;
                }
                if (typedNumber)
                {
                    plan.SkippedTypedNumber++;
                    continue;
                }
                plan.MathType.Add(new WordFormulaHostDescriptor
                {
                    Kind = WordFormulaHostKind.MathType,
                    DisplayMode = "block",
                    Range = new WordFormulaRangeAddress
                    {
                        StoryType = range.StoryType,
                        Start = range.Start,
                        End = range.End,
                    },
                });
            }
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
            Release(probe);
        }
    }

    internal WordDisplayNumberingResult NumberDisplayFormulasCore(WordDisplayNumberingPlan plan)
    {
        if (plan is null) throw new ArgumentNullException(nameof(plan));
        var watch = Stopwatch.StartNew();
        using var numberFormatScope = WordEquationNumberFormatAuthority.BeginOperationScope();
        Document? document = null;
        Range? content = null;
        WordNativeNumberingAdoptionPlan? nativeNumbers = null;
        WordFormulaNumberingKernel.VisualTeXConversionNumberPlan? visualTeXNumbers = null;
        MathTypeSectionStaging? mathTypeSection = null;
        var mathTypeShapes = new List<InlineShape>();
        Range? mathTypePlaceRefSource = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            EnsureEquationFieldResultsVisible(document);
            content = document.Content;
            if (!string.Equals(DocumentIdentity(document), plan.DocumentId, StringComparison.Ordinal)
                || content.End != plan.ContentEnd)
                throw new InvalidOperationException(
                    "The document changed after the formulas were counted. Run batch numbering again.");

            var formatId = EquationNumberFormat.Resolve(
                WordEquationNumberFormatAuthority.Resolve(document)).Id;
            var format = EquationNumberFormat.Resolve(formatId);

            // ---- Preflight: every read happens before Word's undo record. ----
            var preparedOmml = new Dictionary<int, string>();
            if (plan.Omml.Count > 0)
            {
                var semantic = new Dictionary<int, string>();
                using (WordOperationMetrics.Measure("DisplayNumbering.Omml.ReadEquations"))
                {
                    foreach (var host in plan.Omml)
                    {
                        Range? exact = null;
                        try
                        {
                            exact = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
                            semantic[host.Range.Start] = WordNativeOmmlNumbering.ReadSemanticOmml(exact);
                        }
                        finally { Release(exact); }
                        ReportOperationProgress(OfficePluginLanguage.Text("正在读取公式…", "Reading equations…"), semantic.Count, plan.Count);
                    }
                }
                nativeNumbers = WordNativeNumberingAdoptionPlan.PrepareBatchInsertion(
                    document, semantic.Keys.ToArray());
                using (WordOperationMetrics.Measure("DisplayNumbering.Omml.Prepare"))
                    foreach (var pair in semantic)
                        preparedOmml[pair.Key] = nativeNumbers.PrepareBatchInsertionOmml(pair.Key, pair.Value);
            }

            var visualTeXOperations = new Dictionary<int, string>();
            var sectionTabs = Array.Empty<(int Start, int End, float Center, float Right)>();
            if (plan.VisualTeX.Count > 0)
            {
                foreach (var host in plan.VisualTeX)
                    visualTeXOperations[host.Range.Start] = Guid.NewGuid().ToString("N");
                visualTeXNumbers = WordFormulaNumberingKernel.VisualTeXConversionNumberPlan.Capture(
                    document,
                    visualTeXOperations.Select(pair => (pair.Value, pair.Key)),
                    formatId);
                visualTeXNumbers.EnableBodyFieldCopies();
                if (plan.VisualTeXNeedsLayout.Count > 0)
                    sectionTabs = CaptureConversionSectionTabStops(document);
            }

            MathTypeWordOpenXml.NumberTemplate? mathTypeTemplate = null;
            var mathTypePosition = GetMathTypeNumberPositionPreference();
            var mathTypeColumns = Array.Empty<(int Start, int End, float? Width)>();
            WordCharacterFormatting? mathTypeBodyFormatting = null;
            if (plan.MathType.Count > 0)
            {
                // Per-formula layout queries (section number, page setup) cost
                // ~30 ms each; capture every section's column width once.
                mathTypeColumns = CaptureMathTypeSectionColumnWidths(document);
                mathTypeBodyFormatting = WordCharacterFormatting.CaptureParagraphMarkAtPosition(
                    document, plan.MathType.Max(host => host.Range.Start));
                using (WordOperationMetrics.Measure("DisplayNumbering.MathType.Template"))
                {
                    // One nearest-template search for the whole batch keeps the
                    // style of MathType numbers that already exist.
                    Field? nearest = null;
                    try
                    {
                        nearest = FindNearestMathTypePlaceRefField(
                            document, plan.MathType.Min(host => host.Range.Start), -1, -1);
                        mathTypeTemplate = nearest is not null
                            ? ReadMathTypePlaceRefTemplate(document, nearest)
                            : MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(formatId);
                    }
                    finally { Release(nearest); }
                }
                // Opening or closing a hidden staging document ends Word's custom
                // undo record: prepare the one section paragraph before it opens.
                if (format.UsesHeading && MathTypeNumberTemplateUsesHeading(mathTypeTemplate!))
                    mathTypeSection = new MathTypeSectionStaging(document, CreateMathTypeSectionFlatOpc(0, 0));
            }

            using var screenUpdating = NativeOmmlScreenUpdatingScope.Suspend(_application);
            using var minimized = plan.Count >= MinimizeWindowFormulaThreshold
                ? MinimizedWindowScope.Enter(document, OperationWindowMinimized)
                : null;
            var native = nativeNumbers;
            var result = WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Number Display Formulas",
                () =>
                {
                    var written = new WordDisplayNumberingResult();
                    // Write from the end of the document backwards: every target
                    // still to be written lies before the ones already changed,
                    // so its captured address stays exact.
                    var targets = plan.Omml.Concat(plan.VisualTeX).Concat(plan.MathType)
                        .OrderByDescending(host => host.Range.Start)
                        .ToArray();
                    // MathType's native paragraph styles are document styles:
                    // define them once for the batch, not once per equation.
                    if (plan.MathType.Count > 0)
                        using (WordOperationMetrics.Measure("DisplayNumbering.MathType.Styles"))
                            EnsureMathTypeNativeStyles(document, mathTypeBodyFormatting);
                    foreach (var host in targets)
                    {
                        ReportOperationProgress(OfficePluginLanguage.Text("正在添加公式编号…", "Adding equation numbers…"), written.Count, plan.Count);
                        switch (host.Kind)
                        {
                            case WordFormulaHostKind.Omml:
                                // Register the new native owner while its address
                                // is fresh; earlier targets written later shift it.
                                native!.AddBatchInserted(WordNativeOmmlNumbering.ReplaceWithPreparedNumbered(
                                    document, host, preparedOmml[host.Range.Start]));
                                written.Omml++;
                                break;
                            case WordFormulaHostKind.VisualTeX:
                            {
                                var target = plan.VisualTeXNeedsLayout.Contains(host.Range.Start)
                                    ? PrepareVisualTeXNumberedLayout(document, host, sectionTabs)
                                    : host;
                                using (WordOperationMetrics.Measure("DisplayNumbering.VisualTeX.Attach"))
                                    _ = WordFormulaNumberingKernel.AttachCanonical(
                                        document,
                                        target,
                                        deferFinalization: true,
                                        preparedNumberPlan: visualTeXNumbers!.For(visualTeXOperations[host.Range.Start]));
                                written.VisualTeX++;
                                break;
                            }
                            case WordFormulaHostKind.MathType:
                            {
                                var shape = FindMathTypeOleByRange(
                                        document,
                                        RangeReferenceFromAddress(host.Range),
                                        allowGlobalFallback: false)
                                    ?? throw new InvalidDataException(
                                        "A MathType equation to number moved before it was numbered.");
                                mathTypeShapes.Add(shape);
                                using (WordOperationMetrics.Measure("DisplayNumbering.MathType.Scaffold"))
                                    RebuildMathTypeDisplayScaffold(
                                        document,
                                        shape,
                                        numbered: true,
                                        mathTypePosition,
                                        mathTypeTemplate,
                                        updateNumberFields: false,
                                        nativeStylesAlreadyPrepared: true,
                                        knownDisplayColumnWidth: mathTypeColumns
                                            .FirstOrDefault(item => item.Start <= host.Range.Start && item.End > host.Range.Start)
                                            .Width,
                                        placeRefSource: mathTypePlaceRefSource,
                                        keepExistingCenterTab: true);
                                // The first finished number is copied to every
                                // later one; the final pass evaluates them all.
                                mathTypePlaceRefSource ??= CaptureMathTypePlaceRefSource(document, shape, mathTypePosition);
                                written.MathType++;
                                break;
                            }
                        }
                    }
                    ReportOperationProgress(OfficePluginLanguage.Text("正在更新编号…", "Updating numbers…"), written.Count, plan.Count);

                    // ---- One renumbering pass per family, in document order. ----
                    if (native is not null)
                        using (WordOperationMetrics.Measure("DisplayNumbering.Omml.Renumber"))
                            native.ApplyBatch(written.Omml);
                    if (visualTeXNumbers is not null)
                    {
                        using var renumber = WordOperationMetrics.Measure("DisplayNumbering.VisualTeX.Renumber");
                        visualTeXNumbers.ApplyExistingTables(document);
                        _ = WordFormulaNumberingKernel.RefreshCanonicalNumbers(
                            document,
                            kindFilter: WordFormulaHostKind.VisualTeX,
                            minimumExpectedNumbered: written.VisualTeX,
                            preparedVisualTeX: visualTeXNumbers,
                            bulkUpdate: true);
                    }
                    if (mathTypeShapes.Count > 0)
                    {
                        using var renumber = WordOperationMetrics.Measure("DisplayNumbering.MathType.Renumber");
                        if (mathTypeSection is not null)
                            EnsureExistingMathTypeHeadingScopes(
                                document,
                                formatId,
                                mathTypeShapes.Select(ReadShapeParagraphStart).ToList(),
                                mathTypeSection.Paragraph);
                        _ = MathTypeEquationNumbering.UpdateEquationNumbers(document, bulkUpdate: true);
                    }
                    return written;
                },
                verifyRollback: native is null ? null : new Action(native.VerifyRollback));
            result.ElapsedSeconds = watch.Elapsed.TotalSeconds;
            return result;
        }
        finally
        {
            Release(mathTypePlaceRefSource);
            foreach (var shape in mathTypeShapes) Release(shape);
            mathTypeSection?.Dispose();
            visualTeXNumbers?.Dispose();
            nativeNumbers?.Dispose();
            Release(content);
            Release(document);
        }
    }

    // An older VisualTeX display paragraph (centered, no tab stops) receives the
    // justified center/right tab layout a numbered display needs. A leading TAB
    // may be inserted, so the OLE's address is read again from the object.
    private static WordFormulaHostDescriptor PrepareVisualTeXNumberedLayout(
        Document document,
        WordFormulaHostDescriptor host,
        (int Start, int End, float Center, float Right)[] sectionTabs)
    {
        using var metric = WordOperationMetrics.Measure("DisplayNumbering.VisualTeX.Layout");
        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? moved = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
            shapes = range.InlineShapes;
            if (shapes.Count != 1)
                throw new InvalidDataException("The VisualTeX OLE to number moved before it was numbered.");
            shape = shapes[1];
            var section = sectionTabs.FirstOrDefault(item => item.Start <= host.Range.Start && item.End > host.Range.Start);
            WordEquationNumbering.PrepareVisualTeXDisplayHostLayout(
                document,
                range,
                section.End > section.Start ? section.Center : null,
                section.End > section.Start ? section.Right : null);
            moved = shape.Range;
            return new WordFormulaHostDescriptor
            {
                Kind = host.Kind,
                DisplayMode = host.DisplayMode,
                FormulaId = host.FormulaId,
                Metadata = host.Metadata,
                MetadataAuthoritative = host.MetadataAuthoritative,
                Numbering = host.Numbering,
                WithinTable = host.WithinTable,
                Range = new WordFormulaRangeAddress
                {
                    StoryType = moved.StoryType,
                    Start = moved.Start,
                    End = moved.End,
                },
            };
        }
        finally
        {
            Release(moved);
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }

    // The display column width of every single-column section (null for
    // multi-column sections, whose width depends on the paragraph position).
    private static (int Start, int End, float? Width)[] CaptureMathTypeSectionColumnWidths(Document document)
    {
        using var metric = WordOperationMetrics.Measure("DisplayNumbering.MathType.SectionWidths");
        var result = new List<(int Start, int End, float? Width)>();
        Sections? sections = null;
        try
        {
            sections = document.Sections;
            for (var index = 1; index <= sections.Count; index++)
            {
                Section? section = null;
                Range? range = null;
                PageSetup? setup = null;
                TextColumns? columns = null;
                try
                {
                    section = sections[index];
                    range = section.Range;
                    setup = section.PageSetup;
                    columns = setup.TextColumns;
                    float? width = columns is not null && columns.Count > 1
                        ? null
                        : Math.Max(72f, setup.PageWidth - setup.LeftMargin - setup.RightMargin);
                    result.Add((range.Start, range.End, width));
                }
                finally { Release(columns); Release(setup); Release(range); Release(section); }
            }
        }
        finally { Release(sections); }
        return result.ToArray();
    }

    // Exactly the MTPlaceRef field the scaffold just wrote for this equation:
    // TAB OLE TAB [FIELD] for a right number, [FIELD] TAB OLE for a left one.
    private static Range CaptureMathTypePlaceRefSource(Document document, InlineShape shape, string position)
    {
        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Field? placeRef = null;
        Range? code = null;
        Range? source = null;
        Fields? copied = null;
        Field? first = null;
        Range? firstCode = null;
        try
        {
            shapeRange = shape.Range;
            paragraphs = shapeRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            var numberOnLeft = string.Equals(position, "left", StringComparison.OrdinalIgnoreCase);
            placeRef = FindMathTypePlaceRefFieldForShape(paragraphRange, shapeRange, numberOnLeft)
                ?? throw new InvalidDataException("The first batch MathType equation has no MTPlaceRef to copy.");
            code = placeRef.Code;
            var start = code.Start - 1;
            var end = numberOnLeft ? shapeRange.Start - 1 : paragraphRange.End - 1;
            source = document.Range(start, end);
            copied = source.Fields;
            if (copied.Count < 1)
                throw new InvalidDataException("The batch MathType number source holds no field.");
            first = copied[1];
            firstCode = first.Code;
            if (firstCode.Start != code.Start || end <= code.End)
                throw new InvalidDataException("The batch MathType number source is not exactly one MTPlaceRef field.");
            var result = source;
            source = null;
            return result;
        }
        finally
        {
            Release(firstCode);
            Release(first);
            Release(copied);
            Release(source);
            Release(code);
            Release(placeRef);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    private static int ReadShapeParagraphStart(InlineShape shape)
    {
        Range? range = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        try
        {
            range = shape.Range;
            paragraphs = range.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            return paragraphRange.Start;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
        }
    }
}

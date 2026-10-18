using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Canonical single-paragraph numbering for VisualTeX OLE display equations.
///
/// The physical Word layout intentionally mirrors MathType's mature model:
///
///   TAB + VisualTeX OLE + TAB + MACROBUTTON VisualTeXPlaceRef(...) + paragraph mark
///
/// VisualTeXPlaceRef owns a hidden SEQ increment plus the visible current value.
/// No hidden caption paragraph, Frame, VTEqCap_* or always-on VTEqNum_* bookmark
/// is required.  A VTEqNum_* bookmark is created lazily only when the user inserts
/// a body reference to this equation.
/// </summary>
internal static partial class WordVisualTeXParagraphNumbering
{
    private const string PlaceRefInstruction = "VisualTeXPlaceRef";
    private const string NumberBookmarkPrefix = "VTEqNum_";

    // Created only by the writer after it has configured this new body host.
    // Live ranges survive removal of the temporary OMML boundary guard before
    // attachment; their ownership transfers once to Attach.
    internal sealed class WrittenBodyLayout : IDisposable
    {
        private Range? _host;
        private Range? _paragraph;

        internal WrittenBodyLayout(Range host)
        {
            Paragraphs? paragraphs = null;
            Paragraph? paragraph = null;
            try
            {
                _host = host.Duplicate;
                paragraphs = host.Paragraphs;
                paragraph = paragraphs[1];
                _paragraph = paragraph.Range.Duplicate;
            }
            catch { Dispose(); throw; }
            finally { Release(paragraph); Release(paragraphs); }
        }

        internal (Range Host, Range Paragraph) Take()
        {
            var result = (_host!, _paragraph!);
            _host = null;
            _paragraph = null;
            return result;
        }

        public void Dispose()
        {
            Release(_paragraph); _paragraph = null;
            Release(_host); _host = null;
        }
    }

    // Physical fields produced by one new body scaffold. Their roles and order
    // come from the prepared segments, not a later COM collection/code reread.
    internal sealed class CreatedBodyFields : IDisposable
    {
        private Field? _outer;
        private readonly List<(int Order, string Instruction, Field Field)> _nested = new();
        internal Field Outer => _outer!;

        internal void Add(int order, string instruction, Field field) =>
            _nested.Add((order, instruction.TrimStart(), field));

        internal void TakeOuter(Field outer)
        {
            _outer = outer;
            _nested.Sort((left, right) => left.Order.CompareTo(right.Order));
        }

        internal void Update()
        {
            using var metric = WordOperationMetrics.Measure("FormatConversion.VisualTeXNumbering.UpdateCreatedBodyFields");
            UpdatePlaceRefTree(_outer!);
        }

        public void Dispose()
        {
            foreach (var item in _nested) Release(item.Field);
            _nested.Clear();
            Release(_outer); _outer = null;
        }
    }

    internal static WordFormulaHostDescriptor Attach(
        Document document,
        WordFormulaHostDescriptor host,
        bool deferFinalization = false,
        WordFormulaNumberingKernel.NumberPlan? preparedNumberPlan = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.VisualTeXNumbering.Attach");
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.VisualTeX
            || !host.Display
            || host.WithinTable)
            throw new ArgumentException(
                "Single-paragraph VisualTeX numbering requires one body display OLE.",
                nameof(host));
        if (!Guid.TryParse(host.FormulaId, out _))
            throw new InvalidDataException(
                "Single-paragraph VisualTeX numbering requires a stable FormulaId.");

        Range? hostRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Range? suffix = null;
        Range? tabInsertion = null;
        Field? placeRef = null;
        try
        {
            var writtenLayout = preparedNumberPlan?.WrittenBodyLayout;
            if (writtenLayout is not null)
            {
                (hostRange, paragraphRange) = writtenLayout.Take();
                preparedNumberPlan!.WrittenBodyLayout = null;
            }
            else
            {
                hostRange = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
                paragraphs = hostRange.Paragraphs;
                if (paragraphs.Count != 1)
                    throw new InvalidDataException(
                        "A VisualTeX numbered display host must occupy one Word paragraph.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                if (WordEquationNumbering.IsWithinTableLayoutFree(paragraphRange))
                    throw new InvalidDataException(
                        "Body paragraph numbering cannot be attached inside a user table.");
                ValidateDisplayParagraphGeometry(document, hostRange, paragraph, paragraphRange);
            }

            var editableEnd =
                Math.Max(
                    paragraphRange.Start,
                    paragraphRange.End - 1);
            if (hostRange.End > editableEnd)
                throw new InvalidDataException(
                    "The VisualTeX OLE escaped its display paragraph.");

            suffix =
                document.Range(
                    hostRange.End,
                    editableEnd);
            if (!ContainsOnlyScaffoldWhitespace(
                    suffix.Text))
                throw new InvalidDataException(
                    "VisualTeX refused to overwrite user content after the display OLE while attaching numbering.");
            suffix.Text = string.Empty;

            tabInsertion =
                document.Range(
                    hostRange.End,
                    hostRange.End);
            tabInsertion.Text = "\t";

            ResolvedNumberPlan numberPlan;
            using (WordOperationMetrics.Measure("Numbering.Attach.ResolvePlan"))
                numberPlan = preparedNumberPlan?.BodyNumbering
                    ?? ResolveNumberPlan(document, hostRange.Start);
            using (WordOperationMetrics.Measure("Numbering.Attach.HeadingStateFields"))
            {
                if (writtenLayout is not null)
                    InsertHeadingStateFields(document, hostRange, numberPlan, deferFinalization);
                else
                    RebuildHeadingStateFields(document, hostRange, numberPlan);
            }

            CreatedBodyFields? createdFields = null;
            if (deferFinalization && preparedNumberPlan is not null)
            {
                createdFields = new CreatedBodyFields();
                preparedNumberPlan.CreatedBodyFields = createdFields;
            }
            using (WordOperationMetrics.Measure("Numbering.Attach.CreatePlaceRef"))
                placeRef = CreatePlaceRef(document, hostRange.End + 1, numberPlan, createdFields);
            if (createdFields is not null)
            {
                createdFields.TakeOuter(placeRef);
                placeRef = null;
            }
            if (!deferFinalization)
            {
                using (WordOperationMetrics.Measure("Numbering.Attach.UpdateFields"))
                {
                    UpdateHeadingStateFields(
                        document,
                        hostRange);
                    UpdatePlaceRefFields(placeRef!);
                }

                using (WordOperationMetrics.Measure("Numbering.Attach.ResolveRequired"))
                    return ResolveRequired(
                        document,
                        hostRange,
                        host.FormulaId!);
            }

            // Batch conversion materializes the complete target stream before
            // evaluating it. Keep the fully materialized field topology, but defer
            // field evaluation and the expensive resolver round-trip until the
            // single batch finalization pass. If any scaffold is malformed, that
            // final pass throws inside the same mutation transaction and Word
            // rolls the complete batch back.
            if (writtenLayout is not null)
                host.Range = Address(hostRange);
            host.Numbering =
                new WordFormulaNumberingDescriptor
                {
                    Numbered = true,
                    FormulaId = host.FormulaId,
                    Position = "right",
                    ContainerKind =
                        WordFormulaNumberingContainerKind
                            .CanonicalBodyTabParagraph,
                    SelfContainedVisualTeXParagraph = true,
                    ContainerRange =
                        Address(paragraphRange),
                };
            return host;
        }
        finally
        {
            Release(placeRef);
            Release(tabInsertion);
            Release(suffix);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(hostRange);
        }
    }

    internal static WordFormulaHostDescriptor Detach(
        Document document,
        WordFormulaHostDescriptor host,
        bool sourceWillBeRemovedImmediately = false,
        BoundBodyNumber? preparedBody = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.VisualTeXNumbering.Detach");
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.VisualTeX)
            throw new ArgumentException(
                "VisualTeX paragraph numbering detach received another host kind.",
                nameof(host));

        Range? hostRange = null;
        Field? placeRef = null;
        Range? code = null;
        Range? result = null;
        Range? fullField = null;
        Range? separator = null;
        try
        {
            if (preparedBody is not null)
            {
                hostRange = preparedBody.Remove();
                if (!sourceWillBeRemovedImmediately)
                    return ResolveUnnumbered(document, hostRange, host);
                host.Range = Address(hostRange);
                host.Numbering = new WordFormulaNumberingDescriptor
                {
                    Numbered = false,
                    FormulaId = host.FormulaId,
                    Position = "right",
                    ContainerKind = WordFormulaNumberingContainerKind.None,
                };
                return host;
            }

            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            if (!TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null)
                return host;

            code = placeRef.Code.Duplicate;
            result = placeRef.Result.Duplicate;
            var fieldStart =
                code.Start - 1;
            var fieldEnd =
                ResolveOuterFieldEndExclusive(
                    document,
                    code,
                    result);
            if (fieldStart < hostRange.End)
                throw new InvalidDataException(
                    "VisualTeX paragraph number moved before or into its OLE.");

            separator =
                document.Range(
                    hostRange.End,
                    fieldStart);
            if (!string.Equals(
                    separator.Text,
                    "\t",
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "VisualTeX paragraph number lost its single OLE-to-number TAB separator.");

            fullField =
                document.Range(
                    fieldStart,
                    fieldEnd);
            fullField.Delete();
            separator =
                document.Range(
                    hostRange.End,
                    Math.Min(
                        document.Content.End,
                        hostRange.End + 1));
            if (string.Equals(
                    separator.Text,
                    "\t",
                    StringComparison.Ordinal))
                separator.Delete();

            RemoveHeadingStateFields(
                document,
                hostRange);

            // Any lazily-created VTEqNum_* bookmark lived wholly inside the deleted
            // visible number. Word removes it naturally; existing body REF fields
            // then enter Word's ordinary missing-source state.
            //
            // A batch cross-format replacement deletes this exact OLE immediately
            // after detach. Reopening the same embedded OLE and re-running the full
            // numbering resolver just to prove an intermediate unnumbered state is
            // duplicate work; the detach operations above are exact and any COM
            // failure aborts the surrounding transaction. Preserve the updated
            // Word range and mark the transient descriptor unnumbered so the caller
            // can delete the host directly. Interactive detach keeps the independent
            // post-detach resolver/validation below.
            if (sourceWillBeRemovedImmediately)
            {
                host.Range = Address(hostRange);
                host.Numbering =
                    new WordFormulaNumberingDescriptor
                    {
                        Numbered = false,
                        FormulaId = host.FormulaId,
                        Position = "right",
                        ContainerKind =
                            WordFormulaNumberingContainerKind.None,
                    };
                return host;
            }

            return ResolveUnnumbered(
                document,
                hostRange,
                host);
        }
        finally
        {
            Release(separator);
            Release(fullField);
            Release(result);
            Release(code);
            Release(placeRef);
            Release(hostRange);
        }
    }

    internal static bool TryResolve(
        Document document,
        Range hostRange,
        string? expectedFormulaId,
        out WordFormulaNumberingDescriptor descriptor) =>
        TryResolveCore(document, hostRange, expectedFormulaId, false, out descriptor, out _);

    internal static bool TryResolve(
        Document document,
        Range hostRange,
        string? expectedFormulaId,
        out WordFormulaNumberingDescriptor descriptor,
        out BoundBodyNumber? body) =>
        TryResolveCore(document, hostRange, expectedFormulaId, true, out descriptor, out body);

    private static bool TryResolveCore(
        Document document,
        Range hostRange,
        string? expectedFormulaId,
        bool retainBody,
        out WordFormulaNumberingDescriptor descriptor,
        out BoundBodyNumber? body)
    {
        body = null;
        BoundBodyNumber? capturedBody = null;
        descriptor = null!;
        if (document is null
            || hostRange is null
            || !Guid.TryParse(
                expectedFormulaId,
                out var parsed))
            return false;

        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        Range? leading = null;
        Field? placeRef = null;
        Range? visibleNumber = null;
        try
        {
            if (WordEquationNumbering.IsWithinTableLayoutFree(hostRange))
                return false;
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (hostRange.Start <= paragraphRange.Start)
                return false;

            leading =
                document.Range(
                    hostRange.Start - 1,
                    hostRange.Start);
            if (!string.Equals(
                    leading.Text,
                    "\t",
                    StringComparison.Ordinal))
                return false;

            format = paragraph.Format;
            if (format.Alignment !=
                WdParagraphAlignment.wdAlignParagraphJustify)
                return false;
            tabs = format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1;
                 index <= tabs.Count;
                 index++)
            {
                Release(tab);
                tab = tabs[index];
                hasCenter |=
                    tab.Alignment ==
                    WdTabAlignment.wdAlignTabCenter;
                hasRight |=
                    tab.Alignment ==
                    WdTabAlignment.wdAlignTabRight;
            }
            if (!hasCenter || !hasRight)
                return false;

            if (retainBody) capturedBody = new BoundBodyNumber(hostRange);
            if (!TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef,
                    capturedBody)
                || placeRef is null)
                return false;

            if (!TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    includeParentheses: false,
                    out visibleNumber,
                    capturedBody?.Code)
                || visibleNumber is null)
                return false;

            descriptor =
                new WordFormulaNumberingDescriptor
                {
                    Numbered = true,
                    FormulaId =
                        parsed.ToString("D"),
                    Position = "right",
                    ContainerKind =
                        WordFormulaNumberingContainerKind
                            .CanonicalBodyTabParagraph,
                    SelfContainedVisualTeXParagraph = true,
                    ContainerRange =
                        Address(paragraphRange),
                    NumberRange =
                        Address(visibleNumber),
                };
            body = capturedBody;
            capturedBody = null;
            return true;
        }
        catch
        {
            // A requested owning binding must not turn COM acquisition failure
            // into a legacy/no-number fallback. General recognition keeps its
            // existing observational false-on-failure contract.
            if (retainBody) throw;
            descriptor = null!;
            return false;
        }
        finally
        {
            capturedBody?.Dispose();
            Release(visibleNumber);
            Release(placeRef);
            Release(leading);
            Release(tab);
            Release(tabs);
            Release(format);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal static void Refresh(
        Document document,
        WordFormulaHostDescriptor host,
        bool rebuildForCurrentFormat)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.VisualTeXNumbering.Refresh");
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.VisualTeX
            || !host.Display
            || host.WithinTable)
            throw new ArgumentException(
                "VisualTeX paragraph-number refresh requires one body display OLE.",
                nameof(host));

        Range? hostRange = null;
        Field? placeRef = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            if (!TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null)
                throw new InvalidDataException(
                    "The VisualTeX paragraph number disappeared before refresh.");

            var currentFormat =
                EquationNumberFormat.Resolve(
                    WordEquationNumberFormatAuthority.Resolve(
                        document));
            var mustRebuild =
                rebuildForCurrentFormat
                || currentFormat.UsesHeading;

            if (mustRebuild)
            {
                // Heading-aware VisualTeX numbering stores its private chapter /
                // section state in hidden fields in this same paragraph, outside
                // the VisualTeXPlaceRef outer field (matching MathType's topology).
                // Re-read Heading scopes on every explicit refresh so moving or
                // renumbering headings cannot leave stale private state.
                RebuildPlaceRefPreservingReferenceBookmark(
                    document,
                    hostRange,
                    host.FormulaId
                    ?? throw new InvalidDataException(
                        "VisualTeX paragraph numbering lost its FormulaId."),
                    placeRef);
                Release(placeRef);
                placeRef = null;
                if (!TryFindPlaceRef(
                        document,
                        hostRange,
                        out placeRef)
                    || placeRef is null)
                    throw new InvalidDataException(
                        "The VisualTeX paragraph number disappeared after format rebuild.");
            }

            UpdateHeadingStateFields(
                document,
                hostRange);
            UpdatePlaceRefFields(placeRef);
        }
        finally
        {
            Release(placeRef);
            Release(hostRange);
        }
    }

    internal static void EnsureReferenceBookmark(
        Document document,
        WordFormulaHostDescriptor host,
        string bookmarkName)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.VisualTeX)
            throw new ArgumentException(
                "VisualTeX reference target requires a VisualTeX host.",
                nameof(host));
        if (string.IsNullOrWhiteSpace(
                bookmarkName))
            throw new ArgumentException(
                "VisualTeX reference target bookmark is empty.",
                nameof(bookmarkName));

        Range? hostRange = null;
        Field? placeRef = null;
        Range? numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? existing = null;
        Range? existingRange = null;
        Bookmark? created = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            if (!TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null
                || !TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    includeParentheses: false,
                    out numberRange)
                || numberRange is null)
                throw new InvalidDataException(
                    "The VisualTeX equation no longer exposes one self-contained number.");

            bookmarks = document.Bookmarks;
            if (bookmarks.Exists(
                    bookmarkName))
            {
                existing =
                    bookmarks[bookmarkName];
                existingRange =
                    existing.Range.Duplicate;
                if (existingRange.StoryType ==
                        numberRange.StoryType
                    && existingRange.Start ==
                        numberRange.Start
                    && existingRange.End ==
                        numberRange.End)
                    return;
                existing.Delete();
                Release(existing);
                existing = null;
            }

            created =
                bookmarks.Add(
                    bookmarkName,
                    numberRange);
            Release(existingRange);
            existingRange =
                created.Range.Duplicate;
            if (existingRange.StoryType !=
                    numberRange.StoryType
                || existingRange.Start !=
                    numberRange.Start
                || existingRange.End !=
                    numberRange.End)
                throw new InvalidDataException(
                    "Word changed the lazy VisualTeX reference bookmark range.");
        }
        finally
        {
            Release(created);
            Release(existingRange);
            Release(existing);
            Release(bookmarks);
            Release(numberRange);
            Release(placeRef);
            Release(hostRange);
        }
    }

    internal static bool IsSelfContainedHost(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? hostRange = null;
        try
        {
            if (host is null
                || host.Kind != WordFormulaHostKind.VisualTeX
                || string.IsNullOrWhiteSpace(host.FormulaId))
                return false;
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            return TryResolve(
                document,
                hostRange,
                host.FormulaId,
                out _);
        }
        catch
        {
            return false;
        }
        finally
        {
            Release(hostRange);
        }
    }

    internal static string ReadVisibleNumberText(
        Document document,
        WordFormulaHostDescriptor host)
    {
        Range? hostRange = null;
        Field? placeRef = null;
        try
        {
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            if (!TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null)
                return string.Empty;
            return ReadVisibleNumberText(
                    document,
                    placeRef)
                .Trim()
                .Trim('(', ')')
                .Trim();
        }
        finally
        {
            Release(placeRef);
            Release(hostRange);
        }
    }

    private static Field CreatePlaceRef(
        Document document,
        int position,
        ResolvedNumberPlan plan,
        CreatedBodyFields? createdFields = null)
    {
        Range? insertion = null;
        Range? outerCode = null;
        Fields? nestedFields = null;
        Field? outer = null;
        Field? nested = null;
        try
        {
            insertion =
                document.Range(
                    position,
                    position);
            outer =
                document.Fields.Add(
                    insertion,
                    WdFieldType.wdFieldMacroButton,
                    PlaceRefInstruction,
                    false);
            outer.ShowCodes = true;

            outerCode =
                outer.Code.Duplicate;
            var nativeCodeText =
                outerCode.Text
                ?? string.Empty;
            if (nativeCodeText.Length == 0
                || !char.IsWhiteSpace(
                    nativeCodeText[
                        nativeCodeText.Length - 1]))
                throw new InvalidOperationException(
                    "Word did not create VisualTeXPlaceRef with its native instruction separator.");

            var insertionPosition =
                outerCode.End;
            Release(insertion);
            insertion =
                document.Range(
                    insertionPosition,
                    insertionPosition);
            insertion.InsertAfter(")");

            Release(outerCode);
            outerCode =
                outer.Code.Duplicate;
            if (insertionPosition <=
                    outerCode.Start
                || insertionPosition >=
                    outerCode.End)
                throw new InvalidOperationException(
                    "Word did not retain VisualTeXPlaceRef closing punctuation inside the outer field.");

            var segments = plan.Segments ?? BuildSegments(plan);
            for (var index =
                     segments.Count - 1;
                 index >= 0;
                 index--)
            {
                var segment =
                    segments[index];
                Release(insertion);
                insertion =
                    document.Range(
                        insertionPosition,
                        insertionPosition);
                if (!segment.IsField)
                {
                    if (!string.IsNullOrEmpty(
                            segment.Value))
                        insertion.InsertAfter(
                            segment.Value);
                    continue;
                }

                Release(nested);
                Release(nestedFields);
                Release(outerCode);
                outerCode =
                    outer.Code.Duplicate;
                nestedFields =
                    outerCode.Fields;
                nested =
                    nestedFields.Add(
                        insertion,
                        WdFieldType.wdFieldEmpty,
                        segment.Value,
                        false);
                try
                {
                    nested.ShowCodes = false;
                }
                catch { }
                if (createdFields is not null)
                {
                    createdFields.Add(index, segment.Value, nested);
                    nested = null;
                }
            }

            Release(outerCode);
            outerCode =
                outer.Code.Duplicate;
            var completed =
                outerCode.Text
                ?? string.Empty;
            if (completed.Length == 0
                || char.IsWhiteSpace(
                    completed[
                        completed.Length - 1]))
                throw new InvalidOperationException(
                    "Word left trailing whitespace after the VisualTeXPlaceRef number.");

            try
            {
                outer.ShowCodes = false;
            }
            catch { }

            var result = outer;
            outer = null;
            return result;
        }
        finally
        {
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
            Release(insertion);
            Release(outer);
        }
    }

    internal sealed class ResolvedNumberPlan
    {
        internal EquationNumberFormat Format { get; set; } =
            EquationNumberFormat.Resolve(
                EquationNumberFormat.ContinuousId);
        internal int Chapter { get; set; }
        internal int Section { get; set; }
        internal IReadOnlyList<NumberSegment>? Segments { get; set; }
    }

    private static ResolvedNumberPlan ResolveNumberPlan(
        Document document,
        int formulaPosition)
    {
        var format =
            EquationNumberFormat.Resolve(
                WordEquationNumberFormatAuthority.Resolve(
                    document));
        var plan =
            new ResolvedNumberPlan
            {
                Format = format,
            };
        if (!format.UsesHeading)
            return plan;

        // One formula: the nearest explicit heading found by a backward native
        // Find decides, exactly as the full anchor scan would (that scan remains
        // the fallback); the whole-document heading scan is for batches.
        string numberText;
        using (WordOperationMetrics.Measure("Numbering.Attach.HeadingScope"))
            numberText =
                WordEquationNumbering
                    .ResolveHeadingScopeAtPosition(
                        document,
                        formulaPosition,
                        format.Id)
                    .NumberText;
        return PrepareNumberPlan(format, numberText);
    }

    internal static ResolvedNumberPlan PrepareNumberPlan(
        EquationNumberFormat format, string? numberText)
    {
        var plan = new ResolvedNumberPlan { Format = format };
        var parts =
            (numberText ?? string.Empty)
                .Split(
                    new[] { '.' },
                    StringSplitOptions.RemoveEmptyEntries)
                .Select(part =>
                    part.Trim())
                .ToArray();

        static int ReadComponent(
            IReadOnlyList<string> values,
            int index)
        {
            if (index >= values.Count)
                return 0;
            return int.TryParse(
                    values[index],
                    System.Globalization.NumberStyles.Integer,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var parsed)
                && parsed >= 0
                    ? parsed
                    : 0;
        }

        plan.Chapter =
            ReadComponent(
                parts,
                0);
        if (format.HeadingLevel >= 2)
        {
            plan.Section =
                ReadComponent(
                    parts,
                    1);
        }

        plan.Segments = BuildSegments(plan);
        return plan;
    }

    private static void RebuildHeadingStateFields(
        Document document,
        Range hostRange,
        ResolvedNumberPlan plan)
    {
        RemoveHeadingStateFields(document, hostRange);
        InsertHeadingStateFields(document, hostRange, plan);
    }

    private static void InsertHeadingStateFields(
        Document document, Range hostRange, ResolvedNumberPlan plan,
        bool deferEvaluation = false)
    {
        if (plan.Format.HeadingLevel <= 0)
            return;

        void InsertStateField(
            string instruction)
        {
            Range? insertion = null;
            Field? field = null;
            try
            {
                // The canonical display host always owns exactly one leading TAB
                // immediately before the OLE. Insert private heading state just
                // before that TAB. Because hostRange is live, its Start shifts
                // after each insertion and the next state field naturally lands
                // after the previous one while remaining before the TAB.
                insertion =
                    document.Range(
                        hostRange.Start - 1,
                        hostRange.Start - 1);
                field =
                    document.Fields.Add(
                        insertion,
                        WdFieldType.wdFieldEmpty,
                        instruction,
                        false);
                try
                {
                    field.ShowCodes = false;
                }
                catch { }
                if (!deferEvaluation) field.Update();
            }
            finally
            {
                Release(field);
                Release(insertion);
            }
        }

        InsertStateField(
            $"SEQ {ChapterSequenceName} \\r {plan.Chapter} \\h \\* MERGEFORMAT");
        if (plan.Format.HeadingLevel >= 2)
        {
            InsertStateField(
                $"SEQ {SectionSequenceName} \\r {plan.Section} \\h \\* MERGEFORMAT");
        }
    }

    private static void UpdateHeadingStateFields(
        Document document,
        Range hostRange)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            paragraphs =
                hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "VisualTeX heading-state refresh requires one display paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            fields =
                paragraphRange.Fields;

            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                code =
                    field.Code.Duplicate;
                if (!IsHeadingStateRestartCode(
                        code.Text)
                    || code.Start - 1 >=
                        hostRange.Start)
                    continue;
                field.Update();
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void RemoveHeadingStateFields(
        Document document,
        Range hostRange)
    {
        _ = document;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            paragraphs =
                hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "VisualTeX heading-state cleanup requires one display paragraph.");
            paragraph =
                paragraphs[1];
            paragraphRange =
                paragraph.Range.Duplicate;
            fields =
                paragraphRange.Fields;

            // Word field control characters are not ordinary Range.Text for
            // empty-result \h fields. Delete them through Field.Delete(), from
            // the end of the paragraph toward the start, exactly as Word itself
            // manages field ownership.
            for (var index =
                     fields.Count;
                 index >= 1;
                 index--)
            {
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                code =
                    field.Code.Duplicate;
                if (!IsHeadingStateRestartCode(
                        code.Text)
                    || code.Start - 1 >=
                        hostRange.Start)
                    continue;
                field.Delete();
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    internal sealed class NumberSegment
    {
        internal bool IsField { get; set; }
        internal string Value { get; set; } =
            string.Empty;
    }

    private static IReadOnlyList<NumberSegment>
        BuildSegments(
            ResolvedNumberPlan plan)
    {
        var format =
            plan.Format;
        var result =
            new List<NumberSegment>();

        // The private chapter/section restart state lives before the
        // leading TAB in this same paragraph, outside the outer field. This is
        // the same topology MathType uses for section state. VisualTeXPlaceRef
        // therefore contains only the formula increment plus visible current
        // chapter/section/equation values.
        result.Add(
            new NumberSegment
            {
                IsField = true,
                Value =
                    format.UsesHeading
                        ? $"SEQ {SequenceName} \\s {format.HeadingLevel} \\h \\* MERGEFORMAT"
                        : $"SEQ {SequenceName} \\h \\* MERGEFORMAT",
            });
        result.Add(
            new NumberSegment
            {
                IsField = false,
                Value = "(",
            });

        if (format.HeadingLevel >= 1)
        {
            result.Add(
                new NumberSegment
                {
                    IsField = true,
                    Value =
                        $"SEQ {ChapterSequenceName} \\c \\* ARABIC \\* MERGEFORMAT",
                });
        }
        if (format.HeadingLevel >= 2)
        {
            result.Add(
                new NumberSegment
                {
                    IsField = false,
                    Value = ".",
                });
            result.Add(
                new NumberSegment
                {
                    IsField = true,
                    Value =
                        $"SEQ {SectionSequenceName} \\c \\* ARABIC \\* MERGEFORMAT",
                });
        }
        if (format.UsesHeading)
        {
            result.Add(
                new NumberSegment
                {
                    IsField = false,
                    Value = format.Separator,
                });
        }

        result.Add(
            new NumberSegment
            {
                IsField = true,
                Value =
                    $"SEQ {SequenceName} \\c \\* ARABIC \\* MERGEFORMAT",
            });

        // The final ')' is materialized before these reverse-inserted segments.
        return result;
    }

    internal static int RefreshBatchFields(
        Document document,
        int minimumExpectedNumbered = 0,
        ISet<string>? targetBookmarkNames = null,
        WordFormulaNumberingKernel.VisualTeXConversionNumberPlan? preparedVisualTeX = null)
    {
        using var operationMetric =
            WordOperationMetrics.Measure(
                "FormatConversion.VisualTeXNumbering.RefreshBatchFields");
        if (document is null)
            throw new ArgumentNullException(nameof(document));

        WordNumberingFieldCollection? fields = null;
        WordNumberingFieldCollection.Entry? entry = null;
        var placeRefCount = 0;
        var hasUserReferenceFields = false;
        // End of the PlaceRef code whose nested fields were just evaluated.
        var evaluatedPlaceRefEnd = -1;
        var createdFields = preparedVisualTeX?.CreateBodyFieldCursor();
        try
        {
            // Capture only physically bounded numbering/reference fields. Global
            // document.Fields[index] can enter a long WWLIB enumeration over
            // unrelated historical EMBED fields (CP110). The operation-local
            // snapshot preserves story order and all existing validation below.
            using (WordOperationMetrics.Measure("Numbering.BatchFields.Capture"))
                fields = WordNumberingFieldCollection.Capture(document);
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                entry?.Dispose();
                entry = null;
                entry = fields.TakeEntry(index);
                var field = entry.Field;
                var code = entry.Code;

                if (entry.Type is WdFieldType.wdFieldRef or WdFieldType.wdFieldGoToButton)
                    hasUserReferenceFields = true;

                var instruction = entry.Instruction;

                // The snapshot also lists the fields nested in a PlaceRef. Those
                // were evaluated together with their PlaceRef just before, so a
                // second Update of each nested sequence would only repeat it.
                if (entry.Type == WdFieldType.wdFieldSequence
                    && code.Start < evaluatedPlaceRefEnd)
                    continue;

                if (IsHeadingStateRestartCode(
                        instruction))
                {
                    using (WordOperationMetrics.Measure("Numbering.BatchFields.HeadingRestartUpdate"))
                        field.Update();
                    continue;
                }

                if (entry.Type ==
                        WdFieldType.wdFieldMacroButton
                    && IsPlaceRefCode(
                        instruction))
                {
                    bool updatedFromPlan;
                    using (WordOperationMetrics.Measure("Numbering.BatchFields.PlannedUpdate"))
                        updatedFromPlan = createdFields?.TryUpdate(code) == true;
                    if (!updatedFromPlan)
                        using (WordOperationMetrics.Measure("Numbering.BatchFields.PlaceRefUpdate"))
                            UpdatePlaceRefTree(field);
                    evaluatedPlaceRefEnd = code.End;
                    placeRefCount++;
                    continue;
                }

                // Older canonical VisualTeX numbering can own a standalone
                // top-level sequence field instead of the current self-contained
                // PlaceRef tree. Keep it numerically coherent if such a host
                // coexists with newly converted formulas, without rebuilding or
                // globally rediscovering that legacy-compatible host.
                if (IsStandaloneEquationSequenceCode(
                        instruction))
                    field.Update();
            }
        }
        finally
        {
            try { entry?.Dispose(); }
            finally { fields?.Dispose(); }
        }

        createdFields?.EnsureComplete();
        if (placeRefCount <
            minimumExpectedNumbered)
            throw new InvalidDataException(
                $"VisualTeX batch numbering materialized only {placeRefCount} numbered hosts; expected at least {minimumExpectedNumbered}.");

        // Formula references are uncommon. Do not perform another document-wide
        // field pass when the document contains no REF/GOTOBUTTON at all.
        if (hasUserReferenceFields)
            using (WordOperationMetrics.Measure("Numbering.BatchFields.UpdateReferences"))
                _ =
                    WordEquationReferenceFields
                        .UpdateReferences(
                            document,
                            targetBookmarkNames);

        return placeRefCount;
    }

    private static bool IsStandaloneEquationSequenceCode(
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(
                instruction))
            return false;
        var normalized =
            instruction!
                .TrimStart();
        return normalized.StartsWith(
                   "SEQ "
                   + SequenceName
                   + " ",
                   StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith(
                   "SEQ \""
                   + SequenceName
                   + "\" ",
                   StringComparison.OrdinalIgnoreCase);
    }

    // Batch finalization: a PlaceRef tree holds only SEQ fields (no STYLEREF, see
    // the parenthesis acceptance), so one Update of the outer field evaluates
    // them in code order (counter increment first, then the displayed values),
    // as Word's own full field update does. Measured on 200 formulas: the same
    // numbers at less than half the cost of one Update per nested field.
    private static void UpdatePlaceRefTree(Field placeRef)
    {
        placeRef.Update();
        try { placeRef.ShowCodes = false; } catch { }
    }

    private static void UpdatePlaceRefFields(Field placeRef)
    {
        Range? code = null;
        Fields? nested = null;
        Field? field = null;
        Range? fieldCode = null;
        var headingFields = new List<Field>();
        var sequenceFields = new List<Field>();
        try
        {
            code = placeRef.Code;
            nested = code.Fields;
            var count = nested.Count;
            // Classify each physical field once. Keep owned COM references while
            // evaluating the same heading-before-sequence dependency order;
            // reading the collection again cannot add information to this tree.
            for (var index = 1; index <= count; index++)
            {
                field = nested[index];
                fieldCode = field.Code;
                var instruction = (fieldCode.Text ?? string.Empty).TrimStart();
                Release(fieldCode); fieldCode = null;
                if (instruction.StartsWith("STYLEREF ", StringComparison.OrdinalIgnoreCase))
                    headingFields.Add(field);
                else if (instruction.StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase))
                    sequenceFields.Add(field);
                else Release(field);
                field = null;
            }
            foreach (var heading in headingFields) heading.Update();
            foreach (var sequence in sequenceFields) sequence.Update();
            try { placeRef.ShowCodes = false; } catch { }
        }
        finally
        {
            Release(fieldCode);
            Release(field);
            foreach (var heading in headingFields) Release(heading);
            foreach (var sequence in sequenceFields) Release(sequence);
            Release(nested);
            Release(code);
        }
    }

    private static void RebuildPlaceRefPreservingReferenceBookmark(
        Document document,
        Range hostRange,
        string formulaId,
        Field placeRef)
    {
        var bookmarkName =
            NumberBookmarkPrefix
            + Guid.Parse(
                formulaId)
                .ToString("N");

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        Range? code = null;
        Range? result = null;
        Range? fullField = null;
        Range? insertion = null;
        Field? rebuilt = null;
        try
        {
            bookmarks =
                document.Bookmarks;
            var preserveBookmark =
                bookmarks.Exists(
                    bookmarkName);
            if (preserveBookmark)
            {
                bookmark =
                    bookmarks[bookmarkName];
                bookmarkRange =
                    bookmark.Range.Duplicate;
            }

            code =
                placeRef.Code.Duplicate;
            result =
                placeRef.Result.Duplicate;
            var fieldStart =
                code.Start - 1;
            var fieldEnd =
                ResolveOuterFieldEndExclusive(
                    document,
                    code,
                    result);
            fullField =
                document.Range(
                    fieldStart,
                    fieldEnd);
            fullField.Delete();

            RemoveHeadingStateFields(
                document,
                hostRange);
            var plan =
                ResolveNumberPlan(
                    document,
                    hostRange.Start);
            RebuildHeadingStateFields(
                document,
                hostRange,
                plan);

            rebuilt =
                CreatePlaceRef(
                    document,
                    hostRange.End + 1,
                    plan);
            UpdateHeadingStateFields(
                document,
                hostRange);
            UpdatePlaceRefFields(
                rebuilt);

            if (preserveBookmark)
            {
                Range? numberRange = null;
                try
                {
                    if (!TryGetVisibleNumberRange(
                            document,
                            rebuilt,
                            includeParentheses: false,
                            out numberRange)
                        || numberRange is null)
                        throw new InvalidDataException(
                            "The rebuilt VisualTeXPlaceRef exposes no visible number for an existing reference bookmark.");

                    Release(bookmark);
                    bookmark = null;
                    if (bookmarks.Exists(
                            bookmarkName))
                    {
                        bookmark =
                            bookmarks[bookmarkName];
                        bookmark.Delete();
                        Release(bookmark);
                        bookmark = null;
                    }
                    bookmark =
                        bookmarks.Add(
                            bookmarkName,
                            numberRange);
                }
                finally
                {
                    Release(numberRange);
                }
            }
        }
        finally
        {
            Release(rebuilt);
            Release(insertion);
            Release(fullField);
            Release(result);
            Release(code);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static string ReadVisibleNumberText(
        Document document,
        Field placeRef)
    {
        Range? outerCode = null;
        Fields? nested = null;
        Field? field = null;
        Range? fieldCode = null;
        Range? fieldResult = null;
        try
        {
            outerCode =
                placeRef.Code.Duplicate;
            var stream =
                outerCode.Text
                ?? string.Empty;
            nested =
                outerCode.Fields;
            if (nested.Count == 0)
                return string.Empty;

            var hiddenEnd = -1;
            var nestedOrdinal = 0;
            for (var index = 0;
                 index < stream.Length;
                 index++)
            {
                if (stream[index] != '\u0013')
                    continue;
                nestedOrdinal++;
                var end =
                    stream.IndexOf(
                        '\u0015',
                        index + 1);
                if (end < 0)
                    return string.Empty;

                if (nestedOrdinal <=
                    nested.Count)
                {
                    Release(fieldCode);
                    fieldCode = null;
                    Release(field);
                    field =
                        nested[nestedOrdinal];
                    fieldCode =
                        field.Code.Duplicate;
                    if (IsHiddenIncrement(
                            document,
                            fieldCode.Text))
                    {
                        hiddenEnd = end;
                        break;
                    }
                }

                index = end;
            }

            if (hiddenEnd < 0)
                return string.Empty;

            var visible =
                new StringBuilder();
            var fieldOrdinal =
                nestedOrdinal;
            for (var index =
                     hiddenEnd + 1;
                 index < stream.Length;)
            {
                if (stream[index] !=
                    '\u0013')
                {
                    if (stream[index] is not
                        '\u0014'
                        and not '\u0015')
                        visible.Append(
                            stream[index]);
                    index++;
                    continue;
                }

                var end =
                    stream.IndexOf(
                        '\u0015',
                        index + 1);
                if (end < 0)
                    break;
                fieldOrdinal++;
                if (fieldOrdinal <=
                    nested.Count)
                {
                    Release(fieldResult);
                    fieldResult = null;
                    Release(field);
                    field =
                        nested[fieldOrdinal];
                    fieldResult =
                        field.Result.Duplicate;
                    visible.Append(
                        fieldResult.Text
                        ?? string.Empty);
                }
                index =
                    end + 1;
            }

            return visible
                .ToString()
                .Trim();
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            Release(fieldResult);
            Release(fieldCode);
            Release(field);
            Release(nested);
            Release(outerCode);
        }
    }

    private static void ValidateDisplayParagraphGeometry(
        Document document,
        Range hostRange,
        Paragraph paragraph,
        Range paragraphRange)
    {
        Range? leading = null;
        ParagraphFormat? format = null;
        TabStops? tabs = null;
        TabStop? tab = null;
        try
        {
            if (hostRange.Start <=
                paragraphRange.Start)
                throw new InvalidDataException(
                    "The VisualTeX display OLE has no leading center-tab slot.");
            leading =
                document.Range(
                    hostRange.Start - 1,
                    hostRange.Start);
            if (!string.Equals(
                    leading.Text,
                    "\t",
                    StringComparison.Ordinal))
                throw new InvalidDataException(
                    "The VisualTeX display OLE lost its leading center TAB.");

            format =
                paragraph.Format;
            if (format.Alignment !=
                WdParagraphAlignment.wdAlignParagraphJustify)
                throw new InvalidDataException(
                    "The VisualTeX display paragraph is not justified.");
            tabs =
                format.TabStops;
            var hasCenter = false;
            var hasRight = false;
            for (var index = 1;
                 index <= tabs.Count;
                 index++)
            {
                Release(tab);
                tab =
                    tabs[index];
                hasCenter |=
                    tab.Alignment ==
                    WdTabAlignment.wdAlignTabCenter;
                hasRight |=
                    tab.Alignment ==
                    WdTabAlignment.wdAlignTabRight;
            }
            if (!hasCenter || !hasRight)
                throw new InvalidDataException(
                    "The VisualTeX display paragraph lost its center/right tab stops.");
        }
        finally
        {
            Release(tab);
            Release(tabs);
            Release(format);
            Release(leading);
        }
    }

    private static WordFormulaHostDescriptor ResolveRequired(
        Document document,
        Range hostRange,
        string formulaId)
    {
        var host =
            WordFormulaHostResolver.ResolveLocal(
                document,
                hostRange,
                WordFormulaHostKind.VisualTeX)
            ?? throw new InvalidDataException(
                "The VisualTeX OLE disappeared while attaching paragraph numbering.");
        if (string.IsNullOrWhiteSpace(
                host.FormulaId))
            host.FormulaId =
                formulaId;
        var numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
        if (!numbering.Numbered
            || numbering.ContainerKind !=
                WordFormulaNumberingContainerKind
                    .CanonicalBodyTabParagraph)
            throw new InvalidDataException(
                "The VisualTeX OLE did not converge to the canonical self-contained paragraph number.");
        host.Numbering =
            numbering;
        return host;
    }

    private static WordFormulaHostDescriptor ResolveUnnumbered(
        Document document,
        Range hostRange,
        WordFormulaHostDescriptor fallback)
    {
        var host =
            WordFormulaHostResolver.ResolveLocal(
                document,
                hostRange,
                WordFormulaHostKind.VisualTeX)
            ?? throw new InvalidDataException(
                "The VisualTeX OLE disappeared while detaching paragraph numbering.");
        if (string.IsNullOrWhiteSpace(
                host.FormulaId))
            host.FormulaId =
                fallback.FormulaId;
        host.Numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
        if (host.Numbering.Numbered
            || host.Numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.None)
            throw new InvalidDataException(
                "VisualTeX paragraph numbering survived canonical detach.");
        return host;
    }

    private static WordFormulaRangeAddress Address(
        Range range) =>
        new()
        {
            StoryType = range.StoryType,
            Start = range.Start,
            End = range.End,
        };

}

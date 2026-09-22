using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
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
internal static class WordVisualTeXParagraphNumbering
{
    private const string PlaceRefInstruction = "VisualTeXPlaceRef";
    private const string PlaceRefMarker = "MACROBUTTON VisualTeXPlaceRef";
    private const string SequenceName = "VisualTeXEquation";
    private const string ChapterSequenceName = "VisualTeXChapter";
    private const string SectionSequenceName = "VisualTeXSection";
    private const string NumberBookmarkPrefix = "VTEqNum_";

    internal static WordFormulaHostDescriptor Attach(
        Document document,
        WordFormulaHostDescriptor host)
    {
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
            hostRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    host.Range);
            paragraphs = hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "A VisualTeX numbered display host must occupy one Word paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range.Duplicate;
            if (Convert.ToBoolean(
                    paragraphRange.get_Information(
                        WdInformation.wdWithInTable)))
                throw new InvalidDataException(
                    "Body paragraph numbering cannot be attached inside a user table.");

            ValidateDisplayParagraphGeometry(
                document,
                hostRange,
                paragraph,
                paragraphRange);

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

            var numberPlan =
                ResolveNumberPlan(
                    document,
                    hostRange.Start);
            RebuildHeadingStateFields(
                document,
                hostRange,
                numberPlan);

            placeRef =
                CreatePlaceRef(
                    document,
                    hostRange.End + 1,
                    numberPlan);
            UpdateHeadingStateFields(
                document,
                hostRange);
            UpdatePlaceRefFields(placeRef);

            return ResolveRequired(
                document,
                hostRange,
                host.FormulaId!);
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
        WordFormulaHostDescriptor host)
    {
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
        out WordFormulaNumberingDescriptor descriptor)
    {
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
            if (Convert.ToBoolean(
                    hostRange.get_Information(
                        WdInformation.wdWithInTable)))
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

            if (!TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null)
                return false;

            if (!TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    includeParentheses: false,
                    out visibleNumber)
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
                    ContainerRange =
                        Address(paragraphRange),
                    NumberRange =
                        Address(visibleNumber),
                };
            return true;
        }
        catch
        {
            descriptor = null!;
            return false;
        }
        finally
        {
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
                    WordEquationNumbering
                        .GetEquationNumberFormatId(
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

    internal static Range? FindVisibleLabelRange(
        Document document,
        string formulaId)
    {
        Range? hostRange = null;
        Field? placeRef = null;
        Range? visible = null;
        try
        {
            hostRange =
                ResolveVisualTeXHostRangeByFormulaId(
                    document,
                    formulaId);
            if (hostRange is null
                || !TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null
                || !TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    includeParentheses: true,
                    out visible)
                || visible is null)
                return null;

            var result =
                visible.Duplicate;
            return result;
        }
        finally
        {
            Release(visible);
            Release(placeRef);
            Release(hostRange);
        }
    }

    internal static Range? FindOwnerParagraphRange(
        Document document,
        string formulaId)
    {
        Range? hostRange = null;
        Field? placeRef = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        try
        {
            hostRange =
                ResolveVisualTeXHostRangeByFormulaId(
                    document,
                    formulaId);
            if (hostRange is null
                || !TryFindPlaceRef(
                    document,
                    hostRange,
                    out placeRef)
                || placeRef is null)
                return null;

            paragraphs =
                hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return null;
            paragraph =
                paragraphs[1];
            owner =
                paragraph.Range.Duplicate;
            var result = owner;
            owner = null;
            return result;
        }
        finally
        {
            Release(owner);
            Release(paragraph);
            Release(paragraphs);
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

    internal static bool IsPlaceRefCode(
        string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(
            PlaceRefMarker,
            StringComparison.OrdinalIgnoreCase) >= 0;

    private static Field CreatePlaceRef(
        Document document,
        int position,
        ResolvedNumberPlan plan)
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

            var segments =
                BuildSegments(
                    plan);
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

    private sealed class ResolvedNumberPlan
    {
        internal EquationNumberFormat Format { get; set; } =
            EquationNumberFormat.Resolve(
                EquationNumberFormat.ContinuousId);
        internal int Chapter { get; set; }
        internal int Section { get; set; }
    }

    private static ResolvedNumberPlan ResolveNumberPlan(
        Document document,
        int formulaPosition)
    {
        var format =
            EquationNumberFormat.Resolve(
                WordEquationNumbering
                    .GetEquationNumberFormatId(
                        document));
        var plan =
            new ResolvedNumberPlan
            {
                Format = format,
            };
        if (!format.UsesHeading)
            return plan;

        var scopes =
            WordEquationNumbering
                .CaptureHeadingScopesAtPositions(
                    document,
                    format.Id,
                    new[] { formulaPosition });
        var numberText =
            scopes.TryGetValue(
                    formulaPosition,
                    out var scope)
                ? scope.NumberText
                : string.Empty;
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

        return plan;
    }

    private static void RebuildHeadingStateFields(
        Document document,
        Range hostRange,
        ResolvedNumberPlan plan)
    {
        RemoveHeadingStateFields(
            document,
            hostRange);
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
                field.Update();
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

    private static bool IsHeadingStateRestartCode(
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(
                instruction))
            return false;
        var normalized =
            instruction!
                .TrimStart();
        var isOwnedSequence =
            normalized.StartsWith(
                "SEQ "
                + ChapterSequenceName
                + " ",
                StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith(
                "SEQ "
                + SectionSequenceName
                + " ",
                StringComparison.OrdinalIgnoreCase);
        return isOwnedSequence
            && normalized.IndexOf(
                   "\\r",
                   StringComparison.OrdinalIgnoreCase) >= 0
            && normalized.IndexOf(
                   "\\h",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private sealed class NumberSegment
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

    private static void UpdatePlaceRefFields(
        Field placeRef)
    {
        Range? code = null;
        Fields? nested = null;
        Field? field = null;
        Range? fieldCode = null;
        try
        {
            code =
                placeRef.Code.Duplicate;
            nested =
                code.Fields;

            // Heading state is read before current-value SEQ. The hidden increment
            // itself can update first because its \s switch derives from Word's
            // heading outline state rather than the STYLEREF result.
            for (var index = 1;
                 index <= nested.Count;
                 index++)
            {
                Release(fieldCode);
                fieldCode = null;
                Release(field);
                field =
                    nested[index];
                fieldCode =
                    field.Code.Duplicate;
                var instruction =
                    (fieldCode.Text
                        ?? string.Empty)
                    .TrimStart();
                if (instruction.StartsWith(
                        "STYLEREF ",
                        StringComparison.OrdinalIgnoreCase))
                    field.Update();
            }

            for (var index = 1;
                 index <= nested.Count;
                 index++)
            {
                Release(fieldCode);
                fieldCode = null;
                Release(field);
                field =
                    nested[index];
                fieldCode =
                    field.Code.Duplicate;
                var instruction =
                    (fieldCode.Text
                        ?? string.Empty)
                    .TrimStart();
                if (instruction.StartsWith(
                        "SEQ ",
                        StringComparison.OrdinalIgnoreCase))
                    field.Update();
            }

            try
            {
                placeRef.ShowCodes = false;
            }
            catch { }
        }
        finally
        {
            Release(fieldCode);
            Release(field);
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

    private static bool TryFindPlaceRef(
        Document document,
        Range hostRange,
        out Field? placeRef)
    {
        placeRef = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? separator = null;
        try
        {
            paragraphs =
                hostRange.Paragraphs;
            if (paragraphs.Count != 1)
                return false;
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
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field =
                    fields[index];
                if (field.Type !=
                    WdFieldType.wdFieldMacroButton)
                    continue;
                code =
                    field.Code.Duplicate;
                if (!IsPlaceRefCode(
                        code.Text))
                    continue;

                var fieldStart =
                    code.Start - 1;
                if (fieldStart < hostRange.End)
                    return false;
                separator =
                    document.Range(
                        hostRange.End,
                        fieldStart);
                if (!string.Equals(
                        separator.Text,
                        "\t",
                        StringComparison.Ordinal))
                    return false;

                result =
                    field.Result.Duplicate;
                var fieldEnd =
                    ResolveOuterFieldEndExclusive(
                        document,
                        code,
                        result);
                var editableEnd =
                    Math.Max(
                        paragraphRange.Start,
                        paragraphRange.End - 1);
                if (fieldEnd > editableEnd)
                    return false;

                Range? suffix = null;
                try
                {
                    suffix =
                        document.Range(
                            fieldEnd,
                            editableEnd);
                    if (!ContainsOnlyScaffoldWhitespace(
                            suffix.Text))
                        return false;
                }
                finally
                {
                    Release(suffix);
                }

                var found = field;
                field = null;
                placeRef = found;
                return true;
            }

            return false;
        }
        finally
        {
            Release(separator);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static bool TryGetVisibleNumberRange(
        Document document,
        Field placeRef,
        bool includeParentheses,
        out Range? numberRange)
    {
        numberRange = null;
        Range? outerCode = null;
        Fields? nested = null;
        Field? field = null;
        Range? fieldCode = null;
        Range? hiddenResult = null;
        try
        {
            outerCode =
                placeRef.Code.Duplicate;
            if (!IsPlaceRefCode(
                    outerCode.Text))
                return false;
            nested =
                outerCode.Fields;
            for (var index = 1;
                 index <= nested.Count;
                 index++)
            {
                Release(fieldCode);
                fieldCode = null;
                Release(field);
                field =
                    nested[index];
                fieldCode =
                    field.Code.Duplicate;
                if (!IsHiddenIncrement(
                        document,
                        fieldCode.Text))
                    continue;

                hiddenResult =
                    field.Result.Duplicate;
                var start =
                    hiddenResult.End;
                var end =
                    outerCode.End;

                var outerText =
                    outerCode.Text
                    ?? string.Empty;
                var trimmedLength =
                    outerText
                        .TrimEnd()
                        .Length;
                end -=
                    outerText.Length
                    - trimmedLength;
                if (!includeParentheses)
                {
                    start++;
                    end--;
                }
                if (end <= start)
                    return false;

                numberRange =
                    document.Range(
                        start,
                        end);
                return true;
            }

            return false;
        }
        finally
        {
            Release(hiddenResult);
            Release(fieldCode);
            Release(field);
            Release(nested);
            Release(outerCode);
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

    private static bool IsHiddenIncrement(
        Document document,
        string? instruction)
    {
        if (string.IsNullOrWhiteSpace(
                instruction))
            return false;
        var normalized =
            " "
            + instruction!
                .Replace('\t', ' ')
                .Replace('\r', ' ')
                .Replace('\n', ' ')
            + " ";
        return normalized.IndexOf(
                   " SEQ "
                   + SequenceName
                   + " ",
                   StringComparison.OrdinalIgnoreCase) >= 0
            && normalized.IndexOf(
                   "\\h",
                   StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static int ResolveOuterFieldEndExclusive(
        Document document,
        Range code,
        Range result)
    {
        var codeBoundary =
            ReadCharacter(
                document,
                code.End);
        if (codeBoundary ==
            '\u0015')
            return code.End + 1;
        if (codeBoundary ==
                '\u0014'
            && ReadCharacter(
                    document,
                    result.End)
                == '\u0015')
            return result.End + 1;
        throw new InvalidDataException(
            $"VisualTeXPlaceRef has an invalid outer field boundary at code={code.End}, result={result.End}.");
    }

    private static char ReadCharacter(
        Document document,
        int position)
    {
        Range? probe = null;
        try
        {
            if (position <
                    document.Content.Start
                || position >=
                    document.Content.End)
                return '\0';
            probe =
                document.Range(
                    position,
                    position + 1);
            var text =
                probe.Text
                ?? string.Empty;
            return text.Length == 1
                ? text[0]
                : '\0';
        }
        finally
        {
            Release(probe);
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

    private static bool ContainsOnlyScaffoldWhitespace(
        string? text)
    {
        if (string.IsNullOrEmpty(
                text))
            return true;
        foreach (var character in text!)
        {
            if (character is
                    '\t'
                    or '\r'
                    or '\n'
                    or '\v'
                    or '\u0013'
                    or '\u0014'
                    or '\u0015'
                || char.IsWhiteSpace(
                    character))
                continue;
            return false;
        }
        return true;
    }

    private static Range? ResolveVisualTeXHostRangeByFormulaId(
        Document document,
        string formulaId)
    {
        if (!Guid.TryParse(
                formulaId,
                out var parsed))
            return null;

        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? owner = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            bookmarks =
                document.Bookmarks;
            var identityName =
                WordFormulaMetadataReader
                    .IdentityBookmarkName(
                        parsed.ToString("D"));
            if (!bookmarks.Exists(
                    identityName))
                return null;

            bookmark =
                bookmarks[identityName];
            owner =
                bookmark.Range.Duplicate;
            shapes =
                owner.InlineShapes;
            if (shapes.Count != 1)
                return null;
            shape =
                shapes[1];
            if (!WordFormulaMetadataReader.IsNativeOle(
                    shape))
                return null;
            shapeRange =
                shape.Range.Duplicate;
            if (shapeRange.StoryType !=
                    owner.StoryType
                || shapeRange.Start !=
                    owner.Start
                || shapeRange.End !=
                    owner.End)
                return null;

            var result =
                shapeRange;
            shapeRange = null;
            return result;
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(owner);
            Release(bookmark);
            Release(bookmarks);
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

    private static void Release(
        object? value)
    {
        if (value is null
            || !Marshal.IsComObject(
                value))
            return;
        try
        {
            Marshal.ReleaseComObject(
                value);
        }
        catch { }
    }
}

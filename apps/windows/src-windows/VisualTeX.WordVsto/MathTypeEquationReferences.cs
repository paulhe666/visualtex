using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Native MathType equation-reference compatibility for Word.
///
/// MathType owns numbered equations with an outer MACROBUTTON MTPlaceRef field.
/// Its hidden first SEQ MTEqn \h field increments the equation counter; only the
/// visible portion after that hidden field is bookmarked.  Native MathType
/// references are a GOTOBUTTON containing a nested REF ... \! field.  Keeping
/// exactly that structure lets VisualTeX reference numbered equations created by
/// either MathType itself or VisualTeX without inventing a second numbering
/// system.
/// </summary>
internal enum EquationReferenceBookmarkSpan
{
    NumberOnly,
    VisibleNumber,
}

internal sealed class EquationReferenceBookmarkAlias
{
    internal string Name { get; set; } = string.Empty;
    internal EquationReferenceBookmarkSpan Span { get; set; }
}

internal static class MathTypeEquationReferences
{
    private const string PlaceRefMarker = "MACROBUTTON MTPlaceRef";
    private const string EquationBookmarkPrefix = "ZEqnNum";
    private const string MathTypeSectionStyleName = "MTEquationSection";

    internal static IReadOnlyList<EquationReferenceTarget> GetTargets(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        var targets = new List<EquationReferenceTarget>();
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;

                var position = Math.Max(document.Content.Start, code.Start - 1);
                Range? numberRange = null;
                try
                {
                    if (!TryGetVisibleNumberRange(document, field, out numberRange)
                        || numberRange is null)
                        continue;
                    if (!TryFindMathTypeEquationLatex(field, out var latexPreview))
                        continue;

                    var numberText = ReadVisibleNumberText(field);
                    if (string.IsNullOrWhiteSpace(numberText))
                        numberText = (numberRange.Text ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(numberText)) continue;

                    targets.Add(new EquationReferenceTarget(
                        $"mathtype:{position}",
                        -1,
                        numberText,
                        latexPreview,
                        position,
                        EquationReferenceSource.MathType));
                }
                catch
                {
                    // A malformed legacy MTPlaceRef must not hide other valid
                    // numbered MathType equations from the reference picker.
                }
                finally { Release(numberRange); }
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        return targets
            .OrderBy(target => target.Position)
            .ToArray();
    }

    internal static void InsertReference(
        Document document,
        Selection selection,
        EquationReferenceTarget target,
        WdColor? preferredInsertionColor = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        if (target is null) throw new ArgumentNullException(nameof(target));
        if (target.Source != EquationReferenceSource.MathType)
            throw new InvalidOperationException("The selected equation is not a MathType reference target.");
        if (document.ReadOnly)
            throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");

        Field? placeRef = null;
        Range? numberRange = null;
        try
        {
            placeRef = ResolvePlaceRef(document, target)
                ?? throw new InvalidDataException("The MathType reference target is missing.");
            if (!TryGetVisibleNumberRange(document, placeRef, out numberRange) || numberRange is null)
                throw new InvalidDataException("The MathType reference target has no visible number.");
            var bookmarkName = EnsureNativeMathTypeNumberBookmark(document, numberRange);
            WordEquationReferenceFields.InsertNavigableReference(document, selection, bookmarkName,
                string.Empty, string.Empty, preferredInsertionColor);
        }
        finally { Release(numberRange); Release(placeRef); }
    }
    internal static IReadOnlyList<string> CaptureReferenceBookmarkAliases(
        Document document,
        InlineShape equationShape)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (equationShape is null) throw new ArgumentNullException(nameof(equationShape));

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            shapeRange = equationShape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return Array.Empty<string>();
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;
                if (!TryGetVisibleNumberRange(document, field, out numberRange)
                    || numberRange is null)
                    continue;

                var aliases = new List<string>();
                bookmarks = document.Bookmarks;
                for (var bookmarkIndex = 1; bookmarkIndex <= bookmarks.Count; bookmarkIndex++)
                {
                    Release(bookmarkRange);
                    bookmarkRange = null;
                    Release(bookmark);
                    bookmark = bookmarks[bookmarkIndex];
                    if (!bookmark.Name.StartsWith(
                            EquationBookmarkPrefix,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    bookmarkRange = bookmark.Range;
                    if (bookmarkRange.Start != numberRange.Start
                        || bookmarkRange.End != numberRange.End)
                        continue;
                    aliases.Add(bookmark.Name);
                }
                return aliases
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
            }
            return Array.Empty<string>();
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(numberRange);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static IReadOnlyList<EquationReferenceBookmarkAlias> CaptureFormatConversionAliasesFromMathType(
        Document document,
        InlineShape equationShape)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (equationShape is null) throw new ArgumentNullException(nameof(equationShape));

        Range? shapeRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? ownerRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? visibleRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            shapeRange = equationShape.Range;
            paragraphs = shapeRange.Paragraphs;
            if (paragraphs.Count != 1) return Array.Empty<EquationReferenceBookmarkAlias>();
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range.Duplicate;
            fields = WordFormulaHost.GetLocalFields(ownerRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;
                if (!TryGetVisibleNumberRange(document, field, out visibleRange)
                    || visibleRange is null)
                    continue;
                break;
            }
            if (visibleRange is null) return Array.Empty<EquationReferenceBookmarkAlias>();

            Range? numberOnlyRange = null;
            try
            {
                numberOnlyRange = NumberInsideVisibleDelimiters(document, visibleRange);
                return CaptureNumberAliases(document, ownerRange, visibleRange, numberOnlyRange);
            }
            finally { Release(numberOnlyRange); }
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(visibleRange);
            Release(code);
            Release(field);
            Release(fields);
            Release(ownerRange);
            Release(paragraph);
            Release(paragraphs);
            Release(shapeRange);
        }
    }

    internal static IReadOnlyList<EquationReferenceBookmarkAlias> CaptureFormatConversionAliasesFromVisualTeX(
        Document document,
        string formulaId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new ArgumentException("FormulaId is required.", nameof(formulaId));

        var aliasWatch = System.Diagnostics.Stopwatch.StartNew();
        var aliasTrace = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_TRACE_FORMAT_PERF") == "1";
        long aliasCheckpoint = 0;
        void TraceAlias(string stage)
        {
            if (!aliasTrace) return;
            var elapsed = aliasWatch.ElapsedMilliseconds;
            WordDoubleClickHook.TraceMessage($"conversion-alias-perf stage={stage} deltaMs={elapsed - aliasCheckpoint} totalMs={elapsed}");
            aliasCheckpoint = elapsed;
        }
        Range? ownerRange = null;
        Range? visibleNumberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId);
            TraceAlias("owner");
            visibleNumberRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                document,
                formulaId);
            if (ownerRange is null || visibleNumberRange is null)
                return Array.Empty<EquationReferenceBookmarkAlias>();

            bookmarks = document.Bookmarks;
            var nativeName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(nativeName))
                throw new InvalidDataException($"The numbered formula {formulaId} has no number identity.");
            bookmark = bookmarks[nativeName];
            bookmarkRange = bookmark.Range.Duplicate;
            TraceAlias("number-ranges");
            var result = CaptureNumberAliases(document, ownerRange, visibleNumberRange, bookmarkRange);
            TraceAlias("local-aliases-and-references");
            return result;
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(visibleNumberRange);
            Release(ownerRange);
        }
    }

    private static IReadOnlyList<EquationReferenceBookmarkAlias> CaptureNumberAliases(
        Document document, Range ownerRange, Range visibleRange, Range numberOnlyRange)
    {
        var aliases = new List<EquationReferenceBookmarkAlias>();
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Bookmarks? bookmarks = null;
        try
        {
            // An alias must be fully contained in one of these two number spans
            // to be transferable. Enumerating all document bookmarks only to
            // reject unrelated ranges made one conversion O(total formulas).
            foreach (var numberSpan in new[] { visibleRange, numberOnlyRange })
            {
            Release(bookmarks); bookmarks = null;
            bookmarks = numberSpan.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    var name = bookmark.Name;
                    if (!visited.Add(name)) continue;
                    var numberOnly = MathTypeWordOpenXml.IsVisualTeXNumberAlias(name);
                    if (!numberOnly && !name.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var expected = numberOnly ? numberOnlyRange : visibleRange;
                    range = bookmark.Range;
                    if (range.StoryType != expected.StoryType || range.Start < expected.Start
                        || range.End > expected.End || range.Start == range.End)
                        continue;
                    if (!HasExternalReferenceToBookmark(document, name, ownerRange)) continue;
                    if (!string.Equals(range.Text, expected.Text, StringComparison.Ordinal))
                        throw new InvalidDataException($"Reference alias '{name}' does not own its complete equation number.");
                    aliases.Add(new EquationReferenceBookmarkAlias
                    {
                        Name = name,
                        Span = numberOnly ? EquationReferenceBookmarkSpan.NumberOnly
                            : EquationReferenceBookmarkSpan.VisibleNumber,
                    });
                }
                finally { Release(range); Release(bookmark); }
            }
            }
            return aliases.OrderBy(alias => alias.Name, StringComparer.OrdinalIgnoreCase).ToArray();
        }
        finally { Release(bookmarks); }
    }

    internal static int RestoreFormatConversionAliasesToVisualTeX(
        Document document,
        string formulaId,
        IReadOnlyCollection<EquationReferenceBookmarkAlias> aliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (aliases is null || aliases.Count == 0) return 0;

        Bookmarks? bookmarks = null;
        Bookmark? nativeBookmark = null;
        Range? numberOnlyRange = null;
        Range? visibleRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            var nativeName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(nativeName))
                throw new InvalidDataException(
                    $"Converted VisualTeX formula {formulaId} has no durable number bookmark {nativeName}.");
            nativeBookmark = bookmarks[nativeName];
            numberOnlyRange = nativeBookmark.Range.Duplicate;

            visibleRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(
                    $"Converted VisualTeX formula {formulaId} has no visible equation-number slot.");

            var restored = 0;
            foreach (var alias in aliases
                         .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                         .GroupBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                var targetRange = alias.Span == EquationReferenceBookmarkSpan.VisibleNumber
                    ? visibleRange
                    : numberOnlyRange;
                var expectedSpan = MathTypeWordOpenXml.IsVisualTeXNumberAlias(alias.Name)
                    ? EquationReferenceBookmarkSpan.NumberOnly : EquationReferenceBookmarkSpan.VisibleNumber;
                if ((!MathTypeWordOpenXml.IsVisualTeXNumberAlias(alias.Name)
                        && !alias.Name.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                    || alias.Span != expectedSpan)
                    throw new InvalidDataException($"Number alias '{alias.Name}' changed its reference span contract.");
                BindAliasRange(document, targetRange, alias.Name);
                restored++;
            }
            return restored;
        }
        finally
        {
            Release(visibleRange);
            Release(numberOnlyRange);
            Release(nativeBookmark);
            Release(bookmarks);
        }
    }

    internal static Range ResolveNumberAliasRange(Document document, Range visibleRange, string name)
    {
        if (name.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
            return visibleRange.Duplicate;
        if (!MathTypeWordOpenXml.IsVisualTeXNumberAlias(name))
            throw new InvalidDataException($"Unknown MathType number alias '{name}'.");
        return NumberInsideVisibleDelimiters(document, visibleRange);
    }

    private static Range NumberInsideVisibleDelimiters(Document document, Range visibleRange)
    {
        Range? first = null;
        Range? last = null;
        try
        {
            if (visibleRange.End - visibleRange.Start < 2)
                throw new InvalidDataException("The MathType visible number has no paired delimiters.");
            first = document.Range(visibleRange.Start, visibleRange.Start + 1);
            last = document.Range(visibleRange.End - 1, visibleRange.End);
            if (first.Text != "(" || last.Text != ")")
                throw new InvalidDataException("The VisualTeX compatibility alias must own only the number inside MathType's parentheses.");
            return document.Range(visibleRange.Start + 1, visibleRange.End - 1);
        }
        finally { Release(last); Release(first); }
    }

    internal static void BindNumberAlias(Document document, Range visibleRange, string name)
    {
        Range? expected = null;
        try
        {
            expected = ResolveNumberAliasRange(document, visibleRange, name);
            BindAliasRange(document, expected, name);
        }
        finally { Release(expected); }
    }

    private static void BindAliasRange(Document document, Range expected, string name)
    {
        Range? actual = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            bookmark = bookmarks.Add(name, expected);
            actual = bookmark.Range;
            if (actual.Start != expected.Start || actual.End != expected.End || actual.StoryType != expected.StoryType)
                throw new InvalidDataException($"Word did not retain the exact number range for alias '{name}'.");
        }
        finally { Release(actual); Release(bookmark); Release(bookmarks); }
    }

    internal static int RestoreFormatConversionAliasesToMathType(
        Document document,
        string targetBookmarkName,
        IReadOnlyCollection<EquationReferenceBookmarkAlias> aliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (aliases is null || aliases.Count == 0) return 0;
        if (string.IsNullOrWhiteSpace(targetBookmarkName))
            throw new ArgumentException("Target bookmark is required.", nameof(targetBookmarkName));

        Bookmarks? bookmarks = null;
        Bookmark? targetBookmark = null;
        Range? targetRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? visibleRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(targetBookmarkName))
                throw new InvalidDataException(
                    $"Converted MathType target locator {targetBookmarkName} is missing.");
            targetBookmark = bookmarks[targetBookmarkName];
            targetRange = targetBookmark.Range;
            paragraphs = targetRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("Converted MathType target is not in one stable paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            fields = WordFormulaHost.GetLocalFields(paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;
                if (!TryGetVisibleNumberRange(document, field, out visibleRange)
                    || visibleRange is null)
                    continue;
                break;
            }
            if (visibleRange is null)
                throw new InvalidDataException("Converted MathType target has no MTPlaceRef visible number range.");

            var restored = 0;
            foreach (var alias in aliases
                         .Where(item => !string.IsNullOrWhiteSpace(item.Name))
                         .GroupBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                         .Select(group => group.First()))
            {
                var expectedSpan = MathTypeWordOpenXml.IsVisualTeXNumberAlias(alias.Name)
                    ? EquationReferenceBookmarkSpan.NumberOnly : EquationReferenceBookmarkSpan.VisibleNumber;
                if (alias.Span != expectedSpan)
                    throw new InvalidDataException($"Number alias '{alias.Name}' changed its reference span contract.");
                BindNumberAlias(document, visibleRange, alias.Name);
                restored++;
            }
            return restored;
        }
        finally
        {
            Release(visibleRange);
            Release(code);
            Release(field);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(targetRange);
            Release(targetBookmark);
            Release(bookmarks);
        }
    }

    internal static IReadOnlyDictionary<string, IReadOnlyList<WordCharacterFormatting>>
        CaptureReferenceCharacterFormatting(
            Document document,
            IEnumerable<string> bookmarkAliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var aliases = new HashSet<string>(
            bookmarkAliases.Where(name => !string.IsNullOrWhiteSpace(name)),
            StringComparer.OrdinalIgnoreCase);
        if (aliases.Count == 0)
            return new Dictionary<string, IReadOnlyList<WordCharacterFormatting>>(
                StringComparer.OrdinalIgnoreCase);

        var captured = aliases.ToDictionary(
            alias => alias,
            _ => new List<(int Start, WordCharacterFormatting Formatting)>(),
            StringComparer.OrdinalIgnoreCase);
        var bookmarkRanges = new Dictionary<string, (int Start, int End)>(
            StringComparer.OrdinalIgnoreCase);
        var sourceOwnerRanges = new Dictionary<string, (int Start, int End)>(
            StringComparer.OrdinalIgnoreCase);
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var alias in aliases)
            {
                if (bookmarks.Exists(alias))
                {
                    Release(bookmarkRange);
                    bookmarkRange = null;
                    Release(bookmark);
                    bookmark = bookmarks[alias];
                    bookmarkRange = bookmark.Range;
                    bookmarkRanges[alias] = (bookmarkRange.Start, bookmarkRange.End);
                }

                const string nativeNumberPrefix = "VTEqNum_";
                if (!alias.StartsWith(nativeNumberPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                var rawFormulaId = alias.Substring(nativeNumberPrefix.Length);
                if (!Guid.TryParseExact(rawFormulaId, "N", out var formulaGuid))
                    continue;
                Range? ownerRange = null;
                try
                {
                    ownerRange = WordEquationNumbering.FindNumberingOwnerRange(
                        document,
                        formulaGuid.ToString("D"));
                    if (ownerRange is null) continue;
                    sourceOwnerRanges[alias] = (ownerRange.Start, ownerRange.End);
                }
                finally
                {
                    Release(ownerRange);
                }
            }

            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result);
                result = null;
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty).TrimStart();
                if (!text.StartsWith("REF ", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!WordEquationReferenceFields.TryReadVisualTeXNumberBookmark(text, out var alias)
                    || !aliases.Contains(alias)) continue;
                if (sourceOwnerRanges.TryGetValue(alias, out var ownerRange)
                    && code.Start >= ownerRange.Start
                    && code.Start < ownerRange.End)
                    continue;
                result = field.Result;
                if (bookmarkRanges.TryGetValue(alias, out var bookmarkTarget)
                    && result.Start < bookmarkTarget.End
                    && result.End > bookmarkTarget.Start)
                    continue;
                captured[alias].Add((
                    code.Start,
                    WordCharacterFormatting.Capture(result)));
            }
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }

        return captured.ToDictionary(
            entry => entry.Key,
            entry => (IReadOnlyList<WordCharacterFormatting>)entry.Value
                .OrderBy(item => item.Start)
                .Select(item => item.Formatting)
                .ToArray(),
            StringComparer.OrdinalIgnoreCase);
    }

    internal static int RestoreReferenceCharacterFormatting(
        Document document,
        IReadOnlyDictionary<string, IReadOnlyList<WordCharacterFormatting>> captured)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (captured is null || captured.Count == 0) return 0;

        var live = captured.Keys.ToDictionary(
            alias => alias,
            _ => new List<(int Start, Field Field)>(),
            StringComparer.OrdinalIgnoreCase);
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                var text = (code.Text ?? string.Empty).TrimStart();
                if (!text.StartsWith("REF ", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!WordEquationReferenceFields.TryReadVisualTeXNumberBookmark(text, out var alias)
                    || !captured.ContainsKey(alias)) continue;
                live[alias].Add((code.Start, field));
                field = null;
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        var restored = 0;
        try
        {
            foreach (var entry in captured)
            {
                var liveFields = live[entry.Key]
                    .OrderBy(item => item.Start)
                    .ToArray();
                if (entry.Value.Count != liveFields.Length)
                    throw new InvalidDataException($"Reference alias '{entry.Key}' retained {liveFields.Length}/{entry.Value.Count} fields.");
                var count = liveFields.Length;
                for (var index = 0; index < count; index++)
                {
                    Range? result = null;
                    try
                    {
                        result = liveFields[index].Field.Result;
                        entry.Value[index].Apply(result);
                        Range? referenceCode = null;
                        try { referenceCode = liveFields[index].Field.Code; entry.Value[index].Apply(referenceCode); }
                        finally { Release(referenceCode); }
                        restored++;
                    }
                    finally { Release(result); }
                }
            }
            return restored;
        }
        finally
        {
            foreach (var entries in live.Values)
            foreach (var item in entries)
                Release(item.Field);
        }
    }

    internal static int FreezeReferencesToPlainText(
        Document document,
        IReadOnlyCollection<string> bookmarkAliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (bookmarkAliases is null || bookmarkAliases.Count == 0) return 0;

        var aliases = new HashSet<string>(
            bookmarkAliases.Where(name => !string.IsNullOrWhiteSpace(name)),
            StringComparer.OrdinalIgnoreCase);
        if (aliases.Count == 0) return 0;

        var replacements = new List<(int Start, int End, string Text)>();
        Fields? fields = null;
        Field? outer = null;
        Range? outerCode = null;
        Range? outerResult = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? nestedCode = null;
        Range? nestedResult = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(nestedResult);
                nestedResult = null;
                Release(nestedCode);
                nestedCode = null;
                Release(nested);
                nested = null;
                Release(nestedFields);
                nestedFields = null;
                Release(outerResult);
                outerResult = null;
                Release(outerCode);
                outerCode = null;
                Release(outer);
                outer = fields[index];
                outerCode = outer.Code;
                var outerText = (outerCode.Text ?? string.Empty)
                    .Replace('\r', ' ')
                    .Replace('\n', ' ')
                    .Replace('\t', ' ')
                    .TrimStart();
                if (!outerText.StartsWith("GOTOBUTTON ", StringComparison.OrdinalIgnoreCase)
                    || !aliases.Any(alias =>
                        outerText.IndexOf(alias, StringComparison.OrdinalIgnoreCase) >= 0))
                    continue;

                nestedFields = outerCode.Fields;
                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(nestedResult);
                    nestedResult = null;
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nested);
                    nested = nestedFields[nestedIndex];
                    nestedCode = nested.Code;
                    var nestedText = nestedCode.Text ?? string.Empty;
                    if (!nestedText.TrimStart().StartsWith("REF ", StringComparison.OrdinalIgnoreCase)
                        || !aliases.Any(alias =>
                            nestedText.IndexOf(alias, StringComparison.OrdinalIgnoreCase) >= 0))
                        continue;
                    nestedResult = nested.Result;
                    outerResult = outer.Result;
                    var fullStart = Math.Max(document.Content.Start, outerCode.Start - 1);
                    var fullEnd = Math.Min(
                        document.Content.End,
                        Math.Max(fullStart, outerResult.End + 1));
                    replacements.Add((
                        fullStart,
                        fullEnd,
                        nestedResult.Text ?? string.Empty));
                    break;
                }
            }
        }
        finally
        {
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerResult);
            Release(outerCode);
            Release(outer);
            Release(fields);
        }

        var frozen = 0;
        foreach (var replacement in replacements
                     .OrderByDescending(item => item.Start))
        {
            Range? range = null;
            try
            {
                range = document.Range(replacement.Start, replacement.End);
                range.Text = replacement.Text;
                frozen++;
            }
            finally { Release(range); }
        }
        return frozen;
    }

    internal static int RefreshReferences(Document document, ISet<string> bookmarkAliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (bookmarkAliases is null || bookmarkAliases.Count == 0) return 0;
        return WordEquationReferenceFields.UpdateReferences(document, bookmarkAliases);
    }
    private static bool HasExternalReferenceToBookmark(
        Document document,
        string bookmarkName,
        Range ownerRange)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                var type = field.Type;
                if (type != WdFieldType.wdFieldRef && type != WdFieldType.wdFieldGoToButton) continue;
                code = field.Code;
                if (code.Start >= ownerRange.Start && code.Start < ownerRange.End)
                    continue;
                var text = code.Text ?? string.Empty;
                if (text.IndexOf(bookmarkName, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                if (text.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) >= 0
                    || text.IndexOf("GOTOBUTTON ", StringComparison.OrdinalIgnoreCase) >= 0)
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

    private static Field? ResolvePlaceRef(
        Document document,
        EquationReferenceTarget target)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Field? best = null;
        var bestDistance = int.MaxValue;
        try
        {
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypePlaceRefCode(code.Text)) continue;
                var position = Math.Max(document.Content.Start, code.Start - 1);
                var distance = Math.Abs(position - target.Position);
                if (distance > 8 || distance >= bestDistance) continue;

                Range? numberRange = null;
                try
                {
                    if (!TryGetVisibleNumberRange(document, field, out numberRange)
                        || numberRange is null)
                        continue;
                    var numberText = ReadVisibleNumberText(field);
                    if (!string.Equals(
                            numberText.Trim(),
                            target.NumberText.Trim(),
                            StringComparison.Ordinal))
                        continue;
                }
                finally { Release(numberRange); }

                Release(best);
                best = field;
                field = null;
                bestDistance = distance;
                if (distance == 0) break;
            }
            return best;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    internal static bool TryGetVisibleNumberRange(
        Document document,
        Field placeRef,
        out Range? numberRange)
    {
        numberRange = null;
        Range? outerCode = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? nestedCode = null;
        Range? hiddenResult = null;
        try
        {
            outerCode = placeRef.Code;
            if (!IsMathTypePlaceRefCode(outerCode.Text)) return false;
            nestedFields = outerCode.Fields;
            for (var index = 1; index <= nestedFields.Count; index++)
            {
                Release(nestedCode);
                nestedCode = null;
                Release(nested);
                nested = nestedFields[index];
                nestedCode = nested.Code;
                if (!IsHiddenMathTypeEquationIncrement(nestedCode.Text)) continue;
                hiddenResult = nested.Result;
                var start = hiddenResult.End;
                var end = outerCode.End;

                // The visible MathType number ends at its closing punctuation. A
                // malformed/legacy MTPlaceRef can contain one or more direct spaces
                // after that punctuation; bookmarking them makes every REF result
                // visibly render the same unwanted blanks. Trim only trailing
                // whitespace owned by the outer field code, never number content.
                var outerText = outerCode.Text ?? string.Empty;
                var trimmedLength = outerText.TrimEnd().Length;
                end -= outerText.Length - trimmedLength;
                if (end <= start) return false;
                numberRange = document.Range(start, end);
                return true;
            }
            return false;
        }
        finally
        {
            Release(hiddenResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
        }
    }

    internal static string ReadVisibleNumberText(Field placeRef)
    {
        Range? outerCode = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? nestedCode = null;
        Range? nestedResult = null;
        try
        {
            outerCode = placeRef.Code;
            var stream = outerCode.Text ?? string.Empty;
            nestedFields = outerCode.Fields;
            if (nestedFields.Count == 0) return string.Empty;

            var hiddenControlStart = -1;
            var hiddenControlEnd = -1;
            var nestedOrdinal = 0;
            for (var index = 0; index < stream.Length; index++)
            {
                if (stream[index] != '\u0013') continue;
                nestedOrdinal++;
                var end = stream.IndexOf('\u0015', index + 1);
                if (end < 0) return string.Empty;
                if (nestedOrdinal <= nestedFields.Count)
                {
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nested);
                    nested = nestedFields[nestedOrdinal];
                    nestedCode = nested.Code;
                    if (IsHiddenMathTypeEquationIncrement(nestedCode.Text))
                    {
                        hiddenControlStart = index;
                        hiddenControlEnd = end;
                        break;
                    }
                }
                index = end;
            }
            if (hiddenControlStart < 0 || hiddenControlEnd < 0) return string.Empty;

            var visible = new System.Text.StringBuilder();
            var fieldOrdinal = nestedOrdinal;
            for (var index = hiddenControlEnd + 1; index < stream.Length;)
            {
                if (stream[index] != '\u0013')
                {
                    if (stream[index] is not '\u0014' and not '\u0015')
                        visible.Append(stream[index]);
                    index++;
                    continue;
                }

                var end = stream.IndexOf('\u0015', index + 1);
                if (end < 0) break;
                fieldOrdinal++;
                if (fieldOrdinal <= nestedFields.Count)
                {
                    Release(nestedResult);
                    nestedResult = null;
                    Release(nested);
                    nested = nestedFields[fieldOrdinal];
                    nestedResult = nested.Result;
                    visible.Append(nestedResult.Text ?? string.Empty);
                }
                index = end + 1;
            }
            return visible.ToString().Trim();
        }
        catch { return string.Empty; }
        finally
        {
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
        }
    }

    private static bool TryFindMathTypeEquationLatex(Field placeRef, out string latex)
    {
        // Equation-reference discovery must stay a lightweight Word-field query.
        // Reading Equation Native here used to open every OLE storage merely to
        // populate the picker preview; after save/reopen Word can synchronously
        // block while materializing that OLE storage. A native reference only
        // needs to prove that the MTPlaceRef belongs to an Equation.DSMT4 object.
        // Keep the preview generic and leave MTEF reads to the explicit editor.
        latex = "MathType 公式";
        Range? code = null;
        Range? probe = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? paragraphRange = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        try
        {
            code = placeRef.Code;
            var document = code.Document;
            var position = Math.Max(document.Content.Start, code.Start - 1);
            probe = document.Range(position, position);
            paragraphs = probe.Paragraphs;
            if (paragraphs.Count != 1) return false;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                Release(shape);
                shape = shapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) return true;
            }
            return false;
        }
        finally
        {
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(probe);
            Release(code);
        }
    }

    private static string EnsureNativeMathTypeNumberBookmark(
        Document document,
        Range numberRange)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!bookmark.Name.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start == numberRange.Start
                    && bookmarkRange.End == numberRange.End)
                    return bookmark.Name;
            }

            var seed = 100000 + Math.Abs(numberRange.Start % 900000);
            for (var offset = 0; offset < 900000; offset++)
            {
                var value = 100000 + ((seed - 100000 + offset) % 900000);
                var name = EquationBookmarkPrefix
                    + value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (bookmarks.Exists(name)) continue;
                Release(bookmark);
                bookmark = bookmarks.Add(name, numberRange);
                return name;
            }
            throw new InvalidOperationException("无法为 MathType 公式编号创建唯一引用书签。");
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static bool IsMathTypePlaceRefCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(PlaceRefMarker, StringComparison.OrdinalIgnoreCase) >= 0;

    internal static bool IsHiddenMathTypeEquationIncrement(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var normalized = " " + code!.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' ') + " ";
        return normalized.IndexOf(" SEQ MTEqn ", StringComparison.OrdinalIgnoreCase) >= 0
            && normalized.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); }
        catch { }
    }
}

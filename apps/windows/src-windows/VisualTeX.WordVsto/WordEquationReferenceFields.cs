using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Creates a real Word equation reference: an outer GOTOBUTTON field containing
/// one nested REF field. The nested REF keeps the displayed number live, while
/// the outer field makes double-click navigation work consistently for VisualTeX,
/// OMML and MathType numbering bookmarks.
/// </summary>
internal static class WordEquationReferenceFields
{
    private const string MathTypeSectionStyleName = "MTEquationSection";

    internal static void InsertNavigableReference(
        Document document,
        Selection selection,
        string bookmarkName,
        string prefix,
        string suffix,
        WdColor? preferredInsertionColor = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        if (string.IsNullOrWhiteSpace(bookmarkName))
            throw new ArgumentException("Equation reference bookmark is required.", nameof(bookmarkName));
        if (document.ReadOnly)
            throw new UnauthorizedAccessException("当前 Word 文档为只读状态。");

        Bookmarks? bookmarks = null;
        Range? sourceFormattingRange = null;
        Range? insertion = null;
        Field? goToField = null;
        Range? goToCode = null;
        Range? nestedInsertion = null;
        Field? refField = null;
        Fields? insertionFields = null;
        Fields? finalNestedFields = null;
        Field? finalRefField = null;
        Range? finalRefCode = null;
        Range? finalRefResult = null;
        Range? goToResult = null;
        Range? selectionRange = null;
        var phase = "capture insertion formatting";
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
                throw new InvalidDataException(
                    $"公式引用目标书签“{bookmarkName}”已不存在。文档内容可能已发生变化。");

            sourceFormattingRange = selection.Range.Duplicate;
            sourceFormattingRange.Collapse(WdCollapseDirection.wdCollapseStart);
            var formatting = WordCharacterFormatting.Capture(sourceFormattingRange);
            if (preferredInsertionColor.HasValue)
            {
                var requested = preferredInsertionColor.Value;
                formatting.Color = requested == WdColor.wdColorAutomatic
                    || (int)requested >= 0
                        ? requested
                        : WdColor.wdColorAutomatic;
            }

            var referenceStart = selection.Start;
            phase = "insert reference delimiters";
            // Keep both literal delimiters outside the field before creating it.
            // Typing a suffix at a stale field-end Range can split the nested REF.
            if (!string.IsNullOrEmpty(prefix) || !string.IsNullOrEmpty(suffix))
                selection.TypeText(prefix + suffix);
            selection.SetRange(referenceStart + prefix.Length, referenceStart + prefix.Length);

            insertion = selection.Range.Duplicate;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            var placeholder = "VTREF_" + Guid.NewGuid().ToString("N");
            phase = "create GOTOBUTTON field";
            goToField = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldGoToButton,
                ResolveNavigationBookmark(document, bookmarkName) + " " + placeholder,
                false);
            if (goToField is null)
                throw new InvalidDataException("Word did not return the inserted GOTOBUTTON field.");
            goToField.ShowCodes = true;

            // The nested REF is deliberately placed inside GOTOBUTTON.Code. Word
            // renders its result as the visible number and routes a double-click
            // on the enclosing field to the bookmark.
            goToCode = goToField.Code;
            phase = "resolve nested REF insertion slot";
            NormalizeInternalStyle(goToCode);
            formatting.Apply(goToCode);
            var placeholderOffset = (goToCode.Text ?? throw new InvalidDataException("Word returned an empty GOTOBUTTON code."))
                .IndexOf(placeholder, StringComparison.Ordinal);
            if (placeholderOffset < 0)
                throw new InvalidDataException("Word lost the reference's internal insertion slot.");
            nestedInsertion = document.Range(goToCode.Start + placeholderOffset,
                goToCode.Start + placeholderOffset + placeholder.Length);
            if (nestedInsertion.Text != placeholder)
                throw new InvalidDataException("The nested reference insertion slot is not inside its field code.");
            // Remove only the exact private slot we just created. Word's
            // Fields.Add does not reliably replace noncollapsed code text.
            // A collapsed insertion inside the code avoids that replacement.
            var nestedStart = nestedInsertion.Start;
            nestedInsertion.Text = string.Empty;
            nestedInsertion.SetRange(nestedStart, nestedStart);
            phase = "insert nested REF field";
            insertionFields = nestedInsertion.Fields;
            refField = insertionFields.Add(
                nestedInsertion,
                WdFieldType.wdFieldRef,
                bookmarkName + " \\* CHARFORMAT \\!",
                false);
            // Word can return null for a successful insertion into another
            // field's code. Read the actual field tree; never repeat the insert.
            Release(refField);
            refField = null;
            Release(goToCode);
            goToCode = goToField.Code;
            finalNestedFields = goToCode.Fields;
            if (finalNestedFields.Count != 1)
                throw new InvalidDataException($"Word created {finalNestedFields.Count} nested reference fields; exactly one is required.");
            refField = finalNestedFields[1];
            finalRefCode = refField.Code;
            if (refField.Type != WdFieldType.wdFieldRef
                || !TryReadVisualTeXNumberBookmark(finalRefCode.Text, out var insertedBookmark)
                || !string.Equals(insertedBookmark, bookmarkName, StringComparison.OrdinalIgnoreCase)
                || finalRefCode.Start <= goToCode.Start || finalRefCode.End >= goToCode.End)
                throw new InvalidDataException("Word created a reference outside its intended field or to a different bookmark.");
            Release(finalRefCode);
            finalRefCode = null;
            Release(finalNestedFields);
            finalNestedFields = null;
            phase = "refresh nested REF field";
            refField.ShowCodes = false;
            goToField.ShowCodes = false;

            // Updating the nested field can rematerialize the complete outer field
            // and reapply Word's default red GOTOBUTTON formatting. Reacquire and
            // normalize the final field tree after the update.
            refField.Update();
            Release(goToCode);
            goToCode = goToField.Code;
            NormalizeInternalStyle(goToCode);
            formatting.Apply(goToCode);

            finalNestedFields = goToCode.Fields;
            phase = "validate refreshed reference field tree";
            if (finalNestedFields.Count != 1)
                throw new InvalidDataException(
                    "公式引用未能保留一个完整的嵌套 REF 字段。");
            finalRefField = finalNestedFields[1];
            finalRefCode = finalRefField.Code;
            NormalizeInternalStyle(finalRefCode);
            formatting.Apply(finalRefCode);
            finalRefField.Update();

            Release(goToCode);
            goToCode = goToField.Code;
            NormalizeInternalStyle(goToCode);
            formatting.Apply(goToCode);

            Release(finalRefResult);
            finalRefResult = null;
            Release(finalRefCode);
            finalRefCode = null;
            Release(finalRefField);
            finalRefField = null;
            Release(finalNestedFields);
            finalNestedFields = null;
            finalNestedFields = goToCode.Fields;
            if (finalNestedFields.Count != 1)
                throw new InvalidDataException(
                    "公式引用在刷新后丢失了嵌套 REF 字段。");
            finalRefField = finalNestedFields[1];
            finalRefResult = finalRefField.Result;
            NormalizeInternalStyle(finalRefResult);
            formatting.Apply(finalRefResult);
            if ((goToCode.Text ?? string.Empty).IndexOf("VTREF_", StringComparison.Ordinal) >= 0
                || !HasCurrentReferenceResult(document, bookmarkName, finalRefResult))
                throw new InvalidDataException("The created reference does not display its current equation number cleanly.");
            Range? trailingCode = null;
            try
            {
                trailingCode = document.Range(finalRefResult.End + 1, goToCode.End);
                if (!string.IsNullOrWhiteSpace(trailingCode.Text))
                    throw new InvalidDataException("Unexpected text follows the number in the navigable reference.");
                // GOTOBUTTON renders code whitespace after its display text.
                // Remove only the verified whitespace in this new field's own
                // code, leaving field markers and document text untouched.
                if (trailingCode.End > trailingCode.Start) trailingCode.Text = string.Empty;
            }
            finally { Release(trailingCode); }
            Release(goToCode);
            goToCode = goToField.Code;

            goToResult = goToField.Result;
            phase = "restore insertion point after reference";
            var after = Math.Max(goToResult.End + 1, goToCode.End + 2);
            after = Math.Max(
                document.Content.Start,
                Math.Min(after, Math.Max(document.Content.Start, document.Content.End - 1)));
            selection.SetRange(after + suffix.Length, after + suffix.Length);
            selectionRange = selection.Range;
            NormalizeInternalStyle(selectionRange);
            formatting.Apply(selectionRange);
        }
        catch (Exception error)
        {
            throw new InvalidOperationException($"Could not {phase} for equation reference '{bookmarkName}'.", error);
        }
        finally
        {
            Release(selectionRange);
            Release(goToResult);
            Release(finalRefResult);
            Release(finalRefCode);
            Release(finalRefField);
            Release(finalNestedFields);
            Release(refField);
            Release(insertionFields);
            Release(nestedInsertion);
            Release(goToCode);
            Release(goToField);
            Release(insertion);
            Release(sourceFormattingRange);
            Release(bookmarks);
        }
    }

    internal static int FreezeNavigableReferences(
        Document document,
        string targetBookmarkName)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(targetBookmarkName)) return 0;

        Fields? outerFields = null;
        var frozen = 0;
        try
        {
            outerFields = document.Fields;
            for (var outerIndex = outerFields.Count; outerIndex >= 1; outerIndex--)
            {
                Field? outerField = null;
                Range? outerCode = null;
                Fields? nestedFields = null;
                Field? nestedField = null;
                Range? nestedCode = null;
                try
                {
                    outerField = outerFields[outerIndex];
                    if (outerField.Type != WdFieldType.wdFieldGoToButton)
                        continue;
                    outerCode = outerField.Code;
                    nestedFields = outerCode.Fields;
                    var matches = false;
                    for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                    {
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nestedField);
                        nestedField = nestedFields[nestedIndex];
                        if (nestedField.Type != WdFieldType.wdFieldRef)
                            continue;
                        nestedCode = nestedField.Code;
                        if (!TryReadVisualTeXNumberBookmark(
                                nestedCode.Text,
                                out var bookmarkName)
                            || !string.Equals(
                                bookmarkName,
                                targetBookmarkName,
                                StringComparison.OrdinalIgnoreCase))
                            continue;
                        RefreshFieldPreservingFormatting(nestedField);
                        matches = true;
                        break;
                    }
                    if (!matches) continue;

                    // The nested REF lives in GOTOBUTTON.Code, so document.Fields
                    // never enumerates it as an ordinary top-level REF. Unlinking
                    // only after the nested result is current lets Word replace the
                    // complete navigable field tree with exactly the visible number.
                    // This is the required semantic when the target equation itself
                    // is restored to plain LaTeX and its bookmark is about to vanish.
                    outerField.Unlink();
                    frozen++;
                }
                finally
                {
                    Release(nestedCode);
                    Release(nestedField);
                    Release(nestedFields);
                    Release(outerCode);
                    Release(outerField);
                }
            }
        }
        finally { Release(outerFields); }
        return frozen;
    }

    internal static int UpdateNavigableReferences(
        Document document,
        ISet<string>? targetBookmarkNames = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        Fields? outerFields = null;
        var updated = 0;
        var failures = new List<Exception>();
        try
        {
            outerFields = document.Fields;
            for (var outerIndex = 1; outerIndex <= outerFields.Count; outerIndex++)
            {
                Field? outerField = null;
                Range? outerCode = null;
                Fields? nestedFields = null;
                Field? nestedField = null;
                Range? nestedCode = null;
                Range? nestedResult = null;
                try
                {
                    outerField = outerFields[outerIndex];
                    if (outerField.Type != WdFieldType.wdFieldGoToButton)
                        continue;

                    outerCode = outerField.Code;
                    nestedFields = outerCode.Fields;
                    for (var nestedIndex = 1;
                         nestedIndex <= nestedFields.Count;
                         nestedIndex++)
                    {
                        Release(nestedResult);
                        nestedResult = null;
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nestedField);
                        nestedField = nestedFields[nestedIndex];
                        if (nestedField.Type != WdFieldType.wdFieldRef)
                            continue;

                        nestedCode = nestedField.Code;
                        if (!TryReadVisualTeXNumberBookmark(
                                nestedCode.Text,
                                out var bookmarkName))
                            continue;
                        if (targetBookmarkNames is not null
                            && !targetBookmarkNames.Contains(bookmarkName))
                            continue;

                        nestedResult = nestedField.Result;
                        if (HasCurrentReferenceResult(document, bookmarkName, nestedResult))
                        {
                            if (UpdateNavigationTarget(document, outerField, bookmarkName)) updated++;
                            break;
                        }
                        var formatting = WordCharacterFormatting.Capture(nestedResult);

                        // The REF is nested inside GOTOBUTTON.Code and therefore is
                        // not part of document.Fields' top-level enumeration. Update
                        // it directly. Never touch the target #(SEQ) field or its
                        // mathematical Field.Code.Text.
                        nestedField.Update();

                        // Word may rematerialize the outer field tree after a nested
                        // update. Reacquire it before restoring the user-visible
                        // character formatting inherited when the reference was
                        // inserted.
                        Release(nestedResult);
                        nestedResult = null;
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nestedField);
                        nestedField = null;
                        Release(nestedFields);
                        nestedFields = null;
                        Release(outerCode);
                        outerCode = outerField.Code;
                        NormalizeInternalStyle(outerCode);
                        nestedFields = outerCode.Fields;
                        for (var refreshedIndex = 1;
                             refreshedIndex <= nestedFields.Count;
                             refreshedIndex++)
                        {
                            Release(nestedCode);
                            nestedCode = null;
                            Release(nestedField);
                            nestedField = nestedFields[refreshedIndex];
                            if (nestedField.Type != WdFieldType.wdFieldRef)
                                continue;
                            nestedCode = nestedField.Code;
                            if (!TryReadVisualTeXNumberBookmark(
                                    nestedCode.Text,
                                    out var refreshedBookmark)
                                || !string.Equals(
                                    refreshedBookmark,
                                    bookmarkName,
                                    StringComparison.OrdinalIgnoreCase))
                                continue;
                            NormalizeInternalStyle(nestedCode);
                            formatting.Apply(nestedCode);
                            nestedResult = nestedField.Result;
                            NormalizeInternalStyle(nestedResult);
                            formatting.Apply(nestedResult);
                            break;
                        }
                        outerField.ShowCodes = false;
                        UpdateNavigationTarget(document, outerField, bookmarkName);
                        updated++;
                        break;
                    }
                }
                catch (COMException error)
                {
                    failures.Add(new InvalidOperationException(
                        $"Word could not refresh navigable formula reference {outerIndex}.", error));
                }
                finally
                {
                    Release(nestedResult);
                    Release(nestedCode);
                    Release(nestedField);
                    Release(nestedFields);
                    Release(outerCode);
                    Release(outerField);
                }
            }
        }
        finally { Release(outerFields); }
        if (failures.Count > 0)
            throw new AggregateException("Some formula references could not be refreshed.", failures);
        return updated;
    }

    private static string ResolveNavigationBookmark(Document document, string numberBookmarkName)
    {
        Bookmarks? bookmarks = null;
        Bookmark? numberBookmark = null;
        Range? numberRange = null;
        Frames? frames = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(numberBookmarkName))
                throw new InvalidDataException($"Reference number bookmark '{numberBookmarkName}' is missing.");
            numberBookmark = bookmarks[numberBookmarkName];
            numberRange = numberBookmark.Range;
            frames = numberRange.Frames;
            if (frames.Count == 0) return numberBookmarkName;

            // A clipped caption remains the REF value source. Navigation must use
            // the corresponding visible number. Compatibility aliases can retain
            // an older FormulaId, so prove ownership through the exact live number
            // range and an existing canonical visible bookmark, never proximity.
            string? navigationName = null;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Bookmark? candidate = null;
                Bookmark? visible = null;
                Range? candidateRange = null;
                Range? visibleRange = null;
                Range? visibleNumberText = null;
                Frames? visibleFrames = null;
                try
                {
                    candidate = bookmarks[index];
                    var name = candidate.Name;
                    const string prefix = "VTEqNum_";
                    if (!name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                        || !Guid.TryParseExact(name.Substring(prefix.Length), "N", out var formulaId))
                        continue;
                    candidateRange = candidate.Range;
                    if (candidateRange.StoryType != numberRange.StoryType
                        || candidateRange.Start != numberRange.Start || candidateRange.End != numberRange.End)
                        continue;
                    var visibleName = WordEquationNumbering.EquationBookmarkName(formulaId.ToString("D"));
                    if (!bookmarks.Exists(visibleName)) continue;
                    visible = bookmarks[visibleName];
                    visibleRange = visible.Range;
                    visibleNumberText = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                        document, formulaId.ToString("D"));
                    visibleFrames = visibleRange.Frames;
                    if (visibleFrames.Count != 0 || visibleRange.StoryType != WdStoryType.wdMainTextStory
                        || visibleNumberText is null || visibleNumberText.StoryType != visibleRange.StoryType
                        || visibleNumberText.Start < visibleRange.Start || visibleNumberText.End > visibleRange.End
                        || visibleNumberText.Text != "(" + numberRange.Text + ")")
                        continue;
                    if (navigationName is not null && navigationName != visibleName)
                        throw new InvalidDataException("The clipped number has multiple visible navigation owners.");
                    navigationName = visibleName;
                }
                finally
                {
                    Release(visibleFrames);
                    Release(visibleNumberText);
                    Release(visibleRange);
                    Release(candidateRange);
                    Release(visible);
                    Release(candidate);
                }
            }
            return navigationName
                ?? throw new InvalidDataException("The clipped number has no verified visible navigation owner.");
        }
        finally
        {
            Release(frames);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
        }
    }

    private static bool UpdateNavigationTarget(Document document, Field outerField, string numberBookmarkName)
    {
        Range? code = null;
        Range? token = null;
        Fields? nested = null;
        Field? firstNested = null;
        Range? firstCode = null;
        try
        {
            var target = ResolveNavigationBookmark(document, numberBookmarkName);
            code = outerField.Code;
            var match = Regex.Match(code.Text ?? string.Empty,
                @"^\s*GOTOBUTTON\s+(?<target>[^\s\\]+)", RegexOptions.IgnoreCase);
            if (!match.Success) throw new InvalidDataException("The navigable reference lost its GOTOBUTTON target.");
            var group = match.Groups["target"];
            if (string.Equals(group.Value, target, StringComparison.OrdinalIgnoreCase)) return false;
            nested = code.Fields;
            if (nested.Count != 1) throw new InvalidDataException("The navigable reference must contain exactly one REF.");
            firstNested = nested[1];
            firstCode = firstNested.Code;
            if (firstNested.Type != WdFieldType.wdFieldRef || code.Start + group.Index + group.Length >= firstCode.Start - 1)
                throw new InvalidDataException("The navigation token overlaps the nested reference.");
            token = code.Duplicate;
            token.SetRange(code.Start + group.Index, code.Start + group.Index + group.Length);
            // Replace only the leading bookmark token; assigning Code.Text would
            // destroy the nested REF and its user character formatting.
            token.Text = target;
            return true;
        }
        finally
        {
            Release(firstCode);
            Release(firstNested);
            Release(nested);
            Release(token);
            Release(code);
        }
    }

    internal static void RefreshFieldPreservingFormatting(Field field)
    {
        Range? code = null;
        Range? result = null;
        try
        {
            code = field.Code;
            result = field.Result;
            var formatting = WordCharacterFormatting.Capture(result);
            // CHARFORMAT must use the reference's own typeface; the target can
            // be a hidden number field with a different size or internal style.
            formatting.Apply(code);
            field.Update();
            Release(result);
            result = field.Result;
            formatting.Apply(result);
        }
        finally { Release(result); Release(code); }
    }

    private static bool IsGeneratedNumberReference(Document document, Field field, string name) =>
        name.StartsWith("VTEqNum_", StringComparison.OrdinalIgnoreCase)
        && WordEquationNumbering.IsGeneratedEquationNumberReference(document, field, name.Substring(8));

    internal static IReadOnlyDictionary<string, int> CaptureReferenceCounts(Document document)
    {
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        Fields? fields = null;
        Bookmarks? bookmarks = null;
        try
        {
            fields = document.Fields;
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!TryReadVisualTeXNumberBookmark(code.Text, out var name)
                        || !bookmarks.Exists(name) || IsGeneratedNumberReference(document, field, name))
                        continue;
                    counts.TryGetValue(name, out var count);
                    counts[name] = count + 1;
                }
                finally { Release(code); Release(field); }
            }
            return counts;
        }
        finally { Release(bookmarks); Release(fields); }
    }

    internal static void ValidateReferences(Document document, IReadOnlyDictionary<string, int> expectedCounts)
    {
        var actualCounts = expectedCounts.Keys.ToDictionary(name => name, _ => 0, StringComparer.OrdinalIgnoreCase);
        Bookmarks? bookmarks = null;
        Fields? fields = null;
        try
        {
            bookmarks = document.Bookmarks;
            foreach (var name in expectedCounts.Keys)
                if (!bookmarks.Exists(name))
                    throw new InvalidDataException($"Conversion lost reference target '{name}'.");
            fields = document.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!TryReadVisualTeXNumberBookmark(code.Text, out var name)
                        || !actualCounts.ContainsKey(name) || IsGeneratedNumberReference(document, field, name)) continue;
                    result = field.Result;
                    if (!HasCurrentReferenceResult(document, name, result))
                        throw new InvalidDataException($"Reference '{name}' does not display its current target number.");
                    actualCounts[name]++;
                }
                finally { Release(result); Release(code); Release(field); }
            }
            foreach (var expected in expectedCounts)
                if (actualCounts[expected.Key] != expected.Value)
                    throw new InvalidDataException($"Reference '{expected.Key}' retained {actualCounts[expected.Key]}/{expected.Value} fields.");
        }
        finally { Release(fields); Release(bookmarks); }
    }

    private static bool HasCurrentReferenceResult(Document document, string bookmarkName, Range result)
    {
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? target = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName)) return false;
            bookmark = bookmarks[bookmarkName];
            target = bookmark.Range;
            return string.Equals(result.Text, target.Text, StringComparison.Ordinal);
        }
        finally
        {
            Release(target);
            Release(bookmark);
            Release(bookmarks);
        }
    }

    internal static bool TryReadVisualTeXNumberBookmark(
        string? code,
        out string bookmarkName)
    {
        bookmarkName = string.Empty;
        if (string.IsNullOrWhiteSpace(code)) return false;
        var match = Regex.Match(
            code!,
            @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!match.Success) return false;
        var candidate = match.Groups["quoted"].Success
            ? match.Groups["quoted"].Value
            : match.Groups["plain"].Value;
        var visualTeX = candidate.StartsWith("VTEqNum_", StringComparison.OrdinalIgnoreCase)
            && candidate.Length == 40 && Guid.TryParseExact(candidate.Substring(8), "N", out _);
        if (!visualTeX && !candidate.StartsWith("ZEqnNum", StringComparison.OrdinalIgnoreCase))
            return false;
        bookmarkName = candidate;
        return true;
    }

    internal static int UpdateReferences(Document document, ISet<string>? targetBookmarkNames = null)
    {
        var updated = UpdateNavigableReferences(document, targetBookmarkNames);
        Fields? fields = null;
        var nestedStarts = new HashSet<int>();
        var failures = new List<Exception>();
        try
        {
            fields = document.Fields;
            // Word exposes nested REF both through GOTOBUTTON.Code.Fields and
            // Document.Fields. Refresh each once, preserving the host formatting.
            for (var index = 1; index <= fields.Count; index++)
            {
                Field? outer = null;
                Range? code = null;
                Fields? nested = null;
                try
                {
                    outer = fields[index];
                    if (outer.Type != WdFieldType.wdFieldGoToButton) continue;
                    code = outer.Code;
                    nested = code.Fields;
                    for (var child = 1; child <= nested.Count; child++)
                    {
                        Field? item = null;
                        Range? childCode = null;
                        try { item = nested[child]; childCode = item.Code; nestedStarts.Add(childCode.Start); }
                        finally { Release(childCode); Release(item); }
                    }
                }
                finally { Release(nested); Release(code); Release(outer); }
            }
            for (var index = fields.Count; index >= 1; index--)
            {
                Field? field = null;
                Range? code = null;
                Range? result = null;
                try
                {
                    field = fields[index];
                    if (field.Type != WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (nestedStarts.Contains(code.Start)
                        || !TryReadVisualTeXNumberBookmark(code.Text, out var name)
                        || (targetBookmarkNames is not null && !targetBookmarkNames.Contains(name))) continue;
                    result = field.Result;
                    if (HasCurrentReferenceResult(document, name, result)) continue;
                    RefreshFieldPreservingFormatting(field);
                    updated++;
                }
                catch (COMException error) { failures.Add(new InvalidOperationException($"Word could not refresh formula reference {index}.", error)); }
                finally { Release(result); Release(code); Release(field); }
            }
        }
        finally { Release(fields); }
        if (failures.Count > 0) throw new AggregateException("Some formula references could not be refreshed.", failures);
        return updated;
    }

    private static void NormalizeInternalStyle(Range range)
    {
        Style? style = null;
        try
        {
            style = range.get_Style() as Style;
            var styleName = style?.NameLocal ?? string.Empty;
            if (!string.Equals(
                    styleName,
                    MathTypeSectionStyleName,
                    StringComparison.OrdinalIgnoreCase))
                return;

            object defaultParagraphFont = WdBuiltinStyle.wdStyleDefaultParagraphFont;
            range.set_Style(ref defaultParagraphFont);
            Microsoft.Office.Interop.Word.Font? font = null;
            try { font = range.Font; font.Hidden = 0; }
            finally { Release(font); }
        }
        finally { Release(style); }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

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

    internal static void InsertNativeWordEquationReference(
        Document document,
        Selection selection,
        int nativeReferenceItem,
        string prefix,
        string suffix,
        WdColor? preferredInsertionColor = null)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (selection is null) throw new ArgumentNullException(nameof(selection));
        if (nativeReferenceItem <= 0)
            throw new ArgumentOutOfRangeException(
                nameof(nativeReferenceItem));
        if (document.ReadOnly)
            throw new UnauthorizedAccessException(
                "当前 Word 文档为只读状态。");

        Range? sourceFormattingRange = null;
        Range? insertion = null;
        Range? probe = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Range? selectedResult = null;
        Range? selectionRange = null;
        try
        {
            sourceFormattingRange =
                selection.Range.Duplicate;
            sourceFormattingRange.Collapse(
                WdCollapseDirection.wdCollapseStart);
            var formatting =
                WordCharacterFormatting.Capture(
                    sourceFormattingRange);
            if (preferredInsertionColor.HasValue)
            {
                var requested =
                    preferredInsertionColor.Value;
                formatting.Color =
                    requested == WdColor.wdColorAutomatic
                    || (int)requested >= 0
                        ? requested
                        : WdColor.wdColorAutomatic;
            }

            if (!string.IsNullOrEmpty(prefix))
                selection.TypeText(prefix);

            var referenceStart =
                selection.Start;
            insertion =
                selection.Range.Duplicate;
            insertion.Collapse(
                WdCollapseDirection.wdCollapseStart);
            insertion.InsertCrossReference(
                ReferenceType:
                    WdCaptionLabelID.wdCaptionEquation,
                ReferenceKind:
                    WdReferenceKind.wdOnlyLabelAndNumber,
                ReferenceItem:
                    nativeReferenceItem,
                InsertAsHyperlink: true,
                IncludePosition: false);

            var probeEnd =
                Math.Min(
                    document.Content.End,
                    referenceStart + 256);
            probe =
                document.Range(
                    referenceStart,
                    probeEnd);
            fields = probe.Fields;
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
                if (field.Type !=
                    WdFieldType.wdFieldRef)
                    continue;
                code =
                    field.Code.Duplicate;
                result =
                    field.Result.Duplicate;
                if (result.End < referenceStart)
                    continue;
                if (selectedResult is not null
                    && selectedResult.Start <=
                        result.Start)
                    continue;
                Release(selectedResult);
                selectedResult =
                    result.Duplicate;
            }

            if (selectedResult is null)
                throw new InvalidDataException(
                    "Word did not materialize a native REF field for the selected Equation item.");

            // Keep Word's native equation-reference result exactly as Word
            // materialized it. A native REF result can itself be OMath, and
            // forcing ordinary text properties such as Subscript/Superscript
            // onto that mathematical range is invalid in desktop Word.

            var after =
                Math.Max(
                    document.Content.Start,
                    Math.Min(
                        selectedResult.End + 1,
                        Math.Max(
                            document.Content.Start,
                            document.Content.End - 1)));
            selection.SetRange(
                after,
                after);
            if (!string.IsNullOrEmpty(suffix))
                selection.TypeText(suffix);

            selectionRange =
                selection.Range;
            NormalizeInternalStyle(
                selectionRange);
            formatting.Apply(
                selectionRange);
        }
        finally
        {
            Release(selectionRange);
            Release(selectedResult);
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(probe);
            Release(insertion);
            Release(sourceFormattingRange);
        }
    }

    internal static void InsertDirectReference(
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
            throw new ArgumentException(
                "Equation reference bookmark is required.",
                nameof(bookmarkName));
        if (document.ReadOnly)
            throw new UnauthorizedAccessException(
                "当前 Word 文档为只读状态。");

        Bookmarks? bookmarks = null;
        Range? sourceFormattingRange = null;
        Range? insertion = null;
        Field? refField = null;
        Range? refCode = null;
        Range? refResult = null;
        Range? selectionRange = null;
        try
        {
            bookmarks = document.Bookmarks;
            if (!bookmarks.Exists(bookmarkName))
                throw new InvalidDataException(
                    $"公式引用目标书签“{bookmarkName}”已不存在。文档内容可能已发生变化。");

            sourceFormattingRange =
                selection.Range.Duplicate;
            sourceFormattingRange.Collapse(
                WdCollapseDirection.wdCollapseStart);
            var formatting =
                WordCharacterFormatting.Capture(
                    sourceFormattingRange);
            if (preferredInsertionColor.HasValue)
            {
                var requested =
                    preferredInsertionColor.Value;
                formatting.Color =
                    requested == WdColor.wdColorAutomatic
                    || (int)requested >= 0
                        ? requested
                        : WdColor.wdColorAutomatic;
            }

            if (!string.IsNullOrEmpty(prefix))
                selection.TypeText(prefix);

            insertion =
                selection.Range.Duplicate;
            insertion.Collapse(
                WdCollapseDirection.wdCollapseStart);

            // One native REF field is sufficient. The \h switch gives Word's own
            // hyperlink/navigation behavior without the historical GOTOBUTTON
            // wrapper or a nested field tree. Avoid Field.Update here: Fields.Add
            // materializes the current REF result, while explicit Update can make
            // desktop Word close the surrounding Custom UndoRecord.
            refField = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldRef,
                bookmarkName + " \\h \\* CHARFORMAT",
                false);
            if (refField is null)
                throw new InvalidDataException(
                    "Word did not return the inserted REF field.");

            refCode = refField.Code;
            if (refField.Type != WdFieldType.wdFieldRef
                || !TryReadVisualTeXNumberBookmark(
                    refCode.Text,
                    out var insertedBookmark)
                || !string.Equals(
                    insertedBookmark,
                    bookmarkName,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "Word created a REF field for a different equation target.");

            refField.ShowCodes = false;
            refResult = refField.Result;
            NormalizeInternalStyle(
                refResult);
            formatting.Apply(
                refResult);
            if (!HasCurrentReferenceResult(
                    document,
                    bookmarkName,
                    refResult))
                throw new InvalidDataException(
                    "Word did not materialize the current equation number in the new REF field.");

            var after =
                Math.Max(
                    document.Content.Start,
                    Math.Min(
                        refResult.End + 1,
                        Math.Max(
                            document.Content.Start,
                            document.Content.End - 1)));
            selection.SetRange(
                after,
                after);
            if (!string.IsNullOrEmpty(suffix))
                selection.TypeText(suffix);

            selectionRange =
                selection.Range;
            NormalizeInternalStyle(
                selectionRange);
            formatting.Apply(
                selectionRange);
        }
        finally
        {
            Release(selectionRange);
            Release(refResult);
            Release(refCode);
            Release(refField);
            Release(insertion);
            Release(sourceFormattingRange);
            Release(bookmarks);
        }
    }

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
                Range? directResult = null;
                Fields? nestedFields = null;
                Field? nestedField = null;
                Range? nestedCode = null;
                Range? nestedResult = null;
                try
                {
                    outerField = outerFields[outerIndex];

                    if (outerField.Type == WdFieldType.wdFieldRef)
                    {
                        outerCode = outerField.Code;
                        if (!TryReadVisualTeXNumberBookmark(
                                outerCode.Text,
                                out var directBookmark)
                            || (targetBookmarkNames is not null
                                && !targetBookmarkNames.Contains(
                                    directBookmark)))
                            continue;

                        directResult =
                            outerField.Result;
                        if (!HasCurrentReferenceResult(
                                document,
                                directBookmark,
                                directResult))
                        {
                            var formatting =
                                WordCharacterFormatting.Capture(
                                    directResult);
                            outerField.Update();
                            Release(directResult);
                            directResult =
                                outerField.Result;
                            NormalizeInternalStyle(
                                directResult);
                            formatting.Apply(
                                directResult);
                            updated++;
                        }
                        continue;
                    }

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
                    Release(directResult);
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

    internal static string CreateNativeOmmlNumberReferenceBookmark(
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (host is null
            || host.Kind != WordFormulaHostKind.Omml)
            throw new ArgumentException(
                "A native Word reference bookmark requires one OMML host.",
                nameof(host));

        var numbering =
            WordFormulaNumberingResolver.ResolveLocal(
                document,
                host);
        if (!numbering.Numbered
            || numbering.ContainerKind !=
                WordFormulaNumberingContainerKind.CanonicalNativeOmml
            || numbering.NumberRange is null)
            throw new InvalidDataException(
                "The converted OMML host has no canonical native equation-number range.");

        Range? numberRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? created = null;
        Range? createdRange = null;
        try
        {
            numberRange =
                WordFormulaHostSemanticReader.CreateRange(
                    document,
                    numbering.NumberRange);
            bookmarks = document.Bookmarks;

            string? name = null;
            for (var attempt = 0;
                 attempt < 64;
                 attempt++)
            {
                var bytes =
                    Guid.NewGuid().ToByteArray();
                var value =
                    BitConverter.ToUInt32(
                        bytes,
                        0)
                    % 1_000_000_000u;
                var candidate =
                    "_Ref"
                    + value.ToString("D9");
                if (bookmarks.Exists(candidate))
                    continue;
                name = candidate;
                break;
            }
            if (name is null)
                throw new InvalidOperationException(
                    "VisualTeX could not allocate a unique native Word _Ref bookmark.");

            // A referenced pure OMML equation may own Word's ordinary hidden
            // _Ref bookmark, but never a VisualTeX FormulaId/private numbering
            // bookmark. Bind only the exact native number result (STYLEREF+SEQ
            // when applicable), so an existing number-only REF keeps displaying
            // exactly the same label after conversion.
            created =
                bookmarks.Add(
                    name,
                    numberRange);
            createdRange =
                created.Range.Duplicate;
            if (createdRange.StoryType !=
                    numberRange.StoryType
                || createdRange.Start !=
                    numberRange.Start
                || createdRange.End !=
                    numberRange.End)
                throw new InvalidDataException(
                    "Word changed the native equation-number range while creating its _Ref target.");

            return name;
        }
        finally
        {
            Release(createdRange);
            Release(created);
            Release(bookmarks);
            Release(numberRange);
        }
    }

    internal static int MigrateReferenceTargets(
        Document document,
        IReadOnlyDictionary<string, string> replacements,
        IReadOnlyDictionary<string, int> expectedReferenceCounts)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (replacements is null)
            throw new ArgumentNullException(nameof(replacements));
        if (expectedReferenceCounts is null)
            throw new ArgumentNullException(nameof(expectedReferenceCounts));
        if (replacements.Count == 0)
            return 0;

        var migrated =
            replacements.Keys.ToDictionary(
                name => name,
                _ => 0,
                StringComparer.OrdinalIgnoreCase);
        var total = 0;

        // First repair legacy navigable references. Their outer GOTOBUTTON target
        // and nested REF must move together to the same native _Ref bookmark.
        Fields? outerFields = null;
        try
        {
            outerFields = document.Fields;
            for (var index = 1;
                 index <= outerFields.Count;
                 index++)
            {
                Field? outer = null;
                Range? outerCode = null;
                Fields? nestedFields = null;
                Field? nested = null;
                Range? nestedCode = null;
                try
                {
                    outer = outerFields[index];
                    if (outer.Type !=
                        WdFieldType.wdFieldGoToButton)
                        continue;
                    outerCode = outer.Code;
                    nestedFields = outerCode.Fields;
                    for (var child = 1;
                         child <= nestedFields.Count;
                         child++)
                    {
                        Release(nestedCode);
                        nestedCode = null;
                        Release(nested);
                        nested = nestedFields[child];
                        if (nested.Type !=
                            WdFieldType.wdFieldRef)
                            continue;
                        nestedCode =
                            nested.Code.Duplicate;
                        if (!TryReadReferenceBookmark(
                                nestedCode.Text,
                                out var oldName)
                            || !replacements.TryGetValue(
                                oldName,
                                out var newName))
                            continue;

                        _ = UpdateNavigationTarget(
                            document,
                            outer,
                            newName);
                        ReplaceReferenceTargetToken(
                            nestedCode,
                            oldName,
                            newName);
                        RefreshMigratedReference(
                            document,
                            nested,
                            newName);
                        outer.ShowCodes = false;
                        migrated[oldName]++;
                        total++;
                        break;
                    }
                }
                finally
                {
                    Release(nestedCode);
                    Release(nested);
                    Release(nestedFields);
                    Release(outerCode);
                    Release(outer);
                }
            }
        }
        finally { Release(outerFields); }

        // Reacquire the document field inventory because updating a nested REF can
        // rematerialize its outer field tree. Direct REF fields are then migrated
        // in-place by replacing only their target token; switches, parentheses,
        // surrounding prose and character formatting remain user-owned.
        Fields? fields = null;
        try
        {
            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    if (field.Type !=
                        WdFieldType.wdFieldRef)
                        continue;
                    code = field.Code.Duplicate;
                    if (!TryReadReferenceBookmark(
                            code.Text,
                            out var oldName)
                        || !replacements.TryGetValue(
                            oldName,
                            out var newName))
                        continue;

                    ReplaceReferenceTargetToken(
                        code,
                        oldName,
                        newName);
                    RefreshMigratedReference(
                        document,
                        field,
                        newName);
                    field.ShowCodes = false;
                    migrated[oldName]++;
                    total++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
        }
        finally { Release(fields); }

        foreach (var replacement in
                 replacements)
        {
            expectedReferenceCounts.TryGetValue(
                replacement.Key,
                out var expected);
            if (migrated[replacement.Key] !=
                expected)
                throw new InvalidDataException(
                    $"Equation reference migration retained {migrated[replacement.Key]}/{expected} references for '{replacement.Key}'.");

            if (CountReferencesToBookmark(
                    document,
                    replacement.Key)
                != 0)
                throw new InvalidDataException(
                    $"Equation reference migration left a stale REF target '{replacement.Key}' in the document.");
        }

        return total;
    }

    private static void RefreshMigratedReference(
        Document document,
        Field field,
        string bookmarkName)
    {
        Range? result = null;
        OMaths? maths = null;
        try
        {
            result = field.Result.Duplicate;
            var formatting =
                WordCharacterFormatting.Capture(
                    result);
            field.Update();

            Release(result);
            result = field.Result.Duplicate;
            maths = result.OMaths;
            if (maths.Count == 0)
            {
                NormalizeInternalStyle(
                    result);
                formatting.Apply(
                    result);
            }

            if (!HasCurrentReferenceResult(
                    document,
                    bookmarkName,
                    result))
                throw new InvalidDataException(
                    $"Migrated equation REF '{bookmarkName}' does not display its current native number.");
        }
        finally
        {
            Release(maths);
            Release(result);
        }
    }

    private static int CountReferencesToBookmark(
        Document document,
        string bookmarkName)
    {
        Fields? fields = null;
        var count = 0;
        try
        {
            fields = document.Fields;
            for (var index = 1;
                 index <= fields.Count;
                 index++)
            {
                Field? field = null;
                Range? code = null;
                try
                {
                    field = fields[index];
                    if (field.Type !=
                        WdFieldType.wdFieldRef)
                        continue;
                    code = field.Code;
                    if (TryReadReferenceBookmark(
                            code.Text,
                            out var current)
                        && string.Equals(
                            current,
                            bookmarkName,
                            StringComparison.OrdinalIgnoreCase))
                        count++;
                }
                finally
                {
                    Release(code);
                    Release(field);
                }
            }
            return count;
        }
        finally { Release(fields); }
    }

    private static void ReplaceReferenceTargetToken(
        Range code,
        string expectedBookmark,
        string replacementBookmark)
    {
        var text =
            code.Text
            ?? string.Empty;
        var match =
            Regex.Match(
                text,
                @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
                RegexOptions.IgnoreCase
                | RegexOptions.CultureInvariant);
        if (!match.Success)
            throw new InvalidDataException(
                "The equation reference lost its REF target token.");

        var group =
            match.Groups["quoted"].Success
                ? match.Groups["quoted"]
                : match.Groups["plain"];
        if (!string.Equals(
                group.Value,
                expectedBookmark,
                StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException(
                $"The equation REF changed from '{expectedBookmark}' before migration.");

        Range? token = null;
        try
        {
            token = code.Duplicate;
            token.SetRange(
                code.Start + group.Index,
                code.Start + group.Index + group.Length);
            token.Text =
                replacementBookmark;
        }
        finally { Release(token); }
    }

    private static bool TryReadReferenceBookmark(
        string? code,
        out string bookmarkName)
    {
        bookmarkName = string.Empty;
        if (string.IsNullOrWhiteSpace(code))
            return false;
        var match =
            Regex.Match(
                code!,
                @"^\s*REF\s+(?:""(?<quoted>[^""]+)""|(?<plain>[^\s\\]+))",
                RegexOptions.IgnoreCase
                | RegexOptions.CultureInvariant);
        if (!match.Success)
            return false;
        bookmarkName =
            match.Groups["quoted"].Success
                ? match.Groups["quoted"].Value
                : match.Groups["plain"].Value;
        return !string.IsNullOrWhiteSpace(
            bookmarkName);
    }

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
                        // VTEq_*'s own visible-number REF is an internal numbering
                        // artifact, not a user equation reference. Numbering has
                        // already finalized and verified this field before the
                        // document-wide reference pass. Updating it again can make
                        // Word 2021 absorb the adjacent ')' into Field.Result.
                        // CaptureReferenceCounts/ValidateReferences already exclude
                        // the same generated field; refresh must use that contract.
                        || IsGeneratedNumberReference(document, field, name)
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

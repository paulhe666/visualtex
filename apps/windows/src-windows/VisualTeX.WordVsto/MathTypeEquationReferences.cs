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
        Range? insertion = null;
        Field? goToField = null;
        Range? goToCode = null;
        Range? nestedInsertion = null;
        Field? refField = null;
        Range? goToResult = null;
        Fields? finalNestedFields = null;
        Field? finalRefField = null;
        Range? finalRefCode = null;
        Range? finalRefResult = null;
        Microsoft.Office.Interop.Word.Font? insertionFont = null;
        Microsoft.Office.Interop.Word.Font? goToCodeFont = null;
        Microsoft.Office.Interop.Word.Font? refCodeFont = null;
        Microsoft.Office.Interop.Word.Font? refResultFont = null;
        Microsoft.Office.Interop.Word.Font? finalRefCodeFont = null;
        Microsoft.Office.Interop.Word.Font? finalRefResultFont = null;
        Microsoft.Office.Interop.Word.Font? selectionFont = null;
        Range? selectionRange = null;
        var insertionColor = WdColor.wdColorAutomatic;
        try
        {
            placeRef = ResolvePlaceRef(document, target)
                ?? throw new InvalidOperationException(
                    "找不到目标 MathType 公式编号。文档内容可能已在引用窗口打开后发生变化。");
            if (!TryGetVisibleNumberRange(document, placeRef, out numberRange)
                || numberRange is null)
                throw new InvalidDataException("MathType 公式编号缺少可引用的可见编号范围。");

            var bookmarkName = EnsureNativeMathTypeNumberBookmark(document, numberRange);

            insertion = selection.Range;
            insertion.Collapse(WdCollapseDirection.wdCollapseStart);
            insertionFont = insertion.Font;
            var requestedColor = preferredInsertionColor ?? insertionFont.Color;
            insertionColor = requestedColor == WdColor.wdColorAutomatic
                || (int)requestedColor >= 0
                ? requestedColor
                : WdColor.wdColorAutomatic;
            goToField = document.Fields.Add(
                insertion,
                WdFieldType.wdFieldGoToButton,
                bookmarkName + " ",
                true);

            // MathType inserts the REF field *inside the GOTOBUTTON field code*.
            // Word then renders the nested REF result as the visible reference,
            // while double-clicking the outer field navigates back to the number.
            goToCode = goToField.Code;
            // Word's built-in GOTOBUTTON field is created in red. MathType's own
            // equation-reference command immediately normalizes that temporary
            // field formatting before inserting its nested REF. Do the same here,
            // otherwise \\* Charformat makes the visible equation reference red
            // and leaves Word's typing color red after the field.
            NormalizeMathTypeInternalReferenceStyle(goToCode);
            goToCodeFont = goToCode.Font;
            goToCodeFont.Color = insertionColor;
            nestedInsertion = document.Range(goToCode.End, goToCode.End);
            refField = document.Fields.Add(
                nestedInsertion,
                WdFieldType.wdFieldRef,
                bookmarkName + " \\* Charformat \\!",
                true);
            var refCode = refField.Code;
            try
            {
                NormalizeMathTypeInternalReferenceStyle(refCode);
                refCodeFont = refCode.Font;
                refCodeFont.Color = insertionColor;
            }
            finally { Release(refCode); }
            try { refField.Update(); } catch { }
            var refResult = refField.Result;
            try
            {
                NormalizeMathTypeInternalReferenceStyle(refResult);
                refResultFont = refResult.Font;
                refResultFont.Color = insertionColor;
            }
            finally { Release(refResult); }
            try { refField.ShowCodes = false; } catch { }
            try { goToField.ShowCodes = false; } catch { }

            // Adding the nested REF causes Word to materialize the GOTOBUTTON
            // field tree a second time. On desktop Word that second materialization
            // reapplies GOTOBUTTON's built-in red character formatting, so any
            // color written to the pre-nesting Code range above is stale. Re-open
            // the *final* field tree after it is complete and normalize every
            // visible/code range once more. This mirrors MathType's own macro and
            // is required on real ribbon/dialog insertion, not just isolated tests.
            Release(goToCodeFont);
            goToCodeFont = null;
            Release(goToCode);
            goToCode = goToField.Code;
            NormalizeMathTypeInternalReferenceStyle(goToCode);
            goToCodeFont = goToCode.Font;
            goToCodeFont.Color = insertionColor;

            finalNestedFields = goToCode.Fields;
            if (finalNestedFields.Count != 1)
                throw new InvalidDataException(
                    "MathType GOTOBUTTON reference did not retain exactly one nested REF field.");
            finalRefField = finalNestedFields[1];
            finalRefCode = finalRefField.Code;
            NormalizeMathTypeInternalReferenceStyle(finalRefCode);
            finalRefCodeFont = finalRefCode.Font;
            finalRefCodeFont.Color = insertionColor;
            try { finalRefField.Update(); } catch { }

            // REF.Update() can recreate the result run and, on some Word builds,
            // repaint the enclosing GOTOBUTTON code red yet again. Normalize the
            // outer code after that final update, then the actual REF result.
            Release(goToCodeFont);
            goToCodeFont = null;
            Release(goToCode);
            goToCode = goToField.Code;
            NormalizeMathTypeInternalReferenceStyle(goToCode);
            goToCodeFont = goToCode.Font;
            goToCodeFont.Color = insertionColor;

            Release(finalRefResult);
            finalRefResult = null;
            Release(finalRefField);
            finalRefField = null;
            Release(finalNestedFields);
            finalNestedFields = null;
            finalNestedFields = goToCode.Fields;
            finalRefField = finalNestedFields[1];
            finalRefResult = finalRefField.Result;
            NormalizeMathTypeInternalReferenceStyle(finalRefResult);
            finalRefResultFont = finalRefResult.Font;
            finalRefResultFont.Color = insertionColor;

            Release(goToResult);
            goToResult = goToField.Result;
            // GOTOBUTTON's Result is collapsed immediately before its outer field
            // end. goToCode.End + 1 is still the red field boundary on desktop
            // Word, so a caret placed there inherits the field's typing format.
            // Advance past the outer field end and clamp to a legal document caret.
            var after = Math.Max(goToResult.End + 1, goToCode.End + 2);
            after = Math.Max(
                document.Content.Start,
                Math.Min(after, Math.Max(document.Content.Start, document.Content.End - 1)));
            selection.SetRange(after, after);
            selectionRange = selection.Range;
            NormalizeMathTypeInternalReferenceStyle(selectionRange);
            selectionFont = selection.Font;
            selectionFont.Color = insertionColor;
        }
        finally
        {
            Release(selectionRange);
            Release(selectionFont);
            Release(finalRefResultFont);
            Release(finalRefCodeFont);
            Release(refResultFont);
            Release(refCodeFont);
            Release(goToCodeFont);
            Release(insertionFont);
            Release(finalRefResult);
            Release(finalRefCode);
            Release(finalRefField);
            Release(finalNestedFields);
            Release(goToResult);
            Release(refField);
            Release(nestedInsertion);
            Release(goToCode);
            Release(goToField);
            Release(insertion);
            Release(numberRange);
            Release(placeRef);
        }
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
            fields = paragraphRange.Fields;
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

    internal static int RestoreReferenceBookmarkAliases(
        Document document,
        string formulaId,
        IReadOnlyCollection<string> bookmarkAliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (bookmarkAliases is null || bookmarkAliases.Count == 0) return 0;
        if (string.IsNullOrWhiteSpace(formulaId))
            throw new ArgumentException("FormulaId is required.", nameof(formulaId));

        Bookmarks? bookmarks = null;
        Bookmark? targetBookmark = null;
        Table? numberedTable = null;
        Cell? numberCell = null;
        Range? targetRange = null;
        Bookmark? aliasBookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            var targetName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            if (!bookmarks.Exists(targetName))
                throw new InvalidDataException(
                    $"Converted OMML formula {formulaId} has no durable number bookmark {targetName}.");
            targetBookmark = bookmarks[targetName];

            numberedTable = WordEquationNumbering.FindNumberedEquationTable(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Converted OMML formula {formulaId} has no numbered equation table.");
            numberCell = numberedTable.Cell(1, 3);
            targetRange = numberCell.Range.Duplicate;
            targetRange.End = Math.Max(targetRange.Start, targetRange.End - 1);
            if (string.IsNullOrWhiteSpace(targetRange.Text))
                throw new InvalidDataException(
                    $"Converted OMML formula {formulaId} has no visible number text for MathType reference compatibility.");

            var restored = 0;
            foreach (var alias in bookmarkAliases
                         .Where(name => !string.IsNullOrWhiteSpace(name))
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (!alias.StartsWith(EquationBookmarkPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (bookmarks.Exists(alias))
                {
                    Release(aliasBookmark);
                    aliasBookmark = bookmarks[alias];
                    aliasBookmark.Delete();
                }
                Release(aliasBookmark);
                aliasBookmark = bookmarks.Add(alias, targetRange);
                restored++;
            }
            return restored;
        }
        finally
        {
            Release(aliasBookmark);
            Release(targetRange);
            Release(numberCell);
            Release(numberedTable);
            Release(targetBookmark);
            Release(bookmarks);
        }
    }

    internal static int RefreshReferences(
        Document document,
        ISet<string> bookmarkAliases)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        if (bookmarkAliases is null || bookmarkAliases.Count == 0) return 0;

        Fields? fields = null;
        Field? outer = null;
        Range? outerCode = null;
        Fields? nestedFields = null;
        Field? nested = null;
        Range? nestedCode = null;
        Range? nestedResult = null;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
        Microsoft.Office.Interop.Word.Font? codeFont = null;
        Microsoft.Office.Interop.Word.Font? outerCodeFont = null;
        var updated = 0;
        try
        {
            fields = document.Fields;
            for (var outerIndex = 1; outerIndex <= fields.Count; outerIndex++)
            {
                Release(outerCodeFont);
                outerCodeFont = null;
                Release(nestedResult);
                nestedResult = null;
                Release(nestedCode);
                nestedCode = null;
                Release(nested);
                nested = null;
                Release(nestedFields);
                nestedFields = null;
                Release(outerCode);
                outerCode = null;
                Release(outer);
                outer = fields[outerIndex];
                outerCode = outer.Code;
                nestedFields = outerCode.Fields;
                if (nestedFields.Count == 0) continue;

                for (var nestedIndex = 1; nestedIndex <= nestedFields.Count; nestedIndex++)
                {
                    Release(codeFont);
                    codeFont = null;
                    Release(resultFont);
                    resultFont = null;
                    Release(nestedResult);
                    nestedResult = null;
                    Release(nestedCode);
                    nestedCode = null;
                    Release(nested);
                    nested = nestedFields[nestedIndex];
                    nestedCode = nested.Code;
                    var nestedText = nestedCode.Text ?? string.Empty;
                    if (!bookmarkAliases.Any(alias =>
                            nestedText.IndexOf(alias, StringComparison.OrdinalIgnoreCase) >= 0))
                        continue;

                    var preferredColor = WdColor.wdColorAutomatic;
                    try
                    {
                        nestedResult = nested.Result;
                        resultFont = nestedResult.Font;
                        var current = resultFont.Color;
                        if (current == WdColor.wdColorAutomatic || (int)current >= 0)
                            preferredColor = current;
                    }
                    catch { }

                    NormalizeMathTypeInternalReferenceStyle(nestedCode);
                    try
                    {
                        codeFont = nestedCode.Font;
                        codeFont.Color = preferredColor;
                    }
                    catch { }
                    try { nested.Update(); } catch { }

                    Release(nestedResult);
                    nestedResult = null;
                    Release(resultFont);
                    resultFont = null;
                    try
                    {
                        nestedResult = nested.Result;
                        NormalizeMathTypeInternalReferenceStyle(nestedResult);
                        resultFont = nestedResult.Font;
                        resultFont.Color = preferredColor;
                    }
                    catch { }

                    // Updating a nested REF can rematerialize the enclosing
                    // GOTOBUTTON code and reapply Word's built-in red style. Reopen
                    // that final outer code and normalize it after every update.
                    Release(outerCodeFont);
                    outerCodeFont = null;
                    Release(outerCode);
                    outerCode = outer.Code;
                    NormalizeMathTypeInternalReferenceStyle(outerCode);
                    try
                    {
                        outerCodeFont = outerCode.Font;
                        outerCodeFont.Color = preferredColor;
                    }
                    catch { }
                    updated++;
                }
            }
            return updated;
        }
        finally
        {
            Release(outerCodeFont);
            Release(codeFont);
            Release(resultFont);
            Release(nestedResult);
            Release(nestedCode);
            Release(nested);
            Release(nestedFields);
            Release(outerCode);
            Release(outer);
            Release(fields);
        }
    }

    private static void NormalizeMathTypeInternalReferenceStyle(Range range)
    {
        Style? style = null;
        try
        {
            try { style = range.get_Style() as Style; }
            catch { }
            var styleName = string.Empty;
            try { styleName = style?.NameLocal ?? string.Empty; }
            catch { }
            if (!string.Equals(
                    styleName,
                    MathTypeSectionStyleName,
                    StringComparison.OrdinalIgnoreCase))
                return;

            // MTEquationSection is MathType's internal hidden/red character style.
            // It must never escape onto a visible equation reference. Reset only
            // this internal style; legitimate user character styles are preserved.
            object defaultParagraphFont = WdBuiltinStyle.wdStyleDefaultParagraphFont;
            try { range.set_Style(ref defaultParagraphFont); }
            catch { }
            try { range.Font.Hidden = 0; } catch { }
        }
        finally { Release(style); }
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
        try { Marshal.FinalReleaseComObject(value); }
        catch { }
    }
}

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

/// <summary>
/// Updates MathType's native Word-field numbering without introducing a second
/// VisualTeX sequence. Numbered MathType equations remain MACROBUTTON MTPlaceRef
/// fields backed by MTChap/MTSec/MTEqn, and native ZEqnNum references keep their
/// bookmark names while the visible number template is rewritten.
/// </summary>
internal static class MathTypeEquationNumbering
{
    private const string SectionBreakMarker = "MACROBUTTON MTEditEquationSection2";
    private const string ReferenceBookmarkPrefix = "ZEqnNum";

    internal static int UpdateEquationNumbers(Document document)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));

        var numberedEquationCount = CountPlaceRefFields(document);
        if (numberedEquationCount == 0) return 0;

        // Update only MathType-owned sequence fields. A document-wide
        // Fields.Update() would also refresh TOCs, citations and unrelated fields.
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
                if (!IsMathTypeSequenceFieldCode(code.Text)) continue;
                try { field.Update(); } catch { }
            }

            // MathType equation references use a nested REF ZEqnNum... field
            // inside GOTOBUTTON. The REF fields are also exposed by
            // Document.Fields, so update those explicitly after all sequences.
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(code);
                code = null;
                Release(field);
                field = fields[index];
                code = field.Code;
                if (!IsMathTypeReferenceFieldCode(code.Text)) continue;
                try { field.Update(); } catch { }
            }
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }

        return numberedEquationCount;
    }

    internal static int SetEquationNumberFormat(Document document, string? formatId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var format = EquationNumberFormat.Resolve(formatId);
        var starts = CollectPlaceRefCodeStarts(document);
        if (starts.Count == 0) return 0;

        // VisualTeX heading-aware numbering follows real Word Heading paragraphs.
        // Rebuild MathType's hidden MTChap/MTSec state from those headings whenever
        // the user explicitly changes the numbering preset. A headingless document
        // therefore remains in chapter/section zero instead of manufacturing 1/1.
        WordFormulaService.RemoveAllMathTypeSectionBreakFields(document);
        if (format.UsesHeading)
        {
            var guard = 0;
            while (guard++ < Math.Max(1, starts.Count))
            {
                var inserted = false;
                var currentStarts = CollectPlaceRefCodeStarts(document);
                foreach (var start in currentStarts)
                {
                    if (WordFormulaService.EnsureMathTypeHeadingScopeState(
                            document,
                            Math.Max(document.Content.Start, start - 1),
                            format) <= 0)
                        continue;
                    inserted = true;
                    break;
                }
                if (!inserted) break;
            }
            starts = CollectPlaceRefCodeStarts(document);
        }
        else
        {
            starts = CollectPlaceRefCodeStarts(document);
        }

        var template = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(format.Id);
        var rewritten = 0;
        // Process from the end of the document toward the start. Replacing the
        // field-code tail changes its character length; descending order keeps all
        // yet-to-be-processed stored positions stable.
        foreach (var start in starts.OrderByDescending(value => value))
        {
            Field? placeRef = null;
            try
            {
                placeRef = ResolvePlaceRefAtCodeStart(document, start);
                if (placeRef is null) continue;
                RewriteVisibleNumberTemplate(document, placeRef, template);
                rewritten++;
            }
            finally { Release(placeRef); }
        }

        UpdateEquationNumbers(document);
        return rewritten;
    }

    internal static int CountPlaceRefFields(Document document)
    {
        if (document is null) return 0;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var count = 0;
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
                if (MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    count++;
            }
            return count;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static List<int> CollectPlaceRefCodeStarts(Document document)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        var starts = new List<int>();
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
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                starts.Add(code.Start);
            }
            return starts;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static Field? ResolvePlaceRefAtCodeStart(Document document, int codeStart)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Field? result = null;
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
                if (code.Start != codeStart
                    || !MathTypeEquationReferences.IsMathTypePlaceRefCode(code.Text))
                    continue;
                result = field;
                field = null;
                break;
            }
            return result;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static void RewriteVisibleNumberTemplate(
        Document document,
        Field placeRef,
        MathTypeWordOpenXml.NumberTemplate template)
    {
        Range? numberRange = null;
        Range? outerCode = null;
        Range? rebuiltCode = null;
        Fields? rebuiltNestedFields = null;
        Field? rebuilt = null;
        Field? createdField = null;
        Range? createdCode = null;
        Microsoft.Office.Interop.Word.Font? outerFont = null;
        Microsoft.Office.Interop.Word.Font? rebuiltFont = null;
        Microsoft.Office.Interop.Word.Font? createdFont = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        Range? bookmarkRange = null;
        var bookmarkNames = new List<string>();
        var codeColor = WdColor.wdColorAutomatic;
        var fieldStart = -1;
        try
        {
            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    out numberRange)
                || numberRange is null)
                throw new InvalidDataException(
                    "MathType MTPlaceRef has no visible equation-number range.");

            outerCode = placeRef.Code;
            fieldStart = Math.Max(document.Content.Start, outerCode.Start - 1);
            outerFont = outerCode.Font;
            if (outerFont.Color == WdColor.wdColorAutomatic || (int)outerFont.Color >= 0)
                codeColor = outerFont.Color;

            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[index];
                if (!bookmark.Name.StartsWith(
                        ReferenceBookmarkPrefix,
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start < numberRange.Start
                    || bookmarkRange.End > numberRange.End)
                    continue;
                bookmarkNames.Add(bookmark.Name);
            }

            // Word does not reliably delete a nested field tree when only the
            // visible tail of a MACROBUTTON field is deleted. Rebuild this one
            // MTPlaceRef field atomically instead. The surrounding tabs, equation
            // OLE and paragraph layout are outside the field and remain untouched.
            placeRef.Delete();
            Release(numberRange);
            numberRange = null;
            Release(outerFont);
            outerFont = null;
            Release(outerCode);
            outerCode = null;

            // Use the same structurally verified MTPlaceRef constructor as direct
            // MathType insertion. The old formatter rebuilt nested SEQ fields at
            // Field.Code.End while the outer field was in result view, which can
            // detach MTEqn/MTChap into sibling document fields just like the former
            // left-number insertion bug.
            rebuilt = WordFormulaService.CreateIndependentMathTypePlaceRef(
                document,
                fieldStart,
                template);

            rebuiltCode = rebuilt.Code;
            rebuiltFont = rebuiltCode.Font;
            rebuiltFont.Color = codeColor;
            rebuiltNestedFields = rebuiltCode.Fields;
            for (var index = 1; index <= rebuiltNestedFields.Count; index++)
            {
                Release(createdField);
                createdField = rebuiltNestedFields[index];
                Release(createdCode);
                createdCode = createdField.Code;
                Release(createdFont);
                createdFont = createdCode.Font;
                createdFont.Color = codeColor;
            }
            try { rebuilt.ShowCodes = false; } catch { }

            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    rebuilt,
                    out numberRange)
                || numberRange is null)
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef has no visible equation-number range.");

            // Preserve MathType's native ZEqnNum bookmark identity. Existing
            // GOTOBUTTON/REF fields keep pointing to the same bookmark names and
            // only need their results refreshed after the rewrite.
            foreach (var name in bookmarkNames)
            {
                if (bookmarks.Exists(name))
                {
                    Release(bookmark);
                    bookmark = bookmarks[name];
                    try { bookmark.Delete(); } catch { }
                }
                Release(bookmark);
                bookmark = bookmarks.Add(name, numberRange);
            }
        }
        finally
        {
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(createdFont);
            Release(createdCode);
            Release(createdField);
            Release(rebuiltNestedFields);
            Release(rebuiltFont);
            Release(rebuiltCode);
            Release(rebuilt);
            Release(outerFont);
            Release(outerCode);
            Release(numberRange);
        }
    }

    private static bool IsMathTypeSequenceFieldCode(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var normalized = code!
            .Replace('\t', ' ')
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .TrimStart();
        return normalized.StartsWith("SEQ MTEqn ", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("SEQ MTSec ", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("SEQ MTChap ", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsMathTypeReferenceFieldCode(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) return false;
        var normalized = code!
            .Replace('\t', ' ')
            .Replace('\r', ' ')
            .Replace('\n', ' ')
            .TrimStart();
        return normalized.StartsWith("REF " + ReferenceBookmarkPrefix,
            StringComparison.OrdinalIgnoreCase);
    }

    internal static bool IsMathTypeSectionBreakCode(string? code) =>
        !string.IsNullOrWhiteSpace(code)
        && code!.IndexOf(SectionBreakMarker, StringComparison.OrdinalIgnoreCase) >= 0;

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); }
        catch { }
    }
}

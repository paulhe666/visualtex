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

    private sealed class PlaceRefRewritePlan
    {
        internal int FieldStart { get; set; }
        internal int ParagraphStart { get; set; }
        internal bool NumberOnLeft { get; set; }
        internal WdColor CodeColor { get; set; } = WdColor.wdColorAutomatic;
        internal string OleProgId { get; set; } = string.Empty;
        internal string SourceFlatOpc { get; set; } = string.Empty;
        internal string RewrittenFlatOpc { get; set; } = string.Empty;
        internal string[] BookmarkNames { get; set; } = Array.Empty<string>();
    }

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

    internal static int ValidateEquationNumberFormat(
        Document document,
        string? formatId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        return CreatePlaceRefRewritePlans(document, formatId).Count;
    }

    internal static int SetEquationNumberFormat(Document document, string? formatId)
    {
        if (document is null) throw new ArgumentNullException(nameof(document));
        var plans = CreatePlaceRefRewritePlans(document, formatId);
        if (plans.Count == 0) return 0;

        // MathType was used only to observe the native Word/OpenXML contract.
        // Production never calls a MathType macro or process: VisualTeX prepares
        // every replacement package itself and writes one exact MTPlaceRef field
        // range at a time. MTEditEquationSection2, tabs, OLE objects and paragraph
        // marks remain outside those transactions.
        var paragraphCountBefore = ReadDocumentParagraphCount(document);
        var inlineShapeCountBefore = ReadDocumentInlineShapeCount(document);
        var applied = new List<PlaceRefRewritePlan>(plans.Count);
        Document? stagingDocument = null;
        try
        {
            // Do not FinalRelease document.Application/Documents here. Office may
            // return the same RCW held by the caller; disconnecting it would make
            // the active Word session unusable after an otherwise successful
            // rewrite. The short-lived staging Document is the only owned RCW.
            stagingDocument = document.Application.Documents.Add(Visible: false);
            try
            {
                foreach (var plan in plans)
                {
                    applied.Add(plan);
                    ApplyPlaceRefRewritePlan(
                        document,
                        stagingDocument,
                        plan);
                }

                UpdateEquationNumbers(document);
                if (ReadDocumentParagraphCount(document) != paragraphCountBefore)
                    throw new InvalidDataException(
                        "MathType number-format rewrite changed the document paragraph count.");
                if (ReadDocumentInlineShapeCount(document) != inlineShapeCountBefore)
                    throw new InvalidDataException(
                        "MathType number-format rewrite changed the Equation.DSMT4 object count.");
                return plans.Count;
            }
            catch (Exception error)
            {
                var rollbackErrors = new List<Exception>();
                foreach (var plan in applied.AsEnumerable().Reverse())
                {
                    try
                    {
                        RestorePlaceRefRewritePlan(
                            document,
                            stagingDocument,
                            plan);
                    }
                    catch (Exception rollbackError)
                    {
                        rollbackErrors.Add(rollbackError);
                    }
                }
                try { UpdateEquationNumbers(document); } catch { }
                if (rollbackErrors.Count > 0)
                    throw new AggregateException(
                        "MathType number-format rewrite failed and one or more native MTPlaceRef fields could not be restored.",
                        new[] { error }.Concat(rollbackErrors));
                throw;
            }
        }
        finally
        {
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(stagingDocument);
        }
    }

    private static List<PlaceRefRewritePlan> CreatePlaceRefRewritePlans(
        Document document,
        string? formatId)
    {
        var format = EquationNumberFormat.Resolve(formatId);
        var starts = CollectPlaceRefCodeStarts(document)
            .OrderByDescending(value => value)
            .ToArray();
        if (starts.Length == 0) return new List<PlaceRefRewritePlan>();

        var template = MathTypeWordOpenXml.CreateVisualTeXNumberTemplate(format.Id);
        var plans = new List<PlaceRefRewritePlan>(starts.Length);
        foreach (var start in starts)
        {
            Field? placeRef = null;
            try
            {
                placeRef = ResolvePlaceRefAtCodeStart(document, start)
                    ?? throw new InvalidDataException(
                        $"MathType MTPlaceRef at code position {start} disappeared before format preparation.");
                plans.Add(CreatePlaceRefRewritePlan(
                    document,
                    placeRef,
                    template));
            }
            finally { Release(placeRef); }
        }
        return plans;
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

    private static PlaceRefRewritePlan CreatePlaceRefRewritePlan(
        Document document,
        Field placeRef,
        MathTypeWordOpenXml.NumberTemplate template)
    {
        Range? code = null;
        Range? result = null;
        Range? fieldRange = null;
        Range? paragraphRange = null;
        Range? shapeRange = null;
        Range? separator = null;
        Range? bookmarkRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OLEFormat? oleFormat = null;
        Microsoft.Office.Interop.Word.Font? codeFont = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        var bookmarkNames = new List<string>();
        var previousShowHidden = false;
        try
        {
            code = placeRef.Code;
            result = placeRef.Result;
            var fieldStart = code.Start - 1;
            var fieldEnd = ResolvePlaceRefFieldEndExclusive(
                document,
                code,
                result);
            fieldRange = document.Range(fieldStart, fieldEnd);
            paragraphs = fieldRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "MathType MTPlaceRef rewrite requires one owner paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            shapes = paragraphRange.InlineShapes;
            if (shapes.Count != 1)
                throw new InvalidDataException(
                    "MathType MTPlaceRef owner paragraph must contain exactly one Equation.DSMT4 object.");
            shape = shapes[1];
            shapeRange = shape.Range;

            bool numberOnLeft;
            if (fieldEnd <= shapeRange.Start)
            {
                numberOnLeft = true;
                separator = document.Range(fieldEnd, shapeRange.Start);
            }
            else if (fieldStart >= shapeRange.End)
            {
                numberOnLeft = false;
                separator = document.Range(shapeRange.End, fieldStart);
            }
            else
            {
                throw new InvalidDataException(
                    "MathType MTPlaceRef overlaps its Equation.DSMT4 object.");
            }
            if (!string.Equals(separator.Text, "\t", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "MathType MTPlaceRef is not separated from its Equation.DSMT4 object by exactly one native tab.");

            oleFormat = shape.OLEFormat;
            var progId = oleFormat.ProgID ?? string.Empty;
            if (!progId.StartsWith("Equation.DSMT4", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    $"MathType MTPlaceRef owner is not Equation.DSMT4: '{progId}'.");

            codeFont = code.Font;
            var codeColor = codeFont.Color;
            if (codeColor != WdColor.wdColorAutomatic && (int)codeColor < 0)
                codeColor = WdColor.wdColorAutomatic;

            bookmarks = document.Bookmarks;
            try
            {
                previousShowHidden = bookmarks.ShowHidden;
                bookmarks.ShowHidden = true;
            }
            catch { }
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
                if (bookmarkRange.Start < fieldStart
                    || bookmarkRange.End > fieldEnd)
                    continue;
                bookmarkNames.Add(bookmark.Name);
            }

            var sourceFlatOpc = fieldRange.WordOpenXML;
            var rewrittenFlatOpc =
                MathTypeWordOpenXml.RewriteMathTypePlaceRefFieldFlatOpc(
                    sourceFlatOpc,
                    template);
            return new PlaceRefRewritePlan
            {
                FieldStart = fieldStart,
                ParagraphStart = paragraphRange.Start,
                NumberOnLeft = numberOnLeft,
                CodeColor = codeColor,
                OleProgId = progId,
                SourceFlatOpc = sourceFlatOpc,
                RewrittenFlatOpc = rewrittenFlatOpc,
                BookmarkNames = bookmarkNames.ToArray(),
            };
        }
        finally
        {
            if (bookmarks is not null)
            {
                try { bookmarks.ShowHidden = previousShowHidden; } catch { }
            }
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(codeFont);
            Release(oleFormat);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(separator);
            Release(fieldRange);
            Release(result);
            Release(code);
        }
    }

    private static void ApplyPlaceRefRewritePlan(
        Document document,
        Document stagingDocument,
        PlaceRefRewritePlan plan)
    {
        Field? current = null;
        Field? rebuilt = null;
        Range? code = null;
        Range? result = null;
        Range? fieldRange = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        try
        {
            current = ResolvePlaceRefAtCodeStart(
                    document,
                    plan.FieldStart + 1)
                ?? throw new InvalidDataException(
                    $"MathType MTPlaceRef at {plan.FieldStart} disappeared before its OpenXML transaction.");
            code = current.Code;
            result = current.Result;
            var fieldEnd = ResolvePlaceRefFieldEndExclusive(
                document,
                code,
                result);
            fieldRange = document.Range(plan.FieldStart, fieldEnd);
            ReplacePlaceRefRangeFromFlatOpc(
                stagingDocument,
                fieldRange,
                plan.RewrittenFlatOpc);

            Release(fieldRange);
            fieldRange = null;
            Release(result);
            result = null;
            Release(code);
            code = null;
            Release(current);
            current = null;

            rebuilt = ResolvePlaceRefAtCodeStart(
                    document,
                    plan.FieldStart + 1)
                ?? throw new InvalidDataException(
                    $"Word did not materialize the rebuilt MathType MTPlaceRef at {plan.FieldStart}.");
            try { rebuilt.ShowCodes = false; } catch { }
            code = rebuilt.Code;
            font = code.Font;
            font.Color = plan.CodeColor;
            RestoreNativeBookmarkRanges(
                document,
                rebuilt,
                plan.BookmarkNames);
            ValidatePlaceRefLayout(
                document,
                rebuilt,
                plan,
                requireNativeFieldEnd: true);
        }
        finally
        {
            Release(font);
            Release(fieldRange);
            Release(result);
            Release(code);
            Release(rebuilt);
            Release(current);
        }
    }

    private static void RestorePlaceRefRewritePlan(
        Document document,
        Document stagingDocument,
        PlaceRefRewritePlan plan)
    {
        Field? current = null;
        Field? restored = null;
        Range? code = null;
        Range? result = null;
        Range? fieldRange = null;
        try
        {
            current = ResolvePlaceRefAtCodeStart(
                    document,
                    plan.FieldStart + 1)
                ?? throw new InvalidDataException(
                    $"MathType MTPlaceRef at {plan.FieldStart} is unavailable for rollback.");
            code = current.Code;
            result = current.Result;
            var fieldEnd = ResolvePlaceRefFieldEndExclusive(
                document,
                code,
                result);
            fieldRange = document.Range(plan.FieldStart, fieldEnd);
            ReplacePlaceRefRangeFromFlatOpc(
                stagingDocument,
                fieldRange,
                plan.SourceFlatOpc);

            Release(fieldRange);
            fieldRange = null;
            Release(result);
            result = null;
            Release(code);
            code = null;
            Release(current);
            current = null;

            restored = ResolvePlaceRefAtCodeStart(
                    document,
                    plan.FieldStart + 1)
                ?? throw new InvalidDataException(
                    $"Word did not restore the original MathType MTPlaceRef at {plan.FieldStart}.");
            try { restored.ShowCodes = false; } catch { }
            RestoreNativeBookmarkRanges(
                document,
                restored,
                plan.BookmarkNames);
            ValidatePlaceRefLayout(
                document,
                restored,
                plan,
                requireNativeFieldEnd: false);
        }
        finally
        {
            Release(fieldRange);
            Release(result);
            Release(code);
            Release(restored);
            Release(current);
        }
    }

    private static void ReplacePlaceRefRangeFromFlatOpc(
        Document stagingDocument,
        Range targetRange,
        string fieldFlatOpc)
    {
        Range? stagingContent = null;
        Range? insertion = null;
        Range? candidateCode = null;
        Range? stagedCode = null;
        Range? stagedResult = null;
        Range? stagedFieldRange = null;
        Fields? fields = null;
        Field? candidate = null;
        Field? stagedField = null;
        try
        {
            stagingContent = stagingDocument.Content;
            stagingContent.Text = string.Empty;
            Release(stagingContent);
            stagingContent = stagingDocument.Content;
            insertion = stagingDocument.Range(
                stagingContent.Start,
                stagingContent.Start);
            insertion.InsertXML(fieldFlatOpc);

            fields = stagingDocument.Fields;
            var placeRefCount = 0;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(candidateCode);
                candidateCode = null;
                Release(candidate);
                candidate = fields[index];
                candidateCode = candidate.Code;
                if (!MathTypeEquationReferences.IsMathTypePlaceRefCode(
                        candidateCode.Text))
                    continue;
                placeRefCount++;
                if (stagedField is null)
                {
                    stagedField = candidate;
                    candidate = null;
                }
            }
            if (placeRefCount != 1 || stagedField is null)
                throw new InvalidDataException(
                    $"MathType staging document materialized {placeRefCount} MTPlaceRef fields instead of one.");

            stagedCode = stagedField.Code;
            stagedResult = stagedField.Result;
            var stagedStart = stagedCode.Start - 1;
            var stagedEnd = ResolvePlaceRefFieldEndExclusive(
                stagingDocument,
                stagedCode,
                stagedResult);
            stagedFieldRange = stagingDocument.Range(
                stagedStart,
                stagedEnd);

            // Assigning only the exact staged field range avoids the paragraph
            // split caused by inserting a Flat OPC <w:p> wrapper directly into a
            // live MTDisplayEquation paragraph. This is pure Word automation and
            // does not load or call MathType.
            targetRange.FormattedText = stagedFieldRange;
        }
        finally
        {
            Release(stagedFieldRange);
            Release(stagedResult);
            Release(stagedCode);
            Release(stagedField);
            Release(candidateCode);
            Release(candidate);
            Release(fields);
            Release(insertion);
            Release(stagingContent);
        }
    }

    private static void RestoreNativeBookmarkRanges(
        Document document,
        Field placeRef,
        IReadOnlyList<string> bookmarkNames)
    {
        if (bookmarkNames.Count == 0) return;
        Range? visibleRange = null;
        Range? bookmarkRange = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        try
        {
            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    out visibleRange)
                || visibleRange is null)
                throw new InvalidDataException(
                    "MathType MTPlaceRef has no visible range for its ZEqnNum bookmark.");

            bookmarks = document.Bookmarks;
            foreach (var name in bookmarkNames)
            {
                if (bookmarks.Exists(name))
                {
                    Release(bookmark);
                    bookmark = bookmarks[name];
                    try { bookmark.Delete(); } catch { }
                }
                Release(bookmarkRange);
                bookmarkRange = visibleRange.Duplicate;
                Release(bookmark);
                bookmark = bookmarks.Add(name, bookmarkRange);
            }
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
            Release(bookmarkRange);
            Release(visibleRange);
        }
    }

    private static void ValidatePlaceRefLayout(
        Document document,
        Field placeRef,
        PlaceRefRewritePlan plan,
        bool requireNativeFieldEnd)
    {
        Range? code = null;
        Range? result = null;
        Range? fieldRange = null;
        Range? paragraphRange = null;
        Range? shapeRange = null;
        Range? separator = null;
        Range? bookmarkRange = null;
        Range? visibleNumberRange = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        OLEFormat? oleFormat = null;
        Bookmarks? bookmarks = null;
        Bookmark? bookmark = null;
        var previousShowHidden = false;
        try
        {
            code = placeRef.Code;
            result = placeRef.Result;
            var fieldStart = code.Start - 1;
            if (fieldStart != plan.FieldStart)
                throw new InvalidDataException(
                    $"MathType MTPlaceRef moved from {plan.FieldStart} to {fieldStart}.");
            var fieldEnd = ResolvePlaceRefFieldEndExclusive(
                document,
                code,
                result);
            if (requireNativeFieldEnd
                && ReadDocumentCharacter(document, code.End) != '\u0015')
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef retained an outer separator/result instead of MathType's native field end.");

            fieldRange = document.Range(fieldStart, fieldEnd);
            paragraphs = fieldRange.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef spans more than one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            if (paragraphRange.Start != plan.ParagraphStart)
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef moved to another paragraph.");
            shapes = paragraphRange.InlineShapes;
            if (shapes.Count != 1)
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef owner paragraph changed its OLE object count.");
            shape = shapes[1];
            shapeRange = shape.Range;
            oleFormat = shape.OLEFormat;
            if (!string.Equals(
                    oleFormat.ProgID,
                    plan.OleProgId,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef changed the Equation.DSMT4 ProgID.");

            if (plan.NumberOnLeft)
            {
                if (fieldEnd > shapeRange.Start)
                    throw new InvalidDataException(
                        "Rebuilt left MathType number moved after or into its equation object.");
                separator = document.Range(fieldEnd, shapeRange.Start);
            }
            else
            {
                if (fieldStart < shapeRange.End)
                    throw new InvalidDataException(
                        "Rebuilt right MathType number moved before or into its equation object.");
                separator = document.Range(shapeRange.End, fieldStart);
            }
            if (!string.Equals(separator.Text, "\t", StringComparison.Ordinal))
                throw new InvalidDataException(
                    "Rebuilt MathType number lost its single native tab separator.");

            if (!MathTypeEquationReferences.TryGetVisibleNumberRange(
                    document,
                    placeRef,
                    out visibleNumberRange)
                || visibleNumberRange is null)
                throw new InvalidDataException(
                    "Rebuilt MathType MTPlaceRef has no visible number range.");

            bookmarks = document.Bookmarks;
            try
            {
                previousShowHidden = bookmarks.ShowHidden;
                bookmarks.ShowHidden = true;
            }
            catch { }
            foreach (var bookmarkName in plan.BookmarkNames)
            {
                if (!bookmarks.Exists(bookmarkName))
                    throw new InvalidDataException(
                        $"Rebuilt MathType MTPlaceRef lost native bookmark '{bookmarkName}'.");
                Release(bookmarkRange);
                bookmarkRange = null;
                Release(bookmark);
                bookmark = bookmarks[bookmarkName];
                bookmarkRange = bookmark.Range;
                if (bookmarkRange.Start < visibleNumberRange.Start
                    || bookmarkRange.End > visibleNumberRange.End)
                    throw new InvalidDataException(
                        $"Rebuilt MathType bookmark '{bookmarkName}' no longer wraps the visible equation number.");
            }
        }
        finally
        {
            if (bookmarks is not null)
            {
                try { bookmarks.ShowHidden = previousShowHidden; } catch { }
            }
            Release(bookmarkRange);
            Release(bookmark);
            Release(bookmarks);
            Release(visibleNumberRange);
            Release(oleFormat);
            Release(separator);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(fieldRange);
            Release(result);
            Release(code);
        }
    }

    private static int ResolvePlaceRefFieldEndExclusive(
        Document document,
        Range code,
        Range result)
    {
        var codeBoundary = ReadDocumentCharacter(document, code.End);
        if (codeBoundary == '\u0015') return code.End + 1;
        if (codeBoundary == '\u0014'
            && ReadDocumentCharacter(document, result.End) == '\u0015')
            return result.End + 1;
        throw new InvalidDataException(
            $"MathType MTPlaceRef has an invalid outer field boundary at code={code.End}, result={result.End}.");
    }

    private static char ReadDocumentCharacter(Document document, int position)
    {
        Range? probe = null;
        try
        {
            if (position < document.Content.Start
                || position >= document.Content.End)
                return '\0';
            probe = document.Range(position, position + 1);
            var text = probe.Text ?? string.Empty;
            return text.Length == 1 ? text[0] : '\0';
        }
        finally { Release(probe); }
    }

    private static int ReadDocumentParagraphCount(Document document)
    {
        Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            return paragraphs.Count;
        }
        finally { Release(paragraphs); }
    }

    private static int ReadDocumentInlineShapeCount(Document document)
    {
        InlineShapes? shapes = null;
        try
        {
            shapes = document.InlineShapes;
            return shapes.Count;
        }
        finally { Release(shapes); }
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

using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using Range = Microsoft.Office.Interop.Word.Range;
using WordApplication = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>
/// Narrow adapter around third-party MathType OLE hosts. The rebuilt core does
/// not modify MTEF, storage, rendering, or MathType numbering implementation.
/// This adapter only resolves/reads/deletes one exact local Equation.DSMT4 host
/// when a route crosses into OMML/VisualTeX.
/// </summary>
internal static class WordMathTypeHostAdapter
{
    internal sealed class Semantic
    {
        internal FormulaMetadata Metadata { get; set; } = new();
        internal string MathMl { get; set; } = string.Empty;
    }

    internal static IReadOnlyList<WordFormulaHostDescriptor> CaptureDocumentIndex(
        WordApplication application,
        Document document)
    {
        Range? content = null;
        try
        {
            content = document.Content;
            return CaptureScopeIndex(
                application,
                document,
                content);
        }
        finally { Release(content); }
    }

    internal static IReadOnlyList<WordFormulaHostDescriptor> CaptureScopeIndex(
        WordApplication application,
        Document document,
        Range scope,
        IDictionary<WordFormulaHostDescriptor, Semantic>? capturedSemantics = null)
    {
        var result = new List<WordFormulaHostDescriptor>();
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? range = null;
        try
        {
            shapes = scope.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                shape = shapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape))
                {
                    Release(shape);
                    shape = null;
                    continue;
                }

                range = shape.Range.Duplicate;
                // Conversion needs both the host and its source MathML in this
                // same synchronous capture. Export/decode the object once and
                // share that evidence only with this operation's target builder.
                Semantic? semantic;
                FormulaMetadata metadata;
                try
                {
                    semantic = capturedSemantics is null
                        ? null
                        : ReadShapeSemantic(application, shape);
                    metadata = semantic?.Metadata
                        ?? MathTypeOleInterop.ReadMetadata(application, shape);
                }
                catch (InvalidDataException error) when (MathTypeOleInterop.IsBlankEquationError(error))
                {
                    // Nothing to convert or number: leave the empty object as it is.
                    Release(range); range = null;
                    Release(shape); shape = null;
                    continue;
                }
                var host = new WordFormulaHostDescriptor
                {
                    Kind = WordFormulaHostKind.MathType,
                    Range = new WordFormulaRangeAddress
                    {
                        StoryType = range.StoryType,
                        Start = range.Start,
                        End = range.End,
                    },
                    DisplayMode = string.Equals(
                            metadata.DisplayMode,
                            "block",
                            StringComparison.OrdinalIgnoreCase)
                        ? "block"
                        : "inline",
                    FormulaId = metadata.FormulaId,
                    Metadata = metadata,
                    MetadataAuthoritative = true,
                    Numbering = new WordFormulaNumberingDescriptor
                    {
                        Numbered = metadata.Numbered,
                        FormulaId = metadata.FormulaId,
                    },
                    WithinTable = IsWithinTable(range),
                    Latex = metadata.Latex,
                };
                result.Add(host);
                if (capturedSemantics is not null)
                    capturedSemantics.Add(host, semantic!);
                Release(range); range = null;
                Release(shape); shape = null;
            }
            return result;
        }
        finally
        {
            Release(range);
            Release(shape);
            Release(shapes);
        }
    }

    internal static Semantic ReadSemantic(
        WordApplication application,
        Document document,
        WordFormulaHostDescriptor host)
    {
        if (host.Kind != WordFormulaHostKind.MathType)
            throw new ArgumentException(
                "MathType semantic reader received another host kind.");

        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? shape = null;
        Range? shapeRange = null;
        try
        {
            range = WordFormulaHostSemanticReader.CreateRange(
                document,
                host.Range);
            shapes = range.InlineShapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                var candidate = shapes[index];
                Range? candidateRange = null;
                try
                {
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate))
                        continue;
                    candidateRange = candidate.Range.Duplicate;
                    if (!WordFormulaHostSemanticReader.SameAddress(
                            candidateRange,
                            host.Range))
                        continue;
                    if (shape is not null)
                        throw new InvalidDataException(
                            "The MathType host range contains more than one Equation.DSMT4 object.");
                    shape = candidate;
                    candidate = null;
                    shapeRange = candidateRange;
                    candidateRange = null;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            if (shape is null)
                throw new InvalidDataException(
                    "The captured MathType host no longer exists.");

            return ReadShapeSemantic(application, shape);
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }
    private static Semantic ReadShapeSemantic(
        WordApplication application,
        InlineShape shape)
    {
        var compound = MathTypeOleStorage.CaptureCompoundFile(shape);
        var native = MathTypeOleStorage.ReadEquationNative(compound);
        var mathMl = MathTypeMtefCodec.ReadEquationNativeMathMl(native);
        var metadata = MathTypeOleInterop.ReadMetadata(
            application, shape, knownMathMl: mathMl, knownCompoundFile: compound);
        return new Semantic { Metadata = metadata, MathMl = mathMl };
    }

    internal static WordFormulaHostDescriptor? ResolveLocal(
        WordApplication application,
        Document document,
        Range selectionRange)
    {
        using var binding = BindLocal(application, document, selectionRange, requireExact: false);
        return binding?.Host;
    }

    internal static WordNativeConversionSource BindCapturedHost(
        WordApplication application,
        Document document,
        Range capturedRange) =>
        BindLocal(application, document, capturedRange, requireExact: true)
        ?? throw new InvalidDataException("The captured MathType host no longer exists at its exact range.");

    private static WordNativeConversionSource? BindLocal(
        WordApplication application,
        Document document,
        Range selectionRange,
        bool requireExact)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? selected = null;
        Range? selectedRange = null;
        try
        {
            probe = selectionRange.Duplicate;
            if (!requireExact)
            {
                try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch { }
                try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch { }
            }
            shapes = probe.InlineShapes;

            for (var index = 1; index <= shapes.Count; index++)
            {
                InlineShape? candidate = null;
                Range? candidateRange = null;
                try
                {
                    candidate = shapes[index];
                    if (!MathTypeOleInterop.IsMathTypeOle(candidate))
                        continue;
                    candidateRange = candidate.Range.Duplicate;
                    var matches = requireExact
                        ? candidateRange.StoryType == selectionRange.StoryType
                            && candidateRange.Start == selectionRange.Start
                            && candidateRange.End == selectionRange.End
                        : Touches(selectionRange, candidateRange);
                    if (!matches)
                        continue;

                    if (selected is not null)
                        throw new InvalidDataException(
                            "The local range contains more than one MathType host.");
                    selected = candidate;
                    candidate = null;
                    selectedRange = candidateRange;
                    candidateRange = null;
                }
                finally
                {
                    Release(candidateRange);
                    Release(candidate);
                }
            }

            if (selected is null || selectedRange is null)
                return null;

            WordFormulaMutationTransaction.TraceBoundary(application, "mathtype-resolve-before-metadata");
            var semantic = ReadShapeSemantic(application, selected);
            var metadata = semantic.Metadata;
            WordFormulaMutationTransaction.TraceBoundary(application, "mathtype-resolve-after-metadata");
            var displayMode = string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase)
                ? "block"
                : "inline";

            var host = new WordFormulaHostDescriptor
            {
                Kind = WordFormulaHostKind.MathType,
                Range = new WordFormulaRangeAddress
                {
                    StoryType = selectedRange.StoryType,
                    Start = selectedRange.Start,
                    End = selectedRange.End,
                },
                DisplayMode = displayMode,
                FormulaId = metadata.FormulaId,
                Metadata = metadata,
                MetadataAuthoritative = true,
                Numbering = new WordFormulaNumberingDescriptor
                {
                    Numbered = metadata.Numbered,
                    FormulaId = metadata.FormulaId,
                    Position = "right",
                },
                WithinTable = IsWithinTable(selectedRange),
                Latex = metadata.Latex,
            };
            var ownedShape = selected;
            selected = null;
            return WordNativeConversionSource.AdoptMathType(
                document, host, ownedShape, semantic.MathMl);
        }
        finally
        {
            Release(selectedRange);
            Release(selected);
            Release(shapes);
            Release(probe);
        }
    }

    internal static Range DeleteExact(
        Document document,
        WordFormulaHostDescriptor source,
        InlineShape? preparedSource = null)
    {
        if (source.Kind != WordFormulaHostKind.MathType)
            throw new ArgumentException(
                "MathType adapter received another host kind.");

        Range? range = null;
        InlineShapes? shapes = null;
        InlineShape? selected = null;
        Range? selectedRange = null;
        Range? anchor = null;
        try
        {
            if (preparedSource is not null)
            {
                selected = preparedSource;
                selectedRange = selected.Range;
            }
            else
            {
                range = WordFormulaHostSemanticReader.CreateRange(
                    document,
                    source.Range);
                shapes = range.InlineShapes;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    InlineShape? candidate = null;
                    Range? candidateRange = null;
                    try
                    {
                        candidate = shapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(candidate))
                            continue;
                        candidateRange = candidate.Range.Duplicate;
                        if (!WordFormulaHostSemanticReader.SameAddress(
                                candidateRange,
                                source.Range))
                            continue;

                        if (selected is not null)
                            throw new InvalidDataException(
                                "The MathType source range contains multiple matching hosts.");
                        selected = candidate;
                        candidate = null;
                        selectedRange = candidateRange;
                        candidateRange = null;
                    }
                    finally
                    {
                        Release(candidateRange);
                        Release(candidate);
                    }
                }
            }

            if (selected is null || selectedRange is null)
                throw new InvalidDataException(
                    "The MathType source host moved before deletion.");

            var display = string.Equals(
                source.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase);
            if (display)
            {
                Paragraphs? paragraphs = null;
                Paragraph? paragraph = null;
                Range? paragraphRange = null;
                Range? body = null;
                try
                {
                    paragraphs = selectedRange.Paragraphs;
                    if (paragraphs.Count != 1)
                        throw new InvalidDataException(
                            "The MathType display source no longer occupies one paragraph.");
                    paragraph = paragraphs[1];
                    paragraphRange = paragraph.Range.Duplicate;
                    if (!IsSafeMathTypeDisplayParagraph(paragraphRange))
                        throw new InvalidDataException(
                            "The MathType display paragraph contains user content or unsupported fields.");

                    // Delete the exact OLE owner first. Clearing a paragraph
                    // containing both external OLE storage and nested numbering
                    // fields makes Word serialize/reflow that mixed structure.
                    // The remaining text is the scaffold proved above, not an
                    // object deletion disguised as a paragraph text assignment.
                    using (WordOperationMetrics.Measure("FormatConversion.MathTypeDelete.DisplayOle"))
                        selectedRange.Text = string.Empty;
                    Release(paragraphRange);
                    paragraphRange = paragraph.Range.Duplicate;
                    var bodyEnd = Math.Max(
                        paragraphRange.Start,
                        paragraphRange.End - 1);
                    anchor = document.Range(
                        paragraphRange.Start,
                        paragraphRange.Start);
                    body = document.Range(
                        paragraphRange.Start,
                        bodyEnd);
                    using (WordOperationMetrics.Measure("FormatConversion.MathTypeDelete.DisplayScaffold"))
                        body.Text = string.Empty;

                    Range? verify = null;
                    InlineShapes? verifyShapes = null;
                    OMaths? verifyMaths = null;
                    Fields? verifyFields = null;
                    try
                    {
                        verify = paragraph.Range.Duplicate;
                        verifyShapes = verify.InlineShapes;
                        verifyMaths = verify.OMaths;
                        verifyFields = WordFormulaHost.GetLocalFields(verify);
                        if (verifyShapes.Count != 0
                            || verifyMaths.Count != 0
                            || verifyFields.Count != 0)
                            throw new InvalidDataException(
                                "The MathType display source did not clear to one empty paragraph.");
                    }
                    finally
                    {
                        Release(verifyFields);
                        Release(verifyMaths);
                        Release(verifyShapes);
                        Release(verify);
                    }
                }
                finally
                {
                    Release(body);
                    Release(paragraphRange);
                    Release(paragraph);
                    Release(paragraphs);
                }
            }
            else
            {
                anchor = selectedRange.Duplicate;
                anchor.Collapse(WdCollapseDirection.wdCollapseStart);
                using (WordOperationMetrics.Measure("FormatConversion.MathTypeDelete.InlineOle"))
                    selectedRange.Text = string.Empty;
            }

            // Physical deletion normally retires its range bookmark too;
            // remove only any remaining marker, after the owner is gone.
            using (WordOperationMetrics.Measure("FormatConversion.MathTypeDelete.RemoveIdentity"))
                WordFormulaIdentityStore.RemoveMathType(document, source.FormulaId);
            var returned = anchor;
            anchor = null;
            return returned;
        }
        finally
        {
            Release(anchor);
            Release(selectedRange);
            if (preparedSource is null) Release(selected);
            Release(shapes);
            Release(range);
        }
    }

    // An external MathType number is a sibling scaffold, not part of the OLE
    // atom. Prepare its exact live ranges outside the target transaction; remove
    // only that proven scaffold while retaining the source OLE for the shared
    // physical OMML replacement. Never clear an entire historical paragraph.
    internal sealed class NumberedScaffold : IDisposable
    {
        private Range? _left;
        private Range? _right;
        private readonly List<string> _numberBookmarks;

        internal NumberedScaffold(Range left, Range right, List<string> numberBookmarks)
        { _left = left; _right = right; _numberBookmarks = numberBookmarks; }

        internal void Remove()
        {
            if (_left is null || _right is null)
                throw new InvalidOperationException("The prepared MathType numbering scaffold is unavailable.");
            WordFormulaMutationTransaction.TraceBoundary(_right.Application, "number-scaffold-before-right-delete");
            if (_right.End > _right.Start) _right.Text = string.Empty;
            WordFormulaMutationTransaction.TraceBoundary(_left.Application, "number-scaffold-before-left-delete");
            if (_left.End > _left.Start) _left.Text = string.Empty;
            WordFormulaMutationTransaction.TraceBoundary(_left.Application, "number-scaffold-after-delete");
        }

        internal void RestoreNumberBookmarks(Document document, Range equation)
        {
            if (_numberBookmarks.Count == 0) return;
            Fields? fields = null;
            Field? field = null;
            Range? code = null;
            Range? result = null;
            Range? number = null;
            Bookmarks? bookmarks = null;
            try
            {
                fields = equation.Fields;
                var start = int.MaxValue;
                var end = int.MinValue;
                var sequences = 0;
                for (var index = 1; index <= fields.Count; index++)
                {
                    Release(result); result = null;
                    Release(code); code = null;
                    Release(field); field = fields[index];
                    code = field.Code;
                    var instruction = (code.Text ?? string.Empty).TrimStart();
                    if (!instruction.StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase)
                        && !instruction.StartsWith("STYLEREF ", StringComparison.OrdinalIgnoreCase)) continue;
                    if (instruction.StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase)) sequences++;
                    result = field.Result;
                    start = Math.Min(start, result.Start);
                    end = Math.Max(end, result.End);
                }
                if (sequences != 1 || start > end)
                    throw new InvalidDataException("The replacement did not retain one native number for the source references.");
                number = WordNativeOmmlNumbering.ResolveNumberContentRange(equation);
                if (number.Start > start || number.End < end)
                    throw new InvalidDataException("The native number argument does not contain the captured field results.");
                bookmarks = document.Bookmarks;
                foreach (var name in _numberBookmarks)
                {
                    Bookmark? bookmark = null;
                    try { bookmark = bookmarks.Add(name, number); }
                    finally { Release(bookmark); }
                }
            }
            finally
            {
                Release(bookmarks); Release(number); Release(result); Release(code); Release(field); Release(fields);
            }
        }

        public void Dispose()
        { Release(_right); _right = null; Release(_left); _left = null; }
    }

    internal static NumberedScaffold PrepareNumberedScaffold(Document document, Range source, string ownerXml)
    {
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        Range? left = null;
        Range? right = null;
        InlineShapes? leftShapes = null;
        InlineShapes? rightShapes = null;
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        Range? result = null;
        Bookmarks? bookmarks = null;
        try
        {
            paragraphs = source.Paragraphs;
            if (paragraphs.Count != 1)
                throw new InvalidDataException("A numbered MathType source must have one exact paragraph owner.");
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            if (!IsSafeMathTypeDisplayParagraph(owner))
                throw new InvalidDataException("The MathType number scaffold contains unowned text or fields.");
            left = source.Duplicate; left.SetRange(owner.Start, source.Start);
            right = source.Duplicate; right.SetRange(source.End, Math.Max(source.End, owner.End - 1));
            leftShapes = left.InlineShapes; rightShapes = right.InlineShapes;
            if (leftShapes.Count != 0 || rightShapes.Count != 0)
                throw new InvalidDataException("The MathType number scaffold contains another current or historical OLE.");
            fields = owner.Fields;
            var numberStart = -1;
            var numberEnd = -1;
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                var tokens = (code.Text ?? string.Empty).Trim().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (tokens.Length < 2 || !string.Equals(tokens[0], "MACROBUTTON", StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(tokens[1], "MTPlaceRef", StringComparison.OrdinalIgnoreCase)) continue;
                result = field.Result;
                // A code-only MACROBUTTON has no separate/result atom: Word's
                // empty Result already ends after its field-end character.
                // Result.End+1 would include the user's paragraph mark. Prove
                // code and result against the already-owned sibling range; do
                // not manufacture a field extent with control-character offsets.
                var inLeft = code.Start >= left.Start && code.End <= left.End
                    && result.Start >= left.Start && result.End <= left.End;
                var inRight = code.Start >= right.Start && code.End <= right.End
                    && result.Start >= right.Start && result.End <= right.End;
                if (numberStart >= 0 || inLeft == inRight
                    || !code.InStory(source) || !result.InStory(source))
                    throw new InvalidDataException("The source does not own one unambiguous sibling MathType number.");
                numberStart = inLeft ? left.Start : right.Start;
                numberEnd = inLeft ? left.End : right.End;
            }
            if (numberStart < 0)
                throw new InvalidDataException("The numbered MathType source has no exact MTPlaceRef field.");
            var names = new List<string>();
            bookmarks = document.Bookmarks;
            foreach (var name in MathTypeLocalFragment.BookmarkNames(ownerXml))
            {
                Bookmark? bookmark = null;
                Range? range = null;
                try
                {
                    if (!bookmarks.Exists(name)) continue;
                    bookmark = bookmarks[name]; range = bookmark.Range;
                    if (range.InStory(source) && range.Start >= numberStart && range.End <= numberEnd)
                        names.Add(name);
                }
                finally { Release(range); Release(bookmark); }
            }
            var plan = new NumberedScaffold(left, right, names);
            left = null; right = null;
            return plan;
        }
        finally
        {
            Release(bookmarks); Release(result); Release(code); Release(field); Release(fields);
            Release(rightShapes); Release(leftShapes); Release(right); Release(left); Release(owner); Release(paragraph); Release(paragraphs);
        }
    }

    private static bool IsSafeMathTypeDisplayParagraph(
        Range paragraphRange)
    {
        Fields? fields = null;
        Field? field = null;
        Range? code = null;
        try
        {
            fields = WordFormulaHost.GetLocalFields(
                paragraphRange);
            for (var index = 1; index <= fields.Count; index++)
            {
                field = fields[index];
                code = field.Code;
                if (!IsKnownMathTypeDisplayFieldCode(code.Text))
                    return false;
                Release(code); code = null;
                Release(field); field = null;
            }

            var text = paragraphRange.Text ?? string.Empty;
            var fieldDepth = 0;
            foreach (var character in text)
            {
                if (character == '\u0013')
                {
                    fieldDepth++;
                    continue;
                }
                if (character == '\u0015')
                {
                    if (fieldDepth > 0) fieldDepth--;
                    continue;
                }
                if (fieldDepth > 0 || character == '\u0014')
                    continue;

                if (char.IsWhiteSpace(character)
                    || char.IsDigit(character))
                    continue;
                if (character < ' ') continue;
                if (character is '\u0001' or '\uFFFC')
                    continue;
                if ("()[]{}.-–—_:;,+/\\".IndexOf(character) >= 0)
                    continue;
                return false;
            }
            return fieldDepth == 0;
        }
        finally
        {
            Release(code);
            Release(field);
            Release(fields);
        }
    }

    private static bool IsKnownMathTypeDisplayFieldCode(
        string? value)
    {
        var code = (value ?? string.Empty).Trim();
        if (code.Length == 0) return false;
        if (code.IndexOf(
                "EMBED Equation.DSMT4",
                StringComparison.OrdinalIgnoreCase) >= 0)
            return true;
        if (code.StartsWith(
                "MACROBUTTON MTPlaceRef",
                StringComparison.OrdinalIgnoreCase))
            return true;
        return code.StartsWith(
                "SEQ MTEqn ",
                StringComparison.OrdinalIgnoreCase)
            || code.StartsWith(
                "SEQ MTChap ",
                StringComparison.OrdinalIgnoreCase)
            || code.StartsWith(
                "SEQ MTSec ",
                StringComparison.OrdinalIgnoreCase);
    }

    private static bool Touches(Range selection, Range candidate)
    {
        if (selection.StoryType != candidate.StoryType)
            return false;
        if (selection.Start == selection.End)
            return candidate.Start <= selection.Start
                && selection.Start <= candidate.End;
        return selection.Start < candidate.End
            && selection.End > candidate.Start;
    }

    private static bool IsWithinTable(Range range)
    {
        try
        {
            // This adapter supplies an exact InlineShape range. Table ownership
            // is structural; asking Word's pagination-backed Information getter
            // here can spend seconds laying out the entire historical document.
            // Reuse the same bounded proof as native-host lookup. No XML export,
            // guessed cell offset or document-wide table enumeration is required.
            return WordFormulaHostLookup.IsWhollyInsideTable(range);
        }
        catch { return false; }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

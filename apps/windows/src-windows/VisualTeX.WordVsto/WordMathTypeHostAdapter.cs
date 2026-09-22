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
        Range scope)
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
                var metadata = MathTypeOleInterop.ReadMetadata(
                    application,
                    shape);
                result.Add(new WordFormulaHostDescriptor
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
                });
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

            var compound =
                MathTypeOleStorage.CaptureCompoundFile(shape);
            var native =
                MathTypeOleStorage.ReadEquationNative(compound);
            var mathMl =
                MathTypeMtefCodec.ReadEquationNativeMathMl(native);
            var metadata =
                MathTypeOleInterop.ReadMetadata(
                    application,
                    shape,
                    knownMathMl: mathMl,
                    knownCompoundFile: compound);
            return new Semantic
            {
                Metadata = metadata,
                MathMl = mathMl,
            };
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(range);
        }
    }
    internal static WordFormulaHostDescriptor? ResolveLocal(
        WordApplication application,
        Document document,
        Range selectionRange)
    {
        Range? probe = null;
        InlineShapes? shapes = null;
        InlineShape? selected = null;
        Range? selectedRange = null;
        try
        {
            probe = selectionRange.Duplicate;
            try { probe.MoveStart(WdUnits.wdCharacter, -2); } catch { }
            try { probe.MoveEnd(WdUnits.wdCharacter, 2); } catch { }
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
                    if (!Touches(selectionRange, candidateRange))
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

            var metadata = MathTypeOleInterop.ReadMetadata(
                application,
                selected);
            var displayMode = string.Equals(
                    metadata.DisplayMode,
                    "block",
                    StringComparison.OrdinalIgnoreCase)
                ? "block"
                : "inline";

            return new WordFormulaHostDescriptor
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
        WordFormulaHostDescriptor source)
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

            if (selected is null || selectedRange is null)
                throw new InvalidDataException(
                    "The MathType source host moved before deletion.");

            var display = string.Equals(
                source.DisplayMode,
                "block",
                StringComparison.OrdinalIgnoreCase);
            WordFormulaIdentityStore.RemoveMathType(
                document,
                source.FormulaId);
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

                    var bodyEnd = Math.Max(
                        paragraphRange.Start,
                        paragraphRange.End - 1);
                    anchor = document.Range(
                        paragraphRange.Start,
                        paragraphRange.Start);
                    body = document.Range(
                        paragraphRange.Start,
                        bodyEnd);
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
                selected.Delete();
            }

            var returned = anchor;
            anchor = null;
            return returned;
        }
        finally
        {
            Release(anchor);
            Release(selectedRange);
            Release(selected);
            Release(shapes);
            Release(range);
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
            return Convert.ToBoolean(
                range.get_Information(WdInformation.wdWithInTable));
        }
        catch { return false; }
    }

    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.ReleaseComObject(value); } catch { }
    }
}

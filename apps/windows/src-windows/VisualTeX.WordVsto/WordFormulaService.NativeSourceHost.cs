using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal sealed partial class WordFormulaService
{
    private static bool IsSafeMathTypeRowWithSectionPrefix(Document document, Range paragraph, Range equation)
    {
        if (!MathTypeSourceHost.TryGetSectionPrefixSplit(document, paragraph, equation, out var split))
            return false;
        Range? row = null;
        try
        {
            row = document.Range(split, paragraph.End);
            return row.InlineShapes.Count == 1 && IsSafeMathTypeDisplayParagraph(row);
        }
        finally { Release(row); }
    }

    // Called once for this exact source, in descending source order, inside the
    // already-established conversion Undo. Never changes other MathType sources
    // or discards their chapter/section reset fields.
    private void IsolateMathTypeSectionPrefix(Document document, WordFormulaFormatConversionTarget target)
    {
        InlineShape? shape = null;
        Range? equation = null;
        Paragraphs? paragraphs = null;
        Paragraph? paragraph = null;
        Range? owner = null;
        Range? prefix = null;
        Range? splitRange = null;
        Range? state = null;
        Range? row = null;
        try
        {
            shape = FindMathTypeOleByRange(document, target.SourceObjectId, allowGlobalFallback: false)
                ?? throw new InvalidDataException("The native MathType source moved before section-state isolation.");
            equation = shape.Range;
            paragraphs = equation.Paragraphs;
            if (paragraphs.Count != 1) throw new InvalidDataException("Native MathType source has no unique paragraph.");
            paragraph = paragraphs[1];
            owner = paragraph.Range;
            if (!MathTypeSourceHost.TryGetSectionPrefixSplit(document, owner, equation, out var split)
                || !IsSafeMathTypeRowWithSectionPrefix(document, owner, equation))
                throw new InvalidDataException("MathType section state changed after conversion preflight.");
            var stateStart = owner.Start;
            prefix = document.Range(stateStart, split);
            var originalFieldCodes = CaptureHostFieldCodes(prefix);
            var beforeTables = document.Tables.Count;
            using var rowFormatting = WordParagraphFormatting.Capture(owner);
            splitRange = document.Range(split, split);
            splitRange.Text = "\r";
            Release(equation); equation = shape.Range;
            Release(paragraphs); paragraphs = equation.Paragraphs;
            Release(paragraph); paragraph = paragraphs[1];
            row = paragraph.Range;
            state = document.Range(stateStart, row.Start);
            if (document.Tables.Count != beforeTables || state.InlineShapes.Count != 0
                || state.OMaths.Count != 0 || !originalFieldCodes.SequenceEqual(CaptureHostFieldCodes(state))
                || !IsSafeMathTypeDisplayParagraph(row))
                throw new InvalidDataException("Native MathType section-state isolation failed its ownership validation.");
            rowFormatting.Apply(row);
            // This is MathType's existing state style, not a new global/default
            // font setting. Only the isolated state paragraph receives it.
            object sectionStyle = "MTEquationSection";
            state.set_Style(ref sectionStyle);
            target.SourceObjectId = $"{RangeReferencePrefix}{equation.Start}:{equation.End}";
            target.SourceStart = equation.Start;
            target.SourceHasMathTypeSectionPrefix = false;
            WordDoubleClickHook.TraceMessage($"format-conversion-native-section-preserved sourceId={target.SourceFormulaId} state={stateStart}:{row.Start} formula={equation.Start}:{equation.End}");
        }
        finally
        {
            Release(row); Release(state); Release(splitRange); Release(prefix);
            Release(owner); Release(paragraph); Release(paragraphs); Release(equation); Release(shape);
        }
    }

    private static string[] CaptureHostFieldCodes(Range range)
    {
        Fields? fields = null;
        var result = new List<string>();
        try
        {
            fields = WordFormulaHost.GetLocalFields(range);
            for (var i = 1; i <= fields.Count; i++)
            {
                Field? field = null;
                Range? code = null;
                try { field = fields[i]; code = field.Code; result.Add(code.Text ?? string.Empty); }
                finally { Release(code); Release(field); }
            }
            return result.ToArray();
        }
        finally { Release(fields); }
    }
}

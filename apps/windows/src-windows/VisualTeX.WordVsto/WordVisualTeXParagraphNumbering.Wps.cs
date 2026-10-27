using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// In WPS the number tree is inserted as Word's exact field XML (WpsFieldXml): Fields.Add
// would give the MACROBUTTON and the hidden SEQ fields a result that WPS then displays
// instead of the visible "(n)". Only fields that legitimately carry a result are updated.
// ShowCodes is never toggled in WPS: a failed call inside a custom UndoRecord makes WPS
// roll the whole record back.
internal static partial class WordVisualTeXParagraphNumbering
{
    private static bool IsHiddenSequenceInstruction(string? instruction) =>
        WordFieldCode.IsHiddenSequence(instruction);

    private static Field CreatePlaceRefWps(
        Document document,
        int position,
        ResolvedNumberPlan plan,
        CreatedBodyFields? createdFields)
    {
        var segments = plan.Segments ?? BuildSegments(plan);
        var runs = new List<XElement>
        {
            WpsFieldXml.FieldChar("begin"),
            WpsFieldXml.Instruction(" " + PlaceRefMarker + " "),
        };
        var fieldSegments = new List<(int Index, string Instruction)>();
        for (var index = 0; index < segments.Count; index++)
        {
            var segment = segments[index];
            if (!segment.IsField)
            {
                if (!string.IsNullOrEmpty(segment.Value))
                    runs.Add(WpsFieldXml.Instruction(segment.Value));
                continue;
            }
            runs.AddRange(WpsFieldXml.ComplexField(
                segment.Value,
                IsHiddenSequenceInstruction(segment.Value) ? null : "1"));
            fieldSegments.Add((index, segment.Value));
        }
        runs.Add(WpsFieldXml.Instruction(")"));
        runs.Add(WpsFieldXml.FieldChar("end"));

        WpsFieldXml.InsertRuns(document, position, runs);
        var outer = WpsFieldXml.FindFieldStartingAt(
                document,
                position,
                field => field.Type == WdFieldType.wdFieldMacroButton
                    && IsPlaceRefCode(ReadCode(field)))
            ?? throw new InvalidDataException(
                "WPS did not materialize the VisualTeXPlaceRef field at its insertion point.");

        if (createdFields is not null)
            RegisterNestedWpsFields(outer, fieldSegments, createdFields);
        return outer;
    }

    private static Field InsertHiddenStateFieldWps(
        Document document,
        int position,
        string instruction)
    {
        WpsFieldXml.InsertRuns(
            document,
            position,
            WpsFieldXml.ComplexField(instruction, result: null).ToList());
        return WpsFieldXml.FindFieldStartingAt(
                document,
                position,
                field => IsSameInstruction(ReadCode(field), instruction))
            ?? throw new InvalidDataException(
                "WPS did not materialize the VisualTeX heading-state field at its insertion point.");
    }

    // WPS: one Update of the outer MACROBUTTON gives it a separator and an empty result.
    // Update only the nested fields that carry a visible result, headings first.
    private static void UpdatePlaceRefTreeWps(Field placeRef) =>
        UpdatePlaceRefFields(placeRef);

    private static void RegisterNestedWpsFields(
        Field outer,
        IReadOnlyList<(int Index, string Instruction)> fieldSegments,
        CreatedBodyFields createdFields)
    {
        Range? code = null;
        Fields? nested = null;
        try
        {
            code = outer.Code;
            nested = code.Fields;
            var ordered = new List<(int Start, Field Field)>();
            for (var index = 1 + WordFieldCode.OwnerEntries(nested, code); index <= nested.Count; index++)
            {
                var field = nested[index];
                Range? fieldCode = null;
                try
                {
                    fieldCode = field.Code;
                    ordered.Add((fieldCode.Start, field));
                }
                finally { Release(fieldCode); }
            }
            ordered.Sort((left, right) => left.Start.CompareTo(right.Start));
            if (ordered.Count != fieldSegments.Count)
            {
                foreach (var item in ordered) Release(item.Field);
                throw new InvalidDataException(
                    $"WPS materialized {ordered.Count} nested number fields; expected {fieldSegments.Count}.");
            }
            for (var index = 0; index < ordered.Count; index++)
                createdFields.Add(fieldSegments[index].Index, fieldSegments[index].Instruction, ordered[index].Field);
        }
        finally
        {
            Release(nested);
            Release(code);
        }
    }

    private static string ReadCode(Field field)
    {
        Range? code = null;
        try
        {
            code = field.Code;
            return code.Text ?? string.Empty;
        }
        finally { Release(code); }
    }
}

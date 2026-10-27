using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

internal static class WordFieldCode
{
    // Range.Fields of a field's code. Word lists only the fields nested in that code;
    // WPS lists the owner field (and its ancestors) first. The number of those leading
    // entries, always 0 in Word; nested field n is codeFields[n + OwnerEntries].
    internal static int OwnerEntries(Fields codeFields, Range code)
    {
        if (!OfficeHostEnvironment.IsWps) return 0;
        var entries = 0;
        for (var index = 1; index <= codeFields.Count; index++)
        {
            Field? field = null;
            Range? fieldCode = null;
            try
            {
                field = codeFields[index];
                fieldCode = field.Code;
                if (fieldCode.Start > code.Start) break;
                entries++;
            }
            finally
            {
                if (fieldCode is not null) Marshal.ReleaseComObject(fieldCode);
                if (field is not null) Marshal.ReleaseComObject(field);
            }
        }
        return entries;
    }

    // WPS reads each "*" of a field instruction inside an equation as U+2217, and a field
    // whose stored instruction holds U+2217 (a Word-made number opened in WPS) updates to
    // "Error! Bookmark not defined.". Writing the code back with "*" before an update repairs
    // it, as the macOS WPS add-in does. Word: no change (never rewrite a code inside an OMath).
    internal static void RepairWpsEquationInstruction(Range code, ref string instruction)
    {
        if (!OfficeHostEnvironment.IsWps || instruction.IndexOf('\u2217') < 0) return;
        var repaired = instruction.Replace('\u2217', '*');
        code.Text = repaired;
        instruction = repaired;
    }

    // The fields of every equation in the main story whose instruction WPS reads with U+2217:
    // Word writes the STYLEREF/SEQ of an equation's #(number) as math text, and WPS turns its
    // "*" into U+2217 when it opens the document. WPS's own update then shows "Error! Bookmark
    // not defined." and saving writes U+2217 into the file, which Word reads the same way.
    // Code.Text keeps reading U+2217 after the repair; the stored instruction has "*".
    internal static int RepairWpsEquationInstructions(Document document)
    {
        if (!OfficeHostEnvironment.IsWps) return 0;
        var repaired = 0;
        OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                Fields? fields = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    fields = range.Fields;
                    for (var position = 1; position <= fields.Count; position++)
                    {
                        Field? field = null;
                        Range? code = null;
                        try
                        {
                            field = fields[position];
                            code = field.Code;
                            var instruction = code.Text ?? string.Empty;
                            if (instruction.IndexOf('\u2217') < 0) continue;
                            RepairWpsEquationInstruction(code, ref instruction);
                            repaired++;
                        }
                        finally
                        {
                            if (code is not null) Marshal.ReleaseComObject(code);
                            if (field is not null) Marshal.ReleaseComObject(field);
                        }
                    }
                }
                finally
                {
                    if (fields is not null) Marshal.ReleaseComObject(fields);
                    if (range is not null) Marshal.ReleaseComObject(range);
                    if (math is not null) Marshal.ReleaseComObject(math);
                }
            }
        }
        finally
        {
            if (maths is not null) Marshal.ReleaseComObject(maths);
        }
        return repaired;
    }

    // WPS gives a MACROBUTTON or a hidden (\h) SEQ a separator and an empty result
    // when it updates the field, and then displays that empty result.
    internal static bool IsHiddenSequence(string? instruction)
    {
        var normalized = (instruction ?? string.Empty).TrimStart();
        return normalized.StartsWith("SEQ ", StringComparison.OrdinalIgnoreCase)
            && normalized.IndexOf("\\h", StringComparison.OrdinalIgnoreCase) >= 0;
    }
}

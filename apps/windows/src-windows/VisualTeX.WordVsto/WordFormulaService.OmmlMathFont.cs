using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using static VisualTeX.WordVsto.WordOmmlMathFont;

namespace VisualTeX.WordVsto;

/// <summary>
/// OMML Font: the document math font from the ribbon choice, and existing equations
/// brought in line with the document math font. Word draws a math run without w:rFonts
/// in the document math font, but VisualTeX (and Word itself) usually writes the font
/// on every run, so changing Document.OMathFontName alone can leave old equations in
/// the old font. Only formatting changes: Find/Replace of exactly the run fonts that
/// <see cref="WordOmmlMathFontPolicy.AlignRunFonts"/> selects, inside each OMath range,
/// so text, structure, fields (#(SEQ) numbers), bookmarks and VisualTeX metadata stay.
/// </summary>
internal sealed partial class WordFormulaService
{
    /// <summary>
    /// Ribbon OMML Font choice: applies the chosen font to the active document and
    /// aligns its equations, in one Undo record. Null when there is nothing to do
    /// (follow, WPS, or no open document). WPS: its Find never searches inside
    /// equations, so existing equations keep their fonts until they are edited.
    /// </summary>
    public OmmlMathFontAlignmentResult? ApplyOmmlMathFontToActiveDocument()
    {
        if (OfficeHostEnvironment.IsWps || WordOmmlMathFontPolicy.PreferredFont is null) return null;
        Document? document = null;
        UndoRecord? undo = null;
        try
        {
            try
            {
                if (_application.Documents.Count == 0) return null;
                document = _application.ActiveDocument;
            }
            catch (COMException) { return null; }
            if (document is null) return null;
            EnsureWritable(document);
            undo = _application.UndoRecord;
            if (undo.IsRecordingCustomRecord || undo.CustomRecordLevel > 0)
                throw new InvalidOperationException(
                    "The OMML font cannot be applied inside another Word undo transaction.");
            undo.StartCustomRecord("VisualTeX OMML Font");
            try { return AlignOmmlMathFont(document); }
            finally { undo.EndCustomRecord(); }
        }
        finally
        {
            Release(undo);
            Release(document);
        }
    }
}

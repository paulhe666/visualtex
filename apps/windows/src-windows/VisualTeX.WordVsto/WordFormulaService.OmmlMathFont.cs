using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

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
    internal sealed class OmmlMathFontAlignmentResult
    {
        /// <summary>The document math font after the operation.</summary>
        internal string MathFont { get; set; } = string.Empty;

        /// <summary>The previous document math font when it changed; otherwise null.</summary>
        internal string? PreviousMathFont { get; set; }

        /// <summary>Equations whose run fonts were rewritten.</summary>
        internal int Equations { get; set; }

        /// <summary>m:nor (normal text) runs that keep their own font.</summary>
        internal int KeptNormalTextRuns { get; set; }
    }

    internal sealed class OmmlMathFontScan
    {
        internal string MathFont { get; set; } = string.Empty;

        /// <summary>The current document math font when it will change; otherwise null.</summary>
        internal string? CurrentMathFont { get; set; }

        internal int EquationsToAlign { get; set; }

        internal int Total => (CurrentMathFont is null ? 0 : 1) + EquationsToAlign;
    }

    /// <summary>
    /// Makes <paramref name="requested"/> the document's native Office Math font
    /// (loading it first when it is one of the bundled fonts). Installed as the policy's
    /// document-font applier. WPS has no document math font (it reports the placeholder
    /// "default", and a rejected call inside WPS's custom Undo record rolls back the whole
    /// write): there the font only goes into the equation's runs.
    /// </summary>
    internal static void ApplyDocumentOmmlMathFont(Document document, string requested)
    {
        if (WordOfficeMathFontLoader.Find(requested) is { } bundled)
            WordOfficeMathFontLoader.EnsureLoaded(bundled);
        if (OfficeHostEnvironment.IsWps) return;

        string current;
        try { current = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not read this document's native Office Math font setting.",
                error);
        }
        if (string.Equals(current, requested, StringComparison.OrdinalIgnoreCase))
            return;

        try { document.OMathFontName = requested; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                $"Word could not set this document's native Office Math font to '{requested}'.",
                error);
        }

        string applied;
        try { applied = document.OMathFontName ?? string.Empty; }
        catch (COMException error)
        {
            throw new InvalidOperationException(
                "Word could not verify the document-level Office Math font after applying it.",
                error);
        }
        if (!string.Equals(applied, requested, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Word rejected '{requested}' as the document's native Office Math font. "
                + "Install a valid OpenType math font or repair the VisualTeX Office integration.");
        }
    }

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

    private static string ReadDocumentMathFont(Document document)
    {
        try { return document.OMathFontName ?? string.Empty; }
        catch (COMException) { return string.Empty; }
    }

    // Read-only: what AlignOmmlMathFont would change.
    private static OmmlMathFontScan ScanOmmlMathFont(Document document)
    {
        var scan = new OmmlMathFontScan();
        if (OfficeHostEnvironment.IsWps) return scan;
        var current = ReadDocumentMathFont(document);
        var target = WordOmmlMathFontPolicy.PreferredFont;
        if (target is null)
        {
            // Following the document: the legacy CFF Latin Modern Math is migrated to
            // the TrueType build by the same update, before the run fonts are aligned.
            target = string.Equals(current, WordOfficeMathFontLoader.LegacyLatinModernMathFamily,
                StringComparison.OrdinalIgnoreCase)
                ? WordOfficeMathFontLoader.LatinModernMathFamily
                : current;
        }
        else if (!string.Equals(current, target, StringComparison.OrdinalIgnoreCase))
        {
            scan.CurrentMathFont = current;
        }
        scan.MathFont = target;
        if (string.IsNullOrWhiteSpace(target)) return scan;
        var count = 0;
        _ = AnyStory(document, story =>
        {
            count += AnalyzeStoryRunFonts(story, target, out _)?.EquationsToAlign ?? 0;
            return false;
        });
        scan.EquationsToAlign = count;
        return scan;
    }

    // Applies the chosen font (if any) as the document math font, then makes every
    // OMath in every story use the document math font. Caller owns the Undo record.
    private static OmmlMathFontAlignmentResult AlignOmmlMathFont(Document document)
    {
        var result = new OmmlMathFontAlignmentResult();
        if (OfficeHostEnvironment.IsWps) return result;
        var previous = ReadDocumentMathFont(document);
        WordOmmlMathFontPolicy.ApplyPreferredToDocument(document);
        var font = ReadDocumentMathFont(document);
        result.MathFont = font;
        if (!string.Equals(previous, font, StringComparison.OrdinalIgnoreCase))
            result.PreviousMathFont = previous;
        if (string.IsNullOrWhiteSpace(font)) return result;
        _ = AnyStory(document, story =>
        {
            AlignStoryRunFonts(story, font, result);
            return false;
        });
        return result;
    }

    private static WordOmmlMathFontPolicy.RunFontAlignment? AnalyzeStoryRunFonts(
        Range story,
        string mathFont,
        out int equationCount)
    {
        equationCount = 0;
        OMaths? maths = null;
        try
        {
            maths = story.OMaths;
            equationCount = maths.Count;
            if (equationCount == 0) return null;
        }
        catch (COMException) { return null; }
        finally { Release(maths); }
        string xml;
        try { xml = story.WordOpenXML ?? string.Empty; }
        catch (COMException) { return null; }
        return string.IsNullOrEmpty(xml) ? null : WordOmmlMathFontPolicy.AlignRunFonts(xml, mathFont);
    }

    private static void AlignStoryRunFonts(Range story, string mathFont, OmmlMathFontAlignmentResult result)
    {
        var plan = AnalyzeStoryRunFonts(story, mathFont, out var count);
        if (plan is null || count == 0) return;
        result.KeptNormalTextRuns += plan.KeptNormalTextRuns;
        if (plan.EquationsToAlign == 0) return;
        // The story package lists the body's top-level m:oMath elements in document
        // order, as Range.OMaths does. If the counts differ (for example several
        // m:oMath in one m:oMathPara), each OMath is analyzed on its own instead.
        var indexed = plan.EquationFonts.Count == count;
        OMaths? maths = null;
        try
        {
            maths = story.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                OMath? math = null;
                Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range;
                    var fonts = indexed && index <= plan.EquationFonts.Count
                        ? plan.EquationFonts[index - 1]
                        : EquationRunFonts(range, mathFont);
                    if (fonts.Count == 0) continue;
                    var changed = false;
                    foreach (var font in fonts)
                        changed |= ReplaceFont(range, font, mathFont);
                    if (changed) result.Equations++;
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
        }
        finally { Release(maths); }
    }

    private static IReadOnlyCollection<string> EquationRunFonts(Range equation, string mathFont)
    {
        string xml;
        try { xml = equation.WordOpenXML ?? string.Empty; }
        catch (COMException) { return Array.Empty<string>(); }
        if (string.IsNullOrEmpty(xml)) return Array.Empty<string>();
        return WordOmmlMathFontPolicy.AlignRunFonts(xml, mathFont).EquationFonts
            .SelectMany(fonts => fonts)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}

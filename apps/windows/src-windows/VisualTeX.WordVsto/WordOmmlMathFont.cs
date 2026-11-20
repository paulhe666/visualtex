using System.Runtime.InteropServices;
using System.Xml.Linq;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;
using static VisualTeX.WordVsto.WordFormulaService;
using static VisualTeX.WordVsto.WordBulkInsertion;

namespace VisualTeX.WordVsto;

/// <summary>
/// The math font of OMML: applying a document math font, aligning the run fonts of every story, and Chinese
/// typography in OMML.
/// </summary>
internal static class WordOmmlMathFont
{
    // Word exposes one native OpenType MATH face per document. This is intentionally
    // document-scoped: changing it reflows every existing OMath in the document; it is
    // not a per-formula font override and must never be emulated with m:nor/run splitting.
    // The string overload (WordFormulaService.OmmlMathFont.cs) is the OMML Font
    // menu's document-font applier.
    private static void ApplyDocumentOmmlMathFont(
        Document document,
        FormulaMetadata metadata) =>
        ApplyDocumentOmmlMathFont(document, ResolveDocumentOmmlMathFont(metadata.FormulaLetterFont));

    private static string ResolveDocumentOmmlMathFont(string? preference)
    {
        // Word native OMath has one document-wide math font and requires an
        // OpenType MATH table. VisualTeX's text-face choices (Times, Palatino,
        // Helvetica/Arial) cannot safely replace that math font, so they fall back
        // to the bundled Latin Modern Math instead of flattening individual runs.
        return (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "cambria" => "Cambria Math",
            "stix" => "STIX Two Math",
            _ => WordOfficeMathFontLoader.LatinModernMathFamily,
        };
    }

    private static string ApplyOmmlTypographyXml(
        string omml,
        double fontSizePt,
        FormulaMetadata metadata)
    {
        if (string.IsNullOrWhiteSpace(omml))
            throw new InvalidDataException("VisualTeX produced an empty OMML payload.");
        var equation = XElement.Parse(
            omml,
            LoadOptions.PreserveWhitespace);
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        ApplyBulkOmmlTypographyXml(
            equation,
            fontSizePt,
            metadata,
            word,
            math);
        return equation.ToString(SaveOptions.DisableFormatting);
    }

    internal static string ResolveOmmlChineseFont(string? preference) =>
        (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "songti" => "SimSun",
            "kaiti" => "KaiTi",
            "heiti" => "SimHei",
            _ => "Microsoft YaHei",
        };

    internal static bool IsChineseOmmlCodePoint(int codePoint) =>
        codePoint is >= 0x2E80 and <= 0x2FFF
        || codePoint is >= 0x3000 and <= 0x303F
        || codePoint is >= 0x31C0 and <= 0x31EF
        || codePoint is >= 0x3400 and <= 0x4DBF
        || codePoint is >= 0x4E00 and <= 0x9FFF
        || codePoint is >= 0xF900 and <= 0xFAFF
        || codePoint is >= 0xFF00 and <= 0xFFEF
        || codePoint is >= 0x20000 and <= 0x2FA1F;
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

    private static string ReadDocumentMathFont(Document document)
    {
        try { return document.OMathFontName ?? string.Empty; }
        catch (COMException) { return string.Empty; }
    }

    // Read-only: what AlignOmmlMathFont would change.
    internal static OmmlMathFontScan ScanOmmlMathFont(Document document)
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
    internal static OmmlMathFontAlignmentResult AlignOmmlMathFont(Document document)
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

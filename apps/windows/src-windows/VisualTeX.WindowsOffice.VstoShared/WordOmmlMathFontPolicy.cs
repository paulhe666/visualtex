using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace VisualTeX.WindowsOffice.VstoShared;

/// <summary>The per-user "OMML Font" ribbon choice of the Word add-in.</summary>
internal enum OmmlMathFontChoice
{
    /// <summary>Keep the document's own math font (VisualTeX never changes it).</summary>
    FollowDocument = 0,
    CambriaMath,
    StixTwoMath,
    LatinModernMath,
    TexGyreTermesMath,
}

/// <summary>
/// Word gives every document one Office Math font (Document.OMathFontName). When the
/// user picks a font in the OMML Font menu, every VisualTeX OMML write first makes it
/// the document math font, so the converter's existing "mathFontName =
/// document.OMathFontName" plumbing writes it. The Word add-in installs the provider at
/// startup; PowerPoint never does (it has no document math font). WPS has no document
/// math font either (it reports the placeholder "default"): there the chosen font is
/// only written into the runs of each equation VisualTeX writes.
/// </summary>
internal static class WordOmmlMathFontPolicy
{
    internal const string FollowDocumentId = "follow";
    internal const string CambriaMathId = "cambria";
    internal const string StixTwoMathId = "stix";
    internal const string LatinModernMathId = "latin-modern";
    internal const string TexGyreTermesMathId = "termes";

    internal const string CambriaMathFamily = "Cambria Math";
    internal const string StixTwoMathFamily = "STIX Two Math";
    // The bundled TrueType build; equals WordOfficeMathFontLoader.LatinModernMathFamily.
    internal const string LatinModernMathFamily = "VisualTeX Latin Modern Math";
    // The bundled TrueType build of TeX Gyre Termes Math (Times-style); equals
    // WordOfficeMathFontLoader.TexGyreTermesMathFamily.
    internal const string TexGyreTermesMathFamily = "VisualTeX TeX Gyre Termes Math";

    private const string MathNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static readonly OmmlMathFontChoice[] Choices =
    {
        OmmlMathFontChoice.FollowDocument,
        OmmlMathFontChoice.CambriaMath,
        OmmlMathFontChoice.StixTwoMath,
        OmmlMathFontChoice.TexGyreTermesMath,
        OmmlMathFontChoice.LatinModernMath,
    };

    internal static OmmlMathFontChoice Parse(string? id) =>
        (id ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            CambriaMathId => OmmlMathFontChoice.CambriaMath,
            StixTwoMathId => OmmlMathFontChoice.StixTwoMath,
            LatinModernMathId => OmmlMathFontChoice.LatinModernMath,
            TexGyreTermesMathId => OmmlMathFontChoice.TexGyreTermesMath,
            _ => OmmlMathFontChoice.FollowDocument,
        };

    internal static string Id(OmmlMathFontChoice choice) =>
        choice switch
        {
            OmmlMathFontChoice.CambriaMath => CambriaMathId,
            OmmlMathFontChoice.StixTwoMath => StixTwoMathId,
            OmmlMathFontChoice.LatinModernMath => LatinModernMathId,
            OmmlMathFontChoice.TexGyreTermesMath => TexGyreTermesMathId,
            _ => FollowDocumentId,
        };

    /// <summary>The math font family of a choice; null for "follow the document".</summary>
    internal static string? FontFamily(OmmlMathFontChoice choice) =>
        choice switch
        {
            OmmlMathFontChoice.CambriaMath => CambriaMathFamily,
            OmmlMathFontChoice.StixTwoMath => StixTwoMathFamily,
            OmmlMathFontChoice.LatinModernMath => LatinModernMathFamily,
            OmmlMathFontChoice.TexGyreTermesMath => TexGyreTermesMathFamily,
            _ => null,
        };

    /// <summary>
    /// The font every OMML write must use, or null to follow the document. A font that is
    /// not available (not installed / bundled font not loaded) is never applied silently:
    /// the ribbon disables it, and this falls back to follow.
    /// </summary>
    internal static string? EffectiveFontFamily(
        OmmlMathFontChoice choice,
        Func<string, bool> isAvailable)
    {
        var family = FontFamily(choice);
        return family is not null && isAvailable(family) ? family : null;
    }

    private static Func<string?>? _preferredFont;
    // The Word Document is passed as object so this file needs no Office interop types.
    private static Action<object, string>? _applyDocumentFont;

    /// <summary>Installed by the Word add-in; null arguments restore "follow".</summary>
    internal static void Configure(
        Func<string?>? preferredFont,
        Action<object, string>? applyDocumentFont)
    {
        _preferredFont = preferredFont;
        _applyDocumentFont = applyDocumentFont;
    }

    /// <summary>The configured font family, or null to follow the document.</summary>
    internal static string? PreferredFont
    {
        get
        {
            var provider = _preferredFont;
            if (provider is null) return null;
            string? family;
            try { family = provider(); }
            catch { return null; }
            return string.IsNullOrWhiteSpace(family) ? null : family!.Trim();
        }
    }

    /// <summary>
    /// The choke point of every OMML write into a Word document (WordOmmlConverter's
    /// Insert / ReplaceWithPreparedOmml* entries and the exact OLE-to-OMML replacement):
    /// with a configured font, makes it the document math font and returns it as the
    /// font to write; otherwise returns the caller's font unchanged.
    /// </summary>
    internal static string? ForWrite(object document, string? mathFontName)
    {
        var preferred = PreferredFont;
        if (preferred is null) return mathFontName;
        // A document that refuses the font (not loaded on this PC, a .doc in
        // compatibility mode) still takes the equation in its current font.
        try { _applyDocumentFont?.Invoke(document, preferred); }
        catch (Exception error) when (error is InvalidOperationException or System.Runtime.InteropServices.COMException
            or ArgumentException or IOException or InvalidDataException)
        {
            return mathFontName;
        }
        return preferred;
    }

    /// <summary>Applies the configured font to the document; false when following.</summary>
    internal static bool ApplyPreferredToDocument(object document)
    {
        var preferred = PreferredFont;
        if (preferred is null) return false;
        _applyDocumentFont?.Invoke(document, preferred);
        return true;
    }

    /// <summary>
    /// A font family with an OpenType MATH table by naming convention ("Cambria Math",
    /// "STIX Two Math", "VisualTeX Latin Modern Math", "VisualTeX TeX Gyre Termes Math",
    /// "XITS Math" ...).
    /// </summary>
    internal static bool IsMathFontFamily(string? family) =>
        !string.IsNullOrWhiteSpace(family)
        && family!.Trim().EndsWith(" Math", StringComparison.OrdinalIgnoreCase);

    internal sealed class RunFontAlignment
    {
        internal RunFontAlignment(
            string xml,
            IReadOnlyList<IReadOnlyCollection<string>> equationFonts,
            int rewrittenRuns,
            int keptNormalTextRuns)
        {
            Xml = xml;
            EquationFonts = equationFonts;
            RewrittenRuns = rewrittenRuns;
            KeptNormalTextRuns = keptNormalTextRuns;
        }

        /// <summary>The input with the run fonts aligned.</summary>
        internal string Xml { get; }

        /// <summary>
        /// Per top-level m:oMath in document order: the explicit run/control font names
        /// replaced by the math font (empty: that equation already follows it).
        /// </summary>
        internal IReadOnlyList<IReadOnlyCollection<string>> EquationFonts { get; }

        /// <summary>Math runs and control properties whose fonts were rewritten.</summary>
        internal int RewrittenRuns { get; }

        /// <summary>m:nor runs that keep their own (text) font.</summary>
        internal int KeptNormalTextRuns { get; }

        internal int EquationsToAlign => EquationFonts.Count(fonts => fonts.Count > 0);
    }

    /// <summary>
    /// Makes every math run of every m:oMath in <paramref name="xml"/> (bare OMML or a
    /// Range.WordOpenXML package: only the document body, text boxes excluded since
    /// they are separate stories) follow <paramref name="mathFont"/>: explicit
    /// w:rFonts ascii/hAnsi on non-m:nor runs and on m:ctrlPr are rewritten to it.
    /// m:nor runs (\mathrm, \text, number prefixes) keep their text font: a font also
    /// used by m:nor text in the same equation is left alone unless it is itself a math
    /// font (the normal text then only followed the old math font, as Word writes it).
    /// The rewritten font names per equation drive the live Word rewrite, which replaces
    /// exactly those fonts inside each OMath range.
    /// </summary>
    internal static RunFontAlignment AlignRunFonts(string xml, string mathFont)
    {
        if (xml is null) throw new ArgumentNullException(nameof(xml));
        if (string.IsNullOrWhiteSpace(mathFont)) throw new ArgumentException("A math font is required.", nameof(mathFont));
        mathFont = mathFont.Trim();
        XNamespace math = MathNamespace;
        XNamespace word = WordNamespace;
        var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        var body = document.Descendants(word + "body").FirstOrDefault();
        var scope = body ?? document.Root;
        var equations = scope is null
            ? new List<XElement>()
            : scope.DescendantsAndSelf(math + "oMath")
                .Where(element => !element.Ancestors(math + "oMath").Any()
                    && !element.Ancestors(word + "txbxContent").Any())
                .ToList();

        var perEquation = new List<IReadOnlyCollection<string>>(equations.Count);
        var rewrittenRuns = 0;
        var keptNormalTextRuns = 0;
        var fontAttributes = new[] { word + "ascii", word + "hAnsi" };

        foreach (var equation in equations)
        {
            var mathTargets = new List<XElement>();
            var normalTargets = new List<XElement>();
            foreach (var run in equation.Descendants(math + "r"))
            {
                var fonts = run.Element(word + "rPr")?.Element(word + "rFonts");
                if (fonts is null) continue;
                (IsNormalText(run, math) ? normalTargets : mathTargets).Add(fonts);
            }
            foreach (var control in equation.Descendants(math + "ctrlPr"))
            {
                var fonts = control.Element(word + "rPr")?.Element(word + "rFonts");
                if (fonts is not null) mathTargets.Add(fonts);
            }

            var normalFonts = new HashSet<string>(
                normalTargets.SelectMany(fonts => FontNames(fonts, fontAttributes)),
                StringComparer.OrdinalIgnoreCase);
            var replace = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var font in mathTargets.SelectMany(fonts => FontNames(fonts, fontAttributes)))
            {
                if (string.Equals(font, mathFont, StringComparison.OrdinalIgnoreCase)) continue;
                if (normalFonts.Contains(font) && !IsMathFontFamily(font)) continue;
                replace.Add(font);
            }

            foreach (var fonts in mathTargets.Concat(normalTargets))
            {
                var changed = false;
                foreach (var attribute in fontAttributes)
                {
                    var value = ((string?)fonts.Attribute(attribute))?.Trim();
                    if (string.IsNullOrEmpty(value) || !replace.Contains(value!)) continue;
                    fonts.SetAttributeValue(attribute, mathFont);
                    changed = true;
                }
                if (changed) rewrittenRuns++;
            }
            keptNormalTextRuns += normalTargets.Count(fonts =>
                FontNames(fonts, fontAttributes).Any(font =>
                    !string.Equals(font, mathFont, StringComparison.OrdinalIgnoreCase)));
            perEquation.Add(replace.OrderBy(font => font, StringComparer.OrdinalIgnoreCase).ToArray());
        }

        var output = document.Declaration is null
            ? document.Root?.ToString(SaveOptions.DisableFormatting) ?? xml
            : document.Declaration + document.Root?.ToString(SaveOptions.DisableFormatting);
        return new RunFontAlignment(output, perEquation, rewrittenRuns, keptNormalTextRuns);
    }

    private static IEnumerable<string> FontNames(XElement fonts, IEnumerable<XName> attributes) =>
        attributes
            .Select(attribute => ((string?)fonts.Attribute(attribute))?.Trim())
            .Where(value => !string.IsNullOrEmpty(value))
            .Select(value => value!);

    private static bool IsNormalText(XElement run, XNamespace math)
    {
        var normal = run.Element(math + "rPr")?.Element(math + "nor");
        if (normal is null) return false;
        var value = ((string?)normal.Attribute(math + "val"))?.Trim().ToLowerInvariant();
        return value is null or "" or "1" or "on" or "true";
    }
}

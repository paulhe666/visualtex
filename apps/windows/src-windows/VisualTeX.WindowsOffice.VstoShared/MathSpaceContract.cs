using System.Text.RegularExpressions;

namespace VisualTeX.WordVsto;

// One source-width contract for MathML recovery and target-specific spacing.
// Numeric equality is lexical decimal normalization, never floating-point or
// decimal rounding. In particular, a near-1em value must not become native quad.
internal static class MathSpaceContract
{
    private static readonly Regex DimensionPattern = new(
        @"\A(?<number>[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+))\s*(?<unit>em|ex|pt|pc|px|in|cm|mm)\z",
        RegexOptions.CultureInvariant);

    // Only native widths measured using native WMF character origins belong
    // here. The table is shared by native read/write and LaTeX source recovery.
    private static readonly MathSpaceNativeIdentity[] NativeSpaces =
    {
        new(0xEF05, "1em", @"\quad ", "Native WMF character-origin spacing verified for MTCode EF05."),
        new(0xEF06, "2em", @"\qquad ", "Native WMF character-origin spacing verified for MTCode EF06."),
        // MathType's own spaces for TeX's thin, medium, thick and negative thin
        // space: what MathType 7.9's TeX and MathML importers store for \, \: \;
        // and \! (the MathML importer uses EF03 for \; as well). MathType's
        // renderer gives them their width, which depends on the neighbours.
        new(0xEF07, "0.167em", @"\,", "MathType 7.9 TeX and MathML import store \\, (0.167em) as MTCode EF07."),
        new(0xEF03, "0.222em", @"\:", "MathType 7.9 TeX and MathML import store \\: (0.222em) as MTCode EF03."),
        new(0xEF04, "0.278em", @"\;", "MathType 7.9 TeX import stores \\; (0.278em) as MTCode EF04."),
        new(0xEF01, "-0.167em", @"\!", "MathType 7.9 TeX and MathML import store \\! (-0.167em) as MTCode EF01."),
    };


    // OfficeMath represents supported explicit gaps as mathematical Unicode
    // spaces in m:t, not as an arbitrary-width spacing object. This table is
    // shared by the forward projection and reverse decoding. It is separate
    // from measured MTEF identities: evidence for one target is not evidence for
    // the other. Native host persistence/layout is checked by the Word oracle.
    // https://learn.microsoft.com/en-us/office/math/latex#spacing-commands
    // https://www.unicode.org/notes/tn28/UTN28-PlainTextMath-v3.3.pdf section 3.16
    private static readonly MathSpaceOmmlIdentity[] OmmlSpaces =
    {
        new("1em", "\u2003"),
        new("2em", "\u2003\u2003"),
    };

    internal static int MaximumOmmlSpaceTextLength =>
        OmmlSpaces.Max(space => space.Text.Length);

    internal static MathSpaceOmmlIdentity? FindOmmlSpaceText(string text) =>
        OmmlSpaces.FirstOrDefault(space =>
            string.Equals(space.Text, text, StringComparison.Ordinal));

    internal static MathSpaceOmmlCapability DescribeOmmlSpace(
        string? width, string? height = null, string? depth = null,
        string? linebreak = null)
    {
        var descriptor = ParseWidth(width);
        var identity = OmmlSpaces.FirstOrDefault(space =>
            string.Equals(space.MathMlWidth, descriptor.CanonicalWidth, StringComparison.Ordinal));
        var zeroWidth = HaveSameWidth(width, "0em");
        var heightSupported = HaveSameWidth(height, "0em");
        var depthSupported = HaveSameWidth(depth, "0em");
        var linebreakSupported = string.IsNullOrWhiteSpace(linebreak)
            || string.Equals(linebreak!.Trim(), "auto", StringComparison.Ordinal);
        var limits = new List<string>();
        if (identity is null && !zeroWidth)
            limits.Add("OMML has no verified exact mapping for width '" + descriptor.SourceWidth
                + "'; the width is not rounded to another Unicode space.");
        if (!heightSupported)
            limits.Add("OMML explicit-space height is not represented: '" + height + "'.");
        if (!depthSupported)
            limits.Add("OMML explicit-space depth is not represented: '" + depth + "'.");
        if (!linebreakSupported)
            limits.Add("OMML explicit-space linebreak is not represented: '" + linebreak + "'.");
        return new MathSpaceOmmlCapability(descriptor, identity?.Text ?? (zeroWidth ? "" : null),
            heightSupported, depthSupported, linebreakSupported, limits);
    }

    internal static MathSpaceNativeIdentity? FindNativeSpace(int mtCode) =>
        NativeSpaces.FirstOrDefault(space => space.MtCode == mtCode);

    internal static MathSpaceDescriptor ParseWidth(string? width)
    {
        var source = (width ?? string.Empty).Trim();
        // The MathML default is zero, including a linebreak-only mspace.
        if (source.Length == 0) source = "0em";
        var match = DimensionPattern.Match(source);
        if (match.Success)
        {
            var number = match.Groups["number"].Value;
            var unit = match.Groups["unit"].Value;
            var canonical = NormalizeDecimal(number) + unit;
            var native = NativeSpaces.FirstOrDefault(space => space.MathMlWidth == canonical);
            // Preserve the original decimal digits and unit. Do not convert mu
            // or absolute units via em: MathJax rounds converted mu to 3 decimals.
            var dimension = number + unit;
            return new MathSpaceDescriptor(source, canonical,
                native?.LatexSource ?? @"\hspace{" + dimension + "}",
                native, MathSpaceSourceSupport.ExactDimension, null);
        }

        var symbolic = source.StartsWith("negative", StringComparison.Ordinal)
            ? source.Substring("negative".Length) : source;
        if (symbolic is "veryverythinmathspace" or "verythinmathspace"
            or "thinmathspace" or "mediummathspace" or "thickmathspace"
            or "verythickmathspace" or "veryverythickmathspace")
        {
            // These are exact rational eighteenth-em lengths, not the rounded
            // decimal emitted by the current MathJax mu-command export path.
            return new MathSpaceDescriptor(source, source, null, null,
                MathSpaceSourceSupport.SymbolicMathSpace,
                "Named MathML spaces need an exact symbolic roundtrip; current mu export rounds to decimal em.");
        }
        if (source.EndsWith("%", StringComparison.Ordinal))
            return new MathSpaceDescriptor(source, null, null, null,
                MathSpaceSourceSupport.NeedsContext,
                "Percentage spacing requires its MathML reference-width context.");
        return new MathSpaceDescriptor(source, null, null, null,
            MathSpaceSourceSupport.NotRepresented,
            "This width has no verified exact source representation in the current editor/export path.");
    }

    // A separate width-fidelity assertion: semantic token comparisons may
    // intentionally ignore spaces, but that cannot prove spacing fidelity.
    internal static bool HaveSameWidth(string? first, string? second)
    {
        var left = ParseWidth(first);
        var right = ParseWidth(second);
        if (left.SourceSupported && right.SourceSupported
            && left.CanonicalWidth!.Substring(0, left.CanonicalWidth.Length - 2) == "0"
            && right.CanonicalWidth!.Substring(0, right.CanonicalWidth.Length - 2) == "0")
            return true;
        return left.CanonicalWidth is not null && right.CanonicalWidth is not null
            && string.Equals(left.CanonicalWidth, right.CanonicalWidth, StringComparison.Ordinal);
    }

    private static string NormalizeDecimal(string value)
    {
        var negative = value[0] == '-';
        if (negative || value[0] == '+') value = value.Substring(1);
        var dot = value.IndexOf('.');
        var whole = (dot < 0 ? value : value.Substring(0, dot)).TrimStart('0');
        var fraction = dot < 0 ? string.Empty : value.Substring(dot + 1).TrimEnd('0');
        if (whole.Length == 0) whole = "0";
        var result = whole + (fraction.Length == 0 ? string.Empty : "." + fraction);
        return negative && result != "0" ? "-" + result : result;
    }
}

internal enum MathSpaceSourceSupport
{
    ExactDimension,
    SymbolicMathSpace,
    NeedsContext,
    NotRepresented,
}

internal sealed class MathSpaceDescriptor
{
    internal MathSpaceDescriptor(string sourceWidth, string? canonicalWidth,
        string? latexSource, MathSpaceNativeIdentity? nativeSpace,
        MathSpaceSourceSupport sourceSupport, string? unsupportedReason)
    {
        SourceWidth = sourceWidth;
        CanonicalWidth = canonicalWidth;
        LatexSource = latexSource;
        NativeSpace = nativeSpace;
        SourceSupport = sourceSupport;
        UnsupportedReason = unsupportedReason;
    }

    internal string SourceWidth { get; }
    internal string? CanonicalWidth { get; }
    internal string? LatexSource { get; }
    internal MathSpaceNativeIdentity? NativeSpace { get; }
    internal MathSpaceSourceSupport SourceSupport { get; }
    internal bool SourceSupported => SourceSupport == MathSpaceSourceSupport.ExactDimension;
    internal string? UnsupportedReason { get; }
    internal string NativeSupport => NativeSpace?.SupportEvidence
        ?? "Native MTEF width fidelity is uncharacterized for this source width.";
}

internal sealed class MathSpaceNativeIdentity
{
    internal MathSpaceNativeIdentity(int mtCode, string mathMlWidth, string latexSource, string supportEvidence)
    {
        MtCode = mtCode;
        MathMlWidth = mathMlWidth;
        LatexSource = latexSource;
        SupportEvidence = supportEvidence;
    }

    internal int MtCode { get; }
    internal string MathMlWidth { get; }
    internal string LatexSource { get; }
    internal string SupportEvidence { get; }
}

// CanProject describes the exact, documented projection subset. It does not
// claim that every host version has passed persistence and width measurements.
internal sealed class MathSpaceOmmlCapability
{
    internal MathSpaceOmmlCapability(MathSpaceDescriptor width, string? text,
        bool heightSupported, bool depthSupported, bool linebreakSupported,
        IReadOnlyList<string> auditLimits)
    {
        Width = width;
        Text = text;
        HeightSupported = heightSupported;
        DepthSupported = depthSupported;
        LinebreakSupported = linebreakSupported;
        AuditLimits = auditLimits;
    }

    internal MathSpaceDescriptor Width { get; }
    internal string? Text { get; }
    internal bool WidthSupported => Text is not null;
    internal bool HeightSupported { get; }
    internal bool DepthSupported { get; }
    internal bool LinebreakSupported { get; }
    internal IReadOnlyList<string> AuditLimits { get; }
    internal bool CanProject => AuditLimits.Count == 0;
}

internal sealed class MathSpaceOmmlIdentity
{
    internal MathSpaceOmmlIdentity(string mathMlWidth, string text)
    {
        MathMlWidth = mathMlWidth;
        Text = text;
    }

    internal string MathMlWidth { get; }
    internal string Text { get; }
}

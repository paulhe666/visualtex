using System.Globalization;
using System.Text.RegularExpressions;

namespace VisualTeX.WordVsto;

internal sealed partial class WordNativeNumberingAdoptionPlan
{
    internal sealed class NumberStyle
    {
        internal string Prefix = string.Empty;
        internal string Separator = string.Empty;
        internal int Ordinal;
        internal int HeadingLevel;
        internal string SequenceName = string.Empty;
        internal bool AutomaticHeading;
        internal string FormatId => HeadingLevel == 0 ? EquationNumberFormat.ContinuousId
            : HeadingLevel == 1 ? Separator == "." ? EquationNumberFormat.Heading1DotId : EquationNumberFormat.Heading1DashId
            : Separator == "." ? EquationNumberFormat.Heading2DotId : EquationNumberFormat.Heading2DashId;
        internal NumberStyle Copy() => (NumberStyle)MemberwiseClone();
    }

    internal static bool TryParseLabel(string text, out NumberStyle style)
    {
        style = new NumberStyle();
        // The structural #( ... ) proof happens separately. This parser accepts
        // numbering labels, not arbitrary mathematical delimiter expressions.
        var match = Regex.Match(text ?? string.Empty,
            @"^(?:(?<prefix>(?:[0-9]+|[A-Za-z]+)(?:\.[0-9]+)*)(?<separator>[.\-\u2010-\u2013\u2212]))?(?<ordinal>[0-9]+)$",
            RegexOptions.CultureInvariant);
        if (!match.Success || !int.TryParse(match.Groups["ordinal"].Value,
                NumberStyles.None, CultureInfo.InvariantCulture, out var ordinal) || ordinal < 1)
            return false;
        // A leading-zero picture is a distinct native number format, not ARABIC.
        if (match.Groups["ordinal"].Value.Length > 1 && match.Groups["ordinal"].Value[0] == '0') return false;
        style.Prefix = match.Groups["prefix"].Value;
        style.Separator = match.Groups["separator"].Value;
        style.Ordinal = ordinal;
        style.HeadingLevel = style.Prefix.Length == 0 ? 0 : style.Prefix.Contains(".") ? 2 : 1;
        return true;
    }

}

using System.Text;
using System.Xml.Linq;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// Where WPS Presentation's native equations are in a text frame. WPS has no
/// MathZones (it returns a placeholder range); its text spells every equation
/// inline: the math characters as WPS stores them (an italic x is U+1D465),
/// each structure (fraction, script, radical, delimiter, matrix, ...) opened by
/// U+001A and each of its arguments closed by U+001B, in the order the
/// arguments appear in the OMML. Spelling WPS's own saved XML of the shape the
/// same way gives each equation's exact character range in the live text.
/// </summary>
internal static class WpsMathText
{
    internal const char StructureStart = '\u001A';
    internal const char ArgumentEnd = '\u001B';

    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XName Zone = XName.Get("m", "http://schemas.microsoft.com/office/drawing/2010/main");

    private static readonly HashSet<string> Structures = new(StringComparer.Ordinal)
    {
        "acc", "bar", "borderBox", "box", "d", "eqArr", "f", "func", "groupChr", "limLow",
        "limUpp", "m", "nary", "phant", "rad", "sPre", "sSub", "sSubSup", "sSup",
    };

    /// <summary>
    /// False when the text certainly holds no structure or math letter. Only a
    /// cheap filter before a saved copy is read (an equation of digits and
    /// operators alone passes it unnoticed).
    /// </summary>
    internal static bool MayContainMath(string text) =>
        text.IndexOf(StructureStart) >= 0 || text.IndexOf('\uD835') >= 0 || text.IndexOf('\u210E') >= 0;

    /// <summary>
    /// The 1-based start and length of each native equation (a14:m) of a saved
    /// text shape, in document order; null when the saved shape does not spell
    /// <paramref name="liveText"/> exactly, so no position could be trusted.
    /// </summary>
    internal static IReadOnlyList<(int Start, int Length)>? Locate(XElement shape, string liveText)
    {
        var zones = new List<(int Start, int Length)>();
        var text = new StringBuilder(liveText.Length);
        var paragraphs = shape.Element(P + "txBody")?.Elements(A + "p").ToArray() ?? Array.Empty<XElement>();
        for (var index = 0; index < paragraphs.Length; index++)
        {
            if (index > 0) text.Append('\r');
            foreach (var child in LiveChildren(paragraphs[index]))
            {
                // WPS saves the paragraph break before a display equation as a CR
                // inside the preceding run (Before text.&#13;); serialized XML
                // reads it back as a line feed.
                if (child.Name == A + "r" || child.Name == A + "fld")
                    text.Append(((string?)child.Element(A + "t") ?? string.Empty).Replace("\r\n", "\r").Replace('\n', '\r'));
                else if (child.Name == A + "br")
                    text.Append('\v');
                else if (child.Name == Zone)
                {
                    var start = text.Length;
                    SpellChildren(child, text);
                    zones.Add((start + 1, text.Length - start));
                }
            }
        }
        return string.Equals(text.ToString(), liveText, StringComparison.Ordinal) ? zones : null;
    }

    /// <summary>The live text WPS shows for one a14:m equation.</summary>
    internal static string Spell(XElement zone)
    {
        var text = new StringBuilder();
        SpellChildren(zone, text);
        return text.ToString();
    }

    // The zone of a selection: the one equation that holds the caret, or that a
    // selection overlaps. -1 when there is no such single equation.
    internal static int ZoneAt(IReadOnlyList<(int Start, int Length)> zones, int selectedStart, int selectedLength)
    {
        var found = -1;
        for (var index = 0; index < zones.Count; index++)
        {
            var (start, length) = zones[index];
            var hit = selectedLength == 0
                ? selectedStart >= start && selectedStart < start + length
                : selectedStart < start + length && selectedStart + selectedLength > start;
            if (!hit) continue;
            if (found >= 0) return -1;
            found = index;
        }
        return found;
    }

    private static void Spell(XElement element, StringBuilder text)
    {
        if (element.Name.Namespace != M) return;
        var name = element.Name.LocalName;
        if (name.EndsWith("Pr", StringComparison.Ordinal)) return;
        if (name == "r")
        {
            foreach (var run in element.Elements(M + "t")) text.Append(run.Value);
            return;
        }
        if (!Structures.Contains(name))
        {
            SpellChildren(element, text);
            return;
        }
        text.Append(StructureStart);
        foreach (var argument in element.Elements())
        {
            if (argument.Name.Namespace != M || argument.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)) continue;
            // A matrix row adds no mark of its own: each cell is an argument.
            foreach (var cell in argument.Name == M + "mr" ? argument.Elements(M + "e") : new[] { argument })
            {
                SpellChildren(cell, text);
                text.Append(ArgumentEnd);
            }
        }
    }

    private static void SpellChildren(XElement element, StringBuilder text)
    {
        foreach (var child in LiveChildren(element)) Spell(child, text);
    }

    // Child elements, an mc:AlternateContent replaced by its first choice (the
    // branch WPS and PowerPoint show).
    private static IEnumerable<XElement> LiveChildren(XElement element)
    {
        foreach (var child in element.Elements())
        {
            if (child.Name != Mc + "AlternateContent")
            {
                yield return child;
                continue;
            }
            var live = child.Elements(Mc + "Choice").FirstOrDefault() ?? child.Element(Mc + "Fallback");
            if (live is null) continue;
            foreach (var nested in live.Elements()) yield return nested;
        }
    }
}

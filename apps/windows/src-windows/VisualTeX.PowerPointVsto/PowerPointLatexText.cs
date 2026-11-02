using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.PowerPointVsto;

// One complete LaTeX formula in the text of a PowerPoint (or WPS) text box: the
// 1-based character range of its code, delimiters included, as the text
// range's Characters(start, length) counts it (UTF-16 units).
internal sealed class PowerPointLatexSpan
{
    internal PowerPointLatexSpan(int start, int length, string latex, bool display)
    {
        Start = start;
        Length = length;
        Latex = latex;
        Display = display;
    }

    internal int Start { get; }
    internal int Length { get; }
    // 1-based, exclusive.
    internal int End => Start + Length;
    internal string Latex { get; }
    // $$...$$, \[...\] and display environments; $...$ and \(...\) are inline.
    internal bool Display { get; }
    internal string DisplayMode => Display ? "block" : "inline";
}

// Pure text rules of LaTeX redraw in PowerPoint and WPS text boxes and of turning
// formulas back into LaTeX code. Nothing here touches COM; the unit tests share it.
internal static class PowerPointLatexText
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private static readonly XName MathZone = XName.Get("m", "http://schemas.microsoft.com/office/drawing/2010/main");

    // What the scanner reads in place of a native equation's characters: Word's
    // table-cell mark, which ends every delimiter search of the shared scanner.
    // So nothing an equation spells is read as LaTeX, and no LaTeX formula is
    // read across an equation.
    internal const char EquationMask = '\a';

    // A cheap filter before any LaTeX is looked for: every delimiter the
    // scanner knows starts with "$" or a backslash.
    internal static bool MayContainLatex(string? text) =>
        !string.IsNullOrEmpty(text) && (text!.IndexOf('$') >= 0 || text.IndexOf('\\') >= 0);

    // The text the scanner reads. PowerPoint and WPS end a paragraph with U+000D
    // and spell a line break inside one (Shift+Enter) as U+000B; the scanner sees
    // that line break as a new line, as TeX sees a source line break (inline math
    // ends there, a blank line ends display math). Native equations (1-based
    // ranges) are masked. Every character keeps its position.
    internal static string ScanText(string? text, IEnumerable<(int Start, int Length)>? equations)
    {
        var characters = (text ?? string.Empty).ToCharArray();
        for (var index = 0; index < characters.Length; index++)
            if (characters[index] == '\v') characters[index] = '\n';
        foreach (var (start, length) in equations ?? Array.Empty<(int Start, int Length)>())
        {
            var from = Math.Max(0, start - 1);
            var to = Math.Min(characters.Length, start - 1 + Math.Max(0, length));
            for (var index = from; index < to; index++) characters[index] = EquationMask;
        }
        return new string(characters);
    }

    // Every complete LaTeX formula of a text, in text order, with the delimiter
    // rules of Word's LaTeX redraw (WordBulkImportParser.FindFormulaSpans).
    internal static List<PowerPointLatexSpan> FindFormulas(string? text, IReadOnlyList<(int Start, int Length)>? equations) =>
        WordBulkImportParser.FindFormulaSpans(ScanText(text, equations))
            .Select(span => new PowerPointLatexSpan(
                span.Start + 1,
                span.Length,
                span.Latex,
                string.Equals(span.DisplayMode, "block", StringComparison.Ordinal)))
            .OrderBy(span => span.Start)
            .ToList();

    // A text selection converts only the formulas whose whole code it holds.
    internal static bool Inside(PowerPointLatexSpan span, int selectionStart, int selectionLength) =>
        span.Start >= selectionStart && span.End <= selectionStart + selectionLength;

    internal static bool Overlaps(int start, int length, int otherStart, int otherLength) =>
        start < otherStart + otherLength && otherStart < start + length;

    // The text each formula's conversion replaces, in the order of the spans
    // (which must be in text order). An inline formula replaces its code. A
    // display formula becomes a paragraph of its own, so it also takes the
    // spaces next to its code and one line break (U+000B) on each side: the
    // paragraph breaks then stand where they were, instead of an empty line or
    // a stray space. Neighbouring ranges never overlap, and none reaches into a
    // native equation.
    internal static List<(int Start, int Length)> ReplacedRanges(
        string? text,
        IReadOnlyList<(int Start, int Length)>? equations,
        IReadOnlyList<PowerPointLatexSpan> spans)
    {
        var scan = ScanText(text, equations);
        var result = new List<(int Start, int Length)>(spans.Count);
        foreach (var span in spans)
        {
            if (!span.Display)
            {
                result.Add((span.Start, span.Length));
                continue;
            }
            var start = ExtendLeft(scan, span.Start - 1);
            var end = ExtendRight(scan, span.Start - 1 + span.Length);
            result.Add((start + 1, end - start));
        }
        for (var index = 0; index + 1 < result.Count; index++)
        {
            var (start, length) = result[index];
            var next = result[index + 1].Start;
            if (start + length > next)
                result[index] = (start, Math.Max(spans[index].End, next) - start);
        }
        return result;
    }

    private static int ExtendLeft(string scan, int start)
    {
        var index = SkipBlanksLeft(scan, start);
        return index > 0 && scan[index - 1] == '\n' ? SkipBlanksLeft(scan, index - 1) : index;
    }

    private static int ExtendRight(string scan, int end)
    {
        var index = SkipBlanksRight(scan, end);
        return index < scan.Length && scan[index] == '\n' ? SkipBlanksRight(scan, index + 1) : index;
    }

    private static int SkipBlanksLeft(string scan, int index)
    {
        while (index > 0 && scan[index - 1] is ' ' or '\t') index--;
        return index;
    }

    private static int SkipBlanksRight(string scan, int index)
    {
        while (index < scan.Length && scan[index] is ' ' or '\t') index++;
        return index;
    }

    // Whether a text box holds exactly one LaTeX formula and nothing else but
    // white space: only such a box is replaced by a VisualTeX or MathType object.
    // A native equation, a field or any other character makes it mixed.
    internal static bool IsSingleFormula(
        string? text,
        IReadOnlyList<(int Start, int Length)>? equations,
        IReadOnlyList<PowerPointLatexSpan> spans)
    {
        if (spans.Count != 1 || equations is { Count: > 0 }) return false;
        var value = text ?? string.Empty;
        var span = spans[0];
        for (var index = 0; index < value.Length; index++)
        {
            if (index >= span.Start - 1 && index < span.End - 1) continue;
            if (!IsWhiteSpace(value[index])) return false;
        }
        return true;
    }

    // What holds a text, by its Office shape type (MsoShapeType).
    internal enum TextHost
    {
        TextBox,
        Placeholder,
        AutoShape,
        Other,
    }

    internal static TextHost HostOf(int msoShapeType) => msoShapeType switch
    {
        17 => TextHost.TextBox,     // msoTextBox
        14 => TextHost.Placeholder, // msoPlaceholder
        1 => TextHost.AutoShape,    // msoAutoShape
        _ => TextHost.Other,
    };

    // Only a text box is replaced by a formula object: a text box, a placeholder,
    // or an autoshape that shows nothing but its text (no visible fill or line).
    // A drawn shape (a filled circle labelled $v_1$) keeps its LaTeX code;
    // LaTeX -> OMML still works in it.
    internal static bool CanBecomeObject(TextHost host, bool fillVisible, bool lineVisible) =>
        host is TextHost.TextBox or TextHost.Placeholder
        || host == TextHost.AutoShape && !fillVisible && !lineVisible;

    // The rotation that shows a formula the way a text box with this rotation
    // and flip shows its text: PowerPoint never mirrors text (a horizontal flip
    // changes nothing) and turns it upside down in a vertically flipped box.
    internal static float TextRotation(float rotation, bool flipVertical) =>
        flipVertical ? (rotation + 180f) % 360f : rotation;

    // White space, plus the zero-width space, word joiner and byte order mark.
    internal static bool IsWhiteSpace(char character) =>
        char.IsWhiteSpace(character) || IsZeroWidth(character);

    // Spaces inside a line (no line or paragraph break).
    private static bool IsBlank(char character) =>
        character is ' ' or '\t' or (char)0x00A0 or (char)0x3000 || IsZeroWidth(character);

    private static bool IsZeroWidth(char character) =>
        character is (char)0x200B or (char)0x2060 or (char)0xFEFF;

    // A native equation alone in its paragraph (nothing but spaces around it)
    // is display math and becomes "$$...$$"; any other one "$...$". A line break
    // (U+000B) next to it means the paragraph has other lines: inline.
    internal static bool IsAloneInParagraph(string? text, int start, int length)
    {
        var value = text ?? string.Empty;
        for (var index = Math.Min(value.Length, start - 1) - 1; index >= 0 && value[index] != '\r'; index--)
            if (!IsBlank(value[index])) return false;
        for (var index = Math.Max(0, start - 1 + length); index < value.Length && value[index] != '\r'; index++)
            if (!IsBlank(value[index])) return false;
        return true;
    }

    // The LaTeX code that replaces a formula: "$...$" for inline math and
    // "$$...$$" for display math (an equation alone in its paragraph, and every
    // formula object). Display code keeps its line breaks as PowerPoint line
    // breaks (U+000B) but no empty line, which would end display math; inline
    // code is one line. An equation tag is kept for display math only, as in
    // Word's conversion to LaTeX code.
    internal static string BuildSource(string? latex, bool display, string? equationTag = null)
    {
        var body = (latex ?? string.Empty)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace('\v', '\n')
            .Trim();
        if (display)
        {
            body = FormulaEquationTag.Attach(body, equationTag);
            var lines = body.Split('\n').Select(line => line.Trim()).Where(line => line.Length > 0).ToList();
            if (lines.Count == 0) throw EmptySource();
            return "$$" + string.Join("\v", lines) + "$$";
        }
        body = Regex.Replace(FormulaEquationTag.Extract(body).Latex, @"[ \t]*\n[ \t\n]*", " ").Trim();
        if (body.Length == 0) throw EmptySource();
        return "$" + body + "$";
    }

    private static Exception EmptySource() =>
        new InvalidDataException(OfficePluginLanguage.Text(
            "该公式没有可写出的 LaTeX 源码。",
            "The equation has no LaTeX source to write."));

    // LaTeX code shown in a color other than black keeps that color when it
    // becomes a VisualTeX or MathType object, which draw uncolored parts black
    // (on a dark slide the code is white, and a black formula would vanish).
    internal static string WithTextColor(string latex, string? color) =>
        PowerPointConversionColors.NeedsExplicit(color, PowerPointConversionColors.Black)
            ? @"\textcolor{" + color!.ToUpperInvariant() + "}{" + latex + "}"
            : latex;

    // Where each replacement stands once all replacements of one text have been
    // written (from last to first, so each one's position is still the scanned
    // one when it is written): every earlier replacement moves the later text by
    // its change of length. Input and output in text order.
    internal static List<(int Start, int Length)> FinalRanges(
        IReadOnlyList<(int Start, int Length, int NewLength)> replacements)
    {
        var result = new List<(int Start, int Length)>(replacements.Count);
        var shift = 0;
        foreach (var (start, length, newLength) in replacements)
        {
            result.Add((start + shift, newLength));
            shift += newLength - length;
        }
        return result;
    }

    // The character of ordinary text whose format LaTeX code takes when it
    // replaces the native equation at [start, start + length): the nearest one
    // of the equation's paragraph (the one before it first), else the nearest in
    // the whole text. 0 when the text has none (an equation-only box).
    internal static int FormatDonor(string? text, IReadOnlyList<(int Start, int Length)>? equations, int start, int length)
    {
        var value = text ?? string.Empty;
        var zones = equations ?? Array.Empty<(int Start, int Length)>();
        bool Ordinary(int position) =>
            position >= 1 && position <= value.Length
            && value[position - 1] is not ('\r' or '\v')
            && !zones.Any(zone => position >= zone.Start && position < zone.Start + zone.Length);

        for (var position = Math.Min(start - 1, value.Length); position >= 1 && value[position - 1] != '\r'; position--)
            if (Ordinary(position)) return position;
        for (var position = start + length; position <= value.Length && value[position - 1] != '\r'; position++)
            if (Ordinary(position)) return position;
        for (var distance = 1; distance <= value.Length; distance++)
        {
            if (Ordinary(start - distance)) return start - distance;
            if (Ordinary(start + length - 1 + distance)) return start + length - 1 + distance;
        }
        return 0;
    }

    // A text box's font size as text: PowerPoint accepts 1-4000 pt.
    internal static float TextFontSize(double formulaFontSizePt)
    {
        if (!(formulaFontSizePt > 0) || double.IsInfinity(formulaFontSizePt)) return 18f;
        return (float)Math.Round(Math.Min(4000, Math.Max(1, formulaFontSizePt)), 1);
    }

    // PowerPoint's "shrink text on overflow" draws a text body smaller than its
    // runs' sizes, which Font.Size reports unscaled; the scale is saved as
    // bodyPr/normAutofit/@fontScale (1/1000 %, or "62.5%"). 1 when not shrunk.
    internal static double AutofitFontScale(XElement? textShape)
    {
        var value = (string?)textShape?.Element(P + "txBody")?.Element(A + "bodyPr")
            ?.Element(A + "normAutofit")?.Attribute("fontScale");
        if (string.IsNullOrWhiteSpace(value)) return 1;
        double scale;
        var trimmed = value!.Trim();
        if (trimmed.EndsWith("%", StringComparison.Ordinal))
        {
            if (!double.TryParse(trimmed.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var percent))
                return 1;
            scale = percent / 100;
        }
        else
        {
            if (!double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out var thousandths))
                return 1;
            scale = thousandths / 100000;
        }
        return scale > 0 && scale <= 1 ? scale : 1;
    }

    // The LaTeX formulas in the saved text bodies below an element (the text
    // shapes of a group, the cells of a table), reported as left alone.
    internal static int CountFormulasInTextBodies(XElement element)
    {
        var count = 0;
        foreach (var body in PowerPointSlideXml.LiveDescendants(element)
                     .Where(candidate => candidate.Name == P + "txBody" || candidate.Name == A + "txBody"))
        {
            var text = SavedText(body);
            if (!MayContainLatex(text)) continue;
            try { count += WordBulkImportParser.FindFormulaSpans(ScanText(text, null)).Count; }
            catch (InvalidDataException) { count++; }
        }
        return count;
    }

    // A saved text body spelled like the live text: a paragraph ends with U+000D,
    // a:br is U+000B, runs and fields give their text, and a native equation one
    // masked character.
    internal static string SavedText(XElement body)
    {
        var text = new StringBuilder();
        var paragraphs = body.Elements(A + "p").ToList();
        for (var index = 0; index < paragraphs.Count; index++)
        {
            if (index > 0) text.Append('\r');
            foreach (var child in LiveChildren(paragraphs[index]))
            {
                if (child.Name == A + "r" || child.Name == A + "fld")
                    text.Append((string?)child.Element(A + "t") ?? string.Empty);
                else if (child.Name == A + "br")
                    text.Append('\v');
                else if (child.Name == MathZone)
                    text.Append(EquationMask);
            }
        }
        return text.ToString();
    }

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

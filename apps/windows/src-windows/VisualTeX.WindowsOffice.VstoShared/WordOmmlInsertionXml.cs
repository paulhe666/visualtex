using System.Xml.Linq;

namespace VisualTeX.WordVsto;

/// <summary>
/// Pure XML splice used by paragraph-local OMML insertion. Word owns the
/// surrounding run-level wrappers (tracked insertion, content controls,
/// hyperlinks, custom XML, and similar containers). VisualTeX replaces only the
/// marker run at the exact location Word chose; it never promotes that run out of
/// its existing wrapper.
/// </summary>
internal static class WordOmmlInsertionXml
{
    private static readonly XNamespace Word =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // The character formatting of the last paragraph mark in a WordOpenXML
    // export (a paragraph's own export ends with that paragraph), without
    // revision marks; null when the mark has none.
    internal static string? ParagraphMarkRunProperties(string? wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml)) return null;
        var body = XDocument.Parse(wordOpenXml!, LoadOptions.PreserveWhitespace)
            .Descendants(Word + "body")
            .FirstOrDefault();
        var properties = body?
            .Elements(Word + "p")
            .LastOrDefault()?
            .Element(Word + "pPr")?
            .Element(Word + "rPr");
        if (properties is null) return null;
        var copy = new XElement(properties);
        copy.Elements()
            .Where(element => element.Name == Word + "ins"
                || element.Name == Word + "del"
                || element.Name == Word + "moveFrom"
                || element.Name == Word + "moveTo"
                || element.Name == Word + "rPrChange")
            .Remove();
        return copy.HasElements ? copy.ToString(SaveOptions.DisableFormatting) : null;
    }

    internal static void ReplaceOrdinaryMarkerRun(
        XElement markerNode,
        string markerText,
        XElement equation,
        XElement targetParagraph)
    {
        if (markerNode is null) throw new ArgumentNullException(nameof(markerNode));
        if (equation is null) throw new ArgumentNullException(nameof(equation));
        if (targetParagraph is null) throw new ArgumentNullException(nameof(targetParagraph));
        if (string.IsNullOrEmpty(markerText))
            throw new ArgumentException("The OMML insertion marker is empty.", nameof(markerText));
        if (markerNode.Name != Word + "t")
            throw new InvalidDataException("The OMML insertion marker is not an ordinary Word text node.");

        var sourceRun = markerNode.Parent;
        if (sourceRun is null || sourceRun.Name != Word + "r")
            throw new InvalidDataException(
                "The OMML insertion marker is not inside one ordinary Word run.");

        var ownerParagraphs = sourceRun.Ancestors(Word + "p").ToArray();
        if (ownerParagraphs.Length != 1
            || !ReferenceEquals(ownerParagraphs[0], targetParagraph))
            throw new InvalidDataException(
                "The OMML insertion marker run is not owned by the expected Word paragraph.");

        var markerIndex = markerNode.Value.IndexOf(markerText, StringComparison.Ordinal);
        if (markerIndex < 0
            || markerNode.Value.IndexOf(
                markerText,
                markerIndex + markerText.Length,
                StringComparison.Ordinal) >= 0)
            throw new InvalidDataException(
                "The ordinary Word run does not contain exactly one OMML insertion marker.");

        var prefix = markerNode.Value.Substring(0, markerIndex);
        var suffix = markerNode.Value.Substring(markerIndex + markerText.Length);

        XElement CreateSiblingRun(string text)
        {
            var run = new XElement(
                Word + "r",
                sourceRun.Attributes().Select(attribute => new XAttribute(attribute)));
            var runProperties = sourceRun.Element(Word + "rPr");
            if (runProperties is not null)
                run.Add(new XElement(runProperties));

            var textElement = new XElement(Word + "t", text);
            if (text.Length > 0
                && (char.IsWhiteSpace(text[0])
                    || char.IsWhiteSpace(text[text.Length - 1])))
            {
                textElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            }
            run.Add(textElement);
            return run;
        }

        var replacements = new List<object>();
        if (!string.IsNullOrEmpty(prefix))
            replacements.Add(CreateSiblingRun(prefix));
        replacements.Add(new XElement(equation));
        if (!string.IsNullOrEmpty(suffix))
            replacements.Add(CreateSiblingRun(suffix));

        // Replace in place. The source run may be a direct child of w:p, or it
        // may be nested under Word-owned run-level wrappers such as w:ins. Keeping
        // the current parent preserves revision provenance and other host
        // structure while changing only the marker atom.
        sourceRun.ReplaceWith(replacements.ToArray());
    }
}

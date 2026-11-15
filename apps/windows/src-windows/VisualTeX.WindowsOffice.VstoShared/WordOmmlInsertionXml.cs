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

    // The paragraph properties of the last paragraph in a WordOpenXML export (a
    // paragraph's own export ends with that paragraph): style, list numbering,
    // indentation, spacing and the mark's character formatting. When an equation
    // is imported into an empty paragraph Word keeps the package's paragraph
    // properties, so they must be the target paragraph's own, or its style and
    // numbering are lost. Revision marks and a section break are not carried.
    // Null when the paragraph has none.
    // The definition of the paragraph style that the last paragraph of a WordOpenXML
    // export names. An import package carries it in a styles part: without one,
    // Word drops the package paragraph's style reference. Null for the default style.
    internal static string? ParagraphStyleDefinition(string? wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml)) return null;
        var package = XDocument.Parse(wordOpenXml!, LoadOptions.PreserveWhitespace);
        var styleId = package.Descendants(Word + "body").FirstOrDefault()?
            .Elements(Word + "p").LastOrDefault()?
            .Element(Word + "pPr")?.Element(Word + "pStyle")?
            .Attribute(Word + "val")?.Value;
        if (string.IsNullOrEmpty(styleId)) return null;
        var style = package.Descendants(Word + "style")
            .FirstOrDefault(element => (string?)element.Attribute(Word + "styleId") == styleId);
        return style?.ToString(SaveOptions.DisableFormatting);
    }

    internal static string? ParagraphProperties(string? wordOpenXml)
    {
        if (string.IsNullOrWhiteSpace(wordOpenXml)) return null;
        var body = XDocument.Parse(wordOpenXml!, LoadOptions.PreserveWhitespace)
            .Descendants(Word + "body")
            .FirstOrDefault();
        var properties = body?
            .Elements(Word + "p")
            .LastOrDefault()?
            .Element(Word + "pPr");
        if (properties is null) return null;
        var copy = new XElement(properties);
        copy.Elements(Word + "pPrChange").Remove();
        copy.Elements(Word + "sectPr").Remove();
        copy.Element(Word + "rPr")?.Elements()
            .Where(element => element.Name == Word + "ins"
                || element.Name == Word + "del"
                || element.Name == Word + "moveFrom"
                || element.Name == Word + "moveTo"
                || element.Name == Word + "rPrChange")
            .Remove();
        if (copy.Element(Word + "rPr") is { HasElements: false } emptyRunProperties) emptyRunProperties.Remove();
        return copy.HasElements ? copy.ToString(SaveOptions.DisableFormatting) : null;
    }

    // Paragraph properties without a display host's own layout: tab stops and
    // justification belong to the VisualTeX display layout, not to the paragraph.
    internal static string? WithoutDisplayLayout(string? paragraphProperties)
    {
        if (string.IsNullOrWhiteSpace(paragraphProperties)) return null;
        var properties = XElement.Parse(paragraphProperties!, LoadOptions.PreserveWhitespace);
        properties.Elements(Word + "tabs").Remove();
        properties.Elements(Word + "jc").Remove();
        return properties.HasElements ? properties.ToString(SaveOptions.DisableFormatting) : null;
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

        // Word often keeps text, tabs and breaks of the same formatting in the
        // marker's run. Everything before the marker stays in the run before the
        // equation and everything after it in the run after; nothing is dropped.
        var before = markerNode.NodesBeforeSelf()
            .Where(node => node is not XElement element || element.Name != Word + "rPr").ToArray();
        var after = markerNode.NodesAfterSelf().ToArray();

        XElement CreateTextElement(string text)
        {
            var textElement = new XElement(Word + "t", text);
            if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1]))
                textElement.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            return textElement;
        }

        XElement CreateSiblingRun(IEnumerable<XNode> leading, string text, IEnumerable<XNode> trailing)
        {
            var run = new XElement(
                Word + "r",
                sourceRun.Attributes().Select(attribute => new XAttribute(attribute)));
            var runProperties = sourceRun.Element(Word + "rPr");
            if (runProperties is not null)
                run.Add(new XElement(runProperties));
            foreach (var node in leading) run.Add(CloneNode(node));
            if (text.Length > 0) run.Add(CreateTextElement(text));
            foreach (var node in trailing) run.Add(CloneNode(node));
            return run;
        }

        static XNode CloneNode(XNode node) => node switch
        {
            XElement element => new XElement(element),
            XText text => new XText(text),
            _ => node,
        };

        var replacements = new List<object>();
        if (!string.IsNullOrEmpty(prefix) || before.OfType<XElement>().Any())
            replacements.Add(CreateSiblingRun(before, prefix, Array.Empty<XNode>()));
        replacements.Add(new XElement(equation));
        if (!string.IsNullOrEmpty(suffix) || after.OfType<XElement>().Any())
            replacements.Add(CreateSiblingRun(Array.Empty<XNode>(), suffix, after));

        // Replace in place. The source run may be a direct child of w:p, or it
        // may be nested under Word-owned run-level wrappers such as w:ins. Keeping
        // the current parent preserves revision provenance and other host
        // structure while changing only the marker atom.
        sourceRun.ReplaceWith(replacements.ToArray());
    }
}

using System.Xml.Linq;

namespace VisualTeX.WordVsto;

// Pure comparisons of Word recovery evidence. Mathematical content, fonts,
// geometry, explicit language assignments and field results remain significant.
internal static class WordRecoveryXmlEquivalence
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    internal static string NormalizeHeaderFooterPayload(XElement xmlData)
    {
        var copy = new XElement(xmlData);
        foreach (var attribute in copy.DescendantsAndSelf().Attributes()
                     .Where(a => a.Name.Namespace == W && a.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).ToArray())
            attribute.Remove();
        return copy.ToString(SaveOptions.DisableFormatting);
    }

    internal static bool HasOnlyAutomaticChineseProofingAdditions(string before, string after)
    {
        var additions = 0;
        var equal = Compare(XElement.Parse(before), XElement.Parse(after), false, ref additions);
        return equal && additions > 0;
    }

    private static bool Compare(XElement before, XElement after, bool chineseParagraph, ref int additions)
    {
        if (before.Name != after.Name) return false;
        if (before.Name == W + "p")
            chineseParagraph = before.Descendants(W + "t").Any(t => t.Value.Any(c => c >= '\u4E00' && c <= '\u9FFF'));
        var oldAttributes = before.Attributes().Where(a => !a.IsNamespaceDeclaration).ToDictionary(a => a.Name, a => a.Value);
        var newAttributes = after.Attributes().Where(a => !a.IsNamespaceDeclaration).ToDictionary(a => a.Name, a => a.Value);
        foreach (var item in oldAttributes)
            if (!newAttributes.TryGetValue(item.Key, out var actual) || item.Value != actual) return false;
        foreach (var item in newAttributes.Where(p => !oldAttributes.ContainsKey(p.Key)))
        {
            if (!chineseParagraph || before.Name != W + "lang" || item.Key != W + "eastAsia" || item.Value != "zh-CN") return false;
            additions++;
        }
        var oldNodes = before.Nodes().ToArray();
        var newNodes = after.Nodes().ToArray();
        var i = 0;
        foreach (var node in newNodes)
        {
            var original = i < oldNodes.Length ? oldNodes[i] : null;
            if (node is XElement current && original is XElement expected && current.Name == expected.Name)
            {
                if (!Compare(expected, current, chineseParagraph, ref additions)) return false;
                i++;
            }
            else if (node is XElement extra && chineseParagraph
                     && ((before.Name == W + "rPr" && extra.Name == W + "lang"
                          && !before.Elements(W + "lang").Any() && IsAutomaticChineseLanguage(extra))
                         || ((before.Name == W + "r" || before.Name == W + "pPr")
                             && extra.Name == W + "rPr" && !before.Elements(W + "rPr").Any()
                             && string.IsNullOrWhiteSpace(extra.Value)
                             && !extra.Attributes().Any(a => !a.IsNamespaceDeclaration)
                             && extra.Elements().Count() == 1 && IsAutomaticChineseLanguage(extra.Elements().Single()))))
                additions++;
            else
            {
                if (original is null || !XNode.DeepEquals(original, node)) return false;
                i++;
            }
        }
        return i == oldNodes.Length;
    }

    private static bool IsAutomaticChineseLanguage(XElement element) =>
        element.Name == W + "lang" && !element.HasElements && string.IsNullOrWhiteSpace(element.Value)
        && element.Attributes().Count(a => !a.IsNamespaceDeclaration) == 1
        && (string?)element.Attribute(W + "eastAsia") == "zh-CN";
}

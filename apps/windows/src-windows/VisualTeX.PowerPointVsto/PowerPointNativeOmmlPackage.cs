using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.PowerPoint;

namespace VisualTeX.PowerPointVsto;

// On-command native source capture. Never write VisualTeX metadata into an
// Office-owned equation and never use the clipboard during selection events.
internal static class PowerPointNativeOmmlPackage
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    private const string Drawing2010 = "http://schemas.microsoft.com/office/drawing/2010/main";

    internal static string ReadStandalone(Presentation presentation, int slideId, int shapeId)
    {
        var path = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-omml-{Guid.NewGuid():N}.pptx");
        try
        {
            presentation.SaveCopyAs(path, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using var stream = File.OpenRead(path);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            return ReadStandalone(archive, slideId, shapeId);
        }
        finally { try { File.Delete(path); } catch { } }
    }

    internal static string ReadStandalone(ZipArchive archive, int slideId, int shapeId)
    {
        var shape = ReadTextShape(archive, slideId, shapeId);
        var equations = shape.Descendants(M + "oMath").ToArray();
        var surroundingText = shape.Descendants(A + "t").Any(element =>
            !element.Ancestors(M + "oMath").Any() && !string.IsNullOrWhiteSpace(element.Value));
        var hasOtherContent = shape.Descendants(A + "fld").Any() || shape.Descendants(A + "br").Any();
        var paragraphs = shape.Element(P + "txBody")?.Elements(A + "p").ToArray() ?? Array.Empty<XElement>();
        var hasExtraText = shape.Descendants(A + "t").Any();
        if (equations.Length != 1 || surroundingText || hasOtherContent || hasExtraText
            || paragraphs.Length != 1 || !paragraphs[0].Descendants(M + "oMath").Any())
            throw new InvalidOperationException("请选择只包含一个原生公式的独立文本框；混排文字或多个公式不能作为整个对象替换。");
        return equations[0].ToString(SaveOptions.DisableFormatting);
    }

    internal static string ReadTextShapeXml(Presentation presentation, int slideId, int shapeId)
    {
        var path = Path.Combine(Path.GetTempPath(), $"visualtex-ppt-text-{Guid.NewGuid():N}.pptx");
        try
        {
            presentation.SaveCopyAs(path, PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            using var stream = File.OpenRead(path);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            return ReadTextShape(archive, slideId, shapeId).ToString(SaveOptions.DisableFormatting);
        }
        finally { try { File.Delete(path); } catch { } }
    }

    internal static string ReadEquation(string shapeXml, int ordinal, int expectedCount)
    {
        var shape = XElement.Parse(shapeXml, LoadOptions.PreserveWhitespace);
        var equations = shape.Descendants(M + "oMath").ToArray();
        if (equations.Length != expectedCount || ordinal < 1 || ordinal > equations.Length)
            throw new InvalidDataException("The native math ranges do not match the saved PowerPoint equations.");
        return equations[ordinal - 1].ToString(SaveOptions.DisableFormatting);
    }

    private static XElement ReadTextShape(ZipArchive archive, int slideId, int shapeId)
    {
        var slide = ReadXml(archive, "ppt/presentation.xml").Descendants(P + "sldId")
            .SingleOrDefault(element => (int?)element.Attribute("id") == slideId)
            ?? throw new InvalidOperationException("The native equation slide no longer exists.");
        var id = (string?)slide.Attribute(R + "id");
        var relationship = ReadXml(archive, "ppt/_rels/presentation.xml.rels").Descendants(Rel + "Relationship")
            .SingleOrDefault(element => (string?)element.Attribute("Id") == id)
            ?? throw new InvalidDataException("The native equation slide relationship is missing.");
        if ((string?)relationship.Attribute("TargetMode") == "External")
            throw new InvalidDataException("The native equation slide is external.");
        var target = (string?)relationship.Attribute("Target")
            ?? throw new InvalidDataException("The native equation slide has no target.");
        var uri = new Uri(new Uri("https://package.invalid/ppt/presentation.xml"), target);
        if (uri.Host != "package.invalid") throw new InvalidDataException("Invalid native equation slide target.");
        var slideXml = ReadXml(archive, Uri.UnescapeDataString(uri.AbsolutePath.TrimStart('/')));
        return SelectTextShape(slideXml, shapeId);
    }

    internal static XElement SelectTextShape(XDocument slideXml, int shapeId)
    {
        // PowerPoint saves the a14 native-math choice and an older-host fallback
        // with the SAME shape ID. The fallback can itself be a text shape.
        // Resolve compatibility branches before enforcing exact object identity;
        // unrelated duplicate IDs and unknown representations remain errors.
        var matches = slideXml.Descendants(P + "sp").Where(element =>
            (int?)element.Element(P + "nvSpPr")?.Element(P + "cNvPr")?.Attribute("id") == shapeId
            && element.Element(P + "txBody") is not null).Where(IsSupportedShapeBranch).ToArray();
        return matches.SingleOrDefault()
            ?? throw new InvalidOperationException("The exact text shape no longer exists.");
    }

    private static bool IsSupportedShapeBranch(XElement shape)
    {
        foreach (var branch in shape.Ancestors().Where(element => element.Parent?.Name == Mc + "AlternateContent"))
        {
            var choices = branch.Parent!.Elements(Mc + "Choice").ToArray();
            // An unknown earlier Choice may be supported by this Office version.
            // Skipping it would fingerprint a different representation. Only the
            // observed native-math single-choice contract is editable here.
            if (choices.Length != 1)
                throw new InvalidDataException("The text shape has multiple compatibility choices; its active native representation is unverified.");
            var supported = choices[0];
            var requires = ((string?)supported.Attribute("Requires") ?? "")
                .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (requires.Length == 0 || requires.Any(prefix => supported.GetNamespaceOfPrefix(prefix)?.NamespaceName != Drawing2010))
                throw new InvalidDataException("The text shape uses an unsupported compatibility representation; its fallback is not an editable native source.");
            if (!ReferenceEquals(branch, supported)) return false;
        }
        return true;
    }

    private static XDocument ReadXml(ZipArchive archive, string part)
    {
        var entry = archive.GetEntry(part) ?? throw new InvalidDataException("Missing native equation part: " + part);
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = 32 * 1024 * 1024,
        });
        return XDocument.Load(reader);
    }
}

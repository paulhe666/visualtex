using System.IO;
using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WpsWriterVsto;

/// <summary>
/// Extracts the embedded MathType compound file from a range's Flat OPC.
/// WPS exposes WordOpenXML but its clipboard data object rejects the
/// EmbeddedObject/EmbedSource FORMATETC, so this is the primary read path.
/// </summary>
internal static class WpsMathTypeOpenXml
{
    private const int MaximumFlatOpcCharacters = 32 * 1024 * 1024;

    private static readonly XNamespace PackageNamespace =
        "http://schemas.microsoft.com/office/2006/xmlPackage";

    internal static bool TryReadCompoundFile(string? flatOpc, out byte[] compoundFile)
    {
        compoundFile = Array.Empty<byte>();
        if (string.IsNullOrWhiteSpace(flatOpc)) return false;
        if (flatOpc!.Length > MaximumFlatOpcCharacters) return false;

        try
        {
            var document = XDocument.Parse(flatOpc, LoadOptions.None);
            foreach (var part in document.Descendants(PackageNamespace + "part"))
            {
                var contentType = (string?)part.Attribute(PackageNamespace + "contentType");
                if (contentType is null
                    || contentType.IndexOf("oleObject", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                var binary = part.Element(PackageNamespace + "binaryData");
                if (binary is null) continue;

                var encoded = new string(
                    binary.Value.Where(character => !char.IsWhiteSpace(character)).ToArray());
                byte[] bytes;
                try { bytes = Convert.FromBase64String(encoded); }
                catch (FormatException) { continue; }
                if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(bytes)) continue;

                compoundFile = bytes;
                return true;
            }
            return false;
        }
        catch (Exception error)
        {
            WpsLog.WriteException("Parsing range WordOpenXML failed", error);
            return false;
        }
    }
}

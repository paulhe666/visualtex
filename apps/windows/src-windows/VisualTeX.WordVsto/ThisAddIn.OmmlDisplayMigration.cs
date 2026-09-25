using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace VisualTeX.WordVsto;

public sealed partial class ThisAddIn
{
    /// <summary>
    /// Explicit compatibility migration for documents that still contain a
    /// retired VisualTeX numbering topology.
    ///
    /// Current OMath.Type is authoritative and is never rewritten from legacy
    /// metadata. The host core performs one immutable inventory pass only to
    /// decide whether a legacy numbering container is actually present. The old
    /// numbering reconciler is invoked only as a one-shot legacy importer; normal
    /// open/read/edit/save paths never enter it and all newly written structures
    /// are owned by the canonical numbering core.
    /// </summary>
    internal static int MigrateManagedOmmlDisplayTypes(
        Document document)
    {
        if (document is null)
            throw new ArgumentNullException(nameof(document));
        if (document.ReadOnly)
            throw new UnauthorizedAccessException(
                "Cannot migrate formula numbering in a read-only document.");

        var index =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        var legacyCount = 0;

        foreach (var host in index.Omml
                     .Concat(index.VisualTeX))
        {
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.Legacy)
                legacyCount++;
        }

        if (legacyCount == 0)
            return 0;

        // Compatibility boundary only. This method exists for explicit upgrade
        // of retired documents/acceptance fixtures. It is intentionally not
        // subscribed to DocumentChange/DocumentOpen and is never a fallback from
        // a local operation.
        var migrated =
            WordEquationNumbering.UpdateEquationNumbers(
                document);

        // Re-prove the result from current local Word structure. A legacy importer
        // may not declare success merely because its own bookkeeping completed.
        var after =
            WordFormulaHostResolver.CaptureDocumentIndex(
                document);
        foreach (var host in after.Omml
                     .Concat(after.VisualTeX))
        {
            var numbering =
                WordFormulaNumberingResolver.ResolveLocal(
                    document,
                    host);
            if (numbering.ContainerKind ==
                WordFormulaNumberingContainerKind.Legacy)
                throw new InvalidDataException(
                    "Legacy numbering migration left a retired formula host in the document.");
        }

        return migrated;
    }
}

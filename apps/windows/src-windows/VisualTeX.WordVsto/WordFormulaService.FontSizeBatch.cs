using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// One size for every inline or every display formula of the document (Ribbon 行内 / 行间, as on the Mac):
// VisualTeX objects and Word equations in one Undo step, then MathType objects one by one (a MathType
// rewrite exports the object's XML first, which Word does not allow inside an open Undo record).
internal sealed partial class WordFormulaService
{
    internal sealed class FormulaFontSizeBatchResult
    {
        internal int Resized { get; set; }
        internal List<string> Failures { get; } = new();
    }

    internal FormulaFontSizeBatchResult ApplyFontSizeToDocumentFormulas(bool display, double requestedFontSizePt)
    {
        var size = FormulaFontSize.Normalize(requestedFontSizePt);
        var result = new FormulaFontSizeBatchResult();
        Document? document = null;
        Selection? selection = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(OfficePluginLanguage.Text(
                    "没有打开的 Word 文档。", "No Word document is open."));
            EnsureWritable(document);
            var target = document;

            // A numbered equation is a display formula, whatever its OMath type.
            var hosts = new List<WordFormulaHostDescriptor>();
            var index = WordFormulaHostResolver.CaptureDocumentIndex(target);
            foreach (var host in index.Omml.Concat(index.VisualTeX))
            {
                try
                {
                    host.Numbering = WordFormulaNumberingResolver.ResolveLocal(target, host);
                    if ((host.Display || host.Numbering.Numbered) == display) hosts.Add(host);
                }
                catch (Exception error)
                {
                    result.Failures.Add(error.Message);
                }
            }
            hosts.Sort((left, right) => right.Range.Start.CompareTo(left.Range.Start));
            if (hosts.Count > 0)
            {
                result.Resized = WordFormulaMutationTransaction.Execute(
                    _application,
                    target,
                    "VisualTeX Formula Font Size",
                    () =>
                    {
                        foreach (var host in hosts) ResizeHostInTransaction(target, host, size);
                        return hosts.Count;
                    });
            }

            var mathType = WordMathTypeHostAdapter.CaptureDocumentIndex(_application, target)
                .Where(host => (host.Display || host.Numbering.Numbered) == display)
                .OrderByDescending(host => host.Range.Start)
                .ToList();
            if (mathType.Count > 0)
            {
                // The single-formula command resizes the selected MathType object; the user's
                // selection comes back afterwards.
                selection = _application.Selection;
                var start = selection.Start;
                var end = selection.End;
                foreach (var host in mathType)
                {
                    Range? range = null;
                    try
                    {
                        range = WordFormulaHostSemanticReader.CreateRange(target, host.Range);
                        range.Select();
                        SetSelectedFormulaFontSizeCore(size);
                        result.Resized++;
                    }
                    catch (Exception error)
                    {
                        result.Failures.Add(error.Message);
                    }
                    finally { Release(range); }
                }
                try { selection.SetRange(start, end); } catch (Exception ignored) { IgnoredErrors.Ignore(ignored); }
            }
            return result;
        }
        finally
        {
            Release(selection);
            Release(document);
        }
    }

    // As the single-formula command does (SetSelectedFormulaFontSizeCore), inside the batch's transaction.
    private static void ResizeHostInTransaction(Document document, WordFormulaHostDescriptor host, float size)
    {
        if (host.Kind == WordFormulaHostKind.Omml)
        {
            var owned = WordFormulaHostMutationKernel.AdoptUnownedOmmlForMutation(
                document,
                host,
                host.FormulaId ?? Guid.NewGuid().ToString("D"));
            Range? range = null;
            try
            {
                range = WordFormulaHostSemanticReader.CreateRange(document, owned.Range);
                WordFormulaHostLayout.ApplyOmmlLocalTypography(range, size);
            }
            finally { Release(range); }
            return;
        }
        var updated = WordVisualTeXHostWriter.UpdateFontSize(document, host, size);
        if (updated.Numbering.ContainerKind == WordFormulaNumberingContainerKind.CanonicalBodyTabParagraph
            && WordVisualTeXParagraphNumbering.IsSelfContainedHost(document, updated)
            && updated.Metadata is not null)
            WordVisualTeXParagraphNumbering.ApplyNumberFormat(document, updated, updated.Metadata);
    }
}

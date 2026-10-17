using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

/// <summary>
/// Update Equation Format: brings every equation number an earlier VisualTeX for
/// Windows wrote into the current format (docs/equation-format-v2.md), the
/// format VisualTeX for macOS writes too. Each formula keeps its kind; only
/// number structure and retired VisualTeX state change.
/// </summary>
internal sealed partial class WordFormulaService
{
    internal sealed class EquationFormatUpgradeReport
    {
        /// <summary>Formulas still numbered by a retired caption/frame/table layout.</summary>
        internal int RetiredLayouts { get; set; }

        /// <summary>Native #(...) numbers not in the current canonical form.</summary>
        internal int NonCanonicalNumbers { get; set; }

        /// <summary>The retired document variable that stored the numbering style.</summary>
        internal bool LegacyFormatVariable { get; set; }

        internal int Total =>
            RetiredLayouts + NonCanonicalNumbers + (LegacyFormatVariable ? 1 : 0);
    }

    public EquationFormatUpgradeReport ScanEquationFormat()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            return ScanEquationFormat(document);
        }
        finally { Release(document); }
    }

    private static EquationFormatUpgradeReport ScanEquationFormat(Document document) =>
        new()
        {
            RetiredLayouts = CountRetiredNumberingLayouts(document),
            NonCanonicalNumbers = WordNativeNumberingAdoptionPlan.CountNonCanonicalNumbers(
                document, WordEquationNumberFormatAuthority.Resolve(document)),
            LegacyFormatVariable = WordEquationNumbering.HasLegacyEquationNumberFormat(document),
        };

    public int UpgradeEquationFormat()
    {
        Document? document = null;
        UndoRecord? undo = null;
        var undoSteps = 0;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException("No active Word document.");
            EnsureWritable(document);
            var count = 0;
            // Every formula keeps its kind: the update changes number structure
            // and retired VisualTeX state only, never how many equations or
            // pictures the document holds.
            var inventory = EquationInventory(document);

            try
            {
                // Retired layouts first: the explicit one-shot importer turns them
                // into canonical native numbers (ThisAddIn.OmmlDisplayMigration).
                // It owns no Undo record, so give it exactly one.
                if (CountRetiredNumberingLayouts(document) > 0)
                {
                    undo = _application.UndoRecord;
                    if (undo.IsRecordingCustomRecord || undo.CustomRecordLevel > 0)
                        throw new InvalidOperationException(
                            "The equation format update cannot run inside another Word undo transaction.");
                    undo.StartCustomRecord("VisualTeX Update Equation Format");
                    try { count += ThisAddIn.MigrateManagedOmmlDisplayTypes(document); }
                    finally { undo.EndCustomRecord(); }
                    undoSteps++;
                }

                // Then every native number in the document's current style, written
                // by the shared number builder in one Undo record (PlanFormat).
                var formatId = WordEquationNumberFormatAuthority.Resolve(document);
                if (WordNativeNumberingAdoptionPlan.CountNonCanonicalNumbers(document, formatId) > 0)
                {
                    count += SetEquationNumberFormatCore(formatId);
                    undoSteps++;
                }

                WordEquationNumbering.RetireLegacyEquationNumberFormat(document);

                var after = ScanEquationFormat(document);
                if (after.Total != 0)
                    throw new InvalidDataException(
                        $"The equation format update left {after.RetiredLayouts} retired layout(s) and "
                        + $"{after.NonCanonicalNumbers} number(s) in an older form.");
                if (!string.Equals(EquationInventory(document), inventory, StringComparison.Ordinal))
                    throw new InvalidDataException(
                        "The equation format update changed how many equations or formula pictures the document holds.");
                return count;
            }
            catch
            {
                // Roll the whole update back: exactly the Undo records it made.
                for (var step = 0; step < undoSteps; step++)
                {
                    try { document.Undo(); } catch { break; }
                }
                throw;
            }
        }
        finally
        {
            Release(undo);
            Release(document);
        }
    }

    private static string EquationInventory(Document document)
    {
        OMaths? maths = null;
        InlineShapes? shapes = null;
        try
        {
            maths = document.OMaths;
            shapes = document.InlineShapes;
            return maths.Count + "|" + shapes.Count;
        }
        finally
        {
            Release(shapes);
            Release(maths);
        }
    }

    private static int CountRetiredNumberingLayouts(Document document)
    {
        var index = WordFormulaHostResolver.CaptureDocumentIndex(document);
        return index.Omml
            .Concat(index.VisualTeX)
            .Count(host =>
                WordFormulaNumberingResolver.ResolveLocal(document, host).ContainerKind
                == WordFormulaNumberingContainerKind.Legacy);
    }
}

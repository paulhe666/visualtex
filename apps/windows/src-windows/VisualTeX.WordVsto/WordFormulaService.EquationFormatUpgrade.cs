using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

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

        /// <summary>
        /// Equations still in the CFF Latin Modern Math (VisualTeX before 1.2.8
        /// made it the document math font). Word's Save as PDF rasterizes that
        /// font and drops the parentheses of equation numbers; the TrueType
        /// VisualTeX Latin Modern Math has the same glyphs and is embedded.
        /// </summary>
        internal bool LegacyMathFont { get; set; }

        internal int Total =>
            RetiredLayouts + NonCanonicalNumbers + (LegacyFormatVariable ? 1 : 0) + (LegacyMathFont ? 1 : 0);
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
            LegacyMathFont = UsesLegacyLatinModernMath(document),
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

                if (UsesLegacyLatinModernMath(document))
                {
                    undo ??= _application.UndoRecord;
                    if (undo.IsRecordingCustomRecord || undo.CustomRecordLevel > 0)
                        throw new InvalidOperationException(
                            "The equation format update cannot run inside another Word undo transaction.");
                    undo.StartCustomRecord("VisualTeX Update Equation Format");
                    try { MigrateLegacyLatinModernMath(document); }
                    finally { undo.EndCustomRecord(); }
                    undoSteps++;
                }

                var after = ScanEquationFormat(document);
                if (after.Total != 0)
                    throw new InvalidDataException(
                        $"The equation format update left {after.RetiredLayouts} retired layout(s), "
                        + $"{after.NonCanonicalNumbers} number(s) in an older form"
                        + (after.LegacyMathFont ? " and equations in Latin Modern Math." : "."));
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

    private static bool UsesLegacyLatinModernMath(Document document)
    {
        string? mathFont = null;
        try { mathFont = document.OMathFontName; }
        catch (COMException) { }
        return string.Equals(
                mathFont, WordOfficeMathFontLoader.LegacyLatinModernMathFamily, StringComparison.OrdinalIgnoreCase)
            || AnyStory(document, story => ReplaceFont(story, WordOfficeMathFontLoader.LegacyLatinModernMathFamily, null));
    }

    // The document math font only governs new equations; every existing run
    // and delimiter names its font explicitly, so the old name is replaced in
    // every story (body, notes, headers, text boxes) as well. The glyphs are
    // the same, so nothing reflows.
    private static void MigrateLegacyLatinModernMath(Document document)
    {
        var family = WordOfficeMathFontLoader.LatinModernMathFamily;
        WordOfficeMathFontLoader.EnsureLoaded();
        string? mathFont = null;
        try { mathFont = document.OMathFontName; }
        catch (COMException) { }
        if (string.Equals(
                mathFont, WordOfficeMathFontLoader.LegacyLatinModernMathFamily, StringComparison.OrdinalIgnoreCase))
        {
            document.OMathFontName = family;
            if (!string.Equals(document.OMathFontName, family, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException(
                    $"Word did not accept '{family}' as this document's math font. Reinstall the VisualTeX Office integration.");
        }
        _ = AnyStory(document, story => ReplaceFont(story, WordOfficeMathFontLoader.LegacyLatinModernMathFamily, family));
    }

    // Visits every story range, including linked header/footer and text-box
    // stories; stops early when visit returns true.
    private static bool AnyStory(Document document, Func<Range, bool> visit)
    {
        StoryRanges? stories = null;
        try
        {
            stories = document.StoryRanges;
            foreach (Range first in stories)
            {
                var story = first;
                while (story is not null)
                {
                    Range? next = null;
                    try
                    {
                        if (visit(story)) return true;
                        next = story.NextStoryRange;
                    }
                    finally { Release(story); }
                    story = next;
                }
            }
            return false;
        }
        finally { Release(stories); }
    }

    // replacement null: report whether the story holds the font at all.
    private static bool ReplaceFont(Range story, string fontName, string? replacement)
    {
        Range? search = null;
        Find? find = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Replacement? replace = null;
        Microsoft.Office.Interop.Word.Font? replaceFont = null;
        try
        {
            search = story.Duplicate;
            find = search.Find;
            find.ClearFormatting();
            font = find.Font;
            font.Name = fontName;
            find.Text = string.Empty;
            find.Format = true;
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;
            find.MatchWildcards = false;
            if (replacement is null)
                return find.Execute();
            replace = find.Replacement;
            replace.ClearFormatting();
            replace.Text = string.Empty;
            replaceFont = replace.Font;
            replaceFont.Name = replacement;
            object replaceAll = WdReplace.wdReplaceAll;
            return find.Execute(Replace: ref replaceAll);
        }
        finally
        {
            Release(replaceFont);
            Release(replace);
            Release(font);
            Release(find);
            Release(search);
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

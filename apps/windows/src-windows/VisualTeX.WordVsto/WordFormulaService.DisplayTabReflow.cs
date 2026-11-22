using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using Range = Microsoft.Office.Interop.Word.Range;

namespace VisualTeX.WordVsto;

// A display formula keeps the tab stops of the column it was written in. When its column is another one now (the
// section got columns, the formula moved to the other column, the margins changed), Update Equation Format and
// renumbering set just those stops again from the column (or cell) the formula is in (WordDisplayColumns): the
// centre stop, and the right stop where the formula has one. A first-line indent left on the formula's paragraph
// (2 characters of Chinese body text, which older writers could not clear) goes too. Nothing else changes.
internal sealed partial class WordFormulaService
{
    /// <summary>How many VisualTeX display formulas of the active document have stops their column does not give.</summary>
    internal int CountDisplayTabStopReflow()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(NoActiveDocumentMessage);
            return ReflowDisplayTabStops(document, apply: false, multiColumnOnly: false);
        }
        finally { Release(document); }
    }

    /// <summary>Sets those stops again, in one Undo step; returns how many formulas changed.</summary>
    internal int ApplyDisplayTabStopReflow()
    {
        Document? document = null;
        try
        {
            document = _application.ActiveDocument
                ?? throw new InvalidOperationException(NoActiveDocumentMessage);
            EnsureWritable(document);
            return WordFormulaMutationTransaction.Execute(
                _application,
                document,
                "VisualTeX Reset Display Tab Stops",
                () => ReflowDisplayTabStops(document, apply: true, multiColumnOnly: false));
        }
        finally { Release(document); }
    }

    /// <summary>
    /// The VisualTeX display formulas (outside tables) whose centre or right stop is not where their column puts it;
    /// with <paramref name="apply"/> their stops are set again. <paramref name="multiColumnOnly"/>: a document
    /// without newspaper columns is skipped at once (renumbering keeps its speed; there page-width stops are right).
    /// </summary>
    internal int ReflowDisplayTabStops(Document document, bool apply, bool multiColumnOnly)
    {
        using var metric = WordOperationMetrics.Measure("Display.TabStops.Reflow");
        if (multiColumnOnly && !HasMultiColumnSection(document)) return 0;
        // Setting tab stops moves no text, so the formulas' positions stay valid for the section lookup.
        var columns = WordDisplayColumns.BatchCache.BySectionPosition(document);
        var index = WordFormulaHostResolver.CaptureDocumentIndex(document);
        var count = 0;
        foreach (var host in index.VisualTeX)
        {
            if (!host.Display || host.WithinTable) continue;
            Range? range = null;
            Paragraphs? paragraphs = null;
            Paragraph? paragraph = null;
            Range? paragraphRange = null;
            ParagraphFormat? format = null;
            TabStops? tabs = null;
            TabStop? tab = null;
            try
            {
                range = WordFormulaHostSemanticReader.CreateRange(document, host.Range);
                paragraphs = range.Paragraphs;
                if (paragraphs.Count != 1) continue;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                format = paragraph.Format;
                tabs = format.TabStops;
                float? center = null;
                float? right = null;
                for (var position = 1; position <= tabs.Count; position++)
                {
                    Release(tab);
                    tab = tabs[position];
                    if (tab.Alignment == WdTabAlignment.wdAlignTabCenter) center ??= tab.Position;
                    else if (tab.Alignment == WdTabAlignment.wdAlignTabRight) right ??= tab.Position;
                }
                Release(tab);
                tab = null;
                // a display formula without the centre stop is not in the tab layout: left as it is
                if (center is null) continue;
                if (columns.ColumnWidthFor(document, paragraphRange) is not { } width) continue;
                // A first-line indent (one an older writer could not clear: WordEquationParagraphLayout.ZeroIndents)
                // moves a tab-laid line off its stops: it goes, and the stops are those of a line without one.
                var indented = Math.Abs(format.FirstLineIndent) > 0.05f;
                var expected = DisplayColumnGeometry.TabStops(width, format.LeftIndent, 0f, format.RightIndent);
                if (!indented
                    && Math.Abs(center.Value - expected.Center) <= 0.5f
                    && (right is null || Math.Abs(right.Value - expected.Right) <= 0.5f))
                    continue;
                count++;
                if (!apply) continue;
                if (indented) WordEquationParagraphLayout.ZeroIndents(format, firstLineOnly: true);
                tabs.ClearAll();
                Release(tabs);
                tabs = format.TabStops;
                tab = tabs.Add(expected.Center, WdTabAlignment.wdAlignTabCenter, WdTabLeader.wdTabLeaderSpaces);
                if (right is not null)
                {
                    Release(tab);
                    Release(tabs);
                    tabs = format.TabStops;
                    tab = tabs.Add(expected.Right, WdTabAlignment.wdAlignTabRight, WdTabLeader.wdTabLeaderSpaces);
                }
            }
            catch (Exception error) when (!apply && error is COMException)
            {
                // a formula Word cannot describe now is not counted; the update leaves it as it is
                IgnoredErrors.Ignore(error);
            }
            finally
            {
                Release(tab);
                Release(tabs);
                Release(format);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(range);
            }
        }
        return count;
    }

    private static bool HasMultiColumnSection(Document document)
    {
        Sections? sections = null;
        try
        {
            sections = document.Sections;
            for (var index = 1; index <= sections.Count; index++)
            {
                Section? section = null;
                PageSetup? setup = null;
                TextColumns? columns = null;
                try
                {
                    section = sections[index];
                    setup = section.PageSetup;
                    columns = setup.TextColumns;
                    if (columns is not null && columns.Count > 1) return true;
                }
                finally { Release(columns); Release(setup); Release(section); }
            }
            return false;
        }
        catch (Exception ignored)
        {
            IgnoredErrors.Ignore(ignored);
            return true;
        }
        finally { Release(sections); }
    }
}

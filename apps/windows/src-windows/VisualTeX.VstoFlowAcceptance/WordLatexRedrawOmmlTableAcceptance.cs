using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordLatexRedrawOmmlTableAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);

        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mi>a</mi><mo>+</mo><mi>b</mi>"
            + "</math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Table? table = null;
        Word.Range? anchor = null;
        Word.Cell? cell = null;
        Word.Range? cellRange = null;
        Word.Range? probe = null;
        try
        {
            if (AttachActiveWord)
                throw new InvalidOperationException(
                    "OMML table redraw acceptance must never attach the user's active Word.");

            application =
                CreateFreshAcceptanceAutomationWord(
                    artifactRoot);
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();

            anchor =
                document.Range(
                    0,
                    0);
            table =
                document.Tables.Add(
                    anchor,
                    2,
                    4);

            for (var row = 1; row <= 2; row++)
            {
                Release(cellRange);
                cellRange = null;
                Release(cell);
                cell = table.Cell(
                    row,
                    1);
                cellRange =
                    cell.Range.Duplicate;
                cellRange.End =
                    Math.Max(
                        cellRange.Start,
                        cellRange.End - 1);
                cellRange.Text =
                    "$$a+b$$";
            }

            Release(cellRange);
            cellRange = null;
            Release(cell);
            cell = table.Cell(
                1,
                2);
            cellRange =
                cell.Range.Duplicate;
            cellRange.End =
                Math.Max(
                    cellRange.Start,
                    cellRange.End - 1);
            cellRange.Text =
                "left $a+b$ right";

            static float ReadFirstRowRenderedHeight(Word.Table value)
            {
                Word.Cell? first = null;
                Word.Cell? second = null;
                Word.Range? firstStart = null;
                Word.Range? secondStart = null;
                try
                {
                    first = value.Cell(1, 1);
                    second = value.Cell(2, 1);
                    firstStart = first.Range.Duplicate;
                    secondStart = second.Range.Duplicate;
                    firstStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    secondStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    var firstY = Convert.ToSingle(
                        firstStart.Information[
                            Word.WdInformation.wdVerticalPositionRelativeToPage]);
                    var secondY = Convert.ToSingle(
                        secondStart.Information[
                            Word.WdInformation.wdVerticalPositionRelativeToPage]);
                    return secondY - firstY;
                }
                finally
                {
                    Release(secondStart);
                    Release(firstStart);
                    Release(second);
                    Release(first);
                }
            }

            var sourceFirstRowHeight =
                ReadFirstRowRenderedHeight(table);
            Console.WriteLine(
                $"[TABLE OMML HEIGHT BEFORE] firstRow={sourceFirstRowHeight:0.###}pt.");

            var service =
                new WordFormulaService(
                    application);
            var plan =
                service.CaptureLatexRedrawPlan(
                    wholeDocument: true);

            AssertEqual(
                3,
                plan.Targets.Count,
                "Table OMML redraw fixture did not capture two display formulas plus one inline formula.");
            AssertEqual(
                2,
                plan.Targets.Count(target =>
                    string.Equals(
                        target.DisplayMode,
                        "block",
                        StringComparison.Ordinal)),
                "Table OMML redraw fixture did not classify both display formulas.");
            AssertEqual(
                1,
                plan.Targets.Count(target =>
                    string.Equals(
                        target.DisplayMode,
                        "inline",
                        StringComparison.Ordinal)),
                "Table OMML redraw fixture did not classify the inline mixed-text formula.");

            Console.WriteLine(
                "[TABLE OMML PLAN] "
                + string.Join(
                    "; ",
                    plan.Targets.Select((target, index) =>
                        $"#{index + 1} range={target.AbsoluteStart}:{target.AbsoluteEnd} dedicated={target.DedicatedDisplayParagraph} preserve={target.PreserveDisplayParagraphBoundary}")));

            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                prepared[target.Id] =
                    new PreparedWordBulkFormula
                    {
                        Run = new WordBulkRun
                        {
                            Id = target.Id,
                            IsFormula = true,
                            Latex = target.Latex,
                            DisplayMode =
                                target.DisplayMode,
                        },
                        Session =
                            CreateOmmlMathTypeAcceptanceSession(
                                mathMl,
                                target.DisplayMode,
                                numbered: false,
                                FormulaOleContract.WordOmmlMode),
                        MathMl = mathMl,
                    };
            }

            var result =
                service.ApplyLatexRedrawPlan(
                    plan,
                    prepared);
            AssertEqual(
                3,
                result.FormulaCount,
                "Table OMML redraw did not insert all three formulas.");
            AssertEqual(
                3,
                document.OMaths.Count,
                "Table OMML redraw created the wrong OMath count.");

            void AssertCellStructure(
                int row,
                string stage)
            {
                Release(cellRange);
                cellRange = null;
                Release(cell);
                cell = table.Cell(
                    row,
                    1);
                cellRange =
                    cell.Range.Duplicate;

                var flatOpc =
                    XDocument.Parse(
                        cellRange.WordOpenXML,
                        LoadOptions.PreserveWhitespace);
                var body =
                    WordDocumentXml.GetRequiredMainDocumentBody(
                        flatOpc);
                XNamespace w =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace m =
                    "http://schemas.openxmlformats.org/officeDocument/2006/math";

                var outerTable =
                    body.Elements(w + "tbl")
                        .SingleOrDefault()
                    ?? throw new InvalidDataException(
                        $"{stage}: cell wrapper has no outer table.");
                var outerCell =
                    outerTable
                        .Descendants(w + "tc")
                        .FirstOrDefault()
                    ?? throw new InvalidDataException(
                        $"{stage}: cell wrapper has no cell.");

                var directNestedTables =
                    outerCell
                        .Elements(w + "tbl")
                        .Count();
                var directParagraphs =
                    outerCell
                        .Elements(w + "p")
                        .Count();
                var mathCount =
                    outerCell
                        .Descendants(m + "oMath")
                        .Count();

                Word.OMath? diagnosticMath = null;
                Word.Range? diagnosticMathRange = null;
                try
                {
                    if (cellRange.OMaths.Count == 1)
                    {
                        diagnosticMath = cellRange.OMaths[1];
                        diagnosticMathRange = diagnosticMath.Range.Duplicate;
                    }
                    var rawCellText = cellRange.Text ?? string.Empty;
                    var codes = string.Join(
                        ",",
                        rawCellText.Select(character => ((int)character).ToString("X4")));
                    Console.WriteLine(
                        $"[TABLE OMML CELL {row}] {stage}: cell={cellRange.Start}:{cellRange.End}; math={(diagnosticMathRange is null ? "<none>" : diagnosticMathRange.Start + ":" + diagnosticMathRange.End)}; nestedTables={directNestedTables}; directParagraphs={directParagraphs}; maths={mathCount}; text='{rawCellText.Replace("\r", "<CR>").Replace("\a", "<CELL>")}'; codes={codes}.");
                }
                finally
                {
                    Release(diagnosticMathRange);
                    Release(diagnosticMath);
                }

                AssertEqual(
                    0,
                    directNestedTables,
                    $"{stage}: OMML redraw created a nested table in cell {row},1.");
                AssertEqual(
                    1,
                    directParagraphs,
                    $"{stage}: OMML redraw left extra paragraphs in cell {row},1.");
                AssertEqual(
                    1,
                    mathCount,
                    $"{stage}: cell {row},1 does not contain exactly one OMML formula.");
            }

            AssertCellStructure(
                1,
                "after redraw");
            AssertCellStructure(
                2,
                "after redraw");

            void AssertInlineMixedCell(string stage)
            {
                Release(cellRange);
                cellRange = null;
                Release(cell);
                cell = table.Cell(
                    1,
                    2);
                cellRange =
                    cell.Range.Duplicate;

                var flatOpc =
                    XDocument.Parse(
                        cellRange.WordOpenXML,
                        LoadOptions.PreserveWhitespace);
                var body =
                    WordDocumentXml.GetRequiredMainDocumentBody(
                        flatOpc);
                XNamespace w =
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace m =
                    "http://schemas.openxmlformats.org/officeDocument/2006/math";
                var outerTable =
                    body.Elements(w + "tbl")
                        .SingleOrDefault()
                    ?? throw new InvalidDataException(
                        $"{stage}: inline cell wrapper has no outer table.");
                var outerCell =
                    outerTable.Descendants(w + "tc")
                        .FirstOrDefault()
                    ?? throw new InvalidDataException(
                        $"{stage}: inline cell wrapper has no target cell.");
                var directNestedTables =
                    outerCell.Elements(w + "tbl").Count();
                var directParagraphs =
                    outerCell.Elements(w + "p").Count();
                var maths =
                    outerCell.Descendants(m + "oMath").Count();
                var visibleText =
                    string.Concat(
                        outerCell.Descendants(w + "t")
                            .Select(node => node.Value));

                var rawInlineCellText = cellRange.Text ?? string.Empty;
                var rawInlineCodes = string.Join(
                    ",",
                    rawInlineCellText.Select(character => ((int)character).ToString("X4")));
                Console.WriteLine(
                    $"[TABLE OMML INLINE] {stage}: cell={cellRange.Start}:{cellRange.End}; nestedTables={directNestedTables}; directParagraphs={directParagraphs}; maths={maths}; text='{visibleText}'; raw='{rawInlineCellText.Replace("\r", "<CR>").Replace("\a", "<CELL>")}'; codes={rawInlineCodes}.");
                for (var paragraphIndex = 1;
                     paragraphIndex <= cellRange.Paragraphs.Count;
                     paragraphIndex++)
                {
                    Word.Paragraph? diagnosticParagraph = null;
                    Word.Range? diagnosticParagraphRange = null;
                    try
                    {
                        diagnosticParagraph =
                            cellRange.Paragraphs[paragraphIndex];
                        diagnosticParagraphRange =
                            diagnosticParagraph.Range.Duplicate;
                        var paragraphText =
                            diagnosticParagraphRange.Text
                            ?? string.Empty;
                        var paragraphCodes =
                            string.Join(
                                ",",
                                paragraphText.Select(character => ((int)character).ToString("X4")));
                        Console.WriteLine(
                            $"[TABLE OMML INLINE P{paragraphIndex}] {stage}: range={diagnosticParagraphRange.Start}:{diagnosticParagraphRange.End}; text='{paragraphText.Replace("\r", "<CR>").Replace("\a", "<CELL>")}'; codes={paragraphCodes}; maths={diagnosticParagraphRange.OMaths.Count}.");
                    }
                    finally
                    {
                        Release(diagnosticParagraphRange);
                        Release(diagnosticParagraph);
                    }
                }

                AssertEqual(
                    0,
                    directNestedTables,
                    $"{stage}: inline OMML redraw created a nested table.");
                AssertEqual(
                    1,
                    directParagraphs,
                    $"{stage}: inline OMML redraw changed paragraph topology.");
                AssertEqual(
                    1,
                    maths,
                    $"{stage}: inline mixed cell does not contain exactly one OMath.");
                AssertTrue(
                    visibleText.IndexOf("left", StringComparison.Ordinal) >= 0
                    && visibleText.IndexOf("right", StringComparison.Ordinal) >= 0,
                    $"{stage}: inline OMML redraw lost surrounding prose.");
            }

            AssertInlineMixedCell(
                "after redraw");

            var redrawFirstRowHeight =
                ReadFirstRowRenderedHeight(table);
            Console.WriteLine(
                $"[TABLE OMML HEIGHT AFTER] firstRow={redrawFirstRowHeight:0.###}pt.");
            AssertTrue(
                redrawFirstRowHeight > 0f
                && redrawFirstRowHeight < 48f,
                "Simple table display OMML redraw expanded a row beyond a normal one-line formula height.");

            var outputPath =
                Path.Combine(
                    artifactRoot,
                    "Word-Latex-Redraw-OMML-Table.docx");
            document.SaveAs2(
                outputPath,
                Word.WdSaveFormat
                    .wdFormatXMLDocument,
                AddToRecentFiles: false);

            document.Close(
                Word.WdSaveOptions
                    .wdSaveChanges);
            Release(table);
            table = null;
            Release(document);
            document = null;

            document =
                application.Documents.Open(
                    outputPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
            document.Activate();
            table =
                document.Tables[1];

            AssertEqual(
                3,
                document.OMaths.Count,
                "Save/reopen lost a table OMML formula.");
            AssertCellStructure(
                1,
                "after reopen");
            AssertCellStructure(
                2,
                "after reopen");
            AssertInlineMixedCell(
                "after reopen");
            var reopenedFirstRowHeight =
                ReadFirstRowRenderedHeight(table);
            Console.WriteLine(
                $"[TABLE OMML HEIGHT REOPEN] firstRow={reopenedFirstRowHeight:0.###}pt.");
            AssertTrue(
                reopenedFirstRowHeight > 0f
                && reopenedFirstRowHeight < 48f,
                "Save/reopen expanded a simple table OMML row beyond one-line formula height.");

            Console.WriteLine(
                "[TABLE OMML REDRAW PASS] 2x4 table retained one paragraph and no nested table in both formula cells.");
        }
        finally
        {
            Release(probe);
            Release(cellRange);
            Release(cell);
            Release(table);
            Release(anchor);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions
                            .wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            try
            {
                QuitWordApplicationIfOwned(
                    application);
            }
            catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}

using System.Text;
using System.Xml.Linq;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportMathTypeSkeletonAcceptance(
        VisualTeXSessionClient client, string artifactRoot)
    {
        if (AttachActiveWord) throw new InvalidOperationException("Skeleton acceptance must create its own Word process.");
        Directory.CreateDirectory(artifactRoot);
        var twenty = CreateWordBulkImportMathTypeHundredSource(10);
        const string displayEdges = "\\[a+1\\]\n\nMiddle prose $b+2$ after inline.\n\n\\[c+3\\]";
        const string rich = "# Imported heading\n\n**Bold run** and *italic run* with $a+1$ plain tail.\n\n"
            + "- Bullet text $b+2$ after bullet.\n\n> Quoted prose.\n\n"
            + "```text\ncode one\ncode two\n```\n\n$$c+3$$\n\nFinal prose $d+4$ end.";
        const string chapters = "# 4 First chapter\n\n$$a+1$$\n\n# 12 Second chapter\n\n$$b+2$$";
        var cases = new[]
        {
            (Name: "body20", Table: false, Track: false, Source: twenty, Format: "latex", Numbered: false, Fault: false),
            (Name: "body20-tracked", Table: false, Track: true, Source: twenty, Format: "latex", Numbered: false, Fault: false),
            (Name: "table20", Table: true, Track: false, Source: twenty, Format: "latex", Numbered: false, Fault: false),
            (Name: "table20-tracked", Table: true, Track: true, Source: twenty, Format: "latex", Numbered: false, Fault: false),
            (Name: "display-edges", Table: false, Track: false, Source: displayEdges, Format: "latex", Numbered: false, Fault: false),
            (Name: "rich-numbered", Table: false, Track: false, Source: rich, Format: "markdown", Numbered: true, Fault: false),
            (Name: "body-fault-retry", Table: false, Track: false, Source: displayEdges, Format: "latex", Numbered: false, Fault: true),
            (Name: "table-fault-retry", Table: true, Track: false, Source: displayEdges, Format: "latex", Numbered: false, Fault: true),
            (Name: "manual-chapters", Table: false, Track: false, Source: chapters, Format: "markdown", Numbered: true, Fault: false),
            (Name: "numbered-fault-retry", Table: false, Track: false, Source: chapters, Format: "markdown", Numbered: true, Fault: true),
        };
        var selected = Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_SKELETON_CASE");
        var reports = new List<object>();
        foreach (var test in cases.Where(item => string.IsNullOrEmpty(selected) || item.Name == selected))
        {
            var output = Path.Combine(artifactRoot, test.Name);
            Directory.CreateDirectory(output);
            var previous = new Dictionary<string, string?>();
            void Set(string key, string? value)
            {
                if (!previous.ContainsKey(key)) previous[key] = Environment.GetEnvironmentVariable(key);
                Environment.SetEnvironmentVariable(key, value);
            }
            Word.Application? app = null;
            Word.Document? doc = null;
            ThisAddIn? addIn = null;
            Array custom = Array.Empty<object>();
            try
            {
                var parsed = WordBulkImportParser.Parse(test.Source,
                    test.Format == "latex" ? WordBulkSourceFormat.Latex : WordBulkSourceFormat.Markdown,
                    WordBulkFormulaObjectMode.MathType);
                var sourceFile = Path.Combine(output, "source.txt");
                File.WriteAllText(sourceFile, test.Source, new UTF8Encoding(false));
                Set("VISUALTEX_VSTO_BULK_SOURCE", null);
                Set("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourceFile);
                Set("VISUALTEX_VSTO_BULK_FORMAT", test.Format);
                Set("VISUALTEX_VSTO_BULK_OBJECT_MODE", "mathtype");
                Set("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS", test.Numbered ? "1" : "0");
                Set("VISUALTEX_VSTO_TRACE_RECOVERY_XML", "1");
                Set("VISUALTEX_WORD_HOOK_TRACE_PATH", Path.Combine(output, "word-hook.log"));
                Set("VISUALTEX_VSTO_TRACE_FORMAT_PERF", "1");
                app = CreateWordApplication(visible: test.Numbered);
                doc = app.Documents.Add(Visible: test.Numbered);
                doc.Activate();
                var content = doc.Content;
                try { content.Text = "OUTER-BEFORE\rHOST-PREFIX REPLACE_ME HOST-SUFFIX\rOUTER-AFTER\r"; }
                finally { Release(content); }
                if (test.Table)
                {
                    var at = doc.Range(13, 13);
                    Word.Table? table = null;
                    Word.Range? cellRange = null;
                    Word.Cell? cell = null;
                    try
                    {
                        table = doc.Tables.Add(at, 1, 2);
                        cell = table.Cell(1, 1); cellRange = cell.Range;
                        cellRange.End -= 1;
                        cellRange.Text = "HOST-PREFIX REPLACE_ME HOST-SUFFIX";
                        Release(cellRange); cellRange = null;
                        Release(cell); cell = table.Cell(1, 2); cellRange = cell.Range;
                        cellRange.End -= 1; cellRange.Text = "UNTOUCHED-CELL";
                        // Keep only one replacement token: the original body host
                        // is deliberately unrelated to the imported table cell.
                        var originalHost = doc.Content;
                        try
                        {
                            var lastToken = (originalHost.Text ?? "").LastIndexOf("REPLACE_ME", StringComparison.Ordinal);
                            var original = doc.Range(lastToken, lastToken + "REPLACE_ME".Length);
                            try { original.Text = "BODY-UNCHANGED"; } finally { Release(original); }
                        }
                        finally { Release(originalHost); }
                    }
                    finally { Release(cellRange); Release(cell); Release(table); Release(at); }
                }
                SelectSkeletonReplaceToken(app, doc);
                doc.TrackRevisions = test.Track;
                if (test.Numbered)
                    WordEquationNumbering.SetEquationNumberFormatPreference(doc, EquationNumberFormat.Heading1DotId);
                addIn = new ThisAddIn();
                addIn.OnConnection(app, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);
                var baseline = CaptureSkeletonFullSignature(doc, Path.Combine(output, "before.xml"));
                var originalTables = doc.Tables.Count;
                if (test.Fault)
                {
                    var faultLog = Path.Combine(output, "fault.log");
                    Set("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", faultLog);
                    Set("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE", "after-flat-opc");
                    addIn.OnBulkImport(new object());
                    var failed = false;
                    try { WaitForBulkImportCompletion(faultLog, TimeSpan.FromSeconds(90)); }
                    catch (InvalidOperationException) { failed = true; }
                    WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                    AssertTrue(failed, "Injected native insertion fault unexpectedly succeeded.");
                    var faultText = File.ReadAllText(faultLog);
                    AssertTrue(faultText.Contains("Injected MathType Flat OPC failure") && faultText.Contains("0x8007000E"),
                        "Injected fault did not reach the expected native Flat OPC stage/HRESULT.");
                    AssertEqual(baseline, CaptureSkeletonFullSignature(doc, Path.Combine(output, "fault-rollback.xml")),
                        "Bulk fault rollback did not restore the complete original document.");
                    AssertEqual(0, doc.InlineShapes.Count, "Bulk fault left a residual native OLE.");
                    AssertNoSkeletonLocators(doc);
                    Set("VISUALTEX_ACCEPTANCE_MATHTYPE_FAIL_STAGE", null);
                    SelectSkeletonReplaceToken(app, doc);
                }
                var log = Path.Combine(output, "import.log");
                Set("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", log);
                addIn.OnBulkImport(new object());
                var elapsed = WaitForBulkImportCompletion(log, TimeSpan.FromSeconds(120));
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));
                AssertEqual(test.Track, doc.TrackRevisions, "Bulk import changed Track Changes.");
                AssertEqual(originalTables, doc.Tables.Count, "Bulk import changed table count.");
                var after = CaptureSkeletonFullSignature(doc, Path.Combine(output, "after.xml"));
                AssertNoSkeletonLocators(doc);
                AssertSkeletonProseAndPlacement(doc, parsed, test.Table, test.Numbered);
                if (test.Numbered)
                    AssertSkeletonNativeNumberResults(doc, test.Source == chapters ? new[] { "(4.1)", "(12.1)" } : new[] { "(1.1)" });
                object one = 1;
                AssertTrue(doc.Undo(ref one), "The completed bulk import has no native Undo entry.");
                AssertEqual(baseline, CaptureSkeletonFullSignature(doc, Path.Combine(output, "undo.xml")),
                    "One native Undo did not restore the full original fixture.");
                AssertTrue(doc.Redo(ref one), "The completed bulk import has no native Redo entry.");
                AssertEqual(after, CaptureSkeletonFullSignature(doc, Path.Combine(output, "redo.xml")),
                    "One native Redo did not recreate the full imported fixture.");
                var metadata = ReadBulkMathTypeMetadata(new WordFormulaService(app), doc, test.Name);
                AssertEqual(parsed.FormulaCount, metadata.Count, "Skeleton formula count changed.");
                AssertEqual(metadata.Count, metadata.Select(item => item.FormulaId).Distinct().Count(), "Skeleton IDs are not unique.");
                var expected = parsed.Blocks.SelectMany(block => block.Runs).Where(run => run.IsFormula).ToArray();
                for (var i = 0; i < expected.Length; i++)
                {
                    AssertEqual(NormalizeMathTypePerformanceLatex(expected[i].Latex),
                        NormalizeMathTypePerformanceLatex(metadata[i].Latex), "Skeleton formula source/order changed.");
                    AssertEqual(expected[i].DisplayMode, metadata[i].DisplayMode, "Skeleton inline/display mode changed.");
                    AssertEqual(test.Numbered && expected[i].DisplayMode == "block", metadata[i].Numbered,
                        "Skeleton numbered state changed.");
                }
                var file = Path.Combine(output, "result.docx");
                doc.SaveAs2(file, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
                doc.Close(Word.WdSaveOptions.wdSaveChanges); Release(doc); doc = null;
                doc = app.Documents.Open(file, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                doc.Activate();
                AssertSkeletonProseAndPlacement(doc, parsed, test.Table, test.Numbered);
                AssertNoSkeletonLocators(doc);
                if (test.Numbered)
                    AssertSkeletonNativeNumberResults(doc, test.Source == chapters ? new[] { "(4.1)", "(12.1)" } : new[] { "(1.1)" });
                var reopened = ReadBulkMathTypeMetadata(new WordFormulaService(app), doc, test.Name + " reopen");
                AssertEqual(metadata.Count, reopened.Count, "Reopening lost editable formulas.");
                for (var i = 0; i < metadata.Count; i++)
                    AssertEqual(metadata[i].FormulaId, reopened[i].FormulaId, "Reopening changed a formula identity.");
                reports.Add(new { test.Name, Passed = true, test.Table, test.Track, test.Fault,
                    FormulaCount = parsed.FormulaCount, OperationMs = elapsed, StrictUndo = true, StrictRedo = true,
                    SaveReopen = true, NativeStorageVerified = true, Candidate = typeof(ThisAddIn).Assembly.Location });
                Console.WriteLine($"[SKELETON PASS] {test.Name} count={metadata.Count} operation={elapsed}ms strictUndo/Redo/save-reopen=true");
            }
            catch (Exception error)
            {
                reports.Add(new { test.Name, Passed = false, Error = error.ToString() });
                File.WriteAllText(Path.Combine(output, "error.txt"), error.ToString());
                throw;
            }
            finally
            {
                File.WriteAllText(Path.Combine(artifactRoot, "summary.json"),
                    System.Text.Json.JsonSerializer.Serialize(reports,
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
                if (addIn is not null)
                    try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
                try { doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(app); } catch { }
                Release(doc); Release(app); ForceComCleanup();
                foreach (var item in previous) Environment.SetEnvironmentVariable(item.Key, item.Value);
            }
        }
        AssertTrue(reports.Count > 0, "No skeleton acceptance case matched the filter.");
    }

    private static void AssertSkeletonNativeNumberResults(Word.Document doc, string[] expected)
    {
        var actual = new List<string>();
        var fields = doc.Fields;
        try
        {
            for (var index = 1; index <= fields.Count; index++)
            {
                var field = fields[index];
                Word.Range? code = null;
                Word.Range? result = null;
                try
                {
                    code = field.Code;
                    if (!(code.Text ?? "").TrimStart().StartsWith("MACROBUTTON MTPlaceRef", StringComparison.OrdinalIgnoreCase)) continue;
                    // Native MACROBUTTON display text is nested inside its
                    // instruction, not the normally empty outer Result range.
                    actual.Add(MathTypeEquationReferences.ReadVisibleNumberText(field).Trim());
                }
                finally { Release(result); Release(code); Release(field); }
            }
        }
        finally { Release(fields); }
        AssertTrue(actual.SequenceEqual(expected),
            "Native MathType chapter/sequence results differ: count=" + actual.Count + " values=" + string.Join(",", actual));
    }

    private static void SelectSkeletonReplaceToken(Word.Application app, Word.Document doc)
    {
        var content = doc.Content;
        try
        {
            var token = (content.Text ?? "").IndexOf("REPLACE_ME", StringComparison.Ordinal);
            if (token < 0) throw new InvalidDataException("The owned fixture replacement token was not found.");
            var selected = app.Selection;
            try { selected.SetRange(token, token + "REPLACE_ME".Length); }
            finally { Release(selected); }
        }
        finally { Release(content); }
    }

    private static string CaptureSkeletonFullSignature(Word.Document doc, string path)
    {
        var content = doc.Content;
        try
        {
            var xml = WordDocumentXml.Read(doc, content);
            var geometry = WordInlineObjectGeometry.Capture(content);
            File.WriteAllText(path, WordLocalEditSnapshot.NormalizedBody(xml, geometry), new UTF8Encoding(false));
            return WordLocalEditSnapshot.Signature(xml, geometry);
        }
        finally { Release(content); }
    }

    private static void AssertNoSkeletonLocators(Word.Document doc)
    {
        var bookmarks = doc.Bookmarks;
        try
        {
            for (var i = 1; i <= bookmarks.Count; i++)
            {
                var bookmark = bookmarks[i];
                try { AssertTrue(!bookmark.Name.StartsWith("VTBS", StringComparison.Ordinal), "Temporary skeleton locator survived."); }
                finally { Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
    }

    private static void AssertSkeletonProseAndPlacement(Word.Document doc,
        WordBulkImportDocument source, bool inTable, bool numbered)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var package = XDocument.Parse(WordDocumentXml.Read(doc));
        var body = package.Descendants(w + "body").Single();
        var visible = new StringBuilder();
        foreach (var node in body.Descendants().Where(node => !node.Ancestors().Any(
                     owner => owner.Name == w + "del" || owner.Name == w + "moveFrom")))
        {
            if (node.Name == w + "t") visible.Append(node.Value);
            else if (node.Name == w + "tab") visible.Append('\t');
            else if (node.Name == w + "br") visible.Append('\v');
        }
        var text = visible.ToString();
        if (source.Blocks.Any(block => block.Kind == WordBulkBlockKind.Heading))
        {
            var prefixParagraph = body.Descendants(w + "p").First(p => p.Descendants(w + "t").Any(t => t.Value.Contains("HOST-PREFIX")));
            AssertTrue(prefixParagraph.Element(w + "pPr")?.Element(w + "pStyle") is null,
                "An imported structural block changed the original prefix paragraph style.");
        }
        foreach (var run in source.Blocks.SelectMany(block => block.Runs).Where(run => !run.IsFormula && (run.Bold || run.Italic)))
        {
            var node = body.Descendants(w + "t").FirstOrDefault(t => t.Value.Contains(run.Text));
            AssertTrue(node is not null, "Styled native prose run is missing.");
            var properties = node!.Ancestors(w + "r").First().Element(w + "rPr");
            if (run.Bold) AssertTrue(properties?.Element(w + "b") is not null, "Bold prose style was lost.");
            if (run.Italic) AssertTrue(properties?.Element(w + "i") is not null, "Italic prose style was lost.");
        }
        foreach (var block in source.Blocks.Where(block => block.Kind == WordBulkBlockKind.Code))
        {
            var codeText = block.Runs[0].Text.Replace("\r\n", "\n").Replace('\r', '\n');
            var firstLine = codeText.Split('\n')[0];
            var paragraph = body.Descendants(w + "p").First(p => p.Descendants(w + "t").Any(t => t.Value.Contains(firstLine)));
            AssertTrue(paragraph.Descendants(w + "rFonts").Any(font => (string?)font.Attribute(w + "ascii") == "Consolas"),
                "Native code font was lost.");
            AssertEqual(codeText.Count(character => character == '\n'), paragraph.Descendants(w + "br").Count(),
                "Native code line breaks changed.");
        }
        AssertTrue(!text.Contains("REPLACE_ME"), "Bulk import retained the replaced current source text.");
        AssertTrue(!text.Contains("\uE000") && !text.Contains("\uE001"), "Bulk import left a placeholder character.");
        AssertTrue(text.Contains("OUTER-BEFORE") && text.Contains("OUTER-AFTER")
            && text.Contains("HOST-PREFIX") && text.Contains("HOST-SUFFIX"), "Bulk import damaged surrounding prose.");
        var cursor = text.IndexOf("HOST-PREFIX", StringComparison.Ordinal);
        foreach (var run in source.Blocks.SelectMany(block => block.Runs).Where(run => !run.IsFormula && run.Text.Length > 0))
        {
            var expected = run.Text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\n', '\v');
            var found = text.IndexOf(expected, cursor, StringComparison.Ordinal);
            AssertTrue(found >= cursor, $"Bulk prose is absent or out of order: {expected}");
            cursor = found + expected.Length;
        }
        AssertTrue(text.IndexOf("HOST-SUFFIX", cursor, StringComparison.Ordinal) >= cursor,
            "The original suffix moved before imported prose.");
        var shapes = doc.InlineShapes;
        try
        {
            AssertEqual(source.FormulaCount, shapes.Count, "Bulk native shape count changed.");
            for (var index = 1; index <= shapes.Count; index++)
            {
                var shape = shapes[index];
                Word.Range? range = null;
                Word.Tables? tables = null;
                try
                {
                    AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape), "Bulk formula is not genuine MathType OLE.");
                    range = shape.Range; tables = range.Tables;
                    AssertEqual(inTable ? 1 : 0, tables.Count, "Bulk OLE escaped its owned table/body context.");
                }
                finally { Release(tables); Release(range); Release(shape); }
            }
        }
        finally { Release(shapes); }
        if (inTable)
        {
            AssertTrue(text.Contains("UNTOUCHED-CELL") && text.Contains("BODY-UNCHANGED"),
                "Bulk import damaged a neighboring table cell or body paragraph.");
            var table = doc.Tables[1];
            try { AssertEqual(2, table.Columns.Count, "Bulk import changed the table column topology."); }
            finally { Release(table); }
        }
    }
}

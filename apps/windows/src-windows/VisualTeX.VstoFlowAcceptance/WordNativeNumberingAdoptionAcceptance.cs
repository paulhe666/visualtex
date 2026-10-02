using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunNativeNumberAdoptionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var rows = new List<object>();
        var oldPreference = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        try
        {
            foreach (var table in new[] { false, true })
            foreach (var tracking in new[] { false, true })
                RunNativeNumberAdoptionCase(artifactRoot, table, tracking, rows);
            RunManagedNativeNumberFormatRegression(artifactRoot, rows);
            File.WriteAllText(Path.Combine(artifactRoot, "summary.json"), JsonSerializer.Serialize(
                new { Time = DateTimeOffset.Now, Passed = rows.Count, Rows = rows }, new JsonSerializerOptions { WriteIndented = true }));
        }
        finally { WordEquationNumbering.SetDefaultEquationNumberFormatPreference(oldPreference); }
    }

    private static void RunNativeNumberAdoptionCase(string root, bool table, bool tracking, List<object> rows)
    {
        var label = (table ? "table" : "body") + (tracking ? "-tracked" : "-plain");
        var output = Path.Combine(root, label); Directory.CreateDirectory(output);
        Word.Application? application = null; Word.Document? document = null;
        Word.OMath? formula = null; Word.Range? insert = null; Word.Table? grid = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "Chapter\r\rRIGHT\r";
            Word.Range? heading = null;
            try
            {
                heading = document.Paragraphs[1].Range;
                object style = Word.WdBuiltinStyle.wdStyleHeading1;
                heading.set_Style(ref style);
                heading.ListFormat.ApplyNumberDefault();
            }
            finally { Release(heading); }
            insert = document.Paragraphs[2].Range;
            if (table)
            {
                grid = document.Tables.Add(insert, 1, 2);
                grid.Cell(1, 2).Range.Text = "UNTOUCHED CELL";
                Release(insert); insert = grid.Cell(1, 1).Range;
                insert.End--;
            }
            insert.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            const string mn = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            const string wn = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace m = mn; XNamespace w = wn;
            var math = XElement.Parse(WordOmmlConverter.BuildWordNativeNumberedOmml(
                "<m:oMath xmlns:m=\"" + mn + "\"><m:r><m:t>x+1</m:t></m:r></m:oMath>",
                "Equation", null, 1, "-", "7", "1"));
            // The input is a genuine Word #(literal) host, not a VT numbered ID.
            // Preserve structural proof but replace its number argument with text.
            var argument = math.Element(m + "eqArr")!.Element(m + "e")!.Element(m + "d")!.Element(m + "e")!;
            argument.ReplaceNodes(new XElement(m + "r", new XElement(m + "rPr", new XElement(m + "nor")), new XElement(m + "t", "1-7")));
            var seeded = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(document, insert, math.ToString(), display: true);
            Release(seeded);
            var candidates = document.OMaths;
            try { AssertEqual(1, candidates.Count, "Fixture must own exactly one native OMath."); formula = candidates[1]; }
            finally { Release(candidates); }
            formula.Type = Word.WdOMathType.wdOMathDisplay;
            var initial = ReadAdoptionLabel(formula);
            AssertEqual("1-7", initial, "Literal native input changed.");
            Word.Range? exact = null; Word.Range? number = null; Word.Bookmark? alias = null;
            try
            {
                exact = formula.Range; number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                alias = document.Bookmarks.Add("_RefNativeAdoption", number);
            }
            finally { Release(alias); Release(number); Release(exact); }
            document.TrackRevisions = tracking;
            File.WriteAllText(Path.Combine(output, "fixture-before.xml"), document.Content.WordOpenXML);
            Console.WriteLine($"Fixture {label}: type={formula.Type}, math={formula.Range.Text}");
            if (table)
            {
                var scope = document.Content;
                try
                {
                    var cursor = scope.Start;
                    for (var probe = 0; probe < 4; probe++)
                    {
                        var search = scope.Duplicate; search.SetRange(cursor, scope.End);
                        var retrieval = search.TextRetrievalMode; retrieval.IncludeFieldCodes = true; retrieval.IncludeHiddenText = true;
                        var find = search.Find; find.ClearFormatting(); find.Text = "#"; find.Forward = true; find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.MatchWildcards = false; find.MatchWholeWord = false;
                        if (!find.Execute()) { Release(find); Release(retrieval); Release(search); break; }
                        var expanded = search.Duplicate; expanded.Expand(Word.WdUnits.wdParagraph);
                        var physical = search.Paragraphs[1].Range;
                        Console.WriteLine($"Table discovery: cursor={cursor} scope={scope.Start}:{scope.End} match={search.Start}:{search.End} expanded={expanded.Start}:{expanded.End} physical={physical.Start}:{physical.End} stories={scope.StoryType}/{expanded.StoryType} inStory={expanded.InStory(scope)}");
                        var next = expanded.End;
                        var boundary = scope.Duplicate; boundary.SetRange(next, Math.Min(next + 2, scope.End));
                        Console.WriteLine("Following boundary=" + string.Join(",", (boundary.Text ?? string.Empty).Select(ch => ((int)ch).ToString("X4"))));
                        Release(boundary);
                        if (probe == 0) next++;
                        Release(physical); Release(expanded); Release(find); Release(retrieval); Release(search);
                        if (next <= cursor) break; cursor = next;
                    }
                }
                finally { Release(scope); }
            }
            var service = new WordFormulaService(application);
            var paragraphCount = document.Paragraphs.Count; var tableCount = document.Tables.Count;
            // Test recovery AFTER actual field replacement and native refresh,
            // not a preflight rejection. Only a diagnostic callback injects fault.
            using (var plan = WordNativeNumberingAdoptionPlan.Prepare(document, numberFormat: EquationNumberFormat.Heading1DotId))
            {
                AssertEqual(1, plan.Count, "The native adoption preflight missed its literal fixture.");
                var injected = false;
                try
                {
                    WordFormulaMutationTransaction.Execute(application, document, "Native field adoption terminal fault", () =>
                    {
                        plan.Apply();
                        throw new COMException("Injected native-number adoption terminal fault", unchecked((int)0x8007000E));
#pragma warning disable CS0162
                        return 0;
#pragma warning restore CS0162
                    }, verifyRollback: plan.VerifyRollback);
                }
                catch (Exception error) when (error.ToString().Contains("Injected native-number adoption terminal fault")) { injected = true; }
                AssertTrue(injected, "The terminal fault did not run.");
                AssertEqual("1-7", ReadAdoptionLabel(document.OMaths[1]), "Fault rollback did not restore the native literal.");
                AssertEqual(paragraphCount, document.Paragraphs.Count, "Fault changed paragraph topology.");
                AssertEqual(tableCount, document.Tables.Count, "Fault changed table ownership.");
            }
            var cases = new[] {
                (Format: EquationNumberFormat.Heading1DotId, Label: "1.1", Fields: 2),
                (Format: EquationNumberFormat.ContinuousId, Label: "1", Fields: 1),
                (Format: EquationNumberFormat.Heading1DashId, Label: "1-1", Fields: 2),
            };
            foreach (var item in cases)
            {
                Release(formula); formula = document.OMaths[1];
                exact = formula.Range;
                using var recovery = MathTypeLocalEditScope.Capture(document, exact);
                Release(exact); exact = null;
                var watch = Stopwatch.StartNew();
                AssertEqual(1, service.SetEquationNumberFormat(item.Format), "Format count differs.");
                watch.Stop();
                Release(formula); formula = document.OMaths[1];
                File.WriteAllText(Path.Combine(output, item.Format + "-observed.xml"), formula.Range.WordOpenXML);
                AssertEqual(item.Label, ReadAdoptionLabel(formula), "Native automatic-heading format result differs.");
                exact = formula.Range; number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                try
                {
                    AssertEqual(item.Fields, number.Fields.Count, "STYLEREF/SEQ native field topology differs.");
                    var range = document.Bookmarks["_RefNativeAdoption"].Range;
                    try { AssertEqual(number.Start, range.Start, "Alias start drifted."); AssertEqual(number.End, range.End, "Alias end drifted."); }
                    finally { Release(range); }
                    File.WriteAllText(Path.Combine(output, item.Format + ".xml"), exact.WordOpenXML);
                }
                finally { Release(number); number = null; Release(exact); exact = null; }
                document.Undo(1); recovery.VerifyRollback();
                document.Redo(1);
                Release(formula); formula = document.OMaths[1];
                AssertEqual(item.Label, ReadAdoptionLabel(formula), "Redo changed the number.");
                AssertEqual(tracking, document.TrackRevisions, "Revision mode changed.");
                AssertEqual(paragraphCount, document.Paragraphs.Count, "Format changed paragraphs.");
                AssertEqual(tableCount, document.Tables.Count, "Format changed table ownership.");
                Console.WriteLine($"PASS {label} {item.Format} {watch.ElapsedMilliseconds}ms Undo/Redo.");
            }
            var file = Path.Combine(output, "native-adoption.docx");
            document.SaveAs2(file, Word.WdSaveFormat.wdFormatXMLDocument);
            Release(formula); formula = null; document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(document); document = null;
            document = application.Documents.Open(file, ReadOnly: false, Visible: false);
            formula = document.OMaths[1];
            AssertEqual("1-1", ReadAdoptionLabel(formula), "Reopen changed automatic native numbering.");
            rows.Add(new { Name = label, TerminalFaultRollback = true, Formats = 3, UndoRedo = true, SaveReopen = true });
            File.WriteAllText(Path.Combine(root, "progress.json"), JsonSerializer.Serialize(rows));
        }
        finally
        {
            Release(formula); Release(grid); Release(insert);
            if (document is not null) { try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { } }
            Release(document); try { QuitWordApplicationIfOwned(application); } catch { } Release(application); ForceComCleanup();
        }
    }

    private static void RunManagedNativeNumberFormatRegression(
        string root,
        List<object> rows)
    {
        var output = Path.Combine(root, "managed-field-format");
        Directory.CreateDirectory(output);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.OMath? formula = null;
        Word.Range? insertion = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "Chapter\r\r";
            Word.Range? heading = null;
            try
            {
                heading = document.Paragraphs[1].Range;
                object style = Word.WdBuiltinStyle.wdStyleHeading1;
                heading.set_Style(ref style);
                heading.ListFormat.ApplyNumberDefault();
            }
            finally { Release(heading); }

            insertion = document.Paragraphs[2].Range;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            const string mn = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var prepared = WordOmmlConverter.BuildWordNativeNumberedOmml(
                "<m:oMath xmlns:m=\"" + mn + "\"><m:r><m:t>x+1</m:t></m:r></m:oMath>",
                "Equation",
                numberBookmarkName: null,
                restartHeadingLevel: 0,
                headingSeparator: string.Empty,
                initialSequenceResult: "1");
            var inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(
                document,
                insertion,
                prepared,
                display: true);
            Release(inserted);

            var maths = document.OMaths;
            try
            {
                AssertEqual(1, maths.Count, "Managed native regression fixture must own one equation.");
                formula = maths[1];
            }
            finally { Release(maths); }
            formula.Type = Word.WdOMathType.wdOMathDisplay;
            AssertEqual("1", ReadAdoptionLabel(formula), "Managed continuous fixture changed before rewrite.");

            var service = new WordFormulaService(application);
            AssertEqual(
                1,
                service.SetEquationNumberFormat(EquationNumberFormat.Heading1DashId),
                "Managed native field format rewrite count differs.");
            Release(formula);
            formula = document.OMaths[1];
            AssertEqual(
                "1-1",
                ReadAdoptionLabel(formula),
                "Managed native field did not adopt the requested heading format.");

            _ = service.SetEquationNumberFormat(EquationNumberFormat.Heading2DotId);
            Release(formula);
            formula = document.OMaths[1];
            _ = ReadAdoptionLabel(formula);
            _ = service.SetEquationNumberFormat(EquationNumberFormat.Heading2DashId);
            Release(formula);
            formula = document.OMaths[1];
            _ = ReadAdoptionLabel(formula);

            rows.Add(new
            {
                Name = "managed-field-format",
                ContinuousToHeading = true,
                Heading2Transitions = true,
            });
            File.WriteAllText(
                Path.Combine(output, "after.xml"),
                formula.Range.WordOpenXML);
        }
        finally
        {
            Release(formula);
            Release(insertion);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static string ReadAdoptionLabel(Word.OMath formula)
    {
        Word.Range? range = null;
        try { range = formula.Range; return WordOmmlConverter.ReadWordNativeEquationNumberText(range.WordOpenXML); }
        finally { Release(range); }
    }
}

using System.Text.Json;
using System.Runtime.InteropServices;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunNativeLiteralLifecycle(string root)
    {
        Directory.CreateDirectory(root);
        var previous = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        var observations = new List<object>();
        Word.Application? app = null;
        try
        {
            app = CreateWordApplication(visible: false);
            foreach (var table in new[] { false, true })
            foreach (var tracking in new[] { false, true })
            foreach (var level in new[] { 1, 2 })
            foreach (var separator in new[] { "-", "." })
            {
                var id = (table ? "table" : "body") + "-" + tracking + "-h" + level + (separator == "-" ? "-dash" : "-dot");
                var only = Environment.GetEnvironmentVariable("VISUALTEX_LITERAL_CASE");
                if (!string.IsNullOrWhiteSpace(only) && !string.Equals(only, id, StringComparison.Ordinal))
                    continue;
                var output = Path.Combine(root, id); Directory.CreateDirectory(output);
                Word.Document? doc = null; Word.Table? grid = null;
                try
                {
                    doc = app.Documents.Add();
                    doc.Content.Text = "4 Chapter\r4.2 Section\r\r";
                    for (var i = 1; i <= 2; i++)
                    {
                        var heading = doc.Paragraphs[i].Range;
                        try { object style = i == 1 ? Word.WdBuiltinStyle.wdStyleHeading1 : Word.WdBuiltinStyle.wdStyleHeading2; heading.set_Style(ref style); }
                        finally { Release(heading); }
                    }
                    var slot = doc.Paragraphs[3].Range;
                    try
                    {
                        slot.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        if (table)
                        {
                            grid = doc.Tables.Add(slot, 1, 2);
                            grid.Cell(1, 2).Range.Text = "UNTOUCHED";
                            Release(slot); slot = grid.Cell(1, 1).Range;
                            slot.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        }
                        InsertNativeNumberedFixture(doc, slot.Duplicate, "x+1", 0, "", "1", 1);
                    }
                    finally { Release(slot); }
                    var math = doc.OMaths[1]; var exact = math.Range;
                    try
                    {
                        var number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                        try { var bookmark = doc.Bookmarks.Add("_RefLiteralLifecycle", number); Release(bookmark); }
                        finally { Release(number); }
                    }
                    finally { Release(exact); Release(math); }
                    var reference = RootAppendParagraph(doc, "Reference ");
                    try
                    {
                        reference.End--; reference.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        var field = reference.Fields.Add(reference, Word.WdFieldType.wdFieldRef, "_RefLiteralLifecycle \\h", true);
                        try { field.Update(); } finally { Release(field); }
                    }
                    finally { Release(reference); }
                    doc.TrackRevisions = tracking;
                    var format = level == 1 ? (separator == "-" ? EquationNumberFormat.Heading1DashId : EquationNumberFormat.Heading1DotId)
                        : (separator == "-" ? EquationNumberFormat.Heading2DashId : EquationNumberFormat.Heading2DotId);
                    var service = new WordFormulaService(app);
                    service.SetEquationNumberFormat(format);
                    var original = (level == 1 ? "4" : "4.2") + separator + "1";
                    AssertLiteralLifecycleState(doc, original);
                    // Changing a heading must flow through ordinary Update,
                    // not a second explicit set-format call that could mask a
                    // separate broken literal-prefix writer.
                    var changedHeading = doc.Paragraphs[level].Range;
                    try { changedHeading.End--; changedHeading.Text = level == 1 ? "9 Chapter" : "9.7 Section"; }
                    finally { Release(changedHeading); }
                    var expected = (level == 1 ? "9" : "9.7") + separator + "1";
                    Console.WriteLine("START literal lifecycle " + id);
                    using var before = new RootDocumentEvidence(doc, Path.Combine(output, "before-update"));
                    using (var plan = WordNativeNumberingAdoptionPlan.Prepare(doc))
                    {
                        var injected = false;
                        try
                        {
                            WordFormulaMutationTransaction.Execute<int>(app, doc, "Native literal update terminal fault", () =>
                            {
                                plan.Apply();
                                throw new COMException("LITERAL_UPDATE_TERMINAL_FAULT", unchecked((int)0x8007000E));
                            }, verifyRollback: plan.VerifyRollback);
                        }
                        catch (COMException error) when (error.Message == "LITERAL_UPDATE_TERMINAL_FAULT") { injected = true; }
                        AssertTrue(injected, "The literal update fault was not reached or native rollback failed.");
                        before.VerifyRollback();
                        Console.WriteLine("VERIFY terminal rollback " + id);
                        AssertLiteralLifecycleState(doc, original);
                    }
                    service.UpdateEquationNumbers();
                    AssertLiteralLifecycleState(doc, expected);
                    File.WriteAllText(Path.Combine(output, "after-update.xml"), doc.Content.WordOpenXML);
                    AssertTrue(doc.Undo(1), "Literal prefix update did not own one Undo.");
                    before.VerifyRollback(); AssertLiteralLifecycleState(doc, original);
                    AssertTrue(doc.Redo(1), "Literal prefix update failed Redo.");
                    AssertLiteralLifecycleState(doc, expected);
                    var file = Path.Combine(output, "literal-lifecycle.docx");
                    doc.SaveAs2(file, Word.WdSaveFormat.wdFormatXMLDocument);
                    Release(grid); grid = null;
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); doc = null;
                    doc = app.Documents.Open(file, ReadOnly: false, Visible: false, AddToRecentFiles: false);
                    AssertLiteralLifecycleState(doc, expected);
                    AssertEqual(tracking, doc.TrackRevisions, "Literal update changed Track Changes.");
                    if (table) AssertTrue(doc.Content.Text.Contains("UNTOUCHED"), "Literal update lost adjacent cell.");
                    observations.Add(new { Id = id, TerminalFaultRollback = true, UpdateUndoRedo = true, SaveReopenReferences = true });
                    File.WriteAllText(Path.Combine(root, "progress.json"), JsonSerializer.Serialize(observations, new JsonSerializerOptions { WriteIndented = true }));
                    Console.WriteLine("PASS literal lifecycle " + id);
                }
                catch { try { File.WriteAllText(Path.Combine(output, "failed.xml"), doc!.Content.WordOpenXML); } catch { } throw; }
                finally { Release(grid); grid = null; try { doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { } Release(doc); }
            }
            AssertEqual(16, observations.Count, "Literal lifecycle matrix is incomplete.");
            File.WriteAllText(Path.Combine(root, "summary.json"), JsonSerializer.Serialize(observations, new JsonSerializerOptions { WriteIndented = true }));
        }
        finally { try { QuitWordApplicationIfOwned(app); } catch { } Release(app); WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previous); ForceComCleanup(); }
    }

    private static void AssertLiteralLifecycleState(Word.Document document, string expected)
    {
        var math = document.OMaths[1]; var range = math.Range;
        Word.Range? number = null; Word.Bookmark? bookmark = null; Word.Range? alias = null;
        try
        {
            AssertEqual(expected, WordOmmlConverter.ReadWordNativeEquationNumberText(range.WordOpenXML), "The literal prefix changed mathematical representation.");
            number = WordNativeOmmlNumbering.ResolveNumberContentRange(range);
            AssertEqual(expected, number.Text, "Live native number differs from literal display.");
            bookmark = document.Bookmarks["_RefLiteralLifecycle"]; alias = bookmark.Range;
            AssertTrue(alias.InStory(number) && alias.Start == number.Start && alias.End == number.End, "Literal number alias lost its physical owner.");
            var found = 0;
            using var fields = WordNumberingFieldCollection.Capture(document);
            for (var i = 1; i <= fields.Count; i++)
            {
                var field = fields[i]; Word.Range? code = null; Word.Range? result = null;
                try
                {
                    if (field.Type != Word.WdFieldType.wdFieldRef) continue;
                    code = field.Code;
                    if (!WordEquationReferenceFields.TryReadReferenceBookmark(code.Text, out var name) || name != "_RefLiteralLifecycle") continue;
                    result = field.Result; AssertEqual(expected, result.Text, "Native literal reference is stale."); found++;
                }
                finally { Release(result); Release(code); Release(field); }
            }
            if (found != 1)
            {
                // Diagnostic-only whole collection on this tiny owned fixture.
                // It independently distinguishes an omitted discovery result
                // from an actually deleted native field without changing view.
                var raw = document.Fields;
                try
                {
                    for (var i = 1; i <= raw.Count; i++)
                    {
                        var f = raw[i]; Word.Range? c = null; Word.Range? r = null;
                        try
                        {
                            c = f.Code; r = f.Result;
                            Console.WriteLine($"FIELD {i} type={f.Type} current={WordNumberingFieldCollection.IsCurrentField(f)} code={c.Start}:{c.End} [{c.Text}] result={r.Start}:{r.End} [{r.Text}]");
                            foreach (var span in new[] { c, r })
                            {
                                var revisions = span.Revisions;
                                try
                                {
                                    for (var j = 1; j <= revisions.Count; j++)
                                    {
                                        var revision = revisions[j]; var rr = revision.Range;
                                        try { Console.WriteLine($" REV span={span.Start}:{span.End} kind={revision.Type} range={rr.Start}:{rr.End} [{rr.Text}]"); }
                                        finally { Release(rr); Release(revision); }
                                    }
                                }
                                finally { Release(revisions); }
                            }
                        }
                        finally { Release(r); Release(c); Release(f); }
                    }
                }
                finally { Release(raw); }
                var paragraphCollection = document.Paragraphs;
                try
                {
                    for (var p = 1; p <= paragraphCollection.Count; p++)
                    {
                        var paragraph = paragraphCollection[p]; var pr = paragraph.Range; Word.Fields? pf = null;
                        try
                        {
                            pf = WordFormulaHost.GetLocalFields(pr);
                            if (pf.Count != 0)
                                Console.WriteLine($" PARAGRAPH {p} range={pr.Start}:{pr.End} fields={pf.Count} text=[{pr.Text}]");
                        }
                        finally { Release(pf); Release(pr); Release(paragraph); }
                    }
                }
                finally { Release(paragraphCollection); }

                foreach (var pattern in new[] { "REF", "^d REF", "^d", "^19" })
                {
                    var scope = document.Content; var cursor = scope.Start;
                    try
                    {
                        for (var pass = 0; pass < 8 && cursor < scope.End; pass++)
                        {
                            var search = scope.Duplicate; search.SetRange(cursor, scope.End);
                            var retrieval = search.TextRetrievalMode;
                            var find = search.Find;
                            Word.Range? unit = null; Word.Range? physical = null; Word.Fields? local = null;
                            try
                            {
                                retrieval.IncludeFieldCodes = true; retrieval.IncludeHiddenText = true;
                                find.ClearFormatting(); find.Text = pattern; find.Forward = true; find.Wrap = Word.WdFindWrap.wdFindStop;
                                find.Format = false; find.MatchWildcards = false; find.MatchCase = false; find.MatchWholeWord = pattern == "REF";
                                find.MatchSoundsLike = false; find.MatchAllWordForms = false;
                                if (!find.Execute()) { Console.WriteLine($"FIND [{pattern}] cursor={cursor} none"); break; }
                                unit = WordNumberingFieldCollection.CaptureSearchUnit(search);
                                var paragraphs = search.Paragraphs; var paragraph = paragraphs[1];
                                try { physical = paragraph.Range; } finally { Release(paragraph); Release(paragraphs); }
                                local = WordFormulaHost.GetLocalFields(unit);
                                Console.WriteLine($"FIND [{pattern}] cursor={cursor} match={search.Start}:{search.End} unit={unit.Start}:{unit.End} physical={physical.Start}:{physical.End} local={local.Count}");
                                for (var j = 1; j <= local.Count; j++)
                                {
                                    var f = local[j]; var c = f.Code;
                                    try { Console.WriteLine($" LOCAL {f.Type} code={c.Start}:{c.End} [{c.Text}] current={WordNumberingFieldCollection.IsCurrentField(f)}"); }
                                    finally { Release(c); Release(f); }
                                }
                                if (unit.End <= cursor) break;
                                cursor = unit.End;
                            }
                            finally { Release(local); Release(physical); Release(unit); Release(find); Release(retrieval); Release(search); }
                        }
                    }
                    finally { Release(scope); }
                }
            }
            AssertEqual(1, found, "Literal lifecycle reference disappeared.");
        }
        finally { Release(alias); Release(bookmark); Release(number); Release(range); Release(math); }
    }
}

using System.Diagnostics;
using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunNumberingDiscoveryDensity(string root)
    {
        Directory.CreateDirectory(root);
        var results = new List<object>();
        Word.Application? app = null;
        try
        {
            app = CreateWordApplication(visible: false);
            foreach (var paragraphs in new[] { 0, 128, 512 })
            {
                Word.Document? doc = null;
                try
                {
                    doc = app.Documents.Add();
                    doc.Content.Text = "TARGET VALUE\r" + string.Concat(Enumerable.Repeat("OLD unrelated body content\r", paragraphs)) + "\r";
                    var target = doc.Range(0, 12);
                    try { var bookmark = doc.Bookmarks.Add("_RefDiscoveryTarget", target); Release(bookmark); }
                    finally { Release(target); }
                    doc.TrackRevisions = true;
                    if (paragraphs > 0)
                    {
                        var scope = doc.Content; var find = scope.Find; var replacement = find.Replacement;
                        try
                        {
                            find.ClearFormatting(); replacement.ClearFormatting();
                            find.Text = "OLD"; replacement.Text = "NEW";
                            find.Forward = true; find.Wrap = Word.WdFindWrap.wdFindStop;
                            find.MatchCase = true; find.MatchWholeWord = true; find.MatchWildcards = false; find.Format = false;
                            AssertTrue(find.Execute(Replace: Word.WdReplace.wdReplaceAll), "Fixture did not generate unrelated revisions.");
                        }
                        finally { Release(replacement); Release(find); Release(scope); }
                    }
                    foreach (var n in Enumerable.Range(0, 4))
                    {
                        var slot = RootAppendParagraph(doc, "Field "); Word.Field? field = null;
                        try
                        {
                            slot.End--; slot.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                            field = slot.Fields.Add(slot, n % 2 == 0 ? Word.WdFieldType.wdFieldSequence : Word.WdFieldType.wdFieldRef,
                                n % 2 == 0 ? "Discovery \\* ARABIC" : "_RefDiscoveryTarget \\h", true);
                            field.Update();
                        }
                        finally { Release(field); Release(slot); }
                    }
                    // Test-only complete native collection on this owned fixture
                    // is the independent oracle, not a production fallback.
                    var expected = new List<string>(); var raw = doc.Fields;
                    try
                    {
                        for (var i = 1; i <= raw.Count; i++)
                        {
                            var field = raw[i]; var code = field.Code;
                            try { if (WordNumberingFieldCollection.IsCurrentField(field)) expected.Add(code.Start + ":" + code.Text); }
                            finally { Release(code); Release(field); }
                        }
                    }
                    finally { Release(raw); }
                    var before = new RootDocumentEvidence(doc, Path.Combine(root, "density-" + paragraphs));
                    var watch = Stopwatch.StartNew();
                    var actual = new List<string>();
                    using (var snapshot = WordNumberingFieldCollection.Capture(doc))
                    {
                        for (var i = 1; i <= snapshot.Count; i++)
                        {
                            var field = snapshot[i]; var code = field.Code;
                            try { actual.Add(code.Start + ":" + code.Text); }
                            finally { Release(code); Release(field); }
                        }
                    }
                    watch.Stop();
                    AssertEqual(string.Join("|", expected.OrderBy(x => x)), string.Join("|", actual.OrderBy(x => x)), "Bounded discovery missed or invented a native field.");
                    before.VerifyRollback(); before.Dispose();
                    var revisions = doc.Revisions;
                    try { results.Add(new { UnrelatedParagraphs = paragraphs, Revisions = revisions.Count, Fields = actual.Count, Milliseconds = watch.ElapsedMilliseconds, ExactOracle = true, NoMutation = true }); }
                    finally { Release(revisions); }
                    File.WriteAllText(Path.Combine(root, "progress.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
                    Console.WriteLine($"PASS field discovery density={paragraphs} fields={actual.Count} ms={watch.ElapsedMilliseconds}");
                }
                finally { try { doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { } Release(doc); }
            }
            File.WriteAllText(Path.Combine(root, "summary.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
        }
        finally { try { QuitWordApplicationIfOwned(app); } catch { } Release(app); ForceComCleanup(); }
    }
}

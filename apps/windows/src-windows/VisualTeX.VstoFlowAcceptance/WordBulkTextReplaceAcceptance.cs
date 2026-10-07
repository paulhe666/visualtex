using System.Diagnostics;
using System.Text.Json;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    // Real native Word Find.Execute timings. No formula conversion, package rewrite,
    // or manual per-match loop belongs inside the measured replacement operation.
    private static void RunOwnedBulkTextReplace(string root)
    {
        var inputs = Environment.GetEnvironmentVariable("VISUALTEX_REPLACE_FIXTURES")
            ?? throw new InvalidOperationException("Explicit owned fixture directory required.");
        var enabled = Environment.GetEnvironmentVariable("VISUALTEX_REPLACE_ADDIN") == "1";
        var useSelection = Environment.GetEnvironmentVariable("VISUALTEX_REPLACE_SELECTION") == "1";
        var installed = Environment.GetEnvironmentVariable("VISUALTEX_REPLACE_INSTALLED") == "1";
        AssertTrue(!(enabled && installed), "Never load two add-in instances.");
        if (enabled && string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_OFFICE_APP_DATA_ROOT")))
            throw new InvalidOperationException("Manual add-in requires isolated Companion environment.");
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", Path.Combine(root, "callbacks.log"));
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_SELECTION_PERF", "1");
        Environment.SetEnvironmentVariable("VISUALTEX_OLE_TRACE_PATH", Path.Combine(root, "native-ole.log"));
        var rows = new List<object>();
        var pattern = Environment.GetEnvironmentVariable("VISUALTEX_REPLACE_PATTERN") ?? "*.docx";
        foreach (var input in Directory.GetFiles(inputs, pattern).OrderBy(p => p, StringComparer.Ordinal))
        {
            Word.Application? app = null;
            VisualTeX.WordVsto.ThisAddIn? addin = null;
            Array custom = Array.Empty<object>();
            try
            {
                app = CreateWordApplication(visible: false);
                if (enabled)
                {
                    addin = new VisualTeX.WordVsto.ThisAddIn();
                    addin.OnConnection(app, Extensibility.ext_ConnectMode.ext_cm_AfterStartup, addin, ref custom);
                    addin.OnRibbonLoad(new AcceptanceRibbonProbe());
                }
                for (var round = 1; round <= 3; round++)
                {
                    var key = Path.GetFileNameWithoutExtension(input) + "-r" + round;
                    var copy = Path.Combine(root, key + ".docx"); File.Copy(input, copy);
                    Word.Document? doc = null; Word.Range? range = null; Word.Find? find = null;
                    try
                    {
                        Console.WriteLine("[REPLACE] opening " + key);
                        var clock = Stopwatch.StartNew();
                        doc = app.Documents.Open(copy, ReadOnly: false, AddToRecentFiles: false, Visible: useSelection);
                        if (useSelection) { app.Visible = true; doc.Activate(); }
                        var openMs = clock.Elapsed.TotalMilliseconds;
                        app.ScreenUpdating = true;
                        var before = doc.Content.Text;
                        var beforeObjects = doc.InlineShapes.Count; var beforeMaths = doc.OMaths.Count;
                        AssertTrue(!doc.TrackRevisions, "Fixture unexpectedly tracks changes.");
                        range = doc.Content;
                        if (useSelection)
                        {
                            var selection = app.Selection;
                            try { selection.SetRange(0, 0); find = selection.Find; }
                            finally { Release(selection); }
                        }
                        else find = range.Find;
                        find.ClearFormatting(); find.Replacement.ClearFormatting();
                        System.Windows.Forms.Application.DoEvents();
                        var events = 0;
                        Word.ApplicationEvents4_WindowSelectionChangeEventHandler count = _ => events++;
                        app.WindowSelectionChange += count;
                        try
                        {
                            WordDoubleClickHook.TraceMessage("replace-phase begin " + key);
                            clock.Restart();
                            var matched = find.Execute(FindText: " ", MatchCase: false, MatchWholeWord: false,
                                MatchWildcards: false, MatchSoundsLike: false, MatchAllWordForms: false,
                                Forward: true, Wrap: useSelection ? Word.WdFindWrap.wdFindContinue : Word.WdFindWrap.wdFindStop, Format: false,
                                ReplaceWith: "", Replace: Word.WdReplace.wdReplaceAll);
                            var replaceMs = clock.Elapsed.TotalMilliseconds;
                            WordDoubleClickHook.TraceMessage("replace-phase native-return " + key);
                            clock.Restart(); System.Windows.Forms.Application.DoEvents();
                            var callbacksMs = clock.Elapsed.TotalMilliseconds;
                            clock.Restart(); doc.Repaginate(); var repaginateMs = clock.Elapsed.TotalMilliseconds;
                            clock.Restart(); app.ScreenRefresh(); var refreshMs = clock.Elapsed.TotalMilliseconds;
                            var after = doc.Content.Text;
                            AssertEqual(before.Replace(" ", ""), after, "Native replace must remove exactly ASCII spaces.");
                            AssertEqual(beforeObjects, doc.InlineShapes.Count, "OLE inventory changed.");
                            AssertEqual(beforeMaths, doc.OMaths.Count, "OMML inventory changed.");
                            clock.Restart(); doc.Save(); var saveMs = clock.Elapsed.TotalMilliseconds;
                            var row = new { key, enabled, installed, useSelection, openMs, replaceMs, callbacksMs, repaginateMs, refreshMs, saveMs,
                                matched, spacesBefore=before.Count(c => c==' '), spacesAfter=after.Count(c => c==' '),
                                beforeObjects, beforeMaths, events, passed=true, wordVersion=app.Version };
                            rows.Add(row);
                            File.AppendAllText(Path.Combine(root, "samples.ndjson"), JsonSerializer.Serialize(row) + "\n");
                            Console.WriteLine($"[REPLACE] {key} spaces={row.spacesBefore} replace={replaceMs:F1}ms layout={repaginateMs:F1}ms events={events}");
                        }
                        finally { app.WindowSelectionChange -= count; }
                    }
                    finally
                    {
                        Release(find); Release(range);
                        if (doc is not null) { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); }
                    }
                }
            }
            finally
            {
                if (addin is not null) addin.OnDisconnection(Extensibility.ext_DisconnectMode.ext_dm_UserClosed, ref custom);
                QuitWordApplicationIfOwned(app); Release(app); ForceComCleanup();
            }
        }
        File.WriteAllText(Path.Combine(root, "results.json"), JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented=true }));
    }
}

using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordStagedOmmlRollbackAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixture = Path.GetFullPath(Path.Combine("docs", "remediation-3d207d7", "evidence",
            "perf100-fix9b-omml-vt-source-copy.docx"));
        var originalHash = System.Security.Cryptography.SHA256.Create();
        string HashFile(string path)
        {
            using var stream = File.OpenRead(path);
            return Convert.ToBase64String(originalHash.ComputeHash(stream));
        }
        var fixtureHash = HashFile(fixture);
        var environmentNames = new[]
        {
            "VISUALTEX_EXPERIMENTAL_STAGED_OMML_SOURCE_BATCH",
            "VISUALTEX_EXPERIMENTAL_DEFERRED_VT_BATCH_WRITER",
            "VISUALTEX_EXPERIMENTAL_DELETE_PROVEN_OMML_TABLE",
            "VISUALTEX_ACCEPTANCE_FAIL_AFTER_SOURCE_STAGING",
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
            "VISUALTEX_WORD_HOOK_TRACE_PATH",
        };
        var previous = environmentNames.ToDictionary(name => name, Environment.GetEnvironmentVariable);
        var previewRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX", "office", "temp", "staged-source-rollback-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(previewRoot);
        var png = Path.Combine(previewRoot, "rollback-test.png");
        var svg = Path.Combine(previewRoot, "rollback-test.svg");
        WriteAcceptancePng(png, "rollback fixture", 260, 96);
        File.WriteAllText(svg, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-size=\"48\">x+1</text></svg>");
        var emf = OfficeOlePreview.CreateVectorEmfFromSvg(svg, 260, 96);
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            Environment.SetEnvironmentVariable(environmentNames[0], "1");
            Environment.SetEnvironmentVariable(environmentNames[1], "0");
            Environment.SetEnvironmentVariable(environmentNames[2], "0");
            if (AttachActiveWord) throw new InvalidOperationException("Rollback acceptance must never attach an active user Word.");
            application = CreateFreshRollbackAutomationWord(artifactRoot);
            foreach (var failurePoint in new[] { "after-source-stage", "after-two-targets" })
            {
                var path = Path.Combine(artifactRoot, failurePoint + ".docx");
                CopyNumberedOmmlPerformanceFixture(fixture, path, 3);
                document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                document.Activate();
                AssertEqual(3, document.OMaths.Count, "Offline rollback fixture lost OMaths while opening.");
                AssertEqual(document.FullName, application.ActiveDocument.FullName, "Rollback source is not the active owned document.");
                var service = new WordFormulaService(application);
                var plan = service.CaptureFormulaFormatConversionPlan(true,
                    FormulaOleContract.WordOmmlMode, FormulaOleContract.NativeOleMode);
                AssertEqual(3, plan.Targets.Count, "Rollback fixture did not contain three sources.");
                var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
                foreach (var target in plan.Targets)
                {
                    var mathMl = target.SourceMathMl ?? throw new InvalidDataException("Missing source MathML.");
                    prepared.Add(target.Id, new PreparedWordBulkFormula
                    {
                        Run = new WordBulkRun { Id = target.Id, IsFormula = true, Latex = target.Latex, DisplayMode = target.DisplayMode },
                        Session = CreateSimpleFormatTargetSession(target, FormulaOleContract.NativeOleMode, mathMl),
                        MathMl = mathMl,
                        PngPath = png,
                        EmfPath = emf,
                    });
                }
                var bodyBefore = CaptureDeepRollbackBodySignature(document);
                var textBefore = document.Content.Text;
                var bookmarksBefore = CaptureDeepRollbackBookmarks(document);
                var metadataBefore = CaptureDeepRollbackCustomXml(document);
                var variablesBefore = CaptureDeepRollbackVariables(document);
                var undoBefore = CaptureDeepRollbackUndoHistory(application);
                var paragraphsBefore = document.Paragraphs.Count;
                var fieldsBefore = document.Fields.Count;
                var sourceIdsBefore = plan.Targets.OrderBy(target => target.SourceStart)
                    .Select(target => target.SourceFormulaId + "|" + target.Latex + "|" + target.Numbered).ToArray();
                var trace = Path.Combine(artifactRoot, failurePoint + ".log");
                Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", trace);
                Environment.SetEnvironmentVariable("VISUALTEX_ACCEPTANCE_FAIL_AFTER_SOURCE_STAGING",
                    failurePoint == "after-source-stage" ? "1" : null);
                Environment.SetEnvironmentVariable("VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                    failurePoint == "after-two-targets" ? plan.Targets.OrderBy(target => target.SourceStart).First().SourceFormulaId : null);
                Exception? failure = null;
                try { service.ApplyFormulaFormatConversionPlan(plan, prepared); }
                catch (Exception error) { failure = error; }
                finally
                {
                    Environment.SetEnvironmentVariable("VISUALTEX_ACCEPTANCE_FAIL_AFTER_SOURCE_STAGING", null);
                    Environment.SetEnvironmentVariable("VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE", null);
                }
                var traceText = File.ReadAllText(trace);
                var restored = service.CaptureFormulaFormatConversionPlan(true,
                    FormulaOleContract.WordOmmlMode, FormulaOleContract.NativeOleMode);
                var checks = new Dictionary<string, bool>
                {
                    ["stageReallyUsed"] = traceText.Contains("format-conversion-source-staging-complete count=3"),
                    ["faultReached"] = failure is not null && failure is not AggregateException
                        && failure.ToString().Contains(failurePoint == "after-source-stage"
                            ? "Injected failure after the proven OMML source stage."
                            : "Injected format-conversion failure after deleting the source host."),
                    ["partialTargetsExercised"] = failurePoint != "after-two-targets"
                        || traceText.Split(new[] { "format-conversion-bookmarked-visualtex-target-stable" }, StringSplitOptions.None).Length - 1 == 2,
                    ["body"] = bodyBefore == CaptureDeepRollbackBodySignature(document),
                    ["text"] = textBefore == document.Content.Text,
                    ["bookmarks"] = bookmarksBefore.SequenceEqual(CaptureDeepRollbackBookmarks(document)),
                    ["metadata"] = metadataBefore.SequenceEqual(CaptureDeepRollbackCustomXml(document)),
                    ["variables"] = variablesBefore.SequenceEqual(CaptureDeepRollbackVariables(document)),
                    ["undoBoundary"] = undoBefore.SequenceEqual(CaptureDeepRollbackUndoHistory(application)),
                    ["sourceSemantics"] = sourceIdsBefore.SequenceEqual(restored.Targets.OrderBy(target => target.SourceStart)
                        .Select(target => target.SourceFormulaId + "|" + target.Latex + "|" + target.Numbered)),
                    ["counts"] = document.OMaths.Count == 3 && document.InlineShapes.Count == 0
                        && document.Tables.Count == 3 && CountVisualTeXNumberingBookmarkTriples(document) == 3,
                    ["paragraphs"] = document.Paragraphs.Count == paragraphsBefore,
                    ["fields"] = document.Fields.Count == fieldsBefore,
                    ["sourceFileUnchanged"] = fixtureHash == HashFile(fixture),
                };
                var report = new { failurePoint, passed = checks.Values.All(value => value), checks, failure = failure?.ToString() };
                File.WriteAllText(Path.Combine(artifactRoot, failurePoint + "-result.json"),
                    System.Text.Json.JsonSerializer.Serialize(report, new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
                Console.WriteLine("[STAGED OMML ROLLBACK] " + failurePoint + ": "
                    + string.Join(", ", checks.Select(pair => pair.Key + "=" + pair.Value)));
                if (!report.passed) throw new InvalidDataException("Staged rollback failed: "
                    + string.Join(", ", checks.Where(pair => !pair.Value).Select(pair => pair.Key)));
                document.Save();
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = application.Documents.Open(path, ReadOnly: false, AddToRecentFiles: false, Visible: false);
                AssertEqual(3, document.OMaths.Count, "Rollback save/reopen lost source formulas.");
                AssertEqual(0, document.InlineShapes.Count, "Rollback save/reopen left target OLEs.");
                AssertEqual(3, CountVisualTeXNumberingBookmarkTriples(document), "Rollback save/reopen lost numbering.");
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;
            }
        }
        finally
        {
            foreach (var pair in previous) Environment.SetEnvironmentVariable(pair.Key, pair.Value);
            if (document is not null) try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            originalHash.Dispose();
            ForceComCleanup();
        }
    }
}

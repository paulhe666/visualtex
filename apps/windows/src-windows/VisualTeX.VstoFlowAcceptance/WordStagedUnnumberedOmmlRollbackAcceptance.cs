using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordStagedUnnumberedOmmlRollbackAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var source = Path.GetFullPath(Path.Combine(
            "docs", "remediation-3d207d7", "evidence",
            "baseline-body100-s0-omml-audit.xml"));

        using var sha256 = System.Security.Cryptography.SHA256.Create();
        string HashFile(string path)
        {
            using var stream = File.OpenRead(path);
            return Convert.ToBase64String(sha256.ComputeHash(stream));
        }
        var sourceHash = HashFile(source);

        var environmentNames = new[]
        {
            "VISUALTEX_EXPERIMENTAL_STAGED_OMML_SOURCE_BATCH",
            "VISUALTEX_EXPERIMENTAL_DEFERRED_VT_BATCH_WRITER",
            "VISUALTEX_EXPERIMENTAL_DELETE_PROVEN_OMML_TABLE",
            "VISUALTEX_ACCEPTANCE_FAIL_AFTER_SOURCE_STAGING",
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
            "VISUALTEX_WORD_HOOK_TRACE_PATH",
        };
        var previous = environmentNames.ToDictionary(
            name => name, Environment.GetEnvironmentVariable);

        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX", "office", "temp",
            "staged-unnumbered-rollback-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(previewRoot);
        var png = Path.Combine(previewRoot, "rollback-test.png");
        var svg = Path.Combine(previewRoot, "rollback-test.svg");
        WriteAcceptancePng(png, "rollback fixture", 260, 96);
        File.WriteAllText(
            svg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-size=\"48\">x+1</text></svg>");
        var emf = OfficeOlePreview.CreateVectorEmfFromSvg(svg, 260, 96);

        Word.Application? application = null;
        Word.Document? document = null;

        void PromoteFixtureToManagedUnnumbered(Word.Document targetDocument)
        {
            var customXmlCountBefore =
                CaptureDeepRollbackCustomXml(targetDocument).Length;
            Word.Bookmarks? bookmarks = null;
            Word.OMaths? maths = null;
            var ranges = new List<Word.Range>();
            try
            {
                bookmarks = targetDocument.Bookmarks;
                var staleNames = new List<string>();
                for (var index = 1; index <= bookmarks.Count; index++)
                {
                    Word.Bookmark? bookmark = null;
                    try
                    {
                        bookmark = bookmarks[index];
                        if ((bookmark.Name ?? string.Empty).StartsWith(
                                WordOmmlFormulaStore.BookmarkPrefix,
                                StringComparison.OrdinalIgnoreCase))
                            staleNames.Add(bookmark.Name);
                    }
                    finally { Release(bookmark); }
                }
                foreach (var name in staleNames)
                {
                    Word.Bookmark? bookmark = null;
                    try
                    {
                        if (!bookmarks.Exists(name)) continue;
                        bookmark = bookmarks[name];
                        bookmark.Delete();
                    }
                    finally { Release(bookmark); }
                }

                maths = targetDocument.OMaths;
                for (var index = 1; index <= maths.Count; index++)
                {
                    Word.OMath? math = null;
                    Word.Range? range = null;
                    try
                    {
                        math = maths[index];
                        range = math.Range.Duplicate;
                        ranges.Add(range);
                        range = null;
                    }
                    finally
                    {
                        Release(range);
                        Release(math);
                    }
                }
            }
            finally
            {
                Release(maths);
                Release(bookmarks);
            }

            foreach (var range in ranges)
            {
                Word.Bookmark? identity = null;
                try
                {
                    var metadata = WordOmmlNativeSource.CreateForNative(
                        targetDocument,
                        range);
                    if (metadata.Numbered)
                        throw new InvalidDataException(
                            "Managed rollback fixture unexpectedly became numbered.");
                    identity = WordOmmlFormulaStore.Wrap(
                        targetDocument,
                        range,
                        metadata,
                        replaceExisting: true);
                    WordOmmlFormulaStore.Save(
                        targetDocument,
                        metadata);
                }
                finally
                {
                    Release(identity);
                    Release(range);
                }
            }

            AssertEqual(
                customXmlCountBefore + 3,
                CaptureDeepRollbackCustomXml(targetDocument).Length,
                "Managed rollback fixture did not add one metadata part per OMath.");
        }

        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_EXPERIMENTAL_STAGED_OMML_SOURCE_BATCH", "1");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_EXPERIMENTAL_DEFERRED_VT_BATCH_WRITER", "0");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_EXPERIMENTAL_DELETE_PROVEN_OMML_TABLE", "0");

            if (AttachActiveWord)
                throw new InvalidOperationException(
                    "Unnumbered rollback acceptance must never attach an active user Word.");

            application = CreateFreshRollbackAutomationWord(artifactRoot);

            foreach (var failurePoint in new[]
                     {
                         "after-source-stage",
                         "after-two-targets",
                     })
            {
                var path = Path.Combine(
                    artifactRoot,
                    "unnumbered-" + failurePoint + ".docx");
                MaterializeFlatOpcOmmlPerformanceFixture(source, path, 3);

                document = application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
                document.Activate();

                AssertEqual(
                    3,
                    document.OMaths.Count,
                    "Unnumbered rollback fixture lost OMaths while opening.");
                AssertEqual(
                    0,
                    document.InlineShapes.Count,
                    "Unnumbered rollback fixture unexpectedly contains OLEs.");
                AssertEqual(
                    0,
                    document.Tables.Count,
                    "Unnumbered rollback fixture unexpectedly contains tables.");

                PromoteFixtureToManagedUnnumbered(document);
                document.Save();

                var service = new WordFormulaService(application);
                var plan = service.CaptureFormulaFormatConversionPlan(
                    true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
                AssertEqual(
                    3,
                    plan.Targets.Count,
                    "Unnumbered rollback fixture did not contain three sources.");
                AssertTrue(
                    plan.Targets.All(target => !target.Numbered),
                    "Unnumbered rollback fixture captured a numbered source.");

                var prepared =
                    new Dictionary<string, PreparedWordBulkFormula>(
                        StringComparer.Ordinal);
                foreach (var target in plan.Targets)
                {
                    var mathMl = target.SourceMathMl
                        ?? throw new InvalidDataException(
                            "Missing unnumbered rollback source MathML.");
                    prepared.Add(
                        target.Id,
                        new PreparedWordBulkFormula
                        {
                            Run = new WordBulkRun
                            {
                                Id = target.Id,
                                IsFormula = true,
                                Latex = target.Latex,
                                DisplayMode = target.DisplayMode,
                            },
                            Session = CreateSimpleFormatTargetSession(
                                target,
                                FormulaOleContract.NativeOleMode,
                                mathMl),
                            MathMl = mathMl,
                            PngPath = png,
                            EmfPath = emf,
                        });
                }

                string[] SourceSemantics(
                    WordFormulaFormatConversionPlan captured) =>
                    captured.Targets
                        .OrderBy(target => target.SourceStart)
                        .Select(target =>
                            target.SourceFormulaId
                            + "|" + target.SourceStart
                            + "|" + target.SourceObjectId
                            + "|" + target.DisplayMode
                            + "|" + target.Numbered
                            + "|" + target.SourceWithinTable
                            + "|" + target.Latex)
                        .ToArray();

                var bodyBefore =
                    CaptureDeepRollbackBodySignature(document);
                var textBefore = document.Content.Text;
                var bookmarksBefore =
                    CaptureDeepRollbackBookmarks(document);
                var metadataBefore =
                    CaptureDeepRollbackCustomXml(document);
                var variablesBefore =
                    CaptureDeepRollbackVariables(document);
                var undoBefore =
                    CaptureDeepRollbackUndoHistory(application);
                var paragraphsBefore = document.Paragraphs.Count;
                var fieldsBefore = document.Fields.Count;
                var sourceSemanticsBefore = SourceSemantics(plan);

                var trace = Path.Combine(
                    artifactRoot,
                    "unnumbered-" + failurePoint + ".log");
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH",
                    trace);
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_ACCEPTANCE_FAIL_AFTER_SOURCE_STAGING",
                    failurePoint == "after-source-stage" ? "1" : null);
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                    failurePoint == "after-two-targets"
                        ? plan.Targets
                            .OrderBy(target => target.SourceStart)
                            .First()
                            .SourceFormulaId
                        : null);

                Exception? failure = null;
                try
                {
                    service.ApplyFormulaFormatConversionPlan(
                        plan,
                        prepared);
                }
                catch (Exception error)
                {
                    failure = error;
                }
                finally
                {
                    Environment.SetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_FAIL_AFTER_SOURCE_STAGING",
                        null);
                    Environment.SetEnvironmentVariable(
                        "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                        null);
                }

                var traceText = File.ReadAllText(trace);
                var restored = service.CaptureFormulaFormatConversionPlan(
                    true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);

                var checks = new Dictionary<string, bool>
                {
                    ["stageReallyUsed"] =
                        traceText.Contains(
                            "format-conversion-source-staging-complete count=3")
                        && traceText.Contains("numbered=False"),
                    ["faultReached"] =
                        failure is not null
                        && failure is not AggregateException
                        && failure.ToString().Contains(
                            failurePoint == "after-source-stage"
                                ? "Injected failure after the proven OMML source stage."
                                : "Injected format-conversion failure after deleting the source host."),
                    ["partialTargetsExercised"] =
                        failurePoint != "after-two-targets"
                        || traceText.Split(
                                new[]
                                {
                                    "format-conversion-bookmarked-visualtex-target-stable",
                                },
                                StringSplitOptions.None)
                            .Length
                            - 1
                            == 2,
                    ["body"] =
                        bodyBefore
                        == CaptureDeepRollbackBodySignature(document),
                    ["text"] = textBefore == document.Content.Text,
                    ["bookmarks"] =
                        bookmarksBefore.SequenceEqual(
                            CaptureDeepRollbackBookmarks(document)),
                    ["metadata"] =
                        metadataBefore.SequenceEqual(
                            CaptureDeepRollbackCustomXml(document)),
                    ["variables"] =
                        variablesBefore.SequenceEqual(
                            CaptureDeepRollbackVariables(document)),
                    ["undoBoundary"] =
                        undoBefore.SequenceEqual(
                            CaptureDeepRollbackUndoHistory(application)),
                    ["sourceSemantics"] =
                        sourceSemanticsBefore.SequenceEqual(
                            SourceSemantics(restored)),
                    ["counts"] =
                        document.OMaths.Count == 3
                        && document.InlineShapes.Count == 0
                        && document.Tables.Count == 0
                        && CountVisualTeXNumberingBookmarkTriples(document)
                        == 0,
                    ["paragraphs"] =
                        document.Paragraphs.Count == paragraphsBefore,
                    ["fields"] =
                        document.Fields.Count == fieldsBefore,
                    ["sourceFileUnchanged"] =
                        sourceHash == HashFile(source),
                };

                var report = new
                {
                    failurePoint,
                    passed = checks.Values.All(value => value),
                    checks,
                    failure = failure?.ToString(),
                };
                File.WriteAllText(
                    Path.Combine(
                        artifactRoot,
                        "unnumbered-" + failurePoint + "-result.json"),
                    System.Text.Json.JsonSerializer.Serialize(
                        report,
                        new System.Text.Json.JsonSerializerOptions
                        {
                            WriteIndented = true,
                        }));

                Console.WriteLine(
                    "[STAGED UNNUMBERED OMML ROLLBACK] "
                    + failurePoint
                    + ": "
                    + string.Join(
                        ", ",
                        checks.Select(pair =>
                            pair.Key + "=" + pair.Value)));

                if (!report.passed)
                    throw new InvalidDataException(
                        "Staged unnumbered rollback failed: "
                        + string.Join(
                            ", ",
                            checks
                                .Where(pair => !pair.Value)
                                .Select(pair => pair.Key)));

                document.Save();
                document.Close(
                    Word.WdSaveOptions.wdSaveChanges);
                Release(document);

                document = application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
                AssertEqual(
                    3,
                    document.OMaths.Count,
                    "Unnumbered rollback save/reopen lost source formulas.");
                AssertEqual(
                    0,
                    document.InlineShapes.Count,
                    "Unnumbered rollback save/reopen left target OLEs.");
                AssertEqual(
                    0,
                    document.Tables.Count,
                    "Unnumbered rollback save/reopen created a table.");
                AssertTrue(
                    bookmarksBefore.SequenceEqual(
                        CaptureDeepRollbackBookmarks(document)),
                    "Unnumbered rollback save/reopen changed source bookmarks.");
                AssertTrue(
                    metadataBefore.SequenceEqual(
                        CaptureDeepRollbackCustomXml(document)),
                    "Unnumbered rollback save/reopen changed source metadata.");

                document.Close(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;
            }
        }
        finally
        {
            foreach (var pair in previous)
                Environment.SetEnvironmentVariable(
                    pair.Key,
                    pair.Value);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions.wdDoNotSaveChanges);
                }
                catch
                {
                }
            }
            Release(document);
            try
            {
                QuitWordApplicationIfOwned(application);
            }
            catch
            {
            }
            Release(application);
            ForceComCleanup();
            try
            {
                Directory.Delete(previewRoot, recursive: true);
            }
            catch
            {
            }
        }
    }
}

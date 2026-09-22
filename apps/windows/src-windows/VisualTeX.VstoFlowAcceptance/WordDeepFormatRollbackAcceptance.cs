using Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordDeepFormatRollbackAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixturePath = Path.GetFullPath(Path.Combine(
            "docs", "remediation-3d207d7", "evidence",
            "native-mt-r25-source-original-mathtype.docx"));
        if (!File.Exists(fixturePath))
            throw new FileNotFoundException(
                "The fixed five-equation native MathType fixture is required for deep rollback acceptance.",
                fixturePath);
        var workingPath = Path.Combine(artifactRoot, "deep-finalize-rollback-source.docx");
        File.Copy(fixturePath, workingPath, overwrite: false);

        var oldStage = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_STAGE");
        var oldDeleteFailure = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
        var oldTrace = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var tracePath = Path.Combine(artifactRoot, "deep-finalize-rollback-word-hook.log");

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            if (AttachActiveWord)
                throw new InvalidOperationException(
                    "Deep format/table acceptance must never attach an active user Word.");
            application = CreateFreshAcceptanceAutomationWord(artifactRoot);
            document = application.Documents.Open(
                workingPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            var service = new WordFormulaService(application);

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(5, plan.Targets.Count,
                "Deep rollback fixture did not expose exactly five native MathType formulas.");
            AssertEqual(5, CountMathTypeOleShapes(document),
                "Deep rollback fixture did not contain exactly five Equation.DSMT4 objects.");
            AssertEqual(0, document.OMaths.Count,
                "Deep rollback fixture unexpectedly contains OMML before conversion.");
            AssertEqual(1, document.Tables.Count,
                "Deep rollback fixture must begin with exactly one user table.");
            AssertEqual(2, document.Tables[1].Rows.Count,
                "Deep rollback fixture user table row count changed.");
            AssertEqual(2, document.Tables[1].Columns.Count,
                "Deep rollback fixture user table column count changed.");
            var tableTargetCount = plan.Targets.Count(target => target.SourceWithinTable);
            AssertTrue(tableTargetCount > 0,
                "Deep rollback fixture does not exercise a formula inside the user table.");
            AssertTrue(tableTargetCount < plan.Targets.Count,
                "Deep rollback fixture does not exercise a formula outside/near the user table.");

            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                var sourceMathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"Native MathType source {target.SourceFormulaId} lost MathML before deep rollback acceptance.");
                prepared[target.Id] = new PreparedWordBulkFormula
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
                        FormulaOleContract.WordOmmlMode,
                        sourceMathMl),
                    MathMl = sourceMathMl,
                };
            }

            var bodyBefore = CaptureDeepRollbackBodySignature(document);
            var textBefore = document.Content.Text ?? string.Empty;
            var bookmarkBefore = CaptureDeepRollbackBookmarks(document);
            var customXmlBefore = CaptureDeepRollbackCustomXml(document);
            var variablesBefore = CaptureDeepRollbackVariables(document);
            var sourceSemanticsBefore = CaptureDeepRollbackMathTypeSemantics(plan);
            var undoBefore = CaptureDeepRollbackUndoHistory(application);
            var paragraphsBefore = document.Paragraphs.Count;
            var tablesBefore = document.Tables.Count;
            var fieldsBefore = document.Fields.Count;
            var firstTableRowsBefore = document.Tables[1].Rows.Count;
            var firstTableColumnsBefore = document.Tables[1].Columns.Count;

            if (File.Exists(tracePath)) File.Delete(tracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE", null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_STAGE",
                "after-final-omml-fingerprint-refresh");

            Exception? expectedFailure = null;
            try
            {
                _ = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            }
            catch (Exception error)
            {
                expectedFailure = error;
            }
            finally
            {
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_STAGE", oldStage);
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE", oldDeleteFailure);
                Environment.SetEnvironmentVariable(
                    "VISUALTEX_WORD_HOOK_TRACE_PATH", oldTrace);
            }
            if (expectedFailure is null)
                throw new InvalidDataException(
                    "Deep finalization rollback fault injection did not fail the conversion.");

            WordFormulaFormatConversionPlan? restoredPlan = null;
            Exception? restoredPlanError = null;
            try
            {
                restoredPlan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.WordOmmlMode);
            }
            catch (Exception error)
            {
                restoredPlanError = error;
            }

            var bodyAfter = CaptureDeepRollbackBodySignature(document);
            var textAfter = document.Content.Text ?? string.Empty;
            var bookmarkAfter = CaptureDeepRollbackBookmarks(document);
            var customXmlAfter = CaptureDeepRollbackCustomXml(document);
            var variablesAfter = CaptureDeepRollbackVariables(document);
            var undoAfter = CaptureDeepRollbackUndoHistory(application);
            var ownedUndo = ReadOwnedUndoDepthFromTrace(tracePath);
            var restoredSemantics = restoredPlan is null
                ? Array.Empty<string>()
                : CaptureDeepRollbackMathTypeSemantics(restoredPlan);
            var failureText = expectedFailure.ToString();

            var checks = new Dictionary<string, bool>(StringComparer.Ordinal)
            {
                ["deepUndo"] = ownedUndo == 1,
                ["primaryFailurePreserved"] =
                    failureText.IndexOf(
                        "Injected format-conversion failure after post-commit OMML validation.",
                        StringComparison.OrdinalIgnoreCase) >= 0
                    && expectedFailure is not AggregateException,
                ["body"] = string.Equals(bodyBefore, bodyAfter, StringComparison.Ordinal),
                ["text"] = string.Equals(textBefore, textAfter, StringComparison.Ordinal),
                ["bookmarks"] = bookmarkBefore.SequenceEqual(bookmarkAfter, StringComparer.Ordinal),
                ["customXml"] = customXmlBefore.SequenceEqual(customXmlAfter, StringComparer.Ordinal),
                ["variables"] = variablesBefore.SequenceEqual(variablesAfter, StringComparer.Ordinal),
                ["sourceSemantics"] = sourceSemanticsBefore.SequenceEqual(restoredSemantics, StringComparer.Ordinal),
                ["restoredPlan"] = restoredPlanError is null && restoredPlan?.Targets.Count == 5,
                ["mathTypeCount"] = CountMathTypeOleShapes(document) == 5,
                ["ommlCount"] = document.OMaths.Count == 0,
                ["paragraphs"] = document.Paragraphs.Count == paragraphsBefore,
                ["tables"] = document.Tables.Count == tablesBefore,
                ["fields"] = document.Fields.Count == fieldsBefore,
                ["userTableShape"] = document.Tables.Count == 1
                    && document.Tables[1].Rows.Count == firstTableRowsBefore
                    && document.Tables[1].Columns.Count == firstTableColumnsBefore,
                ["undoBoundary"] = undoBefore.SequenceEqual(undoAfter, StringComparer.Ordinal),
            };

            var report = new
            {
                passed = checks.Values.All(value => value),
                ownedUndo,
                undoDepthBefore = undoBefore.Length,
                undoDepthAfter = undoAfter.Length,
                failureType = expectedFailure.GetType().FullName,
                failure = failureText,
                restoredPlanError = restoredPlanError?.ToString(),
                checks,
            };
            File.WriteAllText(
                Path.Combine(artifactRoot, "deep-finalize-rollback-result.json"),
                System.Text.Json.JsonSerializer.Serialize(
                    report,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine(
                $"[DEEP FINALIZE ROLLBACK] ownedUndo={ownedUndo}; "
                + string.Join(", ", checks.Select(pair => pair.Key + "=" + pair.Value)));
            if (!report.passed)
                throw new InvalidDataException(
                    "Deep finalization rollback did not restore every verified boundary. "
                    + string.Join(", ", checks.Where(pair => !pair.Value).Select(pair => pair.Key)));

            var verifiedRestoredPlan = restoredPlan
                ?? throw new InvalidDataException(
                    "Deep table success phase has no restored MathType conversion plan.");
            var successPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            foreach (var target in verifiedRestoredPlan.Targets)
            {
                var sourceMathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"Restored MathType source {target.SourceFormulaId} lost MathML before table success conversion.");
                successPrepared[target.Id] = new PreparedWordBulkFormula
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
                        FormulaOleContract.WordOmmlMode,
                        sourceMathMl),
                    MathMl = sourceMathMl,
                };
            }

            var successResult = service.ApplyFormulaFormatConversionPlan(
                verifiedRestoredPlan,
                successPrepared);
            AssertEqual(5, successResult.FormulaCount,
                "User-table MathType→OMML success phase did not convert all five formulas.");
            AssertEqual(0, successResult.FailedFormulaCount,
                "User-table MathType→OMML success phase reported a conversion failure: "
                + string.Join(" | ", successResult.Failures));
            AssertEqual(0, CountMathTypeOleShapes(document),
                "User-table MathType→OMML success phase left MathType sources behind.");
            AssertEqual(5, document.OMaths.Count,
                "User-table MathType→OMML success phase did not leave five OMath targets.");

            int CountOriginalTwoByTwoUserTables()
            {
                var count = 0;
                for (var tableIndex = 1;
                     tableIndex <= document.Tables.Count;
                     tableIndex++)
                {
                    Word.Table? table = null;
                    Word.Rows? rows = null;
                    Word.Columns? columns = null;
                    try
                    {
                        table = document.Tables[tableIndex];
                        rows = table.Rows;
                        columns = table.Columns;
                        if (rows.Count == firstTableRowsBefore
                            && columns.Count == firstTableColumnsBefore)
                            count++;
                    }
                    finally
                    {
                        Release(columns);
                        Release(rows);
                        Release(table);
                    }
                }
                return count;
            }

            void VerifyConvertedTableOwnership(string phase)
            {
                AssertEqual(1, CountOriginalTwoByTwoUserTables(),
                    phase + ": the original 2x2 user table was removed, duplicated or reshaped.");
                foreach (var target in verifiedRestoredPlan.Targets)
                {
                    var convertedFormulaId =
                        successPrepared[target.Id].Session.FormulaId;
                    Word.Bookmark? bookmark = null;
                    Word.Range? range = null;
                    Word.Tables? ownerTables = null;
                    Word.Table? ownerTable = null;
                    Word.Rows? ownerRows = null;
                    Word.Columns? ownerColumns = null;
                    try
                    {
                        bookmark = WordOmmlFormulaStore.FindByFormulaId(
                                document,
                                convertedFormulaId)
                            ?? throw new InvalidDataException(
                                phase + ": converted OMML target lost its VTOMML identity: "
                                + convertedFormulaId);
                        range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                        var insideTable =
                            WordEquationNumbering.RangeIsWhollyWithinTable(range);
                        if (target.SourceWithinTable)
                        {
                            AssertTrue(insideTable,
                                phase + ": a formula originally inside the user table moved into the body.");
                            ownerTables = range.Tables;
                            AssertEqual(1, ownerTables.Count,
                                phase + ": a user-table formula has an ambiguous table owner.");
                            ownerTable = ownerTables[1];
                            ownerRows = ownerTable.Rows;
                            ownerColumns = ownerTable.Columns;
                            AssertEqual(firstTableRowsBefore, ownerRows.Count,
                                phase + ": a user-table formula moved into a managed/non-user row topology.");
                            AssertEqual(firstTableColumnsBefore, ownerColumns.Count,
                                phase + ": a user-table formula moved into a managed/non-user column topology.");
                        }
                        else if (!target.Numbered)
                        {
                            AssertTrue(!insideTable,
                                phase + ": an unnumbered body formula was pulled into a table.");
                        }
                        else if (insideTable)
                        {
                            ownerTables = range.Tables;
                            AssertEqual(1, ownerTables.Count,
                                phase + ": a numbered body formula has an ambiguous managed table owner.");
                            ownerTable = ownerTables[1];
                            ownerRows = ownerTable.Rows;
                            ownerColumns = ownerTable.Columns;
                            AssertEqual(1, ownerRows.Count,
                                phase + ": a numbered body formula did not use a standard managed row.");
                            AssertEqual(3, ownerColumns.Count,
                                phase + ": a numbered body formula did not use a standard 1x3 managed host.");
                        }
                    }
                    finally
                    {
                        Release(ownerColumns);
                        Release(ownerRows);
                        Release(ownerTable);
                        Release(ownerTables);
                        Release(range);
                        Release(bookmark);
                    }
                }
            }

            VerifyConvertedTableOwnership("user-table success");

            var successPath = Path.Combine(
                artifactRoot,
                "deep-user-table-mathtype-to-omml-success.docx");
            document.SaveAs2(
                successPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                successPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertEqual(0, CountMathTypeOleShapes(document),
                "User-table MathType→OMML save/reopen restored MathType sources.");
            AssertEqual(5, document.OMaths.Count,
                "User-table MathType→OMML save/reopen changed the OMath count.");
            VerifyConvertedTableOwnership("user-table save/reopen");
            Console.WriteLine(
                $"[USER TABLE FORMAT PASS] inside={tableTargetCount}; outside={plan.Targets.Count - tableTargetCount}; "
                + "MathType→OMML conversion and save/reopen preserved the original 2x2 user table.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_STAGE", oldStage);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE", oldDeleteFailure);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH", oldTrace);
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

    private static string CaptureDeepRollbackBodySignature(Word.Document document)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var xml = WordDocumentXml.Read(document, content);
            return WordLocalEditSnapshot.Signature(
                xml,
                WordInlineObjectGeometry.Capture(content));
        }
        finally { Release(content); }
    }

    private static string[] CaptureDeepRollbackMathTypeSemantics(
        WordFormulaFormatConversionPlan plan) =>
        plan.Targets
            .OrderBy(target => target.SourceStart)
            .Select(target =>
            {
                var mathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"MathType source {target.SourceFormulaId} has no MathML during rollback verification.");
                return $"{target.DisplayMode}|{target.Numbered}|{MathTypeMtefCodec.SemanticSignature(mathMl)}";
            })
            .ToArray();

    private static string[] CaptureDeepRollbackBookmarks(Word.Document document)
    {
        Word.Bookmarks? bookmarks = null;
        var result = new List<string>();
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Word.Bookmark? bookmark = null;
                Word.Range? range = null;
                try
                {
                    bookmark = bookmarks[index];
                    range = bookmark.Range;
                    result.Add($"{bookmark.Name}|{range.StoryType}|{range.Start}|{range.End}");
                }
                finally { Release(range); Release(bookmark); }
            }
        }
        finally { Release(bookmarks); }
        return result.OrderBy(value => value, StringComparer.Ordinal).ToArray();
    }

    private static string[] CaptureDeepRollbackCustomXml(Word.Document document)
    {
        object? parts = null;
        var values = new List<string>();
        try
        {
            parts = ((dynamic)document).CustomXMLParts;
            for (var index = 1; index <= (int)((dynamic)parts).Count; index++)
            {
                object? part = null;
                try
                {
                    part = ((dynamic)parts)[index];
                    values.Add((string)((dynamic)part).XML);
                }
                finally { Release(part); }
            }
        }
        finally { Release(parts); }
        return values.OrderBy(value => value, StringComparer.Ordinal).ToArray();
    }

    private static string[] CaptureDeepRollbackVariables(Word.Document document)
    {
        Word.Variables? variables = null;
        var values = new List<string>();
        try
        {
            variables = document.Variables;
            for (var index = 1; index <= variables.Count; index++)
            {
                Word.Variable? variable = null;
                try
                {
                    variable = variables[index];
                    values.Add(variable.Name + "=" + variable.Value);
                }
                finally { Release(variable); }
            }
        }
        finally { Release(variables); }
        return values.OrderBy(value => value, StringComparer.Ordinal).ToArray();
    }

    private static string[] CaptureDeepRollbackUndoHistory(Word.Application application)
    {
        CommandBars? bars = null;
        CommandBar? standard = null;
        CommandBarControl? control = null;
        try
        {
            bars = application.CommandBars;
            standard = bars["Standard"];
            control = standard.FindControl(Id: 128)
                ?? throw new InvalidOperationException("Word's native undo-history control is unavailable.");
            var history = control as CommandBarComboBox
                ?? throw new InvalidOperationException("Word's native undo-history control is not a combo box.");
            if (!bars.GetEnabledMso("Undo")) return Array.Empty<string>();
            var result = new string[history.ListCount];
            for (var index = 0; index < result.Length; index++)
                result[index] = history.get_List(index + 1);
            return result;
        }
        finally { Release(control); Release(standard); Release(bars); }
    }

    private static int ReadOwnedUndoDepthFromTrace(string tracePath)
    {
        if (!File.Exists(tracePath))
            throw new InvalidDataException("Deep rollback trace was not created.");
        var matches = System.Text.RegularExpressions.Regex.Matches(
            File.ReadAllText(tracePath),
            @"word-undo-history-recovery original=(\d+) current=(\d+) owned=(\d+)");
        if (matches.Count == 0)
            throw new InvalidDataException("Deep rollback trace contains no undo-depth record.");
        return int.Parse(matches[matches.Count - 1].Groups[3].Value,
            System.Globalization.CultureInfo.InvariantCulture);
    }
}

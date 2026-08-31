using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeToOmmlReeditRegressionAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The MathType→OMML re-edit regression must not attach to the user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-to-omml-reedit.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"420\" height=\"120\" viewBox=\"0 0 420 120\"><text x=\"8\" y=\"82\" font-family=\"Cambria Math\" font-size=\"42\">VisualTeX OMML regression</text></svg>",
            new UTF8Encoding(false));
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 420, 120);
        var previousPreviewDisable = Environment.GetEnvironmentVariable(
            "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW");
        var previousPerfTrace = Environment.GetEnvironmentVariable(
            "VISUALTEX_NUMBERED_PERF_TRACE");
        var previousOmmlFailureStage = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_OMML_FAIL_STAGE");
        Word.Application? application = null;
        Word.Document? document = null;
        bool? originalShowFieldCodes = null;
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");
            Environment.SetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE", "1");
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            originalShowFieldCodes = document.ActiveWindow.View.ShowFieldCodes;
            document.ActiveWindow.View.ShowFieldCodes = true;
            Console.WriteLine("[USER REPORT FIXTURE] Word field-code view forced ON for MathType/OMML regression.");
            ConfigureOmmlTableNumberPage(document);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            Console.WriteLine(
                $"[USER REPORT RUNTIME] serviceAssembly={typeof(WordFormulaService).Assembly.Location}");
            var allFormulas = CreateUserReportedMathTypeFormulaSet();
            (string Latex, string MathMl)[] formulas;
            if (int.TryParse(
                    Environment.GetEnvironmentVariable("VISUALTEX_USER_REPORT_FORMULA_INDEX"),
                    out var selectedFormulaIndex))
            {
                selectedFormulaIndex = Math.Max(
                    1,
                    Math.Min(allFormulas.Count, selectedFormulaIndex));
                var selected = allFormulas[selectedFormulaIndex - 1];
                var repeat = 1;
                if (int.TryParse(
                        Environment.GetEnvironmentVariable("VISUALTEX_USER_REPORT_REPEAT_SELECTED"),
                        out var parsedRepeat))
                    repeat = Math.Max(1, Math.Min(8, parsedRepeat));
                formulas = Enumerable.Repeat(selected, repeat).ToArray();
            }
            else
            {
                var requestedFormulaCount = allFormulas.Count;
                if (int.TryParse(
                        Environment.GetEnvironmentVariable("VISUALTEX_USER_REPORT_FORMULA_COUNT"),
                        out var parsedFormulaCount))
                    requestedFormulaCount = Math.Max(
                        1,
                        Math.Min(allFormulas.Count, parsedFormulaCount));
                formulas = allFormulas.Take(requestedFormulaCount).ToArray();
            }
            var deferLiveGeometryAssertions = string.Equals(
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_USER_REPORT_DEFER_LIVE_GEOMETRY"),
                "1",
                StringComparison.Ordinal);
            var expectInjectedRollback = string.Equals(
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_USER_REPORT_EXPECT_ROLLBACK"),
                "1",
                StringComparison.Ordinal);
            Console.WriteLine(
                $"[USER REPORT FIXTURE] formulaCount={formulas.Length} deferLiveGeometry={deferLiveGeometryAssertions} expectRollback={expectInjectedRollback}");

            for (var index = 0; index < formulas.Length; index++)
            {
                SelectDocumentEnd(document);
                var source = formulas[index];
                service.InsertMathTypeOle(
                    CreateOmmlMathTypeAcceptanceSession(
                        source.MathMl,
                        "block",
                        numbered: true,
                        FormulaOleContract.MathTypeOleMode),
                    source.MathMl,
                    emfPath,
                    updateCreatedMathTypeNumberFields: true);
                if (index < formulas.Length - 1)
                    AppendAcceptanceText(document, "\r");
            }

            if (string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_USER_REPORT_REOPEN_MATHTYPE_SOURCE"),
                    "1",
                    StringComparison.Ordinal))
            {
                var sourcePath = Path.Combine(
                    artifactRoot,
                    "MathType-Source-Before-Conversion.docx");
                document.SaveAs2(
                    sourcePath,
                    Word.WdSaveFormat.wdFormatXMLDocument,
                    AddToRecentFiles: false);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;
                document = application.Documents.Open(
                    sourcePath,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: false);
                document.Activate();
                service = new WordFormulaService(application);
                Console.WriteLine(
                    "[USER REPORT FIXTURE] saved and reopened MathType source before conversion.");
            }

            document.ActiveWindow.View.ShowFieldCodes = true;
            AssertTrue(document.ActiveWindow.View.ShowFieldCodes,
                "The MathType→OMML conversion fixture could not re-enter Word field-code view after source creation.");
            DumpUserReportMathTypeSourceStructure(document, "before-conversion");
            AssertEqual(formulas.Length, CountMathTypeOleShapes(document),
                "The user-report fixture did not create all numbered MathType equations.");
            AssertEqual(formulas.Length, CountMathTypePlaceRefFields(document),
                "The user-report fixture did not create one MTPlaceRef number per MathType equation.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(formulas.Length, plan.Targets.Count,
                "MathType→OMML capture did not find all user-report equations.");
            var orderedTargets = plan.Targets
                .OrderBy(target => target.SourceStart)
                .ToArray();
            var prepared = PrepareOmmlMathTypeTargets(plan, emfPath);
            var conversionWatch = Stopwatch.StartNew();
            var conversion = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            conversionWatch.Stop();
            Console.WriteLine(
                $"[USER REPORT MT→OMML] converted={conversion.FormulaCount} failed={conversion.FailedFormulaCount} elapsedMs={conversionWatch.ElapsedMilliseconds} failures={string.Join(" | ", conversion.Failures)}");
            var diagnosticPartialConversion =
                int.TryParse(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_MAX_FORMAT_ITEMS"),
                    out var maximumAcceptanceItems)
                && maximumAcceptanceItems > 0
                || string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_RETURN_AFTER_SOURCE_DELETE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_STOP_AFTER_MATHTYPE_PLACEREF_DELETE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_STOP_AFTER_MATHTYPE_SHAPE_DELETE"),
                    "1",
                    StringComparison.Ordinal);
            if (diagnosticPartialConversion
                && conversion.FormulaCount < formulas.Length)
            {
                DumpUserReportMathTypeSourceStructure(
                    document,
                    "after-partial-conversion");
                TrySaveUserReportFailureArtifact(
                    document,
                    Path.Combine(
                        artifactRoot,
                        "MathType-To-OMML-Partial.docx"));
                Console.WriteLine(
                    "[diagnostic] partial MathType→OMML conversion remained callable.");
                return;
            }
            AssertEqual(formulas.Length, conversion.FormulaCount,
                "MathType→OMML did not convert all user-report equations.");
            AssertEqual(0, conversion.FailedFormulaCount,
                "MathType→OMML conversion failed before the reported re-edit step: "
                + string.Join(" | ", conversion.Failures));

            var formulaIds = orderedTargets
                .Select(target => prepared[target.Id].Session.FormulaId)
                .ToArray();
            var stopAfterNumberingStage = string.Equals(
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_ACCEPTANCE_RETURN_AFTER_OMML_NUMBERING_BUILD"),
                "1",
                StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_RETURN_AFTER_OMML_SEQUENCE_FINALIZE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_RETURN_AFTER_OMML_FIELD_FINALIZE"),
                    "1",
                    StringComparison.Ordinal)
                || string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_ACCEPTANCE_RETURN_AFTER_OMML_BLANK_RESTORE"),
                    "1",
                    StringComparison.Ordinal);
            if (stopAfterNumberingStage)
            {
                var liveMathCount = document.OMaths.Count;
                var liveTableCount = document.Tables.Count;
                Console.WriteLine(
                    $"[diagnostic] numbered-stage return remained callable: maths={liveMathCount} tables={liveTableCount}.");
                DumpUserReportOmmlStructure(
                    document,
                    formulaIds,
                    "after-numbering-stage-return");
                TrySaveUserReportFailureArtifact(
                    document,
                    Path.Combine(
                        artifactRoot,
                        "MathType-To-OMML-Numbering-Stage.docx"));
                return;
            }
            if (string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_USER_REPORT_FORCE_SCREEN_REFRESH"),
                    "1",
                    StringComparison.Ordinal))
            {
                document.Repaginate();
                application.ScreenRefresh();
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
                Console.WriteLine("[USER REPORT FIXTURE] forced one synchronous layout refresh.");
            }
            if (string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_USER_REPORT_TOGGLE_VIEW"),
                    "1",
                    StringComparison.Ordinal))
            {
                var activeView = document.ActiveWindow.View;
                var originalView = activeView.Type;
                activeView.Type = originalView == Word.WdViewType.wdNormalView
                    ? Word.WdViewType.wdPrintView
                    : Word.WdViewType.wdNormalView;
                System.Windows.Forms.Application.DoEvents();
                activeView.Type = originalView;
                document.Repaginate();
                application.ScreenRefresh();
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
                Release(activeView);
                Console.WriteLine("[USER REPORT FIXTURE] toggled Word view once.");
            }
            var beforeEditPath = Path.Combine(
                artifactRoot,
                "MathType-To-OMML-Before-Reedit.docx");
            TrySaveUserReportFailureArtifact(document, beforeEditPath);
            if (string.Equals(
                    Environment.GetEnvironmentVariable(
                        "VISUALTEX_USER_REPORT_REOPEN_BEFORE_ASSERT"),
                    "1",
                    StringComparison.Ordinal))
            {
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(document);
                document = null;
                document = application.Documents.Open(
                    beforeEditPath,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false,
                    OpenAndRepair: false);
                document.Activate();
                service = new WordFormulaService(application);
                Console.WriteLine("[USER REPORT FIXTURE] reopened before live assertions.");
            }
            AssertEqual(formulas.Length, document.OMaths.Count,
                "MathType→OMML conversion lost an OMath before re-editing.");
            AssertEqual(formulas.Length, document.Tables.Count,
                "MathType→OMML conversion did not create one managed 1x3 host per equation.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "MathType→OMML conversion left a MathType OLE source behind.");
            AssertEqual(0, CountMathTypePlaceRefFields(document),
                "MathType→OMML conversion left an MTPlaceRef field behind.");
            DumpUserReportOmmlStructure(document, formulaIds, "after-conversion");
            AssertManagedNativeOmmlInterTableSeparatorsCompact(
                document,
                formulaIds,
                "MathType→OMML after conversion");
            AssertManagedOmmlMathRunsUseDocumentMathFont(
                document,
                formulaIds,
                "MathType→OMML after conversion");
            if (!deferLiveGeometryAssertions)
            {
                foreach (var formulaId in formulaIds)
                    AssertOmmlTableNumberLifecyclePhase(
                        application,
                        document,
                        formulaId,
                        "MathType→OMML before re-edit");
            }

            // Match the user's screenshot: edit the final integral formula and append
            // Euler's identity while the other converted formulas remain in the same
            // document. A failure must never delete this table or any sibling formula.
            var editedIndex = formulaIds.Length - 1;
            var editedFormulaId = formulaIds[editedIndex];
            var originalMetadata = WordOmmlFormulaStore.TryRead(document, editedFormulaId)
                ?? throw new InvalidDataException(
                    "The converted integral formula lost metadata before re-editing.");
            Word.Range? editedRange = null;
            var injectedFailureObserved = false;
            var originalLiveFingerprint = string.Empty;
            try
            {
                editedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    editedFormulaId,
                    originalMetadata);
                DumpRangeBoundary(
                    document,
                    editedRange,
                    "converted-integral-before-visualtex-reedit");
                originalLiveFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                    WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                        document,
                        editedRange,
                        editedFormulaId));
                var editedMathMl = formulas[editedIndex].MathMl.Replace(
                    "</math>",
                    "<mo>=</mo><msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>");
                var editSession = CreateNumberedOmmlTabSession(
                    editedFormulaId,
                    document.FullName,
                    editedRange.Start,
                    editedRange.End,
                    formulas[editedIndex].Latex + @"=e^{i\pi}+1=0",
                    originalMetadata);
                editSession.FontSizePt = originalMetadata.FontSizePt ?? 12;
                editSession.ExportResult ??= new OfficeExportDocument();
                editSession.ExportResult.MathMl = editedMathMl;

                var editWatch = Stopwatch.StartNew();
                if (expectInjectedRollback)
                    Environment.SetEnvironmentVariable(
                        "VISUALTEX_VSTO_OMML_FAIL_STAGE",
                        "after-direct-table-replacement");
                try
                {
                    service.ReplaceOmml(editSession, editedMathMl);
                }
                catch (Exception error)
                {
                    editWatch.Stop();
                    Console.WriteLine(
                        $"[USER REPORT RE-EDIT FAILURE] ReplaceOmml failed after {editWatch.ElapsedMilliseconds}ms type={error.GetType().FullName} hresult=0x{error.HResult:X8} message={error.Message}");
                    DumpUserReportOmmlStructure(document, formulaIds, "after-reedit-failure");
                    TrySaveUserReportFailureArtifact(
                        document,
                        Path.Combine(artifactRoot, "MathType-To-OMML-After-Reedit-Failure.docx"));
                    if (!expectInjectedRollback)
                        throw;
                    injectedFailureObserved = true;
                }
                finally
                {
                    Environment.SetEnvironmentVariable(
                        "VISUALTEX_VSTO_OMML_FAIL_STAGE",
                        previousOmmlFailureStage);
                }
                editWatch.Stop();
                Console.WriteLine(
                    expectInjectedRollback
                        ? $"[USER REPORT MT→OMML ROLLBACK] elapsedMs={editWatch.ElapsedMilliseconds} failureObserved={injectedFailureObserved}"
                        : $"[USER REPORT MT→OMML RE-EDIT] elapsedMs={editWatch.ElapsedMilliseconds}");
            }
            finally { Release(editedRange); }

            if (expectInjectedRollback)
            {
                AssertTrue(injectedFailureObserved,
                    "The direct-table failure injection did not reach the rollback path.");
                var restoredMetadata = WordOmmlFormulaStore.TryRead(
                        document,
                        editedFormulaId)
                    ?? throw new InvalidDataException(
                        "The direct-table rollback lost the original formula metadata.");
                Word.Range? restoredRange = null;
                try
                {
                    restoredRange = WordOmmlFormulaStore
                        .GetEquationRangeVerifiedForStructuralEdit(
                            document,
                            editedFormulaId,
                            restoredMetadata);
                    var restoredFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(
                        WordOmmlNativeSource.ReadCompleteEquationWordOpenXml(
                            document,
                            restoredRange,
                            editedFormulaId));
                    AssertEqual(originalLiveFingerprint, restoredFingerprint,
                        "The direct-table rollback did not restore the original center OMath.");
                    AssertEqual(originalMetadata.NativeOmmlFingerprint,
                        restoredMetadata.NativeOmmlFingerprint,
                        "The direct-table rollback changed the original OMML metadata fingerprint.");
                }
                finally { Release(restoredRange); }
            }

            AssertEqual(formulas.Length, document.OMaths.Count,
                "Re-editing one converted OMML equation deleted another document equation.");
            AssertEqual(formulas.Length, document.Tables.Count,
                "Re-editing one converted OMML equation deleted or duplicated a 1x3 table.");
            DumpUserReportOmmlStructure(
                document,
                formulaIds,
                expectInjectedRollback
                    ? "after-successful-rollback"
                    : "after-successful-reedit");
            AssertManagedNativeOmmlInterTableSeparatorsCompact(
                document,
                formulaIds,
                expectInjectedRollback
                    ? "MathType→OMML after rollback"
                    : "MathType→OMML after re-edit");
            AssertManagedOmmlMathRunsUseDocumentMathFont(
                document,
                formulaIds,
                expectInjectedRollback
                    ? "MathType→OMML after rollback"
                    : "MathType→OMML after re-edit");
            if (!deferLiveGeometryAssertions)
            {
                foreach (var formulaId in formulaIds)
                    AssertOmmlTableNumberLifecyclePhase(
                        application,
                        document,
                        formulaId,
                        expectInjectedRollback
                            ? "MathType→OMML after rollback"
                            : "MathType→OMML after re-edit");
            }

            var afterEditPath = Path.Combine(
                artifactRoot,
                expectInjectedRollback
                    ? "MathType-To-OMML-After-Rollback.docx"
                    : "MathType-To-OMML-After-Reedit.docx");
            document.SaveAs2(
                afterEditPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                afterEditPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertEqual(formulas.Length, document.OMaths.Count,
                "Save/reopen after converted OMML re-edit lost an equation.");
            AssertEqual(formulas.Length, document.Tables.Count,
                "Save/reopen after converted OMML re-edit changed the 1x3 table count.");
            AssertManagedNativeOmmlInterTableSeparatorsCompact(
                document,
                formulaIds,
                expectInjectedRollback
                    ? "MathType→OMML rollback save/reopen"
                    : "MathType→OMML re-edit save/reopen");
            AssertManagedOmmlMathRunsUseDocumentMathFont(
                document,
                formulaIds,
                expectInjectedRollback
                    ? "MathType→OMML rollback save/reopen"
                    : "MathType→OMML re-edit save/reopen");
            ExpandManagedNativeOmmlInterTableSeparatorsForLegacyFixture(
                document,
                formulaIds);
            var refreshedAfterOpen =
                WordEquationNumbering.RefreshNumberedOmmlTabLayouts(document);
            AssertEqual(formulaIds.Length, refreshedAfterOpen,
                "Document-open OMML layout refresh did not visit every converted 1x3 formula.");
            AssertManagedNativeOmmlInterTableSeparatorsCompact(
                document,
                formulaIds,
                "legacy visible separators after document-open refresh");
            foreach (var formulaId in formulaIds)
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    formulaId,
                    expectInjectedRollback
                        ? "MathType→OMML rollback save/reopen"
                        : "MathType→OMML re-edit save/reopen");
            Console.WriteLine(
                expectInjectedRollback
                    ? "MathType→OMML rollback regression passed: the injected post-replacement failure restored the original center OMath, preserved every sibling formula/table/numbering identity, and survived save/reopen."
                    : "MathType→OMML re-edit regression passed: four numbered Equation.DSMT4 sources converted to four direct-SEQ 1x3 OMML hosts, one converted equation was edited through VisualTeX, and all four survived live state plus save/reopen.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                previousPreviewDisable);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_NUMBERED_PERF_TRACE",
                previousPerfTrace);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_OMML_FAIL_STAGE",
                previousOmmlFailureStage);
            if (document is not null)
            {
                if (originalShowFieldCodes.HasValue)
                {
                    try { document.ActiveWindow.View.ShowFieldCodes = originalShowFieldCodes.Value; } catch { }
                }
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { File.Delete(emfPath); } catch { }
        }
    }

    private static void RunWordUnsavedFirstNumberedOmmlVstoRegressionAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The unsaved-first numbered OMML regression must own its Word process.");
        Directory.CreateDirectory(artifactRoot);
        AssertFailedFreshNumberedOmmlInsertRollsBackAtomically();
        using var host = new WordPerformanceHost(documentPath: null);
        AssertTrue(string.IsNullOrWhiteSpace(host.Document.Path),
            "The unsaved-first regression fixture was unexpectedly saved before insertion.");

        Array startupCustom = Array.Empty<object>();
        host.AddIn.OnStartupComplete(ref startupCustom);
        host.Document.Activate();
        host.Document.ActiveWindow.View.ShowFieldCodes = true;
        AssertTrue(host.Document.ActiveWindow.View.ShowFieldCodes,
            "The fresh first-insert fixture could not enter Word field-code view.");
        WordEquationNumbering.SetEquationNumberFormatPreference(
            host.Document,
            EquationNumberFormat.ContinuousId);
        host.Application.Selection.SetRange(
            host.Document.Content.Start,
            host.Document.Content.Start);

        var existing = SnapshotSessionIds();
        host.AddIn.OnInsertDisplayOmml(new object());
        var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
        var session = client.GetSessionAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        AssertEqual("wordOmml", session.ObjectMode,
            "The first fresh-document OMML command did not create a wordOmml Session.");
        AssertEqual("create", session.Mode,
            "The first fresh-document OMML command did not create a create Session.");

        Commit(
            client,
            session,
            "block",
            "wordOmml",
            @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
            numbered: true,
            mathMl: QuadraticFormulaMathMl());
        var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
        client.CloseEditorAsync(sessionId, CancellationToken.None)
            .GetAwaiter().GetResult();
        WaitForAddInIdle(host.AddIn, TimeSpan.FromSeconds(15));

        Console.WriteLine(
            $"[UNSAVED FIRST OMML VSTO] status={final.Status} error={final.Error ?? string.Empty} "
            + $"tables={host.Document.Tables.Count} maths={host.Document.OMaths.Count} "
            + $"fields={host.Document.Fields.Count} bookmarks={host.Document.Bookmarks.Count}.");
        if (!string.Equals(final.Status, "completed", StringComparison.Ordinal))
        {
            DumpNumberedOmmlBodyParagraphs(host.Document, "unsaved-first-vsto-failure");
            throw new InvalidDataException(
                final.Error ?? "The first numbered OMML insertion in a fresh Word process failed.");
        }

        var formulaId = final.FormulaId
            ?? throw new InvalidDataException(
                "The first numbered OMML insertion completed without a FormulaId.");
        AssertEqual(1, host.Document.Tables.Count,
            "The first numbered OMML insertion did not leave exactly one 1x3 host.");
        AssertEqual(1, host.Document.OMaths.Count,
            "The first numbered OMML insertion lost its center OMath.");
        AssertNumberedOmmlFieldResultsVisible(
            host.Document,
            formulaId,
            "unsaved fresh Word process first numbered OMML insertion");
        AssertOmmlTableNumberLifecyclePhase(
            host.Application,
            host.Document,
            formulaId,
            "unsaved fresh Word process first numbered OMML insertion");

        var path = Path.Combine(
            artifactRoot,
            "unsaved-first-numbered-omml-vsto.docx");
        host.Save(path);
        Console.WriteLine(
            "Unsaved first numbered OMML VSTO regression passed: a freshly started Word instance, "
            + "unsaved Document1, real Ribbon Session/Commit path and first numbered display OMML "
            + "retained one center OMath plus one direct-SEQ 1x3 host.");
    }

    private static void RunWordInstalledUnsavedFirstNumberedOmmlAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previousAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousPerfTrace = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_TRACE_FORMAT_PERF");
        var tracePath = Path.Combine(
            artifactRoot,
            "installed-unsaved-first-omml.trace.log");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            // The installed COM add-in suppresses itself when the acceptance flag is
            // present. Clear it before starting this dedicated Word process so this
            // test exercises the exact Program Files VSTO binary and its real Ribbon
            // callbacks rather than the manually hosted source assembly.
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            try { File.Delete(tracePath); } catch { }
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                tracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_TRACE_FORMAT_PERF",
                "1");
            application = StartInstalledWordNormallyAndWaitForAutomation();
            document = application.Documents.Add();
            document.Activate();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(500);
            callbacksObject = WaitForInstalledWordAddInAutoLoad(
                application,
                out addIns,
                out installedAddIn);
            dynamic callbacks = callbacksObject;
            application.Selection.SetRange(
                document.Content.Start,
                document.Content.Start);
            AssertTrue(string.IsNullOrWhiteSpace(document.Path),
                "The installed first-insert fixture was unexpectedly saved before insertion.");
            document.ActiveWindow.View.ShowFieldCodes = true;
            AssertTrue(document.ActiveWindow.View.ShowFieldCodes,
                "The installed first-insert fixture could not enter Word field-code view.");

            var existing = SnapshotSessionIds();
            callbacks.OnInsertDisplayOmml(null);
            var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
            var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            Commit(
                client,
                session,
                "block",
                "wordOmml",
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                numbered: true,
                mathMl: QuadraticFormulaMathMl());
            var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
            client.CloseEditorAsync(sessionId, CancellationToken.None)
                .GetAwaiter().GetResult();
            WaitForInstalledRibbonSessionRelease(sessionId);

            Console.WriteLine(
                $"[INSTALLED UNSAVED FIRST OMML] status={final.Status} error={final.Error ?? string.Empty} "
                + $"tables={document.Tables.Count} maths={document.OMaths.Count} "
                + $"fields={document.Fields.Count} bookmarks={document.Bookmarks.Count}.");
            AssertEqual("completed", final.Status,
                final.Error ?? "Installed VSTO failed the first numbered OMML insertion in a fresh Word process.");
            AssertEqual(1, document.Tables.Count,
                "Installed VSTO first numbered OMML insertion did not leave one 1x3 host.");
            AssertEqual(1, document.OMaths.Count,
                "Installed VSTO first numbered OMML insertion lost the center OMath.");
            var formulaId = final.FormulaId
                ?? throw new InvalidDataException(
                    "Installed VSTO first numbered OMML insertion completed without FormulaId.");
            AssertNumberedOmmlFieldResultsVisible(
                document,
                formulaId,
                "installed fresh Word process first numbered OMML insertion");
            DumpUserReportOmmlStructure(
                document,
                new[] { formulaId },
                "installed-unsaved-first-numbered-omml");
            Word.View? lifecycleView = null;
            var restoreFieldCodes = false;
            try
            {
                lifecycleView = document.ActiveWindow.View;
                restoreFieldCodes = lifecycleView.ShowFieldCodes;
                if (restoreFieldCodes)
                {
                    // The user's Word may intentionally be in Alt+F9 / field-code
                    // view. Structural validation must preserve that preference, but
                    // the physical right-edge assertion is defined for the rendered
                    // equation-number result, not for the long SEQ instruction text.
                    lifecycleView.ShowFieldCodes = false;
                    document.Repaginate();
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(100);
                }
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    formulaId,
                    "installed fresh Word process first numbered OMML insertion");
            }
            finally
            {
                if (lifecycleView is not null && restoreFieldCodes)
                {
                    try { lifecycleView.ShowFieldCodes = true; } catch { }
                }
                Release(lifecycleView);
            }
            var path = Path.Combine(
                artifactRoot,
                "installed-unsaved-first-numbered-omml.docx");
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            Console.WriteLine(
                "Installed fresh-process OMML acceptance passed: a new WINWORD process loaded the Program Files add-in, "
                + "created unsaved Document1 and committed the first numbered display OMML without losing its center formula.");
        }
        finally
        {
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_ACCEPTANCE",
                previousAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_TRACE_FORMAT_PERF",
                previousPerfTrace);
        }
    }

    private static void AssertNumberedOmmlFieldResultsVisible(
        Word.Document document,
        string formulaId,
        string stage)
    {
        Word.View? view = null;
        Word.Table? table = null;
        Word.Cell? numberCell = null;
        Word.Range? numberRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        try
        {
            view = document.ActiveWindow.View;
            AssertTrue(!view.ShowFieldCodes,
                stage + ": VisualTeX left the Word window in field-code view.");
            table = WordEquationNumbering.FindNumberedEquationTable(document, formulaId)
                ?? throw new InvalidDataException(stage + ": numbered 1x3 host is missing.");
            numberCell = table.Cell(1, 3);
            numberRange = numberCell.Range.Duplicate;
            fields = numberRange.Fields;
            AssertEqual(1, fields.Count,
                stage + ": right number cell does not contain exactly one SEQ field.");
            field = fields[1];
            AssertTrue(!field.ShowCodes,
                stage + ": direct SEQ field is still showing its instruction code.");
            var text = numberRange.Text ?? string.Empty;
            AssertTrue(text.IndexOf("SEQ VisualTeXEquation", StringComparison.OrdinalIgnoreCase) < 0,
                stage + ": right number cell still exposes the SEQ instruction text.");
            AssertTrue(text.IndexOf("(1)", StringComparison.Ordinal) >= 0,
                stage + $": rendered first equation number is not visible as '(1)': '{text}'.");
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(numberRange);
            Release(numberCell);
            Release(table);
            Release(view);
        }
    }

    private static Word.Application StartInstalledWordNormallyAndWaitForAutomation()
    {
        var existingWordProcesses = Process.GetProcessesByName("WINWORD");
        try
        {
            AssertEqual(0, existingWordProcesses.Length,
                "The installed first-insert acceptance requires Word to be completely closed before startup.");
        }
        finally
        {
            foreach (var process in existingWordProcesses) process.Dispose();
        }

        var candidates = new[]
        {
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
        };
        var wordPath = candidates.FirstOrDefault(File.Exists)
            ?? throw new FileNotFoundException(
                "The installed WINWORD.EXE could not be located for fresh-process acceptance.");
        var startInfo = new ProcessStartInfo
        {
            FileName = wordPath,
            UseShellExecute = false,
            WorkingDirectory = Path.GetDirectoryName(wordPath) ?? string.Empty,
        };
        startInfo.EnvironmentVariables.Remove("VISUALTEX_VSTO_ACCEPTANCE");
        startInfo.EnvironmentVariables.Remove("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        using var started = Process.Start(startInfo)
            ?? throw new InvalidOperationException(
                "Windows did not start WINWORD.EXE for installed first-insert acceptance.");

        Word.Application? application = null;
        Exception? lastError = null;
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            try
            {
                application = Marshal.GetActiveObject("Word.Application") as Word.Application;
                if (application is not null) break;
            }
            catch (Exception error)
            {
                lastError = error;
            }
        }
        if (application is null)
            throw new TimeoutException(
                "A normally launched WINWORD.EXE did not publish Word.Application to the ROT within 30 seconds.",
                lastError);
        application.Visible = true;
        application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
        Console.WriteLine(
            $"[INSTALLED WORD START] pid={started.Id} path={wordPath} launchedNormally=true.");
        return application;
    }

    private static object WaitForInstalledWordAddInAutoLoad(
        Word.Application application,
        out Microsoft.Office.Core.COMAddIns addIns,
        out Microsoft.Office.Core.COMAddIn installedAddIn)
    {
        addIns = application.COMAddIns;
        object addInKey = "VisualTeX.WordVsto";
        installedAddIn = addIns.Item(ref addInKey);
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            if (installedAddIn.Connect && installedAddIn.Object is not null)
            {
                Console.WriteLine(
                    "[INSTALLED WORD ADDIN] VisualTeX.WordVsto auto-loaded without modifying COMAddIn.Connect.");
                return installedAddIn.Object;
            }
        }
        throw new InvalidOperationException(
            $"VisualTeX.WordVsto did not auto-load in the normally started Word process. Connect={installedAddIn.Connect}, ObjectAvailable={installedAddIn.Object is not null}.");
    }

    private static void AssertFailedFreshNumberedOmmlInsertRollsBackAtomically()
    {
        var previousFailureStage = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_OMML_FAIL_STAGE");
        using var host = new WordPerformanceHost(documentPath: null);
        try
        {
            Array startupCustom = Array.Empty<object>();
            host.AddIn.OnStartupComplete(ref startupCustom);
            host.Document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                host.Document,
                EquationNumberFormat.ContinuousId);
            host.Application.Selection.SetRange(
                host.Document.Content.Start,
                host.Document.Content.Start);
            var formulaId = Guid.NewGuid().ToString("D");
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                host.Document.FullName,
                host.Document.Content.Start,
                host.Document.Content.Start,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            var service = new WordFormulaService(host.Application);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_OMML_FAIL_STAGE",
                "after-numbered-insert-reconcile");
            var failed = false;
            try
            {
                service.InsertOmml(session, QuadraticFormulaMathMl());
            }
            catch (InvalidOperationException error) when (
                error.Message.IndexOf(
                    "Injected failure after numbered OMML insertion reconciliation",
                    StringComparison.Ordinal) >= 0)
            {
                failed = true;
            }
            AssertTrue(failed,
                "The fresh numbered OMML rollback injection did not reach the intended post-reconcile stage.");
            AssertEqual(0, host.Document.Tables.Count,
                "A failed fresh numbered OMML insertion left a number-only 1x3 table behind.");
            AssertEqual(0, host.Document.OMaths.Count,
                "A failed fresh numbered OMML insertion left an orphan OMath behind.");
            AssertEqual(0, host.Document.Fields.Count,
                "A failed fresh numbered OMML insertion left a SEQ/REF field behind.");
            AssertTrue(
                WordOmmlFormulaStore.FindByFormulaId(host.Document, formulaId) is null,
                "A failed fresh numbered OMML insertion left its VTOMML identity bookmark behind.");
            Console.WriteLine(
                "[UNSAVED FIRST OMML ROLLBACK] injected post-reconcile failure left tables=0 maths=0 fields=0 and no FormulaId bookmark.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_OMML_FAIL_STAGE",
                previousFailureStage);
        }
    }

    private static void RunWordOmmlEmptyLineInsertionRegressionAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The empty-line insertion regression must not attach to the user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            RunEmptyLineInsertionCase(
                application,
                artifactRoot,
                "blank-document",
                initializeDocument: _ => { },
                resolveInsertion: document => document.Content.Start,
                expectedTableStart: _ => 0);
            RunEmptyLineInsertionCase(
                application,
                artifactRoot,
                "between-text-paragraphs",
                initializeDocument: document => document.Content.Text = "VT-BEFORE\r\rVT-AFTER",
                resolveInsertion: document =>
                {
                    Word.Paragraph? paragraph = null;
                    Word.Range? range = null;
                    try
                    {
                        paragraph = document.Paragraphs[2];
                        range = paragraph.Range;
                        AssertTrue(IsOnlyParagraphOrCellMarks(range.Text),
                            "The middle fixture paragraph is not empty before insertion.");
                        return range.Start;
                    }
                    finally
                    {
                        Release(range);
                        Release(paragraph);
                    }
                },
                expectedTableStart: insertionStart => insertionStart);
            Console.WriteLine(
                "Empty-line numbered OMML insertion regression passed: the current empty paragraph was consumed in a blank document and between surrounding text; no undeletable blank paragraph was left before the 1x3 table.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void ExpandManagedNativeOmmlInterTableSeparatorsForLegacyFixture(
        Word.Document document,
        IReadOnlyList<string> formulaIds)
    {
        var tables = new List<(int Start, int End)>();
        foreach (var formulaId in formulaIds)
        {
            Word.Table? table = null;
            Word.Range? tableRange = null;
            try
            {
                table = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        formulaId)
                    ?? throw new InvalidDataException(
                        "Legacy-separator fixture lost a managed 1x3 table.");
                tableRange = table.Range;
                tables.Add((tableRange.Start, tableRange.End));
            }
            finally
            {
                Release(tableRange);
                Release(table);
            }
        }
        tables.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (var index = 1; index < tables.Count; index++)
        {
            Word.Range? separator = null;
            Word.Font? font = null;
            Word.ParagraphFormat? format = null;
            try
            {
                separator = document.Range(tables[index - 1].End, tables[index].Start);
                AssertEqual("\r", separator.Text,
                    "Legacy-separator fixture expected one structural paragraph between tables.");
                font = separator.Font;
                font.Size = 10.5f;
                font.Hidden = 0;
                format = separator.ParagraphFormat;
                format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
                Console.WriteLine(
                    $"[LEGACY VISIBLE OMML TABLE SEPARATOR] index={index} range={separator.Start}:{separator.End} font={font.Size:0.###} lineRule={format.LineSpacingRule}.");
            }
            finally
            {
                Release(format);
                Release(font);
                Release(separator);
            }
        }
    }

    private static void AssertManagedNativeOmmlInterTableSeparatorsCompact(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        string stage)
    {
        var tables = new List<(int Start, int End, string FormulaId)>();
        foreach (var formulaId in formulaIds)
        {
            Word.Table? table = null;
            Word.Range? tableRange = null;
            try
            {
                table = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        formulaId)
                    ?? throw new InvalidDataException(
                        $"{stage}: formula {formulaId} lost its 1x3 table.");
                tableRange = table.Range;
                tables.Add((tableRange.Start, tableRange.End, formulaId));
            }
            finally
            {
                Release(tableRange);
                Release(table);
            }
        }

        tables.Sort((left, right) => left.Start.CompareTo(right.Start));
        for (var index = 1; index < tables.Count; index++)
        {
            var previous = tables[index - 1];
            var current = tables[index];
            Word.Range? separator = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            Word.Font? font = null;
            Word.ParagraphFormat? format = null;
            try
            {
                separator = document.Range(previous.End, current.Start);
                AssertEqual("\r", separator.Text,
                    $"{stage}: formulas {index}/{index + 1} are not separated by exactly one structural paragraph.");
                AssertTrue(!(bool)separator.get_Information(
                               Word.WdInformation.wdWithInTable)
                           && separator.Tables.Count == 0
                           && separator.OMaths.Count == 0
                           && separator.Fields.Count == 0
                           && separator.Bookmarks.Count == 0
                           && separator.InlineShapes.Count == 0,
                    $"{stage}: the table separator owns formula, field, bookmark, OLE, or table content.");
                paragraphs = separator.Paragraphs;
                AssertEqual(1, paragraphs.Count,
                    $"{stage}: the table separator is not one body paragraph.");
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range.Duplicate;
                font = paragraphRange.Font;
                format = paragraphRange.ParagraphFormat;
                AssertTrue(font.Size <= 1.25f,
                    $"{stage}: table separator {index} remains visible at {font.Size:0.###}pt.");
                AssertEqual(Word.WdLineSpacing.wdLineSpaceExactly,
                    format.LineSpacingRule,
                    $"{stage}: table separator {index} is not exact-height.");
                AssertTrue(format.LineSpacing <= 1.25f,
                    $"{stage}: table separator {index} remains {format.LineSpacing:0.###}pt high.");
                Console.WriteLine(
                    $"[COMPACT OMML TABLE SEPARATOR] stage={stage} index={index} range={separator.Start}:{separator.End} font={font.Size:0.###} line={format.LineSpacing:0.###}.");
            }
            finally
            {
                Release(format);
                Release(font);
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(separator);
            }
        }
    }

    private static void RunWordOmmlNumberEndEnterRegressionAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The number-end Enter regression must not attach to the user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.Bookmark? visibleBookmark = null;
        Word.Range? visibleRange = null;
        Word.Selection? selection = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            for (var index = 0; index < 2; index++)
            {
                SelectDocumentEnd(document);
                insertion = application.Selection.Range;
                var formulaId = Guid.NewGuid().ToString("D");
                formulaIds.Add(formulaId);
                var session = CreateNumberedOmmlTabSession(
                    formulaId,
                    document.FullName,
                    insertion.Start,
                    insertion.End,
                    index == 0 ? @"e^{i\pi}+1=0" : @"a^2+b^2=c^2",
                    originalMetadata: null);
                service.InsertOmml(
                    session,
                    index == 0
                        ? CreateUserReportedMathTypeFormulaSet()[1].MathMl
                        : CreateUserReportedMathTypeFormulaSet()[2].MathMl);
                Release(insertion);
                insertion = null;
            }

            DumpNumberedOmmlBodyParagraphs(document, "before-number-end-enter");
            var visibleName = "VTEq_" + Guid.Parse(formulaIds[0]).ToString("N");
            visibleBookmark = document.Bookmarks[visibleName];
            visibleRange = visibleBookmark.Range.Duplicate;
            selection = application.Selection;
            selection.SetRange(visibleRange.End, visibleRange.End);
            Console.WriteLine(
                $"[NUMBER END ENTER] before selection={selection.Start}:{selection.End} visible={visibleRange.Start}:{visibleRange.End} tables={document.Tables.Count} paragraphs={document.Paragraphs.Count}");
            selection.TypeParagraph();
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            Console.WriteLine(
                $"[NUMBER END ENTER] native-after selection={selection.Start}:{selection.End} withinTable={selection.get_Information(Word.WdInformation.wdWithInTable)} tables={document.Tables.Count} paragraphs={document.Paragraphs.Count}");
            AssertTrue((bool)selection.get_Information(
                    Word.WdInformation.wdWithInTable),
                "The regression did not reproduce Word's native extra paragraph inside the number cell.");
            DumpNumberedOmmlBodyParagraphs(document, "after-native-number-end-enter");

            var redirected =
                WordEquationNumbering.TryRedirectManagedNativeOmmlNumberEndEnter(
                    selection);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(100);
            Console.WriteLine(
                $"[NUMBER END ENTER] repaired={redirected} selection={selection.Start}:{selection.End} withinTable={selection.get_Information(Word.WdInformation.wdWithInTable)} tables={document.Tables.Count} paragraphs={document.Paragraphs.Count}");
            DumpNumberedOmmlBodyParagraphs(
                document,
                redirected
                    ? "after-number-end-enter-repair"
                    : "after-number-end-enter-repair-failure");
            AssertTrue(redirected,
                "VisualTeX did not recognize and repair the extra right-cell paragraph created by Enter.");
            AssertTrue(!(bool)selection.get_Information(
                    Word.WdInformation.wdWithInTable),
                "VisualTeX repaired the number cell but left the caret inside the table.");
            AssertEqual(2, document.Tables.Count,
                "Repairing number-end Enter merged or duplicated the two 1x3 tables.");
            AssertEqual(2, document.OMaths.Count,
                "Repairing number-end Enter changed the formula count.");
            foreach (var formulaId in formulaIds)
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    formulaId,
                    "number-end Enter repaired host");
            AssertNumberEndEnterTypingParagraphIsDeletable(
                document,
                selection,
                formulaIds[0]);

            for (var index = 1; index <= document.Tables.Count; index++)
            {
                Word.Table? table = null;
                Word.Range? tableRange = null;
                try
                {
                    table = document.Tables[index];
                    tableRange = table.Range;
                    Console.WriteLine(
                        $"[NUMBER END ENTER TABLE] #{index} range={tableRange.Start}:{tableRange.End} rows={table.Rows.Count} columns={table.Columns.Count} codes={FormatCharacterCodes(tableRange.Text)}");
                }
                finally
                {
                    Release(tableRange);
                    Release(table);
                }
            }

            var path = Path.Combine(artifactRoot, "OMML-Number-End-Enter-Regression.docx");
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
        }
        finally
        {
            Release(selection);
            Release(visibleRange);
            Release(visibleBookmark);
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

    private static void AssertNumberEndEnterTypingParagraphIsDeletable(
        Word.Document document,
        Word.Selection selection,
        string precedingFormulaId)
    {
        Word.Table? table = null;
        Word.Range? tableRange = null;
        Word.Range? selectionRange = null;
        Word.Paragraphs? typingParagraphs = null;
        Word.Paragraph? typingParagraph = null;
        Word.Range? typingRange = null;
        Word.Font? typingFont = null;
        Word.ParagraphFormat? typingFormat = null;
        Word.Range? separatorProbe = null;
        Word.Paragraphs? separatorParagraphs = null;
        Word.Paragraph? separatorParagraph = null;
        Word.Range? separatorRange = null;
        Word.Font? separatorFont = null;
        Word.ParagraphFormat? separatorFormat = null;
        Word.Table? refreshedTable = null;
        Word.Range? refreshedTableRange = null;
        Word.Range? retainedSeparatorProbe = null;
        Word.Paragraphs? retainedSeparatorParagraphs = null;
        Word.Paragraph? retainedSeparatorParagraph = null;
        Word.Range? retainedSeparatorRange = null;
        Word.Font? retainedSeparatorFont = null;
        Word.ParagraphFormat? retainedSeparatorFormat = null;
        try
        {
            table = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    precedingFormulaId)
                ?? throw new InvalidDataException(
                    "The number-end Enter repair lost the preceding 1x3 table.");
            tableRange = table.Range;

            selectionRange = selection.Range.Duplicate;
            typingParagraphs = selectionRange.Paragraphs;
            AssertEqual(1, typingParagraphs.Count,
                "The redirected caret did not resolve to one ordinary paragraph.");
            typingParagraph = typingParagraphs[1];
            typingRange = typingParagraph.Range.Duplicate;
            typingFont = typingRange.Font;
            typingFormat = typingRange.ParagraphFormat;
            AssertEqual(tableRange.End, typingRange.Start,
                "The ordinary typing paragraph is not immediately after the numbered table.");
            AssertEqual("\r", typingRange.Text,
                "The redirected typing paragraph is not truly empty.");
            AssertTrue(!(bool)typingRange.get_Information(
                           Word.WdInformation.wdWithInTable)
                       && typingRange.Tables.Count == 0
                       && typingRange.OMaths.Count == 0
                       && typingRange.Fields.Count == 0
                       && typingRange.Bookmarks.Count == 0
                       && typingRange.InlineShapes.Count == 0,
                "The redirected typing paragraph still owns table, formula, field, bookmark, or OLE state.");
            AssertTrue(typingFont.Size > 1.25f,
                "The redirected typing paragraph inherited the internal 1pt separator format.");
            AssertEqual(Word.WdLineSpacing.wdLineSpaceSingle,
                typingFormat.LineSpacingRule,
                "The redirected typing paragraph is not an ordinary single-line paragraph.");

            separatorProbe = document.Range(
                typingRange.End,
                Math.Min(document.Content.End, typingRange.End + 1));
            AssertTrue(!(bool)separatorProbe.get_Information(
                    Word.WdInformation.wdWithInTable),
                "The mandatory post-typing separator remained inside the next table.");
            separatorParagraphs = separatorProbe.Paragraphs;
            AssertEqual(1, separatorParagraphs.Count,
                "The number-end Enter repair did not retain one structural separator paragraph.");
            separatorParagraph = separatorParagraphs[1];
            separatorRange = separatorParagraph.Range.Duplicate;
            separatorFont = separatorRange.Font;
            separatorFormat = separatorRange.ParagraphFormat;
            AssertEqual("\r", separatorRange.Text,
                "The table separator contains non-structural content.");
            AssertTrue(separatorRange.Tables.Count == 0
                       && separatorRange.OMaths.Count == 0
                       && separatorRange.Fields.Count == 0
                       && separatorRange.Bookmarks.Count == 0,
                "The compact table separator owns formula, field, table, or bookmark state.");
            AssertTrue(separatorFont.Size <= 1.25f,
                $"The mandatory table separator remains visible at {separatorFont.Size:0.###}pt.");
            AssertEqual(Word.WdLineSpacing.wdLineSpaceExactly,
                separatorFormat.LineSpacingRule,
                "The mandatory table separator is not an exact-height line.");
            AssertTrue(separatorFormat.LineSpacing <= 1.25f,
                $"The mandatory table separator remains {separatorFormat.LineSpacing:0.###}pt high.");

            var tableCount = document.Tables.Count;
            var mathCount = document.OMaths.Count;
            var compactFontSize = separatorFont.Size;
            var compactLineSpacing = separatorFormat.LineSpacing;
            typingRange.Delete();
            AssertEqual(tableCount, document.Tables.Count,
                "Deleting the ordinary number-end Enter line merged the two 1x3 tables.");
            AssertEqual(mathCount, document.OMaths.Count,
                "Deleting the ordinary number-end Enter line removed a formula.");

            refreshedTable = WordEquationNumbering.FindNumberedEquationTable(
                    document,
                    precedingFormulaId)
                ?? throw new InvalidDataException(
                    "Deleting the ordinary line lost the preceding 1x3 table.");
            refreshedTableRange = refreshedTable.Range;
            retainedSeparatorProbe = document.Range(
                refreshedTableRange.End,
                Math.Min(document.Content.End, refreshedTableRange.End + 1));
            AssertTrue(!(bool)retainedSeparatorProbe.get_Information(
                    Word.WdInformation.wdWithInTable),
                "Deleting the ordinary line removed the internal table separator.");
            retainedSeparatorParagraphs = retainedSeparatorProbe.Paragraphs;
            AssertEqual(1, retainedSeparatorParagraphs.Count,
                "Deleting the ordinary line did not leave one compact separator.");
            retainedSeparatorParagraph = retainedSeparatorParagraphs[1];
            retainedSeparatorRange = retainedSeparatorParagraph.Range.Duplicate;
            retainedSeparatorFont = retainedSeparatorRange.Font;
            retainedSeparatorFormat = retainedSeparatorRange.ParagraphFormat;
            AssertEqual("\r", retainedSeparatorRange.Text,
                "The retained table separator is no longer structurally empty.");
            AssertTrue(retainedSeparatorFont.Size <= 1.25f
                       && retainedSeparatorFormat.LineSpacingRule
                           == Word.WdLineSpacing.wdLineSpaceExactly
                       && retainedSeparatorFormat.LineSpacing <= 1.25f,
                "Deleting the ordinary line expanded or removed the compact table separator.");
            Console.WriteLine(
                $"[NUMBER END ENTER DELETE] ordinary paragraph deleted; tables={document.Tables.Count} maths={document.OMaths.Count} compactSeparator={compactFontSize:0.###}pt/{compactLineSpacing:0.###}pt retained.");
        }
        finally
        {
            Release(retainedSeparatorFormat);
            Release(retainedSeparatorFont);
            Release(retainedSeparatorRange);
            Release(retainedSeparatorParagraph);
            Release(retainedSeparatorParagraphs);
            Release(retainedSeparatorProbe);
            Release(refreshedTableRange);
            Release(refreshedTable);
            Release(separatorFormat);
            Release(separatorFont);
            Release(separatorRange);
            Release(separatorParagraph);
            Release(separatorParagraphs);
            Release(separatorProbe);
            Release(typingFormat);
            Release(typingFont);
            Release(typingRange);
            Release(typingParagraph);
            Release(typingParagraphs);
            Release(selectionRange);
            Release(tableRange);
            Release(table);
        }
    }

    private static void DumpNumberedOmmlBodyParagraphs(
        Word.Document document,
        string stage)
    {
        Console.WriteLine(
            $"[NUMBERED OMML BODY] stage={stage} content={document.Content.Start}:{document.Content.End} paragraphs={document.Paragraphs.Count} tables={document.Tables.Count}");
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Font? font = null;
            Word.ParagraphFormat? format = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range.Duplicate;
                font = range.Font;
                format = range.ParagraphFormat;
                Console.WriteLine(
                    $"  p#{index}={range.Start}:{range.End} inTable={range.get_Information(Word.WdInformation.wdWithInTable)} tables={range.Tables.Count} maths={range.OMaths.Count} fields={range.Fields.Count} bookmarks={range.Bookmarks.Count} fontSize={font.Size:0.###} hidden={font.Hidden} lineRule={format.LineSpacingRule} line={format.LineSpacing:0.###} before={format.SpaceBefore:0.###} after={format.SpaceAfter:0.###} codes={FormatCharacterCodes(range.Text)}");
            }
            finally
            {
                Release(format);
                Release(font);
                Release(range);
                Release(paragraph);
            }
        }
    }

    private static void RunWordOmmlExistingSeparatorRepairAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Activate();
            var service = new WordFormulaService(application);
            for (var index = 0; index < 3; index++)
            {
                SelectDocumentEnd(document);
                Word.Range? insertion = null;
                try
                {
                    insertion = application.Selection.Range.Duplicate;
                    service.InsertOmml(
                        CreateNumberedOmmlTabSession(
                            Guid.NewGuid().ToString("D"),
                            document.FullName,
                            insertion.Start,
                            insertion.End,
                            @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                            originalMetadata: null),
                        QuadraticFormulaMathMl());
                }
                finally { Release(insertion); }
            }

            var formulaIds = WordOmmlFormulaStore.FormulaIds(document).ToArray();
            AssertEqual(3, formulaIds.Length,
                "Existing-separator repair fixture did not create three numbered OMML formulas.");
            ExpandManagedOmmlTableSeparatorsForFixture(
                document,
                formulaIds,
                "live expanded fixture");

            var beforePath = Path.Combine(
                artifactRoot,
                "OMML-Visible-Table-Separators-Before-Repair.docx");
            document.SaveAs2(
                beforePath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                beforePath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertManagedOmmlTableSeparatorsExpanded(
                document,
                formulaIds,
                "expanded fixture save/reopen");

            var refreshed = WordEquationNumbering.RefreshNumberedOmmlTabLayouts(
                document);
            AssertEqual(formulaIds.Length, refreshed,
                "Document-open OMML layout refresh did not process every managed formula.");
            AssertManagedNativeOmmlInterTableSeparatorsCompact(
                document,
                formulaIds,
                "existing visible separator repair");
            foreach (var formulaId in formulaIds)
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    formulaId,
                    "existing separator repair live");

            var repairedPath = Path.Combine(
                artifactRoot,
                "OMML-Visible-Table-Separators-After-Repair.docx");
            document.SaveAs2(
                repairedPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                repairedPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertManagedNativeOmmlInterTableSeparatorsCompact(
                document,
                formulaIds,
                "existing separator repair save/reopen");
            foreach (var formulaId in formulaIds)
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    formulaId,
                    "existing separator repair save/reopen");
            Console.WriteLine(
                "Existing visible OMML separator repair passed: three 10.5pt/12pt body paragraphs matching the user document were compacted to sterile 1pt fixed-line table separators by the real document-open refresh path and remained compact after save/reopen.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
                catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void ExpandManagedOmmlTableSeparatorsForFixture(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        string stage)
    {
        var tables = new List<(int Start, int End)>();
        foreach (var formulaId in formulaIds)
        {
            Word.Table? table = null;
            Word.Range? range = null;
            try
            {
                table = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        formulaId)
                    ?? throw new InvalidDataException(
                        $"{stage}: formula {formulaId} lost its numbered table.");
                range = table.Range.Duplicate;
                tables.Add((range.Start, range.End));
            }
            finally
            {
                Release(range);
                Release(table);
            }
        }

        var ordered = tables.OrderBy(item => item.Start).ToArray();
        for (var index = 1; index < ordered.Length; index++)
        {
            Word.Range? separator = null;
            Word.Font? font = null;
            Word.ParagraphFormat? format = null;
            try
            {
                separator = document.Range(
                    ordered[index - 1].End,
                    ordered[index].Start);
                AssertEqual("\r", separator.Text,
                    $"{stage}: table pair {index}/{index + 1} has an unexpected separator topology.");
                object normalStyle = Word.WdBuiltinStyle.wdStyleNormal;
                separator.set_Style(ref normalStyle);
                font = separator.Font;
                try { font.Reset(); } catch { }
                font.Size = 10.5f;
                font.Hidden = 0;
                font.Position = 0;
                format = separator.ParagraphFormat;
                try { format.Reset(); } catch { }
                format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                format.LineSpacing = 12f;
                format.SpaceBefore = 0f;
                format.SpaceAfter = 0f;
            }
            finally
            {
                Release(format);
                Release(font);
                Release(separator);
            }
        }
        AssertManagedOmmlTableSeparatorsExpanded(document, formulaIds, stage);
    }

    private static void AssertManagedOmmlTableSeparatorsExpanded(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        string stage)
    {
        var tables = new List<(int Start, int End)>();
        foreach (var formulaId in formulaIds)
        {
            Word.Table? table = null;
            Word.Range? range = null;
            try
            {
                table = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        formulaId)
                    ?? throw new InvalidDataException(
                        $"{stage}: formula {formulaId} lost its numbered table.");
                range = table.Range.Duplicate;
                tables.Add((range.Start, range.End));
            }
            finally
            {
                Release(range);
                Release(table);
            }
        }
        var ordered = tables.OrderBy(item => item.Start).ToArray();
        for (var index = 1; index < ordered.Length; index++)
        {
            Word.Range? separator = null;
            Word.Font? font = null;
            Word.ParagraphFormat? format = null;
            Word.Tables? nestedTables = null;
            Word.OMaths? maths = null;
            Word.Fields? fields = null;
            Word.Bookmarks? bookmarks = null;
            Word.InlineShapes? shapes = null;
            Word.Frames? frames = null;
            try
            {
                separator = document.Range(
                    ordered[index - 1].End,
                    ordered[index].Start);
                AssertEqual("\r", separator.Text,
                    $"{stage}: table pair {index}/{index + 1} does not have one body separator.");
                nestedTables = separator.Tables;
                maths = separator.OMaths;
                fields = separator.Fields;
                bookmarks = separator.Bookmarks;
                shapes = separator.InlineShapes;
                frames = separator.Frames;
                AssertTrue(nestedTables.Count == 0
                           && maths.Count == 0
                           && fields.Count == 0
                           && bookmarks.Count == 0
                           && shapes.Count == 0
                           && frames.Count == 0,
                    $"{stage}: expanded separator owns unexpected managed content.");
                font = separator.Font;
                format = separator.ParagraphFormat;
                AssertTrue(font.Size >= 9f,
                    $"{stage}: separator was not expanded to the user-reported visible size; font={font.Size:0.###}pt.");
                AssertTrue(format.LineSpacing >= 10f,
                    $"{stage}: separator was not expanded to the user-reported visible height; line={format.LineSpacing:0.###}pt.");
                Console.WriteLine(
                    $"[VISIBLE 1X3 SEPARATOR FIXTURE] stage={stage} pair={index}/{index + 1} range={separator.Start}:{separator.End} font={font.Size:0.###}pt line={format.LineSpacing:0.###}pt.");
            }
            finally
            {
                Release(frames);
                Release(shapes);
                Release(bookmarks);
                Release(fields);
                Release(maths);
                Release(nestedTables);
                Release(format);
                Release(font);
                Release(separator);
            }
        }
    }

    private static void RunWordOmmlDeferredFinalizationCostAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The deferred-finalization cost probe must not attach to the user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            var service = new WordFormulaService(application);
            var formulaIds = new List<string>();
            for (var index = 0; index < 8; index++)
            {
                SelectDocumentEnd(document);
                var formulaId = Guid.NewGuid().ToString("D");
                formulaIds.Add(formulaId);
                var insertion = application.Selection.Range;
                try
                {
                    var session = CreateNumberedOmmlTabSession(
                        formulaId,
                        document.FullName,
                        insertion.Start,
                        insertion.End,
                        @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                        originalMetadata: null);
                    service.InsertOmml(session, QuadraticFormulaMathMl());
                }
                finally { Release(insertion); }
            }

            var retiredScheduler = typeof(ThisAddIn).GetMethod(
                "ScheduleNumberedOmmlDisplayShapeFinalization",
                System.Reflection.BindingFlags.Instance
                | System.Reflection.BindingFlags.NonPublic);
            AssertTrue(retiredScheduler is null,
                "The retired five-turn numbered-OMML Shape finalizer was reintroduced into ThisAddIn.");

            var targetId = formulaIds[formulaIds.Count / 2];
            var metadata = WordOmmlFormulaStore.TryRead(document, targetId)
                ?? throw new InvalidDataException("Cost-probe target metadata is missing.");
            Word.Range? targetRange = null;
            var editWatch = Stopwatch.StartNew();
            try
            {
                targetRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    targetId,
                    metadata);
                var session = CreateNumberedOmmlTabSession(
                    targetId,
                    document.FullName,
                    targetRange.Start,
                    targetRange.End,
                    @"y=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                    metadata);
                session.FontSizePt = metadata.FontSizePt ?? 12;
                service.ReplaceOmml(
                    session,
                    QuadraticFormulaMathMl().Replace("<mi>x</mi>", "<mi>y</mi>"));
            }
            finally
            {
                editWatch.Stop();
                Release(targetRange);
            }

            AssertTrue(editWatch.ElapsedMilliseconds < 2_000,
                $"A direct 1x3 OMML edit took {editWatch.ElapsedMilliseconds}ms; the synchronous path should complete without a post-edit busy period.");
            AssertEqual(formulaIds.Count, document.Tables.Count,
                "The direct performance edit changed the managed 1x3 table count.");
            AssertEqual(formulaIds.Count, document.OMaths.Count,
                "The direct performance edit changed the document OMath count.");
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                targetId,
                "post-edit no-deferred-finalization");
            Console.WriteLine(
                $"[NO DEFERRED FINALIZATION] formulas={formulaIds.Count} editMs={editWatch.ElapsedMilliseconds} retiredSchedulerPresent={retiredScheduler is not null}");

            var path = Path.Combine(artifactRoot, "OMML-Deferred-Finalization-Cost.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
        }
        finally
        {
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

    private static void RunEmptyLineInsertionCase(
        Word.Application application,
        string artifactRoot,
        string caseName,
        Action<Word.Document> initializeDocument,
        Func<Word.Document, int> resolveInsertion,
        Func<int, int> expectedTableStart)
    {
        Word.Document? document = null;
        Word.Range? insertionRange = null;
        Word.Table? table = null;
        Word.Range? tableRange = null;
        Word.Range? beforeTable = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            initializeDocument(document);
            var insertionStart = resolveInsertion(document);
            insertionRange = document.Range(insertionStart, insertionStart);
            insertionRange.Select();
            var formulaId = Guid.NewGuid().ToString("D");
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertionStart,
                insertionStart,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            var service = new WordFormulaService(application);
            service.InsertOmml(session, QuadraticFormulaMathMl());

            table = WordEquationNumbering.FindNumberedEquationTable(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Empty-line case '{caseName}' did not create a managed 1x3 table.");
            tableRange = table.Range;
            var expectedStart = expectedTableStart(insertionStart);
            beforeTable = document.Range(document.Content.Start, tableRange.Start);
            Console.WriteLine(
                $"[EMPTY LINE INSERTION] case={caseName} insertionStart={insertionStart} tableStart={tableRange.Start} expected={expectedStart} beforeCodes={FormatCharacterCodes(beforeTable.Text)} contentCodes={FormatCharacterCodes(document.Content.Text)}");
            AssertEqual(expectedStart, tableRange.Start,
                $"Empty-line case '{caseName}' placed the formula after the current empty paragraph instead of consuming that paragraph.");
            DumpUserReportOmmlStructure(
                document,
                new[] { formulaId },
                "empty-line-" + caseName);
            AssertOmmlTableNumberLifecyclePhase(
                application,
                document,
                formulaId,
                "empty-line insertion " + caseName);

            var path = Path.Combine(artifactRoot, $"OMML-Empty-Line-{caseName}.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
        }
        finally
        {
            Release(beforeTable);
            Release(tableRange);
            Release(table);
            Release(insertionRange);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static IReadOnlyList<(string Latex, string MathMl)> CreateUserReportedMathTypeFormulaSet() =>
        new[]
        {
            (
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                QuadraticFormulaMathMl()),
            (
                @"e^{i\pi}+1=0",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>e</mi><mrow><mi>i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>"),
            (
                @"a^2+b^2=c^2",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>"),
            (
                @"\int_{-\infty}^{\infty}e^{-x^2}\,dx=\sqrt{\pi}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msubsup><mo>∫</mo><mrow><mo>−</mo><mi>∞</mi></mrow><mi>∞</mi></msubsup><msup><mi>e</mi><mrow><mo>−</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></msup><mi>d</mi><mi>x</mi><mo>=</mo><msqrt><mi>π</mi></msqrt></math>"),
        };

    private static void AssertManagedOmmlMathRunsUseDocumentMathFont(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        string stage)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var expected = (document.OMathFontName ?? string.Empty).Trim();
        AssertTrue(!string.IsNullOrWhiteSpace(expected),
            stage + ": document OMathFontName is empty.");
        var explicitMathRunFonts = 0;
        var explicitControlFonts = 0;
        foreach (var formulaId in formulaIds)
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    $"{stage}: formula {formulaId} lost metadata before font verification.");
            Word.Range? range = null;
            try
            {
                range = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                var xml = range.WordOpenXML ?? string.Empty;
                var package = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
                XNamespace math = MathNamespace;
                XNamespace word = WordNamespace;

                foreach (var run in package.Descendants(math + "r"))
                {
                    if (run.Element(math + "rPr")?.Element(math + "nor") is not null)
                        continue;
                    var fonts = run.Element(word + "rPr")?.Element(word + "rFonts");
                    if (fonts is null) continue;
                    explicitMathRunFonts++;
                    AssertExplicitMathFont(fonts, word, expected, stage, formulaId);
                }
                foreach (var control in package.Descendants(math + "ctrlPr"))
                {
                    var fonts = control.Element(word + "rPr")?.Element(word + "rFonts");
                    if (fonts is null) continue;
                    explicitControlFonts++;
                    AssertExplicitMathFont(fonts, word, expected, stage, formulaId);
                }
            }
            finally { Release(range); }
        }
        Console.WriteLine(
            $"[OMML MATH FONT STABLE] stage={stage} documentMathFont='{expected}' explicitMathRuns={explicitMathRunFonts} explicitControls={explicitControlFonts} formulas={formulaIds.Count}.");
    }

    private static void AssertExplicitMathFont(
        XElement fonts,
        XNamespace word,
        string expected,
        string stage,
        string formulaId)
    {
        foreach (var attributeName in new[] { "ascii", "hAnsi" })
        {
            var value = ((string?)fonts.Attribute(word + attributeName) ?? string.Empty).Trim();
            if (value.Length == 0) continue;
            AssertTrue(string.Equals(value, expected, StringComparison.OrdinalIgnoreCase),
                $"{stage}: formula {formulaId} carries explicit {attributeName} math font '{value}', expected document math font '{expected}'.");
        }
    }

    private static void DumpUserReportMathTypeSourceStructure(
        Word.Document document,
        string stage)
    {
        Console.WriteLine(
            $"[USER REPORT MATHTYPE STRUCTURE] stage={stage} paragraphs={document.Paragraphs.Count} inlineShapes={document.InlineShapes.Count} tables={document.Tables.Count} fields={document.Fields.Count} contentCodes={FormatCharacterCodes(document.Content.Text)}");
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Style? style = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range.Duplicate;
                object styleObject = range.get_Style();
                style = styleObject as Word.Style;
                var styleName = style?.NameLocal
                    ?? Convert.ToString(styleObject)
                    ?? string.Empty;
                Console.WriteLine(
                    $"  p#{index} range={range.Start}:{range.End} style={styleName} shapes={range.InlineShapes.Count} fields={range.Fields.Count} maths={range.OMaths.Count} tables={range.Tables.Count} codes={FormatCharacterCodes(range.Text)}");
            }
            finally
            {
                Release(style);
                Release(range);
                Release(paragraph);
            }
        }
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shape = document.InlineShapes[index];
                range = shape.Range.Duplicate;
                Console.WriteLine(
                    $"  shape#{index} range={range.Start}:{range.End} progId={shape.OLEFormat.ProgID}");
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }
    }

    private static void DumpUserReportOmmlStructure(
        Word.Document document,
        IReadOnlyList<string> formulaIds,
        string stage)
    {
        Console.WriteLine(
            $"[USER REPORT STRUCTURE] stage={stage} maths={document.OMaths.Count} tables={document.Tables.Count} inlineShapes={document.InlineShapes.Count} shapes={document.Shapes.Count} formulaIds={WordOmmlFormulaStore.FormulaIds(document).Count}");
        for (var index = 0; index < formulaIds.Count; index++)
        {
            var formulaId = formulaIds[index];
            Word.Bookmark? bookmark = null;
            Word.Range? equationRange = null;
            Word.Table? table = null;
            Word.Range? tableRange = null;
            Word.Cell? centerCell = null;
            Word.Range? centerRange = null;
            Word.Cell? numberCell = null;
            Word.Range? numberRange = null;
            Word.Range? tableStart = null;
            Word.Range? numberEnd = null;
            Word.ParagraphFormat? tableParagraphFormat = null;
            Word.Style? tableParagraphStyle = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is not null)
                    equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                table = WordEquationNumbering.FindNumberedEquationTable(document, formulaId);
                if (table is not null)
                {
                    tableRange = table.Range;
                    centerCell = table.Cell(1, 2);
                    centerRange = centerCell.Range;
                    numberCell = table.Cell(1, 3);
                    numberRange = numberCell.Range;
                    tableStart = tableRange.Duplicate;
                    tableStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    var visibleName = "VTEq_" + Guid.Parse(formulaId).ToString("N");
                    if (document.Bookmarks.Exists(visibleName))
                    {
                        var visible = document.Bookmarks[visibleName];
                        try
                        {
                            numberEnd = visible.Range.Duplicate;
                            numberEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        finally { Release(visible); }
                    }
                    tableParagraphFormat = tableRange.ParagraphFormat;
                    object styleObject = tableRange.get_Style();
                    tableParagraphStyle = styleObject as Word.Style;
                }
                var tablePageX = tableStart is null
                    ? float.NaN
                    : Convert.ToSingle(tableStart.get_Information(
                        Word.WdInformation.wdHorizontalPositionRelativeToPage));
                var tableTextX = tableStart is null
                    ? float.NaN
                    : Convert.ToSingle(tableStart.get_Information(
                        Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary));
                var numberPageX = numberEnd is null
                    ? float.NaN
                    : Convert.ToSingle(numberEnd.get_Information(
                        Word.WdInformation.wdHorizontalPositionRelativeToPage));
                var paragraphStyleName = tableParagraphStyle?.NameLocal ?? string.Empty;
                var tableLeftIndent = tableParagraphFormat?.LeftIndent ?? float.NaN;
                var tableRows = table?.Rows;
                Console.WriteLine(
                    $"  #{index + 1} id={formulaId} bookmark={(bookmark is null ? "missing" : bookmark.Range.Start + ":" + bookmark.Range.End)} equation={(equationRange is null ? "missing" : equationRange.Start + ":" + equationRange.End)} table={(tableRange is null ? "missing" : tableRange.Start + ":" + tableRange.End)} view={document.ActiveWindow.View.Type} tablePageX={tablePageX:0.###} tableTextX={tableTextX:0.###} numberPageX={numberPageX:0.###} rowAlignment={tableRows?.Alignment} rowLeftIndent={tableRows?.LeftIndent:0.###} wrap={tableRows?.WrapAroundText} horizontal={tableRows?.HorizontalPosition:0.###} relativeHorizontal={tableRows?.RelativeHorizontalPosition} vertical={tableRows?.VerticalPosition:0.###} relativeVertical={tableRows?.RelativeVerticalPosition} tableFrames={tableRange?.Frames.Count} equationFrames={equationRange?.Frames.Count} tableParagraphStyle={paragraphStyleName} tableParagraphLeftIndent={tableLeftIndent:0.###} centerCodes={FormatCharacterCodes(centerRange?.Text)} numberCodes={FormatCharacterCodes(numberRange?.Text)} metadata={(WordOmmlFormulaStore.TryRead(document, formulaId) is null ? "missing" : "present")}");
                Release(tableRows);
            }
            catch (Exception error)
            {
                Console.WriteLine(
                    $"  #{index + 1} id={formulaId} inspection-failed type={error.GetType().Name} hresult=0x{error.HResult:X8} message={error.Message}");
            }
            finally
            {
                Release(tableParagraphStyle);
                Release(tableParagraphFormat);
                Release(numberEnd);
                Release(tableStart);
                Release(numberRange);
                Release(numberCell);
                Release(centerRange);
                Release(centerCell);
                Release(tableRange);
                Release(table);
                Release(equationRange);
                Release(bookmark);
            }
        }
    }

    private static void DumpRangeBoundary(
        Word.Document document,
        Word.Range equationRange,
        string stage)
    {
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Range? context = null;
        try
        {
            paragraphs = equationRange.Paragraphs;
            if (paragraphs.Count > 0)
            {
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
            }
            var contextStart = Math.Max(document.Content.Start, equationRange.Start - 3);
            var contextEnd = Math.Min(document.Content.End, equationRange.End + 3);
            context = document.Range(contextStart, contextEnd);
            Console.WriteLine(
                $"[OMML RANGE BOUNDARY] stage={stage} equation={equationRange.Start}:{equationRange.End} equationCodes={FormatCharacterCodes(equationRange.Text)} paragraph={(paragraphRange is null ? "missing" : paragraphRange.Start + ":" + paragraphRange.End)} paragraphCodes={FormatCharacterCodes(paragraphRange?.Text)} context={context.Start}:{context.End} contextCodes={FormatCharacterCodes(context.Text)} withinTable={equationRange.get_Information(Word.WdInformation.wdWithInTable)}");
        }
        finally
        {
            Release(context);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
        }
    }

    private static void TrySaveUserReportFailureArtifact(
        Word.Document document,
        string path)
    {
        try
        {
            document.SaveAs2(
                path,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            Console.WriteLine($"[USER REPORT FAILURE ARTIFACT] saved={path}");
        }
        catch (Exception saveError)
        {
            Console.WriteLine(
                $"[USER REPORT FAILURE ARTIFACT] save-failed type={saveError.GetType().Name} hresult=0x{saveError.HResult:X8} message={saveError.Message}");
        }
    }

    private static bool IsOnlyParagraphOrCellMarks(string? text)
    {
        if (string.IsNullOrEmpty(text)) return true;
        return text.All(character => character is '\r' or '\a' or '\n');
    }

    private static string FormatCharacterCodes(string? text)
    {
        if (text is null) return "<null>";
        if (text.Length == 0) return "<empty>";
        return string.Join(",", text.Select(character => $"U+{(int)character:X4}"));
    }
}

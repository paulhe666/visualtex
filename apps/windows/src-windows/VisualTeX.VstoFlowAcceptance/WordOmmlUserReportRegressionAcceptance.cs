using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
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
        try
        {
            Environment.SetEnvironmentVariable(
                "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW",
                "1");
            Environment.SetEnvironmentVariable("VISUALTEX_NUMBERED_PERF_TRACE", "1");
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
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
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { File.Delete(emfPath); } catch { }
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

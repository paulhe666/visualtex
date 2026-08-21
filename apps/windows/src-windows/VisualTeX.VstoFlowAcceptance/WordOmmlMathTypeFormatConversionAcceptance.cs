using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordInstalledOmmlMathTypeFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var fixturePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_OMML_MATHTYPE_INSTALLED_FIXTURE");
        if (string.IsNullOrWhiteSpace(fixturePath) || !File.Exists(fixturePath))
            throw new InvalidOperationException(
                "Installed OMML↔MathType acceptance requires VISUALTEX_OMML_MATHTYPE_INSTALLED_FIXTURE pointing to the saved OMML core fixture.");

        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var previousFormatAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var tracePath = Path.Combine(
            artifactRoot,
            "installed-omml-mathtype-format-conversion.trace.log");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                "1");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                tracePath);

            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            if (mathTypeBaseline.Count != 0)
                throw new InvalidOperationException(
                    "Installed OMML↔MathType acceptance requires MathType.exe process count to be zero before Word starts.");

            var writablePath = Path.Combine(
                artifactRoot,
                "installed-omml-mathtype-input.docx");
            File.Copy(Path.GetFullPath(fixturePath), writablePath, overwrite: true);

            application = CreateWordApplication(visible: true);
            document = application.Documents.Open(
                writablePath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(7, document.OMaths.Count,
                "Installed OMML↔MathType fixture must start with the seven saved OMML equations from core acceptance.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed OMML↔MathType fixture unexpectedly contains a MathType object before the installed add-in test.");

            // Add one equation through Word itself after the fixture opens. This
            // deliberately has no VisualTeX custom XML/bookmark metadata and proves
            // the installed Ribbon path also handles arbitrary native OMath content.
            const string nativeToken = "VT_INSTALLED_PURE_NATIVE_OMML";
            AppendAcceptanceText(
                document,
                $" installed-before-native {nativeToken} installed-after-native\r");
            Word.Range? nativeRange = null;
            try
            {
                nativeRange = InsertPureNativeOmml(document, nativeToken, "r+2");
            }
            finally { Release(nativeRange); }
            AssertEqual(8, document.OMaths.Count,
                "Installed acceptance did not create the additional pure Word-native OMath source.");

            addIns = application.COMAddIns;
            object addInKey = "VisualTeX.WordVsto";
            installedAddIn = addIns.Item(ref addInKey);
            if (!installedAddIn.Connect)
                installedAddIn.Connect = true;
            for (var index = 0; index < 80 && installedAddIn.Object is null; index++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
            callbacksObject = installedAddIn.Object
                ?? throw new InvalidOperationException(
                    "Installed VisualTeX.WordVsto automation object was unavailable. The acceptance refuses to fall back to a locally constructed ThisAddIn.");
            dynamic callbacks = callbacksObject;

            SelectFirstOmmlEquations(document, 2);
            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertOmmlToMathTypeSelection(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline);
            AssertEqual(6, document.OMaths.Count,
                "Installed OMML→MathType selection callback changed the wrong OMML count.");
            AssertEqual(2, CountMathTypeOleShapes(document),
                "Installed OMML→MathType selection callback did not create two MathType objects.");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertOmmlToMathTypeDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline);
            AssertEqual(0, document.OMaths.Count,
                "Installed OMML→MathType document callback left OMML sources behind in the mixed document.");
            AssertEqual(8, CountMathTypeOleShapes(document),
                "Installed OMML→MathType document callback did not preserve the two already-converted MathType objects while converting the remaining six OMML equations.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Installed OMML→MathType conversion did not preserve the single numbered equation as MTPlaceRef.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after installed OMML→MathType callbacks");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after installed OMML→MathType callbacks");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "installed OMML→MathType callbacks");

            var mathTypePath = Path.Combine(
                artifactRoot,
                "Installed-OMML-To-MathType-Acceptance.docx");
            document.SaveAs2(mathTypePath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                mathTypePath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(8, CountMathTypeOleShapes(document),
                "Installed OMML→MathType save/reopen lost a MathType object.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Installed OMML→MathType save/reopen changed the numbered equation.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after installed OMML→MathType save/reopen");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after installed OMML→MathType save/reopen");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "installed OMML→MathType save/reopen");

            SelectFirstMathTypeEquations(document, 2);
            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertMathTypeToOmmlSelection(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=MathType target=OMML",
                mathTypeBaseline);
            AssertEqual(2, document.OMaths.Count,
                "Installed MathType→OMML selection callback did not create two OMath equations.");
            AssertEqual(6, CountMathTypeOleShapes(document),
                "Installed MathType→OMML selection callback changed the wrong MathType count.");

            ResetInstalledFormatConversionTrace(tracePath);
            callbacks.OnConvertMathTypeToOmmlDocument(null);
            WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=MathType target=OMML",
                mathTypeBaseline);
            AssertEqual(8, document.OMaths.Count,
                "Installed MathType→OMML document callback did not restore all eight equations in the mixed document.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed MathType→OMML document callback left MathType sources behind.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "Installed MathType→OMML callbacks did not restore the single numbered OMML equation.");
            AssertOmmlSemanticCoverage(document);
            AssertInstalledOmmlMathTypeProseSurvived(document);
            AssertNoNewMathTypeProcess(mathTypeBaseline, "installed MathType→OMML callbacks");

            var finalPath = Path.Combine(
                artifactRoot,
                "Installed-MathType-To-OMML-Acceptance.docx");
            document.SaveAs2(finalPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(8, document.OMaths.Count,
                "Installed MathType→OMML save/reopen lost a native equation.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Installed MathType→OMML save/reopen restored a MathType object.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "Installed MathType→OMML save/reopen changed numbering state.");
            AssertOmmlSemanticCoverage(document);
            AssertInstalledOmmlMathTypeProseSurvived(document);
            AssertNoNewMathTypeProcess(mathTypeBaseline, "full installed OMML↔MathType acceptance");

            Console.WriteLine(
                "[OMML↔MATHTYPE INSTALLED] Installed VisualTeX.WordVsto COM automation object executed all four new Ribbon callbacks. Selection/document conversion in both directions passed with a mixed document, one runtime-created pure Word OMath, numbering, complex semantics, Equation.DSMT4/Equation Native integrity, live metafile ink, no OlePres, save/reopen persistence, and MathTypeProcessCount=0.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void ResetInstalledFormatConversionTrace(string tracePath)
    {
        try { File.Delete(tracePath); } catch { }
    }

    private static int WaitForInstalledOmmlMathTypeConversion(
        string tracePath,
        string directionMarker,
        IReadOnlyCollection<int> mathTypeBaseline,
        bool allowTransientMathTypeProcess = false)
    {
        var deadline = DateTime.UtcNow.AddSeconds(120);
        var completionMarker = "format-conversion-complete " + directionMarker;
        var stoppedMarker = "format-conversion-stopped " + directionMarker;
        var peakAdditionalMathTypeProcesses = 0;
        while (DateTime.UtcNow < deadline)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(200);
            var startedMathType = SnapshotMathTypeProcessIds()
                .Except(mathTypeBaseline)
                .ToArray();
            peakAdditionalMathTypeProcesses = Math.Max(
                peakAdditionalMathTypeProcesses,
                startedMathType.Length);
            if (!allowTransientMathTypeProcess && startedMathType.Length > 0)
                throw new InvalidOperationException(
                    "Installed OMML↔MathType conversion started MathType.exe: "
                    + string.Join(", ", startedMathType));
            if (!File.Exists(tracePath)) continue;
            string trace;
            try { trace = File.ReadAllText(tracePath); }
            catch { continue; }
            var stoppedIndex = trace.IndexOf(stoppedMarker, StringComparison.Ordinal);
            if (stoppedIndex >= 0)
                throw new InvalidOperationException(
                    "Installed Ribbon conversion reported a stopped transaction: "
                    + trace.Substring(stoppedIndex).Trim());
            if (trace.IndexOf(completionMarker, StringComparison.Ordinal) < 0)
                continue;
            for (var settle = 0; settle < 50; settle++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
                var remainingMathType = SnapshotMathTypeProcessIds()
                    .Except(mathTypeBaseline)
                    .ToArray();
                peakAdditionalMathTypeProcesses = Math.Max(
                    peakAdditionalMathTypeProcesses,
                    remainingMathType.Length);
                if (remainingMathType.Length == 0)
                    return peakAdditionalMathTypeProcesses;
                if (!allowTransientMathTypeProcess)
                    throw new InvalidOperationException(
                        "Installed OMML↔MathType conversion left MathType.exe running: "
                        + string.Join(", ", remainingMathType));
            }
            throw new InvalidOperationException(
                "Installed Ribbon conversion completed but its transient MathType native-preview helper did not exit.");
        }
        throw new TimeoutException(
            $"Installed Ribbon conversion did not report '{completionMarker}' within the acceptance deadline.");
    }

    private static void AssertInstalledOmmlMathTypeProseSurvived(Word.Document document)
    {
        var requiredInOrder = new[]
        {
            "before-inline-hbar-greek",
            "after-inline-hbar-greek",
            "before-display-fraction-integral",
            "after-display-fraction-integral",
            "before-display-matrix",
            "after-display-matrix",
            "before-inline-accents-vector",
            "after-inline-accents-vector",
            "before-pure-native-omml",
            "after-pure-native-omml",
            "installed-before-native",
            "installed-after-native",
        };
        var text = document.Content.Text ?? string.Empty;
        var previousIndex = -1;
        foreach (var marker in requiredInOrder)
        {
            var index = text.IndexOf(marker, StringComparison.Ordinal);
            AssertTrue(index > previousIndex,
                $"Installed OMML↔MathType conversion removed or reordered adjacent prose marker '{marker}'.");
            previousIndex = index;
        }
    }

    private static void RunWordOmmlMathTypeFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var pngPath = Path.Combine(artifactRoot, "omml-mathtype-format-conversion.png");
        var svgPath = Path.Combine(artifactRoot, "omml-mathtype-format-conversion.svg");
        WriteAcceptancePng(pngPath, "OMML↔MT", 360, 112);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><text x=\"8\" y=\"78\" font-family=\"Cambria Math\" font-size=\"46\">α + ℏ + ∫ + A</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 112);

        var sources = new[]
        {
            new OmmlMathTypeAcceptanceFormula(
                "inline-hbar-greek",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>ℏ</mi><mo>+</mo><mi>α</mi><mo>=</mo><mi>β</mi></mrow></math>",
                "inline",
                false),
            new OmmlMathTypeAcceptanceFormula(
                "display-fraction-integral",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mfrac><mrow><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msub><mi>y</mi><mn>1</mn></msub></mrow><mi>z</mi></mfrac><mo>+</mo><msubsup><mo>∫</mo><mn>0</mn><mi>∞</mi></msubsup><mi>f</mi><mo>⁡</mo><mfenced><mi>x</mi></mfenced></mrow></math>",
                "block",
                true),
            new OmmlMathTypeAcceptanceFormula(
                "display-matrix",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>A</mi><mo>=</mo><mfenced open=\"[\" close=\"]\"><mtable><mtr><mtd><mi>α</mi></mtd><mtd><mi>β</mi></mtd></mtr><mtr><mtd><mi>γ</mi></mtd><mtd><mi>δ</mi></mtd></mtr></mtable></mfenced></mrow></math>",
                "block",
                false),
            new OmmlMathTypeAcceptanceFormula(
                "inline-accents-vector",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mover accent=\"true\"><mi>v</mi><mo>→</mo></mover><mo>+</mo><mover accent=\"true\"><mi>x</mi><mo>¯</mo></mover><mo>+</mo><mover accent=\"true\"><mi>y</mi><mo>^</mo></mover></mrow></math>",
                "inline",
                false),
        };
        const string existingMathTypeMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>q</mi><mo>=</mo><mn>7</mn></math>";
        const string ommlSentinelMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>s</mi><mo>=</mo><mn>9</mn></math>";

        var mathTypeProcessesBefore = SnapshotMathTypeProcessIds();
        if (mathTypeProcessesBefore.Count != 0)
            throw new InvalidOperationException(
                "OMML↔MathType acceptance requires MathType.exe to be absent before the test starts.");

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Add();
            document.Content.Text = "OMML MathType format-conversion acceptance\r";
            var service = new WordFormulaService(application);

            for (var index = 0; index < sources.Length; index++)
                InsertManagedOmmlBetweenProse(document, service, sources[index]);

            SelectDocumentEnd(document);
            service.InsertMathTypeOle(
                CreateOmmlMathTypeAcceptanceSession(
                    existingMathTypeMathMl,
                    "inline",
                    false,
                    FormulaOleContract.MathTypeOleMode),
                existingMathTypeMathMl,
                emfPath);
            AppendAcceptanceText(document, " existing-mathtype\r");

            AssertEqual(4, document.OMaths.Count,
                "OMML→MathType setup did not create exactly four OMML source equations.");
            AssertEqual(1, CountMathTypeOleShapes(document),
                "OMML→MathType setup did not retain exactly one pre-existing MathType object.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "OMML→MathType setup lost the numbered OMML source state.");

            SelectFirstOmmlEquations(document, 2);
            var selectionPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(2, selectionPlan.Targets.Count,
                "OMML→MathType selection capture did not isolate exactly two OMML equations.");
            AssertTrue(selectionPlan.Targets.All(target => !string.IsNullOrWhiteSpace(target.SourceMathMl)),
                "OMML→MathType selection capture did not preserve canonical source MathML.");
            var selectionResult = service.ApplyFormulaFormatConversionPlan(
                selectionPlan,
                PrepareOmmlMathTypeTargets(selectionPlan, emfPath));
            AssertEqual(2, selectionResult.FormulaCount,
                "OMML→MathType selection conversion did not replace two equations. Failures: "
                + string.Join(" | ", selectionResult.Failures));
            AssertEqual(0, selectionResult.FailedFormulaCount,
                $"OMML→MathType selection conversion failed: {string.Join(" | ", selectionResult.Failures)}");
            AssertEqual(2, document.OMaths.Count,
                "OMML→MathType selection conversion changed the wrong OMML count.");
            AssertEqual(3, CountMathTypeOleShapes(document),
                "OMML→MathType selection conversion did not preserve the pre-existing MathType object.");

            var documentPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(2, documentPlan.Targets.Count,
                "OMML→MathType document capture did not find only the remaining two OMML equations.");
            var documentResult = service.ApplyFormulaFormatConversionPlan(
                documentPlan,
                PrepareOmmlMathTypeTargets(documentPlan, emfPath));
            AssertEqual(2, documentResult.FormulaCount,
                "OMML→MathType document conversion did not replace the remaining two equations. Failures: "
                + string.Join(" | ", documentResult.Failures));
            AssertEqual(0, documentResult.FailedFormulaCount,
                $"OMML→MathType document conversion failed: {string.Join(" | ", documentResult.Failures)}");
            AssertEqual(0, document.OMaths.Count,
                "OMML→MathType document conversion left OMML equations behind.");
            AssertEqual(5, CountMathTypeOleShapes(document),
                "OMML→MathType document conversion produced the wrong MathType count.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "OMML→MathType conversion did not recreate the numbered source as one MTPlaceRef field.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after OMML→MathType core conversion");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after OMML→MathType core conversion");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "OMML→MathType core conversion");

            var mathTypePath = Path.Combine(
                artifactRoot,
                "OMML-To-MathType-Core-Acceptance.docx");
            document.SaveAs2(mathTypePath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                mathTypePath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(5, CountMathTypeOleShapes(document),
                "Saved/reopened OMML→MathType document lost a MathType object.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Saved/reopened OMML→MathType document changed numbering.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after OMML→MathType save/reopen");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after OMML→MathType save/reopen");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "OMML→MathType save/reopen");

            service = new WordFormulaService(application);
            SelectDocumentEnd(document);
            service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    ommlSentinelMathMl,
                    "inline",
                    false,
                    FormulaOleContract.WordOmmlMode),
                ommlSentinelMathMl);
            AppendAcceptanceText(document, " existing-omml-sentinel\r");
            AssertEqual(1, document.OMaths.Count,
                "MathType→OMML mixed-document setup did not create the OMML sentinel.");

            SelectFirstMathTypeEquations(document, 2);
            var reverseSelectionPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(2, reverseSelectionPlan.Targets.Count,
                "MathType→OMML selection capture did not isolate exactly two MathType equations.");
            AssertTrue(reverseSelectionPlan.Targets.All(target => !string.IsNullOrWhiteSpace(target.SourceMathMl)),
                "MathType→OMML selection capture did not preserve Equation Native MathML.");
            var reverseSelectionResult = service.ApplyFormulaFormatConversionPlan(
                reverseSelectionPlan,
                PrepareOmmlMathTypeTargets(reverseSelectionPlan, emfPath));
            AssertEqual(2, reverseSelectionResult.FormulaCount,
                "MathType→OMML selection conversion did not replace two equations. Failures: "
                + string.Join(" | ", reverseSelectionResult.Failures));
            AssertEqual(0, reverseSelectionResult.FailedFormulaCount,
                $"MathType→OMML selection conversion failed: {string.Join(" | ", reverseSelectionResult.Failures)}");
            AssertEqual(3, CountMathTypeOleShapes(document),
                "MathType→OMML selection conversion removed the wrong MathType objects.");
            AssertEqual(3, document.OMaths.Count,
                "MathType→OMML selection conversion did not retain the pre-existing OMML sentinel.");

            var reverseDocumentPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(3, reverseDocumentPlan.Targets.Count,
                "MathType→OMML document capture did not find only the remaining MathType objects.");
            var reverseDocumentResult = service.ApplyFormulaFormatConversionPlan(
                reverseDocumentPlan,
                PrepareOmmlMathTypeTargets(reverseDocumentPlan, emfPath));
            AssertEqual(3, reverseDocumentResult.FormulaCount,
                "MathType→OMML document conversion did not replace the remaining MathType equations. Failures: "
                + string.Join(" | ", reverseDocumentResult.Failures));
            AssertEqual(0, reverseDocumentResult.FailedFormulaCount,
                $"MathType→OMML document conversion failed: {string.Join(" | ", reverseDocumentResult.Failures)}");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "MathType→OMML document conversion left MathType OLE objects behind.");
            AssertEqual(6, document.OMaths.Count,
                "MathType→OMML document conversion produced the wrong OMML count or removed the sentinel.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "MathType→OMML conversion did not restore the numbered formula state.");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "MathType→OMML core conversion");

            var finalPath = Path.Combine(
                artifactRoot,
                "MathType-To-OMML-Core-Acceptance.docx");
            document.SaveAs2(finalPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Saved/reopened MathType→OMML document restored a MathType OLE object.");
            AssertEqual(6, document.OMaths.Count,
                "Saved/reopened MathType→OMML document lost a native OMML equation.");
            AssertEqual(1, CountManagedNumberedOmml(document),
                "Saved/reopened MathType→OMML document changed the numbered OMML state.");
            AssertOmmlSemanticCoverage(document);
            AssertOmmlConversionProseSurvived(document, sources);

            service = new WordFormulaService(application);
            const string nativeToken = "VT_PURE_NATIVE_OMML_SOURCE";
            AppendAcceptanceText(
                document,
                $" before-pure-native-omml {nativeToken} after-pure-native-omml\r");
            Word.Range? nativeRange = null;
            try
            {
                nativeRange = InsertPureNativeOmml(document, nativeToken, "n+1");
                nativeRange.Select();
                var nativePlan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
                AssertEqual(1, nativePlan.Targets.Count,
                    "Pure Word-native OMath selection was not recognized as one OMML conversion source.");
                AssertTrue(!nativePlan.Targets[0].SourceIsManagedOmml,
                    "Pure Word-native OMath was incorrectly treated as VisualTeX-managed OMML.");
                var nativeResult = service.ApplyFormulaFormatConversionPlan(
                    nativePlan,
                    PrepareOmmlMathTypeTargets(nativePlan, emfPath));
                AssertEqual(1, nativeResult.FormulaCount,
                    "Pure Word-native OMath did not convert to MathType. Failures: "
                    + string.Join(" | ", nativeResult.Failures));
            }
            finally { Release(nativeRange); }
            AssertEqual(1, CountMathTypeOleShapes(document),
                "Pure Word-native OMML conversion did not create exactly one Equation.DSMT4 object.");
            AssertEveryMathTypeProgId(document);
            AssertNoUnknownMathTypeGlyphTokens(document, "after pure native OMML→MathType");
            AssertVisibleMathTypePreviewsWithClipboardRetry(
                document,
                "after pure native OMML→MathType");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "pure native OMML→MathType");

            SelectFirstMathTypeEquations(document, 1);
            var nativeReversePlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: false,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(1, nativeReversePlan.Targets.Count,
                "Pure native round-trip MathType source was not captured for OMML restoration.");
            var nativeReverseResult = service.ApplyFormulaFormatConversionPlan(
                nativeReversePlan,
                PrepareOmmlMathTypeTargets(nativeReversePlan, emfPath));
            AssertEqual(1, nativeReverseResult.FormulaCount,
                "Pure native MathType→OMML round-trip failed: "
                + string.Join(" | ", nativeReverseResult.Failures));
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Pure native OMML round-trip left a MathType object behind.");
            AssertEqual(7, document.OMaths.Count,
                "Pure native OMML round-trip produced the wrong final OMath count.");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                finalPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(7, document.OMaths.Count,
                "Pure Word-native OMML round-trip did not survive save/reopen.");
            AssertEqual(0, CountMathTypeOleShapes(document),
                "Pure Word-native OMML round-trip restored a MathType object after save/reopen.");
            AssertTrue((document.Content.Text ?? string.Empty).IndexOf(
                           "before-pure-native-omml",
                           StringComparison.Ordinal) >= 0
                       && (document.Content.Text ?? string.Empty).IndexOf(
                           "after-pure-native-omml",
                           StringComparison.Ordinal) >= 0,
                "Pure Word-native OMML round-trip damaged adjacent prose.");
            AssertNoNewMathTypeProcess(mathTypeProcessesBefore, "full OMML↔MathType acceptance");

            Console.WriteLine(
                "[OMML↔MATHTYPE CORE] Selection + document conversion in both directions passed with VisualTeX-managed and pure Word-native OMath sources, mixed source types, inline/display formulas, numbering, hbar/Greek/fraction/integral/subscript/superscript/matrix/accent/vector semantics, adjacent prose preservation, non-empty MathType live previews, no OlePres, save/reopen persistence, and MathTypeProcessCount=0.");
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

    private sealed class OmmlMathTypeAcceptanceFormula
    {
        internal OmmlMathTypeAcceptanceFormula(
            string name,
            string mathMl,
            string displayMode,
            bool numbered)
        {
            Name = name;
            MathMl = mathMl;
            DisplayMode = displayMode;
            Numbered = numbered;
        }

        internal string Name { get; }
        internal string MathMl { get; }
        internal string DisplayMode { get; }
        internal bool Numbered { get; }
    }

    private static OfficeSessionDocument CreateOmmlMathTypeAcceptanceSession(
        string mathMl,
        string displayMode,
        bool numbered,
        string objectMode)
    {
        var latex = MathMlToLatexConverter.Convert(mathMl).Trim();
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "OMML MathType conversion acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = objectMode,
            Numbered = numbered,
            MathTypeNumberPosition = "right",
            FontSizePt = 12,
            ExportResult = new OfficeExportDocument
            {
                MathMl = mathMl,
                Width = 360,
                Height = 112,
                Baseline = 82,
            },
        };
    }

    private static IReadOnlyDictionary<string, PreparedWordBulkFormula> PrepareOmmlMathTypeTargets(
        WordFormulaFormatConversionPlan plan,
        string emfPath)
    {
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        foreach (var target in plan.Targets)
        {
            var mathMl = target.SourceMathMl
                ?? throw new InvalidDataException(
                    $"Format-conversion target '{target.Latex}' has no canonical source MathML.");
            prepared[target.Id] = new PreparedWordBulkFormula
            {
                Run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                },
                Session = CreateOmmlMathTypeAcceptanceSession(
                    mathMl,
                    target.DisplayMode,
                    target.Numbered,
                    plan.TargetMode),
                MathMl = mathMl,
                EmfPath = string.Equals(
                    plan.TargetMode,
                    FormulaOleContract.MathTypeOleMode,
                    StringComparison.Ordinal)
                    ? emfPath
                    : null,
            };
        }
        return prepared;
    }

    private static void SelectDocumentEnd(Word.Document document)
    {
        Word.Range? range = null;
        try
        {
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Select();
        }
        finally { Release(range); }
    }

    private static void AppendAcceptanceText(Word.Document document, string text)
    {
        Word.Range? range = null;
        try
        {
            range = document.Range(document.Content.End - 1, document.Content.End - 1);
            range.Text = text;
        }
        finally { Release(range); }
    }

    private static void SelectFirstOmmlEquations(Word.Document document, int count)
    {
        Word.OMaths? maths = null;
        Word.OMath? first = null;
        Word.OMath? last = null;
        Word.Range? firstRange = null;
        Word.Range? lastRange = null;
        Word.Range? selection = null;
        try
        {
            maths = document.OMaths;
            if (maths.Count < count)
                throw new InvalidDataException(
                    $"Expected at least {count} OMML equations, actual {maths.Count}.");
            first = maths[1];
            last = maths[count];
            firstRange = first.Range;
            lastRange = last.Range;
            selection = document.Range(firstRange.Start, lastRange.End);
            selection.Select();
        }
        finally
        {
            Release(selection);
            Release(lastRange);
            Release(firstRange);
            Release(last);
            Release(first);
            Release(maths);
        }
    }

    private static void SelectFirstMathTypeEquations(Word.Document document, int count)
    {
        var starts = new List<(int Start, int End)>();
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            Word.Range? range = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                range = shape.Range;
                starts.Add((range.Start, range.End));
            }
            finally
            {
                Release(range);
                Release(shape);
            }
        }
        starts.Sort((left, right) => left.Start.CompareTo(right.Start));
        if (starts.Count < count)
            throw new InvalidDataException(
                $"Expected at least {count} MathType equations, actual {starts.Count}.");
        Word.Range? selection = null;
        try
        {
            selection = document.Range(starts[0].Start, starts[count - 1].End);
            selection.Select();
        }
        finally { Release(selection); }
    }

    private static int CountManagedNumberedOmml(Word.Document document)
    {
        var count = 0;
        foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
        {
            Word.Bookmark? bookmark = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                var metadata = WordOmmlFormulaStore.TryRead(document, bookmark);
                if (metadata?.Numbered == true) count++;
            }
            finally { Release(bookmark); }
        }
        return count;
    }

    private static void AssertEveryMathTypeProgId(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                count++;
                AssertEqual("Equation.DSMT4", shape.OLEFormat.ProgID,
                    $"MathType conversion output #{count} lost ProgID=Equation.DSMT4.");
                var mathMl = MathTypeOleStorage.ReadMathMl(shape);
                AssertTrue(!string.IsNullOrWhiteSpace(mathMl),
                    $"MathType conversion output #{count} has no readable Equation Native MathML.");
            }
            finally { Release(shape); }
        }
        AssertTrue(count > 0, "No MathType conversion outputs were available for ProgID validation.");
    }

    private static void AssertNoNewMathTypeProcess(
        IReadOnlyCollection<int> baseline,
        string stage)
    {
        var started = SnapshotMathTypeProcessIds().Except(baseline).ToArray();
        AssertEqual(0, started.Length,
            $"MathType.exe started during {stage}: {string.Join(", ", started)}");
    }

    private static void InsertManagedOmmlBetweenProse(
        Word.Document document,
        WordFormulaService service,
        OmmlMathTypeAcceptanceFormula source)
    {
        var token = "VT_OMML_SOURCE_" + Guid.NewGuid().ToString("N");
        var surroundingText = string.Equals(
                source.DisplayMode,
                "block",
                StringComparison.Ordinal)
            ? $"before-{source.Name}\r{token}\r after-{source.Name}\r"
            : $"before-{source.Name} {token} after-{source.Name}\r";
        AppendAcceptanceText(document, surroundingText);
        Word.Range? insertion = null;
        try
        {
            insertion = FindAcceptanceTextRange(document, token);
            insertion.Text = string.Empty;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            insertion.Select();
            service.InsertOmml(
                CreateOmmlMathTypeAcceptanceSession(
                    source.MathMl,
                    source.DisplayMode,
                    source.Numbered,
                    FormulaOleContract.WordOmmlMode),
                source.MathMl);
        }
        finally { Release(insertion); }
    }

    private static Word.Range FindAcceptanceTextRange(
        Word.Document document,
        string text)
    {
        Word.Range? range = null;
        Word.Find? find = null;
        try
        {
            range = document.Content.Duplicate;
            find = range.Find;
            find.ClearFormatting();
            find.Text = text;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            if (!find.Execute())
                throw new InvalidDataException(
                    $"Acceptance placeholder '{text}' was not found in the Word document.");
            var result = range.Duplicate;
            Release(range);
            range = null;
            return result;
        }
        finally
        {
            Release(find);
            Release(range);
        }
    }

    private static void AssertVisibleMathTypePreviewsWithClipboardRetry(
        Word.Document document,
        string stage)
    {
        Exception? last = null;
        for (var attempt = 1; attempt <= 8; attempt++)
        {
            try
            {
                AssertVisibleMathTypeEmfPreviews(document, stage);
                return;
            }
            catch (InvalidDataException error)
            {
                last = error;
                if (error.ToString().IndexOf(
                        "CLIPBRD_E_CANT_OPEN",
                        StringComparison.OrdinalIgnoreCase) < 0
                    && error.ToString().IndexOf(
                        "clipboard stayed busy",
                        StringComparison.OrdinalIgnoreCase) < 0)
                    throw;
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(400 * attempt);
            }
        }
        throw new InvalidDataException(
            $"Word clipboard remained busy after repeated live-preview validation attempts {stage}.",
            last);
    }

    private static Word.Range InsertPureNativeOmml(
        Word.Document document,
        string placeholder,
        string linearText)
    {
        Word.Range? source = null;
        Word.Range? added = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? result = null;
        try
        {
            source = FindAcceptanceTextRange(document, placeholder);
            var start = source.Start;
            source.Text = linearText;
            source.SetRange(start, start + linearText.Length);
            added = document.OMaths.Add(source);
            maths = added.OMaths;
            if (maths.Count != 1)
                throw new InvalidDataException(
                    "Word did not create exactly one native OMath from the linear source range.");
            math = maths[1];
            math.BuildUp();
            result = math.Range.Duplicate;
            var duplicate = result;
            result = null;
            return duplicate;
        }
        finally
        {
            Release(result);
            Release(math);
            Release(maths);
            Release(added);
            Release(source);
        }
    }

    private static void AssertOmmlConversionProseSurvived(
        Word.Document document,
        IReadOnlyList<OmmlMathTypeAcceptanceFormula> sources)
    {
        var text = document.Content.Text ?? string.Empty;
        var previous = -1;
        foreach (var source in sources)
        {
            var before = "before-" + source.Name;
            var after = "after-" + source.Name;
            var beforeIndex = text.IndexOf(before, StringComparison.Ordinal);
            var afterIndex = text.IndexOf(after, StringComparison.Ordinal);
            AssertTrue(beforeIndex >= 0 && afterIndex > beforeIndex,
                $"OMML↔MathType conversion damaged adjacent prose around '{source.Name}'.");
            AssertTrue(beforeIndex > previous,
                $"OMML↔MathType conversion changed the document order near '{source.Name}'.");
            previous = afterIndex;
        }
    }

    private static void AssertOmmlSemanticCoverage(Word.Document document)
    {
        var latex = new List<string>();
        Word.OMaths? maths = null;
        try
        {
            maths = document.OMaths;
            for (var index = 1; index <= maths.Count; index++)
            {
                Word.OMath? math = null;
                Word.Range? range = null;
                try
                {
                    math = maths[index];
                    range = math.Range.Duplicate;
                    var metadata = WordOmmlNativeSource.CreateForNative(document, range);
                    latex.Add((metadata.Latex ?? string.Empty).Replace(" ", string.Empty));
                }
                finally
                {
                    Release(range);
                    Release(math);
                }
            }
        }
        finally { Release(maths); }

        var joined = string.Join(" | ", latex);
        AssertTrue(joined.IndexOf("hbar", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("ℏ", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost hbar: {joined}");
        AssertTrue(joined.IndexOf("alpha", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("α", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost Greek alpha: {joined}");
        AssertTrue(joined.IndexOf("frac", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("/", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost fraction structure: {joined}");
        AssertTrue(joined.IndexOf("int", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("∫", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost integral structure: {joined}");
        AssertTrue(joined.IndexOf("matrix", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("begin", StringComparison.OrdinalIgnoreCase) >= 0,
            $"Final OMML semantics lost matrix structure: {joined}");
        AssertTrue(joined.IndexOf("vec", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("over", StringComparison.OrdinalIgnoreCase) >= 0
                   || joined.IndexOf("→", StringComparison.Ordinal) >= 0,
            $"Final OMML semantics lost vector/accent structure: {joined}");
    }
}

using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordBulkImportOmmlStressAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var logPath = Path.Combine(artifactRoot, "omml-bulk-stress.log");
        var outputPath = Path.Combine(artifactRoot, "OMML-Bulk-Stress.docx");
        var initialOmmlProbePath = Path.Combine(
            artifactRoot,
            "OMML-Before-First-MathType-Roundtrip.docx");
        var mathTypeIntermediateOpenXmlPath = Path.Combine(
            artifactRoot,
            "MathType-Intermediate-Document.xml");
        // VisualTeX.Formula.1 accepts previews only from the product-owned Office
        // temp root. Keep both OLE target previews there so this same stress
        // document can traverse OMML, MathType and VisualTeX without weakening
        // the production path validation.
        var previewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "omml-bulk-stress-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(previewRoot);
        var pngPath = Path.Combine(previewRoot, "omml-bulk-roundtrip-preview.png");
        var svgPath = Path.Combine(previewRoot, "omml-bulk-roundtrip-preview.svg");
        WriteAcceptancePng(pngPath, "omml-bulk-roundtrip", 360, 112);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><text x=\"8\" y=\"78\" font-family=\"Cambria Math\" font-size=\"42\">∑∫ ∂ψ/∂x</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
            svgPath,
            360,
            112);
        try { File.Delete(logPath); } catch { }
        try { File.Delete(outputPath); } catch { }
        try { File.Delete(initialOmmlProbePath); } catch { }
        try { File.Delete(mathTypeIntermediateOpenXmlPath); } catch { }

        var source = CreateOmmlBulkStressSource();
        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.Omml);
        parsed.NumberDisplayFormulas = true;
        AssertEqual(20, parsed.FormulaCount,
            "OMML stress source must contain exactly twenty formulas.");
        AssertEqual(10, parsed.DisplayFormulaCount,
            "OMML stress source must contain exactly ten display formulas.");
        AssertEqual(10, parsed.InlineFormulaCount,
            "OMML stress source must contain exactly ten inline formulas.");
        AssertEqual(0, parsed.Warnings.Count,
            "OMML stress source unexpectedly produced parser warnings.");

        var previousSource = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE");
        var previousSourcePath = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH");
        var previousFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT");
        var previousMode = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE");
        var previousNumber = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS");
        var previousLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");

        Word.Application? application = null;
        Word.Document? returnDocument = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            // Default to an isolated WINWORD instance so a currently loaded
            // installed add-in cannot race this source-tree acceptance's
            // DocumentOpen handler or touch the user's live documents. Explicit
            // VISUALTEX_VSTO_ACCEPTANCE_ATTACH_WORD=1 retains the diagnostic mode
            // for an intentionally selected active instance.
            application = CreateWordApplication(visible: false);
            if (AttachActiveWord)
                returnDocument = application.ActiveDocument;
            document = application.Documents.Add();
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE", source);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "omml");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);

            addIn = new ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);
            addIn.OnBulkImport(new object());
            var elapsedMs = WaitForBulkImportCompletion(
                logPath,
                TimeSpan.FromMinutes(4));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(45));

            var service = new WordFormulaService(application);
            AssertOmmlBulkStressState(document, service, "fresh OMML stress import");
            AssertOmmlStressProseOrderAndStyles(document);
            var separatorCount = CountStructuralParagraphsBetweenTables(document);
            Console.WriteLine(
                $"[OMML BULK STRUCTURE] paragraphs={document.Paragraphs.Count}; tables={document.Tables.Count}; "
                + $"interTableParagraphs={separatorCount}; fields={document.Fields.Count}; bookmarks={document.Bookmarks.Count}");
            AssertEqual(0, separatorCount,
                "Direct OMML bulk import left structural blank paragraphs between consecutive numbered formulas.");
            AssertEqual(10, document.Fields.Count,
                "Direct OMML bulk import did not retain all ten SEQ fields before save.");
            AssertEqual(50, document.Bookmarks.Count,
                "Direct OMML bulk import did not retain all twenty VTOMML and thirty numbering bookmarks before save.");

            document.SaveAs2(
                outputPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            File.Copy(outputPath, initialOmmlProbePath, overwrite: true);

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);
            reopened.Activate();
            var reopenedService = new WordFormulaService(application);
            AssertOmmlBulkStressState(
                reopened,
                reopenedService,
                "save/reopened OMML stress import");
            AssertOmmlStressProseOrderAndStyles(reopened);
            AssertEqual(0, CountStructuralParagraphsBetweenTables(reopened),
                "Save/reopen exposed structural blank paragraphs between numbered OMML formulas.");
            AssertEqual(10, reopened.Fields.Count,
                "Save/reopen lost one or more numbered OMML SEQ fields.");
            AssertEqual(50, reopened.Bookmarks.Count,
                "Save/reopen lost one or more VTOMML or numbering identity bookmarks.");
            var baselineOmmlParagraphCount = reopened.Paragraphs.Count;

            var ommlToMathTypePlan =
                reopenedService.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
            AssertEqual(20, ommlToMathTypePlan.Targets.Count,
                "The grouped OMML document did not expose all twenty OMML→MathType targets.");
            var originalSignatures = ommlToMathTypePlan.Targets
                .Select(target => MathTypeMtefCodec.SemanticSignature(
                    target.SourceMathMl
                    ?? throw new InvalidDataException(
                        "A grouped OMML source has no canonical MathML.")))
                .OrderBy(signature => signature, StringComparer.Ordinal)
                .ToArray();
            var toMathTypeResult = reopenedService.ApplyFormulaFormatConversionPlan(
                ommlToMathTypePlan,
                PrepareOmmlMathTypeTargets(ommlToMathTypePlan, emfPath));
            Console.WriteLine(
                $"[GROUPED OMML→MATHTYPE] converted={toMathTypeResult.FormulaCount}; failed={toMathTypeResult.FailedFormulaCount}; failures={string.Join(" | ", toMathTypeResult.Failures)}");
            AssertEqual(20, toMathTypeResult.FormulaCount,
                "Grouped OMML→MathType did not convert the complete batch.");
            AssertEqual(0, toMathTypeResult.FailedFormulaCount,
                "Grouped OMML→MathType reported failures.");
            AssertEqual(0, reopened.OMaths.Count,
                "Grouped OMML→MathType left native OMath sources behind.");
            AssertEqual(20, CountMathTypeOleShapes(reopened),
                "Grouped OMML→MathType did not create twenty MathType OLE objects.");
            AssertEqual(10, CountMathTypePlaceRefFields(reopened),
                "Grouped OMML→MathType did not preserve all ten equation numbers.");
            AssertEqual(0, reopened.Tables.Count,
                "Grouped OMML→MathType left an empty OMML numbering table behind.");
            AssertOleStressProseOrderAndStyles(
                reopened,
                FormulaOleContract.MathTypeOleMode,
                "first OMML→MathType leg");
            WriteDocumentOpenXmlProbe(
                reopened,
                mathTypeIntermediateOpenXmlPath);

            var mathTypeToOmmlPlan =
                reopenedService.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(20, mathTypeToOmmlPlan.Targets.Count,
                "The MathType result did not expose all twenty reverse targets.");
            Console.WriteLine(
                "[MATHTYPE PRECEDING BLANKS] "
                + string.Join(
                    ",",
                    mathTypeToOmmlPlan.Targets
                        .Where(target => target.Numbered)
                        .OrderBy(target => target.SourceStart)
                        .Select(target =>
                            $"{target.SourceStart}:{target.PrecedingPlainBlankParagraphCount}")));
            var mathTypeSignatures = mathTypeToOmmlPlan.Targets
                .Select(target => MathTypeMtefCodec.SemanticSignature(
                    target.SourceMathMl
                    ?? throw new InvalidDataException(
                        "A round-trip MathType source has no canonical MathML.")))
                .OrderBy(signature => signature, StringComparer.Ordinal)
                .ToArray();
            AssertEqual(
                string.Join("\n", originalSignatures),
                string.Join("\n", mathTypeSignatures),
                "Grouped OMML→MathType changed formula semantics before the reverse leg.");

            var toOmmlResult = reopenedService.ApplyFormulaFormatConversionPlan(
                mathTypeToOmmlPlan,
                PrepareOmmlMathTypeTargets(mathTypeToOmmlPlan, emfPath));
            Console.WriteLine(
                $"[GROUPED MATHTYPE→OMML] converted={toOmmlResult.FormulaCount}; failed={toOmmlResult.FailedFormulaCount}; failures={string.Join(" | ", toOmmlResult.Failures)}");
            AssertEqual(20, toOmmlResult.FormulaCount,
                "Grouped MathType→OMML did not convert the complete batch.");
            AssertEqual(0, toOmmlResult.FailedFormulaCount,
                "Grouped MathType→OMML reported failures.");
            AssertOmmlBulkStressState(
                reopened,
                reopenedService,
                "OMML→MathType→OMML grouped round-trip");
            AssertOmmlStressProseOrderAndStyles(reopened);
            AssertEqual(0, CountStructuralParagraphsBetweenTables(reopened),
                "OMML→MathType→OMML recreated structural blank paragraphs.");
            reopened.Save();
            Console.WriteLine(
                $"[FIRST ROUNDTRIP STRUCTURE] beforeParagraphs={baselineOmmlParagraphCount}; afterParagraphs={reopened.Paragraphs.Count}; beforeProbe={initialOmmlProbePath}; afterProbe={outputPath}");
            AssertEqual(baselineOmmlParagraphCount, reopened.Paragraphs.Count,
                "OMML→MathType→OMML changed the canonical OMML paragraph count.");
            AssertEqual(1, reopened.Tables.Count,
                "OMML→MathType→OMML did not regroup consecutive numbered formulas into one table.");
            AssertEqual(10, reopened.Fields.Count,
                "OMML→MathType→OMML did not restore all ten SEQ fields.");
            AssertEqual(50, reopened.Bookmarks.Count,
                "OMML→MathType→OMML did not restore all OMML identity and numbering bookmarks.");

            var ommlToVisualTeXPlan =
                reopenedService.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(20, ommlToVisualTeXPlan.Targets.Count,
                "The first OMML round-trip result did not expose twenty OMML→VisualTeX targets.");
            AssertStressSemanticSignatures(
                originalSignatures,
                ommlToVisualTeXPlan,
                "OMML→VisualTeX capture");
            var visualTeXRenderMathMl = ommlToVisualTeXPlan.Targets.ToDictionary(
                StressFormulaIdentity,
                target => target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"OMML→VisualTeX target '{target.Latex}' has no canonical MathML."),
                StringComparer.Ordinal);
            var toVisualTeXResult = reopenedService.ApplyFormulaFormatConversionPlan(
                ommlToVisualTeXPlan,
                PrepareOmmlVisualTeXStressTargets(
                    ommlToVisualTeXPlan,
                    pngPath,
                    emfPath));
            Console.WriteLine(
                $"[GROUPED OMML→VISUALTEX] converted={toVisualTeXResult.FormulaCount}; failed={toVisualTeXResult.FailedFormulaCount}; failures={string.Join(" | ", toVisualTeXResult.Failures)}");
            AssertEqual(20, toVisualTeXResult.FormulaCount,
                "Grouped OMML→VisualTeX did not convert the complete batch.");
            AssertEqual(0, toVisualTeXResult.FailedFormulaCount,
                "Grouped OMML→VisualTeX reported failures.");
            AssertEqual(0, reopened.OMaths.Count,
                "Grouped OMML→VisualTeX left native OMath sources behind.");
            AssertEqual(20, CountVisualTeXNativeOleShapes(reopened),
                "Grouped OMML→VisualTeX did not create twenty VisualTeX OLE objects.");
            AssertEqual(0, CountMathTypeOleShapes(reopened),
                "Grouped OMML→VisualTeX left MathType objects behind.");
            AssertEqual(10, CountInstalledVisualTeXNumberedFormulaHosts(reopened),
                "Grouped OMML→VisualTeX did not preserve ten numbered formula hosts.");
            AssertEqual(10, CountVisualTeXNumberingBookmarkTriples(reopened),
                "Grouped OMML→VisualTeX did not preserve all ten numbering bookmark triples.");
            AssertEqual(0, reopened.Tables.Count,
                "Grouped OMML→VisualTeX left an empty OMML numbering table behind.");
            AssertOleStressProseOrderAndStyles(
                reopened,
                FormulaOleContract.NativeOleMode,
                "OMML→VisualTeX leg");

            var visualTeXToOmmlPlan =
                reopenedService.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.NativeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(20, visualTeXToOmmlPlan.Targets.Count,
                "The VisualTeX result did not expose twenty reverse OMML targets.");
            AssertStressSemanticSignatures(
                originalSignatures,
                visualTeXToOmmlPlan,
                "VisualTeX→OMML capture",
                visualTeXRenderMathMl);
            var visualTeXToOmmlResult =
                reopenedService.ApplyFormulaFormatConversionPlan(
                    visualTeXToOmmlPlan,
                    PrepareOmmlVisualTeXStressTargets(
                        visualTeXToOmmlPlan,
                        pngPath,
                        emfPath,
                        visualTeXRenderMathMl));
            Console.WriteLine(
                $"[GROUPED VISUALTEX→OMML] converted={visualTeXToOmmlResult.FormulaCount}; failed={visualTeXToOmmlResult.FailedFormulaCount}; failures={string.Join(" | ", visualTeXToOmmlResult.Failures)}");
            AssertEqual(20, visualTeXToOmmlResult.FormulaCount,
                "Grouped VisualTeX→OMML did not convert the complete batch.");
            AssertEqual(0, visualTeXToOmmlResult.FailedFormulaCount,
                "Grouped VisualTeX→OMML reported failures.");
            AssertCanonicalOmmlStressState(
                reopened,
                reopenedService,
                originalSignatures,
                baselineOmmlParagraphCount,
                "OMML→VisualTeX→OMML round-trip");

            // Persist and reopen between format families. This exercises the real
            // DocumentOpen numbering/identity repair path before a second
            // conversion cycle instead of validating only one in-memory graph.
            reopened.Save();
            reopened.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(reopened);
            reopened = null;
            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);
            reopened.Activate();
            reopenedService = new WordFormulaService(application);
            AssertCanonicalOmmlStressState(
                reopened,
                reopenedService,
                originalSignatures,
                baselineOmmlParagraphCount,
                "post-VisualTeX save/reopen");

            var secondOmmlToMathTypePlan =
                reopenedService.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
            AssertEqual(20, secondOmmlToMathTypePlan.Targets.Count,
                "Second OMML→MathType cycle did not capture twenty formulas.");
            AssertStressSemanticSignatures(
                originalSignatures,
                secondOmmlToMathTypePlan,
                "second OMML→MathType capture");
            var secondToMathTypeResult =
                reopenedService.ApplyFormulaFormatConversionPlan(
                    secondOmmlToMathTypePlan,
                    PrepareOmmlMathTypeTargets(
                        secondOmmlToMathTypePlan,
                        emfPath));
            Console.WriteLine(
                $"[SECOND OMML→MATHTYPE] converted={secondToMathTypeResult.FormulaCount}; failed={secondToMathTypeResult.FailedFormulaCount}; failures={string.Join(" | ", secondToMathTypeResult.Failures)}");
            AssertEqual(20, secondToMathTypeResult.FormulaCount,
                "Second OMML→MathType cycle did not convert the complete batch.");
            AssertEqual(0, secondToMathTypeResult.FailedFormulaCount,
                "Second OMML→MathType cycle reported failures.");
            AssertEqual(20, CountMathTypeOleShapes(reopened),
                "Second OMML→MathType cycle did not retain twenty MathType objects.");
            AssertEqual(10, CountMathTypePlaceRefFields(reopened),
                "Second OMML→MathType cycle did not retain ten MTPlaceRef fields.");
            AssertEqual(0, reopened.OMaths.Count,
                "Second OMML→MathType cycle left OMML sources behind.");
            AssertEqual(0, reopened.Tables.Count,
                "Second OMML→MathType cycle left an empty numbering table behind.");
            AssertOleStressProseOrderAndStyles(
                reopened,
                FormulaOleContract.MathTypeOleMode,
                "second OMML→MathType leg");

            var secondMathTypeToOmmlPlan =
                reopenedService.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.MathTypeOleMode,
                    FormulaOleContract.WordOmmlMode);
            AssertEqual(20, secondMathTypeToOmmlPlan.Targets.Count,
                "Second MathType→OMML cycle did not capture twenty formulas.");
            AssertStressSemanticSignatures(
                originalSignatures,
                secondMathTypeToOmmlPlan,
                "second MathType→OMML capture");
            var secondToOmmlResult =
                reopenedService.ApplyFormulaFormatConversionPlan(
                    secondMathTypeToOmmlPlan,
                    PrepareOmmlMathTypeTargets(
                        secondMathTypeToOmmlPlan,
                        emfPath));
            Console.WriteLine(
                $"[SECOND MATHTYPE→OMML] converted={secondToOmmlResult.FormulaCount}; failed={secondToOmmlResult.FailedFormulaCount}; failures={string.Join(" | ", secondToOmmlResult.Failures)}");
            AssertEqual(20, secondToOmmlResult.FormulaCount,
                "Second MathType→OMML cycle did not convert the complete batch.");
            AssertEqual(0, secondToOmmlResult.FailedFormulaCount,
                "Second MathType→OMML cycle reported failures.");
            AssertCanonicalOmmlStressState(
                reopened,
                reopenedService,
                originalSignatures,
                baselineOmmlParagraphCount,
                "second OMML→MathType→OMML round-trip");

            reopened.Save();
            reopened.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(reopened);
            reopened = null;
            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: true);
            reopened.Activate();
            reopenedService = new WordFormulaService(application);
            AssertCanonicalOmmlStressState(
                reopened,
                reopenedService,
                originalSignatures,
                baselineOmmlParagraphCount,
                "final multi-round save/reopen");

            Console.WriteLine(
                $"[OMML BULK STRESS PASS] 10 inline + 10 display formulas retained semantics, native OMML structure, identities, numbering, prose and zero table gaps through OMML↔MathType, OMML↔VisualTeX, a second OMML↔MathType cycle and two save/reopen boundaries; initialImportMs={elapsedMs}.");
            Console.WriteLine("Artifact: " + outputPath);
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE", previousSource);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", previousSourcePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", previousFormat);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", previousMode);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS", previousNumber);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", previousLog);
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); }
                catch { }
            }
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { returnDocument?.Activate(); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(returnDocument);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordOmmlStressProbeComparison(
        string artifactRoot)
    {
        var beforePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_OMML_STRESS_BEFORE");
        var afterPath = Environment.GetEnvironmentVariable(
            "VISUALTEX_OMML_STRESS_AFTER");
        if (string.IsNullOrWhiteSpace(beforePath)
            || !File.Exists(beforePath)
            || string.IsNullOrWhiteSpace(afterPath)
            || !File.Exists(afterPath))
            throw new FileNotFoundException(
                "Set VISUALTEX_OMML_STRESS_BEFORE and VISUALTEX_OMML_STRESS_AFTER to the two probe DOCX files.");

        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);

            WordFormulaFormatConversionPlan Capture(string path)
            {
                document = application.Documents.Open(
                    path,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
                document.Activate();
                var service = new WordFormulaService(application);
                var plan = service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: true,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.MathTypeOleMode);
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                Release(document);
                document = null;
                return plan;
            }

            var beforePlan = Capture(beforePath);
            var afterPlan = Capture(afterPath);
            var beforeTargets = beforePlan.Targets
                .OrderBy(target => target.SourceStart)
                .ToArray();
            var afterTargets = afterPlan.Targets
                .OrderBy(target => target.SourceStart)
                .ToArray();
            AssertEqual(20, beforeTargets.Length,
                "The before probe does not contain twenty OMML formulas.");
            AssertEqual(20, afterTargets.Length,
                "The after probe does not contain twenty OMML formulas.");

            var mismatchCount = 0;
            for (var index = 0; index < beforeTargets.Length; index++)
            {
                var beforeMathMl = beforeTargets[index].SourceMathMl
                    ?? throw new InvalidDataException(
                        $"Before probe formula {index + 1} has no MathML.");
                var afterMathMl = afterTargets[index].SourceMathMl
                    ?? throw new InvalidDataException(
                        $"After probe formula {index + 1} has no MathML.");
                var beforeSignature =
                    MathTypeMtefCodec.SemanticSignature(beforeMathMl);
                var afterSignature =
                    MathTypeMtefCodec.SemanticSignature(afterMathMl);
                if (string.Equals(
                        beforeSignature,
                        afterSignature,
                        StringComparison.Ordinal))
                    continue;

                mismatchCount++;
                var prefix = $"semantic-diff-{index + 1:D2}";
                File.WriteAllText(
                    Path.Combine(artifactRoot, prefix + "-before.xml"),
                    beforeMathMl);
                File.WriteAllText(
                    Path.Combine(artifactRoot, prefix + "-after.xml"),
                    afterMathMl);
                var differenceIndex = 0;
                var sharedLength = Math.Min(
                    beforeSignature.Length,
                    afterSignature.Length);
                while (differenceIndex < sharedLength
                       && beforeSignature[differenceIndex]
                           == afterSignature[differenceIndex])
                    differenceIndex++;
                var windowStart = Math.Max(0, differenceIndex - 80);
                var beforeWindow = beforeSignature.Substring(
                    windowStart,
                    Math.Min(240, beforeSignature.Length - windowStart));
                var afterWindow = afterSignature.Substring(
                    windowStart,
                    Math.Min(240, afterSignature.Length - windowStart));
                Console.WriteLine(
                    $"[OMML SEMANTIC DIFF] index={index + 1}; beforeLatex={beforeTargets[index].Latex}; afterLatex={afterTargets[index].Latex}; firstDifference={differenceIndex}");
                Console.WriteLine("  before=" + beforeWindow);
                Console.WriteLine("  after =" + afterWindow);
            }

            Console.WriteLine(
                $"[OMML SEMANTIC PROBE] compared={beforeTargets.Length}; mismatches={mismatchCount}; artifacts={artifactRoot}");
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); }
            catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void WriteDocumentOpenXmlProbe(
        Word.Document document,
        string path)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            File.WriteAllText(path, content.WordOpenXML);
        }
        finally
        {
            Release(content);
        }
    }

    private static string StressFormulaIdentity(
        WordFormulaFormatConversionTarget target) =>
        $"{target.DisplayMode}|{target.Numbered}|{target.Latex.Trim()}";

    private static IReadOnlyDictionary<string, PreparedWordBulkFormula>
        PrepareOmmlVisualTeXStressTargets(
            WordFormulaFormatConversionPlan plan,
            string pngPath,
            string emfPath,
            IReadOnlyDictionary<string, string>? renderedMathMl = null)
    {
        var targetIsVisualTeX = string.Equals(
            plan.TargetMode,
            FormulaOleContract.NativeOleMode,
            StringComparison.Ordinal);
        var targetIsOmml = string.Equals(
            plan.TargetMode,
            FormulaOleContract.WordOmmlMode,
            StringComparison.Ordinal);
        if (!targetIsVisualTeX && !targetIsOmml)
            throw new InvalidDataException(
                $"Unsupported OMML/VisualTeX stress target mode '{plan.TargetMode}'.");

        var prepared =
            new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        foreach (var target in plan.Targets)
        {
            var mathMl = target.SourceMathMl;
            if (string.IsNullOrWhiteSpace(mathMl)
                && renderedMathMl is not null)
                renderedMathMl.TryGetValue(StressFormulaIdentity(target), out mathMl);
            if (string.IsNullOrWhiteSpace(mathMl))
                throw new InvalidDataException(
                    $"OMML/VisualTeX stress target '{target.Latex}' has no canonical MathML.");
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
                    plan.TargetMode,
                    mathMl),
                MathMl = mathMl,
                PngPath = targetIsVisualTeX ? pngPath : null,
                EmfPath = targetIsVisualTeX ? emfPath : null,
            };
        }
        return prepared;
    }

    private static void AssertStressSemanticSignatures(
        IReadOnlyList<string> expectedSignatures,
        WordFormulaFormatConversionPlan plan,
        string context,
        IReadOnlyDictionary<string, string>? renderedMathMl = null)
    {
        AssertEqual(20, plan.Targets.Count,
            $"{context}: conversion plan did not retain twenty formulas.");
        AssertEqual(
            10,
            plan.Targets.Count(target =>
                string.Equals(
                    target.DisplayMode,
                    "inline",
                    StringComparison.Ordinal)),
            $"{context}: inline target count changed.");
        AssertEqual(
            10,
            plan.Targets.Count(target =>
                string.Equals(
                    target.DisplayMode,
                    "block",
                    StringComparison.Ordinal)),
            $"{context}: display target count changed.");
        AssertEqual(
            10,
            plan.Targets.Count(target => target.Numbered),
            $"{context}: numbered target count changed.");

        var actualSignatures = plan.Targets
            .Select(target =>
            {
                var mathMl = target.SourceMathMl;
                if (string.IsNullOrWhiteSpace(mathMl)
                    && renderedMathMl is not null)
                    renderedMathMl.TryGetValue(StressFormulaIdentity(target), out mathMl);
                if (string.IsNullOrWhiteSpace(mathMl))
                    throw new InvalidDataException(
                        $"{context}: target '{target.Latex}' has no canonical MathML.");
                return MathTypeMtefCodec.SemanticSignature(mathMl);
            })
            .OrderBy(signature => signature, StringComparer.Ordinal)
            .ToArray();
        AssertEqual(
            string.Join("\n", expectedSignatures),
            string.Join("\n", actualSignatures),
            $"{context}: formula semantics changed.");
    }

    private static void AssertCanonicalOmmlStressState(
        Word.Document document,
        WordFormulaService service,
        IReadOnlyList<string> expectedSignatures,
        int expectedParagraphCount,
        string context)
    {
        AssertOmmlBulkStressState(document, service, context);
        AssertOmmlStressProseOrderAndStyles(document);
        AssertEqual(0, CountStructuralParagraphsBetweenTables(document),
            $"{context}: structural blank paragraphs appeared between numbered OMML rows.");
        AssertEqual(expectedParagraphCount, document.Paragraphs.Count,
            $"{context}: canonical OMML paragraph count changed.");
        AssertEqual(1, document.Tables.Count,
            $"{context}: numbered OMML rows are not grouped into one managed table.");
        AssertEqual(10, document.Fields.Count,
            $"{context}: expected ten direct SEQ fields.");
        AssertEqual(50, document.Bookmarks.Count,
            $"{context}: OMML identity/numbering bookmark inventory changed.");

        var verificationPlan = service.CaptureFormulaFormatConversionPlan(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.MathTypeOleMode);
        AssertStressSemanticSignatures(
            expectedSignatures,
            verificationPlan,
            context + " semantic recapture");
    }

    private static void AssertOleStressProseOrderAndStyles(
        Word.Document document,
        string objectMode,
        string context)
    {
        var text = document.Content.Text ?? string.Empty;
        var cursor = -1;
        for (var index = 1; index <= 10; index++)
        {
            var marker = index <= 5
                ? $"OMML-PROSE-{index:00}-BEFORE"
                : $"OMML-PROSE-{index:00}-AFTER";
            var position = text.IndexOf(
                marker,
                cursor + 1,
                StringComparison.Ordinal);
            AssertTrue(position > cursor,
                $"{context}: prose marker {marker} was lost or reordered.");
            cursor = position;
        }

        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Style? style = null;
            Word.InlineShapes? shapes = null;
            Word.InlineShape? shape = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                var paragraphText = range.Text ?? string.Empty;
                if (paragraphText.IndexOf(
                        "OMML-PROSE-",
                        StringComparison.Ordinal) < 0)
                    continue;

                style = range.get_Style() as Word.Style;
                var styleName = style?.NameLocal ?? string.Empty;
                AssertTrue(
                    styleName.IndexOf(
                        "MTDisplayEquation",
                        StringComparison.OrdinalIgnoreCase) < 0,
                    $"{context}: ordinary prose paragraph {index} was polluted with MTDisplayEquation style.");
                AssertEqual(0, range.OMaths.Count,
                    $"{context}: ordinary prose paragraph {index} retained an OMML source.");

                shapes = range.InlineShapes;
                AssertEqual(1, shapes.Count,
                    $"{context}: ordinary prose paragraph {index} did not retain exactly one inline OLE formula.");
                shape = shapes[1];
                if (string.Equals(
                        objectMode,
                        FormulaOleContract.NativeOleMode,
                        StringComparison.Ordinal))
                {
                    AssertTrue(
                        WordFormulaMetadataReader.IsNativeOle(shape),
                        $"{context}: prose paragraph {index} does not contain a VisualTeX OLE.");
                    var metadata = WordFormulaMetadataReader.TryRead(shape)
                        ?? throw new InvalidDataException(
                            $"{context}: VisualTeX inline OLE in paragraph {index} has no metadata.");
                    AssertEqual("inline", metadata.DisplayMode,
                        $"{context}: VisualTeX prose formula {index} changed display mode.");
                    AssertTrue(!metadata.Numbered,
                        $"{context}: VisualTeX prose formula {index} became numbered.");
                }
                else if (string.Equals(
                             objectMode,
                             FormulaOleContract.MathTypeOleMode,
                             StringComparison.Ordinal))
                {
                    AssertTrue(
                        MathTypeOleInterop.IsMathTypeOle(shape),
                        $"{context}: prose paragraph {index} does not contain a MathType OLE.");
                }
                else
                {
                    throw new InvalidDataException(
                        $"{context}: unsupported OLE mode '{objectMode}'.");
                }
            }
            finally
            {
                Release(shape);
                Release(shapes);
                Release(style);
                Release(range);
                Release(paragraph);
            }
        }
    }

    private static string CreateOmmlBulkStressSource()
    {
        var displays = new[]
        {
            @"\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}a_{nm}X_n(x)Y_m(y)+\sum_{p=0}^{P}\sum_{q=0}^{Q}\frac{b_{pq}}{1+x^2}",
            @"\int_{0}^{1}\frac{x^2}{1+x^4}\,\mathrm{d}x+\oint_{\Gamma}\mathbf{F}\cdot\mathrm{d}\mathbf{r}",
            @"\frac{\partial^2 u}{\partial x^2}+\frac{\partial^2 u}{\partial y^2}=f(x,y)",
            @"\langle f|L|g\rangle-\langle g|L^{\dagger}|f\rangle=Q[f^{\ast},g]\Big|_{a}^{b}",
            @"\begin{aligned}a_1x+b_1y&=c_1\\a_2x+b_2y&=c_2\\x-y&=\lambda\end{aligned}",
            @"A_{i_{1}i_{2}}^{j^{2}k_{\ell}}=\frac{B_{i_1}^{j_1}+C_{i_2}^{j_2}}{D_{\ell}^{\,3}}",
            @"\lim_{x\to0}\frac{\sin x}{x}=1,\qquad\lim_{n\to\infty}\left(1+\frac1n\right)^n=\mathrm e",
            @"\sqrt{\frac{1+\sqrt{1-x^2}}{2}}+\sqrt[3]{\alpha^3+\beta^3}",
            @"\prod_{i=1}^{N}\left(1+\frac{\lambda_i}{i^2}\right)=\sum_{k=0}^{N}\binom{N}{k}p^k(1-p)^{N-k}",
            @"\left.\frac{\partial v(x,t)}{\partial t}\right|_{t=0}=-\lambda v(x,0),\quad 0\le x\le L",
        };

        var source = new StringBuilder();
        for (var index = 1; index <= 5; index++)
        {
            source.AppendLine(
                $"OMML-PROSE-{index:00}-BEFORE 普通正文，行内公式 "
                + $"$x_{{{index}}}^{{2}}+y_{{{index},j}}=z^{{\\ast}}$"
                + " 之后正文必须保持同一段和原顺序。");
            source.AppendLine();
        }
        for (var index = 0; index < displays.Length; index++)
        {
            source.AppendLine("\\begin{equation*}");
            source.AppendLine(displays[index]);
            source.AppendLine("\\end{equation*}");
            source.AppendLine();
        }
        for (var index = 6; index <= 10; index++)
        {
            source.AppendLine(
                $"OMML-PROSE-{index:00}-AFTER 普通正文，行内公式 "
                + $"$\\alpha_{{{index}}}+\\beta^{{{index}}}=\\gamma_{{i_{{{index}}}}}$"
                + " 之后正文不得被公式或样式覆盖。");
            source.AppendLine();
        }
        return source.ToString();
    }

    private static void AssertOmmlBulkStressState(
        Word.Document document,
        WordFormulaService service,
        string context)
    {
        AssertEqual(20, document.OMaths.Count,
            $"{context}: native OMath count changed.");
        AssertEqual(0, document.InlineShapes.Count,
            $"{context}: direct OMML import unexpectedly created InlineShape/OLE objects.");
        AssertEqual(20, WordOmmlFormulaStore.StoredFormulaIds(document).Count,
            $"{context}: stale or missing OMML CustomXML metadata parts were detected.");
        AssertEqual(20, WordOmmlFormulaStore.BookmarkedFormulaIds(document).Count,
            $"{context}: a managed OMML formula lost its VTOMML bookmark.");
        AssertEqual(20, WordOmmlFormulaStore.FormulaIds(document).Count,
            $"{context}: managed OMML identity reconciliation did not yield twenty formulas.");

        var inlineCount = 0;
        var displayCount = 0;
        var numberedCount = 0;
        var verifiedFingerprints = 0;
        var doubleSumNaryCount = -1;
        var totalNaryCount = 0;
        var emptyNaryOperands = 0;
        var eqArrCount = 0;
        var matrixCount = 0;
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";

        foreach (var formulaId in WordOmmlFormulaStore.FormulaIds(document))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? range = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                    ?? throw new InvalidDataException(
                        $"{context}: formula {formulaId} has no readable OMML metadata.");
                if (string.Equals(metadata.DisplayMode, "inline", StringComparison.Ordinal))
                    inlineCount++;
                else if (string.Equals(metadata.DisplayMode, "block", StringComparison.Ordinal))
                    displayCount++;
                if (metadata.Numbered) numberedCount++;

                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                    ?? throw new InvalidDataException(
                        $"{context}: formula {formulaId} has no VTOMML bookmark.");
                range = WordOmmlFormulaStore.GetEquationRange(bookmark);
                var xml = range.WordOpenXML;
                var liveFingerprint = WordOmmlConverter.ComputeOmmlFingerprint(xml);
                AssertEqual(
                    metadata.NativeOmmlFingerprint ?? string.Empty,
                    liveFingerprint,
                    $"{context}: formula {formulaId} has a stale OMML fingerprint.");
                verifiedFingerprints++;

                var tree = XDocument.Parse(xml);
                var naries = tree.Descendants(math + "nary").ToArray();
                totalNaryCount += naries.Length;
                emptyNaryOperands += naries.Count(nary =>
                {
                    var operand = nary.Element(math + "e");
                    return operand is null
                        || !operand
                            .Descendants(math + "t")
                            .Any(text => !string.IsNullOrWhiteSpace(text.Value));
                });
                eqArrCount += tree.Descendants(math + "eqArr").Count();
                matrixCount += tree.Descendants(math + "m").Count();
                if ((metadata.Latex ?? string.Empty).IndexOf(
                        "a_{nm}",
                        StringComparison.Ordinal) >= 0)
                    doubleSumNaryCount = naries.Length;
            }
            finally
            {
                Release(range);
                Release(bookmark);
            }
        }

        AssertEqual(10, inlineCount,
            $"{context}: inline OMML count changed.");
        AssertEqual(10, displayCount,
            $"{context}: display OMML count changed.");
        AssertEqual(10, numberedCount,
            $"{context}: numbered OMML count changed.");
        AssertEqual(20, verifiedFingerprints,
            $"{context}: not every formula retained a live fingerprint.");
        AssertEqual(4, doubleSumNaryCount,
            $"{context}: nested double summation did not retain four m:nary nodes.");
        AssertEqual(0, emptyNaryOperands,
            $"{context}: at least one native n-ary operator has an empty m:e operand.");
        AssertEqual(8, totalNaryCount,
            $"{context}: one or more expected sum, integral, contour-integral or product m:nary nodes changed.");
        AssertTrue(eqArrCount + matrixCount > 0,
            $"{context}: the aligned-system formula lost both m:eqArr and m:m structure.");

        var conversionPlan = service.CaptureFormulaFormatConversionPlan(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.MathTypeOleMode);
        AssertEqual(20, conversionPlan.Targets.Count,
            $"{context}: OMML conversion capture did not discover every formula.");
        var doubleSumTarget = conversionPlan.Targets.Single(target =>
            target.Latex.IndexOf("a_{nm}", StringComparison.Ordinal) >= 0
            || target.Latex.IndexOf("a_{n m}", StringComparison.Ordinal) >= 0);
        var roundTripDoubleSumCount = XDocument
            .Parse(doubleSumTarget.SourceMathMl
                ?? throw new InvalidDataException(
                    $"{context}: double summation target has no reverse MathML."))
            .Descendants()
            .Count(element =>
                element.Name.LocalName == "mo"
                && element.Value == "∑");
        AssertEqual(4, roundTripDoubleSumCount,
            $"{context}: OMML→MathML reverse conversion changed double summation semantics.");
    }

    private static void AssertOmmlStressProseOrderAndStyles(
        Word.Document document)
    {
        var text = document.Content.Text ?? string.Empty;
        var cursor = -1;
        for (var index = 1; index <= 10; index++)
        {
            var marker = index <= 5
                ? $"OMML-PROSE-{index:00}-BEFORE"
                : $"OMML-PROSE-{index:00}-AFTER";
            var position = text.IndexOf(marker, cursor + 1, StringComparison.Ordinal);
            AssertTrue(position > cursor,
                $"OMML stress prose marker {marker} was lost or reordered.");
            cursor = position;
        }

        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.Style? style = null;
            Word.OMath? proseMath = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                var paragraphText = range.Text ?? string.Empty;
                if (paragraphText.IndexOf(
                        "OMML-PROSE-",
                        StringComparison.Ordinal) < 0)
                    continue;
                style = range.get_Style() as Word.Style;
                var styleName = style?.NameLocal ?? string.Empty;
                AssertTrue(
                    styleName.IndexOf(
                        "MTDisplayEquation",
                        StringComparison.OrdinalIgnoreCase) < 0,
                    $"Ordinary prose paragraph {index} was polluted with MTDisplayEquation style.");
                AssertEqual(1, range.OMaths.Count,
                    $"Ordinary prose paragraph {index} did not retain exactly one inline OMath.");
                proseMath = range.OMaths[1];
                AssertEqual(
                    Word.WdOMathType.wdOMathInline,
                    proseMath.Type,
                    $"Ordinary prose paragraph {index} absorbed a display OMath.");
            }
            finally
            {
                Release(proseMath);
                Release(style);
                Release(range);
                Release(paragraph);
            }
        }
    }

    private static int CountStructuralParagraphsBetweenTables(
        Word.Document document)
    {
        const string wordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var tree = XDocument.Parse(content.WordOpenXML);
            XNamespace word = wordNamespace;
            XNamespace math = mathNamespace;
            var body = tree.Descendants(word + "body").SingleOrDefault()
                ?? throw new InvalidDataException(
                    "The OMML stress document has no WordOpenXML body.");
            var children = body
                .Elements()
                .Where(element =>
                    element.Name == word + "p"
                    || element.Name == word + "tbl")
                .ToArray();

            bool IsBlankBodyParagraph(XElement element)
            {
                if (element.Name != word + "p") return false;
                return !element.Descendants().Any(descendant =>
                    (descendant.Name == word + "t"
                     && !string.IsNullOrWhiteSpace(descendant.Value))
                    || descendant.Name == word + "tab"
                    || descendant.Name == word + "br"
                    || descendant.Name == word + "drawing"
                    || descendant.Name == word + "object"
                    || descendant.Name == word + "pict"
                    || descendant.Name == word + "fldSimple"
                    || descendant.Name == math + "oMath"
                    || descendant.Name == math + "oMathPara");
            }

            var count = 0;
            for (var index = 0; index < children.Length; index++)
            {
                if (!IsBlankBodyParagraph(children[index])) continue;
                var previous = index - 1;
                while (previous >= 0
                       && IsBlankBodyParagraph(children[previous]))
                    previous--;
                var next = index + 1;
                while (next < children.Length
                       && IsBlankBodyParagraph(children[next]))
                    next++;
                if (previous >= 0
                    && next < children.Length
                    && children[previous].Name == word + "tbl"
                    && children[next].Name == word + "tbl")
                    count++;
            }
            return count;
        }
        finally
        {
            Release(content);
        }
    }
}

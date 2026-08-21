using System.Diagnostics;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunUserOmmlMathTypeConversionAcceptance(string artifactRoot)
    {
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER_OMML_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "VISUALTEX_USER_OMML_SOURCE must point to the 100-formula OMML document.",
                sourcePath);

        Directory.CreateDirectory(artifactRoot);
        var outputPath = Path.Combine(artifactRoot, "user-100-omml-to-mathtype.docx");
        var tracePath = Path.Combine(artifactRoot, "user-100-omml-to-mathtype.trace.log");
        File.Copy(Path.GetFullPath(sourcePath), outputPath, overwrite: true);
        try { File.Delete(tracePath); } catch { }

        var previousFormatAcceptance = Environment.GetEnvironmentVariable(
            "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
        var previousTracePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_WORD_HOOK_TRACE_PATH");
        var previousAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        Word.Application? application = null;
        Word.Document? document = null;
        Microsoft.Office.Core.COMAddIns? addIns = null;
        Microsoft.Office.Core.COMAddIn? installedAddIn = null;
        object? callbacksObject = null;
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", null);
            var mathTypeBaseline = SnapshotMathTypeProcessIds();
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            callbacksObject = GetInstalledStressCallbacks(
                application,
                out addIns,
                out installedAddIn);
            dynamic callbacks = callbacksObject;

            var sourceService = new WordFormulaService(application);
            var plan = sourceService.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(100, plan.Targets.Count,
                "OMML→MathType source did not contain exactly 100 formulas.");
            var expected = plan.Targets
                .OrderBy(target => target.SourceStart)
                .Select(target => MathTypeMtefCodec.SemanticSignature(
                    target.SourceMathMl
                    ?? throw new InvalidDataException("OMML target has no source MathML.")))
                .ToArray();

            var watch = Stopwatch.StartNew();
            callbacks.OnConvertOmmlToMathTypeDocument(new object());
            var peakMathTypeProcessCount = WaitForInstalledOmmlMathTypeConversion(
                tracePath,
                "source=OMML target=MathType",
                mathTypeBaseline,
                allowTransientMathTypeProcess: true);
            watch.Stop();

            ValidateUserOmmlMathTypeResult(document, expected, "live document");
            AssertNoNewMathTypeProcess(mathTypeBaseline, "user-100 OMML→MathType");
            document.Save();
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = application.Documents.Open(
                outputPath,
                ReadOnly: true,
                AddToRecentFiles: false,
                Visible: false);
            ValidateUserOmmlMathTypeResult(document, expected, "saved/reopened document");
            Console.WriteLine(
                $"[USER100 OMML→MT PASS] 100/100 converted in {watch.Elapsed.TotalSeconds:0.00}s; "
                + $"peakTransientMathTypeProcessCount={peakMathTypeProcessCount}; finalMathTypeProcessCount=0; output={outputPath}");
        }
        finally
        {
            Release(callbacksObject);
            Release(installedAddIn);
            Release(addIns);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
            Environment.SetEnvironmentVariable(
                "VISUALTEX_WORD_HOOK_TRACE_PATH",
                previousTracePath);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE",
                previousFormatAcceptance);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", previousAcceptance);
        }
    }

    private static void ValidateUserOmmlMathTypeResult(
        Word.Document document,
        IReadOnlyList<string> expected,
        string phase)
    {
        AssertEqual(0, document.OMaths.Count,
            $"OMML→MathType left OMML source formulas behind in the {phase}.");
        AssertEqual(100, CountMathTypeOleShapes(document),
            $"OMML→MathType did not create exactly 100 MathType formulas in the {phase}.");
        var seen = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (!MathTypeOleInterop.IsMathTypeOle(shape)) continue;
                var actual = MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(shape));
                AssertEqual(expected[seen], actual,
                    $"OMML→MathType formula #{seen + 1} changed semantics in the {phase}.");
                seen++;
            }
            finally { Release(shape); }
        }
        AssertEqual(100, seen,
            $"OMML→MathType did not validate 100 targets in the {phase}.");
    }

    private static int CountStructurallyBlankParagraphs(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.InlineShapes? shapes = null;
            Word.OMaths? maths = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                shapes = range.InlineShapes;
                maths = range.OMaths;
                var text = (range.Text ?? string.Empty)
                    .Trim('\r', '\a', '\v', '\f', '\t', ' ');
                if (text.Length == 0 && shapes.Count == 0 && maths.Count == 0)
                    count++;
            }
            finally
            {
                Release(maths);
                Release(shapes);
                Release(range);
                Release(paragraph);
            }
        }
        return count;
    }

    private static void RunUserOmmlMathTypeSourceAudit()
    {
        var sourcePath = Environment.GetEnvironmentVariable("VISUALTEX_USER_OMML_SOURCE");
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "VISUALTEX_USER_OMML_SOURCE must point to the user's OMML Word document.",
                sourcePath);

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            var writablePath = Path.Combine(
                Path.GetTempPath(),
                $"VisualTeX-user-omml-audit-{Guid.NewGuid():N}.docx");
            File.Copy(Path.GetFullPath(sourcePath), writablePath, overwrite: true);
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                writablePath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var service = new WordFormulaService(application);
            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            Console.WriteLine(
                $"[USER OMML AUDIT] targets={plan.Targets.Count}; omaths={document.OMaths.Count}; "
                + $"paragraphs={document.Paragraphs.Count}; inlineShapes={document.InlineShapes.Count}");

            var mismatches = 0;
            foreach (var pair in plan.Targets
                         .OrderBy(target => target.SourceStart)
                         .Select((target, index) => (Target: target, Index: index + 1)))
            {
                var target = pair.Target;
                var sourceMathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"OMML target #{pair.Index} has no source MathML.");
                var generated = MathTypeMtefCodec.CreateEquationNative(
                    sourceMathMl,
                    inline: string.Equals(
                        target.DisplayMode,
                        "inline",
                        StringComparison.OrdinalIgnoreCase));
                var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
                var generatedMathMl = MathTypeOleStorage.ReadMathMl(compound);
                var expected = MathTypeMtefCodec.SemanticSignature(sourceMathMl);
                var actual = MathTypeMtefCodec.SemanticSignature(generatedMathMl);
                if (string.Equals(expected, actual, StringComparison.Ordinal))
                    continue;

                mismatches++;
                Console.WriteLine(
                    $"MISMATCH|{pair.Index}|start={target.SourceStart}|display={target.DisplayMode}"
                    + $"|latex={target.Latex}|expected={expected}|actual={actual}"
                    + $"|sourceMathMl={sourceMathMl}|generatedMathMl={generatedMathMl}");
            }

            Console.WriteLine($"[USER OMML AUDIT DONE] mismatches={mismatches}/{plan.Targets.Count}");
            if (mismatches != 0)
                throw new InvalidDataException(
                    $"User OMML audit found {mismatches} standalone MTEF semantic mismatch(es).");
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }
}

using System.Diagnostics;
using System.Text;
using Extensibility;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const int BulkMathTypeHundredFormulaCount = 100;
    private const long BulkMathTypeHundredLimitMilliseconds = 30_000;

    private static void RunWordBulkImportMathTypeHundredPerformance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var pairCountText = Environment.GetEnvironmentVariable("VISUALTEX_ACCEPTANCE_MATHTYPE_PAIRS");
        var pairCount = string.IsNullOrWhiteSpace(pairCountText)
            ? 50
            : int.TryParse(pairCountText, out var parsedPairCount)
                && parsedPairCount is >= 1 and <= 50
                    ? parsedPairCount
                    : throw new InvalidDataException(
                        "VISUALTEX_ACCEPTANCE_MATHTYPE_PAIRS must be an integer from 1 to 50.");
        var formulaCount = pairCount * 2;
        var sourcePath = Path.Combine(artifactRoot, "word-bulk-import-mathtype-100.tex");
        var logPath = Path.Combine(artifactRoot, "word-bulk-import-mathtype-100.log");
        var outputPath = Path.Combine(artifactRoot, "word-bulk-import-mathtype-100.docx");
        var reportPath = Path.Combine(artifactRoot, "word-bulk-import-mathtype-100-performance.json");
        var source = CreateWordBulkImportMathTypeHundredSource(pairCount);
        File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
        DeleteBulkPerformanceArtifact(logPath);
        DeleteBulkPerformanceArtifact(outputPath);
        DeleteBulkPerformanceArtifact(reportPath);

        var parsed = WordBulkImportParser.Parse(
            source,
            WordBulkSourceFormat.Latex,
            WordBulkFormulaObjectMode.MathType);
        AssertEqual(formulaCount, parsed.FormulaCount,
            "MathType performance source formula count changed.");
        AssertEqual(pairCount, parsed.InlineFormulaCount,
            "MathType performance source inline count changed.");
        AssertEqual(pairCount, parsed.DisplayFormulaCount,
            "MathType performance source display count changed.");
        AssertEqual(0, parsed.Warnings.Count,
            "MathType 100 performance source unexpectedly produced parser warnings.");
        var expectedLatex = parsed.Blocks
            .SelectMany(block => block.Runs)
            .Where(run => run.IsFormula)
            .Select(run => run.Latex)
            .ToArray();

        var previousSource = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE");
        var previousSourcePath = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH");
        var previousFormat = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT");
        var previousMode = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE");
        var previousNumber = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS");
        var previousLog = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE", null);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_SOURCE_PATH", sourcePath);
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_FORMAT", "latex");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_OBJECT_MODE", "mathtype");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_NUMBER_DISPLAY_FORMULAS", "0");
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG", logPath);

            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            addIn = new ThisAddIn();
            addIn.OnConnection(
                application,
                ext_ConnectMode.ext_cm_AfterStartup,
                addIn,
                ref custom);

            var wall = Stopwatch.StartNew();
            addIn.OnBulkImport(new object());
            var operationMilliseconds = WaitForBulkImportCompletion(
                logPath,
                TimeSpan.FromMinutes(5));
            WaitForAddInIdle(addIn, TimeSpan.FromSeconds(60));
            wall.Stop();

            AssertEqual(0, document.OMaths.Count,
                "MathType 100 performance import unexpectedly created native OMML.");
            AssertEqual(formulaCount, CountMathTypeOleShapes(document),
                "MathType performance import produced the wrong MathType OLE count.");
            AssertEqual(formulaCount, document.InlineShapes.Count,
                "MathType performance import produced an unexpected non-MathType inline shape.");

            var service = new WordFormulaService(application);
            var metadata = ReadBulkMathTypeMetadata(
                service,
                document,
                "MathType 100 performance import");
            AssertEqual(formulaCount, metadata.Count,
                "MathType performance import did not leave the expected editable formulas.");
            AssertEqual(formulaCount,
                metadata.Select(item => item.FormulaId)
                    .Distinct(StringComparer.OrdinalIgnoreCase).Count(),
                "MathType performance import reused a formula ID.");
            for (var index = 0; index < expectedLatex.Length; index++)
            {
                AssertEqual(
                    NormalizeMathTypePerformanceLatex(expectedLatex[index]),
                    NormalizeMathTypePerformanceLatex(metadata[index].Latex),
                    $"MathType performance formula {index + 1} source/order changed.");
            }

            document.SaveAs2(
                outputPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            reopened = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(0, reopened.OMaths.Count,
                "MathType 100 save/reopen unexpectedly created native OMML.");
            AssertEqual(formulaCount, CountMathTypeOleShapes(reopened),
                "MathType performance save/reopen lost a MathType OLE.");
            AssertEqual(formulaCount, reopened.InlineShapes.Count,
                "MathType performance save/reopen changed inline-shape count.");

            var reopenedService = new WordFormulaService(application);
            var reopenedMetadata = ReadBulkMathTypeMetadata(
                reopenedService,
                reopened,
                "MathType 100 save/reopen");
            AssertEqual(formulaCount, reopenedMetadata.Count,
                "MathType performance save/reopen lost editable metadata.");

            var report = new
            {
                Time = DateTimeOffset.Now,
                CandidateWordAssembly = typeof(ThisAddIn).Assembly.Location,
                FormulaCount = formulaCount,
                InlineCount = pairCount,
                DisplayCount = pairCount,
                OperationMilliseconds = operationMilliseconds,
                WallMilliseconds = wall.ElapsedMilliseconds,
                LimitMilliseconds = BulkMathTypeHundredLimitMilliseconds,
                Passed = operationMilliseconds <= BulkMathTypeHundredLimitMilliseconds,
                Output = outputPath,
            };
            File.WriteAllText(
                reportPath,
                System.Text.Json.JsonSerializer.Serialize(
                    report,
                    new System.Text.Json.JsonSerializerOptions { WriteIndented = true }),
                new UTF8Encoding(false));

            AssertTrue(
                operationMilliseconds <= BulkMathTypeHundredLimitMilliseconds,
                $"MathType 100 bulk import took {operationMilliseconds} ms; "
                + $"limit is {BulkMathTypeHundredLimitMilliseconds} ms.");

            Console.WriteLine(
                $"[MATHTYPE PERFORMANCE PASS] formulas={formulaCount}, "
                + $"operation={operationMilliseconds}ms, wall={wall.ElapsedMilliseconds}ms, "
                + $"limit={BulkMathTypeHundredLimitMilliseconds}ms.");
            Console.WriteLine($"Artifact: {outputPath}");
            Console.WriteLine($"Report: {reportPath}");
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
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static string CreateWordBulkImportMathTypeHundredSource(int pairCount)
    {
        var source = new StringBuilder();
        for (var index = 1; index <= pairCount; index++)
        {
            source.AppendLine(
                $"PERF-{index:00}-BEFORE ordinary prose before inline formula "
                + $"$x_{{{index}}}^2+y_{{{index}}}^2=z_{{{index}}}^2$ "
                + "and ordinary prose after it.");
            source.AppendLine();
            source.AppendLine("\\[");
            source.AppendLine(
                $"\\frac{{\\partial^2 u_{{{index}}}}}{{\\partial x^2}}"
                + $"+\\lambda_{{{index}}}u_{{{index}}}=f_{{{index}}}(x)");
            source.AppendLine("\\]");
            source.AppendLine(
                $"PERF-{index:00}-AFTER prose immediately after the display formula.");
            source.AppendLine();
        }
        return source.ToString();
    }

    private static string NormalizeMathTypePerformanceLatex(string? value)
    {
        var normalized = NormalizeStressLatex(value ?? string.Empty);
        return System.Text.RegularExpressions.Regex.Replace(
            normalized,
            @"([_^])\{([A-Za-z0-9])\}",
            "$1$2");
    }
}

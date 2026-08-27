using System.Collections;
using System.Diagnostics;
using System.Reflection;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeInsertionComplexityAcceptance(string artifactRoot)
    {
        if (!MathTypeOleInterop.TryResolveCapabilities(
                MathTypeOleInterop.CanonicalProgId,
                out var capabilities))
        {
            Console.WriteLine(
                "[MATHTYPE INSERT COMPLEXITY] Skipped: no registered MathType OLE server is available.");
            return;
        }

        Directory.CreateDirectory(artifactRoot);
        var svgPath = Path.Combine(artifactRoot, "mathtype-insert-scaling.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-family=\"Cambria Math\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 260, 96);
        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";

        var mathTypeProcessCountBefore = CountMathTypeInsertionComplexityProcesses();
        var staticCollectionCountsBefore = SnapshotMathTypeInsertionStaticCollections();
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType insertion complexity acceptance\r";
            var service = new WordFormulaService(application);
            var serviceCollectionCountsBefore =
                SnapshotMathTypeInsertionInstanceCollections(service);

            var inlineTimes = InsertMathTypeScalingSeries(
                application,
                document,
                service,
                emfPath,
                mathMl,
                display: false,
                numbered: false,
                count: 24);
            AssertBoundedMathTypeInsertionGrowth("inline", inlineTimes);

            var numberedTimes = InsertMathTypeScalingSeries(
                application,
                document,
                service,
                emfPath,
                mathMl,
                display: true,
                numbered: true,
                count: 24);
            AssertBoundedMathTypeInsertionGrowth("numbered display", numberedTimes);

            AssertEqual(48, CountMathTypeOleShapes(document),
                "MathType insertion complexity acceptance created the wrong number of MathType OLE objects.");
            var mathTypeProcessCountDuring = CountMathTypeInsertionComplexityProcesses();
            AssertTrue(
                mathTypeProcessCountDuring <= mathTypeProcessCountBefore + 1,
                $"MathType insertion leaked one process per formula: before={mathTypeProcessCountBefore}, during={mathTypeProcessCountDuring}.");
            AssertBoundedMathTypeInsertionStaticState(
                staticCollectionCountsBefore,
                SnapshotMathTypeInsertionStaticCollections());
            AssertBoundedMathTypeInsertionInstanceState(
                serviceCollectionCountsBefore,
                SnapshotMathTypeInsertionInstanceCollections(service));

            var outputPath = Path.Combine(
                artifactRoot,
                "MathType-Insertion-Complexity.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            document = application.Documents.Open(
                outputPath,
                ReadOnly: false,
                AddToRecentFiles: false);
            AssertEqual(48, CountMathTypeOleShapes(document),
                "Save/reopen changed the MathType object inventory after the insertion scaling run.");
            AssertMathTypeInsertionComplexityFormula(
                application,
                document,
                1,
                expectedNumbered: false,
                expectedDisplay: false);
            AssertMathTypeInsertionComplexityFormula(
                application,
                document,
                24,
                expectedNumbered: false,
                expectedDisplay: false);
            AssertMathTypeInsertionComplexityFormula(
                application,
                document,
                25,
                expectedNumbered: true,
                expectedDisplay: true);
            AssertMathTypeInsertionComplexityFormula(
                application,
                document,
                48,
                expectedNumbered: true,
                expectedDisplay: true);
            Console.WriteLine(
                $"[MATHTYPE INSERT COMPLEXITY] Passed with server={capabilities.ServerPath ?? capabilities.ResolvedClsid.ToString("D")}; "
                + $"inline first/last median={MathTypeInsertionComplexityMedian(inlineTimes.Take(6))}/{MathTypeInsertionComplexityMedian(inlineTimes.Skip(Math.Max(0, inlineTimes.Count - 6)))}ms, "
                + $"numbered first/last median={MathTypeInsertionComplexityMedian(numberedTimes.Take(6))}/{MathTypeInsertionComplexityMedian(numberedTimes.Skip(Math.Max(0, numberedTimes.Count - 6)))}ms. "
                + $"path={outputPath}");

            // Release Word and all RCWs before checking the external MathType
            // process inventory. This catches a per-insertion server/process leak
            // without terminating any MathType process the user already had open.
            document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(document);
            document = null;
            QuitWordApplicationIfOwned(application);
            Release(application);
            application = null;
            ForceComCleanup();
            var mathTypeProcessCountAfter = WaitForMathTypeInsertionProcessCount(
                mathTypeProcessCountBefore,
                TimeSpan.FromSeconds(12));
            AssertTrue(
                mathTypeProcessCountAfter <= mathTypeProcessCountBefore,
                $"MathType insertion left a background process after Word cleanup: before={mathTypeProcessCountBefore}, after={mathTypeProcessCountAfter}.");
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

    private static IReadOnlyList<long> InsertMathTypeScalingSeries(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string emfPath,
        string mathMl,
        bool display,
        bool numbered,
        int count)
    {
        var results = new List<long>(count);
        for (var index = 0; index < count; index++)
        {
            Word.Range? insertion = null;
            try
            {
                insertion = document.Range(
                    document.Content.End - 1,
                    document.Content.End - 1);
                insertion.Select();
            }
            finally { Release(insertion); }

            var latex = $"x_{{{index + 1}}}+1";
            var session = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "create",
                Host = "word",
                FormulaId = Guid.NewGuid().ToString("D"),
                Title = "MathType insertion complexity acceptance",
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
                },
                CodeFormat = "latex",
                DisplayMode = display ? "block" : "inline",
                ObjectMode = FormulaOleContract.MathTypeOleMode,
                Numbered = numbered,
                FontSizePt = 12,
                ExportResult = new OfficeExportDocument
                {
                    MathMl = mathMl,
                    Width = 260,
                    Height = 96,
                    Baseline = 72,
                },
            };

            var watch = Stopwatch.StartNew();
            service.InsertMathTypeOle(session, mathMl, emfPath);
            watch.Stop();
            results.Add(watch.ElapsedMilliseconds);
            System.Windows.Forms.Application.DoEvents();
        }
        return results;
    }

    private static void AssertBoundedMathTypeInsertionGrowth(
        string seriesName,
        IReadOnlyList<long> elapsedMilliseconds)
    {
        AssertTrue(elapsedMilliseconds.Count >= 12,
            $"MathType {seriesName} scaling series is too short.");
        var firstMedian = MathTypeInsertionComplexityMedian(elapsedMilliseconds.Take(6));
        var lastMedian = MathTypeInsertionComplexityMedian(
            elapsedMilliseconds.Skip(Math.Max(0, elapsedMilliseconds.Count - 6)));
        var allowedLastMedian = Math.Max(firstMedian * 2.5, firstMedian + 650);
        AssertTrue(lastMedian <= allowedLastMedian,
            $"MathType {seriesName} insertion time grows with document size: first median={firstMedian}ms, "
            + $"last median={lastMedian}ms, allowed={allowedLastMedian:F0}ms, "
            + $"samples=[{string.Join(",", elapsedMilliseconds)}].");

        var slope = MathTypeInsertionComplexitySlope(elapsedMilliseconds);
        AssertTrue(slope <= 25,
            $"MathType {seriesName} insertion has an excessive positive slope of {slope:F2}ms/formula: "
            + $"samples=[{string.Join(",", elapsedMilliseconds)}].");
    }

    private static void AssertMathTypeInsertionComplexityFormula(
        Word.Application application,
        Word.Document document,
        int index,
        bool expectedNumbered,
        bool expectedDisplay)
    {
        Word.InlineShape? shape = null;
        try
        {
            shape = document.InlineShapes[index];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                $"Reopened formula {index} is not recognized as MathType OLE.");
            var metadata = MathTypeOleInterop.ReadMetadata(application, shape);
            AssertEqual(expectedNumbered, metadata.Numbered,
                $"Reopened MathType formula {index} has the wrong Numbered state.");
            AssertEqual(expectedDisplay ? "block" : "inline", metadata.DisplayMode,
                $"Reopened MathType formula {index} has the wrong display mode.");
            var mathMl = MathTypeOleStorage.ReadMathMl(shape);
            AssertTrue(
                !string.IsNullOrWhiteSpace(mathMl)
                && mathMl.IndexOf("<math", StringComparison.OrdinalIgnoreCase) >= 0,
                $"Reopened MathType formula {index} has no recoverable MathML.");
        }
        finally { Release(shape); }
    }

    private static IReadOnlyDictionary<string, int>
        SnapshotMathTypeInsertionInstanceCollections(object instance)
    {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var type = instance.GetType(); type is not null; type = type.BaseType)
        {
            foreach (var field in type.GetFields(
                         BindingFlags.Instance
                         | BindingFlags.Public
                         | BindingFlags.NonPublic
                         | BindingFlags.DeclaredOnly))
            {
                object? value;
                try { value = field.GetValue(instance); }
                catch { continue; }
                int? count = value switch
                {
                    IDictionary dictionary => dictionary.Count,
                    ICollection collection => collection.Count,
                    _ => null,
                };
                if (count.HasValue)
                    result[type.FullName + "." + field.Name] = count.Value;
            }
        }
        return result;
    }

    private static void AssertBoundedMathTypeInsertionInstanceState(
        IReadOnlyDictionary<string, int> before,
        IReadOnlyDictionary<string, int> after)
    {
        foreach (var entry in after)
        {
            before.TryGetValue(entry.Key, out var originalCount);
            AssertTrue(entry.Value <= originalCount + 16,
                $"MathType insertion leaked per-service collection state in {entry.Key}: "
                + $"before={originalCount}, after={entry.Value}.");
        }
    }

    private static IReadOnlyDictionary<string, int> SnapshotMathTypeInsertionStaticCollections()
    {
        var result = new Dictionary<string, int>(StringComparer.Ordinal);
        var types = new[]
        {
            typeof(MathTypeOleInterop),
            typeof(MathTypeOleClipboardProxy),
            typeof(MathTypeWordCommandsBridge),
            typeof(MathTypeNativePreviewRenderer),
            typeof(WordFormulaService),
        };
        foreach (var type in types)
        {
            foreach (var field in type.GetFields(
                         BindingFlags.Static
                         | BindingFlags.Public
                         | BindingFlags.NonPublic))
            {
                object? value;
                try { value = field.GetValue(null); }
                catch { continue; }
                int? count = value switch
                {
                    IDictionary dictionary => dictionary.Count,
                    ICollection collection => collection.Count,
                    _ => null,
                };
                if (count.HasValue)
                    result[type.FullName + "." + field.Name] = count.Value;
            }
        }
        return result;
    }

    private static void AssertBoundedMathTypeInsertionStaticState(
        IReadOnlyDictionary<string, int> before,
        IReadOnlyDictionary<string, int> after)
    {
        foreach (var entry in after)
        {
            before.TryGetValue(entry.Key, out var originalCount);
            var allowedGrowth = entry.Key.EndsWith(".CapabilityCache", StringComparison.Ordinal)
                ? 32
                : 16;
            AssertTrue(entry.Value <= originalCount + allowedGrowth,
                $"MathType insertion leaked static collection state in {entry.Key}: "
                + $"before={originalCount}, after={entry.Value}, allowedGrowth={allowedGrowth}.");
            if (entry.Key.EndsWith(".CapabilityCache", StringComparison.Ordinal))
                AssertTrue(entry.Value <= 32,
                    $"MathType capability cache exceeded its hard bound: {entry.Value} entries.");
        }
    }

    private static int WaitForMathTypeInsertionProcessCount(
        int targetCount,
        TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        var current = CountMathTypeInsertionComplexityProcesses();
        while (current > targetCount && DateTime.UtcNow < deadline)
        {
            Thread.Sleep(150);
            current = CountMathTypeInsertionComplexityProcesses();
        }
        return current;
    }

    private static int CountMathTypeInsertionComplexityProcesses()
    {
        var processes = Process.GetProcessesByName("MathType");
        try { return processes.Length; }
        finally
        {
            foreach (var process in processes)
                process.Dispose();
        }
    }

    private static double MathTypeInsertionComplexityMedian(IEnumerable<long> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        if (ordered.Length == 0) return 0;
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2d
            : ordered[middle];
    }

    private static double MathTypeInsertionComplexitySlope(IReadOnlyList<long> values)
    {
        var count = values.Count;
        var meanX = (count - 1) / 2d;
        var meanY = values.Average(value => (double)value);
        double numerator = 0;
        double denominator = 0;
        for (var index = 0; index < count; index++)
        {
            var dx = index - meanX;
            numerator += dx * (values[index] - meanY);
            denominator += dx * dx;
        }
        return denominator <= 0 ? 0 : numerator / denominator;
    }
}

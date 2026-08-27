using System.Diagnostics;
using System.Globalization;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class MathTypeInsertScalingMetric
    {
        internal int Index { get; set; }
        internal double ElapsedMilliseconds { get; set; }
    }

    private static void RunWordMathTypeInsertScalingAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const int inlineCount = 60;
        const int numberedCount = 40;
        const int distantNumberedCount = 24;
        var svgPath = Path.Combine(artifactRoot, "mathtype-insert-scaling.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"160\" height=\"48\" viewBox=\"0 0 160 48\"><text x=\"4\" y=\"34\" font-family=\"Cambria Math\" font-size=\"28\">x₁ = 1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 160, 48);
        var wmfPath = Path.Combine(artifactRoot, "mathtype-insert-scaling.wmf");
        File.WriteAllBytes(
            wmfPath,
            MathTypeWordOpenXml.ConvertEnhancedMetafileToPlaceableWmf(
                emfPath,
                widthPt: 120,
                heightPt: 36));

        Word.Application? application = null;
        Word.Document? inlineDocument = null;
        Word.Document? numberedDocument = null;
        Word.Document? distantNumberedDocument = null;
        var baseline = SnapshotMathTypeProcessIds();
        try
        {
            if (baseline.Count > 0)
                Console.WriteLine(
                    $"[MATHTYPE INSERT SCALING] Preserving {baseline.Count} pre-existing MathType process(es); the acceptance requires that insertion create no additional process.");
            application = CreateWordApplication(visible: false);
            try { application.ScreenUpdating = false; } catch { }

            inlineDocument = application.Documents.Add();
            inlineDocument.Content.Text = "MathType inline insertion scaling\r";
            var inlineService = new WordFormulaService(application);
            var inlineMetrics = InsertMathTypeScalingSeries(
                application,
                inlineDocument,
                inlineService,
                inlineCount,
                displayMode: "inline",
                numbered: false,
                emfPath,
                wmfPath);
            AssertMathTypeInsertScaling("inline", inlineMetrics);
            AssertEqual(
                inlineCount,
                CountScalingMathTypeObjects(inlineDocument),
                "MathType inline scaling acceptance produced the wrong object count.");
            AssertNoNewMathTypeProcess(baseline, "MathType inline insertion scaling");

            numberedDocument = application.Documents.Add();
            numberedDocument.Content.Text = "MathType numbered insertion scaling\r";
            WordEquationNumbering.SetEquationNumberFormat(numberedDocument, "continuous");
            var numberedService = new WordFormulaService(application);
            var numberedMetrics = InsertMathTypeScalingSeries(
                application,
                numberedDocument,
                numberedService,
                numberedCount,
                displayMode: "block",
                numbered: true,
                emfPath,
                wmfPath);
            AssertMathTypeInsertScaling("numbered", numberedMetrics);
            AssertEqual(
                numberedCount,
                CountScalingMathTypeObjects(numberedDocument),
                "MathType numbered scaling acceptance produced the wrong object count.");
            AssertEqual(
                numberedCount,
                CountScalingMathTypeNumberFields(numberedDocument),
                "MathType numbered scaling acceptance produced the wrong MTPlaceRef count.");
            AssertNoNewMathTypeProcess(baseline, "MathType numbered insertion scaling");

            // Force the nearest-template lookup outside its 4K local window. This
            // catches a subtler O(N²) regression that adjacent formulas cannot see:
            // every new equation must not rescan all earlier Word fields merely
            // because ordinary prose separates the numbered formulas.
            distantNumberedDocument = application.Documents.Add();
            distantNumberedDocument.Content.Text =
                "MathType distant numbered insertion scaling\r";
            WordEquationNumbering.SetEquationNumberFormat(
                distantNumberedDocument,
                "continuous");
            var distantNumberedService = new WordFormulaService(application);
            var distantNumberedMetrics = InsertMathTypeScalingSeries(
                application,
                distantNumberedDocument,
                distantNumberedService,
                distantNumberedCount,
                displayMode: "block",
                numbered: true,
                emfPath,
                wmfPath,
                fillerCharactersBeforeEach: 5000);
            AssertMathTypeInsertScaling("distant-numbered", distantNumberedMetrics);
            AssertEqual(
                distantNumberedCount,
                CountScalingMathTypeObjects(distantNumberedDocument),
                "MathType distant-numbered scaling acceptance produced the wrong object count.");
            AssertEqual(
                distantNumberedCount,
                CountScalingMathTypeNumberFields(distantNumberedDocument),
                "MathType distant-numbered scaling acceptance produced the wrong MTPlaceRef count.");
            AssertNoNewMathTypeProcess(
                baseline,
                "MathType distant-numbered insertion scaling");

            WriteMathTypeInsertScalingCsv(
                Path.Combine(artifactRoot, "mathtype-insert-scaling-inline.csv"),
                inlineMetrics);
            WriteMathTypeInsertScalingCsv(
                Path.Combine(artifactRoot, "mathtype-insert-scaling-numbered.csv"),
                numberedMetrics);
            WriteMathTypeInsertScalingCsv(
                Path.Combine(
                    artifactRoot,
                    "mathtype-insert-scaling-distant-numbered.csv"),
                distantNumberedMetrics);

            var inlinePath = Path.Combine(
                artifactRoot,
                "MathType-Insert-Scaling-Inline.docx");
            var numberedPath = Path.Combine(
                artifactRoot,
                "MathType-Insert-Scaling-Numbered.docx");
            var distantNumberedPath = Path.Combine(
                artifactRoot,
                "MathType-Insert-Scaling-Distant-Numbered.docx");
            inlineDocument.SaveAs2(inlinePath, Word.WdSaveFormat.wdFormatXMLDocument);
            numberedDocument.SaveAs2(numberedPath, Word.WdSaveFormat.wdFormatXMLDocument);
            distantNumberedDocument.SaveAs2(
                distantNumberedPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            inlineDocument.Close(Word.WdSaveOptions.wdSaveChanges);
            numberedDocument.Close(Word.WdSaveOptions.wdSaveChanges);
            distantNumberedDocument.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(inlineDocument);
            Release(numberedDocument);
            Release(distantNumberedDocument);
            inlineDocument = application.Documents.Open(
                inlinePath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            numberedDocument = application.Documents.Open(
                numberedPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            distantNumberedDocument = application.Documents.Open(
                distantNumberedPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(
                inlineCount,
                CountScalingMathTypeObjects(inlineDocument),
                "Save/reopen lost an inline MathType scaling formula.");
            AssertEqual(
                numberedCount,
                CountScalingMathTypeObjects(numberedDocument),
                "Save/reopen lost a numbered MathType scaling formula.");
            AssertEqual(
                numberedCount,
                CountScalingMathTypeNumberFields(numberedDocument),
                "Save/reopen changed the MathType scaling number-field count.");
            AssertEqual(
                distantNumberedCount,
                CountScalingMathTypeObjects(distantNumberedDocument),
                "Save/reopen lost a distant-numbered MathType scaling formula.");
            AssertEqual(
                distantNumberedCount,
                CountScalingMathTypeNumberFields(distantNumberedDocument),
                "Save/reopen changed the distant-numbered MathType number-field count.");
            AssertNoNewMathTypeProcess(baseline, "MathType insertion scaling save/reopen");

            Console.WriteLine(
                "[MATHTYPE INSERT SCALING PASS] Interactive insertion remained effectively flat through "
                + $"{inlineCount} inline, {numberedCount} adjacent-numbered and "
                + $"{distantNumberedCount} prose-separated numbered Equation.DSMT4 formulas; "
                + "all objects and number fields survived save/reopen and MathTypeProcessCount=0.");
        }
        finally
        {
            if (inlineDocument is not null)
            {
                try { inlineDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            if (numberedDocument is not null)
            {
                try { numberedDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            if (distantNumberedDocument is not null)
            {
                try { distantNumberedDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(inlineDocument);
            Release(numberedDocument);
            Release(distantNumberedDocument);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordMathTypePreviewFallbackAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const string disableVariable = "VISUALTEX_DISABLE_MATHTYPE_NATIVE_PREVIEW";
        const string sourceLatex = @"x=\frac{-b+\sqrt{b^{2}-4ac}}{2a}";
        const string editedLatex = @"\frac{a+b}{c}=d";
        const string sourceMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>+</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
        const string editedMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow><mi>c</mi></mfrac><mo>=</mo><mi>d</mi></math>";

        var svgPath = Path.Combine(artifactRoot, "mathtype-preview-fallback.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"320\" height=\"96\" viewBox=\"0 0 320 96\"><text x=\"6\" y=\"68\" font-family=\"Cambria Math\" font-size=\"42\">(a+b)/c = d</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 320, 96);
        var baseline = SnapshotMathTypeProcessIds();
        var oldDisable = Environment.GetEnvironmentVariable(disableVariable);
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.InlineShape? shape = null;
        try
        {
            Environment.SetEnvironmentVariable(disableVariable, "1");
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            document.Content.Text = "MathType preview fallback acceptance\r";
            WordEquationNumbering.SetEquationNumberFormat(document, "continuous");
            var service = new WordFormulaService(application);

            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.Select();
            Release(insertion);
            insertion = null;
            service.InsertMathTypeOle(
                CreateMathTypeCreateSession(
                    displayMode: "block",
                    numbered: true,
                    latex: sourceLatex,
                    mathTypeNumberPosition: "right"),
                sourceMathMl,
                emfPath);
            AssertEqual(1, document.InlineShapes.Count,
                "MathType preview fallback setup did not create exactly one OLE formula.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "MathType preview fallback setup did not create exactly one native number.");
            AssertNoNewMathTypeProcess(baseline,
                "MathType preview-disabled insertion");

            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Preview-disabled insertion did not create a recognized MathType OLE object.");
            shape.Range.Select();
            var selection = service.ReadSelection();
            var editSession = CreateMathTypeEditSession(
                selection,
                editedLatex,
                mathTypeNumberPosition: "right");
            editSession.ExportResult = new OfficeExportDocument
            {
                MathMl = editedMathMl,
                Width = 320,
                Height = 96,
                Baseline = 72,
            };
            service.ReplaceMathTypeOle(editSession, editedMathMl, emfPath);

            Release(shape);
            shape = document.InlineShapes[1];
            AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                "Preview-disabled edit changed the MathType OLE class.");
            var actualMathMl = MathTypeOleStorage.ReadMathMl(shape);
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(editedMathMl),
                MathTypeMtefCodec.SemanticSignature(actualMathMl),
                "Preview-disabled edit changed the MathType equation semantics.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Preview-disabled edit lost or duplicated the native equation number.");
            var preview = ReadInlineShapeEnhancedMetafile(shape);
            AssertTrue(
                !string.Equals(
                    DescribeEmfInkBounds(preview),
                    "empty",
                    StringComparison.Ordinal),
                "Preview-disabled edit left a blank Word OLE presentation.");
            AssertNoNewMathTypeProcess(baseline,
                "MathType preview-disabled edit");

            var path = Path.Combine(
                artifactRoot,
                "MathType-Preview-Fallback-Edit.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(
                path,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "Save/reopen lost the preview-fallback MathType formula.");
            Release(shape);
            shape = document.InlineShapes[1];
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(editedMathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(shape)),
                "Save/reopen changed the preview-fallback MathType formula.");
            AssertEqual(1, CountMathTypePlaceRefFields(document),
                "Save/reopen changed the preview-fallback equation number count.");
            AssertTrue(
                !string.Equals(
                    DescribeEmfInkBounds(ReadInlineShapeEnhancedMetafile(shape)),
                    "empty",
                    StringComparison.Ordinal),
                "Save/reopen left the preview-fallback MathType formula blank.");
            AssertNoNewMathTypeProcess(baseline,
                "MathType preview fallback save/reopen");
            Console.WriteLine(
                "[MATHTYPE PREVIEW FALLBACK PASS] Insertion, numbered edit, semantic readback, visible EMF presentation and save/reopen passed with native MathPage rendering forcibly unavailable and no additional MathType process.");
        }
        finally
        {
            Environment.SetEnvironmentVariable(disableVariable, oldDisable);
            Release(shape);
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

    private static List<MathTypeInsertScalingMetric> InsertMathTypeScalingSeries(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        int count,
        string displayMode,
        bool numbered,
        string emfPath,
        string wmfPath,
        int fillerCharactersBeforeEach = 0)
    {
        var metrics = new List<MathTypeInsertScalingMetric>(count);
        for (var index = 1; index <= count; index++)
        {
            Word.Range? insertion = null;
            Word.Range? filler = null;
            try
            {
                document.Activate();
                if (index > 1 && fillerCharactersBeforeEach > 0)
                {
                    var fillerPosition = Math.Max(
                        document.Content.Start,
                        document.Content.End - 1);
                    filler = document.Range(fillerPosition, fillerPosition);
                    filler.Text = new string('p', fillerCharactersBeforeEach) + "\r";
                }
                Release(filler);
                filler = null;
                var position = Math.Max(
                    document.Content.Start,
                    document.Content.End - 1);
                insertion = document.Range(position, position);
                insertion.Select();
            }
            finally
            {
                Release(filler);
                Release(insertion);
            }

            var mathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mi>x</mi><mn>"
                + index.ToString(CultureInfo.InvariantCulture)
                + "</mn></msub><mo>=</mo><mn>"
                + index.ToString(CultureInfo.InvariantCulture)
                + "</mn></math>";
            var session = CreateMathTypeScalingSession(
                index,
                displayMode,
                numbered,
                mathMl);
            var watch = Stopwatch.StartNew();
            service.InsertMathTypeOle(
                session,
                mathMl,
                emfPath,
                isolatedNativePreviewWmfPath: wmfPath,
                isolatedNativePreviewWidthPt: 120,
                isolatedNativePreviewHeightPt: 36,
                isolatedNativePreviewWordPosition: numbered || displayMode == "inline"
                    ? -3
                    : 0,
                isolatedNativePreviewAttempted: true);
            watch.Stop();
            metrics.Add(new MathTypeInsertScalingMetric
            {
                Index = index,
                ElapsedMilliseconds = watch.Elapsed.TotalMilliseconds,
            });
            if (string.Equals(displayMode, "inline", StringComparison.Ordinal))
                application.Selection.TypeText(" ");
            if (index % 10 == 0 || index == count)
                Console.WriteLine(
                    $"[MATHTYPE INSERT {displayMode.ToUpperInvariant()}] "
                    + $"{index}/{count} elapsed={watch.Elapsed.TotalMilliseconds:0}ms");
        }
        return metrics;
    }

    private static OfficeSessionDocument CreateMathTypeScalingSession(
        int index,
        string displayMode,
        bool numbered,
        string mathMl)
    {
        var latex = $"x_{{{index}}}={index}";
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Host = "word",
            Mode = "create",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "MathType insertion scaling",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = displayMode,
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = numbered,
            MathTypeNumberPosition = "right",
            FontSizePt = 12,
            ExportResult = new OfficeExportDocument
            {
                MathMl = mathMl,
                Width = 160,
                Height = 48,
                Baseline = 36,
            },
        };
    }

    private static void AssertMathTypeInsertScaling(
        string context,
        IReadOnlyList<MathTypeInsertScalingMetric> metrics)
    {
        AssertTrue(metrics.Count >= 20, context + ": insufficient scaling samples.");
        var values = metrics.Select(item => item.ElapsedMilliseconds).ToArray();
        var window = Math.Min(12, Math.Max(6, values.Length / 4));
        var early = values.Skip(Math.Min(4, values.Length - window)).Take(window).ToArray();
        var late = values.Skip(values.Length - window).Take(window).ToArray();
        var earlyMedian = MedianScalingMilliseconds(early);
        var lateMedian = MedianScalingMilliseconds(late);
        var ratio = earlyMedian > 0 ? lateMedian / earlyMedian : 0;
        var slope = MathTypeScalingSlope(values);
        Console.WriteLine(
            $"[MATHTYPE INSERT SCALING {context.ToUpperInvariant()}] "
            + $"count={values.Length} earlyMedian={earlyMedian:0.0}ms "
            + $"lateMedian={lateMedian:0.0}ms ratio={ratio:0.000} "
            + $"slope={slope:0.000}ms/formula total={values.Sum():0.0}ms");
        AssertTrue(
            ratio <= 1.35,
            $"{context}: MathType insertion tail median grew by {ratio:0.000}x.");
        AssertTrue(
            slope <= 4.0,
            $"{context}: MathType insertion slope is {slope:0.000}ms per existing formula.");
    }

    private static double MedianScalingMilliseconds(IReadOnlyList<double> values)
    {
        var ordered = values.OrderBy(value => value).ToArray();
        var middle = ordered.Length / 2;
        return ordered.Length % 2 == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2d
            : ordered[middle];
    }

    private static double MathTypeScalingSlope(IReadOnlyList<double> values)
    {
        var meanX = (values.Count + 1) / 2d;
        var meanY = values.Average();
        double numerator = 0;
        double denominator = 0;
        for (var index = 0; index < values.Count; index++)
        {
            var x = index + 1;
            numerator += (x - meanX) * (values[index] - meanY);
            denominator += (x - meanX) * (x - meanX);
        }
        return denominator > 0 ? numerator / denominator : 0;
    }

    private static int CountScalingMathTypeObjects(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (MathTypeOleInterop.IsMathTypeOle(shape)) count++;
            }
            finally { Release(shape); }
        }
        return count;
    }

    private static int CountScalingMathTypeNumberFields(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.Fields.Count; index++)
        {
            Word.Field? field = null;
            Word.Range? code = null;
            try
            {
                field = document.Fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        "MACROBUTTON MTPlaceRef",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    count++;
            }
            finally
            {
                Release(code);
                Release(field);
            }
        }
        return count;
    }

    private static void WriteMathTypeInsertScalingCsv(
        string path,
        IReadOnlyList<MathTypeInsertScalingMetric> metrics)
    {
        var lines = new List<string> { "index,elapsed_ms" };
        lines.AddRange(metrics.Select(item =>
            item.Index.ToString(CultureInfo.InvariantCulture)
            + ","
            + item.ElapsedMilliseconds.ToString("0.000", CultureInfo.InvariantCulture)));
        File.WriteAllLines(path, lines);
    }
}

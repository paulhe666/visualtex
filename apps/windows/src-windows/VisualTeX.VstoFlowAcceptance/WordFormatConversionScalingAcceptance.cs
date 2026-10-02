using System.Diagnostics;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class FormatScalingMetric
    {
        internal long Calls { get; set; }
        internal double Milliseconds { get; set; }
    }

    private sealed class FormatScalingSample
    {
        internal string SourceMode { get; set; } = string.Empty;
        internal string TargetMode { get; set; } = string.Empty;
        internal int Count { get; set; }
        internal double CaptureMilliseconds { get; set; }
        internal double ApplyMilliseconds { get; set; }
        internal IReadOnlyDictionary<string, FormatScalingMetric> CaptureMetrics { get; set; } =
            new Dictionary<string, FormatScalingMetric>();
        internal IReadOnlyDictionary<string, FormatScalingMetric> ApplyMetrics { get; set; } =
            new Dictionary<string, FormatScalingMetric>();
    }

    private static void RunWordFormatConversionScalingAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var countsText =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_FORMAT_SCALING_COUNTS");
        var counts =
            string.IsNullOrWhiteSpace(countsText)
                ? new[] { 10, 20, 40 }
                : countsText!
                    .Split(',')
                    .Select(value =>
                        int.Parse(
                            value.Trim(),
                            System.Globalization.CultureInfo.InvariantCulture))
                    .Where(value => value > 0)
                    .Distinct()
                    .OrderBy(value => value)
                    .ToArray();
        if (counts.Length == 0)
            throw new InvalidDataException(
                "Format-conversion scaling requires at least one positive count.");

        var assetRoot = Path.Combine(
            Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "format-scaling-"
            + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        var pngPath =
            Path.Combine(
                assetRoot,
                "format-scaling.png");
        var svgPath =
            Path.Combine(
                assetRoot,
                "format-scaling.svg");
        WriteAcceptancePng(
            pngPath,
            "x+1",
            260,
            96);
        File.WriteAllText(
            svgPath,
            @"<svg xmlns=""http://www.w3.org/2000/svg"" width=""260"" height=""96"" viewBox=""0 0 260 96""><text x=""4"" y=""64"" font-family=""Times New Roman"" font-size=""44"">x+1</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                260,
                96);

        var modes =
            new[]
            {
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode,
            };
        var requestedSource =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_FORMAT_SCALING_SOURCE");
        var requestedTarget =
            Environment.GetEnvironmentVariable(
                "VISUALTEX_FORMAT_SCALING_TARGET");
        var reuseExistingSource =
            string.Equals(
                Environment.GetEnvironmentVariable(
                    "VISUALTEX_FORMAT_SCALING_REUSE_EXISTING_SOURCE"),
                "1",
                StringComparison.Ordinal);
        var samples =
            new List<FormatScalingSample>();

        try
        {
            foreach (var count in counts)
            {
                foreach (var sourceMode in modes)
                {
                    if (!string.IsNullOrWhiteSpace(requestedSource)
                        && !string.Equals(
                            ModeKey(sourceMode),
                            requestedSource,
                            StringComparison.OrdinalIgnoreCase))
                        continue;
                    var sourceFixture =
                        Path.Combine(
                            artifactRoot,
                            $"format-scaling-{ModeKey(sourceMode)}-{count}.docx");
                    if (reuseExistingSource
                        && File.Exists(sourceFixture))
                    {
                        Console.WriteLine(
                            $"[FORMAT SCALING] reusing source fixture {sourceFixture}");
                    }
                    else
                    {
                        CreateFormatScalingSourceFixture(
                            artifactRoot,
                            sourceFixture,
                            sourceMode,
                            count,
                            pngPath,
                            emfPath);
                    }

                    foreach (var targetMode in modes)
                    {
                        if (string.Equals(
                                sourceMode,
                                targetMode,
                                StringComparison.Ordinal))
                            continue;
                        if (!string.IsNullOrWhiteSpace(requestedTarget)
                            && !string.Equals(
                                ModeKey(targetMode),
                                requestedTarget,
                                StringComparison.OrdinalIgnoreCase))
                            continue;

                        var conversionPath =
                            Path.Combine(
                                artifactRoot,
                                $"format-scaling-{ModeKey(sourceMode)}-to-{ModeKey(targetMode)}-{count}.docx");
                        File.Copy(
                            sourceFixture,
                            conversionPath,
                            overwrite: true);
                        var sample =
                            MeasureFormatScalingConversion(
                                artifactRoot,
                                conversionPath,
                                sourceMode,
                                targetMode,
                                count,
                                pngPath,
                                emfPath);
                        samples.Add(sample);
                        Console.WriteLine(
                            $"[FORMAT SCALING] {ModeKey(sourceMode)}->{ModeKey(targetMode)} "
                            + $"count={count} captureMs={sample.CaptureMilliseconds:0.0} "
                            + $"applyMs={sample.ApplyMilliseconds:0.0} "
                            + $"capturePer={sample.CaptureMilliseconds / count:0.00} "
                            + $"applyPer={sample.ApplyMilliseconds / count:0.00}");
                    }
                }
            }

            var report =
                new List<string>
                {
                    "source,target,count,capture_ms,apply_ms,capture_per_formula_ms,apply_per_formula_ms"
                };
            foreach (var sample in samples
                         .OrderBy(item => item.SourceMode, StringComparer.Ordinal)
                         .ThenBy(item => item.TargetMode, StringComparer.Ordinal)
                         .ThenBy(item => item.Count))
            {
                report.Add(
                    string.Join(
                        ",",
                        ModeKey(sample.SourceMode),
                        ModeKey(sample.TargetMode),
                        sample.Count.ToString(
                            System.Globalization.CultureInfo.InvariantCulture),
                        sample.CaptureMilliseconds.ToString(
                            "0.###",
                            System.Globalization.CultureInfo.InvariantCulture),
                        sample.ApplyMilliseconds.ToString(
                            "0.###",
                            System.Globalization.CultureInfo.InvariantCulture),
                        (sample.CaptureMilliseconds / sample.Count).ToString(
                            "0.###",
                            System.Globalization.CultureInfo.InvariantCulture),
                        (sample.ApplyMilliseconds / sample.Count).ToString(
                            "0.###",
                            System.Globalization.CultureInfo.InvariantCulture)));
            }
            File.WriteAllLines(
                Path.Combine(
                    artifactRoot,
                    "format-conversion-scaling.csv"),
                report);

            var metricLines =
                new List<string>();
            foreach (var sample in samples)
            {
                metricLines.Add(
                    $"[{ModeKey(sample.SourceMode)}->{ModeKey(sample.TargetMode)} count={sample.Count}]");
                foreach (var pair in sample.CaptureMetrics
                             .OrderByDescending(pair => pair.Value.Milliseconds))
                    metricLines.Add(
                        $"capture|{pair.Key}|calls={pair.Value.Calls}|ms={pair.Value.Milliseconds:0.###}");
                foreach (var pair in sample.ApplyMetrics
                             .OrderByDescending(pair => pair.Value.Milliseconds))
                    metricLines.Add(
                        $"apply|{pair.Key}|calls={pair.Value.Calls}|ms={pair.Value.Milliseconds:0.###}");
            }
            File.WriteAllLines(
                Path.Combine(
                    artifactRoot,
                    "format-conversion-scaling-metrics.txt"),
                metricLines);

            Console.WriteLine(
                $"[FORMAT SCALING COMPLETE] samples={samples.Count}; report={Path.Combine(artifactRoot, "format-conversion-scaling.csv")}");
        }
        finally
        {
            try
            {
                Directory.Delete(
                    assetRoot,
                    recursive: true);
            }
            catch { }
        }
    }

    private static void CreateFormatScalingSourceFixture(
        string artifactRoot,
        string fixturePath,
        string sourceMode,
        int count,
        string pngPath,
        string emfPath,
        bool numbered = true,
        bool mixedDisplay = false)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            // Always isolate scaling fixtures from the user's live Word.
            // Office can otherwise reuse/switch the ROT-active WINWORD instance
            // between fixture creation and conversion, which correctly trips the
            // product's source-document identity guard but invalidates the
            // benchmark itself.
            application =
                CreateFreshFormatScalingWord(
                    artifactRoot,
                    $"source-{ModeKey(sourceMode)}-{count}");
            document =
                application.Documents.Add();
            document.Activate();
            document.Content.Text =
                "VisualTeX format conversion scaling source\r";
            WordEquationNumbering
                .SetEquationNumberFormatPreference(
                    document,
                    EquationNumberFormat.ContinuousId);
            var service =
                new WordFormulaService(
                    application);

            for (var index = 1;
                 index <= count;
                 index++)
            {
                SelectDocumentEnd(
                    document);
                var displayMode = mixedDisplay && index % 2 == 1 ? "inline" : "block";
                if (displayMode == "inline")
                {
                    var prefix = "INLINE[" + index + "] ";
                    var inlineStart = application.Selection.Start;
                    application.Selection.TypeText(prefix + " [END-INLINE]");
                    application.Selection.SetRange(inlineStart + prefix.Length, inlineStart + prefix.Length);
                }
                var mathMl =
                    FormatScalingMathMl(
                        index);
                var session =
                    CreateOmmlMathTypeAcceptanceSession(
                        mathMl,
                        displayMode,
                        numbered: numbered && displayMode == "block",
                        sourceMode);
                session.ExportResult =
                    new OfficeExportDocument
                    {
                        MathMl = mathMl,
                        Width = 260,
                        Height = 96,
                        Baseline = 72,
                    };

                if (string.Equals(
                        sourceMode,
                        FormulaOleContract.WordOmmlMode,
                        StringComparison.Ordinal))
                {
                    service.InsertOmml(
                        session,
                        mathMl);
                }
                else if (string.Equals(
                             sourceMode,
                             FormulaOleContract.NativeOleMode,
                             StringComparison.Ordinal))
                {
                    service.InsertOle(
                        session,
                        pngPath,
                        emfPath);
                }
                else
                {
                    service.InsertMathTypeOle(
                        session,
                        mathMl,
                        emfPath,
                        isolatedNativePreviewAttempted: true);
                }

                if (index < count)
                {
                    SelectDocumentEnd(
                        document);
                    application.Selection.TypeParagraph();
                }
            }

            if (string.Equals(
                    sourceMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
                AssertEqual(
                    count,
                    document.OMaths.Count,
                    "Scaling OMML source fixture has the wrong formula count.");
            else if (string.Equals(
                         sourceMode,
                         FormulaOleContract.NativeOleMode,
                         StringComparison.Ordinal))
                AssertEqual(
                    count,
                    CountVisualTeXNativeOleShapes(
                        document),
                    "Scaling VisualTeX source fixture has the wrong formula count.");
            else
                AssertEqual(
                    count,
                    CountMathTypeOleShapes(
                        document),
                    "Scaling MathType source fixture has the wrong formula count.");

            document.SaveAs2(
                fixturePath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
        }
        finally
        {
            try
            {
                document?.Close(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
            try
            {
                QuitWordApplicationIfOwned(
                    application);
            }
            catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static FormatScalingSample MeasureFormatScalingConversion(
        string artifactRoot,
        string conversionPath,
        string sourceMode,
        string targetMode,
        int count,
        string pngPath,
        string emfPath)
    {
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application =
                CreateFreshFormatScalingWord(
                    artifactRoot,
                    $"convert-{ModeKey(sourceMode)}-to-{ModeKey(targetMode)}-{count}");
            document =
                application.Documents.Open(
                    conversionPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
            document.Activate();
            var service =
                new WordFormulaService(
                    application);

            WordFormulaFormatConversionPlan plan;
            var captureWatch =
                Stopwatch.StartNew();
            IReadOnlyDictionary<string, FormatScalingMetric>
                captureEntries;
            using (var metrics =
                   new WordOperationMetrics())
            {
                plan =
                    service
                        .CaptureFormulaFormatConversionPlan(
                            wholeDocument: true,
                            sourceMode,
                            targetMode);
                captureEntries =
                    SnapshotFormatScalingMetrics(
                        metrics);
            }
            captureWatch.Stop();

            AssertEqual(
                count,
                plan.Targets.Count,
                $"Scaling {ModeKey(sourceMode)}->{ModeKey(targetMode)} capture returned the wrong formula count.");

            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                var mathMl =
                    target.SourceMathMl
                    ?? FormatScalingMathMlFromLatex(
                        target.Latex);
                var session =
                    CreateSimpleFormatTargetSession(
                        target,
                        targetMode,
                        mathMl);
                prepared[target.Id] =
                    new PreparedWordBulkFormula
                    {
                        Run =
                            new WordBulkRun
                            {
                                Id = target.Id,
                                IsFormula = true,
                                Latex = target.Latex,
                                DisplayMode =
                                    target.DisplayMode,
                            },
                        Session = session,
                        MathMl = mathMl,
                        PngPath =
                            string.Equals(
                                targetMode,
                                FormulaOleContract.NativeOleMode,
                                StringComparison.Ordinal)
                                ? pngPath
                                : null,
                        EmfPath =
                            targetMode is
                                FormulaOleContract.NativeOleMode
                                or FormulaOleContract.MathTypeOleMode
                                ? emfPath
                                : null,
                        MathTypeNativePreviewAttempted =
                            string.Equals(
                                targetMode,
                                FormulaOleContract.MathTypeOleMode,
                                StringComparison.Ordinal),
                    };
            }

            WordFormulaFormatConversionResult result;
            var applyWatch =
                Stopwatch.StartNew();
            IReadOnlyDictionary<string, FormatScalingMetric>
                applyEntries;
            using (var metrics =
                   new WordOperationMetrics())
            {
                try
                {
                    result =
                        service
                            .ApplyFormulaFormatConversionPlan(
                                plan,
                                prepared);
                }
                catch (Exception error)
                {
                    throw new InvalidDataException(
                        $"Scaling {ModeKey(sourceMode)}->{ModeKey(targetMode)} apply failed for count={count}. "
                        + FormatScalingExceptionTree(error),
                        error);
                }
                applyEntries =
                    SnapshotFormatScalingMetrics(
                        metrics);
            }
            applyWatch.Stop();

            AssertEqual(
                count,
                result.FormulaCount,
                $"Scaling {ModeKey(sourceMode)}->{ModeKey(targetMode)} did not convert every formula.");
            AssertEqual(
                0,
                result.FailedFormulaCount,
                $"Scaling {ModeKey(sourceMode)}->{ModeKey(targetMode)} reported failures.");

            if (string.Equals(
                    targetMode,
                    FormulaOleContract.WordOmmlMode,
                    StringComparison.Ordinal))
            {
                AssertEqual(
                    count,
                    document.OMaths.Count,
                    "Scaling target OMML formula count changed.");
            }
            else if (string.Equals(
                         targetMode,
                         FormulaOleContract.NativeOleMode,
                         StringComparison.Ordinal))
            {
                AssertEqual(
                    count,
                    CountVisualTeXNativeOleShapes(
                        document),
                    "Scaling target VisualTeX formula count changed.");
            }
            else
            {
                AssertEqual(
                    count,
                    CountMathTypeOleShapes(
                        document),
                    "Scaling target MathType formula count changed.");
            }

            document.Save();
            return new FormatScalingSample
            {
                SourceMode =
                    sourceMode,
                TargetMode =
                    targetMode,
                Count =
                    count,
                CaptureMilliseconds =
                    captureWatch.Elapsed
                        .TotalMilliseconds,
                ApplyMilliseconds =
                    applyWatch.Elapsed
                        .TotalMilliseconds,
                CaptureMetrics =
                    captureEntries,
                ApplyMetrics =
                    applyEntries,
            };
        }
        finally
        {
            try
            {
                document?.Close(
                    Word.WdSaveOptions.wdDoNotSaveChanges);
            }
            catch { }
            try
            {
                QuitWordApplicationIfOwned(
                    application);
            }
            catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunWordFormatConversionMultiBodyFlatOpcAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "format-multibody-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        var pngPath = Path.Combine(assetRoot, "preview.png");
        var svgPath = Path.Combine(assetRoot, "preview.svg");
        WriteAcceptancePng(pngPath, "x+y", 280, 96);
        File.WriteAllText(
            svgPath,
            @"<svg xmlns=""http://www.w3.org/2000/svg"" width=""280"" height=""96"" viewBox=""0 0 280 96""><text x=""6"" y=""66"" font-family=""Cambria Math"" font-size=""44"">x+y</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 280, 96);

        Word.Application? application = null;
        Word.Document? document = null;

        string AddAuxiliaryWordprocessingBody(
            string flatOpc,
            string marker)
        {
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var package = XDocument.Parse(flatOpc);
            if (package.Root?.Name != pkg + "package")
                throw new InvalidDataException(
                    "Multi-body acceptance expected a Flat OPC package root.");
            package.Root.Add(
                new XElement(
                    pkg + "part",
                    new XAttribute(
                        pkg + "name",
                        "/customXml/visualtex-acceptance-" + marker + ".xml"),
                    new XAttribute(
                        pkg + "contentType",
                        "application/xml"),
                    new XElement(
                        pkg + "xmlData",
                        new XElement(
                            w + "document",
                            new XElement(
                                w + "body",
                                new XElement(
                                    w + "p",
                                    new XElement(
                                        w + "r",
                                        new XElement(
                                            w + "t",
                                            "auxiliary recovery evidence " + marker))),
                                new XElement(w + "sectPr"))))));
            return package.ToString(SaveOptions.DisableFormatting);
        }

        void AssertMultiBodyRecoveryEvidence(
            string flatOpc,
            int expectedMathCount,
            int expectedOleCount,
            string stage)
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            XNamespace o = "urn:schemas-microsoft-com:office:office";
            var package = XDocument.Parse(flatOpc);
            var allBodyCount = package.Descendants(w + "body").Count();
            AssertTrue(allBodyCount >= 2,
                stage + ": injected Flat OPC did not contain multiple w:body elements.");
            var mainBody = WordDocumentXml.GetRequiredMainDocumentBody(package);
            AssertEqual(expectedMathCount, mainBody.Descendants(m + "oMath").Count(),
                stage + ": main /word/document.xml OMML count changed during body selection.");
            AssertEqual(expectedOleCount, mainBody.Descendants(o + "OLEObject").Count(),
                stage + ": main /word/document.xml OLE count changed during body selection.");

            // This is the exact recovery/signature path that previously called
            // Descendants(w:body).Single() and threw on a valid multi-body package.
            var normalizedBody = WordLocalEditSnapshot.NormalizedBody(flatOpc);
            var normalized = XElement.Parse(normalizedBody);
            AssertEqual(expectedMathCount, normalized.Descendants(m + "oMath").Count(),
                stage + ": recovery normalization selected a non-main Word body.");
            AssertEqual(expectedOleCount, normalized.Descendants(o + "OLEObject").Count(),
                stage + ": recovery normalization selected a non-main OLE body.");
            AssertTrue(!string.IsNullOrWhiteSpace(WordLocalEditSnapshot.Signature(flatOpc)),
                stage + ": recovery signature was not produced for multi-body Flat OPC.");
        }

        try
        {
            application = CreateFreshFormatScalingWord(
                artifactRoot,
                "multibody-flatopc");
            document = application.Documents.Add();
            document.Activate();
            document.Content.Text = "multi-body OMML source\r";
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);

            for (var index = 1; index <= 3; index++)
            {
                SelectDocumentEnd(document);
                var mathMl = FormatScalingMathMl(index);
                service.InsertOmml(
                    CreateOmmlMathTypeAcceptanceSession(
                        mathMl,
                        "block",
                        numbered: index != 2,
                        FormulaOleContract.WordOmmlMode),
                    mathMl);
                if (index < 3)
                {
                    SelectDocumentEnd(document);
                    application.Selection.TypeParagraph();
                }
            }
            AssertEqual(3, document.OMaths.Count,
                "Multi-body fixture did not create three OMML formulas.");

            // Start from Word's actual Flat OPC for this real 3-OMML document,
            // then add the legal auxiliary WordprocessingML part that triggered
            // the former global Descendants(w:body).Single() assumption.
            var beforeFlatOpc = AddAuxiliaryWordprocessingBody(
                document.Content.WordOpenXML,
                "before");
            AssertMultiBodyRecoveryEvidence(
                beforeFlatOpc,
                expectedMathCount: 3,
                expectedOleCount: 0,
                stage: "before OMML→VisualTeX");
            var beforeBodyCount = XDocument.Parse(beforeFlatOpc)
                .Descendants(XName.Get("body", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                .Count();
            File.WriteAllText(
                Path.Combine(artifactRoot, "multibody-before-flatopc.xml"),
                beforeFlatOpc);

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(3, plan.Targets.Count,
                "Multi-body OMML→VisualTeX capture did not resolve three formulas.");

            var prepared = new Dictionary<string, PreparedWordBulkFormula>(
                StringComparer.Ordinal);
            foreach (var target in plan.Targets)
            {
                var mathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        "Multi-body OMML source exposed no MathML.");
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
                        FormulaOleContract.NativeOleMode,
                        mathMl),
                    MathMl = mathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                };
            }

            var result = service.ApplyFormulaFormatConversionPlan(
                plan,
                prepared);
            AssertEqual(3, result.FormulaCount,
                "Multi-body OMML→VisualTeX did not convert all formulas.");
            AssertEqual(0, result.FailedFormulaCount,
                "Multi-body OMML→VisualTeX reported a failed formula.");
            AssertEqual(0, document.OMaths.Count,
                "Multi-body OMML→VisualTeX left OMML sources behind.");
            AssertEqual(3, CountVisualTeXNativeOleShapes(document),
                "Multi-body OMML→VisualTeX created the wrong VisualTeX OLE count.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Multi-body OMML→VisualTeX changed numbered state.");

            var afterFlatOpc = AddAuxiliaryWordprocessingBody(
                document.Content.WordOpenXML,
                "after");
            AssertMultiBodyRecoveryEvidence(
                afterFlatOpc,
                expectedMathCount: 0,
                expectedOleCount: 3,
                stage: "after OMML→VisualTeX");
            var afterBodyCount = XDocument.Parse(afterFlatOpc)
                .Descendants(XName.Get("body", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                .Count();
            File.WriteAllText(
                Path.Combine(artifactRoot, "multibody-after-flatopc.xml"),
                afterFlatOpc);

            var outputPath = Path.Combine(
                artifactRoot,
                "format-conversion-multibody-flatopc.docx");
            document.SaveAs2(
                outputPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            Console.WriteLine(
                $"[FORMAT MULTIBODY FLAT OPC] bodies={beforeBodyCount}->{afterBodyCount}; "
                + $"OMML=3→0; VisualTeX=0→3; numbered=2; output={outputPath}");
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(document);
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
        }
    }

    private static void RunWordFormatConversionTableMixedAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "format-table-mixed-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(assetRoot);
        var pngPath = Path.Combine(assetRoot, "preview.png");
        var svgPath = Path.Combine(assetRoot, "preview.svg");
        WriteAcceptancePng(pngPath, "mixed", 300, 96);
        File.WriteAllText(
            svgPath,
            @"<svg xmlns=""http://www.w3.org/2000/svg"" width=""300"" height=""96"" viewBox=""0 0 300 96""><text x=""6"" y=""66"" font-family=""Cambria Math"" font-size=""44"">mixed</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 300, 96);

        const string bodyInline =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mn>1</mn></math>";
        const string bodyNumbered =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>";
        const string tableInline =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>p</mi><mo>=</mo><mi>q</mi></math>";
        const string tableNumbered =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>F</mi><mo>=</mo><mi>m</mi><mi>a</mi></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        Word.Table? userTable = null;
        try
        {
            application = CreateFreshFormatScalingWord(
                artifactRoot,
                "table-mixed");
            document = application.Documents.Add();
            document.Activate();
            document.Content.Text =
                "body-left BODY_INLINE body-right\r"
                + "body-before-display\rBODY_DISPLAY\rbody-after-display\r"
                + "table-follows\r";
            var service = new WordFormulaService(application);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);

            void ReplaceTokenWithOmml(
                string token,
                string mathMl,
                string displayMode,
                bool numbered)
            {
                Word.Range? insertion = null;
                try
                {
                    insertion = FindAcceptanceTextRange(document, token);
                    insertion.Text = string.Empty;
                    insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    insertion.Select();
                    service.InsertOmml(
                        CreateOmmlMathTypeAcceptanceSession(
                            mathMl,
                            displayMode,
                            numbered,
                            FormulaOleContract.WordOmmlMode),
                        mathMl);
                }
                finally { Release(insertion); }
            }

            ReplaceTokenWithOmml(
                "BODY_INLINE",
                bodyInline,
                "inline",
                numbered: false);
            ReplaceTokenWithOmml(
                "BODY_DISPLAY",
                bodyNumbered,
                "block",
                numbered: true);

            Word.Range? tableInsertion = null;
            try
            {
                tableInsertion = document.Range(
                    document.Content.End - 1,
                    document.Content.End - 1);
                userTable = document.Tables.Add(
                    tableInsertion,
                    2,
                    2);
            }
            finally { Release(tableInsertion); }

            void InsertCellOmml(
                int row,
                int column,
                string mathMl,
                string displayMode,
                bool numbered,
                bool addInlineGuards)
            {
                Word.Cell? cell = null;
                Word.Range? body = null;
                Word.Range? insertion = null;
                try
                {
                    cell = userTable!.Cell(row, column);
                    body = cell.Range.Duplicate;
                    body.End = Math.Max(body.Start, body.End - 1);
                    if (addInlineGuards)
                    {
                        body.Text = "L R";
                        insertion = document.Range(body.Start + 1, body.Start + 1);
                    }
                    else
                    {
                        body.Text = string.Empty;
                        insertion = document.Range(body.Start, body.Start);
                    }
                    insertion.Select();
                    service.InsertOmml(
                        CreateOmmlMathTypeAcceptanceSession(
                            mathMl,
                            displayMode,
                            numbered,
                            FormulaOleContract.WordOmmlMode),
                        mathMl);
                }
                finally
                {
                    Release(insertion);
                    Release(body);
                    Release(cell);
                }
            }

            InsertCellOmml(
                1,
                1,
                tableInline,
                "inline",
                numbered: false,
                addInlineGuards: true);
            InsertCellOmml(
                2,
                2,
                tableNumbered,
                "block",
                numbered: true,
                addInlineGuards: false);

            int CountOmmlInsideUserTable(Word.Document current)
            {
                var count = 0;
                for (var index = 1; index <= current.OMaths.Count; index++)
                {
                    Word.OMath? math = null;
                    Word.Range? range = null;
                    try
                    {
                        math = current.OMaths[index];
                        range = math.Range.Duplicate;
                        if (Convert.ToBoolean(
                                range.get_Information(
                                    Word.WdInformation.wdWithInTable)))
                            count++;
                    }
                    finally
                    {
                        Release(range);
                        Release(math);
                    }
                }
                return count;
            }

            int CountVisualTeXInsideUserTable(Word.Document current)
            {
                var count = 0;
                for (var index = 1; index <= current.InlineShapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    Word.Range? range = null;
                    try
                    {
                        shape = current.InlineShapes[index];
                        if (!WordFormulaMetadataReader.IsNativeOle(shape))
                            continue;
                        range = shape.Range.Duplicate;
                        if (Convert.ToBoolean(
                                range.get_Information(
                                    Word.WdInformation.wdWithInTable)))
                            count++;
                    }
                    finally
                    {
                        Release(range);
                        Release(shape);
                    }
                }
                return count;
            }

            int CountMathTypeInsideUserTable(Word.Document current)
            {
                var count = 0;
                for (var index = 1; index <= current.InlineShapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    Word.Range? range = null;
                    try
                    {
                        shape = current.InlineShapes[index];
                        if (!MathTypeOleInterop.IsMathTypeOle(shape))
                            continue;
                        range = shape.Range.Duplicate;
                        if (Convert.ToBoolean(
                                range.get_Information(
                                    Word.WdInformation.wdWithInTable)))
                            count++;
                    }
                    finally
                    {
                        Release(range);
                        Release(shape);
                    }
                }
                return count;
            }

            void AssertUserTableShape(string stage)
            {
                AssertEqual(1, document.Tables.Count,
                    stage + ": conversion added or removed a user table.");
                Word.Table? table = null;
                Word.Rows? rows = null;
                Word.Columns? columns = null;
                try
                {
                    table = document.Tables[1];
                    rows = table.Rows;
                    columns = table.Columns;
                    AssertEqual(2, rows.Count,
                        stage + ": user-table row count changed.");
                    AssertEqual(2, columns.Count,
                        stage + ": user-table column count changed.");
                }
                finally
                {
                    Release(columns);
                    Release(rows);
                    Release(table);
                }
                var text = document.Content.Text ?? string.Empty;
                AssertTrue(
                    text.IndexOf("body-left", StringComparison.Ordinal) >= 0
                    && text.IndexOf("body-right", StringComparison.Ordinal) >= 0
                    && text.IndexOf("body-before-display", StringComparison.Ordinal) >= 0
                    && text.IndexOf("body-after-display", StringComparison.Ordinal) >= 0
                    && text.IndexOf("table-follows", StringComparison.Ordinal) >= 0,
                    stage + ": surrounding prose changed.");
            }

            Dictionary<string, PreparedWordBulkFormula> PrepareVisualTeX(
                WordFormulaFormatConversionPlan plan)
            {
                var prepared =
                    new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
                foreach (var target in plan.Targets)
                {
                    var mathMl = target.SourceMathMl
                        ?? throw new InvalidDataException(
                            "Table-mixed OMML target exposed no MathML.");
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
                            FormulaOleContract.NativeOleMode,
                            mathMl),
                        MathMl = mathMl,
                        PngPath = pngPath,
                        EmfPath = emfPath,
                    };
                }
                return prepared;
            }

            AssertEqual(4, document.OMaths.Count,
                "Table-mixed fixture did not create four OMML formulas.");
            AssertEqual(2, CountOmmlInsideUserTable(document),
                "Table-mixed fixture did not keep two OMML formulas in the user table.");
            AssertEqual(2, CountManagedNumberedOmml(document),
                "Table-mixed fixture did not create two native numbered OMML formulas.");
            AssertUserTableShape("initial OMML fixture");

            var ommlToVisualTeX = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(4, ommlToVisualTeX.Targets.Count,
                "Table-mixed OMML→VisualTeX did not capture four formulas.");
            AssertEqual(2, ommlToVisualTeX.Targets.Count(item => item.SourceWithinTable),
                "Table-mixed OMML→VisualTeX lost table ownership at capture.");
            var sourceMathMlByLatex =
                ommlToVisualTeX.Targets.ToDictionary(
                    target => target.Latex,
                    target => target.SourceMathMl
                        ?? throw new InvalidDataException(
                            "Table-mixed OMML source exposed no MathML."),
                    StringComparer.Ordinal);
            var vtResult = service.ApplyFormulaFormatConversionPlan(
                ommlToVisualTeX,
                PrepareVisualTeX(ommlToVisualTeX));
            AssertEqual(4, vtResult.FormulaCount,
                "Table-mixed OMML→VisualTeX did not convert all formulas.");
            AssertEqual(0, vtResult.FailedFormulaCount,
                "Table-mixed OMML→VisualTeX reported a failed formula.");
            AssertEqual(0, document.OMaths.Count,
                "Table-mixed OMML→VisualTeX left OMath sources.");
            AssertEqual(4, CountVisualTeXNativeOleShapes(document),
                "Table-mixed OMML→VisualTeX produced the wrong OLE count.");
            AssertEqual(2, CountVisualTeXInsideUserTable(document),
                "Table-mixed OMML→VisualTeX moved a formula across the table boundary.");
            AssertEqual(2, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "Table-mixed OMML→VisualTeX changed numbered state.");
            AssertUserTableShape("after OMML→VisualTeX");

            var visualTeXToOmml = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.WordOmmlMode);
            var visualTeXToOmmlPrepared =
                new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            foreach (var target in visualTeXToOmml.Targets)
            {
                if (!sourceMathMlByLatex.TryGetValue(
                        target.Latex,
                        out var mathMl))
                    throw new InvalidDataException(
                        "Table-mixed VisualTeX source could not be matched to its original MathML.");
                visualTeXToOmmlPrepared[target.Id] = new PreparedWordBulkFormula
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
                        mathMl),
                    MathMl = mathMl,
                };
            }
            var ommlBackResult = service.ApplyFormulaFormatConversionPlan(
                visualTeXToOmml,
                visualTeXToOmmlPrepared);
            AssertEqual(4, ommlBackResult.FormulaCount,
                "Table-mixed VisualTeX→OMML did not convert all formulas.");
            AssertEqual(0, ommlBackResult.FailedFormulaCount,
                "Table-mixed VisualTeX→OMML reported a failed formula.");
            AssertEqual(4, document.OMaths.Count,
                "Table-mixed VisualTeX→OMML produced the wrong OMath count.");
            AssertEqual(2, CountOmmlInsideUserTable(document),
                "Table-mixed VisualTeX→OMML moved a formula across the table boundary.");
            AssertEqual(2, CountManagedNumberedOmml(document),
                "Table-mixed VisualTeX→OMML changed native numbered state.");
            AssertEqual(0, CountVisualTeXNumberingBookmarks(document),
                "Table-mixed VisualTeX→OMML left VisualTeX numbering bookmarks.");
            AssertEqual(0, CountVisualTeXPlaceRefFields(document),
                "Table-mixed VisualTeX→OMML left VisualTeXPlaceRef fields.");
            AssertUserTableShape("after VisualTeX→OMML");

            var ommlToMathType = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.MathTypeOleMode);
            var mtResult = service.ApplyFormulaFormatConversionPlan(
                ommlToMathType,
                PrepareOmmlMathTypeTargets(ommlToMathType, emfPath));
            AssertEqual(4, mtResult.FormulaCount,
                "Table-mixed OMML→MathType did not convert all formulas.");
            AssertEqual(0, mtResult.FailedFormulaCount,
                "Table-mixed OMML→MathType reported a failed formula.");
            AssertEqual(4, CountMathTypeOleShapes(document),
                "Table-mixed OMML→MathType produced the wrong OLE count.");
            AssertEqual(2, CountMathTypeInsideUserTable(document),
                "Table-mixed OMML→MathType moved a formula across the table boundary.");
            AssertUserTableShape("after OMML→MathType");

            var mathTypeToOmml = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.MathTypeOleMode,
                FormulaOleContract.WordOmmlMode);
            var finalResult = service.ApplyFormulaFormatConversionPlan(
                mathTypeToOmml,
                PrepareOmmlMathTypeTargets(mathTypeToOmml, emfPath));
            AssertEqual(4, finalResult.FormulaCount,
                "Table-mixed MathType→OMML did not convert all formulas.");
            AssertEqual(0, finalResult.FailedFormulaCount,
                "Table-mixed MathType→OMML reported a failed formula.");
            AssertEqual(4, document.OMaths.Count,
                "Table-mixed MathType→OMML produced the wrong OMath count.");
            AssertEqual(2, CountOmmlInsideUserTable(document),
                "Table-mixed MathType→OMML moved a formula across the table boundary.");
            AssertEqual(2, CountManagedNumberedOmml(document),
                "Table-mixed MathType→OMML changed native numbered state.");
            AssertUserTableShape("after MathType→OMML");

            var outputPath = Path.Combine(
                artifactRoot,
                "format-conversion-table-mixed.docx");
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
                AddToRecentFiles: false);
            document = reopened;
            reopened = null;
            AssertEqual(4, document.OMaths.Count,
                "Table-mixed save/reopen changed the OMath count.");
            AssertEqual(2, CountOmmlInsideUserTable(document),
                "Table-mixed save/reopen moved a formula across the table boundary.");
            AssertEqual(2, CountManagedNumberedOmml(document),
                "Table-mixed save/reopen changed native numbering.");
            AssertUserTableShape("after save/reopen");
            Console.WriteLine(
                "[FORMAT TABLE MIXED] OMML→VisualTeX→OMML→MathType→OMML preserved four formulas, two user-table owners, inline/display and numbered states, prose, and save/reopen.");
        }
        finally
        {
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(userTable);
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
        }
    }

    private static void RunWordExistingOmmlToVisualTeXCaptureAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_EXISTING_OMML_VT_SOURCE_PATH");
        if (string.IsNullOrWhiteSpace(sourcePath)
            || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "Existing OMML→VisualTeX capture requires VISUALTEX_EXISTING_OMML_VT_SOURCE_PATH.",
                sourcePath);

        var localPath = Path.Combine(
            artifactRoot,
            "existing-omml-vt-source-copy.docx");
        File.Copy(
            sourcePath,
            localPath,
            overwrite: true);

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateFreshFormatScalingWord(
                artifactRoot,
                "existing-omml-vt-capture");
            document = application.Documents.Open(
                localPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var beforeOmml = document.OMaths.Count;
            var beforeVisualTeX =
                CountVisualTeXNativeOleShapes(
                    document);
            Console.WriteLine(
                $"[EXISTING OMML→VT] before OMML={beforeOmml} VT={beforeVisualTeX} path={localPath}");

            var service =
                new WordFormulaService(
                    application);
            WordFormulaFormatConversionPlan plan;
            try
            {
                plan =
                    service.CaptureFormulaFormatConversionPlan(
                        wholeDocument: true,
                        FormulaOleContract.WordOmmlMode,
                        FormulaOleContract.NativeOleMode);
            }
            catch (Exception error)
            {
                var diagnosticLines =
                    new List<string>
                    {
                        "whole-document=" + FormatScalingExceptionTree(error),
                        error.ToString(),
                    };
                File.WriteAllLines(
                    Path.Combine(
                        artifactRoot,
                        "existing-omml-vt-capture-primary-error.txt"),
                    diagnosticLines);
                var failureFound = false;
                for (var index = 1;
                     index <= document.OMaths.Count;
                     index++)
                {
                    Word.OMath? math = null;
                    Word.Range? mathRange = null;
                    try
                    {
                        math = document.OMaths[index];
                        mathRange = math.Range.Duplicate;
                        mathRange.Select();
                        try
                        {
                            var localPlan =
                                service.CaptureFormulaFormatConversionPlan(
                                    wholeDocument: false,
                                    FormulaOleContract.WordOmmlMode,
                                    FormulaOleContract.NativeOleMode);
                            diagnosticLines.Add(
                                $"formula={index} range={mathRange.Start}:{mathRange.End} type={math.Type} targets={localPlan.Targets.Count} ok");
                        }
                        catch (Exception localError)
                        {
                            failureFound = true;
                            diagnosticLines.Add(
                                $"formula={index} range={mathRange.Start}:{mathRange.End} type={math.Type} error={FormatScalingExceptionTree(localError)}");
                            try
                            {
                                File.WriteAllText(
                                    Path.Combine(
                                        artifactRoot,
                                        $"existing-omml-vt-failing-{index:D4}.xml"),
                                    mathRange.WordOpenXML);
                            }
                            catch { }

                            Word.Paragraphs? paragraphs = null;
                            Word.Paragraph? paragraph = null;
                            Word.Range? paragraphRange = null;
                            try
                            {
                                paragraphs = mathRange.Paragraphs;
                                if (paragraphs.Count > 0)
                                {
                                    paragraph = paragraphs[1];
                                    paragraphRange = paragraph.Range.Duplicate;
                                    File.WriteAllText(
                                        Path.Combine(
                                            artifactRoot,
                                            $"existing-omml-vt-failing-{index:D4}-paragraph.xml"),
                                        paragraphRange.WordOpenXML);
                                }
                            }
                            catch { }
                            finally
                            {
                                Release(paragraphRange);
                                Release(paragraph);
                                Release(paragraphs);
                            }
                            break;
                        }
                    }
                    finally
                    {
                        Release(mathRange);
                        Release(math);
                    }
                }

                File.WriteAllLines(
                    Path.Combine(
                        artifactRoot,
                        "existing-omml-vt-capture-diagnostic.txt"),
                    diagnosticLines);
                throw new InvalidDataException(
                    "Existing OMML→VisualTeX capture failed. "
                    + FormatScalingExceptionTree(error)
                    + (failureFound
                        ? " A per-formula reproducer was captured."
                        : " No individual formula reproduced the failure."),
                    error);
            }

            Console.WriteLine(
                $"[EXISTING OMML→VT CAPTURE] targets={plan.Targets.Count} numbered={plan.Targets.Count(item => item.Numbered)} inline={plan.Targets.Count(item => string.Equals(item.DisplayMode, "inline", StringComparison.OrdinalIgnoreCase))} display={plan.Targets.Count(item => string.Equals(item.DisplayMode, "block", StringComparison.OrdinalIgnoreCase))}");
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


    private static void RunWordExistingOmmlToVisualTeXApplyAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Environment.GetEnvironmentVariable(
            "VISUALTEX_EXISTING_OMML_VT_SOURCE_PATH");
        if (string.IsNullOrWhiteSpace(sourcePath)
            || !File.Exists(sourcePath))
            throw new FileNotFoundException(
                "Existing OMML→VisualTeX apply requires VISUALTEX_EXISTING_OMML_VT_SOURCE_PATH.",
                sourcePath);

        var localPath = Path.Combine(
            artifactRoot,
            "existing-omml-vt-apply-source-copy.docx");
        File.Copy(
            sourcePath,
            localPath,
            overwrite: true);

        var assetRoot = Path.Combine(
            artifactRoot,
            "existing-omml-vt-apply-assets");
        Directory.CreateDirectory(assetRoot);
        var pngPath = Path.Combine(
            assetRoot,
            "preview.png");
        var svgPath = Path.Combine(
            assetRoot,
            "preview.svg");
        WriteAcceptancePng(
            pngPath,
            "P(A_9,B)",
            320,
            96);
        File.WriteAllText(
            svgPath,
            @"<svg xmlns=""http://www.w3.org/2000/svg"" width=""320"" height=""96"" viewBox=""0 0 320 96""><text x=""6"" y=""66"" font-family=""Cambria Math"" font-size=""42"">P(A_9,B)</text></svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                320,
                96);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? reopened = null;
        Word.OMath? formula = null;
        Word.Range? formulaRange = null;
        try
        {
            application = CreateFreshFormatScalingWord(
                artifactRoot,
                "existing-omml-vt-apply");
            document = application.Documents.Open(
                localPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            var beforeOmml =
                document.OMaths.Count;
            var beforeVisualTeX =
                CountVisualTeXNativeOleShapes(
                    document);
            if (beforeOmml < 15)
                throw new InvalidDataException(
                    "Existing OMML→VisualTeX reproducer no longer contains formula #15.");

            formula =
                document.OMaths[15];
            formulaRange =
                formula.Range.Duplicate;
            var sourceStart =
                formulaRange.Start;
            var sourceEnd =
                formulaRange.End;
            formulaRange.Select();

            var service =
                new WordFormulaService(
                    application);
            var plan =
                service.CaptureFormulaFormatConversionPlan(
                    wholeDocument: false,
                    FormulaOleContract.WordOmmlMode,
                    FormulaOleContract.NativeOleMode);
            AssertEqual(
                1,
                plan.Targets.Count,
                "Formula #15 selection did not resolve exactly one OMML source.");
            var target =
                plan.Targets[0];
            AssertEqual(
                sourceStart,
                target.SourceStart,
                "Formula #15 capture moved to a different source range.");
            var mathMl =
                target.SourceMathMl
                ?? throw new InvalidDataException(
                    "Formula #15 exposed no MathML after the fenced-expression fix.");
            AssertTrue(
                target.Latex.IndexOf(
                    "A_{9}",
                    StringComparison.Ordinal) >= 0
                && target.Latex.IndexOf(
                    "B",
                    StringComparison.Ordinal) >= 0,
                "Formula #15 lost its ordinary multi-element fenced expression during OMML→LaTeX capture.");

            var prepared =
                new Dictionary<string, PreparedWordBulkFormula>(
                    StringComparer.Ordinal)
                {
                    [target.Id] =
                        new PreparedWordBulkFormula
                        {
                            Run =
                                new WordBulkRun
                                {
                                    Id = target.Id,
                                    IsFormula = true,
                                    Latex = target.Latex,
                                    DisplayMode =
                                        target.DisplayMode,
                                },
                            Session =
                                CreateSimpleFormatTargetSession(
                                    target,
                                    FormulaOleContract.NativeOleMode,
                                    mathMl),
                            MathMl =
                                mathMl,
                            PngPath =
                                pngPath,
                            EmfPath =
                                emfPath,
                        },
                };

            var result =
                service.ApplyFormulaFormatConversionPlan(
                    plan,
                    prepared);
            AssertEqual(
                1,
                result.FormulaCount,
                "Formula #15 OMML→VisualTeX apply did not convert exactly one formula.");
            AssertEqual(
                0,
                result.FailedFormulaCount,
                "Formula #15 OMML→VisualTeX apply reported a failure.");
            AssertEqual(
                beforeOmml - 1,
                document.OMaths.Count,
                "Formula #15 OMML→VisualTeX apply changed the wrong number of OMath objects.");
            AssertEqual(
                beforeVisualTeX + 1,
                CountVisualTeXNativeOleShapes(
                    document),
                "Formula #15 OMML→VisualTeX apply did not create exactly one VisualTeX OLE.");

            Release(formulaRange);
            formulaRange = null;
            Release(formula);
            formula = null;

            var outputPath = Path.Combine(
                artifactRoot,
                "existing-omml-vt-formula15-fixed.docx");
            document.SaveAs2(
                outputPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(
                Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            reopened =
                application.Documents.Open(
                    outputPath,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);
            document = reopened;
            reopened = null;
            AssertEqual(
                beforeOmml - 1,
                document.OMaths.Count,
                "Formula #15 OMML→VisualTeX save/reopen changed the OMath count.");
            AssertEqual(
                beforeVisualTeX + 1,
                CountVisualTeXNativeOleShapes(
                    document),
                "Formula #15 OMML→VisualTeX save/reopen lost the new VisualTeX OLE.");
            Console.WriteLine(
                $"[EXISTING OMML→VT APPLY] formula=15 range={sourceStart}:{sourceEnd} converted=1 "
                + $"sourceOMML={beforeOmml} finalOMML={document.OMaths.Count} "
                + $"sourceVT={beforeVisualTeX} finalVT={CountVisualTeXNativeOleShapes(document)} "
                + $"latex={target.Latex} output={outputPath}");
        }
        finally
        {
            Release(formulaRange);
            Release(formula);
            try { reopened?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(reopened);
            Release(document);
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
        }
    }

    private static string FormatScalingExceptionTree(
        Exception error)
    {
        var entries = new List<string>();
        static void Append(
            Exception current,
            int depth,
            List<string> output)
        {
            output.Add(
                $"depth={depth} type={current.GetType().FullName} "
                + $"hr=0x{current.HResult:X8} message={current.Message}");
            if (current is AggregateException aggregate)
            {
                foreach (var inner in aggregate.InnerExceptions)
                    Append(inner, depth + 1, output);
                return;
            }
            if (current.InnerException is not null)
                Append(current.InnerException, depth + 1, output);
        }

        Append(error, 0, entries);
        return string.Join(" | ", entries);
    }

    private static IReadOnlyDictionary<string, FormatScalingMetric>
        SnapshotFormatScalingMetrics(
            WordOperationMetrics metrics)
    {
        var result =
            new Dictionary<string, FormatScalingMetric>(
                StringComparer.Ordinal);
        foreach (var pair in metrics.Entries)
        {
            result[pair.Key] =
                new FormatScalingMetric
                {
                    Calls =
                        pair.Value.Calls,
                    Milliseconds =
                        pair.Value.Milliseconds,
                };
        }
        return result;
    }

    private static string FormatScalingMathMl(
        int index) =>
        $"<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msub><mi>x</mi><mn>{index}</mn></msub><mo>+</mo><msub><mi>y</mi><mn>{index}</mn></msub><mo>=</mo><mn>{index}</mn></math>";

    private static string FormatScalingMathMlFromLatex(
        string latex)
    {
        var match =
            System.Text.RegularExpressions.Regex.Match(
                latex ?? string.Empty,
                @"(?:x|y)_?{?(?<n>\d+)\}?");
        var index =
            match.Success
            && int.TryParse(
                match.Groups["n"].Value,
                out var parsed)
                ? parsed
                : 1;
        return FormatScalingMathMl(
            index);
    }

    private static Word.Application CreateFreshFormatScalingWord(
        string artifactRoot,
        string label)
    {
        var ownedRoot =
            Path.Combine(
                artifactRoot,
                "owned-word-"
                + label
                + "-"
                + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(
            ownedRoot);
        return CreateFreshAcceptanceAutomationWord(
            ownedRoot);
    }

    private static string ModeKey(
        string mode) =>
        string.Equals(
            mode,
            FormulaOleContract.WordOmmlMode,
            StringComparison.Ordinal)
            ? "omml"
            : string.Equals(
                mode,
                FormulaOleContract.NativeOleMode,
                StringComparison.Ordinal)
                ? "visualtex"
                : "mathtype";
}

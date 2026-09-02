using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordSimpleFormatConversionNumberingAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var olePreviewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "simple-format-conversion-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(olePreviewRoot);
        var pngPath = Path.Combine(olePreviewRoot, "simple-format-conversion.png");
        var svgPath = Path.Combine(olePreviewRoot, "simple-format-conversion.svg");
        WriteAcceptancePng(pngPath, "x+1", 260, 96);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-family=\"Cambria Math\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 260, 96);

        var formulas = new[]
        {
            (
                Latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>",
                Numbered: true),
            (
                Latex: "\\mathrm{e}^{\\mathrm{i}\\pi}+1=0",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>",
                Numbered: true),
            (
                Latex: "a^2+b^2=c^2",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>",
                Numbered: true),
        };

        var previousDefaultNumberFormat =
            WordEquationNumbering.GetDefaultEquationNumberFormatId();
        Word.Application? application = null;
        Word.Document? sourceDocument = null;
        Word.Range? sourceHostRange = null;
        Word.Document? document = null;
        var ownsApplication = false;
        var useActiveWord = false;
        try
        {
            useActiveWord = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_USE_ACTIVE_WORD"),
                "1",
                StringComparison.Ordinal);
            var includeExistingMathType = string.Equals(
                Environment.GetEnvironmentVariable("VISUALTEX_INCLUDE_EXISTING_MATHTYPE"),
                "1",
                StringComparison.Ordinal);
            if (useActiveWord)
            {
                application = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject(
                    "Word.Application");
                sourceDocument = application.ActiveDocument
                    ?? throw new InvalidOperationException("No active Word document is available as the real VisualTeX source fixture.");
                Word.InlineShape? seedShape = null;
                Word.Range? seedOwnerRange = null;
                Word.Bookmark? seedCaptionBookmark = null;
                Word.Range? seedCaptionRange = null;
                try
                {
                    for (var index = 1; index <= sourceDocument.InlineShapes.Count; index++)
                    {
                        Release(seedShape);
                        seedShape = sourceDocument.InlineShapes[index];
                        if (!WordFormulaMetadataReader.IsNativeOle(seedShape)) continue;
                        var metadata = WordFormulaMetadataReader.TryRead(seedShape);
                        if (metadata?.Numbered != true
                            || !string.Equals(metadata.DisplayMode, "block", StringComparison.OrdinalIgnoreCase))
                            continue;
                        Release(seedOwnerRange);
                        seedOwnerRange = WordEquationNumbering.FindNumberingOwnerRange(
                            sourceDocument,
                            metadata.FormulaId);
                        if (seedOwnerRange is null) continue;
                        var hostStart = seedOwnerRange.Start;
                        var hostEnd = seedOwnerRange.End;
                        var captionName = WordEquationNumbering.NativeCaptionBookmarkName(metadata.FormulaId);
                        if (sourceDocument.Bookmarks.Exists(captionName))
                        {
                            seedCaptionBookmark = sourceDocument.Bookmarks[captionName];
                            seedCaptionRange = seedCaptionBookmark.Range;
                            hostEnd = Math.Max(hostEnd, seedCaptionRange.End);
                        }
                        sourceHostRange = sourceDocument.Range(hostStart, hostEnd);
                        break;
                    }
                    if (sourceHostRange is null)
                        throw new InvalidOperationException(
                            "The active Word document does not contain a complete numbered VisualTeX display formula to clone.");
                }
                finally
                {
                    Release(seedCaptionRange);
                    Release(seedCaptionBookmark);
                    Release(seedOwnerRange);
                    Release(seedShape);
                }
            }
            else
            {
                application = CreateWordApplication(visible: false);
                ownsApplication = true;
            }
            document = application.Documents.Add();
            document.Content.Text = "VisualTeX simple format conversion numbering acceptance\r";
            if (useActiveWord && sourceDocument is not null)
            {
                var sourceNumberFormat = WordEquationNumbering.GetEquationNumberFormatId(sourceDocument);
                WordEquationNumbering.SetEquationNumberFormat(document, sourceNumberFormat);
                Console.WriteLine($"[CONSECUTIVE NUMBERED] inherited equation-number format '{sourceNumberFormat}'.");
            }
            else
            {
                // Never let this acceptance depend on the current user's registry
                // default. Its expected 0.1/0.2/0.3 captions deliberately exercise
                // heading-aware conversion even on a clean or differently configured
                // machine.
                WordEquationNumbering.SetEquationNumberFormat(
                    document,
                    EquationNumberFormat.Heading1DotId);
            }
            var service = new WordFormulaService(application);
            RunSimpleFormatConversionRollbackBridgeAcceptance(
                application,
                service,
                pngPath,
                emfPath,
                artifactRoot);
            if (includeExistingMathType)
            {
                var existingInsertion = document.Range(
                    document.Content.End - 1,
                    document.Content.End - 1);
                try { existingInsertion.Select(); }
                finally { Release(existingInsertion); }
                var existingSession = CreateSimpleVisualTeXSourceSession("y=x", numbered: true);
                existingSession.ObjectMode = FormulaOleContract.MathTypeOleMode;
                service.InsertMathTypeOle(
                    existingSession,
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>y</mi><mo>=</mo><mi>x</mi></math>",
                    emfPath);
                Console.WriteLine(
                    "[CONSECUTIVE NUMBERED] Added one existing numbered MathType formula before the three VisualTeX sources.");
            }

            if (useActiveWord)
            {
                for (var copyIndex = 0; copyIndex < 3; copyIndex++)
                {
                    var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                    Word.InlineShape? pastedShape = null;
                    try
                    {
                        sourceHostRange!.Copy();
                        insertion.Paste();
                        System.Windows.Forms.Application.DoEvents();
                        Thread.Sleep(120);
                        if (document.InlineShapes.Count <= copyIndex)
                            throw new InvalidOperationException(
                                $"Copy {copyIndex + 1} did not materialize a VisualTeX OLE object.");
                        pastedShape = document.InlineShapes[document.InlineShapes.Count];
                        pastedShape.Range.Select();
                        _ = service.ReadSelection();
                        System.Windows.Forms.Application.DoEvents();
                        Thread.Sleep(120);

                        Release(pastedShape);
                        pastedShape = document.InlineShapes[document.InlineShapes.Count];
                        var desiredFormula = formulas[copyIndex];
                        var copiedMetadata = WordFormulaMetadataReader.TryRead(pastedShape)
                            ?? throw new InvalidDataException(
                                $"Copy {copyIndex + 1} lost its VisualTeX metadata after identity repair.");
                        copiedMetadata.Latex = desiredFormula.Latex;
                        copiedMetadata.Lines = new List<FormulaLine>
                        {
                            new()
                            {
                                Id = copiedMetadata.Lines.FirstOrDefault()?.Id
                                    ?? Guid.NewGuid().ToString("D"),
                                Latex = desiredFormula.Latex,
                            },
                        };
                        WordFormulaMetadataReader.CacheMetadata(pastedShape, copiedMetadata);
                    }
                    finally
                    {
                        Release(pastedShape);
                        Release(insertion);
                    }
                }
            }
            else
            {
                foreach (var formula in formulas)
                {
                    var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                    try { insertion.Select(); }
                    finally { Release(insertion); }

                    service.InsertOle(
                        CreateSimpleVisualTeXSourceSession(formula.Latex, formula.Numbered),
                        pngPath,
                        emfPath);
                }
            }

            var expectedTotalFormulaObjectsBeforeConversion = includeExistingMathType ? 4 : 3;
            var expectedMathTypeObjectsAfterConversion = includeExistingMathType ? 4 : 3;
            AssertEqual(expectedTotalFormulaObjectsBeforeConversion, document.InlineShapes.Count,
                "Consecutive-numbered conversion setup created the wrong total formula-object count.");
            AssertEqual(3, CountVisualTeXNativeOleShapes(document),
                "Consecutive-numbered conversion setup did not create exactly three VisualTeX source formulas.");
            AssertEqual(3, CountVisualTeXNumberingBookmarkTriples(document),
                "Consecutive-numbered conversion setup did not create three numbered VisualTeX hosts.");

            // Fresh VisualTeX OLE formulas now follow MathType's native Word
            // geometry. Verify every source before format conversion so this
            // acceptance cannot silently regress to the legacy 1x3 table host.
            for (var sourceIndex = 1; sourceIndex <= document.InlineShapes.Count; sourceIndex++)
            {
                Word.InlineShape? sourceShape = null;
                try
                {
                    sourceShape = document.InlineShapes[sourceIndex];
                    if (!WordFormulaMetadataReader.IsNativeOle(sourceShape)) continue;
                    var sourceMetadata = WordFormulaMetadataReader.TryRead(sourceShape)
                        ?? throw new InvalidDataException(
                            $"VisualTeX source {sourceIndex} lost metadata before format conversion.");
                    AssertVisualTeXNumberedTabHost(
                        document,
                        sourceMetadata.FormulaId,
                        updateReference: true,
                        context: $"format-conversion source {sourceIndex}");
                }
                finally { Release(sourceShape); }
            }

            var sourceFixturePath = Path.Combine(
                artifactRoot,
                "VisualTeX-Three-Numbered-Heading1-Source.docx");
            document.SaveAs2(sourceFixturePath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine($"[CONSECUTIVE NUMBERED SOURCE] {sourceFixturePath}");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(3, plan.Targets.Count,
                "Consecutive-numbered conversion capture did not find three VisualTeX formulas.");
            for (var index = 0; index < formulas.Length; index++)
            {
                AssertEqual(formulas[index].Numbered, plan.Targets[index].Numbered,
                    $"Simple conversion capture lost Numbered state at formula {index + 1}.");
            }

            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            for (var index = 0; index < plan.Targets.Count; index++)
            {
                var target = plan.Targets[index];
                var source = formulas[index];
                prepared[target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleMathTypeTargetSession(target, source.MathMl),
                    MathMl = source.MathMl,
                    EmfPath = emfPath,
                };
            }

            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            Console.WriteLine(
                $"[CONSECUTIVE NUMBERED RESULT] converted={result.FormulaCount} failed={result.FailedFormulaCount} failures={string.Join(" | ", result.Failures)}");
            AssertEqual(3, result.FormulaCount,
                "Consecutive-numbered VisualTeX -> MathType conversion did not replace three formulas.");
            AssertEqual(0, result.FailedFormulaCount,
                $"Consecutive-numbered VisualTeX -> MathType conversion failed: {string.Join(" | ", result.Failures)}");
            AssertEqual(expectedMathTypeObjectsAfterConversion, document.InlineShapes.Count,
                "Consecutive-numbered VisualTeX -> MathType conversion changed the total formula-object count.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                "Simple VisualTeX -> MathType conversion left VisualTeX OLE objects behind.");
            AssertEqual(expectedMathTypeObjectsAfterConversion, CountMathTypeOleShapes(document),
                "Consecutive-numbered VisualTeX -> MathType conversion produced the wrong Equation.DSMT4 object count.");
            AssertEqual(0, CountVisualTeXNumberingBookmarks(document),
                "Simple VisualTeX -> MathType conversion left old VTEq/VTEqCap/VTEqNum bookmarks behind.");
            AssertEqual(expectedMathTypeObjectsAfterConversion, CountMathTypePlaceRefFields(document),
                "Consecutive-numbered VisualTeX -> MathType conversion produced the wrong MTPlaceRef field count.");
            if (includeExistingMathType)
                AssertMathTypeNumberTexts(document, "(0.1)", "(0.2)", "(0.3)", "(0.4)");
            else
                AssertMathTypeNumberTexts(document, "(0.1)", "(0.2)", "(0.3)");
            AssertNativeMathTypeSectionBreak(document, 0);

            for (var index = 1; index <= document.InlineShapes.Count; index++)
            {
                Word.InlineShape? shape = null;
                try
                {
                    shape = document.InlineShapes[index];
                    AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                        $"Converted formula {index} is not Equation.DSMT4.");
                    var metadata = MathTypeOleInterop.ReadMetadata(application, shape);
                    AssertEqual(true, metadata.Numbered,
                        $"MathType formula {index} has the wrong numbered state.");
                }
                finally { Release(shape); }
            }

            var outputPath = Path.Combine(
                artifactRoot,
                "VisualTeX-Simple-Format-Conversion-Numbering.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "[CONSECUTIVE NUMBERED] Three consecutive numbered VisualTeX display formulas converted to three MathType Equation.DSMT4 objects with three fresh MTPlaceRef fields and no VisualTeX numbering bookmarks.");

            RunOmmlVisualTeXNumberedRoundTripAcceptance(
                application,
                pngPath,
                emfPath,
                formulas,
                artifactRoot);
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            Release(sourceHostRange);
            Release(sourceDocument);
            if (ownsApplication)
            {
                try { QuitWordApplicationIfOwned(application); } catch { }
            }
            Release(application);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(
                previousDefaultNumberFormat);
            ForceComCleanup();
            try { Directory.Delete(olePreviewRoot, recursive: true); } catch { }
        }
    }

    private static void RunWordOmmlVisualTeXNumberedRoundTripAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var olePreviewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "omml-visualtex-numbered-roundtrip-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(olePreviewRoot);
        var pngPath = Path.Combine(olePreviewRoot, "omml-visualtex-numbered-roundtrip.png");
        var svgPath = Path.Combine(olePreviewRoot, "omml-visualtex-numbered-roundtrip.svg");
        WriteAcceptancePng(pngPath, "x+1", 260, 96);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"96\" viewBox=\"0 0 260 96\"><text x=\"6\" y=\"66\" font-family=\"Cambria Math\" font-size=\"48\">x+1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 260, 96);
        var formulas = new[]
        {
            (
                Latex: "x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>",
                Numbered: true),
            (
                Latex: "\\mathrm{e}^{\\mathrm{i}\\pi}+1=0",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>",
                Numbered: true),
            (
                Latex: "a^2+b^2=c^2",
                MathMl: "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msup><mi>a</mi><mn>2</mn></msup><mo>+</mo><msup><mi>b</mi><mn>2</mn></msup><mo>=</mo><msup><mi>c</mi><mn>2</mn></msup></math>",
                Numbered: true),
        };
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            RunOmmlVisualTeXNumberedRoundTripAcceptance(
                application,
                pngPath,
                emfPath,
                formulas,
                artifactRoot);
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(olePreviewRoot, recursive: true); } catch { }
        }
    }

    private static void RunOmmlVisualTeXNumberedRoundTripAcceptance(
        Word.Application application,
        string pngPath,
        string emfPath,
        IReadOnlyList<(string Latex, string MathMl, bool Numbered)> formulas,
        string artifactRoot)
    {
        Word.Document? document = null;
        try
        {
            document = application.Documents.Add();
            document.Content.Text = "OMML to VisualTeX numbered roundtrip acceptance\r";
            var service = new WordFormulaService(application);
            foreach (var formula in formulas)
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
                service.InsertOle(
                    CreateSimpleVisualTeXSourceSession(formula.Latex, numbered: true),
                    pngPath,
                    emfPath);
            }

            AssertEqual(formulas.Count, CountVisualTeXNumberingBookmarkTriples(document),
                "OMML→VisualTeX roundtrip setup did not create one numbered VisualTeX host per formula.");

            var toOmmlPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(formulas.Count, toOmmlPlan.Targets.Count,
                "OMML→VisualTeX roundtrip setup did not capture every VisualTeX source formula.");
            var toOmmlPrepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            for (var index = 0; index < toOmmlPlan.Targets.Count; index++)
            {
                var target = toOmmlPlan.Targets[index];
                var source = formulas[index];
                toOmmlPrepared[target.Id] = new PreparedWordBulkFormula
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
                        source.MathMl),
                    MathMl = source.MathMl,
                };
            }
            var toOmmlResult = service.ApplyFormulaFormatConversionPlan(
                toOmmlPlan,
                toOmmlPrepared);
            AssertEqual(formulas.Count, toOmmlResult.FormulaCount,
                "VisualTeX→OMML roundtrip leg did not convert every numbered formula.");
            AssertEqual(0, toOmmlResult.FailedFormulaCount,
                $"VisualTeX→OMML roundtrip leg failed: {string.Join(" | ", toOmmlResult.Failures)}");
            AssertEqual(formulas.Count, document.OMaths.Count,
                "VisualTeX→OMML roundtrip leg did not leave one OMath per formula.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                "VisualTeX→OMML roundtrip leg left VisualTeX OLE sources behind.");
            AssertEqual(formulas.Count, CountVisualTeXNumberingBookmarkTriples(document),
                "VisualTeX→OMML roundtrip leg did not preserve every numbered host.");
            var baselineOmmlBlankParagraphs =
                CountPureBlankParagraphsImmediatelyBeforeTables(document);

            var toVisualTeXPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(formulas.Count, toVisualTeXPlan.Targets.Count,
                "OMML→VisualTeX roundtrip leg did not capture every numbered OMML formula.");
            AssertTrue(toVisualTeXPlan.Targets.All(target => target.Numbered),
                "OMML→VisualTeX roundtrip capture lost numbered state before conversion.");

            var toVisualTeXPrepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            foreach (var target in toVisualTeXPlan.Targets)
            {
                var mathMl = target.SourceMathMl
                    ?? throw new InvalidDataException(
                        $"OMML target {target.Id} did not expose source MathML.");
                toVisualTeXPrepared[target.Id] = new PreparedWordBulkFormula
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
            var toVisualTeXResult = service.ApplyFormulaFormatConversionPlan(
                toVisualTeXPlan,
                toVisualTeXPrepared);
            Console.WriteLine(
                $"[OMML→VISUALTEX NUMBERED RESULT] converted={toVisualTeXResult.FormulaCount} failed={toVisualTeXResult.FailedFormulaCount} triples={CountVisualTeXNumberingBookmarkTriples(document)}");
            AssertEqual(formulas.Count, toVisualTeXResult.FormulaCount,
                "OMML→VisualTeX roundtrip leg did not convert every numbered formula.");
            AssertEqual(0, toVisualTeXResult.FailedFormulaCount,
                $"OMML→VisualTeX roundtrip leg failed: {string.Join(" | ", toVisualTeXResult.Failures)}");
            AssertEqual(0, document.OMaths.Count,
                "OMML→VisualTeX roundtrip leg left OMML sources behind.");
            AssertEqual(formulas.Count, CountVisualTeXNativeOleShapes(document),
                "OMML→VisualTeX roundtrip leg did not create one VisualTeX OLE per source formula.");
            AssertEqual(formulas.Count, CountInstalledVisualTeXNumberedFormulaHosts(document),
                "OMML→VisualTeX roundtrip leg lost Numbered metadata on a converted formula.");
            AssertEqual(formulas.Count, CountVisualTeXNumberingBookmarkTriples(document),
                "OMML→VisualTeX roundtrip leg lost a VTEq/VTEqCap/VTEqNum numbering host.");

            var backToOmmlPlan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.WordOmmlMode);
            AssertEqual(formulas.Count, backToOmmlPlan.Targets.Count,
                "OMML→VisualTeX→OMML roundtrip did not recapture every VisualTeX formula.");
            var backToOmmlPrepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
            for (var index = 0; index < backToOmmlPlan.Targets.Count; index++)
            {
                var target = backToOmmlPlan.Targets[index];
                var source = formulas[index];
                backToOmmlPrepared[target.Id] = new PreparedWordBulkFormula
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
                        source.MathMl),
                    MathMl = source.MathMl,
                };
            }
            var backToOmmlResult = service.ApplyFormulaFormatConversionPlan(
                backToOmmlPlan,
                backToOmmlPrepared);
            AssertEqual(formulas.Count, backToOmmlResult.FormulaCount,
                "OMML→VisualTeX→OMML roundtrip did not convert every formula back to OMML.");
            AssertEqual(0, backToOmmlResult.FailedFormulaCount,
                $"OMML→VisualTeX→OMML roundtrip failed: {string.Join(" | ", backToOmmlResult.Failures)}");
            AssertEqual(formulas.Count, document.OMaths.Count,
                "OMML→VisualTeX→OMML roundtrip did not leave one OMath per formula.");
            AssertEqual(formulas.Count, CountVisualTeXNumberingBookmarkTriples(document),
                "OMML→VisualTeX→OMML roundtrip lost a numbering host.");
            AssertEqual(
                baselineOmmlBlankParagraphs,
                CountPureBlankParagraphsImmediatelyBeforeTables(document),
                "OMML→VisualTeX→OMML roundtrip introduced an extra plain blank paragraph before a numbered OMML table.");

            var outputPath = Path.Combine(
                artifactRoot,
                "OMML-To-VisualTeX-Numbered-Roundtrip.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"[OMML→VISUALTEX NUMBERED] Preserved {formulas.Count} independent numbered VisualTeX hosts. path={outputPath}");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
        }
    }

    private static void RunWordSingleNumberedOmmlToVisualTeXAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        // VisualTeX.Formula.1 intentionally refuses preview files outside the
        // product-owned Office temp root. Put this acceptance's PNG/SVG/EMF in
        // the same root used by production Session exports; the DOCX artifact
        // itself still belongs under artifactRoot.
        var olePreviewRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            "single-numbered-omml-to-visualtex-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(olePreviewRoot);
        var pngPath = Path.Combine(olePreviewRoot, "single-numbered-omml-to-visualtex.png");
        var svgPath = Path.Combine(olePreviewRoot, "single-numbered-omml-to-visualtex.svg");
        WriteAcceptancePng(pngPath, "quadratic", 360, 112);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><text x=\"8\" y=\"76\" font-family=\"Cambria Math\" font-size=\"44\">quadratic</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 112);

        const string sourceLatex = @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}";
        const string sourceMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
            + "<msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt>"
            + "</mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";
        const string secondLatex = @"\mathrm{e}^{\mathrm{i}\pi}+1=0";
        const string secondMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
            + "<msup><mi mathvariant=\"normal\">e</mi><mrow><mi mathvariant=\"normal\">i</mi><mi>π</mi></mrow></msup>"
            + "<mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn></math>";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Selection? selection = null;
        Word.Range? ownerRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? captionBookmark = null;
        Word.Bookmark? firstNumberBookmark = null;
        Word.Bookmark? secondNumberBookmark = null;
        Word.Range? captionRange = null;
        Word.Range? firstNumberRange = null;
        Word.Range? secondNumberRange = null;
        Word.Frames? captionFrames = null;
        Word.Fields? captionFields = null;
        Word.Field? captionField = null;
        Word.Range? captionCode = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add();
            var service = new WordFormulaService(application);
            WordEquationNumbering.SetEquationNumberFormat(document, "continuous");

            var sourceFormulaId = Guid.NewGuid().ToString("D");
            var sourceSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "create",
                Host = "word",
                FormulaId = sourceFormulaId,
                Title = "Single numbered OMML conversion source",
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = sourceLatex },
                },
                CodeFormat = "latex",
                DisplayMode = "block",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = true,
                FontSizePt = 12,
                ExportResult = new OfficeExportDocument
                {
                    MathMl = sourceMathMl,
                    Width = 360,
                    Height = 112,
                    Baseline = 84,
                },
            };
            selection = application.Selection;
            selection.SetRange(document.Content.Start, document.Content.Start);
            service.InsertOmml(sourceSession, sourceMathMl);
            Release(selection);
            selection = null;

            AssertEqual(1, document.OMaths.Count,
                "Single numbered OMML setup did not create exactly one Word equation.");
            AssertEqual(0, CountVisualTeXNativeOleShapes(document),
                "Single numbered OMML setup unexpectedly created a VisualTeX OLE.");
            AssertEqual(1, CountVisualTeXNumberingBookmarkTriples(document),
                "Single numbered OMML setup did not create one numbering identity triplet.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Single numbered OMML conversion did not capture exactly one target.");
            var target = plan.Targets[0];
            AssertTrue(target.Numbered,
                "Single numbered OMML conversion lost numbered state during capture.");
            AssertTrue(target.SourceIsManagedOmml,
                "Single numbered OMML conversion did not recognize the managed OMML source.");
            var targetMathMl = target.SourceMathMl ?? sourceMathMl;
            var targetSession = CreateSimpleFormatTargetSession(
                target,
                FormulaOleContract.NativeOleMode,
                targetMathMl);
            targetSession.ExportResult!.Width = 360;
            targetSession.ExportResult.Height = 112;
            targetSession.ExportResult.Baseline = 84;
            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal)
            {
                [target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = targetSession,
                    MathMl = targetMathMl,
                    PngPath = pngPath,
                    EmfPath = emfPath,
                },
            };

            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            AssertEqual(1, result.FormulaCount,
                "Single numbered OMML→VisualTeX conversion did not commit its target.");
            AssertEqual(0, result.FailedFormulaCount,
                $"Single numbered OMML→VisualTeX conversion failed: {string.Join(" | ", result.Failures)}");
            AssertEqual(0, document.OMaths.Count,
                "Single numbered OMML→VisualTeX conversion left the source OMath behind.");
            AssertEqual(1, CountVisualTeXNativeOleShapes(document),
                "Single numbered OMML→VisualTeX conversion did not leave exactly one VisualTeX OLE.");
            AssertEqual(0, document.Tables.Count,
                "Single numbered OMML→VisualTeX conversion left a numbering table behind.");
            AssertEqual(1, CountVisualTeXNumberingBookmarkTriples(document),
                "Single numbered OMML→VisualTeX conversion did not converge to one numbering triplet.");
            AssertTrue(!document.Bookmarks.Exists(WordOmmlFormulaStore.BookmarkName(sourceFormulaId)),
                "Single numbered OMML→VisualTeX conversion left the source VTOMML identity behind.");

            var convertedFormulaId = targetSession.FormulaId;
            AssertVisualTeXNumberedTabHost(
                document,
                convertedFormulaId,
                updateReference: true,
                context: "single numbered OMML→VisualTeX");

            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, convertedFormulaId)
                ?? throw new InvalidDataException(
                    "Single numbered OMML→VisualTeX conversion lost its owner paragraph after validation.");
            bookmarks = document.Bookmarks;
            var captionName = WordEquationNumbering.NativeCaptionBookmarkName(convertedFormulaId);
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(convertedFormulaId);
            AssertTrue(bookmarks.Exists(captionName) && bookmarks.Exists(numberName),
                "Single numbered OMML→VisualTeX conversion lost its hidden caption identities.");
            captionBookmark = bookmarks[captionName];
            firstNumberBookmark = bookmarks[numberName];
            captionRange = captionBookmark.Range;
            firstNumberRange = firstNumberBookmark.Range;
            AssertTrue(captionRange.Start >= ownerRange.End,
                "Single numbered OMML→VisualTeX hidden SEQ caption overlaps the OLE owner paragraph.");
            AssertTrue(firstNumberRange.Start >= captionRange.Start
                && firstNumberRange.End <= captionRange.End,
                "Single numbered OMML→VisualTeX VTEqNum bookmark escapes its hidden caption.");
            captionFrames = captionRange.Frames;
            AssertEqual(1, captionFrames.Count,
                "Single numbered OMML→VisualTeX hidden SEQ caption is not isolated in exactly one clipping frame.");
            captionFields = captionRange.Fields;
            var sequenceCount = 0;
            for (var index = 1; index <= captionFields.Count; index++)
            {
                Release(captionCode);
                captionCode = null;
                Release(captionField);
                captionField = captionFields[index];
                captionCode = captionField.Code;
                if ((captionCode.Text ?? string.Empty).IndexOf(
                        "SEQ VisualTeXEquation",
                        StringComparison.OrdinalIgnoreCase) >= 0)
                    sequenceCount++;
            }
            AssertEqual(1, sequenceCount,
                "Single numbered OMML→VisualTeX hidden caption does not own exactly one VisualTeX SEQ field.");
            AssertEqual("1", (firstNumberRange.Text ?? string.Empty).Trim(),
                "Single numbered OMML→VisualTeX target did not retain equation number 1.");

            // Reproduce the user's strongest orphan check: inserting a second
            // numbered OMML must legitimately become equation 2 because the first
            // converted VisualTeX OLE owns one healthy SEQ, not because an invisible
            // abandoned source scaffold is still incrementing the sequence.
            var secondFormulaId = Guid.NewGuid().ToString("D");
            var secondSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "create",
                Host = "word",
                FormulaId = secondFormulaId,
                Title = "Second numbered OMML after conversion",
                Lines = new List<FormulaLine>
                {
                    new() { Id = Guid.NewGuid().ToString("D"), Latex = secondLatex },
                },
                CodeFormat = "latex",
                DisplayMode = "block",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = true,
                FontSizePt = 12,
                ExportResult = new OfficeExportDocument
                {
                    MathMl = secondMathMl,
                    Width = 260,
                    Height = 96,
                    Baseline = 72,
                },
            };
            selection = application.Selection;
            var insertion = Math.Max(document.Content.Start, document.Content.End - 1);
            selection.SetRange(insertion, insertion);
            service.InsertOmml(secondSession, secondMathMl);
            Release(selection);
            selection = null;
            AssertTrue(WordEquationNumbering.UpdateEquationNumbers(document) >= 2,
                "Updating numbers after inserting the second OMML did not see both legitimate numbered formulas.");

            Release(bookmarks);
            bookmarks = document.Bookmarks;
            Release(firstNumberBookmark);
            firstNumberBookmark = bookmarks[numberName];
            Release(firstNumberRange);
            firstNumberRange = firstNumberBookmark.Range;
            var secondNumberName = WordEquationNumbering.NativeNumberBookmarkName(secondFormulaId);
            AssertTrue(bookmarks.Exists(secondNumberName),
                "Second numbered OMML lost its VTEqNum identity.");
            secondNumberBookmark = bookmarks[secondNumberName];
            secondNumberRange = secondNumberBookmark.Range;
            AssertEqual("1", (firstNumberRange.Text ?? string.Empty).Trim(),
                "First converted VisualTeX number changed after inserting a second OMML.");
            AssertEqual("2", (secondNumberRange.Text ?? string.Empty).Trim(),
                "Second numbered OMML did not receive equation number 2.");
            AssertEqual(2, CountVisualTeXNumberingBookmarkTriples(document),
                "Inserting the second numbered OMML revealed an orphan or missing numbering triplet.");

            var outputPath = Path.Combine(
                artifactRoot,
                "Single-Numbered-OMML-To-VisualTeX.docx");
            document.SaveAs2(outputPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = application.Documents.Open(outputPath, ReadOnly: false, Visible: false);
            AssertEqual(1, document.OMaths.Count,
                "Save/reopen changed the post-conversion OMML inventory.");
            AssertEqual(1, CountVisualTeXNativeOleShapes(document),
                "Save/reopen changed the converted VisualTeX OLE inventory.");
            AssertEqual(2, CountVisualTeXNumberingBookmarkTriples(document),
                "Save/reopen changed the valid numbering identities.");
            AssertVisualTeXNumberedTabHost(
                document,
                convertedFormulaId,
                updateReference: true,
                context: "reopened single numbered OMML→VisualTeX");
            Console.WriteLine(
                $"[SINGLE NUMBERED OMML→VISUALTEX] Converted one managed numbered OMML to one healthy table-free VisualTeX OLE; a subsequently inserted numbered OMML became 2 with no orphan sequence, and save/reopen preserved both identities. path={outputPath}");
        }
        finally
        {
            Release(captionCode);
            Release(captionField);
            Release(captionFields);
            Release(captionFrames);
            Release(secondNumberRange);
            Release(secondNumberBookmark);
            Release(firstNumberRange);
            Release(firstNumberBookmark);
            Release(captionRange);
            Release(captionBookmark);
            Release(bookmarks);
            Release(ownerRange);
            Release(selection);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
            try { Directory.Delete(olePreviewRoot, recursive: true); } catch { }
        }
    }

    private static void RunSimpleFormatConversionRollbackBridgeAcceptance(
        Word.Application application,
        WordFormulaService service,
        string pngPath,
        string emfPath,
        string artifactRoot)
    {
        const string latex = @"(a+b)^{n}=\sum_{k=0}^{n}\left( \begin{matrix}n \\ k\end{matrix}\right) a^{n-k}b^{k}";
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mo stretchy=\"false\">(</mo><mi>a</mi><mo>+</mo><mi>b</mi>"
            + "<msup><mo stretchy=\"false\">)</mo><mi>n</mi></msup><mo>=</mo>"
            + "<munderover><mo data-mjx-texclass=\"OP\">∑</mo><mrow><mi>k</mi><mo>=</mo><mn>0</mn></mrow><mi>n</mi></munderover>"
            + "<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\">(</mo>"
            + "<mtable><mtr><mtd><mi>n</mi></mtd></mtr><mtr><mtd><mi>k</mi></mtd></mtr></mtable>"
            + "<mo data-mjx-texclass=\"CLOSE\">)</mo></mrow>"
            + "<msup><mi>a</mi><mrow><mi>n</mi><mo>−</mo><mi>k</mi></mrow></msup>"
            + "<msup><mi>b</mi><mi>k</mi></msup></math>";

        Word.Document? rollbackDocument = null;
        Word.Range? insertion = null;
        var oldAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE");
        var oldFailure = Environment.GetEnvironmentVariable(
            "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE");
        try
        {
            rollbackDocument = application.Documents.Add();
            rollbackDocument.Content.Text = "before rollback\r";
            insertion = rollbackDocument.Range(
                rollbackDocument.Content.End - 1,
                rollbackDocument.Content.End - 1);
            insertion.Select();
            Release(insertion);
            insertion = null;
            service.InsertOle(
                CreateSimpleVisualTeXSourceSession(latex, numbered: false),
                pngPath,
                emfPath);
            AssertEqual(1, CountVisualTeXNativeOleShapes(rollbackDocument),
                "Rollback bridge setup did not create exactly one VisualTeX source formula.");

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.NativeOleMode,
                FormulaOleContract.MathTypeOleMode);
            AssertEqual(1, plan.Targets.Count,
                "Rollback bridge setup did not capture exactly one VisualTeX source formula.");
            var target = plan.Targets[0];
            var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal)
            {
                [target.Id] = new PreparedWordBulkFormula
                {
                    Run = new WordBulkRun
                    {
                        Id = target.Id,
                        IsFormula = true,
                        Latex = target.Latex,
                        DisplayMode = target.DisplayMode,
                    },
                    Session = CreateSimpleMathTypeTargetSession(target, mathMl),
                    MathMl = mathMl,
                    EmfPath = emfPath,
                },
            };

            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", "1");
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                target.SourceFormulaId);
            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            AssertEqual(0, result.FormulaCount,
                "Injected rollback conversion unexpectedly reported a converted formula.");
            AssertEqual(1, result.FailedFormulaCount,
                "Injected rollback conversion did not report exactly one failure.");
            AssertEqual(1, CountVisualTeXNativeOleShapes(rollbackDocument),
                "Injected rollback did not restore exactly one VisualTeX source formula.");
            AssertEqual(0, CountMathTypeOleShapes(rollbackDocument),
                "Injected rollback left a MathType target behind.");
            var bridge = "$" + latex + "$";
            AssertTrue(
                (rollbackDocument.Content.Text ?? string.Empty).IndexOf(
                    bridge,
                    StringComparison.Ordinal) < 0,
                "Injected rollback restored the source formula but left the temporary LaTeX bridge in Word text.");

            var path = Path.Combine(
                artifactRoot,
                "VisualTeX-Format-Conversion-Rollback-No-Bridge.docx");
            rollbackDocument.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                "[FORMAT CONVERSION ROLLBACK] Injected post-delete failure restored the VisualTeX OLE and left no temporary LaTeX bridge.");
        }
        finally
        {
            Environment.SetEnvironmentVariable("VISUALTEX_VSTO_ACCEPTANCE", oldAcceptance);
            Environment.SetEnvironmentVariable(
                "VISUALTEX_VSTO_FORMAT_CONVERSION_FAIL_AFTER_DELETE",
                oldFailure);
            Release(insertion);
            if (rollbackDocument is not null)
            {
                try { rollbackDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(rollbackDocument);
        }
    }

    private static OfficeSessionDocument CreateSimpleVisualTeXSourceSession(
        string latex,
        bool numbered)
    {
        var formulaId = Guid.NewGuid().ToString("D");
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = formulaId,
            Title = "Simple format conversion source",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = latex },
            },
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.NativeOleMode,
            Numbered = numbered,
            FontSizePt = 12,
            ExportResult = new OfficeExportDocument
            {
                Width = 260,
                Height = 96,
                Baseline = 72,
            },
        };
    }

    private static OfficeSessionDocument CreateSimpleMathTypeTargetSession(
        WordFormulaFormatConversionTarget target,
        string mathMl) =>
        CreateSimpleFormatTargetSession(
            target,
            FormulaOleContract.MathTypeOleMode,
            mathMl);

    private static OfficeSessionDocument CreateSimpleFormatTargetSession(
        WordFormulaFormatConversionTarget target,
        string objectMode,
        string mathMl)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = Guid.NewGuid().ToString("D"),
            Title = "Simple format conversion target",
            Lines = new List<FormulaLine>
            {
                new() { Id = Guid.NewGuid().ToString("D"), Latex = target.Latex },
            },
            CodeFormat = "latex",
            DisplayMode = target.DisplayMode,
            ObjectMode = objectMode,
            Numbered = target.Numbered,
            MathTypeNumberPosition = target.MathTypeNumberPosition,
            FontSizePt = target.FontSizePt,
            OriginalMetadata = target.Metadata,
            ExportResult = new OfficeExportDocument
            {
                MathMl = mathMl,
                Width = 260,
                Height = 96,
                Baseline = 72,
            },
        };
    }

    private static int CountVisualTeXNativeOleShapes(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.InlineShapes.Count; index++)
        {
            Word.InlineShape? shape = null;
            try
            {
                shape = document.InlineShapes[index];
                if (WordFormulaMetadataReader.IsNativeOle(shape)) count++;
            }
            finally { Release(shape); }
        }
        return count;
    }

    private static int CountMathTypeOleShapes(Word.Document document)
    {
        const int RpcCallRejected = unchecked((int)0x80010001);
        for (var attempt = 0; attempt < 80; attempt++)
        {
            Word.InlineShapes? shapes = null;
            try
            {
                shapes = document.InlineShapes;
                var count = 0;
                for (var index = 1; index <= shapes.Count; index++)
                {
                    Word.InlineShape? shape = null;
                    try
                    {
                        shape = shapes[index];
                        if (MathTypeOleInterop.IsMathTypeOle(shape)) count++;
                    }
                    finally { Release(shape); }
                }
                return count;
            }
            catch (System.Runtime.InteropServices.COMException error)
                when (error.ErrorCode == RpcCallRejected && attempt < 79)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(100);
            }
            finally { Release(shapes); }
        }
        throw new TimeoutException(
            "Word remained busy while the installed acceptance counted MathType OLE objects.");
    }

    private static int CountVisualTeXNumberingBookmarks(Word.Document document)
    {
        var count = 0;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? bookmark = null;
        try
        {
            bookmarks = document.Bookmarks;
            for (var index = 1; index <= bookmarks.Count; index++)
            {
                Release(bookmark);
                bookmark = bookmarks[index];
                var name = bookmark.Name;
                if (name.StartsWith("VTEq_", StringComparison.Ordinal)
                    || name.StartsWith("VTEqCap_", StringComparison.Ordinal)
                    || name.StartsWith("VTEqNum_", StringComparison.Ordinal))
                    count++;
            }
            return count;
        }
        finally
        {
            Release(bookmark);
            Release(bookmarks);
        }
    }

    private static int CountPureBlankParagraphsImmediatelyBeforeTables(Word.Document document)
    {
        var count = 0;
        for (var index = 1; index <= document.Tables.Count; index++)
        {
            Word.Table? table = null;
            Word.Range? probe = null;
            Word.Paragraphs? paragraphs = null;
            Word.Paragraph? paragraph = null;
            Word.Range? paragraphRange = null;
            try
            {
                table = document.Tables[index];
                var tableStart = table.Range.Start;
                if (tableStart <= document.Content.Start) continue;
                probe = document.Range(tableStart - 1, tableStart);
                if ((bool)probe.get_Information(Word.WdInformation.wdWithInTable)) continue;
                paragraphs = probe.Paragraphs;
                if (paragraphs.Count != 1) continue;
                paragraph = paragraphs[1];
                paragraphRange = paragraph.Range;
                if (paragraphRange.End != tableStart
                    || !string.Equals(paragraphRange.Text, "\r", StringComparison.Ordinal))
                    continue;
                if (paragraphRange.Tables.Count == 0
                    && paragraphRange.InlineShapes.Count == 0
                    && paragraphRange.OMaths.Count == 0
                    && paragraphRange.Fields.Count == 0
                    && paragraphRange.Bookmarks.Count == 0
                    && paragraphRange.Frames.Count == 0)
                    count++;
            }
            finally
            {
                Release(paragraphRange);
                Release(paragraph);
                Release(paragraphs);
                Release(probe);
                Release(table);
            }
        }
        return count;
    }

    private static int CountVisualTeXNumberingBookmarkTriples(Word.Document document) =>
        CountVisualTeXNumberingBookmarks(document) / 3;
}

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
        var pngPath = Path.Combine(artifactRoot, "simple-format-conversion.png");
        var svgPath = Path.Combine(artifactRoot, "simple-format-conversion.svg");
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
                Word.Table? seedTable = null;
                Word.Bookmark? seedCaptionBookmark = null;
                Word.Range? seedTableRange = null;
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
                        seedTable = WordEquationNumbering.FindNumberedEquationTable(
                            sourceDocument,
                            metadata.FormulaId);
                        if (seedTable is null) continue;
                        if (seedTable.Rows.Count != 1 || seedTable.Columns.Count != 3)
                        {
                            Release(seedTable);
                            seedTable = null;
                            continue;
                        }
                        seedTableRange = seedTable.Range;
                        var hostStart = seedTableRange.Start;
                        var hostEnd = seedTableRange.End;
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
                    Release(seedTableRange);
                    Release(seedCaptionBookmark);
                    Release(seedTable);
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

            // Reproduce the real Word quirk from 文档2: after an earlier MathType
            // numbered equation, the first newly-created VisualTeX numbered host can
            // retain one completely empty trailing table row. Conversion must accept
            // and normalize that benign 2x3 host instead of treating it as corruption.
            Word.InlineShape? firstSourceShape = null;
            Word.Table? firstSourceTable = null;
            Word.Row? addedEmptyRow = null;
            try
            {
                firstSourceShape = document.InlineShapes[includeExistingMathType ? 2 : 1];
                var firstMetadata = WordFormulaMetadataReader.TryRead(firstSourceShape)
                    ?? throw new InvalidDataException(
                        "The first VisualTeX source lost metadata before the empty-row regression setup.");
                firstSourceTable = WordEquationNumbering.FindNumberedEquationTable(
                        document,
                        firstMetadata.FormulaId)
                    ?? throw new InvalidDataException(
                        "The first VisualTeX source lost its numbered table before the empty-row regression setup.");
                addedEmptyRow = firstSourceTable.Rows.Add();
                AssertEqual(2, firstSourceTable.Rows.Count,
                    "The empty-row regression setup did not produce a 2x3 VisualTeX numbering table.");
                AssertEqual(3, firstSourceTable.Columns.Count,
                    "The empty-row regression setup changed the VisualTeX numbering-table column count.");
                AssertEqual(1, firstSourceTable.Range.InlineShapes.Count,
                    "The empty-row regression setup unexpectedly duplicated the formula object.");
                Console.WriteLine(
                    "[CONSECUTIVE NUMBERED] Reproduced benign 2x3 host: first VisualTeX numbering table has one empty trailing row.");
            }
            finally
            {
                Release(addedEmptyRow);
                Release(firstSourceTable);
                Release(firstSourceShape);
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
            ForceComCleanup();
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
            ObjectMode = FormulaOleContract.MathTypeOleMode,
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

    private static int CountVisualTeXNumberingBookmarkTriples(Word.Document document) =>
        CountVisualTeXNumberingBookmarks(document) / 3;
}

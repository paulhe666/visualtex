using System.Text;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string RecoveryMathMlA =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mi>a</mi><mo>+</mo><mi>b</mi><mo>=</mo><mi>c</mi></math>";
    private const string RecoveryMathMlB =
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><msup><mi>x</mi><mn>2</mn></msup><mo>+</mo><msup><mi>y</mi><mn>2</mn></msup><mo>=</mo><msup><mi>r</mi><mn>2</mn></msup></math>";

    private static void RunWordOmmlComplexRecoveryAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        RunDenseProseOmmlRedrawAndConversion(artifactRoot);
        RunDeletedNumberedOmmlFormatConversionRecovery(artifactRoot);
        RunDeletedNumberedOmmlRedrawRecovery(artifactRoot);
    }

    private static void RunDenseProseOmmlRedrawAndConversion(string artifactRoot)
    {
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-dense-prose-redraw-conversion.docx");
        TryDeleteAcceptanceFile(documentPath);
        using (var host = new WordPerformanceHost(documentPath: null))
        {
            var service = new WordFormulaService(host.Application);
            var corpus = CreateDenseOmmlRedrawCorpus();
            host.Document.Content.Text = corpus.Source;
            var redrawPlan = service.CaptureLatexRedrawPlan(wholeDocument: true);
            AssertEqual(20, redrawPlan.Targets.Count,
                "Dense prose redraw capture did not find twenty formulas.");
            AssertEqual(0, redrawPlan.DeletedNumberedOmmlResidues.Count,
                "Healthy dense prose was misclassified as deleted OMML residue.");
            var redrawPrepared = PrepareDirectOmmlRedraw(
                redrawPlan,
                corpus.MathMlByTarget);
            var redrawResult = service.ApplyLatexRedrawPlan(
                redrawPlan,
                redrawPrepared);
            AssertEqual(20, redrawResult.FormulaCount,
                "Dense prose redraw did not commit twenty OMML formulas.");
            AssertDenseUnnumberedOmmlState(
                host.Document,
                service,
                "dense prose after OMML redraw");

            var assetRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "VisualTeX",
                "office",
                "temp",
                $"word-omml-complex-{Guid.NewGuid():N}");
            Directory.CreateDirectory(assetRoot);
            var pngPath = Path.Combine(assetRoot, "preview.png");
            var svgPath = Path.Combine(assetRoot, "preview.svg");
            WriteAcceptancePng(pngPath, "dense", 360, 112);
            File.WriteAllText(
                svgPath,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><rect width=\"360\" height=\"112\" fill=\"white\"/><text x=\"16\" y=\"72\" font-family=\"Cambria Math\" font-size=\"42\">x+y</text></svg>",
                Encoding.UTF8);
            var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                360,
                112);

            var plan = service.CaptureFormulaFormatConversionPlan(
                wholeDocument: true,
                FormulaOleContract.WordOmmlMode,
                FormulaOleContract.NativeOleMode);
            AssertEqual(20, plan.Targets.Count,
                "Dense OMML format conversion did not capture twenty formulas.");
            AssertEqual(0, plan.DeletedNumberedOmmlResidues.Count,
                "Healthy dense OMML was misclassified as deleted numbering residue.");
            var prepared = PrepareOmmlVisualTeXStressTargets(
                plan,
                pngPath,
                emfPath);
            var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
            AssertEqual(20, result.FormulaCount,
                "Dense OMML-to-VisualTeX conversion did not commit all formulas.");
            AssertEqual(0, result.FailedFormulaCount,
                "Dense OMML-to-VisualTeX conversion reported failures.");
            AssertEqual(0, host.Document.OMaths.Count,
                "Dense OMML-to-VisualTeX conversion left native OMath sources.");
            AssertEqual(20, CountVisualTeXNativeOleShapes(host.Document),
                "Dense OMML-to-VisualTeX conversion lost OLE targets.");
            AssertDenseVisualTeXOleProseState(
                host.Document,
                "dense prose after OMML-to-VisualTeX conversion");
            host.Save(documentPath);
            Console.WriteLine(
                $"[OMML COMPLEX PASS] Twenty inline/display formulas interleaved with prose redrew without numbering and converted atomically without losing prose: {documentPath}");
        }
    }

    private static void RunDeletedNumberedOmmlFormatConversionRecovery(
        string artifactRoot)
    {
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-partial-delete-format-conversion.docx");
        TryDeleteAcceptanceFile(documentPath);
        using var host = new WordPerformanceHost(documentPath: null);
        var service = new WordFormulaService(host.Application);
        var damagedFormulaId = InsertRecoveryNumberedOmml(
            host,
            service,
            "FORMAT-BEGIN 正文段落",
            @"a+b=c",
            RecoveryMathMlA);
        _ = InsertRecoveryNumberedOmml(
            host,
            service,
            "FORMAT-MIDDLE 正文段落",
            @"x^2+y^2=r^2",
            RecoveryMathMlB);
        AppendRecoveryProse(host, "FORMAT-END 正文段落");
        DeleteOnlyManagedOmmlEquation(host.Document, damagedFormulaId);
        AssertTrue(
            WordEquationNumbering.TryCaptureDeletedNumberedOmmlResidue(
                host.Document,
                damagedFormulaId,
                out _,
                out _),
            "The partial-delete fixture was not recognized as a proven managed residue.");

        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"word-omml-residue-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var pngPath = Path.Combine(assetRoot, "preview.png");
        var svgPath = Path.Combine(assetRoot, "preview.svg");
        WriteAcceptancePng(pngPath, "recovery", 300, 96);
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"300\" height=\"96\" viewBox=\"0 0 300 96\"><rect width=\"300\" height=\"96\" fill=\"white\"/><text x=\"12\" y=\"62\" font-size=\"36\">x²+y²</text></svg>",
            Encoding.UTF8);
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 300, 96);

        var plan = service.CaptureFormulaFormatConversionPlan(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.NativeOleMode);
        AssertEqual(1, plan.Targets.Count,
            "Partial-delete conversion did not retain the healthy numbered OMML target.");
        AssertEqual(1, plan.DeletedNumberedOmmlResidues.Count,
            "Partial-delete conversion did not capture exactly one residue.");
        AssertEqual(damagedFormulaId, plan.DeletedNumberedOmmlResidues[0].FormulaId,
            "Partial-delete conversion captured the wrong residue identity.");
        var prepared = PrepareOmmlVisualTeXStressTargets(plan, pngPath, emfPath);
        var result = service.ApplyFormulaFormatConversionPlan(plan, prepared);
        AssertEqual(1, result.FormulaCount,
            "Partial-delete conversion did not convert the healthy formula.");
        AssertEqual(0, result.FailedFormulaCount,
            "Partial-delete conversion reported a formula failure.");
        AssertEqual(1, CountVisualTeXNativeOleShapes(host.Document),
            "Partial-delete conversion did not retain one healthy VisualTeX target.");
        AssertEqual(0, host.Document.OMaths.Count,
            "Partial-delete conversion left a native OMML source.");
        AssertDeletedOmmlIdentityRemoved(host.Document, damagedFormulaId);
        AssertRecoveryProse(host.Document, "FORMAT-BEGIN", "FORMAT-MIDDLE", "FORMAT-END");
        host.Save(documentPath);
        Console.WriteLine(
            $"[OMML RESIDUE FORMAT PASS] Healthy numbered OMML converted while the proven empty numbered row and its stale identity were removed: {documentPath}");
    }

    private static void RunDeletedNumberedOmmlRedrawRecovery(string artifactRoot)
    {
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-partial-delete-redraw.docx");
        TryDeleteAcceptanceFile(documentPath);
        using (var host = new WordPerformanceHost(documentPath: null))
        {
            var service = new WordFormulaService(host.Application);
            AppendRecoveryProse(
                host,
                "REDRAW-BEGIN 正文夹着 $u_1+v_1=w_1$ REDRAW-MIDDLE 正文继续 REDRAW-END");
            var damagedFormulaId = InsertRecoveryNumberedOmml(
                host,
                service,
                "REDRAW-RESIDUE-BEFORE 正文段落",
                @"a+b=c",
                RecoveryMathMlA);
            DeleteOnlyManagedOmmlEquation(host.Document, damagedFormulaId);
            AssertTrue(
                WordEquationNumbering.TryCaptureDeletedNumberedOmmlResidue(
                    host.Document,
                    damagedFormulaId,
                    out _,
                    out _),
                "The redraw partial-delete fixture was not recognized as residue.");

            var redrawPlan = service.CaptureLatexRedrawPlan(wholeDocument: true);
            AssertEqual(1, redrawPlan.Targets.Count,
                "Residue-aware redraw did not capture the raw inline LaTeX.");
            AssertEqual(1, redrawPlan.DeletedNumberedOmmlResidues.Count,
                "Residue-aware redraw did not capture exactly one deleted numbered row.");
            var redrawPrepared = PrepareDirectOmmlRedraw(
                redrawPlan,
                new[]
                {
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msub><mi>u</mi><mn>1</mn></msub><mo>+</mo><msub><mi>v</mi><mn>1</mn></msub><mo>=</mo><msub><mi>w</mi><mn>1</mn></msub></math>",
                });
            var redrawResult = service.ApplyLatexRedrawPlan(
                redrawPlan,
                redrawPrepared);
            AssertEqual(1, redrawResult.FormulaCount,
                "Residue-aware redraw did not insert its OMML target.");
            AssertEqual(1, host.Document.OMaths.Count,
                "Residue-aware redraw created the wrong OMath count.");
            AssertEqual(0, host.Document.Tables.Count,
                "Residue-aware redraw left the empty numbered table behind.");
            AssertDeletedOmmlIdentityRemoved(host.Document, damagedFormulaId);
            AssertRecoveryProse(host.Document, "REDRAW-BEGIN", "REDRAW-MIDDLE", "REDRAW-END");
            var newFormulaId = WordOmmlFormulaStore.FormulaIds(host.Document).Single();
            var newMetadata = WordOmmlFormulaStore.TryRead(host.Document, newFormulaId)
                ?? throw new InvalidDataException("Redrawn OMML lost its metadata.");
            AssertTrue(!newMetadata.Numbered,
                "OMML redraw changed current product behavior by automatically numbering the display residue replacement.");
            host.Save(documentPath);
            Console.WriteLine(
                $"[OMML RESIDUE REDRAW PASS] Raw prose LaTeX redrew with numbering still disabled while the proven residue and its stale identity were removed: {documentPath}");
        }
    }

    private static void AssertDenseUnnumberedOmmlState(
        Word.Document document,
        WordFormulaService service,
        string context)
    {
        AssertEqual(20, document.OMaths.Count,
            $"{context}: OMath count changed.");
        AssertEqual(0, document.InlineShapes.Count,
            $"{context}: OMML redraw unexpectedly created OLE objects.");
        AssertEqual(0, document.Tables.Count,
            $"{context}: OMML redraw unexpectedly created numbered tables.");
        AssertEqual(20, WordOmmlFormulaStore.StoredFormulaIds(document).Count,
            $"{context}: OMML metadata count changed.");
        AssertEqual(20, WordOmmlFormulaStore.FormulaIds(document).Count,
            $"{context}: OMML identity count changed.");
        var metadata = WordOmmlFormulaStore.FormulaIds(document)
            .Select(id => WordOmmlFormulaStore.TryRead(document, id)
                ?? throw new InvalidDataException($"{context}: metadata {id} is missing."))
            .ToArray();
        AssertEqual(10, metadata.Count(item => item.DisplayMode == "inline"),
            $"{context}: inline formula count changed.");
        AssertEqual(10, metadata.Count(item => item.DisplayMode == "block"),
            $"{context}: display formula count changed.");
        AssertEqual(0, metadata.Count(item => item.Numbered),
            $"{context}: OMML redraw unexpectedly enabled numbering.");
        AssertDenseProseMarkers(document);
        var recapture = service.CaptureFormulaFormatConversionPlan(
            wholeDocument: true,
            FormulaOleContract.WordOmmlMode,
            FormulaOleContract.NativeOleMode);
        AssertEqual(20, recapture.Targets.Count,
            $"{context}: format conversion recapture lost formulas.");
        AssertEqual(0, recapture.DeletedNumberedOmmlResidues.Count,
            $"{context}: healthy formulas were classified as residue.");
    }

    private static string InsertRecoveryNumberedOmml(
        WordPerformanceHost host,
        WordFormulaService service,
        string prose,
        string latex,
        string mathMl)
    {
        var selection = host.Application.Selection;
        selection.EndKey(Word.WdUnits.wdStory);
        selection.TypeText(prose);
        selection.TypeParagraph();
        selection.EndKey(Word.WdUnits.wdStory);
        var formulaId = Guid.NewGuid().ToString("D");
        var session = CreateNumberedOmmlTabSession(
            formulaId,
            host.Document.FullName,
            selection.Start,
            selection.End,
            latex,
            originalMetadata: null);
        service.InsertOmml(session, mathMl);
        return formulaId;
    }

    private static void AppendRecoveryProse(
        WordPerformanceHost host,
        string text)
    {
        var selection = host.Application.Selection;
        selection.EndKey(Word.WdUnits.wdStory);
        selection.TypeParagraph();
        selection.TypeText(text);
    }

    private static void DeleteOnlyManagedOmmlEquation(
        Word.Document document,
        string formulaId)
    {
        Word.Range? equation = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Cannot create partial-delete fixture: metadata {formulaId} is missing.");
            equation = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                formulaId,
                metadata);
            equation.Delete();
        }
        finally { Release(equation); }
    }

    private static void AssertDeletedOmmlIdentityRemoved(
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            AssertTrue(bookmark is null,
                "The deleted OMML VTOMML identity survived residue cleanup.");
        }
        finally { Release(bookmark); }
        AssertTrue(WordOmmlFormulaStore.TryRead(document, formulaId) is null,
            "The deleted OMML CustomXML metadata survived residue cleanup.");
        var names = new[]
        {
            WordEquationNumbering.EquationBookmarkName(formulaId),
            WordEquationNumbering.NativeCaptionBookmarkName(formulaId),
            WordEquationNumbering.NativeNumberBookmarkName(formulaId),
        };
        foreach (var name in names)
            AssertTrue(!document.Bookmarks.Exists(name),
                $"Generated numbering bookmark {name} survived residue cleanup.");
    }

    private static void AssertRecoveryProse(
        Word.Document document,
        params string[] markers)
    {
        var text = document.Content.Text ?? string.Empty;
        var cursor = -1;
        foreach (var marker in markers)
        {
            var next = text.IndexOf(marker, cursor + 1, StringComparison.Ordinal);
            AssertTrue(next > cursor,
                $"Recovery prose marker {marker} was lost or reordered.");
            cursor = next;
        }
    }

    private static (string Source, IReadOnlyList<string> MathMlByTarget)
        CreateDenseOmmlRedrawCorpus()
    {
        var source = new StringBuilder();
        var mathMl = new List<string>();
        for (var index = 1; index <= 10; index++)
        {
            source.Append(
                $"DENSE-{index:00}-BEFORE 中文正文夹着行内公式 "
                + $"$x_{{{index}}}+y_{{{index}}}=z_{{{index}}}$"
                + $" DENSE-{index:00}-AFTER，正文继续。\r");
            mathMl.Add(
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + $"<msub><mi>x</mi><mn>{index}</mn></msub><mo>+</mo>"
                + $"<msub><mi>y</mi><mn>{index}</mn></msub><mo>=</mo>"
                + $"<msub><mi>z</mi><mn>{index}</mn></msub></math>");
            source.Append(
                $"\\[\\sum_{{k=1}}^{{{index + 2}}}\\frac{{a_{{k,{index}}}}}{{1+k^2}}=S_{{{index}}}\\]\r");
            mathMl.Add(
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow>"
                + "<munderover><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow>"
                + $"<mn>{index + 2}</mn></munderover><mfrac><msub><mi>a</mi>"
                + $"<mrow><mi>k</mi><mo>,</mo><mn>{index}</mn></mrow></msub>"
                + "<mrow><mn>1</mn><mo>+</mo><msup><mi>k</mi><mn>2</mn></msup></mrow></mfrac>"
                + $"<mo>=</mo><msub><mi>S</mi><mn>{index}</mn></msub></mrow></math>");
        }
        return (source.ToString(), mathMl);
    }

    private static Dictionary<string, PreparedWordBulkFormula> PrepareDirectOmmlRedraw(
        WordLatexRedrawPlan plan,
        IReadOnlyList<string> mathMlByTarget)
    {
        AssertEqual(plan.Targets.Count, mathMlByTarget.Count,
            "Direct OMML redraw corpus and capture counts differ.");
        var prepared = new Dictionary<string, PreparedWordBulkFormula>(StringComparer.Ordinal);
        for (var index = 0; index < plan.Targets.Count; index++)
        {
            var target = plan.Targets[index];
            var mathMl = mathMlByTarget[index];
            prepared[target.Id] = new PreparedWordBulkFormula
            {
                Run = new WordBulkRun
                {
                    Id = target.Id,
                    IsFormula = true,
                    Latex = target.Latex,
                    DisplayMode = target.DisplayMode,
                },
                Session = new OfficeSessionDocument
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Mode = "create",
                    Host = "word",
                    FormulaId = Guid.NewGuid().ToString("D"),
                    SourceDocumentId = plan.DocumentId,
                    Title = "Direct complex OMML redraw acceptance",
                    CodeFormat = "latex",
                    DisplayMode = target.DisplayMode,
                    ObjectMode = FormulaOleContract.WordOmmlMode,
                    Numbered = false,
                    FontSizePt = target.FontSizePt,
                    Lines = new List<FormulaLine>
                    {
                        new() { Id = Guid.NewGuid().ToString("D"), Latex = target.Latex },
                    },
                    ExportResult = new OfficeExportDocument
                    {
                        MathMl = mathMl,
                        Width = 360,
                        Height = 112,
                        Baseline = 84,
                    },
                },
                MathMl = mathMl,
            };
        }
        return prepared;
    }

    private static void AssertDenseProseMarkers(Word.Document document)
    {
        var text = document.Content.Text ?? string.Empty;
        var cursor = -1;
        for (var index = 1; index <= 10; index++)
        {
            foreach (var marker in new[]
                     {
                         $"DENSE-{index:00}-BEFORE",
                         $"DENSE-{index:00}-AFTER",
                     })
            {
                var next = text.IndexOf(marker, cursor + 1, StringComparison.Ordinal);
                AssertTrue(next > cursor,
                    $"Dense prose marker {marker} was lost or reordered.");
                cursor = next;
            }
        }
    }

    private static void AssertDenseVisualTeXOleProseState(
        Word.Document document,
        string context)
    {
        AssertDenseProseMarkers(document);
        var proseParagraphs = 0;
        for (var index = 1; index <= document.Paragraphs.Count; index++)
        {
            Word.Paragraph? paragraph = null;
            Word.Range? range = null;
            Word.InlineShapes? shapes = null;
            Word.InlineShape? shape = null;
            try
            {
                paragraph = document.Paragraphs[index];
                range = paragraph.Range;
                if ((range.Text ?? string.Empty).IndexOf(
                        "DENSE-",
                        StringComparison.Ordinal) < 0)
                    continue;
                proseParagraphs++;
                AssertEqual(0, range.OMaths.Count,
                    $"{context}: prose paragraph {index} retained an OMath source.");
                shapes = range.InlineShapes;
                AssertEqual(1, shapes.Count,
                    $"{context}: prose paragraph {index} does not contain exactly one inline OLE.");
                shape = shapes[1];
                AssertTrue(WordFormulaMetadataReader.IsNativeOle(shape),
                    $"{context}: prose paragraph {index} contains a non-VisualTeX OLE.");
                var metadata = WordFormulaMetadataReader.TryRead(shape)
                    ?? throw new InvalidDataException(
                        $"{context}: prose OLE {index} has no metadata.");
                AssertEqual("inline", metadata.DisplayMode,
                    $"{context}: prose OLE {index} changed display mode.");
                AssertTrue(!metadata.Numbered,
                    $"{context}: prose OLE {index} became numbered.");
            }
            finally
            {
                Release(shape);
                Release(shapes);
                Release(range);
                Release(paragraph);
            }
        }
        AssertEqual(10, proseParagraphs,
            $"{context}: prose paragraph count changed.");
    }
}

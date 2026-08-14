using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class InlineFractionGridMetric
    {
        internal float FormulaHeightPx { get; set; }
        internal float LinePitchPt { get; set; }
        internal int DisableLineHeightGrid { get; set; }
        internal float SpaceBefore { get; set; }
        internal float SpaceAfter { get; set; }
        internal Word.WdLineSpacing LineSpacingRule { get; set; }
    }

    private static void RunWordInlineFractionLineGridAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);

            var denominatorLower = MeasureInlineFractionGridFormula(
                application,
                artifactRoot,
                "denominator-lower-p",
                @"\lambda=\frac{h}{p}",
                "<mi>λ</mi><mo>=</mo><mfrac><mi>h</mi><mi>p</mi></mfrac>");
            var denominatorUpper = MeasureInlineFractionGridFormula(
                application,
                artifactRoot,
                "denominator-upper-P",
                @"\lambda=\frac{h}{P}",
                "<mi>λ</mi><mo>=</mo><mfrac><mi>h</mi><mi>P</mi></mfrac>");
            var numeratorLower = MeasureInlineFractionGridFormula(
                application,
                artifactRoot,
                "numerator-lower-p",
                @"\lambda=\frac{p}{h}",
                "<mi>λ</mi><mo>=</mo><mfrac><mi>p</mi><mi>h</mi></mfrac>");
            var numeratorUpper = MeasureInlineFractionGridFormula(
                application,
                artifactRoot,
                "numerator-upper-P",
                @"\lambda=\frac{P}{h}",
                "<mi>λ</mi><mo>=</mo><mfrac><mi>P</mi><mi>h</mi></mfrac>");

            AssertEqual(
                -1,
                denominatorLower.DisableLineHeightGrid,
                "Inline OMML fractions must opt their paragraph out of Word line-grid height quantization.");
            AssertEqual(
                -1,
                numeratorLower.DisableLineHeightGrid,
                "A lowercase descender in the numerator must receive the same line-grid protection.");
            AssertTrue(
                denominatorLower.FormulaHeightPx <= denominatorUpper.FormulaHeightPx * 1.35f + 2f,
                $"Lowercase p in an inline denominator still triggered a grid-height jump: p={denominatorLower.FormulaHeightPx:0.##}px, P={denominatorUpper.FormulaHeightPx:0.##}px.");
            AssertTrue(
                Math.Abs(denominatorLower.LinePitchPt - denominatorUpper.LinePitchPt) <= 5f,
                $"Lowercase p in an inline denominator still changed the next-line pitch excessively: p={denominatorLower.LinePitchPt:0.##}pt, P={denominatorUpper.LinePitchPt:0.##}pt.");
            AssertTrue(
                numeratorLower.FormulaHeightPx <= numeratorUpper.FormulaHeightPx * 1.35f + 2f,
                $"Lowercase p in an inline numerator still triggered a grid-height jump: p={numeratorLower.FormulaHeightPx:0.##}px, P={numeratorUpper.FormulaHeightPx:0.##}px.");
            AssertTrue(
                Math.Abs(numeratorLower.LinePitchPt - numeratorUpper.LinePitchPt) <= 5f,
                $"Lowercase p in an inline numerator still changed the next-line pitch excessively: p={numeratorLower.LinePitchPt:0.##}pt, P={numeratorUpper.LinePitchPt:0.##}pt.");

            foreach (var metric in new[]
                     {
                         denominatorLower,
                         denominatorUpper,
                         numeratorLower,
                         numeratorUpper,
                     })
            {
                AssertTrue(
                    Math.Abs(metric.SpaceBefore) < 0.01f && Math.Abs(metric.SpaceAfter) < 0.01f,
                    "Inline fraction line-grid protection changed paragraph spacing before/after.");
                AssertEqual(
                    Word.WdLineSpacing.wdLineSpaceSingle,
                    metric.LineSpacingRule,
                    "Inline fraction line-grid protection changed the user's line-spacing rule.");
            }

            var plainInlineGridState = MeasurePlainInlineGridState(application);
            AssertEqual(
                0,
                plainInlineGridState,
                "A plain inline OMML formula without a stacked fraction must not change the paragraph's line-grid setting.");

            var edited = RunInlineFractionGridEditAcceptance(application, artifactRoot);
            AssertEqual(
                -1,
                edited.DisableLineHeightGrid,
                "Editing an existing inline fraction to lowercase p re-enabled Word line-grid quantization.");
            AssertTrue(
                edited.FormulaHeightPx <= denominatorUpper.FormulaHeightPx * 1.35f + 2f,
                "Editing an existing inline fraction from P to p reintroduced the doubled line box.");

            Console.WriteLine(
                "Word inline fraction line-grid acceptance passed: "
                + $"den p/P={denominatorLower.FormulaHeightPx:0.##}/{denominatorUpper.FormulaHeightPx:0.##}px, "
                + $"num p/P={numeratorLower.FormulaHeightPx:0.##}/{numeratorUpper.FormulaHeightPx:0.##}px, "
                + $"plainGrid={plainInlineGridState}.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static InlineFractionGridMetric MeasureInlineFractionGridFormula(
        Word.Application application,
        string artifactRoot,
        string artifactName,
        string latex,
        string mathMlBody)
    {
        Word.Document? document = null;
        try
        {
            document = CreateLineGridInlineAcceptanceDocument(application);
            var formulaId = Guid.NewGuid().ToString("D");
            InsertInlineAcceptanceOmml(
                application,
                document,
                formulaId,
                latex,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\">"
                + mathMlBody
                + "</math>");
            var metric = ReadInlineFractionGridMetric(application, document, formulaId);
            document.SaveAs2(
                Path.Combine(artifactRoot, artifactName + ".docx"),
                Word.WdSaveFormat.wdFormatXMLDocument);
            return metric;
        }
        finally
        {
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
        }
    }

    private static int MeasurePlainInlineGridState(Word.Application application)
    {
        Word.Document? document = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? paragraphFormat = null;
        try
        {
            document = CreateLineGridInlineAcceptanceDocument(application);
            var formulaId = Guid.NewGuid().ToString("D");
            InsertInlineAcceptanceOmml(
                application,
                document,
                formulaId,
                "x+p",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi>x</mi><mo>+</mo><mi>p</mi></math>");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Plain inline grid acceptance lost the OMML bookmark.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            paragraphs = range.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphFormat = paragraphRange.ParagraphFormat;
            return paragraphFormat.DisableLineHeightGrid;
        }
        finally
        {
            Release(paragraphFormat);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(range);
            Release(bookmark);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
        }
    }

    private static InlineFractionGridMetric RunInlineFractionGridEditAcceptance(
        Word.Application application,
        string artifactRoot)
    {
        Word.Document? document = null;
        Word.Bookmark? bookmark = null;
        Word.Range? range = null;
        try
        {
            document = CreateLineGridInlineAcceptanceDocument(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var lineId = Guid.NewGuid().ToString("D");
            InsertInlineAcceptanceOmml(
                application,
                document,
                formulaId,
                @"\lambda=\frac{h}{P}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi>λ</mi><mo>=</mo><mfrac><mi>h</mi><mi>P</mi></mfrac></math>",
                lineId);

            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Inline fraction edit grid acceptance lost its bookmark.");
            var stored = WordOmmlFormulaStore.TryRead(document, bookmark)
                ?? throw new InvalidDataException("Inline fraction edit grid acceptance lost metadata.");
            range = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var editSession = new OfficeSessionDocument
            {
                Id = Guid.NewGuid().ToString("D"),
                Mode = "edit",
                Host = "word",
                FormulaId = formulaId,
                SourceDocumentId = document.FullName,
                SourceObjectId = WordRangeReference(range.Start, range.End),
                Title = "VisualTeX inline fraction line-grid edit acceptance",
                CodeFormat = "latex",
                DisplayMode = "inline",
                ObjectMode = FormulaOleContract.WordOmmlMode,
                Numbered = false,
                FontSizePt = 14,
                OriginalMetadata = stored,
                Lines = new List<FormulaLine>
                {
                    new() { Id = lineId, Latex = @"\lambda=\frac{h}{p}" },
                },
                ExportResult = new OfficeExportDocument
                {
                    FormulaLetterFont = "katex",
                    FormulaChineseFont = "songti",
                },
            };
            Release(range);
            range = null;
            Release(bookmark);
            bookmark = null;
            new WordFormulaService(application).ReplaceOmml(
                editSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"inline\"><mi>λ</mi><mo>=</mo><mfrac><mi>h</mi><mi>p</mi></mfrac></math>");

            var metric = ReadInlineFractionGridMetric(application, document, formulaId);
            document.SaveAs2(
                Path.Combine(artifactRoot, "edit-denominator-P-to-p.docx"),
                Word.WdSaveFormat.wdFormatXMLDocument);
            return metric;
        }
        finally
        {
            Release(range);
            Release(bookmark);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
        }
    }

    private static Word.Document CreateLineGridInlineAcceptanceDocument(Word.Application application)
    {
        var document = application.Documents.Add();
        document.Activate();
        Word.PageSetup? pageSetup = null;
        Word.Range? content = null;
        Microsoft.Office.Interop.Word.Font? bodyFont = null;
        Word.ParagraphFormat? paragraphFormat = null;
        try
        {
            pageSetup = document.PageSetup;
            pageSetup.LayoutMode = Word.WdLayoutMode.wdLayoutModeLineGrid;
            pageSetup.LinesPage = 28;

            content = document.Content;
            content.Text = "除此以外 公式 后文\r横纵坐标的单位往往\r";
            bodyFont = content.Font;
            bodyFont.Name = "SimSun";
            bodyFont.Size = 14;
            paragraphFormat = content.ParagraphFormat;
            paragraphFormat.SpaceBefore = 0;
            paragraphFormat.SpaceAfter = 0;
            paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            paragraphFormat.DisableLineHeightGrid = 0;
            return document;
        }
        catch
        {
            try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            throw;
        }
        finally
        {
            Release(paragraphFormat);
            Release(bodyFont);
            Release(content);
            Release(pageSetup);
        }
    }

    private static void InsertInlineAcceptanceOmml(
        Word.Application application,
        Word.Document document,
        string formulaId,
        string latex,
        string mathMl,
        string? lineId = null)
    {
        var insertionPosition = "除此以外 ".Length;
        application.Selection.SetRange(insertionPosition, insertionPosition);
        var session = new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = document.FullName,
            SourceObjectId = WordRangeReference(insertionPosition, insertionPosition),
            Title = "VisualTeX inline fraction line-grid acceptance",
            CodeFormat = "latex",
            DisplayMode = "inline",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = false,
            FontSizePt = 14,
            Lines = new List<FormulaLine>
            {
                new() { Id = lineId ?? Guid.NewGuid().ToString("D"), Latex = latex },
            },
            ExportResult = new OfficeExportDocument
            {
                FormulaLetterFont = "katex",
                FormulaChineseFont = "songti",
            },
        };
        new WordFormulaService(application).InsertOmml(session, mathMl);
    }

    private static InlineFractionGridMetric ReadInlineFractionGridMetric(
        Word.Application application,
        Word.Document document,
        string formulaId)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Word.Range? firstLine = null;
        Word.Range? secondLine = null;
        Word.Paragraph? secondParagraph = null;
        Word.Window? window = null;
        try
        {
            document.Repaginate();
            Thread.Sleep(80);
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("Inline fraction grid acceptance lost the OMML bookmark.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            paragraphs = formulaRange.Paragraphs;
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            paragraphFormat = paragraphRange.ParagraphFormat;
            firstLine = document.Range(paragraphRange.Start, paragraphRange.Start);
            secondParagraph = document.Paragraphs[2];
            secondLine = document.Range(secondParagraph.Range.Start, secondParagraph.Range.Start);
            var firstY = Convert.ToSingle(
                firstLine.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            var secondY = Convert.ToSingle(
                secondLine.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);

            var left = 0;
            var top = 0;
            var width = 0;
            var height = 0;
            try
            {
                window = application.ActiveWindow;
                window.ScrollIntoView(formulaRange, true);
                Thread.Sleep(40);
                window.GetPoint(out left, out top, out width, out height, formulaRange);
            }
            catch { }

            return new InlineFractionGridMetric
            {
                FormulaHeightPx = height,
                LinePitchPt = secondY - firstY,
                DisableLineHeightGrid = paragraphFormat.DisableLineHeightGrid,
                SpaceBefore = paragraphFormat.SpaceBefore,
                SpaceAfter = paragraphFormat.SpaceAfter,
                LineSpacingRule = paragraphFormat.LineSpacingRule,
            };
        }
        finally
        {
            Release(window);
            Release(secondLine);
            Release(secondParagraph);
            Release(firstLine);
            Release(paragraphFormat);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(formulaRange);
            Release(bookmark);
        }
    }
}

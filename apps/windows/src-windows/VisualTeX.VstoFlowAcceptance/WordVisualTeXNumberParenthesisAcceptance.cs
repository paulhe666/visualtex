using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordVisualTeXNumberParenthesisAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "visualtex-number-parenthesis.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"number-parenthesis-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "visualtex-number-parenthesis.svg");
        File.WriteAllText(svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"220\" height=\"70\" viewBox=\"0 0 220 70\"><text x=\"4\" y=\"48\" font-size=\"36\">x = 1</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 220, 70);
        var pngDataUrl = CreatePngDataUrl("number-parenthesis", 220, 70);
        var pngPath = Path.Combine(assetRoot, "visualtex-number-parenthesis.png");
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var createSession = CreateNumberedPerformanceSession(
                    "create",
                    formulaId,
                    document.FullName,
                    WordRangeReference(insertion.Start, insertion.End),
                    originalMetadata: null,
                    latex: @"x=1");
                createSession.ExportResult = new OfficeExportDocument
                {
                    Width = 220,
                    Height = 70,
                    Baseline = 52.5f,
                };
                service.InsertOle(createSession, pngPath, emfPath);
            }
            finally { Release(insertion); }

            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "fresh numbered VisualTeX insertion");

            shape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var originalMetadata = WordFormulaMetadataReader.TryRead(shape)
                ?? throw new InvalidDataException("Fresh numbered VisualTeX OLE lost metadata before edit.");
            shapeRange = shape.Range;
            var editSession = CreateNumberedPerformanceSession(
                "edit",
                formulaId,
                document.FullName,
                WordRangeReference(shapeRange.Start, shapeRange.End),
                originalMetadata,
                latex: @"x=2");
            editSession.ExportResult = new OfficeExportDocument
            {
                Width = 220,
                Height = 70,
                Baseline = 52.5f,
            };
            service.ReplaceOle(editSession, pngPath, emfPath);
            Release(shapeRange); shapeRange = null;
            Release(shape); shape = null;

            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "numbered VisualTeX edit/reconcile");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(documentPath, ReadOnly: false, Visible: false);
            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "numbered VisualTeX save/reopen");

            Console.WriteLine(
                "VisualTeX tab-numbering acceptance passed: fresh insert, REF update, edit/reconcile and save/reopen kept one justified paragraph with centered/right tab stops and parentheses outside REF.Result.");
        }
        finally
        {
            Release(shapeRange);
            Release(shape);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            try { Directory.Delete(assetRoot, recursive: true); } catch { }
            ForceComCleanup();
        }
    }

    private static void AssertVisualTeXNumberedTabHost(
        Word.Document document,
        string formulaId,
        bool updateReference,
        string context,
        bool requireNativeOle = true,
        bool requireFormulaMetadata = true)
    {
        Word.Range? ownerRange = null;
        Word.Range? visibleRange = null;
        Word.Range? visibleTextRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Word.InlineShapes? shapes = null;
        Word.InlineShape? shape = null;
        Word.Range? shapeRange = null;
        Word.Range? objectResultRange = null;
        Word.Font? objectResultFont = null;
        Word.Range? precedingShape = null;
        Word.Range? betweenFormulaAndNumber = null;
        Word.Range? numberEnd = null;
        Word.Font? visibleFont = null;
        Word.Range? paragraphMark = null;
        Word.Font? paragraphMarkFont = null;
        Word.Fields? fields = null;
        Word.Field? reference = null;
        Word.Range? candidateCode = null;
        Word.Range? result = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        try
        {
            document.Repaginate();
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": VisualTeX numbering owner is missing.");
            AssertTrue(
                !(bool)ownerRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": numbered VisualTeX OLE is still hosted by a table instead of a MathType-style tab paragraph.");

            paragraphs = ownerRange.Paragraphs;
            AssertEqual(1, paragraphs.Count, context + ": numbered VisualTeX OLE spans multiple paragraphs.");
            paragraph = paragraphs[1];
            format = paragraph.Format;
            AssertEqual(
                Word.WdParagraphAlignment.wdAlignParagraphJustify,
                format.Alignment,
                context + ": numbered VisualTeX paragraph does not match MathType's justified display style.");
            AssertNear(0f, format.LeftIndent, 0.5f,
                context + ": numbered VisualTeX paragraph has an unexpected left indent.");
            AssertNear(0f, format.RightIndent, 0.5f,
                context + ": numbered VisualTeX paragraph has an unexpected right indent.");

            tabStops = format.TabStops;
            var sawCenter = false;
            var sawRight = false;
            var centerPosition = 0f;
            var rightPosition = 0f;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop); tabStop = tabStops[index];
                if (tabStop.Alignment == Word.WdTabAlignment.wdAlignTabCenter)
                {
                    sawCenter = true;
                    centerPosition = tabStop.Position;
                }
                if (tabStop.Alignment == Word.WdTabAlignment.wdAlignTabRight)
                {
                    sawRight = true;
                    rightPosition = tabStop.Position;
                }
            }
            AssertTrue(sawCenter && sawRight,
                context + ": center/right equation tab stops are missing.");

            sections = ownerRange.Sections;
            AssertTrue(sections.Count > 0, context + ": paragraph has no Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            var expected = WordEquationNumbering.CalculateEquationTabStops(
                pageSetup.PageWidth,
                pageSetup.LeftMargin,
                pageSetup.RightMargin,
                0f,
                0f);
            AssertNear(expected.Center, centerPosition, 0.75f,
                context + ": equation center tab is not at the writable-page midpoint.");
            AssertNear(expected.Right, rightPosition, 0.75f,
                context + ": equation-number tab is not at the writable-page right edge.");

            shapes = ownerRange.InlineShapes;
            AssertEqual(1, shapes.Count,
                context + ": tab paragraph does not contain exactly one VisualTeX OLE object.");
            shape = shapes[1];
            if (requireNativeOle)
            {
                AssertTrue(WordFormulaMetadataReader.IsNativeOle(shape),
                    context + ": the tab paragraph object is not a VisualTeX native OLE formula.");
            }
            FormulaMetadata? metadata = null;
            if (requireFormulaMetadata)
            {
                metadata = WordFormulaMetadataReader.TryRead(shape)
                    ?? throw new InvalidDataException(context + ": VisualTeX OLE metadata is missing.");
                AssertEqual(formulaId, metadata.FormulaId,
                    context + ": the tab paragraph belongs to another VisualTeX formula.");
            }
            shapeRange = shape.Range;

            if (requireNativeOle && metadata is not null)
            {
                for (var position = shapeRange.Start; position < shapeRange.End; position++)
                {
                    Release(objectResultRange); objectResultRange = null;
                    objectResultRange = document.Range(position, position + 1);
                    if (!string.Equals(objectResultRange.Text, "\u0001", StringComparison.Ordinal))
                        continue;
                    objectResultFont = objectResultRange.Font;
                    break;
                }
                if (objectResultFont is null)
                    throw new InvalidDataException(context + ": VisualTeX OLE has no U+0001 object-result character.");
                var semanticFontSize = FormulaFontSize.ResolveSemanticFontSize(metadata);
                var expectedOlePosition = WordInlineAlignment.CalculateFontPositionWithLegacyFallback(
                    shape.Height,
                    (float)(metadata.RenderHeightPx ?? 0d),
                    metadata.Baseline.HasValue ? (float?)metadata.Baseline.Value : null,
                    existingFontPosition: null,
                    sourceSemanticFontSizePoints: semanticFontSize,
                    targetSemanticFontSizePoints: semanticFontSize);
                AssertNear(expectedOlePosition, objectResultFont.Position, 0.1f,
                    context + ": VisualTeX OLE is not using its own exported baseline for display placement.");
                AssertTrue(objectResultFont.Position <= 0,
                    context + ": numbered VisualTeX OLE was raised above the Word text baseline.");
            }

            Word.Range? visibleObjectStart = null;
            try
            {
                visibleObjectStart = (objectResultRange ?? shapeRange).Duplicate;
                visibleObjectStart.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                var shapeStartX = Convert.ToSingle(
                    visibleObjectStart.Information[Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary]);
                var shapeCenterX = shapeStartX + shape.Width / 2f;
                AssertNear(expected.Center, shapeCenterX, 0.75f,
                    context + ": VisualTeX OLE is not physically centered on the center tab stop.");
            }
            finally { Release(visibleObjectStart); }

            AssertTrue(shapeRange.Start > ownerRange.Start,
                context + ": formula has no leading center-tab character.");
            precedingShape = document.Range(shapeRange.Start - 1, shapeRange.Start);
            AssertEqual("\t", precedingShape.Text,
                context + ": formula is not positioned after the center tab.");

            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": visible equation-number range is missing.");
            AssertTrue(visibleRange.Start >= shapeRange.End,
                context + ": visible equation number precedes or overlaps the formula.");
            betweenFormulaAndNumber = document.Range(shapeRange.End, visibleRange.Start);
            var ownerText = ownerRange.Text ?? string.Empty;
            var betweenText = betweenFormulaAndNumber.Text ?? string.Empty;
            var rawVisible = visibleRange.Text ?? string.Empty;
            Console.WriteLine(
                $"  {context}: raw owner codes={string.Join(",", ownerText.Select(character => $"U+{(int)character:X4}"))}, owner={ownerRange.Start}-{ownerRange.End}, shape={shapeRange.Start}-{shapeRange.End}, visible={visibleRange.Start}-{visibleRange.End} codes={string.Join(",", rawVisible.Select(character => $"U+{(int)character:X4}"))}, between={betweenFormulaAndNumber.Start}-{betweenFormulaAndNumber.End} codes={string.Join(",", betweenText.Select(character => $"U+{(int)character:X4}"))}.");
            AssertTrue(
                ownerText.Count(character => character == '\t') >= 2
                && (betweenText.IndexOf('\t') >= 0 || rawVisible.StartsWith("\t", StringComparison.Ordinal)),
                context + ": formula and equation number are not separated by the right tab.");

            var visible = rawVisible.TrimStart('\t').TrimEnd('\r', '\a');
            visibleTextRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(context + ": pure visible equation-number text range is missing.");
            AssertEqual(
                visible,
                (visibleTextRange.Text ?? string.Empty).TrimEnd('\r', '\a'),
                context + ": pure equation-number text range still contains a layout tab or lost visible text.");
            AssertTrue(
                visible.StartsWith("(", StringComparison.Ordinal)
                && visible.EndsWith(")", StringComparison.Ordinal),
                context + ": visible equation number is not enclosed by both parentheses: '" + visible + "'.");

            visibleFont = visibleTextRange.Font;
            AssertNear(0f, visibleFont.Position, 0.1f,
                context + ": tab-layout equation number has a manual vertical offset.");
            paragraphMark = document.Range(ownerRange.End - 1, ownerRange.End);
            paragraphMarkFont = paragraphMark.Font;
            AssertNear(0f, paragraphMarkFont.Position, 0.1f,
                context + ": display paragraph mark has a manual vertical offset.");
            AssertNear(paragraphMarkFont.Size, visibleFont.Size, 0.1f,
                context + ": VisualTeX number does not inherit the display paragraph point size.");
            var paragraphFontName = paragraphMarkFont.NameAscii ?? paragraphMarkFont.Name ?? string.Empty;
            var visibleFontName = visibleFont.NameAscii ?? visibleFont.Name ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(paragraphFontName)
                && !string.IsNullOrWhiteSpace(visibleFontName))
            {
                AssertEqual(paragraphFontName, visibleFontName,
                    context + ": VisualTeX number does not inherit the display paragraph typeface.");
            }
            numberEnd = visibleTextRange.Duplicate;
            numberEnd.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            var numberEndX = Convert.ToSingle(
                numberEnd.Information[Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary]);
            AssertNear(expected.Right, numberEndX, 0.75f,
                context + ": visible equation number does not end on the right tab stop.");
            var numberY = Convert.ToSingle(
                visibleTextRange.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            var paragraphMarkY = Convert.ToSingle(
                paragraphMark.Information[Word.WdInformation.wdVerticalPositionRelativeToPage]);
            AssertNear(paragraphMarkY, numberY, 0.75f,
                context + ": equation number and paragraph mark are not on the same Word baseline.");

            fields = visibleRange.Fields;
            for (var index = 1; index <= fields.Count; index++)
            {
                Word.Field? candidate = null;
                try
                {
                    candidate = fields[index];
                    Release(candidateCode); candidateCode = candidate.Code;
                    if ((candidateCode.Text ?? string.Empty).IndexOf(
                            "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    reference = candidate;
                    candidate = null;
                    break;
                }
                finally { Release(candidate); }
            }
            if (reference is null)
                throw new InvalidDataException(context + ": visible number REF field is missing.");

            if (updateReference)
            {
                reference.Update();
                if (requireNativeOle)
                {
                    AssertTrue(
                        WordEquationNumbering.UpdateEquationNumbers(document) >= 1,
                        context + ": explicit equation-number refresh did not see the numbered VisualTeX formula.");
                    WordEquationNumbering.UpdateNativeCrossReferences(document);
                }
            }
            Release(visibleRange); visibleRange = null;
            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": REF update removed the visible equation-number range.");
            visible = (visibleRange.Text ?? string.Empty)
                .TrimStart('\t')
                .TrimEnd('\r', '\a');
            AssertTrue(
                visible.StartsWith("(", StringComparison.Ordinal)
                && visible.EndsWith(")", StringComparison.Ordinal),
                context + ": REF update removed a parenthesis: '" + visible + "'.");

            Release(visibleFont); visibleFont = null;
            Release(visibleTextRange); visibleTextRange = null;
            visibleTextRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(
                    document,
                    formulaId)
                ?? throw new InvalidDataException(context + ": number update removed the pure equation-number text range.");
            visibleFont = visibleTextRange.Font;
            AssertNear(0f, visibleFont.Position, 0.1f,
                context + ": F9/update-number moved the tab-layout number off the paragraph baseline.");
            AssertNear(paragraphMarkFont.Size, visibleFont.Size, 0.1f,
                context + ": F9/update-number replaced the inherited paragraph point size.");
            visibleFontName = visibleFont.NameAscii ?? visibleFont.Name ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(paragraphFontName)
                && !string.IsNullOrWhiteSpace(visibleFontName))
            {
                AssertEqual(paragraphFontName, visibleFontName,
                    context + ": F9/update-number replaced the inherited paragraph typeface.");
            }

            result = reference.Result;
            var resultText = result.Text ?? string.Empty;
            AssertTrue(
                resultText.IndexOf('(') < 0 && resultText.IndexOf(')') < 0,
                context + ": parenthesis leaked inside REF.Result: '" + resultText + "'.");
            AssertTrue(
                result.Start >= visibleRange.Start && result.End < visibleRange.End,
                context + ": REF.Result reaches outside the visible-number bookmark; ')' is not safely outside the field result.");
            Console.WriteLine(
                $"  {context}: owner={ownerRange.Start}-{ownerRange.End}, formula={shapeRange.Start}-{shapeRange.End}, visible='{visible}', tabs={centerPosition:0.##}/{rightPosition:0.##}, olePosition={objectResultFont?.Position ?? 0}, numberPosition={visibleFont.Position}, numberFont='{visibleFontName}', numberY={numberY:0.##}, markY={paragraphMarkY:0.##}, REF.Result='{resultText}'.");
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(result);
            Release(reference);
            Release(paragraphMarkFont);
            Release(paragraphMark);
            Release(visibleFont);
            Release(numberEnd);
            Release(candidateCode);
            Release(fields);
            Release(betweenFormulaAndNumber);
            Release(precedingShape);
            Release(objectResultFont);
            Release(objectResultRange);
            Release(shapeRange);
            Release(shape);
            Release(shapes);
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(visibleTextRange);
            Release(visibleRange);
            Release(ownerRange);
        }
    }
}

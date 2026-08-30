using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlTabFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "visualtex-omml-tab-format-roundtrip.docx");
        var assetRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "VisualTeX",
            "office",
            "temp",
            $"omml-tab-roundtrip-{Guid.NewGuid():N}");
        Directory.CreateDirectory(assetRoot);
        var svgPath = Path.Combine(assetRoot, "formula.svg");
        var pngPath = Path.Combine(assetRoot, "formula.png");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"112\" viewBox=\"0 0 360 112\"><text x=\"8\" y=\"78\" font-family=\"Cambria Math\" font-size=\"44\">x=(-b±√(b²-4ac))/(2a)</text></svg>");
        var emfPath = OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 360, 112);
        var pngDataUrl = CreatePngDataUrl("omml-tab-format-roundtrip", 360, 112);
        File.WriteAllBytes(
            pngPath,
            Convert.FromBase64String(
                pngDataUrl.Substring(pngDataUrl.IndexOf(',') + 1)));

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.InlineShape? oleShape = null;
        Word.Range? oleRange = null;
        Word.Range? referenceInsertion = null;
        Word.Fields? referenceFields = null;
        Word.Field? externalReference = null;
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
            var insertionPosition = document.Content.End - 1;
            application.Selection.SetRange(insertionPosition, insertionPosition);
            var ommlSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                insertionPosition,
                insertionPosition,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(ommlSession, QuadraticFormulaMathMl());
            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "OMML→OLE→OMML source");

            referenceInsertion = document.Range(
                document.Content.End - 1,
                document.Content.End - 1);
            referenceInsertion.Text = "External equation reference: ";
            referenceInsertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            referenceFields = referenceInsertion.Fields;
            object referenceType = Word.WdFieldType.wdFieldEmpty;
            object referenceCode =
                $"REF {WordEquationNumbering.NativeNumberBookmarkName(formulaId)} \\h";
            object preserveFormatting = true;
            externalReference = referenceFields.Add(
                referenceInsertion,
                ref referenceType,
                ref referenceCode,
                ref preserveFormatting);
            externalReference.Update();
            AssertExternalEquationReference(document, formulaId, "initial OMML");
            Release(externalReference); externalReference = null;
            Release(referenceFields); referenceFields = null;
            Release(referenceInsertion); referenceInsertion = null;
            TraceOmmlOleRoundtripParagraphs(document, "before OMML→OLE");
            AssertEqual(1, document.Tables.Count,
                "The source numbered OMML must own exactly one direct-SEQ 1x3 table.");
            AssertEqual(0, document.Frames.Count,
                "The source numbered OMML unexpectedly owns a hidden caption Frame.");

            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException("OMML roundtrip source bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var sourceMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException("OMML roundtrip source metadata is missing.");
            var oleSession = CreateNumberedOleRoundtripSession(
                formulaId,
                document.FullName,
                equationRange.Start,
                equationRange.End,
                sourceMetadata);
            service.ReplaceOle(oleSession, pngPath, emfPath);
            Release(equationRange); equationRange = null;
            Release(bookmark); bookmark = null;

            AssertVisualTeXNumberedTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "OMML→VisualTeX OLE tab conversion");
            TraceOmmlOleRoundtripParagraphs(document, "after OMML→OLE");
            AssertEqual(0, document.Tables.Count,
                "OMML→VisualTeX OLE left the native OMML 1x3 table behind.");
            AssertEqual(1, document.Frames.Count,
                "OMML→VisualTeX OLE must own exactly one clipped native SEQ caption Frame.");
            AssertExternalEquationReference(document, formulaId, "OMML→VisualTeX OLE");

            oleShape = FindVisualTeXOleByFormulaIdForNumberToggle(document, formulaId);
            var oleMetadata = WordFormulaMetadataReader.TryRead(oleShape)
                ?? throw new InvalidDataException("Converted VisualTeX OLE lost metadata.");
            oleRange = oleShape.Range;
            var returnSession = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                oleRange.Start,
                oleRange.End,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                oleMetadata);
            service.ReplaceOmml(returnSession, QuadraticFormulaMathMl());
            Release(oleRange); oleRange = null;
            Release(oleShape); oleShape = null;
            FinalizeNumberedOmmlShapesAcrossOfficeTurns(
                document,
                expectedFormulaCount: 1,
                context: "VisualTeX OLE→OMML native #SEQ host");

            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "VisualTeX OLE→OMML tab conversion");
            AssertEqual(1, document.Tables.Count,
                "VisualTeX OLE→OMML did not restore exactly one direct-SEQ 1x3 table.");
            AssertEqual(0, document.Frames.Count,
                "VisualTeX OLE→OMML left the OLE clipped caption Frame behind.");
            AssertExternalEquationReference(document, formulaId, "VisualTeX OLE→OMML");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertNumberedOmmlTabHost(
                document,
                formulaId,
                updateReference: true,
                context: "OMML tab format roundtrip save/reopen");
            AssertEqual(1, document.Tables.Count,
                "OMML tab roundtrip save/reopen did not retain exactly one 1x3 table.");
            AssertEqual(0, document.Frames.Count,
                "OMML tab roundtrip save/reopen recreated a hidden caption Frame.");
            AssertExternalEquationReference(document, formulaId, "save/reopen");

            Console.WriteLine(
                "OMML↔VisualTeX OLE numbered roundtrip acceptance passed: OMML used the minimal direct-SEQ 1x3 host with no hidden caption Frame, OLE used its clipped SEQ caption plus center/right-tab REF row, and returning to OMML removed that Frame while preserving dynamic numbering, body references and save/reopen.");
        }
        finally
        {
            Release(externalReference);
            Release(referenceFields);
            Release(referenceInsertion);
            Release(oleRange);
            Release(oleShape);
            Release(equationRange);
            Release(bookmark);
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

    private static void TraceOmmlOleRoundtripParagraphs(
        Word.Document document,
        string context)
    {
        Word.Paragraphs? paragraphs = null;
        try
        {
            paragraphs = document.Paragraphs;
            Console.WriteLine($"  [OMML↔OLE paragraphs] {context}: count={paragraphs.Count}");
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Word.Paragraph? paragraph = null;
                Word.Range? range = null;
                Word.OMaths? maths = null;
                Word.InlineShapes? shapes = null;
                Word.Fields? fields = null;
                Word.Frames? frames = null;
                try
                {
                    paragraph = paragraphs[index];
                    range = paragraph.Range;
                    maths = range.OMaths;
                    shapes = range.InlineShapes;
                    fields = range.Fields;
                    frames = range.Frames;
                    var codes = string.Join(",", (range.Text ?? string.Empty)
                        .Select(character => $"U+{(int)character:X4}"));
                    Console.WriteLine(
                        $"    p#{index}={range.Start}:{range.End} maths={maths.Count} shapes={shapes.Count} fields={fields.Count} frames={frames.Count} codes={codes}");
                }
                finally
                {
                    Release(frames);
                    Release(fields);
                    Release(shapes);
                    Release(maths);
                    Release(range);
                    Release(paragraph);
                }
            }
        }
        finally { Release(paragraphs); }
    }

    private static void AssertExternalEquationReference(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Range? content = null;
        Word.Range? ownerRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? code = null;
        Word.Range? result = null;
        try
        {
            content = document.Content;
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Numbering owner disappeared before checking the external reference after {context}.");
            fields = content.Fields;
            var expectedCode =
                "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            for (var index = 1; index <= fields.Count; index++)
            {
                Release(result); result = null;
                Release(code); code = null;
                Release(field); field = fields[index];
                code = field.Code;
                if ((code.Text ?? string.Empty).IndexOf(
                        expectedCode,
                        StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                field.Update();
                result = field.Result;
                if (result.Start >= ownerRange.Start && result.End <= ownerRange.End)
                    continue;
                var text = (result.Text ?? string.Empty).Trim();
                AssertTrue(
                    !string.IsNullOrWhiteSpace(text)
                    && text.IndexOf("Error", StringComparison.OrdinalIgnoreCase) < 0
                    && text.IndexOf("未找到", StringComparison.OrdinalIgnoreCase) < 0,
                    $"External equation reference failed after {context}: '{text}'.");
                return;
            }
            throw new InvalidDataException(
                $"External equation reference field disappeared after {context}.");
        }
        finally
        {
            Release(result);
            Release(code);
            Release(field);
            Release(fields);
            Release(ownerRange);
            Release(content);
        }
    }

    private static OfficeSessionDocument CreateNumberedOleRoundtripSession(
        string formulaId,
        string documentId,
        int start,
        int end,
        FormulaMetadata originalMetadata)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "edit",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = documentId,
            SourceObjectId = WordRangeReference(start, end),
            Title = "VisualTeX OMML tab format roundtrip",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.NativeOleMode,
            Numbered = true,
            FontSizePt = 14,
            OriginalMetadata = originalMetadata,
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                },
            },
            ExportResult = new OfficeExportDocument
            {
                Width = 360,
                Height = 112,
                Baseline = 84,
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
    }
}

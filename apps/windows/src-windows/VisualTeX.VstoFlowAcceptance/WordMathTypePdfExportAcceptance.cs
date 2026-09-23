using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeWmfWindowPdfAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);

        const string mathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow>"
            + "<mrow><mi>y</mi><mo>+</mo><mn>2</mn></mrow></mfrac>"
            + "<mo>=</mo><msqrt><mrow><mi>a</mi><mo>+</mo><mi>b</mi></mrow></msqrt>"
            + "</math>";
        const string latex =
            @"\frac{x+1}{y+2}=\sqrt{a+b}";
        const float widthPt = 135f;
        const float heightPt = 48f;

        static byte[] RemoveLogicalWindowRecords(byte[] source)
        {
            const int firstRecordOffset = 40;
            const ushort setWindowOrg = 0x020B;
            const ushort setWindowExt = 0x020C;
            var removals = new List<(int Offset, int Size)>();
            for (var offset = firstRecordOffset;
                 offset + 6 <= source.Length;)
            {
                var words = BitConverter.ToUInt32(source, offset);
                if (words < 3 || words > int.MaxValue / 2)
                    throw new InvalidDataException(
                        "WMF regression fixture contains an invalid record length.");
                var size = checked((int)words * 2);
                if (offset + size > source.Length)
                    throw new InvalidDataException(
                        "WMF regression fixture contains a truncated record.");
                var function = BitConverter.ToUInt16(source, offset + 4);
                if (function is setWindowOrg or setWindowExt)
                    removals.Add((offset, size));
                offset += size;
                if (function == 0) break;
            }

            AssertEqual(
                2,
                removals.Count,
                "WMF regression source does not contain exactly one SETWINDOWORG and one SETWINDOWEXT record.");
            AssertTrue(
                removals.All(item => item.Size == 10),
                "WMF regression logical-window records do not use the canonical five-word form.");

            var result = new byte[source.Length - 20];
            var sourceOffset = 0;
            var targetOffset = 0;
            foreach (var removal in removals.OrderBy(item => item.Offset))
            {
                var count = removal.Offset - sourceOffset;
                Buffer.BlockCopy(
                    source,
                    sourceOffset,
                    result,
                    targetOffset,
                    count);
                sourceOffset = removal.Offset + removal.Size;
                targetOffset += count;
            }
            Buffer.BlockCopy(
                source,
                sourceOffset,
                result,
                targetOffset,
                source.Length - sourceOffset);

            var declaredWords = BitConverter.ToUInt32(result, 28);
            if (declaredWords < 10)
                throw new InvalidDataException(
                    "WMF regression source has an invalid METAHEADER size.");
            var newDeclaredWords = declaredWords - 10;
            result[28] = (byte)newDeclaredWords;
            result[29] = (byte)(newDeclaredWords >> 8);
            result[30] = (byte)(newDeclaredWords >> 16);
            result[31] = (byte)(newDeclaredWords >> 24);
            return result;
        }

        var svgPath =
            Path.Combine(
                artifactRoot,
                "mathtype-pdf-window-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"360\" height=\"128\" viewBox=\"0 0 360 128\">"
            + "<text x=\"8\" y=\"86\" font-family=\"Cambria Math\" font-size=\"52\">(x+1)/(y+2)=sqrt(a+b)</text>"
            + "</svg>");
        var emfPath =
            OfficeOlePreview.CreateVectorEmfFromSvg(
                svgPath,
                360,
                128);

        var prepared =
            WordFormulaService
                .PrepareStandaloneMathTypeOleData(
                    mathMl,
                    inline: true,
                    fontSizePt: 12,
                    latex);
        var sourceWmf =
            MathTypeWordOpenXml
                .ConvertEnhancedMetafileToPlaceableWmf(
                    emfPath,
                    widthPt,
                    heightPt);
        var rawWmf =
            RemoveLogicalWindowRecords(
                sourceWmf);
        var expectedFixedWmf =
            WordFormulaService
                .EnsureMathTypePlaceableWmfWindowRecords(
                    rawWmf);

        AssertEqual(
            rawWmf.Length + 20,
            expectedFixedWmf.Length,
            "The raw VisualTeX EMF→WMF fixture unexpectedly already contains both logical-window records.");

        // Product invariant under test: every new MathType Flat OPC package must
        // normalize its external WMF before Word ever sees the Equation.DSMT4.
        var flatOpc =
            MathTypeWordOpenXml
                .CreateWithPlaceableWmf(
                    prepared.CompoundFile,
                    rawWmf,
                    widthPt,
                    heightPt,
                    display: false);
        var fragment =
            MathTypeWordOpenXml.Read(
                flatOpc);
        AssertTrue(
            fragment.PreviewWmf.SequenceEqual(
                expectedFixedWmf),
            "CreateWithPlaceableWmf did not persist SETWINDOWORG + SETWINDOWEXT.");
        AssertTrue(
            WordFormulaService
                .EnsureMathTypePlaceableWmfWindowRecords(
                    fragment.PreviewWmf)
                .SequenceEqual(
                    fragment.PreviewWmf),
            "The persisted MathType preview still needs logical-window repair.");

        File.WriteAllBytes(
            Path.Combine(
                artifactRoot,
                "raw-preview-without-window-records.wmf"),
            rawWmf);
        File.WriteAllBytes(
            Path.Combine(
                artifactRoot,
                "normalized-preview-with-window-records.wmf"),
            fragment.PreviewWmf);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? insertion = null;
        Word.InlineShape? shape = null;
        try
        {
            application =
                CreateFreshAcceptanceAutomationWord(
                    artifactRoot);
            application.Visible = false;
            document =
                application.Documents.Add(
                    Visible: false);
            document.Activate();

            insertion =
                document.Range(
                    0,
                    0);
            insertion.InsertXML(
                flatOpc);

            AssertEqual(
                1,
                document.InlineShapes.Count,
                "Normalized MathType Flat OPC did not materialize one Equation.DSMT4 object.");
            shape =
                document.InlineShapes[1];
            AssertTrue(
                MathTypeOleInterop.IsMathTypeOle(
                    shape),
                "Normalized PDF fixture is no longer Equation.DSMT4.");
            AssertEqual(
                MathTypeMtefCodec.SemanticSignature(
                    mathMl),
                MathTypeMtefCodec.SemanticSignature(
                    MathTypeOleStorage.ReadMathMl(
                        shape)),
                "WMF normalization changed MathType equation semantics.");

            var screenPreview =
                ReadInlineShapeEnhancedMetafile(
                    shape);
            var ink =
                DescribeEmfInkBounds(
                    screenPreview);
            AssertTrue(
                !string.Equals(
                    ink,
                    "empty",
                    StringComparison.Ordinal),
                "Normalized MathType formula became blank in Word.");

            var docxPath =
                Path.Combine(
                    artifactRoot,
                    "MathType-WMF-Window-PDF-Regression.docx");
            var pdfPath =
                Path.Combine(
                    artifactRoot,
                    "MathType-WMF-Window-PDF-Regression.pdf");
            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat
                    .wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.ExportAsFixedFormat(
                pdfPath,
                Word.WdExportFormat
                    .wdExportFormatPDF,
                OpenAfterExport: false,
                OptimizeFor:
                    Word.WdExportOptimizeFor
                        .wdExportOptimizeForPrint,
                Range:
                    Word.WdExportRange
                        .wdExportAllDocument,
                Item:
                    Word.WdExportItem
                        .wdExportDocumentContent,
                IncludeDocProps: true,
                KeepIRM: true,
                CreateBookmarks:
                    Word.WdExportCreateBookmarks
                        .wdExportCreateNoBookmarks,
                DocStructureTags: true,
                BitmapMissingFonts: true,
                UseISO19005_1: false);

            AssertTrue(
                File.Exists(
                    pdfPath)
                && new FileInfo(
                    pdfPath).Length > 0,
                "Word did not emit the MathType PDF regression artifact.");

            Console.WriteLine(
                $"[MATHTYPE WMF WINDOW PDF PASS] rawWmf={rawWmf.Length}; normalizedWmf={fragment.PreviewWmf.Length}; screenInk={ink}; pdf={pdfPath}");
        }
        finally
        {
            Release(shape);
            Release(insertion);
            if (document is not null)
            {
                try
                {
                    document.Close(
                        Word.WdSaveOptions
                            .wdDoNotSaveChanges);
                }
                catch { }
            }
            Release(document);
            try
            {
                QuitWordApplicationIfOwned(
                    application);
            }
            catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}

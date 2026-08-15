using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeAddOleFromCfbAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourceDocx = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "..",
            "artifacts", "mathtype-native-editor",
            "VisualTeX-MathType7-NativeEditor-5f04f8b3545e444a824705446e314ba1.docx"));
        if (!File.Exists(sourceDocx))
            throw new FileNotFoundException("Synchronized MathType source fixture is missing.", sourceDocx);
        var target = Path.Combine(artifactRoot, "MathType7-Proxy-Paste.docx");
        File.Copy(sourceDocx, target, overwrite: true);

        const string targetMathMl =
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
            + "<mfrac><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow><mi>y</mi></mfrac>"
            + "</math>";
        var previewSvg = Path.Combine(artifactRoot, "proxy-preview.svg");
        File.WriteAllText(
            previewSvg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"180\" height=\"64\" viewBox=\"0 0 180 64\"><text x=\"4\" y=\"44\" font-size=\"36\">VisualTeX</text></svg>");
        var previewEmf = OfficeOlePreview.CreateVectorEmfFromSvg(previewSvg, 180, 64);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Document? stagingDocument = null;
        Word.InlineShape? sourceShape = null;
        Word.InlineShape? stagingShape = null;
        Word.InlineShape? pastedShape = null;
        Word.Range? stagingInsertion = null;
        Word.Range? insertion = null;
        Word.OLEFormat? format = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(target, ReadOnly: false, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "Proxy paste fixture must begin with one MathType OLE.");
            sourceShape = document.InlineShapes[1];
            var originalPreview = ReadInlineShapeEnhancedMetafile(sourceShape);
            stagingDocument = application.Documents.Add(Visible: false);

            Console.WriteLine("[MathType proxy paste 1/5] Capturing the live Word OLE IDataObject and rewriting its CFB offline...");
            var before = SnapshotMathTypeProcessIds();
            using (var transaction = MathTypeOleStorage.BeginClipboardTransaction(sourceShape))
            {
                var rewritten = MathTypeOleStorage.RewriteMathTypeCompoundFile(
                    transaction.CompoundFile,
                    targetMathMl,
                    inline: true);
                var cachedProbe = MathTypeOleStorage.AddEnhancedMetafilePresentationCache(
                    rewritten.CompoundFile,
                    previewEmf);
                var cachedEntries = MathTypeOleStorage.ListCompoundFileEntries(cachedProbe);
                Console.WriteLine(
                    "  pre-Word Windows OLE cache entries="
                    + string.Join(", ", cachedEntries.Select(name =>
                        name.Replace("\u0001", "\\x01").Replace("\u0002", "\\x02").Replace("\u0003", "\\x03"))));
                var cachedPreview = MathTypeOleStorage.ReadEnhancedMetafilePresentationCache(cachedProbe);
                var preWordPreviewDifference = MeasureEmfPixelDifference(
                    File.ReadAllBytes(previewEmf),
                    cachedPreview);
                Console.WriteLine(
                    $"  pre-Word OLE cache preview pixel difference={preWordPreviewDifference:0.0000}.");
                AssertTrue(
                    preWordPreviewDifference < 0.08,
                    "Windows OLE cache did not persist the VisualTeX EMF before Word paste.");
                transaction.SetReplacementClipboard(
                    cachedProbe,
                    previewEmf,
                    preferEmbedSource: false,
                    standaloneExternal: true);

                Console.WriteLine("[MathType proxy paste 2/5] Pasting the independent OLE source into a blank staging document...");
                stagingInsertion = stagingDocument.Range(0, 0);
                stagingInsertion.PasteSpecial(
                    Link: false,
                    DataType: Word.WdPasteDataType.wdPasteOLEObject,
                    Placement: Word.WdOLEPlacement.wdInLine,
                    DisplayAsIcon: false);
            }
            Thread.Sleep(500);
            var after = SnapshotMathTypeProcessIds();
            var started = after.Except(before).ToArray();
            Console.WriteLine(
                "  MathType processes started during capture/rewrite/paste: "
                + (started.Length == 0 ? "none" : string.Join(", ", started)));
            if (started.Length > 0)
                throw new InvalidDataException(
                    "Proxy PasteSpecial unexpectedly activated MathType: "
                    + string.Join(", ", started));

            Console.WriteLine("[MathType proxy paste 3/5] Verifying the staging OLE preview and MTEF...");
            AssertEqual(1, stagingDocument.InlineShapes.Count,
                "Standalone proxy did not create exactly one staging inline OLE object.");
            stagingShape = stagingDocument.InlineShapes[1];
            format = stagingShape.OLEFormat;
            var stagingProgId = format.ProgID ?? string.Empty;
            Console.WriteLine(
                $"  staging ProgID='{stagingProgId}', type={stagingShape.Type}, size={stagingShape.Width:0.##}x{stagingShape.Height:0.##}.");
            AssertEqual("Equation.DSMT4", stagingProgId,
                "Staging PasteSpecial changed the MathType OLE ProgID.");
            Release(format); format = null;
            // Match the physical box to the VisualTeX preview before comparing
            // pixels. Word initially inherits the source MathType object's old
            // width/height, which can heavily distort a differently-shaped edit.
            stagingShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            stagingShape.Width = 135f;
            stagingShape.Height = 48f;
            var expectedPreview = File.ReadAllBytes(previewEmf);
            var stagingPreview = ReadInlineShapeEnhancedMetafile(stagingShape);
            var stagingPreviewDifference = MeasureEmfPixelDifference(
                expectedPreview,
                stagingPreview);
            var originalToExpectedDifference = MeasureEmfPixelDifference(
                expectedPreview,
                originalPreview);
            var stagingToOriginalDifference = MeasureEmfPixelDifference(
                stagingPreview,
                originalPreview);
            Console.WriteLine(
                $"  staging->VisualTeX diff={stagingPreviewDifference:0.0000}, "
                + $"original->VisualTeX diff={originalToExpectedDifference:0.0000}, "
                + $"staging->original diff={stagingToOriginalDifference:0.0000}.");
            AssertTrue(
                stagingPreviewDifference + 0.02 < originalToExpectedDifference,
                "Blank staging Word document remains closer to the old MathType preview than the VisualTeX replacement preview.");
            var stagingCaptured = MathTypeOleStorage.CaptureCompoundFile(stagingShape);
            var stagingLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(stagingCaptured))
                .Trim();
            AssertMathTypeLatexEquivalent(
                "\\frac{x+1}{y}",
                stagingLatex,
                "Staging MathType OLE contains the wrong MTEF source.");

            Console.WriteLine("[MathType proxy paste 4/5] Copying the validated staging OLE back into the original document...");
            var stagingRange = stagingShape.Range;
            try { stagingRange.Copy(); }
            finally { Release(stagingRange); }
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.InsertParagraphBefore();
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.PasteSpecial(
                Link: false,
                DataType: Word.WdPasteDataType.wdPasteOLEObject,
                Placement: Word.WdOLEPlacement.wdInLine,
                DisplayAsIcon: false);
            Thread.Sleep(300);
            AssertEqual(2, document.InlineShapes.Count,
                "Staging-to-original copy did not create exactly one additional inline OLE object.");
            pastedShape = document.InlineShapes[2];
            format = pastedShape.OLEFormat;
            var progId = format.ProgID ?? string.Empty;
            Console.WriteLine(
                $"  final ProgID='{progId}', type={pastedShape.Type}, size={pastedShape.Width:0.##}x{pastedShape.Height:0.##}.");
            AssertEqual("Equation.DSMT4", progId,
                "Staging-to-original copy changed the MathType OLE ProgID.");
            pastedShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
            pastedShape.Width = 135f;
            pastedShape.Height = 48f;
            var pastedPreview = ReadInlineShapeEnhancedMetafile(pastedShape);
            var previewDifference = MeasureEmfPixelDifference(expectedPreview, pastedPreview);
            var pastedToOriginalDifference = MeasureEmfPixelDifference(pastedPreview, originalPreview);
            Console.WriteLine(
                $"  final->VisualTeX diff={previewDifference:0.0000}, "
                + $"final->original diff={pastedToOriginalDifference:0.0000}.");
            AssertTrue(
                previewDifference + 0.02 < originalToExpectedDifference,
                "Staging-to-original copy regressed toward the old MathType preview instead of preserving the VisualTeX replacement preview.");
            var captured = MathTypeOleStorage.CaptureCompoundFile(pastedShape);
            AssertTrue(MathTypeOleStorage.LooksLikeMathTypeCompoundFile(captured),
                "Proxy-pasted Word OLE is no longer a MathType Compound File immediately after paste.");
            var directLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(captured))
                .Trim();
            Console.WriteLine($"  immediate post-paste MTEF={directLatex}.");
            var afterSize = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(MathTypeOleStorage.CaptureCompoundFile(pastedShape)))
                .Trim();
            Console.WriteLine($"  after Width/Height MTEF={afterSize}.");
            var shapeRange = pastedShape.Range;
            try
            {
                var shapeFont = shapeRange.Font;
                try { shapeFont.Position = 0; }
                finally { Release(shapeFont); }
            }
            finally { Release(shapeRange); }
            var afterPosition = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(MathTypeOleStorage.CaptureCompoundFile(pastedShape)))
                .Trim();
            Console.WriteLine($"  after Font.Position MTEF={afterPosition}.");
            AssertMathTypeLatexEquivalent(
                "\\frac{x+1}{y}",
                directLatex,
                "Proxy-pasted MathType OLE contains the wrong MTEF source.");

            Console.WriteLine("[MathType proxy paste 5/5] Saving/reopening Word and validating the pasted CFB directly...");
            document.Save();
            stagingDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(stagingShape); stagingShape = null;
            Release(stagingInsertion); stagingInsertion = null;
            Release(stagingDocument); stagingDocument = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(format); format = null;
            Release(sourceShape); sourceShape = null;
            Release(pastedShape); pastedShape = null;
            Release(insertion); insertion = null;
            Release(document); document = null;
            application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(application); application = null;

            var embeddings = ReadDocxOleEmbeddings(target);
            AssertEqual(2, embeddings.Count,
                "Saved proxy-paste document did not retain two embedded OLE objects.");
            AssertTrue(
                embeddings.Any(bytes =>
                {
                    try
                    {
                        if (!MathTypeOleStorage.LooksLikeMathTypeCompoundFile(bytes)) return false;
                        var latex = MathMlToLatexConverter.Convert(MathTypeOleStorage.ReadMathMl(bytes)).Trim();
                        return NormalizeMathTypeLatex(latex) == NormalizeMathTypeLatex("\\frac{x+1}{y}");
                    }
                    catch { return false; }
                }),
                "Saved proxy-paste document does not contain the rewritten fraction CFB.");
            Console.WriteLine(
                "MathType proxy-paste acceptance passed: Word embedded a VisualTeX-rewritten Equation.DSMT4 CFB through a custom IDataObject without starting MathType.");
        }
        finally
        {
            Release(format);
            Release(pastedShape);
            Release(stagingShape);
            Release(sourceShape);
            Release(stagingInsertion);
            Release(insertion);
            if (stagingDocument is not null)
            {
                try { stagingDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(stagingDocument);
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

    private static List<byte[]> ReadDocxOleEmbeddings(string path)
    {
        using var archive = System.IO.Compression.ZipFile.OpenRead(path);
        var result = new List<byte[]>();
        foreach (var entry in archive.Entries.Where(entry =>
                     entry.FullName.StartsWith("word/embeddings/oleObject", StringComparison.OrdinalIgnoreCase)))
        {
            using var stream = entry.Open();
            using var memory = new MemoryStream();
            stream.CopyTo(memory);
            result.Add(memory.ToArray());
        }
        return result;
    }

    private static string NormalizeMathTypeLatex(string value) =>
        value.Replace(" ", string.Empty).Replace("{", string.Empty).Replace("}", string.Empty);
}

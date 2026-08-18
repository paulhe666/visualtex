using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeDirectPasteSymbolsAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var previewSvg = Path.Combine(artifactRoot, "direct-paste-symbols-preview.svg");
        File.WriteAllText(
            previewSvg,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"260\" height=\"70\" viewBox=\"0 0 260 70\"><text x=\"4\" y=\"48\" font-size=\"36\">MathType symbols</text></svg>");
        var previewEmf = OfficeOlePreview.CreateVectorEmfFromSvg(previewSvg, 260, 70);

        var cases = new[]
        {
            new
            {
                Name = "quadratic",
                MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>"
            },
            new
            {
                Name = "gaussian",
                MathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><msubsup><mo>∫</mo><mrow><mo>−</mo><mo>∞</mo></mrow><mo>∞</mo></msubsup><msup><mi mathvariant=\"normal\">e</mi><mrow><mo>−</mo><msup><mi>x</mi><mn>2</mn></msup></mrow></msup><mi mathvariant=\"normal\">d</mi><mi>x</mi><mo>=</mo><msqrt><mi>π</mi></msqrt></math>"
            },
        };

        foreach (var testCase in cases)
        {
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Document? reopened = null;
            Word.Range? insertion = null;
            Word.InlineShape? shape = null;
            try
            {
                var mathTypeBefore = SnapshotMathTypeProcessIds();
                AssertEqual(0, mathTypeBefore.Count,
                    $"{testCase.Name}: direct-paste acceptance requires MathType.exe to be absent before insertion.");

                var generated = MathTypeMtefCodec.CreateEquationNative(testCase.MathMl, inline: false);
                var compound = MathTypeOleStorage.CreateStandaloneCompoundFile(generated);
                var expected = MathTypeMtefCodec.SemanticSignature(testCase.MathMl);
                AssertEqual(expected,
                    MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(compound)),
                    $"{testCase.Name}: generated standalone CFB is already semantically wrong before Word.");

                application = CreateWordApplication(visible: false);
                document = application.Documents.Add(Visible: false);
                insertion = document.Range(0, 0);
                using (var transaction = MathTypeOleStorage.BeginStandaloneClipboardTransaction(
                           compound,
                           previewEmf,
                           195f,
                           52.5f))
                {
                    insertion.PasteSpecial(
                        Link: false,
                        DataType: Word.WdPasteDataType.wdPasteOLEObject,
                        Placement: Word.WdOLEPlacement.wdInLine,
                        DisplayAsIcon: false);
                    AssertTrue(transaction.ReplacementStorageWriteCount > 0,
                        $"{testCase.Name}: Word never requested the standalone MathType CFB storage.");
                }

                AssertEqual(1, document.InlineShapes.Count,
                    $"{testCase.Name}: direct PasteSpecial did not create exactly one OLE object.");
                shape = document.InlineShapes[1];
                AssertTrue(MathTypeOleInterop.IsMathTypeOle(shape),
                    $"{testCase.Name}: direct PasteSpecial did not create Equation.DSMT4.");
                var startedAfterPaste = SnapshotMathTypeProcessIds().Except(mathTypeBefore).ToArray();
                AssertEqual(0, startedAfterPaste.Length,
                    $"{testCase.Name}: VisualTeX standalone IDataObject PasteSpecial started MathType.exe.");
                var livePreview = ReadInlineShapeEnhancedMetafile(shape);
                var liveInk = DescribeEmfInkBounds(livePreview);
                AssertTrue(!string.Equals(liveInk, "empty", StringComparison.Ordinal),
                    $"{testCase.Name}: direct PasteSpecial created a visually blank live Word OLE.");
                Console.WriteLine($"{testCase.Name} live preview={liveInk}; bytes={livePreview.Length}.");
                var immediate = MathTypeOleStorage.ReadMathMl(shape);
                Console.WriteLine($"{testCase.Name} immediate MathML={immediate}");
                AssertEqual(expected,
                    MathTypeMtefCodec.SemanticSignature(immediate),
                    $"{testCase.Name}: Word changed the MTEF during direct PasteSpecial.");

                var path = Path.Combine(artifactRoot, $"direct-paste-{testCase.Name}.docx");
                document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdSaveChanges);
                Release(shape); shape = null;
                Release(insertion); insertion = null;
                Release(document); document = null;

                reopened = application.Documents.Open(path, ReadOnly: true, Visible: false);
                AssertEqual(1, reopened.InlineShapes.Count,
                    $"{testCase.Name}: save/reopen changed the OLE count.");
                shape = reopened.InlineShapes[1];
                var reopenedPreview = ReadInlineShapeEnhancedMetafile(shape);
                var reopenedInk = DescribeEmfInkBounds(reopenedPreview);
                AssertTrue(!string.Equals(reopenedInk, "empty", StringComparison.Ordinal),
                    $"{testCase.Name}: save/reopen produced a visually blank MathType OLE.");
                var persisted = MathTypeOleStorage.ReadMathMl(shape);
                Console.WriteLine($"{testCase.Name} reopened MathML={persisted}");
                AssertEqual(expected,
                    MathTypeMtefCodec.SemanticSignature(persisted),
                    $"{testCase.Name}: save/reopen changed the directly pasted MathType MTEF.");
                var mathTypeAfter = SnapshotMathTypeProcessIds().Except(mathTypeBefore).ToArray();
                AssertEqual(0, mathTypeAfter.Length,
                    $"{testCase.Name}: standalone IDataObject insertion/preview/save-reopen started MathType.exe.");
                Console.WriteLine($"{testCase.Name}: direct CFB PasteSpecial stayed visible, started no MathType process, and save/reopen preserved the exact semantic signature.");
            }
            finally
            {
                Release(shape);
                Release(insertion);
                if (reopened is not null)
                {
                    try { reopened.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                }
                Release(reopened);
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

        Console.WriteLine("MathType direct-paste symbol acceptance passed for quadratic and Gaussian formulas.");
    }
}

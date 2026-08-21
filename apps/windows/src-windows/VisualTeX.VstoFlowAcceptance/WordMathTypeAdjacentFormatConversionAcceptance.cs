using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeAdjacentFormatConversionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        const string firstLatex = @"\frac{d}{dx}a=b";
        const string secondLatex = @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}";
        const string firstMathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mi>d</mi><mrow><mi>d</mi><mi>x</mi></mrow></mfrac><mi>a</mi><mo>=</mo><mi>b</mi></math>";
        const string secondMathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";

        var svgPath = Path.Combine(artifactRoot, "adjacent-format-preview.svg");
        File.WriteAllText(
            svgPath,
            "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"96\" viewBox=\"0 0 240 96\"><text x=\"4\" y=\"64\" font-family=\"Times New Roman\" font-size=\"48\">x+1</text></svg>");
        var emfPath = VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview.CreateVectorEmfFromSvg(svgPath, 240, 96);

        ProbeAdjacentOmmlGroupInsertion();
        RunDirection(targetVisualTeX: true);
        RunDirection(targetVisualTeX: false);
        Console.WriteLine("[ADJACENT FORMAT PASS] zero-gap MathType OLE pair survived MT→VisualTeX and MT→OMML without cross-formula contamination.");

        void ProbeAdjacentOmmlGroupInsertion()
        {
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Range? target = null;
            WordOmmlConverter.BatchSource? source = null;
            IReadOnlyList<Word.Range>? inserted = null;
            try
            {
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                document.Content.Text = "AB";
                var firstId = Guid.NewGuid().ToString("D");
                var secondId = Guid.NewGuid().ToString("D");
                source = WordOmmlConverter.CreateBatchSource(
                    application,
                    new[]
                    {
                        (FormulaId: firstId, MathMl: firstMathMl),
                        (FormulaId: secondId, MathMl: secondMathMl),
                    });
                target = document.Range(0, 2);
                inserted = source.InsertAdjacentInlineGroup(
                    application,
                    document,
                    target,
                    new[] { firstId, secondId });
                AssertEqual(2, inserted.Count,
                    "Adjacent OMML group paste did not return two equation ranges.");
                AssertEqual(2, document.OMaths.Count,
                    "Adjacent OMML group paste was normalized into one OMath by Word.");
                Console.WriteLine("[ADJACENT GROUP PROBE] one-shot sibling OMath paste retained 2 independent equations.");
            }
            finally
            {
                if (inserted is not null)
                    foreach (var range in inserted) Release(range);
                source?.Dispose();
                Release(target);
                try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(document);
                Release(application);
                ForceComCleanup();
            }
        }

        void RunDirection(bool targetVisualTeX)
        {
            var previousFormatAcceptance = Environment.GetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE");
            var previousTracePath = Environment.GetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH");
            var tracePath = Path.Combine(
                artifactRoot,
                targetVisualTeX ? "adjacent-mt-to-vt.trace.log" : "adjacent-mt-to-omml.trace.log");
            try { File.Delete(tracePath); } catch { }
            Word.Application? application = null;
            Word.Document? document = null;
            Word.Range? insertion = null;
            Word.InlineShape? first = null;
            Word.InlineShape? second = null;
            VisualTeX.WordVsto.ThisAddIn? addIn = null;
            Array custom = Array.Empty<object>();
            try
            {
                Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", "1");
                Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", tracePath);
                var mathTypeBaseline = SnapshotMathTypeProcessIds();
                application = CreateWordApplication(visible: false);
                document = application.Documents.Add();
                document.Activate();
                var service = new WordFormulaService(application);

                insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
                insertion.Select();
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession("inline", false, firstLatex),
                    firstMathMl,
                    emfPath,
                    createdObjectBookmarkName: "VTMT_ADJ_FMT_FIRST");
                Release(insertion); insertion = null;

                first = document.InlineShapes[1];
                var boundary = first.Range.End;
                insertion = document.Range(boundary, boundary);
                insertion.Select();
                service.InsertMathTypeOle(
                    CreateMathTypeCreateSession("inline", false, secondLatex),
                    secondMathMl,
                    emfPath,
                    createdObjectBookmarkName: "VTMT_ADJ_FMT_SECOND");
                Release(insertion); insertion = null;
                Release(first); first = null;

                AssertEqual(2, document.InlineShapes.Count,
                    "Zero-gap MathType setup did not retain exactly two OLE equations.");
                first = document.InlineShapes[1];
                second = document.InlineShapes[2];
                AssertEqual(first.Range.End, second.Range.Start,
                    "Zero-gap MathType setup unexpectedly inserted a separator between the two OLE equations.");
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(firstMathMl),
                    MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(first)),
                    "Zero-gap first MathType equation changed before batch conversion.");
                AssertEqual(
                    MathTypeMtefCodec.SemanticSignature(secondMathMl),
                    MathTypeMtefCodec.SemanticSignature(MathTypeOleStorage.ReadMathMl(second)),
                    "Zero-gap second MathType equation changed before batch conversion.");
                Release(first); first = null;
                Release(second); second = null;

                addIn = new VisualTeX.WordVsto.ThisAddIn();
                addIn.OnConnection(
                    application,
                    Extensibility.ext_ConnectMode.ext_cm_AfterStartup,
                    addIn,
                    ref custom);
                if (targetVisualTeX)
                    addIn.OnConvertMathTypeToVisualTeXDocument(new object());
                else
                    addIn.OnConvertMathTypeToOmmlDocument(new object());
                WaitForInstalledOmmlMathTypeConversion(
                    tracePath,
                    targetVisualTeX
                        ? "source=MathType target=VisualTeX"
                        : "source=MathType target=OMML",
                    mathTypeBaseline);
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(30));

                if (targetVisualTeX)
                {
                    AssertEqual(2, CountInstalledVisualTeXOleShapes(document),
                        "Zero-gap MT→VisualTeX did not create exactly two VisualTeX formulas.");
                    var expected = new[]
                    {
                        NormalizeStressLatex(MathMlToLatexConverter.Convert(firstMathMl)),
                        NormalizeStressLatex(MathMlToLatexConverter.Convert(secondMathMl)),
                    };
                    var seen = 0;
                    for (var index = 1; index <= document.InlineShapes.Count; index++)
                    {
                        Word.InlineShape? shape = null;
                        try
                        {
                            shape = document.InlineShapes[index];
                            if (!WordFormulaMetadataReader.IsNativeOle(shape)) continue;
                            var metadata = WordFormulaMetadataReader.TryRead(shape)
                                ?? throw new InvalidDataException($"Zero-gap VisualTeX target #{seen + 1} has no metadata.");
                            AssertEqual(expected[seen], NormalizeStressLatex(metadata.Latex ?? string.Empty),
                                $"Zero-gap MT→VisualTeX formula #{seen + 1} inherited content from its neighbor.");
                            seen++;
                        }
                        finally { Release(shape); }
                    }
                    AssertEqual(2, seen, "Zero-gap MT→VisualTeX did not inspect two targets.");
                }
                else
                {
                    AssertEqual(2, document.OMaths.Count,
                        "Zero-gap MT→OMML did not create exactly two OMath equations.");
                    var expectedSignatures = new[]
                    {
                        MathTypeMtefCodec.SemanticSignature(firstMathMl),
                        MathTypeMtefCodec.SemanticSignature(secondMathMl),
                    };
                    for (var index = 1; index <= 2; index++)
                    {
                        Word.OMath? math = null;
                        Word.Range? range = null;
                        try
                        {
                            math = document.OMaths[index];
                            range = math.Range;
                            var roundTrip = WordOmmlConverter.TransformOmmlToMathMl(range.WordOpenXML, display: false);
                            AssertEqual(expectedSignatures[index - 1], MathTypeMtefCodec.SemanticSignature(roundTrip),
                                $"Zero-gap MT→OMML formula #{index} inherited content from its neighbor.");
                        }
                        finally
                        {
                            Release(range);
                            Release(math);
                        }
                    }
                }
                AssertNoNewMathTypeProcess(
                    mathTypeBaseline,
                    targetVisualTeX ? "zero-gap MT→VisualTeX" : "zero-gap MT→OMML");
                Console.WriteLine($"[ADJACENT FORMAT] {(targetVisualTeX ? "MT→VisualTeX" : "MT→OMML")} passed.");
            }
            finally
            {
                Release(first);
                Release(second);
                Release(insertion);
                if (addIn is not null)
                {
                    try
                    {
                        addIn.OnDisconnection(
                            Extensibility.ext_DisconnectMode.ext_dm_UserClosed,
                            ref custom);
                    }
                    catch { }
                }
                try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                try { QuitWordApplicationIfOwned(application); } catch { }
                Release(document);
                Release(application);
                ForceComCleanup();
                Environment.SetEnvironmentVariable("VISUALTEX_FORMAT_CONVERSION_ACCEPTANCE", previousFormatAcceptance);
                Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", previousTracePath);
            }
        }
    }
}

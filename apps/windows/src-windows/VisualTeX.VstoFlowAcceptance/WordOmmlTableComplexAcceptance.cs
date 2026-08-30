using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class OmmlTableComplexPair
    {
        internal OmmlTableComplexPair(
            NumberedOmmlDisplayCase testCase,
            string numberedId,
            string plainId,
            string semanticSignature)
        {
            TestCase = testCase;
            NumberedId = numberedId;
            PlainId = plainId;
            SemanticSignature = semanticSignature;
        }
        internal NumberedOmmlDisplayCase TestCase { get; }
        internal string NumberedId { get; }
        internal string PlainId { get; }
        internal string SemanticSignature { get; }
    }

    private static void RunWordOmmlTableComplexAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The complex 1x3 acceptance refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(artifactRoot, "word-omml-1x3-complex.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureOmmlTableNumberPage(document);
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var pairs = new List<OmmlTableComplexPair>();

            foreach (var testCase in CreateNumberedOmmlDisplayCases())
            {
                var numberedId = Guid.NewGuid().ToString("D");
                InsertComplexDisplayFormula(
                    application,
                    document,
                    service,
                    testCase,
                    numberedId,
                    numbered: true);
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    numberedId,
                    "complex numbered " + testCase.Name);

                var plainId = Guid.NewGuid().ToString("D");
                InsertComplexDisplayFormula(
                    application,
                    document,
                    service,
                    testCase,
                    plainId,
                    numbered: false);
                var signature = AssertComplexTablePair(
                    application,
                    document,
                    testCase,
                    numberedId,
                    plainId,
                    "initial");
                pairs.Add(new OmmlTableComplexPair(
                    testCase,
                    numberedId,
                    plainId,
                    signature));
            }

            AssertEqual(pairs.Count, document.Tables.Count,
                "Complex 1x3 acceptance created a wrong table count.");
            document.Fields.Update();
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();

            AssertEqual(pairs.Count, document.Tables.Count,
                "Complex 1x3 save/reopen changed the table count.");
            foreach (var pair in pairs)
            {
                AssertOmmlTableNumberLifecyclePhase(
                    application,
                    document,
                    pair.NumberedId,
                    "complex reopen " + pair.TestCase.Name);
                var reopenedSignature = AssertComplexTablePair(
                    application,
                    document,
                    pair.TestCase,
                    pair.NumberedId,
                    pair.PlainId,
                    "save/reopen");
                AssertEqual(pair.SemanticSignature, reopenedSignature,
                    pair.TestCase.Name + ": save/reopen changed the semantic OMath structure.");
            }

            Console.WriteLine(
                "Word OMML complex 1x3 acceptance passed: fraction/radical, n-ary limits, nested fractions and matrices retained the same semantic OMath structure, native font size and horizontal glyph geometry as ordinary wdOMathDisplay references, while numbering stayed outside math across F9/save/reopen; Word's table-specific vertical character Range boxes were recorded separately rather than treated as glyph scaling.");
        }
        finally
        {
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

    private static string AssertComplexTablePair(
        Word.Application application,
        Word.Document document,
        NumberedOmmlDisplayCase testCase,
        string numberedId,
        string plainId,
        string phase)
    {
        Word.Range? numberedRange = null;
        Word.Range? plainRange = null;
        Word.Window? window = null;
        try
        {
            var numberedMetadata = WordOmmlFormulaStore.TryRead(document, numberedId)
                ?? throw new InvalidDataException(testCase.Name + ": numbered metadata missing.");
            var plainMetadata = WordOmmlFormulaStore.TryRead(document, plainId)
                ?? throw new InvalidDataException(testCase.Name + ": plain metadata missing.");
            numberedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document, numberedId, numberedMetadata);
            plainRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document, plainId, plainMetadata);

            AssertEqual(Word.WdOMathType.wdOMathDisplay, numberedRange.OMaths[1].Type,
                phase + " " + testCase.Name + ": numbered formula is not Display.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, plainRange.OMaths[1].Type,
                phase + " " + testCase.Name + ": plain formula is not Display.");
            AssertEqual(0, numberedRange.Fields.Count,
                phase + " " + testCase.Name + ": a number field entered the complex OMath.");
            AssertTrue(!WordOmmlConverter.HasVisualTeXNativeEquationNumber(
                    numberedRange.WordOpenXML ?? string.Empty),
                phase + " " + testCase.Name + ": retired #()/eqArr numbering entered the complex OMath.");

            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var math = (XNamespace)MathNamespace;
            var numberedEquation = XElement.Parse(
                WordOmmlConverter.ExtractSingleOMath(numberedRange.WordOpenXML),
                LoadOptions.PreserveWhitespace);
            var plainEquation = XElement.Parse(
                WordOmmlConverter.ExtractSingleOMath(plainRange.WordOpenXML),
                LoadOptions.PreserveWhitespace);
            var numberedSignature = BuildOmmlSemanticStructureSignature(numberedEquation);
            var plainSignature = BuildOmmlSemanticStructureSignature(plainEquation);
            AssertEqual(plainSignature, numberedSignature,
                phase + " " + testCase.Name + ": 1x3 changed semantic OMath structure.");
            foreach (var requirement in testCase.RequiredElements)
            {
                var count = numberedEquation
                    .DescendantsAndSelf(math + requirement.ElementName)
                    .Count();
                AssertTrue(count >= requirement.MinimumCount,
                    phase + " " + testCase.Name + $": expected m:{requirement.ElementName}>={requirement.MinimumCount}, found {count}.");
            }

            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(numberedRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(70);
            var numberedBox = ReadVisibleMathInkBox(
                document,
                window,
                numberedRange,
                phase + " " + testCase.Name + " numbered ink");
            window.ScrollIntoView(plainRange, ref scrollStart);
            Thread.Sleep(70);
            var plainBox = ReadVisibleMathInkBox(
                document,
                window,
                plainRange,
                phase + " " + testCase.Name + " plain ink");
            AssertMetricRatio(plainBox.Width, numberedBox.Width, 0.96, 1.04,
                phase + " " + testCase.Name + ": 1x3 changed formula width.");
            var numberedFontSize = numberedRange.Font.Size;
            var plainFontSize = plainRange.Font.Size;
            AssertNear(plainFontSize, numberedFontSize, 0.1f,
                phase + " " + testCase.Name + ": 1x3 changed the native OMath font size.");
            var heightRatio = plainBox.Height > 0
                ? numberedBox.Height / (double)plainBox.Height
                : 1.0;
            Console.WriteLine(
                $"  {phase} complex {testCase.Name}: numbered={numberedBox.Width}x{numberedBox.Height}px, plain={plainBox.Width}x{plainBox.Height}px, heightRangeRatio={heightRatio:0.###}, font={numberedFontSize:0.##}/{plainFontSize:0.##}pt, signature={numberedSignature.Length}.");
            return numberedSignature;
        }
        finally
        {
            Release(window);
            Release(plainRange);
            Release(numberedRange);
        }
    }
}

using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlNativeHashComplexFontAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-native-hash-complex-font.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        var formulaId = Guid.NewGuid().ToString("D");
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.SaveAs2(documentPath, Word.WdSaveFormat.wdFormatXMLDocument);
            document.Activate();
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.ContinuousId);
            var service = new WordFormulaService(application);
            var insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            try
            {
                application.Selection.SetRange(insertion.Start, insertion.End);
                var session = CreateNumberedOmmlTabSession(
                    formulaId,
                    document.FullName,
                    insertion.Start,
                    insertion.End,
                    NativeHashComplexLatex(14),
                    originalMetadata: null);
                session.FontSizePt = 14;
                service.InsertOmml(session, NativeHashComplexMathMl());
            }
            finally { Release(insertion); }

            AssertNativeHashComplexFontState(
                document,
                formulaId,
                expectedFontSize: 14f,
                context: "complex native hash at 14pt");

            ReplaceNativeHashComplexFont(
                application,
                document,
                service,
                formulaId,
                fontSize: 12f);
            AssertNativeHashComplexFontState(
                document,
                formulaId,
                expectedFontSize: 12f,
                context: "complex native hash at 12pt");

            ReplaceNativeHashComplexFont(
                application,
                document,
                service,
                formulaId,
                fontSize: 16f);
            AssertNativeHashComplexFontState(
                document,
                formulaId,
                expectedFontSize: 16f,
                context: "complex native hash at 16pt");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;
            ForceComCleanup();

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            WordEquationNumbering.UpdateEquationNumbers(document);
            AssertNativeHashComplexFontState(
                document,
                formulaId,
                expectedFontSize: 16f,
                context: "complex native hash 16pt save/reopen");

            Console.WriteLine(
                "Production OMML native #(SEQ) complex/font acceptance passed: matrix + Σ/Π/∫ survived 14→12→16pt editing and save/reopen as table/Shape-free wdOMathDisplay.");
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

    private static void ReplaceNativeHashComplexFont(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        string formulaId,
        float fontSize)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Complex numbered OMML {formulaId} lost metadata before {fontSize:0.##}pt edit.");
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    $"Complex numbered OMML {formulaId} lost its VTOMML bookmark before edit.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var session = CreateNumberedOmmlTabSession(
                formulaId,
                document.FullName,
                formulaRange.Start,
                formulaRange.End,
                NativeHashComplexLatex(fontSize),
                metadata);
            session.FontSizePt = fontSize;
            service.ReplaceOmml(session, NativeHashComplexMathMl());
        }
        finally
        {
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static void AssertNativeHashComplexFontState(
        Word.Document document,
        string formulaId,
        float expectedFontSize,
        string context)
    {
        UpdateNativeHashProductionFields(document, new[] { formulaId });
        AssertNativeHashProductionNumbers(
            document,
            new[] { formulaId },
            new[] { "1" },
            context + " numbering");

        var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
            ?? throw new InvalidDataException(context + ": metadata is missing.");
        AssertNear(expectedFontSize, (float)metadata.FontSizePt, 0.01f,
            context + ": persisted semantic font size differs.");

        Word.Bookmark? bookmark = null;
        Word.Range? formulaRange = null;
        Word.Range? tokenRange = null;
        Microsoft.Office.Interop.Word.Font? tokenFont = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(context + ": VTOMML bookmark is missing.");
            formulaRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            var semantic = WordOmmlConverter.StripVisualTeXNativeEquationNumber(
                formulaRange.WordOpenXML ?? string.Empty);
            var xml = XElement.Parse(semantic, LoadOptions.PreserveWhitespace);
            XNamespace math =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            AssertTrue(xml.Descendants(math + "m").Any(),
                context + ": native matrix structure is missing.");
            // Office normalizes large operators as m:nary structures. Depending on
            // the Word build/operator defaults, the concrete glyph is not guaranteed
            // to remain as a literal m:chr/m:t character in Range.WordOpenXML. The
            // established complex-OMML acceptance separately compares the full
            // stripped semantic structure against an ordinary Display OMath. This
            // combined font regression only needs to prove that all three source
            // large operators survived as native n-ary nodes while font-size edits
            // and #(SEQ) wrapping are applied.
            AssertTrue(xml.Descendants(math + "nary").Count() >= 3,
                context + ": Σ/Π/∫ native n-ary structures are missing.");

            var linearText = formulaRange.Text ?? string.Empty;
            var separatorOffset = linearText.LastIndexOf('#');
            var bodyLength = separatorOffset >= 0 ? separatorOffset : linearText.Length;
            var tokenOffset = -1;
            for (var index = 0; index < bodyLength; index++)
            {
                var character = linearText[index];
                if (character is '\r' or '\n' or '\v' or '\a' or '\t'
                    || char.IsWhiteSpace(character))
                    continue;
                tokenOffset = index;
                break;
            }
            AssertTrue(tokenOffset >= 0,
                context + ": no visible semantic token was found before #().");
            tokenRange = document.Range(
                formulaRange.Start + tokenOffset,
                formulaRange.Start + tokenOffset + 1);
            tokenFont = tokenRange.Font;
            AssertNear(expectedFontSize, tokenFont.Size, 0.2f,
                context + ": first semantic token did not receive the requested font size.");
        }
        finally
        {
            Release(tokenFont);
            Release(tokenRange);
            Release(formulaRange);
            Release(bookmark);
        }
    }

    private static string NativeHashComplexLatex(float fontSize) =>
        @"A=\begin{pmatrix}a&b\\c&d\end{pmatrix}"
        + @"+\sum_{k=1}^{n}k+\prod_{j=1}^{m}j+\int_{0}^{1}f(x)\,dx"
        + $"% {fontSize:0.##}pt";

    private static string NativeHashComplexMathMl() =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
        + "<mrow><mi>A</mi><mo>=</mo><mfenced open=\"(\" close=\")\"><mtable>"
        + "<mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr>"
        + "<mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr>"
        + "</mtable></mfenced><mo>+</mo>"
        + "<munderover><mo>∑</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><mi>k</mi>"
        + "<mo>+</mo><munderover><mo>∏</mo><mrow><mi>j</mi><mo>=</mo><mn>1</mn></mrow><mi>m</mi></munderover><mi>j</mi>"
        + "<mo>+</mo><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo><mi>d</mi><mi>x</mi>"
        + "</mrow></math>";
}

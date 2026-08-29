using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class NumberedOmmlDisplayCase
    {
        internal NumberedOmmlDisplayCase(
            string name,
            string latex,
            string mathMl,
            params (string ElementName, int MinimumCount)[] requiredElements)
        {
            Name = name;
            Latex = latex;
            MathMl = mathMl;
            RequiredElements = requiredElements;
        }

        internal string Name { get; }
        internal string Latex { get; }
        internal string MathMl { get; }
        internal IReadOnlyList<(string ElementName, int MinimumCount)> RequiredElements { get; }
    }

    private sealed class NumberedOmmlDisplayPair
    {
        internal NumberedOmmlDisplayPair(
            NumberedOmmlDisplayCase testCase,
            string numberedFormulaId,
            string plainFormulaId)
        {
            TestCase = testCase;
            NumberedFormulaId = numberedFormulaId;
            PlainFormulaId = plainFormulaId;
        }

        internal NumberedOmmlDisplayCase TestCase { get; }
        internal string NumberedFormulaId { get; }
        internal string PlainFormulaId { get; }
    }

    private static void RunWordNumberedOmmlTrueDisplayComplexAcceptance(
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-numbered-omml-true-display-complex.docx");
        var cases = CreateNumberedOmmlDisplayCases();

        Word.Application? application = null;
        Word.Document? document = null;
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
            var pairs = new List<NumberedOmmlDisplayPair>(cases.Count);

            foreach (var testCase in cases)
            {
                var numberedFormulaId = Guid.NewGuid().ToString("D");
                InsertComplexDisplayFormula(
                    application,
                    document,
                    service,
                    testCase,
                    numberedFormulaId,
                    numbered: true);
                AssertOmmlTabNumberingHost(
                    document,
                    numberedFormulaId,
                    context: $"complex true-display numbered {testCase.Name}",
                    updateReference: true);

                var plainFormulaId = Guid.NewGuid().ToString("D");
                InsertComplexDisplayFormula(
                    application,
                    document,
                    service,
                    testCase,
                    plainFormulaId,
                    numbered: false);
                pairs.Add(new NumberedOmmlDisplayPair(
                    testCase,
                    numberedFormulaId,
                    plainFormulaId));
                document.Save();
                AssertComplexNumberedOmmlDisplayPair(
                    application,
                    document,
                    pairs[pairs.Count - 1],
                    "initial insertion");
            }

            AssertEqual(0, document.Tables.Count,
                "Complex numbered OMML insertion created a Word table.");
            AssertEqual(0, document.Shapes.Count,
                "Complex numbered OMML insertion created a legacy floating number Shape.");
            document.Fields.Update();
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertEqual(0, document.Tables.Count,
                "Complex numbered OMML save/reopen recreated a Word table.");
            foreach (var pair in pairs)
            {
                AssertOmmlTabNumberingHost(
                    document,
                    pair.NumberedFormulaId,
                    context: $"complex true-display save/reopen {pair.TestCase.Name}",
                    updateReference: true);
                AssertComplexNumberedOmmlDisplayPair(
                    application,
                    document,
                    pair,
                    "save/reopen");
            }
            AssertEqual(0, document.Shapes.Count,
                "Complex numbered OMML save/reopen recreated a legacy number Shape.");
            document.Save();

            Console.WriteLine(
                "Word complex numbered-OMML native-#SEQ acceptance passed: fraction/radical, sum/product/integral limits, large operators, nested fractions and matrices retained the same semantic OMath structure and body geometry as ordinary wdOMathDisplay formulas, with one direct mathematical SEQ and zero Shape/Table objects after F9 plus save/reopen.");
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

    private static void InsertComplexDisplayFormula(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        NumberedOmmlDisplayCase testCase,
        string formulaId,
        bool numbered)
    {
        var insertion = document.Content.End - 1;
        application.Selection.SetRange(insertion, insertion);
        var session = new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = document.FullName,
            SourceObjectId = WordRangeReference(insertion, insertion),
            Title = "VisualTeX complex genuine-display OMML acceptance",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = numbered,
            FontSizePt = 14,
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = testCase.Latex,
                },
            },
            ExportResult = new OfficeExportDocument
            {
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
        service.InsertOmml(session, testCase.MathMl);
    }

    private static void AssertComplexNumberedOmmlDisplayPair(
        Word.Application application,
        Word.Document document,
        NumberedOmmlDisplayPair pair,
        string phase)
    {
        Word.Range? numberedRange = null;
        Word.Range? plainRange = null;
        Word.OMaths? numberedMaths = null;
        Word.OMaths? plainMaths = null;
        Word.OMath? numberedMath = null;
        Word.OMath? plainMath = null;
        Word.Paragraphs? numberedParagraphs = null;
        Word.Paragraphs? plainParagraphs = null;
        Word.Paragraph? numberedParagraph = null;
        Word.Paragraph? plainParagraph = null;
        Word.Range? numberedOwner = null;
        Word.Range? plainOwner = null;
        Word.Range? numberedBodyRange = null;
        Word.Window? window = null;
        try
        {
            var numberedMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    pair.NumberedFormulaId)
                ?? throw new InvalidDataException(
                    $"{phase} {pair.TestCase.Name}: numbered metadata is missing.");
            var plainMetadata = WordOmmlFormulaStore.TryRead(
                    document,
                    pair.PlainFormulaId)
                ?? throw new InvalidDataException(
                    $"{phase} {pair.TestCase.Name}: plain-display metadata is missing.");
            numberedRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                pair.NumberedFormulaId,
                numberedMetadata);
            plainRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                document,
                pair.PlainFormulaId,
                plainMetadata);
            numberedMaths = numberedRange.OMaths;
            plainMaths = plainRange.OMaths;
            AssertEqual(1, numberedMaths.Count,
                $"{phase} {pair.TestCase.Name}: numbered host does not contain exactly one OMath.");
            AssertEqual(1, plainMaths.Count,
                $"{phase} {pair.TestCase.Name}: plain host does not contain exactly one OMath.");
            numberedMath = numberedMaths[1];
            plainMath = plainMaths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, numberedMath.Type,
                $"{phase} {pair.TestCase.Name}: numbered formula is not genuine Word Display math.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, plainMath.Type,
                $"{phase} {pair.TestCase.Name}: comparison formula is not genuine Word Display math.");
            Console.WriteLine(
                $"  {phase} {pair.TestCase.Name}: numberedRange={numberedRange.Start}:{numberedRange.End} fields={numberedRange.Fields.Count}; plainRange={plainRange.Start}:{plainRange.End} fields={plainRange.Fields.Count}; documentMaths={document.OMaths.Count}.");
            AssertEqual(1, numberedRange.Fields.Count,
                $"{phase} {pair.TestCase.Name}: numbered formula must contain exactly one mathematical SEQ field.");
            AssertEqual(0, plainRange.Fields.Count,
                $"{phase} {pair.TestCase.Name}: plain formula contains an unexpected field inside OMath.");
            Word.Field? sequenceField = null;
            Word.Range? sequenceCode = null;
            try
            {
                sequenceField = numberedRange.Fields[1];
                sequenceCode = sequenceField.Code;
                AssertTrue(WordEquationNumbering.IsVisualTeXSequenceFieldCode(sequenceCode.Text),
                    $"{phase} {pair.TestCase.Name}: the numbered formula field is not SEQ VisualTeXEquation.");
                AssertTrue((sequenceCode.Text ?? string.Empty).IndexOf(
                        "REF VTEqNum_",
                        StringComparison.OrdinalIgnoreCase) < 0,
                    $"{phase} {pair.TestCase.Name}: a REF field was embedded in the mathematical #() label.");
            }
            finally
            {
                Release(sequenceCode);
                Release(sequenceField);
            }

            numberedParagraphs = numberedRange.Paragraphs;
            plainParagraphs = plainRange.Paragraphs;
            AssertEqual(1, numberedParagraphs.Count,
                $"{phase} {pair.TestCase.Name}: numbered formula spans multiple paragraphs.");
            AssertEqual(1, plainParagraphs.Count,
                $"{phase} {pair.TestCase.Name}: plain formula spans multiple paragraphs.");
            numberedParagraph = numberedParagraphs[1];
            plainParagraph = plainParagraphs[1];
            numberedOwner = numberedParagraph.Range;
            plainOwner = plainParagraph.Range;
            AssertTrue(!(bool)numberedOwner.get_Information(
                    Word.WdInformation.wdWithInTable),
                $"{phase} {pair.TestCase.Name}: numbered formula is inside a table.");
            AssertTrue(!(bool)plainOwner.get_Information(
                    Word.WdInformation.wdWithInTable),
                $"{phase} {pair.TestCase.Name}: comparison formula is inside a table.");

            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var math = (XNamespace)MathNamespace;
            var word = (XNamespace)WordNamespace;
            var numberedOwnerXml = XDocument.Parse(
                numberedOwner.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            var plainOwnerXml = XDocument.Parse(
                plainOwner.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            AssertEqual(1, numberedOwnerXml.Descendants(math + "oMathPara").Count(),
                $"{phase} {pair.TestCase.Name}: numbered formula lost m:oMathPara.");
            AssertEqual(1, plainOwnerXml.Descendants(math + "oMathPara").Count(),
                $"{phase} {pair.TestCase.Name}: plain formula lost m:oMathPara.");
            AssertEqual(1, numberedOwnerXml.Descendants(math + "eqArr").Count(),
                $"{phase} {pair.TestCase.Name}: native #() did not retain exactly one legal m:eqArr.");
            AssertTrue(numberedOwnerXml.Descendants(word + "fldChar").Any(),
                $"{phase} {pair.TestCase.Name}: mathematical SEQ controls are missing.");
            AssertTrue(!numberedOwnerXml.Descendants(word + "instrText").Any(node =>
                    node.Value.IndexOf("REF VTEqNum_", StringComparison.OrdinalIgnoreCase) >= 0),
                $"{phase} {pair.TestCase.Name}: a REF field leaked inside the mathematical label.");

            var numberedEquation = XElement.Parse(
                WordOmmlConverter.StripVisualTeXNativeEquationNumber(
                    numberedRange.WordOpenXML),
                LoadOptions.PreserveWhitespace);
            var plainEquation = XElement.Parse(
                WordOmmlConverter.ExtractSingleOMath(plainRange.WordOpenXML),
                LoadOptions.PreserveWhitespace);
            var numberedSignature = BuildOmmlSemanticStructureSignature(numberedEquation);
            var plainSignature = BuildOmmlSemanticStructureSignature(plainEquation);
            AssertEqual(plainSignature, numberedSignature,
                $"{phase} {pair.TestCase.Name}: numbered and ordinary Display OMML structures diverged.");
            foreach (var requirement in pair.TestCase.RequiredElements)
            {
                var count = numberedEquation
                    .DescendantsAndSelf(math + requirement.ElementName)
                    .Count();
                AssertTrue(count >= requirement.MinimumCount,
                    $"{phase} {pair.TestCase.Name}: expected at least {requirement.MinimumCount} m:{requirement.ElementName} nodes, found {count}.");
            }

            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(numberedRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(80);
            numberedBodyRange = ResolveNativeHashSequenceFormulaBodyRange(
                document,
                numberedRange,
                $"{phase} {pair.TestCase.Name}");
            var numberedBox = ReadVisibleMathInkBox(
                document,
                window,
                numberedBodyRange,
                $"{phase} {pair.TestCase.Name} numbered body ink");
            var numberedWholeBox = ReadVisibleMathInkBox(
                document,
                window,
                numberedRange,
                $"{phase} {pair.TestCase.Name} numbered whole ink");
            window.ScrollIntoView(plainRange, ref scrollStart);
            Thread.Sleep(80);
            var plainBox = ReadVisibleMathInkBox(
                document,
                window,
                plainRange,
                $"{phase} {pair.TestCase.Name} ordinary Display ink");
            // Word's native #() host is an m:eqArr, so the formula-body Range can
            // include a few pixels of equation-array alignment padding even after
            // the number slot is excluded. Reject real scaling/deformation, but do
            // not require the impossible byte-identical box of a plain oMath body.
            AssertMetricRatio(plainBox.Width, numberedBox.Width, 0.90, 1.10,
                $"{phase} {pair.TestCase.Name}: numbered/plain Display body widths diverged.");
            // A partial OMath Range ending immediately before # can make Word's
            // pixel API omit n-ary upper/lower limits. The complete OMath height is
            // authoritative; the trailing number is shorter and cannot inflate it.
            AssertMetricRatio(plainBox.Height, numberedWholeBox.Height, 0.90, 1.10,
                $"{phase} {pair.TestCase.Name}: numbered/plain Display heights diverged.");
            Console.WriteLine(
                $"  {phase} {pair.TestCase.Name}: structureLength={numberedSignature.Length}, numberedBody={numberedBox.Width}x{numberedBox.Height}px, numberedWhole={numberedWholeBox.Width}x{numberedWholeBox.Height}px, plainBox={plainBox.Width}x{plainBox.Height}px, type={numberedMath.Type}, tables={document.Tables.Count}.");
        }
        finally
        {
            Release(window);
            Release(numberedBodyRange);
            Release(plainOwner);
            Release(numberedOwner);
            Release(plainParagraph);
            Release(numberedParagraph);
            Release(plainParagraphs);
            Release(numberedParagraphs);
            Release(plainMath);
            Release(numberedMath);
            Release(plainMaths);
            Release(numberedMaths);
            Release(plainRange);
            Release(numberedRange);
        }
    }

    private static Word.Range ResolveNativeHashSequenceFormulaBodyRange(
        Word.Document document,
        Word.Range equationRange,
        string context)
    {
        var text = equationRange.Text ?? string.Empty;
        var separatorOffset = text.LastIndexOf('#');
        if (separatorOffset <= 0)
            throw new InvalidDataException(
                context + ": the native #(SEQ) separator is missing from OMath.Range.Text.");
        var body = document.Range(
            equationRange.Start,
            equationRange.Start + separatorOffset);
        var bodyText = body.Text ?? string.Empty;
        while (body.End > body.Start
               && bodyText.Length > 0
               && (char.IsWhiteSpace(bodyText[bodyText.Length - 1])
                   || bodyText[bodyText.Length - 1] is '\r' or '\n' or '\v' or '\a'))
        {
            body.MoveEnd(Word.WdUnits.wdCharacter, -1);
            bodyText = body.Text ?? string.Empty;
        }
        if (body.End <= body.Start)
        {
            Release(body);
            throw new InvalidDataException(
                context + ": the native #(SEQ) formula body is empty.");
        }
        return body;
    }

    private static string BuildOmmlSemanticStructureSignature(XElement equation)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var copy = new XElement(equation);
        foreach (var runProperties in copy.Descendants(word + "rPr").ToArray())
            runProperties.Remove();
        foreach (var controlProperties in copy.Descendants(math + "ctrlPr").ToArray())
            controlProperties.Remove();
        foreach (var attribute in copy
                     .DescendantsAndSelf()
                     .Attributes()
                     .Where(attribute => attribute.IsNamespaceDeclaration)
                     .ToArray())
            attribute.Remove();
        return copy.ToString(SaveOptions.DisableFormatting);
    }

    private static IReadOnlyList<NumberedOmmlDisplayCase>
        CreateNumberedOmmlDisplayCases() =>
        new NumberedOmmlDisplayCase[]
        {
            new(
                "fraction-radical",
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>x</mi><mo>=</mo><mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo><msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup><mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow><mrow><mn>2</mn><mi>a</mi></mrow></mfrac></mrow></math>",
                ("f", 1),
                ("rad", 1)),
            new(
                "sum-product-integral-limits",
                @"\sum_{i=1}^{n}a_i+\prod_{j=1}^{m}b_j+\int_{0}^{1}f(x)\,dx",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><msub><mi>a</mi><mi>i</mi></msub><mo>+</mo><munderover><mo>∏</mo><mrow><mi>j</mi><mo>=</mo><mn>1</mn></mrow><mi>m</mi></munderover><msub><mi>b</mi><mi>j</mi></msub><mo>+</mo><msubsup><mo>∫</mo><mn>0</mn><mn>1</mn></msubsup><mi>f</mi><mo>(</mo><mi>x</mi><mo>)</mo><mi>d</mi><mi>x</mi></mrow></math>",
                ("nary", 3)),
            new(
                "large-operator",
                @"\bigcup_{k=1}^{n}A_k\subseteq\bigcap_{k=1}^{n}B_k",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><munderover><mo>⋃</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><msub><mi>A</mi><mi>k</mi></msub><mo>⊆</mo><munderover><mo>⋂</mo><mrow><mi>k</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><msub><mi>B</mi><mi>k</mi></msub></mrow></math>",
                ("nary", 2)),
            new(
                "nested-fraction",
                @"z=\frac{1+\frac{a}{b}}{c+\frac{d}{\frac{e}{f}}}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>z</mi><mo>=</mo><mfrac><mrow><mn>1</mn><mo>+</mo><mfrac><mi>a</mi><mi>b</mi></mfrac></mrow><mrow><mi>c</mi><mo>+</mo><mfrac><mi>d</mi><mfrac><mi>e</mi><mi>f</mi></mfrac></mfrac></mrow></mfrac></mrow></math>",
                ("f", 4)),
            new(
                "matrix",
                @"M=\begin{pmatrix}a&b&c\\d&e&f\\g&h&i\end{pmatrix}",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>M</mi><mo>=</mo><mfenced open=\"(\" close=\")\"><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd><mtd><mi>c</mi></mtd></mtr><mtr><mtd><mi>d</mi></mtd><mtd><mi>e</mi></mtd><mtd><mi>f</mi></mtd></mtr><mtr><mtd><mi>g</mi></mtd><mtd><mi>h</mi></mtd><mtd><mi>i</mi></mtd></mtr></mtable></mfenced></mrow></math>",
                ("m", 1),
                ("mr", 3)),
        };
}

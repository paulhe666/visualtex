using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class ProductionHashFormula
    {
        internal ProductionHashFormula(string formulaId, string latex, string mathMl)
        {
            FormulaId = formulaId;
            Latex = latex;
            MathMl = mathMl;
        }

        internal string FormulaId { get; }
        internal string Latex { get; }
        internal string MathMl { get; }
    }

    private static void RunWordOmmlHashNumberProductionAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-hash-number-production.docx");
        Word.Application? application = null;
        Word.Document? document = null;
        Word.Field? bodyReference = null;
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
            var formulas = new[]
            {
                new ProductionHashFormula(
                    Guid.NewGuid().ToString("D"),
                    @"x=\frac{-b+\sqrt{b^2-4ac}}{2a}",
                    QuadraticFormulaMathMl()),
                new ProductionHashFormula(
                    Guid.NewGuid().ToString("D"),
                    @"S=\sum_{i=1}^{n}i^2",
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>S</mi><mo>=</mo><munderover><mo>∑</mo><mrow><mi>i</mi><mo>=</mo><mn>1</mn></mrow><mi>n</mi></munderover><msup><mi>i</mi><mn>2</mn></msup></mrow></math>"),
                new ProductionHashFormula(
                    Guid.NewGuid().ToString("D"),
                    @"A=\begin{pmatrix}a&b\\c&d\end{pmatrix}",
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\"><mrow><mi>A</mi><mo>=</mo><mfenced open=\"(\" close=\")\"><mtable><mtr><mtd><mi>a</mi></mtd><mtd><mi>b</mi></mtd></mtr><mtr><mtd><mi>c</mi></mtd><mtd><mi>d</mi></mtd></mtr></mtable></mfenced></mrow></math>"),
            };

            for (var index = 0; index < formulas.Length; index++)
            {
                Console.WriteLine(
                    $"  production #(SEQ) insert stage={index + 1}/{formulas.Length}");
                InsertProductionHashFormula(
                    application,
                    document,
                    service,
                    formulas[index],
                    fontSizePt: 14f);
            }

            UpdateProductionHashFields(document, formulas);
            for (var index = 0; index < formulas.Length; index++)
            {
                AssertProductionHashFormula(
                    document,
                    formulas[index].FormulaId,
                    expectedNumber: (index + 1).ToString(
                        System.Globalization.CultureInfo.InvariantCulture),
                    context: $"production #(SEQ) formula {index + 1} after F9");
            }

            bodyReference = AppendProductionHashBodyReference(
                document,
                formulas[1].FormulaId);
            bodyReference.Update();
            AssertEqual(
                "2",
                NormalizeEquationNumberText(bodyReference.Result.Text),
                "The production body REF did not read the mathematical VTEqNum bookmark.");
            AssertEqual(Word.WdStoryType.wdMainTextStory, bodyReference.Result.StoryType,
                "The production body REF is not ordinary main-story text.");
            Word.Range? bodyReferenceCode = null;
            Word.OMaths? bodyReferenceCodeMaths = null;
            try
            {
                // A REF targeting VTEqNum inside OMath can make Word expose the
                // target association through Result.OMaths even though the body
                // field itself is ordinary text. Field.Code is the authoritative
                // host range for deciding whether the REF was inserted in math.
                bodyReferenceCode = bodyReference.Code;
                bodyReferenceCodeMaths = bodyReferenceCode.OMaths;
                AssertEqual(0, bodyReferenceCodeMaths.Count,
                    "The production body REF field code was inserted inside an OMath zone.");
            }
            finally
            {
                Release(bodyReferenceCodeMaths);
                Release(bodyReferenceCode);
            }

            document.Save();
            Release(bodyReference);
            bodyReference = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document);
            document = null;

            document = application.Documents.Open(
                documentPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            UpdateProductionHashFields(document, formulas);
            for (var index = 0; index < formulas.Length; index++)
            {
                AssertProductionHashFormula(
                    document,
                    formulas[index].FormulaId,
                    expectedNumber: (index + 1).ToString(
                        System.Globalization.CultureInfo.InvariantCulture),
                    context: $"production #(SEQ) formula {index + 1} save/reopen");
            }
            bodyReference = FindExternalEquationReference(
                    document,
                    formulas[1].FormulaId)
                ?? throw new InvalidDataException(
                    "Save/reopen lost the production body REF field.");
            bodyReference.Update();
            AssertEqual(
                "2",
                NormalizeEquationNumberText(bodyReference.Result.Text),
                "Save/reopen disconnected the production body REF from VTEqNum.");

            Console.WriteLine(
                "Production OMML #(SEQ) acceptance passed: three WordFormulaService insertions remained wdOMathDisplay, numbered 1/2/3 under F9, exposed mathematical VTEqNum bookmarks to an ordinary body REF, contained no table or Shape, and survived save/reopen.");
        }
        finally
        {
            Release(bodyReference);
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

    private static void InsertProductionHashFormula(
        Word.Application application,
        Word.Document document,
        WordFormulaService service,
        ProductionHashFormula formula,
        float fontSizePt)
    {
        Word.Paragraph? paragraph = null;
        Word.Range? insertion = null;
        try
        {
            paragraph = document.Paragraphs.Add();
            insertion = paragraph.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            application.Selection.SetRange(insertion.Start, insertion.Start);
            var session = CreateNumberedOmmlTabSession(
                formula.FormulaId,
                document.FullName,
                insertion.Start,
                insertion.Start,
                formula.Latex,
                originalMetadata: null);
            session.FontSizePt = fontSizePt;
            service.InsertOmml(session, formula.MathMl);
        }
        finally
        {
            Release(insertion);
            Release(paragraph);
        }
    }

    private static void UpdateProductionHashFields(
        Word.Document document,
        IReadOnlyList<ProductionHashFormula> formulas)
    {
        foreach (var formula in formulas)
        {
            Word.Bookmark? bookmark = null;
            Word.Range? equationRange = null;
            Word.Fields? fields = null;
            try
            {
                bookmark = WordOmmlFormulaStore.FindByFormulaId(
                        document,
                        formula.FormulaId)
                    ?? throw new InvalidDataException(
                        $"Production formula {formula.FormulaId} lost VTOMML identity before F9.");
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                fields = equationRange.Fields;
                fields.Update();
            }
            finally
            {
                Release(fields);
                Release(equationRange);
                Release(bookmark);
            }
        }
        Word.Fields? documentFields = null;
        try
        {
            documentFields = document.Fields;
            if (documentFields.Count > 0) documentFields.Update();
        }
        finally { Release(documentFields); }
    }

    private static void AssertProductionHashFormula(
        Word.Document document,
        string formulaId,
        string expectedNumber,
        string context)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Range? equationRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? fieldCode = null;
        Word.Range? fieldResult = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        Word.Sections? numberSections = null;
        Word.Section? numberSection = null;
        Word.PageSetup? numberPageSetup = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        try
        {
            AssertEqual(0, document.Tables.Count,
                context + ": a Word table exists in the production document.");
            AssertEqual(0, document.Shapes.Count,
                context + ": a floating Shape exists in the production document.");
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(context + ": VTOMML identity is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": VTOMML does not resolve exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": OMath degraded from wdOMathDisplay.");
            mathRange = math.Range;
            fields = mathRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": the native #(SEQ) host does not contain exactly one live field.");
            field = fields[1];
            fieldCode = field.Code;
            fieldResult = field.Result;
            var normalizedCode = (fieldCode.Text ?? string.Empty)
                .Replace("\r", " ")
                .Replace("\n", " ");
            AssertTrue(
                normalizedCode.IndexOf(
                    "SEQ VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + $": mathematical field is not SEQ VisualTeXEquation: '{normalizedCode}'.");
            AssertTrue(
                normalizedCode.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) < 0,
                context + ": a forbidden REF field exists inside the OMath.");
            AssertEqual(expectedNumber, NormalizeEquationNumberText(fieldResult.Text),
                context + ": mathematical SEQ result is stale.");

            bookmarks = document.Bookmarks;
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertTrue(bookmarks.Exists(numberName),
                context + ": VTEqNum bookmark is missing.");
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            AssertEqual(expectedNumber, NormalizeEquationNumberText(numberRange.Text),
                context + ": VTEqNum does not expose the rendered number.");
            AssertTrue(
                numberRange.Start >= mathRange.Start
                && numberRange.End <= mathRange.End,
                context + ": VTEqNum is not hosted inside the numbered OMath.");

            // Word's professional #(...) layout owns the equation centering and
            // right-margin label geometry. A structurally valid m:eqArr can still
            // look wrong if its field-control runs are emitted as ordinary math:
            // the hidden SEQ instruction then consumes layout width and drags the
            // visible number back toward the formula. Guard the actual on-page
            // geometry, not only the XML shape.
            numberSections = numberRange.Sections;
            AssertTrue(numberSections.Count >= 1,
                context + ": VTEqNum has no owning Word section.");
            numberSection = numberSections[1];
            numberPageSetup = numberSection.PageSetup;
            var rightTextEdge = numberPageSetup.PageWidth - numberPageSetup.RightMargin;
            // VTEqNum encloses the complete prefix + SEQ field, including hidden
            // field-control/code runs. For a continuous number with no visible
            // prefix, Range.Information therefore reports the hidden field-code
            // start rather than the on-page digit. Field.Result is the actual
            // rendered mathematical label and is the correct geometry probe.
            var numberX = Convert.ToSingle(fieldResult.get_Information(
                Word.WdInformation.wdHorizontalPositionRelativeToPage));
            AssertTrue(numberX >= rightTextEdge - 40f,
                context + $": native #() visible SEQ result is not right aligned; x={numberX:0.##}, rightEdge={rightTextEdge:0.##}.");

            paragraphs = mathRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                context + ": numbered OMath spans more than one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            AssertEqual(mathRange.End + 1, paragraphRange.End,
                context + ": content exists between OMath end and the paragraph mark.");
            AssertTrue((paragraphRange.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal),
                context + ": numbered OMath paragraph lost its normal paragraph mark.");

            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var mathNs = (XNamespace)MathNamespace;
            var wordNs = (XNamespace)WordNamespace;
            var openXml = XDocument.Parse(
                paragraphRange.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            AssertEqual(1, openXml.Descendants(mathNs + "oMathPara").Count(),
                context + ": production host does not retain one m:oMathPara.");
            AssertEqual(1, openXml.Descendants(mathNs + "eqArr").Count(),
                context + ": production host does not retain Word's legal m:eqArr.");
            AssertTrue(openXml.Descendants(mathNs + "t").Any(node => node.Value == "#"),
                context + ": native # separator is missing.");
            AssertTrue(!openXml.Descendants(wordNs + "tbl").Any(),
                context + ": table XML leaked into the production host.");
            AssertTrue(!openXml.Descendants(wordNs + "txbxContent").Any(),
                context + ": TextBox XML leaked into the production host.");
            AssertTrue(!openXml.Descendants(wordNs + "instrText").Any(node =>
                    node.Value.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) >= 0),
                context + ": mathematical #(REF) leaked into the production host.");

            Console.WriteLine(
                $"  {context}: range={mathRange.Start}:{mathRange.End}, number={expectedNumber}, field='{normalizedCode.Trim()}'.");
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(numberPageSetup);
            Release(numberSection);
            Release(numberSections);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(fieldResult);
            Release(fieldCode);
            Release(field);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(equationRange);
            Release(formulaBookmark);
        }
    }

    private static Word.Field AppendProductionHashBodyReference(
        Word.Document document,
        string formulaId)
    {
        Word.Paragraph? paragraph = null;
        Word.Range? insertion = null;
        Word.OMaths? maths = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        try
        {
            paragraph = document.Paragraphs.Add();
            insertion = paragraph.Range.Duplicate;
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            maths = insertion.OMaths;
            if (maths.Count != 0)
                throw new InvalidOperationException(
                    "The production body REF insertion point is still inside OMath.");
            fields = insertion.Fields;
            object fieldType = Word.WdFieldType.wdFieldRef;
            object fieldCode =
                WordEquationNumbering.NativeNumberBookmarkName(formulaId) + " \\h";
            object preserveFormatting = true;
            field = fields.Add(
                insertion,
                ref fieldType,
                ref fieldCode,
                ref preserveFormatting);
            field.Update();
            var result = field;
            field = null;
            return result;
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(maths);
            Release(insertion);
            Release(paragraph);
        }
    }
}

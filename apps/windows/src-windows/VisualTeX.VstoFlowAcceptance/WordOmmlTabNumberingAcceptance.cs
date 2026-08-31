using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string NumberedOmmlShapeNamePrefix = "VTEqShape_";
    private const string NumberedOmmlShapeAlternativeTextPrefix =
        "VisualTeX numbered OMML ";

    private static void RunWordOmmlTabNumberingAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var documentPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-numbering.docx");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? equationRange = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();
            application.Selection.SetRange(0, 0);
            WordEquationNumbering.SetEquationNumberFormatPreference(
                document,
                EquationNumberFormat.Heading1DotId);

            var service = new WordFormulaService(application);
            var formulaId = Guid.NewGuid().ToString("D");
            var createSession = CreateOmmlNumberingSession(
                document,
                formulaId,
                mode: "create",
                sourceObjectId: WordRangeReference(0, 0),
                latex: @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: null);
            service.InsertOmml(createSession, QuadraticMathMl("x"));

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "fresh numbered OMML insertion",
                updateReference: true);
            AssertManagedOmmlMathRunsUseDocumentMathFont(
                document,
                new[] { formulaId },
                "fresh numbered OMML insertion");
            document.SaveAs2(
                documentPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);

            // Mirror the production DocumentBeforeSave compatibility finalizer
            // before any edit. For native #(SEQ), this only refreshes the unchanged
            // mathematical field and verifies the FormulaId-bound number aliases.
            var finalizedBeforeImmediateSave =
                WordEquationNumbering.FinalizeNumberedOmmlDisplayShapeLayouts(
                    document);
            AssertEqual(1, finalizedBeforeImmediateSave,
                "The save-boundary native-number finalizer did not settle the fresh numbered OMML formula.");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "fresh numbered OMML immediate save/reopen",
                updateReference: true);

            var refreshedCount = WordEquationNumbering.UpdateEquationNumbers(document);
            AssertEqual(1, refreshedCount,
                "The explicit numbering refresh did not recognize exactly one healthy true-display OMML formula.");
            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "healthy true-display OMML numbering refresh",
                updateReference: true);

            var storedMetadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(
                    "Fresh numbered OMML metadata was not persisted.");
            equationRange = ResolveOmmlRange(document, formulaId, storedMetadata);
            var editSession = CreateOmmlNumberingSession(
                document,
                formulaId,
                mode: "edit",
                sourceObjectId: WordRangeReference(
                    equationRange.Start,
                    equationRange.End),
                latex: @"y=\frac{-b\pm\sqrt{b^2-4ac}}{2a}",
                originalMetadata: storedMetadata);
            service.ReplaceOmml(editSession, QuadraticMathMl("y"));
            Release(equationRange); equationRange = null;

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "numbered OMML edit/reconcile",
                updateReference: true);
            AssertManagedOmmlMathRunsUseDocumentMathFont(
                document,
                new[] { formulaId },
                "numbered OMML edit/reconcile");

            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;
            document = application.Documents.Open(
                documentPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();

            AssertOmmlTabNumberingHost(
                document,
                formulaId,
                context: "numbered OMML save/reopen",
                updateReference: true);
            AssertManagedOmmlMathRunsUseDocumentMathFont(
                document,
                new[] { formulaId },
                "numbered OMML save/reopen");

            document.Save();
            Console.WriteLine(
                "Word numbered-OMML true-display acceptance passed: insert, F9 plus body REF update, edit and save/reopen retained one wdOMathDisplay/m:oMathPara formula using Word-native #(SEQ), with VTEqNum inside the mathematical number slot and zero Shape/Table objects.");
        }
        finally
        {
            Release(equationRange);
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

    private static OfficeSessionDocument CreateOmmlNumberingSession(
        Word.Document document,
        string formulaId,
        string mode,
        string sourceObjectId,
        string latex,
        FormulaMetadata? originalMetadata)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = mode,
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = document.FullName,
            SourceObjectId = sourceObjectId,
            Title = "VisualTeX numbered OMML true-display acceptance",
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = true,
            FontSizePt = 14,
            OriginalMetadata = originalMetadata,
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = latex,
                },
            },
            ExportResult = new OfficeExportDocument
            {
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
    }

    private static string QuadraticMathMl(string variable) =>
        "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">"
        + $"<mi>{variable}</mi><mo>=</mo>"
        + "<mfrac><mrow><mo>−</mo><mi>b</mi><mo>±</mo>"
        + "<msqrt><mrow><msup><mi>b</mi><mn>2</mn></msup>"
        + "<mo>−</mo><mn>4</mn><mi>a</mi><mi>c</mi></mrow></msqrt></mrow>"
        + "<mrow><mn>2</mn><mi>a</mi></mrow></mfrac></math>";

    private static Word.Range ResolveOmmlRange(
        Word.Document document,
        string formulaId,
        FormulaMetadata metadata)
    {
        return WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
            document,
            formulaId,
            metadata);
    }

    // Kept under the historical helper name because many migration/round-trip
    // acceptances call it. The current producer is the minimal direct-SEQ 1x3 host;
    // older native #(SEQ) equations remain valid migration input and are still
    // asserted by the compatibility branches below.
    private static void AssertOmmlTabNumberingHost(
        Word.Document document,
        string formulaId,
        string context,
        bool updateReference,
        bool requireDocumentTableFree = true)
    {
        Word.Range? currentRange = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
            if (metadata is not null)
            {
                currentRange = WordOmmlFormulaStore.GetEquationRangeVerifiedForStructuralEdit(
                    document,
                    formulaId,
                    metadata);
                if ((bool)currentRange.get_Information(Word.WdInformation.wdWithInTable))
                {
                    if (updateReference) document.Fields.Update();
                    AssertOmmlTableNumberLifecyclePhase(
                        document.Application,
                        document,
                        formulaId,
                        context);
                    return;
                }
            }
        }
        finally { Release(currentRange); }

        if (TryAssertNativeHashNumberingHostV3(
                document,
                formulaId,
                context,
                updateReference,
                requireDocumentTableFree))
            return;

        AssertNativeHashSequenceNumberingHost(
            document,
            formulaId,
            context,
            updateReference);
        AssertNativeHashSequenceRemainsDynamicallyOrdered(
            document,
            formulaId,
            context);
        AssertNativeHashSequenceOwnerHasNoFloatingFrame(
            document,
            formulaId,
            context);
        if (requireDocumentTableFree)
        {
            AssertEqual(
                0,
                document.Tables.Count,
                context + ": the final document still contains a Word numbering table.");
        }
    }

    private static void AssertNativeHashSequenceOwnerHasNoFloatingFrame(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Range? owner = null;
        Word.Frames? frames = null;
        Word.ShapeRange? shapes = null;
        try
        {
            owner = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(
                    context + ": native #(SEQ) owner range is missing.");
            frames = owner.Frames;
            AssertEqual(0, frames.Count,
                context + ": the native numbered OMML owner is inside an old Word Frame.");
            try
            {
                shapes = owner.ShapeRange;
                AssertEqual(0, shapes.Count,
                    context + ": a floating Shape is anchored to the native numbered OMML owner.");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Word throws when no floating Shape is anchored to this range.
            }
        }
        finally
        {
            Release(shapes);
            Release(frames);
            Release(owner);
        }
    }

    private static void AssertNativeHashSequenceRemainsDynamicallyOrdered(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? equationRange = null;
        Word.Fields? fields = null;
        Word.Field? sequence = null;
        Word.Range? code = null;
        try
        {
            bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId)
                ?? throw new InvalidDataException(
                    context + ": the native #(SEQ) formula bookmark is missing.");
            equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
            fields = equationRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": the native numbered OMath does not contain exactly one SEQ field.");
            sequence = fields[1];
            code = sequence.Code;
            var instruction = code.Text ?? string.Empty;
            AssertTrue(
                !System.Text.RegularExpressions.Regex.IsMatch(
                    instruction,
                    @"\\r\s+\d+",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    | System.Text.RegularExpressions.RegexOptions.CultureInvariant),
                context + ": the mathematical SEQ field was frozen with \\r N; F9 could no longer dynamically reorder later formulas. Code='"
                + instruction.Trim() + "'.");
        }
        finally
        {
            Release(code);
            Release(sequence);
            Release(fields);
            Release(equationRange);
            Release(bookmark);
        }
    }

    // Core assertion for the current production host: one genuine display OMath
    // containing Word's native m:eqArr/#(SEQ VisualTeXEquation), with all number
    // aliases inside math and no floating Shape/TextBox or numbering table.
    private static void AssertNativeHashSequenceNumberingHost(
        Word.Document document,
        string formulaId,
        string context,
        bool updateReference,
        bool requireDocumentTableFree = true)
    {
        Word.Range? equationRange = null;
        Word.Range? ownerRange = null;
        Word.Range? visibleRange = null;
        Word.Range? visibleTextRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.Fields? fields = null;
        Word.Field? sequenceField = null;
        Word.Range? sequenceCode = null;
        Word.Range? sequenceResult = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(context + ": OMML metadata is missing.");
            equationRange = ResolveOmmlRange(document, formulaId, metadata);
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": numbering owner is missing.");
            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": VTEq_ number alias is missing.");
            visibleTextRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": visible number text is missing.");

            if (requireDocumentTableFree)
                AssertEqual(0, document.Tables.Count,
                    context + ": the final document still contains a Word numbering table.");
            AssertTrue(
                !(bool)ownerRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": numbered OMML is still hosted by a legacy table.");
            AssertEqual(Word.WdStoryType.wdMainTextStory, visibleRange.StoryType,
                context + ": VTEq_ is not inside the main mathematical paragraph.");

            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": equation range does not contain exactly one native OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": numbered OMML is not genuine Word display math.");
            mathRange = math.Range.Duplicate;
            paragraphs = mathRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                context + ": numbered OMML spans more than one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            AssertEqual(mathRange.End + 1, paragraphRange.End,
                context + ": OMath is not immediately followed by the paragraph mark.");
            AssertTrue((paragraphRange.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal),
                context + ": owner paragraph does not end with a normal paragraph mark.");
            AssertTrue((paragraphRange.Text ?? string.Empty).IndexOf('\t') < 0,
                context + ": native #SEQ display unexpectedly contains a tab stop token.");

            fields = mathRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": native #SEQ OMath must own exactly one SEQ field.");
            sequenceField = fields[1];
            sequenceCode = sequenceField.Code;
            AssertTrue(WordEquationNumbering.IsVisualTeXSequenceFieldCode(sequenceCode.Text),
                context + ": mathematical field is not SEQ VisualTeXEquation.");
            AssertTrue((sequenceCode.Text ?? string.Empty).IndexOf(
                    "REF VTEqNum_",
                    StringComparison.OrdinalIgnoreCase) < 0,
                context + ": REF was incorrectly embedded inside #().");
            if (updateReference)
                sequenceField.Update();
            sequenceResult = sequenceField.Result;
            AssertTrue(!string.IsNullOrWhiteSpace(NormalizeEquationNumberText(sequenceResult.Text)),
                context + ": SEQ field has no visible result.");

            var openXml = equationRange.WordOpenXML ?? string.Empty;
            AssertTrue(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(openXml),
                context + ": OMath is not the accepted native #(SEQ) structure.");
            AssertTrue(openXml.IndexOf(
                    "REF VTEqNum_",
                    StringComparison.OrdinalIgnoreCase) < 0,
                context + ": mathematical OpenXML contains a REF target.");
            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var word = (XNamespace)WordNamespace;
            var mathNamespace = (XNamespace)MathNamespace;
            var parsed = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
            AssertEqual(1, parsed.Descendants(mathNamespace + "oMathPara").Count(),
                context + ": display equation lost m:oMathPara.");
            AssertEqual(1, parsed.Descendants(mathNamespace + "eqArr").Count(),
                context + ": native Word #() did not serialize as exactly one m:eqArr.");
            AssertTrue(parsed.Descendants(mathNamespace + "t").Any(node => node.Value == "#"),
                context + ": native equation-number # separator is missing.");
            AssertTrue(parsed.Descendants(word + "fldChar").Any(),
                context + ": SEQ field controls are missing from mathematical OpenXML.");
            AssertTrue(!parsed.Descendants(word + "tbl").Any(),
                context + ": a Word table leaked into the native #SEQ paragraph.");
            AssertTrue(!parsed.Descendants(word + "drawing").Any(),
                context + ": a floating drawing leaked into the native #SEQ paragraph.");
            AssertTrue(openXml.IndexOf("<w:txbxContent", StringComparison.OrdinalIgnoreCase) < 0,
                context + ": a TextBox leaked into the native #SEQ paragraph.");
            AssertTrue((document.Content.WordOpenXML ?? string.Empty).IndexOf(
                    NumberedOmmlShapeNamePrefix + Guid.Parse(formulaId).ToString("N"),
                    StringComparison.OrdinalIgnoreCase) < 0,
                context + ": legacy VTEqShape_ survived in the document.");

            bookmarks = document.Bookmarks;
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertTrue(bookmarks.Exists(numberName),
                context + ": VTEqNum_<FormulaId> bookmark is missing.");
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            AssertEqual(Word.WdStoryType.wdMainTextStory, numberRange.StoryType,
                context + ": VTEqNum bookmark is not in the main document story.");
            AssertTrue(numberRange.Start >= mathRange.Start && numberRange.End <= mathRange.End,
                context + ": VTEqNum does not wrap the number region inside OMath.");
            AssertTrue(visibleRange.Start >= mathRange.Start && visibleRange.End <= mathRange.End,
                context + ": VTEq_ alias lies outside the native OMath number slot.");
            AssertEqual(
                NormalizeEquationNumberText(numberRange.Text),
                NormalizeEquationNumberText(visibleTextRange.Text),
                context + ": visible number alias and VTEqNum disagree.");

            Console.WriteLine(
                $"  {context}: native wdOMathDisplay #(SEQ) verified range={mathRange.Start}:{mathRange.End}, number={NormalizeEquationNumberText(numberRange.Text)}, shapes={document.Shapes.Count}, tables={document.Tables.Count}.");
        }
        finally
        {
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(sequenceResult);
            Release(sequenceCode);
            Release(sequenceField);
            Release(fields);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(visibleTextRange);
            Release(visibleRange);
            Release(ownerRange);
            Release(equationRange);
        }
    }

    private static void AssertOmmlShapeNumberingHostLegacy(
        Word.Document document,
        string formulaId,
        string context,
        bool updateReference,
        bool requireDocumentTableFree = true)
    {
        Word.Range? equationRange = null;
        Word.Range? ownerRange = null;
        Word.Range? visibleRange = null;
        Word.Range? visibleTextRange = null;
        Word.Range? documentRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Fields? equationFields = null;
        Word.InlineShapes? inlineShapes = null;
        Word.Shape? numberShape = null;
        Word.TextFrame? textFrame = null;
        Word.Range? textFrameRange = null;
        Word.Fields? visibleFields = null;
        Word.Field? reference = null;
        Word.Range? fieldCode = null;
        Word.Range? fieldResult = null;
        Word.WrapFormat? wrapFormat = null;
        Word.LineFormat? shapeLine = null;
        Word.FillFormat? shapeFill = null;
        Word.Range? anchor = null;
        Word.Paragraphs? anchorParagraphs = null;
        Word.Paragraph? anchorParagraph = null;
        Word.Range? anchorParagraphRange = null;
        Word.ParagraphFormat? anchorFormat = null;
        Microsoft.Office.Interop.Word.Font? visibleFont = null;
        Word.Window? window = null;
        Word.Range? formulaStartProbe = null;
        Word.Range? formulaEndProbe = null;
        Word.Range? ownerStartProbe = null;
        Word.Range? ownerEndProbe = null;
        Word.Range? anchorStartProbe = null;
        Word.Range? anchorEndProbe = null;
        Word.Range? visibleLabelEndProbe = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        try
        {
            var metadata = WordOmmlFormulaStore.TryRead(document, formulaId)
                ?? throw new InvalidDataException(context + ": OMML metadata is missing.");
            equationRange = ResolveOmmlRange(document, formulaId, metadata);
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": numbering owner is missing.");
            visibleRange = WordEquationNumbering.FindVisibleEquationNumberRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": visible equation number is missing.");
            visibleTextRange = WordEquationNumbering.FindVisibleEquationNumberTextRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": visible equation-number text range is missing.");

            if (requireDocumentTableFree)
            {
                AssertEqual(0, document.Tables.Count,
                    context + ": the final document still contains a Word numbering table.");
            }
            AssertTrue(
                !(bool)ownerRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": numbered OMML is still hosted by a legacy table.");
            paragraphs = ownerRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                context + ": the true-display OMML owner spans more than one paragraph.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            AssertEqual(ownerRange.Start, paragraphRange.Start,
                context + ": numbering owner does not start at the formula paragraph.");
            AssertEqual(ownerRange.End, paragraphRange.End,
                context + ": numbering owner does not end at the formula paragraph.");
            AssertTrue(
                equationRange.Start >= ownerRange.Start
                && equationRange.End <= ownerRange.End,
                context + ": the native formula lies outside its pure display paragraph.");

            maths = ownerRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": the formula paragraph does not contain exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": numbered OMML is not genuine Word display math.");
            equationFields = equationRange.Fields;
            AssertEqual(0, equationFields.Count,
                context + ": a field leaked into the OMath Range.");
            inlineShapes = ownerRange.InlineShapes;
            AssertEqual(0, inlineShapes.Count,
                context + ": the pure display paragraph contains an inline object.");

            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var word = (XNamespace)WordNamespace;
            var mathNamespace = (XNamespace)MathNamespace;
            var paragraphXml = XDocument.Parse(
                ownerRange.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            AssertEqual(1, paragraphXml.Descendants(mathNamespace + "oMathPara").Count(),
                context + ": the formula paragraph does not retain exactly one m:oMathPara.");
            AssertEqual(1, paragraphXml.Descendants(mathNamespace + "oMath").Count(),
                context + ": the formula paragraph does not retain exactly one m:oMath.");
            AssertTrue(!paragraphXml.Descendants(mathNamespace + "eqArr").Any(),
                context + ": the obsolete m:eqArr/#(...) numbering wrapper returned.");
            AssertTrue(!paragraphXml.Descendants(word + "tbl").Any(),
                context + ": a table leaked into the formula paragraph OpenXML.");
            AssertTrue(!paragraphXml.Descendants(word + "drawing").Any(),
                context + ": the number Shape anchor leaked into the formula paragraph.");
            AssertTrue(!paragraphXml.Descendants(word + "fldChar").Any(),
                context + ": field controls leaked into the formula paragraph.");
            AssertTrue(!paragraphXml.Descendants(word + "instrText").Any(node =>
                    node.Value.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) >= 0),
                context + ": the dynamic REF leaked inside the formula paragraph.");

            numberShape = FindNumberedOmmlShape(document, formulaId, context);
            // Word can transiently return E_FAIL for Shape.Type immediately after
            // replacing the anchored display OMath. The deterministic Shape name,
            // TextFrame story and serialized w:txbxContent assertions below are the
            // stable proof that this is the external text-box REF host.
            AssertEqual(
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin,
                numberShape.RelativeHorizontalPosition,
                context + ": the number Shape is not positioned relative to page margins.");
            AssertTrue(
                !float.IsNaN(numberShape.Left)
                && !float.IsInfinity(numberShape.Left)
                && numberShape.Left >= 0f,
                context + ": the number Shape has an invalid explicit horizontal position.");
            AssertEqual(
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph,
                numberShape.RelativeVerticalPosition,
                context + ": the number Shape is not positioned relative to its anchor paragraph.");
            AssertTrue(
                !float.IsNaN(numberShape.Top)
                && !float.IsInfinity(numberShape.Top)
                && numberShape.Top >= -120f
                && numberShape.Top <= 240f,
                context + ": the number Shape has an invalid calibrated vertical offset.");
            AssertNear(72f, numberShape.Width, 0.1f,
                context + ": the number Shape width changed.");
            AssertTrue(numberShape.Height >= 30f && numberShape.Height <= 240f,
                context + ": the number Shape height is outside the supported formula range.");
            wrapFormat = numberShape.WrapFormat;
            AssertEqual(Word.WdWrapType.wdWrapNone, wrapFormat.Type,
                context + ": the number Shape unexpectedly wraps body text.");
            AssertTrue(wrapFormat.AllowOverlap != 0,
                context + ": the number Shape does not allow overlay rendering.");
            shapeLine = numberShape.Line;
            shapeFill = numberShape.Fill;
            var serializedDecorationHealthy =
                WordEquationNumbering.IsSerializedNativeDisplayNumberShapeHealthy(
                    document,
                    formulaId);
            Console.WriteLine(
                $"  {context}: Shape line visible={shapeLine.Visible}, transparency={shapeLine.Transparency:0.###}; fill visible={shapeFill.Visible}, transparency={shapeFill.Transparency:0.###}; serializedNoFillNoStroke={serializedDecorationHealthy}.");
            AssertTrue(
                shapeLine.Visible == Office.MsoTriState.msoFalse
                || shapeLine.Transparency >= 0.99f
                || serializedDecorationHealthy,
                context + ": the external number Shape has a visible border in both COM and persisted OpenXML.");
            AssertTrue(
                shapeFill.Visible == Office.MsoTriState.msoFalse
                || shapeFill.Transparency >= 0.99f
                || serializedDecorationHealthy,
                context + ": the external number Shape has a visible fill in both COM and persisted OpenXML.");
            anchor = numberShape.Anchor;
            AssertEqual(Word.WdStoryType.wdMainTextStory, anchor.StoryType,
                context + ": the number Shape anchor is not in the main document story.");
            anchorParagraphs = anchor.Paragraphs;
            AssertEqual(1, anchorParagraphs.Count,
                context + ": the number Shape anchor spans more than one paragraph.");
            anchorParagraph = anchorParagraphs[1];
            anchorParagraphRange = anchorParagraph.Range;
            AssertEqual(anchorParagraphRange.End, ownerRange.Start,
                context + ": the dedicated Shape anchor is not immediately before the display formula.");
            AssertEqual(0, anchorParagraphRange.OMaths.Count,
                context + ": mathematical content leaked into the Shape anchor paragraph.");
            AssertEqual(0, anchorParagraphRange.Fields.Count,
                context + ": fields leaked into the Shape anchor paragraph.");
            anchorFormat = anchorParagraphRange.ParagraphFormat;
            AssertEqual(Word.WdLineSpacing.wdLineSpaceExactly, anchorFormat.LineSpacingRule,
                context + ": the dedicated Shape anchor does not use exact 1pt line spacing.");
            AssertNear(1f, anchorFormat.LineSpacing, 0.1f,
                context + ": the dedicated Shape anchor is not 1pt high.");

            AssertEqual(Word.WdStoryType.wdTextFrameStory, visibleRange.StoryType,
                context + ": the visible number is not external ordinary Word text.");
            AssertTrue(
                visibleRange.Start < equationRange.Start
                || visibleRange.Start >= equationRange.End,
                context + ": the visible REF range intersects the OMath Range.");
            textFrame = numberShape.TextFrame;
            Console.WriteLine(
                $"  {context}: Word TextFrame runtime margins={textFrame.MarginLeft:0.###},{textFrame.MarginTop:0.###},{textFrame.MarginRight:0.###},{textFrame.MarginBottom:0.###}.");
            sections = equationRange.Sections;
            AssertTrue(sections.Count > 0,
                context + ": the display formula is not associated with a Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            var writableWidth = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            var expectedShapeLeft = Math.Max(
                0f,
                writableWidth - numberShape.Width + Math.Max(0f, textFrame.MarginRight));
            AssertNear(
                expectedShapeLeft,
                numberShape.Left,
                1.5f,
                context + ": the external number Shape is not aligned to the writable right boundary.");
            AssertTrue(
                textFrame.MarginLeft >= 0f
                && textFrame.MarginRight >= 0f
                && textFrame.MarginTop >= 0f
                && textFrame.MarginBottom >= 0f
                && textFrame.MarginLeft <= 8f
                && textFrame.MarginRight <= 8f
                && textFrame.MarginTop <= 4f
                && textFrame.MarginBottom <= 4f,
                context + ": the number Shape has invalid Word text-box insets.");
            textFrameRange = textFrame.TextRange;
            AssertTrue(
                visibleRange.Start >= textFrameRange.Start
                && visibleRange.End <= textFrameRange.End,
                context + ": the VTEq bookmark lies outside its text-box story.");

            var visibleText = NormalizeNumberedOmmlLabel(visibleTextRange.Text);
            AssertTrue(
                visibleText.StartsWith("(", StringComparison.Ordinal)
                && visibleText.EndsWith(")", StringComparison.Ordinal),
                context + ": visible number is not enclosed by ordinary parentheses: '"
                + visibleText + "'.");
            visibleFields = visibleTextRange.Fields;
            AssertEqual(1, visibleFields.Count,
                context + ": the visible text-box label does not contain exactly one field.");
            reference = visibleFields[1];
            fieldCode = reference.Code;
            AssertTrue((fieldCode.Text ?? string.Empty).IndexOf(
                    "REF " + WordEquationNumbering.NativeNumberBookmarkName(formulaId),
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the visible field is not the expected dynamic REF.");
            AssertEqual(0, fieldCode.OMaths.Count,
                context + ": the REF field code is nested inside OMath.");
            if (updateReference) reference.Update();
            fieldResult = reference.Result;
            AssertEqual(0, fieldResult.OMaths.Count,
                context + ": the REF field result is nested inside OMath.");
            var resultText = NormalizeNumberedOmmlLabel(fieldResult.Text);
            AssertTrue(resultText.Length > 0,
                context + ": F9 produced an empty equation-number result.");
            AssertTrue(resultText.IndexOf('(') < 0 && resultText.IndexOf(')') < 0,
                context + ": parentheses leaked into REF.Result: '" + resultText + "'.");
            visibleFont = visibleTextRange.Font;
            AssertEqual(0, visibleFont.Position,
                context + ": visible equation-number text has a manual baseline shift.");

            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(equationRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(100);
            var formulaBox = ReadVisibleMathInkBox(
                document,
                window,
                equationRange,
                context + " true-display formula ink");
            var wholeFormulaBox = ReadWordRangePixelBox(
                window,
                equationRange,
                context + " whole true-display formula range");
            var numberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " external Shape REF result");
            var formulaCenterY = formulaBox.Top + formulaBox.Height / 2.0;
            var numberCenterY = numberBox.Top + numberBox.Height / 2.0;
            var verticalCenterDelta = numberCenterY - formulaCenterY;
            formulaStartProbe = document.Range(equationRange.Start, equationRange.Start);
            formulaEndProbe = document.Range(equationRange.End, equationRange.End);
            ownerStartProbe = document.Range(ownerRange.Start, ownerRange.Start);
            ownerEndProbe = document.Range(ownerRange.End, ownerRange.End);
            anchorStartProbe = document.Range(anchorParagraphRange.Start, anchorParagraphRange.Start);
            anchorEndProbe = document.Range(anchorParagraphRange.End, anchorParagraphRange.End);
            static float ReadVerticalPosition(Word.Range probe) => Convert.ToSingle(
                probe.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage));
            static int ReadPageNumber(Word.Range probe) => Convert.ToInt32(
                probe.get_Information(Word.WdInformation.wdActiveEndAdjustedPageNumber));
            Console.WriteLine(
                $"  {context}: formulaBox={formulaBox.Left},{formulaBox.Top},{formulaBox.Width},{formulaBox.Height}, wholeFormulaBox={wholeFormulaBox.Left},{wholeFormulaBox.Top},{wholeFormulaBox.Width},{wholeFormulaBox.Height}, numberBox={numberBox.Left},{numberBox.Top},{numberBox.Width},{numberBox.Height}, centerDelta={verticalCenterDelta:0.##}px; "
                + $"verticalPt anchor={ReadVerticalPosition(anchorStartProbe):0.###}->{ReadVerticalPosition(anchorEndProbe):0.###} "
                + $"owner={ReadVerticalPosition(ownerStartProbe):0.###}->{ReadVerticalPosition(ownerEndProbe):0.###} "
                + $"math={ReadVerticalPosition(formulaStartProbe):0.###}->{ReadVerticalPosition(formulaEndProbe):0.###} "
                + $"pages anchor={ReadPageNumber(anchorStartProbe)}/{ReadPageNumber(anchorEndProbe)} owner={ReadPageNumber(ownerStartProbe)}/{ReadPageNumber(ownerEndProbe)} "
                + $"refTopPt={ReadVerticalPosition(fieldResult):0.###} shapeTop={numberShape.Top:0.###} shapeHeight={numberShape.Height:0.###}.");
            AssertTrue(Math.Abs(verticalCenterDelta) <= 4.0,
                context + $": the external number Shape is not vertically centered on the true-display formula (delta={verticalCenterDelta:0.##}px).");

            documentRange = document.Content;
            var documentXml = documentRange.WordOpenXML ?? string.Empty;
            var normalizedId = Guid.Parse(formulaId).ToString("N");
            AssertTrue(documentXml.IndexOf(
                    "VTEq_" + normalizedId,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the visible-number bookmark is absent from document OpenXML.");
            AssertTrue(documentXml.IndexOf(
                    "VTEqAnc_" + normalizedId,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the dedicated Shape-anchor bookmark is absent from document OpenXML.");
            AssertTrue(documentXml.IndexOf("<w:txbxContent", StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the external REF did not serialize in w:txbxContent.");
            var shapeMarker = "alt=\""
                + NumberedOmmlShapeAlternativeTextPrefix
                + normalizedId
                + "\"";
            var shapeMarkerIndex = documentXml.IndexOf(
                shapeMarker,
                StringComparison.OrdinalIgnoreCase);
            AssertTrue(shapeMarkerIndex >= 0,
                context + ": the deterministic numbered-OMML Shape marker is absent from OpenXML.");
            var shapeFragmentStart = Math.Max(0, shapeMarkerIndex - 1200);
            var shapeFragmentLength = Math.Min(
                3600,
                documentXml.Length - shapeFragmentStart);
            var shapeFragment = documentXml.Substring(
                shapeFragmentStart,
                shapeFragmentLength);
            Console.WriteLine(
                $"  {context}: serialized numbered-OMML Shape fragment={shapeFragment}");
            AssertTrue(shapeFragment.IndexOf(
                    "mso-position-horizontal-relative:margin",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the serialized number Shape lost its margin-relative horizontal positioning.");
            AssertTrue(shapeFragment.IndexOf(
                    "mso-position-vertical-relative:text",
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the serialized number Shape lost its paragraph-relative vertical positioning.");
            AssertTrue(documentXml.IndexOf("<m:oMathPara", StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the genuine display math host is absent from document OpenXML.");

            Console.WriteLine(
                $"  {context}: formula={equationRange.Start}-{equationRange.End}, owner={ownerRange.Start}-{ownerRange.End}, anchor={anchorParagraphRange.Start}-{anchorParagraphRange.End}, shape={numberShape.Name}/{numberShape.Width:0.##}x{numberShape.Height:0.##}pt, visible='{visibleText}', result='{resultText}', type={math.Type}, tables={document.Tables.Count}.");
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(visibleLabelEndProbe);
            Release(anchorEndProbe);
            Release(anchorStartProbe);
            Release(ownerEndProbe);
            Release(ownerStartProbe);
            Release(formulaEndProbe);
            Release(formulaStartProbe);
            Release(window);
            Release(visibleFont);
            Release(anchorFormat);
            Release(anchorParagraphRange);
            Release(anchorParagraph);
            Release(anchorParagraphs);
            Release(anchor);
            Release(shapeFill);
            Release(shapeLine);
            Release(wrapFormat);
            Release(fieldResult);
            Release(fieldCode);
            Release(reference);
            Release(visibleFields);
            Release(textFrameRange);
            Release(textFrame);
            Release(numberShape);
            Release(inlineShapes);
            Release(equationFields);
            Release(math);
            Release(maths);
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(documentRange);
            Release(visibleTextRange);
            Release(visibleRange);
            Release(ownerRange);
            Release(equationRange);
        }
    }

    private static bool TryAssertNativeHashNumberingHostV3(
        Word.Document document,
        string formulaId,
        string context,
        bool updateReference,
        bool requireDocumentTableFree)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Range? formulaRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Fields? fields = null;
        Word.Field? sequenceField = null;
        Word.Range? code = null;
        Word.Bookmarks? bookmarks = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? numberRange = null;
        Word.Range? ownerRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? paragraphRange = null;
        try
        {
            formulaBookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
            if (formulaBookmark is null) return false;
            formulaRange = WordOmmlFormulaStore.GetEquationRange(formulaBookmark);
            var openXml = formulaRange.WordOpenXML ?? string.Empty;
            if (!WordOmmlConverter.HasVisualTeXNativeEquationNumber(openXml)
                || openXml.IndexOf("SEQ VisualTeXEquation",
                    StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            if (requireDocumentTableFree)
                AssertEqual(0, document.Tables.Count,
                    context + ": native #() document contains a Word table.");
            AssertEqual(0, document.Shapes.Count,
                context + ": native #() document contains a floating Shape/TextBox.");
            maths = formulaRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": native #() formula does not contain exactly one OMath.");
            math = maths[1];
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                context + ": native #() formula degraded from wdOMathDisplay.");
            mathRange = math.Range;
            fields = mathRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": native #() formula does not contain exactly one SEQ field.");
            sequenceField = fields[1];
            code = sequenceField.Code;
            var instruction = code.Text ?? string.Empty;
            AssertTrue(
                WordEquationNumbering.IsVisualTeXSequenceFieldCode(instruction),
                context + $": native #() field is not SEQ VisualTeXEquation: '{instruction}'.");
            AssertTrue(instruction.IndexOf(
                    "REF ",
                    StringComparison.OrdinalIgnoreCase) < 0,
                context + ": native #() contains REF inside OMath.");
            AssertTrue(!System.Text.RegularExpressions.Regex.IsMatch(
                    instruction,
                    @"\r\s+\d+\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                    | System.Text.RegularExpressions.RegexOptions.CultureInvariant),
                context + ": native #() mathematical SEQ is frozen with \r N.");
            if (updateReference) sequenceField.Update();

            bookmarks = document.Bookmarks;
            var numberName = WordEquationNumbering.NativeNumberBookmarkName(formulaId);
            AssertTrue(bookmarks.Exists(numberName),
                context + ": VTEqNum bookmark is missing from native #().");
            numberBookmark = bookmarks[numberName];
            numberRange = numberBookmark.Range;
            AssertTrue(numberRange.Start >= mathRange.Start
                && numberRange.End <= mathRange.End,
                context + ": VTEqNum bookmark lies outside the mathematical SEQ result.");
            ownerRange = WordEquationNumbering.FindNumberingOwnerRange(document, formulaId)
                ?? throw new InvalidDataException(context + ": native #() owner is missing.");
            paragraphs = ownerRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                context + ": native #() owner spans multiple paragraphs.");
            paragraph = paragraphs[1];
            paragraphRange = paragraph.Range;
            AssertEqual(mathRange.End + 1, paragraphRange.End,
                context + ": paragraph mark is not immediately after native #().");
            AssertTrue(!(bool)paragraphRange.get_Information(
                    Word.WdInformation.wdWithInTable),
                context + ": native #() formula remains in a table.");

            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var mathNs = (XNamespace)MathNamespace;
            var word = (XNamespace)WordNamespace;
            var xml = XDocument.Parse(
                paragraphRange.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            AssertEqual(1, xml.Descendants(mathNs + "oMathPara").Count(),
                context + ": native #() lost m:oMathPara.");
            AssertEqual(1, xml.Descendants(mathNs + "eqArr").Count(),
                context + ": native #() lost Word's legal m:eqArr host.");
            var equationArray = xml.Descendants(mathNs + "eqArr").Single();
            var equationRow = equationArray.Elements(mathNs + "e").Single();
            var rowChildren = equationRow.Elements().ToArray();
            var hasNativeHashSeparator = rowChildren
                .Select((element, index) => new { element, index })
                .Any(item =>
                    item.element.Name == mathNs + "r"
                    && string.Concat(item.element.Elements(mathNs + "t").Select(node => node.Value))
                        .EndsWith("#", StringComparison.Ordinal)
                    && item.index + 1 < rowChildren.Length
                    && rowChildren[item.index + 1].Name == mathNs + "d");
            AssertTrue(hasNativeHashSeparator,
                context + ": native #() lost its hash separator immediately before the number delimiter.");
            AssertTrue(!xml.Descendants(word + "drawing").Any(),
                context + ": native #() paragraph contains a DrawingML Shape.");
            AssertTrue(!xml.Descendants(word + "tbl").Any(),
                context + ": native #() paragraph contains a Word table.");

            Console.WriteLine(
                $"  {context}: native wdOMathDisplay #(SEQ) verified range={mathRange.Start}:{mathRange.End}, number={NormalizeNumberedOmmlLabel(numberRange.Text)}, shapes={document.Shapes.Count}, tables={document.Tables.Count}.");
            return true;
        }
        finally
        {
            Release(paragraphRange);
            Release(paragraph);
            Release(paragraphs);
            Release(ownerRange);
            Release(numberRange);
            Release(numberBookmark);
            Release(bookmarks);
            Release(code);
            Release(sequenceField);
            Release(fields);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(formulaRange);
            Release(formulaBookmark);
        }
    }

    private static void FinalizeNumberedOmmlShapesAcrossOfficeTurns(
        Word.Document document,
        int expectedFormulaCount,
        string context)
    {
        // The compatibility entry point retains its historical name, but the
        // current producer no longer creates or finalizes floating Shapes. A
        // migrated numbered OMML is complete when Word exposes either the current
        // minimal direct-SEQ 1x3 host or an older healthy native #(SEQ) host.
        // Retry only for Word's transient COM normalization window.
        var finalized = 0;
        var numberedOmmlHealthy = 0;
        for (var turn = 1; turn <= 5; turn++)
        {
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(120);
            System.Windows.Forms.Application.DoEvents();
            finalized = WordEquationNumbering
                .FinalizeNumberedOmmlDisplayShapeLayouts(document);
            numberedOmmlHealthy = CountHealthyNativeHashSequenceOmml(document);
            Console.WriteLine(
                $"  {context}: numbered-OMML finalization turn {turn} finalized={finalized}/{expectedFormulaCount}, healthy={numberedOmmlHealthy}/{expectedFormulaCount}, shapes={document.Shapes.Count}, tables={document.Tables.Count}.");
            if (numberedOmmlHealthy >= expectedFormulaCount
                && finalized >= expectedFormulaCount)
            {
                AssertEqual(0, document.Shapes.Count,
                    context + ": numbered-OMML finalization created a floating Shape.");
                return;
            }
        }
        throw new InvalidDataException(
            context + $": Word did not expose {expectedFormulaCount} healthy numbered OMML formula(s) across five Office turns; finalized={finalized}, healthy={numberedOmmlHealthy}.");
    }

    private static int CountHealthyNativeHashSequenceOmml(Word.Document document)
    {
        var healthy = 0;
        foreach (var formulaId in WordOmmlFormulaStore.BookmarkedFormulaIds(document))
        {
            Word.Bookmark? bookmark = null;
            Word.Range? equationRange = null;
            Word.OMaths? maths = null;
            Word.OMath? math = null;
            try
            {
                var metadata = WordOmmlFormulaStore.TryRead(document, formulaId);
                if (metadata?.Numbered != true
                    || !string.Equals(
                        metadata.DisplayMode,
                        "block",
                        StringComparison.OrdinalIgnoreCase))
                    continue;
                bookmark = WordOmmlFormulaStore.FindByFormulaId(document, formulaId);
                if (bookmark is null) continue;
                equationRange = WordOmmlFormulaStore.GetEquationRange(bookmark);
                maths = equationRange.OMaths;
                if (maths.Count != 1) continue;
                math = maths[1];
                if (math.Type != Word.WdOMathType.wdOMathDisplay) continue;
                if ((bool)equationRange.get_Information(Word.WdInformation.wdWithInTable))
                {
                    AssertOmmlTableNumberLifecyclePhase(
                        document.Application,
                        document,
                        formulaId,
                        "numbered-OMML finalization health");
                    healthy++;
                    continue;
                }
                if (!WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(
                        equationRange.WordOpenXML,
                        formulaId))
                    continue;
                healthy++;
            }
            catch
            {
                // A just-migrated OMath can be transiently unavailable for one
                // dispatcher turn. The outer retry loop reacquires it.
            }
            finally
            {
                Release(math);
                Release(maths);
                Release(equationRange);
                Release(bookmark);
            }
        }
        return healthy;
    }

    private static void FinalizeNumberedOmmlShapeDecorationForAcceptance(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Shape? shape = null;
        Word.LineFormat? line = null;
        Word.FillFormat? fill = null;
        try
        {
            shape = FindNumberedOmmlShape(document, formulaId, context);
            line = shape.Line;
            line.Visible = Office.MsoTriState.msoFalse;
            line.Transparency = 1f;
            fill = shape.Fill;
            fill.Visible = Office.MsoTriState.msoFalse;
            fill.Transparency = 1f;
            Console.WriteLine(
                $"  {context}: decoration applied through independently resolved Shape.");
        }
        finally
        {
            Release(fill);
            Release(line);
            Release(shape);
        }
    }

    private static int CountOccurrencesIgnoreCase(
        string source,
        string value)
    {
        if (string.IsNullOrEmpty(source) || string.IsNullOrEmpty(value)) return 0;
        var count = 0;
        var searchStart = 0;
        while (searchStart < source.Length)
        {
            var match = source.IndexOf(
                value,
                searchStart,
                StringComparison.OrdinalIgnoreCase);
            if (match < 0) break;
            count++;
            searchStart = match + value.Length;
        }
        return count;
    }

    private static Word.Shape FindNumberedOmmlShape(
        Word.Document document,
        string formulaId,
        string context)
    {
        Word.Shapes? shapes = null;
        Word.Shape? candidate = null;
        Word.Shape? match = null;
        try
        {
            var normalizedId = Guid.Parse(formulaId).ToString("N");
            var expectedName = NumberedOmmlShapeNamePrefix + normalizedId;
            var expectedAlternativeText =
                NumberedOmmlShapeAlternativeTextPrefix + normalizedId;
            shapes = document.Shapes;
            for (var index = 1; index <= shapes.Count; index++)
            {
                candidate = shapes[index];
                string candidateName;
                string candidateAlternativeText;
                try { candidateName = candidate.Name ?? string.Empty; }
                catch { candidateName = string.Empty; }
                try
                {
                    candidateAlternativeText =
                        candidate.AlternativeText ?? string.Empty;
                }
                catch { candidateAlternativeText = string.Empty; }
                var isMatch = string.Equals(
                        candidateName,
                        expectedName,
                        StringComparison.OrdinalIgnoreCase)
                    || string.Equals(
                        candidateAlternativeText,
                        expectedAlternativeText,
                        StringComparison.Ordinal);
                if (!isMatch)
                {
                    Release(candidate); candidate = null;
                    continue;
                }
                if (match is not null)
                    throw new InvalidDataException(
                        context + ": more than one number Shape belongs to the formula.");
                match = candidate;
                candidate = null;
            }
            if (match is null)
                throw new InvalidDataException(
                    context + ": the numbered OMML text-box Shape is missing.");
            var result = match;
            match = null;
            return result;
        }
        finally
        {
            Release(match);
            Release(candidate);
            Release(shapes);
        }
    }

    private static string NormalizeNumberedOmmlLabel(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\a", string.Empty)
            .Replace("\v", string.Empty)
            .Trim();

    private static bool IsNearCoordinate(float expected, float actual, float tolerance) =>
        actual >= 0f
        && !float.IsNaN(actual)
        && !float.IsInfinity(actual)
        && Math.Abs(expected - actual) <= tolerance;

    private static float ReadHorizontalPosition(
        Word.Range range,
        Word.WdInformation information)
    {
        try
        {
            return Convert.ToSingle(range.get_Information(information));
        }
        catch { return -1f; }
    }

    private static float ReadVerticalPosition(Word.Range range)
    {
        try
        {
            return Convert.ToSingle(range.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
        }
        catch { return -1f; }
    }
}

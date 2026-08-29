using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string TrueDisplayShapeAnchorBookmark = "VTShapeAnchor";
    private const string TrueDisplayShapeFormulaBookmark = "VTShapeFormula";
    private const string TrueDisplayShapeReferenceBookmark = "VTShapeReference";
    private const string TrueDisplayShapeTargetBookmark = "VTShapeTarget";
    private const string TrueDisplayShapeName = "VisualTeXTrueDisplayNumber";
    private const string TrueDisplayShapeAlternativeText =
        "VisualTeX genuine display equation number REF host";

    private static void RunWordOmmlTrueDisplayShapePrototypeAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-shape-prototype-source.docx");
        var roundTripPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-shape-prototype-roundtrip.docx");
        var initialPdfPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-shape-prototype-initial.pdf");
        var roundTripPdfPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-shape-prototype-roundtrip.pdf");
        var semanticOmml = WordOmmlConverter.ExtractSingleOMath(
            WordOmmlConverter.TransformMathMlToOmml(QuadraticFormulaMathMl()));
        WriteTrueDisplayShapeProbeDocx(sourcePath, semanticOmml);

        Word.Application? application = null;
        Word.Document? document = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(
                sourcePath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            document.Fields.Update();
            CreateTrueDisplayReferenceShape(document, "initial shape layout");
            UpdateTrueDisplayShapeFields(document, "initial shape layout");
            ExportTrueDisplayShapeDiagnostic(document, initialPdfPath, "initial shape layout");
            AssertTrueDisplayShapeProbe(
                application,
                document,
                semanticOmml,
                "initial shape layout");
            document.SaveAs2(
                roundTripPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                roundTripPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            document.Fields.Update();
            UpdateTrueDisplayShapeFields(
                document,
                "shape save/reopen normalization");
            ExportTrueDisplayShapeDiagnostic(
                document,
                roundTripPdfPath,
                "shape save/reopen normalization");
            AssertTrueDisplayShapeProbe(
                application,
                document,
                semanticOmml,
                "shape save/reopen normalization");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            Console.WriteLine(
                "Word true-display shape prototype acceptance passed: the formula remained genuine wdOMathDisplay/m:oMathPara, the ordinary dynamic REF stayed outside OMath in a table-free Word text-box story, and the overlay host survived F9 plus save/reopen.");
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

    private static void CreateTrueDisplayReferenceShape(
        Word.Document document,
        string context)
    {
        Word.Bookmark? targetBookmark = null;
        Word.Range? targetBookmarkRange = null;
        Word.Paragraphs? targetParagraphs = null;
        Word.Paragraph? targetParagraph = null;
        Word.Range? targetParagraphRange = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Range? formulaBookmarkRange = null;
        Word.Paragraphs? formulaParagraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaParagraphRange = null;
        Word.Shapes? shapes = null;
        Word.Shape? shape = null;
        Word.TextFrame? textFrame = null;
        Word.Range? textRange = null;
        Word.Range? insertionRange = null;
        Word.Field? referenceField = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Word.WrapFormat? wrapFormat = null;
        Word.LineFormat? line = null;
        Word.FillFormat? fill = null;
        object? anchor = null;
        try
        {
            targetBookmark = document.Bookmarks[TrueDisplayShapeAnchorBookmark];
            targetBookmarkRange = targetBookmark.Range;
            targetParagraphs = targetBookmarkRange.Paragraphs;
            AssertEqual(1, targetParagraphs.Count,
                context + ": the dedicated Shape anchor is not in exactly one paragraph.");
            targetParagraph = targetParagraphs[1];
            targetParagraphRange = targetParagraph.Range;

            formulaBookmark = document.Bookmarks[TrueDisplayShapeFormulaBookmark];
            formulaBookmarkRange = formulaBookmark.Range;
            formulaParagraphs = formulaBookmarkRange.Paragraphs;
            AssertEqual(1, formulaParagraphs.Count,
                context + ": the display formula is not in exactly one paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            shapes = document.Shapes;
            AssertEqual(0, shapes.Count,
                context + ": the source probe unexpectedly already contains a Shape.");

            anchor = targetParagraphRange;
            shape = shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                0f,
                0f,
                72f,
                30f,
                ref anchor);
            shape.Name = TrueDisplayShapeName;
            shape.AlternativeText = TrueDisplayShapeAlternativeText;
            shape.RelativeHorizontalPosition =
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            shape.Left = (float)Word.WdShapePosition.wdShapeRight;
            shape.RelativeVerticalPosition =
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
            shape.Top = 0f;
            shape.LockAnchor = -1;
            wrapFormat = shape.WrapFormat;
            wrapFormat.Type = Word.WdWrapType.wdWrapNone;
            wrapFormat.AllowOverlap = -1;
            line = shape.Line;
            line.Visible = Office.MsoTriState.msoFalse;
            fill = shape.Fill;
            fill.Visible = Office.MsoTriState.msoFalse;

            textFrame = shape.TextFrame;
            textFrame.MarginLeft = 0f;
            textFrame.MarginRight = 0f;
            textFrame.MarginTop = 0f;
            textFrame.MarginBottom = 0f;
            textFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
            textRange = textFrame.TextRange;
            textRange.Text = "()";
            paragraphFormat = textRange.ParagraphFormat;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            paragraphFormat.SpaceBefore = 0f;
            paragraphFormat.SpaceAfter = 0f;
            paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            font = textRange.Font;
            font.Name = "Times New Roman";
            font.NameAscii = "Times New Roman";
            font.NameFarEast = "Microsoft YaHei";
            font.NameOther = "Times New Roman";
            font.Size = 11f;
            font.Position = 0;
            font.Hidden = 0;
            font.Color = Word.WdColor.wdColorAutomatic;

            insertionRange = textRange.Duplicate;
            insertionRange.SetRange(textRange.Start + 1, textRange.Start + 1);
            referenceField = textRange.Fields.Add(
                insertionRange,
                Word.WdFieldType.wdFieldRef,
                TrueDisplayShapeTargetBookmark + " \\h \\* CHARFORMAT",
                PreserveFormatting: true);
            referenceField.Update();
            Word.Bookmark? visibleNumberBookmark = null;
            try
            {
                visibleNumberBookmark = document.Bookmarks.Add(
                    "VTEq_11111111111111111111111111111111",
                    textRange);
            }
            finally { Release(visibleNumberBookmark); }
            Console.WriteLine(
                $"  {context}: created Shape name='{shape.Name}', left={shape.Left:0.###}, top={shape.Top:0.###}, size={shape.Width:0.###}x{shape.Height:0.###}pt, text='{NormalizeTrueDisplayShapeText(textRange.Text)}'.");
        }
        finally
        {
            Release(fill);
            Release(line);
            Release(wrapFormat);
            Release(font);
            Release(paragraphFormat);
            Release(referenceField);
            Release(insertionRange);
            Release(textRange);
            Release(textFrame);
            Release(shape);
            Release(shapes);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(formulaBookmarkRange);
            Release(formulaBookmark);
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
            Release(targetBookmarkRange);
            Release(targetBookmark);
        }
    }

    private static void UpdateTrueDisplayShapeFields(
        Word.Document document,
        string context)
    {
        Word.Shape? shape = null;
        Word.TextFrame? textFrame = null;
        Word.Range? textRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        try
        {
            shape = FindTrueDisplayReferenceShape(document, context);
            textFrame = shape.TextFrame;
            textRange = textFrame.TextRange;
            fields = textRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": the Shape does not contain exactly one REF field.");
            field = fields[1];
            field.Update();
            Console.WriteLine(
                $"  {context}: updated text-box REF result='{NormalizeTrueDisplayShapeText(field.Result.Text)}'.");
        }
        finally
        {
            Release(field);
            Release(fields);
            Release(textRange);
            Release(textFrame);
            Release(shape);
        }
    }

    private static Word.Shape FindTrueDisplayReferenceShape(
        Word.Document document,
        string context)
    {
        Word.Shapes? shapes = null;
        Word.Shape? candidate = null;
        try
        {
            shapes = document.Shapes;
            Word.Shape? match = null;
            for (var index = 1; index <= shapes.Count; index++)
            {
                candidate = shapes[index];
                var isMatch = string.Equals(
                    candidate.AlternativeText,
                    TrueDisplayShapeAlternativeText,
                    StringComparison.Ordinal)
                    || string.Equals(
                        candidate.Name,
                        TrueDisplayShapeName,
                        StringComparison.Ordinal);
                if (isMatch)
                {
                    if (match is not null)
                    {
                        Release(match);
                        throw new InvalidDataException(
                            context + ": more than one true-display number Shape exists.");
                    }
                    match = candidate;
                    candidate = null;
                }
                Release(candidate); candidate = null;
            }
            if (match is null)
                throw new InvalidDataException(
                    context + ": the true-display number Shape is missing.");
            return match;
        }
        finally
        {
            Release(candidate);
            Release(shapes);
        }
    }

    private static void ExportTrueDisplayShapeDiagnostic(
        Word.Document document,
        string path,
        string context)
    {
        document.ExportAsFixedFormat(
            path,
            Word.WdExportFormat.wdExportFormatPDF,
            OpenAfterExport: false,
            OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
            Range: Word.WdExportRange.wdExportAllDocument,
            Item: Word.WdExportItem.wdExportDocumentContent,
            IncludeDocProps: false,
            KeepIRM: false,
            CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
            DocStructureTags: true,
            BitmapMissingFonts: true,
            UseISO19005_1: false);
        Console.WriteLine(
            $"  {context}: exported fixed-layout diagnostic '{path}'.");
    }

    private static void AssertTrueDisplayShapeProbe(
        Word.Application application,
        Word.Document document,
        string semanticOmml,
        string context)
    {
        Word.Bookmark? targetBookmark = null;
        Word.Range? targetBookmarkRange = null;
        Word.Paragraphs? targetParagraphs = null;
        Word.Paragraph? targetParagraph = null;
        Word.Range? targetParagraphRange = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Bookmark? referenceBookmark = null;
        Word.Range? formulaBookmarkRange = null;
        Word.Range? referenceBookmarkRange = null;
        Word.OMaths? formulaMaths = null;
        Word.OMaths? referenceMaths = null;
        Word.OMath? formulaMath = null;
        Word.OMath? referenceMath = null;
        Word.Range? formulaMathRange = null;
        Word.Range? referenceMathRange = null;
        Word.Paragraphs? formulaParagraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaParagraphRange = null;
        Word.Shape? shape = null;
        Word.TextFrame? textFrame = null;
        Word.Range? textRange = null;
        Word.Fields? fields = null;
        Word.Field? field = null;
        Word.Range? fieldCode = null;
        Word.Range? fieldResult = null;
        Word.WrapFormat? wrapFormat = null;
        Word.Range? anchor = null;
        Word.Window? window = null;
        try
        {
            AssertEqual(0, document.Tables.Count,
                context + ": the true-display Shape prototype created a Word table.");
            targetBookmark = document.Bookmarks[TrueDisplayShapeAnchorBookmark];
            targetBookmarkRange = targetBookmark.Range;
            targetParagraphs = targetBookmarkRange.Paragraphs;
            AssertEqual(1, targetParagraphs.Count,
                context + ": the dedicated Shape anchor is not in exactly one paragraph.");
            targetParagraph = targetParagraphs[1];
            targetParagraphRange = targetParagraph.Range;

            formulaBookmark = document.Bookmarks[TrueDisplayShapeFormulaBookmark];
            referenceBookmark = document.Bookmarks[TrueDisplayShapeReferenceBookmark];
            formulaBookmarkRange = formulaBookmark.Range;
            referenceBookmarkRange = referenceBookmark.Range;
            formulaMaths = formulaBookmarkRange.OMaths;
            referenceMaths = referenceBookmarkRange.OMaths;
            AssertEqual(1, formulaMaths.Count,
                context + ": the numbered formula bookmark does not contain exactly one OMath.");
            AssertEqual(1, referenceMaths.Count,
                context + ": the plain comparison bookmark does not contain exactly one OMath.");
            formulaMath = formulaMaths[1];
            referenceMath = referenceMaths[1];
            formulaMathRange = formulaMath.Range;
            referenceMathRange = referenceMath.Range;
            AssertEqual(Word.WdOMathType.wdOMathDisplay, formulaMath.Type,
                context + ": the numbered formula is not genuine Word display math.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, referenceMath.Type,
                context + ": the comparison formula is not genuine Word display math.");
            AssertEqual(0, formulaMathRange.Fields.Count,
                context + ": the external REF leaked into the formula OMath range.");

            formulaParagraphs = formulaMathRange.Paragraphs;
            AssertEqual(1, formulaParagraphs.Count,
                context + ": the display formula spans more than one paragraph.");
            formulaParagraph = formulaParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            AssertTrue(!(bool)formulaParagraphRange.get_Information(
                    Word.WdInformation.wdWithInTable),
                context + ": the display formula paragraph is inside a table.");

            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var word = (XNamespace)WordNamespace;
            var math = (XNamespace)MathNamespace;
            var formulaXml = XDocument.Parse(
                formulaParagraphRange.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            AssertEqual(1, formulaXml.Descendants(math + "oMathPara").Count(),
                context + ": the formula paragraph did not retain exactly one m:oMathPara.");
            AssertEqual(1, formulaXml.Descendants(math + "oMath").Count(),
                context + ": the formula paragraph did not retain exactly one m:oMath.");
            AssertTrue(!formulaXml.Descendants(math + "eqArr").Any(),
                context + ": the obsolete m:eqArr numbering wrapper returned.");
            AssertTrue(!formulaXml.Descendants(word + "fldChar").Any(),
                context + ": field controls leaked into the display formula paragraph.");
            AssertTrue(!formulaXml.Descendants(word + "instrText").Any(node =>
                    node.Value.IndexOf("REF ", StringComparison.OrdinalIgnoreCase) >= 0),
                context + ": a REF instruction leaked into the display formula paragraph.");

            shape = FindTrueDisplayReferenceShape(document, context);
            AssertEqual(Office.MsoShapeType.msoTextBox, shape.Type,
                context + ": the number host is not a Word text box Shape.");
            AssertEqual(
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin,
                shape.RelativeHorizontalPosition,
                context + ": the number Shape is not positioned relative to the margin.");
            AssertNear(
                (float)Word.WdShapePosition.wdShapeRight,
                shape.Left,
                0.1f,
                context + ": the number Shape is not aligned to the right margin.");
            AssertEqual(
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph,
                shape.RelativeVerticalPosition,
                context + ": the number Shape is not positioned relative to its hidden anchor paragraph.");
            AssertNear(0f, shape.Top, 0.1f,
                context + ": the number Shape vertical offset changed.");
            wrapFormat = shape.WrapFormat;
            AssertEqual(Word.WdWrapType.wdWrapNone, wrapFormat.Type,
                context + ": the number Shape unexpectedly wraps body text.");
            AssertTrue(wrapFormat.AllowOverlap != 0,
                context + ": the number Shape does not allow overlap.");
            anchor = shape.Anchor;
            AssertTrue(anchor.Start >= targetParagraphRange.Start
                    && anchor.Start <= targetParagraphRange.End,
                context + ": the number Shape anchor is not in the dedicated anchor paragraph.");
            AssertTrue(anchor.Start < formulaParagraphRange.Start,
                context + ": the number Shape anchor leaked into the display formula paragraph.");

            textFrame = shape.TextFrame;
            textRange = textFrame.TextRange;
            AssertEqual(Word.WdStoryType.wdTextFrameStory, textRange.StoryType,
                context + ": the number REF is not in a Word text-frame story.");
            fields = textRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": the number Shape does not contain exactly one field.");
            field = fields[1];
            fieldCode = field.Code;
            AssertTrue((fieldCode.Text ?? string.Empty).IndexOf(
                    "REF " + TrueDisplayShapeTargetBookmark,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the Shape field is not the expected dynamic REF.");
            field.Update();
            fieldResult = field.Result;
            AssertEqual("1", (fieldResult.Text ?? string.Empty).Trim(),
                context + ": F9 did not preserve the hidden SEQ result.");
            AssertEqual("(1)", NormalizeTrueDisplayShapeText(textRange.Text),
                context + ": the visible number label is not parenthesized.");
            AssertTrue(document.Bookmarks.Exists("VTEq_11111111111111111111111111111111"),
                context + ": the visible-number bookmark did not survive in the text-box story.");
            Word.Range? documentContent = null;
            try
            {
                documentContent = document.Content;
                var contentXml = documentContent.WordOpenXML ?? string.Empty;
                AssertTrue(contentXml.IndexOf(
                        "VTEq_11111111111111111111111111111111",
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    context + ": the text-box visible-number bookmark is absent from document WordOpenXML.");
                AssertTrue(contentXml.IndexOf("<w:txbxContent", StringComparison.OrdinalIgnoreCase) >= 0,
                    context + ": the visible-number Shape did not serialize as a Word text-box story.");
            }
            finally { Release(documentContent); }

            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(formulaMathRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(100);
            var formulaBox = ReadVisibleMathInkBox(
                document,
                window,
                formulaMathRange,
                context + " genuine display formula ink");
            var referenceBox = ReadVisibleMathInkBox(
                document,
                window,
                referenceMathRange,
                context + " plain display reference ink");
            AssertTrue(Math.Abs(formulaBox.Width - referenceBox.Width) <= 3,
                context + ": the numbered and plain display widths diverged.");
            AssertTrue(Math.Abs(formulaBox.Height - referenceBox.Height) <= 3,
                context + ": the numbered and plain display heights diverged.");
            var numberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " visible Shape REF result");
            var formulaCenterY = formulaBox.Top + formulaBox.Height / 2.0;
            var numberCenterY = numberBox.Top + numberBox.Height / 2.0;
            var verticalCenterDelta = numberCenterY - formulaCenterY;
            var formulaTopPoints = Convert.ToSingle(formulaMathRange.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            var numberTopPoints = Convert.ToSingle(fieldResult.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            Console.WriteLine(
                $"  {context}: Shape REF box={numberBox.Left},{numberBox.Top},{numberBox.Width},{numberBox.Height}, formula/number centers={formulaCenterY:0.##}/{numberCenterY:0.##}, delta={verticalCenterDelta:0.##}px, topPoints={formulaTopPoints:0.###}/{numberTopPoints:0.###}.");
            AssertTrue(Math.Abs(verticalCenterDelta) <= 4.0,
                context + $": the Shape REF is not vertically centered on the true-display formula (delta={verticalCenterDelta:0.##}px).");

            var candidateMetric = ReadWordOmmlLayoutMetric(
                application,
                document,
                formulaMathRange,
                context + " candidate metric",
                semanticOmml,
                measureVisibleCharacterInk: true);
            var referenceMetric = ReadWordOmmlLayoutMetric(
                application,
                document,
                referenceMathRange,
                context + " reference metric",
                semanticOmml,
                measureVisibleCharacterInk: true);
            AssertEqual(referenceMetric.Type, candidateMetric.Type,
                context + ": candidate/reference OMath types differ.");
            AssertTrue(Math.Abs(candidateMetric.WidthPx - referenceMetric.WidthPx) <= 3,
                context + ": candidate/reference visible widths differ.");
            AssertTrue(Math.Abs(candidateMetric.HeightPx - referenceMetric.HeightPx) <= 3,
                context + ": candidate/reference visible heights differ.");
            Console.WriteLine(
                $"  {context}: formula={formulaBox.Left},{formulaBox.Top},{formulaBox.Width},{formulaBox.Height}, reference={referenceBox.Left},{referenceBox.Top},{referenceBox.Width},{referenceBox.Height}, shape={shape.Left:0.###},{shape.Top:0.###},{shape.Width:0.###},{shape.Height:0.###}pt, label='{NormalizeTrueDisplayShapeText(textRange.Text)}'.");
        }
        finally
        {
            Release(window);
            Release(anchor);
            Release(wrapFormat);
            Release(fieldResult);
            Release(fieldCode);
            Release(field);
            Release(fields);
            Release(textRange);
            Release(textFrame);
            Release(shape);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(referenceMathRange);
            Release(formulaMathRange);
            Release(referenceMath);
            Release(formulaMath);
            Release(referenceMaths);
            Release(formulaMaths);
            Release(referenceBookmarkRange);
            Release(formulaBookmarkRange);
            Release(referenceBookmark);
            Release(formulaBookmark);
            Release(targetParagraphRange);
            Release(targetParagraph);
            Release(targetParagraphs);
            Release(targetBookmarkRange);
            Release(targetBookmark);
        }
    }

    private static void WriteTrueDisplayShapeProbeDocx(
        string path,
        string semanticOmml)
    {
        if (File.Exists(path)) File.Delete(path);
        using var stream = new FileStream(
            path,
            FileMode.CreateNew,
            FileAccess.ReadWrite,
            FileShare.None);
        using var archive = new ZipArchive(
            stream,
            ZipArchiveMode.Create,
            leaveOpen: false);

        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        const string RelationshipsNamespace =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        const string ContentTypesNamespace =
            "http://schemas.openxmlformats.org/package/2006/content-types";

        WriteTrueDisplayShapeProbeEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Types xmlns=\"{ContentTypesNamespace}\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>"
            + "</Types>");
        WriteTrueDisplayShapeProbeEntry(
            archive,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Relationships xmlns=\"{RelationshipsNamespace}\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>");
        WriteTrueDisplayShapeProbeEntry(
            archive,
            "word/_rels/document.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Relationships xmlns=\"{RelationshipsNamespace}\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>"
            + "</Relationships>");
        WriteTrueDisplayShapeProbeEntry(
            archive,
            "word/settings.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:settings xmlns:w=\"{WordNamespace}\">"
            + "<w:updateFields w:val=\"true\"/>"
            + "</w:settings>");

        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + "<w:body>"
            + "<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>"
            + $"<w:bookmarkStart w:id=\"1\" w:name=\"{TrueDisplayShapeTargetBookmark}\"/>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:instrText xml:space=\"preserve\"> SEQ VTShape \\* ARABIC </w:instrText></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:bookmarkEnd w:id=\"1\"/>"
            + "</w:p>"
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"20\" w:lineRule=\"exact\"/><w:rPr><w:sz w:val=\"2\"/><w:szCs w:val=\"2\"/></w:rPr></w:pPr>"
            + $"<w:bookmarkStart w:id=\"2\" w:name=\"{TrueDisplayShapeAnchorBookmark}\"/>"
            + "<w:bookmarkEnd w:id=\"2\"/>"
            + "</w:p>"
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"
            + $"<w:bookmarkStart w:id=\"3\" w:name=\"{TrueDisplayShapeFormulaBookmark}\"/>"
            + "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>"
            + semanticOmml
            + "</m:oMathPara>"
            + "<w:bookmarkEnd w:id=\"3\"/>"
            + "</w:p>"
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"
            + $"<w:bookmarkStart w:id=\"4\" w:name=\"{TrueDisplayShapeReferenceBookmark}\"/>"
            + "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>"
            + semanticOmml
            + "</m:oMathPara>"
            + "<w:bookmarkEnd w:id=\"4\"/>"
            + "</w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
            + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
            + "</w:sectPr></w:body></w:document>";
        WriteTrueDisplayShapeProbeEntry(archive, "word/document.xml", documentXml);
    }

    private static string NormalizeTrueDisplayShapeText(string? value) =>
        (value ?? string.Empty)
            .Replace("\r", string.Empty)
            .Replace("\a", string.Empty)
            .Replace("\v", string.Empty)
            .Trim();

    private static void WriteTrueDisplayShapeProbeEntry(
        ZipArchive archive,
        string path,
        string content)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var stream = entry.Open();
        using var writer = new StreamWriter(
            stream,
            new UTF8Encoding(false));
        writer.Write(content);
    }
}

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string TrueDisplayFrameFormulaBookmark = "VTFrameFormula";
    private const string TrueDisplayFrameReferenceBookmark = "VTFrameReference";
    private const string TrueDisplayFrameTargetBookmark = "VTFrameTarget";
    private const string TrueDisplayFrameVisibleBookmark = "VTFrameVisible";

    private static void RunWordOmmlTrueDisplayFramePrototypeAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-frame-prototype-source.docx");
        var roundTripPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-frame-prototype-roundtrip.docx");
        var diagnosticPdfPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-frame-prototype-initial.pdf");
        var roundTripPdfPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-frame-prototype-roundtrip.pdf");
        var semanticOmml = WordOmmlConverter.ExtractSingleOMath(
            WordOmmlConverter.TransformMathMlToOmml(QuadraticFormulaMathMl()));
        WriteTrueDisplayFrameProbeDocx(sourcePath, semanticOmml);

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
            ConfigureTrueDisplayReferenceFrame(document, "initial frame layout");
            document.ExportAsFixedFormat(
                diagnosticPdfPath,
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
            Console.WriteLine($"  initial frame layout: exported fixed-layout diagnostic '{diagnosticPdfPath}'.");
            AssertTrueDisplayFrameProbe(
                application,
                document,
                semanticOmml,
                "initial frame layout");
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
            document.ExportAsFixedFormat(
                roundTripPdfPath,
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
            Console.WriteLine($"  frame save/reopen normalization: exported fixed-layout diagnostic '{roundTripPdfPath}'.");
            AssertTrueDisplayFrameProbe(
                application,
                document,
                semanticOmml,
                "frame save/reopen normalization");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            Console.WriteLine(
                "Word true-display frame prototype acceptance passed: the formula remained genuine wdOMathDisplay/m:oMathPara, the ordinary dynamic REF stayed outside OMath in a table-free Word Frame, and visual centering/right alignment survived F9 plus save/reopen.");
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

    private static void ConfigureTrueDisplayReferenceFrame(
        Word.Document document,
        string context)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Bookmark? visibleBookmark = null;
        Word.Range? formulaRange = null;
        Word.Range? visibleRange = null;
        Word.Paragraphs? visibleParagraphs = null;
        Word.Paragraph? visibleParagraph = null;
        Word.Range? visibleParagraphRange = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Microsoft.Office.Interop.Word.Font? visibleFont = null;
        Word.Frames? frames = null;
        Word.Frame? frame = null;
        Word.Borders? borders = null;
        try
        {
            formulaBookmark = document.Bookmarks[TrueDisplayFrameFormulaBookmark];
            formulaRange = formulaBookmark.Range;
            visibleBookmark = document.Bookmarks[TrueDisplayFrameVisibleBookmark];
            visibleRange = visibleBookmark.Range;
            visibleParagraphs = visibleRange.Paragraphs;
            AssertEqual(1, visibleParagraphs.Count,
                context + ": the visible REF label is not in exactly one paragraph.");
            visibleParagraph = visibleParagraphs[1];
            visibleParagraphRange = visibleParagraph.Range;

            paragraphFormat = visibleParagraphRange.ParagraphFormat;
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            paragraphFormat.SpaceBefore = 0f;
            paragraphFormat.SpaceAfter = 0f;
            paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            paragraphFormat.KeepTogether = -1;
            paragraphFormat.KeepWithNext = 0;
            paragraphFormat.WidowControl = 0;
            visibleFont = visibleRange.Font;
            visibleFont.Hidden = 0;
            visibleFont.Size = 11f;
            visibleFont.Color = Word.WdColor.wdColorAutomatic;
            visibleFont.Position = 0;

            frames = visibleParagraphRange.Frames;
            frame = frames.Count > 0
                ? frames[1]
                : frames.Add(visibleParagraphRange);
            frame.WidthRule = Word.WdFrameSizeRule.wdFrameExact;
            frame.HeightRule = Word.WdFrameSizeRule.wdFrameExact;
            frame.Width = 72f;
            frame.Height = 18f;
            frame.RelativeHorizontalPosition =
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            frame.RelativeVerticalPosition =
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
            frame.HorizontalPosition = (float)Word.WdFramePosition.wdFrameRight;
            frame.VerticalPosition = 0f;
            frame.TextWrap = false;
            frame.LockAnchor = true;
            borders = visibleParagraphRange.Borders;
            borders.Enable = 0;
            CalibrateTrueDisplayReferenceFrameVerticalPosition(
                document,
                formulaRange,
                frame,
                context);

            Console.WriteLine(
                $"  {context}: configured genuine-display REF Frame at calibrated verticalPosition={frame.VerticalPosition:0.###}pt.");
        }
        finally
        {
            Release(borders);
            Release(frame);
            Release(frames);
            Release(visibleFont);
            Release(paragraphFormat);
            Release(visibleParagraphRange);
            Release(visibleParagraph);
            Release(visibleParagraphs);
            Release(visibleRange);
            Release(formulaRange);
            Release(visibleBookmark);
            Release(formulaBookmark);
        }
    }

    private static void CalibrateTrueDisplayReferenceFrameVerticalPosition(
        Word.Document document,
        Word.Range formulaBookmarkRange,
        Word.Frame frame,
        string context)
    {
        Word.Bookmark? visibleBookmark = null;
        Word.Range? visibleRange = null;
        Word.Fields? fields = null;
        Word.Field? referenceField = null;
        Word.Range? fieldResult = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? frameRange = null;
        Word.Window? window = null;
        try
        {
            visibleBookmark = document.Bookmarks[TrueDisplayFrameVisibleBookmark];
            visibleRange = visibleBookmark.Range;
            fields = visibleRange.Fields;
            AssertEqual(1, fields.Count,
                context + ": the floating label does not contain exactly one REF field.");
            referenceField = fields[1];
            referenceField.Update();
            fieldResult = referenceField.Result;
            maths = formulaBookmarkRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": the formula bookmark does not contain exactly one OMath.");
            math = maths[1];
            mathRange = math.Range;
            frameRange = frame.Range;
            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(mathRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(100);

            frame.VerticalPosition = 0f;
            window.ScrollIntoView(mathRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(120);
            WordRangePixelBox formulaBox;
            try
            {
                formulaBox = ReadVisibleMathInkBox(
                    document,
                    window,
                    mathRange,
                    context + " display formula calibration ink");
            }
            catch (InvalidDataException)
            {
                // A newly materialized Frame can make Word temporarily reject
                // per-character GetPoint calls inside professional OMML. The OMath
                // Range box is still stable and is sufficient for vertical-center
                // calibration; the final acceptance later rechecks visible glyph ink.
                window.ScrollIntoView(mathRange, ref scrollStart);
                document.Repaginate();
                Thread.Sleep(120);
                formulaBox = ReadWordRangePixelBox(
                    window,
                    mathRange,
                    context + " display formula calibration range");
            }
            window.ScrollIntoView(frameRange, ref scrollStart);
            Thread.Sleep(80);
            Console.WriteLine(
                $"  {context}: formulaPageY={Convert.ToSingle(mathRange.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)):0.##}pt, framePageY={Convert.ToSingle(frameRange.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)):0.##}pt, resultPageY={Convert.ToSingle(fieldResult.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)):0.##}pt, framePos={frame.VerticalPosition:0.##}pt.");
            var baselineNumberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " frame baseline calibration");

            const float CalibrationPoints = 10f;
            frame.VerticalPosition = CalibrationPoints;
            document.Repaginate();
            Thread.Sleep(80);
            window.ScrollIntoView(frameRange, ref scrollStart);
            Thread.Sleep(80);
            var movedNumberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " frame moved calibration");
            var measuredPixelsPerPoint =
                (movedNumberBox.Top - baselineNumberBox.Top) / CalibrationPoints;
            if (Math.Abs(measuredPixelsPerPoint) < 0.1)
                throw new InvalidDataException(
                    context + ": Word Frame vertical position did not affect screen geometry.");
            var absolutePixelsPerPoint = Math.Abs(measuredPixelsPerPoint);
            var formulaHeightPoints = formulaBox.Height / absolutePixelsPerPoint;
            var numberHeightPoints = baselineNumberBox.Height / absolutePixelsPerPoint;
            var calibratedPosition = 0f;
            var formulaPageTop = 0f;
            var actualNumberPageTop = 0f;
            var desiredNumberPageTop = 0.0;
            for (var iteration = 0; iteration < 6; iteration++)
            {
                frame.VerticalPosition = calibratedPosition;
                document.Repaginate();
                Thread.Sleep(120);
                // Converting the paragraph to a Frame can itself move the following
                // display paragraph by roughly one line. Re-read both page positions
                // on every iteration instead of calibrating against the stale
                // pre-Frame formula position.
                formulaPageTop = Convert.ToSingle(mathRange.get_Information(
                    Word.WdInformation.wdVerticalPositionRelativeToPage));
                actualNumberPageTop = Convert.ToSingle(fieldResult.get_Information(
                    Word.WdInformation.wdVerticalPositionRelativeToPage));
                desiredNumberPageTop = formulaPageTop
                    + (formulaHeightPoints - numberHeightPoints) / 2.0;
                var error = desiredNumberPageTop - actualNumberPageTop;
                if (Math.Abs(error) <= 0.35) break;
                calibratedPosition = (float)Math.Max(
                    -120.0,
                    Math.Min(120.0, calibratedPosition + error));
            }
            // The final iteration may compute a corrected position after measuring
            // the previous one. Apply that last value explicitly, then report the
            // geometry Word actually committed.
            frame.VerticalPosition = calibratedPosition;
            document.Repaginate();
            Thread.Sleep(120);
            formulaPageTop = Convert.ToSingle(mathRange.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            actualNumberPageTop = Convert.ToSingle(fieldResult.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            desiredNumberPageTop = formulaPageTop
                + (formulaHeightPoints - numberHeightPoints) / 2.0;
            Console.WriteLine(
                $"  {context}: frame calibration formula={formulaBox.Left},{formulaBox.Top},{formulaBox.Width},{formulaBox.Height}, number0={baselineNumberBox.Left},{baselineNumberBox.Top},{baselineNumberBox.Width},{baselineNumberBox.Height}, pixelsPerPoint={measuredPixelsPerPoint:0.###}, formulaPageY={formulaPageTop:0.###}pt, desiredPageY={desiredNumberPageTop:0.###}pt, actualPageY={actualNumberPageTop:0.###}pt, verticalPosition={frame.VerticalPosition:0.###}pt.");
        }
        finally
        {
            Release(window);
            Release(frameRange);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(fieldResult);
            Release(referenceField);
            Release(fields);
            Release(visibleRange);
            Release(visibleBookmark);
        }
    }

    private static void AssertTrueDisplayFrameProbe(
        Word.Application application,
        Word.Document document,
        string semanticOmml,
        string context)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Bookmark? referenceBookmark = null;
        Word.Bookmark? visibleBookmark = null;
        Word.Range? formulaBookmarkRange = null;
        Word.Range? referenceBookmarkRange = null;
        Word.Range? visibleRange = null;
        Word.OMaths? formulaMaths = null;
        Word.OMaths? referenceMaths = null;
        Word.OMath? formulaMath = null;
        Word.OMath? referenceMath = null;
        Word.Range? formulaMathRange = null;
        Word.Range? referenceMathRange = null;
        Word.Paragraphs? formulaParagraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? formulaParagraphRange = null;
        Word.Fields? visibleFields = null;
        Word.Field? visibleField = null;
        Word.Range? fieldCode = null;
        Word.Range? fieldResult = null;
        Word.Frames? visibleFrames = null;
        Word.Frame? visibleFrame = null;
        Word.Window? window = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Word.Range? numberEnd = null;
        try
        {
            AssertEqual(0, document.Tables.Count,
                context + ": the true-display numbering prototype created a Word table.");
            formulaBookmark = document.Bookmarks[TrueDisplayFrameFormulaBookmark];
            referenceBookmark = document.Bookmarks[TrueDisplayFrameReferenceBookmark];
            visibleBookmark = document.Bookmarks[TrueDisplayFrameVisibleBookmark];
            formulaBookmarkRange = formulaBookmark.Range;
            referenceBookmarkRange = referenceBookmark.Range;
            visibleRange = visibleBookmark.Range;
            formulaMaths = formulaBookmarkRange.OMaths;
            referenceMaths = referenceBookmarkRange.OMaths;
            AssertEqual(1, formulaMaths.Count,
                context + ": the numbered formula bookmark does not contain exactly one OMath.");
            AssertEqual(1, referenceMaths.Count,
                context + ": the plain reference bookmark does not contain exactly one OMath.");
            formulaMath = formulaMaths[1];
            referenceMath = referenceMaths[1];
            formulaMathRange = formulaMath.Range;
            referenceMathRange = referenceMath.Range;
            AssertEqual(Word.WdOMathType.wdOMathDisplay, formulaMath.Type,
                context + ": the numbered formula is not genuine Word display math.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, referenceMath.Type,
                context + ": the plain comparison formula is not display math.");
            AssertEqual(0, formulaMathRange.Fields.Count,
                context + ": the visible REF leaked into the formula OMath range.");

            formulaParagraphs = formulaMathRange.Paragraphs;
            AssertEqual(1, formulaParagraphs.Count,
                context + ": the display formula spans more than one Word paragraph.");
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

            visibleFields = visibleRange.Fields;
            AssertEqual(1, visibleFields.Count,
                context + ": the floating label does not contain exactly one field.");
            visibleField = visibleFields[1];
            fieldCode = visibleField.Code;
            AssertTrue((fieldCode.Text ?? string.Empty).IndexOf(
                    "REF " + TrueDisplayFrameTargetBookmark,
                    StringComparison.OrdinalIgnoreCase) >= 0,
                context + ": the floating label is not the expected dynamic REF field.");
            visibleField.Update();
            fieldResult = visibleField.Result;
            AssertEqual("1", (fieldResult.Text ?? string.Empty).Trim(),
                context + ": F9 did not preserve the live hidden SEQ result.");
            AssertTrue(visibleRange.Start < formulaMathRange.Start
                    || visibleRange.Start >= formulaMathRange.End,
                context + ": the visible REF range intersects the formula OMath.");

            visibleFrames = visibleRange.Frames;
            AssertEqual(1, visibleFrames.Count,
                context + ": the visible REF label is not hosted by exactly one Word Frame.");
            visibleFrame = visibleFrames[1];
            AssertEqual(
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin,
                visibleFrame.RelativeHorizontalPosition,
                context + ": the visible REF Frame is not positioned relative to the margin.");
            AssertEqual(
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph,
                visibleFrame.RelativeVerticalPosition,
                context + ": the visible REF Frame is not anchored relative to the formula paragraph.");
            AssertNear(
                (float)Word.WdFramePosition.wdFrameRight,
                visibleFrame.HorizontalPosition,
                0.1f,
                context + ": the visible REF Frame is not right aligned.");
            AssertTrue(!visibleFrame.TextWrap,
                context + ": the visible REF Frame unexpectedly wraps body text.");

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
            var numberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " floating REF result");
            var referenceBox = ReadVisibleMathInkBox(
                document,
                window,
                referenceMathRange,
                context + " plain display reference ink");
            AssertTrue(Math.Abs(formulaBox.Width - referenceBox.Width) <= 3,
                context + ": the numbered and plain display widths diverged.");
            AssertTrue(Math.Abs(formulaBox.Height - referenceBox.Height) <= 3,
                context + ": the numbered and plain display heights diverged.");
            var formulaCenterX = formulaBox.Left + formulaBox.Width / 2.0;
            var referenceCenterX = referenceBox.Left + referenceBox.Width / 2.0;
            AssertTrue(Math.Abs(formulaCenterX - referenceCenterX) <= 4.0,
                context + ": the numbered formula is not at the native display center.");
            var formulaPageTop = Convert.ToSingle(formulaMathRange.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            var numberPageTop = Convert.ToSingle(fieldResult.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            var pixelsPerPoint = Math.Max(0.25, numberBox.Height / Math.Max(1f, visibleFrame.Height));
            var formulaCenterOnPage = formulaPageTop
                + formulaBox.Height / pixelsPerPoint / 2.0;
            var numberCenterOnPage = numberPageTop
                + numberBox.Height / pixelsPerPoint / 2.0;
            Console.WriteLine(
                $"  {context}: page geometry formulaY={formulaPageTop:0.###}pt, numberY={numberPageTop:0.###}pt, formulaHeight={formulaBox.Height}px, numberHeight={numberBox.Height}px, frameHeight={visibleFrame.Height:0.###}pt, pixelsPerPoint={pixelsPerPoint:0.###}, centers={formulaCenterOnPage:0.###}/{numberCenterOnPage:0.###}pt.");
            var formulaCenterOnScreen = formulaBox.Top + formulaBox.Height / 2.0;
            var numberCenterOnScreen = numberBox.Top + numberBox.Height / 2.0;
            Console.WriteLine(
                $"  {context}: screen centers formula/number={formulaCenterOnScreen:0.###}/{numberCenterOnScreen:0.###}px, delta={numberCenterOnScreen - formulaCenterOnScreen:0.###}px.");
            AssertTrue(
                !float.IsNaN(visibleFrame.VerticalPosition)
                && !float.IsInfinity(visibleFrame.VerticalPosition)
                && visibleFrame.VerticalPosition >= -120f
                && visibleFrame.VerticalPosition <= 120f,
                context + ": the calibrated REF Frame vertical offset is invalid.");
            AssertTrue(Math.Abs(numberCenterOnScreen - formulaCenterOnScreen) <= 4.0,
                context + $": the external REF Frame is not vertically centered on the true-display formula (delta={numberCenterOnScreen - formulaCenterOnScreen:0.###}px).");

            sections = formulaParagraphRange.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            numberEnd = document.Range(fieldResult.End, fieldResult.End);
            var numberRightOnPage = Convert.ToSingle(numberEnd.get_Information(
                Word.WdInformation.wdHorizontalPositionRelativeToPage));
            var expectedRightOnPage = pageSetup.PageWidth - pageSetup.RightMargin;
            Console.WriteLine(
                $"  {context}: COM frame right diagnostic={numberRightOnPage:0.##}/{expectedRightOnPage:0.##}pt; fixed-layout PDF geometry is authoritative for Frame rendering.");

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
                context + ": candidate/reference visible math widths differ.");
            AssertTrue(Math.Abs(candidateMetric.HeightPx - referenceMetric.HeightPx) <= 3,
                context + ": candidate/reference visible math heights differ.");
            Console.WriteLine(
                $"  {context}: formula={formulaBox.Left},{formulaBox.Top},{formulaBox.Width},{formulaBox.Height}, number={numberBox.Left},{numberBox.Top},{numberBox.Width},{numberBox.Height}, frameY={visibleFrame.VerticalPosition:0.###}pt, numberRight={numberRightOnPage:0.##}/{expectedRightOnPage:0.##}, referenceTop={referenceBox.Top}.");
        }
        finally
        {
            Release(numberEnd);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(window);
            Release(visibleFrame);
            Release(visibleFrames);
            Release(fieldResult);
            Release(fieldCode);
            Release(visibleField);
            Release(visibleFields);
            Release(formulaParagraphRange);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(referenceMathRange);
            Release(formulaMathRange);
            Release(referenceMath);
            Release(formulaMath);
            Release(referenceMaths);
            Release(formulaMaths);
            Release(visibleRange);
            Release(referenceBookmarkRange);
            Release(formulaBookmarkRange);
            Release(visibleBookmark);
            Release(referenceBookmark);
            Release(formulaBookmark);
        }
    }

    private static void WriteTrueDisplayFrameProbeDocx(
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

        WriteTrueDisplayFrameProbeEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Types xmlns=\"{ContentTypesNamespace}\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>"
            + "</Types>");
        WriteTrueDisplayFrameProbeEntry(
            archive,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Relationships xmlns=\"{RelationshipsNamespace}\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>");
        WriteTrueDisplayFrameProbeEntry(
            archive,
            "word/_rels/document.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Relationships xmlns=\"{RelationshipsNamespace}\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>"
            + "</Relationships>");
        WriteTrueDisplayFrameProbeEntry(
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
            + $"<w:bookmarkStart w:id=\"1\" w:name=\"{TrueDisplayFrameTargetBookmark}\"/>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:instrText xml:space=\"preserve\"> SEQ VTFrame \\* ARABIC </w:instrText></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:bookmarkEnd w:id=\"1\"/>"
            + "</w:p>"
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"
            + $"<w:bookmarkStart w:id=\"2\" w:name=\"{TrueDisplayFrameFormulaBookmark}\"/>"
            + "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>"
            + semanticOmml
            + "</m:oMathPara>"
            + "<w:bookmarkEnd w:id=\"2\"/>"
            + "</w:p>"
            + "<w:p><w:pPr><w:jc w:val=\"right\"/><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"
            + $"<w:bookmarkStart w:id=\"3\" w:name=\"{TrueDisplayFrameVisibleBookmark}\"/>"
            + "<w:r><w:t>(</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>"
            + $"<w:r><w:instrText xml:space=\"preserve\"> REF {TrueDisplayFrameTargetBookmark} \\h \\* CHARFORMAT </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:t>1</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:r><w:t>)</w:t></w:r>"
            + "<w:bookmarkEnd w:id=\"3\"/>"
            + "</w:p>"
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"
            + $"<w:bookmarkStart w:id=\"4\" w:name=\"{TrueDisplayFrameReferenceBookmark}\"/>"
            + "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>"
            + semanticOmml
            + "</m:oMathPara>"
            + "<w:bookmarkEnd w:id=\"4\"/>"
            + "</w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
            + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
            + "</w:sectPr></w:body></w:document>";
        WriteTrueDisplayFrameProbeEntry(archive, "word/document.xml", documentXml);
    }

    private static void WriteTrueDisplayFrameProbeEntry(
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

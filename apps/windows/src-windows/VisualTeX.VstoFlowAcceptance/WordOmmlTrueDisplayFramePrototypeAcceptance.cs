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
        Word.TabStops? tabStops = null;
        Word.TabStop? rightTab = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
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
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            paragraphFormat.SpaceBefore = 0f;
            paragraphFormat.SpaceAfter = 0f;
            paragraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            paragraphFormat.KeepTogether = -1;
            paragraphFormat.KeepWithNext = 0;
            paragraphFormat.WidowControl = 0;
            paragraphFormat.LeftIndent = 0f;
            paragraphFormat.RightIndent = 0f;
            paragraphFormat.FirstLineIndent = 0f;

            sections = formulaRange.Sections;
            AssertTrue(sections.Count > 0,
                context + ": the display formula has no owning Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            var writableWidth = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            const float numberFrameWidth = 72f;
            tabStops = paragraphFormat.TabStops;
            tabStops.ClearAll();
            rightTab = tabStops.Add(
                numberFrameWidth,
                Word.WdTabAlignment.wdAlignTabRight,
                Word.WdTabLeader.wdTabLeaderSpaces);

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
            frame.Width = numberFrameWidth;
            frame.Height = 18f;
            frame.RelativeHorizontalPosition =
                Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin;
            frame.RelativeVerticalPosition =
                Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph;
            frame.HorizontalPosition = (float)Word.WdFramePosition.wdFrameRight;
            frame.VerticalPosition = 0f;
            frame.HorizontalDistanceFromText = 0f;
            frame.VerticalDistanceFromText = 0f;
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
                $"  {context}: configured right-edge genuine-display REF Frame width={frame.Width:0.###}pt, rightTab={rightTab.Position:0.###}pt, bodyWidth={writableWidth:0.###}pt, calibrated verticalPosition={frame.VerticalPosition:0.###}pt.");
        }
        finally
        {
            Release(borders);
            Release(frame);
            Release(frames);
            Release(visibleFont);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(rightTab);
            Release(tabStops);
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
        Word.ParagraphFormat? frameParagraphFormat = null;
        Microsoft.Office.Interop.Word.Font? visibleFont = null;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
        Microsoft.Office.Interop.Word.Font? leftParenthesisFont = null;
        Microsoft.Office.Interop.Word.Font? rightParenthesisFont = null;
        Word.Range? leftParenthesisRange = null;
        Word.Range? rightParenthesisRange = null;
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
            frameParagraphFormat = frameRange.ParagraphFormat;
            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(mathRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(100);

            var candidates = new[]
            {
                (Name: "paragraph-exact1", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 0f),
                (Name: "paragraph-exact1-dist-2", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: -2f, ExactLine: 0f),
                (Name: "paragraph-exact1-dist-4", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: -4f, ExactLine: 0f),
                (Name: "paragraph-exact1-dist-8", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: -8f, ExactLine: 0f),
                (Name: "paragraph-exact1-line1", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 1f),
                (Name: "paragraph-exact1-line2", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 2f),
                (Name: "paragraph-exact1-line4", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 4f),
                (Name: "paragraph-exact18", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 18f, Distance: 0f, ExactLine: 0f),
                (Name: "page-exact1", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 0f),
                (Name: "page-exact1-dist-2", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: -2f, ExactLine: 0f),
                (Name: "page-exact1-dist-4", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: -4f, ExactLine: 0f),
                (Name: "page-exact1-dist-8", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: -8f, ExactLine: 0f),
                (Name: "page-exact1-line1", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 1f),
                (Name: "page-exact1-line2", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 2f),
                (Name: "page-exact1-line4", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 1f, Distance: 0f, ExactLine: 4f),
                (Name: "page-exact18", Relative: Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage, Wrap: false, Rule: Word.WdFrameSizeRule.wdFrameExact, Height: 18f, Distance: 0f, ExactLine: 0f),
            };
            var bestResidual = double.MaxValue;
            var bestRelative = frame.RelativeVerticalPosition;
            var bestWrap = frame.TextWrap;
            var bestRule = frame.HeightRule;
            var bestHeight = frame.Height;
            var bestDistance = frame.VerticalDistanceFromText;
            var bestExactLine = 0f;
            var bestPosition = frame.VerticalPosition;
            const float CalibrationPoints = 10f;
            foreach (var candidate in candidates)
            {
                try
                {
                    frame.RelativeVerticalPosition = candidate.Relative;
                    frame.TextWrap = candidate.Wrap;
                    frame.HeightRule = candidate.Rule;
                    if (candidate.Rule != Word.WdFrameSizeRule.wdFrameAuto)
                        frame.Height = candidate.Height;
                    frame.VerticalDistanceFromText = candidate.Distance;
                    if (candidate.ExactLine > 0f)
                    {
                        frameParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                        frameParagraphFormat.LineSpacing = candidate.ExactLine;
                    }
                    else
                    {
                        frameParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    }
                    frame.VerticalPosition = 0f;
                    document.Repaginate();
                    Thread.Sleep(100);
                    window.ScrollIntoView(mathRange, ref scrollStart);
                    Thread.Sleep(60);
                    var formula0 = ReadWordRangePixelBox(
                        window,
                        mathRange,
                        context + " " + candidate.Name + " formula0");
                    var number0 = ReadWordRangePixelBox(
                        window,
                        fieldResult,
                        context + " " + candidate.Name + " number0");
                    var formulaCenter0 = formula0.Top + formula0.Height / 2.0;
                    var numberCenter0 = number0.Top + number0.Height / 2.0;
                    var delta0 = numberCenter0 - formulaCenter0;

                    frame.VerticalPosition = CalibrationPoints;
                    document.Repaginate();
                    Thread.Sleep(100);
                    window.ScrollIntoView(mathRange, ref scrollStart);
                    Thread.Sleep(60);
                    var formula10 = ReadWordRangePixelBox(
                        window,
                        mathRange,
                        context + " " + candidate.Name + " formula10");
                    var number10 = ReadWordRangePixelBox(
                        window,
                        fieldResult,
                        context + " " + candidate.Name + " number10");
                    var delta10 = (number10.Top + number10.Height / 2.0)
                        - (formula10.Top + formula10.Height / 2.0);
                    var slope = (delta10 - delta0) / CalibrationPoints;
                    if (Math.Abs(slope) < 0.05)
                    {
                        Console.WriteLine(
                            $"  {context}: frame candidate {candidate.Name} is immovable delta0={delta0:0.###}px delta10={delta10:0.###}px actualPos={frame.VerticalPosition:0.###}pt, heightRule={frame.HeightRule}, height={frame.Height:0.###}pt, numberBox={number0.Width}x{number0.Height}px.");
                        continue;
                    }

                    var targetPosition = (float)Math.Max(
                        -120.0,
                        Math.Min(120.0, -delta0 / slope));
                    frame.VerticalPosition = targetPosition;
                    document.Repaginate();
                    Thread.Sleep(120);
                    window.ScrollIntoView(mathRange, ref scrollStart);
                    Thread.Sleep(60);
                    var formulaFinal = ReadWordRangePixelBox(
                        window,
                        mathRange,
                        context + " " + candidate.Name + " formula-final");
                    var numberFinal = ReadWordRangePixelBox(
                        window,
                        fieldResult,
                        context + " " + candidate.Name + " number-final");
                    var finalDelta = (numberFinal.Top + numberFinal.Height / 2.0)
                        - (formulaFinal.Top + formulaFinal.Height / 2.0);
                    var residual = Math.Abs(finalDelta);
                    Console.WriteLine(
                        $"  {context}: frame candidate {candidate.Name} delta0={delta0:0.###}px delta10={delta10:0.###}px slope={slope:0.###}px/pt target={targetPosition:0.###}pt actual={frame.VerticalPosition:0.###}pt finalDelta={finalDelta:0.###}px heightRule={frame.HeightRule}, height={frame.Height:0.###}pt, numberBox={numberFinal.Width}x{numberFinal.Height}px, formulaPageY={Convert.ToSingle(mathRange.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)):0.###}pt numberPageY={Convert.ToSingle(fieldResult.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)):0.###}pt.");
                    if (residual < bestResidual)
                    {
                        bestResidual = residual;
                        bestRelative = candidate.Relative;
                        bestWrap = candidate.Wrap;
                        bestRule = frame.HeightRule;
                        bestHeight = frame.Height;
                        bestDistance = frame.VerticalDistanceFromText;
                        bestExactLine = candidate.ExactLine;
                        bestPosition = frame.VerticalPosition;
                    }
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"  {context}: frame candidate {candidate.Name} rejected: {error.GetType().Name}: {error.Message}");
                }
            }

            AssertTrue(bestResidual <= 30.0,
                context + $": legacy Frame paragraph avoidance remains too large even with a 1pt collision box; best residual={bestResidual:0.###}px.");
            frame.RelativeVerticalPosition = bestRelative;
            frame.TextWrap = bestWrap;
            frame.HeightRule = bestRule;
            if (bestRule != Word.WdFrameSizeRule.wdFrameAuto)
                frame.Height = bestHeight;
            frame.VerticalDistanceFromText = bestDistance;
            if (bestExactLine > 0f)
            {
                frameParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                frameParagraphFormat.LineSpacing = bestExactLine;
            }
            else
            {
                frameParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
            }
            frame.VerticalPosition = bestPosition;
            document.Repaginate();
            Thread.Sleep(120);
            window.ScrollIntoView(mathRange, ref scrollStart);
            Thread.Sleep(60);
            var finalFormulaBox = ReadWordRangePixelBox(
                window,
                mathRange,
                context + " selected frame formula");
            var finalNumberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " selected frame number");
            var selectedDelta = (finalNumberBox.Top + finalNumberBox.Height / 2.0)
                - (finalFormulaBox.Top + finalFormulaBox.Height / 2.0);

            // A legacy Word Frame always reserves a paragraph-level vertical band,
            // even when TextWrap=false and the frame is only 1pt tall. Keep the
            // Frame geometry itself at Word's stable minimum, then test whether an
            // ordinary Word run baseline shift can close only the remaining visual
            // center gap without changing font family, size, TabStop or OMath.
            visibleFont = visibleRange.Font;
            resultFont = fieldResult.Font;
            leftParenthesisRange = document.Range(visibleRange.Start, visibleRange.Start + 1);
            rightParenthesisRange = document.Range(visibleRange.End - 1, visibleRange.End);
            leftParenthesisFont = leftParenthesisRange.Font;
            rightParenthesisFont = rightParenthesisRange.Font;
            visibleFont.Position = 0;
            resultFont.Position = 0;
            leftParenthesisFont.Position = 0;
            rightParenthesisFont.Position = 0;
            if (Math.Abs(selectedDelta) > 2.0)
            {
                const int PositionProbePoints = 4;
                resultFont.Position = PositionProbePoints;
                leftParenthesisFont.Position = PositionProbePoints;
                rightParenthesisFont.Position = PositionProbePoints;
                document.Repaginate();
                Thread.Sleep(100);
                window.ScrollIntoView(mathRange, ref scrollStart);
                Thread.Sleep(60);
                var raisedNumberBox = ReadWordRangePixelBox(
                    window,
                    fieldResult,
                    context + " raised-number baseline probe");
                var raisedDelta = (raisedNumberBox.Top + raisedNumberBox.Height / 2.0)
                    - (finalFormulaBox.Top + finalFormulaBox.Height / 2.0);
                var positionSlope = (raisedDelta - selectedDelta) / PositionProbePoints;
                AssertTrue(Math.Abs(positionSlope) >= 0.1,
                    context + ": ordinary Word Font.Position did not move the Frame number glyphs.");
                var targetPosition = (int)Math.Round(
                    -selectedDelta / positionSlope,
                    MidpointRounding.AwayFromZero);
                targetPosition = Math.Max(-40, Math.Min(40, targetPosition));
                visibleFont.Position = targetPosition;
                resultFont.Position = targetPosition;
                leftParenthesisFont.Position = targetPosition;
                rightParenthesisFont.Position = targetPosition;
                document.Repaginate();
                Thread.Sleep(120);
                window.ScrollIntoView(mathRange, ref scrollStart);
                Thread.Sleep(60);
                finalFormulaBox = ReadWordRangePixelBox(
                    window,
                    mathRange,
                    context + " baseline-compensated formula");
                finalNumberBox = ReadWordRangePixelBox(
                    window,
                    fieldResult,
                    context + " baseline-compensated number");
                selectedDelta = (finalNumberBox.Top + finalNumberBox.Height / 2.0)
                    - (finalFormulaBox.Top + finalFormulaBox.Height / 2.0);
                Console.WriteLine(
                    $"  {context}: ordinary number baseline compensation slope={positionSlope:0.###}px/pt, visiblePosition={visibleFont.Position}pt, resultPosition={resultFont.Position}pt, finalCenterDelta={selectedDelta:0.###}px.");
            }
            AssertTrue(Math.Abs(selectedDelta) <= 4.0,
                context + $": ordinary Word baseline compensation could not center the Frame number on the display formula; delta={selectedDelta:0.###}px.");
            Console.WriteLine(
                $"  {context}: selected frame mode relative={bestRelative}, wrap={bestWrap}, heightRule={bestRule}, height={frame.Height:0.###}pt, verticalPosition={bestPosition:0.###}pt, numberPosition={resultFont.Position}pt, centerDelta={selectedDelta:0.###}px.");
        }
        finally
        {
            Release(window);
            Release(rightParenthesisFont);
            Release(leftParenthesisFont);
            Release(resultFont);
            Release(visibleFont);
            Release(rightParenthesisRange);
            Release(leftParenthesisRange);
            Release(frameParagraphFormat);
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
        Word.Paragraphs? visibleParagraphs = null;
        Word.Paragraph? visibleParagraph = null;
        Word.Range? visibleParagraphRange = null;
        Word.ParagraphFormat? visibleParagraphFormat = null;
        Word.TabStops? visibleTabStops = null;
        Word.TabStop? visibleTabStop = null;
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
            AssertEqual(0, document.Shapes.Count,
                context + ": the true-display numbering prototype created a Shape/TextBox.");
            AssertEqual(0, document.InlineShapes.Count,
                context + ": the true-display numbering prototype created an InlineShape/OLE host.");
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

            visibleParagraphs = visibleRange.Paragraphs;
            AssertEqual(1, visibleParagraphs.Count,
                context + ": the visible REF is not in exactly one ordinary Word paragraph.");
            visibleParagraph = visibleParagraphs[1];
            visibleParagraphRange = visibleParagraph.Range;
            AssertTrue((visibleParagraphRange.Text ?? string.Empty).StartsWith("\t", StringComparison.Ordinal),
                context + ": the Frame paragraph does not begin with a genuine TAB character.");
            AssertTrue((visibleParagraphRange.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal),
                context + ": the Frame paragraph lost its final ordinary paragraph mark.");
            AssertEqual(visibleParagraphRange.Start + 1, visibleRange.Start,
                context + ": the visible number is not immediately after the right TAB.");
            AssertEqual(visibleParagraphRange.End - 1, visibleRange.End,
                context + ": the visible number is not immediately before the final paragraph mark.");
            visibleParagraphFormat = visibleParagraphRange.ParagraphFormat;
            visibleTabStops = visibleParagraphFormat.TabStops;
            var sawRightTab = false;
            var rightTabPosition = 0f;
            for (var index = 1; index <= visibleTabStops.Count; index++)
            {
                Release(visibleTabStop); visibleTabStop = visibleTabStops[index];
                if (visibleTabStop.Alignment != Word.WdTabAlignment.wdAlignTabRight) continue;
                sawRightTab = true;
                rightTabPosition = visibleTabStop.Position;
            }
            AssertTrue(sawRightTab,
                context + ": the Frame paragraph lost its genuine right TabStop.");

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
                context + ": the visible REF Frame is not anchored at the writable right edge.");
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
            var writableWidth = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            AssertNear(415.3f, writableWidth, 0.7f,
                context + ": prototype writable width is not the intended 415.3pt.");
            AssertNear(72f, visibleFrame.Width, 0.7f,
                context + ": the right-edge Frame width changed from 72pt.");
            AssertNear(visibleFrame.Width, rightTabPosition, 0.7f,
                context + ": the genuine right TabStop is not at the Frame's right edge.");
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
            Release(visibleTabStop);
            Release(visibleTabStops);
            Release(visibleParagraphFormat);
            Release(visibleParagraphRange);
            Release(visibleParagraph);
            Release(visibleParagraphs);
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
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>" + $"<w:bookmarkStart w:id=\"2\" w:name=\"{TrueDisplayFrameFormulaBookmark}\"/><m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>" + semanticOmml + "</m:oMathPara><w:bookmarkEnd w:id=\"2\"/></w:p>"
            + "<w:p><w:pPr><w:tabs><w:tab w:val=\"right\" w:pos=\"8306\"/></w:tabs><w:jc w:val=\"left\"/><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr><w:r><w:tab/></w:r>" + $"<w:bookmarkStart w:id=\"3\" w:name=\"{TrueDisplayFrameVisibleBookmark}\"/><w:r><w:t>(</w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> REF {TrueDisplayFrameTargetBookmark} \\h \\* CHARFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r><w:r><w:t>)</w:t></w:r><w:bookmarkEnd w:id=\"3\"/></w:p>"
            + "<w:p><w:pPr><w:spacing w:before=\"0\" w:after=\"0\"/></w:pPr>"
            + $"<w:bookmarkStart w:id=\"4\" w:name=\"{TrueDisplayFrameReferenceBookmark}\"/>"
            + "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>"
            + semanticOmml
            + "</m:oMathPara>"
            + "<w:bookmarkEnd w:id=\"4\"/>"
            + "</w:p>"
            + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>"
            + "<w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
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

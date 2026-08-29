using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string TrueDisplayProbeFormulaBookmark = "VTProbeFormula";
    private const string TrueDisplayProbeReferenceBookmark = "VTProbeReference";
    private const string TrueDisplayProbeTargetBookmark = "VTProbeTarget";
    private const string TrueDisplayProbeVisibleBookmark = "VTProbeVisible";

    private static void RunWordOmmlTrueDisplayTabPrototypeAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourcePath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-tab-prototype-source.docx");
        var roundTripPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-tab-prototype-roundtrip.docx");
        var strippedPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-tab-prototype-stripped.docx");
        var strippedRoundTripPath = Path.Combine(
            artifactRoot,
            "word-omml-true-display-tab-prototype-stripped-roundtrip.docx");

        var semanticOmml = WordOmmlConverter.ExtractSingleOMath(
            WordOmmlConverter.TransformMathMlToOmml(QuadraticFormulaMathMl()));
        WriteTrueDisplayTabProbeDocx(sourcePath, semanticOmml);

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

            AssertTrueDisplayTabProbe(
                application,
                document,
                semanticOmml,
                "initial OOXML import",
                promoteToDisplay: true,
                alignVisibleReferenceVertically: false,
                requireNumberSameVisualLine: false);
            document.Fields.Update();
            document.SaveAs2(
                roundTripPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            RewriteTrueDisplayProbeWithoutSeparators(
                roundTripPath,
                strippedPath,
                removeLeadingSeparator: false,
                removeTrailingSeparator: true);
            document = application.Documents.Open(
                strippedPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            document.Fields.Update();
            AssertTrueDisplayTabProbe(
                application,
                document,
                semanticOmml,
                "trailing-separator-free import",
                promoteToDisplay: false,
                alignVisibleReferenceVertically: false,
                requireNumberSameVisualLine: true);
            document.SaveAs2(
                strippedRoundTripPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                strippedRoundTripPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            document.Fields.Update();
            AssertTrueDisplayTabProbe(
                application,
                document,
                semanticOmml,
                "trailing-separator-free save/reopen",
                promoteToDisplay: false,
                alignVisibleReferenceVertically: false,
                requireNumberSameVisualLine: true);
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            Console.WriteLine(
                "Word true-display tab prototype acceptance passed: m:oMathPara retained a genuine wdOMathDisplay formula while ordinary TAB/REF runs remained outside m:oMath, table-free, centered/right-aligned, and stable after F9 plus save/reopen.");
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

    private static void AssertTrueDisplayTabProbe(
        Word.Application application,
        Word.Document document,
        string semanticOmml,
        string context,
        bool promoteToDisplay,
        bool alignVisibleReferenceVertically,
        bool requireNumberSameVisualLine)
    {
        Word.Bookmark? candidateBookmark = null;
        Word.Bookmark? referenceBookmark = null;
        Word.Range? candidateRange = null;
        Word.Range? referenceRange = null;
        Word.OMaths? candidateMaths = null;
        Word.OMaths? referenceMaths = null;
        Word.OMath? candidateMath = null;
        Word.OMath? referenceMath = null;
        Word.Range? candidateMathRange = null;
        Word.Range? referenceMathRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.Range? ownerRange = null;
        Word.Fields? ownerFields = null;
        Word.Field? visibleReference = null;
        Word.Range? fieldCode = null;
        Word.Range? fieldResult = null;
        Word.Bookmark? visibleLabelBookmark = null;
        Word.Range? visibleLabelRange = null;
        Microsoft.Office.Interop.Word.Font? visibleLabelFont = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Word.Window? window = null;
        Word.Range? numberEnd = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        try
        {
            AssertEqual(0, document.Tables.Count,
                context + ": the true-display candidate unexpectedly contains a table.");
            AssertTrue(document.Bookmarks.Exists(TrueDisplayProbeFormulaBookmark),
                context + ": the candidate formula bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(TrueDisplayProbeReferenceBookmark),
                context + ": the reference display formula bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(TrueDisplayProbeTargetBookmark),
                context + ": the hidden REF target bookmark is missing.");

            candidateBookmark = document.Bookmarks[TrueDisplayProbeFormulaBookmark];
            referenceBookmark = document.Bookmarks[TrueDisplayProbeReferenceBookmark];
            candidateRange = candidateBookmark.Range;
            referenceRange = referenceBookmark.Range;
            candidateMaths = candidateRange.OMaths;
            referenceMaths = referenceRange.OMaths;
            AssertEqual(1, candidateMaths.Count,
                context + ": the candidate bookmark does not contain exactly one OMath.");
            AssertEqual(1, referenceMaths.Count,
                context + ": the reference bookmark does not contain exactly one OMath.");
            candidateMath = candidateMaths[1];
            referenceMath = referenceMaths[1];
            candidateMathRange = candidateMath.Range;
            referenceMathRange = referenceMath.Range;
            Console.WriteLine(
                $"  {context}: imported candidate type={candidateMath.Type}, range={candidateMathRange.Start}:{candidateMathRange.End}, text='{(candidateMathRange.Text ?? string.Empty).Replace("\r", "\\r").Replace("\v", "\\v").Replace("\t", "\\t")}'.");
            if (candidateMath.Type != Word.WdOMathType.wdOMathDisplay
                && promoteToDisplay)
            {
                candidateMath.Type = Word.WdOMathType.wdOMathDisplay;
                Release(candidateMathRange);
                candidateMathRange = candidateMath.Range;
                Console.WriteLine(
                    $"  {context}: after explicit wdOMathDisplay type={candidateMath.Type}, range={candidateMathRange.Start}:{candidateMathRange.End}, text='{(candidateMathRange.Text ?? string.Empty).Replace("\r", "\\r").Replace("\v", "\\v").Replace("\t", "\\t")}'.");
            }
            var displayDiagnosticPath = Path.Combine(
                document.Path,
                "word-omml-true-display-before-separator-cleanup.docx");
            document.SaveAs2(
                displayDiagnosticPath,
                Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine(
                $"  {context}: saved genuine-display separator diagnostic to '{displayDiagnosticPath}'.");
            Console.WriteLine(
                $"  {context}: finalType={candidateMath.Type}, finalRange={candidateMathRange.Start}:{candidateMathRange.End}.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, candidateMath.Type,
                context + ": Word rejected the explicit genuine-display conversion.");
            AssertEqual(Word.WdOMathType.wdOMathDisplay, referenceMath.Type,
                context + ": the plain Word display reference is not display math.");
            AssertEqual(0, candidateMathRange.Fields.Count,
                context + ": the visible REF field leaked into m:oMath.");

            paragraphs = candidateMathRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                context + ": the numbered candidate spans more than one Word paragraph.");
            paragraph = paragraphs[1];
            ownerRange = paragraph.Range;
            AssertTrue(!(bool)ownerRange.get_Information(Word.WdInformation.wdWithInTable),
                context + ": the numbered candidate is hosted in a table.");
            var ownerText = ownerRange.Text ?? string.Empty;
            AssertTrue(ownerText.Count(character => character == '\t') >= 2,
                context + ": Word did not preserve both ordinary layout TAB characters.");

            const string WordNamespace =
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var ownerXml = XDocument.Parse(
                ownerRange.WordOpenXML ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            var word = (XNamespace)WordNamespace;
            var math = (XNamespace)MathNamespace;
            var mathParagraphs = ownerXml.Descendants(math + "oMathPara").ToArray();
            AssertEqual(1, mathParagraphs.Length,
                context + ": Word did not preserve exactly one m:oMathPara host.");
            var mathParagraph = mathParagraphs[0];
            var formulaNodes = mathParagraph.Descendants(math + "oMath").ToArray();
            AssertEqual(1, formulaNodes.Length,
                context + ": the m:oMathPara host does not contain exactly one formula.");
            var formulaXml = formulaNodes[0].ToString(SaveOptions.DisableFormatting);
            AssertTrue(formulaXml.IndexOf("REF " + TrueDisplayProbeTargetBookmark,
                    StringComparison.OrdinalIgnoreCase) < 0,
                context + ": the visible REF instruction is inside m:oMath.");
            AssertTrue(!formulaNodes[0].Descendants(word + "fldChar").Any(),
                context + ": field control runs leaked inside m:oMath.");
            AssertTrue(ownerXml.Descendants(word + "instrText").Any(text =>
                    text.Value.IndexOf(
                        "REF " + TrueDisplayProbeTargetBookmark,
                        StringComparison.OrdinalIgnoreCase) >= 0),
                context + ": the ordinary REF field is not retained outside m:oMath in the same Word paragraph.");
            AssertTrue(!ownerXml.Descendants(math + "eqArr").Any(),
                context + ": the candidate contains the obsolete m:eqArr numbering wrapper.");

            ownerFields = ownerRange.Fields;
            for (var index = 1; index <= ownerFields.Count; index++)
            {
                Word.Field? candidate = null;
                Word.Range? code = null;
                try
                {
                    candidate = ownerFields[index];
                    code = candidate.Code;
                    if ((code.Text ?? string.Empty).IndexOf(
                            "REF " + TrueDisplayProbeTargetBookmark,
                            StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                    visibleReference = candidate;
                    candidate = null;
                    break;
                }
                finally
                {
                    Release(code);
                    Release(candidate);
                }
            }
            if (visibleReference is null)
                throw new InvalidDataException(
                    context + ": Word COM cannot resolve the visible ordinary REF field.");
            visibleReference.Update();
            fieldCode = visibleReference.Code;
            fieldResult = visibleReference.Result;
            var visibleNumber = (fieldResult.Text ?? string.Empty).Trim();
            AssertEqual("1", visibleNumber,
                context + ": the visible REF result is not the live hidden SEQ value.");

            format = paragraph.Format;
            tabStops = format.TabStops;
            var sawCenter = false;
            var sawRight = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop); tabStop = tabStops[index];
                sawCenter |= tabStop.Alignment == Word.WdTabAlignment.wdAlignTabCenter;
                sawRight |= tabStop.Alignment == Word.WdTabAlignment.wdAlignTabRight;
            }
            AssertTrue(sawCenter && sawRight,
                context + ": the Word paragraph lost its center/right tab stops.");

            window = document.ActiveWindow;
            if (alignVisibleReferenceVertically)
            {
                visibleLabelBookmark = document.Bookmarks[TrueDisplayProbeVisibleBookmark];
                visibleLabelRange = visibleLabelBookmark.Range;
                var labelPosition = AlignTrueDisplayReferenceVertically(
                    document,
                    window,
                    candidateMathRange,
                    fieldResult,
                    context);
                ApplyTrueDisplayParenthesisPosition(
                    document,
                    visibleLabelRange,
                    labelPosition,
                    context);
                visibleLabelFont = fieldResult.Font;
                Console.WriteLine(
                    $"  {context}: persisted visible label position={visibleLabelFont.Position}pt.");
            }
            object scrollStart = true;
            window.ScrollIntoView(ownerRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(100);
            var candidateBox = ReadWordRangePixelBox(
                window,
                candidateMathRange,
                context + " candidate display formula");
            var referenceBox = ReadWordRangePixelBox(
                window,
                referenceMathRange,
                context + " plain display reference");
            AssertTrue(Math.Abs(candidateBox.Width - referenceBox.Width) <= 3,
                context + $": candidate/reference display widths differ ({candidateBox.Width} vs {referenceBox.Width}).");
            AssertTrue(Math.Abs(candidateBox.Height - referenceBox.Height) <= 3,
                context + $": candidate/reference display heights differ ({candidateBox.Height} vs {referenceBox.Height}).");
            var candidateCenter = candidateBox.Left + candidateBox.Width / 2.0;
            var referenceCenter = referenceBox.Left + referenceBox.Width / 2.0;
            AssertTrue(Math.Abs(candidateCenter - referenceCenter) <= 4.0,
                context + $": center TAB did not place the candidate at the normal display center ({candidateCenter:0.##} vs {referenceCenter:0.##}).");
            var numberBox = ReadWordRangePixelBox(
                window,
                fieldResult,
                context + " visible REF result");
            Console.WriteLine(
                $"  {context}: candidateBox={candidateBox.Left},{candidateBox.Top},{candidateBox.Width},{candidateBox.Height}, numberBox={numberBox.Left},{numberBox.Top},{numberBox.Width},{numberBox.Height}, nextDisplayTop={referenceBox.Top}, candidateToNextDisplay={referenceBox.Top - candidateBox.Top}.");
            var formulaCenterY = candidateBox.Top + candidateBox.Height / 2.0;
            var numberCenterY = numberBox.Top + numberBox.Height / 2.0;
            var verticalCenterDelta = numberCenterY - formulaCenterY;
            Console.WriteLine(
                $"  {context}: unadjusted formula/number vertical centers={formulaCenterY:0.##}/{numberCenterY:0.##}, delta={verticalCenterDelta:0.##}px.");
            if (requireNumberSameVisualLine)
            {
                AssertTrue(Math.Abs(verticalCenterDelta) <= 3.0,
                    context + $": removing the trailing display separator did not keep the external REF on the formula's visual line (delta={verticalCenterDelta:0.##}px).");
            }

            sections = ownerRange.Sections;
            AssertTrue(sections.Count > 0,
                context + ": the numbered candidate has no Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            var expectedRight = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            numberEnd = document.Range(fieldResult.End, fieldResult.End);
            var actualRight = Convert.ToSingle(numberEnd.get_Information(
                Word.WdInformation.wdHorizontalPositionRelativeToTextBoundary));
            AssertTrue(actualRight >= 0f,
                context + ": Word could not report the REF result's right edge.");
            if (requireNumberSameVisualLine)
            {
                AssertNear(expectedRight, actualRight, 6f,
                    context + ": right TAB did not align the visible REF to the writable right edge.");
            } // pre-cleanup display separators intentionally defer final tab geometry 

            var candidateMetric = ReadWordOmmlLayoutMetric(
                application,
                document,
                candidateMathRange,
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
                context + ": candidate and plain display reference report different OMath types.");
            AssertTrue(Math.Abs(candidateMetric.WidthPx - referenceMetric.WidthPx) <= 3,
                context + ": candidate and reference visible math ink widths diverged.");
            AssertTrue(Math.Abs(candidateMetric.HeightPx - referenceMetric.HeightPx) <= 3,
                context + ": candidate and reference visible math ink heights diverged.");

            Console.WriteLine(
                $"  {context}: candidateType={candidateMath.Type}, candidateBox={candidateBox.Left},{candidateBox.Top},{candidateBox.Width},{candidateBox.Height}, referenceBox={referenceBox.Left},{referenceBox.Top},{referenceBox.Width},{referenceBox.Height}, number='{visibleNumber}', numberRight={actualRight:0.##}/{expectedRight:0.##}.");
        }
        finally
        {
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(numberEnd);
            Release(window);
            Release(tabStop);
            Release(tabStops);
            Release(format);
            Release(visibleLabelFont);
            Release(visibleLabelRange);
            Release(visibleLabelBookmark);
            Release(fieldResult);
            Release(fieldCode);
            Release(visibleReference);
            Release(ownerFields);
            Release(ownerRange);
            Release(paragraph);
            Release(paragraphs);
            Release(referenceMathRange);
            Release(candidateMathRange);
            Release(referenceMath);
            Release(candidateMath);
            Release(referenceMaths);
            Release(candidateMaths);
            Release(referenceRange);
            Release(candidateRange);
            Release(referenceBookmark);
            Release(candidateBookmark);
        }
    }

    private static void RewriteTrueDisplayProbeWithoutSeparators(
        string sourcePath,
        string targetPath,
        bool removeLeadingSeparator,
        bool removeTrailingSeparator)
    {
        if (File.Exists(targetPath)) File.Delete(targetPath);
        File.Copy(sourcePath, targetPath);
        using var archive = ZipFile.Open(targetPath, ZipArchiveMode.Update);
        var documentEntry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException(
                "The true-display probe DOCX has no word/document.xml part.");
        string documentXml;
        using (var stream = documentEntry.Open())
        using (var reader = new StreamReader(stream, Encoding.UTF8, true, 4096, leaveOpen: false))
            documentXml = reader.ReadToEnd();

        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var word = (XNamespace)WordNamespace;
        var math = (XNamespace)MathNamespace;
        var xml = XDocument.Parse(documentXml, LoadOptions.PreserveWhitespace);
        var bookmarkStart = xml
            .Descendants(word + "bookmarkStart")
            .SingleOrDefault(element => string.Equals(
                (string?)element.Attribute(word + "name"),
                TrueDisplayProbeFormulaBookmark,
                StringComparison.Ordinal))
            ?? throw new InvalidDataException(
                "The saved true-display probe lost its formula bookmark.");
        var paragraph = bookmarkStart.Ancestors(word + "p").FirstOrDefault()
            ?? throw new InvalidDataException(
                "The saved true-display probe formula is not inside a Word paragraph.");
        var mathParagraph = paragraph.Descendants(math + "oMathPara").SingleOrDefault()
            ?? throw new InvalidDataException(
                "The saved true-display probe paragraph does not contain exactly one m:oMathPara.");
        var equation = mathParagraph.Elements(math + "oMath").SingleOrDefault()
            ?? throw new InvalidDataException(
                "The saved true-display probe paragraph does not contain exactly one OMath.");
        var leadingRun = mathParagraph.ElementsBeforeSelf()
            .Where(element => element.Name == word + "r")
            .LastOrDefault(element => element.Descendants(word + "br").Any())
            ?? throw new InvalidDataException(
                "Word did not persist the expected leading display separator run.");
        var trailingRun = equation.Elements().LastOrDefault();
        if (trailingRun?.Name != math + "r"
            || !trailingRun.Descendants(word + "br").Any())
            throw new InvalidDataException(
                "Word did not persist the expected trailing display separator run.");

        if (removeLeadingSeparator) leadingRun.Remove();
        if (removeTrailingSeparator) trailingRun.Remove();

        documentEntry.Delete();
        var replacement = archive.CreateEntry(
            "word/document.xml",
            CompressionLevel.Optimal);
        using (var output = replacement.Open())
        using (var writer = new StreamWriter(
                   output,
                   new UTF8Encoding(false),
                   4096,
                   leaveOpen: false))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            writer.Write(xml.Root!.ToString(SaveOptions.DisableFormatting));
        }
        Console.WriteLine(
            $"Rewrote true-display probe without separators leading={removeLeadingSeparator}, trailing={removeTrailingSeparator}: {targetPath}");
    }

    private static void RemoveTrailingTrueDisplaySeparatorByBookmark(
        Word.Document document,
        string context)
    {
        Word.Bookmark? bookmark = null;
        Word.Range? bookmarkRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        try
        {
            bookmark = document.Bookmarks[TrueDisplayProbeFormulaBookmark];
            bookmarkRange = bookmark.Range;
            maths = bookmarkRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": the formula bookmark no longer contains exactly one OMath.");
            math = maths[1];
            mathRange = math.Range;
            AssertTrue(
                TryRemoveTrailingTrueDisplaySeparator(document, mathRange, context),
                context + ": Word did not expose its trailing display separator for removal.");
        }
        finally
        {
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(bookmarkRange);
            Release(bookmark);
        }
    }

    private static bool TryRemoveTrailingTrueDisplaySeparator(
        Word.Document document,
        Word.Range formulaRange,
        string context)
    {
        Word.Range? trailing = null;
        try
        {
            if (formulaRange.End <= formulaRange.Start) return false;
            trailing = document.Range(formulaRange.End - 1, formulaRange.End);
            var codePoint = string.IsNullOrEmpty(trailing.Text)
                ? -1
                : char.ConvertToUtf32(trailing.Text, 0);
            Console.WriteLine(
                $"  {context}: trailing OMath character at {trailing.Start}:{trailing.End} is U+{codePoint:X4}.");
            if (codePoint != '\r' && codePoint != '\v') return false;
            trailing.Delete();
            Console.WriteLine(
                $"  {context}: deleted the trailing Word display separator.");
            return true;
        }
        finally { Release(trailing); }
    }

    private static int AlignTrueDisplayReferenceVertically(
        Word.Document document,
        Word.Window window,
        Word.Range formulaRange,
        Word.Range fieldResultRange,
        string context)
    {
        _ = document;
        Microsoft.Office.Interop.Word.Font? resultFont = null;
        try
        {
            resultFont = fieldResultRange.Font;
            resultFont.Position = 0;
            object scrollStart = true;
            window.ScrollIntoView(fieldResultRange, ref scrollStart);
            Thread.Sleep(80);
            var baselineLabelBox = ReadWordRangePixelBox(
                window,
                fieldResultRange,
                context + " vertical label baseline calibration");
            const int CalibrationPoints = 10;
            resultFont.Position = CalibrationPoints;
            window.ScrollIntoView(fieldResultRange, ref scrollStart);
            Thread.Sleep(80);
            var raisedCalibrationBox = ReadWordRangePixelBox(
                window,
                fieldResultRange,
                context + " vertical label raised calibration");
            var measuredPixelsPerPoint =
                (baselineLabelBox.Top - raisedCalibrationBox.Top) / (double)CalibrationPoints;
            var pixelsPerPoint = Math.Max(0.25, measuredPixelsPerPoint);
            var appliedPosition = 0;
            resultFont.Position = appliedPosition;
            Console.WriteLine(
                $"    true-display REF calibration: baselineTop={baselineLabelBox.Top}, raisedTop={raisedCalibrationBox.Top}, pixelsPerPoint={measuredPixelsPerPoint:0.###}.");
            for (var iteration = 0; iteration < 3; iteration++)
            {
                window.ScrollIntoView(formulaRange, ref scrollStart);
                Thread.Sleep(80);
                var formulaBox = ReadWordRangePixelBox(
                    window,
                    formulaRange,
                    context + " vertical formula probe");
                var labelBox = ReadWordRangePixelBox(
                    window,
                    fieldResultRange,
                    context + " vertical label probe");
                var desiredTop = formulaBox.Top
                    + (formulaBox.Height - labelBox.Height) / 2.0;
                var deltaPixels = labelBox.Top - desiredTop;
                if (Math.Abs(deltaPixels) <= 1.0) break;
                var deltaPoints = (int)Math.Round(
                    deltaPixels / pixelsPerPoint,
                    MidpointRounding.AwayFromZero);
                appliedPosition = Math.Max(
                    -150,
                    Math.Min(150, appliedPosition + deltaPoints));
                resultFont.Position = appliedPosition;
            }
            window.ScrollIntoView(fieldResultRange, ref scrollStart);
            Thread.Sleep(80);
            Console.WriteLine(
                $"    true-display visible REF vertical position={appliedPosition}pt.");
            return appliedPosition;
        }
        finally { Release(resultFont); }
    }

    private static void ApplyTrueDisplayParenthesisPosition(
        Word.Document document,
        Word.Range visibleLabelRange,
        int position,
        string context)
    {
        Word.Range? leftParenthesis = null;
        Word.Range? rightParenthesis = null;
        Microsoft.Office.Interop.Word.Font? leftFont = null;
        Microsoft.Office.Interop.Word.Font? rightFont = null;
        try
        {
            if (visibleLabelRange.End - visibleLabelRange.Start < 2)
                throw new InvalidDataException(
                    context + ": the visible equation label range is too short.");
            leftParenthesis = document.Range(
                visibleLabelRange.Start,
                visibleLabelRange.Start + 1);
            rightParenthesis = document.Range(
                visibleLabelRange.End - 1,
                visibleLabelRange.End);
            if (!string.Equals(leftParenthesis.Text, "(", StringComparison.Ordinal)
                || !string.Equals(rightParenthesis.Text, ")", StringComparison.Ordinal))
                throw new InvalidDataException(
                    context + ": the visible equation label parentheses are not ordinary Word characters.");
            leftFont = leftParenthesis.Font;
            rightFont = rightParenthesis.Font;
            leftFont.Position = position;
            rightFont.Position = position;
        }
        finally
        {
            Release(rightFont);
            Release(leftFont);
            Release(rightParenthesis);
            Release(leftParenthesis);
        }
    }

    private static void WriteTrueDisplayTabProbeDocx(
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

        WriteTrueDisplayProbeEntry(
            archive,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Types xmlns=\"{ContentTypesNamespace}\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
            + "<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>"
            + "</Types>");
        WriteTrueDisplayProbeEntry(
            archive,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Relationships xmlns=\"{RelationshipsNamespace}\">"
            + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
            + "</Relationships>");
        WriteTrueDisplayProbeEntry(
            archive,
            "word/_rels/document.xml.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<Relationships xmlns=\"{RelationshipsNamespace}\">"
            + "<Relationship Id=\"rIdSettings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/>"
            + "</Relationships>");
        WriteTrueDisplayProbeEntry(
            archive,
            "word/settings.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:settings xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + "<w:compat><w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"12\"/></w:compat>"
            + $"<m:mathPr><m:mathFont m:val=\"{WordOfficeMathFontLoader.LatinModernMathFamily}\"/></m:mathPr>"
            + "</w:settings>");

        var documentXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            + $"<w:document xmlns:w=\"{WordNamespace}\" xmlns:m=\"{MathNamespace}\">"
            + "<w:body>"
            + "<w:p><w:pPr><w:rPr><w:vanish/></w:rPr></w:pPr>"
            + $"<w:bookmarkStart w:id=\"1\" w:name=\"{TrueDisplayProbeTargetBookmark}\"/>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:instrText xml:space=\"preserve\"> SEQ VTProbe \\* ARABIC </w:instrText></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:t>1</w:t></w:r>"
            + "<w:r><w:rPr><w:vanish/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:bookmarkEnd w:id=\"1\"/>"
            + "</w:p>"
            + "<w:p><w:pPr>"
            + "<w:tabs><w:tab w:val=\"center\" w:pos=\"4680\"/><w:tab w:val=\"right\" w:pos=\"9360\"/></w:tabs>"
            + "<w:jc w:val=\"both\"/><w:spacing w:before=\"0\" w:after=\"0\"/>"
            + "</w:pPr>"
            + "<w:r><w:tab/></w:r>"
            + "<m:oMathPara><m:oMathParaPr><m:jc m:val=\"left\"/></m:oMathParaPr>"
            + $"<w:bookmarkStart w:id=\"2\" w:name=\"{TrueDisplayProbeFormulaBookmark}\"/>"
            + semanticOmml
            + "<w:bookmarkEnd w:id=\"2\"/>"
            + "</m:oMathPara>"
            + "<w:r><w:tab/></w:r>"
            + "<w:r><w:tab/></w:r>"
            + $"<w:bookmarkStart w:id=\"4\" w:name=\"{TrueDisplayProbeVisibleBookmark}\"/>" // two-tab display-number host 
            + "<w:r><w:t>(</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\" w:dirty=\"true\"/></w:r>"
            + $"<w:r><w:instrText xml:space=\"preserve\"> REF {TrueDisplayProbeTargetBookmark} \\h \\* CHARFORMAT </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:t>1</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:r><w:t>)</w:t></w:r>"
            + "<w:bookmarkEnd w:id=\"4\"/>"
            + "</w:p>"
            + "<w:p><m:oMathPara>" // Devspace edit boundary 
            + $"<w:bookmarkStart w:id=\"3\" w:name=\"{TrueDisplayProbeReferenceBookmark}\"/>"
            + semanticOmml
            + "<w:bookmarkEnd w:id=\"3\"/>"
            + "</m:oMathPara></w:p>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
            + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>"
            + "</w:sectPr></w:body></w:document>";
        WriteTrueDisplayProbeEntry(archive, "word/document.xml", documentXml);
    }

    private static void WriteTrueDisplayProbeEntry(
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

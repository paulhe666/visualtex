using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlStyleSeparatorFieldProbeAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator field probe refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            RunStyleSeparatorFieldProbeScenario(
                application,
                artifactRoot,
                "display-first-seq",
                useLeadingCenterAnchor: false);
            RunStyleSeparatorFieldProbeScenario(
                application,
                artifactRoot,
                "center-anchor-display-seq",
                useLeadingCenterAnchor: true);
            Console.WriteLine(
                "Word OMML Style Separator field probe completed with a genuine SEQ field through update, save and close/reopen.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void RunStyleSeparatorFieldProbeScenario(
        Word.Application application,
        string artifactRoot,
        string name,
        bool useLeadingCenterAnchor)
    {
        var docxPath = Path.Combine(artifactRoot, name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, name + ".pdf");
        var xmlPath = Path.Combine(artifactRoot, name + "-after-reopen.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(xmlPath),
            name + ": field-probe artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
        Word.Range? emptyParentheses = null;
        Word.Range? fieldInsertion = null;
        Word.Field? sequenceField = null;
        Word.Range? fieldResult = null;
        Word.Range? numberParagraphRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? centerTab = null;
        Word.TabStop? rightTab = null;
        Word.Selection? selection = null;
        Word.Range? separatorPoint = null;
        Word.Range? formulaRange = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Fields? contentFields = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? setup = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            content = document.Content;
            content.Text = (useLeadingCenterAnchor ? "\t\r" : string.Empty)
                + StyleSeparatorFormulaText
                + "\r\t()";

            sections = document.Sections;
            section = sections[1];
            setup = section.PageSetup;
            setup.PageWidth = 595.3f;
            setup.PageHeight = 841.9f;
            setup.LeftMargin = 90f;
            setup.RightMargin = 90f;
            setup.TopMargin = 72f;
            setup.BottomMargin = 72f;

            emptyParentheses = FindUniqueStyleSeparatorTextRange(
                document,
                "()",
                name + " empty number wrapper");
            fieldInsertion = document.Range(
                emptyParentheses.Start + 1,
                emptyParentheses.Start + 1);
            sequenceField = document.Fields.Add(
                fieldInsertion,
                Word.WdFieldType.wdFieldEmpty,
                "SEQ VisualTeXEquation \\r 1 \\* ARABIC",
                PreserveFormatting: true);
            sequenceField.Update();
            fieldResult = sequenceField.Result;
            AssertEqual("1", fieldResult.Text ?? string.Empty,
                name + ": SEQ field did not evaluate to 1.");

            paragraphs = document.Paragraphs;
            var expectedParagraphs = useLeadingCenterAnchor ? 3 : 2;
            AssertEqual(expectedParagraphs, paragraphs.Count,
                name + ": field insertion changed the logical paragraph count.");
            paragraph = paragraphs[expectedParagraphs];
            numberParagraphRange = paragraph.Range.Duplicate;
            numberParagraphRange.MoveStart(Word.WdUnits.wdCharacter, 1);
            numberParagraphRange.MoveEnd(Word.WdUnits.wdCharacter, -1);
            numberFont = numberParagraphRange.Font;
            numberFont.Name = StyleSeparatorNumberFont;
            numberFont.NameAscii = StyleSeparatorNumberFont;
            numberFont.NameOther = StyleSeparatorNumberFont;
            numberFont.Size = StyleSeparatorNumberFontSize;
            numberFont.Bold = 0;
            numberFont.Italic = 0;
            numberFont.Position = 0;
            numberFont.Hidden = 0;

            for (var index = 1; index <= expectedParagraphs; index++)
            {
                Release(rightTab); rightTab = null;
                Release(centerTab); centerTab = null;
                Release(tabs); tabs = null;
                Release(format); format = null;
                Release(paragraph); paragraph = null;
                paragraph = paragraphs[index];
                format = paragraph.Format;
                ConfigureStyleSeparatorMatrixParagraph(
                    format,
                    Word.WdParagraphAlignment.wdAlignParagraphLeft);
                tabs = format.TabStops;
                tabs.ClearAll();
                if (useLeadingCenterAnchor)
                {
                    centerTab = tabs.Add(
                        StyleSeparatorRightTabPosition / 2f,
                        Word.WdTabAlignment.wdAlignTabCenter,
                        Word.WdTabLeader.wdTabLeaderSpaces);
                }
                rightTab = tabs.Add(
                    StyleSeparatorRightTabPosition,
                    Word.WdTabAlignment.wdAlignTabRight,
                    Word.WdTabLeader.wdTabLeaderSpaces);
            }

            selection = application.Selection;
            for (var index = 1; index < expectedParagraphs; index++)
            {
                Release(separatorPoint); separatorPoint = null;
                Release(paragraph); paragraph = null;
                Release(paragraphs); paragraphs = null;
                paragraphs = document.Paragraphs;
                paragraph = paragraphs[index];
                separatorPoint = document.Range(
                    paragraph.Range.End - 1,
                    paragraph.Range.End - 1);
                selection.SetRange(separatorPoint.Start, separatorPoint.End);
                selection.InsertStyleSeparator();
            }

            formulaRange = FindUniqueStyleSeparatorTextRange(
                document,
                StyleSeparatorFormulaText,
                name + " formula");
            addedMathRange = document.OMaths.Add(formulaRange);
            addedMaths = addedMathRange.OMaths;
            AssertEqual(1, addedMaths.Count,
                name + ": Word did not create one OMath after Style Separator insertion.");
            math = addedMaths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            math.Justification = Word.WdOMathJc.wdOMathJcCenterGroup;
            mathRange = math.Range;
            var mathFont = mathRange.Font;
            try { mathFont.Size = 14f; }
            finally { Release(mathFont); }

            contentFields = document.Content.Fields;
            _ = contentFields.Update();
            AssertEqual(1, document.Fields.Count,
                name + ": document no longer contains exactly one SEQ field.");
            document.Repaginate();
            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                docxPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            contentFields = document.Content.Fields;
            _ = contentFields.Update();
            document.Repaginate();

            Release(mathRange); mathRange = null;
            Release(math); math = null;
            var reopenedMaths = document.OMaths;
            try
            {
                AssertEqual(1, reopenedMaths.Count,
                    name + ": reopened document no longer contains one OMath.");
                math = reopenedMaths[1];
                mathRange = math.Range;
            }
            finally { Release(reopenedMaths); }

            var openXml = document.Content.WordOpenXML ?? string.Empty;
            File.WriteAllText(xmlPath, openXml, new UTF8Encoding(false));
            var parsed = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace officeMath = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var mathNodes = parsed.Descendants(officeMath + "oMath").ToArray();
            var mathParagraphs = parsed.Descendants(officeMath + "oMathPara").ToArray();
            var mathOnly = mathNodes.Length == 1
                && string.Equals(
                    string.Concat(mathNodes[0]
                        .Descendants(officeMath + "t")
                        .Select(text => text.Value)),
                    StyleSeparatorFormulaText,
                    StringComparison.Ordinal)
                && !mathNodes[0].Descendants(word + "fldChar").Any()
                && !mathNodes[0].Descendants(word + "instrText").Any()
                && !mathNodes[0].Descendants(word + "tab").Any();
            var fieldsOutsideMath = parsed.Descendants(word + "fldChar").Any()
                && !mathNodes.Any(node => node.Descendants(word + "fldChar").Any());

            ExportStyleSeparatorFixedLayout(document, xpsPath, name);
            ExportStyleSeparatorPdfLayout(document, pdfPath, name);
            var glyphs = ReadStyleSeparatorFixedGlyphs(xpsPath, name);
            var equalsGlyphs = glyphs
                .Where(glyph => glyph.Text.IndexOf('=') >= 0)
                .ToArray();
            var numberGlyph = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('(') >= 0);
            var rightParenthesis = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf(')') >= 0);
            var numberRendered = numberGlyph is not null
                && rightParenthesis is not null;
            var formulaGlyph = equalsGlyphs.FirstOrDefault();
            var sameLine = numberRendered
                && formulaGlyph is not null
                && Math.Abs(formulaGlyph.OriginY - numberGlyph!.OriginY) <= 1d;
            var rightAligned = numberRendered && numberGlyph!.OriginX >= 450d;
            var candidate = math.Type == Word.WdOMathType.wdOMathDisplay
                && mathParagraphs.Length == 1
                && mathOnly
                && fieldsOutsideMath
                && numberRendered
                && sameLine
                && rightAligned
                && equalsGlyphs.Length == 1;
            var glyphSummary = string.Join(
                " | ",
                glyphs.Select(glyph =>
                    $"'{EscapeStyleSeparatorGlyphText(glyph.Text)}'@{glyph.OriginX:0.###},{glyph.OriginY:0.###}"));
            Console.WriteLine(
                $"  field-probe {name}: type={math.Type}, fields={document.Fields.Count}, mathOnly={mathOnly}, fieldOutside={fieldsOutsideMath}, numberRendered={numberRendered}, sameLine={sameLine}, right={rightAligned}, formulaCopies={equalsGlyphs.Length}, candidate={candidate}, glyphs={glyphSummary}.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(setup);
            Release(section);
            Release(sections);
            Release(contentFields);
            Release(mathRange);
            Release(math);
            Release(addedMaths);
            Release(addedMathRange);
            Release(formulaRange);
            Release(separatorPoint);
            Release(selection);
            Release(rightTab);
            Release(centerTab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(numberFont);
            Release(numberParagraphRange);
            Release(fieldResult);
            Release(sequenceField);
            Release(fieldInsertion);
            Release(emptyParentheses);
            Release(content);
            Release(document);
            ForceComCleanup();
        }
    }
}

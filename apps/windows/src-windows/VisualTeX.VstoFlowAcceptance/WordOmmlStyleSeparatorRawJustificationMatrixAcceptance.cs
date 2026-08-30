using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOmmlStyleSeparatorRawJustificationMatrixAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The raw Style Separator justification matrix refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        var variants = new string?[]
        {
            null,
            "centerGroup",
            "center",
            "left",
            "right",
            "inline",
        };

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var candidates = new List<string>();
            foreach (var variant in variants)
            {
                var name = "raw-jc-" + (variant ?? "absent");
                try
                {
                    if (RunStyleSeparatorRawJustificationScenario(
                            application,
                            artifactRoot,
                            name,
                            variant))
                        candidates.Add(name);
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"  raw-jc {name}: ERROR {error.GetType().Name}: {error.Message}");
                }
            }

            Console.WriteLine(
                "Style Separator raw-justification candidates: "
                + (candidates.Count == 0
                    ? "<none>"
                    : string.Join(", ", candidates)));
            Console.WriteLine(
                "Word OMML Style Separator raw-justification matrix completed after package mutation, close/reopen and fixed-layout export.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static bool RunStyleSeparatorRawJustificationScenario(
        Word.Application application,
        string artifactRoot,
        string name,
        string? rawJustification)
    {
        var docxPath = Path.Combine(artifactRoot, name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, name + ".pdf");
        var beforeXmlPath = Path.Combine(artifactRoot, name + "-before-patch.xml");
        var afterXmlPath = Path.Combine(artifactRoot, name + "-after-reopen.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(beforeXmlPath)
                   && !File.Exists(afterXmlPath),
            name + ": raw-justification artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
        Word.Range? formulaRange = null;
        Word.Range? numberRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? format = null;
        Word.TabStops? tabs = null;
        Word.TabStop? centerTab = null;
        Word.TabStop? rightTab = null;
        Word.Selection? selection = null;
        Word.Range? separatorPoint = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? setup = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            content = document.Content;
            content.Text = "\t\r"
                + StyleSeparatorFormulaText
                + "\r\t"
                + StyleSeparatorNumberText;

            sections = document.Sections;
            section = sections[1];
            setup = section.PageSetup;
            setup.PageWidth = 595.3f;
            setup.PageHeight = 841.9f;
            setup.LeftMargin = 90f;
            setup.RightMargin = 90f;
            setup.TopMargin = 72f;
            setup.BottomMargin = 72f;

            paragraphs = document.Paragraphs;
            AssertEqual(3, paragraphs.Count,
                name + ": Word did not create three logical paragraphs.");
            for (var index = 1; index <= paragraphs.Count; index++)
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
                centerTab = tabs.Add(
                    StyleSeparatorRightTabPosition / 2f,
                    Word.WdTabAlignment.wdAlignTabCenter,
                    Word.WdTabLeader.wdTabLeaderSpaces);
                rightTab = tabs.Add(
                    StyleSeparatorRightTabPosition,
                    Word.WdTabAlignment.wdAlignTabRight,
                    Word.WdTabLeader.wdTabLeaderSpaces);
            }

            numberRange = FindUniqueStyleSeparatorTextRange(
                document,
                StyleSeparatorNumberText,
                name + " number");
            numberFont = numberRange.Font;
            numberFont.Name = StyleSeparatorNumberFont;
            numberFont.NameAscii = StyleSeparatorNumberFont;
            numberFont.NameOther = StyleSeparatorNumberFont;
            numberFont.Size = StyleSeparatorNumberFontSize;
            numberFont.Bold = 0;
            numberFont.Italic = 0;
            numberFont.Position = 0;
            numberFont.Hidden = 0;

            selection = application.Selection;
            for (var index = 1; index <= 2; index++)
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
                name + ": Word did not create one OMath after Style Separators.");
            math = addedMaths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            math.Justification = Word.WdOMathJc.wdOMathJcCenterGroup;
            mathRange = math.Range;
            var mathFont = mathRange.Font;
            try { mathFont.Size = 14f; }
            finally { Release(mathFont); }

            document.Repaginate();
            var beforeXml = document.Content.WordOpenXML ?? string.Empty;
            File.WriteAllText(beforeXmlPath, beforeXml, new UTF8Encoding(false));
            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            PatchStyleSeparatorRawJustification(
                docxPath,
                rawJustification,
                name);

            document = application.Documents.Open(
                docxPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            document.Repaginate();

            Release(mathRange); mathRange = null;
            Release(math); math = null;
            var liveMaths = document.OMaths;
            try
            {
                AssertEqual(1, liveMaths.Count,
                    name + ": reopened raw-JC document no longer contains one OMath.");
                math = liveMaths[1];
                mathRange = math.Range;
            }
            finally { Release(liveMaths); }

            var afterXml = document.Content.WordOpenXML ?? string.Empty;
            File.WriteAllText(afterXmlPath, afterXml, new UTF8Encoding(false));
            var parsed = XDocument.Parse(afterXml, LoadOptions.PreserveWhitespace);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace officeMath = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var mathParagraphs = parsed.Descendants(officeMath + "oMathPara").ToArray();
            var mathNodes = parsed.Descendants(officeMath + "oMath").ToArray();
            var persistedJc = mathParagraphs
                .SelectMany(node => node
                    .Elements(officeMath + "oMathParaPr")
                    .Elements(officeMath + "jc"))
                .Select(node => (string?)node.Attribute(officeMath + "val"))
                .FirstOrDefault();
            var mathOnly = mathNodes.Length == 1
                && string.Equals(
                    string.Concat(mathNodes[0]
                        .Descendants(officeMath + "t")
                        .Select(text => text.Value)),
                    StyleSeparatorFormulaText,
                    StringComparison.Ordinal)
                && !mathNodes[0].Descendants(word + "tab").Any()
                && !mathNodes[0].Descendants(word + "instrText").Any();

            ExportStyleSeparatorFixedLayout(document, xpsPath, name);
            ExportStyleSeparatorPdfLayout(document, pdfPath, name);
            var glyphs = ReadStyleSeparatorFixedGlyphs(xpsPath, name);
            var equalsGlyphs = glyphs
                .Where(glyph => glyph.Text.IndexOf('=') >= 0)
                .ToArray();
            var numberGlyph = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('(') >= 0);
            var numberRendered = numberGlyph is not null
                && glyphs.Any(glyph => glyph.Text.IndexOf(')') >= 0);
            var formulaGlyph = equalsGlyphs.FirstOrDefault();
            var sameLine = numberRendered
                && formulaGlyph is not null
                && Math.Abs(formulaGlyph.OriginY - numberGlyph!.OriginY) <= 1d;
            var rightAligned = numberRendered && numberGlyph!.OriginX >= 450d;
            var candidate = math.Type == Word.WdOMathType.wdOMathDisplay
                && mathParagraphs.Length == 1
                && mathOnly
                && numberRendered
                && sameLine
                && rightAligned
                && equalsGlyphs.Length == 1;
            var glyphSummary = string.Join(
                " | ",
                glyphs.Select(glyph =>
                    $"'{EscapeStyleSeparatorGlyphText(glyph.Text)}'@{glyph.OriginX:0.###},{glyph.OriginY:0.###}"));
            Console.WriteLine(
                $"  raw-jc {name}: requested='{rawJustification ?? "<absent>"}', liveType={math.Type}, liveJc={math.Justification}, persistedJc='{persistedJc ?? "<absent>"}', mathOnly={mathOnly}, numberRendered={numberRendered}, sameLine={sameLine}, right={rightAligned}, formulaCopies={equalsGlyphs.Length}, candidate={candidate}, glyphs={glyphSummary}.");
            return candidate;
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
            Release(mathRange);
            Release(math);
            Release(addedMaths);
            Release(addedMathRange);
            Release(separatorPoint);
            Release(selection);
            Release(rightTab);
            Release(centerTab);
            Release(tabs);
            Release(format);
            Release(paragraph);
            Release(paragraphs);
            Release(numberFont);
            Release(numberRange);
            Release(formulaRange);
            Release(content);
            Release(document);
            ForceComCleanup();
        }
    }

    private static void PatchStyleSeparatorRawJustification(
        string docxPath,
        string? rawJustification,
        string context)
    {
        using var archive = ZipFile.Open(
            docxPath,
            ZipArchiveMode.Update);
        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException(
                context + ": DOCX has no word/document.xml part.");
        XDocument documentXml;
        using (var input = entry.Open())
            documentXml = XDocument.Load(input, LoadOptions.PreserveWhitespace);

        XNamespace officeMath = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var mathParagraphs = documentXml.Descendants(officeMath + "oMathPara").ToArray();
        AssertEqual(1, mathParagraphs.Length,
            context + ": package patch did not find exactly one m:oMathPara.");
        var mathParagraph = mathParagraphs[0];
        var properties = mathParagraph.Element(officeMath + "oMathParaPr");
        if (properties is null)
        {
            properties = new XElement(officeMath + "oMathParaPr");
            mathParagraph.AddFirst(properties);
        }
        properties.Elements(officeMath + "jc").Remove();
        if (rawJustification is not null)
        {
            properties.Add(new XElement(
                officeMath + "jc",
                new XAttribute(officeMath + "val", rawJustification)));
        }
        if (!properties.HasElements)
            properties.Remove();

        using var output = entry.Open();
        output.SetLength(0);
        using var writer = new StreamWriter(
            output,
            new UTF8Encoding(false),
            bufferSize: 4096,
            leaveOpen: true);
        documentXml.Save(writer, SaveOptions.DisableFormatting);
        writer.Flush();
    }
}

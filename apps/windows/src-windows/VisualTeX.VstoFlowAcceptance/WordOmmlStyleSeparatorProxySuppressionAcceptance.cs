using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class StyleSeparatorProxySuppressionScenario
    {
        public string Name { get; set; } = string.Empty;
        public float? FirstParagraphExactLineSpacing { get; set; }
        public float? FormulaParagraphExactLineSpacing { get; set; }
        public float FormulaParagraphLeftIndent { get; set; }
        public float FormulaParagraphRightIndent { get; set; }
        public Word.WdOMathJc Justification { get; set; } =
            Word.WdOMathJc.wdOMathJcCenterGroup;
        public bool HideMathRun { get; set; }
        public bool HideFormulaParagraphRange { get; set; }
        public int MathPosition { get; set; }
    }

    private static void RunWordOmmlStyleSeparatorProxySuppressionAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator proxy-suppression matrix refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        var scenarios = new List<StyleSeparatorProxySuppressionScenario>
        {
            new() { Name = "control" },
            new() { Name = "hide-math-run", HideMathRun = true },
            new() { Name = "hide-formula-paragraph-range", HideFormulaParagraphRange = true },
            new() { Name = "math-position-minus-100", MathPosition = -100 },
            new() { Name = "math-position-plus-100", MathPosition = 100 },
        };
        foreach (var spacing in new[] { 1f, 2f, 4f, 6f, 8f, 10f, 12f, 14f, 16f, 20f, 24f, 30f })
        {
            scenarios.Add(new StyleSeparatorProxySuppressionScenario
            {
                Name = "formula-exact-line-" + spacing.ToString("0"),
                FormulaParagraphExactLineSpacing = spacing,
            });
            scenarios.Add(new StyleSeparatorProxySuppressionScenario
            {
                Name = "first-exact-line-" + spacing.ToString("0"),
                FirstParagraphExactLineSpacing = spacing,
            });
        }
        foreach (var indent in new[] { -1000f, -700f, -500f, -350f, -250f, -150f, -72f, 72f, 150f, 250f, 350f })
        {
            scenarios.Add(new StyleSeparatorProxySuppressionScenario
            {
                Name = "formula-left-jc-indent-" + StyleSeparatorSignedToken(indent),
                FormulaParagraphLeftIndent = indent,
                Justification = Word.WdOMathJc.wdOMathJcLeft,
            });
            scenarios.Add(new StyleSeparatorProxySuppressionScenario
            {
                Name = "formula-center-jc-indent-" + StyleSeparatorSignedToken(indent),
                FormulaParagraphLeftIndent = indent,
                Justification = Word.WdOMathJc.wdOMathJcCenterGroup,
            });
        }
        foreach (var indent in new[] { -1000f, -700f, -500f, -350f, -250f, -150f, -72f, 72f, 150f, 250f, 350f })
        {
            scenarios.Add(new StyleSeparatorProxySuppressionScenario
            {
                Name = "formula-right-jc-right-indent-" + StyleSeparatorSignedToken(indent),
                FormulaParagraphRightIndent = indent,
                Justification = Word.WdOMathJc.wdOMathJcRight,
            });
        }
        foreach (var spacing in new[] { 1f, 2f, 4f })
        {
            foreach (var indent in new[] { -1000f, -700f, -500f, -350f })
            {
                scenarios.Add(new StyleSeparatorProxySuppressionScenario
                {
                    Name = "formula-exact-" + spacing.ToString("0")
                        + "-left-indent-" + StyleSeparatorSignedToken(indent),
                    FormulaParagraphExactLineSpacing = spacing,
                    FormulaParagraphLeftIndent = indent,
                    Justification = Word.WdOMathJc.wdOMathJcLeft,
                });
            }
        }

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var candidates = new List<string>();
            foreach (var scenario in scenarios)
            {
                try
                {
                    if (RunStyleSeparatorProxySuppressionScenario(
                            application,
                            artifactRoot,
                            scenario))
                        candidates.Add(scenario.Name);
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"  proxy-suppression {scenario.Name}: ERROR {error.GetType().Name}: {error.Message}");
                }
            }

            Console.WriteLine(
                "Style Separator proxy-suppression candidates: "
                + (candidates.Count == 0
                    ? "<none>"
                    : string.Join(", ", candidates)));
            Console.WriteLine(
                "Word OMML Style Separator proxy-suppression matrix completed; line spacing, independent formula-paragraph indents and hidden/position formatting were evaluated from XPS output.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static bool RunStyleSeparatorProxySuppressionScenario(
        Word.Application application,
        string artifactRoot,
        StyleSeparatorProxySuppressionScenario scenario)
    {
        var docxPath = Path.Combine(artifactRoot, scenario.Name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, scenario.Name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, scenario.Name + ".pdf");
        var xmlPath = Path.Combine(artifactRoot, scenario.Name + "-wordopenxml.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(xmlPath),
            scenario.Name + ": proxy-suppression artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
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
        Word.Range? formulaRange = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Microsoft.Office.Interop.Word.Font? mathFont = null;
        Word.Range? formulaParagraphRange = null;
        Microsoft.Office.Interop.Word.Font? formulaParagraphFont = null;
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
                scenario.Name + ": Word did not create three logical paragraphs.");
            for (var index = 1; index <= 3; index++)
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
                scenario.Name + " number");
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
                scenario.Name + " formula");
            addedMathRange = document.OMaths.Add(formulaRange);
            addedMaths = addedMathRange.OMaths;
            AssertEqual(1, addedMaths.Count,
                scenario.Name + ": Word did not create one OMath.");
            math = addedMaths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            math.Justification = scenario.Justification;
            mathRange = math.Range;
            mathFont = mathRange.Font;
            mathFont.Size = 14f;
            if (scenario.HideMathRun)
                mathFont.Hidden = -1;
            if (scenario.MathPosition != 0)
                mathFont.Position = scenario.MathPosition;

            Release(paragraph); paragraph = null;
            Release(paragraphs); paragraphs = null;
            paragraphs = mathRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                scenario.Name + ": OMath spans multiple logical paragraphs.");
            paragraph = paragraphs[1];
            format = paragraph.Format;
            format.LeftIndent = scenario.FormulaParagraphLeftIndent;
            format.RightIndent = scenario.FormulaParagraphRightIndent;
            if (scenario.FormulaParagraphExactLineSpacing is float formulaSpacing)
            {
                format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                format.LineSpacing = formulaSpacing;
            }
            formulaParagraphRange = paragraph.Range;
            if (scenario.HideFormulaParagraphRange)
            {
                formulaParagraphFont = formulaParagraphRange.Font;
                formulaParagraphFont.Hidden = -1;
            }

            if (scenario.FirstParagraphExactLineSpacing is float firstSpacing)
            {
                Release(format); format = null;
                Release(paragraph); paragraph = null;
                Release(paragraphs); paragraphs = null;
                paragraphs = document.Paragraphs;
                paragraph = paragraphs[1];
                format = paragraph.Format;
                format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                format.LineSpacing = firstSpacing;
            }

            document.Repaginate();
            var openXml = document.Content.WordOpenXML ?? string.Empty;
            File.WriteAllText(xmlPath, openXml, new UTF8Encoding(false));
            var parsed = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace officeMath = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var mathParagraphs = parsed.Descendants(officeMath + "oMathPara").ToArray();
            var mathNodes = parsed.Descendants(officeMath + "oMath").ToArray();
            var mathOnly = mathNodes.Length == 1
                && string.Equals(
                    string.Concat(mathNodes[0]
                        .Descendants(officeMath + "t")
                        .Select(text => text.Value)),
                    StyleSeparatorFormulaText,
                    StringComparison.Ordinal)
                && !mathNodes[0].Descendants(word + "tab").Any()
                && !mathNodes[0].Descendants(word + "instrText").Any();

            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            ExportStyleSeparatorFixedLayout(document, xpsPath, scenario.Name);
            ExportStyleSeparatorPdfLayout(document, pdfPath, scenario.Name);
            var glyphs = ReadStyleSeparatorFixedGlyphs(xpsPath, scenario.Name);
            var equalsGlyphs = glyphs
                .Where(glyph => glyph.Text.IndexOf('=') >= 0)
                .ToArray();
            var numberGlyph = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('(') >= 0);
            var numberRendered = numberGlyph is not null
                && glyphs.Any(glyph => glyph.Text.IndexOf(')') >= 0);
            var firstFormula = equalsGlyphs.FirstOrDefault();
            var sameLine = numberRendered
                && firstFormula is not null
                && Math.Abs(firstFormula.OriginY - numberGlyph!.OriginY) <= 1d;
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
                $"  proxy-suppression {scenario.Name}: liveType={math.Type}, mathOnly={mathOnly}, number={numberRendered}/{sameLine}/{rightAligned}, formulaCopies={equalsGlyphs.Length}, candidate={candidate}, formulaParagraph(line={format.LineSpacingRule}/{format.LineSpacing:0.###}, left={format.LeftIndent:0.###}, right={format.RightIndent:0.###}), glyphs={glyphSummary}.");
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
            Release(formulaParagraphFont);
            Release(formulaParagraphRange);
            Release(mathFont);
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
            Release(numberRange);
            Release(content);
            Release(document);
            ForceComCleanup();
        }
    }

    private static string StyleSeparatorSignedToken(float value)
    {
        return value < 0f
            ? "minus-" + Math.Abs(value).ToString("0")
            : "plus-" + value.ToString("0");
    }
}

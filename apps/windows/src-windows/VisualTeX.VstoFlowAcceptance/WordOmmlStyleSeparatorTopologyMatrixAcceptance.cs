using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class StyleSeparatorTopologyScenario
    {
        public string Name { get; set; } = string.Empty;
        public string[] Segments { get; set; } = Array.Empty<string>();
        public int FormulaSegmentIndex { get; set; }
        public int NumberSegmentIndex { get; set; }
        public int? HiddenSegmentIndex { get; set; }
        public bool ClearFirstSeparatorHidden { get; set; }
    }

    private sealed class StyleSeparatorTopologyResult
    {
        public string Name { get; set; } = string.Empty;
        public bool Display { get; set; }
        public bool MathOnly { get; set; }
        public bool NumberOutside { get; set; }
        public bool NumberRendered { get; set; }
        public bool SameLine { get; set; }
        public bool RightAligned { get; set; }
        public int Paragraphs { get; set; }
        public int Separators { get; set; }
        public double FormulaY { get; set; } = -1d;
        public double NumberY { get; set; } = -1d;
        public double NumberX { get; set; } = -1d;

        public bool Candidate =>
            Display
            && MathOnly
            && NumberOutside
            && NumberRendered
            && SameLine
            && RightAligned;
    }

    private static void RunWordOmmlStyleSeparatorTopologyMatrixAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator topology matrix refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        const string Zws = "\u200B";
        const string WordJoiner = "\u2060";
        var scenarios = new[]
        {
            new StyleSeparatorTopologyScenario
            {
                Name = "formula-number-control",
                Segments = new[] { StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 0,
                NumberSegmentIndex = 1,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-empty-formula-number",
                Segments = new[] { string.Empty, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-zws-formula-number",
                Segments = new[] { Zws, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-hidden-zws-formula-number",
                Segments = new[] { Zws, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
                HiddenSegmentIndex = 0,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-word-joiner-formula-number",
                Segments = new[] { WordJoiner, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-hidden-word-joiner-formula-number",
                Segments = new[] { WordJoiner, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
                HiddenSegmentIndex = 0,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-hidden-space-formula-number",
                Segments = new[] { " ", StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
                HiddenSegmentIndex = 0,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-letter-formula-number-diagnostic",
                Segments = new[] { "A", StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "formula-number-trailing-empty",
                Segments = new[] { StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, string.Empty },
                FormulaSegmentIndex = 0,
                NumberSegmentIndex = 1,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "formula-number-trailing-zws",
                Segments = new[] { StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, Zws },
                FormulaSegmentIndex = 0,
                NumberSegmentIndex = 1,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "formula-number-trailing-hidden-zws",
                Segments = new[] { StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, Zws },
                FormulaSegmentIndex = 0,
                NumberSegmentIndex = 1,
                HiddenSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "formula-number-trailing-word-joiner",
                Segments = new[] { StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, WordJoiner },
                FormulaSegmentIndex = 0,
                NumberSegmentIndex = 1,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "formula-number-trailing-letter-diagnostic",
                Segments = new[] { StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, "Z" },
                FormulaSegmentIndex = 0,
                NumberSegmentIndex = 1,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-zws-formula-number-trailing-empty",
                Segments = new[] { Zws, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, string.Empty },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-zws-formula-number-trailing-zws",
                Segments = new[] { Zws, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, Zws },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-letter-formula-number-trailing-letter-diagnostic",
                Segments = new[] { "A", StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText, "Z" },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
            },
            new StyleSeparatorTopologyScenario
            {
                Name = "leading-zws-formula-number-first-hidden-off",
                Segments = new[] { Zws, StyleSeparatorFormulaText, "\t" + StyleSeparatorNumberText },
                FormulaSegmentIndex = 1,
                NumberSegmentIndex = 2,
                ClearFirstSeparatorHidden = true,
            },
        };

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var results = new List<StyleSeparatorTopologyResult>();
            foreach (var scenario in scenarios)
            {
                try
                {
                    results.Add(RunStyleSeparatorTopologyScenario(
                        application,
                        artifactRoot,
                        scenario));
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"  topology {scenario.Name}: ERROR {error.GetType().Name}: {error.Message}");
                }
            }

            var candidates = results
                .Where(result => result.Candidate)
                .Select(result => result.Name)
                .ToArray();
            Console.WriteLine(
                "Style Separator topology candidates: "
                + (candidates.Length == 0
                    ? "<none>"
                    : string.Join(", ", candidates)));
            Console.WriteLine(
                "Word OMML Style Separator topology matrix completed; all topology results came from live OMath/OOXML plus Word XPS fixed-layout glyphs.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static StyleSeparatorTopologyResult RunStyleSeparatorTopologyScenario(
        Word.Application application,
        string artifactRoot,
        StyleSeparatorTopologyScenario scenario)
    {
        AssertTrue(scenario.Segments.Length >= 2,
            scenario.Name + ": topology requires at least two logical paragraphs.");
        AssertTrue(scenario.FormulaSegmentIndex >= 0
                   && scenario.FormulaSegmentIndex < scenario.Segments.Length,
            scenario.Name + ": invalid formula segment index.");
        AssertTrue(scenario.NumberSegmentIndex >= 0
                   && scenario.NumberSegmentIndex < scenario.Segments.Length,
            scenario.Name + ": invalid number segment index.");
        AssertTrue(scenario.FormulaSegmentIndex != scenario.NumberSegmentIndex,
            scenario.Name + ": formula and number cannot share a logical paragraph.");

        var docxPath = Path.Combine(artifactRoot, scenario.Name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, scenario.Name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, scenario.Name + ".pdf");
        var xmlPath = Path.Combine(artifactRoot, scenario.Name + "-wordopenxml.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(xmlPath),
            scenario.Name + ": topology artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
        Word.Range? formulaSource = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? numberRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Range? hiddenRange = null;
        Microsoft.Office.Interop.Word.Font? hiddenFont = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Word.TabStops? tabs = null;
        Word.TabStop? rightTab = null;
        Word.Selection? selection = null;
        Word.Range? separatorPoint = null;
        Word.Range? firstSeparatorMark = null;
        Microsoft.Office.Interop.Word.Font? firstSeparatorFont = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            content = document.Content;
            content.Text = string.Join("\r", scenario.Segments);

            sections = document.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            pageSetup.PageWidth = 595.3f;
            pageSetup.PageHeight = 841.9f;
            pageSetup.LeftMargin = 90f;
            pageSetup.RightMargin = 90f;
            pageSetup.TopMargin = 72f;
            pageSetup.BottomMargin = 72f;

            paragraphs = document.Paragraphs;
            AssertEqual(scenario.Segments.Length, paragraphs.Count,
                scenario.Name + ": Word materialized a different logical paragraph count.");

            paragraph = paragraphs[scenario.FormulaSegmentIndex + 1];
            formulaSource = document.Range(
                paragraph.Range.Start,
                paragraph.Range.Start + StyleSeparatorFormulaText.Length);
            AssertEqual(StyleSeparatorFormulaText, formulaSource.Text ?? string.Empty,
                scenario.Name + ": formula segment was not found at its expected paragraph start.");
            addedMathRange = document.OMaths.Add(formulaSource);
            addedMaths = addedMathRange.OMaths;
            AssertEqual(1, addedMaths.Count,
                scenario.Name + ": Word did not create exactly one OMath.");
            math = addedMaths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            math.Justification = Word.WdOMathJc.wdOMathJcCenterGroup;
            mathRange = math.Range;
            var mathFont = mathRange.Font;
            try { mathFont.Size = 14f; }
            finally { Release(mathFont); }

            Release(paragraph); paragraph = null;
            Release(paragraphs); paragraphs = null;
            paragraphs = document.Paragraphs;
            paragraph = paragraphs[scenario.NumberSegmentIndex + 1];
            var numberText = paragraph.Range.Text ?? string.Empty;
            var numberOffset = numberText.IndexOf(
                StyleSeparatorNumberText,
                StringComparison.Ordinal);
            AssertTrue(numberOffset >= 0,
                scenario.Name + ": number text is absent from its declared logical paragraph.");
            numberRange = document.Range(
                paragraph.Range.Start + numberOffset,
                paragraph.Range.Start + numberOffset + StyleSeparatorNumberText.Length);
            numberFont = numberRange.Font;
            numberFont.Name = StyleSeparatorNumberFont;
            numberFont.NameAscii = StyleSeparatorNumberFont;
            numberFont.NameOther = StyleSeparatorNumberFont;
            numberFont.Size = StyleSeparatorNumberFontSize;
            numberFont.Bold = 0;
            numberFont.Italic = 0;
            numberFont.Position = 0;
            numberFont.Hidden = 0;

            if (scenario.HiddenSegmentIndex is int hiddenIndex)
            {
                Release(paragraph); paragraph = null;
                Release(paragraphs); paragraphs = null;
                paragraphs = document.Paragraphs;
                paragraph = paragraphs[hiddenIndex + 1];
                var segmentLength = scenario.Segments[hiddenIndex].Length;
                AssertTrue(segmentLength > 0,
                    scenario.Name + ": an empty segment cannot be marked hidden by Range.Font.");
                hiddenRange = document.Range(
                    paragraph.Range.Start,
                    paragraph.Range.Start + segmentLength);
                hiddenFont = hiddenRange.Font;
                hiddenFont.Hidden = -1;
            }

            Release(paragraph); paragraph = null;
            Release(paragraphs); paragraphs = null;
            paragraphs = document.Paragraphs;
            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(rightTab); rightTab = null;
                Release(tabs); tabs = null;
                Release(paragraphFormat); paragraphFormat = null;
                Release(paragraph); paragraph = null;
                paragraph = paragraphs[index];
                paragraphFormat = paragraph.Format;
                ConfigureStyleSeparatorMatrixParagraph(
                    paragraphFormat,
                    Word.WdParagraphAlignment.wdAlignParagraphLeft);
                tabs = paragraphFormat.TabStops;
                tabs.ClearAll();
                rightTab = tabs.Add(
                    StyleSeparatorRightTabPosition,
                    Word.WdTabAlignment.wdAlignTabRight,
                    Word.WdTabLeader.wdTabLeaderSpaces);
            }

            selection = application.Selection;
            for (var index = 1; index < scenario.Segments.Length; index++)
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

            if (scenario.ClearFirstSeparatorHidden)
            {
                Release(paragraph); paragraph = null;
                Release(paragraphs); paragraphs = null;
                paragraphs = document.Paragraphs;
                paragraph = paragraphs[1];
                firstSeparatorMark = document.Range(
                    paragraph.Range.End - 1,
                    paragraph.Range.End);
                firstSeparatorFont = firstSeparatorMark.Font;
                firstSeparatorFont.Hidden = 0;
            }

            document.Repaginate();
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
                && !mathNodes[0].Descendants(word + "tab").Any()
                && !mathNodes[0].Descendants(word + "instrText").Any();
            var numberParagraphs = parsed.Descendants(word + "p")
                .Where(candidate => string.Concat(candidate
                        .Descendants(word + "t")
                        .Select(text => text.Value))
                    .IndexOf(StyleSeparatorNumberText, StringComparison.Ordinal) >= 0)
                .ToArray();
            var numberOutside = numberParagraphs.Length == 1
                && !numberParagraphs[0]
                    .Descendants(officeMath + "oMath")
                    .Any();
            var separators = parsed.Descendants(word + "specVanish").Count();

            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            ExportStyleSeparatorFixedLayout(document, xpsPath, scenario.Name);
            ExportStyleSeparatorPdfLayout(document, pdfPath, scenario.Name);
            var glyphs = ReadStyleSeparatorFixedGlyphs(xpsPath, scenario.Name);
            var formulaGlyph = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('=') >= 0);
            var numberGlyph = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('(') >= 0);
            var numberRendered = numberGlyph is not null
                && glyphs.Any(glyph => glyph.Text.IndexOf(')') >= 0);
            var formulaY = formulaGlyph?.OriginY ?? -1d;
            var numberY = numberGlyph?.OriginY ?? -1d;
            var sameLine = numberRendered
                && formulaY >= 0d
                && Math.Abs(formulaY - numberY) <= 1d;
            var rightAligned = numberRendered
                && numberGlyph!.OriginX >= 450d;

            Release(mathRange); mathRange = null;
            Release(math); math = null;
            var liveMaths = document.OMaths;
            try
            {
                AssertEqual(1, liveMaths.Count,
                    scenario.Name + ": topology no longer contains one OMath.");
                math = liveMaths[1];
                mathRange = math.Range;
            }
            finally { Release(liveMaths); }

            var result = new StyleSeparatorTopologyResult
            {
                Name = scenario.Name,
                Display = math.Type == Word.WdOMathType.wdOMathDisplay
                    && mathParagraphs.Length == 1,
                MathOnly = mathOnly,
                NumberOutside = numberOutside,
                NumberRendered = numberRendered,
                SameLine = sameLine,
                RightAligned = rightAligned,
                Paragraphs = document.Paragraphs.Count,
                Separators = separators,
                FormulaY = formulaY,
                NumberY = numberY,
                NumberX = numberGlyph?.OriginX ?? -1d,
            };
            var glyphText = EscapeStyleSeparatorGlyphText(
                string.Concat(glyphs.Select(glyph => glyph.Text)));
            Console.WriteLine(
                $"  topology {scenario.Name}: display={result.Display}, mathOnly={result.MathOnly}, numberOutside={result.NumberOutside}, rendered={result.NumberRendered}, sameLine={result.SameLine}, right={result.RightAligned}, paragraphs={result.Paragraphs}, separators={result.Separators}, formulaY={result.FormulaY:0.###}, number={result.NumberX:0.###},{result.NumberY:0.###}, candidate={result.Candidate}, glyphText='{glyphText}'.");
            return result;
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(firstSeparatorFont);
            Release(firstSeparatorMark);
            Release(separatorPoint);
            Release(selection);
            Release(rightTab);
            Release(tabs);
            Release(paragraphFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(hiddenFont);
            Release(hiddenRange);
            Release(numberFont);
            Release(numberRange);
            Release(mathRange);
            Release(math);
            Release(addedMaths);
            Release(addedMathRange);
            Release(formulaSource);
            Release(content);
            Release(document);
            ForceComCleanup();
        }
    }
}

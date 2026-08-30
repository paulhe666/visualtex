using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class StyleSeparatorRescueScenario
    {
        public string Name { get; set; } = string.Empty;
        public bool DistinctCustomStyles { get; set; }
        public bool ClearSeparatorMarkHidden { get; set; }
        public bool PrintHiddenText { get; set; }
        public string FormulaBridgeText { get; set; } = string.Empty;
        public bool FormulaBridgeHidden { get; set; }
        public bool UseMiddleParagraph { get; set; }
        public string MiddleParagraphText { get; set; } = string.Empty;
        public bool MiddleParagraphHidden { get; set; }
    }

    private sealed class StyleSeparatorRescueResult
    {
        public string Name { get; set; } = string.Empty;
        public bool GenuineDisplay { get; set; }
        public bool MathContainsOnlyFormula { get; set; }
        public bool NumberOutsideMath { get; set; }
        public bool NumberRendered { get; set; }
        public bool SameFixedLine { get; set; }
        public bool RightAligned { get; set; }
        public int StyleSeparatorCount { get; set; }
        public int ParagraphCount { get; set; }
        public int NumberHidden { get; set; }
        public double FormulaOriginX { get; set; } = -1d;
        public double FormulaOriginY { get; set; } = -1d;
        public double NumberOriginX { get; set; } = -1d;
        public double NumberOriginY { get; set; } = -1d;

        public bool IsCandidate =>
            GenuineDisplay
            && MathContainsOnlyFormula
            && NumberOutsideMath
            && NumberRendered
            && SameFixedLine
            && RightAligned
            && StyleSeparatorCount >= 1
            && NumberHidden == 0;
    }

    private static void RunWordOmmlStyleSeparatorRescueMatrixAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator rescue matrix refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        var scenarios = new[]
        {
            new StyleSeparatorRescueScenario
            {
                Name = "baseline-two-paragraph",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "separator-mark-hidden-off",
                ClearSeparatorMarkHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "distinct-custom-styles",
                DistinctCustomStyles = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "distinct-styles-hidden-off",
                DistinctCustomStyles = true,
                ClearSeparatorMarkHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "print-hidden-text",
                PrintHiddenText = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "formula-hidden-space",
                FormulaBridgeText = " ",
                FormulaBridgeHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "formula-hidden-zero-width-space",
                FormulaBridgeText = "\u200B",
                FormulaBridgeHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "formula-hidden-word-joiner",
                FormulaBridgeText = "\u2060",
                FormulaBridgeHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "formula-visible-zero-width-space",
                FormulaBridgeText = "\u200B",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "formula-visible-word-joiner",
                FormulaBridgeText = "\u2060",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "formula-visible-letter-diagnostic",
                FormulaBridgeText = "B",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-empty-middle",
                UseMiddleParagraph = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-hidden-space",
                UseMiddleParagraph = true,
                MiddleParagraphText = " ",
                MiddleParagraphHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-hidden-zero-width-space",
                UseMiddleParagraph = true,
                MiddleParagraphText = "\u200B",
                MiddleParagraphHidden = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-visible-zero-width-space",
                UseMiddleParagraph = true,
                MiddleParagraphText = "\u200B",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-visible-word-joiner",
                UseMiddleParagraph = true,
                MiddleParagraphText = "\u2060",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-visible-letter-diagnostic",
                UseMiddleParagraph = true,
                MiddleParagraphText = "B",
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-distinct-styles",
                UseMiddleParagraph = true,
                DistinctCustomStyles = true,
            },
            new StyleSeparatorRescueScenario
            {
                Name = "chain-hidden-off",
                UseMiddleParagraph = true,
                ClearSeparatorMarkHidden = true,
            },
        };

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var results = new List<StyleSeparatorRescueResult>();
            foreach (var scenario in scenarios)
            {
                try
                {
                    results.Add(RunStyleSeparatorRescueScenario(
                        application,
                        artifactRoot,
                        scenario));
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"  rescue {scenario.Name}: ERROR {error.GetType().Name}: {error.Message}");
                }
            }

            var candidates = results
                .Where(result => result.IsCandidate)
                .Select(result => result.Name)
                .ToArray();
            Console.WriteLine(
                "Style Separator rescue candidates: "
                + (candidates.Length == 0
                    ? "<none>"
                    : string.Join(", ", candidates)));
            Console.WriteLine(
                "Word OMML Style Separator rescue matrix completed; every result was evaluated from live OMath type, WordOpenXML and XPS glyph coordinates.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static StyleSeparatorRescueResult RunStyleSeparatorRescueScenario(
        Word.Application application,
        string artifactRoot,
        StyleSeparatorRescueScenario scenario)
    {
        var docxPath = Path.Combine(artifactRoot, scenario.Name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, scenario.Name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, scenario.Name + ".pdf");
        var xmlPath = Path.Combine(artifactRoot, scenario.Name + "-wordopenxml.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(xmlPath),
            scenario.Name + ": rescue artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
        Word.Range? formulaSource = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? formulaBridgeRange = null;
        Microsoft.Office.Interop.Word.Font? formulaBridgeFont = null;
        Word.Range? middleRange = null;
        Microsoft.Office.Interop.Word.Font? middleFont = null;
        Word.Range? numberRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Word.TabStops? paragraphTabs = null;
        Word.TabStop? paragraphTab = null;
        Word.Styles? styles = null;
        Word.Style? customStyle = null;
        Word.Selection? selection = null;
        Word.Range? separatorPoint = null;
        Word.Range? separatorMark = null;
        Microsoft.Office.Interop.Word.Font? separatorMarkFont = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Word.Options? options = null;
        var previousPrintHiddenText = false;
        var printHiddenCaptured = false;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            var middle = scenario.UseMiddleParagraph
                ? "\r" + scenario.MiddleParagraphText
                : string.Empty;
            content = document.Content;
            content.Text = StyleSeparatorFormulaText
                + scenario.FormulaBridgeText
                + middle
                + "\r\t"
                + StyleSeparatorNumberText;

            sections = document.Sections;
            section = sections[1];
            pageSetup = section.PageSetup;
            pageSetup.PageWidth = 595.3f;
            pageSetup.PageHeight = 841.9f;
            pageSetup.LeftMargin = 90f;
            pageSetup.RightMargin = 90f;
            pageSetup.TopMargin = 72f;
            pageSetup.BottomMargin = 72f;

            formulaSource = document.Range(0, StyleSeparatorFormulaText.Length);
            addedMathRange = document.OMaths.Add(formulaSource);
            addedMaths = addedMathRange.OMaths;
            AssertEqual(1, addedMaths.Count,
                scenario.Name + ": Word did not create exactly one OMath.");
            math = addedMaths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            mathRange = math.Range;
            var mathFont = mathRange.Font;
            try { mathFont.Size = 14f; }
            finally { Release(mathFont); }

            if (scenario.FormulaBridgeText.Length > 0)
            {
                formulaBridgeRange = document.Range(
                    mathRange.End,
                    mathRange.End + scenario.FormulaBridgeText.Length);
                formulaBridgeFont = formulaBridgeRange.Font;
                formulaBridgeFont.Hidden = scenario.FormulaBridgeHidden ? -1 : 0;
            }

            if (scenario.UseMiddleParagraph
                && scenario.MiddleParagraphText.Length > 0)
            {
                var middleStart = StyleSeparatorFormulaText.Length
                    + scenario.FormulaBridgeText.Length
                    + 1;
                middleRange = document.Range(
                    middleStart,
                    middleStart + scenario.MiddleParagraphText.Length);
                middleFont = middleRange.Font;
                middleFont.Hidden = scenario.MiddleParagraphHidden ? -1 : 0;
            }

            paragraphs = document.Paragraphs;
            var expectedParagraphs = scenario.UseMiddleParagraph ? 3 : 2;
            AssertEqual(expectedParagraphs, paragraphs.Count,
                scenario.Name + ": rescue scenario started with the wrong paragraph count.");

            if (scenario.DistinctCustomStyles)
            {
                styles = document.Styles;
                for (var index = 1; index <= expectedParagraphs; index++)
                {
                    Release(customStyle); customStyle = null;
                    Release(paragraph); paragraph = null;
                    customStyle = styles.Add(
                        "VTStyleSepRescue" + index,
                        Word.WdStyleType.wdStyleTypeParagraph);
                    paragraph = paragraphs[index];
                    var styleName = (object)customStyle.NameLocal;
                    var styledRange = paragraph.Range;
                    try { styledRange.set_Style(ref styleName); }
                    finally { Release(styledRange); }
                }
            }

            for (var index = 1; index <= expectedParagraphs; index++)
            {
                Release(paragraphTab); paragraphTab = null;
                Release(paragraphTabs); paragraphTabs = null;
                Release(paragraphFormat); paragraphFormat = null;
                Release(paragraph); paragraph = null;
                paragraph = paragraphs[index];
                paragraphFormat = paragraph.Format;
                ConfigureStyleSeparatorMatrixParagraph(
                    paragraphFormat,
                    Word.WdParagraphAlignment.wdAlignParagraphLeft);
                paragraphTabs = paragraphFormat.TabStops;
                paragraphTabs.ClearAll();
                paragraphTab = paragraphTabs.Add(
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

            if (scenario.ClearSeparatorMarkHidden)
            {
                for (var index = 1; index < expectedParagraphs; index++)
                {
                    Release(separatorMarkFont); separatorMarkFont = null;
                    Release(separatorMark); separatorMark = null;
                    Release(paragraph); paragraph = null;
                    Release(paragraphs); paragraphs = null;
                    paragraphs = document.Paragraphs;
                    paragraph = paragraphs[index];
                    separatorMark = document.Range(
                        paragraph.Range.End - 1,
                        paragraph.Range.End);
                    separatorMarkFont = separatorMark.Font;
                    separatorMarkFont.Hidden = 0;
                }
            }

            if (scenario.PrintHiddenText)
            {
                options = application.Options;
                previousPrintHiddenText = options.PrintHiddenText;
                printHiddenCaptured = true;
                options.PrintHiddenText = true;
            }

            document.Repaginate();
            Release(mathRange); mathRange = null;
            Release(math); math = null;
            Release(addedMaths); addedMaths = null;
            Release(addedMathRange); addedMathRange = null;
            var documentMaths = document.OMaths;
            try
            {
                AssertEqual(1, documentMaths.Count,
                    scenario.Name + ": rescue scenario no longer has exactly one OMath.");
                math = documentMaths[1];
                mathRange = math.Range;
            }
            finally { Release(documentMaths); }

            Release(numberFont); numberFont = null;
            Release(numberRange); numberRange = null;
            numberRange = FindUniqueStyleSeparatorTextRange(
                document,
                StyleSeparatorNumberText,
                scenario.Name + " post-separator number");
            numberFont = numberRange.Font;

            var openXml = document.Content.WordOpenXML ?? string.Empty;
            File.WriteAllText(xmlPath, openXml, new UTF8Encoding(false));
            var parsed = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace officeMath = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var mathParagraphs = parsed.Descendants(officeMath + "oMathPara").ToArray();
            var mathNodes = parsed.Descendants(officeMath + "oMath").ToArray();
            var specVanishCount = parsed.Descendants(word + "specVanish").Count();
            var mathContainsOnlyFormula = mathNodes.Length == 1
                && string.Equals(
                    string.Concat(mathNodes[0]
                        .Descendants(officeMath + "t")
                        .Select(text => text.Value)),
                    StyleSeparatorFormulaText,
                    StringComparison.Ordinal)
                && !mathNodes[0].Descendants(word + "tab").Any()
                && !mathNodes[0].Descendants(word + "instrText").Any();
            var numberXmlParagraphs = parsed.Descendants(word + "p")
                .Where(candidate => string.Concat(
                        candidate.Descendants(word + "t")
                            .Select(text => text.Value))
                    .IndexOf(StyleSeparatorNumberText, StringComparison.Ordinal) >= 0)
                .ToArray();
            var numberOutsideMath = numberXmlParagraphs.Length == 1
                && !numberXmlParagraphs[0]
                    .Descendants(officeMath + "oMath")
                    .Any();

            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            ExportStyleSeparatorFixedLayout(document, xpsPath, scenario.Name);
            ExportStyleSeparatorPdfLayout(document, pdfPath, scenario.Name);
            var glyphs = ReadStyleSeparatorFixedGlyphs(xpsPath, scenario.Name);
            var numberGlyph = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('(') >= 0);
            var formulaGlyph = glyphs.FirstOrDefault(glyph =>
                !string.IsNullOrWhiteSpace(glyph.Text)
                && glyph.Text.IndexOf('(') < 0
                && glyph.Text.IndexOf(')') < 0);
            var numberRendered = numberGlyph is not null
                && glyphs.Any(glyph => glyph.Text.IndexOf(')') >= 0);
            var formulaY = formulaGlyph?.OriginY ?? -1d;
            var numberY = numberGlyph?.OriginY ?? -1d;
            var sameFixedLine = numberRendered
                && formulaY >= 0d
                && Math.Abs(formulaY - numberY) <= 1d;
            var rightAligned = numberRendered
                && numberGlyph!.OriginX >= 450d;
            var result = new StyleSeparatorRescueResult
            {
                Name = scenario.Name,
                GenuineDisplay = math.Type == Word.WdOMathType.wdOMathDisplay
                    && mathParagraphs.Length == 1,
                MathContainsOnlyFormula = mathContainsOnlyFormula,
                NumberOutsideMath = numberOutsideMath,
                NumberRendered = numberRendered,
                SameFixedLine = sameFixedLine,
                RightAligned = rightAligned,
                StyleSeparatorCount = specVanishCount,
                ParagraphCount = document.Paragraphs.Count,
                NumberHidden = numberFont.Hidden,
                FormulaOriginX = formulaGlyph?.OriginX ?? -1d,
                FormulaOriginY = formulaY,
                NumberOriginX = numberGlyph?.OriginX ?? -1d,
                NumberOriginY = numberY,
            };
            var glyphText = EscapeStyleSeparatorGlyphText(
                string.Concat(glyphs.Select(glyph => glyph.Text)));
            Console.WriteLine(
                $"  rescue {scenario.Name}: display={result.GenuineDisplay}, mathOnly={result.MathContainsOnlyFormula}, numberOutside={result.NumberOutsideMath}, rendered={result.NumberRendered}, sameLine={result.SameFixedLine}, right={result.RightAligned}, specVanish={result.StyleSeparatorCount}, paragraphs={result.ParagraphCount}, numberHidden={result.NumberHidden}, formula={result.FormulaOriginX:0.###},{result.FormulaOriginY:0.###}, number={result.NumberOriginX:0.###},{result.NumberOriginY:0.###}, candidate={result.IsCandidate}, glyphText='{glyphText}'.");
            return result;
        }
        finally
        {
            if (printHiddenCaptured && options is not null)
            {
                try { options.PrintHiddenText = previousPrintHiddenText; } catch { }
            }
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(options);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(separatorMarkFont);
            Release(separatorMark);
            Release(separatorPoint);
            Release(selection);
            Release(customStyle);
            Release(styles);
            Release(paragraphTab);
            Release(paragraphTabs);
            Release(paragraphFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(numberFont);
            Release(numberRange);
            Release(middleFont);
            Release(middleRange);
            Release(formulaBridgeFont);
            Release(formulaBridgeRange);
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

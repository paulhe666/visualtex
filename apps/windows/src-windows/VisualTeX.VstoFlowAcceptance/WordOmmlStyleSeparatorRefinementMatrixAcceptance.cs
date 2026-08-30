using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private enum StyleSeparatorMathCreationOrder
    {
        DisplayBeforeSeparators,
        InlineBeforeDisplayAfterSeparators,
        CreateAfterSeparators,
    }

    private sealed class StyleSeparatorRefinementScenario
    {
        public string Name { get; set; } = string.Empty;
        public string AnchorText { get; set; } = "\u200B";
        public bool AnchorStartsWithCenterTab { get; set; }
        public StyleSeparatorMathCreationOrder CreationOrder { get; set; }
        public Word.WdOMathJc Justification { get; set; } =
            Word.WdOMathJc.wdOMathJcCenterGroup;
        public bool TrimLeadingMathWhitespace { get; set; }
        public bool BuildUpAfterSeparators { get; set; }
    }

    private sealed class StyleSeparatorRefinementResult
    {
        public string Name { get; set; } = string.Empty;
        public bool Display { get; set; }
        public bool MathOnly { get; set; }
        public bool NumberRendered { get; set; }
        public bool SameLine { get; set; }
        public bool RightAligned { get; set; }
        public bool FormulaRenderedOnce { get; set; }
        public int MathParagraphs { get; set; }
        public int MathNodes { get; set; }
        public int FormulaGlyphGroups { get; set; }
        public double PrimaryFormulaX { get; set; } = -1d;
        public double PrimaryFormulaY { get; set; } = -1d;
        public double NumberX { get; set; } = -1d;
        public double NumberY { get; set; } = -1d;

        public bool Candidate =>
            Display
            && MathOnly
            && NumberRendered
            && SameLine
            && RightAligned
            && FormulaRenderedOnce;
    }

    private static void RunWordOmmlStyleSeparatorRefinementMatrixAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator refinement matrix refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        var scenarios = new List<StyleSeparatorRefinementScenario>();
        foreach (var justification in new[]
                 {
                     Word.WdOMathJc.wdOMathJcCenterGroup,
                     Word.WdOMathJc.wdOMathJcCenter,
                     Word.WdOMathJc.wdOMathJcLeft,
                     Word.WdOMathJc.wdOMathJcRight,
                     Word.WdOMathJc.wdOMathJcInline,
                 })
        {
            scenarios.Add(new StyleSeparatorRefinementScenario
            {
                Name = "zws-before-separators-" + justification,
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
                Justification = justification,
            });
            scenarios.Add(new StyleSeparatorRefinementScenario
            {
                Name = "zws-before-separators-trim-" + justification,
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
                Justification = justification,
                TrimLeadingMathWhitespace = true,
            });
            scenarios.Add(new StyleSeparatorRefinementScenario
            {
                Name = "zws-inline-then-display-" + justification,
                CreationOrder = StyleSeparatorMathCreationOrder.InlineBeforeDisplayAfterSeparators,
                Justification = justification,
            });
            scenarios.Add(new StyleSeparatorRefinementScenario
            {
                Name = "zws-create-after-separators-" + justification,
                CreationOrder = StyleSeparatorMathCreationOrder.CreateAfterSeparators,
                Justification = justification,
            });
        }

        scenarios.AddRange(new[]
        {
            new StyleSeparatorRefinementScenario
            {
                Name = "center-tab-before-separators-center-group",
                AnchorText = "\t",
                AnchorStartsWithCenterTab = true,
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "center-tab-before-separators-inline-jc",
                AnchorText = "\t",
                AnchorStartsWithCenterTab = true,
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
                Justification = Word.WdOMathJc.wdOMathJcInline,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "center-tab-create-after-separators-center-group",
                AnchorText = "\t",
                AnchorStartsWithCenterTab = true,
                CreationOrder = StyleSeparatorMathCreationOrder.CreateAfterSeparators,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "center-tab-create-after-separators-inline-jc",
                AnchorText = "\t",
                AnchorStartsWithCenterTab = true,
                CreationOrder = StyleSeparatorMathCreationOrder.CreateAfterSeparators,
                Justification = Word.WdOMathJc.wdOMathJcInline,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "center-tab-zws-before-separators-center-group",
                AnchorText = "\t\u200B",
                AnchorStartsWithCenterTab = true,
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "center-tab-zws-create-after-separators-center-group",
                AnchorText = "\t\u200B",
                AnchorStartsWithCenterTab = true,
                CreationOrder = StyleSeparatorMathCreationOrder.CreateAfterSeparators,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "zws-before-separators-rebuild",
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
                BuildUpAfterSeparators = true,
            },
            new StyleSeparatorRefinementScenario
            {
                Name = "zws-before-separators-trim-rebuild",
                CreationOrder = StyleSeparatorMathCreationOrder.DisplayBeforeSeparators,
                TrimLeadingMathWhitespace = true,
                BuildUpAfterSeparators = true,
            },
        });

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var results = new List<StyleSeparatorRefinementResult>();
            foreach (var scenario in scenarios)
            {
                try
                {
                    results.Add(RunStyleSeparatorRefinementScenario(
                        application,
                        artifactRoot,
                        scenario));
                }
                catch (Exception error)
                {
                    Console.WriteLine(
                        $"  refinement {scenario.Name}: ERROR {error.GetType().Name}: {error.Message}");
                }
            }

            var candidates = results
                .Where(result => result.Candidate)
                .Select(result => result.Name)
                .ToArray();
            Console.WriteLine(
                "Style Separator refinement candidates: "
                + (candidates.Length == 0
                    ? "<none>"
                    : string.Join(", ", candidates)));
            Console.WriteLine(
                "Word OMML Style Separator refinement matrix completed; creation order, math justification, leading-space cleanup and center-tab anchoring were all evaluated from live Word plus fixed-layout output.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static StyleSeparatorRefinementResult RunStyleSeparatorRefinementScenario(
        Word.Application application,
        string artifactRoot,
        StyleSeparatorRefinementScenario scenario)
    {
        var docxPath = Path.Combine(artifactRoot, scenario.Name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, scenario.Name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, scenario.Name + ".pdf");
        var xmlPath = Path.Combine(artifactRoot, scenario.Name + "-wordopenxml.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(xmlPath),
            scenario.Name + ": refinement artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
        Word.Range? formulaRange = null;
        Word.Range? numberRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? trimRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? paragraph = null;
        Word.ParagraphFormat? paragraphFormat = null;
        Word.TabStops? tabs = null;
        Word.TabStop? centerTab = null;
        Word.TabStop? rightTab = null;
        Word.Selection? selection = null;
        Word.Range? separatorPoint = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            content = document.Content;
            content.Text = scenario.AnchorText
                + "\r"
                + StyleSeparatorFormulaText
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

            paragraphs = document.Paragraphs;
            AssertEqual(3, paragraphs.Count,
                scenario.Name + ": refinement scenario did not start with three paragraphs.");

            for (var index = 1; index <= paragraphs.Count; index++)
            {
                Release(rightTab); rightTab = null;
                Release(centerTab); centerTab = null;
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
                if (scenario.AnchorStartsWithCenterTab)
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

            formulaRange = FindUniqueStyleSeparatorTextRange(
                document,
                StyleSeparatorFormulaText,
                scenario.Name + " formula");
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

            if (scenario.CreationOrder
                != StyleSeparatorMathCreationOrder.CreateAfterSeparators)
            {
                addedMathRange = document.OMaths.Add(formulaRange);
                addedMaths = addedMathRange.OMaths;
                AssertEqual(1, addedMaths.Count,
                    scenario.Name + ": Word did not create one pre-separator OMath.");
                math = addedMaths[1];
                math.BuildUp();
                math.Type = scenario.CreationOrder
                    == StyleSeparatorMathCreationOrder.DisplayBeforeSeparators
                    ? Word.WdOMathType.wdOMathDisplay
                    : Word.WdOMathType.wdOMathInline;
                math.Justification = scenario.Justification;
                mathRange = math.Range;
                var mathFont = mathRange.Font;
                try { mathFont.Size = 14f; }
                finally { Release(mathFont); }
            }

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

            if (scenario.CreationOrder
                == StyleSeparatorMathCreationOrder.CreateAfterSeparators)
            {
                Release(formulaRange); formulaRange = null;
                formulaRange = FindUniqueStyleSeparatorTextRange(
                    document,
                    StyleSeparatorFormulaText,
                    scenario.Name + " post-separator formula");
                addedMathRange = document.OMaths.Add(formulaRange);
                addedMaths = addedMathRange.OMaths;
                AssertEqual(1, addedMaths.Count,
                    scenario.Name + ": Word did not create one post-separator OMath.");
                math = addedMaths[1];
                math.BuildUp();
                math.Type = Word.WdOMathType.wdOMathDisplay;
                math.Justification = scenario.Justification;
                mathRange = math.Range;
                var mathFont = mathRange.Font;
                try { mathFont.Size = 14f; }
                finally { Release(mathFont); }
            }
            else if (scenario.CreationOrder
                     == StyleSeparatorMathCreationOrder.InlineBeforeDisplayAfterSeparators)
            {
                AssertTrue(math is not null,
                    scenario.Name + ": pre-separator inline OMath disappeared.");
                math!.Type = Word.WdOMathType.wdOMathDisplay;
                math.Justification = scenario.Justification;
            }

            if (scenario.TrimLeadingMathWhitespace)
            {
                Release(mathRange); mathRange = null;
                AssertTrue(math is not null,
                    scenario.Name + ": cannot trim a missing OMath.");
                mathRange = math!.Range;
                var textBeforeTrim = mathRange.Text ?? string.Empty;
                var leading = 0;
                while (leading < textBeforeTrim.Length
                       && char.IsWhiteSpace(textBeforeTrim[leading]))
                    leading++;
                if (leading > 0)
                {
                    trimRange = document.Range(
                        mathRange.Start,
                        mathRange.Start + leading);
                    trimRange.Delete();
                    Release(trimRange); trimRange = null;
                }
                math.Type = Word.WdOMathType.wdOMathDisplay;
                math.Justification = scenario.Justification;
            }

            if (scenario.BuildUpAfterSeparators)
            {
                AssertTrue(math is not null,
                    scenario.Name + ": cannot rebuild a missing OMath.");
                math!.BuildUp();
                math.Type = Word.WdOMathType.wdOMathDisplay;
                math.Justification = scenario.Justification;
            }

            document.Repaginate();
            Release(mathRange); mathRange = null;
            Release(math); math = null;
            var liveMaths = document.OMaths;
            try
            {
                AssertEqual(1, liveMaths.Count,
                    scenario.Name + ": refinement no longer contains exactly one OMath.");
                math = liveMaths[1];
                mathRange = math.Range;
            }
            finally { Release(liveMaths); }

            var liveMathText = mathRange.Text ?? string.Empty;
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
            var primaryFormula = equalsGlyphs.FirstOrDefault();
            var primaryFormulaY = primaryFormula?.OriginY ?? -1d;
            var numberY = numberGlyph?.OriginY ?? -1d;
            var sameLine = numberRendered
                && primaryFormulaY >= 0d
                && Math.Abs(primaryFormulaY - numberY) <= 1d;
            var rightAligned = numberRendered
                && numberGlyph!.OriginX >= 450d;
            var formulaRenderedOnce = equalsGlyphs.Length == 1;
            var result = new StyleSeparatorRefinementResult
            {
                Name = scenario.Name,
                Display = math.Type == Word.WdOMathType.wdOMathDisplay
                    && mathParagraphs.Length == 1,
                MathOnly = mathOnly,
                NumberRendered = numberRendered,
                SameLine = sameLine,
                RightAligned = rightAligned,
                FormulaRenderedOnce = formulaRenderedOnce,
                MathParagraphs = mathParagraphs.Length,
                MathNodes = mathNodes.Length,
                FormulaGlyphGroups = equalsGlyphs.Length,
                PrimaryFormulaX = primaryFormula?.OriginX ?? -1d,
                PrimaryFormulaY = primaryFormulaY,
                NumberX = numberGlyph?.OriginX ?? -1d,
                NumberY = numberY,
            };
            var glyphSummary = string.Join(
                " | ",
                glyphs.Select(glyph =>
                    $"'{EscapeStyleSeparatorGlyphText(glyph.Text)}'@{glyph.OriginX:0.###},{glyph.OriginY:0.###}"));
            Console.WriteLine(
                $"  refinement {scenario.Name}: type={math.Type}, jc={math.Justification}, liveText='{EscapeStyleSeparatorGlyphText(liveMathText)}', mathOnly={result.MathOnly}, rendered={result.NumberRendered}, sameLine={result.SameLine}, right={result.RightAligned}, formulaCopies={result.FormulaGlyphGroups}, formula={result.PrimaryFormulaX:0.###},{result.PrimaryFormulaY:0.###}, number={result.NumberX:0.###},{result.NumberY:0.###}, candidate={result.Candidate}, glyphs={glyphSummary}.");
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
            Release(separatorPoint);
            Release(selection);
            Release(rightTab);
            Release(centerTab);
            Release(tabs);
            Release(paragraphFormat);
            Release(paragraph);
            Release(paragraphs);
            Release(trimRange);
            Release(mathRange);
            Release(math);
            Release(addedMaths);
            Release(addedMathRange);
            Release(numberFont);
            Release(numberRange);
            Release(formulaRange);
            Release(content);
            Release(document);
            ForceComCleanup();
        }
    }
}

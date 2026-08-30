using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class StyleSeparatorRenderScenario
    {
        public string Name { get; set; } = string.Empty;
        public bool UseMath { get; set; }
        public bool DisplayMath { get; set; }
        public bool LeadingTab { get; set; }
        public bool RightTabOnFirstParagraph { get; set; }
        public bool RightTabOnSecondParagraph { get; set; }
        public Word.WdParagraphAlignment FirstAlignment { get; set; }
        public Word.WdParagraphAlignment SecondAlignment { get; set; }
    }

    private sealed class StyleSeparatorFixedGlyph
    {
        public string Text { get; set; } = string.Empty;
        public double OriginX { get; set; }
        public double OriginY { get; set; }
        public double EmSize { get; set; }
    }

    private static void RunWordOmmlStyleSeparatorRenderMatrixAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator render matrix refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);

        var scenarios = new[]
        {
            new StyleSeparatorRenderScenario
            {
                Name = "plain-no-tab",
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "plain-tab-second",
                LeadingTab = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "plain-tab-first",
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "plain-tab-both",
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "inline-no-tab",
                UseMath = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "inline-tab-second",
                UseMath = true,
                LeadingTab = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "inline-tab-first",
                UseMath = true,
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "inline-tab-both",
                UseMath = true,
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-no-tab-left",
                UseMath = true,
                DisplayMath = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-tab-second-center",
                UseMath = true,
                DisplayMath = true,
                LeadingTab = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-tab-first-center",
                UseMath = true,
                DisplayMath = true,
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-tab-both-center",
                UseMath = true,
                DisplayMath = true,
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-tab-second-left",
                UseMath = true,
                DisplayMath = true,
                LeadingTab = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-tab-first-left",
                UseMath = true,
                DisplayMath = true,
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-tab-both-left",
                UseMath = true,
                DisplayMath = true,
                LeadingTab = true,
                RightTabOnFirstParagraph = true,
                RightTabOnSecondParagraph = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
            },
            new StyleSeparatorRenderScenario
            {
                Name = "display-no-tab-second-right",
                UseMath = true,
                DisplayMath = true,
                FirstAlignment = Word.WdParagraphAlignment.wdAlignParagraphLeft,
                SecondAlignment = Word.WdParagraphAlignment.wdAlignParagraphRight,
            },
        };

        Word.Application? application = null;
        try
        {
            application = CreateWordApplication(visible: false);
            var plainControlRendered = false;
            var displayCandidates = new List<string>();
            foreach (var scenario in scenarios)
            {
                var result = RunStyleSeparatorRenderScenario(
                    application,
                    artifactRoot,
                    scenario);
                if (scenario.Name == "plain-no-tab")
                    plainControlRendered = result.NumberRendered;
                if (scenario.DisplayMath
                    && result.NumberRendered
                    && (!scenario.LeadingTab || result.NumberOriginX >= 450d))
                    displayCandidates.Add(scenario.Name);
            }

            AssertTrue(plainControlRendered,
                "The plain Style Separator control did not render its second paragraph in XPS; the fixed-layout probe is not trustworthy.");
            Console.WriteLine(
                "Style Separator render matrix display candidates: "
                + (displayCandidates.Count == 0
                    ? "<none>"
                    : string.Join(", ", displayCandidates)));
            Console.WriteLine(
                "Word OMML Style Separator render matrix acceptance completed; see per-scenario DOCX/XPS/XML artifacts and glyph coordinates.");
        }
        finally
        {
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static (bool NumberRendered, double NumberOriginX) RunStyleSeparatorRenderScenario(
        Word.Application application,
        string artifactRoot,
        StyleSeparatorRenderScenario scenario)
    {
        var docxPath = Path.Combine(artifactRoot, scenario.Name + ".docx");
        var xpsPath = Path.Combine(artifactRoot, scenario.Name + ".xps");
        var pdfPath = Path.Combine(artifactRoot, scenario.Name + ".pdf");
        var xmlPath = Path.Combine(artifactRoot, scenario.Name + "-wordopenxml.xml");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(xmlPath),
            scenario.Name + ": artifacts already exist; use a fresh artifact root.");

        Word.Document? document = null;
        Word.Range? content = null;
        Word.Range? firstSource = null;
        Word.Range? addedRange = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? firstParagraph = null;
        Word.Paragraph? secondParagraph = null;
        Word.ParagraphFormat? firstFormat = null;
        Word.ParagraphFormat? secondFormat = null;
        Word.TabStops? firstTabs = null;
        Word.TabStops? secondTabs = null;
        Word.TabStop? firstTab = null;
        Word.TabStop? secondTab = null;
        Word.Range? numberRange = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Range? separator = null;
        Word.Selection? selection = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? setup = null;
        try
        {
            document = application.Documents.Add(Visible: false);
            document.Activate();
            var firstText = scenario.UseMath ? "x=1" : "LEFT";
            content = document.Content;
            content.Text = firstText
                + "\r"
                + (scenario.LeadingTab ? "\t" : string.Empty)
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

            if (scenario.UseMath)
            {
                firstSource = document.Range(0, firstText.Length);
                addedRange = document.OMaths.Add(firstSource);
                addedMaths = addedRange.OMaths;
                AssertEqual(1, addedMaths.Count,
                    scenario.Name + ": Word did not create one OMath.");
                math = addedMaths[1];
                math.BuildUp();
                math.Type = scenario.DisplayMath
                    ? Word.WdOMathType.wdOMathDisplay
                    : Word.WdOMathType.wdOMathInline;
                mathRange = math.Range;
                mathRange.Font.Size = 14f;
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

            paragraphs = document.Paragraphs;
            AssertEqual(2, paragraphs.Count,
                scenario.Name + ": scenario did not start with two paragraphs.");
            firstParagraph = paragraphs[1];
            secondParagraph = paragraphs[2];
            firstFormat = firstParagraph.Format;
            secondFormat = secondParagraph.Format;
            ConfigureStyleSeparatorMatrixParagraph(
                firstFormat,
                scenario.FirstAlignment);
            ConfigureStyleSeparatorMatrixParagraph(
                secondFormat,
                scenario.SecondAlignment);
            firstTabs = firstFormat.TabStops;
            secondTabs = secondFormat.TabStops;
            firstTabs.ClearAll();
            secondTabs.ClearAll();
            if (scenario.RightTabOnFirstParagraph)
            {
                firstTab = firstTabs.Add(
                    StyleSeparatorRightTabPosition,
                    Word.WdTabAlignment.wdAlignTabRight,
                    Word.WdTabLeader.wdTabLeaderSpaces);
            }
            if (scenario.RightTabOnSecondParagraph)
            {
                secondTab = secondTabs.Add(
                    StyleSeparatorRightTabPosition,
                    Word.WdTabAlignment.wdAlignTabRight,
                    Word.WdTabLeader.wdTabLeaderSpaces);
            }

            separator = document.Range(
                firstParagraph.Range.End - 1,
                firstParagraph.Range.End - 1);
            separator.Select();
            selection = application.Selection;
            selection.InsertStyleSeparator();
            document.Repaginate();

            AssertEqual(2, document.Paragraphs.Count,
                scenario.Name + ": Style Separator did not preserve two logical paragraphs.");
            var openXml = document.Content.WordOpenXML ?? string.Empty;
            AssertTrue(openXml.IndexOf("specVanish", StringComparison.Ordinal) >= 0,
                scenario.Name + ": Style Separator did not create w:specVanish.");
            File.WriteAllText(xmlPath, openXml, new UTF8Encoding(false));
            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            ExportStyleSeparatorFixedLayout(document, xpsPath, scenario.Name);
            ExportStyleSeparatorPdfLayout(document, pdfPath, scenario.Name);

            var glyphs = ReadStyleSeparatorFixedGlyphs(xpsPath, scenario.Name);
            var renderedText = string.Concat(glyphs.Select(glyph => glyph.Text));
            var leftParenthesis = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf('(') >= 0);
            var rightParenthesis = glyphs.FirstOrDefault(glyph =>
                glyph.Text.IndexOf(')') >= 0);
            var numberRendered = leftParenthesis is not null
                && rightParenthesis is not null;
            var numberOriginX = leftParenthesis?.OriginX ?? -1d;
            var mathType = math is null ? "n/a" : math.Type.ToString();
            var glyphSummary = string.Join(
                " | ",
                glyphs.Select(glyph =>
                    $"'{EscapeStyleSeparatorGlyphText(glyph.Text)}'@{glyph.OriginX:0.###},{glyph.OriginY:0.###}/{glyph.EmSize:0.###}"));
            Console.WriteLine(
                $"  matrix {scenario.Name}: mathType={mathType}, numberRendered={numberRendered}, numberX={numberOriginX:0.###}, text='{EscapeStyleSeparatorGlyphText(renderedText)}', glyphs={glyphSummary}.");
            return (numberRendered, numberOriginX);
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(separator);
            Release(selection);
            Release(numberFont);
            Release(numberRange);
            Release(secondTab);
            Release(firstTab);
            Release(secondTabs);
            Release(firstTabs);
            Release(secondFormat);
            Release(firstFormat);
            Release(secondParagraph);
            Release(firstParagraph);
            Release(paragraphs);
            Release(mathRange);
            Release(math);
            Release(addedMaths);
            Release(addedRange);
            Release(firstSource);
            Release(setup);
            Release(section);
            Release(sections);
            Release(content);
            Release(document);
            ForceComCleanup();
        }
    }

    private static void ExportStyleSeparatorPdfLayout(
        Word.Document document,
        string path,
        string context)
    {
        AssertTrue(!File.Exists(path),
            context + ": PDF artifact already exists; use a fresh artifact root.");
        document.ExportAsFixedFormat(
            path,
            Word.WdExportFormat.wdExportFormatPDF,
            OpenAfterExport: false,
            OptimizeFor: Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
            Range: Word.WdExportRange.wdExportAllDocument,
            From: 1,
            To: 1,
            Item: Word.WdExportItem.wdExportDocumentContent,
            IncludeDocProps: true,
            KeepIRM: true,
            CreateBookmarks: Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks,
            DocStructureTags: true,
            BitmapMissingFonts: true,
            UseISO19005_1: false);
        var info = new FileInfo(path);
        AssertTrue(info.Exists && info.Length > 0,
            context + ": Word did not create a non-empty PDF snapshot.");
    }

    private static void RunWordOmmlStyleSeparatorVisibleProbeAcceptance(
        string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The visible Style Separator probe refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var docxPath = Path.Combine(artifactRoot, "visible-style-separator.docx");
        var xpsPath = Path.Combine(artifactRoot, "visible-style-separator.xps");
        var pdfPath = Path.Combine(artifactRoot, "visible-style-separator.pdf");
        var pngPath = Path.Combine(artifactRoot, "visible-style-separator-window.png");
        AssertTrue(!File.Exists(docxPath)
                   && !File.Exists(xpsPath)
                   && !File.Exists(pdfPath)
                   && !File.Exists(pngPath),
            "Visible Style Separator artifacts already exist; use a fresh artifact root.");

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Bookmark? formulaBookmark = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? formulaRange = null;
        Word.Range? numberRange = null;
        Word.Paragraphs? formulaParagraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.ParagraphFormat? formulaFormat = null;
        Word.TabStops? formulaTabs = null;
        Word.TabStop? formulaTab = null;
        Word.Window? window = null;
        Word.View? view = null;
        Word.Zoom? zoom = null;
        try
        {
            application = CreateWordApplication(visible: true);
            document = application.Documents.Add(Visible: true);
            document.Activate();
            ConfigureStyleSeparatorPrototypeDocument(application, document);

            formulaBookmark = document.Bookmarks[StyleSeparatorFormulaBookmark];
            formulaRange = formulaBookmark.Range;
            formulaParagraphs = formulaRange.Paragraphs;
            formulaParagraph = formulaParagraphs[1];
            formulaFormat = formulaParagraph.Format;
            formulaFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            formulaTabs = formulaFormat.TabStops;
            formulaTabs.ClearAll();
            formulaTab = formulaTabs.Add(
                StyleSeparatorRightTabPosition,
                Word.WdTabAlignment.wdAlignTabRight,
                Word.WdTabLeader.wdTabLeaderSpaces);

            InsertStyleSeparatorAtFormulaBoundary(application, document);
            document.SaveAs2(
                docxPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            document.Repaginate();

            Release(formulaRange); formulaRange = null;
            Release(formulaBookmark); formulaBookmark = null;
            formulaBookmark = document.Bookmarks[StyleSeparatorFormulaBookmark];
            numberBookmark = document.Bookmarks[StyleSeparatorNumberBookmark];
            formulaRange = formulaBookmark.Range;
            numberRange = numberBookmark.Range;

            window = document.ActiveWindow;
            window.WindowState = Word.WdWindowState.wdWindowStateNormal;
            window.Left = 40;
            window.Top = 40;
            window.Width = 1100;
            window.Height = 780;
            view = window.View;
            view.Type = Word.WdViewType.wdPrintView;
            zoom = view.Zoom;
            zoom.Percentage = 125;
            window.Activate();
            object scrollStart = true;
            window.ScrollIntoView(formulaRange, ref scrollStart);
            System.Windows.Forms.Application.DoEvents();
            Thread.Sleep(750);

            window.GetPoint(
                out var formulaLeft,
                out var formulaTop,
                out var formulaWidth,
                out var formulaHeight,
                formulaRange);
            var numberPointAvailable = true;
            var numberLeft = -1;
            var numberTop = -1;
            var numberWidth = -1;
            var numberHeight = -1;
            string? numberPointError = null;
            try
            {
                window.GetPoint(
                    out numberLeft,
                    out numberTop,
                    out numberWidth,
                    out numberHeight,
                    numberRange);
            }
            catch (Exception error)
            {
                numberPointAvailable = false;
                numberPointError = error.GetType().Name + ":" + error.Message;
            }

            AssertTrue(GetWindowRect(
                    new IntPtr(window.Hwnd),
                    out var rectangle),
                "Visible Style Separator probe could not read the isolated Word window rectangle.");
            var captureWidth = Math.Max(1, rectangle.Right - rectangle.Left);
            var captureHeight = Math.Max(1, rectangle.Bottom - rectangle.Top);
            using (var bitmap = new System.Drawing.Bitmap(
                       captureWidth,
                       captureHeight,
                       System.Drawing.Imaging.PixelFormat.Format32bppArgb))
            using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(
                    rectangle.Left,
                    rectangle.Top,
                    0,
                    0,
                    new System.Drawing.Size(captureWidth, captureHeight),
                    System.Drawing.CopyPixelOperation.SourceCopy);
                bitmap.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
            }

            ExportStyleSeparatorFixedLayout(
                document,
                xpsPath,
                "visible Style Separator probe");
            ExportStyleSeparatorPdfLayout(
                document,
                pdfPath,
                "visible Style Separator probe");
            var glyphs = ReadStyleSeparatorFixedGlyphs(
                xpsPath,
                "visible Style Separator probe");
            var fixedNumberRendered = glyphs.Any(glyph =>
                glyph.Text.IndexOf('(') >= 0)
                && glyphs.Any(glyph => glyph.Text.IndexOf(')') >= 0);

            Console.WriteLine(
                $"Visible Style Separator probe: formulaRect={formulaLeft},{formulaTop},{formulaWidth},{formulaHeight}; numberPointAvailable={numberPointAvailable}; numberRect={numberLeft},{numberTop},{numberWidth},{numberHeight}; numberPointError='{numberPointError ?? string.Empty}'; formulaY={formulaRange.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)}; numberY={numberRange.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage)}; fixedNumberRendered={fixedNumberRendered}; screenshot='{pngPath}'.");
            Console.WriteLine(
                "Word OMML Style Separator visible probe completed in a separately created Word instance; no existing document was opened.");
        }
        finally
        {
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(zoom);
            Release(view);
            Release(window);
            Release(formulaTab);
            Release(formulaTabs);
            Release(formulaFormat);
            Release(formulaParagraph);
            Release(formulaParagraphs);
            Release(numberRange);
            Release(formulaRange);
            Release(numberBookmark);
            Release(formulaBookmark);
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static void ConfigureStyleSeparatorMatrixParagraph(
        Word.ParagraphFormat format,
        Word.WdParagraphAlignment alignment)
    {
        format.Alignment = alignment;
        format.SpaceBefore = 0f;
        format.SpaceAfter = 0f;
        format.LeftIndent = 0f;
        format.RightIndent = 0f;
        format.FirstLineIndent = 0f;
    }

    private static IReadOnlyList<StyleSeparatorFixedGlyph> ReadStyleSeparatorFixedGlyphs(
        string xpsPath,
        string context)
    {
        using var archive = ZipFile.OpenRead(xpsPath);
        var pageEntries = archive.Entries
            .Where(entry => entry.FullName.EndsWith(
                ".fpage",
                StringComparison.OrdinalIgnoreCase))
            .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
            .ToArray();
        AssertTrue(pageEntries.Length > 0,
            context + ": XPS contains no FixedPage.");

        XNamespace xps = "http://schemas.microsoft.com/xps/2005/06";
        var glyphs = new List<StyleSeparatorFixedGlyph>();
        foreach (var pageEntry in pageEntries)
        {
            using var stream = pageEntry.Open();
            using var reader = new StreamReader(
                stream,
                Encoding.Unicode,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 4096,
                leaveOpen: false);
            var page = XDocument.Parse(
                reader.ReadToEnd(),
                LoadOptions.PreserveWhitespace);
            foreach (var glyph in page.Descendants(xps + "Glyphs"))
            {
                glyphs.Add(new StyleSeparatorFixedGlyph
                {
                    Text = (string?)glyph.Attribute("UnicodeString")
                        ?? string.Empty,
                    OriginX = ParseStyleSeparatorFixedDouble(
                        glyph.Attribute("OriginX"),
                        context + " OriginX"),
                    OriginY = ParseStyleSeparatorFixedDouble(
                        glyph.Attribute("OriginY"),
                        context + " OriginY"),
                    EmSize = ParseStyleSeparatorFixedDouble(
                        glyph.Attribute("FontRenderingEmSize"),
                        context + " FontRenderingEmSize"),
                });
            }
        }
        return glyphs;
    }

    private static double ParseStyleSeparatorFixedDouble(
        XAttribute? attribute,
        string context)
    {
        AssertTrue(attribute is not null,
            context + ": required XPS coordinate is missing.");
        AssertTrue(double.TryParse(
                attribute!.Value,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var value),
            context + $": invalid XPS coordinate '{attribute.Value}'.");
        return value;
    }

    private static string EscapeStyleSeparatorGlyphText(string text)
    {
        return (text ?? string.Empty)
            .Replace("\r", "\\r")
            .Replace("\n", "\\n")
            .Replace("\t", "\\t")
            .Replace("'", "\\'");
    }
}

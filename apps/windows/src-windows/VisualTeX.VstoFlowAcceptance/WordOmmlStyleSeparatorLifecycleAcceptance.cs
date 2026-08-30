using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private const string StyleSeparatorFormulaBookmark = "VTStyleSepFormula";
    private const string StyleSeparatorNumberBookmark = "VTStyleSepNumber";
    private const string StyleSeparatorFormulaText = "x=1";
    private const string StyleSeparatorNumberText = "(1)";
    private const string StyleSeparatorNumberFont = "Times New Roman";
    private const float StyleSeparatorNumberFontSize = 11f;
    private const float StyleSeparatorRightTabPosition = 415.3f;

    private static void RunWordOmmlStyleSeparatorLifecycleAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord,
            "The Style Separator prototype refuses to attach to a user's active Word instance.");
        Directory.CreateDirectory(artifactRoot);
        var savedPath = Path.Combine(
            artifactRoot,
            "word-omml-style-separator-lifecycle.docx");
        AssertTrue(!File.Exists(savedPath),
            "The Style Separator lifecycle artifact already exists; use a fresh artifact root.");

        Word.Application? application = null;
        Word.Document? document = null;
        var fixedLayoutFailures = new List<string>();
        try
        {
            application = CreateWordApplication(visible: false);
            Console.WriteLine(
                $"Style Separator Word runtime: version={application.Version}, build={application.Build}.");
            document = application.Documents.Add(Visible: false);
            document.Activate();
            ConfigureStyleSeparatorPrototypeDocument(application, document);
            InsertStyleSeparatorAtFormulaBoundary(application, document);

            AssertStyleSeparatorLifecyclePhase(
                application,
                document,
                artifactRoot,
                "01-after-insert",
                fixedLayoutFailures);

            UpdateStyleSeparatorPrototypeLikeF9(document);
            AssertStyleSeparatorLifecyclePhase(
                application,
                document,
                artifactRoot,
                "02-after-f9",
                fixedLayoutFailures);

            document.SaveAs2(
                savedPath,
                Word.WdSaveFormat.wdFormatXMLDocument,
                AddToRecentFiles: false);
            AssertTrue(File.Exists(savedPath),
                "Word did not create the Style Separator lifecycle DOCX.");
            AssertStyleSeparatorLifecyclePhase(
                application,
                document,
                artifactRoot,
                "03-after-save",
                fixedLayoutFailures);

            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                savedPath,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false,
                OpenAndRepair: false);
            document.Activate();
            AssertStyleSeparatorLifecyclePhase(
                application,
                document,
                artifactRoot,
                "04-after-reopen",
                fixedLayoutFailures);
            document.Save();

            AssertEqual(
                0,
                fixedLayoutFailures.Count,
                "Style Separator lifecycle is structurally stable but visually invalid: "
                + string.Join("; ", fixedLayoutFailures));
            Console.WriteLine(
                "Word OMML Style Separator lifecycle acceptance passed structurally and in fixed layout.");
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

    private static void ConfigureStyleSeparatorPrototypeDocument(
        Word.Application application,
        Word.Document document)
    {
        Word.Range? content = null;
        Word.Range? formulaSource = null;
        Word.Range? addedMathRange = null;
        Word.OMaths? maths = null;
        Word.OMaths? addedMaths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Range? numberRange = null;
        Word.Bookmarks? bookmarks = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Paragraph? numberParagraph = null;
        Word.ParagraphFormat? formulaFormat = null;
        Word.ParagraphFormat? numberFormat = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Microsoft.Office.Interop.Word.Font? mathFont = null;
        try
        {
            content = document.Content;
            content.Text = StyleSeparatorFormulaText
                + "\r\t"
                + StyleSeparatorNumberText;

            sections = document.Sections;
            AssertEqual(1, sections.Count,
                "The Style Separator prototype must start with one Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            pageSetup.PageWidth = 595.3f;
            pageSetup.PageHeight = 841.9f;
            pageSetup.LeftMargin = 90f;
            pageSetup.RightMargin = 90f;
            pageSetup.TopMargin = 72f;
            pageSetup.BottomMargin = 72f;
            AssertNear(
                StyleSeparatorRightTabPosition,
                pageSetup.PageWidth - pageSetup.LeftMargin - pageSetup.RightMargin,
                0.6f,
                "The prototype page geometry does not expose the required 415.3pt text width.");

            formulaSource = document.Range(0, StyleSeparatorFormulaText.Length);
            maths = document.OMaths;
            addedMathRange = maths.Add(formulaSource);
            addedMaths = addedMathRange.OMaths;
            AssertEqual(1, addedMaths.Count,
                "Word did not create exactly one OMath for the Style Separator prototype.");
            math = addedMaths[1];
            math.BuildUp();
            math.Type = Word.WdOMathType.wdOMathDisplay;
            mathRange = math.Range;
            mathFont = mathRange.Font;
            mathFont.Size = 14f;

            numberRange = FindUniqueStyleSeparatorTextRange(
                document,
                StyleSeparatorNumberText,
                "prototype number");
            numberFont = numberRange.Font;
            numberFont.Name = StyleSeparatorNumberFont;
            numberFont.NameAscii = StyleSeparatorNumberFont;
            numberFont.NameOther = StyleSeparatorNumberFont;
            numberFont.Size = StyleSeparatorNumberFontSize;
            numberFont.Bold = 0;
            numberFont.Italic = 0;
            numberFont.Position = 0;

            bookmarks = document.Bookmarks;
            bookmarks.Add(StyleSeparatorFormulaBookmark, mathRange);
            bookmarks.Add(StyleSeparatorNumberBookmark, numberRange);

            paragraphs = document.Paragraphs;
            AssertEqual(2, paragraphs.Count,
                "The Style Separator prototype did not materialize exactly two logical paragraphs before joining.");
            formulaParagraph = paragraphs[1];
            numberParagraph = paragraphs[2];

            formulaFormat = formulaParagraph.Format;
            formulaFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            formulaFormat.SpaceBefore = 0f;
            formulaFormat.SpaceAfter = 0f;
            formulaFormat.LeftIndent = 0f;
            formulaFormat.RightIndent = 0f;
            formulaFormat.FirstLineIndent = 0f;

            numberFormat = numberParagraph.Format;
            numberFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            numberFormat.SpaceBefore = 0f;
            numberFormat.SpaceAfter = 0f;
            numberFormat.LeftIndent = 0f;
            numberFormat.RightIndent = 0f;
            numberFormat.FirstLineIndent = 0f;
            tabStops = numberFormat.TabStops;
            tabStops.ClearAll();
            tabStop = tabStops.Add(
                StyleSeparatorRightTabPosition,
                Word.WdTabAlignment.wdAlignTabRight,
                Word.WdTabLeader.wdTabLeaderSpaces);

            document.Repaginate();
            Console.WriteLine(
                $"  Style Separator prototype prepared: paragraphs={document.Paragraphs.Count}, formulaType={math.Type}, rightTab={tabStop.Position:0.###}pt.");
        }
        finally
        {
            Release(mathFont);
            Release(numberFont);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(tabStop);
            Release(tabStops);
            Release(numberFormat);
            Release(formulaFormat);
            Release(numberParagraph);
            Release(formulaParagraph);
            Release(paragraphs);
            Release(bookmarks);
            Release(numberRange);
            Release(mathRange);
            Release(math);
            Release(addedMaths);
            Release(maths);
            Release(addedMathRange);
            Release(formulaSource);
            Release(content);
        }
    }

    private static void InsertStyleSeparatorAtFormulaBoundary(
        Word.Application application,
        Word.Document document)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Range? formulaRange = null;
        Word.Paragraphs? paragraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Range? paragraphRange = null;
        Word.Selection? selection = null;
        try
        {
            formulaBookmark = document.Bookmarks[StyleSeparatorFormulaBookmark];
            formulaRange = formulaBookmark.Range;
            paragraphs = formulaRange.Paragraphs;
            AssertEqual(1, paragraphs.Count,
                "The formula bookmark must occupy exactly one paragraph before Style Separator insertion.");
            formulaParagraph = paragraphs[1];
            paragraphRange = formulaParagraph.Range;
            var separatorPosition = paragraphRange.End - 1;
            AssertTrue(separatorPosition >= formulaRange.End,
                "The formula paragraph mark could not be isolated for Style Separator insertion.");

            selection = application.Selection;
            selection.SetRange(separatorPosition, separatorPosition);
            selection.InsertStyleSeparator();
            document.Repaginate();
            Console.WriteLine(
                $"  InsertStyleSeparator completed at document position {separatorPosition}; paragraphs={document.Paragraphs.Count}.");
        }
        finally
        {
            Release(selection);
            Release(paragraphRange);
            Release(formulaParagraph);
            Release(paragraphs);
            Release(formulaRange);
            Release(formulaBookmark);
        }
    }

    private static void UpdateStyleSeparatorPrototypeLikeF9(Word.Document document)
    {
        Word.Range? content = null;
        Word.Fields? fields = null;
        try
        {
            content = document.Content;
            fields = content.Fields;
            _ = fields.Update();
            document.Repaginate();
            Console.WriteLine(
                $"  F9-equivalent Fields.Update completed; fields={fields.Count}.");
        }
        finally
        {
            Release(fields);
            Release(content);
        }
    }

    private static void AssertStyleSeparatorLifecyclePhase(
        Word.Application application,
        Word.Document document,
        string artifactRoot,
        string phase,
        ICollection<string> fixedLayoutFailures)
    {
        Word.Bookmark? formulaBookmark = null;
        Word.Bookmark? numberBookmark = null;
        Word.Range? formulaBookmarkRange = null;
        Word.Range? numberRange = null;
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? mathRange = null;
        Word.Paragraphs? formulaParagraphs = null;
        Word.Paragraphs? numberParagraphs = null;
        Word.Paragraph? formulaParagraph = null;
        Word.Paragraph? numberParagraph = null;
        Word.Range? formulaParagraphRange = null;
        Word.Range? numberParagraphRange = null;
        Word.ParagraphFormat? numberFormat = null;
        Word.TabStops? tabStops = null;
        Word.TabStop? tabStop = null;
        Microsoft.Office.Interop.Word.Font? numberFont = null;
        Word.Sections? sections = null;
        Word.Section? section = null;
        Word.PageSetup? pageSetup = null;
        Word.Range? content = null;
        try
        {
            AssertTrue(document.Bookmarks.Exists(StyleSeparatorFormulaBookmark),
                phase + ": formula bookmark is missing.");
            AssertTrue(document.Bookmarks.Exists(StyleSeparatorNumberBookmark),
                phase + ": number bookmark is missing.");
            AssertEqual(0, document.Tables.Count,
                phase + ": Style Separator prototype unexpectedly contains a table.");
            AssertEqual(0, document.Shapes.Count,
                phase + ": Style Separator prototype unexpectedly contains a Shape/TextBox.");
            AssertEqual(0, document.InlineShapes.Count,
                phase + ": Style Separator prototype unexpectedly contains an InlineShape/OLE object.");

            formulaBookmark = document.Bookmarks[StyleSeparatorFormulaBookmark];
            numberBookmark = document.Bookmarks[StyleSeparatorNumberBookmark];
            formulaBookmarkRange = formulaBookmark.Range;
            numberRange = numberBookmark.Range;
            maths = formulaBookmarkRange.OMaths;
            AssertEqual(1, maths.Count,
                phase + ": formula bookmark does not contain exactly one OMath.");
            math = maths[1];
            mathRange = math.Range;
            AssertEqual(Word.WdOMathType.wdOMathDisplay, math.Type,
                phase + ": formula degraded from genuine wdOMathDisplay.");
            AssertTrue((mathRange.Text ?? string.Empty).IndexOf(
                    StyleSeparatorNumberText,
                    StringComparison.Ordinal) < 0,
                phase + ": number leaked into OMath.Range.");
            AssertTrue((mathRange.Text ?? string.Empty).IndexOf('\t') < 0,
                phase + ": layout TAB leaked into OMath.Range.");
            AssertEqual(0, mathRange.Fields.Count,
                phase + ": a Word field leaked into OMath.Range.");
            AssertEqual(StyleSeparatorNumberText, numberRange.Text ?? string.Empty,
                phase + ": ordinary equation number text changed.");

            formulaParagraphs = formulaBookmarkRange.Paragraphs;
            numberParagraphs = numberRange.Paragraphs;
            AssertEqual(1, formulaParagraphs.Count,
                phase + ": formula bookmark spans multiple logical paragraphs.");
            AssertEqual(1, numberParagraphs.Count,
                phase + ": number bookmark spans multiple logical paragraphs.");
            formulaParagraph = formulaParagraphs[1];
            numberParagraph = numberParagraphs[1];
            formulaParagraphRange = formulaParagraph.Range;
            numberParagraphRange = numberParagraph.Range;
            AssertTrue(formulaParagraphRange.Start != numberParagraphRange.Start,
                phase + ": Style Separator collapsed formula and number into one logical paragraph.");
            AssertEqual(2, document.Paragraphs.Count,
                phase + ": the document no longer exposes exactly two logical paragraphs.");
            AssertTrue((formulaParagraphRange.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal),
                phase + ": formula logical paragraph lost its paragraph mark.");
            AssertTrue((numberParagraphRange.Text ?? string.Empty).StartsWith("\t", StringComparison.Ordinal),
                phase + ": number logical paragraph no longer begins with a real TAB character.");
            AssertTrue((numberParagraphRange.Text ?? string.Empty).EndsWith("\r", StringComparison.Ordinal),
                phase + ": number logical paragraph lost its final ordinary paragraph mark.");
            AssertEqual(numberParagraphRange.Start + 1, numberRange.Start,
                phase + ": number is not immediately after the leading right TAB.");
            AssertEqual(numberParagraphRange.End - 1, numberRange.End,
                phase + ": number is not immediately before the final paragraph mark.");

            numberFont = numberRange.Font;
            AssertEqual(StyleSeparatorNumberFont, numberFont.Name,
                phase + ": number font changed.");
            AssertNear(StyleSeparatorNumberFontSize, numberFont.Size, 0.1f,
                phase + ": number font size changed.");
            AssertEqual(0, numberFont.Bold,
                phase + ": number unexpectedly became bold.");
            AssertEqual(0, numberFont.Italic,
                phase + ": number unexpectedly became italic.");
            AssertEqual(0, numberFont.Position,
                phase + ": number baseline position changed.");

            numberFormat = numberParagraph.Format;
            tabStops = numberFormat.TabStops;
            var sawRequiredRightTab = false;
            for (var index = 1; index <= tabStops.Count; index++)
            {
                Release(tabStop); tabStop = tabStops[index];
                if (tabStop.Alignment == Word.WdTabAlignment.wdAlignTabRight
                    && Math.Abs(tabStop.Position - StyleSeparatorRightTabPosition) <= 0.6f)
                    sawRequiredRightTab = true;
            }
            AssertTrue(sawRequiredRightTab,
                phase + ": number paragraph lost the genuine 415.3pt right TabStop.");

            content = document.Content;
            var openXml = content.WordOpenXML ?? string.Empty;
            File.WriteAllText(
                Path.Combine(artifactRoot, phase + "-wordopenxml.xml"),
                openXml,
                new UTF8Encoding(false));
            var parsed = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
            XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace officeMath = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var mathParagraphs = parsed.Descendants(officeMath + "oMathPara").ToArray();
            var mathNodes = parsed.Descendants(officeMath + "oMath").ToArray();
            AssertEqual(1, mathParagraphs.Length,
                phase + ": m:oMathPara was not preserved exactly once.");
            AssertEqual(1, mathNodes.Length,
                phase + ": m:oMath was not preserved exactly once.");
            var formulaXmlParagraph = mathParagraphs[0].Ancestors(word + "p").FirstOrDefault()
                ?? throw new InvalidDataException(phase + ": m:oMathPara has no w:p owner.");
            var numberXmlParagraph = parsed.Descendants(word + "p")
                .SingleOrDefault(candidate =>
                    !ReferenceEquals(candidate, formulaXmlParagraph)
                    && string.Equals(
                        string.Concat(candidate.Descendants(word + "t").Select(text => text.Value)),
                        StyleSeparatorNumberText,
                        StringComparison.Ordinal))
                ?? throw new InvalidDataException(
                    phase + ": ordinary number paragraph could not be found outside OMML.");
            AssertTrue(formulaXmlParagraph.Descendants(word + "specVanish").Any(),
                phase + ": formula paragraph mark lost w:specVanish Style Separator semantics.");
            AssertTrue(!numberXmlParagraph.Descendants(word + "specVanish").Any(),
                phase + ": final number paragraph mark incorrectly became a Style Separator.");
            AssertTrue(numberXmlParagraph.Descendants(word + "tab").Any(element =>
                    element.Attribute(word + "val") is null),
                phase + ": OOXML no longer contains a real w:tab character before the number.");
            AssertTrue(numberXmlParagraph.Descendants(word + "tab").Any(element =>
                    string.Equals(
                        (string?)element.Attribute(word + "val"),
                        "right",
                        StringComparison.Ordinal)
                    && string.Equals(
                        (string?)element.Attribute(word + "pos"),
                        "8306",
                        StringComparison.Ordinal)),
                phase + ": OOXML no longer contains the exact 415.3pt right TabStop.");
            AssertTrue(!numberXmlParagraph.Descendants(officeMath + "oMath").Any(),
                phase + ": number paragraph unexpectedly contains OMML.");
            AssertTrue(!formulaXmlParagraph.Descendants(word + "tab").Any(),
                phase + ": formula paragraph unexpectedly contains a layout TAB.");
            AssertTrue(!formulaXmlParagraph.Descendants(word + "t").Any(),
                phase + ": formula paragraph contains ordinary visible Word text outside math.");
            AssertTrue(!mathNodes[0].Descendants(word + "fldChar").Any(),
                phase + ": field controls leaked into m:oMath.");
            AssertTrue(!mathNodes[0].Descendants(word + "tab").Any(),
                phase + ": a Word TAB leaked into m:oMath.");
            AssertEqual(
                StyleSeparatorFormulaText,
                string.Concat(mathNodes[0].Descendants(officeMath + "t").Select(text => text.Value)),
                phase + ": m:oMath contains anything other than the formula body.");
            AssertEqual(1, parsed.Descendants(word + "specVanish").Count(),
                phase + ": lifecycle no longer contains exactly one Style Separator marker.");
            AssertTrue(!parsed.Descendants(word + "br").Any()
                       && !parsed.Descendants(word + "cr").Any(),
                phase + ": an explicit visible line break was introduced.");

            document.Repaginate();
            var formulaY = Convert.ToSingle(formulaBookmarkRange.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            var numberY = Convert.ToSingle(numberRange.get_Information(
                Word.WdInformation.wdVerticalPositionRelativeToPage));
            AssertTrue(formulaY >= 0f && numberY >= 0f,
                phase + ": Word did not expose formula/number vertical page coordinates.");
            AssertNear(formulaY, numberY, 1.5f,
                phase + ": formula and number are not on the same visual line.");

            sections = numberRange.Sections;
            AssertTrue(sections.Count > 0,
                phase + ": number has no Word section.");
            section = sections[1];
            pageSetup = section.PageSetup;
            var expectedRight = pageSetup.PageWidth
                - pageSetup.LeftMargin
                - pageSetup.RightMargin;
            AssertNear(StyleSeparatorRightTabPosition, expectedRight, 0.6f,
                phase + ": writable text width changed from 415.3pt.");
            var expectedRightOnPage = pageSetup.PageWidth - pageSetup.RightMargin;
            var numberXDiagnostic = Convert.ToSingle(numberRange.get_Information(
                Word.WdInformation.wdHorizontalPositionRelativeToPage));
            var fixedLayoutPath = Path.Combine(artifactRoot, phase + "-layout.xps");
            var pdfLayoutPath = Path.Combine(artifactRoot, phase + "-layout.pdf");
            ExportStyleSeparatorFixedLayout(document, fixedLayoutPath, phase);
            ExportStyleSeparatorPdfLayout(document, pdfLayoutPath, phase);
            var fixedGlyphs = ReadStyleSeparatorFixedGlyphs(fixedLayoutPath, phase);
            var fixedNumberRendered = fixedGlyphs.Any(glyph =>
                    glyph.Text.IndexOf('(') >= 0)
                && fixedGlyphs.Any(glyph => glyph.Text.IndexOf(')') >= 0);
            var fixedFormulaCopies = fixedGlyphs.Count(glyph =>
                glyph.Text.IndexOf('=') >= 0);
            if (!fixedNumberRendered)
            {
                fixedLayoutFailures.Add(
                    phase + ": Word XPS fixed layout omitted the ordinary equation number even though Range.Information reported the same Y coordinate");
            }
            if (fixedFormulaCopies != 1)
            {
                fixedLayoutFailures.Add(
                    phase + $": fixed layout rendered {fixedFormulaCopies} formula copies instead of one");
            }

            Console.WriteLine(
                $"  {phase}: type={math.Type}, paragraphs={document.Paragraphs.Count}, specVanish={parsed.Descendants(word + "specVanish").Count()}, rangeFormulaY={formulaY:0.###}pt, rangeNumberY={numberY:0.###}pt, COM-numberX={numberXDiagnostic:0.###}pt, rightEdge={expectedRightOnPage:0.###}pt, rightTab={expectedRight:0.###}pt, fixedNumberRendered={fixedNumberRendered}, fixedFormulaCopies={fixedFormulaCopies}, XPS='{fixedLayoutPath}', PDF='{pdfLayoutPath}', font='{numberFont.Name}' {numberFont.Size:0.##}pt.");
        }
        finally
        {
            Release(content);
            Release(pageSetup);
            Release(section);
            Release(sections);
            Release(numberFont);
            Release(tabStop);
            Release(tabStops);
            Release(numberFormat);
            Release(numberParagraphRange);
            Release(formulaParagraphRange);
            Release(numberParagraph);
            Release(formulaParagraph);
            Release(numberParagraphs);
            Release(formulaParagraphs);
            Release(mathRange);
            Release(math);
            Release(maths);
            Release(numberRange);
            Release(formulaBookmarkRange);
            Release(numberBookmark);
            Release(formulaBookmark);
        }
    }

    private static void ExportStyleSeparatorFixedLayout(
        Word.Document document,
        string path,
        string context)
    {
        AssertTrue(!File.Exists(path),
            context + ": fixed-layout artifact already exists; use a fresh artifact root.");
        document.ExportAsFixedFormat(
            path,
            Word.WdExportFormat.wdExportFormatXPS,
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
            context + ": Word did not create a non-empty XPS fixed-layout snapshot.");
    }

    private static Word.Range FindUniqueStyleSeparatorTextRange(
        Word.Document document,
        string text,
        string context)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var documentText = content.Text ?? string.Empty;
            var first = documentText.IndexOf(text, StringComparison.Ordinal);
            var last = documentText.LastIndexOf(text, StringComparison.Ordinal);
            if (first < 0 || first != last)
                throw new InvalidDataException(
                    $"Could not locate one unique {context} '{text}' in '{documentText.Replace("\r", "\\r").Replace("\t", "\\t")}'.");
            return document.Range(
                content.Start + first,
                content.Start + first + text.Length);
        }
        finally { Release(content); }
    }
}

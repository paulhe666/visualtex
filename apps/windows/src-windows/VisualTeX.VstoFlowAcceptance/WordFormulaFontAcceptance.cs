using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private sealed class WordOmmlLayoutMetric
    {
        internal string Context { get; set; } = string.Empty;
        internal Word.WdOMathType Type { get; set; }
        internal float FontSizePt { get; set; }
        internal int WidthPx { get; set; }
        internal int HeightPx { get; set; }
        internal int EqualsWidthPx { get; set; }
        internal int PlusMinusWidthPx { get; set; }
        internal int RadicalSpanWidthPx { get; set; }
        internal int FractionVerticalGapPx { get; set; }
        internal int FractionCount { get; set; }
        internal int RadicalCount { get; set; }
        internal int NormalTextRunCount { get; set; }

        public override string ToString() =>
            $"{Context}: type={Type}, size={FontSizePt:0.##}pt, "
            + $"box={WidthPx}x{HeightPx}px, equals={EqualsWidthPx}px, "
            + $"plusMinus={PlusMinusWidthPx}px, radicalSpan={RadicalSpanWidthPx}px, "
            + $"fractionVerticalGap={FractionVerticalGapPx}px, "
            + $"fractions={FractionCount}, radicals={RadicalCount}, "
            + $"normalTextRuns={NormalTextRunCount}";
    }

    private sealed class WordRangePixelBox
    {
        internal int Left { get; set; }
        internal int Top { get; set; }
        internal int Width { get; set; }
        internal int Height { get; set; }
    }

    private static void RunWordFormulaFontAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var artifactPath = Path.Combine(artifactRoot, "word-formula-fonts.docx");
        var metricPath = Path.Combine(artifactRoot, "word-formula-fonts.metrics.txt");
        const string nativeToken = "VT_NATIVE_QUADRATIC_FONT_CONTROL";
        const string nativeLinear = "x=(-b±√(b^2-4ac))/(2a)";

        Word.Application? application = null;
        Word.Document? document = null;
        Word.Range? nativeRange = null;
        Word.Range? insertion = null;
        Word.Bookmark? visualBookmark = null;
        Word.Range? visualRange = null;
        Word.Bookmark? textBookmark = null;
        Word.Range? textRange = null;
        Microsoft.Office.Interop.Word.Font? nativeFont = null;
        var metrics = new List<WordOmmlLayoutMetric>();
        WordOmmlLayoutMetric? nativeCambria = null;
        WordOmmlLayoutMetric? nativeLatinModern = null;
        WordOmmlLayoutMetric? visualLatinModern = null;
        WordOmmlLayoutMetric? reopenedNative = null;
        WordOmmlLayoutMetric? reopenedVisual = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Add(Visible: false);
            document.Activate();

            // Establish an explicit Word-native control before VisualTeX changes the
            // document-wide Office Math font. This proves the product semantics on
            // pre-existing OMath, rather than merely checking the newly inserted one.
            document.OMathFontName = "Cambria Math";
            document.Content.Text = nativeToken + "\r";
            nativeRange = InsertPureNativeOmml(document, nativeToken, nativeLinear);
            nativeFont = nativeRange.Font;
            nativeFont.Position = 0;
            nativeFont.Size = 14f;
            try { nativeFont.SizeBi = 14f; } catch { }
            Release(nativeFont); nativeFont = null;
            AssertNativeQuadraticOmml(nativeRange, "Word-native Cambria control");
            nativeCambria = ReadWordOmmlLayoutMetric(
                application,
                document,
                nativeRange,
                "Word-native before document mathFont change");
            metrics.Add(nativeCambria);

            var service = new WordFormulaService(application);
            var visualFormulaId = Guid.NewGuid().ToString("D");
            insertion = AppendAcceptanceParagraph(document);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var visualSession = CreateFontAcceptanceSession(
                document,
                visualFormulaId,
                insertion.Start,
                insertion.End,
                @"x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}");
            service.InsertOmml(visualSession, QuadraticFormulaMathMl());
            Release(insertion); insertion = null;

            AssertDocumentOmmlMathFont(
                document,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                "VisualTeX quadratic insertion");
            AssertTrue(WordOfficeMathFontLoader.IsLoaded,
                "Word did not load VisualTeX's verified Latin Modern Math payload.");
            AssertTrue(File.Exists(WordOfficeMathFontLoader.LoadedPath),
                "The verified Latin Modern Math payload disappeared while Word was using it.");

            AssertEqual(2, document.OMaths.Count,
                "The Word control and VisualTeX quadratic did not remain as two native OMath objects.");
            Release(nativeRange); nativeRange = null;
            nativeRange = GetDocumentOmmlRange(document, 1);
            visualBookmark = WordOmmlFormulaStore.FindByFormulaId(document, visualFormulaId)
                ?? throw new InvalidDataException("VisualTeX quadratic lost its OMML bookmark.");
            visualRange = WordOmmlFormulaStore.GetEquationRange(visualBookmark);

            AssertNativeQuadraticOmml(nativeRange,
                "pre-existing Word OMath after document mathFont change");
            AssertNativeQuadraticOmml(visualRange, "VisualTeX quadratic");
            nativeLatinModern = ReadWordOmmlLayoutMetric(
                application,
                document,
                nativeRange,
                "Word-native after VisualTeX set Latin Modern Math");
            visualLatinModern = ReadWordOmmlLayoutMetric(
                application,
                document,
                visualRange,
                "VisualTeX Latin Modern Math quadratic");
            metrics.Add(nativeLatinModern);
            metrics.Add(visualLatinModern);
            AssertComparableNativeMathLayout(
                nativeLatinModern,
                visualLatinModern,
                "Word-native versus VisualTeX quadratic");

            // A document-level math font must reflow OMath that already existed
            // before VisualTeX inserted anything. Word can defer final pixel layout
            // until pagination/save, so the save/reopen check below accepts either
            // the live or reopened reflow as the observable geometry change.
            var liveControlReflowed = HasObservableMathLayoutChange(
                nativeCambria,
                nativeLatinModern);

            var textFormulaId = Guid.NewGuid().ToString("D");
            insertion = AppendAcceptanceParagraph(document);
            application.Selection.SetRange(insertion.Start, insertion.End);
            var textSession = CreateFontAcceptanceSession(
                document,
                textFormulaId,
                insertion.Start,
                insertion.End,
                @"x+\text{中文}+2");
            service.InsertOmml(
                textSession,
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">"
                + "<mi>x</mi><mo>+</mo><mtext>中文</mtext><mo>+</mo><mn>2</mn></math>");
            Release(insertion); insertion = null;
            textBookmark = WordOmmlFormulaStore.FindByFormulaId(document, textFormulaId)
                ?? throw new InvalidDataException("VisualTeX text formula lost its OMML bookmark.");
            textRange = WordOmmlFormulaStore.GetEquationRange(textBookmark);
            AssertTrueTextOmmlTypography(
                textRange,
                ResolveExpectedAcceptanceChineseFont("system"),
                "VisualTeX \\text{中文}");
            AssertDocumentOmmlMathFont(
                document,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                "mixed text formula");

            document.SaveAs2(artifactPath, Word.WdSaveFormat.wdFormatXMLDocument);
            AssertSavedDocumentOmmlMathFont(
                artifactPath,
                WordOfficeMathFontLoader.LatinModernMathFamily);
            AssertSavedDocumentOrdinaryMathIsNative(artifactPath);

            Release(textRange); textRange = null;
            Release(textBookmark); textBookmark = null;
            Release(visualRange); visualRange = null;
            Release(visualBookmark); visualBookmark = null;
            Release(nativeRange); nativeRange = null;
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(document); document = null;

            document = application.Documents.Open(
                artifactPath,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: false);
            document.Activate();
            AssertDocumentOmmlMathFont(
                document,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                "save/reopen");
            AssertEqual(3, document.OMaths.Count,
                "Save/reopen did not retain the Word control and both VisualTeX OMath formulas.");

            nativeRange = GetDocumentOmmlRange(document, 1);
            visualBookmark = WordOmmlFormulaStore.FindByFormulaId(document, visualFormulaId)
                ?? throw new InvalidDataException("Save/reopen lost the VisualTeX quadratic bookmark.");
            visualRange = WordOmmlFormulaStore.GetEquationRange(visualBookmark);
            textBookmark = WordOmmlFormulaStore.FindByFormulaId(document, textFormulaId)
                ?? throw new InvalidDataException("Save/reopen lost the VisualTeX text bookmark.");
            textRange = WordOmmlFormulaStore.GetEquationRange(textBookmark);

            AssertNativeQuadraticOmml(nativeRange, "save/reopen Word-native quadratic");
            AssertNativeQuadraticOmml(visualRange, "save/reopen VisualTeX quadratic");
            AssertTrueTextOmmlTypography(
                textRange,
                ResolveExpectedAcceptanceChineseFont("system"),
                "save/reopen VisualTeX \\text{中文}");

            reopenedNative = ReadWordOmmlLayoutMetric(
                application,
                document,
                nativeRange,
                "save/reopen Word-native Latin Modern quadratic");
            reopenedVisual = ReadWordOmmlLayoutMetric(
                application,
                document,
                visualRange,
                "save/reopen VisualTeX Latin Modern quadratic");
            metrics.Add(reopenedNative);
            metrics.Add(reopenedVisual);
            AssertStableMathLayoutAfterReopen(
                nativeLatinModern,
                reopenedNative,
                "Word-native quadratic");
            AssertStableMathLayoutAfterReopen(
                visualLatinModern,
                reopenedVisual,
                "VisualTeX quadratic");
            AssertComparableNativeMathLayout(
                reopenedNative,
                reopenedVisual,
                "save/reopen Word-native versus VisualTeX quadratic");
            AssertTrue(
                liveControlReflowed || HasObservableMathLayoutChange(nativeCambria, reopenedNative),
                "Changing Document.OMathFontName from Cambria Math to Latin Modern Math did not visibly reflow the pre-existing Word OMath control.");

            File.WriteAllText(
                metricPath,
                string.Join(Environment.NewLine, metrics.Select(item => item.ToString()))
                + Environment.NewLine,
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            Console.WriteLine(
                "Word document-level OMML font acceptance passed: pre-existing and VisualTeX OMath share Latin Modern Math, ordinary math stayed native, true Chinese text stayed text, and save/reopen geometry remained stable.");
            Console.WriteLine($"Word OMML font metrics: {metricPath}");
        }
        finally
        {
            Release(nativeFont);
            Release(textRange);
            Release(textBookmark);
            Release(visualRange);
            Release(visualBookmark);
            Release(insertion);
            Release(nativeRange);
            try { document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }

    private static OfficeSessionDocument CreateFontAcceptanceSession(
        Word.Document document,
        string formulaId,
        int start,
        int end,
        string latex)
    {
        return new OfficeSessionDocument
        {
            Id = Guid.NewGuid().ToString("D"),
            Mode = "create",
            Host = "word",
            FormulaId = formulaId,
            SourceDocumentId = document.FullName,
            SourceObjectId = WordRangeReference(start, end),
            Title = "VisualTeX document-level Office Math font acceptance",
            CodeFormat = "latex",
            DisplayMode = "inline",
            ObjectMode = FormulaOleContract.WordOmmlMode,
            Numbered = false,
            FontSizePt = 14,
            Lines = new List<FormulaLine>
            {
                new()
                {
                    Id = Guid.NewGuid().ToString("D"),
                    Latex = latex,
                },
            },
            ExportResult = new OfficeExportDocument
            {
                FormulaLetterFont = "katex",
                FormulaChineseFont = "system",
            },
        };
    }

    private static Word.Range AppendAcceptanceParagraph(Word.Document document)
    {
        Word.Range? end = null;
        Word.Range? result = null;
        try
        {
            var position = Math.Max(document.Content.Start, document.Content.End - 1);
            end = document.Range(position, position);
            end.InsertParagraphAfter();
            position = Math.Max(document.Content.Start, document.Content.End - 1);
            result = document.Range(position, position);
            var duplicate = result;
            result = null;
            return duplicate;
        }
        finally
        {
            Release(result);
            Release(end);
        }
    }

    private static Word.Range GetDocumentOmmlRange(Word.Document document, int index)
    {
        Word.OMaths? maths = null;
        Word.OMath? math = null;
        Word.Range? range = null;
        try
        {
            maths = document.OMaths;
            if (index < 1 || index > maths.Count)
                throw new InvalidDataException(
                    $"Word OMath index {index} is outside the document's {maths.Count} equations.");
            math = maths[index];
            range = math.Range.Duplicate;
            var duplicate = range;
            range = null;
            return duplicate;
        }
        finally
        {
            Release(range);
            Release(math);
            Release(maths);
        }
    }

    private static void AssertDocumentOmmlMathFont(
        Word.Document document,
        string expected,
        string context)
    {
        var actual = document.OMathFontName ?? string.Empty;
        AssertTrue(
            string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase),
            $"{context}: document OMathFontName is '{actual}', expected '{expected}'.");
    }

    private static void AssertNativeQuadraticOmml(
        Word.Range equationRange,
        string context)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var document = XDocument.Parse(
            equationRange.WordOpenXML ?? string.Empty,
            LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        AssertTrue(document.Descendants(math + "oMath").Any(),
            context + ": the formula is not a native Word OMath.");
        AssertTrue(document.Descendants(math + "f").Any(),
            context + ": the quadratic formula lost its native fraction structure.");
        AssertTrue(document.Descendants(math + "rad").Any(),
            context + ": the quadratic formula lost its native radical structure.");

        var visibleRuns = document
            .Descendants(math + "r")
            .Where(run => run.Elements(math + "t")
                .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            .ToArray();
        AssertTrue(visibleRuns.Length > 0,
            context + ": the formula contains no visible native math runs.");
        var normalTextRuns = visibleRuns
            .Where(run => run.Element(math + "rPr")?.Element(math + "nor") is not null)
            .ToArray();
        AssertEqual(0, normalTextRuns.Length,
            context + ": ordinary quadratic variables/operators were flattened to m:nor normal text.");

        var semantic = NormalizeOmmlSemanticText(string.Concat(
            visibleRuns.SelectMany(run => run.Elements(math + "t"))
                .Select(text => text.Value)));
        foreach (var token in new[] { "x", "=", "b", "±", "2", "4", "a", "c" })
        {
            AssertTrue(semantic.IndexOf(token, StringComparison.Ordinal) >= 0,
                context + $": quadratic semantic token '{token}' is missing. Semantic='{semantic}'.");
        }
    }

    private static void AssertTrueTextOmmlTypography(
        Word.Range equationRange,
        string expectedChineseFont,
        string context)
    {
        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var document = XDocument.Parse(
            equationRange.WordOpenXML ?? string.Empty,
            LoadOptions.PreserveWhitespace);
        var word = (XNamespace)WordNamespace;
        var math = (XNamespace)MathNamespace;
        var visibleRuns = document
            .Descendants(math + "r")
            .Where(run => run.Elements(math + "t")
                .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            .ToArray();
        var chineseRuns = visibleRuns
            .Where(run => NormalizeOmmlSemanticText(string.Concat(
                    run.Elements(math + "t").Select(text => text.Value)))
                .IndexOf("中文", StringComparison.Ordinal) >= 0)
            .ToArray();
        AssertTrue(chineseRuns.Length > 0,
            context + ": the true Chinese text run is missing.");
        AssertTrue(chineseRuns.All(run =>
                run.Element(math + "rPr")?.Element(math + "nor") is not null),
            context + ": true Chinese text is not marked as Word normal text.");
        AssertTrue(chineseRuns.Any(run =>
        {
            var fonts = run.Element(word + "rPr")?.Element(word + "rFonts");
            return FontNameMatches(
                (string?)fonts?.Attribute(word + "eastAsia"),
                expectedChineseFont);
        }), context + $": true Chinese text does not use {expectedChineseFont} as its East-Asian text font.");

        var ordinaryRuns = visibleRuns.Except(chineseRuns).ToArray();
        AssertTrue(ordinaryRuns.All(run =>
                run.Element(math + "rPr")?.Element(math + "nor") is null),
            context + ": ordinary variables, digits or operators were flattened to m:nor alongside true text.");
        AssertOmmlCharacterFont(
            equationRange,
            "x",
            WordOfficeMathFontLoader.LatinModernMathFamily,
            context + " mathematical variable");
        AssertOmmlCharacterFont(
            equationRange,
            "2",
            WordOfficeMathFontLoader.LatinModernMathFamily,
            context + " mathematical digit");
    }

    private static WordOmmlLayoutMetric ReadWordOmmlLayoutMetric(
        Word.Application application,
        Word.Document document,
        Word.Range equationRange,
        string context,
        string? semanticOmmlOverride = null,
        bool measureVisibleCharacterInk = false)
    {
        Word.OMaths? maths = null;
        Word.OMath? equation = null;
        Microsoft.Office.Interop.Word.Font? font = null;
        Word.Window? window = null;
        try
        {
            document.Repaginate();
            Thread.Sleep(80);
            maths = equationRange.OMaths;
            AssertEqual(1, maths.Count,
                context + ": metric range does not contain exactly one OMath.");
            equation = maths[1];
            font = equationRange.Font;
            window = document.ActiveWindow;
            object scrollStart = true;
            window.ScrollIntoView(equationRange, ref scrollStart);
            document.Repaginate();
            Thread.Sleep(80);
            var box = measureVisibleCharacterInk
                ? ReadVisibleMathInkBox(document, window, equationRange, context + " formula ink")
                : ReadWordRangePixelBox(window, equationRange, context + " formula box");

            const string MathNamespace =
                "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var openXml = XDocument.Parse(
                semanticOmmlOverride
                    ?? equationRange.WordOpenXML
                    ?? string.Empty,
                LoadOptions.PreserveWhitespace);
            var math = (XNamespace)MathNamespace;
            var metric = new WordOmmlLayoutMetric
            {
                Context = context,
                Type = equation.Type,
                FontSizePt = font.Size,
                WidthPx = box.Width,
                HeightPx = box.Height,
                EqualsWidthPx = ReadSemanticCharacterWidth(
                    document,
                    window,
                    equationRange,
                    "=",
                    useLast: false),
                PlusMinusWidthPx = ReadSemanticCharacterWidth(
                    document,
                    window,
                    equationRange,
                    "±",
                    useLast: false),
                RadicalSpanWidthPx = ReadRadicalSemanticSpanWidth(
                    document,
                    window,
                    equationRange),
                FractionVerticalGapPx = ReadFractionVerticalGap(
                    document,
                    window,
                    equationRange),
                FractionCount = openXml.Descendants(math + "f").Count(),
                RadicalCount = openXml.Descendants(math + "rad").Count(),
                NormalTextRunCount = openXml
                    .Descendants(math + "r")
                    .Count(run => run.Element(math + "rPr")?.Element(math + "nor") is not null),
            };
            AssertTrue(metric.FontSizePt > 0f,
                context + ": Word reported an invalid OMath font size.");
            AssertTrue(metric.WidthPx > 0 && metric.HeightPx > 0,
                context + ": Word reported an empty physical OMath box.");
            AssertTrue(metric.EqualsWidthPx > 0,
                context + ": Word could not measure the native relation operator.");
            AssertTrue(metric.PlusMinusWidthPx > 0,
                context + ": Word could not measure the native plus/minus operator.");
            AssertEqual(0, metric.NormalTextRunCount,
                context + ": the quadratic formula contains m:nor normal-text runs.");
            Console.WriteLine("  " + metric);
            return metric;
        }
        finally
        {
            Release(window);
            Release(font);
            Release(equation);
            Release(maths);
        }
    }

    private static WordRangePixelBox ReadVisibleMathInkBox(
        Word.Document document,
        Word.Window window,
        Word.Range range,
        string context)
    {
        var text = range.Text ?? string.Empty;
        var left = int.MaxValue;
        var top = int.MaxValue;
        var right = int.MinValue;
        var bottom = int.MinValue;
        var measured = 0;
        for (var offset = 0; offset < text.Length; offset++)
        {
            var character = text[offset];
            if (character is '\r' or '\n' or '\v' or '\a' or '\t'
                or '\u200b' or '\u200c' or '\u2060' or '\ufeff'
                || char.IsWhiteSpace(character))
                continue;
            Word.Range? characterRange = null;
            try
            {
                characterRange = document.Range(
                    range.Start + offset,
                    range.Start + offset + 1);
                window.GetPoint(
                    out var characterLeft,
                    out var characterTop,
                    out var characterWidth,
                    out var characterHeight,
                    characterRange);
                if (characterWidth <= 0 || characterHeight <= 0) continue;
                left = Math.Min(left, characterLeft);
                top = Math.Min(top, characterTop);
                right = Math.Max(right, characterLeft + characterWidth);
                bottom = Math.Max(bottom, characterTop + characterHeight);
                measured++;
            }
            catch
            {
                // Structural control positions inside a professional OMath can
                // reject GetPoint. Visible neighboring characters still define the
                // painted mathematical ink box.
            }
            finally { Release(characterRange); }
        }
        AssertTrue(measured > 0 && right > left && bottom > top,
            context + ": Word did not expose any visible mathematical glyph boxes.");
        return new WordRangePixelBox
        {
            Left = left,
            Top = top,
            Width = right - left,
            Height = bottom - top,
        };
    }

    private static WordRangePixelBox ReadWordRangePixelBox(
        Word.Window window,
        Word.Range range,
        string context)
    {
        window.GetPoint(
            out var left,
            out var top,
            out var width,
            out var height,
            range);
        AssertTrue(width > 0 && height > 0,
            context + $": Word returned invalid geometry {left},{top},{width},{height}.");
        return new WordRangePixelBox
        {
            Left = left,
            Top = top,
            Width = width,
            Height = height,
        };
    }

    private static int ReadSemanticCharacterWidth(
        Word.Document document,
        Word.Window window,
        Word.Range equationRange,
        string token,
        bool useLast)
    {
        var text = equationRange.Text ?? string.Empty;
        var index = useLast
            ? text.LastIndexOf(token, StringComparison.Ordinal)
            : text.IndexOf(token, StringComparison.Ordinal);
        if (index < 0 || index + token.Length > equationRange.End - equationRange.Start)
            return 0;
        Word.Range? tokenRange = null;
        try
        {
            tokenRange = document.Range(
                equationRange.Start + index,
                equationRange.Start + index + token.Length);
            return ReadWordRangePixelBox(window, tokenRange, "OMath token '" + token + "'").Width;
        }
        catch
        {
            return 0;
        }
        finally { Release(tokenRange); }
    }

    private static int ReadRadicalSemanticSpanWidth(
        Word.Document document,
        Word.Window window,
        Word.Range equationRange)
    {
        var text = equationRange.Text ?? string.Empty;
        var start = text.IndexOf("√", StringComparison.Ordinal);
        if (start < 0) return 0;
        var end = text.IndexOf("c", start, StringComparison.Ordinal);
        if (end < start) return 0;
        Word.Range? span = null;
        try
        {
            span = document.Range(
                equationRange.Start + start,
                equationRange.Start + end + 1);
            return ReadWordRangePixelBox(window, span, "OMath radical semantic span").Width;
        }
        catch
        {
            return 0;
        }
        finally { Release(span); }
    }

    private static int ReadFractionVerticalGap(
        Word.Document document,
        Word.Window window,
        Word.Range equationRange)
    {
        var text = equationRange.Text ?? string.Empty;
        var numeratorIndex = text.IndexOf("b", StringComparison.Ordinal);
        var denominatorIndex = text.LastIndexOf("2", StringComparison.Ordinal);
        if (numeratorIndex < 0 || denominatorIndex < 0 || numeratorIndex == denominatorIndex)
            return 0;
        Word.Range? numerator = null;
        Word.Range? denominator = null;
        try
        {
            numerator = document.Range(
                equationRange.Start + numeratorIndex,
                equationRange.Start + numeratorIndex + 1);
            denominator = document.Range(
                equationRange.Start + denominatorIndex,
                equationRange.Start + denominatorIndex + 1);
            var numeratorBox = ReadWordRangePixelBox(window, numerator, "OMath fraction numerator probe");
            var denominatorBox = ReadWordRangePixelBox(window, denominator, "OMath fraction denominator probe");
            return Math.Abs(denominatorBox.Top - numeratorBox.Top);
        }
        catch
        {
            return 0;
        }
        finally
        {
            Release(denominator);
            Release(numerator);
        }
    }

    private static void AssertComparableNativeMathLayout(
        WordOmmlLayoutMetric expected,
        WordOmmlLayoutMetric actual,
        string context)
    {
        AssertEqual(expected.Type, actual.Type,
            context + ": OMath types differ.");
        AssertNear(expected.FontSizePt, actual.FontSizePt, 0.25f,
            context + ": actual font sizes differ.");
        AssertMetricRatio(expected.WidthPx, actual.WidthPx, 0.72, 1.38,
            context + ": physical widths are not comparable.");
        AssertMetricRatio(expected.HeightPx, actual.HeightPx, 0.72, 1.38,
            context + ": physical heights are not comparable.");
        AssertMetricRatio(expected.EqualsWidthPx, actual.EqualsWidthPx, 0.50, 2.00,
            context + ": relation-operator spacing is not comparable.");
        AssertMetricRatio(expected.PlusMinusWidthPx, actual.PlusMinusWidthPx, 0.50, 2.00,
            context + ": plus/minus operator spacing is not comparable.");
        if (expected.RadicalSpanWidthPx > 0 && actual.RadicalSpanWidthPx > 0)
            AssertMetricRatio(expected.RadicalSpanWidthPx, actual.RadicalSpanWidthPx, 0.65, 1.55,
                context + ": radical span spacing is not comparable.");
        if (expected.FractionVerticalGapPx > 0 && actual.FractionVerticalGapPx > 0)
            AssertMetricRatio(expected.FractionVerticalGapPx, actual.FractionVerticalGapPx, 0.60, 1.70,
                context + ": numerator/denominator spacing is not comparable.");
        AssertEqual(expected.FractionCount, actual.FractionCount,
            context + ": native fraction structure count differs.");
        AssertEqual(expected.RadicalCount, actual.RadicalCount,
            context + ": native radical structure count differs.");
    }

    private static void AssertStableMathLayoutAfterReopen(
        WordOmmlLayoutMetric before,
        WordOmmlLayoutMetric after,
        string context)
    {
        AssertEqual(before.Type, after.Type,
            context + ": save/reopen changed the OMath type.");
        AssertNear(before.FontSizePt, after.FontSizePt, 0.25f,
            context + ": save/reopen changed the OMath font size.");
        AssertNear(before.WidthPx, after.WidthPx, 4f,
            context + ": save/reopen changed the physical OMath width.");
        AssertNear(before.HeightPx, after.HeightPx, 4f,
            context + ": save/reopen changed the physical OMath height.");
        AssertMetricRatio(before.EqualsWidthPx, after.EqualsWidthPx, 0.66, 1.50,
            context + ": save/reopen changed relation-operator spacing.");
        AssertMetricRatio(before.PlusMinusWidthPx, after.PlusMinusWidthPx, 0.66, 1.50,
            context + ": save/reopen changed plus/minus spacing.");
        AssertEqual(before.FractionCount, after.FractionCount,
            context + ": save/reopen changed the native fraction count.");
        AssertEqual(before.RadicalCount, after.RadicalCount,
            context + ": save/reopen changed the native radical count.");
    }

    private static void AssertMetricRatio(
        double expected,
        double actual,
        double minimum,
        double maximum,
        string message)
    {
        AssertTrue(expected > 0 && actual > 0,
            message + $" Invalid values expected={expected:0.###}, actual={actual:0.###}.");
        var ratio = actual / expected;
        AssertTrue(ratio >= minimum && ratio <= maximum,
            message + $" Ratio={ratio:0.###}, expected range={minimum:0.###}..{maximum:0.###}; values={expected:0.###}/{actual:0.###}.");
    }

    private static bool HasObservableMathLayoutChange(
        WordOmmlLayoutMetric before,
        WordOmmlLayoutMetric after) =>
        Math.Abs(before.WidthPx - after.WidthPx) >= 1
        || Math.Abs(before.HeightPx - after.HeightPx) >= 1
        || Math.Abs(before.EqualsWidthPx - after.EqualsWidthPx) >= 1
        || Math.Abs(before.PlusMinusWidthPx - after.PlusMinusWidthPx) >= 1
        || Math.Abs(before.RadicalSpanWidthPx - after.RadicalSpanWidthPx) >= 1
        || Math.Abs(before.FractionVerticalGapPx - after.FractionVerticalGapPx) >= 1;

    private static void AssertSavedDocumentOmmlMathFont(
        string documentPath,
        string expected)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        using var archive = ZipFile.OpenRead(documentPath);
        var settingsEntry = archive.GetEntry("word/settings.xml")
            ?? throw new InvalidDataException("Saved DOCX has no word/settings.xml.");
        using var stream = settingsEntry.Open();
        var settings = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        var actual = settings
            .Descendants(math + "mathFont")
            .Select(element =>
                (string?)element.Attribute(math + "val")
                ?? (string?)element.Attribute("val")
                ?? string.Empty)
            .FirstOrDefault()
            ?? string.Empty;
        AssertTrue(string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase),
            $"Saved DOCX mathFont is '{actual}', expected '{expected}'.");
    }

    private static void AssertSavedDocumentOrdinaryMathIsNative(string documentPath)
    {
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        using var archive = ZipFile.OpenRead(documentPath);
        var documentEntry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("Saved DOCX has no word/document.xml.");
        using var stream = documentEntry.Open();
        var wordDocument = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        var equations = wordDocument.Descendants(math + "oMath").ToArray();
        AssertTrue(equations.Length >= 3,
            "Saved DOCX did not retain all native OMath equations.");
        foreach (var equation in equations.Take(2))
        {
            var visibleRuns = equation
                .Descendants(math + "r")
                .Where(run => run.Elements(math + "t")
                    .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                .ToArray();
            AssertTrue(visibleRuns.All(run =>
                    run.Element(math + "rPr")?.Element(math + "nor") is null),
                "Saved quadratic OMath contains m:nor normal-text math runs.");
        }
    }

    private static string ResolveExpectedAcceptanceChineseFont(string? preference) =>
        (preference ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "songti" => "SimSun",
            "kaiti" => "KaiTi",
            "heiti" => "SimHei",
            _ => "Microsoft YaHei",
        };

    // Shared by the tab-numbering acceptance. For mathematical tokens the native
    // font is document-wide, so the assertion verifies OMathFontName plus the
    // absence of m:nor. For true text it verifies the run-level East-Asian face.
    private static void AssertOmmlCharacterFont(
        Word.Range equationRange,
        string expectedText,
        string expectedFont,
        string context)
    {
        const string WordNamespace =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string MathNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var openXml = equationRange.WordOpenXML ?? string.Empty;
        var document = XDocument.Parse(openXml, LoadOptions.PreserveWhitespace);
        var word = (XNamespace)WordNamespace;
        var math = (XNamespace)MathNamespace;
        var normalizedExpectedText = NormalizeOmmlSemanticText(expectedText);
        var matchingRuns = document
            .Descendants(math + "r")
            .Where(run => run.Elements(math + "t").Any(text =>
                NormalizeOmmlSemanticText(text.Value).IndexOf(
                    normalizedExpectedText,
                    StringComparison.Ordinal) >= 0))
            .ToArray();
        AssertTrue(matchingRuns.Length > 0,
            $"{context}: character '{expectedText}' was not found in native OMath. XML={openXml}");

        if (string.Equals(
                expectedFont,
                WordOfficeMathFontLoader.LatinModernMathFamily,
                StringComparison.OrdinalIgnoreCase)
            || string.Equals(expectedFont, "Cambria Math", StringComparison.OrdinalIgnoreCase)
            || string.Equals(expectedFont, "STIX Two Math", StringComparison.OrdinalIgnoreCase))
        {
            Word.Document? owner = null;
            try
            {
                owner = equationRange.Document;
                AssertDocumentOmmlMathFont(owner, expectedFont, context);
            }
            finally { Release(owner); }
            AssertTrue(matchingRuns.All(run =>
                    run.Element(math + "rPr")?.Element(math + "nor") is null),
                $"{context}: mathematical token '{expectedText}' was flattened to m:nor normal text.");
            return;
        }

        var matchedFont = matchingRuns.Any(run =>
        {
            var fonts = run.Element(word + "rPr")?.Element(word + "rFonts");
            return FontNameMatches(
                (string?)fonts?.Attribute(word + "eastAsia"),
                expectedFont);
        });
        AssertTrue(matchedFont,
            $"{context}: true-text token '{expectedText}' does not use {expectedFont} as its East-Asian font. XML={openXml}");
    }

    private static string NormalizeOmmlSemanticText(string? text)
    {
        var value = text ?? string.Empty;
        try { return value.Normalize(NormalizationForm.FormKC); }
        catch (ArgumentException) { return value; }
    }

    private static bool FontNameMatches(string? actual, string expected)
    {
        if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
            return true;
        if (string.Equals(expected, "SimSun", StringComparison.OrdinalIgnoreCase))
            return string.Equals(actual, "宋体", StringComparison.Ordinal)
                || string.Equals(actual, "新宋体", StringComparison.Ordinal)
                || string.Equals(actual, "NSimSun", StringComparison.OrdinalIgnoreCase);
        if (string.Equals(expected, "KaiTi", StringComparison.OrdinalIgnoreCase))
            return string.Equals(actual, "楷体", StringComparison.Ordinal);
        if (string.Equals(expected, "Microsoft YaHei", StringComparison.OrdinalIgnoreCase))
            return string.Equals(actual, "微软雅黑", StringComparison.Ordinal);
        if (string.Equals(expected, "SimHei", StringComparison.OrdinalIgnoreCase))
            return string.Equals(actual, "黑体", StringComparison.Ordinal);
        return false;
    }
}

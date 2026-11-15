using System.Xml.Linq;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Pure rules of PowerPoint LaTeX redraw and of turning formulas into LaTeX code:
// where formulas are in a text box's text, what a conversion replaces, the
// LaTeX code it writes and the Ribbon commands.
public sealed class PowerPointLatexRedrawTests
{
    private const string Ole = FormulaOleContract.NativeOleMode;
    private static readonly (int Start, int Length)[] NoEquations = Array.Empty<(int Start, int Length)>();

    // The 1-based Characters(start, length) text of a span.
    private static string Code(string text, PowerPointLatexSpan span) => text.Substring(span.Start - 1, span.Length);

    [Fact]
    public void FormulasAreFoundWithTheirPowerPointPositions()
    {
        const string text = "Let $x^2$ grow.\r$$\\frac{a}{b}$$\rAnd \\(y\\) too.";
        var spans = PowerPointLatexText.FindFormulas(text, NoEquations);
        Assert.Equal(3, spans.Count);
        Assert.Equal(5, spans[0].Start);
        Assert.Equal("$x^2$", Code(text, spans[0]));
        Assert.Equal("x^2", spans[0].Latex);
        Assert.False(spans[0].Display);
        Assert.Equal("$$\\frac{a}{b}$$", Code(text, spans[1]));
        Assert.True(spans[1].Display);
        Assert.Equal("block", spans[1].DisplayMode);
        Assert.Equal("\\(y\\)", Code(text, spans[2]));
        Assert.Equal("inline", spans[2].DisplayMode);
    }

    [Fact]
    public void ANativeEquationIsNeverReadAsLatex()
    {
        // What PowerPoint spells for the equation (positions 7-11) holds "$a$".
        const string text = "Area: x$a$y and $b$";
        var equation = new[] { (Start: 7, Length: 5) };
        var spans = PowerPointLatexText.FindFormulas(text, equation);
        var span = Assert.Single(spans);
        Assert.Equal("$b$", Code(text, span));
        Assert.Equal("b", span.Latex);

        // A formula cannot reach across an equation either.
        const string across = "$a EQ b$ and then $c$";
        var spansAcross = PowerPointLatexText.FindFormulas(across, new[] { (Start: 4, Length: 2) });
        Assert.Equal("$c$", Code(across, Assert.Single(spansAcross)));

        // Masking keeps every position.
        var scan = PowerPointLatexText.ScanText(text, equation);
        Assert.Equal(text.Length, scan.Length);
        Assert.Equal(new string(PowerPointLatexText.EquationMask, 5), scan.Substring(6, 5));
        Assert.Equal(text.Substring(11), scan.Substring(11));
    }

    [Fact]
    public void ASoftLineBreakIsReadAsANewLineWithoutMovingAnything()
    {
        // Inline math ends at a line break (Shift+Enter, U+000B).
        Assert.Empty(PowerPointLatexText.FindFormulas("$a\vb$", NoEquations));
        // Display math may span line breaks...
        const string display = "Then\v$$\va+b\v$$\vdone";
        var span = Assert.Single(PowerPointLatexText.FindFormulas(display, NoEquations));
        Assert.Equal("$$\va+b\v$$", Code(display, span));
        Assert.Equal("a+b", span.Latex);
        // ...but a blank line ends it, as in TeX.
        Assert.Empty(PowerPointLatexText.FindFormulas("$$a\v\vb$$", NoEquations));
        // Positions after a line break are PowerPoint's.
        const string later = "x\v\v$y$";
        Assert.Equal(4, Assert.Single(PowerPointLatexText.FindFormulas(later, NoEquations)).Start);
        Assert.Equal("x\n\n$y$", PowerPointLatexText.ScanText(later, null));
    }

    [Fact]
    public void PricesAndStrayDollarsAreNotFormulas()
    {
        Assert.Empty(PowerPointLatexText.FindFormulas("It costs $5 and $10.", NoEquations));
        Assert.False(PowerPointLatexText.MayContainLatex("plain text"));
        Assert.True(PowerPointLatexText.MayContainLatex("\\alpha"));
    }

    [Fact]
    public void ATextSelectionConvertsOnlyTheFormulasItHolds()
    {
        const string text = "$a$ and $b$";
        var spans = PowerPointLatexText.FindFormulas(text, NoEquations);
        Assert.True(PowerPointLatexText.Inside(spans[0], 1, 3));
        Assert.False(PowerPointLatexText.Inside(spans[0], 2, 5));
        Assert.True(PowerPointLatexText.Inside(spans[1], 5, 7));
        Assert.True(PowerPointLatexText.Overlaps(5, 3, 7, 1));
        Assert.False(PowerPointLatexText.Overlaps(5, 3, 8, 1));
    }

    [Fact]
    public void ADisplayFormulaTakesTheLineBreaksAndSpacesNextToIt()
    {
        const string text = "Text:  \v $$x^2$$ \vMore $y$ end";
        var spans = PowerPointLatexText.FindFormulas(text, NoEquations);
        var ranges = PowerPointLatexText.ReplacedRanges(text, NoEquations, spans);
        // From the first blank after "Text:" to the line break after the code.
        Assert.Equal((6, 13), ranges[0]);
        Assert.Equal("  \v $$x^2$$ \v", text.Substring(ranges[0].Start - 1, ranges[0].Length));
        // An inline formula replaces just its code.
        Assert.Equal((spans[1].Start, spans[1].Length), ranges[1]);
    }

    [Fact]
    public void NeighbouringDisplayRangesNeverOverlapAndNeverReachIntoAnEquation()
    {
        const string text = "$$a$$ \v $$b$$";
        var spans = PowerPointLatexText.FindFormulas(text, NoEquations);
        var ranges = PowerPointLatexText.ReplacedRanges(text, NoEquations, spans);
        Assert.Equal(2, ranges.Count);
        Assert.True(ranges[0].Start + ranges[0].Length <= ranges[1].Start);
        Assert.Equal(1, ranges[0].Start);
        Assert.Equal(text.Length + 1, ranges[1].Start + ranges[1].Length);

        // A blank that belongs to an equation stays with it.
        const string withEquation = "E $$x$$";
        var spansNext = PowerPointLatexText.FindFormulas(withEquation, new[] { (Start: 1, Length: 2) });
        var range = PowerPointLatexText.ReplacedRanges(withEquation, new[] { (Start: 1, Length: 2) }, spansNext).Single();
        Assert.Equal(3, range.Start);
    }

    [Theory]
    [InlineData("$$x^2$$", true)]
    [InlineData("  \\[ x \\]\r\v ", true)]
    [InlineData("\u200B$a$\u00A0", true)]
    [InlineData("Area: $x^2$", false)]
    [InlineData("$a$ $b$", false)]
    [InlineData("$a$ 1", false)]
    public void OnlyABoxHoldingJustOneFormulaBecomesAnObject(string text, bool single)
    {
        var spans = PowerPointLatexText.FindFormulas(text, NoEquations);
        Assert.Equal(single, PowerPointLatexText.IsSingleFormula(text, NoEquations, spans));
    }

    // MsoShapeType: msoAutoShape 1, msoFreeform 5, msoPicture 13, msoPlaceholder 14, msoTextBox 17.
    [Theory]
    [InlineData(17, true, true, true)]
    [InlineData(17, false, false, true)]
    [InlineData(14, true, false, true)]
    [InlineData(1, false, false, true)]
    [InlineData(1, true, false, false)]
    [InlineData(1, false, true, false)]
    [InlineData(5, false, false, false)]
    [InlineData(13, false, false, false)]
    public void OnlyATextBoxBecomesAFormulaObject(int shapeType, bool fillVisible, bool lineVisible, bool becomes)
    {
        Assert.Equal(becomes, PowerPointLatexText.CanBecomeObject(
            PowerPointLatexText.HostOf(shapeType), fillVisible, lineVisible));
    }

    [Fact]
    public void AShapesLatexLabelIsReportedWithTheWayToKeepIt()
    {
        var english = PowerPointConversionRules.SkipReasonText(PowerPointConversionSkipReason.LatexInShape, 3, english: true);
        Assert.StartsWith("3 LaTeX formula(s) are the text of a drawn shape", english, StringComparison.Ordinal);
        Assert.Contains("only a text box that holds just the formula becomes an object", english, StringComparison.Ordinal);
        Assert.Contains("LaTeX → OMML still works there", english, StringComparison.Ordinal);
        var chinese = PowerPointConversionRules.SkipReasonText(PowerPointConversionSkipReason.LatexInShape, 3, english: false);
        Assert.StartsWith("3 个 LaTeX 公式是图形", chinese, StringComparison.Ordinal);
        Assert.Contains("只有仅含该公式的文本框才会变为对象", chinese, StringComparison.Ordinal);
        Assert.Contains("“LaTeX → OMML”", chinese, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0f, false, 0f)]
    [InlineData(30f, false, 30f)]
    [InlineData(0f, true, 180f)]
    [InlineData(270f, true, 90f)]
    public void AFlippedTextBoxIsNeverMirroredAndUpsideDownWhenFlippedVertically(float rotation, bool flipVertical, float shown)
    {
        Assert.Equal(shown, PowerPointLatexText.TextRotation(rotation, flipVertical), 3);
    }

    [Fact]
    public void ACaretConvertsTheEquationItIsIn()
    {
        // "→ LaTeX code" with only a caret: the equation holding it (its first
        // character up to its last), else none, and then the whole box.
        var zones = new[] { (Start: 5, Length: 3), (Start: 12, Length: 2) };
        Assert.Equal(0, WpsMathText.ZoneAt(zones, 5, 0));
        Assert.Equal(0, WpsMathText.ZoneAt(zones, 7, 0));
        Assert.Equal(-1, WpsMathText.ZoneAt(zones, 8, 0));
        Assert.Equal(1, WpsMathText.ZoneAt(zones, 13, 0));
        Assert.Equal(-1, WpsMathText.ZoneAt(zones, 1, 0));
    }

    [Fact]
    public void ANativeEquationMakesTheBoxMixed()
    {
        const string text = "$$x$$EQ";
        var equations = new[] { (Start: 6, Length: 2) };
        var spans = PowerPointLatexText.FindFormulas(text, equations);
        Assert.Single(spans);
        Assert.False(PowerPointLatexText.IsSingleFormula(text, equations, spans));
    }

    [Theory]
    [InlineData("EQ", 1, 2, true)]
    [InlineData("Before\r  EQ \rAfter", 10, 2, true)]
    [InlineData("Let EQ be", 5, 2, false)]
    [InlineData("Line\vEQ", 6, 2, false)]
    [InlineData("EQ\vmore", 1, 2, false)]
    [InlineData("A\rEQEQ", 3, 2, false)]
    public void AnEquationAloneInItsParagraphIsDisplayCode(string text, int start, int length, bool alone)
    {
        Assert.Equal(alone, PowerPointLatexText.IsAloneInParagraph(text, start, length));
    }

    [Fact]
    public void LatexCodeIsWrittenTheWayWordWritesIt()
    {
        Assert.Equal("$x^2$", PowerPointLatexText.BuildSource(" x^2 ", display: false));
        Assert.Equal("$$\\frac{a}{b}$$", PowerPointLatexText.BuildSource("\\frac{a}{b}", display: true));
        // The equation tag stays with display code only.
        Assert.Equal("$$a=b\\tag{3}$$", PowerPointLatexText.BuildSource("a=b", display: true, equationTag: "3"));
        Assert.Equal("$a=b$", PowerPointLatexText.BuildSource("a=b\\tag{3}", display: false));
        // Lines of display code become PowerPoint line breaks, without empty
        // lines; inline code is one line.
        Assert.Equal("$$a\\\\\vb$$", PowerPointLatexText.BuildSource("a\\\\\r\n\r\n  b", display: true));
        // A formula's rows without their own environment become one aligned
        // block (gathered without alignment points).
        Assert.Equal("$$\\begin{aligned}x+y&=1 \\\\ x-y&=0\\end{aligned}$$",
            PowerPointLatexText.BuildSource("x+y&=1\nx-y&=0", display: true));
        Assert.Equal("$$\\begin{gathered}a=1 \\\\ b=2\\end{gathered}\\tag{2}$$",
            PowerPointLatexText.BuildSource("a=1\r\nb=2", display: true, equationTag: "2"));
        Assert.Equal("$$a\\&b$$", PowerPointLatexText.BuildSource("a\\&b", display: true));
        Assert.Equal("$a b$", PowerPointLatexText.BuildSource("a\n  b", display: false));
        Assert.Throws<InvalidDataException>(() => PowerPointLatexText.BuildSource("  ", display: true));
        Assert.Throws<InvalidDataException>(() => PowerPointLatexText.BuildSource("\\tag{1}", display: false));
    }

    [Theory]
    [InlineData("x^2", false)]
    [InlineData("\\frac{a}{b}", true)]
    [InlineData("\\begin{aligned}a&=b\\\\\nc&=d\\end{aligned}", true)]
    public void WrittenCodeIsFoundAgainByTheRedraw(string latex, bool display)
    {
        var source = PowerPointLatexText.BuildSource(latex, display);
        var text = "Before " + source + " after";
        var span = Assert.Single(PowerPointLatexText.FindFormulas(text, NoEquations));
        Assert.Equal(source, Code(text, span));
        Assert.Equal(display, span.Display);
        Assert.Equal(latex.Replace("\n", " ").Replace("\\\\ ", "\\\\"), span.Latex.Replace("\\\\ ", "\\\\"));
    }

    [Fact]
    public void WrittenRowsAreFoundAgainAsOneAlignedFormula()
    {
        var source = PowerPointLatexText.BuildSource("x+y&=1\nx-y&=0", display: true);
        var text = "Before " + source + " after";
        var span = Assert.Single(PowerPointLatexText.FindFormulas(text, NoEquations));
        Assert.Equal(source, Code(text, span));
        Assert.True(span.Display);
        Assert.Equal(
            "\\begin{aligned}x+y&=1\\\\x-y&=0\\end{aligned}",
            span.Latex.Replace(" \\\\ ", "\\\\"));
    }

    [Fact]
    public void ColoredCodeKeepsItsColorAsAnObject()
    {
        Assert.Equal("x", PowerPointLatexText.WithTextColor("x", "#000000"));
        Assert.Equal("x", PowerPointLatexText.WithTextColor("x", null));
        Assert.Equal("\\textcolor{#FFFFFF}{x^2}", PowerPointLatexText.WithTextColor("x^2", "#ffffff"));
    }

    [Fact]
    public void FinalRangesFollowTheChangesBeforeThem()
    {
        var final = PowerPointLatexText.FinalRanges(new[]
        {
            (Start: 3, Length: 2, NewLength: 7),
            (Start: 10, Length: 4, NewLength: 3),
            (Start: 20, Length: 1, NewLength: 5),
        });
        Assert.Equal(new[] { (3, 7), (15, 3), (24, 5) }, final.ToArray());
    }

    [Fact]
    public void CodeTakesTheFormatOfTheNearestOrdinaryText()
    {
        // Before the equation in its paragraph first, then after it.
        Assert.Equal(4, PowerPointLatexText.FormatDonor("Let EQ be", new[] { (5, 2) }, 5, 2));
        Assert.Equal(3, PowerPointLatexText.FormatDonor("EQab", new[] { (1, 2) }, 1, 2));
        // An equation alone in its paragraph takes the next paragraph's text.
        Assert.Equal(4, PowerPointLatexText.FormatDonor("EQ\rab", new[] { (1, 2) }, 1, 2));
        // Other equations are not text; an equation-only box has none.
        Assert.Equal(0, PowerPointLatexText.FormatDonor("EQFG", new[] { (1, 2), (3, 2) }, 1, 2));
        Assert.Equal(0, PowerPointLatexText.FormatDonor("EQ", new[] { (1, 2) }, 1, 2));
    }

    [Theory]
    [InlineData(24.0, 24f)]
    [InlineData(32.84, 32.8f)]
    [InlineData(0.2, 1f)]
    [InlineData(9000.0, 4000f)]
    [InlineData(0.0, 18f)]
    public void TextBoxFontSizeStaysInPowerPointsRange(double formula, float text)
    {
        Assert.Equal(text, PowerPointLatexText.TextFontSize(formula), 3);
    }

    [Theory]
    [InlineData("<a:normAutofit fontScale=\"62500\" lnSpcReduction=\"20000\"/>", 0.625)]
    [InlineData("<a:normAutofit fontScale=\"62.5%\"/>", 0.625)]
    [InlineData("<a:normAutofit/>", 1.0)]
    [InlineData("<a:spAutoFit/>", 1.0)]
    [InlineData("<a:normAutofit fontScale=\"abc\"/>", 1.0)]
    public void ShrunkTextIsDrawnAtItsAutofitScale(string autofit, double scale)
    {
        var shape = XElement.Parse(
            "<p:sp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
            + "<p:txBody><a:bodyPr>" + autofit + "</a:bodyPr><a:p><a:r><a:t>$x$</a:t></a:r></a:p></p:txBody></p:sp>");
        Assert.Equal(scale, PowerPointLatexText.AutofitFontScale(shape), 6);
        Assert.Equal(1.0, PowerPointLatexText.AutofitFontScale(null), 6);
    }

    [Fact]
    public void FormulasInGroupsAndTablesAreCountedFromTheSavedSlide()
    {
        var group = XElement.Parse(
            "<p:grpSp xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">"
            + "<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>Let $a$ and</a:t></a:r><a:br/><a:r><a:t>$$b$$</a:t></a:r></a:p></p:txBody></p:sp>"
            + "<p:sp><p:txBody><a:bodyPr/><a:p><a:r><a:t>$</a:t></a:r><a14:m><m:oMath><m:r><m:t>$x$</m:t></m:r></m:oMath></a14:m><a:r><a:t>$</a:t></a:r></a:p></p:txBody></p:sp>"
            + "<p:graphicFrame><a:graphic><a:graphicData><a:tbl><a:tr><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>\\(c\\)</a:t></a:r></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>"
            + "</p:grpSp>");
        // $a$, $$b$$ and \(c\); the equation's own text and the dollars around it are not formulas.
        Assert.Equal(3, PowerPointLatexText.CountFormulasInTextBodies(group));
        var body = group.Descendants().First(element => element.Name.LocalName == "txBody");
        Assert.Equal("Let $a$ and\v$$b$$", PowerPointLatexText.SavedText(body));
    }

    [Theory]
    [InlineData("latex|wordOmml|selection", "latex", "wordOmml", "wordOmml", "Selection")]
    [InlineData("latex|visualtex|slide", "latex", "visualtex", "nativeOle", "Slide")]
    [InlineData("latex|mathTypeOle|presentation", "latex", "mathTypeOle", "mathTypeOle", "Presentation")]
    [InlineData("any|latex|selection", "any", "latex", "latex", "Selection")]
    [InlineData("visualtex|latex|slide", "visualtex", "latex", "latex", "Slide")]
    [InlineData("mathTypeOle|latex|presentation", "mathTypeOle", "latex", "latex", "Presentation")]
    [InlineData("wordOmml|latex|selection", "wordOmml", "latex", "latex", "Selection")]
    public void LatexRibbonTagsResolveToARequest(string tag, string source, string target, string mode, string scope)
    {
        Assert.True(PowerPointConversionRequest.TryParse(tag, Ole, out var request));
        Assert.Equal(source, request.SourceFamily);
        Assert.Equal(target, request.TargetFamily);
        Assert.Equal(mode, request.TargetMode);
        Assert.Equal(scope, request.Scope.ToString());
    }

    [Theory]
    [InlineData("latex|latex|slide")]
    [InlineData("latex|any|slide")]
    [InlineData("tex|wordOmml|slide")]
    public void MalformedLatexTagsAreRejected(string tag)
    {
        Assert.False(PowerPointConversionRequest.TryParse(tag, Ole, out _));
    }

    [Fact]
    public void LatexCodeIsRedrawnOnlyByTheLatexCommands()
    {
        Assert.True(PowerPointConversionRequest.TryParse("latex|wordOmml|slide", Ole, out var redraw));
        Assert.True(redraw.Accepts(PowerPointFormulaKind.Latex));
        Assert.False(redraw.Accepts(PowerPointFormulaKind.Omml));
        Assert.False(redraw.Accepts(PowerPointFormulaKind.VisualTeXOle));
        Assert.False(redraw.Accepts(PowerPointFormulaKind.MathType));

        // "Unify" never touches LaTeX code.
        foreach (var tag in new[] { "any|wordOmml|slide", "any|visualtex|presentation", "any|mathTypeOle|slide" })
        {
            Assert.True(PowerPointConversionRequest.TryParse(tag, Ole, out var unify));
            Assert.False(unify.Accepts(PowerPointFormulaKind.Latex));
        }
    }

    [Fact]
    public void EveryFormulaKindTurnsIntoLatexCode()
    {
        Assert.True(PowerPointConversionRequest.TryParse("any|latex|presentation", Ole, out var all));
        Assert.True(all.Unify);
        Assert.True(all.Accepts(PowerPointFormulaKind.VisualTeXOle));
        Assert.True(all.Accepts(PowerPointFormulaKind.VisualTeXPicture));
        Assert.True(all.Accepts(PowerPointFormulaKind.MathType));
        Assert.True(all.Accepts(PowerPointFormulaKind.Omml));
        Assert.False(all.Accepts(PowerPointFormulaKind.Latex));

        Assert.True(PowerPointConversionRequest.TryParse("mathTypeOle|latex|slide", Ole, out var mathType));
        Assert.True(mathType.Accepts(PowerPointFormulaKind.MathType));
        Assert.False(mathType.Accepts(PowerPointFormulaKind.Omml));
        Assert.False(mathType.Accepts(PowerPointFormulaKind.VisualTeXOle));
    }

    [Fact]
    public void LatexTargetsAndSkipsAreWorded()
    {
        Assert.Equal("LaTeX code", PowerPointConversionRules.TargetLabel(PowerPointConversionRequest.LatexFamily, english: true));
        Assert.Equal("LaTeX 代码", PowerPointConversionRules.TargetLabel(PowerPointConversionRequest.LatexFamily, english: false));
        Assert.Equal("LaTeX", PowerPointConversionRules.FamilyLabel(PowerPointConversionRequest.LatexFamily));
        Assert.Equal(PowerPointConversionRequest.LatexFamily, PowerPointConversionRules.FamilyOf(PowerPointFormulaKind.Latex));
        // A text box holding LaTeX code can be rotated; it cannot be mirrored.
        Assert.True(PowerPointConversionRules.TargetSupportsRotation(PowerPointConversionRequest.LatexFamily));
        Assert.False(PowerPointConversionRules.TargetSupportsFlip(PowerPointConversionRequest.LatexFamily));

        var english = PowerPointConversionRules.SkipReasonText(PowerPointConversionSkipReason.LatexMixedWithText, 2, english: true);
        Assert.StartsWith("2 LaTeX formula(s) share a text box", english, StringComparison.Ordinal);
        Assert.Contains("LaTeX → OMML", english, StringComparison.Ordinal);
        var chinese = PowerPointConversionRules.SkipReasonText(PowerPointConversionSkipReason.LatexMixedWithText, 2, english: false);
        Assert.StartsWith("2 个 LaTeX 公式", chinese, StringComparison.Ordinal);
        Assert.Contains("“LaTeX → OMML”", chinese, StringComparison.Ordinal);
        var report = PowerPointConversionRules.DescribeSkips(new[]
        {
            new PowerPointConversionSkip(PowerPointConversionSkipReason.LatexMixedWithText, 3, 2),
        }, english: true);
        Assert.Contains("(slide 3)", report, StringComparison.Ordinal);
    }
}

using System.Xml.Linq;
using VisualTeX.PowerPointVsto;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Pure rules behind PowerPoint batch format conversion: Ribbon tags, which
// formulas a command converts, size rules and how saved slide XML is read.
public sealed class PowerPointFormulaConversionModelTests
{
    private const string Ole = FormulaOleContract.NativeOleMode;
    private const string Picture = FormulaOleContract.CrossPlatformPictureMode;

    [Theory]
    [InlineData("visualtex|mathTypeOle|selection", Ole, "visualtex", "mathTypeOle", "mathTypeOle", "Selection")]
    [InlineData("mathTypeOle|visualtex|slide", Ole, "mathTypeOle", "visualtex", "nativeOle", "Slide")]
    [InlineData("mathTypeOle|visualtex|slide", Picture, "mathTypeOle", "visualtex", "crossPlatformPicture", "Slide")]
    [InlineData("wordOmml|visualtex|presentation", Picture, "wordOmml", "visualtex", "crossPlatformPicture", "Presentation")]
    [InlineData("any|wordOmml|presentation", Ole, "any", "wordOmml", "wordOmml", "Presentation")]
    public void RibbonTagsResolveToARequest(string tag, string visualTeXMode, string source, string targetFamily, string targetMode, string scope)
    {
        Assert.True(PowerPointConversionRequest.TryParse(tag, visualTeXMode, out var request));
        Assert.Equal(source, request.SourceFamily);
        Assert.Equal(targetFamily, request.TargetFamily);
        Assert.Equal(targetMode, request.TargetMode);
        Assert.Equal(scope, request.Scope.ToString());
        Assert.Equal(source == "any", request.Unify);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("visualtex|mathTypeOle")]
    [InlineData("visualtex|visualtex|slide")]
    [InlineData("mathTypeOle|mathTypeOle|slide")]
    [InlineData("nativeOle|mathTypeOle|slide")]
    [InlineData("visualtex|any|slide")]
    [InlineData("visualtex|mathTypeOle|document")]
    [InlineData("visualtex|mathTypeOle|slide|extra")]
    public void MalformedTagsAreRejected(string? tag)
    {
        Assert.False(PowerPointConversionRequest.TryParse(tag, Ole, out _));
    }

    [Fact]
    public void DirectionalCommandsConvertOnlyTheirSourceFormat()
    {
        Assert.True(PowerPointConversionRequest.TryParse("visualtex|mathTypeOle|slide", Ole, out var request));
        Assert.True(request.Accepts(PowerPointFormulaKind.VisualTeXOle));
        Assert.True(request.Accepts(PowerPointFormulaKind.VisualTeXPicture));
        Assert.False(request.Accepts(PowerPointFormulaKind.MathType));
        Assert.False(request.Accepts(PowerPointFormulaKind.Omml));
    }

    [Fact]
    public void UnifyConvertsEveryOtherFormatAndLeavesTheTargetFormatAlone()
    {
        Assert.True(PowerPointConversionRequest.TryParse("any|mathTypeOle|presentation", Ole, out var toMathType));
        Assert.True(toMathType.Accepts(PowerPointFormulaKind.VisualTeXOle));
        Assert.True(toMathType.Accepts(PowerPointFormulaKind.VisualTeXPicture));
        Assert.True(toMathType.Accepts(PowerPointFormulaKind.Omml));
        Assert.False(toMathType.Accepts(PowerPointFormulaKind.MathType));

        // VisualTeX OLE and SVG are both VisualTeX: unifying to VisualTeX keeps them.
        Assert.True(PowerPointConversionRequest.TryParse("any|visualtex|slide", Picture, out var toVisualTeX));
        Assert.False(toVisualTeX.Accepts(PowerPointFormulaKind.VisualTeXOle));
        Assert.False(toVisualTeX.Accepts(PowerPointFormulaKind.VisualTeXPicture));
        Assert.True(toVisualTeX.Accepts(PowerPointFormulaKind.MathType));
        Assert.True(toVisualTeX.Accepts(PowerPointFormulaKind.Omml));
    }

    [Fact]
    public void ScaledMathTypeEquationKeepsItsVisibleSize()
    {
        // The reported PowerPoint file: a 12 pt MathType equation whose 52 x 20 pt
        // natural box was enlarged to 142.3 x 54.7 pt on the slide.
        var size = PowerPointConversionRules.ScaledFontSize(12, 142.3, 54.72, 52.0, 19.98);
        Assert.InRange(size, 32.6, 33.0);
        Assert.Equal(12, PowerPointConversionRules.ScaledFontSize(12, 52, 20, 52, 20), 6);
        Assert.Equal(24, PowerPointConversionRules.ScaledFontSize(12, 0, 40, 52, 20), 6);
        Assert.Equal(12, PowerPointConversionRules.ScaledFontSize(12, 0, 0, 0, 0), 6);
    }

    [Theory]
    [InlineData(32.8, 32.8, 1.0)]
    [InlineData(300, 200, 1.5)]
    [InlineData(2.5, 5, 0.5)]
    public void RenderSizeStaysInTheRendererRangeAndPlacementMakesUpTheRest(double effective, float render, float scale)
    {
        var (actualRender, actualScale) = PowerPointConversionRules.ResolveRenderFontSize(effective);
        Assert.Equal(render, actualRender, 2);
        Assert.Equal(scale, actualScale, 3);
    }

    [Fact]
    public void TargetsKeepOnlyTheTransformsTheyCanRepresent()
    {
        Assert.False(PowerPointConversionRules.TargetSupportsRotation(FormulaOleContract.MathTypeOleMode));
        Assert.False(PowerPointConversionRules.TargetSupportsRotation(Ole));
        Assert.True(PowerPointConversionRules.TargetSupportsRotation(Picture));
        Assert.True(PowerPointConversionRules.TargetSupportsRotation(FormulaOleContract.WordOmmlMode));
        Assert.True(PowerPointConversionRules.TargetSupportsFlip(Picture));
        Assert.False(PowerPointConversionRules.TargetSupportsFlip(FormulaOleContract.WordOmmlMode));
    }

    [Fact]
    public void SavedSlideShapesAreClassified()
    {
        var shapes = PowerPointSlideXml.IndexTopLevelShapes(XDocument.Parse(SlideXml));
        Assert.Equal(PowerPointSlideShapeKind.TextWithMath, shapes[2].Kind);
        Assert.Equal(PowerPointSlideShapeKind.TextWithMath, shapes[3].Kind);
        Assert.Equal(PowerPointSlideShapeKind.TextWithMath, shapes[4].Kind);
        Assert.Equal(PowerPointSlideShapeKind.Ole, shapes[5].Kind);
        Assert.Equal("Equation.DSMT4", shapes[5].ProgId);
        Assert.False(shapes[5].Linked);
        Assert.Equal(PowerPointSlideShapeKind.Ole, shapes[6].Kind);
        Assert.True(shapes[6].Linked);
        Assert.Equal(PowerPointSlideShapeKind.Picture, shapes[7].Kind);
        Assert.Equal(PowerPointSlideShapeKind.Group, shapes[8].Kind);
        Assert.Equal(2, shapes[8].NestedFormulaHints);
        Assert.Equal(PowerPointSlideShapeKind.TableWithMath, shapes[11].Kind);
        Assert.Equal(1, shapes[11].NestedFormulaHints);
        Assert.Equal(PowerPointSlideShapeKind.Other, shapes[12].Kind);
    }

    [Fact]
    public void StandaloneDisplayEquationIsConvertible()
    {
        var math = PowerPointSlideXml.IndexTopLevelShapes(XDocument.Parse(SlideXml))[2].Math!;
        Assert.True(math.EquationOnly);
        var equation = Assert.Single(math.Equations);
        Assert.True(equation.Display);
        Assert.Contains("oMath", equation.Omml, StringComparison.Ordinal);
        // The serialized equation must stand alone, carrying its namespaces.
        var parsed = XElement.Parse(equation.Omml);
        Assert.Equal("oMath", parsed.Name.LocalName);
    }

    [Fact]
    public void EquationInsideATextLineStaysNative()
    {
        var math = PowerPointSlideXml.IndexTopLevelShapes(XDocument.Parse(SlideXml))[3].Math!;
        Assert.False(math.EquationOnly);
        Assert.True(math.HasOtherContent);
        Assert.False(Assert.Single(math.Equations).Display);
    }

    [Fact]
    public void EquationOnlyBoxWithSeveralEquationsSplitsInOrder()
    {
        var math = PowerPointSlideXml.IndexTopLevelShapes(XDocument.Parse(SlideXml))[4].Math!;
        Assert.True(math.EquationOnly);
        Assert.Equal(2, math.Equations.Count);
        Assert.Equal(0, math.Equations[0].Paragraph);
        Assert.Equal(1, math.Equations[1].Paragraph);
        Assert.Contains(">x<", math.Equations[0].Omml, StringComparison.Ordinal);
        Assert.Contains(">y<", math.Equations[1].Omml, StringComparison.Ordinal);
    }

    [Fact]
    public void FallbackBranchesAreNotReadAsLiveContent()
    {
        // The picture fallback of a math shape carries a visible-looking space
        // run; reading it would make every equation look mixed with text.
        var shape = PowerPointSlideXml.IndexTopLevelShapes(XDocument.Parse(SlideXml))[2];
        Assert.DoesNotContain(PowerPointSlideXml.LiveDescendants(shape.Element),
            element => element.Name == PowerPointSlideXml.P + "pic" || element.Name.LocalName == "blipFill");
    }

    [Fact]
    public void SkipReportListsEveryFailedFormulaWithSlideAndReason()
    {
        var text = PowerPointConversionRules.DescribeSkips(new[]
        {
            new PowerPointConversionSkip(PowerPointConversionSkipReason.InGroup, 1, 2),
            new PowerPointConversionSkip(PowerPointConversionSkipReason.Failed, 2, 1, "first reason", "\\frac{a}{b}"),
            new PowerPointConversionSkip(PowerPointConversionSkipReason.InGroup, 3, 1),
            new PowerPointConversionSkip(PowerPointConversionSkipReason.Failed, 4, 1, "second reason"),
        }, english: true);
        var lines = text.Split(new[] { "\r\n" }, StringSplitOptions.None);
        Assert.Equal(4, lines.Length);
        Assert.Contains("3 equation(s) are inside groups", lines[0], StringComparison.Ordinal);
        Assert.Contains("(slides 1, 3)", lines[0], StringComparison.Ordinal);
        Assert.Contains("2 equation(s) failed to convert", lines[1], StringComparison.Ordinal);
        Assert.Contains("Slide 2 · \\frac{a}{b}: first reason", lines[2], StringComparison.Ordinal);
        Assert.Contains("Slide 4: second reason", lines[3], StringComparison.Ordinal);
        var many = Enumerable.Range(1, PowerPointConversionRules.DescribedSkipLimit + 3)
            .Select(index => new PowerPointConversionSkip(PowerPointConversionSkipReason.Failed, index, 1, "reason " + index))
            .ToArray();
        var capped = PowerPointConversionRules.DescribeSkips(many, english: true).Split(new[] { "\r\n" }, StringSplitOptions.None);
        Assert.Equal(PowerPointConversionRules.DescribedSkipLimit + 2, capped.Length);
        Assert.Contains("and 3 more", capped[capped.Length - 1], StringComparison.Ordinal);
        Assert.Equal(5, PowerPointConversionRules.CountSkips(new[]
        {
            new PowerPointConversionSkip(PowerPointConversionSkipReason.InTable, 1, 3),
            new PowerPointConversionSkip(PowerPointConversionSkipReason.Linked, 1, 2),
        }));
        Assert.Equal(string.Empty, PowerPointConversionRules.DescribeSkips(Array.Empty<PowerPointConversionSkip>(), english: false));
    }

    // A fake MathType preview renderer: a request either renders its whole group
    // or (crash) nothing; an equation in NeverRender crashes every group it is in.
    private sealed class FakePreviewRenderer
    {
        internal readonly List<IReadOnlyList<string>> Calls = new();
        internal readonly HashSet<string> NeverRender = new(StringComparer.Ordinal);
        internal Func<int, bool> CallSucceeds = _ => true;
        internal DateTime Clock = new(2026, 10, 15, 12, 0, 0, DateTimeKind.Utc);
        internal TimeSpan CallDuration = TimeSpan.FromSeconds(1);

        internal IReadOnlyDictionary<string, int> Render(IReadOnlyList<string> group)
        {
            Calls.Add(group.ToList());
            Clock += CallDuration;
            if (!CallSucceeds(Calls.Count) || group.Any(NeverRender.Contains))
                return new Dictionary<string, int>();
            return group.ToDictionary(id => id, _ => Calls.Count);
        }

        internal NativePreviewRetryOptions Options() => new()
        {
            Now = () => Clock,
            Pause = pause => Clock += pause,
        };
    }

    private static List<string> PreviewIds(int count) => Enumerable.Range(0, count).Select(index => "e" + index).ToList();

    [Fact]
    public void PreviewsRenderInGroupsOfTwentyFour()
    {
        var renderer = new FakePreviewRenderer();
        var rendered = NativePreviewRetry.Render(PreviewIds(30), renderer.Render, renderer.Options());
        Assert.Equal(30, rendered.Count);
        Assert.Equal(new[] { 24, 6 }, renderer.Calls.Select(call => call.Count));
    }

    [Fact]
    public void ACrashedRequestIsRenderedAgainInSmallerGroups()
    {
        var renderer = new FakePreviewRenderer { CallSucceeds = call => call != 1 };
        var rendered = NativePreviewRetry.Render(PreviewIds(30), renderer.Render, renderer.Options());
        Assert.Equal(30, rendered.Count);
        Assert.Equal(new[] { 24, 6, 4, 4, 4, 4, 4, 4 }, renderer.Calls.Select(call => call.Count));
    }

    [Fact]
    public void OnlyTheEquationThatCannotRenderIsLeftOut()
    {
        var renderer = new FakePreviewRenderer();
        renderer.NeverRender.Add("e3");
        var rendered = NativePreviewRetry.Render(PreviewIds(10), renderer.Render, renderer.Options());
        Assert.Equal(9, rendered.Count);
        Assert.False(rendered.ContainsKey("e3"));
        // 10 / groups of 4 / one by one / one recovery try for e3, then stop.
        Assert.Equal(new[] { 10, 4, 4, 2, 1, 1, 1, 1, 1 }, renderer.Calls.Select(call => call.Count));
    }

    [Fact]
    public void ARecoveringRendererGetsExtraOneByOnePasses()
    {
        // A freshly started service: the first three requests crash, then it works.
        var renderer = new FakePreviewRenderer { CallSucceeds = call => call > 3 };
        var rendered = NativePreviewRetry.Render(PreviewIds(4), renderer.Render, renderer.Options());
        Assert.Equal(4, rendered.Count);
        Assert.Equal(new[] { 4, 4, 1, 1, 1, 1, 1 }, renderer.Calls.Select(call => call.Count));
    }

    [Fact]
    public void ABrokenRendererIsGivenUpAfterNinetySecondsWithoutProgress()
    {
        var renderer = new FakePreviewRenderer { CallSucceeds = _ => false, CallDuration = TimeSpan.FromSeconds(15) };
        var rendered = NativePreviewRetry.Render(PreviewIds(60), renderer.Render, renderer.Options());
        Assert.Empty(rendered);
        Assert.InRange(renderer.Calls.Count, 5, 7);
    }

    [Fact]
    public void AStopRequestEndsRenderingAtOnce()
    {
        var renderer = new FakePreviewRenderer { CallSucceeds = _ => false };
        var options = renderer.Options();
        options.StopNow = () => renderer.Calls.Count > 0;
        NativePreviewRetry.Render(PreviewIds(60), renderer.Render, options);
        Assert.Single(renderer.Calls);
    }

    [Fact]
    public void OfficeRgbIsReadAsBlueGreenRed()
    {
        Assert.Equal("#FF0000", PowerPointConversionColors.FromOfficeRgb(0x0000FF));
        Assert.Equal("#1F2A44", PowerPointConversionColors.FromOfficeRgb(0x442A1F));
        Assert.Equal("#FFFFFF", PowerPointConversionColors.FromOfficeRgb(0xFFFFFF));
    }

    [Fact]
    public void ExplicitColorsAreThePlainRgbRunFills()
    {
        const string omml = """
<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
 <m:r><a:rPr><a:solidFill><a:srgbClr val="ff0000"/></a:solidFill></a:rPr><m:t>x</m:t></m:r>
 <m:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><m:t>y</m:t></m:r>
 <m:r><a:rPr><a:ln><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></a:rPr><m:t>z</m:t></m:r>
 <m:r><a:rPr lang="en-US"/><m:t>w</m:t></m:r>
</m:oMath>
""";
        var colors = PowerPointConversionColors.ExplicitOmmlColors(omml);
        Assert.Contains("#FF0000", colors);
        Assert.Contains("#00FF00", colors);
        Assert.Equal(2, colors.Count);
        Assert.Empty(PowerPointConversionColors.ExplicitOmmlColors("not xml"));
    }

    [Fact]
    public void InheritedColorIsTheMostUsedColorThatIsNotExplicit()
    {
        var noExplicit = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        Assert.Equal("#FFFFFF", PowerPointConversionColors.InheritedColor(new[] { ("#ffffff", 5) }, noExplicit));
        // Explicit red parts do not decide the color of the uncolored parts.
        var red = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "#FF0000" };
        Assert.Equal("#FFFFFF", PowerPointConversionColors.InheritedColor(new[] { ("#FF0000", 9), ("#FFFFFF", 2) }, red));
        Assert.Null(PowerPointConversionColors.InheritedColor(new[] { ("#FF0000", 9) }, red));
        Assert.Null(PowerPointConversionColors.InheritedColor(Array.Empty<(string, int)>(), noExplicit));
        Assert.Equal("#000000", PowerPointConversionColors.InheritedColor(new[] { ("#000000", 4), ("#FFFFFF", 1), ("#FFFFFF", 0) }, noExplicit));
    }

    [Theory]
    [InlineData("#FFFFFF", "#000000", true)]
    [InlineData("#000000", "#000000", false)]
    [InlineData("#ffffff", "#FFFFFF", false)]
    [InlineData(null, "#000000", false)]
    [InlineData("#000000", null, false)]
    public void AColorIsWrittenOnlyWhereTheTargetWouldShowAnother(string? source, string? targetDefault, bool expected)
    {
        Assert.Equal(expected, PowerPointConversionColors.NeedsExplicit(source, targetDefault));
    }

    [Fact]
    public void BaseColorColorsTheRootAndLeavesExplicitColorsAlone()
    {
        const string mathMl = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi mathcolor=\"#FF0000\">b</mi></math>";
        var colored = XElement.Parse(PowerPointConversionColors.ApplyBaseColor(mathMl, "#ffffff"));
        Assert.Equal("#FFFFFF", (string?)colored.Attribute("mathcolor"));
        XNamespace ns = "http://www.w3.org/1998/Math/MathML";
        Assert.Equal("#FF0000", (string?)colored.Elements(ns + "mi").Last().Attribute("mathcolor"));
        Assert.Equal(3, colored.Elements().Count());
        // An equation that already has a root color, or no color to apply, is unchanged.
        const string rooted = "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" mathcolor=\"#00FF00\"><mi>a</mi></math>";
        Assert.Equal(rooted, PowerPointConversionColors.ApplyBaseColor(rooted, "#FFFFFF"));
        Assert.Equal(mathMl, PowerPointConversionColors.ApplyBaseColor(mathMl, null));
    }

    // Structure taken from PowerPoint's own save of native equations (a14
    // choice + picture fallback with the same shape id) and of MathType OLE.
    private const string SlideXml = """
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
 <p:cSld><p:spTree>
  <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr/>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
   <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="standalone"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/>
     <p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr>
       <m:oMath><m:f><m:num><m:r><a:rPr lang="en-US" i="1"/><m:t>a</m:t></m:r></m:num><m:den><m:r><a:rPr lang="en-US" i="1"/><m:t>b</m:t></m:r></m:den></m:f></m:oMath>
      </m:oMathPara></a14:m><a:endParaRPr lang="zh-CN"/></a:p>
     </p:txBody></p:sp>
   </mc:Choice>
   <mc:Fallback>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="standalone"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
     <p:spPr><a:blipFill><a:blip r:embed="rId2"/></a:blipFill></p:spPr>
     <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"><a:noFill/></a:rPr><a:t> </a:t></a:r></a:p></p:txBody></p:sp>
   </mc:Fallback>
  </mc:AlternateContent>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
   <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="mixed"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/>
     <p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a:r><a:rPr lang="en-US"/><a:t>Let </a:t></a:r>
       <a14:m><m:oMath><m:sSup><m:e><m:r><a:rPr lang="en-US" i="1"/><m:t>x</m:t></m:r></m:e><m:sup><m:r><a:rPr lang="en-US" i="1"/><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></a14:m>
       <a:r><a:rPr lang="en-US"/><a:t> be large</a:t></a:r></a:p>
     </p:txBody></p:sp>
   </mc:Choice>
   <mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="3" name="mixed"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback>
  </mc:AlternateContent>
  <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
   <mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">
    <p:sp><p:nvSpPr><p:cNvPr id="4" name="twoEquations"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/>
     <p:txBody><a:bodyPr/><a:lstStyle/>
      <a:p><a14:m><m:oMathPara><m:oMath><m:r><a:rPr lang="en-US" i="1"/><m:t>x</m:t></m:r></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="en-US"/></a:p>
      <a:p><a14:m><m:oMathPara><m:oMath><m:r><a:rPr lang="en-US" i="1"/><m:t>y</m:t></m:r></m:oMath></m:oMathPara></a14:m><a:endParaRPr lang="zh-CN"/></a:p>
     </p:txBody></p:sp>
   </mc:Choice>
   <mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="4" name="twoEquations"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback>
  </mc:AlternateContent>
  <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Object 4"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>
   <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole">
    <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
     <mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v"><p:oleObj spid="_x0000_s1047" name="Equation" r:id="rId7" imgW="660240" imgH="253800" progId="Equation.DSMT4"><p:embed/></p:oleObj></mc:Choice>
     <mc:Fallback><p:oleObj name="Equation" r:id="rId7" imgW="660240" imgH="253800" progId="Equation.DSMT4"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId8"/></p:blipFill><p:spPr/></p:pic></p:oleObj></mc:Fallback>
    </mc:AlternateContent>
   </a:graphicData></a:graphic></p:graphicFrame>
  <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="6" name="Linked"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>
   <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Equation" r:id="rId9" progId="Equation.DSMT4"><p:link updateAutomatic="1"/></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>
  <p:pic><p:nvPicPr><p:cNvPr id="7" name="VisualTeX_x" descr="metadata"/><p:cNvPicPr/><p:nvPr><p:custDataLst><p:tags r:id="rId10"/></p:custDataLst></p:nvPr></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>
  <p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Group 7"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
   <p:pic><p:nvPicPr><p:cNvPr id="9" name="inner picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill/><p:spPr/></p:pic>
   <p:sp><p:nvSpPr><p:cNvPr id="10" name="inner math"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMath><m:r><m:t>z</m:t></m:r></m:oMath></a14:m></a:p></p:txBody></p:sp>
  </p:grpSp>
  <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="11" name="Table 10"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/>
   <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tr h="1"><a:tc><a:txBody><a:bodyPr/><a:p><a14:m xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"><m:oMath><m:r><m:t>t</m:t></m:r></m:oMath></a14:m></a:p></a:txBody></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
  <p:sp><p:nvSpPr><p:cNvPr id="12" name="plain text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>
 </p:spTree></p:cSld>
</p:sld>
""";
}

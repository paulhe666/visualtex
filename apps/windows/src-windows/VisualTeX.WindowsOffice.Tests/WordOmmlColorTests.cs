using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlColorTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static string Math(string body) => "<math xmlns=\"http://www.w3.org/1998/Math/MathML\">" + body + "</math>";
    private static XElement Forward(string body) => XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(Math(body)));
    private static string? RunColor(XElement run) => (string?)run.Element(W + "rPr")?.Element(W + "color")?.Attribute(W + "val");

    [Theory]
    [InlineData("#2563eb", "2563EB")]
    [InlineData("#f00", "FF0000")]
    [InlineData("red", "FF0000")]
    public void InheritedColorIsWrittenToEveryRun(string input, string expected)
    {
        var result = Forward("<mstyle mathcolor=\"" + input + "\"><mi>x</mi><mo>+</mo><mn>2</mn><mtext>ok</mtext></mstyle>");
        Assert.NotEmpty(result.Descendants(M + "r"));
        Assert.All(result.Descendants(M + "r"), run => Assert.Equal(expected, RunColor(run)));
    }

    [Fact]
    public void AdjacentSameStyleTokensKeepDistinctColors()
    {
        var result = Forward("<mi mathcolor=\"#ff0000\">x</mi><mi mathcolor=\"#0000ff\">y</mi><mi>z</mi>");
        var runs = result.Descendants(M + "r").ToArray();
        Assert.Equal(3, runs.Length);
        Assert.Equal(new string?[] { "FF0000", "0000FF", null }, runs.Select(RunColor));
        Assert.Equal("xyz", string.Concat(runs.Select(r => (string?)r.Element(M + "t"))));
    }

    [Fact]
    public void NestedOverrideDoesNotBleedIntoFollowingSibling()
    {
        var result = Forward("<mstyle mathcolor=\"#2563eb\"><mi>x</mi><mstyle mathcolor=\"#dc2626\"><mi>y</mi></mstyle><mi>z</mi></mstyle><mi>q</mi>");
        Assert.Equal(new string?[] { "2563EB", "DC2626", "2563EB", null }, result.Descendants(M + "r").Select(RunColor));
    }

    [Fact]
    public void RootColorRetainsBoldAndItalic()
    {
        var xml = Math("<mi mathvariant=\"bold\">x</mi><mi mathvariant=\"italic\">y</mi>").Replace("<math ", "<math mathcolor=\"#2563eb\" ");
        var result = XElement.Parse(WordOmmlConverter.TransformMathMlToOmml(xml));
        var runs = result.Descendants(M + "r").ToArray();
        Assert.All(runs, r => Assert.Equal("2563EB", RunColor(r)));
        Assert.Equal(new[] { "b", "i" }, runs.Select(r => (string?)r.Element(M + "rPr")?.Element(M + "sty")?.Attribute(M + "val")));
    }

    [Theory]
    [InlineData("<mfrac><mi>x</mi><mi>y</mi></mfrac>", "fPr")]
    [InlineData("<msqrt><mi>x</mi></msqrt>", "radPr")]
    [InlineData("<mover accent=\"true\"><mi>x</mi><mo stretchy=\"false\">^</mo></mover>", "accPr")]
    public void StructuralInkReceivesColor(string body, string properties)
    {
        var result = Forward("<mstyle mathcolor=\"#2563eb\">" + body + "</mstyle>");
        Assert.All(result.Descendants(M + "r"), r => Assert.Equal("2563EB", RunColor(r)));
        var control = result.Descendants(M + properties).Single().Element(M + "ctrlPr")?.Element(W + "rPr");
        Assert.Equal("2563EB", (string?)control?.Element(W + "color")?.Attribute(W + "val"));
    }

    [Fact]
    public void UncoloredFormulaDoesNotGainWordColors()
    {
        Assert.Empty(Forward("<mi>x</mi><mo>+</mo><mfrac><mi>a</mi><mi>b</mi></mfrac>").Descendants(W + "color"));
    }

    [Theory]
    [InlineData("<msup><mi>x</mi><mn>2</mn></msup>")]
    [InlineData("<msub><mi>x</mi><mi>i</mi></msub>")]
    [InlineData("<mi mathvariant=\"normal\">sin</mi><mo>⁡</mo><mi>x</mi>")]
    public void ScriptAndFunctionTokenReadbackKeepsColor(string body)
    {
        var result = Forward("<mstyle mathcolor=\"#2563eb\">" + body + "</mstyle>");
        var back = XElement.Parse(WordOmmlConverter.TransformOmmlToMathMl(result.ToString(), false));
        var tokens = back.Descendants().Where(e => e.Name.LocalName is "mi" or "mn").ToArray();
        Assert.NotEmpty(tokens);
        Assert.All(tokens, e => Assert.Equal("#2563EB", MathMlColorContract.Effective(e)));
    }

    [Fact]
    public void NativeRunsRecoverColorsWithoutStoredLatex()
    {
        var xml = "<m:oMath xmlns:m=\"" + M + "\" xmlns:w=\"" + W + "\"><m:r><w:rPr><w:color w:val=\"DC2626\"/></w:rPr><m:t>x</m:t></m:r><m:r><m:t>+</m:t></m:r><m:r><w:rPr><w:color w:val=\"2563EB\"/></w:rPr><m:t>y</m:t></m:r></m:oMath>";
        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(xml, false);
        var source = MathMlToLatexConverter.Convert(mathMl);
        Assert.Contains("#DC2626", source);
        Assert.Contains("#2563EB", source);
    }

    [Fact]
    public void MathMlSourceKeepsNestedColorAndDefaultScope()
    {
        var source = MathMlToLatexConverter.Convert(Math("<mstyle mathcolor=\"#2563eb\"><mi>x</mi><mstyle mathcolor=\"#dc2626\"><mi>y</mi></mstyle><mi>z</mi></mstyle><mi>q</mi>"));
        Assert.Equal(@"\textcolor{#2563EB}{x\textcolor{#DC2626}{y}z}q", source);
    }

    [Fact]
    public void IndependentAccentColorDoesNotRestyleBase()
    {
        var result = Forward("<mover accent=\"true\"><mi mathcolor=\"#2563eb\">x</mi><mo mathcolor=\"#dc2626\" stretchy=\"false\">^</mo></mover>");
        Assert.Equal("2563EB", RunColor(result.Descendants(M + "r").Single()));
        Assert.Equal("DC2626", (string?)result.Descendants(M + "accPr").Single().Element(M + "ctrlPr")?.Element(W + "rPr")?.Element(W + "color")?.Attribute(W + "val"));
        var source = MathMlToLatexConverter.Convert(WordOmmlConverter.TransformOmmlToMathMl(result.ToString(), false));
        Assert.Contains("#DC2626", source);
        Assert.Contains("#2563EB", source);
    }
}

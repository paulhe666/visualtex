using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// Word formulas that previously failed to read, or read differently depending
// on which Office build's OMML2MML.XSL was installed.
public sealed class OmmlReadRegressionTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static string Math(string body) => "<m:oMath xmlns:m=\"" + M + "\">" + body + "</m:oMath>";

    private static string Run(string text, string properties = "") =>
        "<m:r>" + (properties.Length > 0 ? "<m:rPr>" + properties + "</m:rPr>" : string.Empty)
        + "<m:t>" + text + "</m:t></m:r>";

    private static string EquationArray(params string[] rows) =>
        "<m:eqArr>" + string.Concat(rows.Select(row => "<m:e>" + row + "</m:e>")) + "</m:eqArr>";

    private static string Latex(string omml) =>
        MathMlToLatexConverter.Convert(WordOmmlConverter.TransformOmmlToMathMl(omml, false));

    // Microsoft 365 updates ship an OMML2MML.XSL that writes styled letters as
    // Unicode mathematical alphanumerics. VisualTeX reads OMML itself, so the
    // same Word formula always reads back with the same LaTeX styles.
    [Theory]
    [InlineData("<m:sty m:val=\"b\"/>", "x", "\\mathbf{x}")]
    [InlineData("<m:sty m:val=\"bi\"/>", "x", "\\mathbfit{x}")]
    [InlineData("<m:scr m:val=\"fraktur\"/><m:sty m:val=\"p\"/>", "g", "\\mathfrak{g}")]
    [InlineData("<m:scr m:val=\"sans-serif\"/><m:sty m:val=\"p\"/>", "A", "\\mathsf{A}")]
    [InlineData("<m:scr m:val=\"monospace\"/><m:sty m:val=\"p\"/>", "x", "\\mathtt{x}")]
    [InlineData("<m:sty m:val=\"b\"/>", "2", "\\mathbf{2}")]
    public void StyledRunsReadBackWithTheirLatexStyle(string properties, string text, string expected)
        => Assert.Equal(expected, Latex(Math(Run(text, properties))));

    [Fact]
    public void StyledRunsNeverReadBackAsMathematicalAlphanumericCharacters()
    {
        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
            Math(Run("x", "<m:sty m:val=\"b\"/>") + Run("R", "<m:scr m:val=\"double-struck\"/>")), false);
        Assert.DoesNotContain("\uD835", mathMl);
    }

    // Word allows several equation arrays side by side in one equation; the
    // native equation-number lookup used to require exactly one.
    [Fact]
    public void AnEquationWithTwoRootEquationArraysReadsBack()
    {
        var latex = Latex(Math(EquationArray(Run("a&amp;=1")) + Run("+") + EquationArray(Run("b&amp;=2"))));
        Assert.Contains("a", latex);
        Assert.Contains("b", latex);
    }

    // Office's stylesheet split supplementary-plane characters in equation
    // arrays into lone surrogates and failed to write the result.
    [Fact]
    public void ASupplementaryCharacterInsideAnEquationArrayReadsBack()
    {
        var mathMl = WordOmmlConverter.TransformOmmlToMathMl(
            Math(EquationArray(Run("𝑥&amp;=𝜋"))), false);
        var root = XElement.Parse(mathMl);
        Assert.Contains("𝑥", root.Value);
        Assert.Contains("𝜋", root.Value);
    }

    // A switch value outside the OOXML spellings reads as Office reads it.
    [Theory]
    [InlineData("On")]
    [InlineData("bogus")]
    [InlineData("")]
    public void AnUnexpectedSwitchValueReadsBackAsOn(string value)
    {
        var omml = Math("<m:rad><m:radPr><m:degHide m:val=\"" + value + "\"/></m:radPr><m:deg/><m:e>" + Run("x") + "</m:e></m:rad>");
        Assert.Equal("\\sqrt{x}", Latex(omml));
    }
}

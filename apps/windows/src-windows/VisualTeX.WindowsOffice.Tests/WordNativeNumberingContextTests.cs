using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordNativeNumberingContextTests
{
    [Theory]
    [InlineData("2-1", "2", "-", 1, 1)]
    [InlineData("12.18", "12", ".", 18, 1)]
    [InlineData("2.3-9", "2.3", "-", 9, 2)]
    [InlineData("2.3.19", "2.3", ".", 19, 2)]
    [InlineData("IV-7", "IV", "-", 7, 1)]
    [InlineData("A.3", "A", ".", 3, 1)]
    [InlineData("17", "", "", 17, 0)]
    [InlineData("2−4", "2", "−", 4, 1)]
    public void LiteralAndFieldNumberLabelsUseTheSameSeriesModel(string text, string prefix, string separator, int ordinal, int level)
    {
        Assert.True(WordNativeNumberingAdoptionPlan.TryParseLabel(text, out var style));
        Assert.Equal(prefix, style.Prefix); Assert.Equal(separator, style.Separator);
        Assert.Equal(ordinal, style.Ordinal); Assert.Equal(level, style.HeadingLevel);
    }

    [Theory]
    [InlineData("x+1")][InlineData("0")][InlineData("-1")][InlineData("2-0")]
    [InlineData("2147483648")][InlineData("100%")][InlineData("2\\r1")]
    [InlineData("01")][InlineData("2-01")]
    public void NonNumberSyntaxAndDifferentNumericPicturesCannotBeSilentlyAdopted(string text)
        => Assert.False(WordNativeNumberingAdoptionPlan.TryParseLabel(text, out _));

    [Fact]
    public void CurrentLabelIgnoresInstructionTextAndDeletedOrMovedFromNumerals()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var xml = WordOmmlConverter.BuildWordNativeNumberedOmml(
            "<m:oMath xmlns:m=\"" + m + "\"><m:r><m:t>x</m:t></m:r></m:oMath>",
            "User Counter", null, 1, "-", "3", "2");
        var equation = XElement.Parse(xml);
        var argument = equation.Element(m + "eqArr")!.Element(m + "e")!.Element(m + "d")!.Element(m + "e")!;
        argument.Add(new XElement(m + "r", new XElement(w + "del", new XElement(m + "t", "99"))));
        argument.Add(new XElement(w + "moveFrom", new XElement(m + "r", new XElement(m + "t", "88"))));
        Assert.Equal("2-3", WordOmmlConverter.ReadWordNativeEquationNumberText(equation.ToString()));
    }

    [Fact]
    public void WordExportedMathTextInsideFieldCodeIsNotVisibleNumberText()
    {
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var xml = WordOmmlConverter.BuildWordNativeNumberedOmml(
            "<m:oMath xmlns:m=\"" + m + "\"><m:r><m:t>x</m:t></m:r></m:oMath>",
            "Equation", null, 1, "-", "2", "2");
        var equation = XElement.Parse(xml);
        foreach (var instruction in equation.Descendants(w + "instrText").ToArray())
            instruction.ReplaceWith(new XElement(m + "t", instruction.Value));
        Assert.Equal("2-2", WordOmmlConverter.ReadWordNativeEquationNumberText(equation.ToString()));
    }

    [Fact]
    public void NativeResetIsFieldSyntaxNotAnOwnedIdOrLiteralCacheHack()
    {
        var xml = WordOmmlConverter.BuildWordNativeNumberedOmml(
            "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>",
            "Equation", null, 0, "", "17", null, 17);
        Assert.Contains("SEQ Equation \\r 17", xml);
        Assert.True(WordOmmlConverter.HasVisualTeXDirectSequenceEquationNumber(xml));
        Assert.Equal("17", WordOmmlConverter.ReadWordNativeEquationNumberText(xml));
        Assert.DoesNotContain("bookmark", xml);
        Assert.DoesNotContain("VTOMML", xml);
    }
}

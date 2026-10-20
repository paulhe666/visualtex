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

    // A plain SEQ's visible ordinal is Word's count; after the number above it
    // was removed it is still the stale old result, never a \r start.
    [Theory]
    [InlineData(null, 5, 5)]
    [InlineData(" SEQ Equation \\r 4 \\* ARABIC ", 2, 4)]
    [InlineData(" SEQ Equation \\R 3 \\s 1 \\* ARABIC ", 9, 3)]
    [InlineData(" SEQ Equation \\s 1 \\* ARABIC \\* MERGEFORMAT ", 2, null)]
    [InlineData(" SEQ \"User Counter\" \\* ARABIC ", 7, null)]
    public void OnlyAnExplicitStartPassesToANumberInsertedBeforeTheSeries(string? sequenceCode, int visibleOrdinal, int? expected)
        => Assert.Equal(expected, WordNativeNumberingAdoptionPlan.ExplicitStartingValue(sequenceCode, visibleOrdinal));

    [Fact]
    public void BatchNumbersWithoutExistingNumbersStartTheirOwnSeries()
    {
        var mapping = WordNativeNumberingAdoptionPlan.MapBatchInsertions(
            Array.Empty<int>(), Array.Empty<string>(), new[] { 40, 10, 40 });
        Assert.Equal(new[] { 10, 40 }, mapping.Keys.OrderBy(key => key));
        Assert.All(mapping.Values, value => Assert.Equal((-1, false), value));
    }

    [Fact]
    public void OnlyTheEarliestBatchNumberBeforeASeriesTakesItsStart()
    {
        // One continuous series at 100 and 200; new numbers at 50, 80 and 150.
        var mapping = WordNativeNumberingAdoptionPlan.MapBatchInsertions(
            new[] { 100, 200 }, new[] { "s", "s" }, new[] { 150, 80, 50 });
        Assert.Equal((0, true), mapping[50]);
        Assert.Equal((0, false), mapping[80]);
        Assert.Equal((0, false), mapping[150]);
    }

    [Fact]
    public void BatchNumbersJoinTheSeriesOfTheirNearestPrecedingNumber()
    {
        // Chapter 1 numbers at 100 and 200, chapter 2 numbers at 300 and 400.
        var mapping = WordNativeNumberingAdoptionPlan.MapBatchInsertions(
            new[] { 100, 200, 300, 400 }, new[] { "1", "1", "2", "2" }, new[] { 250, 350, 450, 280 });
        Assert.Equal((0, false), mapping[250]);
        Assert.Equal((0, false), mapping[280]);
        Assert.Equal((2, false), mapping[350]);
        Assert.Equal((2, false), mapping[450]);
    }

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

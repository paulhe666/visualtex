using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordFormulaHostTests
{
    [Theory]
    [InlineData(10, 20, 12, 15, true)]
    [InlineData(10, 20, 10, 20, true)]
    [InlineData(10, 20, 20, 21, false)]
    [InlineData(10, 20, 20, 20, false)]
    [InlineData(10, 20, 0, 10, false)]
    [InlineData(10, 20, 9, 21, false)]
    [InlineData(10, 20, 12, 21, false)]
    public void NeighbouringMathDoesNotOwnAnExternalCaption(int os, int oe, int start, int end, bool expected)
        => Assert.Equal(expected, WordFormulaHost.ContainsPhysicalRange(os, oe, start, end));

    [Theory]
    [InlineData("\r\a", 1)]
    [InlineData("\r", 1)]
    [InlineData("\a", 1)]
    [InlineData("x", 0)]
    [InlineData(null, 0)]
    public void StructuralTerminatorCountsWordPositionsNotTextCharacters(string? text, int expected)
        => Assert.Equal(expected, WordFormulaHost.ParagraphTerminatorStoryLength(text));

    [Theory]
    [InlineData("\r\a", true)]
    [InlineData("\a", true)]
    [InlineData("\r", false)]
    [InlineData(" ", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void FieldScopeExcludesOnlyWordCellMarkers(string? text, bool expected)
        => Assert.Equal(expected, WordFormulaHost.IsCellStructuralMarker(text));

    [Theory]
    [InlineData(0, 5, 10, 20, false)]
    [InlineData(10, 15, 10, 20, true)]
    [InlineData(10, 20, 10, 20, true)]
    [InlineData(20, 20, 10, 20, false)]
    [InlineData(15, 21, 10, 20, false)]
    [InlineData(9, 15, 10, 20, false)]
    [InlineData(15, 14, 10, 20, false)]
    [InlineData(15, 15, 10, 20, true)]
    public void OnlyActualContainedSourceBelongsToCell(int start, int end, int ownerStart, int ownerEnd, bool expected)
        => Assert.Equal(expected, WordFormulaHost.IsContainedByCell(start, end, ownerStart, ownerEnd));

    [Theory]
    [InlineData(" MACROBUTTON MTEditEquationSection2 Equation Chapter 1 Section 1", true)]
    [InlineData("macrobutton\tMTEditEquationSection2 ", true)]
    [InlineData("MACROBUTTON MTEditEquationSection22 foo", false)]
    [InlineData("REF MTEditEquationSection2", false)]
    [InlineData("text MACROBUTTON MTEditEquationSection2", false)]
    [InlineData("MACROBUTTON MTPlaceRef", false)]
    public void OnlyExactNativeSectionCommandCanBeIsolated(string code, bool expected)
        => Assert.Equal(expected, MathTypeSourceHost.IsSectionStateCode(code));

    [Theory]
    [InlineData("\t ", true)]
    [InlineData("", true)]
    [InlineData("\u00a0", true)]
    [InlineData("123", false)]
    [InlineData("text", false)]
    [InlineData("\r", false)]
    [InlineData("\a", false)]
    [InlineData("\u0013", false)]
    public void SectionPrefixDoesNotClaimProseOrOtherFields(string text, bool expected)
        => Assert.Equal(expected, MathTypeSourceHost.IsScaffoldWhitespace(text));
}

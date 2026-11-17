using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

// The text of a range read with its field codes (every field begins with chr(19) and its instruction): the tail update of
// the numbers (WordNumberingFieldCollection.BulkUpdateFrom) locks the fields that are not number fields only when the
// range holds one.
public sealed class NumberFieldScanTests
{
    private static readonly string[] Number = { "SEQ", "MACROBUTTON", "EMBED", "STYLEREF" };

    private const string Begin = "\u0013";
    private const string Separator = "\u0014";
    private const string End = "\u0015";

    // a VisualTeX numbered display formula: heading state, the OLE object, the PlaceRef with its nested sequence fields
    private static string Formula(int n) =>
        $"{Begin} SEQ VisualTeXChapter \\r 0 \\h \\* MERGEFORMAT {Separator}{End}\u0001\t" +
        $"{Begin} MACROBUTTON VisualTeXPlaceRef {Begin} SEQ VisualTeXEquation \\s 1 \\h {Separator}{End}({Begin} SEQ VisualTeXChapter \\c {Separator}0{End}." +
        $"{Begin} SEQ VisualTeXEquation \\c {Separator}{n}{End}){Separator}{End}\r" +
        $"{Begin} EMBED VisualTeX.Formula.1 \\s {Separator}{End}";

    [Fact]
    public void ANumberedFormulaIsOnlyNumberFields() =>
        Assert.True(WordFieldCode.AllFieldsAre(Formula(1) + "text\r" + Formula(2), Number));

    [Fact]
    public void TextWithoutFieldsIsOnlyNumberFields() =>
        Assert.True(WordFieldCode.AllFieldsAre("plain text\rmore", Number));

    [Theory]
    [InlineData("TOC \\o \"1-3\" \\h")]
    [InlineData("PAGEREF _Toc1 \\h")]
    [InlineData("REF _Ref123 \\h")]
    [InlineData("DATE \\@ \"yyyy\"")]
    [InlineData("HYPERLINK \"https://example.org\"")]
    [InlineData("= 1+2")]
    public void AnotherFieldIsFound(string instruction) =>
        Assert.False(WordFieldCode.AllFieldsAre($"{Formula(1)}\r{Begin} {instruction} {Separator}x{End}\r{Formula(2)}", Number));

    [Fact]
    public void TheKeywordIsComparedWithoutCaseAndWholeWord() =>
        Assert.False(WordFieldCode.AllFieldsAre($"{Begin} SEQUENCE x {Separator}{End}", Number));

    [Fact]
    public void NoTextCannotBeJudged() =>
        Assert.False(WordFieldCode.AllFieldsAre(null, Number));

    // Renumbering Word's own equations leaves the numbers of the other kinds alone: their fields are recognised by
    // the sequence identifier.
    [Theory]
    [InlineData(" SEQ VisualTeXEquation \\s 1 \\h ")]
    [InlineData(" SEQ MTEqn \\c \\* Arabic ")]
    [InlineData(" SEQ MTSec \\h ")]
    [InlineData(" SEQ visualtexchapter \\c ")]
    public void ANumberOfAnotherKindIsRecognised(string instruction) =>
        Assert.True(WordNumberingFieldCollection.MentionsOtherKind(instruction));

    [Theory]
    [InlineData(" SEQ \u516c\u5f0f \\s 1 \\* ARABIC \\* MERGEFORMAT ")]
    [InlineData(" SEQ Equation \\* ARABIC ")]
    [InlineData(" STYLEREF 1 \\s ")]
    [InlineData("")]
    [InlineData(null)]
    public void ANativeNumberIsNotAnotherKind(string? instruction) =>
        Assert.False(WordNumberingFieldCollection.MentionsOtherKind(instruction));
}

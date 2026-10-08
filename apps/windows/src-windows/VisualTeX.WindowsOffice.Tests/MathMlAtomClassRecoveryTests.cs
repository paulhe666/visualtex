using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlAtomClassRecoveryTests
{
    [Theory]
    [InlineData("mrow", "REL", @"\mathrel{⫽}")]
    [InlineData("mrow", "BIN", @"\mathbin{⫽}")]
    [InlineData("mrow", "OP", @"\mathop{⫽}")]
    [InlineData("mstyle", "REL", @"\mathrel{⫽}")]
    [InlineData("mstyle", "BIN", @"\mathbin{⫽}")]
    [InlineData("mstyle", "OP", @"\mathop{⫽}")]
    [InlineData("mpadded", "REL", @"\mathrel{⫽}")]
    [InlineData("mpadded", "BIN", @"\mathbin{⫽}")]
    [InlineData("mpadded", "OP", @"\mathop{⫽}")]
    public void PreservesOnlyExplicitAtomClassWithoutGuessingPackage(
        string tag, string atomClass, string expected)
        => Assert.Equal(expected, MathMlToLatexConverter.Convert(
            $"<math><{tag} data-mjx-texclass='{atomClass}'><mo>&#x2AFD;</mo></{tag}></math>"));

    [Theory]
    [InlineData("REL", @"\mathrel{x+y}")]
    [InlineData("BIN", @"\mathbin{x+y}")]
    [InlineData("OP", @"\mathop{x+y}")]
    public void PreservesClassForArbitraryCompositeBodies(string atomClass, string expected)
        => Assert.Equal(expected, MathMlToLatexConverter.Convert(
            $"<math><mrow data-mjx-texclass='{atomClass}'><mi>x</mi><mo>+</mo><mi>y</mi></mrow></math>"));

    [Fact]
    public void PreservesNestedClassStyleAndScripts()
        => Assert.Equal(@"\mathbin{\mathrel{\mathit{\parallel}}}_{i}^{2}",
            MathMlToLatexConverter.Convert(
                "<math><msubsup><mrow data-mjx-texclass='BIN'>"
                + "<mstyle mathvariant='italic' data-mjx-texclass='REL'><mo>∥</mo></mstyle>"
                + "</mrow><mi>i</mi><mn>2</mn></msubsup></math>"));

    [Theory]
    [InlineData("ORD")]
    [InlineData("OPEN")]
    [InlineData("CLOSE")]
    [InlineData("INNER")]
    [InlineData("unknown")]
    public void LeavesOtherContainerClassesOnTheirEstablishedPath(string atomClass)
        => Assert.Equal("⫽", MathMlToLatexConverter.Convert(
            $"<math><mrow data-mjx-texclass='{atomClass}'><mo>⫽</mo></mrow></math>"));

    [Fact]
    public void DoesNotInferPackageOrRelationFromBareDoubleSolidus()
        => Assert.Equal("⫽", MathMlToLatexConverter.Convert("<math><mo>⫽</mo></math>"));

    [Fact]
    public void DoesNotStripRelationClassThroughLooseBinomialRecognition()
        => Assert.Equal(@"\mathrel{(}\genfrac{}{}{0pt}{}{n}{k})", MathMlToLatexConverter.Convert(
            "<math><mrow data-mjx-texclass='REL'><mo>(</mo></mrow>"
            + "<mfrac linethickness='0'><mi>n</mi><mi>k</mi></mfrac><mo>)</mo></math>"));

    [Fact]
    public void EmptyClassWrapperRemainsEmpty()
        => Assert.Equal("", MathMlToLatexConverter.Convert(
            "<math><mrow data-mjx-texclass='REL'/></math>"));
}

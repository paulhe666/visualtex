using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeNumberingFastPathTests
{
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void ZeroProofAcceptsBodyWithoutMathTypePlaceRefInstruction()
    {
        var xml = $"<w:document xmlns:w=\"{WordNamespace}\"><w:body><w:p><w:r><w:instrText> SEQ VisualTeXEquation </w:instrText></w:r></w:p></w:body></w:document>";
        Assert.True(WordDocumentXml.CanProveNoMathTypePlaceRefFields(xml));
    }

    [Fact]
    public void ZeroProofRejectsPlaceRefEvenWhenInstructionTokenIsSplitAcrossRuns()
    {
        var xml = $"<w:document xmlns:w=\"{WordNamespace}\"><w:body><w:p><w:r><w:instrText> MACROBUTTON MTPlace</w:instrText></w:r><w:r><w:instrText>Ref </w:instrText></w:r></w:p></w:body></w:document>";
        Assert.False(WordDocumentXml.CanProveNoMathTypePlaceRefFields(xml));
    }

    [Fact]
    public void ZeroProofIgnoresVisibleTextAndRejectsMalformedXml()
    {
        var visible = $"<w:document xmlns:w=\"{WordNamespace}\"><w:body><w:p><w:r><w:t>MTPlaceRef documentation</w:t></w:r></w:p></w:body></w:document>";
        Assert.True(WordDocumentXml.CanProveNoMathTypePlaceRefFields(visible));
        Assert.False(WordDocumentXml.CanProveNoMathTypePlaceRefFields("not xml"));
    }

    [Fact]
    public void VisualTeXZeroProofAcceptsMathTypeOnlyNumberingFields()
    {
        var xml = $"<w:document xmlns:w=\"{WordNamespace}\"><w:body><w:p>"
            + "<w:r><w:instrText> MACROBUTTON MTPlaceRef </w:instrText></w:r>"
            + "<w:r><w:instrText> SEQ MTEqn \\h </w:instrText></w:r>"
            + "<w:r><w:instrText> SEQ MTChap \\c </w:instrText></w:r>"
            + "<w:r><w:instrText> SEQ MTSec \\c </w:instrText></w:r>"
            + "</w:p></w:body></w:document>";
        Assert.True(WordDocumentXml.CanProveNoVisualTeXEquationNumberFields(xml));
    }

    [Fact]
    public void VisualTeXZeroProofRejectsSequenceSplitAcrossInstructionRuns()
    {
        var xml = $"<w:document xmlns:w=\"{WordNamespace}\"><w:body><w:p>"
            + "<w:r><w:instrText> SEQ VisualTeX</w:instrText></w:r>"
            + "<w:r><w:instrText>Equation \\* ARABIC </w:instrText></w:r>"
            + "</w:p></w:body></w:document>";
        Assert.False(WordDocumentXml.CanProveNoVisualTeXEquationNumberFields(xml));
    }

    [Fact]
    public void VisualTeXZeroProofRejectsMalformedXml()
        => Assert.False(WordDocumentXml.CanProveNoVisualTeXEquationNumberFields("not xml"));
}

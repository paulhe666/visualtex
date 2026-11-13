using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlTrackedDeletionCountTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static string Body(string content) =>
        $"<w:document xmlns:w=\"{W}\" xmlns:m=\"{M}\"><w:body><w:p>{content}</w:p></w:body></w:document>";

    [Fact]
    public void ATrackedDeletedCopyIsNotASecondCurrentEquation()
    {
        var xml = Body(
            "<w:del w:id=\"1\" w:author=\"a\"><m:oMath><m:r><m:t>old</m:t></m:r></m:oMath></w:del>"
            + "<m:oMath><m:r><m:t>new</m:t></m:r></m:oMath>");
        Assert.Equal(1, WordOmmlConverter.CountOMathElements(xml));
        Assert.Contains("new", WordOmmlConverter.ExtractSingleOMath(xml));
    }

    [Fact]
    public void TwoCurrentEquationsAreStillTwo()
    {
        var xml = Body(
            "<m:oMath><m:r><m:t>a</m:t></m:r></m:oMath>"
            + "<w:moveFrom w:id=\"2\" w:author=\"a\"><m:oMath><m:r><m:t>moved</m:t></m:r></m:oMath></w:moveFrom>"
            + "<m:oMath><m:r><m:t>b</m:t></m:r></m:oMath>");
        Assert.Equal(2, WordOmmlConverter.CountOMathElements(xml));
        Assert.Throws<InvalidDataException>(() => WordOmmlConverter.ExtractSingleOMath(xml));
    }
}

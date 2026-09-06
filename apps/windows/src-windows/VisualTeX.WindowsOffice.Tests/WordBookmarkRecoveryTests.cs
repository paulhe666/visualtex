using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordBookmarkRecoveryTests
{
    private const string Start = "<w:bookmarkStart w:id='owned' w:name='owned'/>";
    private const string End = "<w:bookmarkEnd w:id='owned'/>";
    private const string Number = "<w:r><w:t>2.0-3</w:t></w:r>";
    private static string Body(string contents) => "<w:body xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:p>" + contents + "</w:p></w:body>";

    [Fact]
    public void UndoShortenedOwnedCaptionCanBeReboundAfterExactContentProof()
        => WordBookmarkRecoverySnapshot.VerifyOnlyBookmarkSpansChanged(
            Body(Start + Number + End), Body(Number + Start + End), new[] { "owned" });

    [Theory]
    [InlineData("<w:r><w:t>2.0-4</w:t></w:r>")]
    [InlineData("<w:r><w:rPr><w:b/></w:rPr><w:t>2.0-3</w:t></w:r>")]
    [InlineData("<w:r><w:instrText> REF another </w:instrText></w:r>")]
    [InlineData("<w:r><w:t>2.0-3</w:t></w:r><w:bookmarkStart w:id='unrelated' w:name='unrelated'/>")]
    public void ContentFormattingFieldsAndUnrelatedBookmarksPreventRecovery(string changed)
        => Assert.Throws<InvalidDataException>(() => WordBookmarkRecoverySnapshot.VerifyOnlyBookmarkSpansChanged(
            Body(Start + Number + End), Body(changed + Start + End), new[] { "owned" }));

    [Fact]
    public void UncapturedNameCannotAuthorizeARepair()
        => Assert.Throws<InvalidDataException>(() => WordBookmarkRecoverySnapshot.VerifyOnlyBookmarkSpansChanged(
            Body(Start + Number + End), Body(Number + Start + End), new[] { "unknown" }));

    [Fact]
    public void UnrelatedBookmarkSpanDriftPreventsRecovery()
    {
        const string other = "<w:bookmarkStart w:id='other' w:name='other'/>";
        const string otherEnd = "<w:bookmarkEnd w:id='other'/>";
        Assert.Throws<InvalidDataException>(() => WordBookmarkRecoverySnapshot.VerifyOnlyBookmarkSpansChanged(
            Body(Start + other + Number + otherEnd + End), Body(Number + Start + other + otherEnd + End), new[] { "owned" }));
    }
}

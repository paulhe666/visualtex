using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordRecoveryXmlEquivalenceTests
{
    private static string Body(string content) =>
        "<w:body xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'><w:p>" + content + "</w:p></w:body>";
    private const string Text = "<w:r><w:rPr><w:sz w:val='16'/></w:rPr><w:t>正文</w:t></w:r>";
    private const string Chinese = "<w:lang w:eastAsia='zh-CN'/>";
    private static bool Equivalent(string before, string after) =>
        WordRecoveryXmlEquivalence.HasOnlyAutomaticChineseProofingAdditions(Body(before), Body(after));

    [Fact]
    public void ImplicitChineseProofingCanBeMadeExplicitWithoutChangingContents()
    {
        Assert.True(Equivalent(Text, Text.Replace("</w:rPr>", Chinese + "</w:rPr>")));
        var before = "<w:pPr><w:jc w:val='center'/></w:pPr><w:r><w:t>正文</w:t></w:r>";
        var after = "<w:pPr><w:jc w:val='center'/><w:rPr>" + Chinese
            + "</w:rPr></w:pPr><w:r><w:rPr>" + Chinese + "</w:rPr><w:t>正文</w:t></w:r>";
        Assert.True(Equivalent(before, after));
        var existing = Text.Replace("</w:rPr>", "<w:lang w:val='en-US'/></w:rPr>");
        Assert.True(Equivalent(existing, existing.Replace("w:val='en-US'", "w:val='en-US' w:eastAsia='zh-CN'")));
        Assert.False(Equivalent(Text, Text));
    }

    [Theory]
    [InlineData("zh-TW")]
    [InlineData("ja-JP")]
    [InlineData("ko-KR")]
    public void ExplicitLanguageAssignmentsCannotBeChangedOrRemoved(string language)
    {
        var before = Text.Replace("</w:rPr>", $"<w:lang w:eastAsia='{language}'/></w:rPr>");
        Assert.False(Equivalent(before, before.Replace(language, "zh-CN")));
        Assert.False(Equivalent(before, Text));
        Assert.False(Equivalent(Text, before));
    }

    [Theory]
    [InlineData("正文", "不同")]
    [InlineData("16", "18")]
    public void FormulaTextAndFontSizesRemainStrict(string from, string to)
    {
        var after = Text.Replace("</w:rPr>", Chinese + "</w:rPr>").Replace(from, to);
        Assert.False(Equivalent(Text, after));
    }

    [Fact]
    public void ProofingExceptionDoesNotApplyToEnglishOrHideOtherProperties()
    {
        var ascii = Text.Replace("正文", "plain");
        Assert.False(Equivalent(ascii, ascii.Replace("</w:rPr>", Chinese + "</w:rPr>")));
        Assert.False(Equivalent(Text, Text.Replace("</w:rPr>", Chinese + "<w:b/></w:rPr>")));
        var before = "<w:r><w:t>正文</w:t></w:r>";
        Assert.False(Equivalent(before, "<w:r><w:rPr>" + Chinese + "<w:b/></w:rPr><w:t>正文</w:t></w:r>"));
        var field = Text + "<w:r><w:instrText> SEQ Example </w:instrText></w:r><w:r><w:t>1</w:t></w:r>";
        Assert.False(Equivalent(field, field.Replace("</w:rPr>", Chinese + "</w:rPr>").Replace(">1<", ">2<")));
    }

    [Fact]
    public void HeaderRevisionStampsAreVolatileButFieldResultsAndFormattingAreNot()
    {
        var xml = XElement.Parse("<w:hdr xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>"
            + "<w:p w:rsidRDefault='00000001'><w:r><w:rPr><w:sz w:val='14'/></w:rPr>"
            + "<w:instrText> PAGE </w:instrText><w:t>1</w:t></w:r></w:p></w:hdr>");
        var value = xml.ToString(SaveOptions.DisableFormatting);
        var canonical = WordRecoveryXmlEquivalence.NormalizeHeaderFooterPayload(xml);
        Assert.Equal(canonical, WordRecoveryXmlEquivalence.NormalizeHeaderFooterPayload(XElement.Parse(value.Replace("00000001", "00000002"))));
        Assert.NotEqual(canonical, WordRecoveryXmlEquivalence.NormalizeHeaderFooterPayload(XElement.Parse(value.Replace(">1<", ">2<"))));
        Assert.NotEqual(canonical, WordRecoveryXmlEquivalence.NormalizeHeaderFooterPayload(XElement.Parse(value.Replace("14", "16"))));
    }
}

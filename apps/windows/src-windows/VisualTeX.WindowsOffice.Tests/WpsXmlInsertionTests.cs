using System.Xml.Linq;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

// The Word field and package rules applied before Range.InsertXML in WPS Office.
public sealed class WpsXmlInsertionTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private static string Package(string body) =>
        "<?xml version=\"1.0\" standalone=\"yes\"?>"
        + "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
        + "<pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData>"
        + $"<w:document xmlns:w=\"{W}\" xmlns:m=\"{M}\"><w:body>{body}<w:sectPr/></w:body></w:document>"
        + "</pkg:xmlData></pkg:part></pkg:package>";

    private static XElement LastParagraph(string flatOpc) =>
        XDocument.Parse(flatOpc).Descendants(W + "body").Single().Elements(W + "p").Last();

    [Fact]
    public void ParagraphEndMarkerIsAddedAsTheLastPropertyOfTheLastParagraphOnly()
    {
        var flatOpc = Package(
            "<w:p><w:r><w:t>first</w:t></w:r></w:p>"
            + "<w:p><w:pPr><w:jc w:val=\"center\"/><w:rPr><w:sz w:val=\"28\"/></w:rPr></w:pPr><w:r><w:t>last</w:t></w:r></w:p>");

        var (xml, endsWithParagraph) = WpsXmlInsertion.PrepareForWps(flatOpc, importParagraphEnd: false);

        Assert.True(endsWithParagraph);
        var paragraphs = XDocument.Parse(xml).Descendants(W + "p").ToArray();
        Assert.Null(paragraphs[0].Element(W + "pPr"));
        var properties = paragraphs[1].Element(W + "pPr")!.Elements().Select(element => element.Name.LocalName).ToArray();
        Assert.Equal(new[] { "jc", "rPr", "doNotImportParagraphEnd" }, properties);
        Assert.StartsWith("<?xml", xml, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphEndMarkerIsNotDuplicatedAndIsRemovedWhenTheMarkIsImported()
    {
        var withMarker = Package("<w:p><w:pPr><w:doNotImportParagraphEnd/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>");

        var kept = WpsXmlInsertion.PrepareForWps(withMarker, importParagraphEnd: false).Xml;
        Assert.Single(LastParagraph(kept).Descendants(W + "doNotImportParagraphEnd"));

        var imported = WpsXmlInsertion.PrepareForWps(withMarker, importParagraphEnd: true).Xml;
        Assert.Empty(LastParagraph(imported).Descendants(W + "doNotImportParagraphEnd"));
    }

    [Fact]
    public void PackageEndingInATableIsLeftAlone()
    {
        var flatOpc = Package("<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>");

        var (xml, endsWithParagraph) = WpsXmlInsertion.PrepareForWps(flatOpc, importParagraphEnd: false);

        Assert.False(endsWithParagraph);
        Assert.Equal(flatOpc, xml);
    }

    [Fact]
    public void ExplicitEmbedFieldAroundAnOleObjectIsRemoved()
    {
        var flatOpc = Package(
            "<w:p>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + "<w:r><w:instrText xml:space=\"preserve\"> EMBED Equation.DSMT4 </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:bookmarkStart w:id=\"0\" w:name=\"keep\"/>"
            + "<w:r><w:object/></w:r>"
            + "<w:bookmarkEnd w:id=\"0\"/>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + "<w:r><w:instrText xml:space=\"preserve\"> SEQ MTEqn \\c </w:instrText></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:t>1</w:t></w:r>"
            + "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "</w:p>");

        var paragraph = LastParagraph(WpsXmlInsertion.PrepareForWps(flatOpc, importParagraphEnd: true).Xml);

        Assert.DoesNotContain(paragraph.Descendants(W + "instrText"), instruction => instruction.Value.Contains("EMBED"));
        Assert.Single(paragraph.Descendants(W + "object"));
        Assert.Single(paragraph.Descendants(W + "bookmarkStart"));
        // The SEQ field keeps all four of its field characters.
        Assert.Equal(3, paragraph.Descendants(W + "fldChar").Count());
        Assert.Single(paragraph.Descendants(W + "instrText"));
    }

    [Fact]
    public void InlineSentinelFollowsTheLastEquationOfTheLastParagraph()
    {
        var flatOpc = Package(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"b\"/><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath><w:bookmarkEnd w:id=\"1\"/></w:p>");

        var paragraph = LastParagraph(WordOmmlConverter.AppendWpsInlineSentinel(flatOpc));

        var afterMath = paragraph.Element(M + "oMath")!.ElementsAfterSelf().First();
        Assert.Equal(W + "r", afterMath.Name);
        Assert.Equal("\u2060", afterMath.Element(W + "t")!.Value);
    }

    [Fact]
    public void NumberedEquationFieldInstructionsBecomeNormalTextMathWithPrefixes()
    {
        var numbered = WordOmmlConverter.BuildWordNativeNumberedOmml(
            $"<m:oMath xmlns:m=\"{M}\"><m:r><m:t>x</m:t></m:r></m:oMath>", "Equation", null, 2, ".",
            restartSequenceValue: 1);
        // The builder spells its new elements with default-namespace declarations.
        Assert.Contains("instrText", numbered, StringComparison.Ordinal);

        var xml = WpsXmlInsertion.PrepareForWps(Package($"<w:p>{numbered}</w:p>"), importParagraphEnd: false).Xml;

        Assert.DoesNotContain("instrText", xml, StringComparison.Ordinal);
        Assert.DoesNotContain("xmlns=", xml, StringComparison.Ordinal);
        var equation = LastParagraph(xml).Element(M + "oMath")!;
        var instructions = equation.Descendants(M + "r")
            .Where(run => run.Element(M + "rPr")?.Element(M + "nor") is not null
                && run.Element(M + "t")!.Value.Contains("MERGEFORMAT"))
            .Select(run => run.Element(M + "t")!.Value)
            .ToArray();
        Assert.Equal(
            new[] { " STYLEREF 2 \\w \\* MERGEFORMAT ", " SEQ Equation \\r 1 \\s 2 \\* ARABIC \\* MERGEFORMAT " },
            instructions);
        // m:rPr comes before w:rPr in a math run.
        var sequenceRun = equation.Descendants(M + "t").Single(text => text.Value.Contains("SEQ")).Parent!;
        Assert.Equal(new[] { "rPr", "rPr", "t" }, sequenceRun.Elements().Select(element => element.Name.LocalName).ToArray());
        Assert.Equal(M + "rPr", sequenceRun.Elements().First().Name);
        // Field characters and results stay as they were.
        Assert.Equal(6, equation.Descendants(W + "fldChar").Count());
        Assert.Contains(equation.Descendants(M + "t"), text => text.Value == "#");
    }

    [Fact]
    public void EquationInstructionReadBackWithU2217IsRepaired()
    {
        var flatOpc = Package(
            "<w:p><m:oMath>"
            + "<m:r><w:fldChar w:fldCharType=\"begin\"/></m:r>"
            + "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t xml:space=\"preserve\"> SEQ Equation \\\u2217 ARABIC </m:t></m:r>"
            + "<m:r><w:fldChar w:fldCharType=\"separate\"/></m:r>"
            + "<m:r><m:t>a\u2217b</m:t></m:r>"
            + "<m:r><w:fldChar w:fldCharType=\"end\"/></m:r>"
            + "<m:r><m:t>c\u2217d</m:t></m:r>"
            + "</m:oMath></w:p>");

        var equation = LastParagraph(WpsXmlInsertion.PrepareForWps(flatOpc, importParagraphEnd: true).Xml).Element(M + "oMath")!;

        var runs = equation.Elements(M + "r").ToArray();
        Assert.Equal(" SEQ Equation \\* ARABIC ", runs[1].Element(M + "t")!.Value);
        Assert.Equal(new[] { "nor" }, runs[1].Element(M + "rPr")!.Elements().Select(element => element.Name.LocalName).ToArray());
        // The field result and the math after the field are not instructions.
        Assert.Equal("a\u2217b", runs[3].Element(M + "t")!.Value);
        Assert.Null(runs[3].Element(M + "rPr"));
        Assert.Equal("c\u2217d", runs[5].Element(M + "t")!.Value);
    }

    [Theory]
    [InlineData("default", "Cambria Math")]
    [InlineData(" Default ", "Cambria Math")]
    [InlineData("", "Cambria Math")]
    [InlineData(null, "Cambria Math")]
    [InlineData("Latin Modern Math", "Latin Modern Math")]
    public void WpsDefaultMathFontPlaceholderIsNotWrittenAsAFont(string? reported, string written) =>
        Assert.Equal(written, WordOmmlConverter.NormalizeMathFontName(reported));

    [Theory]
    [InlineData(" SEQ MTEqn \\h \\* MERGEFORMAT ", true)]
    [InlineData("SEQ VisualTeXEquation \\h", true)]
    [InlineData(" SEQ MTEqn \\c \\* Arabic ", false)]
    [InlineData(" MACROBUTTON MTPlaceRef \\* MERGEFORMAT ", false)]
    [InlineData(null, false)]
    public void HiddenSequenceFieldsAreRecognised(string? instruction, bool hidden) =>
        Assert.Equal(hidden, WordFieldCode.IsHiddenSequence(instruction));

    [Fact]
    public void FieldRunsPackageKeepsTheHostParagraph()
    {
        var runs = new List<XElement> { WpsFieldXml.FieldChar("begin"), WpsFieldXml.Instruction(" GOTOBUTTON b "), WpsFieldXml.FieldChar("end") };

        var paragraph = LastParagraph(WpsFieldXml.BuildRunsFlatOpc(runs));

        Assert.NotNull(paragraph.Element(W + "pPr")!.Element(W + "doNotImportParagraphEnd"));
        Assert.Equal(3, paragraph.Elements(W + "r").Count());
    }
}

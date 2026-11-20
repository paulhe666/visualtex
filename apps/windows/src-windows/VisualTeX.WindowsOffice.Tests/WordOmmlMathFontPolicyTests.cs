using System.Xml.Linq;
using VisualTeX.WindowsOffice.VstoShared;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlMathFontPolicyTests
{
    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace Mn = M;
    private static readonly XNamespace Wn = W;

    private static string Run(string text, string? font, bool normal = false) =>
        "<m:r>"
        + (normal ? "<m:rPr><m:nor/></m:rPr>" : string.Empty)
        + (font is null ? string.Empty : $"<w:rPr><w:rFonts w:ascii=\"{font}\" w:hAnsi=\"{font}\"/></w:rPr>")
        + $"<m:t>{text}</m:t></m:r>";

    private static string Equation(params string[] content) =>
        $"<m:oMath xmlns:m=\"{M}\" xmlns:w=\"{W}\">" + string.Concat(content) + "</m:oMath>";

    private static IEnumerable<(string Text, string? Ascii, string? HAnsi)> Runs(string xml) =>
        XDocument.Parse(xml).Descendants(Mn + "r").Select(run =>
        {
            var fonts = run.Element(Wn + "rPr")?.Element(Wn + "rFonts");
            return ((string)run.Element(Mn + "t")!,
                (string?)fonts?.Attribute(Wn + "ascii"),
                (string?)fonts?.Attribute(Wn + "hAnsi"));
        });

    [Theory]
    [InlineData(null, "follow")]
    [InlineData("", "follow")]
    [InlineData("follow", "follow")]
    [InlineData("unknown", "follow")]
    [InlineData("cambria", "cambria")]
    [InlineData(" STIX ", "stix")]
    [InlineData("latin-modern", "latin-modern")]
    [InlineData("termes", "termes")]
    public void ParsesStoredChoices(string? stored, string expectedId) =>
        Assert.Equal(expectedId, WordOmmlMathFontPolicy.Id(WordOmmlMathFontPolicy.Parse(stored)));

    [Fact]
    public void ChoicesRoundTripAndMapToMathFonts()
    {
        foreach (var choice in WordOmmlMathFontPolicy.Choices)
            Assert.Equal(choice, WordOmmlMathFontPolicy.Parse(WordOmmlMathFontPolicy.Id(choice)));
        Assert.Null(WordOmmlMathFontPolicy.FontFamily(OmmlMathFontChoice.FollowDocument));
        Assert.Equal("Cambria Math", WordOmmlMathFontPolicy.FontFamily(OmmlMathFontChoice.CambriaMath));
        Assert.Equal("STIX Two Math", WordOmmlMathFontPolicy.FontFamily(OmmlMathFontChoice.StixTwoMath));
        // The bundled TrueType build, not the CFF "Latin Modern Math".
        Assert.Equal("VisualTeX Latin Modern Math", WordOmmlMathFontPolicy.FontFamily(OmmlMathFontChoice.LatinModernMath));
        // The bundled TrueType build of TeX Gyre Termes Math.
        Assert.Equal("VisualTeX TeX Gyre Termes Math", WordOmmlMathFontPolicy.FontFamily(OmmlMathFontChoice.TexGyreTermesMath));
        Assert.True(WordOmmlMathFontPolicy.IsMathFontFamily(WordOmmlMathFontPolicy.TexGyreTermesMathFamily));
        Assert.Contains(OmmlMathFontChoice.TexGyreTermesMath, WordOmmlMathFontPolicy.Choices);
    }

    [Fact]
    public void MissingFontsFollowTheDocument()
    {
        Assert.Equal("STIX Two Math",
            WordOmmlMathFontPolicy.EffectiveFontFamily(OmmlMathFontChoice.StixTwoMath, _ => true));
        // Every choice but "follow" names a real font (WPS writes it into the runs).
        foreach (var choice in WordOmmlMathFontPolicy.Choices.Where(choice => choice != OmmlMathFontChoice.FollowDocument))
            Assert.Equal(WordOmmlMathFontPolicy.FontFamily(choice), WordOmmlMathFontPolicy.EffectiveFontFamily(choice, _ => true));
        // A font that is not installed is never applied.
        Assert.Null(WordOmmlMathFontPolicy.EffectiveFontFamily(
            OmmlMathFontChoice.StixTwoMath, family => family != "STIX Two Math"));
        Assert.Null(WordOmmlMathFontPolicy.EffectiveFontFamily(
            OmmlMathFontChoice.FollowDocument, _ => true));
    }

    [Fact]
    public void ForWriteAppliesTheConfiguredFontAndOtherwisePassesThrough()
    {
        var applied = new List<string>();
        try
        {
            WordOmmlMathFontPolicy.Configure(() => null, (_, font) => applied.Add(font));
            Assert.Equal("Cambria Math", WordOmmlMathFontPolicy.ForWrite(null!, "Cambria Math"));
            Assert.Null(WordOmmlMathFontPolicy.ForWrite(null!, null));
            Assert.False(WordOmmlMathFontPolicy.ApplyPreferredToDocument(null!));
            Assert.Empty(applied);

            WordOmmlMathFontPolicy.Configure(() => "STIX Two Math", (_, font) => applied.Add(font));
            Assert.Equal("STIX Two Math", WordOmmlMathFontPolicy.ForWrite(null!, "Cambria Math"));
            Assert.True(WordOmmlMathFontPolicy.ApplyPreferredToDocument(null!));
            Assert.Equal(new[] { "STIX Two Math", "STIX Two Math" }, applied);

            // A failing provider means "follow", never a broken write.
            WordOmmlMathFontPolicy.Configure(() => throw new InvalidOperationException(), (_, font) => applied.Add(font));
            Assert.Equal("Cambria Math", WordOmmlMathFontPolicy.ForWrite(null!, "Cambria Math"));
        }
        finally
        {
            WordOmmlMathFontPolicy.Configure(null, null);
        }
        Assert.Null(WordOmmlMathFontPolicy.PreferredFont);
    }

    [Fact]
    public void RewritesMathRunsAndControlsButKeepsNormalTextFonts()
    {
        var xml = Equation(
            Run("x", "Cambria Math"),
            "<m:f><m:fPr><m:ctrlPr><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/><w:i/></w:rPr></m:ctrlPr></m:fPr>"
            + "<m:num>" + Run("1", "Cambria Math") + "</m:num><m:den>" + Run("2", null) + "</m:den></m:f>",
            Run("d", "Times New Roman", normal: true));

        var result = WordOmmlMathFontPolicy.AlignRunFonts(xml, "STIX Two Math");

        Assert.Equal(1, result.EquationsToAlign);
        Assert.Equal(new[] { "Cambria Math" }, result.EquationFonts.Single());
        Assert.Equal(3, result.RewrittenRuns); // two runs + the fraction control
        Assert.Equal(1, result.KeptNormalTextRuns);
        var runs = Runs(result.Xml).ToDictionary(run => run.Text);
        Assert.Equal(("x", "STIX Two Math", "STIX Two Math"), runs["x"]);
        Assert.Equal(("1", "STIX Two Math", "STIX Two Math"), runs["1"]);
        // A run without explicit fonts already follows the math font: nothing added.
        Assert.Equal(("2", null, null), runs["2"]);
        // m:nor (\mathrm{d}) keeps its text font.
        Assert.Equal(("d", "Times New Roman", "Times New Roman"), runs["d"]);
        var control = XDocument.Parse(result.Xml).Descendants(Mn + "ctrlPr").Single()
            .Element(Wn + "rPr")!.Element(Wn + "rFonts")!;
        Assert.Equal("STIX Two Math", (string?)control.Attribute(Wn + "ascii"));
        Assert.NotNull(XDocument.Parse(result.Xml).Descendants(Wn + "i").SingleOrDefault());
    }

    [Fact]
    public void NormalTextInTheOldMathFontFollowsButBodyFontsSharedWithNormalTextStay()
    {
        // Word writes normal text in the math font: it followed the old math font.
        var mathFontNormal = WordOmmlMathFontPolicy.AlignRunFonts(
            Equation(Run("x", "Cambria Math"), Run("max", "Cambria Math", normal: true)),
            "VisualTeX Latin Modern Math");
        Assert.All(Runs(mathFontNormal.Xml), run => Assert.Equal("VisualTeX Latin Modern Math", run.Ascii));
        Assert.Equal(0, mathFontNormal.KeptNormalTextRuns);

        // A body font on a math run that m:nor text also uses cannot be told apart in
        // Word's Find/Replace, so neither is touched.
        var shared = WordOmmlMathFontPolicy.AlignRunFonts(
            Equation(Run("y", "Times New Roman"), Run("if", "Times New Roman", normal: true), Run("z", "Cambria Math")),
            "STIX Two Math");
        Assert.Equal(new[] { "Cambria Math" }, shared.EquationFonts.Single());
        var runs = Runs(shared.Xml).ToDictionary(run => run.Text);
        Assert.Equal("Times New Roman", runs["y"].Ascii);
        Assert.Equal("Times New Roman", runs["if"].Ascii);
        Assert.Equal("STIX Two Math", runs["z"].Ascii);

        // Explicit m:nor val="0" is a math run.
        var off = WordOmmlMathFontPolicy.AlignRunFonts(
            Equation("<m:r><m:rPr><m:nor m:val=\"0\"/></m:rPr><w:rPr><w:rFonts w:ascii=\"Arial\"/></w:rPr><m:t>q</m:t></m:r>"),
            "Cambria Math");
        Assert.Equal("Cambria Math", Runs(off.Xml).Single().Ascii);
    }

    [Fact]
    public void AlignedEquationsNeedNothing()
    {
        var xml = Equation(Run("x", "cambria math"), Run("y", null), Run("t", "Times New Roman", normal: true));
        var result = WordOmmlMathFontPolicy.AlignRunFonts(xml, "Cambria Math");
        Assert.Equal(0, result.EquationsToAlign);
        Assert.Equal(0, result.RewrittenRuns);
        Assert.Single(result.EquationFonts);
    }

    [Fact]
    public void PackagesListBodyEquationsInOrderWithoutTextBoxesAndKeepNumbering()
    {
        // A numbered display (eqArr + # + SEQ field + VisualTeX bookmarks), an inline
        // equation, and a text-box equation (its own story; excluded here).
        var numbered =
            "<m:oMathPara><m:oMath><m:eqArr><m:e>"
            + Run("E=m", "Cambria Math")
            + "<w:bookmarkStart w:id=\"7\" w:name=\"VTEqNum_abc\"/>"
            + Run("#(", "Cambria Math")
            + "<m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><w:fldChar w:fldCharType=\"begin\"/></m:r>"
            + "<m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><w:instrText xml:space=\"preserve\"> SEQ Equation \\* ARABIC \\* MERGEFORMAT </w:instrText></m:r>"
            + "<m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></m:r>"
            + Run("1", "Cambria Math")
            + "<m:r><w:rPr><w:rFonts w:ascii=\"Cambria Math\" w:hAnsi=\"Cambria Math\"/></w:rPr><w:fldChar w:fldCharType=\"end\"/></m:r>"
            + Run(")", "Cambria Math")
            + "<w:bookmarkEnd w:id=\"7\"/>"
            + "</m:e></m:eqArr></m:oMath></m:oMathPara>";
        var package =
            "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\">"
            + "<pkg:part pkg:name=\"/word/document.xml\"><pkg:xmlData>"
            + $"<w:document xmlns:w=\"{W}\" xmlns:m=\"{M}\"><w:body>"
            + "<w:p>" + numbered + "</w:p>"
            + "<w:p><w:r><w:t>and </w:t></w:r><m:oMath>" + Run("y", null) + "</m:oMath>"
            + "<w:r><w:drawing><w:txbxContent><w:p><m:oMath>" + Run("z", "Cambria Math") + "</m:oMath></w:p></w:txbxContent></w:drawing></w:r></w:p>"
            + "</w:body></w:document></pkg:xmlData></pkg:part>"
            + "<pkg:part pkg:name=\"/word/footnotes.xml\"><pkg:xmlData>"
            + $"<w:footnotes xmlns:w=\"{W}\" xmlns:m=\"{M}\"><w:footnote><w:p><m:oMath>" + Run("f", "Cambria Math") + "</m:oMath></w:p></w:footnote></w:footnotes>"
            + "</pkg:xmlData></pkg:part></pkg:package>";

        var result = WordOmmlMathFontPolicy.AlignRunFonts(package, "STIX Two Math");

        Assert.Equal(2, result.EquationFonts.Count);
        Assert.Equal(new[] { "Cambria Math" }, result.EquationFonts[0]);
        Assert.Empty(result.EquationFonts[1]);
        var after = XDocument.Parse(result.Xml);
        var before = XDocument.Parse(package);
        foreach (var name in new[] { Wn + "fldChar", Wn + "instrText", Wn + "bookmarkStart", Wn + "bookmarkEnd", Mn + "eqArr", Mn + "oMath" })
            Assert.Equal(before.Descendants(name).Count(), after.Descendants(name).Count());
        Assert.Equal(" SEQ Equation \\* ARABIC \\* MERGEFORMAT ", after.Descendants(Wn + "instrText").Single().Value);
        Assert.Equal("VTEqNum_abc", (string?)after.Descendants(Wn + "bookmarkStart").Single().Attribute(Wn + "name"));
        var body = after.Descendants(Wn + "body").Single();
        Assert.All(
            body.Descendants(Mn + "r").Where(run => !run.Ancestors(Wn + "txbxContent").Any())
                .Select(run => run.Element(Wn + "rPr")?.Element(Wn + "rFonts"))
                .Where(fonts => fonts is not null),
            fonts => Assert.Equal("STIX Two Math", (string?)fonts!.Attribute(Wn + "ascii")));
        // The text box and the footnote belong to other stories and are left alone here.
        Assert.Equal("Cambria Math", (string?)body.Descendants(Wn + "txbxContent").Single()
            .Descendants(Wn + "rFonts").Single().Attribute(Wn + "ascii"));
        Assert.Equal("Cambria Math", (string?)after.Descendants(Wn + "footnotes").Single()
            .Descendants(Wn + "rFonts").Single().Attribute(Wn + "ascii"));
    }

    [Theory]
    [InlineData("Cambria Math", true)]
    [InlineData("STIX Two Math", true)]
    [InlineData("VisualTeX Latin Modern Math", true)]
    [InlineData("Times New Roman", false)]
    [InlineData("Mathematica", false)]
    [InlineData("", false)]
    public void RecognizesMathFontFamilies(string family, bool expected) =>
        Assert.Equal(expected, WordOmmlMathFontPolicy.IsMathFontFamily(family));
}

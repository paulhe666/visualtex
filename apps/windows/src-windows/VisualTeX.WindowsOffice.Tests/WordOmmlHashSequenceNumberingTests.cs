using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlHashSequenceNumberingTests
{
    private static readonly XNamespace Math =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private static readonly XNamespace Word =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void BuildImmutableHashSequenceNumberedOmml_UsesDirectSeqAndNoRef()
    {
        var formulaId = Guid.NewGuid();
        var numberName = "VTEqNum_" + formulaId.ToString("N");
        var visibleName = "VTEq_" + formulaId.ToString("N");
        var captionName = "VTEqCap_" + formulaId.ToString("N");

        var numbered = WordOmmlConverter.BuildImmutableHashSequenceNumberedOmml(
            SimpleFormula("x"),
            "VisualTeXEquation",
            numberName,
            visibleName,
            captionName,
            prefix: string.Empty,
            restartHeadingLevel: 0,
            initialSequenceResult: "1");
        var equation = XElement.Parse(numbered);

        Assert.Equal(Math + "oMath", equation.Name);
        var outer = Assert.Single(equation.Elements(Math + "eqArr"));
        var row = Assert.Single(outer.Elements(Math + "e"));
        Assert.Contains(row.Descendants(Math + "t"), text => text.Value == "#");
        Assert.Single(row.Descendants(Math + "d"));

        var instruction = string.Concat(
            row.Descendants(Word + "instrText").Select(node => node.Value));
        Assert.Contains("SEQ VisualTeXEquation", instruction, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\\* ARABIC", instruction, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\\* MERGEFORMAT", instruction, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("REF ", instruction, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("\\r ", instruction, StringComparison.OrdinalIgnoreCase);

        var fieldRuns = row
            .Descendants(Math + "r")
            .Where(run => run.Elements(Word + "fldChar").Any()
                || run.Elements(Word + "instrText").Any())
            .ToArray();
        Assert.NotEmpty(fieldRuns);
        Assert.All(fieldRuns, run =>
            Assert.Null(run.Element(Math + "rPr")?.Element(Math + "nor")));
        var fieldBegin = row
            .Descendants(Word + "fldChar")
            .Single(node => string.Equals(
                (string?)node.Attribute(Word + "fldCharType"),
                "begin",
                StringComparison.OrdinalIgnoreCase));
        Assert.Null(fieldBegin.Attribute(Word + "dirty"));
        Assert.NotNull(fieldBegin.Parent?.Element(Word + "rPr")?.Element(Word + "i"));

        var bookmarkNames = row
            .Descendants(Word + "bookmarkStart")
            .Select(node => (string?)node.Attribute(Word + "name"))
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Cast<string>()
            .ToArray();
        Assert.Contains(numberName, bookmarkNames);
        Assert.Contains(visibleName, bookmarkNames);
        Assert.Contains(captionName, bookmarkNames);
        Assert.Equal(3, bookmarkNames.Distinct(StringComparer.OrdinalIgnoreCase).Count());

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(numbered));
        var semantic = XElement.Parse(
            WordOmmlConverter.StripVisualTeXNativeEquationNumber(numbered));
        Assert.Equal(Math + "oMath", semantic.Name);
        Assert.Empty(semantic.Elements(Math + "eqArr"));
        Assert.Equal("x", string.Concat(semantic.Descendants(Math + "t").Select(node => node.Value)));
        Assert.DoesNotContain("VTEq", semantic.ToString(SaveOptions.DisableFormatting),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void NativeHashSequenceParser_AcceptsWordMergedFormulaTailHashRun()
    {
        var formulaId = Guid.NewGuid();
        var numberName = "VTEqNum_" + formulaId.ToString("N");
        var visibleName = "VTEq_" + formulaId.ToString("N");
        var captionName = "VTEqCap_" + formulaId.ToString("N");
        var numbered = XElement.Parse(
            WordOmmlConverter.BuildImmutableHashSequenceNumberedOmml(
                SimpleFormula("x"),
                "VisualTeXEquation",
                numberName,
                visibleName,
                captionName,
                prefix: string.Empty,
                restartHeadingLevel: 0,
                initialSequenceResult: "1"));
        var row = numbered.Descendants(Math + "e").First();
        var children = row.Elements().ToArray();
        var formulaRun = children[0];
        var hashRun = children[1];
        Assert.Equal("#", string.Concat(hashRun.Elements(Math + "t").Select(node => node.Value)));
        formulaRun.Elements(Math + "t").Last().Value += "#";
        hashRun.Remove();
        var wordNormalized = numbered.ToString(SaveOptions.DisableFormatting);

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(wordNormalized));
        var semantic = XElement.Parse(
            WordOmmlConverter.StripVisualTeXNativeEquationNumber(wordNormalized));
        Assert.Equal("x", string.Concat(semantic.Descendants(Math + "t").Select(node => node.Value)));
    }

    [Fact]
    public void BuildImmutableHashSequenceNumberedOmml_ChapterPrefixIsInsideAliases()
    {
        var formulaId = Guid.NewGuid();
        var numberName = "VTEqNum_" + formulaId.ToString("N");
        var visibleName = "VTEq_" + formulaId.ToString("N");
        var captionName = "VTEqCap_" + formulaId.ToString("N");

        var numbered = WordOmmlConverter.BuildImmutableHashSequenceNumberedOmml(
            SimpleFormula("y"),
            "VisualTeXEquation",
            numberName,
            visibleName,
            captionName,
            prefix: "2.",
            restartHeadingLevel: 1,
            initialSequenceResult: "3");
        var equation = XElement.Parse(numbered);
        var instruction = string.Concat(
            equation.Descendants(Word + "instrText").Select(node => node.Value));
        Assert.Contains("SEQ VisualTeXEquation", instruction, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("\\s 1", instruction, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("REF ", instruction, StringComparison.OrdinalIgnoreCase);

        var numberStart = equation
            .Descendants(Word + "bookmarkStart")
            .Single(node => string.Equals(
                (string?)node.Attribute(Word + "name"),
                numberName,
                StringComparison.OrdinalIgnoreCase));
        var slot = numberStart.Parent;
        Assert.NotNull(slot);
        var slotText = string.Concat(slot!.Descendants(Math + "t").Select(node => node.Value));
        Assert.StartsWith("2.", slotText, StringComparison.Ordinal);
        Assert.EndsWith("3", slotText, StringComparison.Ordinal);

        var starts = slot.Elements(Word + "bookmarkStart").ToArray();
        var ends = slot.Elements(Word + "bookmarkEnd").ToArray();
        Assert.Equal(3, starts.Length);
        Assert.Equal(3, ends.Length);
        Assert.Equal(
            new[] { "31803", "31802", "31801" },
            ends.Select(node => (string?)node.Attribute(Word + "id")).ToArray());
    }

    [Fact]
    public void HasVisualTeXNativeEquationNumber_DoesNotTreatArbitraryEqArrAsNumbering()
    {
        var arbitrary = new XElement(
            Math + "oMath",
            new XElement(
                Math + "eqArr",
                new XElement(
                    Math + "e",
                    new XElement(Math + "r", new XElement(Math + "t", "x=y")))))
            .ToString(SaveOptions.DisableFormatting);

        Assert.False(WordOmmlConverter.HasVisualTeXNativeEquationNumber(arbitrary));
        Assert.Equal(
            arbitrary,
            WordOmmlConverter.StripVisualTeXNativeEquationNumber(arbitrary));
    }

    [Fact]
    public void BuildImmutableHashSequenceNumberedOmml_RejectsInvalidBookmarkName()
    {
        Assert.Throws<ArgumentException>(() =>
            WordOmmlConverter.BuildImmutableHashSequenceNumberedOmml(
                SimpleFormula("z"),
                "VisualTeXEquation",
                "VTEqNum-invalid-name",
                "VTEq_" + Guid.NewGuid().ToString("N"),
                "VTEqCap_" + Guid.NewGuid().ToString("N"),
                prefix: string.Empty,
                restartHeadingLevel: 0,
                initialSequenceResult: "1"));
    }

    private static string SimpleFormula(string value) =>
        new XElement(
            Math + "oMath",
            new XElement(Math + "r", new XElement(Math + "t", value)))
        .ToString(SaveOptions.DisableFormatting);
}

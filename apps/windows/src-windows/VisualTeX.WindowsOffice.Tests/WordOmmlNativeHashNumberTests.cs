using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlNativeHashNumberTests
{
    private const string FormulaGuidN = "0123456789abcdef0123456789abcdef";

    [Fact]
    public void DetectsAndStripsWordNormalizedDirectSequenceWrapper()
    {
        var numbered = BuildNumberedOmml(
            sequenceCode: @" SEQ VisualTeXEquation \s 1 \* ARABIC ",
            includeVisualTeXBookmarks: true);

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(numbered));

        var semantic = WordOmmlConverter.StripVisualTeXNativeEquationNumber(numbered);
        var xml = XElement.Parse(semantic, LoadOptions.PreserveWhitespace);
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        Assert.Equal(math + "oMath", xml.Name);
        Assert.Empty(xml.Descendants(math + "eqArr"));
        Assert.Contains(xml.Descendants(math + "t"), node => node.Value == "x");
        Assert.DoesNotContain(xml.Descendants(math + "t"), node => node.Value == "#");
        Assert.DoesNotContain(xml.Descendants(math + "t"), node =>
            node.Value.Contains("SEQ VisualTeXEquation", StringComparison.OrdinalIgnoreCase));
        Assert.Empty(xml.Descendants(word + "bookmarkStart"));
    }

    [Fact]
    public void DoesNotTreatUnmanagedEquationArrayAsVisualTeXNumbering()
    {
        var userEquationArray = BuildNumberedOmml(
            sequenceCode: @" SEQ UserSequence \* ARABIC ",
            includeVisualTeXBookmarks: false);

        Assert.False(
            WordOmmlConverter.HasVisualTeXNativeEquationNumber(userEquationArray));
        Assert.Equal(
            WordOmmlConverter.ExtractSingleOMath(userEquationArray),
            WordOmmlConverter.StripVisualTeXNativeEquationNumber(userEquationArray));
    }

    [Fact]
    public void DirectSequenceWrapperContainsNoMathematicalReference()
    {
        var numbered = BuildNumberedOmml(
            sequenceCode: @" SEQ VisualTeXEquation \* ARABIC ",
            includeVisualTeXBookmarks: true);
        var xml = XElement.Parse(numbered, LoadOptions.PreserveWhitespace);
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        Assert.Single(xml.Elements(math + "eqArr"));
        Assert.Contains(xml.Descendants(math + "t"), node => node.Value == "#");
        var fieldText = string.Concat(
            xml.Descendants(word + "instrText").Select(node => node.Value)
                .Concat(xml.Descendants(math + "t").Select(node => node.Value)));
        Assert.Contains("SEQ VisualTeXEquation", fieldText, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("REF VTEqNum_", fieldText, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(xml.Descendants(word + "bookmarkStart"), node =>
            string.Equals(
                (string?)node.Attribute(word + "name"),
                "VTEqNum_" + FormulaGuidN,
                StringComparison.OrdinalIgnoreCase));
    }

    private static string BuildNumberedOmml(
        string sequenceCode,
        bool includeVisualTeXBookmarks)
    {
        XNamespace math =
            "http://schemas.openxmlformats.org/officeDocument/2006/math";
        XNamespace word =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        XElement FieldRun(XElement content) =>
            new(
                math + "r",
                new XElement(math + "rPr", new XElement(math + "nor")),
                new XElement(word + "rPr", new XElement(word + "noProof")),
                content);

        var number = new List<object>();
        if (includeVisualTeXBookmarks)
        {
            number.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "10"),
                new XAttribute(word + "name", "VTEq_" + FormulaGuidN)));
            number.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "11"),
                new XAttribute(word + "name", "VTEqCap_" + FormulaGuidN)));
            number.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "12"),
                new XAttribute(word + "name", "VTEqNum_" + FormulaGuidN)));
            number.Add(FieldRun(new XElement(math + "t", "2.")));
        }
        number.Add(FieldRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "begin"))));
        // Real Word commonly normalizes mathematical field instructions from
        // w:instrText into m:t. This is the representation that previously caused
        // the production health parser to concatenate field code into the number.
        number.Add(FieldRun(new XElement(
            math + "t",
            new XAttribute(XNamespace.Xml + "space", "preserve"),
            sequenceCode)));
        number.Add(FieldRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "separate"))));
        number.Add(FieldRun(new XElement(math + "t", "3")));
        number.Add(FieldRun(new XElement(
            word + "fldChar",
            new XAttribute(word + "fldCharType", "end"))));
        if (includeVisualTeXBookmarks)
        {
            number.Add(new XElement(
                word + "bookmarkEnd",
                new XAttribute(word + "id", "12")));
            number.Add(new XElement(
                word + "bookmarkEnd",
                new XAttribute(word + "id", "11")));
            number.Add(new XElement(
                word + "bookmarkEnd",
                new XAttribute(word + "id", "10")));
        }

        return new XElement(
                math + "oMath",
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "e",
                        new XElement(math + "r", new XElement(math + "t", "x")),
                        new XElement(math + "r", new XElement(math + "t", "#")),
                        new XElement(
                            math + "d",
                            new XElement(math + "e", number)))))
            .ToString(SaveOptions.DisableFormatting);
    }
}

using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlNativeSequenceXmlTests
{
    private const string MathNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string WordNamespace =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void DirectVisualTeXSequenceWrapperIsRecognizedAndStrippedAtomically()
    {
        var formulaId = "00112233445566778899aabbccddeeff";
        var omml = BuildDirectSequenceOmml(formulaId, includeVisualTeXAliases: true);

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(omml));

        var semantic = WordOmmlConverter.StripVisualTeXNativeEquationNumber(omml);
        var document = XDocument.Parse(semantic, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;

        Assert.Empty(document.Descendants(math + "eqArr"));
        Assert.Equal("x+1", string.Concat(document.Descendants(math + "t").Select(node => node.Value)));
        Assert.Empty(document.Descendants(word + "fldChar"));
        Assert.Empty(document.Descendants(word + "instrText"));
    }

    [Fact]
    public void WordNormalizedMathRunSequenceInstructionRemainsRecognized()
    {
        var omml = BuildDirectSequenceOmml(
            "00112233445566778899aabbccddeeff",
            includeVisualTeXAliases: true);
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var instruction = document.Descendants(word + "instrText").Single();
        instruction.ReplaceWith(
            new XElement(
                math + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                instruction.Value));
        var normalized = document.Root!.ToString(SaveOptions.DisableFormatting);

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(normalized));
        var semantic = WordOmmlConverter.StripVisualTeXNativeEquationNumber(normalized);
        var semanticDocument = XDocument.Parse(semantic, LoadOptions.PreserveWhitespace);
        Assert.Empty(semanticDocument.Descendants(math + "eqArr"));
        Assert.Equal(
            "x+1",
            string.Concat(semanticDocument.Descendants(math + "t").Select(node => node.Value)));
    }

    [Fact]
    public void SplitMathematicalSequenceInstructionRemainsMigrationInput()
    {
        var omml = BuildDirectSequenceOmml(
            "00112233445566778899aabbccddeeff",
            includeVisualTeXAliases: true);
        var document = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var instructionRun = document
            .Descendants(word + "instrText")
            .Single()
            .Parent!;
        instructionRun.ReplaceWith(
            FieldRun(math, word, new XElement(math + "t", " SEQ ")),
            FieldRun(math, word, new XElement(math + "t", "VisualTeXEquation ")),
            FieldRun(math, word, new XElement(math + "t", "\\* ARABIC ")));
        var split = document.Root!.ToString(SaveOptions.DisableFormatting);

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(split));
        var semantic = WordOmmlConverter.StripVisualTeXNativeEquationNumber(split);
        var semanticDocument = XDocument.Parse(semantic, LoadOptions.PreserveWhitespace);
        Assert.Empty(semanticDocument.Descendants(math + "eqArr"));
        Assert.Equal(
            "x+1",
            string.Concat(semanticDocument.Descendants(math + "t").Select(node => node.Value)));
    }

    [Fact]
    public void UnrelatedEquationArrayWithSequenceIsNotClaimedByVisualTeX()
    {
        var omml = BuildDirectSequenceOmml(
            "00112233445566778899aabbccddeeff",
            includeVisualTeXAliases: false);

        Assert.False(WordOmmlConverter.HasVisualTeXNativeEquationNumber(omml));
        Assert.Equal(
            WordOmmlConverter.ExtractSingleOMath(omml),
            WordOmmlConverter.StripVisualTeXNativeEquationNumber(omml));
    }

    [Fact]
    public void ManagedStripCanRecoverDirectSequenceAfterAllAliasesAreLost()
    {
        var omml = BuildDirectSequenceOmml(
            "00112233445566778899aabbccddeeff",
            includeVisualTeXAliases: false);

        // The public/strict path must leave an unowned equation array alone.
        Assert.False(WordOmmlConverter.HasVisualTeXNativeEquationNumber(omml));
        Assert.Equal(
            WordOmmlConverter.ExtractSingleOMath(omml),
            WordOmmlConverter.StripVisualTeXNativeEquationNumber(omml));

        // A caller that already proved VTOMML ownership may recover the semantic
        // body and atomically rebuild all three aliases.
        var semantic = WordOmmlConverter
            .StripManagedVisualTeXNativeEquationNumber(omml);
        var document = XDocument.Parse(semantic, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        Assert.Empty(document.Descendants(math + "eqArr"));
        Assert.Equal(
            "x+1",
            string.Concat(document.Descendants(math + "t").Select(node => node.Value)));
    }

    [Fact]
    public void DocumentFieldCodeExtractionReassemblesSplitRefAndMathSeqRuns()
    {
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var root = new XElement(
            "root",
            new XAttribute(XNamespace.Xmlns + "m", MathNamespace),
            new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "begin"))),
            FieldRun(
                math,
                word,
                new XElement(word + "instrText", " REF ")),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "instrText",
                    "\"VTEqNum_00112233445566778899aabbccddeeff\" \\h")),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "separate"))),
            FieldRun(math, word, new XElement(word + "t", "RESULT_REF_ONLY")),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "end"))),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "begin"))),
            FieldRun(math, word, new XElement(math + "t", " SEQ ")),
            FieldRun(math, word, new XElement(math + "t", "VisualTeXEquation ")),
            FieldRun(math, word, new XElement(math + "t", "\\* ARABIC ")),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "separate"))),
            FieldRun(math, word, new XElement(math + "t", "RESULT_SEQ_ONLY")),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "end"))));

        var method = typeof(WordEquationNumbering).GetMethod(
            "ExtractFieldCodesFromWordOpenXml",
            System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException(
                "The document field-code XML extractor is missing.");
        var codes = (IReadOnlyList<string>?)method.Invoke(
            null,
            new object[] { root.ToString(SaveOptions.DisableFormatting) })
            ?? throw new InvalidOperationException(
                "The document field-code XML extractor returned null.");

        Assert.Contains(
            codes,
            code => code.Contains(
                "REF \"VTEqNum_00112233445566778899aabbccddeeff\" \\h",
                StringComparison.Ordinal));
        Assert.Contains(
            codes,
            code => code.Contains(
                "SEQ VisualTeXEquation \\* ARABIC",
                StringComparison.Ordinal));
        Assert.DoesNotContain(
            codes,
            code => code.Contains("RESULT_REF_ONLY", StringComparison.Ordinal));
        Assert.DoesNotContain(
            codes,
            code => code.Contains("RESULT_SEQ_ONLY", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void VisibleNumberExtractionSkipsSequenceInstructionRuns(
        bool wordUsesInstructionText)
    {
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var instructionContent = wordUsesInstructionText
            ? new XElement(
                word + "instrText",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                " SEQ VisualTeXEquation \\s 1 \\* ARABIC ")
            : new XElement(
                math + "t",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                " SEQ VisualTeXEquation \\s 1 \\* ARABIC ");
        var segment = string.Concat(
            MathRun(math, "2.").ToString(SaveOptions.DisableFormatting),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "begin")))
                .ToString(SaveOptions.DisableFormatting),
            FieldRun(math, word, instructionContent)
                .ToString(SaveOptions.DisableFormatting),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "separate")))
                .ToString(SaveOptions.DisableFormatting),
            FieldRun(math, word, new XElement(math + "t", "3"))
                .ToString(SaveOptions.DisableFormatting),
            FieldRun(
                math,
                word,
                new XElement(
                    word + "fldChar",
                    new XAttribute(word + "fldCharType", "end")))
                .ToString(SaveOptions.DisableFormatting));

        var method = typeof(WordEquationNumbering).GetMethod(
            "ExtractVisibleEquationNumberTextFromOpenXmlSegment",
            System.Reflection.BindingFlags.NonPublic
            | System.Reflection.BindingFlags.Static)
            ?? throw new InvalidOperationException(
                "The visible native-number XML extractor is missing.");
        var result = (string?)method.Invoke(null, new object[] { segment });

        Assert.Equal("2.3", result);
    }

    [Fact]
    public void RetiredMathematicalRefWrapperRemainsMigrationInput()
    {
        var formulaId = "00112233445566778899aabbccddeeff";
        var omml = BuildRetiredReferenceOmml(formulaId);

        Assert.True(WordOmmlConverter.HasVisualTeXNativeEquationNumber(omml));
        var semantic = WordOmmlConverter.StripVisualTeXNativeEquationNumber(omml);
        var document = XDocument.Parse(semantic, LoadOptions.PreserveWhitespace);
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;

        Assert.Empty(document.Descendants(math + "eqArr"));
        Assert.Equal("x+1", string.Concat(document.Descendants(math + "t").Select(node => node.Value)));
        Assert.Empty(document.Descendants(word + "instrText"));
    }

    private static string BuildDirectSequenceOmml(
        string normalizedFormulaId,
        bool includeVisualTeXAliases)
    {
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var numberName = "VTEqNum_" + normalizedFormulaId;
        var visibleName = "VTEq_" + normalizedFormulaId;
        var captionName = "VTEqCap_" + normalizedFormulaId;

        var numberContent = new List<object>();
        if (includeVisualTeXAliases)
        {
            numberContent.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "11"),
                new XAttribute(word + "name", numberName)));
            numberContent.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "12"),
                new XAttribute(word + "name", visibleName)));
            numberContent.Add(new XElement(
                word + "bookmarkStart",
                new XAttribute(word + "id", "13"),
                new XAttribute(word + "name", captionName)));
        }
        numberContent.Add(FieldRun(
            math,
            word,
            new XElement(
                word + "fldChar",
                new XAttribute(word + "fldCharType", "begin"))));
        numberContent.Add(FieldRun(
            math,
            word,
            new XElement(
                word + "instrText",
                new XAttribute(XNamespace.Xml + "space", "preserve"),
                " SEQ VisualTeXEquation \\* ARABIC ")));
        numberContent.Add(FieldRun(
            math,
            word,
            new XElement(
                word + "fldChar",
                new XAttribute(word + "fldCharType", "separate"))));
        numberContent.Add(FieldRun(math, word, new XElement(math + "t", "1")));
        numberContent.Add(FieldRun(
            math,
            word,
            new XElement(
                word + "fldChar",
                new XAttribute(word + "fldCharType", "end"))));
        if (includeVisualTeXAliases)
        {
            numberContent.Add(new XElement(word + "bookmarkEnd", new XAttribute(word + "id", "13")));
            numberContent.Add(new XElement(word + "bookmarkEnd", new XAttribute(word + "id", "12")));
            numberContent.Add(new XElement(word + "bookmarkEnd", new XAttribute(word + "id", "11")));
        }

        return new XElement(
                math + "oMath",
                new XAttribute(XNamespace.Xmlns + "m", MathNamespace),
                new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "e",
                        MathRun(math, "x"),
                        MathRun(math, "+"),
                        MathRun(math, "1"),
                        MathRun(math, "#"),
                        new XElement(
                            math + "d",
                            new XElement(math + "e", numberContent)))))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static string BuildRetiredReferenceOmml(string normalizedFormulaId)
    {
        var math = (XNamespace)MathNamespace;
        var word = (XNamespace)WordNamespace;
        var numberName = "VTEqNum_" + normalizedFormulaId;
        return new XElement(
                math + "oMath",
                new XAttribute(XNamespace.Xmlns + "m", MathNamespace),
                new XAttribute(XNamespace.Xmlns + "w", WordNamespace),
                new XElement(
                    math + "eqArr",
                    new XElement(
                        math + "e",
                        MathRun(math, "x"),
                        MathRun(math, "+"),
                        MathRun(math, "1"),
                        MathRun(math, "#"),
                        new XElement(
                            math + "d",
                            new XElement(
                                math + "e",
                                FieldRun(
                                    math,
                                    word,
                                    new XElement(
                                        word + "fldChar",
                                        new XAttribute(word + "fldCharType", "begin"))),
                                FieldRun(
                                    math,
                                    word,
                                    new XElement(
                                        word + "instrText",
                                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                                        $" REF {numberName} \\h ")),
                                FieldRun(
                                    math,
                                    word,
                                    new XElement(
                                        word + "fldChar",
                                        new XAttribute(word + "fldCharType", "separate"))),
                                FieldRun(math, word, new XElement(math + "t", "1")),
                                FieldRun(
                                    math,
                                    word,
                                    new XElement(
                                        word + "fldChar",
                                        new XAttribute(word + "fldCharType", "end"))))))))
            .ToString(SaveOptions.DisableFormatting);
    }

    private static XElement MathRun(XNamespace math, string text) =>
        new(math + "r", new XElement(math + "t", text));

    private static XElement FieldRun(
        XNamespace math,
        XNamespace word,
        XElement content) =>
        new(
            math + "r",
            new XElement(math + "rPr", new XElement(math + "nor")),
            new XElement(word + "rPr", new XElement(word + "noProof")),
            content);
}

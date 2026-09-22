using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordOmmlStructuralOrderTests
{
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string P = "http://www.w3.org/1998/Math/MathML";
    private static string Wrap(string content) => $"<m:oMath xmlns:m=\"{M}\">{content}</m:oMath>";
    private static string Run(string value) => $"<m:r><m:t>{value}</m:t></m:r>";
    private static string Pre(string basis) => "<m:sPre><m:e>"+basis+"</m:e><m:sub>"+Run("6")+"</m:sub><m:sup>"+Run("14")+"</m:sup></m:sPre>";

    [Theory]
    [InlineData("msup", "sup", "sub")]
    [InlineData("msub", "sub", "sup")]
    public void SinglePrescriptSuppressesAbsentHalfWithoutAddingText(string tag,string used,string absent)
    {
        var input=$"<math xmlns=\"{P}\"><{tag}><mrow/><mn>14</mn></{tag}><mi>C</mi></math>";
        var xml=WordOmmlConverter.TransformMathMlToOmml(input);
        var document=XDocument.Parse(xml);
        var script=document.Descendants(M+"sPre").Single();
        Assert.Equal("14",script.Element(M+used)!.Value);
        var slot=script.Element(M+absent)!;
        Assert.Empty(slot.Descendants(M+"t"));
        var phantom=Assert.Single(slot.Elements(M+"phant"));
        Assert.Equal("0",(string?)phantom.Element(M+"phantPr")!.Element(M+"show")!.Attribute(M+"val"));
        Assert.All(new[]{"zeroAsc","zeroDesc","zeroWid"},name=>Assert.Equal("1",(string?)phantom.Element(M+"phantPr")!.Element(M+name)!.Attribute(M+"val")));
        WordOmmlConverter.ValidateMaterializedOmml(xml);
        var reverse=WordOmmlConverter.TransformOmmlToMathMl(xml,false);
        Assert.DoesNotContain("mphantom",reverse,StringComparison.Ordinal);
        Assert.Equal(xml,WordOmmlConverter.NormalizeOptionalPrescriptSlots(xml));
    }

    [Fact]
    public void GeneratedPrescriptUsesWordNativeArgumentOrder()
    {
        var input = $"<math xmlns=\"{P}\"><msubsup><mrow/><mn>6</mn><mn>14</mn></msubsup><mi>C</mi></math>";
        var result = XDocument.Parse(WordOmmlConverter.TransformMathMlToOmml(input));
        var script = result.Descendants(M+"sPre").Single();
        Assert.Equal(new[]{"sub","sup","e"}, script.Elements().Where(x=>!x.Name.LocalName.EndsWith("Pr")).Select(x=>x.Name.LocalName));
        Assert.Equal("6",script.Element(M+"sub")!.Value);
        Assert.Equal("14",script.Element(M+"sup")!.Value);
        Assert.Equal("C",script.Element(M+"e")!.Value);
    }

    [Fact]
    public void GeneratedOrderMatchesActualWordRegressionWithoutWeakeningFingerprint()
    {
        // Recorded from normal Ribbon redraw, Word 16.0.14334.20848, 2026-09-19.
        var prepared = Wrap(Pre(Run("C")));
        var actual = Wrap("<m:sPre><m:sPrePr/><m:sub>"+Run("6")+"</m:sub><m:sup>"+Run("14")+"</m:sup><m:e>"+Run("C")+"</m:e></m:sPre>");
        var normalized = WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(prepared);
        Assert.Equal(WordOmmlConverter.ComputeOmmlFingerprint(actual),
            WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(normalized,actual));
        Assert.Throws<InvalidDataException>(()=>WordOmmlConverter.ComputeVerifiedMaterializedOmmlFingerprint(normalized,actual.Replace(">6<",">7<")));
    }

    [Theory]
    [InlineData("f","num","den")]
    [InlineData("rad","deg","e")]
    [InlineData("sSub","e","sub")]
    [InlineData("sSup","e","sup")]
    public void ArgumentGrammarDoesNotChangeRoleValues(string owner,string first,string second)
    {
        var xml=Wrap($"<m:{owner}><m:{second}>{Run("y")}</m:{second}><m:{first}>{Run("x")}</m:{first}></m:{owner}>");
        var result=XDocument.Parse(WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(xml));
        var element=result.Descendants(M+owner).Single();
        Assert.Equal(new[]{first,second},element.Elements().Select(x=>x.Name.LocalName));
        Assert.Equal("x",element.Element(M+first)!.Value);Assert.Equal("y",element.Element(M+second)!.Value);
    }

    [Fact]
    public void NestedPrescriptsAreNormalizedBottomUp()
    {
        var output=WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(Wrap(Pre(Pre(Run("C")))));
        var document=XDocument.Parse(output);
        Assert.Equal(2,document.Descendants(M+"sPre").Count());
        Assert.All(document.Descendants(M+"sPre"),p=>Assert.Equal(new[]{"sub","sup","e"},p.Elements().Select(x=>x.Name.LocalName)));
        Assert.Equal(output,WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(output));
    }

    [Fact]
    public void DuplicateArgumentsAreNotCollapsedToMakeTheEquationPass()
    {
        var input=Wrap(Pre(Run("C")).Replace("</m:sPre>","<m:sub>"+Run("7")+"</m:sub></m:sPre>"));
        var output=WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(input);
        Assert.Equal(2,XDocument.Parse(output).Descendants(M+"sPre").Single().Elements(M+"sub").Count());
        Assert.Throws<InvalidDataException>(()=>WordOmmlConverter.ValidateMaterializedOmml(output));
    }

    [Fact]
    public void UnknownChildCannotBeDiscardedDuringArgumentOrdering()
    {
        var input=Wrap(Pre(Run("C")).Replace("</m:sPre>","<x:extension xmlns:x=\"urn:test\">retain</x:extension></m:sPre>"));
        var output=WordOmmlConverter.NormalizeGeneratedOmmlArgumentOrder(input);
        Assert.Contains("retain",output,StringComparison.Ordinal);
        Assert.Equal("e",XDocument.Parse(output).Descendants(M+"sPre").Single().Elements().First().Name.LocalName);
    }
}

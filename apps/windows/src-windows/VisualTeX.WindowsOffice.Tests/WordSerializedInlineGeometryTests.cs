using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class WordSerializedInlineGeometryTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
    private static readonly XNamespace O = "urn:schemas-microsoft-com:office:office";

    private static FormulaMetadata Metadata(string id) => new()
    {
        FormulaId = id, Latex = "x", DisplayMode = "inline", CodeFormat = "latex",
        FontSizePt = 12, RenderFontSizePt = 12,
        Lines = new() { new() { Id = Guid.NewGuid().ToString(), Latex = "x" } },
    };

    private static XElement Object(FormulaMetadata metadata, string style, string progId = FormulaOleContract.ProgId)
    {
        var id = "shape_" + metadata.FormulaId;
        return new XElement(W + "object",
            new XElement(V + "shape", new XAttribute("id", id),
                new XAttribute("alt", FormulaMetadataCodec.Encode(metadata)), new XAttribute("style", style)),
            new XElement(O + "OLEObject", new XAttribute("ProgID", progId), new XAttribute("ShapeID", id)));
    }

    [Fact]
    public void SerializedGeometryResolvesTheExactFormulaRatherThanItsNeighbor()
    {
        var metadata = Metadata(Guid.NewGuid().ToString());
        var xml = new XElement(W + "document",
            Object(Metadata(Guid.NewGuid().ToString()), "width:500pt;height:100pt"),
            Object(metadata, "height:11.75pt;width:90.05pt"));
        Assert.True(WordInlineObjectGeometry.TryReadSerializedVisualTeXSize(xml.ToString(), metadata.FormulaId, out var width, out var height));
        Assert.Equal(90.05f, width);
        Assert.Equal(11.75f, height);
    }

    [Theory]
    [InlineData("width:-1pt;height:10pt")]
    [InlineData("width:50%;height:10pt")]
    [InlineData("width:30pt;width:40pt;height:10pt")]
    [InlineData("width:30pt")]
    [InlineData("width:NaNpt;height:10pt")]
    public void UnusableOrAmbiguousDimensionsAreNotGuessed(string style)
    {
        var metadata = Metadata(Guid.NewGuid().ToString());
        Assert.False(WordInlineObjectGeometry.TryReadSerializedVisualTeXSize(
            new XElement(W + "document", Object(metadata, style)).ToString(), metadata.FormulaId, out _, out _));
    }

    [Fact]
    public void ARepeatedFormulaIdentityDoesNotSelectTheFirstOccurrence()
    {
        var metadata = Metadata(Guid.NewGuid().ToString());
        var xml = new XElement(W + "document", Object(metadata, "width:30pt;height:10pt"), Object(metadata, "width:60pt;height:20pt"));
        Assert.False(WordInlineObjectGeometry.TryReadSerializedVisualTeXSize(xml.ToString(), metadata.FormulaId, out _, out _));
    }

    [Fact]
    public void MathTypePresentationIsNeverTreatedAsVisualTeXGeometry()
    {
        var metadata = Metadata(Guid.NewGuid().ToString());
        var xml = new XElement(W + "document", Object(metadata, "width:30pt;height:10pt", "Equation.DSMT4"));
        Assert.False(WordInlineObjectGeometry.TryReadSerializedVisualTeXSize(xml.ToString(), metadata.FormulaId, out _, out _));
    }

    [Theory]
    [InlineData(92.55f, 10.3f, 12f)]
    [InlineData(70f, 7f, 12f)]
    [InlineData(180.1f, 23.5f, 24f)]
    [InlineData(45.025f, 5.875f, 6f)]
    public void OnlyUniformHostResizingCanChangeTheInferredSemanticSize(float width, float height, float expected)
    {
        var metadata = Metadata(Guid.NewGuid().ToString());
        metadata.WordInlineOleWidthPt = 90.05;
        metadata.WordInlineOleHeightPt = 11.75;
        Assert.Equal(expected, FormulaFontSize.InferOleFontSize(width, height, metadata));
    }
}

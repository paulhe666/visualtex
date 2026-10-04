using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathMlExplicitFenceContractTests
{
    private static readonly XNamespace P = "http://www.w3.org/1998/Math/MathML";
    private static string Source(string body) => $"<math xmlns=\"{P}\">{body}</math>";
    private static string Row(string body, string open = "⟮", string close = "⟯",
        string openAttributes = "", string closeAttributes = "") =>
        $"<mrow data-mjx-texclass=\"INNER\"><mo data-mjx-texclass=\"OPEN\" {openAttributes}>{open}</mo>"
        + body + $"<mo data-mjx-texclass=\"CLOSE\" {closeAttributes}>{close}</mo></mrow>";
    private static string Signature(string source) => MathTypeMtefCodec.SemanticSignature(source);
    private static string RoundTripOmml(string source) =>
        WordOmmlConverter.TransformOmmlToMathMl(WordOmmlConverter.TransformMathMlToOmml(source), false);
    private static string F09 => Source(Row("<mi>x</mi>") + "<mo>+</mo>"
        + Row("<mi>y</mi>", "⏐", "", "fence=\"true\" stretchy=\"true\" symmetric=\"true\"",
            "fence=\"true\" stretchy=\"true\" symmetric=\"true\""));

    [Fact]
    public void OriginalF09SourceSignatureAgreesWithNativeOmmlFenceStructure()
    {
        var source = F09;
        Assert.Equal(Signature(source), Signature(RoundTripOmml(source)));
    }

    [Fact]
    public void UnsupportedMtefTemplateStillLosesFenceRoleAndFailsStrictComparison()
    {
        var source = F09;
        var encoded = MathTypeMtefCodec.CreateEquationNative(source, false);
        var recovered = MathTypeMtefCodec.ReadEquationNativeMathMl(encoded.EquationNative);
        Assert.NotEqual(Signature(source), Signature(recovered));
    }

    [Fact]
    public void ExplicitPairedFenceIsDistinctFromTheSameFlatLiteralGlyphs()
    {
        Assert.NotEqual(Signature(Source(Row("<mi>x</mi>"))),
            Signature(Source("<mo>⟮</mo><mi>x</mi><mo>⟯</mo>")));
    }

    [Theory]
    [InlineData("fence=\"false\"")]
    [InlineData("stretchy=\"false\"")]
    [InlineData("mathvariant=\"bold\"")]
    [InlineData("mathsize=\"18pt\"")]
    [InlineData("minsize=\"2em\"")]
    [InlineData("data-mjx-alternate=\"1\"")]
    public void IneligibleEndpointDoesNotCollapseIntoTheNewExplicitFenceSignature(string attributes)
    {
        foreach (var onLeft in new[] { true, false })
        {
            var source = Source(Row("<mi>x</mi>", openAttributes: onLeft ? attributes : "",
                closeAttributes: onLeft ? "" : attributes));
            Assert.NotEqual(Signature(Source(Row("<mi>x</mi>"))), Signature(source));
        }
    }

    [Theory]
    [InlineData("⏐", "")]
    [InlineData("", "⟧")]
    public void SharedRecognitionAndOmmlRetainTheExplicitEmptySide(string open, string close)
    {
        var source = Source(Row("<mi>x</mi>", open, close));
        var row = XDocument.Parse(source).Root!.Elements().Single();
        Assert.True(MathMlExplicitFenceContract.TryGetRow(row, out var actualOpen, out var actualClose, out var body));
        Assert.Equal(open, actualOpen);
        Assert.Equal(close, actualClose);
        Assert.Same(row.Elements().ElementAt(1), Assert.Single(body));
        var projected = XDocument.Parse(WordOmmlConverter.NormalizeFencedMathMlTables(source));
        var fence = Assert.Single(projected.Descendants(P + "mfenced"));
        Assert.Equal(open, (string?)fence.Attribute("open"));
        Assert.Equal(close, (string?)fence.Attribute("close"));
        Assert.Equal(Signature(source), Signature(RoundTripOmml(source)));
        // Existing content signatures flatten one-sided fences. Structural
        // assertions above carry the empty-side evidence; do not claim otherwise.
        var flat = Source("<mo>" + open + "</mo><mi>x</mi><mo>" + close + "</mo>");
        Assert.Equal(Signature(source), Signature(flat));
    }

    [Theory]
    [InlineData("")]
    [InlineData("http://www.w3.org/1998/Math/MathML")]
    public void SharedRecognitionSupportsCodecAndPresentationNamespaces(string ns)
    {
        var row = XElement.Parse(Row("<mi>x</mi>"));
        foreach (var element in row.DescendantsAndSelf())
            element.Name = XNamespace.Get(ns) + element.Name.LocalName;
        Assert.True(MathMlExplicitFenceContract.TryGetRow(row, out var open, out var close, out _));
        Assert.Equal("⟮", open);
        Assert.Equal("⟯", close);
        Assert.Equal(Signature(Source(Row("<mi>x</mi>"))),
            Signature("<math>" + row + "</math>"));
    }

    [Fact]
    public void ForeignOrMixedNamespacesDoNotAcquireExplicitFenceAuthority()
    {
        var row = XElement.Parse(Row("<mi>x</mi>"));
        row.Name = XNamespace.Get("urn:other") + "mrow";
        Assert.False(MathMlExplicitFenceContract.TryGetRow(row, out _, out _, out _));
        row.Name = P + "mrow";
        Assert.False(MathMlExplicitFenceContract.TryGetRow(row, out _, out _, out _));
    }

    [Fact]
    public void NestedFenceAndScriptAttachmentShareTheSameSourceAndOmmlSignatures()
    {
        var source = Source("<msubsup>" + Row("<mi>a</mi><mo>+</mo>" + Row("<mi>b</mi>", "⦃", "⦄"))
            + "<mi>i</mi><mn>2</mn></msubsup>");
        Assert.Equal(Signature(source), Signature(RoundTripOmml(source)));
    }
}

using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypePrescriptRoundTripTests
{
    private const string Ns="http://www.w3.org/1998/Math/MathML";
    private static string Math(string body)=>$"<math xmlns=\"{Ns}\">{body}</math>";
    private static string Multi(string basis,string post,string pre)=>"<mmultiscripts>"+basis+post+"<mprescripts/>"+pre+"</mmultiscripts>";
    private static void AssertRoundTrip(string input)
    {
        var created=MathTypeMtefCodec.CreateEquationNative(input,inline:true);
        var actual=MathTypeMtefCodec.ReadEquationNativeMathMl(created.EquationNative);
        Assert.Equal(MathTypeMtefCodec.SemanticSignature(input),MathTypeMtefCodec.SemanticSignature(actual));
        Assert.Contains("mprescripts",actual,StringComparison.Ordinal);
        var omml=WordOmmlConverter.TransformMathMlToOmml(actual);
        WordOmmlConverter.ValidateMaterializedOmml(omml);
        Assert.NotEmpty(XDocument.Parse(omml).Descendants(XName.Get("sPre","http://schemas.openxmlformats.org/officeDocument/2006/math")));
    }

    [Theory]
    [InlineData("<none/><mn>14</mn>")]
    [InlineData("<mn>6</mn><none/>")]
    [InlineData("<mn>6</mn><mn>14</mn>")]
    public void NativeMtefRetainsWhichSideTheScriptsBelongTo(string pair)=>AssertRoundTrip(Math(Multi("<mi>C</mi>","",pair)));

    [Theory]
    [InlineData("<mi>C</mi>")]
    [InlineData("<mrow><mi>A</mi><mo>+</mo><mi>B</mi></mrow>")]
    [InlineData("<mfrac><mi>x</mi><mi>y</mi></mfrac>")]
    public void PreAndPostScriptsRetainTheEntireBase(string basis)=>AssertRoundTrip(Math(Multi(basis,"<mi>i</mi><mn>2</mn>","<mi>a</mi><mi>b</mi>")));

    [Fact]
    public void MultipleLeftPairsKeepTheirOrder()=>AssertRoundTrip(Math(Multi("<mi>T</mi>","","<mi>a</mi><mi>b</mi><mi>c</mi><mi>d</mi>")));

    [Fact]
    public void ChangingLeftScriptToRightScriptIsNotEquivalent()
    {
        var left=Math(Multi("<mi>C</mi>","","<none/><mn>14</mn>"));
        var right=Math("<msup><mi>C</mi><mn>14</mn></msup>");
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(left),MathTypeMtefCodec.SemanticSignature(right));
    }

    [Fact]
    public void UnpairedAndDuplicatePrescriptSeparatorsAreRejected()
    {
        Assert.Throws<InvalidDataException>(()=>MathTypeMtefCodec.SemanticSignature(Math(Multi("<mi>C</mi>","","<mi>a</mi>"))));
        Assert.Throws<InvalidDataException>(()=>MathTypeMtefCodec.SemanticSignature(Math(Multi("<mi>C</mi>","","<mprescripts/><none/><mi>a</mi>"))));
    }
}

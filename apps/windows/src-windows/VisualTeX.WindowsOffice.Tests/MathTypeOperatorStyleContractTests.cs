using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;
namespace VisualTeX.WindowsOffice.Tests;
public sealed class MathTypeOperatorStyleContractTests
{
    [Theory]
    [InlineData("−")][InlineData("∑")][InlineData("∞")][InlineData("≥")]
    [InlineData("⇐")][InlineData("∓")][InlineData("↦")][InlineData("∥")]
    [InlineData("⪯")][InlineData("≍")][InlineData("♢")][InlineData("⨿")]
    public void ExplicitOperatorStyleCannotDisappearAtAnyCharacterBoundary(string glyph)
    {
        foreach(var style in new[]{"bold","italic","bold-italic"})
        foreach(var kind in new[]{"mi","mo"})
        foreach(var accent in new[]{false,true})
        {
            var token=new XElement(kind,new XAttribute("mathvariant",style),glyph);
            var body=accent ? new XElement("mover",new XAttribute("accent","true"),token,new XElement("mo","¨")) : token;
            var source=new XElement("math",body).ToString();
            var plain=source.Replace("mathvariant=\""+style+"\"","mathvariant=\"normal\"");
            Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),MathTypeMtefCodec.SemanticSignature(plain));
            var native=MathTypeMtefCodec.CreateEquationNativeAtFontSize(source,false,12);
            var actual=MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),MathTypeMtefCodec.SemanticSignature(actual));
            var basis=XElement.Parse(actual).Descendants().First(x=>x.Name.LocalName is "mi" or "mo");
            Assert.Equal(style,(string?)basis.Attribute("mathvariant"));
            var rewritten=MathTypeMtefCodec.RewriteEquationNative(native.EquationNative,actual,false);
            Assert.Equal(MathTypeMtefCodec.SemanticSignature(source),
                MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewritten.EquationNative)));
        }
    }
}

using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeCharacterEmissionContractTests
{
    [Fact]
    public void EverySingleCharacterDecorationRetainsItsNativeEncodingBytes()
    {
        // An independent byte-level invariant, not a comparison of two calls to
        // SemanticSignature: decorating a CHAR must preserve its entire font
        // definition, MTCode and encoded position, changing only the EMBELL bit
        // and appending its four-byte attachment. This also covers glyph aliases
        // whose rendered code differs from their public Unicode code point.
        var identifiers = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789αηΓϱϵϰȷℏℜℑℓ";
        var variants = new[]{ "", "bold", "bold-italic", "script", "bold-script", "fraktur", "bold-fraktur", "sans-serif", "monospace" };
        var decorations = new[]{ ("˙",2),("¨",3),("~",8),("^",9),("→",11),("←",12),("↔",13),("¯",17) };
        foreach(var glyph in identifiers)
        foreach(var variant in variants)
        {
            var kind=char.IsDigit(glyph)?"mn":"mi";
            var token=new XElement(kind,variant.Length==0?null:new XAttribute("mathvariant",variant),glyph.ToString());
            var bare=Native(token).Mtef;
            foreach(var mark in decorations)
            {
                // Fixed decorations: a bare ~ ^ → ¯ is wide by the MathML dictionary.
                var decorated=Native(new XElement("mover",new XAttribute("accent","true"),
                    new XElement(token),new XElement("mo",new XAttribute("stretchy","false"),mark.Item1))).Mtef;
                Assert.True(bare.Length+4 == decorated.Length,
                    $"{glyph}/{variant}/{mark.Item1}: bare {bare.Length}, decorated {decorated.Length}");
                Assert.Equal(new byte[]{6,0,(byte)mark.Item2,0,0,0},decorated.Skip(decorated.Length-6).ToArray());
                var changes=Enumerable.Range(0,bare.Length-2).Where(i=>bare[i]!=decorated[i]).ToArray();
                var changed=Assert.Single(changes);
                Assert.True(changed>0 && bare[changed-1]==2,$"{glyph}/{variant}/{mark.Item1} changed a non-CHAR option");
                Assert.Equal((byte)(bare[changed]|1),decorated[changed]);
            }
        }
    }

    [Theory]
    [InlineData("∑")][InlineData("∫")][InlineData("∏")][InlineData("⋃")][InlineData("∩")]
    public void OperatorDecorationsAndLimitsKeepDifferentNativeStructures(string glyph)
    {
        var body="<mo>"+glyph+"</mo><mo>¨</mo>";
        var accent="<math><mover accent='true'>"+body+"</mover></math>";
        var limit="<math><mover accent='false'>"+body+"</mover></math>";
        RoundTrip(accent); RoundTrip(limit);
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(accent),MathTypeMtefCodec.SemanticSignature(limit));
    }

    [Theory]
    [InlineData("ℏ","bold-italic")][InlineData("ȷ","bold-italic")]
    [InlineData("ℜ","bold")][InlineData("ℑ","bold")]
    public void ActualFrontendSpecialGlyphStylesSurviveNestedDecoration(string glyph,string variant)
    {
        var token=new XElement("mi",new XAttribute("mathvariant",variant),glyph);
        var source=new XElement("math",new XElement("mover",new XAttribute("accent","true"),
            new XElement("mover",new XAttribute("accent","true"),token,new XElement("mo","¨")),new XElement("mo","^"))).ToString();
        RoundTrip(source);
        var downgraded=source.Replace(variant,"normal");
        Assert.NotEqual(MathTypeMtefCodec.SemanticSignature(source),MathTypeMtefCodec.SemanticSignature(downgraded));
    }

    [Theory]
    [InlineData("msup")][InlineData("msub")][InlineData("msubsup")]
    public void EveryScriptKindKeepsTextAndNumberNeighborsOutsideItsBase(string kind)
    {
        foreach(var token in new[]{"<mi mathvariant='normal'>e</mi>","<mtext>ab</mtext>","<mn>23</mn>"})
        foreach(var prefix in new[]{"<mtext>%</mtext>","<mn>100</mn>","<mtext>PRE #</mtext>"})
        {
            var source="<math>"+prefix+"<"+kind+">"+token+"<mn>2</mn>"+(kind=="msubsup"?"<mn>3</mn>":"")+"</"+kind+"><mo>+</mo><mn>7</mn></math>";
            RoundTrip(source);
        }
    }

    private static MathTypeMtefCodec.RewriteResult Native(XElement body) =>
        MathTypeMtefCodec.CreateEquationNativeAtFontSize(new XElement("math",body).ToString(),false,12);

    private static void RoundTrip(string source)
    {
        var native=MathTypeMtefCodec.CreateEquationNativeAtFontSize(source,false,12);
        var expected=MathTypeMtefCodec.SemanticSignature(source);
        Assert.Equal(expected,MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative)));
        var rewrite=MathTypeMtefCodec.RewriteEquationNative(native.EquationNative,source,false);
        Assert.Equal(expected,MathTypeMtefCodec.SemanticSignature(MathTypeMtefCodec.ReadEquationNativeMathMl(rewrite.EquationNative)));
    }
}

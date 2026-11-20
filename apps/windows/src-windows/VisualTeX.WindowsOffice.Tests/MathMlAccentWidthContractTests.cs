using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public class MathMlAccentWidthContractTests
{
    [Theory]
    [InlineData("^","true","widehat")]
    [InlineData("^","false","hat")]
    [InlineData("^",null,"widehat")]
    [InlineData("~","true","widetilde")]
    [InlineData("~","false","tilde")]
    [InlineData("~",null,"widetilde")]
    [InlineData("→","true","overrightarrow")]
    [InlineData("→","false","vec")]
    [InlineData("→",null,"overrightarrow")]
    public void RecoversDeclaredOrDictionaryAccentWidth(string mark, string? stretchy, string command)
    {
        var attribute=stretchy is null ? "" : $" stretchy='{stretchy}'";
        var source=$"<math><mover><mrow><mi>x</mi><mi>y</mi></mrow><mo{attribute}>{mark}</mo></mover></math>";
        Assert.Equal("\\"+command+"{xy}",MathMlToLatexConverter.Convert(source));
    }

    [Theory]
    [InlineData("x","normal","^",9)]
    [InlineData("Γ","bold","^",9)]
    [InlineData("ℓ","italic","^",9)]
    [InlineData("x","normal","~",8)]
    [InlineData("Γ","bold","~",8)]
    [InlineData("ℓ","italic","~",8)]
    public void ExplicitFixedAccentPreservesBareCharacterBytes(string glyph, string variant, string mark, int code)
    {
        var token=new XElement("mi",new XAttribute("mathvariant",variant),glyph);
        byte[] Native(XElement body)=>MathTypeMtefCodec.CreateEquationNative(new XElement("math",body).ToString(),false).Mtef;
        var bare=Native(new XElement(token));
        var fixedAccent=Native(new XElement("mover",new XAttribute("accent","true"),new XElement(token),
            new XElement("mo",new XAttribute("stretchy","false"),mark)));
        Assert.Equal(bare.Length+4,fixedAccent.Length);
        Assert.Equal(new byte[]{6,0,(byte)code,0,0,0},fixedAccent.Skip(fixedAccent.Length-6).ToArray());
        var changed=Assert.Single(Enumerable.Range(0,bare.Length-2).Where(i=>bare[i]!=fixedAccent[i]));
        Assert.Equal((byte)(bare[changed]|1),fixedAccent[changed]);
        var wideAccent=Native(new XElement("mover",new XAttribute("accent","true"),new XElement(token),
            new XElement("mo",new XAttribute("stretchy","true"),mark)));
        Assert.False(fixedAccent.SequenceEqual(wideAccent));
    }

    [Theory]
    [InlineData("^",false,false)]
    [InlineData("^",false,true)]
    [InlineData("^",true,false)]
    [InlineData("^",true,true)]
    [InlineData("~",false,false)]
    [InlineData("~",false,true)]
    [InlineData("~",true,false)]
    [InlineData("~",true,true)]
    public void NativeRecordsRetainFixedVersusExpandingGeometry(string mark, bool stretchy, bool group)
    {
        var body=group?"<mrow><mi>x</mi><mi>y</mi></mrow>":"<mi>x</mi>";
        var source=$"<math><mover>{body}<mo stretchy='{stretchy.ToString().ToLowerInvariant()}'>{mark}</mo></mover></math>";
        var native=MathTypeMtefCodec.CreateEquationNative(source,false);
        var actual=XElement.Parse(MathTypeMtefCodec.ReadEquationNativeMathMl(native.EquationNative));
        var accent=Assert.Single(actual.Descendants().Where(e=>e.Name.LocalName=="mover"));
        var token=accent.Elements().Last();
        Assert.Equal(mark,token.Value);
        Assert.Equal(stretchy.ToString().ToLowerInvariant(),(string?)token.Attribute("stretchy"));
    }
}

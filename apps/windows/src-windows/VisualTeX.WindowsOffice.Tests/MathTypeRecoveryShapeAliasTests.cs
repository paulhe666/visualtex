using System.Xml.Linq;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeRecoveryShapeAliasTests
{
    private static readonly XNamespace W="http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace V="urn:schemas-microsoft-com:vml";
    private static readonly XNamespace O="urn:schemas-microsoft-com:office:office";
    private static readonly XNamespace P="http://schemas.microsoft.com/office/2006/xmlPackage";
    private static XDocument Fixture(string a="_a",string b="_b",string sa="_s1",string sb="_s2",bool refAlias=false)
    {
        XElement Shape(string id,string alias,int width)=>new(W+"r",new XElement(W+"object",
            new XElement(V+"shape",new XAttribute("id",id),new XAttribute(O+"spid",alias),new XAttribute("style","width:"+width+"pt;height:14pt")),
            new XElement(O+"OLEObject",new XAttribute("ShapeID",refAlias?alias:id),new XAttribute("ProgID","Equation.DSMT4"))));
        return new XDocument(new XElement(P+"package",new XElement(P+"part",new XAttribute(P+"name","/word/document.xml"),
            new XElement(P+"xmlData",new XElement(W+"document",new XElement(W+"body",new XElement(W+"p",Shape(a,sa,27),Shape(b,sb,40))))))));
    }
    private static string Signature(XDocument document)=>WordLocalEditSnapshot.Signature(document.ToString(SaveOptions.DisableFormatting));

    [Fact]
    public void ShapeIdAndOfficeSpidAreTwoAliasesOfTheSamePhysicalShape()
    {
        var input=Fixture();var untouched=input.ToString();
        Assert.Equal(Signature(input),Signature(Fixture("_newA","_newB","_s2","_s1")));
        Assert.Equal(untouched,input.ToString());
    }
    [Fact]
    public void OleReferenceMayUseEitherProvenAliasWithoutChangingItsOwner()
    {
        Assert.Equal(Signature(Fixture()),Signature(Fixture("_newA","_newB","_s2","_s1",true)));
    }
    [Fact]
    public void IdenticalIdAndSpidOnOneOwnerAreNotAmbiguous()
    {
        Assert.Equal(Signature(Fixture()),Signature(Fixture("_one","_two","_one","_two",true)));
    }
    [Theory]
    [InlineData("reference")]
    [InlineData("width")]
    [InlineData("class")]
    public void RealOwnerGeometryAndClassChangesRemainDetectable(string change)
    {
        var altered=Fixture();
        if(change=="reference") altered.Descendants(O+"OLEObject").First().SetAttributeValue("ShapeID","_b");
        else if(change=="width") altered.Descendants(V+"shape").First().SetAttributeValue("style","width:28pt;height:14pt");
        else altered.Descendants(O+"OLEObject").First().SetAttributeValue("ProgID","Other.Server");
        Assert.NotEqual(Signature(Fixture()),Signature(altered));
    }
    [Theory]
    [InlineData("_s1","_s1")]
    [InlineData("_b","_s2")]
    public void AliasesCannotNameTwoDifferentPhysicalShapes(string firstAlias,string secondAlias)
    {
        Assert.Throws<InvalidDataException>(()=>Signature(Fixture(sa:firstAlias,sb:secondAlias)));
    }
}

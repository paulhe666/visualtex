using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Xunit;
namespace VisualTeX.WindowsOffice.Tests;
public sealed class MathTypeNativeGlyphStyleRenderTests
{
    [Fact]
    public async Task ActualNativeRendererRetainsExplicitCharacterStyleWhenEnabled()
    {
        if(Environment.GetEnvironmentVariable("VISUALTEX_RUN_PRIVATE_MATHTYPE_ACCEPTANCE")!="1") return;
        Assert.True(NativeRenderClient.IsAvailable,"The existing private native-render service is required.");
        var root=Environment.GetEnvironmentVariable("VISUALTEX_CHARACTER_RENDER_ARTIFACT_DIR")
            ?? throw new InvalidOperationException("An explicit owned evidence directory is required.");
        Directory.CreateDirectory(root);
        var request=new NativeRenderRequest();
        var inputs=new Dictionary<string,(string Xml,bool Bold,bool Italic)>();
        foreach(var glyph in new[]{"−","∑","∞","≥","⇐","∓","↦","∥","⪯","≍","♢","⨿","∫","∮"})
        foreach(var style in new[]{"bold","italic","bold-italic"})
        foreach(var accent in new[]{false,true})
        {
            var id=$"u{(int)glyph[0]:X4}-{style}-"+(accent?"accent":"bare");
            var token=new XElement("mo",new XAttribute("mathvariant",style),glyph);
            var body=accent ? new XElement("mover",new XAttribute("accent","true"),token,new XElement("mo","¨")) : token;
            var source=new XElement("math",body).ToString();
            var native=MathTypeMtefCodec.CreateEquationNativeAtFontSize(source,false,12);
            request.Items.Add(new NativeRenderInput{Id=id,Mtef=native.Mtef});
            inputs.Add(id,(source,style.Contains("bold"),style.Contains("italic")));
            File.WriteAllBytes(Path.Combine(root,id+".mtef"),native.Mtef);
        }
        using var timeout=new CancellationTokenSource(TimeSpan.FromSeconds(75));
        var response=await NativeRenderClient.RenderAsync(request,timeout.Token);
        Assert.Equal(string.Empty,response.Error);
        Assert.Equal(request.Items.Count,response.Items.Count);
        var observations=new List<object>();
        foreach(var item in response.Items)
        {
            File.WriteAllBytes(Path.Combine(root,item.Id+".wmf"),item.Wmf);
            var fonts=Fonts(item.Wmf).ToArray();
            observations.Add(new{item.Id,item.WidthPt,item.HeightPt,Fonts=fonts,TextRuns=MathTypeNativeWmfEvidence.TextRuns(item.Wmf),Source=inputs[item.Id].Xml});
        }
        File.WriteAllText(Path.Combine(root,"fonts.json"),JsonSerializer.Serialize(observations,new JsonSerializerOptions{WriteIndented=true}));
        // A declared font that is never selected (or belongs only to the accent)
        // cannot prove the base glyph's style. Track GDI object allocation and
        // SELECTOBJECT through each actual text draw, independently of MTEF.
        foreach(var item in response.Items.Where(item => item.Id.EndsWith("-bare",StringComparison.Ordinal)))
        {
            var expected=inputs[item.Id];
            var plain=MathTypeNativeWmfEvidence.TextRuns(item.Wmf).SelectMany(r => r.Codes.Where(c => c != 32).Select(c => (r.Font,Code:c))).ToArray();
            Assert.NotEmpty(plain);
            Assert.All(plain,g => {
                Assert.Equal(expected.Bold,g.Font.Weight>=600);
                Assert.Equal(expected.Italic,g.Font.Italic);
            });
            var decorated=response.Items.Single(candidate => candidate.Id==item.Id.Replace("-bare","-accent"));
            var actual=MathTypeNativeWmfEvidence.TextRuns(decorated.Wmf).SelectMany(r => r.Codes.Where(c => c != 32).Select(c => (r.Font,Code:c))).ToArray();
            var next=0;
            foreach(var glyph in actual)
                if(next<plain.Length && glyph==plain[next]) next++;
            Assert.True(next==plain.Length,item.Id+" lost/re-encoded a base glyph or its SELECTED font under the accent.");
        }
    }

    private sealed record FontRecord(string Name,int Weight,bool Italic);
    private static IEnumerable<FontRecord> Fonts(byte[] data)
    {
        var start=data.Length>=22 && BitConverter.ToUInt32(data,0)==0x9AC6CDD7 ? 22 : 0;
        if(data.Length<start+18) throw new InvalidDataException("Truncated WMF header.");
        var cursor=start+BitConverter.ToUInt16(data,start+2)*2;
        while(cursor+6<=data.Length)
        {
            var size=checked((int)BitConverter.ToUInt32(data,cursor)*2);
            if(size<6 || cursor+size>data.Length) throw new InvalidDataException("Invalid WMF record extent.");
            var kind=BitConverter.ToUInt16(data,cursor+4);
            if(kind==0x02FB && size>=24)
            {
                var end=cursor+24;
                while(end<cursor+size && data[end]!=0) end++;
                yield return new FontRecord(Encoding.ASCII.GetString(data,cursor+24,end-cursor-24),
                    BitConverter.ToInt16(data,cursor+14),data[cursor+16]!=0);
            }
            cursor+=size;
            if(kind==0) yield break;
        }
    }
}

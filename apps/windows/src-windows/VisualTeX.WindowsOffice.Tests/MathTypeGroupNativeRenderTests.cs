using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Xunit;

namespace VisualTeX.WindowsOffice.Tests;

public sealed class MathTypeGroupNativeRenderTests
{
    [Fact]
    public async Task NativeGroupsRetainEveryBaseGlyphWhenEnabled()
    {
        if(Environment.GetEnvironmentVariable("VISUALTEX_RUN_GROUP_RENDER_ACCEPTANCE")!="1") return;
        var root=Environment.GetEnvironmentVariable("VISUALTEX_CHARACTER_RENDER_ARTIFACT_DIR")
            ?? throw new InvalidOperationException("An owned evidence directory is required.");
        Directory.CreateDirectory(root);
        var rows=MathTypeGroupAccentContractTests.Groups().ToArray();
        var bases=rows.Select(row=>(string)row[0]).Distinct().ToArray();
        var baselines=new Dictionary<string,IReadOnlyList<MathTypeNativeWmfEvidence.Run>>();
        var baseHeights=new Dictionary<string,double>();
        async Task<NativeRenderResponse> Render(NativeRenderRequest request)
        {
            using var timeout=new CancellationTokenSource(TimeSpan.FromSeconds(70));
            var response=await NativeRenderClient.RenderAsync(request,timeout.Token);
            Assert.Equal(string.Empty,response.Error); Assert.Equal(request.Items.Count,response.Items.Count);
            foreach(var item in response.Items) {
                Assert.True(item.WidthPt>0 && item.HeightPt>0,item.Id+" has empty geometry.");
                File.WriteAllBytes(Path.Combine(root,item.Id+".wmf"),item.Wmf);
            }
            return response;
        }
        NativeRenderInput Input(string id,string source)
        {
            var native=MathTypeMtefCodec.CreateEquationNativeAtFontSize(source,false,12);
            File.WriteAllBytes(Path.Combine(root,id+".mtef"),native.Mtef);
            return new NativeRenderInput{Id=id,Mtef=native.Mtef};
        }
        var control=new NativeRenderRequest();
        for(var index=0;index<bases.Length;index++) control.Items.Add(Input("base-"+index,"<math>"+bases[index]+"</math>"));
        control.Items.Add(Input("root-no-color-control","<math><mroot><mi>x</mi><mn>3</mn></mroot></math>"));
        control.Items.Add(Input("empty-token-control","<math><mo>{</mo><mi>x</mi><mo/></math>"));
        var baseResult=await Render(control);
        for(var index=0;index<bases.Length;index++) {
            var text=MathTypeNativeWmfEvidence.TextRuns(baseResult.Items.Single(x=>x.Id=="base-"+index).Wmf);
            Assert.NotEmpty(text); baselines.Add(bases[index],text);
            baseHeights.Add(bases[index],baseResult.Items.Single(x=>x.Id=="base-"+index).HeightPt);
        }
        var observations=new List<object>();
        var completed=0;
        foreach(var batch in rows.Chunk(13))
        {
            var request=new NativeRenderRequest();
            var inputs=new Dictionary<string,(string Source,string Base)>();
            foreach(var row in batch)
            {
                var basis=(string)row[0]; var mark=(string)row[1]; var nested=(bool)row[2];
                var accent=new XElement("mover",new XAttribute("accent","true"),XElement.Parse(basis),new XElement("mo",mark));
                var body=nested ? new XElement("msub",new XElement("mover",new XAttribute("accent","true"),accent,new XElement("mo","^")),new XElement("mn","2")) : accent;
                var source=new XElement("math",body).ToString(); var id="group-"+(completed+inputs.Count);
                request.Items.Add(Input(id,source)); inputs.Add(id,(source,basis));
            }
            var response=await Render(request);
            foreach(var actual in response.Items)
            {
                var input=inputs[actual.Id]; var runs=MathTypeNativeWmfEvidence.TextRuns(actual.Wmf);
                Assert.True(actual.HeightPt>baseHeights[input.Base]+0.25,actual.Id+" has no independent accent extent above its base.");
                var expected=baselines[input.Base].SelectMany(r=>r.Codes.Select(code=>(r.Font,code))).GroupBy(x=>x).ToDictionary(g=>g.Key,g=>g.Count());
                var observed=runs.SelectMany(r=>r.Codes.Select(code=>(r.Font,code))).GroupBy(x=>x).ToDictionary(g=>g.Key,g=>g.Count());
                foreach(var glyph in expected) Assert.True(observed.GetValueOrDefault(glyph.Key)>=glyph.Value,actual.Id+" lost a base glyph or selected font.");
                observations.Add(new{actual.Id,input.Source,actual.WidthPt,actual.HeightPt,Runs=runs,BaseGlyphsRetained=true});
            }
            completed+=batch.Length;
            File.WriteAllText(Path.Combine(root,"progress.json"),JsonSerializer.Serialize(observations,new JsonSerializerOptions{WriteIndented=true}));
        }
        Assert.Equal(156,completed);
    }
}

using System.Diagnostics;
using System.Reflection;
using System.Text.Json;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunOwnedFeedbackAcceptance(string root)
    {
        var input = Environment.GetEnvironmentVariable("VISUALTEX_FEEDBACK_DOCUMENT")
            ?? throw new InvalidOperationException("Explicit saved input copy required.");
        using var corpusJson = JsonDocument.Parse(File.ReadAllText(
            Environment.GetEnvironmentVariable("VISUALTEX_PERF_CORPUS")!));
        var corpus = corpusJson.RootElement.GetProperty("formulas").EnumerateArray().Select(item =>
            new ConversionCorpusItem { Latex=item.GetProperty("latex").GetString()!,
                MathMl=item.GetProperty("mathMl").GetString()!, DisplayMode=item.GetProperty("displayMode").GetString()! }).ToList();
        var mode = Environment.GetEnvironmentVariable("VISUALTEX_FEEDBACK_REDRAW_MODE") ?? FormulaOleContract.MathTypeOleMode;
        var output = Path.Combine(root, "redraw.docx");
        File.Copy(input, output);
        Environment.SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", Path.Combine(root,"trace.log"));
        Environment.SetEnvironmentVariable("VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG", Path.Combine(root,"redraw.log"));
        var rows = new List<object>();
        var errors = new List<object>();
        void Validate(WordFormulaService service, string objectMode, string phase)
        {
            var plan = service.CaptureFormulaFormatConversionPlan(true, objectMode,
                objectMode == FormulaOleContract.MathTypeOleMode ? FormulaOleContract.NativeOleMode : FormulaOleContract.MathTypeOleMode);
            if(plan.Targets.Count!=corpus.Count) errors.Add(new {phase,error="formula count",expected=corpus.Count,actual=plan.Targets.Count});
            for(var i=0;i<Math.Min(plan.Targets.Count,corpus.Count);i++)
            {
                var target=plan.Targets[i];
                var passed=target.DisplayMode==corpus[i].DisplayMode;
                passed &= objectMode==FormulaOleContract.NativeOleMode
                    ? target.Latex==corpus[i].Latex
                    : MathTypeMtefCodec.SemanticSignature(target.SourceMathMl!)==MathTypeMtefCodec.SemanticSignature(corpus[i].MathMl);
                rows.Add(new {phase,index=i+1,passed});
                if(!passed) errors.Add(new {phase,index=i+1,error="content or display differs"});
            }
        }
        void Run(WordPerformanceHost host, string phase, Func<Task> action)
        {
            using var metrics=new WordOperationMetrics();
            var watch=Stopwatch.StartNew();
            var operation=action();
            while(!operation.IsCompleted && watch.Elapsed<TimeSpan.FromMinutes(15))
            { System.Windows.Forms.Application.DoEvents(); Thread.Sleep(20); }
            if(!operation.IsCompleted) throw new TimeoutException("Owned feedback operation exceeded 15 minutes: "+phase);
            operation.GetAwaiter().GetResult();
            rows.Add(new {phase,milliseconds=watch.Elapsed.TotalMilliseconds});
            File.WriteAllText(Path.Combine(root,phase+"-metrics.json"),JsonSerializer.Serialize(metrics.Entries,new JsonSerializerOptions{WriteIndented=true}));
            Console.WriteLine($"[FEEDBACK] {phase} {watch.Elapsed.TotalSeconds:F2}s");
        }
        try
        {
            using(var host=new WordPerformanceHost(output))
            {
                var service=new WordFormulaService(host.Application);
                Run(host,"redraw",()=> (Task)typeof(ThisAddIn).GetMethod("RedrawLatexAsync",BindingFlags.NonPublic|BindingFlags.Instance)!
                    .Invoke(host.AddIn,new object[]{true,mode})!);
                Validate(service,mode,"redraw");
                host.Document.Save();
                if(mode==FormulaOleContract.NativeOleMode && errors.Count==0)
                {
                    File.Copy(output,Path.Combine(root,"visualtex-seed.docx"));
                    Run(host,"conversion",()=> (Task)typeof(ThisAddIn).GetMethod("ConvertFormulaFormatAsync",BindingFlags.NonPublic|BindingFlags.Instance)!
                        .Invoke(host.AddIn,new object[]{true,mode,FormulaOleContract.MathTypeOleMode})!);
                    mode=FormulaOleContract.MathTypeOleMode;
                    Validate(service,mode,"converted");
                    host.Document.Save();
                }
            }
            using(var reopened=new WordPerformanceHost(output))
                Validate(new WordFormulaService(reopened.Application),mode,"reopened");
        }
        catch(Exception error) { errors.Add(new {phase="runner",error=error.ToString()}); }
        finally
        {
            File.WriteAllText(Path.Combine(root,"result.json"),JsonSerializer.Serialize(new {input,mode,count=corpus.Count,rows,errors,
                dll=typeof(ThisAddIn).Assembly.Location,mvid=typeof(ThisAddIn).Assembly.ManifestModule.ModuleVersionId},new JsonSerializerOptions{WriteIndented=true}));
        }
        if(errors.Count!=0) throw new InvalidDataException($"Feedback validation has {errors.Count} failures; see result.json.");
    }
}

using System.Text.Json;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.VstoShared;
using Word = Microsoft.Office.Interop.Word;
namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunMathTypeFormatAuthorityAcceptance(string root)
    {
        var svg = Path.Combine(root,"authority.svg");
        File.WriteAllText(svg,"<svg xmlns='http://www.w3.org/2000/svg' width='80' height='24' viewBox='0 0 80 24'><text x='1' y='18'>x+1</text></svg>");
        var emf = OfficeOlePreview.CreateVectorEmfFromSvg(svg,80,24);
        Word.Application? application=null;
        Word.Document? document=null;
        var rows=new List<object>();
        var previous=WordEquationNumbering.GetDefaultEquationNumberFormatId();
        try
        {
            application=CreateWordApplication(visible:false);
            foreach(var tracking in new[]{false,true})
            foreach(var format in new[]{ EquationNumberFormat.Heading2DotId,EquationNumberFormat.Heading2DashId,
                         EquationNumberFormat.Heading1DotId,EquationNumberFormat.Heading1DashId,EquationNumberFormat.ContinuousId })
            {
                document=application.Documents.Add();
                document.Content.Text="PROTECTED TEXT\r";
                var service=new WordFormulaService(application);
                var oldFormat=format==EquationNumberFormat.ContinuousId ? EquationNumberFormat.Heading1DashId : EquationNumberFormat.ContinuousId;
                WordEquationNumbering.RestoreEquationNumberFormatForConversion(document,oldFormat);
                void Insert()
                {
                    Word.Range? end=null;
                    try
                    {
                        end=document.Range(document.Content.End-1,document.Content.End-1); end.Select();
                        service.InsertMathTypeOle(CreateMathTypeCreateSession(displayMode:"block",numbered:true,
                            latex:"a+b",mathTypeNumberPosition:"right"),FirstNumberedMathMl,emf);
                    }
                    finally{Release(end);}
                }
                Insert();
                var donorLabel=oldFormat==EquationNumberFormat.ContinuousId ? "(1)" : "(0-1)";
                AssertMathTypeNumberTexts(document,donorLabel);
                // Model an older numbered formula pasted into a document with
                // an explicit policy. Changing only the document variable is
                // fixture construction, NOT the product's format-change API.
                WordEquationNumbering.RestoreEquationNumberFormatForConversion(document,format);
                document.TrackRevisions=tracking;
                Insert();
                var prefix=format==EquationNumberFormat.ContinuousId ? "" : format==EquationNumberFormat.Heading1DotId ? "0." :
                    format==EquationNumberFormat.Heading1DashId ? "0-" : format==EquationNumberFormat.Heading2DotId ? "0.0." : "0.0-";
                AssertMathTypeNumberTexts(document,donorLabel,"("+prefix+"2)");
                AssertEqual(format,WordEquationNumberFormatAuthority.Resolve(document),"A donor overrode document authority.");
                File.WriteAllText(Path.Combine(root,format+"-"+tracking+"-before-format.xml"),document.Content.WordOpenXML);
                var recovery=new List<MathTypeLocalEditScope>();
                try
                {
                    var shapes=document.InlineShapes;
                    try
                    {
                        for(var index=1;index<=shapes.Count;index++)
                        {
                            Word.InlineShape? shape=null;Word.Range? range=null;
                            try { shape=shapes[index];range=shape.Range;recovery.Add(MathTypeLocalEditScope.Capture(document,range)); }
                            finally { Release(range);Release(shape); }
                        }
                    }
                    finally { Release(shapes); }
                    service.SetEquationNumberFormat(format);
                    AssertMathTypeNumberTexts(document,"("+prefix+"1)","("+prefix+"2)");
                    AssertTrue(document.Undo(1),"The format transaction did not own one Undo record.");
                    for(var index=0;index<recovery.Count;index++)
                    {
                        try { recovery[index].VerifyRollback(); }
                        catch
                        {
                            File.WriteAllText(Path.Combine(root,format+"-"+tracking+"-scope-"+index+"-before.xml"),recovery[index].ParagraphXml);
                            File.WriteAllText(Path.Combine(root,format+"-"+tracking+"-after-one-undo.xml"),document.Content.WordOpenXML);
                            throw;
                        }
                    }
                    AssertMathTypeNumberTexts(document,donorLabel,"("+prefix+"2)");
                    AssertTrue(document.Redo(1),"The format transaction could not Redo.");
                    AssertMathTypeNumberTexts(document,"("+prefix+"1)","("+prefix+"2)");
                }
                finally { foreach(var scope in recovery) scope.Dispose(); }
                AssertEqual(tracking,document.TrackRevisions,"Format changed tracking mode.");
                var path=Path.Combine(root,format+"-"+tracking+".docx");
                document.SaveAs2(path,Word.WdSaveFormat.wdFormatXMLDocument);
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);Release(document);document=null;
                document=application.Documents.Open(path,ReadOnly:false,Visible:false);
                AssertMathTypeNumberTexts(document,"("+prefix+"1)","("+prefix+"2)");
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);Release(document);document=null;
                rows.Add(new{Format=format,Tracking=tracking,DonorIgnored=true,StrictLocalUndoRedo=true,SaveReopen=true});
                File.WriteAllText(Path.Combine(root,"summary.json"),JsonSerializer.Serialize(rows,new JsonSerializerOptions{WriteIndented=true}));
                Console.WriteLine("PASS donor authority "+format+" tracking="+tracking);
            }
        }
        finally
        {
            try{document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{}
            Release(document);try{QuitWordApplicationIfOwned(application);}catch{}Release(application);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previous);ForceComCleanup();
        }
    }
}

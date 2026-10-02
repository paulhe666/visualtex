using System.Text.Json;
using Word=Microsoft.Office.Interop.Word;
namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunCrossDocumentUndoLifecycle(string root)
    {
        var rows=new List<object>();
        foreach(var mode in new[]{"add-inside","write-staging-inside","close-inside","prepared-outside"})
        {
            Word.Application? app=null;Word.Document? doc=null;Word.Document? stage=null;Word.UndoRecord? undo=null;
            var states=new List<object>();
            try
            {
                app=CreateWordApplication(visible:false);doc=app.Documents.Add();doc.Content.Text="BASE";
                if(mode!="add-inside") { stage=app.Documents.Add(Visible:false);stage.Content.Text="STAGING";doc.Activate(); }
                undo=app.UndoRecord;undo.StartCustomRecord("Owned cross-document lifecycle probe");
                void State(string label) => states.Add(new{Stage=label,Recording=undo.IsRecordingCustomRecord,Level=undo.CustomRecordLevel,Target=doc.Content.Text});
                doc.Content.InsertAfter("A");State("first-write");
                if(mode=="add-inside") stage=app.Documents.Add(Visible:false);
                else if(mode=="write-staging-inside") stage!.Content.Text="UPDATED";
                else if(mode=="close-inside") {stage!.Close(Word.WdSaveOptions.wdDoNotSaveChanges);Release(stage);stage=null;}
                State("staging-action");doc.Content.InsertAfter("B");State("second-write");
                undo.EndCustomRecord();doc.Activate();doc.Undo(1);State("one-undo");
                rows.Add(new{Mode=mode,States=states});
                File.WriteAllText(Path.Combine(root,"observations.json"),JsonSerializer.Serialize(rows,new JsonSerializerOptions{WriteIndented=true}));
                Console.WriteLine(mode+"="+JsonSerializer.Serialize(states));
            }
            finally
            {
                try{if(undo?.IsRecordingCustomRecord==true)undo.EndCustomRecord();}catch{}
                Release(undo);try{stage?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{}Release(stage);
                try{doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{}Release(doc);
                try{QuitWordApplicationIfOwned(app);}catch{}Release(app);ForceComCleanup();
            }
        }
    }
}

using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.VstoShared;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunMathTypeCharacterRootAcceptance(string root)
    {
        Directory.CreateDirectory(root);
        var svg=Path.Combine(root,"character-root.svg");
        File.WriteAllText(svg,"<svg xmlns='http://www.w3.org/2000/svg' width='240' height='80' viewBox='0 0 240 80'><text x='1' y='50'>Native formula</text></svg>");
        var emf=OfficeOlePreview.CreateVectorEmfFromSvg(svg,240,80);
        var percent="<mn>100</mn><mtext>%</mtext><msup><mi mathvariant='normal'>e</mi><mrow><mi mathvariant='normal'>i</mi><mi>π</mi></mrow></msup><mo>+</mo><mn>1</mn><mo>=</mo><mn>0</mn>";
        var inputs=new[]{
            ("percent-script",percent),
            ("bold-ddot-eta","<mover accent='true'><mi mathvariant='bold-italic'>η</mi><mo>¨</mo></mover>"),
            ("group-dot-primes","<mover accent='true'><mrow><mi>′</mi><mi>′</mi><mi>′</mi></mrow><mo>˙</mo></mover>"),
            ("nested-group-ddot","<mover accent='true'><mover accent='true'><mrow><mi>a</mi><mi>b</mi></mrow><mo>¨</mo></mover><mo>^</mo></mover>"),
            ("italic-contour-accent","<mover accent='true'><mo mathvariant='italic'>∮</mo><mo>¨</mo></mover>"),
            ("bolditalic-contour-accent","<mover accent='true'><mo mathvariant='bold-italic'>∮</mo><mo>¨</mo></mover>"),
            ("numeric-operator-styles","<mn mathvariant='bold-italic'>123</mn><mo mathvariant='bold'>−</mo><mi mathvariant='bold-script'>A</mi>"),
            ("native-font-glyphs","<mover accent='true'><mi mathvariant='bold-italic'>ℏ</mi><mo>^</mo></mover><mo>+</mo><mi mathvariant='sans-serif'>ȷ</mi><mi>ϱ</mi>"),
            ("empty-object-scripts","<mi>x</mi><mo>+</mo><munderover><mstyle><mrow/></mstyle><mn>2</mn><mn>3</mn></munderover><mi>H</mi>")
        };
        var results=new List<object>();Word.Application? app=null;Word.Document? doc=null;
        try
        {
            app=CreateWordApplication(visible:false);
            foreach(var table in new[]{false,true})
            foreach(var tracking in new[]{false,true})
            foreach(var input in inputs)
            {
                // Alternate mode across the inventory; the full group/script
                // grammar and both modes are separately exhaustively tested at
                // the MTEF boundary. This gate tests real Word create/edit history.
                var inline=Array.IndexOf(inputs,input)%2==0;
                var id=input.Item1+"-"+(table?"table":"body")+"-"+tracking;
                var output=Path.Combine(root,id);Directory.CreateDirectory(output);
                doc=app.Documents.Add();doc.SaveAs2(Path.Combine(output,"owned.docx"),Word.WdSaveFormat.wdFormatXMLDocument);
                doc.Content.Text="LEFT SENTINEL\rRIGHT SENTINEL\r";
                Word.Table? grid=null;Word.Range? at=null;Word.InlineShape? shape=null;Word.Range? range=null;
                try
                {
                    if(table)
                    {
                        var slot=doc.Paragraphs[2].Range;try{slot.Collapse(Word.WdCollapseDirection.wdCollapseStart);grid=doc.Tables.Add(slot,1,2);}finally{Release(slot);}
                        grid.Cell(1,2).Range.Text="UNTOUCHED CELL";
                        at=grid.Cell(1,1).Range;at.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    }
                    else{at=doc.Paragraphs[2].Range;at.Collapse(Word.WdCollapseDirection.wdCollapseStart);}
                    doc.TrackRevisions=tracking;
                    at.Select();var service=new WordFormulaService(app);
                    var source="<math xmlns='http://www.w3.org/1998/Math/MathML'>"+input.Item2+"</math>";
                    var latex=MathMlToLatexConverter.Convert(source);
                    var beforeCreate=new RootDocumentEvidence(doc,Path.Combine(output,"create"));
                    service.InsertMathTypeOle(CreateMathTypeCreateSession(inline?"inline":"block",false,latex),source,emf);
                    AssertEqual(1,doc.InlineShapes.Count,"Create did not materialize exactly one OLE.");
                    AssertCharacterRootPayload(doc,source);
                    AssertTrue(doc.Undo(1),"Create did not own one Undo.");beforeCreate.VerifyRollback();
                    AssertTrue(doc.Redo(1),"Create failed Redo.");AssertCharacterRootPayload(doc,source);
                    var read=ReadMathTypeProductSelection(service,doc,ref shape,ref range);
                    // This is the same product read/edit path as double-click,
                    // with the actual native stored source, not a fake cache.
                    var changed=XElement.Parse(source);XNamespace ns=changed.Name.Namespace;
                    changed.Add(new XElement(ns+"mo","+"),new XElement(ns+"mn","2"));
                    var next=changed.ToString();
                    var session=CreateMathTypeProductSession(read,MathMlToLatexConverter.Convert(next),240,80);
                    var beforeEdit=new RootDocumentEvidence(doc,Path.Combine(output,"edit"));
                    Release(range);range=null;Release(shape);shape=null;
                    service.ReplaceMathTypeOle(session,next,emf);
                    AssertCharacterRootPayload(doc,next);
                    AssertTrue(doc.Undo(1),"Edit did not own one Undo.");beforeEdit.VerifyRollback();AssertCharacterRootPayload(doc,source);
                    AssertTrue(doc.Redo(1),"Edit failed Redo.");AssertCharacterRootPayload(doc,next);
                    AssertEqual(tracking,doc.TrackRevisions,"Edit changed tracking mode.");
                    AssertTrue(doc.Content.Text.Contains("LEFT SENTINEL")&&doc.Content.Text.Contains("RIGHT SENTINEL"),"Edit lost surrounding prose.");
                    if(table)AssertTrue(doc.Content.Text.Contains("UNTOUCHED CELL"),"Edit touched another cell.");
                    var path=Path.Combine(output,"character-root.docx");doc.SaveAs2(path,Word.WdSaveFormat.wdFormatXMLDocument);
                    Release(at);at=null;Release(grid);grid=null;
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);Release(doc);doc=null;
                    doc=app.Documents.Open(path,ReadOnly:false,Visible:false,AddToRecentFiles:false);
                    AssertCharacterRootPayload(doc,next);
                    read=ReadMathTypeProductSelection(service,doc,ref shape,ref range);
                    AssertTrue(!string.IsNullOrWhiteSpace(read.Metadata?.Latex),"Reopened OLE is not editable through product source read.");
                    results.Add(new{Id=id,Inline=inline,CreateUndoRedo=true,EditUndoRedo=true,SaveReopenEditable=true,Tracking=tracking});
                    File.WriteAllText(Path.Combine(root,"progress.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
                    Console.WriteLine("PASS character root "+id);
                }
                catch{try{File.WriteAllText(Path.Combine(output,"failed.xml"),doc!.Content.WordOpenXML);}catch{}throw;}
                finally{Release(range);Release(shape);Release(at);Release(grid);try{doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{}Release(doc);doc=null;}
            }
            AssertEqual(36,results.Count,"Character root matrix incomplete.");
            File.WriteAllText(Path.Combine(root,"summary.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
        }
        finally{try{doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{}Release(doc);try{QuitWordApplicationIfOwned(app);}catch{}Release(app);ForceComCleanup();}
    }
    private static void AssertCharacterRootPayload(Word.Document doc,string expected)
    {
        Word.InlineShape? shape=null;
        try
        {
            AssertEqual(1,doc.InlineShapes.Count,"The current mathematical object count changed.");
            shape=doc.InlineShapes[1];
            AssertTrue(MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(shape,out var bytes),"Cannot read actual stored MathType payload.");
            var actual=MathTypeOleStorage.ReadMathMl(bytes);
            AssertEqual(MathTypeMtefCodec.SemanticSignature(expected),MathTypeMtefCodec.SemanticSignature(actual),"Actual Word OLE changed MTEF mathematical/style content.");
        }
        finally{Release(shape);}
    }
}

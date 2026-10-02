using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using VisualTeX.WordVsto;
using VisualTeX.WindowsOffice.VstoShared;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunWordNumberingRootMatrix(string root)
    {
        Directory.CreateDirectory(root);
        var svg=Path.Combine(root,"numbering-fixture.svg");
        File.WriteAllText(svg,"<svg xmlns='http://www.w3.org/2000/svg' width='80' height='24' viewBox='0 0 80 24'><text x='1' y='18'>a+b</text></svg>");
        var emf=OfficeOlePreview.CreateVectorEmfFromSvg(svg,80,24);
        var previous=WordEquationNumbering.GetDefaultEquationNumberFormatId();
        var results=new List<object>();
        Word.Application? app=null;
        try
        {
            app=CreateWordApplication(visible:false);
            foreach(var table in new[]{false,true})
            foreach(var manual in new[]{false,true})
            foreach(var tracking in new[]{false,true})
            {
                var name=(table?"table":"body")+"-"+(manual?"manual":"auto")+"-"+tracking;
                var only=Environment.GetEnvironmentVariable("VISUALTEX_NUMBERING_CASE");
                if(!string.IsNullOrWhiteSpace(only) && only!=name) continue;
                RunRootNumberingCase(app,root,name,table,manual,tracking,emf,results);
                File.WriteAllText(Path.Combine(root,"summary.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
            }
        }
        finally
        {
            try{QuitWordApplicationIfOwned(app);}catch{} Release(app);
            WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previous);
            ForceComCleanup();
        }
    }

    private static void RunRootNumberingCase(Word.Application app,string root,string name,bool table,bool manual,bool tracking,string emf,List<object> results)
    {
        var output=Path.Combine(root,name); Directory.CreateDirectory(output);
        Word.Document? doc=null; Word.ListTemplate? template=null;
        try
        {
            doc=app.Documents.Add();
            doc.SaveAs2(Path.Combine(output,"owned-fixture.docx"),Word.WdSaveFormat.wdFormatXMLDocument);
            template=CreateEquationNumberFormatHeadingTemplate(doc);
            var service=new WordFormulaService(app);
            WordEquationNumbering.RestoreEquationNumberFormatForConversion(doc,EquationNumberFormat.ContinuousId);
            var scopes=manual ? new[]{"4.2","4.3","5.1"} : new[]{"1.1","1.2","2.1"};
            for(var i=0;i<3;i++)
            {
                if(i==0 || i==2) RootHeading(app,doc,template,1,scopes[i].Split('.')[0],manual,i!=0);
                RootHeading(app,doc,template,2,scopes[i],manual,true);
                Word.Range? mt=null; Word.Range? native=null; Word.Table? grid=null;
                try
                {
                    if(table)
                    {
                        var slot=RootAppendParagraph(doc,"");
                        try{slot.Collapse(Word.WdCollapseDirection.wdCollapseStart);grid=doc.Tables.Add(slot,1,2);}
                        finally{Release(slot);}
                        mt=grid.Cell(1,1).Range; mt.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    }
                    else {mt=RootAppendParagraph(doc,"");mt.Collapse(Word.WdCollapseDirection.wdCollapseStart);}
                    mt.Select();
                    service.InsertMathTypeOle(CreateMathTypeCreateSession("block",true,"a+b",i==1?"left":"right"),FirstNumberedMathMl,emf);
                    if(table) {native=grid!.Cell(1,2).Range; native.Collapse(Word.WdCollapseDirection.wdCollapseStart);}
                    else {native=RootAppendParagraph(doc,"");native.Collapse(Word.WdCollapseDirection.wdCollapseStart);}
                    InsertNativeNumberedFixture(doc,native.Duplicate,"x+"+(i+1),0,"",(i+1).ToString(),i==0?1:(int?)null);
                    Word.OMath? math=null;Word.Range? exact=null;Word.Range? number=null;Word.Bookmark? alias=null;
                    try
                    {
                        math=doc.OMaths[i+1];exact=math.Range;number=WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                        alias=doc.Bookmarks.Add("_RefRootNative"+i,number);
                    }
                    finally{Release(alias);Release(number);Release(exact);Release(math);}
                    var tail=RootAppendParagraph(doc,"SENTINEL-"+i);Release(tail);
                }
                finally{Release(native);Release(mt);Release(grid);}
            }
            // Native aliases and actual REF fields are owned by Word, not VT IDs.
            for(var i=0;i<3;i++)
            {
                var text=RootAppendParagraph(doc,"Native reference ");
                Word.Field? field=null;
                try{text.End--;text.Collapse(Word.WdCollapseDirection.wdCollapseEnd);field=text.Fields.Add(text,Word.WdFieldType.wdFieldRef,"_RefRootNative"+i+" \\h",true);field.Update();}
                finally{Release(field);Release(text);}
            }
            var targets=MathTypeEquationReferences.GetTargets(doc);
            AssertEqual(3,targets.Count,"Fixture lost native MathType targets.");
            for(var i=0;i<targets.Count;i++)
            {
                var range=RootAppendParagraph(doc,"");Word.Selection? selection=null;
                try{range.Collapse(Word.WdCollapseDirection.wdCollapseStart);selection=app.Selection;selection.SetRange(range.Start,range.Start);MathTypeEquationReferences.InsertReference(doc,selection,targets[i]);}
                finally{Release(selection);Release(range);}
            }
            doc.TrackRevisions=tracking;
            var contentEvidence=RootMathematicalEvidence(doc);
            var tableCount=doc.Tables.Count;
            var formats=new[]{EquationNumberFormat.Heading1DotId,EquationNumberFormat.Heading2DotId,EquationNumberFormat.ContinuousId,EquationNumberFormat.Heading2DashId,EquationNumberFormat.Heading1DashId};
            var requestedFormat=Environment.GetEnvironmentVariable("VISUALTEX_NUMBERING_FORMAT");
            if(!string.IsNullOrWhiteSpace(requestedFormat)) formats=formats.Where(format=>format==requestedFormat).ToArray();
            AssertTrue(formats.Length>0,"No requested format matched the targeted matrix.");
            foreach(var format in formats)
            {
                Console.WriteLine("START root matrix "+name+" "+format);
                var timer=System.Diagnostics.Stopwatch.StartNew();
                void Step(string stage)=>Console.WriteLine("TIMING "+name+" "+format+" "+stage+" "+timer.ElapsedMilliseconds+"ms");
                var before=CaptureRootNumberScopes(doc,Path.Combine(output,format));
                var paragraphs=doc.Paragraphs.Count;
                var originalNumbers=RootNumberLabels(doc);
                try
                {
                    AssertEqual(6,service.SetEquationNumberFormat(format),"Mixed format count must include each family once.");
                    Step("set-format");
                    var expected=RootExpectedLabels(format,scopes);
                    AssertMathTypeNumberTexts(doc,expected.Select(x=>"("+x+")").ToArray());
                    AssertRootNativeLabels(doc,expected);
                    AssertRootReferences(doc);
                    AssertEqual(contentEvidence,RootMathematicalEvidence(doc),"Number format mutated mathematical payload/body.");
                    AssertEqual(tableCount,doc.Tables.Count,"Number format changed table ownership.");
                    AssertEqual(tracking,doc.TrackRevisions,"Tracking flag changed.");
                    Step("verify-applied");
                    AssertTrue(doc.Undo(1),"Format did not own one native Undo record.");
                    foreach(var snapshot in before) snapshot.VerifyRollback();
                    AssertEqual(originalNumbers,RootNumberLabels(doc),"Undo did not restore original labels.");
                    AssertEqual(paragraphs,doc.Paragraphs.Count,"Undo did not remove newly generated section paragraphs.");
                    Step("verify-undo");
                    AssertTrue(doc.Redo(1),"Format failed native Redo.");
                    AssertMathTypeNumberTexts(doc,expected.Select(x=>"("+x+")").ToArray());
                    AssertRootNativeLabels(doc,expected);AssertRootReferences(doc);
                    Step("verify-redo");
                    service.UpdateEquationNumbers();
                    Step("update");
                    AssertMathTypeNumberTexts(doc,expected.Select(x=>"("+x+")").ToArray());
                    AssertRootNativeLabels(doc,expected);AssertRootReferences(doc);
                    AssertEqual(contentEvidence,RootMathematicalEvidence(doc),"Update mutated mathematics.");
                    File.WriteAllText(Path.Combine(output,format+".xml"),doc.Content.WordOpenXML);
                    results.Add(new{Name=name,Format=format,IndependentCounters=true,StrictLocalUndoRedo=true,References=true,MathematicalPayload=true,Tracking=tracking});
                    File.WriteAllText(Path.Combine(root,"progress.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
                    Console.WriteLine("PASS root matrix "+name+" "+format);
                }
                catch
                {
                    File.WriteAllText(Path.Combine(output,"failed-"+format+".xml"),doc.Content.WordOpenXML);
                    throw;
                }
                finally{foreach(var snapshot in before) snapshot.Dispose();}
            }
            // Inject AFTER real native field writes, through the same transaction
            // kernel, without adding a fault hook to any production assembly.
            var rollback=CaptureRootNumberScopes(doc,Path.Combine(output,"terminal-fault"));
            try
            {
                var labels=RootNumberLabels(doc);
                using var mtPlan=MathTypeEquationNumbering.PrepareEquationNumberFormat(doc,EquationNumberFormat.ContinuousId);
                using var nativePlan=WordNativeNumberingAdoptionPlan.Prepare(doc,numberFormat:EquationNumberFormat.ContinuousId);
                var injected=false;
                try
                {
                    WordFormulaMutationTransaction.Execute<int>(app,doc,"Targeted native number terminal fault",()=>{
                        AssertEqual(3,MathTypeEquationNumbering.SetEquationNumberFormat(doc,mtPlan),"Fault did not rewrite all MTPlaceRefs.");
                        nativePlan.Apply();
                        throw new COMException("ROOT_NUMBER_TERMINAL_FAULT",unchecked((int)0x8007000E));
                    },verifyRollback:()=>{foreach(var snapshot in rollback)snapshot.VerifyRollback();});
                }
                catch(COMException e) when(e.Message=="ROOT_NUMBER_TERMINAL_FAULT"){injected=true;}
                AssertTrue(injected,"Terminal fault was not reached or exact rollback failed.");
                AssertEqual(labels,RootNumberLabels(doc),"Fault rollback changed labels.");AssertRootReferences(doc);
                AssertEqual(contentEvidence,RootMathematicalEvidence(doc),"Fault rollback changed mathematics.");
                results.Add(new{Name=name,TerminalModuleFaultRollback=true});
                File.WriteAllText(Path.Combine(root,"progress.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
                Console.WriteLine("PASS terminal module fault "+name);
            }
            finally{foreach(var snapshot in rollback)snapshot.Dispose();}
            var file=Path.Combine(output,"numbering-root.docx");doc.SaveAs2(file,Word.WdSaveFormat.wdFormatXMLDocument);
            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);Release(doc);doc=null;
            doc=app.Documents.Open(file,ReadOnly:false,Visible:false,AddToRecentFiles:false);
            var final=RootExpectedLabels(formats[formats.Length-1],scopes);
            AssertMathTypeNumberTexts(doc,final.Select(x=>"("+x+")").ToArray());AssertRootNativeLabels(doc,final);AssertRootReferences(doc);
            AssertEqual(contentEvidence,RootMathematicalEvidence(doc),"Save/reopen changed math payload/body.");
            results.Add(new{Name=name,SaveReopen=true});
            File.WriteAllText(Path.Combine(root,"progress.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions{WriteIndented=true}));
            Console.WriteLine("PASS save/reopen "+name);
        }
        finally
        {
            Release(template);try{doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges);}catch{}Release(doc);ForceComCleanup();
        }
    }

    private static Word.Range RootAppendParagraph(Word.Document doc,string text)
    {
        Word.Range? end=null;
        try{var start=doc.Content.End-1;end=doc.Range(start,start);end.InsertAfter(text+"\r");return doc.Range(start,start+text.Length+1);}
        finally{Release(end);}
    }
    private static void RootHeading(Word.Application app,Word.Document doc,Word.ListTemplate template,int level,string prefix,bool manual,bool continuation)
    {
        if(!manual){InsertNumberedFormatHeading(app,doc,template,level,"Fixture heading "+prefix,continuation);return;}
        var range=RootAppendParagraph(doc,prefix+" Fixture heading");
        try{object style=level==1?Word.WdBuiltinStyle.wdStyleHeading1:Word.WdBuiltinStyle.wdStyleHeading2;range.set_Style(ref style);}
        finally{Release(range);}
    }
    private static string[] RootExpectedLabels(string format,string[] scopes)
    {
        var f=EquationNumberFormat.Resolve(format);var counts=new Dictionary<string,int>();var labels=new List<string>();
        foreach(var scope in scopes){var prefix=f.HeadingLevel==0?"":f.HeadingLevel==1?scope.Split('.')[0]:scope;counts.TryGetValue(prefix,out var n);counts[prefix]=++n;labels.Add(prefix+(f.UsesHeading?f.Separator:"")+n);}
        return labels.ToArray();
    }
    private static Word.Range RootNativeOwnerRange(Word.Document doc,int index)
    {
        Word.Bookmark? bookmark=null;Word.Range? alias=null;Word.Range? unit=null;Word.OMaths? maths=null;
        try
        {
            bookmark=doc.Bookmarks["_RefRootNative"+index];alias=bookmark.Range;
            unit=WordNumberingFieldCollection.CaptureSearchUnit(alias);
            maths=unit.OMaths;
            Word.Range? found=null;
            for(var i=1;i<=maths.Count;i++)
            {
                Word.OMath? math=null;Word.Range? exact=null;
                try
                {
                    math=maths[i];exact=math.Range;
                    if(!exact.InStory(alias)||exact.Start>alias.Start||exact.End<alias.End)continue;
                    var xml=exact.WordOpenXML;
                    if(!WordOmmlConverter.HasWordNativeEquationNumberHost(xml))continue;
                    if(found is not null)throw new InvalidDataException("Native root alias resolves multiple numbered OMath owners.");
                    found=exact.Duplicate;
                }
                catch(System.Runtime.InteropServices.COMException)
                {
                    // Revised REF-result OMaths can expose invalid transient COM
                    // owners. They are not candidates unless they physically
                    // contain this exact native-number bookmark.
                }
                finally{Release(exact);Release(math);}
            }
            return found??throw new InvalidDataException("Native root alias lost its numbered OMath owner.");
        }
        finally{Release(maths);Release(unit);Release(alias);Release(bookmark);}
    }

    private static string[] RootNativeLabels(Word.Document doc)
    {
        var labels=new List<string>();
        for(var i=0;i<3;i++)
        {
            Word.Range? exact=null;
            try
            {
                exact=RootNativeOwnerRange(doc,i);
                labels.Add(WordOmmlConverter.ReadWordNativeEquationNumberText(exact.WordOpenXML));
            }
            finally{Release(exact);}
        }
        return labels.ToArray();
    }
    private static void AssertRootNativeLabels(Word.Document doc,string[] expected)
    {
        AssertEqual(string.Join("|",expected),string.Join("|",RootNativeLabels(doc)),"Native OMML did not follow document policy.");
        for(var i=0;i<expected.Length;i++)
        {
            Word.Range? range=null;Word.Range? number=null;Word.Bookmark? bookmark=null;Word.Range? alias=null;
            try
            {
                range=RootNativeOwnerRange(doc,i);number=WordNativeOmmlNumbering.ResolveNumberContentRange(range);
                bookmark=doc.Bookmarks["_RefRootNative"+i];alias=bookmark.Range;
                AssertTrue(alias.InStory(number)&&alias.Start==number.Start&&alias.End==number.End,"Native alias lost current argument ownership.");
            }
            finally{Release(alias);Release(bookmark);Release(number);Release(range);}
        }
    }
    private static string RootNumberLabels(Word.Document doc)
    {
        var texts=new List<string>(RootNativeLabels(doc));
        using var fields=WordNumberingFieldCollection.Capture(doc);
        for(var i=1;i<=fields.Count;i++){var f=fields[i];Word.Range? c=null;Word.Range? r=null;try{c=f.Code;if(MathTypeEquationReferences.IsMathTypePlaceRefCode(c.Text)){r=f.Result;texts.Add(r.Text??"");}}finally{Release(r);Release(c);Release(f);}}
        return string.Join("|",texts);
    }
    private sealed class RootDocumentEvidence : IDisposable
    {
        private readonly Word.Document _document;
        private readonly string _signature;
        private readonly int _end;
        private readonly string _evidence;
        internal RootDocumentEvidence(Word.Document document,string evidence)
        {
            _document=document;_evidence=evidence;var range=document.Content;
            try{
                _end=range.End;var xml=range.WordOpenXML;var geometry=WordInlineObjectGeometry.Capture(range);
                _signature=WordLocalEditSnapshot.Signature(xml,geometry);
                File.WriteAllText(_evidence+"-snapshot-before.xml",WordLocalEditSnapshot.NormalizedBody(xml,geometry));
            }
            finally{Release(range);}
        }
        internal void VerifyRollback()
        {
            var range=_document.Content;
            try
            {
                AssertEqual(_end,range.End,"Undo changed the complete owned fixture extent.");
                var xml=range.WordOpenXML;var geometry=WordInlineObjectGeometry.Capture(range);
                var actual=WordLocalEditSnapshot.Signature(xml,geometry);
                if(actual!=_signature)File.WriteAllText(_evidence+"-snapshot-after.xml",WordLocalEditSnapshot.NormalizedBody(xml,geometry));
                AssertEqual(_signature,actual,"Undo changed owned fixture content, revision provenance, bookmarks or live OLE geometry.");
            }
            finally{Release(range);}
        }
        public void Dispose() { }
    }
    private static List<RootDocumentEvidence> CaptureRootNumberScopes(Word.Document doc,string evidence) =>
        // This is a small generated fixture, not a user's large document. One
        // complete strict snapshot is both stronger and cheaper than exporting
        // the same package once for every paragraph. Product scope is unchanged.
        new(){new RootDocumentEvidence(doc,evidence)};
    private static string RootMathematicalEvidence(Word.Document doc)
    {
        var values=new List<string>();var shapes=doc.InlineShapes;var maths=doc.OMaths;
        try
        {
            using var hash=SHA256.Create();
            for(var i=1;i<=shapes.Count;i++){var shape=shapes[i];try{AssertTrue(MathTypeOleStorage.TryCaptureCompoundFileFromWordOpenXml(shape,out var bytes),"Cannot independently read OLE storage.");values.Add(Convert.ToBase64String(hash.ComputeHash(bytes)));}finally{Release(shape);}}
            // A REF result to native math may itself be exposed as an OMath.
            // It is not an independent formula body and must change with its
            // target number. Only the structurally proven #(number) hosts are
            // fixture mathematical payload; references have their own assertions.
            for(var i=0;i<3;i++)
            {
                Word.Range? range=null;
                try
                {
                    range=RootNativeOwnerRange(doc,i);
                    values.Add(WordOmmlConverter.StripWordNativeEquationNumberHost(range.WordOpenXML));
                }
                finally{Release(range);}
            }
        }
        finally{Release(maths);Release(shapes);}return string.Join("\n",values);
    }
    private static void AssertRootReferences(Word.Document doc)
    {
        var found=0;using var fields=WordNumberingFieldCollection.Capture(doc);
        for(var i=1;i<=fields.Count;i++)
        {
            var field=fields[i];Word.Range? code=null;Word.Range? result=null;Word.Bookmark? bookmark=null;Word.Range? target=null;
            try
            {
                if(field.Type!=Word.WdFieldType.wdFieldRef)continue;code=field.Code;
                if(!WordEquationReferenceFields.TryReadReferenceBookmark(code.Text,out var name))continue;
                if(!name.StartsWith("_RefRootNative",StringComparison.Ordinal)&&!name.StartsWith("ZEqnNum",StringComparison.OrdinalIgnoreCase))continue;
                bookmark=doc.Bookmarks[name];target=bookmark.Range;result=field.Result;
                AssertEqual(target.Text,result.Text,"A current reference displays stale/deleted numbering.");found++;
            }
            finally{Release(target);Release(bookmark);Release(result);Release(code);Release(field);}
        }
        AssertTrue(found>=6,"Fixture lost native or MathType references.");
    }
}

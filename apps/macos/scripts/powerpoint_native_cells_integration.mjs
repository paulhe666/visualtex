// Real PowerPoint regression of native equations in table cells: the production capture,
// insertion (linear format + build-up), edit source copy, edit and double-click resolution.
// Operates exclusively on the owned NativeOmmlAcceptance.pptm and temporary sessions.
import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { homedir } from 'node:os';
import { randomUUID } from 'node:crypto';
import { fileURLToPath } from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const out=resolve(root,'build-logs/powerpoint-native-omml');
const sessionRoot=resolve(homedir(),'Library/Application Scripts/com.microsoft.Powerpoint/VisualTeXRuntime/OfficeSessions');
const pointer=resolve(sessionRoot,'powerpoint-active-session.txt');
const originalPointer=existsSync(pointer)?readFileSync(pointer):null;
function as(lines){return execFileSync('/usr/bin/osascript',lines.flatMap(x=>['-e',x]),{encoding:'utf8',timeout:60000}).trim();}
function macro(name,args=[]){return as(['tell application "Microsoft PowerPoint"',`run VB macro macro name ${JSON.stringify(name)} list of parameters {${args.map(x=>JSON.stringify(x)).join(',')}}`,'end tell']);}
const keepClipboard=[
 "set pb to current application's NSPasteboard's generalPasteboard()", "set saved to current application's NSMutableArray's array()",
 "repeat with sourceItem in pb's pasteboardItems()", "set newItem to current application's NSPasteboardItem's alloc()'s init()",
 "repeat with t in sourceItem's types()", "newItem's setData:(sourceItem's dataForType:t) forType:t",'end repeat',"saved's addObject:newItem",'end repeat'];
const restoreClipboard=["if ((pb's changeCount()) as integer) is ownedCount then", "pb's clearContents()","pb's writeObjects:saved",'end if'];
function apply(payload){return as([
 'use framework "AppKit"','use scripting additions',
 `set payload to current application's NSData's dataWithContentsOfFile:${JSON.stringify(payload)}`,
 'if payload is missing value then error "Missing native clipboard fixture"',...keepClipboard,
 "pb's clearContents()", 'set didSet to pb\'s setData:payload forType:"com.microsoft.Art--Text-ClipFormat"',
 "set ownedCount to (pb's changeCount()) as integer",
 'try','if didSet is false then error "Cannot set the native clipboard"',
 'tell application "Microsoft PowerPoint" to run VB macro macro name "VTNativeAcceptanceApply" list of parameters {}',
 'on error messageText number errorNumber',...restoreClipboard,'error messageText number errorNumber','end try',...restoreClipboard,
]);}
// The edit source as macos_offline.rs reads it: Office's own copy of the equation.
function copyEditSource(destination){return as([
 'use framework "AppKit"','use scripting additions',...keepClipboard,
 "set initialCount to (pb's changeCount()) as integer","set ownedCount to initialCount",'try',
 'tell application "Microsoft PowerPoint" to run VB macro macro name "VTNativeAcceptanceCopyEditSource" list of parameters {}',
 "set ownedCount to (pb's changeCount()) as integer",
 'set nativeData to pb\'s dataForType:"com.microsoft.Art--Text-ClipFormat"',
 'if nativeData is missing value then error "PowerPoint did not export the cell equation as Office Math"',
 `if (nativeData's writeToFile:${JSON.stringify(destination)} atomically:true) is false then error "Cannot keep the cell equation"`,
 'on error messageText number errorNumber',"if ((pb's changeCount()) as integer) is ownedCount and ownedCount is not initialCount then","pb's clearContents()","pb's writeObjects:saved",'end if','error messageText number errorNumber','end try',
 "if ((pb's changeCount()) as integer) is ownedCount and ownedCount is not initialCount then","pb's clearContents()","pb's writeObjects:saved",'end if',
]);}
function session(display){
 const id=randomUUID(); const dir=resolve(sessionRoot,id); mkdirSync(dir,{recursive:true});
 writeFileSync(resolve(dir,'dispatch.txt'),`protocolVersion=1\nsessionId=${id}\nhost=powerpoint\naction=commit\nmode=create\nnativeEquation=1\nleadingParagraph=0\ntrailingParagraph=0\ndisplayMode=${display?'block':'inline'}\nfontSizePt=24\n`);
 writeFileSync(pointer,id);
 return {id,dir};
}
function cells(id){
 macro('VTNativeCellReport',[id]);
 const [slide,shapes,...items]=readFileSync(resolve(sessionRoot,id,'acceptance-report.txt'),'utf8').split('|');
 const map={};
 for(const item of items){const [key,rest]=[item.slice(0,item.indexOf('=')),item.slice(item.indexOf('=')+1)];const at=rest.lastIndexOf('#');map[key]={text:rest.slice(0,at),zones:Number(rest.slice(at+1))};}
 return {slide:Number(slide),shapes:Number(shapes),cells:map};
}
function expect(condition,message){if(!condition)throw Error(message);}
function zipHasMath(path){
 return execFileSync('/usr/bin/python3',['-c','import sys,zipfile;z=zipfile.ZipFile(sys.argv[1]);print(any(b"<a14:m" in z.read(n) for n in z.namelist()))',path],{encoding:'utf8'}).trim()==='True';
}
const results=[];
const cleanup=[];
try{
 for(const name of ['cell_inline','cell_display','cell_merged']){
  const display=name==='cell_display';
  const {id,dir}=session(display);cleanup.push(id);
  macro('VTNativeCellPrepare',[name,id]);
  const prepared=readFileSync(resolve(dir,'native-prepare.txt'),'utf8').trim();
  expect(prepared==='ok',name+' prepare: '+prepared);
  const request=JSON.parse(readFileSync(resolve(dir,'native-request.json'),'utf8'));
  expect(request.powerPoint?.nativeTextTarget?.tableCell===true,name+': the request does not name a table cell');
  expect(String(request.powerPoint.shapeName).startsWith('Cells_'),name+': the request geometry is not the table\'s');
  apply(resolve(out,display?'native-display00.bin':'native-inline.bin'));
  const result=readFileSync(resolve(dir,'native-result.txt'),'utf8').trim();
  expect(result==='ok',name+': '+result);
  const after=cells(id);
  expect(after.shapes===1,name+' left '+after.shapes+' shapes on the slide');
  const key={cell_inline:'2,2',cell_display:'3,1',cell_merged:'1,1'}[name];
  expect(after.cells[key].zones===1,name+': no equation in cell '+key+': '+JSON.stringify(after.cells[key]));
  if(name==='cell_inline'){const t=after.cells['2,2'].text;expect(t.startsWith('before ')&&t.endsWith('after'),'cell_inline changed the words: '+t);}
  if(name==='cell_merged'){const t=after.cells['1,1'].text;expect(t.startsWith('mer')&&t.endsWith('ged'),'cell_merged changed the words: '+t);}
  macro('VTNativeAcceptanceApply');
  expect(JSON.stringify(cells(id))===JSON.stringify(after),name+' changed the slide on a repeated callback');
  console.log(name+': '+JSON.stringify(after.cells[key]));
  results.push({name,cell:key,...after.cells[key]});
  if(name!=='cell_inline') continue;

  // Edit the equation just written, from the caret inside it.
  const edit=session(false);cleanup.push(edit.id);
  macro('VTNativeCellEditPrepare',['caret',edit.id]);
  const editPrepared=readFileSync(resolve(edit.dir,'native-edit-prepare.txt'),'utf8');
  expect(/resolved=\d+:\d+/.test(editPrepared)&&editPrepared.includes('cell=2,2'),'cell edit did not resolve: '+editPrepared);
  const editRequest=JSON.parse(readFileSync(resolve(edit.dir,'request.json'),'utf8'));
  expect(editRequest.powerPoint?.nativeTextTarget?.tableCell===true&&editRequest.mode==='edit','cell edit request: '+JSON.stringify(editRequest.powerPoint?.nativeTextTarget));
  const source=resolve(out,'cell-edit-source.bin');
  copyEditSource(source);
  const copyStatus=readFileSync(resolve(edit.dir,'native-edit-copy-status.txt'),'utf8').trim();
  expect(copyStatus==='ok','cell edit source copy: '+copyStatus);
  expect(zipHasMath(source),'the copied cell equation holds no Office Math');
  apply(resolve(out,'native-inline36.bin'));
  const editResult=readFileSync(resolve(edit.dir,'native-result.txt'),'utf8').trim();
  expect(editResult==='ok','cell edit: '+editResult);
  const edited=cells(edit.id);
  const t=edited.cells['2,2'].text;
  expect(edited.cells['2,2'].zones===1&&t.startsWith('before ')&&t.endsWith('after')&&edited.shapes===1,'cell edit result: '+JSON.stringify(edited));
  console.log('cell_edit: '+JSON.stringify(edited.cells['2,2']));
  results.push({name:'cell_edit',cell:'2,2',...edited.cells['2,2']});

  // The double-click resolution at the equation's screen point.
  const point=session(false);cleanup.push(point.id);
  macro('VTNativeCellEditPrepare',['point',point.id]);
  const pointPrepared=readFileSync(resolve(point.dir,'native-edit-prepare.txt'),'utf8');
  expect(/resolved=\d+:\d+/.test(pointPrepared),'the double-click point missed the cell equation: '+pointPrepared);
  console.log('cell_point: '+pointPrepared.replace(/\n/g,' | '));
  results.push({name:'cell_point',report:pointPrepared});
 }
 const resultPath=resolve(out,'native-cells-integration-results.json');
 writeFileSync(resultPath,JSON.stringify(results,null,2));
}finally{
 for(const id of cleanup){try{macro('VTCancelPowerPointNativeEquation',[id]);}catch{} rmSync(resolve(sessionRoot,id),{recursive:true,force:true});}
 if(originalPointer!==null)writeFileSync(pointer,originalPointer);else if(existsSync(pointer))rmSync(pointer);
}
console.log(`Completed ${results.length} real PowerPoint table-cell cases.`);

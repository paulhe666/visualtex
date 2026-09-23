// P08 independent gathered/split/merge Native and WASM equivalence.
// Both actual binaries receive identical font bytes and command streams. This
// does not run a browser, native IME, Office or a production integration path.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { once } from 'node:events';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { readFileSync, mkdirSync, writeFileSync } from 'node:fs';
import { dirname, resolve, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const output=resolve(root,'target/contained-multiline-parity-review',stamp);mkdirSync(output,{recursive:true});
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const paths={math:resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),text:'C:/Windows/Fonts/NotoSansSC-VF.ttf'};
const hashes={math:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f',text:'763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a'};
const nativePath=resolve(root,`target/debug/visualtex-engine-test-cli${process.platform==='win32'?'.exe':''}`);
const wasmPath=resolve(root,'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const binaries={nativeSha256:sha(readFileSync(nativePath)),wasmSha256:sha(readFileSync(wasmPath))};
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const wasm=new EngineRuntime(),fonts={};
for(const [name,path] of Object.entries(paths)){
 const bytes=readFileSync(path);assert.equal(sha(bytes),hashes[name]);
 const response=JSON.parse(wasm.registerFont(bytes,0));assert.equal(response.status,'ok');fonts[name]=response.result.font;
}
const native=spawn(nativePath,['--font',paths.math,'--font',paths.text],{stdio:['pipe','pipe','pipe']});
const lines=createInterface({input:native.stdout,crlfDelay:Infinity})[Symbol.asyncIterator]();
let stderr='';native.stderr.on('data',data=>{stderr=(stderr+String(data)).slice(-16000);});
const timer=setTimeout(()=>native.kill(),90000);
let requests=0,scenes=0,maxNumericDifference=0,sourceOracles=0;
const transcript=createHash('sha256'),checks=[];
function equivalent(a,b,path='response'){
 if(typeof a==='number'&&typeof b==='number'){
  assert.ok(Number.isFinite(a)&&Number.isFinite(b),path);const d=Math.abs(a-b);maxNumericDifference=Math.max(maxNumericDifference,d);
  assert.ok(d<=1e-9,`${path}: ${a} != ${b}`);return;
 }
 if(a===null||b===null||typeof a!=='object'||typeof b!=='object'){assert.equal(a,b,path);return;}
 assert.equal(Array.isArray(a),Array.isArray(b),path);assert.deepEqual(Object.keys(a),Object.keys(b),path);
 for(const key of Object.keys(a))equivalent(a[key],b[key],`${path}.${key}`);
}
async function send(command,expectedError=null){
 const input=JSON.stringify({protocol:1,command});const a=JSON.parse(wasm.dispatch(input));
 if(!native.stdin.write(input+'\n'))await once(native.stdin,'drain');const line=await lines.next();assert.equal(line.done,false,stderr);
 const b=JSON.parse(line.value);equivalent(a,b,`${requests}:${command.op}`);requests++;transcript.update(input+'\n'+line.value+'\n');
 if(expectedError){assert.equal(a.status,'error');assert.equal(a.error.code,expectedError);return a;}
 assert.equal(a.status,'ok',JSON.stringify(a));return a.result;
}
const defaults={mathFont:fonts.math.id,textFont:fonts.text.id,style:'display',fontSizePx:16,paddingEm:.125};
async function layout(handle,revision=0,options={},includeSvg=true){
 const result=await send({op:'layout',handle,expectedRevision:revision,options:{...defaults,...options},includeSvg});
 scenes++;const s=result.snapshot;
 assert.equal(s.width,s.widthEm*s.fontSizePx);assert.equal(s.height,s.heightEm*s.fontSizePx);assert.equal(s.baseline,s.baselineEm*s.fontSizePx);
 assert.equal(new Set(s.regions.map(r=>r.nodeId)).size,s.regions.length);
 const ids=new Set(s.regions.map(r=>r.nodeId)),columnIds=new Set();
 for(const column of s.gridColumns??[]){
  assert.ok(Number.isFinite(column.x)&&Number.isFinite(column.width)&&column.width>=0);
  assert.ok(!ids.has(column.columnId)&&!columnIds.has(column.columnId));columnIds.add(column.columnId);
  const region=s.regions.find(r=>r.nodeId===column.gridId);assert.equal(region?.role,'grid');
  assert.ok(column.x>=region.logical.xMin-1e-9&&column.x+column.width<=region.logical.xMax+1e-9);
 }
 if(includeSvg){assert.ok(s.exportReady);assert.ok(result.svg.startsWith('<svg'));assert.equal(result.svg.match(/<svg\b/g)?.length,1);
  for(const forbidden of ['<text','<image','foreignObject','<script','NaN','Infinity','href='])assert.ok(!result.svg.includes(forbidden),forbidden);
 }
 return result;
}
function semantic(value){
 if(Array.isArray(value))return value.map(semantic);
 if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value)
  .filter(([k])=>!['id','origin','grouping','leading','trailing'].includes(k)&&!k.startsWith('before'))
  .map(([k,v])=>[k,k==='columns'?v.length:semantic(v)]));
 return value;
}
function allIds(node){
 const ids=new Set();
 const visit=value=>{
  if(!value||typeof value!=='object')return;
  if(Array.isArray(value)){for(const entry of value)visit(entry);return;}
  if(value.kind&&typeof value.id==='number'){
   assert.ok(!ids.has(value.id));ids.add(value.id);
   if(value.kind==='grid')for(const column of value.columns){assert.ok(!ids.has(column));ids.add(column);}
  }
  for(const [key,child] of Object.entries(value))if(key!=='origin'&&key!=='columns')visit(child);
 };
 visit(node);return ids;
}
async function oracle(document){
 const fresh=await send({op:'analyzeSource',source:document.source});sourceOracles++;
 assert.equal(fresh.lossless,document.source);
 assert.deepEqual(semantic(fresh.analysis.tree),semantic(document.analysis.tree),'Owned tree differs from fresh parse: '+document.source);
 assert.equal(allIds(document.analysis.tree).size,document.analysis.facts.nodes+document.analysis.facts.gridColumns);
 assert.deepEqual(fresh.analysis.facts,document.analysis.facts);
}
function grid(editor){return editor.document.analysis.tree.entries.find(e=>e.node.kind==='grid').node;}
function cell(editor,i){return grid(editor).rows[i].cells[0];}
const p=(nodeId,offset=0,affinity='after')=>({kind:'row',nodeId,offset,affinity});
const caret=position=>({kind:'range',anchor:position,focus:position});
const sourceList=[
 String.raw`\begin{gathered}x+y=1\\x-y=0\end{gathered}`,
 String.raw`\begin{gathered}\frac{x}{y}\\\sqrt[3]{x_1}\end{gathered}`,
 String.raw`\begin{gathered}\text{速度}\\\text{中文 Café}\end{gathered}`,
 String.raw`\begin{gathered}\sum_i^n x_i\\\int_0^1 x\,\mathrm{d}x\end{gathered}`,
 String.raw`\begin{gathered}\textcolor{red}{\boxed{x}}\\\hat{y}+z\end{gathered}`,
 String.raw`\begin{gathered}\begin{gathered}a\\b\end{gathered}\\c\end{gathered}`,
 String.raw`\frac{\begin{gathered}a\\b\end{gathered}}{c}`,
 String.raw`\left\{\begin{gathered}a=1\\b=2\end{gathered}\right.`,
 String.raw`\vphantom{\begin{gathered}a+b+c\\d\end{gathered}}x`,
 String.raw`\hphantom{\begin{gathered}\frac{a}{b}\\c\end{gathered}}x`,
 String.raw`\phantom{\begin{gathered}a\\b\end{gathered}}x`,
 String.raw`\begin{gathered}\!\\x\end{gathered}`,
 "\\begin % opening\n {gathered}a % not a break \\\\\n +b\\\\c\\end % closing\n {gathered}",
];
let report;
try{
 for(const [i,source] of sourceList.entries()){
  const h=(await send({op:'create',documentId:`contained-geometry-${i}`,source})).handle;
  const document=(await send({op:'inspectDocument',handle:h})).snapshot;await oracle(document);
  assert.ok(document.analysis.complete&&document.analysis.supported,'Unsupported valid gathered source');
  for(const style of ['display','text','script','scriptScript']){
   const first=await layout(h,0,{style});const cached=await layout(h,0,{style});assert.equal(cached.svg,first.svg);assert.equal(cached.snapshot.rebuiltFragments,0);
   const doubled=await layout(h,0,{style,fontSizePx:32});assert.equal(doubled.snapshot.width,first.snapshot.width*2);assert.equal(doubled.snapshot.height,first.snapshot.height*2);
   assert.deepEqual(doubled.snapshot.gridColumns,first.snapshot.gridColumns);assert.deepEqual(doubled.snapshot.glyphs,first.snapshot.glyphs);
  }
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 checks.push('13 nested/hidden/colored/Unicode/comment formulas in four styles and two natural sizes');
 const a=(await send({op:'create',documentId:'formula-A',source:'abc'})).handle;
 const b=(await send({op:'create',documentId:'formula-B',source:'z_2'})).handle;
 const bBefore=(await send({op:'inspectEditor',handle:b})).snapshot;
 let revision=0,sequence=0;
 const action=async action=>{
  const r=await send({op:'editorAction',handle:a,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;
  await oracle(r.result.editor.document);
  assert.deepEqual((await send({op:'inspectEditor',handle:b})).snapshot,bBefore,'Frame B changed');
  return r.result;
 };
 const edit=command=>action({kind:'edit',command,groupId:null});
 await action({kind:'setSelection',selection:caret(p(1,1))});const before=(await send({op:'inspectEditor',handle:a})).snapshot;
 const promoted=await edit({kind:'splitLine'});assert.equal(promoted.editor.undoDepth,before.undoDepth+1);
 assert.equal(promoted.editor.document.analysis.facts.grids,1);assert.equal(grid(promoted.editor).rows.length,2);
 assert.equal(cell(promoted.editor,0).entries[0].node.value,'a');assert.equal(cell(promoted.editor,1).entries[0].node.value,'b');
 await layout(a,revision);
 const savedTree=promoted.editor.document.analysis.tree;
 const undone=await action({kind:'undo'});assert.equal(undone.editor.document.source,'abc');assert.deepEqual(undone.editor.selection,before.selection);
 const redone=await action({kind:'redo'});assert.deepEqual(redone.editor.document.analysis.tree,savedTree);
 const oldRevision=revision;
 const next=await edit({kind:'splitLine'});assert.equal(grid(next.editor).rows.length,3);assert.equal(next.editor.document.analysis.facts.grids,1);
 const replay=await send({op:'editorAction',handle:a,expectedRevision:oldRevision,eventSequence:sequence,action:{kind:'edit',command:{kind:'splitLine'},groupId:null}});
 assert.ok(replay.result.replayed);assert.equal(replay.result.editor.document.revision,revision);assert.equal(grid(replay.result.editor).rows.length,3);
 await edit({kind:'insertText',text:'q',mode:'math'});const merged=await edit({kind:'mergeLine',direction:'left'});
 assert.equal(merged.editor.document.analysis.facts.grids,1);
 const owner=grid(merged.editor);await action({kind:'setSelection',selection:{kind:'node',nodeId:owner.id}});
 const copy=await edit({kind:'duplicateSelection'});const entries=copy.editor.document.analysis.tree.entries;
 assert.equal(entries.length,2);const originalIds=allIds(entries[0].node);for(const id of allIds(entries[1].node))assert.ok(!originalIds.has(id));
 await layout(a,revision);await action({kind:'undo'});
 const current=(await send({op:'inspectEditor',handle:a})).snapshot;
 const last=cell(current,grid(current).rows.length-1);await action({kind:'setSelection',selection:caret(p(last.id,last.entries.length))});
 for(let i=0;i<40;i++){
  await edit({kind:'splitLine'});const typed=await edit({kind:'insertText',text:String(i),mode:'math'});
  const s=await layout(a,revision);assert.equal(typed.editor.document.analysis.facts.grids,1);assert.equal(s.snapshot.gridColumns.length,1);
 }
 checks.push('One-owner split/merge/replay/independent-copy/history and 40 row additions without changing frame B');
 let currentDocument=(await send({op:'inspectDocument',handle:a})).snapshot;
 const editable=currentDocument.source.indexOf('q');
 assert.ok(editable>=0);
 for(let i=0;i<40;i++){
  const replacement=i%2?'q':'r';
  const patched=await send({op:'patchSource',handle:a,expectedRevision:revision,edit:{startUtf16:editable,endUtf16:editable+1,replacement}});
  revision=patched.snapshot.revision;await oracle(patched.snapshot);await layout(a,revision);
  assert.deepEqual((await send({op:'inspectEditor',handle:b})).snapshot,bBefore);
 }
 checks.push('40 source edits validate actual new cell content against fresh parser and regenerate current geometry');
 await send({op:'destroy',handle:a,expectedRevision:revision});await send({op:'destroy',handle:b,expectedRevision:0});
 for(const [i,source] of [String.raw`\begin{gathered}x`,String.raw`\begin{gathered}x\end{matrix}`,String.raw`\begin{gathered}a&b\end{gathered}`,String.raw`\begin{gathered}\unknown{x}\end{gathered}`].entries()){
  const h=(await send({op:'create',documentId:`contained-draft-${i}`,source})).handle;
  const d=(await send({op:'inspectDocument',handle:h})).snapshot;assert.equal(d.source,source);assert.equal(d.analysis.exportable,false);
  const error=d.analysis.complete?'unsupportedSyntax':'incompleteSource';
  await send({op:'layout',handle:h,expectedRevision:0,options:defaults,includeSvg:true},error);
  assert.equal((await send({op:'inspectDocument',handle:h})).snapshot.source,source);await send({op:'destroy',handle:h,expectedRevision:0});
 }
 const tooMany=String.raw`\begin{gathered}`+Array.from({length:257},()=> 'x').join('\\\\')+String.raw`\end{gathered}`;
 await send({op:'create',documentId:'overbudget-contained-grid',source:tooMany},'resourceLimit');
 assert.equal((await send({op:'stats'})).activeSessions,0);
 checks.push('Malformed environments, unknown commands and row budget preserve data and release sessions');
 assert.equal(sha(readFileSync(nativePath)),binaries.nativeSha256,'Native binary changed during review');
 assert.equal(sha(readFileSync(wasmPath)),binaries.wasmSha256,'WASM binary changed during review');
 report={status:'PASS',actualNativeExecution:true,actualWasmExecution:true,requests,scenes,sourceOracles,maxNumericDifference,
  exactSvgParity:true,...binaries,transcriptSha256:transcript.digest('hex'),fontHashes:hashes,checks,
  notCovered:['Matrix/cases/aligned','Real browser keys in this runner','Physical Windows IME','Office/installed UI']};
}catch(error){report={status:'FAILED',error:String(error),requests,scenes,sourceOracles,checks,...binaries};throw error;}
finally{
 if(report){writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(output,'report.json')).replaceAll('\\','/'),...report},null,2));}
 clearTimeout(timer);wasm.free();native.stdin.end();void lines.return?.();if(native.exitCode===null)native.kill();
}

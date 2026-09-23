// P08 Matrix foundation: actual Native/WASM command and SVG equivalence.
// Separate from gathered evidence; no Office or browser acceptance is implied.
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {createInterface} from 'node:readline';
import {once} from 'node:events';
import {createRequire} from 'node:module';
import {createHash} from 'node:crypto';
import {readFileSync,mkdirSync,writeFileSync} from 'node:fs';
import {dirname,resolve,relative} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const out=resolve(root,'target/matrix-parity-review',new Date().toISOString().replaceAll(':','-').replaceAll('.','-'));
mkdirSync(out,{recursive:true});
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const paths={math:resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),text:'C:/Windows/Fonts/NotoSansSC-VF.ttf'};
const hashes={math:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f',text:'763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a'};
const nativePath=resolve(root,'target/debug/visualtex-engine-test-cli'+(process.platform==='win32'?'.exe':''));
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
const timer=setTimeout(()=>native.kill(),100000);
let requests=0,scenes=0,sourceOracles=0,maxNumericDifference=0;
const transcript=createHash('sha256'),checks=[];
function equivalent(a,b,path='reply'){
 if(typeof a==='number'&&typeof b==='number'){
  assert.ok(Number.isFinite(a)&&Number.isFinite(b),path);const d=Math.abs(a-b);maxNumericDifference=Math.max(maxNumericDifference,d);assert.ok(d<=1e-9,path);return;
 }
 if(a===null||b===null||typeof a!=='object'||typeof b!=='object'){assert.equal(a,b,path);return;}
 assert.equal(Array.isArray(a),Array.isArray(b));assert.deepEqual(Object.keys(a),Object.keys(b),path);
 for(const key of Object.keys(a))equivalent(a[key],b[key],path+'.'+key);
}
async function send(command,error=null){
 const input=JSON.stringify({protocol:1,command});const a=JSON.parse(wasm.dispatch(input));
 if(!native.stdin.write(input+'\n'))await once(native.stdin,'drain');const line=await lines.next();assert.equal(line.done,false,stderr);
 const b=JSON.parse(line.value);equivalent(a,b,requests+':'+command.op);requests++;transcript.update(input+'\n'+line.value+'\n');
 if(error){assert.equal(a.status,'error',JSON.stringify(a));assert.equal(a.error.code,error);return a;}
 assert.equal(a.status,'ok',JSON.stringify(a));return a.result;
}
const defaults={mathFont:fonts.math.id,textFont:fonts.text.id,style:'display',fontSizePx:16,paddingEm:.125};
async function layout(handle,revision=0,options={}){
 const result=await send({op:'layout',handle,expectedRevision:revision,options:{...defaults,...options},includeSvg:true});scenes++;
 const s=result.snapshot;assert.equal(s.width,s.widthEm*s.fontSizePx);assert.equal(s.height,s.heightEm*s.fontSizePx);
 const ids=new Set(s.regions.map(r=>r.nodeId)),columns=new Set();assert.equal(ids.size,s.regions.length);
 for(const c of s.gridColumns){
  assert.ok(!ids.has(c.columnId)&&!columns.has(c.columnId));columns.add(c.columnId);
  const owner=s.regions.find(r=>r.nodeId===c.gridId);assert.equal(owner?.role,'grid');
  assert.ok(c.x>=owner.logical.xMin-1e-9&&c.x+c.width<=owner.logical.xMax+1e-9);
 }
 assert.ok(s.exportReady);assert.equal(result.svg.match(/<svg\b/g)?.length,1);
 for(const forbidden of ['<text','<image','foreignObject','<script','NaN','Infinity','href='])assert.ok(!result.svg.includes(forbidden));
 return result;
}
function semantic(value){
 if(Array.isArray(value))return value.map(semantic);
 if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value)
  .filter(([k])=>!['id','origin','grouping','leading','trailing'].includes(k)&&!k.startsWith('before'))
  .map(([k,v])=>[k,k==='columns'?v.length:semantic(v)]));
 return value;
}
function ids(node){
 const found=new Set();function visit(v){
  if(!v||typeof v!=='object')return;
  if(Array.isArray(v)){v.forEach(visit);return;}
  if(v.kind&&typeof v.id==='number'){assert.ok(!found.has(v.id));found.add(v.id);if(v.kind==='grid')for(const id of v.columns){assert.ok(!found.has(id));found.add(id);}}
  for(const [key,child] of Object.entries(v))if(key!=='origin'&&key!=='columns')visit(child);
 }visit(node);return found;
}
async function oracle(document){
 const fresh=await send({op:'analyzeSource',source:document.source});sourceOracles++;
 assert.equal(fresh.lossless,document.source);assert.deepEqual(fresh.analysis.facts,document.analysis.facts);
 assert.deepEqual(semantic(fresh.analysis.tree),semantic(document.analysis.tree),'Fresh parser differs: '+document.source);
 assert.equal(ids(document.analysis.tree).size,document.analysis.facts.nodes+document.analysis.facts.gridColumns);
}
const p=(nodeId,offset=0)=>({kind:'row',nodeId,offset,affinity:'after'});
const caret=position=>({kind:'range',anchor:position,focus:position});
const matrix=e=>e.document.analysis.tree.entries.find(e=>e.node.kind==='grid').node;
const cell=(e,row,column)=>matrix(e).rows[row].cells[column];
const content=e=>({source:e.document.source,tree:e.document.analysis.tree,selection:e.selection,documentId:e.document.documentId});
const generated=(r,c)=>String.raw`\begin{matrix}`+Array.from({length:r},(_,i)=>Array.from({length:c},(_,j)=>(i+j)%5===0?'':`a_{${i+j}}`).join('&')).join('\\\\')+String.raw`\end{matrix}`;
const sourceList=[
 String.raw`\begin{matrix}a&b\\c&d\end{matrix}`,
 generated(3,3),generated(16,16),generated(32,12),
 String.raw`\begin{aligned}a&=b+c\\d&=e+f\end{aligned}`,
 String.raw`\begin{align}long&=b&x&=y\\a&=cd&uvw&=z\end{align}`,
 String.raw`\begin{align*}a&=b&&c&=d\\x&=y&&z&=w\end{align*}`,
 String.raw`\begin{aligned}\frac{x}{y}&=\sqrt[3]{z}\\\sum_i^n x_i&=\int_0^1 x\,\mathrm{d}x\end{aligned}`,
 String.raw`\begin{cases}\begin{aligned}\text{速度}&=x\\\text{位移}&=y\end{aligned}&q\\\begin{matrix}a&b\\c&d\end{matrix}&p\end{cases}`,
 String.raw`\textcolor{blue}{\vphantom{\begin{aligned}a&=b\\\frac{x}{y}&=c\end{aligned}}}z`,
 String.raw`\begin{aligned}`+Array.from({length:32},(_,i)=>`x_{${i}}&=${i}&y_{${i}}&=${i+1}`).join('\\\\')+String.raw`\end{aligned}`,
 String.raw`\begin{cases}x^2&x\ge0\\-x&x<0\end{cases}`,
 String.raw`\begin{cases}\frac{x}{y}&\text{条件一}\\\sqrt[3]{x}&\text{条件二}\end{cases}`,
 String.raw`\textcolor{red}{\begin{cases}\begin{matrix}a&b\\c&d\end{matrix}&x>0\\\begin{gathered}p\\q\end{gathered}&x<0\end{cases}}`,
 String.raw`\vphantom{\begin{cases}a+b&c\\d&e\end{cases}}z`,
 String.raw`\begin{cases}`+Array.from({length:64},(_,i)=>`x_{${i}}&x>${i}`).join('\\\\')+String.raw`\end{cases}`,
 String.raw`\begin{matrix}\frac{x}{y}&\sqrt[3]{x_1}\\\sum_i^n x_i&\int_0^1 x\,\mathrm{d}x\end{matrix}`,
 String.raw`\begin{matrix}\text{速度}&\text{A\&B}\\\textcolor{red}{\boxed{x}}&\hat y\end{matrix}`,
 String.raw`\begin{matrix}\begin{matrix}a&b\\c&d\end{matrix}&x\\y&\begin{gathered}p\\q\end{gathered}\end{matrix}`,
 String.raw`\frac{\begin{matrix}a&b\\c&d\end{matrix}}{z}+\left[\begin{matrix}a&b\end{matrix}\right]`,
 String.raw`\vphantom{\begin{matrix}a&b\\c&d\end{matrix}}x`,
 String.raw`\hphantom{\begin{matrix}a&b\\c&d\end{matrix}}x`,
 String.raw`\begin{matrix}\!&\\x&y\end{matrix}`,
 "\\begin %opening\n {matrix}a% & \\\\\n&b\\\\c&d\\end %closing\n {matrix}",
];
let report;
try{
 for(const [i,source] of sourceList.entries()){
  const h=(await send({op:'create',documentId:'matrix-layout-'+i,source})).handle;
  const d=(await send({op:'inspectDocument',handle:h})).snapshot;assert.ok(d.analysis.complete&&d.analysis.supported);await oracle(d);
  for(const style of ['display','text','script','scriptScript']){
   const a=await layout(h,0,{style}),cached=await layout(h,0,{style}),double=await layout(h,0,{style,fontSizePx:32});
   assert.equal(cached.svg,a.svg);assert.equal(cached.snapshot.rebuiltFragments,0);assert.equal(double.snapshot.width,a.snapshot.width*2);assert.equal(double.snapshot.height,a.snapshot.height*2);
   assert.deepEqual(double.snapshot.gridColumns,a.snapshot.gridColumns);
  }
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 checks.push('24 grid sources including 7 aligned/align/align* and 5 cases formulas, paired/empty columns, 32 aligned rows, 64 case rows, 16x16/32x12 matrices; nesting, hidden paint, Unicode and comments in four styles/two natural sizes');
 const a=(await send({op:'create',documentId:'matrix-A',source:String.raw`\begin{matrix}a&b\\c&d\end{matrix}`})).handle;
 const b=(await send({op:'create',documentId:'independent-B',source:'z_2'})).handle;
 const bBefore=(await send({op:'inspectEditor',handle:b})).snapshot;
 let revision=0,sequence=0,current=(await send({op:'inspectEditor',handle:a})).snapshot;
 async function action(action){
  const r=(await send({op:'editorAction',handle:a,expectedRevision:revision,eventSequence:++sequence,action})).result;
  current=r.editor;revision=current.document.revision;await oracle(current.document);
  assert.deepEqual((await send({op:'inspectEditor',handle:b})).snapshot,bBefore,'Other frame changed');return r;
 }
 const edit=command=>action({kind:'edit',command,groupId:null});
 const select=(r,c,offset=0)=>action({kind:'setSelection',selection:caret(p(cell(current,r,c).id,offset))});
 await select(0,1,1);const initial=content(current),oldIds=ids(matrix(current));
 await edit({kind:'splitLine'});assert.equal(matrix(current).rows.length,3);assert.equal(matrix(current).columns.length,2);
 assert.equal(current.selection.focus.nodeId,cell(current,1,1).id);assert.ok([...oldIds].every(id=>ids(matrix(current)).has(id)));const inserted=content(current);
 await action({kind:'undo'});assert.deepEqual(content(current),initial);await action({kind:'redo'});assert.deepEqual(content(current),inserted);
 await edit({kind:'deleteGridRow'});assert.equal(matrix(current).rows.length,2);
 await select(0,1,0);const beforeColumn=content(current);
 await edit({kind:'insertGridColumn',after:false});assert.equal(matrix(current).columns.length,3);await edit({kind:'deleteGridColumn'});assert.equal(current.document.source,beforeColumn.source);
 await select(0,0,0);await edit({kind:'moveGridCell',direction:'right'});assert.equal(current.selection.focus.nodeId,cell(current,0,1).id);
 const geometry=await layout(a,revision);
 const vertical=await send({op:'queryScene',handle:a,expectedRevision:revision,sceneToken:geometry.sceneToken,query:{kind:'vertical',from:p(cell(current,0,1).id,0),down:true,desiredX:null}});
 assert.equal(vertical.result.position.nodeId,cell(current,1,1).id,'Vertical navigation moved to a different column');
 await select(0,0,0);const noDelete=content(current);await edit({kind:'deleteBackward'});assert.deepEqual(content(current),noDelete);
 await action({kind:'setSelection',selection:{kind:'range',anchor:p(cell(current,0,0).id,0),focus:p(cell(current,1,1).id,1)}});
 const crossCell=await send({op:'inspectEditor',handle:a});
 await send({op:'editorAction',handle:a,expectedRevision:revision,eventSequence:++sequence,action:{kind:'edit',command:{kind:'deleteBackward'},groupId:null}},'invalidPosition');
 await send({op:'copySelection',handle:a,expectedRevision:revision},'invalidPosition');
 assert.deepEqual(await send({op:'inspectEditor',handle:a}),crossCell);
 await select(1,1,1);
 for(let i=0;i<32;i++){
  await edit({kind:'insertGridRow',after:true});await edit({kind:'insertText',text:String(i),mode:'math'});await layout(a,revision);
 }
 for(let i=0;i<12;i++){
  await edit({kind:'insertGridColumn',after:i%2===0});await layout(a,revision);await edit({kind:'deleteGridColumn'});await layout(a,revision);
 }
 for(let i=0;i<24;i++){
  const atom=cell(current,0,0).entries[0].node;assert.equal(atom.value,i%2?'q':'a');
  const at=atom.origin.utf16Start;
  const patched=await send({op:'patchSource',handle:a,expectedRevision:revision,edit:{startUtf16:at,endUtf16:at+1,replacement:i%2?'a':'q'}});
  revision=patched.snapshot.revision;await oracle(patched.snapshot);current=(await send({op:'inspectEditor',handle:a})).snapshot;await layout(a,revision);
  assert.deepEqual((await send({op:'inspectEditor',handle:b})).snapshot,bBefore);
 }
 checks.push('Row/column CRUD, exact undo/redo, 32 row additions, 12 column cycles, 24 source edits, same-column vertical navigation, cross-cell refusal, unchanged frame B');
 await action({kind:'setSelection',selection:{kind:'node',nodeId:matrix(current).id}});
 const copied=(await send({op:'copySelection',handle:a,expectedRevision:revision})).fragment;
 assert.equal(copied.latex,current.document.source);
 const beforeDuplicate=matrix(current);await edit({kind:'duplicateSelection'});
 const entries=current.document.analysis.tree.entries;assert.equal(entries.length,2);const original=ids(entries[0].node);
 assert.ok([...ids(entries[1].node)].every(id=>!original.has(id)));assert.deepEqual(ids(beforeDuplicate),original);await layout(a,revision);
 await action({kind:'undo'});
 const newFrame=(await send({op:'create',documentId:'matrix-copy',source:''})).handle;
 const pasted=(await send({op:'editorAction',handle:newFrame,expectedRevision:0,eventSequence:1,action:{kind:'edit',command:{kind:'pasteLatex',latex:copied.latex},groupId:null}})).result.editor;
 await oracle(pasted.document);assert.equal(pasted.document.analysis.facts.grids,1);await layout(newFrame,pasted.document.revision);
 assert.deepEqual((await send({op:'inspectEditor',handle:b})).snapshot,bBefore);
 await send({op:'destroy',handle:newFrame,expectedRevision:pasted.document.revision});
 const rectangle=(r0,c0,r1,c1)=>action({kind:'setSelection',selection:{kind:'gridCells',gridId:matrix(current).id,
  anchorCell:cell(current,r0,c0).id,focusCell:cell(current,r1,c1).id}});
 await rectangle(1,1,0,0);
 const rectangularPayload=(await send({op:'copySelection',handle:a,expectedRevision:revision})).fragment.latex;
 assert.ok(rectangularPayload.startsWith(String.raw`\begin{matrix}`));
 for(let i=0;i<20;i++){
  await rectangle(2,0,3,1);const before=content(current),gridId=matrix(current).id,undoDepth=current.undoDepth;
  const cellIds=matrix(current).rows.map(r=>r.cells.map(c=>c.id));const columnIds=matrix(current).columns;
  await edit({kind:'clearGridCells'});await edit({kind:'pasteGridCells',latex:rectangularPayload});
  assert.equal(matrix(current).id,gridId);assert.deepEqual(matrix(current).columns,columnIds);
  assert.deepEqual(matrix(current).rows.map(r=>r.cells.map(c=>c.id)),cellIds);
  const scene=await layout(a,revision);
  const query=await send({op:'queryScene',handle:a,expectedRevision:revision,sceneToken:scene.sceneToken,
   query:{kind:'selection',selection:current.selection}});
  assert.equal(query.result.caret,null);assert.equal(query.result.rectangles.length,4);
  assert.ok([...ids(cell(current,2,0))].every(id=>!ids(cell(current,0,0)).has(id)));
  const after=content(current);await action({kind:'undo'});await action({kind:'redo'});assert.deepEqual(content(current),after);
  await action({kind:'undo'});
  // Clearing already-empty rectangles is a no-op and therefore does not add
  // a second undo item. Restore only edits that actually advanced history.
  while(current.undoDepth>undoDepth)await action({kind:'undo'});
  assert.deepEqual(content(current),before);
 }
 await rectangle(0,0,1,1);const rectangularBefore=await send({op:'inspectEditor',handle:a});
 await send({op:'editorAction',handle:a,expectedRevision:revision,eventSequence:++sequence,
  action:{kind:'edit',command:{kind:'pasteGridCells',latex:'q'},groupId:null}},'invalidInput');
 assert.deepEqual(await send({op:'inspectEditor',handle:a}),rectangularBefore);
 await select(0,0,0);await action({kind:'extendGridSelection',direction:'right'});await action({kind:'extendGridSelection',direction:'down'});
 assert.equal(current.selection.kind,'gridCells');await layout(a,revision);
 checks.push('20 rectangular clear/paste/history cycles, four actual highlight regions, same destination cell/column IDs, fresh copied descendants, dimension refusal and keyboard cell extension');
 await send({op:'destroy',handle:a,expectedRevision:revision});await send({op:'destroy',handle:b,expectedRevision:0});
 checks.push('Read-only whole-matrix clipboard, duplicate disjoint IDs, independent pasted document and one-owner layout');
 for(const [i,source] of [String.raw`\begin{matrix}a&b\\c\end{matrix}`,String.raw`\begin{matrix}a&b\\[2pt]c&d\end{matrix}`,String.raw`\begin{matrix}x`,String.raw`\begin{cases}x&y&z\end{cases}`,String.raw`\begin{aligned}[t]x&y\end{aligned}`].entries()){
  const h=(await send({op:'create',documentId:'matrix-draft-'+i,source})).handle;
  const d=(await send({op:'inspectDocument',handle:h})).snapshot;assert.equal(d.source,source);assert.equal(d.analysis.exportable,false);
  await send({op:'layout',handle:h,expectedRevision:0,options:defaults,includeSvg:true},d.analysis.complete?'unsupportedSyntax':'incompleteSource');
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 const ca=(await send({op:'create',documentId:'cases-transactions-A',source:'xy'})).handle;
 const cb=(await send({op:'create',documentId:'cases-independent-B',source:'z_2'})).handle;
 const cbBefore=(await send({op:'inspectEditor',handle:cb})).snapshot;
 let cs=(await send({op:'inspectEditor',handle:ca})).snapshot,cr=0,ce=0;
 const caseAction=async action=>{const result=(await send({op:'editorAction',handle:ca,expectedRevision:cr,eventSequence:++ce,action})).result;
  cs=result.editor;cr=cs.document.revision;await oracle(cs.document);assert.deepEqual((await send({op:'inspectEditor',handle:cb})).snapshot,cbBefore);return result;};
 const caseEdit=command=>caseAction({kind:'edit',command,groupId:null});
 await caseAction({kind:'setSelection',selection:{kind:'range',anchor:p(1,0),focus:p(1,2)}});const caseBefore=content(cs);
 await caseEdit({kind:'insertCases',rows:2});const insertedCases=content(cs);
 assert.equal(matrix(cs).environment,'cases');const caseOwner=matrix(cs).id,caseColumns=matrix(cs).columns.slice();
 await caseAction({kind:'undo'});assert.deepEqual(content(cs),caseBefore);await caseAction({kind:'redo'});assert.deepEqual(content(cs),insertedCases);
 await caseAction({kind:'setSelection',selection:caret(p(cell(cs,1,1).id,0))});
 for(let i=0;i<32;i++){await caseEdit({kind:'splitLine'});await caseEdit({kind:'insertText',text:String(i),mode:'math'});await layout(ca,cr);assert.equal(matrix(cs).id,caseOwner);assert.deepEqual(matrix(cs).columns,caseColumns);}
 for(const command of [{kind:'insertGridColumn',after:true},{kind:'deleteGridColumn'}]){const before=(await send({op:'inspectEditor',handle:ca})).snapshot;
  await send({op:'editorAction',handle:ca,expectedRevision:cr,eventSequence:++ce,action:{kind:'edit',command,groupId:null}},'invalidInput');assert.deepEqual((await send({op:'inspectEditor',handle:ca})).snapshot,before);}
 await caseAction({kind:'setSelection',selection:{kind:'gridCells',gridId:caseOwner,anchorCell:cell(cs,0,0).id,focusCell:cell(cs,1,1).id}});
 const caseCopy=(await send({op:'copySelection',handle:ca,expectedRevision:cr})).fragment;assert.ok(caseCopy.latex.startsWith('\\begin{matrix}'));
 const caseSelected=content(cs);await caseEdit({kind:'clearGridCells'});await caseEdit({kind:'pasteGridCells',latex:caseCopy.latex});await layout(ca,cr);
 assert.equal(matrix(cs).environment,'cases');assert.deepEqual(matrix(cs).columns,caseColumns);
 await caseAction({kind:'undo'});await caseAction({kind:'undo'});assert.deepEqual(content(cs),caseSelected);
 for(let i=0;i<20;i++){const atom=cell(cs,0,0).entries[0].node,offset=atom.origin.utf16Start;
  const patch=await send({op:'patchSource',handle:ca,expectedRevision:cr,edit:{startUtf16:offset,endUtf16:offset+1,replacement:i%2?'x':'q'}});
  cr=patch.snapshot.revision;await oracle(patch.snapshot);cs=(await send({op:'inspectEditor',handle:ca})).snapshot;await layout(ca,cr);assert.equal(matrix(cs).id,caseOwner);assert.deepEqual((await send({op:'inspectEditor',handle:cb})).snapshot,cbBefore);}
 await caseAction({kind:'setSelection',selection:{kind:'node',nodeId:caseOwner}});await caseEdit({kind:'duplicateSelection'});
 const copies=cs.document.analysis.tree.entries;const live=ids(copies[0].node);assert.ok([...ids(copies[1].node)].every(id=>!live.has(id)));await layout(ca,cr);
 await send({op:'destroy',handle:ca,expectedRevision:cr});await send({op:'destroy',handle:cb,expectedRevision:0});
 checks.push('Cases-specific actions: one-owner insertion/undo/redo, 32 internal rows, fixed-column rejection, rectangular clipboard, 20 source oracles, independent copied IDs and unchanged frame B');
 const aa=(await send({op:'create',documentId:'aligned-A',source:'a'})).handle,ab=(await send({op:'create',documentId:'aligned-independent-B',source:'z_2'})).handle;
 const abBefore=(await send({op:'inspectEditor',handle:ab})).snapshot;let as=(await send({op:'inspectEditor',handle:aa})).snapshot,ar=0,ae=0;
 const alignAction=async action=>{const r=(await send({op:'editorAction',handle:aa,expectedRevision:ar,eventSequence:++ae,action})).result;
  as=r.editor;ar=as.document.revision;await oracle(as.document);assert.deepEqual((await send({op:'inspectEditor',handle:ab})).snapshot,abBefore);return r;};
 const alignEdit=command=>alignAction({kind:'edit',command,groupId:null});
 const alignSelect=(r,c,o=0)=>alignAction({kind:'setSelection',selection:caret(p(cell(as,r,c).id,o))});
 await alignAction({kind:'setSelection',selection:{kind:'range',anchor:p(1,0),focus:p(1,1)}});
 const alignBefore=content(as);await alignEdit({kind:'insertAligned',rows:1,pairs:1});assert.equal(matrix(as).environment,'aligned');
 const alignCreated=content(as),alignId=matrix(as).id,alignColumns=matrix(as).columns.slice();
 await alignAction({kind:'undo'});assert.deepEqual(content(as),alignBefore);await alignAction({kind:'redo'});assert.deepEqual(content(as),alignCreated);
 await alignEdit({kind:'moveGridCell',direction:'right'});await alignEdit({kind:'insertText',text:'=b+c',mode:'math'});
 for(let i=0;i<20;i++){
  const before=content(as);await alignEdit({kind:'splitLine'});const newRow=matrix(as).rows.length-1;
  assert.equal(as.selection.focus.nodeId,cell(as,newRow,0).id);assert.equal(matrix(as).id,alignId);assert.deepEqual(matrix(as).columns,alignColumns);
  const split=content(as);await alignEdit({kind:'deleteBackward'});assert.equal(as.document.source,before.source);
  await alignAction({kind:'undo'});assert.deepEqual(content(as),split);
  await alignEdit({kind:'insertText',text:'d',mode:'math'});await alignEdit({kind:'moveGridCell',direction:'right'});await alignEdit({kind:'insertText',text:'=e+f',mode:'math'});
  const geometry=await layout(aa,ar);const equals=matrix(as).rows.map(r=>r.cells[1].entries[0].node.id);
  const xs=equals.map(id=>geometry.snapshot.regions.find(r=>r.nodeId===id).logical.xMin);assert.ok(xs.every(x=>Math.abs(x-xs[0])<1e-9),'Aligned equality anchors diverged');
 }
 for(let i=0;i<10;i++){await alignSelect(0,1,2);const before=content(as);await alignEdit({kind:'splitLine'});assert.equal(matrix(as).id,alignId);
  await alignSelect(1,0);await alignEdit({kind:'mergeLine',direction:'left'});assert.equal(as.document.source,before.source);await layout(aa,ar);}
 await alignAction({kind:'setSelection',selection:{kind:'gridCells',gridId:alignId,anchorCell:cell(as,0,0).id,focusCell:cell(as,1,1).id}});
 const alignPayload=(await send({op:'copySelection',handle:aa,expectedRevision:ar})).fragment.latex;
 const rectangleBefore=content(as);await alignEdit({kind:'clearGridCells'});await alignEdit({kind:'pasteGridCells',latex:alignPayload});assert.equal(as.document.source,rectangleBefore.source);
 assert.equal(matrix(as).id,alignId);assert.deepEqual(matrix(as).columns,alignColumns);await layout(aa,ar);
 for(let i=0;i<20;i++){const atom=cell(as,0,0).entries[0].node,offset=atom.origin.utf16Start;
  const patch=await send({op:'patchSource',handle:aa,expectedRevision:ar,edit:{startUtf16:offset,endUtf16:offset+1,replacement:i%2?'a':'q'}});ar=patch.snapshot.revision;await oracle(patch.snapshot);
  as=(await send({op:'inspectEditor',handle:aa})).snapshot;assert.equal(matrix(as).id,alignId);await layout(aa,ar);assert.deepEqual((await send({op:'inspectEditor',handle:ab})).snapshot,abBefore);}
 await alignAction({kind:'setSelection',selection:{kind:'node',nodeId:alignId}});await alignEdit({kind:'duplicateSelection'});
 const alignCopies=as.document.analysis.tree.entries,alignLive=ids(alignCopies[0].node);assert.ok([...ids(alignCopies[1].node)].every(id=>!alignLive.has(id)));await layout(aa,ar);
 await send({op:'destroy',handle:aa,expectedRevision:ar});await send({op:'destroy',handle:ab,expectedRevision:0});
 checks.push('Aligned-specific: 20 new rows with identical equality anchors, per-column split/merge and exact undo, 10 interior splits/merges, rectangle clipboard, 20 source edits, independent copies and unchanged frame B');
 assert.equal((await send({op:'stats'})).activeSessions,0);
 assert.equal(sha(readFileSync(nativePath)),binaries.nativeSha256);assert.equal(sha(readFileSync(wasmPath)),binaries.wasmSha256);
 report={status:'PASS',requests,scenes,sourceOracles,maxNumericDifference,exactSvgParity:true,actualNativeExecution:true,actualWasmExecution:true,
  ...binaries,fontHashes:hashes,transcriptSha256:transcript.digest('hex'),checks,
  notCovered:['Implicit ragged cells','Aligned placement and row-gap options','Browser input in this runner','Physical Windows IME','Office']};
}catch(error){report={status:'FAILED',error:String(error),requests,scenes,sourceOracles,checks,...binaries};throw error;}
finally{
 if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));}
 clearTimeout(timer);wasm.free();native.stdin.end();void lines.return?.();if(native.exitCode===null)native.kill();
}

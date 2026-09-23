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
const stamp=new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out=resolve(root,'target/matrix-variants-parity-review',stamp);mkdirSync(out,{recursive:true});
const sha=b=>createHash('sha256').update(b).digest('hex');
const paths={math:resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),text:'C:/Windows/Fonts/NotoSansSC-VF.ttf'};
const hashes={math:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f',text:'763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a'};
const nativePath=resolve(root,'target/debug/visualtex-engine-test-cli'+(process.platform==='win32'?'.exe':''));
const wasmPath=resolve(root,'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const wasm=new EngineRuntime(),fonts={};
for(const [name,path] of Object.entries(paths)){const bytes=readFileSync(path);assert.equal(sha(bytes),hashes[name]);const r=JSON.parse(wasm.registerFont(bytes,0));assert.equal(r.status,'ok');fonts[name]=r.result.font;}
const native=spawn(nativePath,['--font',paths.math,'--font',paths.text],{stdio:['pipe','pipe','pipe']});
const lines=createInterface({input:native.stdout,crlfDelay:Infinity})[Symbol.asyncIterator]();let stderr='';native.stderr.on('data',d=>stderr=(stderr+String(d)).slice(-16000));
const timer=setTimeout(()=>native.kill(),90000);let requests=0,scenes=0,sourceOracles=0,maxNumericDifference=0;const transcript=createHash('sha256');
function compare(a,b,path='reply'){if(typeof a==='number'&&typeof b==='number'){assert.ok(Number.isFinite(a)&&Number.isFinite(b),path);const d=Math.abs(a-b);maxNumericDifference=Math.max(maxNumericDifference,d);assert.ok(d<=1e-9,path);return;}if(a===null||b===null||typeof a!=='object'||typeof b!=='object'){assert.equal(a,b,path);return;}assert.equal(Array.isArray(a),Array.isArray(b),path);assert.deepEqual(Object.keys(a),Object.keys(b),path);for(const key of Object.keys(a))compare(a[key],b[key],path+'.'+key);}
async function send(command,error=null){const input=JSON.stringify({protocol:1,command});const a=JSON.parse(wasm.dispatch(input));if(!native.stdin.write(input+'\n'))await once(native.stdin,'drain');const line=await lines.next();assert.equal(line.done,false,stderr);const b=JSON.parse(line.value);compare(a,b,`${requests}:${command.op}`);requests++;transcript.update(input+'\n'+line.value+'\n');if(error){assert.equal(a.status,'error');assert.equal(a.error.code,error);return a;}assert.equal(a.status,'ok',JSON.stringify(a));return a.result;}
const defaults={mathFont:fonts.math.id,textFont:fonts.text.id,style:'display',fontSizePx:18,paddingEm:.125};
async function layout(handle,revision=0,options={}){const r=await send({op:'layout',handle,expectedRevision:revision,options:{...defaults,...options},includeSvg:true});scenes++;assert.ok(r.snapshot.exportReady);assert.ok(r.svg.startsWith('<svg'));for(const f of ['<text','<image','foreignObject','<script','NaN','Infinity'])assert.ok(!r.svg.includes(f));return r;}
function semantic(v){if(Array.isArray(v))return v.map(semantic);if(v&&typeof v==='object')return Object.fromEntries(Object.entries(v).filter(([k])=>!['id','origin','grouping','leading','trailing'].includes(k)&&!k.startsWith('before')).map(([k,x])=>[k,k==='columns'?x.length:semantic(x)]));return v;}
async function oracle(doc){const fresh=await send({op:'analyzeSource',source:doc.source});sourceOracles++;assert.equal(fresh.lossless,doc.source);assert.deepEqual(fresh.analysis.facts,doc.analysis.facts);assert.deepEqual(semantic(fresh.analysis.tree),semantic(doc.analysis.tree));}
const envs=['bmatrix','pmatrix','vmatrix','Bmatrix','Vmatrix'];
const sources=envs.flatMap(env=>[String.raw`\begin{${env}}a&b\\c&d\end{${env}}`,String.raw`\begin{${env}}\frac{x}{y}&\sqrt z\\\text{速度}&\hat q\end{${env}}`]);
const gridOf=s=>s.document.analysis.tree.entries.find(e=>e.node.kind==='grid').node;
const p=(nodeId,offset=0)=>({kind:'row',nodeId,offset,affinity:'after'});const caret=x=>({kind:'range',anchor:x,focus:x});
try{
 for(const [i,source] of sources.entries()){
  const h=(await send({op:'create',documentId:`matrix-variant-${i}`,source})).handle;let doc=(await send({op:'inspectDocument',handle:h})).snapshot;assert.ok(doc.analysis.complete&&doc.analysis.supported&&doc.analysis.exportable,source);await oracle(doc);const grid=doc.analysis.tree.entries[0].node;assert.equal(grid.kind,'grid');assert.ok(envs.includes(grid.environment.replace('Matrix','matrix'))||grid.environment);
  for(const style of ['display','text','script','scriptScript']){const a=await layout(h,0,{style}),b=await layout(h,0,{style}),big=await layout(h,0,{style,fontSizePx:36});assert.equal(a.svg,b.svg);assert.equal(b.snapshot.rebuiltFragments,0);assert.equal(big.snapshot.width,a.snapshot.width*2);assert.equal(big.snapshot.height,a.snapshot.height*2);assert.deepEqual(big.snapshot.gridColumns,a.snapshot.gridColumns);const owner=a.snapshot.regions.find(r=>r.nodeId===grid.id);assert.equal(owner.role,'grid');assert.ok(a.snapshot.gridColumns[0].x>owner.logical.xMin+1e-6,'wrapper did not shift first column');}
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 const h=(await send({op:'create',documentId:'variant-edit',source:String.raw`\begin{bmatrix}a&b\\c&d\end{bmatrix}`})).handle;let revision=0,seq=0,state=(await send({op:'inspectEditor',handle:h})).snapshot;
 async function action(action){const r=(await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++seq,action})).result;state=r.editor;revision=state.document.revision;await oracle(state.document);return r;}
 const edit=command=>action({kind:'edit',command,groupId:null});
 let grid=gridOf(state),survivor=grid.rows[0].cells[0].id,columnIds=grid.columns.slice();await action({kind:'setSelection',selection:caret(p(grid.rows[0].cells[1].id,1))});await edit({kind:'splitLine'});grid=gridOf(state);assert.equal(grid.environment,'bmatrix');assert.equal(grid.rows.length,3);assert.equal(grid.rows[0].cells[0].id,survivor);assert.deepEqual(grid.columns,columnIds);
 await edit({kind:'insertGridColumn',after:true});grid=gridOf(state);assert.equal(grid.columns.length,3);await edit({kind:'deleteGridColumn'});assert.equal(gridOf(state).columns.length,2);
 await action({kind:'setSelection',selection:{kind:'node',nodeId:gridOf(state).id}});const copy=(await send({op:'copySelection',handle:h,expectedRevision:revision})).fragment;assert.match(copy.latex,/\\begin\{bmatrix\}/u);await edit({kind:'duplicateSelection'});const grids=state.document.analysis.tree.entries.filter(e=>e.node.kind==='grid').map(e=>e.node);assert.equal(grids.length,2);const ids=n=>{const s=new Set();(function walk(v){if(!v||typeof v!=='object')return;if(Array.isArray(v)){v.forEach(walk);return;}if(typeof v.id==='number')s.add(v.id);for(const [k,x] of Object.entries(v))if(k!=='origin')walk(x);})(n);return s;};assert.ok([...ids(grids[1])].every(id=>!ids(grids[0]).has(id)));
 await layout(h,revision);await send({op:'destroy',handle:h,expectedRevision:revision});
 for(const source of [String.raw`\begin{bmatrix}a&b\\c\end{bmatrix}`,String.raw`\begin{pmatrix}a&b\\[2pt]c&d\end{pmatrix}`]){const x=(await send({op:'create',documentId:'variant-invalid-'+requests,source})).handle;const d=(await send({op:'inspectDocument',handle:x})).snapshot;assert.equal(d.source,source);assert.equal(d.analysis.exportable,false);await send({op:'layout',handle:x,expectedRevision:0,options:defaults,includeSvg:true},d.analysis.complete?'unsupportedSyntax':'incompleteSource');await send({op:'destroy',handle:x,expectedRevision:0});}
 const report={status:'PASS',requests,scenes,sourceOracles,maxNumericDifference,exactSvgParity:true,actualNativeExecution:true,actualWasmExecution:true,environments:envs,nativeSha256:sha(readFileSync(nativePath)),wasmSha256:sha(readFileSync(wasmPath)),transcriptSha256:transcript.digest('hex'),fontHashes:hashes,checks:['five matrix-family environments in four styles and 2x natural size','real wrapper glyphs shift shared column anchors','row/column edits retain variant and survivor ids','whole-grid copy/duplicate preserves source and allocates fresh ids','ragged/row-gap variants remain nonexportable']};writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));
}finally{clearTimeout(timer);wasm.free();native.stdin.end();lines.return?.();if(native.exitCode===null)native.kill();}

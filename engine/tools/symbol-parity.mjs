// P07.8: execute the same authored-atom, integral and explicit-font-fallback
// requests in native Rust and actual WASM. No production renderer is imported.
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {resolve,dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import {createHash} from 'node:crypto';
import {spawn} from 'node:child_process';
import {createInterface} from 'node:readline';
import {once} from 'node:events';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const sha=b=>createHash('sha256').update(b).digest('hex');
const symbols=readFileSync(resolve(root,'crates/core/src/symbol_catalog.rs'),'utf8').split('\n')
 .filter(line=>/^\s*"\\\\/u.test(line)&&line.includes('=>'))
 .flatMap(line=>[...line.split('=>')[0].matchAll(/"(?:\\.|[^"\\])*"/gu)].map(m=>JSON.parse(m[0])));
const integrals=[...readFileSync(resolve(root,'crates/core/src/rare_operators.rs'),'utf8')
 .matchAll(/RareOperator \{ command: ("(?:\\.|[^"\\])*"), glyph: '([^']+)' \}/gu)].map(m=>({command:JSON.parse(m[1]),glyph:m[2]}));
assert.ok(symbols.length>140);assert.equal(new Set(symbols).size,symbols.length);
assert.equal(integrals.length,25);assert.equal(new Set(integrals.map(entry=>entry.command)).size,integrals.length);
for(const [command,glyph] of [[String.raw`\dashint`,'⨍'],[String.raw`\ddashint`,'⨎'],[String.raw`\varointctrclockwise`,'∳']]){
 const matches=integrals.filter(entry=>entry.command===command);assert.equal(matches.length,1,command);assert.equal(matches[0].glyph,glyph,command);
}
const aliases=[...readFileSync(resolve(root,'crates/core/src/styles.rs'),'utf8').matchAll(/\(("(?:\\.|[^"\\])*"), MathVariant::[A-Za-z]+\)/gu)].map(m=>JSON.parse(m[1]));
assert.equal(aliases.length,17);
const fonts=[
 {path:resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),hash:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f'},
 {path:resolve(root,'target/test-resources/stix-two-2.13/STIXTwoMath-Regular.otf'),hash:'f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72'},
];
const nativePath=resolve(root,'target/debug/visualtex-engine-test-cli.exe');
const wasmPath=resolve(root,'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const runtime=new EngineRuntime();
const ids=fonts.map(font=>{const bytes=readFileSync(font.path);assert.equal(sha(bytes),font.hash);
 const response=JSON.parse(runtime.registerFont(bytes,0));assert.equal(response.status,'ok');return response.result.font.id;});
const native=spawn(nativePath,fonts.flatMap(font=>['--font',font.path]),{stdio:['pipe','pipe','pipe']});
const lines=createInterface({input:native.stdout,crlfDelay:Infinity})[Symbol.asyncIterator]();
let stderr='';native.stderr.on('data',data=>{stderr=(stderr+data).slice(-10000);});
const timeout=setTimeout(()=>native.kill(),90000);
let requests=0,scenes=0,maxDifference=0;const transcript=createHash('sha256');
function equivalent(a,b,path='result'){
 if(typeof a==='number'&&typeof b==='number'){
  assert.ok(Number.isFinite(a)&&Number.isFinite(b),path);maxDifference=Math.max(maxDifference,Math.abs(a-b));
  assert.ok(Math.abs(a-b)<=1e-9,`${path}: ${a} != ${b}`);return;
 }
 if(a===null||b===null||typeof a!=='object'||typeof b!=='object'){assert.equal(a,b,path);return;}
 assert.equal(Array.isArray(a),Array.isArray(b),path);assert.deepEqual(Object.keys(a),Object.keys(b),path);
 for(const k of Object.keys(a))equivalent(a[k],b[k],`${path}.${k}`);
}
async function send(command,error){
 const input=JSON.stringify({protocol:1,command});const a=JSON.parse(runtime.dispatch(input));
 if(!native.stdin.write(input+'\n'))await once(native.stdin,'drain');
 const next=await lines.next();assert.equal(next.done,false,stderr);const b=JSON.parse(next.value);
 equivalent(a,b,`request${requests++}`);transcript.update(input+'\n'+next.value+'\n');
 if(error){assert.equal(a.status,'error');assert.equal(a.error.code,error);return a;}
 assert.equal(a.status,'ok',JSON.stringify(a));return a.result;
}
const defaults={mathFont:ids[0],textFont:null,mathFallbacks:[ids[1]],style:'display',fontSizePx:24,paddingEm:0.125};
async function layout(handle,revision=0,options={},error){
 const r=await send({op:'layout',handle,expectedRevision:revision,options:{...defaults,...options},includeSvg:true},error);
 if(error)return r;
 scenes++;assert.ok(r.snapshot.exportReady&&r.snapshot.width>0&&r.snapshot.height>0);
 assert.ok(!/<(?:text|image|script)\b|foreignObject|href=|NaN|Infinity/u.test(r.svg));
 for(const resource of r.snapshot.resources){assert.ok(resource.shape.glyphId>0);assert.ok(r.snapshot.fonts.some(f=>f.id===resource.fontId));}
 return r;
}
try{
 for(const [i,command] of symbols.entries()){
  const h=(await send({op:'create',documentId:`atom-${i}`,source:`a${command} b`})).handle;
  for(const style of ['display','text','script','scriptScript']){
   const first=await layout(h,0,{style});const again=await layout(h,0,{style});
   assert.equal(again.snapshot.rebuiltFragments,0);assert.equal(again.svg,first.svg);
  }
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 for(const [i,command] of aliases.entries()){
  const source=`${command}{Az09}+x`;
  const h=(await send({op:'create',documentId:`style-alias-${i}`,source})).handle;
  for(const style of ['display','text','script','scriptScript']){
   const first=await layout(h,0,{style});const second=await layout(h,0,{style,fontSizePx:48});
   assert.equal(second.snapshot.width,first.snapshot.width*2);assert.equal(second.snapshot.rebuiltFragments,0);
   assert.deepEqual(second.snapshot.glyphs,first.snapshot.glyphs);
  }
  const p=(await send({op:'inspectDocument',handle:h})).snapshot;
  assert.equal(p.source,source);assert.equal(p.analysis.facts.styles,1);
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 for(const [i,entry] of integrals.entries()){
  for(const mode of ['',String.raw`\limits`,String.raw`\nolimits`,String.raw`\displaylimits`]){
   const source=`${entry.command}${mode}_{a}^{b}x`;
   const h=(await send({op:'create',documentId:`rare-${i}-${mode}`,source})).handle;
   for(const style of ['display','text','script','scriptScript'])await layout(h,0,{style});
   const p=(await send({op:'inspectDocument',handle:h})).snapshot;
   assert.equal(p.source,source);assert.equal(p.analysis.facts.largeOperators,1);
   await send({op:'destroy',handle:h,expectedRevision:0});
  }
 }
 const h=(await send({op:'create',documentId:'symbol-actions',source:''})).handle;
 let revision=0,sequence=0;
 const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
 const edit=command=>action({kind:'edit',command,groupId:null});
 await edit({kind:'insertSymbol',command:String.raw`\partial`});
 const typed=await edit({kind:'insertText',text:'x',mode:'math'});
 assert.equal(typed.editor.document.source,String.raw`\partial x`);assert.equal(typed.editor.document.analysis.facts.fractions,0);
 await edit({kind:'insertSymbol',command:String.raw`\subseteq`});await edit({kind:'insertSymbol',command:String.raw`\digamma`});
 const initial=await layout(h,revision);assert.ok(initial.snapshot.resources.some(r=>r.fontId===ids[1]));
 const snapshot=(await send({op:'inspectEditor',handle:h})).snapshot;
 const root=snapshot.document.analysis.tree.id,p={kind:'row',nodeId:root,offset:0,affinity:'after'};
 await action({kind:'setSelection',selection:{kind:'range',anchor:p,focus:{...p,offset:snapshot.document.analysis.tree.entries.length}}});
 const duplicated=await edit({kind:'duplicateSelection'});
 assert.equal(duplicated.editor.document.analysis.tree.entries.length,snapshot.document.analysis.tree.entries.length*2);
 await action({kind:'undo'});assert.equal((await layout(h,revision)).svg,initial.svg);
 await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:sequence+1,action:{kind:'edit',command:{kind:'insertSymbol',command:String.raw`\frac`}}},'unsupportedSyntax');
 await layout(h,revision,{mathFallbacks:[]},'invalidInput');
 assert.equal((await layout(h,revision)).svg,initial.svg);
 await layout(h,revision,{mathFallbacks:[ids[1],ids[1]]},'invalidInput');
 await send({op:'destroy',handle:h,expectedRevision:revision});
 assert.equal((await send({op:'stats'})).activeSessions,0);
 console.log(JSON.stringify({status:'PASS',actualNativeExecution:true,actualWasmExecution:true,
  atomCommands:symbols.length,rareIntegralCommands:integrals.length,styleAliases:aliases.length,requests,scenes,maxNumericDifference:maxDifference,exactSvgParity:true,
  nativeSha256:sha(readFileSync(nativePath)),wasmSha256:sha(readFileSync(wasmPath)),transcriptSha256:transcript.digest('hex'),
  fontHashes:fonts.map(f=>f.hash),scope:'Authored atoms, Unicode integral variants, explicit missing-glyph fallback, glyph provenance, edits/history. No Office or native IME.'},null,2));
}finally{clearTimeout(timeout);runtime.free();native.stdin.end();native.kill();}

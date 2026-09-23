// Actual Native/WASM fraction-family parity. Existing broad layout tests stay
// unchanged; this adds independent geometry, identities, edits and error checks.
import assert from 'node:assert/strict';
import {spawn} from 'node:child_process';
import {createInterface} from 'node:readline';
import {once} from 'node:events';
import {createRequire} from 'node:module';
import {createHash} from 'node:crypto';
import {readFileSync,mkdirSync,writeFileSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const paths={math:resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'),text:'C:/Windows/Fonts/NotoSansSC-VF.ttf'};
const expected={math:'6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f',text:'763146584cf0710223441356b4395e279021b0806c196614377a7a0174ae074a'};
const nativePath=resolve(root,`target/debug/visualtex-engine-test-cli${process.platform==='win32'?'.exe':''}`);
const wasmPath=resolve(root,'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const {EngineRuntime}=createRequire(import.meta.url)(resolve(root,'web/pkg-node/visualtex_engine_wasm.js'));
const wasm=new EngineRuntime();const fonts={};
for(const [key,path] of Object.entries(paths)) {
 const bytes=readFileSync(path);assert.equal(sha(bytes),expected[key]);
 const r=JSON.parse(wasm.registerFont(bytes,0));assert.equal(r.status,'ok');fonts[key]=r.result.font;
}
const native=spawn(nativePath,['--font',paths.math,'--font',paths.text],{stdio:['pipe','pipe','pipe']});
const lines=createInterface({input:native.stdout,crlfDelay:Infinity})[Symbol.asyncIterator]();
let stderr='';native.stderr.on('data',d=>{stderr=(stderr+String(d)).slice(-16000);});
const timer=setTimeout(()=>native.kill(),90000);
let requests=0,scenes=0,maxNumericDifference=0;const transcript=createHash('sha256');
function compare(a,b,path='response') {
 if(typeof a==='number'&&typeof b==='number') {const difference=Math.abs(a-b);assert.ok(Number.isFinite(a)&&Number.isFinite(b));maxNumericDifference=Math.max(maxNumericDifference,difference);assert.ok(difference<=1e-9,`${path}: ${a} vs ${b}`);return;}
 if(a===null||b===null||typeof a!=='object'||typeof b!=='object') {assert.equal(a,b,path);return;}
 assert.equal(Array.isArray(a),Array.isArray(b),path);assert.deepEqual(Object.keys(a),Object.keys(b),path);
 for(const key of Object.keys(a))compare(a[key],b[key],`${path}.${key}`);
}
async function send(command,error=null) {
 const input=JSON.stringify({protocol:1,command});const a=JSON.parse(wasm.dispatch(input));
 if(!native.stdin.write(input+'\n'))await once(native.stdin,'drain');const line=await lines.next();assert.equal(line.done,false,stderr);
 const b=JSON.parse(line.value);compare(a,b,`${requests}:${command.op}`);requests++;transcript.update(input+'\n'+line.value+'\n');
 if(error){assert.equal(a.status,'error');assert.equal(a.error.code,error);return a;}
 assert.equal(a.status,'ok',JSON.stringify(a));return a.result;
}
const defaults={mathFont:fonts.math.id,textFont:fonts.text.id,style:'display',fontSizePx:16,paddingEm:.125};
async function layout(handle,revision=0,options={},includeSvg=true,error=null) {
 const result=await send({op:'layout',handle,expectedRevision:revision,options:{...defaults,...options},includeSvg},error);
 if(error)return result;
 const s=result.snapshot;scenes++;assert.ok(s.width>0&&s.height>0);assert.equal(s.width,s.widthEm*s.fontSizePx);assert.equal(s.height,s.heightEm*s.fontSizePx);
 assert.equal(new Set(s.regions.map(r=>r.nodeId)).size,s.regions.length);
 for(const g of s.glyphs){const r=s.resources[g.resource];assert.ok(r&&r.shape.glyphId>0);assert.ok(s.fonts.some(f=>f.id===r.fontId));}
 if(includeSvg){assert.ok(s.exportReady);for(const text of ['<text','<image','foreignObject','<script','NaN','Infinity','href='])assert.ok(!result.svg.includes(text),text);}
 return result;
}
const commands=['binom','dbinom','tbinom','nicefrac'];
const sources=commands.flatMap(c=>[`\\${c}{n+1}{k}`,`\\${c} % top\n {n} % bottom\n {k}`,`\\textcolor{red}{\\${c}{n}{k}}`]);
sources.push(String.raw`\binom{\frac{n}{2}}{\sqrt{k}}`,String.raw`\nicefrac{\text{速度}}{y}`,String.raw`\binom{\nicefrac{a}{b}}{\binom{x}{y}}`,String.raw`\sqrt{\binom{n}{k}}`,String.raw`x^{\dbinom{n}{k}}`,String.raw`\phantom{\binom{n}{k}}+x`,String.raw`\binom{\!}{k}`);
try {
 for(const [i,source] of sources.entries()) {
  const h=(await send({op:'create',documentId:`fraction-layout-${i}`,source})).handle;
  for(const style of ['display','text','script','scriptScript']) {
   const first=await layout(h,0,{style});const cached=await layout(h,0,{style});
   assert.equal(cached.snapshot.rebuiltFragments,0);assert.equal(cached.svg,first.svg);
   const enlarged=await layout(h,0,{style,fontSizePx:32});assert.equal(enlarged.snapshot.width,first.snapshot.width*2);assert.equal(enlarged.snapshot.height,first.snapshot.height*2);
   assert.deepEqual(enlarged.snapshot.glyphs,first.snapshot.glyphs);assert.equal(enlarged.snapshot.rebuiltFragments,0);
  }
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 for(const [structure,command] of [['binomial',String.raw`\binom`],['skewedFraction',String.raw`\nicefrac`]]) {
  const h=(await send({op:'create',documentId:`fraction-editor-${structure}`,source:''})).handle;
  let revision=0,sequence=0;
  const action=async action=>{const r=await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:++sequence,action});revision=r.result.editor.document.revision;return r.result;};
  const edit=command=>action({kind:'edit',command,groupId:null});
  await edit({kind:'insertStructure',structure});await edit({kind:'insertText',text:'n',mode:'math'});
  await action({kind:'moveHorizontal',direction:'right',extend:false});
  const typed=await edit({kind:'insertText',text:'k',mode:'math'});assert.equal(typed.editor.document.source,`${command}{n}{k}`);
  const owner=typed.editor.document.analysis.tree.entries[0].node;const first=await layout(h,revision);
  assert.equal(first.snapshot.rules.filter(r=>r.nodeId===owner.id).length,0);
  const p={kind:'row',nodeId:owner.numerator.id,offset:0,affinity:'after'};
  const down=await send({op:'queryScene',handle:h,expectedRevision:revision,sceneToken:first.sceneToken,query:{kind:'vertical',from:p,down:true,desiredX:null}});
  assert.equal(down.result.position.nodeId,owner.denominator.id);
  await action({kind:'setSelection',selection:{kind:'node',nodeId:owner.id}});
  const duplicate=await edit({kind:'duplicateSelection'});const copy=duplicate.editor.document.analysis.tree.entries[1].node;
  assert.notEqual(copy.id,owner.id);assert.notEqual(copy.numerator.id,owner.numerator.id);assert.notEqual(copy.denominator.id,owner.denominator.id);
  await action({kind:'undo'});assert.equal((await layout(h,revision)).svg,first.svg);
  await action({kind:'redo'});await layout(h,revision);
  const input=duplicate.editor.document.source;const at=input.lastIndexOf('{n')+1;
  const patched=await send({op:'patchSource',handle:h,expectedRevision:revision,edit:{startUtf16:at,endUtf16:at+1,replacement:'q+1'}});revision=patched.snapshot.revision;
  assert.equal(patched.snapshot.analysis.tree.entries[0].node.numerator.entries[0].node.value,'n');
  assert.equal(patched.snapshot.analysis.tree.entries[1].node.numerator.entries[0].node.value,'q');await layout(h,revision);
  const before=await send({op:'inspectEditor',handle:h});
  await send({op:'editorAction',handle:h,expectedRevision:revision,eventSequence:sequence+1,action:{kind:'edit',command:{kind:'insertStructure',structure:'unimplementedGeneralFraction'}}},'invalidInput');
  assert.deepEqual(await send({op:'inspectEditor',handle:h}),before);
  await send({op:'destroy',handle:h,expectedRevision:revision});
 }
 for(const [i,source] of commands.flatMap(c=>[`\\${c}{}{k}`,`\\${c}{n}{}`]).entries()) {
  const h=(await send({op:'create',documentId:`fraction-empty-${i}`,source})).handle;const r=await layout(h,0,{},false);assert.equal(r.snapshot.exportReady,false);
  await layout(h,0,{},true,'incompleteSource');await send({op:'destroy',handle:h,expectedRevision:0});
 }
 for(const [i,source] of [String.raw`\binom{n}{`,String.raw`\nicefrac`,String.raw`\futurefrac{n}{k}`].entries()) {
  const h=(await send({op:'create',documentId:`fraction-invalid-${i}`,source})).handle;
  await layout(h,0,{},true,i===2?'unsupportedSyntax':'incompleteSource');assert.equal((await send({op:'inspectDocument',handle:h})).snapshot.source,source);
  await send({op:'destroy',handle:h,expectedRevision:0});
 }
 assert.equal((await send({op:'stats'})).activeSessions,0);
 const report={status:'PASS',actualNativeExecution:true,actualWasmExecution:true,requests,scenes,maxNumericDifference,exactSvgParity:true,
  nativeSha256:sha(readFileSync(nativePath)),wasmSha256:sha(readFileSync(wasmPath)),transcriptSha256:transcript.digest('hex'),fonts,
  scope:'Binomial/skewed fraction real slots, font geometry, source, copy/history and failed input. No browser, OS IME or Office actions.'};
 const output=resolve(root,'target/fraction-form-parity');mkdirSync(output,{recursive:true});
 writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));
 console.log(JSON.stringify(report,null,2));
}finally{clearTimeout(timer);wasm.free();native.stdin.end();lines.return?.();if(native.exitCode===null)native.kill();}

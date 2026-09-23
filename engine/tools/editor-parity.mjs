// Executes the real native process and real WASM instance in lockstep. Actions
// use returned node identities, while source/history/error expectations remain
// independent assertions. No browser, Office object or production app is used.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createInterface } from 'node:readline';
import { once } from 'node:events';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const nativePath = resolve(root, `target/debug/visualtex-engine-test-cli${process.platform === 'win32' ? '.exe' : ''}`);
const wasmPath = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const wasm = new EngineRuntime();
const native = spawn(nativePath, [], { stdio: ['pipe','pipe','pipe'] });
const lines = createInterface({ input: native.stdout, crlfDelay: Infinity })[Symbol.asyncIterator]();
let stderr = ''; native.stderr.on('data', data => { stderr = (stderr + data.toString()).slice(-10000); });
const timer = setTimeout(() => native.kill(), 45000);
const transcript = createHash('sha256');
let cases = 0;
const sessions = new Map();
const row = (nodeId, offset, affinity='after') => ({kind:'row',nodeId,offset,affinity});
const text = (nodeId, grapheme, affinity='after') => ({kind:'text',nodeId,grapheme,affinity});
const caret = position => ({kind:'range',anchor:position,focus:position});
const source = h => sessions.get(h).editor.document.source;
const doc = h => sessions.get(h).editor.document;
const tree = h => doc(h).analysis.tree;
const frame = h => {
  const e = sessions.get(h).editor;
  const {revision,documentId,...document} = e.document;
  return structuredClone({document,selection:e.selection});
};
function nodes(n, found=[]) {
  found.push(n);
  if (n.kind === 'row') for (const e of n.entries) nodes(e.node,found);
  if (n.kind === 'fraction') { nodes(n.numerator,found); nodes(n.denominator,found); }
  if (n.kind === 'scripts') { nodes(n.base,found); for (const s of n.slots) nodes(s.argument,found); }
  return found;
}
function validateEditor(e) {
  const all = nodes(e.document.analysis.tree);
  const map = new Map(all.map(n => [n.id,n]));
  assert.equal(all.length,map.size); assert.equal(all.length,e.document.analysis.facts.nodes);
  const validate = p => {
    const n=map.get(p.nodeId); assert.ok(n);
    if (p.kind==='row') { assert.equal(n.kind,'row'); assert.ok(p.offset>=0 && p.offset<=n.entries.length); }
    else { assert.equal(n.kind,'text'); const count=[...new Intl.Segmenter('en',{granularity:'grapheme'}).segment(n.text)].length;
      assert.ok(p.grapheme>=0 && p.grapheme<=count); }
  };
  if (e.selection.kind==='node') assert.ok(map.has(e.selection.nodeId));
  else { validate(e.selection.anchor); validate(e.selection.focus); }
  assert.ok(e.undoDepth<=128); assert.ok(e.redoDepth<=128); assert.ok(e.historyWeight<=32*1024*1024);
}
async function send(command, error) {
  const input=JSON.stringify({protocol:1,command});
  const wasmResult=JSON.parse(wasm.dispatch(input));
  if (!native.stdin.write(input+'\n')) await once(native.stdin,'drain');
  const next=await lines.next(); assert.equal(next.done,false,stderr);
  const nativeResult=JSON.parse(next.value);
  assert.deepEqual(wasmResult,nativeResult,`Native/WASM mismatch at ${cases}: ${command.op}`);
  transcript.update(input+'\n'+next.value+'\n'); cases++;
  if (error) { assert.equal(wasmResult.status,'error'); assert.equal(wasmResult.error.code,error); return wasmResult; }
  assert.equal(wasmResult.status,'ok',JSON.stringify(wasmResult));
  const result=wasmResult.result;
  if (result.kind==='session') sessions.set(result.handle,{revision:result.snapshot.revision,sequence:0,editor:null});
  if (result.kind==='editor') {
    validateEditor(result.snapshot); Object.assign(sessions.get(result.handle),{editor:result.snapshot,revision:result.snapshot.document.revision});
  }
  if (result.kind==='editorAction') {
    validateEditor(result.result.editor); Object.assign(sessions.get(result.handle),{editor:result.result.editor,revision:result.result.editor.document.revision});
  }
  if (result.kind==='document'||result.kind==='sourcePatched') sessions.get(result.handle).revision=result.snapshot.revision;
  return result;
}
async function create(name,initial='') {
  const result=await send({op:'create',documentId:name,source:initial});
  await send({op:'inspectEditor',handle:result.handle,expectedRevision:0});
  assert.equal(sessions.get(result.handle).editor.undoDepth,0); return result.handle;
}
async function action(h, action, error) {
  const session=sessions.get(h); session.sequence++;
  return send({op:'editorAction',handle:h,expectedRevision:session.revision,eventSequence:session.sequence,action},error);
}
const edit=(h,command,groupId=null,error)=>action(h,{kind:'edit',command,groupId},error);
const input=(h,value,mode='math',groupId=null)=>edit(h,{kind:'insertText',text:value,mode},groupId);
const select=(h,selection)=>action(h,{kind:'setSelection',selection});
const undo=h=>action(h,{kind:'undo'});
const redo=h=>action(h,{kind:'redo'});
try {
  const h=await create('owned-editor-parity');
  const original=frame(h);
  await edit(h,{kind:'insertStructure',structure:'fraction'});
  assert.equal(source(h),String.raw`\frac{}{}`);
  await input(h,'x'); await edit(h,{kind:'insertStructure',structure:'subscript'}); await input(h,'1');
  const fraction=tree(h).entries[0].node;
  await select(h,caret(row(fraction.denominator.id,0))); await input(h,'y');
  await edit(h,{kind:'insertStructure',structure:'superscript'}); await input(h,'2');
  assert.equal(source(h),String.raw`\frac{x_{1}}{y^{2}}`);
  const beforeCopy=frame(h); const originalIds=new Set(nodes(tree(h).entries[0].node).map(n=>n.id));
  await select(h,{kind:'node',nodeId:fraction.id});
  const beforeDuplicate=frame(h);
  await edit(h,{kind:'duplicateSelection'});
  const copies=nodes(tree(h).entries[1].node).map(n=>n.id);
  assert.ok(copies.every(id=>!originalIds.has(id))); assert.equal(tree(h).entries.length,2);
  const copied=frame(h); await undo(h); assert.deepEqual(frame(h),beforeDuplicate); await redo(h); assert.deepEqual(frame(h),copied);
  await select(h,caret(row(tree(h).id,2)));
  const selected=await edit(h,{kind:'deleteBackward'});
  assert.equal(selected.result.contentChanged,false); assert.equal(sessions.get(h).editor.selection.kind,'node');
  const oldSource=source(h), current=sessions.get(h);
  const repeatCommand={op:'editorAction',handle:h,expectedRevision:current.revision,eventSequence:current.sequence,
    action:{kind:'edit',command:{kind:'deleteBackward'},groupId:null}};
  const replay=await send(repeatCommand); assert.equal(replay.result.replayed,true); assert.equal(source(h),oldSource);
  await edit(h,{kind:'deleteBackward'}); assert.equal(tree(h).entries.length,1);
  await undo(h); assert.equal(tree(h).entries.length,2);

  const t=await create('owned-text-parity',String.raw`\text{A😀é}`);
  const textId=tree(t).entries[0].node.id;
  await select(t,caret(text(textId,1))); await edit(t,{kind:'splitText'});
  assert.equal(source(t),String.raw`\text{A}\text{😀é}`);
  await edit(t,{kind:'mergeText',direction:'left'}); assert.equal(source(t),String.raw`\text{A😀é}`);
  await select(t,{kind:'range',anchor:row(tree(t).id,0),focus:text(textId,1)});
  await edit(t,{kind:'deleteBackward'}); assert.equal(source(t),String.raw`\text{😀é}`);
  await undo(t); assert.equal(source(t),String.raw`\text{A😀é}`);
  await select(t,caret(text(textId,3)));
  const beforeIme=frame(t), beforeImeRevision=sessions.get(t).revision, beforeImeUndo=sessions.get(t).editor.undoDepth;
  let begin=await action(t,{kind:'beginComposition',mode:'text'}); let composition=begin.result.editor.composition.id;
  for (const preedit of ['s','su','速','速度']) {
    await action(t,{kind:'updateComposition',id:composition,preedit});
    assert.deepEqual(frame(t),beforeIme); assert.equal(sessions.get(t).revision,beforeImeRevision);
  }
  await action(t,{kind:'commitComposition',id:composition,text:null});
  assert.equal(source(t),String.raw`\text{A😀é速度}`); assert.equal(sessions.get(t).editor.undoDepth,beforeImeUndo+1);
  await undo(t); assert.deepEqual(frame(t),beforeIme);
  begin=await action(t,{kind:'beginComposition',mode:'text'}); composition=begin.result.editor.composition.id;
  await action(t,{kind:'updateComposition',id:composition,preedit:'失效输入'});
  const target=sessions.get(t).editor.selection;
  await action(t,{kind:'moveHorizontal',direction:'left',extend:false}); await select(t,target);
  await action(t,{kind:'commitComposition',id:composition,text:null},'conflict'); assert.deepEqual(frame(t),beforeIme);
  await action(t,{kind:'cancelComposition',id:composition});
  await edit(t,{kind:'pasteLatex',latex:String.raw`\future{x}`},null,'unsupportedSyntax'); assert.deepEqual(frame(t),beforeIme);

  const draft=await create('owned-draft-parity',String.raw`\frac{x}{y}`);
  const valid=frame(draft);
  await send({op:'replaceSource',handle:draft,expectedRevision:0,source:String.raw`\frac{x}{`});
  await send({op:'inspectEditor',handle:draft}); assert.equal(doc(draft).analysis.complete,false);
  const draftFrame=frame(draft);
  await edit(draft,{kind:'insertText',text:'z',mode:'math'},null,'incompleteSource');
  await undo(draft); assert.deepEqual(frame(draft),valid); await redo(draft); assert.deepEqual(frame(draft),draftFrame);
  const end=source(draft).length;
  await send({op:'patchSource',handle:draft,expectedRevision:sessions.get(draft).revision,edit:{startUtf16:end,endUtf16:end,replacement:'z}'}});
  await send({op:'inspectEditor',handle:draft}); assert.equal(source(draft),String.raw`\frac{x}{z}`);
  await undo(draft); assert.deepEqual(frame(draft),draftFrame);

  const grouped=await create('owned-grouped-typing');
  for (const c of ['A','\u0301','+','x']) await input(grouped,c,'math','same-group');
  assert.equal(source(grouped),'A\u0301+x'); assert.equal(sessions.get(grouped).editor.undoDepth,1);
  const groupedFrame=frame(grouped); await undo(grouped); assert.equal(source(grouped),''); await redo(grouped); assert.deepEqual(frame(grouped),groupedFrame);
  await action(grouped,{kind:'moveHorizontal',direction:'left',extend:false});
  await action(grouped,{kind:'moveHorizontal',direction:'right',extend:false});
  await input(grouped,'y','math','same-group'); assert.equal(sessions.get(grouped).editor.undoDepth,2);
  await edit(grouped,{kind:'insertText',text:'',mode:'math'},'', 'invalidInput');

  const history=await create('owned-history-parity');
  const before=[], after=[];
  for (let i=0;i<100;i++) {
    await select(history,caret(row(tree(history).id,tree(history).entries.length)));
    before.push(frame(history));
    if (i%4===0) await edit(history,{kind:'pasteLatex',latex:String.raw`\frac{x_{${i}}}{y}`});
    else if (i%4===1) await input(history,'+x');
    else if (i%4===2) await input(history,'速度😀','text');
    else await edit(history,{kind:'insertStructure',structure:'fraction'});
    after.push(frame(history));
  }
  for (let i=before.length-1;i>=0;i--) { await undo(history); assert.deepEqual(frame(history),before[i]); }
  for (let i=0;i<after.length;i++) { await redo(history); assert.deepEqual(frame(history),after[i]); }
  const unaffected=frame(t);
  await input(history,'+q'); assert.deepEqual(frame(t),unaffected);
  for (const [handle,s] of sessions) await send({op:'destroy',handle,expectedRevision:s.revision});
  const stats=await send({op:'stats'}); assert.equal(stats.activeSessions,0); assert.equal(wasm.activeSessions,0);
  const sha=p=>createHash('sha256').update(readFileSync(p)).digest('hex');
  console.log(JSON.stringify({status:'PASS',actualWasmExecution:true,requestCases:cases,
    fullHistoryStates:100,transcriptSha256:transcript.digest('hex'),
    nativeSha256:sha(nativePath),wasmSha256:sha(wasmPath),
    scope:'Pure mathematical transactions, source/history ownership and composition state; real platform IME and Office are NOT exercised.'},null,2));
} finally {
  clearTimeout(timer); wasm.free(); native.stdin.end(); lines.return?.();
  if (native.exitCode===null) native.kill();
}

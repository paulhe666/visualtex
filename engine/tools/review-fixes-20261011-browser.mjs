// Real browser keyboard input through the editor host for the 2026-10-11
// review fixes: modifier deletion, refused keystrokes, vertical keys with a
// selection. CDP key events are not physical-keyboard acceptance.
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {withOwnedBrowser} from './owned-browser-review.mjs';
import {createSandboxAssets} from './sandbox-assets.mjs';
const currentHostAssets=createSandboxAssets();
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const math=readFileSync(process.env.VISUALTEX_TEST_MATH_FONT??resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'));
const assets=new Map([
 ['/core.js',{type:'text/javascript',bytes:readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm.js'))}],
 ['/core.wasm',{type:'application/wasm',bytes:readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm_bg.wasm'))}],
 ['/math.otf',{type:'font/otf',bytes:math}],
]);
for(const [path,asset] of currentHostAssets.assets)if(path.startsWith('/host/')&&path.endsWith('.js'))assets.set(`/${path.slice('/host/'.length)}`,asset);
const CTRL=2,ALT=1,SHIFT=8;
const results=await withOwnedBrowser(assets,'<!doctype html><meta charset="utf-8"><title>VisualTeX review fixes</title><div id="mount" style="margin:40px"></div>',async({send,evaluate})=>{
 const tests=[];
 const key=async(key,code,virtual,modifiers=0)=>{
  await send('Input.dispatchKeyEvent',{type:'rawKeyDown',key,code,windowsVirtualKeyCode:virtual,nativeVirtualKeyCode:virtual,modifiers});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:virtual,nativeVirtualKeyCode:virtual,modifiers});
 };
 const type=async text=>{for(const ch of text)await send('Input.dispatchKeyEvent',{type:'keyDown',key:ch,text:ch,unmodifiedText:ch});};
 const open=source=>evaluate(`(async()=>{
   if(!window.runtime){const {default:init,EngineRuntime}=await import('/core.js');window.EditorHost=(await import('/editor-host.js')).EngineEditorHost;
     const data=await Promise.all(['/core.wasm','/math.otf'].map(p=>fetch(p).then(r=>r.arrayBuffer())));await init({module_or_path:data[0]});
     window.runtime=new EngineRuntime();const r=JSON.parse(runtime.registerFont(new Uint8Array(data[1]),0));if(r.status!=='ok')throw Error(JSON.stringify(r));window.mathFont=r.result.font.id;}
   window.host?.dispose();document.getElementById('mount').replaceChildren();window.errors=[];
   window.host=new EditorHost(runtime,document.getElementById('mount'),{documentId:'review-'+Math.random(),initialSource:${JSON.stringify(source)},
     layout:{mathFont,textFont:null,style:'display',fontSizePx:28,paddingEm:0.125},onError:e=>errors.push(e.message)});
   host.focus();return true;})()`);
 const state=()=>evaluate('({snapshot:host.inspect(),pending:host.input.pendingInput,errors})');
 const src=async()=>(await state()).snapshot.document.source;
 async function test(name,run){try{await run();tests.push({name,status:'PASS'});}catch(error){tests.push({name,status:'FAILED',error:String(error.stack??error)});}}
 const end=async()=>{await key('End','End',35);};
 for(const [label,modifiers] of [['Ctrl',CTRL],['Alt',ALT]])await test(`${label}+Backspace/Delete delete one atom and typing continues`,async()=>{
  await open('abc');await end();
  await key('Backspace','Backspace',8,modifiers);assert.equal(await src(),'ab');
  await key('Home','Home',36);await key('Delete','Delete',46,modifiers);assert.equal(await src(),'b');
  await type('x');const s=await state();assert.equal(s.snapshot.document.source,'xb');assert.equal(s.pending,null);assert.deepEqual(s.errors,[]);
 });
 await test('prime and tie type as atoms without a pending edit',async()=>{
  await open('f');await end();await type("'~x");const s=await state();
  assert.equal(s.snapshot.document.source,"f'~x");assert.equal(s.pending,null);assert.deepEqual(s.errors,[]);
 });
 await test('a refused keystroke is reported and the next key still types',async()=>{
  await open('a');await end();
  // Inject a core refusal for one character; every other action is real.
  await evaluate(`(()=>{const o=host.owner,real=o.action.bind(o);o.action=a=>{if(a.kind==='edit'&&a.command.kind==='insertText'&&a.command.text==='q')throw Error('injected refusal');return real(a);};})()`);
  await type('q');let s=await state();
  assert.equal(s.snapshot.document.source,'a');assert.equal(s.pending,null,'refused key held a pending edit');assert.equal(s.errors.length,1);
  await type('b');s=await state();assert.equal(s.snapshot.document.source,'ab');assert.equal(s.pending,null);
 });
 await test('ArrowUp with a selection collapses and moves up',async()=>{
  await open(String.raw`\frac{ab}{cd}`);
  const t=(await state()).snapshot.document.analysis.tree;const den=t.entries[0].node.denominator.id,num=t.entries[0].node.numerator.id;
  await evaluate(`host.restoreSelectionToken({...host.captureSelectionToken(),selection:{kind:'range',anchor:{kind:'row',nodeId:${den},offset:0,affinity:'after'},focus:{kind:'row',nodeId:${den},offset:2,affinity:'after'}}})`);
  await key('ArrowUp','ArrowUp',38);const sel=(await state()).snapshot.selection;
  assert.equal(sel.focus.nodeId,num);assert.deepEqual(sel.anchor,sel.focus);
 });
 await test('Shift+ArrowUp/Down extend the selection from its anchor',async()=>{
  await open(String.raw`\frac{ab}{cd}`);
  const t=(await state()).snapshot.document.analysis.tree;const den=t.entries[0].node.denominator.id,num=t.entries[0].node.numerator.id;
  await evaluate(`host.restoreSelectionToken({...host.captureSelectionToken(),selection:{kind:'range',anchor:{kind:'row',nodeId:${den},offset:2,affinity:'after'},focus:{kind:'row',nodeId:${den},offset:2,affinity:'after'}}})`);
  await key('ArrowUp','ArrowUp',38,SHIFT);let sel=(await state()).snapshot.selection;
  assert.equal(sel.anchor.nodeId,den);assert.equal(sel.anchor.offset,2);assert.equal(sel.focus.nodeId,num);
  await key('ArrowDown','ArrowDown',40,SHIFT);sel=(await state()).snapshot.selection;
  assert.equal(sel.anchor.nodeId,den);assert.equal(sel.focus.nodeId,den);
  assert.deepEqual((await state()).errors,[]);
 });
 // Client pixel of a caret stop (em coordinates mapped through the SVG CTM).
 const caretPoint=(nodeId,offset)=>evaluate(`(()=>{const scene=host.layoutSnapshot;const c=scene.carets.find(c=>c.position.nodeId===${nodeId}&&c.position.offset===${offset});
   const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(host.view.svg.getScreenCTM());return {x:p.x,y:p.y};})()`);
 const mouse=async(type,p,clickCount=1)=>send('Input.dispatchMouseEvent',{type,x:p.x,y:p.y,button:'left',buttons:type==='mouseReleased'?0:1,clickCount});
 await test('double click selects the slot under the pointer',async()=>{
  await open(String.raw`a+\frac{xy}{z}`);
  const t=(await state()).snapshot.document.analysis.tree;const num=t.entries[2].node.numerator.id;
  const p=await caretPoint(num,1);
  await mouse('mousePressed',p,1);await mouse('mouseReleased',p,1);await mouse('mousePressed',p,2);await mouse('mouseReleased',p,2);
  const s=await state();assert.deepEqual([s.snapshot.selection.anchor.nodeId,s.snapshot.selection.anchor.offset,s.snapshot.selection.focus.nodeId,s.snapshot.selection.focus.offset],[num,0,num,2]);
  assert.deepEqual(s.errors,[]);
 });
 await test('drag selection lands on the release point after many moves',async()=>{
  await open('abcdef');
  const root=(await state()).snapshot.document.analysis.tree.id;
  const a=await caretPoint(root,1),b=await caretPoint(root,5);
  await mouse('mousePressed',a);
  for(let i=1;i<=20;i++)await mouse('mouseMoved',{x:a.x+(b.x-a.x)*i/20,y:a.y});
  await mouse('mouseReleased',b);
  const sel=(await state()).snapshot.selection;
  assert.deepEqual([sel.anchor.offset,sel.focus.offset],[1,5]);
 });
 return {tests,failed:tests.filter(t=>t.status!=='PASS').length};
});
console.log(JSON.stringify(results,null,2));
assert.equal(results.failed,0,'review-fix browser checks failed');

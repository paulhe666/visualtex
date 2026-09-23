// Actual current sandbox events for pairing, exit, source recovery and fence SVG.
// No physical OS IME or production Office acceptance is claimed here.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/fence-browser');mkdirSync(out,{recursive:true});const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{} )});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const p=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...p,button:'left',clickCount:1});}
 async function type(text){for(const char of text){
  const map={'(':['Digit9',57,8],')':['Digit0',48,8],'[':['BracketLeft',219,0],']':['BracketRight',221,0],'{':['BracketLeft',219,8],'}':['BracketRight',221,8],'^':['Digit6',54,8]};
  if(map[char]){const [code,n,mod]=map[char];await key(char,code,n,mod,char);}else await send('Input.insertText',{text:char});
 }}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth});return s;}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('#focus-editor');await type('(x^2');const inside=await snapshot('paired bracket and script typed from blank');
  assert.equal(inside.document.canonicalLatex,String.raw`\left(x^{2}\right)`);
  await type(')');const exited=await snapshot('typed right bracket exits completed script and fence');
  assert.equal(exited.document.source,inside.document.source);assert.equal(exited.document.revision,inside.document.revision);
  assert.equal(exited.undoDepth,inside.undoDepth);assert.equal(exited.selection.focus.nodeId,exited.document.analysis.tree.id);
  await key('a','KeyA',65,2);await click('[data-fence="brackets"]');
  const wrapped=await snapshot('toolbar wraps existing selection preserving the expression');assert.equal(wrapped.document.canonicalLatex,String.raw`\left[\left(x^{2}\right)\right]`);
  await type(']');await type(')');assert.equal((await inspect()).document.canonicalLatex,String.raw`\left[\left(x^{2}\right)\right])`);
  await key('z','KeyZ',90,2);assert.equal((await snapshot('undo unmatched literal closer')).document.source,wrapped.document.source);
  await source('');await click('#focus-editor');await type('([x)');
  const mismatch=await snapshot('mismatched inner closer remains literal');assert.equal(mismatch.document.canonicalLatex,String.raw`\left(\left[x)\right]\right)`);
  await type('])');assert.equal((await snapshot('matching inner then outer closers only move the caret')).document.source,mismatch.document.source);
  assert.equal((await inspect()).undoDepth,mismatch.undoDepth);
  await source('');await click('[data-structure="text"]');await type('(x)');const text=await snapshot('text-mode brackets do not auto-pair');
  assert.equal(text.document.source,String.raw`\text{(x)}`);assert.equal(text.document.analysis.facts.fences,0);
  const tall=String.raw`\left\{\sqrt[3]{\dfrac{x+1}{\dfrac{y}{z}}}\right.`;
  await source(tall);const oneSide=await snapshot('single-sided fence around a tall indexed root');assert.equal(oneSide.document.analysis.exportable,true);
  const scene=await evaluate('window.__engineSandbox.layout()');const fence=oneSide.document.analysis.tree.entries[0].node;
  assert.equal(fence.right,'.');assert.ok(scene.glyphs.some(g=>g.nodeId===fence.id));
  await click('#export-svg');assert.ok((await evaluate('document.getElementById("export-result").textContent')).includes('<svg'));
  await source(String.raw`\left(x`);const draft=await snapshot('missing closer is retained as a source draft');assert.equal(draft.document.analysis.complete,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  await key('z','KeyZ',90,2);assert.equal((await snapshot('undo invalid fence draft')).document.source,tall);
  await click('#focus-editor');const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'single-sided-root.png'),Buffer.from(shot.data,'base64'));
  return{status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
   actualSandboxControls:true,typedBracketKeyEvents:true,steps,screenshot:'target/fence-browser/single-sided-root.png',
   notCovered:['Windows IME in this test','Office UI','production engine switch']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

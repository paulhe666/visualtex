// Fraction-family UI regression in an exclusively owned test browser.
// All mathematical changes use visible sandbox buttons and browser key/pointer
// events. Evaluated page expressions below only inspect state or geometry.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const output=resolve(root,'target/fraction-form-browser');mkdirSync(output,{recursive:true});
const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 const scene=()=>evaluate('window.__engineSandbox.layout()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function mouse(point){for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function click(selector){await mouse(await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)throw new Error('Missing sandbox control');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`));}
 async function type(text){for(const ch of text){
  if(/^[a-z]$/u.test(ch))await key(ch,`Key${ch.toUpperCase()}`,ch.toUpperCase().charCodeAt(0),0,ch);
  else if(/^[0-9]$/u.test(ch))await key(ch,`Digit${ch}`,ch.charCodeAt(0),0,ch);
  else if(ch==='+')await key('+','Equal',187,8,'+');
  else await send('Input.insertText',{text:ch});
 }}
 async function replaceDraft(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);
  steps.push({label,source:s.document.source,revision:s.document.revision,selection:s.selection,undoDepth:s.undoDepth});return s;}
 async function caretClick(nodeId,offset){await mouse(await evaluate(`(()=>{const s=window.__engineSandbox.layout();const c=s.carets.find(c=>c.position.kind==='row'&&c.position.nodeId===${nodeId}&&c.position.offset===${offset});if(!c)throw new Error('Missing fraction caret');
  const svg=document.querySelector('#editor svg');svg.scrollIntoView({block:'center'});const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(svg.getScreenCTM());const hit=document.elementFromPoint(p.x,p.y);if(!hit||!svg.contains(hit))throw new Error('Pointer target outside owned SVG');return {x:p.x,y:p.y};})()`));}
 async function vector(){await click('#export-svg');return evaluate('document.getElementById("export-result").textContent');}
 const owner=s=>s.document.analysis.tree.entries[0].node;
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  for(const [structure,command] of [['binomial',String.raw`\binom`],['skewedFraction',String.raw`\nicefrac`]]){
   if((await inspect()).document.source!=='')await replaceDraft('');
   await click(`[data-structure="${structure}"]`);await type('n+1');await key('ArrowRight','ArrowRight',39);await type('k');
   const built=await snapshot(`${structure}: constructed from blank using two editable arguments`);assert.equal(built.document.source,`${command}{n+1}{k}`);
   assert.equal(built.document.analysis.exportable,true);const f=owner(built);assert.equal(f.kind,'fraction');assert.notEqual(f.numerator.id,f.denominator.id);
   assert.equal((await scene()).rules.filter(r=>r.nodeId===f.id).length,0);
   await caretClick(f.numerator.id,1);assert.equal((await inspect()).selection.focus.nodeId,f.numerator.id);await type('2');
   assert.equal((await snapshot('pointer edit changes numerator only')).document.source,`${command}{n2+1}{k}`);
   await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,built.document.source);
   await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,`${command}{n2+1}{k}`);
   await caretClick(f.denominator.id,1);await key('ArrowUp','ArrowUp',38);assert.equal((await inspect()).selection.focus.nodeId,f.numerator.id);
   await key('ArrowDown','ArrowDown',40);assert.equal((await inspect()).selection.focus.nodeId,f.denominator.id);
   assert.ok((await vector()).includes('<path'));
  }
  await replaceDraft('n+1');await click('#focus-editor');await key('a','KeyA',65,2);
  const selected=await inspect();assert.notEqual(selected.selection.anchor.offset,selected.selection.focus.offset);
  const originalIds=selected.document.analysis.tree.entries.map(e=>e.node.id);
  await click('[data-structure="binomial"]');const wrapped=await snapshot('selected content wraps while retaining numerator identities');
  assert.equal(wrapped.document.source,String.raw`\binom{n+1}{}`);assert.deepEqual(owner(wrapped).numerator.entries.map(e=>e.node.id),originalIds);
  assert.equal(wrapped.selection.focus.nodeId,owner(wrapped).denominator.id);await type('k');
  const presentation=String.raw`\binom{n+1}{k}+\nicefrac{a+b}{c}+\textcolor{blue}{\binom{\frac{a}{b}}{\sqrt{x}}}`;
  await replaceDraft(presentation);const shown=await snapshot('nested colored forms render in one scene');assert.equal(shown.document.analysis.exportable,true);
  const exported=await vector(),layout=await scene();for(const bad of ['<text','<image','foreignObject','NaN','Infinity'])assert.ok(!exported.includes(bad));
  for(const dimension of ['width','height']){const n=Number(new RegExp(`\\b${dimension}="([^"]+)"`).exec(exported)?.[1]);assert.ok(Math.abs(n-layout[dimension])<1e-5);}
  assert.deepEqual(await evaluate('window.__engineSandbox.diagnostics()'),[],'Unexpected errors before deliberate draft');
  const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(output,'fraction-forms.png'),Buffer.from(shot.data,'base64'));
  await replaceDraft(String.raw`\binom{n}{`);assert.equal((await snapshot('incomplete denominator retained verbatim')).document.analysis.complete,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);await key('z','KeyZ',90,2);
  assert.equal((await snapshot('undo restores complete formula source')).document.source,presentation);
  const events=await evaluate('window.__engineSandbox.events()');
  return {status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
   actualSandboxControls:true,browserKeyAndPointerEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,steps,
   screenshot:'target/fraction-form-browser/fraction-forms.png',notCovered:['Physical OS IME','Office/Ribbon','installed application','P08 Shift+Enter']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(output,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});
writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

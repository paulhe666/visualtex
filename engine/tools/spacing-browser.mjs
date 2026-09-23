// P07.6 actual sandbox UI regression. No Office/installed application input.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/spacing-browser');mkdirSync(out,{recursive:true});const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 const scene=()=>evaluate('window.__engineSandbox.layout()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e||e.disabled)throw new Error('Unavailable control');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();const x=r.x+r.width/2,y=r.y+r.height/2;if(!e.contains(document.elementFromPoint(x,y)))throw new Error('Control is occluded');return{x,y};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function type(text){for(const ch of text){if(/[a-z]/u.test(ch))await key(ch,`Key${ch.toUpperCase()}`,ch.toUpperCase().charCodeAt(0),0,ch);
  else if(ch==='+')await key('+','Equal',187,8,'+');else await send('Input.insertText',{text:ch});}}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);
  steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth,selection:s.selection});return s;}
 async function svg(){await click('#export-svg');return evaluate('document.getElementById("export-result").textContent');}
 async function selectAllMath(){await click('#focus-editor');await key('a','KeyA',65,2);
  const s=await inspect();assert.equal(s.normalizedSelection.kind,'row');assert.equal(s.normalizedSelection.start,0);
  assert.equal(s.normalizedSelection.end,s.document.analysis.tree.entries.length,'Ctrl+A must create an actual mathematical selection');}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('[data-phantom="phantom"]');const blank=await snapshot('empty phantom is an editable draft');assert.equal(blank.document.analysis.exportable,false);
  await type('x');const filled=await snapshot('keyboard fills hidden body once');assert.equal(filled.document.source,String.raw`\phantom{x}`);
  assert.equal((await scene()).glyphs.length,0);assert.ok((await scene()).carets.length>2);
  assert.ok((await evaluate('document.querySelector(".vt-core-position").textContent')).includes('隐藏内容'));
  await key('ArrowRight','ArrowRight',39);assert.equal((await inspect()).document.revision,filled.document.revision);
  await click('[data-space="!"]');await click('#insert-boxed');await type('y');
  const composed=await snapshot('negative space and boxed visible sibling constructed from blank');assert.equal(composed.document.source,String.raw`\phantom{x}\!\boxed{y}`);
  const layout=await scene();assert.equal(layout.glyphs.length,1);assert.equal(layout.rules.length,4);
  const vector=await svg();assert.equal((vector.match(/<path /g)||[]).length,1);assert.equal((vector.match(/<rect /g)||[]).length,4);
  assert.ok(!vector.includes('opacity')&&!vector.includes('<text')&&!vector.includes('<image'));
  await click('#focus-editor');await key('z','KeyZ',90,2);assert.equal((await inspect()).document.analysis.exportable,false);
  await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,composed.document.source);assert.equal(await svg(),vector);

  await source('x+y');await selectAllMath();await click('[data-phantom="hphantom"]');
  const horizontal=await snapshot('selected expression wrapped as width-only phantom');assert.equal(horizontal.document.source,String.raw`\hphantom{x+y}`);
  const hscene=await scene();const hr=hscene.regions.find(r=>r.role==='phantom');assert.ok(hr.logical.xMax>hr.logical.xMin);assert.equal(hr.logical.yMax,hr.logical.yMin);
  assert.equal(hscene.glyphs.length,0);
  await click('#focus-editor');await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,'x+y');
  await selectAllMath();await click('#insert-boxed');assert.equal((await snapshot('boxed wraps the same selection')).document.source,String.raw`\boxed{x+y}`);
  assert.equal((await scene()).rules.length,4);

  await source(String.raw`\vphantom{\frac{x}{y}}+z`);const vscene=await scene();
  const vr=vscene.regions.find(r=>r.role==='phantom');assert.equal(vr.logical.xMax,vr.logical.xMin);assert.ok(vr.logical.yMax>vr.logical.yMin);
  assert.equal(vscene.rules.length,0);assert.equal(vscene.glyphs.length,2);
  await snapshot('height-only phantom has zero width and no hidden fraction rule');

  await source('x+y');const ordinary=(await scene()).advanceEm;
  await selectAllMath();await key('ArrowRight','ArrowRight',39);await click('[data-space=","]');
  const withSpace=await snapshot('explicit thin-space button changes geometry and source');assert.equal(withSpace.document.source,String.raw`x+y\,`);
  assert.ok(Math.abs((await scene()).advanceEm-ordinary-1/6)<1e-9);
  await key('Backspace','Backspace',8);assert.equal((await inspect()).document.source,'x+y');
  await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,withSpace.document.source);

  const shown=String.raw`\boxed{\int_0^1 x\,\mathrm{d}x}+\left.\vphantom{\frac{x}{y}}F\right|_0^1+\boxed{\phantom{x}+y}`;
  await source(shown);assert.equal((await snapshot('integral spacing, evaluation-height phantom and nested box')).document.analysis.exportable,true);
  const displayed=await svg();const saved=await inspect();
  for(const wanted of ['0.5','2']){await click('#zoom');await key(wanted==='0.5'?'Home':'End',wanted==='0.5'?'Home':'End',wanted==='0.5'?36:35);await key('Enter','Enter',13);
    assert.equal((await inspect()).document.revision,saved.document.revision);assert.equal(await svg(),displayed);}
  await click('#zoom');await key('Home','Home',36);await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
  const clip=await evaluate(`(()=>{const e=document.getElementById('editor-scroll');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:Math.max(0,r.x+scrollX),y:Math.max(0,r.y+scrollY),width:Math.min(r.width,1100),height:Math.min(r.height,360),scale:1};})()`);
  const image=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true,clip});writeFileSync(resolve(out,'spacing-phantom-boxed.png'),Buffer.from(image.data,'base64'));
  assert.deepEqual(await evaluate('window.__engineSandbox.diagnostics()'),[],'Unexpected errors before deliberately invalid source');
  await source(String.raw`\phantom{`);assert.equal((await snapshot('unfinished phantom source is retained')).document.analysis.complete,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  await key('z','KeyZ',90,2);assert.equal((await snapshot('undo returns prior valid source')).document.source,shown);
  const events=await evaluate('window.__engineSandbox.events()');
  return{status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
    actualSandboxControls:true,browserKeyAndPointerEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,steps,
    screenshot:'target/spacing-browser/spacing-phantom-boxed.png',notCovered:['Physical Windows IME','Office Ribbon or installed product','production UI integration']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

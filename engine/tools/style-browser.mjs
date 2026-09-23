// P07.5 real sandbox controls/key acceptance for mathematical styles.
// Launches only an owned test browser; no production application or Office process.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/style-browser');mkdirSync(out,{recursive:true});const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 const scene=()=>evaluate('window.__engineSandbox.layout()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)throw new Error('missing ${selector}');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function type(text){for(const ch of text){
  if(/^[a-zA-Z]$/u.test(ch)){const upper=ch.toUpperCase();await key(ch,`Key${upper}`,upper.charCodeAt(0),ch===upper?8:0,ch);}
  else if(ch==='+')await key('+','Equal',187,8,'+');
  else await send('Input.insertText',{text:ch});
 }}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);
  steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth,selection:s.selection});return s;}
 async function svg(){await click('#export-svg');return evaluate('document.getElementById("export-result").textContent');}
 async function selectLeft(count){await click('#focus-editor');for(let i=0;i<count;i++)await key('ArrowLeft','ArrowLeft',37,8);}
 async function selectRight(count){await click('#focus-editor');for(let i=0;i<count;i++)await key('ArrowRight','ArrowRight',39,8);}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('#focus-editor');await type('x+y');await selectLeft(3);await click('[data-style="mathbf"]');
  const bold=await snapshot('keyboard selection wrapped by bold style button');assert.equal(bold.document.source,String.raw`\mathbf{x+y}`);
  assert.equal(bold.document.analysis.facts.styles,1);assert.equal(bold.document.analysis.tree.entries[0].node.kind,'style');
  assert.equal(bold.document.analysis.tree.entries[0].node.variant,'bold');
  assert.ok((await evaluate('document.querySelector(".vt-core-position").textContent')).includes('样式'));
  const boldScene=await scene();const boldSvg=await svg();assert.ok(boldSvg.includes('<path'));
  const boldGlyphs=boldScene.glyphs.map(g=>boldScene.resources[g.resource].shape.glyphId);

  await click('#focus-editor');await key('z','KeyZ',90,2);const plain=await snapshot('undo returns exact unstyled source');assert.equal(plain.document.source,'x+y');
  const plainScene=await scene();const plainGlyphs=plainScene.glyphs.map(g=>plainScene.resources[g.resource].shape.glyphId);assert.notDeepEqual(plainGlyphs,boldGlyphs);
  await click('#focus-editor');await key('y','KeyY',89,2);assert.equal((await snapshot('redo restores same style source')).document.source,String.raw`\mathbf{x+y}`);
  assert.equal(await svg(),boldSvg);

  await source('');await click('[data-style="mathrm"]');let upright=await snapshot('collapsed upright style creates an editable empty body');
  assert.equal(upright.document.source,String.raw`\mathrm{}`);assert.equal(upright.document.analysis.exportable,false);
  await type('d');upright=await snapshot('typing fills upright style body');assert.equal(upright.document.source,String.raw`\mathrm{d}`);assert.equal(upright.document.analysis.exportable,true);

  await source('R');await selectRight(1);await click('[data-style="mathbb"]');const blackboard=await snapshot('blackboard style wraps selected R');
  assert.equal(blackboard.document.source,String.raw`\mathbb{R}`);const blackboardScene=await scene();
  assert.notDeepEqual(blackboardScene.glyphs.map(g=>blackboardScene.resources[g.resource].shape.glyphId),plainGlyphs.slice(0,1));

  await source(String.raw`\boldsymbol{\alpha}+\mathcal{A}+\mathfrak{H}+\sin x+\log y+\operatorname{rank}(A)`);
  const mixed=await snapshot('authored style and upright function source loads without fallback');assert.equal(mixed.document.analysis.supported,true);assert.equal(mixed.document.analysis.exportable,true);
  const vector=await svg();for(const bad of ['<text','<image','foreignObject','NaN','Infinity'])assert.ok(!vector.includes(bad),bad);
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'styles.png'),Buffer.from(screenshot.data,'base64'));
  assert.deepEqual(await evaluate('window.__engineSandbox.diagnostics()'),[],'Unexpected engine errors before deliberate incomplete draft');

  await source(String.raw`\mathbf{`);const draft=await snapshot('incomplete style is retained as source draft');assert.equal(draft.document.analysis.complete,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  assert.ok((await evaluate('window.__engineSandbox.diagnostics()')).some(message=>message.includes('incompleteSource')),'Draft layout refusal was not surfaced');
  await click('#source');await key('z','KeyZ',90,2);const recovered=await snapshot('source undo restores previous valid styled formula');assert.equal(recovered.document.source,mixed.document.source);

  const events=await evaluate('window.__engineSandbox.events()');
  return {status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
    actualSandboxControls:true,browserKeyAndPointerEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,steps,
    screenshot:'target/style-browser/styles.png',notCovered:['Physical Windows IME in this runner','Word or PowerPoint UI','installed application','production UI switch']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

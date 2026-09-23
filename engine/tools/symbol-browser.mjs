// P07.8 real isolated sandbox controls. CDP key/pointer input is browser input,
// not a claim of physical Windows IME or Word Ribbon acceptance.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets({symbolFallback:true});
const out=resolve(root,'target/symbol-browser');mkdirSync(out,{recursive:true});const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 const scene=()=>evaluate('window.__engineSandbox.layout()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e||e.disabled)throw new Error('Unavailable test control');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();const x=r.x+r.width/2,y=r.y+r.height/2;if(!e.contains(document.elementFromPoint(x,y)))throw new Error('Control is occluded');return {x,y};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function type(text){for(const ch of text){
  if(/^[a-z]$/u.test(ch))await key(ch,`Key${ch.toUpperCase()}`,ch.toUpperCase().charCodeAt(0),0,ch);
  else if(/^[0-9]$/u.test(ch))await key(ch,`Digit${ch}`,ch.charCodeAt(0),0,ch);
  else if(ch==='+')await key('+','Equal',187,8,'+');
  else await send('Input.insertText',{text:ch});
 }}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);
  steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth,selection:s.selection});return s;}
 async function svg(){await click('#export-svg');return evaluate('document.getElementById("export-result").textContent');}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('[data-symbol="partial"]');await type('x');
  const partial=await snapshot('partial symbol button does not insert a derivative fraction');
  assert.equal(partial.document.source,String.raw`\partial x`);assert.equal(partial.document.analysis.facts.fractions,0);
  await click('[data-symbol="subseteq"]');await click('[data-symbol="digamma"]');
  const before=await snapshot('relation and missing-primary glyph inserted by controls');
  assert.equal(before.document.source,String.raw`\partial x\subseteq\digamma`);
  const initial=await svg();const glyphScene=await scene();
  const fallback=glyphScene.fonts.find(f=>f.sha256==='f2076b9f1676438439dd41e23676f5ab99056e83d6b8f8c27841591ef2ccfa72');
  assert.ok(fallback);assert.ok(glyphScene.resources.some(r=>r.fontId===fallback.id));
  assert.ok(glyphScene.resources.some(r=>r.fontId!==fallback.id));
  await click('#focus-editor');await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,String.raw`\partial x\subseteq`);
  await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,before.document.source);assert.equal(await svg(),initial);
  await source('');await click('[data-operator="intbar"]');await type('0');await key('ArrowRight','ArrowRight',39);await type('1');
  await key('ArrowRight','ArrowRight',39);await type('x');
  const integral=await snapshot('rare integral built with independent lower/upper slots');assert.equal(integral.document.source,String.raw`\intbar_{0}^{1}x`);
  const integralScene=await scene();assert.ok(integralScene.resources.some(r=>r.fontId===fallback.id));
  await click('[data-symbol="Rightarrow"]');await type('y');await click('[data-symbol="otimes"]');await type('z');
  const final=await snapshot('regular symbols and fallback operator coexist');
  assert.equal(final.document.source,String.raw`\intbar_{0}^{1}x\Rightarrow y\otimes z`);
  const exported=await svg();assert.ok(!/<(?:text|image|script)\b|foreignObject|NaN|Infinity/u.test(exported));
  const prior=await inspect();await click('#zoom');await key('End','End',35);await key('Enter','Enter',13);
  assert.equal((await inspect()).document.source,prior.document.source);assert.equal((await inspect()).document.revision,prior.document.revision);
  assert.equal(await svg(),exported,'Viewport zoom changes natural export');
  await source('Az09');await click('#focus-editor');await key('a','KeyA',65,2);
  const selected=await inspect();assert.notDeepEqual(selected.selection.anchor,selected.selection.focus,'Style wrapping needs an actual selection');
  const unstyled=await svg();await click('[data-style="symbfsfit"]');
  const sans=await snapshot('Unicode sans bold italic wraps selected math through existing host');
  assert.equal(sans.document.source,String.raw`\symbfsfit{Az09}`);assert.notEqual(await svg(),unstyled);
  await click('#focus-editor');await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,'Az09');
  await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,sans.document.source);
  await source('');await click('[data-style="symbfcal"]');await type('z');
  assert.equal((await snapshot('Unicode bold script blank body accepts real typing')).document.source,String.raw`\symbfcal{z}`);
  await svg();
  const presentation=String.raw`\digamma\subseteq\mathbb{R}\quad\intbar_0^1 x\,\mathrm{d}x+\sumint\limits_a^b y+\textcolor{blue}{\symbfsfit{A}\otimes \symbfcal{B}\Rightarrow C}`;
  await source(presentation);await snapshot('mixed glyph/style/color presentation');await svg();
  assert.deepEqual(await evaluate('window.__engineSandbox.diagnostics()'),[],'Unexpected engine errors before deliberate draft');
  await evaluate('document.querySelector("#editor svg").scrollIntoView({block:"center"})');
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'symbols.png'),Buffer.from(screenshot.data,'base64'));
  await source(String.raw`\unknownSymbol{x}`);const invalid=await snapshot('unrecognized command remains preserved and unexportable');
  assert.equal(invalid.document.analysis.supported,false);assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  await key('z','KeyZ',90,2);assert.equal((await snapshot('undo restores supported mixed formula')).document.source,presentation);
  const events=await evaluate('window.__engineSandbox.events()');
  return {status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
    actualSandboxControls:true,browserKeyAndPointerEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,
    explicitFallbackSha256:fallback.sha256,steps,screenshot:'target/symbol-browser/symbols.png',
    notCovered:['Physical Windows IME','Word/PowerPoint Ribbon','production entrypoint','P08 Shift+Enter implementation']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

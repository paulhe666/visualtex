// Current-host P07.3 acceptance through actual browser keys, pointer and UI.
// Does not attach to user browsers, invoke Office, or claim physical Windows IME.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/accent-browser');mkdirSync(out,{recursive:true});const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function type(text){for(const ch of text){
  const letter=/^[a-z]$/u.test(ch);const digit=/^[0-9]$/u.test(ch);
  if(letter)await key(ch,`Key${ch.toUpperCase()}`,ch.toUpperCase().charCodeAt(0),0,ch);
  else if(digit)await key(ch,`Digit${ch}`,ch.charCodeAt(0),0,ch);
  else if(ch==='+')await key('+','Equal',187,8,'+');
  else await send('Input.insertText',{text:ch});
 }}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);
  steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth,selection:s.selection});return s;}
 const children=n=>n.kind==='row'?n.entries.map(e=>e.node):n.kind==='accent'||n.kind==='fence'?[n.body]:n.kind==='annotation'?[n.annotation,n.base]:
  n.kind==='fraction'?[n.numerator,n.denominator]:n.kind==='scripts'?[n.base,...n.slots.map(s=>s.argument)]:n.kind==='radical'?[...(n.degree?[n.degree]:[]),n.radicand]:[];
 function ids(n,out=new Set()){assert.ok(!out.has(n.id));out.add(n.id);for(const child of children(n))ids(child,out);return out;}
 async function caretClick(nodeId,offset){
  const point=await evaluate(`(()=>{const s=window.__engineSandbox.layout();const c=s.carets.find(c=>c.position.kind==='row'&&c.position.nodeId===${nodeId}&&c.position.offset===${offset});if(!c)throw new Error('Missing current caret');
   const svg=document.querySelector('#editor svg');svg.scrollIntoView({block:'center'});const m=svg.getScreenCTM();
   const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(m);return {x:p.x,y:p.y};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});
 }
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('[data-accent="hat"]');await click('[data-accent="vec"]');await type('x');
  const nested=await snapshot('nested accents built from blank by toolbar and actual keys');
  assert.equal(nested.document.canonicalLatex,String.raw`\hat{\vec{x}}`);
  assert.ok((await evaluate('document.querySelector(".vt-core-position").textContent')).includes('重音正文'));
  await key('ArrowRight','ArrowRight',39);const exited=await snapshot('right arrow exits completed nested accents');
  assert.equal(exited.document.revision,nested.document.revision);assert.equal(exited.undoDepth,nested.undoDepth);
  assert.equal(exited.selection.focus.nodeId,exited.document.analysis.tree.id);
  await type('+y');assert.equal((await snapshot('typing after the nested structure')).document.canonicalLatex,String.raw`\hat{\vec{x}}+y`);
  await key('a','KeyA',65,2);await click('[data-accent="overline"]');
  assert.equal((await snapshot('range wrapping retains the existing expression')).document.canonicalLatex,String.raw`\overline{\hat{\vec{x}}+y}`);
  await key('z','KeyZ',90,2);assert.equal((await inspect()).document.canonicalLatex,String.raw`\hat{\vec{x}}+y`);

  await source('');await click('[data-annotation="overset"]');await type('a');
  await key('ArrowRight','ArrowRight',39);await type('x');
  const annotated=await snapshot('annotation label then base filled by keyboard');assert.equal(annotated.document.source,String.raw`\overset{a}{x}`);
  const annotation=annotated.document.analysis.tree.entries[0].node;
  await caretClick(annotation.annotation.id,1);await type('b');
  assert.equal((await snapshot('actual pointer selects upper annotation slot')).document.source,String.raw`\overset{ab}{x}`);
  await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,annotated.document.source);
  await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,String.raw`\overset{ab}{x}`);

  await source('');await click('[data-accent="underbrace"]');await type('x+y');await click('#exit-structure');
  await click('[data-structure="subscript"]');await type('n');
  const brace=await snapshot('underbrace and lower label constructed by actual controls');
  assert.equal(brace.document.canonicalLatex,String.raw`\underbrace{x+y}_{n}`);
  await click('#export-svg');assert.ok((await evaluate('document.getElementById("export-result").textContent')).includes('<svg'));
  const oldRevision=brace.document.revision;const beforeZoom=await evaluate('window.__engineSandbox.stats()');
  // Use real select-control key input, not a script that invokes editor actions.
  await click('#zoom');await key('Home','Home',36);await key('Enter','Enter',13);
  assert.equal((await inspect()).document.revision,oldRevision);assert.equal((await evaluate('window.__engineSandbox.stats()')).sceneBuilds,beforeZoom.sceneBuilds);
  await click('#zoom');await key('End','End',35);await key('Enter','Enter',13);
  assert.equal((await inspect()).document.revision,oldRevision);
  await click('#zoom');await key('Home','Home',36);await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);

  for(const name of ['overgroup','undergroup','overleftharpoon','overrightharpoon','overlinesegment','underlinesegment']) {
    await source('');await click(`[data-accent="${name}"]`);await type('ab');
    const added=await snapshot(`${name}: real insertion and body typing`);assert.equal(added.document.source,`\\${name}{ab}`);
    const mark=added.document.analysis.tree.entries[0].node;
    const before=await inspect();await key('ArrowRight','ArrowRight',39);const exited=await inspect();
    assert.equal(exited.document.revision,before.document.revision);assert.equal(exited.undoDepth,before.undoDepth);
    assert.equal(exited.selection.focus.nodeId,exited.document.analysis.tree.id);
    await caretClick(mark.body.id,1);await type('z');assert.equal((await inspect()).document.source,`\\${name}{azb}`);
    await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,added.document.source);
    await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,`\\${name}{azb}`);
    const markScene=await evaluate('window.__engineSandbox.layout()');
    if(name.endsWith('linesegment')) assert.equal(markScene.rules.filter(r=>r.nodeId===mark.id).length,3);
    else assert.ok(markScene.glyphs.some(g=>g.nodeId===mark.id));
  }
  await source(String.raw`\overgroup{a+b}+\underlinesegment{c+d}+\overleftharpoon{x+y}+\undergroup{z}`);
  await click('#export-svg');
  const extensionShot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
  writeFileSync(resolve(out,'accent-extensions.png'),Buffer.from(extensionShot.data,'base64'));

  const long=String.raw`\overset{n}{\overbrace{a+b+c+d+e+f+g+h+i+j}}+\hat{\vec{x}}`;
  await source(long);const display=await snapshot('font-assembled wide decoration and nested accents');assert.equal(display.document.analysis.exportable,true);
  const scene=await evaluate('window.__engineSandbox.layout()');const overbrace=display.document.analysis.tree.entries[0].node.base.entries[0].node;
  assert.ok(scene.glyphs.filter(g=>g.nodeId===overbrace.id).length>1);
  assert.equal(ids(display.document.analysis.tree).size,display.document.analysis.facts.nodes);
  await click('#export-svg');const svg=await evaluate('document.getElementById("export-result").textContent');
  for(const forbidden of ['<text','<image','<script','foreignObject','NaN','Infinity'])assert.ok(!svg.includes(forbidden));
  const errorsBefore=await evaluate('window.__engineSandbox.diagnostics()');assert.deepEqual(errorsBefore,[],'Unexpected editing/layout errors before deliberate invalid drafts');
  const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'accents-and-annotations.png'),Buffer.from(shot.data,'base64'));
  await source(String.raw`\overset{a}{`);const draft=await snapshot('unclosed annotation source remains a recoverable draft');
  assert.equal(draft.document.analysis.complete,false);assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  await key('z','KeyZ',90,2);assert.equal((await snapshot('undo returns the exact valid source')).document.source,long);
  const events=await evaluate('window.__engineSandbox.events()');
  return {status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
   actualSandboxControls:true,browserKeyEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,
   steps,screenshot:'target/accent-browser/accents-and-annotations.png',
   notCovered:['Physical Windows IME in this test','Office UI','installed application','production engine switch']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

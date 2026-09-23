// P07.4 real sandbox controls and browser key/pointer acceptance. Only an owned
// test browser is launched; no production application or Office process is used.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/operator-browser');mkdirSync(out,{recursive:true});const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 const scene=()=>evaluate('window.__engineSandbox.layout()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function type(text){for(const ch of text){
  if(/^[a-z]$/u.test(ch))await key(ch,`Key${ch.toUpperCase()}`,ch.toUpperCase().charCodeAt(0),0,ch);
  else if(/^[0-9]$/u.test(ch))await key(ch,`Digit${ch}`,ch.charCodeAt(0),0,ch);
  else if(ch==='+')await key('+','Equal',187,8,'+');
  else if(ch==='=')await key('=','Equal',187,0,'=');
  else await send('Input.insertText',{text:ch});
 }}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);
  steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth,selection:s.selection});return s;}
 async function caretClick(nodeId,offset){const point=await evaluate(`(()=>{const s=window.__engineSandbox.layout();const c=s.carets.find(c=>c.position.kind==='row'&&c.position.nodeId===${nodeId}&&c.position.offset===${offset});if(!c)throw new Error('Missing current bound caret');
   const svg=document.querySelector('#editor svg');svg.scrollIntoView({block:'center'});const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(svg.getScreenCTM());return {x:p.x,y:p.y};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function style(inline){await click('#math-style');await key(inline?'End':'Home',inline?'End':'Home',inline?35:36);await key('Enter','Enter',13);}
 const operator=s=>s.document.analysis.tree.entries[0].node;
 const bound=(s,name)=>operator(s).parts.find(p=>p.kind==='script'&&p.slot.script===name).slot.argument;
 const opGlyph=(layout,nodeId)=>{const g=layout.glyphs.find(g=>g.nodeId===nodeId);assert.ok(g);return layout.resources[g.resource].shape.glyphId;};
 async function svg(){await click('#export-svg');return evaluate('document.getElementById("export-result").textContent');}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('[data-operator="sum"]');await type('i=1');await key('ArrowRight','ArrowRight',39);await type('n');
  await key('ArrowRight','ArrowRight',39);await type('x');await click('[data-structure="subscript"]');await type('i');
  const sum=await snapshot('sum constructed from empty input: lower bound, upper bound, operand');
  assert.equal(sum.document.source,String.raw`\sum_{i=1}^{n}x_{i}`);
  const opId=operator(sum).id,lower=bound(sum,'subscript').id,upper=bound(sum,'superscript').id;
  await caretClick(lower,0);assert.ok((await evaluate('document.querySelector(".vt-core-position").textContent')).includes('下限'));
  await key('ArrowUp','ArrowUp',38);assert.equal((await snapshot('up stays in the same operator upper limit')).selection.focus.nodeId,upper);
  await key('ArrowDown','ArrowDown',40);assert.equal((await snapshot('down returns to the same lower limit')).selection.focus.nodeId,lower);
  const displayScene=await scene(),beforeStyle=await inspect();const displayGlyph=opGlyph(displayScene,opId);
  await style(true);const inline=await snapshot('inline display uses a smaller designed operator glyph');
  assert.equal(inline.document.source,beforeStyle.document.source);assert.equal(inline.document.revision,beforeStyle.document.revision);
  assert.equal(inline.undoDepth,beforeStyle.undoDepth);assert.notEqual(opGlyph(await scene(),opId),displayGlyph);
  await style(false);assert.equal(opGlyph(await scene(),opId),displayGlyph);
  await caretClick(lower,0);const priorSource=(await inspect()).document.source;const priorSvg=await svg();
  await click('[data-limits="noLimits"]');const forced=await snapshot('explicit side-bound directive is a source transaction');
  assert.ok(forced.document.source.includes(String.raw`\nolimits`));assert.equal(operator(forced).id,opId);
  assert.equal(bound(forced,'subscript').id,lower);assert.equal(bound(forced,'superscript').id,upper);assert.notEqual(await svg(),priorSvg);
  await click('#focus-editor');await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,priorSource);assert.equal(await svg(),priorSvg);
  await click('#focus-editor');await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,forced.document.source);
  await click('[data-limits="auto"]');assert.equal((await snapshot('automatic layout restored without losing bounds')).document.source,priorSource);

  await source('');await click('[data-operator="int"]');await type('0');await key('ArrowRight','ArrowRight',39);await type('1');
  await key('ArrowRight','ArrowRight',39);await type('x');const integral=await snapshot('integral entered with side bounds and following operand');
  assert.equal(integral.document.source,String.raw`\int_{0}^{1}x`);
  await caretClick(bound(integral,'subscript').id,0);await click('[data-limits="limits"]');
  assert.ok((await snapshot('integral explicit stacked limits')).document.source.includes(String.raw`\limits`));
  await source('');await click('[data-operator="lim"]');await type('x=0');await key('ArrowRight','ArrowRight',39);await type('x');
  assert.equal((await snapshot('upright named operator with one editable limit')).document.source,String.raw`\lim_{x=0}x`);

  await source('');await click('[data-general-operator="bounds"]');await type('a');await key('ArrowRight','ArrowRight',39);
  await type('b');await key('ArrowRight','ArrowRight',39);await type('f');
  const generic=await snapshot('general operator: lower, upper and editable body from blank');
  assert.equal(generic.document.source,String.raw`\mathop{f}_{a}^{b}`);
  const general=operator(generic);assert.equal(general.kind,'generalOperator');
  assert.ok((await evaluate('document.querySelector(".vt-core-position").textContent')).includes('算子本体'));
  await caretClick(general.body.id,1);await type('g');assert.equal((await inspect()).document.source,String.raw`\mathop{fg}_{a}^{b}`);
  await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,generic.document.source);
  await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,String.raw`\mathop{fg}_{a}^{b}`);
  await caretClick(bound(generic,'subscript').id,0);await key('ArrowUp','ArrowUp',38);
  assert.equal((await inspect()).selection.focus.nodeId,general.body.id);
  await key('ArrowUp','ArrowUp',38);assert.equal((await inspect()).selection.focus.nodeId,bound(generic,'superscript').id);
  await key('ArrowDown','ArrowDown',40);assert.equal((await inspect()).selection.focus.nodeId,general.body.id);
  const generalBefore=await inspect();await click('[data-limits="noLimits"]');const generalSide=await snapshot('general operator supports the shared placement controls');
  assert.ok(generalSide.document.source.includes(String.raw`\nolimits`));assert.equal(operator(generalSide).body.id,general.body.id);
  await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,generalBefore.document.source);
  await source('f(x)');await click('#focus-editor');await key('a','KeyA',65,2);
  const selectedBody=(await inspect()).document.analysis.tree.entries.map(e=>e.node.id);
  await click('[data-general-operator="body"]');const wrapped=await snapshot('selection becomes operator body, not the following operand');
  assert.equal(wrapped.document.source,String.raw`\mathop{f(x)}`);assert.deepEqual(operator(wrapped).body.entries.map(e=>e.node.id),selectedBody);
  await source(String.raw`\mathop{\mathrm{rank}}_{A}+\mathop{\sum\sum}\limits_{i,j}^{n}+\mathop{\binom{n}{k}}`);
  await svg();const generalShot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
  writeFileSync(resolve(out,'general-operators.png'),Buffer.from(generalShot.data,'base64'));

  const presentation=String.raw`\sum_{i=1}^{n}x_i+\int_{0}^{1}\frac{x}{1+x^2}dx+\lim_{x\to0}\frac{x}{1+x}`;
  await source(presentation);const shown=await snapshot('sum integral and upright limit coexist in one formula');assert.equal(shown.document.analysis.exportable,true);
  const vector=await svg();for(const bad of ['<text','<image','foreignObject','NaN','Infinity'])assert.ok(!vector.includes(bad));
  assert.deepEqual(await evaluate('window.__engineSandbox.diagnostics()'),[],'Unexpected engine errors before the deliberate draft case');
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'operators-display.png'),Buffer.from(screenshot.data,'base64'));
  const priorRevision=shown.document.revision;await style(true);assert.equal((await inspect()).document.revision,priorRevision);assert.equal((await inspect()).document.source,presentation);
  assert.ok(!(await evaluate('document.getElementById("export-result").textContent')).includes('<svg'),'A changed mathematical style must invalidate the previous vector preview');
  const inlineVector=await svg(),inlineLayout=await scene();assert.notEqual(inlineVector,vector);
  for(const dimension of ['width','height']){
    const value=Number(new RegExp(`\\b${dimension}="([^"]+)"`).exec(inlineVector)?.[1]);
    assert.ok(Math.abs(value-inlineLayout[dimension])<1e-5,`${dimension} does not match the current inline scene`);
  }
  const inlineScreenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'operators-inline.png'),Buffer.from(inlineScreenshot.data,'base64'));
  await style(false);
  await source(String.raw`\sum_{i=1}^{`);assert.equal((await snapshot('incomplete upper bound is retained as a draft')).document.analysis.complete,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  await key('z','KeyZ',90,2);assert.equal((await snapshot('undo recovers the unchanged valid formula')).document.source,presentation);
  const events=await evaluate('window.__engineSandbox.events()');
  return {status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
   actualSandboxControls:true,browserKeyAndPointerEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,steps,
   screenshots:['target/operator-browser/operators-display.png','target/operator-browser/operators-inline.png'],
   notCovered:['Physical Windows IME in this runner','Word or PowerPoint UI','installed application','production UI switch']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

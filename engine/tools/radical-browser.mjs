// P07 root closure through the actual isolated editor buttons and browser input.
// This file does not claim physical OS IME or Office acceptance.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets,sha256} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/radical-browser');mkdirSync(out,{recursive:true});
const steps=[];
const result=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()');
 async function click(selector){const p=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);await point(p);}
 async function point(p){await send('Input.dispatchMouseEvent',{type:'mousePressed',...p,button:'left',clickCount:1});await send('Input.dispatchMouseEvent',{type:'mouseReleased',...p,button:'left',clickCount:1});}
 async function key(key,code,n,modifiers=0){for(const type of ['keyDown','keyUp'])await send('Input.dispatchKeyEvent',{type,key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});}
 const type=text=>send('Input.insertText',{text});
 async function source(value){await click('#source');await key('a','KeyA',65,2);await type(value);}
 async function check(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth});return s;}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');assert.equal((await inspect()).document.source,'');
  await click('[data-structure="indexedRoot"]');await type('3');await key('ArrowRight','ArrowRight',39);
  await click('[data-structure="squareRoot"]');await click('[data-structure="fraction"]');await type('x+1');
  await key('ArrowDown','ArrowDown',40);await type('y');await click('[data-structure="superscript"]');await type('2');
  const built=await check('nested indexed/square roots built from blank by UI');
  assert.equal(built.document.canonicalLatex,String.raw`\sqrt[3]{\sqrt{\frac{x+1}{y^{2}}}}`);assert.equal(built.document.analysis.facts.radicals,2);
  const scene=await evaluate('window.__engineSandbox.layout()');
  const degree=built.document.analysis.tree.entries[0].node.degree.id;
  const c=scene.carets.find(c=>c.position.nodeId===degree&&c.position.kind==='row'&&c.position.offset===1);assert.ok(c);
  // Only the SVG viewport transform is used; no DOM formula boxes determine math placement.
  const p=await evaluate(`(()=>{const svg=document.querySelector('#editor svg');svg.scrollIntoView({block:'center',inline:'nearest'});const m=svg.getScreenCTM();const p=new DOMPoint(${c.x},${(c.top+c.bottom)/2}).matrixTransform(m);return {x:p.x,y:p.y};})()`);
  const hit=await evaluate(`(()=>{const p=${JSON.stringify(p)};const e=document.elementFromPoint(p.x,p.y);return {point:p,viewport:{width:innerWidth,height:innerHeight},scrollY,element:e?.tagName??null,insideEditor:Boolean(e?.closest('#editor'))};})()`);
  steps.push({label:'root degree pointer target visibility',...hit});console.log(JSON.stringify({rootDegreeTarget:hit}));
  assert.ok(hit.insideEditor&&p.x>=0&&p.x<hit.viewport.width&&p.y>=0&&p.y<hit.viewport.height,'Root-degree click must hit the visible editor before input is sent');
  await point(p);await type('+1');const changed=await check('pointer edits the root degree');
  assert.equal(changed.document.canonicalLatex,String.raw`\sqrt[3+1]{\sqrt{\frac{x+1}{y^{2}}}}`);
  const label=await evaluate('document.querySelector(".vt-core-input").getAttribute("aria-label")');assert.ok(label.includes('根指数'),label);
  await key('z','KeyZ',90,2);assert.equal((await check('undo degree edit')).document.source,built.document.source);
  await key('y','KeyY',89,2);assert.equal((await inspect()).document.source,changed.document.source);
  const stable=await inspect();await click('#export-svg');
  const svg=await evaluate('document.getElementById("export-result").textContent');assert.ok(svg.includes('<svg')&&!svg.includes('<text')&&!svg.includes('<image'));
  assert.equal((await inspect()).document.source,stable.document.source);
  await source(String.raw`\sqrt[n`);const draft=await check('incomplete root index remains editable source');assert.equal(draft.document.analysis.complete,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
  await key('z','KeyZ',90,2);assert.equal((await check('undo invalid root draft')).document.source,stable.document.source);
  const tall=String.raw`\sqrt{\dfrac{1}{\dfrac{2}{\dfrac{3}{\dfrac{4}{5}}}}}`;
  await source(tall);const tallState=await check('tall root assembled around nested fractions');
  assert.equal(tallState.document.analysis.exportable,true);
  const tallScene=await evaluate('window.__engineSandbox.layout()');const radicalId=tallState.document.analysis.tree.entries[0].node.id;
  const surdPieces=tallScene.glyphs.filter(g=>g.nodeId===radicalId);
  assert.ok(surdPieces.length>1,'Tall root did not use separate font assembly pieces');
  assert.ok(surdPieces.every(g=>g.scale===1),'Radical was anisotropically enlarged');
  await click('#export-svg');const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
  writeFileSync(resolve(out,'tall-root.png'),Buffer.from(shot.data,'base64'));
  return {status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
   actualSandboxControls:true,startedBlank:true,steps,rootAssemblyPieces:surdPieces.length,accessibleDegree:label,
   nestedSvgSha256:sha256(Buffer.from(svg)),screenshot:'target/radical-browser/tall-root.png',
   notCovered:['actual Windows IME in this root test','Office UI','main application default switch']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});
writeFileSync(resolve(out,'report.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result,null,2));

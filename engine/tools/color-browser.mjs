// P07 colors: actual sandbox controls followed by explicit isolated View faults.
import assert from 'node:assert/strict';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,html,manifest,root}=createSandboxAssets();const out=resolve(root,'target/color-browser');mkdirSync(out,{recursive:true});
const steps=[];
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=()=>evaluate('window.__engineSandbox.inspect()'),scene=()=>evaluate('window.__engineSandbox.layout()');
 async function key(key,code,n,modifiers=0,text){
  await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
  await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
 }
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e||e.disabled)throw new Error('Unavailable color control');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();const x=r.x+r.width/2,y=r.y+r.height/2;if(!e.contains(document.elementFromPoint(x,y)))throw new Error('Control occluded');return{x,y};})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function type(text){for(const ch of text){if(/[a-z]/u.test(ch))await key(ch,`Key${ch.toUpperCase()}`,ch.toUpperCase().charCodeAt(0),0,ch);else if(ch==='+')await key('+','Equal',187,8,'+');else await send('Input.insertText',{text:ch});}}
 async function source(text){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
 async function all(){await click('#focus-editor');await key('a','KeyA',65,2);const s=await inspect();assert.equal(s.normalizedSelection.end,s.document.analysis.tree.entries.length);}
 async function snapshot(label){const s=await inspect();assert.equal(await evaluate('document.getElementById("source").value'),s.document.source);steps.push({label,source:s.document.source,revision:s.document.revision,undoDepth:s.undoDepth,selection:s.selection});return s;}
 async function svg(){await click('#export-svg');return (await evaluate('document.getElementById("export-result").textContent')).split(/(?=<svg )/u).at(-1);}
 const expectedFill=c=>c?'#'+[c.r,c.g,c.b].map(x=>x.toString(16).padStart(2,'0')).join(''):null;
 async function checkPaint(){const s=await scene();const paints=await evaluate(`(()=>{const g=document.querySelector('#editor svg .vt-core-glyphs');return{glyphs:[...g.querySelectorAll('use')].map(e=>e.getAttribute('fill')),rules:[...g.querySelectorAll('rect')].map(e=>e.getAttribute('fill'))};})()`);
  assert.deepEqual(paints.glyphs,s.glyphs.map(g=>expectedFill(g.color)));assert.deepEqual(paints.rules,s.rules.map(r=>expectedFill(r.color)));}
 try{
  await evaluate('window.__engineSandbox.ready.then(()=>true)');await click('#focus-editor');await type('x+y');const plain=await scene();
  await all();await click('[data-color="red"]');const red=await snapshot('real selection colored using toolbar');assert.equal(red.document.source,String.raw`\textcolor[HTML]{FF0000}{x+y}`);
  const redScene=await scene();for(const d of ['width','height','baseline'])assert.equal(redScene[d],plain[d]);await checkPaint();const redSvg=await svg();assert.ok(redSvg.includes('fill="#ff0000"'));
  await click('#focus-editor');await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,'x+y');assert.ok((await scene()).glyphs.every(g=>g.color===undefined));
  await key('y','KeyY',89,2);assert.equal(await svg(),redSvg);await snapshot('undo redo recover the same source and colors');
  await click('#focus-editor');await key('ArrowRight','ArrowRight',39);await click('[data-color="blue"]');await type('z');
  const mixed=await snapshot('new empty color body filled after leaving the first color');assert.ok(mixed.document.source.endsWith(String.raw`\textcolor[HTML]{0000FF}{z}`));await checkPaint();

  const scopes=String.raw`a\color{red}b{\color{blue}c}d\textcolor{green}{e}f`;
  await source(scopes);await snapshot('declaration scope and nested restoration');await checkPaint();
  assert.deepEqual((await scene()).glyphs.map(g=>expectedFill(g.color)),[null,'#ff0000','#0000ff','#ff0000','#00ff00','#ff0000']);
  const before=await scene();await source(scopes.replace('{red}','{blue}'));assert.equal((await scene()).width,before.width);assert.equal((await scene()).glyphs[1].color.b,255);

  const presentation=String.raw`\textcolor[HTML]{173B62}{\boxed{\int_0^1 x\,\mathrm{d}x}}+\textcolor{red!70!black}{\sqrt{\frac{a}{b}}}+\phantom{\textcolor{blue}{z}}`;
  await source(presentation);const shown=await snapshot('colored borders roots and hidden content coexist');await checkPaint();const vector=await svg();
  assert.ok(!vector.includes('#0000ff'));assert.ok(!vector.includes('style=')&&!vector.includes('url('));assert.ok(vector.includes('fill="#173b62"'));
  const colorList=(await scene()).glyphs.map(g=>expectedFill(g.color));
  await click('#zoom');await key('End','End',35);await key('Enter','Enter',13);assert.equal((await inspect()).document.revision,shown.document.revision);assert.equal(await svg(),vector);
  assert.deepEqual((await scene()).glyphs.map(g=>expectedFill(g.color)),colorList);
  await click('#zoom');await key('Home','Home',36);await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
  const clip=await evaluate(`(()=>{const e=document.getElementById('editor-scroll');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:Math.max(0,r.x+scrollX),y:Math.max(0,r.y+scrollY),width:Math.min(r.width,1100),height:Math.min(r.height,360),scale:1};})()`);
  const image=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true,clip});writeFileSync(resolve(out,'colors.png'),Buffer.from(image.data,'base64'));
  assert.deepEqual(await evaluate('window.__engineSandbox.diagnostics()'),[],'Unexpected ordinary editor errors');

  // Deliberate malformed-scene tests are synthetic component calls. They do not
  // masquerade as user keyboard actions, and they never mutate the main editor.
  const failures=await evaluate(`(async()=>{const{SvgSceneView}=await import('/host/svg-view.js');const{rgbHex}=await import('/host/paint.js');
    const parent=document.createElement('div');document.body.append(parent);const view=new SvgSceneView(parent);const good=window.__engineSandbox.layout();view.setSnapshot(good);const original=view.svg.outerHTML;let rejected=0;
    const bad=[null,'url(https://example.invalid)',{r:256,g:0,b:0},{r:-1,g:0,b:0},{r:0.5,g:0,b:0},{r:NaN,g:0,b:0},{r:0,g:Infinity,b:0},{r:0,g:0,b:0,style:'fill:red'}];
    try{for(const value of bad){for(const group of ['glyphs','rules']){const next=structuredClone(good);next[group][0].color=value;let error=false;
      try{view.setSnapshot(next);}catch{error=true;rejected++;}if(!error||view.svg.outerHTML!==original)throw new Error('Malformed paint replaced a valid SVG');}}
      if(rgbHex({r:0,g:128,b:255})!=='#0080ff')throw new Error('RGB conversion failed');return rejected;
    }finally{view.dispose();parent.remove();}})()`);assert.equal(failures,16);

  await source(String.raw`\textcolor{url(https://example.invalid)}{x}`);const invalid=await snapshot('untrusted color source retained without publishing');assert.equal(invalid.document.analysis.supported,false);
  assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,presentation);
  await source(String.raw`\textcolor[HTML`);assert.equal((await snapshot('unfinished color model remains a draft')).document.analysis.complete,false);
  await key('z','KeyZ',90,2);assert.equal((await inspect()).document.source,presentation);await checkPaint();
  const events=await evaluate('window.__engineSandbox.events()');return{status:'PASS',browser:browserVersion.product,hostBuildId:manifest.buildId,
   wasmSha256:manifest.assets[0].sha256,actualSandboxControls:true,browserKeyAndPointerEvents:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,
   syntheticPaintFaults:failures,steps,screenshot:'target/color-browser/colors.png',notCovered:['Physical Windows IME','Office Ribbon','installed product','production UI switch']};
 }catch(error){let state;try{state=await inspect();}catch{}writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));throw error;}
});writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report,null,2));

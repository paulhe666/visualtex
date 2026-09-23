// Actual owned Edge keyboard, pointer and browser clipboard acceptance.
// The test creates its own hosts/profile. No Office or existing browser attach.
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve,relative} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/grid-rectangle-browser-review',new Date().toISOString().replaceAll(':','-').replaceAll('.','-'));
mkdirSync(out,{recursive:true});const sha=b=>createHash('sha256').update(b).digest('hex');
const initial=String.raw`\begin{matrix}a&b&c\\d&\frac{x}{y}&f\\g&h&i\end{matrix}`;
const bootstrap=String.raw`
import init,{EngineRuntime} from '/assets/core.js';
import {VerifiedAssetLoader} from '/host/resources.js';
import {EngineEditorHost} from '/host/editor-host.js';
const hosts={},errors=[],clipboard=[],exports={};let runtime,loader,enters=0;
const ready=(async()=>{
 const manifest=await(await fetch('/manifest.json')).json();
 loader=new VerifiedAssetLoader(manifest,{baseUrl:new URL('/assets/',location.href).href,expectedBuildId:manifest.buildId,expectedEngineProtocol:1});
 await init({module_or_path:(await loader.load('core')).bytes});runtime=new EngineRuntime();const fonts={};
 for(const a of manifest.assets.filter(a=>a.kind==='font')){const r=JSON.parse(runtime.registerFont((await loader.load(a.id)).bytes,0));if(r.status!=='ok')throw Error(JSON.stringify(r));fonts[a.id]=r.result.font.id;}
 for(const id of ['A','B']){
  const source=document.getElementById('source-'+id);
  const host=new EngineEditorHost(runtime,document.getElementById('editor-'+id),{documentId:'rectangle-'+id,
   initialSource:source.value,layout:{mathFont:fonts.math,textFont:fonts.text,style:'display',fontSizePx:24,paddingEm:.125},
   onChange:s=>{source.value=s.document.source;},onError:e=>errors.push(String(e)),onEnter:()=>enters++});hosts[id]=host;
  source.addEventListener('input',()=>{try{host.replaceSource(source.value);}catch(e){errors.push(String(e));}});
  for(const type of ['copy','paste','cut'])host.input.textarea.addEventListener(type,e=>clipboard.push({type,trusted:e.isTrusted,text:e.clipboardData?.getData('text/plain'),latex:e.clipboardData?.getData('application/x-latex'),custom:e.clipboardData?.getData('application/vnd.visualtex.engine-fragment+json')}));
  for(const b of document.querySelectorAll('[data-frame="'+id+'"]')){
   b.addEventListener('pointerdown',e=>e.preventDefault());b.addEventListener('click',()=>{try{
    if(b.dataset.action==='focus')host.focus();if(b.dataset.action==='undo'){host.history(false);host.focus();}
    if(b.dataset.action==='redo'){host.history(true);host.focus();}if(b.dataset.action==='svg')exports[id]=host.exportSvg();
   }catch(e){errors.push(String(e));}});
  }
  document.getElementById('zoom-'+id).addEventListener('change',e=>host.setZoom(Number(e.target.value)));
 }
 hosts.A.focus();
})();
window.__rectangle={ready,inspect:id=>hosts[id].inspect(),scene:id=>hosts[id].layoutSnapshot,
 stats:()=>({sessions:runtime?.activeSessions,enters,errors:errors.slice(),clipboard:clipboard.slice()}),
 events:()=>hosts.A.input.events,svg:()=>exports.A};
window.addEventListener('pagehide',()=>{Object.values(hosts).forEach(h=>h.dispose());runtime?.free();loader?.dispose();},{once:true});
ready.catch(e=>document.getElementById('error').textContent=String(e));
`;
assets.set('/review.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const esc=s=>s.replaceAll('&','&amp;').replaceAll('<','&lt;');
const section=(id,source)=>`<section><h2>Independent formula ${id}</h2>${['focus','undo','redo','svg'].map(a=>`<button id="${a}-${id}" data-frame="${id}" data-action="${a}">${a}</button>`).join('')}<select id="zoom-${id}"><option value=".5">50%</option><option value="1" selected>100%</option><option value="2">200%</option></select><div class="editor" id="editor-${id}"></div><textarea id="source-${id}">${esc(source)}</textarea></section>`;
const html=`<!doctype html><meta charset="utf-8"><title>Identified cell rectangle acceptance</title><style>body{font:16px Arial;margin:24px}section{border:1px solid #888;padding:16px;margin:18px 0}.editor{position:relative;min-height:145px;padding:12px 0;overflow:auto}textarea{width:96%;height:65px}button,select{padding:7px;margin-right:8px}</style><h1>Cell selection and clipboard — one formula owner</h1><p id="error"></p>${section('A',initial)}${section('B','z_2')}<script type="module" src="/review.js"></script>`;
const steps=[];let report;
try{report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const inspect=id=>evaluate(`window.__rectangle.inspect(${JSON.stringify(id)})`);const stats=()=>evaluate('window.__rectangle.stats()');
 const grid=s=>s.document.analysis.tree.entries[0].node;const cell=(s,r,c)=>grid(s).rows[r].cells[c];
 const content=s=>({source:s.document.source,tree:s.document.analysis.tree,selection:s.selection});
 async function key(key,code,n,modifiers=0,text){for(const type of ['keyDown','keyUp'])await send('Input.dispatchKeyEvent',{type,key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(type==='keyDown'&&text?{text}:{})});}
 async function click(selector){const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:r.x+r.width/2,y:r.y+r.height/2}})()`);
  for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});}
 async function point(cellId,offset=0){return evaluate(`(()=>{const s=window.__rectangle.scene('A');const c=s.carets.find(c=>c.position.kind==='row'&&c.position.nodeId===${cellId}&&c.position.offset===${offset});if(!c)throw Error('Missing identified cell caret');const e=document.querySelector('#editor-A'),svg=e.querySelector('svg');svg.scrollIntoView({block:'center'});const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(svg.getScreenCTM());if(!e.contains(document.elementFromPoint(p.x,p.y)))throw Error('Cell target outside viewport');return{x:p.x,y:p.y}})()`);}
 async function pointer(cellId,offset=0){const p=await point(cellId,offset);for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...p,button:'left',clickCount:1});}
 async function drag(a,b){const p=await point(a);await send('Input.dispatchMouseEvent',{type:'mousePressed',...p,button:'left',clickCount:1});const q=await point(b);
  await send('Input.dispatchMouseEvent',{type:'mouseMoved',...q,button:'left',buttons:1});await send('Input.dispatchMouseEvent',{type:'mouseReleased',...q,button:'left',clickCount:1});}
 async function check(label){const a=await inspect('A');assert.deepEqual(await inspect('B'),beforeB);assert.equal((await stats()).sessions,2);
  assert.equal(await evaluate('document.querySelectorAll("#editor-A .vt-core-input").length'),1);assert.equal(await evaluate('document.getElementById("source-A").value'),a.document.source);
  steps.push({label,source:a.document.source,revision:a.document.revision,undoDepth:a.undoDepth,selection:a.selection});return a;}
 async function overlays(n){assert.equal(await evaluate('document.querySelectorAll("#editor-A .vt-core-selection").length'),n);}
 await evaluate('window.__rectangle.ready.then(()=>true)');const beforeB=await inspect('B');assert.deepEqual((await stats()).errors,[]);
 let a=await inspect('A');await drag(cell(a,0,0).id,cell(a,1,1).id);a=await check('Trusted pointer drag creates a four-cell identified rectangle');
 assert.equal(a.selection.kind,'gridCells');await overlays(4);const copiedBefore=await inspect('A');await key('c','KeyC',67,2);
 assert.deepEqual(await inspect('A'),copiedBefore);let clipboard=(await stats()).clipboard;
 assert.equal(clipboard.at(-1).type,'copy');assert.ok(clipboard.at(-1).trusted);assert.ok(clipboard.at(-1).text.startsWith(String.raw`\begin{matrix}`));
 const payload=JSON.parse(clipboard.at(-1).custom);assert.deepEqual(Object.keys(payload).sort(),['latex','schemaVersion','structure']);assert.equal(payload.schemaVersion,2);assert.equal(clipboard.at(-1).latex,clipboard.at(-1).text);assert.equal(payload.structure.schema,'visualtex-engine-fragment');assert.equal(payload.structure.version,1);assert.equal(payload.structure.assets.length,0);const structureJson=JSON.stringify(payload.structure);for(const forbidden of ['"id":','"origin":','"nodeId":','"documentId":','"handle":','"sourceId":'])assert.ok(!structureJson.includes(forbidden),forbidden);
 await check('Trusted browser copy emits versioned identity-free structure plus LaTeX/plain fallbacks');
 await drag(cell(a,1,1).id,cell(a,2,2).id);const beforePaste=await inspect('A');await key('v','KeyV',86,2);
 a=await check('Trusted browser paste replaces the selected 2x2 cell contents');assert.equal(a.document.revision,beforePaste.document.revision+1);
 assert.equal(grid(a).id,grid(beforePaste).id);assert.deepEqual(grid(a).columns,grid(beforePaste).columns);
 assert.deepEqual(grid(a).rows.map(r=>r.cells.map(c=>c.id)),grid(beforePaste).rows.map(r=>r.cells.map(c=>c.id)));
 assert.equal(cell(a,1,1).entries[0].node.value,'a');assert.notEqual(cell(a,1,1).entries[0].node.id,cell(a,0,0).entries[0].node.id);
 assert.equal(cell(a,2,2).entries[0].node.kind,'fraction');assert.equal(a.document.analysis.facts.grids,1);await overlays(4);
 const pasted=content(a);await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(beforePaste));
 await key('y','KeyY',89,2);assert.deepEqual(content(await inspect('A')),pasted);await check('Paste undo/redo restores exact source, selection and internal identities');
 await key('x','KeyX',88,2);a=await check('Trusted browser cut clears contents without removing cells or rows');
 for(let r=1;r<3;r++)for(let c=1;c<3;c++)assert.equal(cell(a,r,c).entries.length,0);assert.equal(cell(a,0,0).entries[0].node.value,'a');
 await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),pasted);
 // Keyboard rectangle extension begins at a cell caret and is content-read-only.
 a=await inspect('A');await pointer(cell(a,0,0).id);const keyboardBefore=await inspect('A');
 await key('ArrowRight','ArrowRight',39,10);await key('ArrowDown','ArrowDown',40,10);a=await check('Ctrl+Shift+arrows extend one same-grid cell rectangle');
 assert.equal(a.selection.kind,'gridCells');assert.equal(a.selection.anchorCell,cell(a,0,0).id);assert.equal(a.selection.focusCell,cell(a,1,1).id);
 assert.equal(a.document.revision,keyboardBefore.document.revision);assert.equal(a.undoDepth,keyboardBefore.undoDepth);await overlays(4);
 await key('ArrowLeft','ArrowLeft',37);assert.equal((await inspect('A')).selection.focus.nodeId,cell(a,0,0).id);
 // A 1x1 rectangle must not accept a 2x2 payload by inventing new columns.
 await key('ArrowLeft','ArrowLeft',37,10);const mismatch=await inspect('A');assert.equal(mismatch.selection.kind,'gridCells');await overlays(1);
 await key('v','KeyV',86,2);assert.deepEqual(await inspect('A'),mismatch);let errors=(await stats()).errors;
 assert.equal(errors.length,1);assert.match(errors[0],/dimensions must exactly match/);await key('Escape','Escape',27);
 await key('q','KeyQ',81,0,'q');assert.ok((await inspect('A')).document.source.includes('qa&b'));await key('z','KeyZ',90,2);
 await check('Wrong-size paste fails atomically; Escape returns to normal cell input');
 const zoomCases=[];const natural=await evaluate('window.__rectangle.scene("A")');const stable=(await inspect('A')).document;
 for(const dpr of [1,1.25,2]){await send('Emulation.setDeviceMetricsOverride',{width:1100,height:760,deviceScaleFactor:dpr,mobile:false});
  for(const [i,zoom] of [.5,1,2].entries()){
   await click('#zoom-A');await key('Home','Home',36);for(let j=0;j<i;j++)await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
   a=await inspect('A');await drag(cell(a,1,1).id,cell(a,0,0).id);a=await inspect('A');assert.equal(a.selection.kind,'gridCells');await overlays(4);
   assert.equal(a.document.source,stable.source);assert.equal(a.document.revision,stable.revision);const scene=await evaluate('window.__rectangle.scene("A")');
   assert.equal(scene.width,natural.width);assert.equal(scene.height,natural.height);assert.deepEqual(scene.gridColumns,natural.gridColumns);
   assert.deepEqual(await inspect('B'),beforeB);zoomCases.push({dpr,zoom,selection:a.selection});
  }
 }
 await send('Emulation.clearDeviceMetricsOverride');await check('Nine zoom/DPR combinations preserve rectangular hit geometry and frame B');
 await click('#svg-A');const svg=await evaluate('window.__rectangle.svg()');assert.equal(svg.match(/<svg\b/g)?.length,1);assert.ok(!svg.includes('vt-core-selection'));assert.ok(!svg.includes('<text'));
 await check('Selection overlays are absent from the single-object vector export');
 clipboard=(await stats()).clipboard;assert.ok(clipboard.some(e=>e.type==='copy'&&e.trusted));assert.ok(clipboard.some(e=>e.type==='paste'&&e.trusted));assert.ok(clipboard.some(e=>e.type==='cut'&&e.trusted));
 const events=await evaluate('window.__rectangle.events()');const image=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'cell-rectangle-two-frames.png'),Buffer.from(image.data,'base64'));
 return {status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,bootstrapSha256:sha(bootstrap),
  steps,zoomCases,actualBrowserKeysAndPointer:true,actualBrowserClipboard:true,structuralClipboardV2:true,clipboardEvents:clipboard.map(({type,trusted,text,latex,custom})=>({type,trusted,text,latex,customSchemaVersion:custom?JSON.parse(custom).schemaVersion:null})),
  events:events.length,trustedEvents:events.filter(e=>e.trusted).length,frameBUnchanged:true,expectedDimensionRefusals:1,physicalWindowsIme:false,
  screenshot:relative(root,resolve(out,'cell-rectangle-two-frames.png')).replaceAll('\\','/'),notCovered:['Physical Windows IME','Office or installed application','Arbitrary external spreadsheet clipboard formats']};
});}catch(error){report={status:'FAILED',error:String(error),steps,hostBuild:manifest.buildId};throw error;}
finally{if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));}}

// P08 Matrix: two actual browser hosts. Setup/inspection is scripted; all
// editing goes through real controls, CDP trusted keys and pointer events.
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {mkdirSync,writeFileSync} from 'node:fs';
import {resolve,relative} from 'node:path';
import {createSandboxAssets} from './sandbox-assets.mjs';
import {withOwnedBrowser} from './owned-browser-review.mjs';
const {assets,manifest,root}=createSandboxAssets();
const out=resolve(root,'target/matrix-browser-review',new Date().toISOString().replaceAll(':','-').replaceAll('.','-'));mkdirSync(out,{recursive:true});
const sha=bytes=>createHash('sha256').update(bytes).digest('hex');
const bootstrap=String.raw`
import init,{EngineRuntime} from '/assets/core.js';
import {EngineEditorHost} from '/host/editor-host.js';
import {VerifiedAssetLoader} from '/host/resources.js';
const hosts={},errors=[],enters={A:0,B:0},exports={};let runtime,loader;
const ready=(async()=>{
 const manifest=await(await fetch('/manifest.json',{cache:'no-store'})).json();
 loader=new VerifiedAssetLoader(manifest,{baseUrl:new URL('/assets/',location.href).href,expectedBuildId:manifest.buildId,expectedEngineProtocol:1});
 const core=await loader.load('core');await init({module_or_path:core.bytes});runtime=new EngineRuntime();
 const fonts={};for(const asset of manifest.assets.filter(a=>a.kind==='font')){
  const data=await loader.load(asset.id);const r=JSON.parse(runtime.registerFont(data.bytes,0));
  if(r.status!=='ok'||r.result.font.sha256!==asset.sha256)throw new Error('Font identity mismatch');fonts[asset.id]=r.result.font.id;
 }
 for(const id of ['A','B']){
  const source=document.getElementById('source-'+id);
  const host=new EngineEditorHost(runtime,document.getElementById('editor-'+id),{
   documentId:'matrix-browser-'+id,initialSource:id==='A'?'':'z_2',
   layout:{mathFont:fonts.math,textFont:fonts.text,style:'display',fontSizePx:25,paddingEm:.125},
   onChange:s=>{source.value=s.document.source;},onError:e=>errors.push({frame:id,message:String(e)}),onEnter:()=>enters[id]++,
  });hosts[id]=host;source.value=host.inspect().document.source;
  source.addEventListener('input',()=>{try{host.replaceSource(source.value);}catch(e){errors.push({frame:id,message:String(e)});}});
  for(const button of document.querySelectorAll('[data-frame="'+id+'"]')){
   button.addEventListener('pointerdown',e=>{if(e.button===0)e.preventDefault();});
   button.addEventListener('click',()=>{try{
    switch(button.dataset.action){
     case 'focus':host.focus();break;
     case 'matrix':host.applyMatrix(2,2);break;
     case 'cases':host.applyCases(2);break;
     case 'aligned':host.applyAligned(1,1);break;
     case 'rowBefore':host.insertGridRow(false);break;
     case 'rowAfter':host.insertGridRow(true);break;
     case 'rowDelete':host.deleteGridRow();break;
     case 'columnBefore':host.insertGridColumn(false);break;
     case 'columnAfter':host.insertGridColumn(true);break;
     case 'columnDelete':host.deleteGridColumn();break;
     case 'undo':host.history(false);host.focus();break;
     case 'redo':host.history(true);host.focus();break;
     case 'export':exports[id]=host.exportSvg();break;
    }
   }catch(e){errors.push({frame:id,message:String(e)});}});
  }
  document.getElementById('zoom-'+id).addEventListener('change',e=>{try{host.setZoom(Number(e.target.value));}catch(e){errors.push({frame:id,message:String(e)});}});
 }
 hosts.A.focus();
})();
window.__matrixReview={ready,inspect:id=>hosts[id]?.inspect(),scene:id=>hosts[id]?.layoutSnapshot,
 stats:()=>({sessions:runtime?.activeSessions,enters:{...enters},errors:errors.slice()}),events:id=>hosts[id]?.input.events??[],exported:id=>exports[id]??null};
window.addEventListener('pagehide',()=>{for(const h of Object.values(hosts))h.dispose();runtime?.free();loader?.dispose();},{once:true});
ready.catch(e=>document.getElementById('error').textContent=String(e));
`;
assets.set('/review.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const controls=[['focus','Focus'],['matrix','2×2 Matrix'],['cases','Cases'],['aligned','Aligned'],['rowBefore','Row before'],['rowAfter','Row after'],['rowDelete','Delete row'],['columnBefore','Column before'],['columnAfter','Column after'],['columnDelete','Delete column'],['undo','Undo'],['redo','Redo'],['export','SVG']];
const section=id=>`<section><h2>Independent formula ${id}</h2><div>${controls.map(([action,label])=>`<button data-frame="${id}" data-action="${action}" id="${action}-${id}">${label}</button>`).join('')}
<select id="zoom-${id}"><option value=".5">50%</option><option value="1" selected>100%</option><option value="2">200%</option></select></div><div id="editor-${id}" class="editor"></div><textarea id="source-${id}" aria-label="Formula ${id} source"></textarea></section>`;
const html=`<!doctype html><html><head><meta charset="utf-8"><title>Matrix contained-grid acceptance</title>
<style>body{font:16px Arial;margin:24px}section{border:1px solid #888;padding:12px;margin:12px 0}.editor{position:relative;min-height:100px;padding:12px 0;overflow:auto}textarea{width:96%;height:54px}button,select{margin:3px;padding:7px}</style></head><body><h1>One session per mathematical matrix</h1><p id="error"></p>${section('A')}${section('B')}<script type="module" src="/review.js"></script></body></html>`;
const steps=[];let report;
try{
 report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
  const inspect=id=>evaluate(`window.__matrixReview.inspect(${JSON.stringify(id)})`);
  const stats=()=>evaluate('window.__matrixReview.stats()');
  async function key(key,code,n,modifiers=0,text){
   await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers,...(text?{text}:{})});
   await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:n,nativeVirtualKeyCode:n,modifiers});
  }
  async function click(selector){
   const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)throw new Error('Missing control');e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();const x=r.x+r.width/2,y=r.y+r.height/2;if(!e.contains(document.elementFromPoint(x,y)))throw new Error('Control not in viewport');return{x,y};})()`);
   for(const type of ['mousePressed','mouseReleased'])await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});
  }
  async function type(text){for(const ch of text){
   if(/^[a-z]$/u.test(ch))await key(ch,'Key'+ch.toUpperCase(),ch.toUpperCase().charCodeAt(0),0,ch);
   else if(ch==='=')await key('=','Equal',187,0,'=');
   else if(ch==='+')await key('+','Equal',187,8,'+');
   else if(ch==='-')await key('-','Minus',189,0,'-');
   else if(/^[0-9]$/u.test(ch))await key(ch,'Digit'+ch,ch.charCodeAt(0),0,ch);
   else throw new Error('Unsupported review key '+ch);
  }}
  async function source(text){await click('#source-A');await key('a','KeyA',65,2);await send('Input.insertText',{text});}
  async function pointer(nodeId,offset=0,phase='click'){
   const point=await evaluate(`(()=>{const s=window.__matrixReview.scene('A');const c=s.carets.find(c=>c.position.kind==='row'&&c.position.nodeId===${nodeId}&&c.position.offset===${offset});if(!c)throw new Error('Missing cell caret');const e=document.getElementById('editor-A'),svg=e.querySelector('svg');svg.scrollIntoView({block:'center'});const p=new DOMPoint(c.x,(c.top+c.bottom)/2).matrixTransform(svg.getScreenCTM());if(!e.contains(document.elementFromPoint(p.x,p.y)))throw new Error('Caret outside viewport');return{x:p.x,y:p.y};})()`);
   if(phase==='end')await send('Input.dispatchMouseEvent',{type:'mouseMoved',...point,button:'left',buttons:1});
   for(const type of (phase==='start'?['mousePressed']:phase==='end'?['mouseReleased']:['mousePressed','mouseReleased']))
    await send('Input.dispatchMouseEvent',{type,...point,button:'left',clickCount:1});
  }
  const grid=s=>s.document.analysis.tree.entries.find(e=>e.node.kind==='grid')?.node;
  const content=s=>({source:s.document.source,tree:s.document.analysis.tree,selection:s.selection,documentId:s.document.documentId});
  const containsNode=(value,id)=>!!value&&typeof value==='object'&&(value.id===id||Object.values(value).some(child=>containsNode(child,id)));
  async function checkpoint(label){
   const a=await inspect('A');assert.equal(await evaluate('document.getElementById("source-A").value'),a.document.source);
   assert.deepEqual(await inspect('B'),bBefore,label+': frame B changed');assert.equal((await stats()).sessions,2);
   assert.equal(await evaluate('document.querySelectorAll("#editor-A .vt-core-input").length'),1,'A must have one input owner');
   steps.push({label,source:a.document.source,revision:a.document.revision,selection:a.selection,undoDepth:a.undoDepth,stats:await stats()});return a;
  }
  await evaluate('window.__matrixReview.ready.then(()=>true)');const bBefore=await inspect('B');assert.deepEqual((await stats()).errors,[]);
  await click('#focus-A');await type('a');await key('a','KeyA',65,2);const single=await inspect('A');
  await click('#matrix-A');const inserted=await checkpoint('Create true 2x2 matrix from selected math');
  assert.equal(grid(inserted).environment,'matrix');assert.equal(grid(inserted).columns.length,2);assert.equal(grid(inserted).rows.length,2);
  assert.equal(inserted.document.documentId,single.document.documentId);assert.equal(inserted.document.analysis.tree.id,single.document.analysis.tree.id);
  assert.equal(grid(inserted).rows[0].cells[0].entries[0].node.id,single.document.analysis.tree.entries[0].node.id);
  for(const letter of ['b','c','d']){await key('Tab','Tab',9);await type(letter);}
  const filled=await checkpoint('Trusted Tab navigates and types into four independent cells');
  assert.deepEqual(grid(filled).rows.map(r=>r.cells.map(c=>c.entries.map(e=>e.node.value).join(''))),[['a','b'],['c','d']]);
  await key('Enter','Enter',13,8);const emptyRow=await checkpoint('Shift+Enter inserts same-width internal row');
  assert.equal(grid(emptyRow).rows.length,3);assert.equal(emptyRow.document.analysis.facts.grids,1);assert.equal((await stats()).enters.A,0);
  assert.equal(emptyRow.selection.focus.nodeId,grid(emptyRow).rows[2].cells[1].id);await type('e');const three=await checkpoint('Empty matrix cell is editable');
  await key('z','KeyZ',90,2);await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(filled));
  await key('y','KeyY',89,2);await key('y','KeyY',89,2);assert.deepEqual(content(await inspect('A')),content(three));await checkpoint('Keyboard undo/redo restores source, selection and exact internal IDs');
  await click('#rowBefore-A');assert.equal(grid(await inspect('A')).rows.length,4);await click('#rowDelete-A');assert.equal((await inspect('A')).document.source,three.document.source);
  await click('#columnBefore-A');assert.equal(grid(await inspect('A')).columns.length,3);await click('#columnDelete-A');assert.equal((await inspect('A')).document.source,three.document.source);
  await checkpoint('Visible row/column buttons preserve matrix content and frame ownership');
  let a=await inspect('A');await pointer(grid(a).rows[0].cells[1].id,1);await key('ArrowDown','ArrowDown',40);a=await inspect('A');
  assert.equal(a.selection.focus.nodeId,grid(a).rows[1].cells[1].id,'Down must retain matrix column');
  await key('Tab','Tab',9,8);a=await inspect('A');assert.equal(a.selection.focus.nodeId,grid(a).rows[1].cells[0].id);
  await pointer(grid(a).rows[0].cells[0].id,0);const boundary=content(await inspect('A'));await key('Backspace','Backspace',8);assert.deepEqual(content(await inspect('A')),boundary);
  await checkpoint('Same-column vertical and reverse-Tab navigation; boundary Backspace does not merge cells');
  const beforeEnter=await inspect('A');await key('Enter','Enter',13);assert.deepEqual(await inspect('A'),beforeEnter);assert.equal((await stats()).enters.A,1);
  await checkpoint('Ordinary Enter remains separate from contained row insertion');
  a=await inspect('A');await pointer(grid(a).rows[0].cells[0].id,0,'start');await pointer(grid(a).rows[1].cells[1].id,1,'end');
  const crossCell=await inspect('A');assert.equal(crossCell.selection.kind,'gridCells');
  await key('Backspace','Backspace',8);const cleared=await inspect('A');
  assert.equal(cleared.document.revision,crossCell.document.revision+1);
  for(const row of grid(cleared).rows.slice(0,2))for(const cell of row.cells)assert.equal(cell.entries.length,0);
  assert.deepEqual(grid(cleared).rows[2].cells.map(c=>c.entries.map(e=>e.node.value)),grid(crossCell).rows[2].cells.map(c=>c.entries.map(e=>e.node.value)));
  assert.equal(grid(cleared).id,grid(crossCell).id);assert.deepEqual(grid(cleared).columns,grid(crossCell).columns);
  await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(crossCell));
  assert.deepEqual((await stats()).errors,[]);
  assert.equal(await evaluate('document.querySelector("#editor-A .vt-core-input").disabled'),false);
  await pointer(grid(a).rows[0].cells[0].id,0);await type('q');assert.ok((await inspect('A')).document.source.includes('qa&b'));
  await key('z','KeyZ',90,2);assert.equal((await inspect('A')).document.source,crossCell.document.source);
  await checkpoint('Trusted cell rectangle drag/delete clears only the selected cells; undo restores identities');
  await source(String.raw`\begin{matrix}a&\begin{matrix}\frac{xy}{z}&b\\c&d\end{matrix}\\e&f\end{matrix}`);
  a=await inspect('A');const inner=grid(a).rows[0].cells[1].entries[0].node;const numerator=inner.rows[0].cells[0].entries[0].node.numerator;
  await pointer(numerator.id,1);await key('Enter','Enter',13,8);a=await checkpoint('Nested fraction insertion targets nearest matrix, not outer matrix/gathered');
  assert.equal(grid(a).rows.length,2);assert.equal(grid(a).rows[0].cells[1].entries[0].node.rows.length,3);assert.equal(a.document.analysis.facts.grids,2);
  assert.equal(grid(a).rows[0].cells[1].entries[0].node.rows[0].cells[0].entries[0].node.numerator.id,numerator.id);
  await source(String.raw`\begin{matrix}a&b&c\\d&&e\\f&g&h\end{matrix}`);
  const reopened=await checkpoint('A complete 3x3 matrix reopens inside one formula frame');
  const natural=await evaluate('window.__matrixReview.scene("A")');const zoomCases=[];
  for(const dpr of [1,1.25,2]){
   await send('Emulation.setDeviceMetricsOverride',{width:1100,height:760,deviceScaleFactor:dpr,mobile:false});
   for(const [i,zoom] of [.5,1,2].entries()){
    await click('#zoom-A');await key('Home','Home',36);for(let j=0;j<i;j++)await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
    assert.equal(await evaluate('Number(document.getElementById("zoom-A").value)'),zoom);
    for(let row=0;row<3;row++)for(let column=0;column<3;column++){
     const c=grid(reopened).rows[row].cells[column];await pointer(c.id,0);const positioned=await inspect('A');
     assert.equal(positioned.selection.focus.nodeId,c.id);assert.equal(positioned.selection.focus.offset,0);
     assert.equal(positioned.document.source,reopened.document.source);assert.equal(positioned.document.revision,reopened.document.revision);
     const scene=await evaluate('window.__matrixReview.scene("A")');assert.equal(scene.width,natural.width);assert.equal(scene.height,natural.height);assert.deepEqual(scene.gridColumns,natural.gridColumns);
     assert.deepEqual(await inspect('B'),bBefore);zoomCases.push({dpr,zoom,row,column,position:positioned.selection.focus});
    }
   }
  }
  await send('Emulation.clearDeviceMetricsOverride');await checkpoint('81 trusted cell-pointer/zoom/DPR checks including empty center cell');
  await click('#export-A');const svg=await evaluate('window.__matrixReview.exported("A")');assert.equal(svg.match(/<svg\b/g)?.length,1);assert.ok(!svg.includes('<text')&&!svg.includes('foreignObject'));
  for(const dimension of ['width','height'])assert.ok(Math.abs(Number(new RegExp('\\b'+dimension+'="([^"]+)"').exec(svg)?.[1])-natural[dimension])<1e-5);
  const beforeComposition=await inspect('A');await click('#focus-A');await send('Input.imeSetComposition',{text:'sudu',selectionStart:4,selectionEnd:4});
  assert.ok((await inspect('A')).composition);await key('Enter','Enter',13,8);a=await inspect('A');assert.equal(a.document.source,beforeComposition.document.source);assert.equal(grid(a).rows.length,3);assert.equal((await stats()).enters.A,1);
  await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});await checkpoint('Pending browser-protocol composition does not insert matrix rows');
  await source('xy');await click('#focus-A');await key('a','KeyA',65,2);const beforeCases=await inspect('A');
  await click('#cases-A');const casesCreated=await checkpoint('Cases button creates two identified columns inside the same formula');
  assert.equal(grid(casesCreated).environment,'cases');assert.equal(grid(casesCreated).columns.length,2);assert.equal(grid(casesCreated).rows.length,2);
  assert.equal(grid(casesCreated).rows[0].cells[0].entries[0].node.id,beforeCases.document.analysis.tree.entries[0].node.id);
  await key('Tab','Tab',9);await type('p');const casesBeforeBreak=await inspect('A');const casesColumns=grid(casesBeforeBreak).columns;
  await key('Enter','Enter',13,8);const caseBreak=await checkpoint('Cases Shift+Enter adds an internal row without nesting gathered');
  assert.equal(grid(caseBreak).rows.length,3);assert.equal(caseBreak.document.analysis.facts.grids,1);assert.deepEqual(grid(caseBreak).columns,casesColumns);
  assert.equal(caseBreak.selection.focus.nodeId,grid(caseBreak).rows[1].cells[1].id);
  await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(casesBeforeBreak));
  await key('y','KeyY',89,2);assert.deepEqual(content(await inspect('A')),content(caseBreak));
  await key('Enter','Enter',13);const caseEnter=await checkpoint('Cases ordinary Enter uses internal row behavior and never calls outer Enter');
  assert.equal(grid(caseEnter).rows.length,4);assert.equal((await stats()).enters.A,1);
  await source(String.raw`\begin{matrix}a&\begin{cases}\frac{xy}{z}&q\\c&d\end{cases}\\e&f\end{matrix}`);
  const caseNested=await inspect('A'),caseInner=grid(caseNested).rows[0].cells[1].entries[0].node,caseNumerator=caseInner.rows[0].cells[0].entries[0].node.numerator;
  await pointer(caseNumerator.id,1);await key('Enter','Enter',13,8);const caseNestedAfter=await checkpoint('Nearest nested Cases gains a row while its enclosing matrix and fraction remain intact');
  assert.equal(grid(caseNestedAfter).rows.length,2);assert.equal(grid(caseNestedAfter).rows[0].cells[1].entries[0].node.rows.length,3);
  assert.equal(grid(caseNestedAfter).rows[0].cells[1].entries[0].node.rows[0].cells[0].entries[0].node.numerator.id,caseNumerator.id);
  await source(String.raw`f(x)=\begin{cases}x^2&x\ge0\\-x&x<0\end{cases}`);
  const caseDisplay=await inspect('A'),caseScene=await evaluate('window.__matrixReview.scene("A")'),casesZoom=[];
  assert.ok(caseScene.glyphs.some(g=>g.nodeId===grid(caseDisplay).id),'Brace must have actual font glyphs owned by cases');
  for(const dpr of [1,1.25,2]){await send('Emulation.setDeviceMetricsOverride',{width:1100,height:760,deviceScaleFactor:dpr,mobile:false});
   for(const [i,zoom]of [.5,1,2].entries()){await click('#zoom-A');await key('Home','Home',36);for(let j=0;j<i;j++)await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
    for(let r=0;r<2;r++)for(let c=0;c<2;c++){const target=grid(caseDisplay).rows[r].cells[c];await pointer(target.id,0);const s=await inspect('A');assert.ok(containsNode(target,s.selection.focus.nodeId),'Pointer landed outside the intended case cell');assert.equal(s.document.source,caseDisplay.document.source);
     const now=await evaluate('window.__matrixReview.scene("A")');assert.equal(now.width,caseScene.width);assert.equal(now.height,caseScene.height);assert.deepEqual(now.gridColumns,caseScene.gridColumns);assert.deepEqual(await inspect('B'),bBefore);casesZoom.push({dpr,zoom,row:r,column:c,targetCell:target.id,position:s.selection.focus});}
   }}
  await send('Emulation.clearDeviceMetricsOverride');await checkpoint('Cases 36 real cell-pointer checks preserve natural geometry across zoom/DPR');
  await pointer(grid(caseDisplay).rows[0].cells[0].id,0,'start');await pointer(grid(caseDisplay).rows[1].cells[1].id,1,'end');
  const casesSelected=await inspect('A');assert.equal(casesSelected.selection.kind,'gridCells');
  await key('c','KeyC',67,2);await key('Backspace','Backspace',8);await key('v','KeyV',86,2);
  const casesPasted=await checkpoint('Cases browser copy and paste replace cells while retaining the cases owner and brace');
  assert.equal(casesPasted.document.source,caseDisplay.document.source);assert.equal(grid(casesPasted).id,grid(caseDisplay).id);
  assert.ok((await evaluate('window.__matrixReview.events("A")')).some(e=>e.type==='paste'&&e.trusted));
  await key('Escape','Escape',27);const beforeCaseIme=await inspect('A');await send('Input.imeSetComposition',{text:'sudu',selectionStart:4,selectionEnd:4});
  assert.ok((await inspect('A')).composition);await key('Enter','Enter',13,8);assert.equal((await inspect('A')).document.source,beforeCaseIme.document.source);
  await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});await checkpoint('Pending protocol composition creates no extra case row');
  await click('#export-A');const casesSvg=await evaluate('window.__matrixReview.exported("A")');assert.equal(casesSvg.match(/<svg\b/g)?.length,1);assert.ok(!casesSvg.includes('<text')&&!casesSvg.includes('vt-core-selection'));
  await source('a');await click('#focus-A');await key('a','KeyA',65,2);const beforeAligned=await inspect('A');
  await click('#aligned-A');const alignedCreated=await checkpoint('Aligned button preserves the selected atom in one internal environment');
  assert.equal(grid(alignedCreated).environment,'aligned');assert.equal(grid(alignedCreated).rows.length,1);assert.equal(grid(alignedCreated).columns.length,2);
  assert.equal(grid(alignedCreated).rows[0].cells[0].entries[0].node.id,beforeAligned.document.analysis.tree.entries[0].node.id);
  await key('Tab','Tab',9);await type('=b+c');const alignedSingle=await inspect('A');const alignedColumns=grid(alignedSingle).columns.slice();
  await key('Enter','Enter',13,8);const alignedSecond=await checkpoint('Aligned Shift+Enter creates a row sharing the same column identities');
  assert.equal(grid(alignedSecond).rows.length,2);assert.equal(alignedSecond.selection.focus.nodeId,grid(alignedSecond).rows[1].cells[0].id);assert.deepEqual(grid(alignedSecond).columns,alignedColumns);
  await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(alignedSingle));await key('y','KeyY',89,2);assert.deepEqual(content(await inspect('A')),content(alignedSecond));
  await type('d');await key('Tab','Tab',9);await type('=e+f');await key('Enter','Enter',13,8);await type('g');await key('Tab','Tab',9);await type('=h');
  const alignedThree=await checkpoint('Three aligned rows live in one formula host and share equality anchors');
  assert.equal(grid(alignedThree).rows.length,3);assert.equal(alignedThree.document.analysis.facts.grids,1);assert.equal((await stats()).enters.A,1);
  assert.deepEqual(grid(alignedThree).columns,alignedColumns);
  const alignedGeometry=await evaluate('window.__matrixReview.scene("A")');const eqXs=grid(alignedThree).rows.map(r=>alignedGeometry.regions.find(n=>n.nodeId===r.cells[1].entries[0].node.id).logical.xMin);
  assert.ok(eqXs.every(x=>Math.abs(x-eqXs[0])<1e-9));
  const alignedEnterBefore=await inspect('A');await key('Enter','Enter',13);assert.deepEqual(await inspect('A'),alignedEnterBefore);assert.equal((await stats()).enters.A,2);
  await pointer(grid(alignedThree).rows[2].cells[0].id,0);const beforeAlignMerge=await inspect('A');await key('Backspace','Backspace',8);
  assert.equal(grid(await inspect('A')).rows.length,2);await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(beforeAlignMerge));
  await pointer(grid(alignedThree).rows[0].cells[1].id,2);const interiorAlign=await inspect('A');await key('Enter','Enter',13,8);
  const interiorSplit=await checkpoint('Aligned interior cell split keeps suffix math in the corresponding column');
  assert.equal(grid(interiorSplit).rows.length,4);assert.equal(interiorSplit.document.analysis.facts.grids,1);
  await key('z','KeyZ',90,2);assert.deepEqual(content(await inspect('A')),content(interiorAlign));
  await source(String.raw`\begin{cases}\begin{aligned}\frac{xy}{z}&=a\\b&=c\end{aligned}&q\\d&e\end{cases}`);
  const nestedAlign=await inspect('A'),innerAlign=grid(nestedAlign).rows[0].cells[0].entries[0].node,alignNumerator=innerAlign.rows[0].cells[0].entries[0].node.numerator;
  await pointer(alignNumerator.id,1);await key('Enter','Enter',13,8);const nestedAlignAfter=await checkpoint('Nearest Aligned inside Cases adds a row without splitting its fraction or outer Cases');
  assert.equal(grid(nestedAlignAfter).rows.length,2);assert.equal(grid(nestedAlignAfter).rows[0].cells[0].entries[0].node.rows.length,3);
  assert.equal(grid(nestedAlignAfter).rows[0].cells[0].entries[0].node.rows[0].cells[0].entries[0].node.numerator.id,alignNumerator.id);
  for(const environment of ['align','align*']){const spelling='\\begin{'+environment+'}a&=b&&c&=d\\\\x&=y&&z&=w\\end{'+environment+'}';await source(spelling);
   const reopened=await inspect('A');assert.equal(reopened.document.source,spelling);assert.equal(reopened.document.analysis.facts.grids,1);assert.equal(grid(reopened).columns.length,5);}
  await source(String.raw`\begin{aligned}long&=b&x&=y\\a&=cd&uvw&=z\end{aligned}`);
  const alignmentDisplay=await inspect('A'),alignmentScene=await evaluate('window.__matrixReview.scene("A")'),alignedZoom=[];
  for(const dpr of [1,1.25,2]){await send('Emulation.setDeviceMetricsOverride',{width:1100,height:760,deviceScaleFactor:dpr,mobile:false});
   for(const [i,zoom]of [.5,1,2].entries()){await click('#zoom-A');await key('Home','Home',36);for(let j=0;j<i;j++)await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);
    for(let r=0;r<2;r++)for(let c=0;c<4;c++){const target=grid(alignmentDisplay).rows[r].cells[c];await pointer(target.id,c%2?1:0);
     const now=await inspect('A');assert.ok(containsNode(target,now.selection.focus.nodeId),'Aligned pointer selected another cell');assert.equal(now.document.source,alignmentDisplay.document.source);
     const geometry=await evaluate('window.__matrixReview.scene("A")');assert.equal(geometry.width,alignmentScene.width);assert.equal(geometry.height,alignmentScene.height);assert.deepEqual(geometry.gridColumns,alignmentScene.gridColumns);assert.deepEqual(await inspect('B'),bBefore);
     alignedZoom.push({dpr,zoom,row:r,column:c,targetCell:target.id,position:now.selection.focus});}
   }}
  await send('Emulation.clearDeviceMetricsOverride');await checkpoint('Aligned 72 cell-pointer cases at nine zoom/DPR settings retain shared columns and frame B');
  await pointer(grid(alignmentDisplay).rows[0].cells[0].id,0,'start');await pointer(grid(alignmentDisplay).rows[1].cells[3].id,1,'end');
  assert.equal((await inspect('A')).selection.kind,'gridCells');await key('c','KeyC',67,2);await key('Backspace','Backspace',8);await key('v','KeyV',86,2);
  const alignPasted=await checkpoint('Actual browser clipboard replaces aligned cell content without changing environment or column IDs');
  assert.equal(alignPasted.document.source,alignmentDisplay.document.source);assert.equal(grid(alignPasted).id,grid(alignmentDisplay).id);assert.deepEqual(grid(alignPasted).columns,grid(alignmentDisplay).columns);
  await key('Escape','Escape',27);const alignedImeBefore=await inspect('A');await send('Input.imeSetComposition',{text:'sudu',selectionStart:4,selectionEnd:4});assert.ok((await inspect('A')).composition);
  await key('Enter','Enter',13,8);assert.equal((await inspect('A')).document.source,alignedImeBefore.document.source);await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});
  await checkpoint('Pending protocol composition cannot insert an additional aligned row');await click('#export-A');const alignedSvg=await evaluate('window.__matrixReview.exported("A")');
  assert.equal(alignedSvg.match(/<svg\b/g)?.length,1);assert.ok(!alignedSvg.includes('<text')&&!alignedSvg.includes('vt-core-selection'));
  const finalErrors=(await stats()).errors;assert.deepEqual(finalErrors,[]);
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});writeFileSync(resolve(out,'matrix-two-independent-frames.png'),Buffer.from(screenshot.data,'base64'));
  const events=await evaluate('window.__matrixReview.events("A")');
  return {status:'PASS',browser:browserVersion.product,hostBuild:manifest.buildId,bootstrapSha256:sha(bootstrap),wasmSha256:manifest.assets.find(a=>a.id==='core').sha256,
   steps,zoomCases,casesZoom,alignedZoom,frameBUnchanged:true,actualBrowserKeysAndPointer:true,events:events.length,trustedEvents:events.filter(e=>e.trusted).length,
   rectangularDragClearUndo:true,unexpectedHostErrors:0,physicalWindowsIme:false,protocolComposition:true,screenshot:relative(root,resolve(out,'matrix-two-independent-frames.png')).replaceAll('\\','/'),
   notCovered:['Implicit ragged cells','Aligned placement/row-gap options','Physical Windows IME','Office/installed application']};
 });
}catch(error){report={status:'FAILED',error:String(error),steps,hostBuild:manifest.buildId,bootstrapSha256:sha(bootstrap)};throw error;}
finally{if(report){writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({receipt:relative(root,resolve(out,'report.json')).replaceAll('\\','/'),...report},null,2));}}

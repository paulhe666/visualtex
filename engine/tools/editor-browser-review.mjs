// End-to-end input through actual browser events and the reusable editor host.
// CDP composition is not physical Windows IME acceptance.
import assert from 'node:assert/strict';
import {createHash} from 'node:crypto';
import {readFileSync,mkdirSync,writeFileSync} from 'node:fs';
import {dirname,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {withOwnedBrowser} from './owned-browser-review.mjs';
import {createSandboxAssets} from './sandbox-assets.mjs';
const currentHostAssets=createSandboxAssets();
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const sha=b=>createHash('sha256').update(b).digest('hex');
const wasm=readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm_bg.wasm'));
const math=readFileSync(process.env.VISUALTEX_TEST_MATH_FONT??resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'));
const text=readFileSync(process.env.VISUALTEX_TEST_TEXT_FONT??'C:/Windows/Fonts/NotoSansSC-VF.ttf');
const assets=new Map([
 ['/core.js',{type:'text/javascript',bytes:readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm.js'))}],
 ['/core.wasm',{type:'application/wasm',bytes:wasm}],['/math.otf',{type:'font/otf',bytes:math}],['/text.ttf',{type:'font/ttf',bytes:text}],
]);
for(const [path,asset] of currentHostAssets.assets){
 if(!path.startsWith('/host/')||!path.endsWith('.js'))continue;
 assets.set(`/${path.slice('/host/'.length)}`,asset);
}
assert.ok(assets.has('/editor-host.js'),'Current compiled editor host module is missing');
const html='<!doctype html><meta charset="utf-8"><title>VisualTeX autonomous editor input test</title><style>body{margin:44px;font:17px sans-serif;color:#111}h1{font-size:25px}button{font:inherit;margin:0 8px 0 0;padding:10px 18px}#mount{margin:48px 24px}.vt-core-position{margin-top:30px;font-size:14px}pre{padding:16px;background:#f2f2f2;white-space:pre-wrap}body{min-height:1000px}</style><h1>VisualTeX 自研编辑器 · 输入验证</h1><p>从空白公式开始，通过按钮、键盘与浏览器组合输入构造。</p><div id="toolbar"><button data-structure="fraction">分式</button><button data-structure="subscript">下标</button><button data-structure="superscript">上标</button><button data-structure="text">文本</button></div><div id="mount"></div><pre id="source"></pre>';
const out=resolve(root,'target/editor-browser-review');mkdirSync(out,{recursive:true});
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion})=>{
 const steps=[];
 const inspect=()=>evaluate('window.host.inspect()');
 async function state(label){const s=await inspect();const errors=await evaluate('window.host.diagnostics');assert.deepEqual(errors,[],label);steps.push({label,source:s.document.source,revision:s.document.revision,selection:s.selection,undoDepth:s.undoDepth});return s;}
 async function click(selector){const p=await evaluate(`(()=>{const r=document.querySelector(${JSON.stringify(selector)}).getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2};})()`);
   await send('Input.dispatchMouseEvent',{type:'mousePressed',...p,button:'left',clickCount:1});await send('Input.dispatchMouseEvent',{type:'mouseReleased',...p,button:'left',clickCount:1});}
 async function key(key,code=key,virtual=0,modifiers=0){await send('Input.dispatchKeyEvent',{type:'keyDown',key,code,windowsVirtualKeyCode:virtual,nativeVirtualKeyCode:virtual,modifiers});await send('Input.dispatchKeyEvent',{type:'keyUp',key,code,windowsVirtualKeyCode:virtual,nativeVirtualKeyCode:virtual,modifiers});}
 const type=text=>send('Input.insertText',{text});
 try{
  await evaluate(`(async()=>{
   const {default:init,EngineRuntime}=await import('/core.js');const {EngineEditorHost}=await import('/editor-host.js');
   const data=await Promise.all(['/core.wasm','/math.otf','/text.ttf'].map(p=>fetch(p).then(r=>r.arrayBuffer())));await init({module_or_path:data[0]});
   window.runtime=new EngineRuntime();const f=data.slice(1).map(b=>{const r=JSON.parse(runtime.registerFont(new Uint8Array(b),0));if(r.status!=='ok')throw Error(JSON.stringify(r));return r.result.font.id;});
   window.enterCalls=0;window.nativeEvents=[];
   window.host=new EngineEditorHost(runtime,document.getElementById('mount'),{documentId:'actual-browser-editor',
      layout:{mathFont:f[0],textFont:f[1],style:'display',fontSizePx:28,paddingEm:0.125},
      onChange:s=>document.getElementById('source').textContent=s.document.source,onEnter:()=>window.enterCalls++});
   for(const button of document.querySelectorAll('[data-structure]'))button.addEventListener('click',()=>host.applyStructure(button.dataset.structure));
   for(const name of ['keydown','beforeinput','input','compositionstart','compositionupdate','compositionend'])host.input.textarea.addEventListener(name,e=>nativeEvents.push({type:e.type,data:e.data??null,inputType:e.inputType??null,key:e.key??null,composing:e.isComposing??false,trusted:e.isTrusted}));
   host.focus();return true;
  })()`);
  assert.equal((await state('empty start')).document.source,'');
  await click('[data-structure="fraction"]');await type('x');await type('_');await type('1');await key('ArrowRight','ArrowRight',39);await type('+');
  await click('[data-structure="text"]');const beforeIme=await state('fraction slots and text caret');
  await send('Input.imeSetComposition',{text:'su',selectionStart:2,selectionEnd:2});
  const preedit=await inspect();assert.equal(preedit.document.source,beforeIme.document.source);assert.ok(preedit.composition);
  await send('Input.imeSetComposition',{text:'速度',selectionStart:2,selectionEnd:2});await type('速度');
  const afterIme=await state('Chinese composition committed once');assert.equal(afterIme.document.source.split('速度').length-1,1);assert.equal(afterIme.composition,null);
  assert.equal(afterIme.undoDepth,beforeIme.undoDepth+1);
  await key('ArrowRight','ArrowRight',39);await key('ArrowDown','ArrowDown',40);await type('y');await type('^');await type('2');
  const complete=await state('fully constructed with UI input');
  assert.equal(complete.document.canonicalLatex,String.raw`\frac{x_{1}+\text{速度}}{y^{2}}`);
  assert.ok(complete.document.analysis.exportable);
  const beforeCancel=complete.document.source;
  await send('Input.imeSetComposition',{text:'temporary',selectionStart:9,selectionEnd:9});
  await send('Input.imeSetComposition',{text:'',selectionStart:0,selectionEnd:0});
  assert.equal((await state('composition cancel')).document.source,beforeCancel);
  const caret=await evaluate(`(()=>{const s=host.layoutSnapshot;const c=s.carets.find(c=>c.position.kind==='text'&&c.position.grapheme===1);return {position:c.position,point:host.view.layoutToClient(c.x,(c.top+c.bottom)/2)};})()`);
  await send('Input.dispatchMouseEvent',{type:'mousePressed',...caret.point,button:'left',clickCount:1});
  await send('Input.dispatchMouseEvent',{type:'mouseReleased',...caret.point,button:'left',clickCount:1});
  const clicked=await state('pointer to text grapheme');assert.equal(clicked.selection.kind,'range');assert.deepEqual(clicked.selection.focus,caret.position);
  await send('Input.imeSetComposition',{text:'率',selectionStart:1,selectionEnd:1});
  await key('Enter','Enter',13);assert.equal(await evaluate('window.enterCalls'),0,'IME candidate Enter also triggered a document action');
  await type('率');const extra=await state('candidate Enter is composition-only');assert.equal(extra.document.source.split('率').length-1,1);
  await key('z','KeyZ',90,2);assert.equal((await state('one undo removes one composition')).document.source,beforeCancel);
  await key('y','KeyY',89,2);assert.equal((await state('redo restores composition')).document.source,extra.document.source);
  await key('z','KeyZ',90,2);
  await key('ArrowLeft','ArrowLeft',37,8);await key('Backspace','Backspace',8);
  const deleted=await state('shift selection deletes one complete Chinese character');
  assert.ok(deleted.document.source.includes(String.raw`\text{度}`),deleted.document.source);
  await key('z','KeyZ',90,2);const restored=await state('undo restores original formula');assert.equal(restored.document.source,beforeCancel);
  const exported=await evaluate('window.host.exportSvg()');assert.ok(exported.startsWith('<svg'));assert.ok(!exported.includes('<text'));
  // Export creates a new scene token. A subsequent pointer query must still work.
  await send('Input.dispatchMouseEvent',{type:'mousePressed',...caret.point,button:'left',clickCount:1});
  await send('Input.dispatchMouseEvent',{type:'mouseReleased',...caret.point,button:'left',clickCount:1});await state('pointer after export');
  const noLayout=await evaluate(`(()=>{const before=host.view.stats;for(let i=0;i<100;i++)host.view.setCaretVisible(i%2===0);return {before,after:host.view.stats};})()`);
  assert.deepEqual(noLayout.before,noLayout.after);
  const accessible=await evaluate('window.host.input.textarea.getAttribute("aria-label")');assert.ok(accessible.includes('分子')&&accessible.includes('文本'),accessible);
  const events=await evaluate('window.nativeEvents');assert.ok(events.length>20);
  const untrusted=events.filter(e=>!e.trusted);
  // This Edge/CDP host marks compositionend as untrusted despite browser-driven
  // compositionstart/update. Preserve this evidence, not a false OS-IME claim.
  assert.ok(events.filter(e=>e.type==='keydown'||e.type==='beforeinput'||e.type==='input').every(e=>e.trusted));
  assert.ok(untrusted.every(e=>e.type==='compositionend'),JSON.stringify(untrusted));
  const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});writeFileSync(resolve(out,'editor-input.png'),Buffer.from(screenshot.data,'base64'));
  await evaluate('window.host.dispose(); window.runtime.free(); true');
  return {status:'PASS',browser:browserVersion.product,startedWithEmptyDocument:true,formula:restored.document.source,
    hostBuildId:currentHostAssets.manifest.buildId,hostModuleHashes:Object.fromEntries(['editor-host','input-bridge','svg-view'].map(name=>[name,sha(assets.get(`/${name}.js`).bytes)])),
    eventCount:events.length,trustedBrowserEventCount:events.filter(e=>e.trusted).length,untrustedCompositionEndCount:untrusted.length,
    physicalWindowsIme:false,steps,nativeEvents:events,accessiblePosition:accessible,
    wasmSha256:sha(wasm),svgSha256:sha(Buffer.from(exported)),screenshot:'target/editor-browser-review/editor-input.png',
    scope:'Reusable host, actual browser button/pointer/keyboard events and CDP composition. No physical Windows IME or Office acceptance.'};
 }catch(error){
   let evidence;try{evidence=await evaluate('({snapshot:window.host?.inspect(),errors:window.host?.diagnostics,events:window.nativeEvents})');}catch{}
   writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,evidence},null,2));throw error;
 }
});
writeFileSync(resolve(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({...report,nativeEvents:`${report.eventCount} events retained in report.json`},null,2));

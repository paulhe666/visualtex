// Production desktop UI, app-private fonts and resource-recovery regression.
// The test never attaches to installed VisualTeX, user browsers, Word or PPT.
import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, statSync, mkdirSync, writeFileSync } from 'node:fs';
import { resolve, dirname, join, relative, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { withOwnedBrowser } from '../../../engine/tools/owned-browser-review.mjs';

const app = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repo = resolve(app, '../..');
const dist = resolve(app, 'dist');
const stamp = new Date().toISOString().replaceAll(':','-').replaceAll('.','-');
const out = resolve(repo, 'engine/target/p16-private-font-product', stamp);
mkdirSync(out, {recursive:true});
const sha = b => createHash('sha256').update(b).digest('hex');
const source = String.raw`\frac{1}{\psi_{1,2}}\frac{\mathrm{d}^{2}}{\mathrm{d}x^{2}}\psi_{1,2}+\frac{2m(E-u)}{\hbar^{2}}=0,\quad\psi_2\frac{\mathrm{d}^2\psi_1}{\mathrm{d}x^2}=\psi_1\frac{\mathrm{d}^2\psi_2}{\mathrm{d}x^2}\Rightarrow\frac{\mathrm{d}}{\mathrm{d}x}\left(\psi_2\frac{\mathrm{d}\psi_1}{\mathrm{d}x}-\psi_1\frac{\mathrm{d}\psi_2}{\mathrm{d}x}\right)=0`;
const assets = new Map();
const mime = p => ({'.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.wasm':'application/wasm','.ttf':'font/ttf','.otf':'font/otf','.png':'image/png','.jpg':'image/jpeg'}[extname(p)] ?? 'application/octet-stream');
function collect(path, prefix='') {
  for (const name of readdirSync(path)) {
    const file = join(path,name), rel = prefix ? `${prefix}/${name}` : name;
    if (statSync(file).isDirectory()) collect(file,rel);
    else if (rel !== 'index.html') assets.set('/'+rel,{type:mime(file),bytes:readFileSync(file)});
  }
}
collect(dist);
const bootstrap = `(() => {
 const q=new URLSearchParams(location.search),font=q.get('font')||'katex',chinese=q.get('chinese')||'system';
 const source=${JSON.stringify(source)};
 const lines=[{id:'reference',latex:source,mode:'display'},{id:'empty-one',latex:'',mode:'display'},{id:'empty-two',latex:'',mode:'display'}];
 localStorage.setItem('visualtex.formula-letter-font',font);localStorage.setItem('visualtex.formula-chinese-font',chinese);
 localStorage.setItem('visualtex-editor',JSON.stringify({version:0,state:{title:'内置字体与原界面验收',lines,activeLineId:'reference',formulaLetterFont:font,formulaChineseFont:chinese,editorLayout:'classic',formulaAlignment:'left',showLineNumbers:false,highlightActiveLine:false,formulaInsetLeft:8,formulaInsetRight:8,formulaRowVerticalInset:5,zoom:.65,language:'cn',theme:'light',sourceOpen:false,latexCodeFormat:'raw',classicTileWidth:300,classicDockHeight:300,checkUpdatesOnStartup:false,inputBehavior:{showStructuredCommandSuggestions:true,showOtherCommandSuggestions:true}}}));
 window.__fontReview={externalRequests:[],errors:[],fontRequests:[],failFont:q.has('fail-font')};
 window.addEventListener('error',e=>window.__fontReview.errors.push(String(e.error||e.message)));
 window.addEventListener('unhandledrejection',e=>window.__fontReview.errors.push(String(e.reason)));
 const realFetch=window.fetch.bind(window);window.fetch=async(input,init)=>{
   const url=new URL(typeof input==='string'?input:input.url,location.href);
   if(url.origin!==location.origin){window.__fontReview.externalRequests.push(url.href);throw Error('External access blocked in offline acceptance');}
   if(/\\.(?:ttf|otf)$/.test(url.pathname)){
     window.__fontReview.fontRequests.push(url.pathname);
     if(window.__fontReview.failFont&&url.pathname.includes('KaTeX_Main-Regular')){return new Response('transient test font failure',{status:503});}
   }
   return realFetch(input,init);
 };
})();`;
assets.set('/font-review-bootstrap.js',{type:'text/javascript; charset=utf-8',bytes:Buffer.from(bootstrap)});
const html = readFileSync(join(dist,'index.html'),'utf8').replace('<script type="module"','<script src="/font-review-bootstrap.js"></script><script type="module"');
const cases=[];let report;
const wait = async(evaluate, expression, label, timeout=20000) => {
  const start=Date.now();let last;
  while(Date.now()-start<timeout){last=await evaluate(`Boolean(${expression})`);if(last)return last;await new Promise(r=>setTimeout(r,80));}
  throw Error(`Timed out: ${label}; ${JSON.stringify(last)}`);
};
try {
  report=await withOwnedBrowser(assets,html,async({send,evaluate,origin,browserVersion})=>{
    await send('Emulation.setDeviceMetricsOverride',{width:1586,height:900,deviceScaleFactor:1,mobile:false});
    await send('Network.enable');await send('Network.setCacheDisabled',{cacheDisabled:true});
    const state=()=>evaluate(`(()=>{const e=document.querySelector('[data-engine="visualtex-core"]');return {ready:e?.dataset.coreReady,requested:e?.dataset.requestedFont,actual:e?.dataset.resolvedFont,source:e?.dataset.fontSource,shell:!!e?.classList.contains('editor-surface'),stack:!!e?.querySelector('.mathfield-stack'),rows:e?.querySelectorAll('.formula-line').length||0,inputs:e?.querySelectorAll('.vt-core-input').length||0,legacy:document.querySelectorAll('math-field').length,ink:[...e?.querySelectorAll('.core-formula-line')||[]].map(x=>x.querySelectorAll('svg use,svg path,svg rect').length),sources:JSON.parse(localStorage.getItem('visualtex-editor')).state.lines.map(x=>x.latex),previews:document.querySelectorAll('[data-preview-renderer="visualtex-core"]').length,previewFailures:[...document.querySelectorAll('[data-preview-renderer="unavailable"]')].map(x=>x.textContent),loading:document.querySelectorAll('[data-preview-renderer="loading"]').length,errors:window.__fontReview.errors,external:window.__fontReview.externalRequests,fontRequests:window.__fontReview.fontRequests}})()`);
    const screenshot=async(name)=>{const shot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});writeFileSync(join(out,name),Buffer.from(shot.data,'base64'));return name;};
    const expected={katex:'KaTeX',times:'TeX Gyre Termes',cambria:'STIX Two Math',stix:'STIX Two Math',palatino:'TeX Gyre Pagella',helvetica:'TeX Gyre Heros'};
    const fonts=process.argv.includes('--quick')?['katex','times']:Object.keys(expected);
    const chinese=process.argv.includes('--quick')?['system']:['system','songti','kaiti','heiti','pingfang'];
    for(const font of fonts)for(const chineseFont of chinese){
      await send('Page.navigate',{url:`${origin}/?font=${font}&chinese=${chineseFont}`});
      await wait(evaluate,`document.querySelector('[data-core-ready="true"]')&&document.querySelectorAll('.vt-core-input').length===3`,'three retained core rows');
      await wait(evaluate,`document.querySelectorAll('[data-preview-renderer="loading"]').length===0`,'palette previews settled');
      const current=await state();
      assert.equal(current.actual,expected[font]);assert.equal(current.source,'bundled');assert.equal(current.shell,true);assert.equal(current.stack,true);assert.equal(current.rows,3);assert.equal(current.legacy,0);assert.ok(current.ink[0]>70,'reference formula has no actual mathematical ink');
      assert.deepEqual(current.sources,[source,'',''],'opening/font resolution must not rewrite source');assert.equal(current.external.length,0);assert.equal(current.errors.length,0);
      assert.ok(current.previews>40,'the toolbar did not render its previews');
      assert.deepEqual(current.previewFailures,[],'a shipped palette template has no mathematical preview');
      if(font==='katex'&&chineseFont==='system'){
        current.screenshot=await screenshot('default-katex-reference.png');
        const geometry=await evaluate(`[...document.querySelectorAll('.formula-line')].map(e=>{const r=e.getBoundingClientRect(),s=getComputedStyle(e);return {x:r.x,y:r.y,w:r.width,h:r.height,border:s.borderBottomWidth}})`);
        assert.ok(geometry.every(x=>x.h>=30&&x.border==='1px'));current.geometry=geometry;
      }
      cases.push({font,chineseFont,...current});
      console.log(JSON.stringify({font,chineseFont,status:'PASS',previews:current.previews,previewFailures:current.previewFailures}));
    }
    // The original product candidate shell and the editable numerator/denominator
    // must remain visible in the same three-row editor, not a test-only host.
    await send('Page.navigate',{url:`${origin}/?font=katex&chinese=system`});
    await wait(evaluate,`document.querySelectorAll('.vt-core-input').length===3`,'candidate fixture');
    await evaluate(`document.querySelectorAll('.vt-core-input')[1].focus();true`);
    const key=async(key,code,vk,text)=>{for(const type of ['keyDown','keyUp'])await send('Input.dispatchKeyEvent',{type,key,code,windowsVirtualKeyCode:vk,nativeVirtualKeyCode:vk,...(type==='keyDown'&&text!==undefined?{text}:{})});};
    await key('\\','Backslash',220,'\\');for(const ch of 'fra')await key(ch,'Key'+ch.toUpperCase(),ch.toUpperCase().charCodeAt(0),ch);
    await wait(evaluate,`Boolean(document.querySelector('.suggestion-popup .suggestion-item.is-selected svg'))`,'product candidate mathematical preview');
    await new Promise(r=>setTimeout(r,220));
    const candidateShot=await screenshot('original-product-candidate.png');
    await key('Enter','Enter',13);
    await wait(evaluate,`document.querySelectorAll('.core-formula-line')[1].querySelectorAll('.vt-core-empty-slot').length===2`,'two fraction slots');
    const structureShot=await screenshot('original-product-fraction-slots.png');
    assert.equal((await state()).rows,3,'candidate confirmation must not split outer rows');
    // A missing private font is a local font-plan failure, never a replacement
    // UI. The old editor shell/source stays mounted on the verified safety face.
    await send('Page.navigate',{url:`${origin}/?font=katex&chinese=system&fail-font=1`});
    await wait(evaluate,`document.querySelector('[data-core-ready="true"]')&&document.querySelectorAll('.vt-core-input').length===3`,'font failure keeps the core editor mounted');
    await wait(evaluate,`document.querySelectorAll('[data-preview-renderer="loading"]').length===0`,'safety previews settled');
    let failed=await state();assert.equal(failed.shell,true);assert.equal(failed.stack,true);assert.equal(failed.rows,3);assert.equal(failed.legacy,0);assert.deepEqual(failed.sources,[source,'','']);assert.equal(failed.external.length,0);assert.ok(failed.previews>40,'font failure blanked the formula palette');
    // The shared retry event refreshes both the editor font plan and static
    // previews without recreating settings/document state.
    await evaluate(`window.__fontReview.failFont=false;window.dispatchEvent(new Event('visualtex-core-preview-retry'));true`);
    await wait(evaluate,`document.querySelector('[data-resolved-font="KaTeX"]')&&document.querySelectorAll('.vt-core-input').length===3`,'exact KaTeX recovery without deleting settings');
    await wait(evaluate,`document.querySelectorAll('[data-preview-renderer="loading"]').length===0`,'palette recovered');
    const recovered=await state();assert.ok(recovered.previews>40);assert.deepEqual(recovered.previewFailures,[]);assert.equal(recovered.errors.length,0);assert.deepEqual(recovered.sources,[source,'','']);
    return {status:'PASS',browser:browserVersion.product,cases,defaultScreenshot:'default-katex-reference.png',candidateShot,structureShot,recovery:{shellPreserved:true,sourcePreserved:true,retrySucceeded:true,recoveredPreviews:recovered.previews},fontAssets: [...assets.keys()].filter(x=>/\.(ttf|otf)$/.test(x)).length,systemFontAccess:false,installedPackageTested:false,officeAutomation:false,distEntrySha256:sha(readFileSync(join(dist,'index.html')))};
  });
} catch(error){report={status:'FAILED',error:String(error.stack??error),cases};process.exitCode=1;}
finally{writeFileSync(join(out,'report.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({status:report.status,report:relative(repo,join(out,'report.json')).replaceAll('\\','/'),error:report.error??null,cases:cases.length},null,2));}

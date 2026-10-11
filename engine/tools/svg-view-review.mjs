// Real isolated browser rendering and pointer checks. No physical OS IME claim.
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
const math=readFileSync(process.env.VISUALTEX_TEST_MATH_FONT??resolve(root,'../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'));
const text=readFileSync(process.env.VISUALTEX_TEST_TEXT_FONT??'C:/Windows/Fonts/NotoSansSC-VF.ttf');
const wasm=readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm_bg.wasm'));
const assets=new Map([
 ['/core.js',{type:'text/javascript',bytes:readFileSync(resolve(root,'web/pkg/visualtex_engine_wasm.js'))}],
 ['/core.wasm',{type:'application/wasm',bytes:wasm}],
 ...[...currentHostAssets.assets].filter(([path])=>path.startsWith('/host/')),
 ['/math.otf',{type:'font/otf',bytes:math}],['/text.ttf',{type:'font/ttf',bytes:text}],
]);
const html='<!doctype html><meta charset="utf-8"><title>VisualTeX autonomous core — view regression</title><style>body{font:18px sans-serif;margin:48px;color:#202020}h1{font-size:25px}#mount{margin:60px 24px}#details{margin-top:48px}body{min-height:1600px}</style><h1>VisualTeX 自研公式引擎 · 隔离验证</h1><p>Rust / WASM 排版，SVG View 绘制。正式编辑器未切换。</p><div id="mount"></div><p id="details">分式、上下标、中文文本与结构光标</p>';
const output=resolve(root,'target/svg-review');mkdirSync(output,{recursive:true});
const report=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion,requests})=>{
 const start=await evaluate(`(async()=>{
   const {default:init,EngineRuntime}=await import('/core.js');
   const {SvgSceneView}=await import('/host/svg-view.js');
   const [wasm,math,text]=await Promise.all(['/core.wasm','/math.otf','/text.ttf'].map(p=>fetch(p).then(r=>r.arrayBuffer())));
   await init({module_or_path:wasm});const rt=new EngineRuntime();
   const mf=JSON.parse(rt.registerFont(new Uint8Array(math),0)).result.font.id;
   const tf=JSON.parse(rt.registerFont(new Uint8Array(text),0)).result.font.id;
   const call=c=>{const r=JSON.parse(rt.dispatch(JSON.stringify({protocol:1,command:c})));if(r.status!=='ok')throw Error(JSON.stringify(r));return r.result;};
   const {handle}=call({op:'create',documentId:'svg-review-document',source:${JSON.stringify(String.raw`\frac{x_1+\text{速度}}{y^2}`)}});
   const opts={mathFont:mf,textFont:tf,style:'display',fontSizePx:32,paddingEm:0.125};
   const rendered=call({op:'layout',handle,expectedRevision:0,options:opts,includeSvg:true});
   window.pointerEvents=[];
   const view=new SvgSceneView(document.getElementById('mount'),{onPointer:e=>window.pointerEvents.push(e)});
   view.setSnapshot(rendered.snapshot,'分式：分子为 x 下标 1 加速度；分母为 y 上标 2。');
   const caret=rendered.snapshot.carets.find(c=>c.position.kind==='text'&&c.position.grapheme===1);
   const geometry=call({op:'queryScene',handle,expectedRevision:0,sceneToken:rendered.sceneToken,query:{kind:'selection',selection:{kind:'range',anchor:caret.position,focus:caret.position}}}).result.caretGeometry;
   view.setOverlay({...view.snapshotIdentity,caret:caret.position,caretGeometry:geometry,rectangles:[]});
   window.probe={rt,call,handle,opts,rendered,view,caret,SvgSceneView};
   return {width:rendered.snapshot.width,height:rendered.snapshot.height,glyphs:rendered.snapshot.glyphs.length,
     domGlyphs:view.svg.querySelectorAll('use').length,stats:view.stats,fonts:rendered.snapshot.fonts};
 })()`);
 assert.equal(start.domGlyphs,start.glyphs);assert.equal(start.fonts.length,2);
 const mapping=[];
 for(const dpr of [1,1.25,2])for(const zoom of [0.5,1,2])for(const scroll of [0,180]){
   await send('Emulation.setDeviceMetricsOverride',{width:1100,height:850,deviceScaleFactor:dpr,mobile:false});
   const c=await evaluate(`(()=>{const p=window.probe;p.view.setZoom(${zoom});document.getElementById('mount').style.marginTop='300px';window.scrollTo(0,${scroll});
      const x=p.caret.x,y=(p.caret.top+p.caret.bottom)/2;const point=p.view.layoutToClient(x,y);
      const round=p.view.clientToLayout(point.x,point.y);window.pointerEvents=[];
      return {point,round,expected:{xEm:x,yEm:y},dpr:devicePixelRatio,stats:p.view.stats};})()`);
   assert.ok(Math.abs(c.round.xEm-c.expected.xEm)<1e-9);assert.ok(Math.abs(c.round.yEm-c.expected.yEm)<1e-9);
   await send('Input.dispatchMouseEvent',{type:'mousePressed',x:c.point.x,y:c.point.y,button:'left',clickCount:1});
   await send('Input.dispatchMouseEvent',{type:'mouseReleased',x:c.point.x,y:c.point.y,button:'left',clickCount:1});
   const events=await evaluate('window.pointerEvents');
   assert.deepEqual(events.map(e=>e.phase),['down','up']);
   for(const e of events){assert.ok(Math.abs(e.xEm-c.expected.xEm)<0.01);assert.ok(Math.abs(e.yEm-c.expected.yEm)<0.01);assert.equal(e.revision,0);}
   assert.equal(c.stats.sceneBuilds,1);mapping.push({zoom,dpr,scroll,events:events.length});
 }
 const overlays=await evaluate(`(()=>{const p=window.probe;const before=p.view.stats;
   for(let i=0;i<100;i++)p.view.setCaretVisible(i%2===0);
   const after=p.view.stats;if(before.sceneBuilds!==after.sceneBuilds||before.overlayBuilds!==after.overlayBuilds)throw Error('Blink rebuilt geometry');
   const old=p.rendered.snapshot;const changed=p.call({op:'replaceSource',handle:p.handle,expectedRevision:0,source:'x+y'});
   const next=p.call({op:'layout',handle:p.handle,expectedRevision:1,options:p.opts,includeSvg:true});p.view.setSnapshot(next.snapshot);
   let staleLayout=false,staleOverlay=false;try{p.view.setSnapshot(old);}catch{staleLayout=true;}
   try{p.view.setOverlay({documentId:old.documentId,revision:old.revision,sourceId:old.sourceId,caret:p.caret.position,caretGeometry:null,rectangles:[]});}catch{staleOverlay=true;}
   const beforeBad=p.view.svg.outerHTML;const bad=structuredClone(next.snapshot);bad.resources[0].shape.path.push({op:'line',points:[NaN,0]});
   let rejectedBad=false;try{p.view.setSnapshot(bad);}catch{rejectedBad=true;}
   if(beforeBad!==p.view.svg.outerHTML)throw Error('Bad snapshot partially changed DOM');
   const other=new p.SvgSceneView(document.getElementById('mount'));other.setSnapshot(next.snapshot);
   const ids=[...document.querySelectorAll('[id^="vt-core-view-"]')].map(n=>n.id);
   if(new Set(ids).size!==ids.length)throw Error('Multiple SVG views share resource IDs');other.dispose();
   const caller=structuredClone(next.snapshot);p.view.setSnapshot(caller);const frozenDom=p.view.svg.outerHTML;caller.glyphs[0].x=999;
   if(p.view.svg.outerHTML!==frozenDom)throw Error('Caller data alias leaked into view');
   return {staleLayout,staleOverlay,rejectedBad,blinkDidNotRebuild:true};})()`);
 assert.ok(overlays.staleLayout&&overlays.staleOverlay&&overlays.rejectedBad);
 const raster=await evaluate(`(async()=>{const p=window.probe;p.view.setZoom(1);document.getElementById('mount').style.marginTop='60px';window.scrollTo(0,0);
   p.call({op:'replaceSource',handle:p.handle,expectedRevision:1,source:${JSON.stringify(String.raw`\frac{x_1+\text{速度}}{y^2}`)}});
   const r=p.call({op:'layout',handle:p.handle,expectedRevision:2,options:p.opts,includeSvg:true});p.view.setSnapshot(r.snapshot);
   const clone=p.view.svg.cloneNode(true);clone.querySelector('.vt-core-overlay').remove();
   const width=Math.ceil(r.snapshot.width),height=Math.ceil(r.snapshot.height);
   async function pixels(svg){const image=new Image();image.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);await image.decode();
     const canvas=document.createElement('canvas');canvas.width=width;canvas.height=height;const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0);return ctx.getImageData(0,0,width,height).data;}
   const a=await pixels(r.svg),b=await pixels(new XMLSerializer().serializeToString(clone));let sum=0,max=0,bigPixels=0;
   for(let i=0;i<a.length;i+=4){let biggest=0;for(let c=0;c<4;c++){const d=Math.abs(a[i+c]-b[i+c]);sum+=d;max=Math.max(max,d);biggest=Math.max(biggest,d);}if(biggest>10)bigPixels++;}
   const caret=r.snapshot.carets.find(c=>c.position.kind==='text'&&c.position.grapheme===1);
   const geometry=p.call({op:'queryScene',handle:p.handle,expectedRevision:2,sceneToken:r.sceneToken,query:{kind:'selection',selection:{kind:'range',anchor:caret.position,focus:caret.position}}}).result.caretGeometry;
   p.view.setOverlay({...p.view.snapshotIdentity,caret:caret.position,caretGeometry:geometry,rectangles:[]});p.view.setZoom(2);
   return {width,height,meanAbsoluteChannelDifference:sum/a.length,maxChannelDifference:max,pixelsAbove10:bigPixels/(width*height)};
 })()`);
 assert.ok(raster.meanAbsoluteChannelDifference<0.5,JSON.stringify(raster));
 assert.ok(raster.pixelsAbove10<0.02,JSON.stringify(raster));
 await send('Emulation.setDeviceMetricsOverride',{width:1100,height:750,deviceScaleFactor:1,mobile:false});
 const png=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:false});
 writeFileSync(resolve(output,'autonomous-core-view.png'),Buffer.from(png.data,'base64'));
 await evaluate(`(()=>{const p=window.probe;p.call({op:'destroy',handle:p.handle,expectedRevision:2});p.view.dispose();p.rt.free();return true;})()`);
 return {status:'PASS',browser:browserVersion.product,actualBrowserPointerEvents:true,mappingCases:mapping.length,
   mapping,overlays,raster,wasmSha256:sha(wasm),mathFontSha256:sha(math),textFontSha256:sha(text),
   networkRequests:Object.fromEntries(requests),screenshot:'target/svg-review/autonomous-core-view.png',
   notCovered:['keyboard editing bridge','physical OS IME','Word Ribbon/UI','production default engine switch']};
});
writeFileSync(resolve(output,'report.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify(report,null,2));

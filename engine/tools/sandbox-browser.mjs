// Tests the actual sandbox entrypoint and its real controls, not private edit callbacks.
import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { withOwnedBrowser } from './owned-browser-review.mjs';
import { createSandboxAssets, sha256 } from './sandbox-assets.mjs';
const { assets, html, manifest, root } = createSandboxAssets();
const out = resolve(root,'target/sandbox-browser');
mkdirSync(out,{recursive:true});
const steps=[];
const result=await withOwnedBrowser(assets,html,async({send,evaluate,browserVersion,requests})=>{
  const inspect=()=>evaluate('window.__engineSandbox.inspect()');
  const phase=async name=>{
    const state=await inspect();
    const source=await evaluate('document.getElementById("source").value');
    assert.equal(source,state.document.source,`${name}: source pane diverged from authoritative session`);
    steps.push({name,source,revision:state.document.revision,undoDepth:state.undoDepth,redoDepth:state.redoDepth});
    return state;
  };
  async function click(selector){
    const point=await evaluate(`(()=>{const element=document.querySelector(${JSON.stringify(selector)});element.scrollIntoView({block:'center'});const r=element.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2};})()`);
    await send('Input.dispatchMouseEvent',{type:'mousePressed',...point,button:'left',clickCount:1});
    await send('Input.dispatchMouseEvent',{type:'mouseReleased',...point,button:'left',clickCount:1});
  }
  async function key(key,code,virtual,modifiers=0){
    for(const type of ['keyDown','keyUp']) await send('Input.dispatchKeyEvent',{type,key,code,windowsVirtualKeyCode:virtual,nativeVirtualKeyCode:virtual,modifiers});
  }
  async function source(value){await click('#source');await key('a','KeyA',65,2);await send('Input.insertText',{text:value});}
  try {
    await evaluate('window.__engineSandbox.ready.then(()=>true)');
    assert.equal(await evaluate('document.body.dataset.ready'),'true');
    assert.equal((await phase('blank sandbox')).document.source,'');
    const styles=await evaluate('({background:getComputedStyle(document.body).backgroundColor,sections:document.querySelectorAll("main section").length})');
    assert.equal(styles.sections,3);
    await click('[data-structure="fraction"]');await send('Input.insertText',{text:'x'});
    await click('[data-structure="subscript"]');await send('Input.insertText',{text:'1'});
    await key('ArrowRight','ArrowRight',39);await send('Input.insertText',{text:'+'});
    await click('[data-structure="text"]');
    const beforeIme=await phase('native control owns empty text slot');
    await send('Input.imeSetComposition',{text:'su',selectionStart:2,selectionEnd:2});
    assert.equal((await inspect()).document.source,beforeIme.document.source);
    await send('Input.imeSetComposition',{text:'速度',selectionStart:2,selectionEnd:2});
    await send('Input.insertText',{text:'速度'});
    await key('ArrowRight','ArrowRight',39);await key('ArrowDown','ArrowDown',40);
    await send('Input.insertText',{text:'y'});await click('[data-structure="superscript"]');await send('Input.insertText',{text:'2'});
    const built=await phase('formula constructed via sandbox controls');
    assert.equal(built.document.canonicalLatex,String.raw`\frac{x_{1}+\text{速度}}{y^{2}}`);
    const beforeExport=await evaluate('window.__engineSandbox.stats()');
    await click('#export-svg');
    const exported=await evaluate('document.getElementById("export-result").textContent');
    assert.ok(exported.includes('<svg ')&&!exported.includes('<text'));
    const afterExport=await evaluate('window.__engineSandbox.stats()');
    assert.equal(beforeExport.sceneBuilds,afterExport.sceneBuilds);
    assert.equal((await inspect()).document.source,built.document.source);

    const valid=String.raw`\frac{a+\text{中文}}{b^2}`;
    await source(valid);await phase('source editor changes same core document');
    const sourceRevision=(await inspect()).document.revision;
    await key('z','KeyZ',90,2);
    assert.equal((await phase('source-pane undo reaches visual history')).document.source,built.document.source);
    await key('y','KeyY',89,2);
    assert.equal((await phase('source-pane redo restores source edit')).document.source,valid);
    assert.ok((await inspect()).document.revision>sourceRevision);

    const draft=String.raw`\frac{a+\text{中文}}{`;
    await source(draft);const invalid=await phase('invalid draft retained');
    assert.equal(invalid.document.analysis.complete,false);
    const visual=await evaluate('window.__engineSandbox.layout()');
    assert.ok(visual===null||visual.revision===invalid.document.revision,'Old formula geometry may not masquerade as the current draft');
    assert.equal(await evaluate('document.getElementById("export-svg").disabled'),true);
    assert.equal(await evaluate('document.getElementById("source").disabled'),false);
    await source(valid);assert.equal((await phase('repair draft')).document.source,valid);
    await key('z','KeyZ',90,2);assert.equal((await phase('undo repair returns the intact draft')).document.source,draft);
    await key('z','KeyZ',90,2);assert.equal((await phase('undo draft restores last valid scene')).document.source,valid);
    await key('y','KeyY',89,2);assert.equal((await phase('redo draft does not leave source pane stale')).document.source,draft);
    await source(valid);await phase('final repair');

    const current=await inspect();
    await click('#zoom');await key('End','End',35);await key('Enter','Enter',13);
    const zoomed=await evaluate('window.__engineSandbox.stats()');assert.equal(zoomed.zoom,2);
    assert.equal((await inspect()).document.source,current.document.source);
    await click('#focus-editor');
    assert.equal(await evaluate('document.activeElement.classList.contains("vt-core-input")'),true);
    const stable=await evaluate('window.__engineSandbox.stats()');
    await new Promise(ok=>setTimeout(ok,1150));
    const blink=await evaluate('window.__engineSandbox.stats()');assert.equal(blink.sceneBuilds,stable.sceneBuilds);assert.equal(blink.overlayBuilds,stable.overlayBuilds);
    await click('#export-svg');await phase('final export');
    const screenshot=await send('Page.captureScreenshot',{format:'png',captureBeyondViewport:true});
    writeFileSync(resolve(out,'sandbox.png'),Buffer.from(screenshot.data,'base64'));
    const requestPaths=[...requests.keys()];
    assert.ok(requestPaths.every(path=>!/(mathlive|mathjax|cdn\.)/iu.test(path)));
    return {status:'PASS',browser:browserVersion.product,buildId:manifest.buildId,wasmSha256:manifest.assets[0].sha256,
      fontHashes:manifest.assets.filter(a=>a.kind==='font').map(a=>a.sha256),steps,
      realSandboxEntrypoint:true,coreAndSourceHistoryShared:true,draftsRetained:true,clipboardNotTouched:true,
      physicalWindowsIme:false,compositionSource:'Chromium browser protocol, separate from native-IME acceptance',
      requests:requestPaths,screenshot:'target/sandbox-browser/sandbox.png',exportSha256:sha256(Buffer.from(exported))};
  }catch(error){
    let state;try{state=await evaluate('({state:window.__engineSandbox?.inspect(),layout:window.__engineSandbox?.layout(),errors:window.__engineSandbox?.diagnostics(),source:document.getElementById("source").value})');}catch{}
    writeFileSync(resolve(out,'failure.json'),JSON.stringify({error:String(error),steps,state},null,2));
    throw error;
  }
});
writeFileSync(resolve(out,'report.json'),JSON.stringify(result,null,2));
console.log(JSON.stringify(result,null,2));

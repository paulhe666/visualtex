import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, readdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import assert from "node:assert/strict";
import process from "node:process";

const offset = process.pid % 1000;
const previewPort = 6500 + offset;
const debugPort = 12500 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const profile = join(tmpdir(), `visualtex-p10-core-${process.pid}`);
const downloadDir = join(profile, "downloads");
const candidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = candidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P10 core acceptance.");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function waitFor(url, timeout = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    try {
      const response = await fetch(url);
      if (response.ok) return response;
    } catch {}
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}
class CdpClient {
  constructor(url) { this.url = url; this.id = 1; this.pending = new Map(); }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result);
    });
  }
  send(method, params = {}) {
    const id = this.id++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }
  close() { this.socket?.close(); }
}

async function main() {
  const preview = spawn(process.execPath,["node_modules/vite/bin/vite.js","preview","--host","127.0.0.1","--port",String(previewPort),"--strictPort"],{cwd:process.cwd(),stdio:"ignore"});
  let browser, client;
  try {
    await mkdir(downloadDir,{recursive:true});
    await waitFor(baseUrl);
    browser = spawn(browserPath,["--headless=new","--disable-gpu","--no-first-run","--no-default-browser-check",`--remote-debugging-port=${debugPort}`,`--user-data-dir=${profile}`,"--window-size=1400,1000",baseUrl],{stdio:"ignore"});
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    let page;
    for (let i=0;i<100&&!page;i++) {
      const targets=await(await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
      page=targets.find((target)=>target.type==="page"&&target.url.startsWith(baseUrl));
      if(!page)await sleep(100);
    }
    if(!page)throw new Error("No VisualTeX product page found");
    client=new CdpClient(page.webSocketDebuggerUrl);await client.connect();await client.send("Runtime.enable");await client.send("Page.enable");
    try{await client.send("Browser.setDownloadBehavior",{behavior:"allow",downloadPath:downloadDir,eventsEnabled:true});}catch{await client.send("Page.setDownloadBehavior",{behavior:"allow",downloadPath:downloadDir});}
    const evaluate=async expression=>{const result=await client.send("Runtime.evaluate",{expression,awaitPromise:true,returnByValue:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description??result.exceptionDetails.text);return result.result.value;};
    const waitEval=async(expression,label,timeout=12000)=>{const started=Date.now();while(Date.now()-started<timeout){if(await evaluate(expression))return;await sleep(80);}throw new Error(`Timed out: ${label}`);};
    const key=async(key,code,vk,modifiers=0,text)=>{for(const type of["keyDown","keyUp"])await client.send("Input.dispatchKeyEvent",{type,key,code,windowsVirtualKeyCode:vk,nativeVirtualKeyCode:vk,modifiers,...(type==="keyDown"&&text!==undefined?{text,unmodifiedText:text}:{})});await sleep(70);};
    const click=async selector=>{const point=await evaluate(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)return null;e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2};})()`);if(!point)throw new Error(`Missing ${selector}`);for(const type of["mousePressed","mouseReleased"])await client.send("Input.dispatchMouseEvent",{type,...point,button:"left",clickCount:1});await sleep(100);};
    const clickMatching=async(selector,text,childSelector=null)=>{const point=await evaluate(`(()=>{let e=[...document.querySelectorAll(${JSON.stringify(selector)})].find(node=>(node.textContent??'').includes(${JSON.stringify(text)}));if(!e)return null;${childSelector?`e=e.querySelector(${JSON.stringify(childSelector)});if(!e)return null;`:''}e.scrollIntoView({block:'center'});const r=e.getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2};})()`);if(!point)throw new Error(`Missing ${selector} containing ${text}`);for(const type of["mousePressed","mouseReleased"])await client.send("Input.dispatchMouseEvent",{type,...point,button:"left",clickCount:1});await sleep(100);};
    const sourceText=()=>evaluate(`(()=>{const lines=[...document.querySelectorAll('.cm-line')];return lines.length?lines.map(x=>x.textContent??'').join('\\n'):null;})()`);
    const waitSource=async(predicate,label,timeout=12000)=>{const started=Date.now();while(Date.now()-started<timeout){const value=await sourceText();if(predicate(value))return value;await sleep(80);}throw new Error(`Timed out: ${label}; actual=${JSON.stringify(await sourceText())}`);};
    const storedLine=()=>evaluate(`(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex??null}catch{return null}})()`);
    const waitDownloadedDocument=async(timeout=8000)=>{const started=Date.now();while(Date.now()-started<timeout){const names=(await readdir(downloadDir).catch(()=>[])).filter(name=>name.endsWith('.visualtex.json'));if(names.length){const path=join(downloadDir,names.sort().at(-1));return{path,text:await readFile(path,'utf8')}}await sleep(80);}throw new Error('Timed out waiting for .visualtex.json download');};
    const customLibrary={version:1,symbols:[{id:'p10-user-symbol',command:'userpsymbol',name:'P10 user symbol',role:'ordinary',limitsBehavior:'auto',metrics:{widthEm:.8,ascentEm:.7,descentEm:.1},artwork:{shapes:[{kind:'rect',x:80,y:120,width:640,height:500,fill:true}]},ommlFallback:null,createdAt:1,updatedAt:1}]};
    await waitEval(`location.origin===${JSON.stringify(new URL(baseUrl).origin)}&&document.readyState!=='loading'`,`product origin committed`);
    await evaluate(`localStorage.setItem('visualtex.custom-symbols.v1',${JSON.stringify(JSON.stringify(customLibrary))});true`);
    await client.send('Page.reload',{ignoreCache:true});await sleep(250);

    await waitEval(`Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`,`core editor mount`);
    const initial=await evaluate(`({core:Boolean(document.querySelector('[data-engine="visualtex-core"]')),legacyFields:document.querySelectorAll('math-field').length,wasm:[...performance.getEntriesByType('resource')].some(x=>String(x.name).includes('visualtex_engine_wasm_bg')),font:[...performance.getEntriesByType('resource')].some(x=>String(x.name).includes('latinmodern-math'))})`);
    assert.equal(initial.core,true);assert.equal(initial.legacyFields,0);assert.equal(initial.wasm,true);assert.equal(initial.font,true);
    await waitEval(`Boolean(document.querySelector('.math-preview[data-static-layout="true"][data-preview-renderer="visualtex-core"] svg'))`,`core static preview`);
    const staticPreviewInitial=await evaluate(`(()=>{const previews=[...document.querySelectorAll('.math-preview[data-static-layout="true"][data-preview-renderer="visualtex-core"]')];return{count:previews.length,hasSvg:previews.every(node=>Boolean(node.querySelector('svg'))),hasMathLive:previews.some(node=>Boolean(node.querySelector('.ML__latex'))),hasInput:previews.some(node=>Boolean(node.querySelector('.vt-core-input')))}})()`);assert.ok(staticPreviewInitial.count>0);assert.equal(staticPreviewInitial.hasSvg,true);assert.equal(staticPreviewInitial.hasMathLive,false);assert.equal(staticPreviewInitial.hasInput,false);

    const initialLatex=await storedLine();assert.equal(typeof initialLatex,'string');
    await evaluate(`document.querySelector('.vt-core-input').focus();true`);
    await key('+','Equal',187,0,'+');for(const [char,code,vk] of [['a','KeyA',65],['l','KeyL',76],['p','KeyP',80],['h','KeyH',72],['a','KeyA',65]])await key(char,code,vk,0,char);
    await click('[data-classic-bottom-view="source"]');
    await waitEval(`Boolean(document.querySelector('.cm-editor'))`,`source editor`);
    const typedSource=await waitSource(value=>Boolean(value?.includes(String.raw`\alpha`)),'alpha inline shortcut');
    const alphaLatex=await storedLine();assert.equal(alphaLatex,`${initialLatex}+\\alpha`);assert.ok(!alphaLatex.endsWith('+alpha'));
    const coreButtons=await evaluate(`(()=>{const b=[...document.querySelectorAll('.edit-actions button')];return{undoDisabled:b[0]?.disabled??null,redoDisabled:b[1]?.disabled??null};})()`);assert.deepEqual(coreButtons,{undoDisabled:false,redoDisabled:true});
    await click('.edit-actions button:nth-child(1)');const buttonUndoSource=await waitSource(value=>value!==typedSource,'core header undo');assert.equal(await storedLine(),initialLatex,'inline shortcut did not merge into one undo item');
    const afterUndoButtons=await evaluate(`(()=>{const b=[...document.querySelectorAll('.edit-actions button')];return{undoDisabled:b[0]?.disabled??null,redoDisabled:b[1]?.disabled??null};})()`);assert.equal(afterUndoButtons.redoDisabled,false);
    await click('.edit-actions button:nth-child(2)');await waitSource(value=>value===typedSource,'core header redo');assert.equal(await storedLine(),alphaLatex);

    await click('[data-classic-bottom-view="tools"]');await evaluate(`document.querySelector('.vt-core-input').focus();true`);
    await key('x','KeyX',88,0,'x');await key('^','Digit6',54,0,'^');await key('2','Digit2',50,0,'2');await key('y','KeyY',89,0,'y');
    await click('[data-classic-bottom-view="source"]');const autoExitLatex=await storedLine();assert.ok(autoExitLatex.endsWith('x^{2}y'),autoExitLatex);const autoExitFormattedSource=await sourceText();

    await click('[data-classic-bottom-view="tools"]');await evaluate(`document.querySelector('[data-engine="visualtex-core"]').dataset.settingsOwner='stable';true`);await click('.canvas-input-behavior-trigger');await waitEval(`Boolean(document.querySelector('.input-behavior-popover'))`,`input behavior popover`);await clickMatching('.input-behavior-option','上标输入后跳出','input[type="checkbox"]');assert.equal(await evaluate(`Boolean(document.querySelector('[data-engine="visualtex-core"][data-settings-owner="stable"]'))`),true);
    await click('.canvas-input-behavior-trigger');await evaluate(`document.querySelector('.vt-core-input').focus();true`);for(const [char,code,vk] of [['+','Equal',187],['x','KeyX',88],['^','Digit6',54],['2','Digit2',50],['y','KeyY',89]])await key(char,code,vk,0,char);const autoExitDisabledLatex=await storedLine();assert.ok(autoExitDisabledLatex.endsWith('+x^{2y}'),autoExitDisabledLatex);
    await click('.canvas-input-behavior-trigger');await waitEval(`Boolean(document.querySelector('.input-behavior-popover'))`,`input behavior popover reopen`);await clickMatching('.input-behavior-option','上标输入后跳出','input[type="checkbox"]');await click('.canvas-input-behavior-trigger');assert.equal(await evaluate(`Boolean(document.querySelector('[data-engine="visualtex-core"][data-settings-owner="stable"]'))`),true);await evaluate(`document.querySelector('.vt-core-input').focus();true`);await key('ArrowRight','ArrowRight',39);for(const [char,code,vk] of [['+','Equal',187],['x','KeyX',88],['^','Digit6',54],['2','Digit2',50],['y','KeyY',89]])await key(char,code,vk,0,char);const autoExitReenabledLatex=await storedLine();assert.ok(autoExitReenabledLatex.endsWith('+x^{2}y'),autoExitReenabledLatex);

    await click('.settings-toggle');await waitEval(`Boolean(document.querySelector('.settings-dialog'))`,`settings dialog`);await click('[data-interface-customization-trigger]');await waitEval(`Boolean(document.querySelector('[data-interface-customization-dialog]'))`,`interface customization dialog`);await waitEval(`Boolean(document.querySelector('[data-formula-letter-font-setting]'))`,`settings font selector`);await waitEval(`document.querySelectorAll('.math-preview[data-static-layout="true"][data-preview-renderer="visualtex-core"]').length>0`,`page core static previews before font change`);await click('[data-formula-letter-font-setting]');await key('ArrowDown','ArrowDown',40);await key('Enter','Enter',13);await sleep(180);assert.equal(await evaluate(`Boolean(document.querySelector('[data-engine="visualtex-core"][data-settings-owner="stable"]'))`),true);assert.equal(await storedLine(),autoExitReenabledLatex);await waitEval(`Boolean(document.querySelector('.core-math-editor-notice'))`,`unresolved browser font notice`);await waitEval(`document.querySelectorAll('.math-preview[data-static-layout="true"][data-preview-renderer="visualtex-core"]').length===0`,`static previews retire old verified-font SVG`);await click('[data-formula-letter-font-setting]');await key('ArrowUp','ArrowUp',38);await key('Enter','Enter',13);await sleep(180);await waitEval(`!document.querySelector('.core-math-editor-notice')`,`font notice clears after verified font restore`);await waitEval(`document.querySelectorAll('.math-preview[data-static-layout="true"][data-preview-renderer="visualtex-core"]').length>0`,`static previews restore core renderer`);assert.equal(await evaluate(`Boolean(document.querySelector('[data-engine="visualtex-core"][data-settings-owner="stable"]'))`),true);await click('[data-interface-customization-close]');await waitEval(`!document.querySelector('[data-interface-customization-dialog]')`,`interface customization closed`);await click('.settings-dialog .dialog-header .icon-button');await waitEval(`!document.querySelector('.settings-dialog')`,`settings dialog closed`);await click('[data-classic-bottom-view="source"]');await waitEval(`Boolean(document.querySelector('.cm-content'))`,`source editor after settings`);const sourceBeforeFracEdit=await sourceText();

    const fracLatex=String.raw`\frac{a}{b}`,fracSource=`$$\n${fracLatex}\n$$`;

    await evaluate(`document.querySelector('.cm-content').focus();true`);
    await key('a','KeyA',65,2);await client.send('Input.insertText',{text:fracSource});await sleep(180);
    const formattedFracSource=await waitSource(value=>Boolean(value?.includes(fracLatex)&&value.startsWith('$$\n')&&value.endsWith('\n$$')),'CodeMirror to core source');
    await click('[data-classic-bottom-view="tools"]');await click('[data-classic-bottom-view="source"]');await waitEval(`Boolean(document.querySelector('.cm-editor'))`,`source remount after source edit`);assert.equal(await sourceText(),formattedFracSource,'source edit did not reach the core/store owner');
    await evaluate(`document.querySelector('.cm-content').focus();true`);
    await key('z','KeyZ',90,2);await waitSource(value=>value===sourceBeforeFracEdit,'core-owned source undo');
    await key('y','KeyY',89,2);await waitSource(value=>value===formattedFracSource,'core-owned source redo');

    await click('[data-classic-bottom-view="tools"]');
    await evaluate(`document.querySelector('.vt-core-input').focus();true`);
    await key('\\','Backslash',220,0,'\\');for(const [char,code,vk] of [['f','KeyF',70],['r','KeyR',82],['a','KeyA',65],['c','KeyC',67]])await key(char,code,vk,0,char);
    const commandPopup=await evaluate(`({visible:!document.querySelector('.vt-core-command-popup')?.hidden,selected:document.querySelector('.vt-core-command-candidate[aria-selected="true"]')?.dataset.commandId??null,count:document.querySelectorAll('.vt-core-command-candidate').length})`);assert.equal(commandPopup.visible,true);assert.equal(commandPopup.selected,'frac');assert.ok(commandPopup.count<=6);
    await key('Enter','Enter',13);await sleep(120);assert.ok((await storedLine())?.includes(String.raw`\frac`));
    const fracUsage=await evaluate(`(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.usage?.frac?.useCount??0}catch{return 0}})()`);assert.ok(fracUsage>0);

    await evaluate(`document.querySelector('.vt-core-input').focus();true`);await key('End','End',35);
    await key('\\','Backslash',220,0,'\\');for(const [char,code,vk] of [['u','KeyU',85],['s','KeyS',83],['e','KeyE',69],['r','KeyR',82],['p','KeyP',80],['s','KeyS',83],['y','KeyY',89],['m','KeyM',77],['b','KeyB',66],['o','KeyO',79],['l','KeyL',76]])await key(char,code,vk,0,char);
    const customCandidate=await evaluate(`document.querySelector('.vt-core-command-candidate[aria-selected="true"]')?.dataset.commandId??null`);assert.equal(customCandidate,'custom-symbol:p10-user-symbol');await key('Enter','Enter',13);await sleep(120);
    assert.ok((await storedLine())?.includes(String.raw`\userpsymbol`));const customArtwork=await evaluate(`document.querySelectorAll('[data-engine="visualtex-core"] path[data-asset-sha256]').length`);assert.ok(customArtwork>0);

    await click('[data-classic-bottom-view="source"]');await waitSource(value=>Boolean(value?.includes(String.raw`\userpsymbol`)),'custom symbol source projection');await click('[data-classic-bottom-view="tools"]');
    await evaluate(`document.querySelector('.vt-core-input').focus();true`);
    await evaluate(`document.querySelector('.vt-core-input').focus();true`);
    const beforeToolbar=await evaluate(`document.querySelector('[data-engine="visualtex-core"] svg')?.outerHTML.length??0`);
    await click('button[data-command-id="frac"]');
    await sleep(180);
    await click('[data-classic-bottom-view="source"]');
    await waitSource(value=>Boolean(value?.includes(String.raw`\frac`)),'toolbar structural fraction reached source');
    const afterToolbar=await evaluate(`document.querySelector('[data-engine="visualtex-core"] svg')?.outerHTML.length??0`);
    assert.ok(afterToolbar>0&&afterToolbar!==beforeToolbar);

    const savedCoreLatex=await storedLine();assert.equal(typeof savedCoreLatex,'string');
    await click('.file-actions button:nth-child(3)');const downloaded=await waitDownloadedDocument();const savedDocument=JSON.parse(downloaded.text);assert.equal(savedDocument.formulas?.[0]?.latex,savedCoreLatex);assert.ok(Array.isArray(savedDocument.formulas)&&savedDocument.formulas.length===1);
    const openedLatex=String.raw`\frac{m}{n}+\sqrt{q}`;savedDocument.title='P10 Opened Core Document';savedDocument.formulas=[{...savedDocument.formulas[0],id:'p10-opened-core',latex:openedLatex,displayMode:'block',updatedAt:Date.now()}];
    await click('[data-classic-bottom-view="tools"]');await evaluate(`document.querySelector('[data-engine="visualtex-core"]').dataset.fileOwner='before-open';true`);
    const openPayload=JSON.stringify(savedDocument);await evaluate(`(()=>{const input=document.querySelector('input[type="file"].visually-hidden');if(!input)return false;const transfer=new DataTransfer();transfer.items.add(new File([${JSON.stringify(openPayload)}],'p10-open.visualtex.json',{type:'application/json'}));input.files=transfer.files;input.dispatchEvent(new Event('change',{bubbles:true}));return true;})()`);
    await waitEval(`${JSON.stringify(openedLatex)}===(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex??null}catch{return null}})()&&Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))&&!document.querySelector('[data-engine="visualtex-core"][data-file-owner]')`,`open document remounts selected core owner`);assert.equal(await evaluate(`document.querySelectorAll('math-field').length`),0);

    const beforeDocumentBoundary=await storedLine();assert.equal(beforeDocumentBoundary,openedLatex);
    await click('[data-classic-bottom-view="tools"]');await waitEval(`Boolean(document.querySelector('[data-engine="visualtex-core"]'))`,`core before new document`);
    await evaluate(`document.querySelector('[data-engine="visualtex-core"]').dataset.sessionMarker='old-owner';true`);
    await click('.file-actions button:nth-child(1)');
    await waitEval(`Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))&&!document.querySelector('[data-engine="visualtex-core"][data-session-marker]')`,`new blank core session`);
    assert.equal(await storedLine(),'');const blankInputEnabled=await evaluate(`!document.querySelector('.vt-core-input')?.disabled`);assert.equal(blankInputEnabled,true);
    const boundaryButtons=await evaluate(`(()=>{const b=[...document.querySelectorAll('.edit-actions button')];return{undoDisabled:b[0]?.disabled??null,redoDisabled:b[1]?.disabled??null};})()`);assert.equal(boundaryButtons.undoDisabled,false);
    await evaluate(`document.querySelector('[data-engine="visualtex-core"]').dataset.sessionMarker='blank-owner';true`);await click('.edit-actions button:nth-child(1)');
    await waitEval(`${JSON.stringify(beforeDocumentBoundary)}===(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex??null}catch{return null}})()&&Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))&&!document.querySelector('[data-engine="visualtex-core"][data-session-marker]')`,`document undo remount`);
    const restoredBoundary=await storedLine();assert.equal(restoredBoundary,beforeDocumentBoundary);
    const afterDocumentUndoButtons=await evaluate(`(()=>{const b=[...document.querySelectorAll('.edit-actions button')];return{undoDisabled:b[0]?.disabled??null,redoDisabled:b[1]?.disabled??null};})()`);assert.equal(afterDocumentUndoButtons.redoDisabled,false);
    await click('.edit-actions button:nth-child(2)');await waitEval(`(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex==='' }catch{return false}})()&&Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`,`document redo blank`);
    await click('.edit-actions button:nth-child(1)');await waitEval(`${JSON.stringify(beforeDocumentBoundary)}===(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex??null}catch{return null}})()&&Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`,`document undo restore second time`);

    // A later unsupported verified-font case must stay on the already-selected
    // core session rather than silently replacing it with MathLive.
    const unsupportedLatex=String.raw`\text{中文}`,unsupportedSource=`$$\n${unsupportedLatex}\n$$`;
    await click('[data-classic-bottom-view="source"]');await waitEval(`Boolean(document.querySelector('.cm-content'))`,`source panel before unsupported-font edit`);
    await evaluate(`document.querySelector('.cm-content').focus();true`);
    await key('a','KeyA',65,2);await client.send('Input.insertText',{text:unsupportedSource});await sleep(220);
    const locked=await evaluate(`({core:Boolean(document.querySelector('[data-engine="visualtex-core"]')),legacyFields:document.querySelectorAll('math-field').length,source:[...document.querySelectorAll('.cm-line')].map(x=>x.textContent??'').join('\\n'),viewDisabled:document.querySelector('.vt-core-scene')?.getAttribute('aria-disabled')})`);
    assert.equal(locked.core,true);assert.equal(locked.legacyFields,0);assert.ok(locked.source?.includes(unsupportedLatex));assert.equal(locked.viewDisabled,'true');

    // P15 removes the editable-engine fallback. Reopening unsupported/unresolved
    // source must keep the exact source in a core-owned session; the visual
    // scene may fail closed until the source/font issue is repaired.
    await client.send('Page.reload',{ignoreCache:true});await sleep(250);
    await waitEval(`Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))&&document.querySelectorAll('math-field').length===0`,`initial unsupported source remains core-owned`);
    const reopened=await evaluate(`({core:Boolean(document.querySelector('[data-engine="visualtex-core"]')),legacyFields:document.querySelectorAll('math-field').length,source:(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex??null}catch{return null}})(),viewDisabled:document.querySelector('.vt-core-scene')?.getAttribute('aria-disabled')})`);
    assert.equal(reopened.core,true);assert.equal(reopened.legacyFields,0);assert.equal(reopened.source,unsupportedLatex);assert.equal(reopened.viewDisabled,'true');
    await click('[data-classic-bottom-view="source"]');await waitEval(`Boolean(document.querySelector('.cm-content'))`,`source repair panel`);
    const repairSource=`$$\nx+1\n$$`;
    await evaluate(`document.querySelector('.cm-content').focus();true`);await key('a','KeyA',65,2);await client.send('Input.insertText',{text:repairSource});await sleep(220);
    await waitEval(`(()=>{try{return JSON.parse(localStorage.getItem('visualtex-editor')??'null')?.state?.lines?.[0]?.latex==='x+1'}catch{return false}})()`,`unsupported source repaired in core`);
    await click('[data-classic-bottom-view="tools"]');await waitEval(`document.querySelector('.vt-core-scene')?.getAttribute('aria-disabled')!=='true'`,`core visual scene restored after source repair`);

    console.log(JSON.stringify({status:'PASS',actualProductPage:true,actualBrowserKeyboard:true,coreResourcesVerified:true,legacyMathfieldAbsent:true,sourceProjection:true,coreOwnedCodeMirrorHistory:true,coreAwareHeaderUndoRedo:true,documentHistoryBoundaryArbitration:true,blankCoreSessionEditable:true,inlineShortcutSingleUndo:true,defaultSuperscriptAutoExit:true,productCommandCandidates:true,personalizedUsageWriteback:true,frozenCustomSymbolInsertion:true,coreStaticPreview:true,staticPreviewSessionLeakGuard:true,staticPreviewFontIdentityRetirement:true,toolbarStructuralAction:true,sessionRemainsCoreOnLateUnsupportedFont:true,unsupportedSourceRemainsCoreAcrossReload:true,sourceRepairRestoresCoreScene:true,initial,staticPreviewInitial,coreButtons,afterUndoButtons,buttonUndoSource,commandPopup,customArtwork,boundaryButtons,afterDocumentUndoButtons,locked,reopened},null,2));
  } finally {
    client?.close();
    browser?.kill();
    preview.kill();
    await rm(profile,{recursive:true,force:true}).catch(()=>{});
  }
}
await main();

// P16 post-install visual-regression acceptance for the core-only Office editor.
// This is an owned Chromium surface only: no Word/PowerPoint automation or install mutation.
import assert from "node:assert/strict";
import { mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from "node:fs";
import { dirname, extname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { withOwnedBrowser } from "../../../engine/tools/owned-browser-review.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const appRoot = resolve(here, "..");
const repoRoot = resolve(appRoot, "../..");
const dist = resolve(appRoot, "dist-office-windows-native");
const stamp = new Date().toISOString().replaceAll(":", "-").replaceAll(".", "-");
const out = resolve(repoRoot, "engine/target/p16-core-visual-regression-browser", stamp);
mkdirSync(out, { recursive: true });

const mime = (path) => ({
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".wasm": "application/wasm",
  ".otf": "font/otf",
  ".woff2": "font/woff2",
}[extname(path)] ?? "application/octet-stream");
const assets = new Map();
function collect(directory, prefix = "") {
  for (const name of readdirSync(directory)) {
    const file = join(directory, name);
    const rel = prefix ? `${prefix}/${name}` : name;
    if (statSync(file).isDirectory()) collect(file, rel);
    else if (rel !== "dialog/index.html") assets.set(`/${rel.replaceAll("\\", "/")}`, { type: mime(file), bytes: readFileSync(file) });
  }
}
collect(dist);

const now = Date.now();
const mock = String.raw`(()=>{
const realFetch=window.fetch.bind(window);const now=${now};
const session={id:'p16visual',mode:'edit',host:'word',formulaId:'p16visual-formula',sourceDocumentId:'owned-p16visual-document',sourceObjectId:'owned-p16visual-object',title:'P16 visual regression',lines:[{id:'line-1',latex:'x'}],activeLineId:'line-1',codeFormat:'raw',displayMode:'inline',objectMode:'nativeOle',numbered:false,mathTypeNumberPosition:'right',fontSizePt:10.5,exportWidth:0,exportHeight:0,exportResult:null,originalMetadata:null,dirty:false,status:'editing',autoCommitOnClose:true,explicitCancel:false,error:null,createdAt:now,updatedAt:now,expiresAt:now+600000};
const state=window.__p16VisualMock={session,patches:[],closeCalls:0};
const json=(value,status=200)=>new Response(JSON.stringify(value),{status,headers:{'Content-Type':'application/json'}});
window.fetch=async(input,init={})=>{const url=new URL(typeof input==='string'?input:input.url,location.href);const method=(init.method||((typeof input==='object'&&input.method)||'GET')).toUpperCase();
 if(!url.pathname.startsWith('/api/v1/'))return realFetch(input,init);
 if(url.pathname==='/api/v1/preferences'&&method==='GET')return json({powerpointDefaultFontSizePt:20,editorPreferences:{settings:{theme:'light',editorLayout:'standard',language:'cn',zoom:.6,formulaLetterFont:'katex',formulaChineseFont:'system',suggestionCount:6,inputBehavior:{showStructuredCommandSuggestions:true,showOtherCommandSuggestions:true}}}});
 if(url.pathname==='/api/v1/theme'&&method==='GET')return json({theme:'light',editorLayout:'standard'});
 if(url.pathname==='/api/v1/app/converter/next-batch'&&method==='GET')return json({sessionIds:[]});
 if(url.pathname==='/api/v1/sessions/p16visual'&&method==='GET')return json(state.session);
 if(url.pathname==='/api/v1/sessions/p16visual'&&method==='PATCH'){const body=JSON.parse(String(init.body||'{}'));state.patches.push(structuredClone(body));state.session={...state.session,...body,updatedAt:Date.now()};return json(state.session);}
 if(url.pathname==='/api/v1/app/sessions/p16visual/close'&&method==='POST'){state.closeCalls++;return new Response(null,{status:204});}
 return json({error:'P16 visual mock has no route'},404);
};})();`;
assets.set("/p16-visual-mock.js", { type: "text/javascript; charset=utf-8", bytes: Buffer.from(mock) });
const originalHtml = readFileSync(resolve(dist, "dialog/index.html"), "utf8");
const html = originalHtml.replace(
  '<script type="module"',
  '<script src="/p16-visual-mock.js"></script>\n    <script type="module"',
);

async function waitFor(evaluate, expression, timeout = 15_000) {
  const start = Date.now();
  let last = null;
  while (Date.now() - start < timeout) {
    try { last = await evaluate(expression); if (last) return last; } catch (error) { last = String(error); }
    await new Promise((resolveWait) => setTimeout(resolveWait, 80));
  }
  throw new Error(`P16 visual browser wait timed out: ${expression}; last=${String(last)}`);
}

const steps = [];
let report;
try {
  report = await withOwnedBrowser(assets, html, async ({ send, evaluate, browserVersion, origin }) => {
    const key = async (keyName, code, vk, text) => {
      await send("Input.dispatchKeyEvent", { type: "keyDown", key: keyName, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, ...(text !== undefined ? { text } : {}) });
      await send("Input.dispatchKeyEvent", { type: "keyUp", key: keyName, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk });
    };
    const typeCommand = async (name) => {
      await key("\\", "Backslash", 220, "\\");
      for (const ch of name) {
        const upper = ch.toUpperCase();
        await key(ch, `Key${upper}`, upper.charCodeAt(0), ch);
      }
    };

    await send("Page.navigate", { url: `${origin}/?runtime=vsto-desktop&sessionId=p16visual` });
    await waitFor(evaluate, `Boolean(document.querySelector('[data-engine="visualtex-core"] .vt-core-input'))`);
    assert.equal(await evaluate(`document.querySelectorAll('math-field').length`), 0);
    steps.push("core-only interactive editor mounted with no MathLive field");

    await evaluate(`document.querySelector('.vt-core-input').focus();true`);
    await typeCommand("fra");
    await waitFor(evaluate, `Boolean(document.querySelector('.suggestion-popup .suggestion-item.is-selected'))`);
    const candidate = await evaluate(`(()=>{const popup=document.querySelector('.suggestion-popup');const selected=popup?.querySelector('.suggestion-item.is-selected');const preview=selected?.querySelector('.suggestion-preview');return popup&&selected?{command:selected.querySelector('.suggestion-command')?.textContent??'',label:selected.querySelector('.suggestion-label')?.textContent??'',previewRenderer:preview?.querySelector('[data-preview-renderer]')?.getAttribute('data-preview-renderer')??null,previewSvg:Boolean(preview?.querySelector('svg')),primitiveHidden:document.querySelector('.vt-core-command-popup')?.hidden===true}:null;})()`);
    assert.ok(candidate);
    assert.match(candidate.command, /\\frac/u);
    assert.equal(candidate.previewRenderer, "visualtex-core");
    assert.equal(candidate.previewSvg, true);
    assert.equal(candidate.primitiveHidden, true);
    steps.push("command candidate restores product command text, label and core-rendered mathematical preview");

    await key("Enter", "Enter", 13);
    await waitFor(evaluate, `window.__p16VisualMock.patches.some(p=>String(p.lines?.[0]?.latex??'').includes('\\\\frac')) || !document.querySelector('.suggestion-popup')`);
    const structural = await evaluate(`(()=>({slots:document.querySelectorAll('.core-formula-line .vt-core-empty-slot').length,popup:Boolean(document.querySelector('.suggestion-popup')),source:window.__p16VisualMock.patches.map(p=>p.lines?.[0]?.latex).filter(Boolean).at(-1)??window.__p16VisualMock.session.lines[0].latex,roles:[...document.querySelectorAll('.core-formula-line [data-node-id]')].map(e=>e.getAttribute('class')).filter(Boolean).slice(0,40),svg:document.querySelector('.core-formula-line .vt-core-svg')?.outerHTML.slice(0,1800)??'',slotStyles:[...document.querySelectorAll('.core-formula-line .vt-core-empty-slot')].slice(0,2).map(e=>{const s=getComputedStyle(e);return{stroke:s.stroke,fill:s.fill,opacity:s.opacity}})}))()`);
    assert.ok(structural.slots >= 2, `Expected visible structural slots after \\frac confirmation: ${JSON.stringify(structural)}`);
    assert.equal(structural.popup, false);
    assert.match(structural.source, /\\frac/u);
    assert.ok(structural.slotStyles.every((style) => style.stroke && style.stroke !== "none"));
    steps.push("structural template confirmation renders visible interactive numerator/denominator slots instead of blank geometry");

    await typeCommand("sqrt");
    await waitFor(evaluate, `Boolean(document.querySelector('.suggestion-popup .suggestion-preview svg'))`);
    await key("Escape", "Escape", 27);
    assert.equal(await evaluate(`Boolean(document.querySelector('.suggestion-popup'))`), false);
    steps.push("rich candidate surface cancels cleanly without leaving stale popup DOM");

    const audit = await evaluate(`({coreEditors:document.querySelectorAll('[data-engine="visualtex-core"]').length,coreInputs:document.querySelectorAll('.vt-core-input').length,mathFields:document.querySelectorAll('math-field').length,richPopups:document.querySelectorAll('.suggestion-popup').length,primitiveVisible:[...document.querySelectorAll('.vt-core-command-popup')].filter(e=>!e.hidden).length,emptySlots:document.querySelectorAll('.vt-core-empty-slot').length})`);
    assert.equal(audit.mathFields, 0);
    assert.equal(audit.primitiveVisible, 0);
    const screenshot = await send("Page.captureScreenshot", { format: "png", captureBeyondViewport: true });
    const screenshotPath = resolve(out, "visual-command-structure.png");
    writeFileSync(screenshotPath, Buffer.from(screenshot.data, "base64"));
    return {
      status: "PASS",
      browser: browserVersion.product,
      checks: steps,
      candidate,
      structural,
      audit,
      screenshot: relative(repoRoot, screenshotPath).replaceAll("\\", "/"),
      officeAutomation: 0,
      installedPackageMutation: 0,
    };
  });
} catch (error) {
  report = { status: "FAILED", error: String(error), checks: steps };
  throw error;
} finally {
  if (report) {
    const receipt = resolve(out, "report.json");
    writeFileSync(receipt, JSON.stringify(report, null, 2));
    console.log(JSON.stringify({ receipt: relative(repoRoot, receipt).replaceAll("\\", "/"), ...report }, null, 2));
  }
}

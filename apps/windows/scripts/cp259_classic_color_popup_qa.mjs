import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { readFile, writeFile, mkdir, rm } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import process from "node:process";
import {
  createBrowserProfilePath,
  resolveChromiumExecutable,
} from "./browser_test_runtime.mjs";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const offset = process.pid % 500;
const port = 22400 + offset;
const debugPort = 23400 + offset;
const baseUrl = `http://127.0.0.1:${port}`;
const sessionId = "11111111-2222-4333-8444-555555555555";
const officeUrl = baseUrl + "/";
const stage = "after-r2";
const output = join(process.cwd(), "../../artifacts/conversion-symbol-audit-20261002/cp259-classic-color-popup");
const distRoot = join(process.cwd(), "../../artifacts/conversion-symbol-audit-20261002/cp258-color-popup/after-r2-desktop");
let editorLayout = "classic";
const chromeProfile = createBrowserProfilePath("visualtex-office-layout");
const chromePath = resolveChromiumExecutable();

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".woff2": "font/woff2",
};

let session = {
  id: sessionId,
  mode: "create",
  host: "word",
  formulaId: "aaaaaaaa-bbbb-4ccc-8ddd-eeeeeeeeeeee",
  sourceDocumentId: "office-layout-regression",
  sourceObjectId: "visualtex-word-vsto-range:0:0",
  title: "Word Formula",
  lines: [
    {
      id: "99999999-8888-4777-8666-555555555555",
      latex: "",
    },
  ],
  activeLineId: "99999999-8888-4777-8666-555555555555",
  codeFormat: "latex",
  displayMode: "inline",
  objectMode: "nativeOle",
  numbered: false,
  mathTypeNumberPosition: "right",
  fontSizePt: 10.5,
  exportWidth: 0,
  exportHeight: 0,
  exportResult: null,
  originalMetadata: null,
  dirty: false,
  status: "editing",
  autoCommitOnClose: true,
  explicitCancel: false,
  error: null,
  createdAt: Date.now(),
  updatedAt: Date.now(),
  expiresAt: Date.now() + 60_000,
};

function writeJson(response, status, value) {
  response.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  response.end(JSON.stringify(value));
}

async function readJsonBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const source = Buffer.concat(chunks).toString("utf8");
  return source ? JSON.parse(source) : {};
}

const server = createServer(async (request, response) => {
  const url = new URL(request.url ?? "/", baseUrl);
  try {
    if (url.pathname === "/api/v1/theme") {
      writeJson(response, 200, { theme: "light", editorLayout });
      return;
    }
    if (url.pathname === "/api/v1/preferences") {
      writeJson(response, 200, {
        powerpointDefaultFontSizePt: 20,
        editorPreferences: {
          settings: {
            editorLayout,
            zoom: 0.5,
            sourceOpen: false,
            classicTileWidth: 300,
            classicDockHeight: 240,
          },
        },
      });
      return;
    }
    if (url.pathname === `/api/v1/sessions/${sessionId}`) {
      if (request.method === "PATCH") {
        session = {
          ...session,
          ...(await readJsonBody(request)),
          updatedAt: Date.now(),
        };
      }
      writeJson(response, 200, session);
      return;
    }
    if (url.pathname === `/api/v1/app/sessions/${sessionId}/close`) {
      writeJson(response, 200, { closed: true });
      return;
    }
    if (url.pathname.startsWith("/api/v1/ocr/")) {
      writeJson(response, 503, { error: "OCR unused in office layout regression" });
      return;
    }
    if (url.pathname === "/") {
      const source = await readFile(join(distRoot, "index.html"), "utf8");
      const meta = [
        '<meta name="visualtex-install-token" content="office-layout-regression" />',
        '<meta name="visualtex-native-powerpoint-commit" content="false" />',
        '<meta name="visualtex-theme" content="light" />',
      ].join("\n");
      response.writeHead(200, {
        "Content-Type": "text/html; charset=utf-8",
        "Cache-Control": "no-store",
      });
      response.end(source.replace("</head>", `${meta}\n</head>`));
      return;
    }
    if (url.pathname.startsWith("/assets/")) {
      const relative = normalize(url.pathname.slice(1));
      if (relative.startsWith("..")) {
        response.writeHead(403).end();
        return;
      }
      const content = await readFile(join(distRoot, relative));
      response.writeHead(200, {
        "Content-Type": mimeTypes[extname(relative)] ?? "application/octet-stream",
        "Cache-Control": "no-store",
      });
      response.end(content);
      return;
    }
    response.writeHead(404).end();
  } catch (error) {
    response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(String(error));
  }
});

class CdpClient {
  constructor(url) {
    this.url = url;
    this.nextId = 1;
    this.pending = new Map();
  }

  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      if (!message.id) return;
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message));
      else pending.resolve(message.result);
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.socket.send(JSON.stringify({ id, method, params }));
    });
  }

  async evaluate(expression) {
    const result = await this.send("Runtime.evaluate", {
      expression,
      awaitPromise: true,
      returnByValue: true,
    });
    if (result.exceptionDetails) {
      throw new Error(
        result.exceptionDetails.exception?.description ||
          result.exceptionDetails.text ||
          "Runtime.evaluate failed",
      );
    }
    return result.result.value;
  }

  close() {
    this.socket?.close();
  }
}

async function waitFor(url, timeoutMs = 15_000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {}
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${url}`);
}

async function waitForPage() {
  const startedAt = Date.now();
  while (Date.now() - startedAt < 15_000) {
    const targets = await (
      await fetch(`http://127.0.0.1:${debugPort}/json/list`)
    ).json();
    const target = targets.find(
      (candidate) => candidate.type === "page" && candidate.url.startsWith(baseUrl),
    );
    if (target) return target;
    await sleep(80);
  }
  throw new Error("Timed out waiting for Office layout regression page");
}

async function waitForEvaluation(client, expression, description, timeoutMs = 12_000) {
  const startedAt = Date.now();
  let lastValue;
  while (Date.now() - startedAt < timeoutMs) {
    try { lastValue = await client.evaluate(expression); }
    catch (error) { if (!/navigated|context|closed/.test(String(error))) throw error; }
    if (lastValue?.ready) return lastValue;
    await sleep(60);
  }
  throw new Error(
    `Timed out waiting for ${description}: ${JSON.stringify(lastValue)}`,
  );
}


async function main() {
  await new Promise((resolve,reject)=>{server.once("error",reject);server.listen(port,"127.0.0.1",resolve)});
  let chrome,client;
  const results=[];
  try {
    chrome=spawn(chromePath,["--headless=new","--disable-gpu","--no-first-run","--no-default-browser-check",
      "--remote-debugging-port="+debugPort,"--user-data-dir="+chromeProfile,"--window-size=1280,760",officeUrl],{stdio:"ignore"});
    await waitFor("http://127.0.0.1:"+debugPort+"/json/list");
    const target=await waitForPage(); client=new CdpClient(target.webSocketDebuggerUrl); await client.connect();
    await client.send("Runtime.enable");await client.send("Page.enable");await mkdir(output,{recursive:true});
    await waitForEvaluation(client,'({ready:!!document.querySelector("math-field")})',"initial render");
    const evaluate=fn=>client.evaluate("("+fn.toString()+")()");
    const record=async (name,fn)=>{
      const value=await evaluate(fn); results.push({name,...value});
      await writeFile(join(output,"final-results.json"),JSON.stringify(results,null,2));
      console.log(JSON.stringify({name,pass:value.pass,latex:value.latex})); return value;
    };
    const shot=async name=>{const r=await client.send("Page.captureScreenshot",{format:"png"});await writeFile(join(output,name+".png"),Buffer.from(r.data,"base64"));};
    const click=async selector=>{
      const p=await client.evaluate("(() => {const e=document.querySelector("+JSON.stringify(selector)+");if(!e)throw Error('missing '+ "+JSON.stringify(selector)+");const r=e.getBoundingClientRect();const x=r.x+r.width/2,y=r.y+r.height/2,hit=document.elementFromPoint(x,y);return{x,y,hit:hit===e||e.contains(hit),cover:hit?.outerHTML.slice(0,300)};})()");
      if(!p.hit){
        await client.evaluate("(()=>{const e=document.querySelector("+JSON.stringify(selector)+"); const p=e?.closest('[data-formula-color-popover]');if(p&&p.scrollHeight>p.clientHeight)p.scrollTop=p.scrollHeight;})()");
        await sleep(120);
        const q=await client.evaluate("(()=>{const e=document.querySelector("+JSON.stringify(selector)+"),r=e.getBoundingClientRect(),x=r.x+r.width/2,y=r.y+r.height/2,h=document.elementFromPoint(x,y);return{x,y,hit:h===e||e.contains(h),cover:h?.className}})()");
        Object.assign(p,q);
        if(!p.hit)throw Error("Button not physically hittable: "+selector+" "+JSON.stringify(p));
      }
      for(const type of ["mouseMoved","mousePressed","mouseReleased"])await client.send("Input.dispatchMouseEvent",{type,x:p.x,y:p.y,button:"left",clickCount:1});
      await sleep(100);
    };
    const esc=async()=>{
      for(const type of ["keyDown","keyUp"])await client.send("Input.dispatchKeyEvent",{type,key:"Escape",code:"Escape",windowsVirtualKeyCode:27});
      await sleep(80);
    };
    const selectX=async()=>{await evaluate(()=>{const f=document.querySelector("math-field");f.setValue("x+y");f.dispatchEvent(new Event("input",{bubbles:true}));f.focus();f.selection={ranges:[[0,1]],direction:"forward"};});await sleep(100);};
    const geometry=()=>{
      const e=document.querySelector("[data-formula-color-popover]");if(!e)return{pass:false,missing:true};
      const r=e.getBoundingClientRect(),a=document.querySelector(".formula-formatting-controls").getBoundingClientRect();
      const points=[...e.querySelectorAll("button,input")].map(n=>{const b=n.getBoundingClientRect(),h=document.elementFromPoint(b.x+b.width/2,b.y+b.height/2);return{label:n.getAttribute("aria-label"),hit:n===h||n.contains(h),rect:b.toJSON(),cover:h?.className}});
      return{pass:r.left>=0&&r.top>=0&&r.right<=innerWidth&&r.bottom<=innerHeight&&points.every(p=>p.hit),
        rect:r.toJSON(),anchor:a.toJSON(),parents:(()=>{const out=[];for(let p=e.parentElement;p;p=p.parentElement){const c=getComputedStyle(p);if(/auto|scroll|hidden|clip/.test(c.overflowX+" "+c.overflowY))out.push({class:p.className,x:c.overflowX,y:c.overflowY,rect:p.getBoundingClientRect().toJSON()});}return out;})(),viewport:{width:innerWidth,height:innerHeight},points,inline:e.parentElement!==document.body,scroll:{height:e.scrollHeight,client:e.clientHeight,top:e.scrollTop},latex:document.querySelector("math-field").value};
    };
    for(const [width,height] of [[1280,760],[560,560],[760,360],[440,360]]){
      const name=width+"x"+height;
      try {
        await client.send("Emulation.setDeviceMetricsOverride",{width,height,deviceScaleFactor:1,mobile:false});
        await evaluate(()=>localStorage.setItem("visualtex-editor",JSON.stringify({state:{editorLayout:"classic",language:"zh",lines:[{id:"cp259",latex:"x+y"}],activeLineId:"cp259"},version:0})));
        await client.send("Page.reload",{ignoreCache:true});
        await waitForEvaluation(client,'({ready:document.querySelector(".workspace")?.getAttribute("data-editor-layout")==="classic"&&!!document.querySelector("math-field")})',"classic ready");
        await sleep(400);await selectX();
        await click("[data-formula-selection-color]");await record(name+" text geometry",geometry);await shot(name+"-final-text");
        await record(name+" trigger remains hittable",()=>{const e=document.querySelector("[data-formula-selection-color]"),r=e.getBoundingClientRect(),h=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);return{pass:h===e||e.contains(h),cover:h?.className}});await esc();
        await click("[data-formula-selection-color]");await esc();await record(name+" Escape preserves selected formula",()=>({pass:!document.querySelector("[data-formula-color-popover]")&&document.querySelector("math-field").value==="x+y",latex:document.querySelector("math-field").value}));
        await click("[data-formula-selection-color]");
        await evaluate(()=>{const p=document.querySelector("math-field").getBoundingClientRect();window.__outside={x:p.x+p.width-12,y:p.y+Math.min(20,p.height/2)};});
        const point=await client.evaluate("window.__outside");
        for(const type of ["mousePressed","mouseReleased"])await client.send("Input.dispatchMouseEvent",{type,...point,button:"left",clickCount:1});
        await sleep(100);await record(name+" outside closes",()=>({pass:!document.querySelector("[data-formula-color-popover]")}));
        await selectX();await click("[data-formula-selection-color]");await click('[data-formula-color="#dc2626"]');
        await record(name+" partial font color",()=>({pass:document.querySelector("math-field").value==="\\textcolor{#dc2626}{x}+y"&&!document.querySelector("[data-formula-color-popover]"),latex:document.querySelector("math-field").value}));
        await selectX();await click("[data-formula-selection-background]");await record(name+" background geometry",geometry);await shot(name+"-final-background");
        await click('[data-formula-color="#fef3c7"]');await record(name+" partial background",()=>{const s=document.querySelector("math-field").value;return{pass:s.includes("#fef3c7")&&s.endsWith("+y")&&!document.querySelector("[data-formula-color-popover]"),latex:s}});
        await evaluate(()=>{const f=document.querySelector("math-field");f.focus();f.selection={ranges:[[0,1]],direction:"forward"}});
        await click("[data-formula-selection-background]");await click("[data-formula-background-none]");
        await record(name+" background none",()=>({pass:document.querySelector("math-field").value==="x+y"&&!document.querySelector("[data-formula-color-popover]"),latex:document.querySelector("math-field").value}));
        await selectX();await click("[data-formula-selection-color]");
        await evaluate(()=>{const e=document.querySelector("[data-formula-color-popover] input");Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,"value").set.call(e,"#123456");e.dispatchEvent(new Event("input",{bubbles:true}));});
        await sleep(100);await record(name+" custom add preserves selection",()=>({pass:!!document.querySelector('[data-formula-custom-color="#123456"]')&&document.querySelector("math-field").value==="x+y"}));
        await record(name+" custom text geometry",geometry);await click('[data-formula-custom-color="#123456"] > .formula-color-swatch');
        await record(name+" custom font applies",()=>({pass:document.querySelector("math-field").value==="\\textcolor{#123456}{x}+y"&&!document.querySelector("[data-formula-color-popover]"),latex:document.querySelector("math-field").value}));
        await selectX();await click("[data-formula-selection-background]");
        await evaluate(()=>{const e=document.querySelector("[data-formula-color-popover] input");Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,"value").set.call(e,"#123456");e.dispatchEvent(new Event("input",{bubbles:true}));});await sleep(100);
        await click('[data-formula-custom-color="#123456"] > .formula-color-swatch');
        await record(name+" custom background applies",()=>{const s=document.querySelector("math-field").value;return{pass:s.includes("#123456")&&s.endsWith("+y")&&!document.querySelector("[data-formula-color-popover]"),latex:s}});
        await click("[data-formula-selection-color]");await click('[data-delete-formula-custom-color="#123456"]');
        await record(name+" custom delete stays open",()=>({pass:!!document.querySelector("[data-formula-color-popover]")&&!document.querySelector('[data-formula-custom-color="#123456"]')}));
        await record(name+" alternate trigger remains hittable",()=>{const e=document.querySelector("[data-formula-selection-background]"),r=e.getBoundingClientRect(),h=document.elementFromPoint(r.x+r.width/2,r.y+r.height/2);return{pass:h===e||e.contains(h),cover:h?.className}});await esc();
      }catch(error){results.push({name:name+" flow",pass:false,error:String(error)});await shot(name+"-continued-failure");await writeFile(join(output,"final-results.json"),JSON.stringify(results,null,2));console.log(JSON.stringify(results.at(-1)));await esc();}
    }
    await esc();await client.send("Emulation.setDeviceMetricsOverride",{width:1280,height:760,deviceScaleFactor:1,mobile:false});await sleep(300);
    await selectX();await click("[data-formula-selection-background]");
    await client.send("Emulation.setDeviceMetricsOverride",{width:440,height:360,deviceScaleFactor:1,mobile:false});await sleep(450);
    await record("live resize background",geometry);await shot("live-resize-background");
    await client.send("Emulation.setDeviceMetricsOverride",{width:1280,height:760,deviceScaleFactor:1,mobile:false});await sleep(350);
    await record("live restore background",geometry);await esc();
    await click("[data-formula-selection-color]");
    const scrolls=await evaluate(()=>[...document.querySelectorAll(".classic-bottom-tabs,.classic-bottom-content,.classic-editor-pane-body,.editor-pane-header-controls,.formula-toolbar")].map(e=>({class:e.className,width:e.clientWidth,scrollWidth:e.scrollWidth,height:e.clientHeight,scrollHeight:e.scrollHeight})));
    results.push({name:"classic scroll surfaces",info:true,scrolls});
    await client.send("Input.dispatchMouseEvent",{type:"mouseWheel",x:320,y:700,deltaX:220,deltaY:100});await sleep(250);
    await record("dock wheel with popup",geometry);await esc();
    await writeFile(join(output,"final-results.json"),JSON.stringify(results,null,2));
    console.log(JSON.stringify({total:results.filter(r=>!r.info).length,passed:results.filter(r=>r.pass).length,failed:results.filter(r=>r.pass===false).length}));
    await client.send("Browser.close");
    assert.ok(results.every(r=>r.pass!==false),"Classic color popup QA has failures; see results.json");
  }finally{
    client?.close();if(chrome&&!chrome.killed)chrome.kill();await new Promise(resolve=>server.close(resolve));
    await rm(chromeProfile,{recursive:true,force:true}).catch(()=>{});
  }
}
await main();

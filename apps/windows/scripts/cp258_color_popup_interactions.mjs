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
const stage = process.argv[2] || "before";
const output = join(process.cwd(), "../../artifacts/conversion-symbol-audit-20261002/cp258-color-popup");
const distRoot = join(output, stage + "-desktop");
let editorLayout = "standard";
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
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });

  let chrome;
  let client;
  try {
    await waitFor(`${baseUrl}/api/v1/preferences`);
    chrome = spawn(
      chromePath,
      [
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        `--remote-debugging-port=${debugPort}`,
        `--user-data-dir=${chromeProfile}`,
        "--window-size=1760,760",
        officeUrl,
      ],
      { stdio: "ignore" },
    );

    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    const target = await waitForPage();
    client = new CdpClient(target.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Page.enable");

    await waitForEvaluation(client, '({ready:!!document.querySelector("math-field")})', "initial desktop render");
    await mkdir(output, {recursive:true});
    const results = [];
    const click = async selector => {
      const p = await client.evaluate('(() => { const e=document.querySelector('+JSON.stringify(selector)+'); if(!e) throw Error("missing "+'+JSON.stringify(selector)+'); const r=e.getBoundingClientRect(); return {x:r.x+r.width/2,y:r.y+r.height/2};})()');
      await client.send("Input.dispatchMouseEvent",{type:"mouseMoved",...p});
      await client.send("Input.dispatchMouseEvent",{type:"mousePressed",button:"left",clickCount:1,...p});
      await client.send("Input.dispatchMouseEvent",{type:"mouseReleased",button:"left",clickCount:1,...p});
      await sleep(100);
    };

    const check = async (name, expression) => {
      const value=await client.evaluate(expression);
      results.push({name,...value});
      await writeFile(join(output,stage+"-interactions.json"),JSON.stringify(results,null,2));
      assert.equal(value.pass,true,name+": "+JSON.stringify(value));
    };
    const selectX = async () => {
      await client.evaluate('(() => {const f=document.querySelector("math-field");f.setValue("x+y");f.dispatchEvent(new Event("input",{bubbles:true}));f.focus();f.selection={ranges:[[0,1]],direction:"forward"};})()');
      await sleep(100);
    };
    await client.evaluate('localStorage.setItem("visualtex-editor",JSON.stringify({state:{editorLayout:"standard",language:"zh",lines:[{id:"cp258",latex:"x+y"}],activeLineId:"cp258"},version:0}))');
    await client.send("Page.reload",{ignoreCache:true});
    await waitForEvaluation(client,'({ready:document.querySelector(".workspace")?.getAttribute("data-editor-layout")==="standard" && !!document.querySelector("math-field")})',"standard desktop ready");
    await sleep(400);
    for(const width of [1280,560]){
      await client.send("Emulation.setDeviceMetricsOverride",{width,height:560,deviceScaleFactor:1,mobile:false});
      await sleep(350);
      await selectX();
      await click("[data-formula-selection-color]");
      await check(width+" physical open", '({pass:document.querySelector("[data-formula-color-popover]")?.parentElement===document.body})');
      await click("[data-formula-selection-color]");
      await check(width+" toggle closed", '({pass:!document.querySelector("[data-formula-color-popover]")})');
      await click("[data-formula-selection-color]");
      await client.send("Input.dispatchKeyEvent",{type:"keyDown",key:"Escape",code:"Escape",windowsVirtualKeyCode:27});
      await client.send("Input.dispatchKeyEvent",{type:"keyUp",key:"Escape",code:"Escape",windowsVirtualKeyCode:27});
      await sleep(80);
      await check(width+" Escape closed", '({pass:!document.querySelector("[data-formula-color-popover]")&&document.querySelector("math-field").value==="x+y",latex:document.querySelector("math-field").value})');
      await click("[data-formula-selection-color]");
      await client.send("Input.dispatchMouseEvent",{type:"mousePressed",x:width-25,y:330,button:"left",clickCount:1});
      await client.send("Input.dispatchMouseEvent",{type:"mouseReleased",x:width-25,y:330,button:"left",clickCount:1});
      await sleep(80);
      await check(width+" outside closed", '({pass:!document.querySelector("[data-formula-color-popover]")})');
      await selectX();
      await click("[data-formula-selection-color]");
      await click('[data-formula-color="#dc2626"]');
      await check(width+" partial text color", String.raw`(() => {const s=document.querySelector("math-field").value;return {pass:s==="\\textcolor{#dc2626}{x}+y"&&!document.querySelector("[data-formula-color-popover]"),latex:s}})()`);
      await selectX();
      await click("[data-formula-selection-background]");
      await click('[data-formula-color="#fef3c7"]');
      await check(width+" partial background", String.raw`(() => {const s=document.querySelector("math-field").value;return {pass:s.includes("#fef3c7")&&s.endsWith("+y")&&!document.querySelector("[data-formula-color-popover]"),latex:s}})()`);
      await client.evaluate('(() => {const f=document.querySelector("math-field");f.focus();f.selection={ranges:[[0,1]],direction:"forward"};})()');
      await click("[data-formula-selection-background]");
      await click("[data-formula-background-none]");
      await check(width+" background none", '({pass:document.querySelector("math-field").value==="x+y"&&!document.querySelector("[data-formula-color-popover]"),latex:document.querySelector("math-field").value})');
      await selectX();
      await click("[data-formula-selection-color]");
      await client.evaluate('(() => {const e=document.querySelector("[data-formula-color-popover] input");const set=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,"value").set;set.call(e,"#123456");e.dispatchEvent(new Event("input",{bubbles:true}));})()');
      await sleep(100);
      await check(width+" add custom preserves selection", '({pass:!!document.querySelector(\'[data-formula-custom-color="#123456"]\')&&document.querySelector("math-field").value==="x+y"})');
      await click('[data-formula-custom-color="#123456"] > .formula-color-swatch');
      await check(width+" custom color applies", String.raw`(() => {const s=document.querySelector("math-field").value;return {pass:s==="\\textcolor{#123456}{x}+y"&&!document.querySelector("[data-formula-color-popover]"),latex:s}})()`);
      await click("[data-formula-selection-color]");
      await click('[data-delete-formula-custom-color="#123456"]');
      await check(width+" custom delete stays open", '({pass:!!document.querySelector("[data-formula-color-popover]")&&!document.querySelector(\'[data-formula-custom-color="#123456"]\')})');
      await click("[data-formula-selection-background]");
      await check(width+" switch palette", '({pass:document.querySelector("[data-formula-color-popover]")?.getAttribute("data-formula-color-popover")==="backgroundColor"})');
      await click("[data-formula-selection-background]");
    }
    await click("[data-formula-selection-color]");
    await client.send("Emulation.setDeviceMetricsOverride",{width:440,height:360,deviceScaleFactor:1,mobile:false});
    await sleep(450);
    await check("resize while open", '(() => {const p=document.querySelector("[data-formula-color-popover]"),r=p.getBoundingClientRect(),a=document.querySelector(".formula-formatting-controls").getBoundingClientRect();return {pass:r.left>=0&&r.right<=innerWidth&&r.top>=0&&r.bottom<=innerHeight&&Math.abs(r.top-a.bottom-7)<1,rect:r.toJSON()}})()');
    await client.evaluate('document.querySelector(".editor-pane-header-controls").scrollLeft=35');
    await sleep(350);
    await check("toolbar scroll tracks anchor", '(() => {const p=document.querySelector("[data-formula-color-popover]"),r=p.getBoundingClientRect(),a=document.querySelector(".formula-formatting-controls").getBoundingClientRect();return {pass:r.left>=0&&r.right<=innerWidth&&Math.abs(r.right-a.right)<1,rect:r.toJSON(),scrollLeft:document.querySelector(".editor-pane-header-controls").scrollLeft}})()');
    console.log(JSON.stringify(results,null,2));
    await client.send("Browser.close");
  } finally {
    client?.close();
    if (chrome && !chrome.killed) chrome.kill();
    await new Promise((resolve) => server.close(resolve));
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();

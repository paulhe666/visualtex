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
const officeUrl =
  `${baseUrl}/dialog/index.html?sessionId=${sessionId}&runtime=vsto`;
const stage = process.argv[2] || "before";
const output = join(process.cwd(), "../../artifacts/conversion-symbol-audit-20261002/cp258-color-popup");
const distRoot = join(output, stage + "-office");
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
    if (url.pathname.startsWith("/dialog/")) {
      const source = await readFile(join(distRoot, "office-dialog.html"), "utf8");
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

    await mkdir(output, {recursive:true});
    const results = [];
    const click = async selector => {
      const p = await client.evaluate('(() => { const e=document.querySelector('+JSON.stringify(selector)+'); if(!e) throw Error("missing "+'+JSON.stringify(selector)+'); const r=e.getBoundingClientRect(); return {x:r.x+r.width/2,y:r.y+r.height/2};})()');
      await client.send("Input.dispatchMouseEvent",{type:"mouseMoved",...p});
      await client.send("Input.dispatchMouseEvent",{type:"mousePressed",button:"left",clickCount:1,...p});
      await client.send("Input.dispatchMouseEvent",{type:"mouseReleased",button:"left",clickCount:1,...p});
      await sleep(100);
    };
    for (const layout of ["standard", "classic"]) {
      editorLayout = layout;
      for (const [width,height] of [[1280,760],[760,560],[560,420]]) {
        await client.send("Emulation.setDeviceMetricsOverride",{width,height,deviceScaleFactor:1,mobile:false});
        await client.send("Page.reload",{ignoreCache:true});
        await waitForEvaluation(client, '({ready:document.querySelector(".workspace")?.getAttribute("data-editor-layout")==='+JSON.stringify(layout)+' && !!document.querySelector("math-field"), actual:document.querySelector(".workspace")?.getAttribute("data-editor-layout"),text:document.body?.innerText.slice(0,800)})',"layout ready");
        await sleep(250);
        for (const kind of ["color","backgroundColor"]) {
          console.log({layout,width,height,kind});
          await client.evaluate('document.querySelector('+JSON.stringify(kind==="color"?"[data-formula-selection-color]":"[data-formula-selection-background]")+').click()');
          await sleep(100);
          await waitForEvaluation(client,'({ready:!!document.querySelector("[data-formula-color-popover]")})',"palette open");
          const geometry=await client.evaluate(String.raw`(() => {
            const el=document.querySelector("[data-formula-color-popover]");
            const r=el.getBoundingClientRect();
            const parents=[];let p=el.parentElement;
            while(p){const c=getComputedStyle(p),b=p.getBoundingClientRect();parents.push({class:p.className,overflowX:c.overflowX,overflowY:c.overflowY,contain:c.contain,rect:{x:b.x,y:b.y,width:b.width,height:b.height}});p=p.parentElement;}
            const points=[...el.querySelectorAll("button,input")].map(e=>{const b=e.getBoundingClientRect(),hit=document.elementFromPoint(b.x+b.width/2,b.y+b.height/2);return {label:e.getAttribute("aria-label"),hit:e===hit||e.contains(hit),rect:{x:b.x,y:b.y,width:b.width,height:b.height},cover:hit?.className};});
            return {rect:{x:r.x,y:r.y,width:r.width,height:r.height},viewport:{width:innerWidth,height:innerHeight},inside:r.left>=0&&r.top>=0&&r.right<=innerWidth&&r.bottom<=innerHeight,points,parents,allHit:points.every(p=>p.hit)};
          })()`);
          results.push({layout,width,height,kind,...geometry});
          await writeFile(join(output,stage+"-geometry.json"),JSON.stringify(results,null,2));
          const shot=await client.send("Page.captureScreenshot",{format:"png"});
          await writeFile(join(output,stage+"-"+layout+"-"+width+"-"+kind+".png"),Buffer.from(shot.data,"base64"));
          await client.send("Input.dispatchKeyEvent",{type:"keyDown",key:"Escape",code:"Escape",windowsVirtualKeyCode:27});
          await client.send("Input.dispatchKeyEvent",{type:"keyUp",key:"Escape",code:"Escape",windowsVirtualKeyCode:27});
          await sleep(60);
        }
      }
    }
    await writeFile(join(output,stage+"-geometry.json"),JSON.stringify(results,null,2));
    console.log(JSON.stringify(results.map(({layout,width,height,kind,inside,allHit,rect})=>({layout,width,height,kind,inside,allHit,rect})),null,2));
    if (stage !== "before") {
      assert.equal(results.length, 12);
      assert.ok(results.every((row) => row.inside && row.allHit), "Every palette control must remain visible and hittable");
    }
    await client.send("Browser.close");
  } finally {
    client?.close();
    if (chrome && !chrome.killed) chrome.kill();
    await new Promise((resolve) => server.close(resolve));
    await rm(chromeProfile, { recursive: true, force: true }).catch(() => {});
  }
}

await main();

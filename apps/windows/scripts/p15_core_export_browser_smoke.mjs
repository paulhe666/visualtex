import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { existsSync } from "node:fs";
import { rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import process from "node:process";

const offset = process.pid % 1000;
const previewPort = 7700 + offset;
const debugPort = 13700 + offset;
const baseUrl = `http://127.0.0.1:${previewPort}`;
const fixtureUrl = `${baseUrl}/scripts/p15_export_browser_fixture.html`;
const profile = join(tmpdir(), `visualtex-p15-export-${process.pid}`);
const candidates = [
  process.env.VISUALTEX_CHROME_PATH,
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env["PROGRAMFILES(X86)"] && join(process.env["PROGRAMFILES(X86)"], "Microsoft", "Edge", "Application", "msedge.exe"),
  process.env.PROGRAMFILES && join(process.env.PROGRAMFILES, "Google", "Chrome", "Application", "chrome.exe"),
  process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, "Google", "Chrome", "Application", "chrome.exe"),
].filter(Boolean);
const browserPath = candidates.find((candidate) => existsSync(candidate));
if (!browserPath) throw new Error("No Chrome/Edge executable found for P15 export acceptance.");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function waitFor(url, timeout = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    try { const response = await fetch(url); if (response.ok) return response; } catch {}
    await sleep(100);
  }
  throw new Error(`Timed out waiting for ${url}`);
}
class CdpClient {
  constructor(url) { this.url = url; this.id = 1; this.pending = new Map(); this.events = []; }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((resolve, reject) => {
      this.socket.addEventListener("open", resolve, { once: true });
      this.socket.addEventListener("error", reject, { once: true });
    });
    this.socket.addEventListener("message", (event) => {
      const message = JSON.parse(event.data);
      const pending = this.pending.get(message.id);
      if (!pending) { this.events.push(message); return; }
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
  const vite = spawn(process.execPath, ["node_modules/vite/bin/vite.js", "--host", "127.0.0.1", "--port", String(previewPort), "--strictPort"], { cwd: process.cwd(), stdio: "ignore" });
  let browser;
  let client;
  try {
    await waitFor(fixtureUrl);
    browser = spawn(browserPath, ["--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check", `--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, "--window-size=1280,900", fixtureUrl], { stdio: "ignore" });
    await waitFor(`http://127.0.0.1:${debugPort}/json/list`);
    let page;
    for (let attempt = 0; attempt < 100 && !page; attempt += 1) {
      const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json();
      page = targets.find((target) => target.type === "page" && target.url.startsWith(fixtureUrl));
      if (!page) await sleep(100);
    }
    if (!page) throw new Error("P15 export fixture page was not found.");
    client = new CdpClient(page.webSocketDebuggerUrl);
    await client.connect();
    await client.send("Runtime.enable");
    await client.send("Log.enable");
    const evaluate = async (expression) => {
      const result = await client.send("Runtime.evaluate", { expression, awaitPromise: true, returnByValue: true });
      if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
      return result.result.value;
    };
    const started = Date.now();
    while (true) {
      const readyState = await evaluate(`document.querySelector('#ready')?.textContent??''`);
      if (readyState === "ready") break;
      if (readyState.startsWith("error:")) throw new Error(readyState);
      if (Date.now() - started > 45000) {
        const pageState = await evaluate(`({url:location.href,html:document.documentElement.outerHTML,scripts:[...document.scripts].map(s=>({type:s.type,src:s.src,text:s.textContent?.slice(0,300)}))})`);
        const diagnostics = client.events.filter((event) => /exception|console|Log\.entry/i.test(event.method ?? "")).slice(-12);
        throw new Error(`P15 export runtime did not become ready: ${readyState}\n${JSON.stringify({ pageState, diagnostics }, null, 2)}`);
      }
      await sleep(80);
    }
    const call = async (name, args) => evaluate(`(async()=>{try{return{ok:true,value:await window.p15ExportRuntime[${JSON.stringify(name)}](...${JSON.stringify(args)})}}catch(error){return{ok:false,error:error instanceof Error?error.message:String(error)}}})()`);

    const formulas = [
      ["fraction", String.raw`\frac{a+b}{c+d}`, /<mfrac>/u],
      ["root", String.raw`\sqrt[n]{x^2+y^2}`, /<mroot>/u],
      ["matrix", String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`, /<mtable(?:\s|>)/u],
      ["multiline", "a=b+c\nd=e-f", /<mtable(?:\s|>)/u],
      ["equation-aligned", String.raw`\begin{equation}\begin{aligned}f^*(\mathbf{x})&=\frac{1}{p(\mathbf{x})}\\&=\mathbb{E}_{t}[t\mid\mathbf{x}]\end{aligned}\end{equation}`, /<mtable(?:\s|>)/u],
      ["upright-constants", String.raw`\mathrm{e}^{\mathrm{i}\pi}+1=0`, /mathvariant="normal"/u],
      ["accent", String.raw`\widehat{abc}+\vec{x}`, /<mover(?:\s|>)/u],
      ["physics", String.raw`\qty(\frac{a}{b})+\dv{f}{x}+\abs{x}`, /<mfrac>/u],
      ["physics-orders", String.raw`\dv[2]{f}{x}+\pdv[3]{g}{y}+\fdv{S}{\phi}`, /<mfrac>/u],
      ["physics-vectors", String.raw`\vb{v}+\va{a}+\vu{n}+\pb{f}{g}+\order{x^2}+\Tr A+\rank A`, /<math\b/u],
      ["physics-matrix-quantity", String.raw`\mqty{a&b\\c&d}+\pmqty{1&0\\0&1}+\vmqty{x&y\\z&w}`, /<mtable(?:\s|>)/u],
      ["siunitx", String.raw`\SI{3}{\meter\per\second}`, /<math\b/u],
      ["siunitx-options", String.raw`\SI[round-mode=places]{3.14}{\kilo\meter\per\second}+\qty[round-mode=figures]{5}{\tesla}`, /<math\b/u],
      ["siunitx-range", String.raw`\qtyrange{1}{10}{\milli\second}+\ang{30}+\unit{\joule\per\mole}`, /<math\b/u],
    ];
    for (const [name, latex, mathMlPattern] of formulas) {
      const svgCall = await call("latexToSvg", [latex, { displayMode: true, fontSizePt: 14, paddingPx: 8, background: "transparent" }]);
      assert.equal(svgCall.ok, true, `${name} SVG failed: ${svgCall.error ?? ""}`);
      const svg = svgCall.value;
      assert.match(svg.svg, /^<svg\b/u, `${name} SVG root`);
      assert.ok(svg.width > 0 && svg.height > 0 && svg.baseline >= 0, `${name} geometry`);
      assert.match(svg.svg, /<rect\b[^>]*fill-opacity="0\.001"/u, `${name} transparent hit target`);
      assert.doesNotMatch(svg.svg, /MathJax|data-mml-node|mjx|<foreignObject\b|<script\b/iu, `${name} retired renderer material`);
      assert.equal(svg.base64, await evaluate(`window.p15ExportRuntime.svgToBase64(${JSON.stringify(svg.svg)})`), `${name} base64`);
      const mathCall = await call("latexToMathMl", [latex, true]);
      assert.equal(mathCall.ok, true, `${name} MathML failed: ${mathCall.error ?? ""}`);
      assert.match(mathCall.value, /^<math\b/u, `${name} MathML root`);
      assert.match(mathCall.value, mathMlPattern, `${name} semantic MathML`);
      assert.doesNotMatch(mathCall.value, /data-visualtex-core-grid-id|data-mtef-tabs|data-visualtex-mtef-ruler-stops/u, `${name} private MathML attributes`);
    }

    const chineseMathMl = await call("latexToMathMl", [String.raw`\text{测试}+\alpha`, false]);
    assert.equal(chineseMathMl.ok, true, `Chinese MathML failed: ${chineseMathMl.error ?? ""}`);
    assert.match(chineseMathMl.value, /测试/u);

    const placeholder = await call("latexToSvg", [String.raw`x+\placeholder{}`, { displayMode: false, fontSizePt: 12, paddingPx: 1, background: "transparent" }]);
    assert.equal(placeholder.ok, false);
    assert.match(placeholder.error, /empty VisualTeX placeholders/i);
    const unknown = await call("latexToMathMl", [String.raw`\definitelyUnknownVisualTeXCommand`, true]);
    assert.equal(unknown.ok, false);
    assert.match(unknown.error, /unsupported|unknown|unparsed|command/i);
    const empty = await call("latexToSvg", ["", { displayMode: true, fontSizePt: 12, paddingPx: 8, background: "transparent" }]);
    assert.equal(empty.ok, false);
    assert.match(empty.error, /Cannot export an empty formula/i);

    const resources = await evaluate(`performance.getEntriesByType('resource').map(x=>({name:String(x.name),transferSize:Number(x.transferSize??0),initiatorType:String(x.initiatorType??'')}))`);
    const wasmResources = resources.filter((item) => item.name.includes("visualtex_engine_wasm_bg"));
    const fontResources = resources.filter((item) => item.name.includes("latinmodern-math"));
    const wasmBodyFetches = wasmResources.filter((item) => item.initiatorType === "fetch" && !item.name.includes("?import&url"));
    const fontBodyFetches = fontResources.filter((item) => item.initiatorType === "fetch" && !item.name.includes("?import&url"));
    assert.equal(wasmBodyFetches.length, 1, "core WASM must perform exactly one body fetch");
    assert.ok(wasmBodyFetches[0].transferSize > 1_000_000, "core WASM body fetch must contain the actual binary payload");
    assert.equal(fontBodyFetches.length, 1, "core math font must perform exactly one body fetch");
    assert.ok(fontBodyFetches[0].transferSize > 100_000, "core math font body fetch must contain the actual font payload");
    assert.equal(resources.some((item) => /mathlive|mathjax|KaTeX_/i.test(item.name)), false, "retired runtime resource loaded");

    console.log(JSON.stringify({ status: "PASS", actualBrowser: true, formulaClasses: formulas.map(([name]) => name), chineseMathMl: true, failClosed: ["placeholder", "unknown-command", "empty"], coreResources: { wasmEntries: wasmResources.length, wasmBodyFetches: wasmBodyFetches.length, mathFontEntries: fontResources.length, mathFontBodyFetches: fontBodyFetches.length, viteUrlResolverEntriesExpected: true }, retiredRuntimeResources: 0 }, null, 2));
  } finally {
    client?.close();
    browser?.kill();
    vite.kill();
    await rm(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 120 }).catch(() => {});
  }
}

await main();

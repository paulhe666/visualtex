import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { readFile, rm } from "node:fs/promises";
import { createServer } from "node:http";
import { fileURLToPath } from "node:url";
import { resolve } from "node:path";
import { latexToMathMl } from "../src/export/runtime.ts";
import { createBrowserProfilePath, resolveChromiumExecutable } from "./browser_test_runtime.mjs";

const root = fileURLToPath(new URL("../../../", import.meta.url));
const probe = resolve(root, "artifacts/conversion-symbol-audit-20261002/spacing-source-probe/bin/Debug/net9.0/SpacingSourceProbe.dll");
function csharp(inputs: unknown[]) {
  const run = spawnSync("dotnet", [probe], { input: inputs.map(x => JSON.stringify(x)).join("\n") + "\n", encoding: "utf8" });
  assert.equal(run.status, 0, run.stderr);
  return run.stdout.trim().split(/\r?\n/).map(x => JSON.parse(x));
}
const widths = [
  "1em", "2.00em", "+001.000em", ".0000001em", "0.123456789em",
  "-0.333333em", "1.00000000000000000000000000001em",
  "0.99999999999999999999999999999em", "-0.00em", " .500 pt ",
  ...["em", "ex", "pt", "pc", "px", "in", "cm", "mm"].flatMap(unit => ["0", "1.25", "-0.375"].map(n => n + unit)),
];
const inputs = widths.map(width => ({ mathMl: '<math><mi>x</mi><mspace width="' + width + '"/><mi>y</mi></math>' }));
const sources = csharp(inputs).map((r, i) => ({ width: widths[i], latex: r.latex }));
function exportedWidth(mathMl: string) {
  const tag = mathMl.match(/<mspace(?:\s[^>]*)?>/)?.[0];
  return tag ? tag.match(/\bwidth="([^"]+)"/)?.[1] ?? "0em" : undefined;
}
for (const source of sources) {
  for (const display of [false, true]) {
    const mathMl = await latexToMathMl(source.latex, display);
    const width = exportedWidth(mathMl);
    assert.ok(width, mathMl);
    assert.equal(csharp([{ compare: [source.width, width] }])[0].equal, true, JSON.stringify({ source, width }));
  }
}
console.log(JSON.stringify({ stage: "production MathML recovery/export", passed: sources.length * 2 }));
const greekCases = [
  {mathMl:"<math><mi>Γ</mi></math>",variant:"italic"},
  {mathMl:"<math><mi mathvariant='normal'>Γ</mi></math>",variant:"normal"},
  {mathMl:"<math><mi mathvariant='italic'>Γ</mi></math>",variant:"italic"},
  {mathMl:"<math><mi>α</mi></math>",variant:"italic"},
  {mathMl:"<math><mi mathvariant='normal'>α</mi></math>",variant:"normal"},
  {mathMl:"<math><mi>ΓΔ</mi></math>",variant:"normal"},
  {mathMl:"<math><mstyle mathvariant='normal'><mi mathvariant='italic'>Γ</mi></mstyle></math>",variant:"italic"},
];
const greekSources = csharp(greekCases);
for (let i = 0; i < greekCases.length; i++) {
  const mml = await latexToMathMl(greekSources[i].latex, false);
  const tokens = [...mml.matchAll(/<mi([^>]*)>([^<]+)<\/mi>/g)];
  assert.ok(tokens.length > 0, mml);
  for (const token of tokens)
    assert.equal(token[1].match(/\bmathvariant="([^"]+)"/)?.[1] ?? "italic", greekCases[i].variant, JSON.stringify({input:greekCases[i],source:greekSources[i],mml}));
}
console.log(JSON.stringify({stage:"shared declared/effective Greek style→source→production export",passed:greekCases.length}));

const library = await readFile(new URL("../node_modules/mathlive/mathlive.mjs", import.meta.url));
const server = createServer((request, response) => {
  if (request.url === "/mathlive.js") { response.setHeader("content-type", "text/javascript"); response.end(library); }
  else if (request.url === "/") { response.setHeader("content-type", "text/html"); response.end("<!doctype html><html><body></body></html>"); }
  else { response.writeHead(404); response.end(); }
});
await new Promise<void>(resolve => server.listen(0, "127.0.0.1", resolve));
const address = server.address() as { port: number };
const baseUrl = "http://127.0.0.1:" + address.port;
const profile = createBrowserProfilePath("visualtex-space-source-roundtrip");
const debugPort = 26000 + process.pid % 1000;
const browser = spawn(resolveChromiumExecutable(), [
  "--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
  "--remote-debugging-port=" + debugPort, "--user-data-dir=" + profile, baseUrl,
], { stdio: "ignore" });
let socket: WebSocket | undefined;
try {
  let targets: any[] = [];
  const started = Date.now();
  while (Date.now() - started < 15000) {
    try {
      const response = await fetch("http://127.0.0.1:" + debugPort + "/json/list");
      targets = await response.json() as any[];
      if (targets.some(x => x.type === "page" && x.url.startsWith(baseUrl))) break;
    } catch {}
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  const page = targets.find(x => x.type === "page" && x.url.startsWith(baseUrl));
  assert.ok(page, "Isolated source test browser did not start");
  socket = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket!.addEventListener("open", resolve, {once:true}); socket!.addEventListener("error", reject, {once:true}); });
  let next = 0;
  const pending = new Map<number, {resolve: (x:any)=>void, reject:(e:unknown)=>void}>();
  socket.addEventListener("message", event => {
    const data = JSON.parse(String(event.data)); const waiter = pending.get(data.id);
    if (waiter) { pending.delete(data.id); data.error ? waiter.reject(data.error) : waiter.resolve(data.result); }
  });
  function send(method: string, params: unknown) {
    return new Promise<any>((resolve, reject) => { const id = ++next; pending.set(id, {resolve, reject}); socket!.send(JSON.stringify({id,method,params})); });
  }
  await send("Page.enable", {});
  await send("Page.navigate", {url:baseUrl});
  for (let attempt = 0; attempt < 150; attempt++) {
    const ready = await send("Runtime.evaluate", {returnByValue:true, expression: "location.href === " + JSON.stringify(baseUrl + "/") + " && document.readyState === 'complete'"});
    if (ready.result?.value) break;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  const evaluated = await send("Runtime.evaluate", { awaitPromise:true, returnByValue:true, expression:
    "(async () => {" +
    " const {MathfieldElement} = await import(" + JSON.stringify(baseUrl + "/mathlive.js") + ");" +
    " MathfieldElement.fontsDirectory = null; MathfieldElement.soundsDirectory = null;" +
    " const field = new MathfieldElement(); field.mathVirtualKeyboardPolicy='manual'; document.body.append(field);" +
    " const result=[]; for(const item of " + JSON.stringify(sources) + ") {" +
    " field.setValue(item.latex,{silenceNotifications:true});" +
    " field.position = field.lastOffset; field.insert('+z'); field.executeCommand('deleteBackward'); field.executeCommand('deleteBackward');" +
    " result.push({...item,source:field.getValue('latex'),productSource:field.getValue('latex-without-placeholders'),value:field.value,expanded:field.getValue('latex-expanded'),errors:field.errors});" +
    " } field.remove(); return result; })()"
  });
  assert.equal(evaluated.exceptionDetails, undefined, JSON.stringify(evaluated.exceptionDetails));
  const edited = evaluated.result.value as Array<{width:string; latex:string; source:string; productSource:string; value:string; expanded:string; errors:unknown[]}>;
  for (const item of edited) {
    assert.deepEqual(item.errors, [], JSON.stringify(item));
    for (const source of [item.source, item.productSource, item.value]) {
      const mathMl = await latexToMathMl(source, false);
      const width = exportedWidth(mathMl);
      assert.ok(width, JSON.stringify(item));
      assert.equal(csharp([{compare:[item.width,width]}])[0].equal, true, JSON.stringify({item,width}));
    }
  }
  console.log(JSON.stringify({stage:"actual MathfieldElement setValue/getValue source→production export",passed:edited.length * 3}));
  // Expanded-source serialization is a separate, explicitly measured boundary:
  // MathLive converts dimension digits to JS numbers when expanding macros.
  const expandedChanges = [];
  for(const item of edited) {
    try {
      const width = exportedWidth(await latexToMathMl(item.expanded, false));
      if (!width || !csharp([{compare:[item.width,width]}])[0].equal) expandedChanges.push({width:item.width, expanded:item.expanded, exportedWidth:width});
    } catch (error) {
      expandedChanges.push({width:item.width, expanded:item.expanded, exportError:(error as Error).message});
    }
  }
  assert.deepEqual(expandedChanges.map(x => x.width), [".0000001em", "1.00000000000000000000000000001em", "0.99999999999999999999999999999em"], "Expanded-source support boundary changed; review rather than suppress it");
  console.log(JSON.stringify({scope:"No pixel/native acceptance; source getValue formats tested. Expanded format limitations measured separately.", expandedChanges}));
} finally {
  socket?.close();
  browser.kill("SIGTERM");
  await new Promise<void>(resolve => server.close(() => resolve()));
  await new Promise(resolve => setTimeout(resolve, 250));
  await rm(profile, {recursive:true,force:true}).catch(() => undefined);
}

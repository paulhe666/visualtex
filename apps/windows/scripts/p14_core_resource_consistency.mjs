import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { createServer } from "node:http";
import { cp, mkdtemp, readFile, readdir, rm, stat } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, extname, join, resolve } from "node:path";
import process from "node:process";

const root = process.cwd();
const repoRoot = resolve(root, "../..");
const enginePkg = resolve(repoRoot, "engine/web/pkg");
const sourceWasm = resolve(enginePkg, "visualtex_engine_wasm_bg.wasm");
const sourceFont = resolve(root, "src-windows/artifacts/office-fonts/latinmodern-math.otf");
const runtimeSource = resolve(root, "src/engine/coreEngineRuntime.ts");
const desktopDist = resolve(root, "dist");
const officeDist = resolve(root, "dist-office-windows-native");

const sha = (bytes) => createHash("sha256").update(bytes).digest("hex");
async function record(path) {
  const bytes = await readFile(path);
  return { path, bytes: bytes.length, sha256: sha(bytes) };
}
async function uniqueAsset(dist, expression, label) {
  const assets = (await readdir(resolve(dist, "assets"))).filter((name) => expression.test(name));
  assert.equal(assets.length, 1, `${label} must contain exactly one matching asset, got ${assets.join(", ")}`);
  return resolve(dist, "assets", assets[0]);
}

const sourceWasmRecord = await record(sourceWasm);
const sourceFontRecord = await record(sourceFont);
const desktopWasm = await uniqueAsset(desktopDist, /^visualtex_engine_wasm_bg-.*\.wasm$/u, "core-enabled desktop build");
const desktopFont = await uniqueAsset(desktopDist, /^latinmodern-math-.*\.otf$/u, "core-enabled desktop build");
const officeWasm = await uniqueAsset(officeDist, /^visualtex_engine_wasm_bg-.*\.wasm$/u, "Office build");
const officeFont = await uniqueAsset(officeDist, /^latinmodern-math-.*\.otf$/u, "Office build");
const desktopWasmRecord = await record(desktopWasm);
const desktopFontRecord = await record(desktopFont);
const officeWasmRecord = await record(officeWasm);
const officeFontRecord = await record(officeFont);

assert.equal(desktopWasmRecord.sha256, sourceWasmRecord.sha256, "desktop WASM differs from engine/web/pkg source");
assert.equal(officeWasmRecord.sha256, sourceWasmRecord.sha256, "Office WASM differs from engine/web/pkg source");
assert.equal(desktopFontRecord.sha256, sourceFontRecord.sha256, "desktop math font differs from source font");
assert.equal(officeFontRecord.sha256, sourceFontRecord.sha256, "Office math font differs from source font");

const runtimeText = await readFile(runtimeSource, "utf8");
const wasmConstant = /EXPECTED_CORE_WASM_SHA256\s*=\s*\n?\s*["']([0-9a-f]{64})["']/u.exec(runtimeText)?.[1];
const fontConstant = /EXPECTED_MATH_FONT_SHA256\s*=\s*\n?\s*["']([0-9a-f]{64})["']/u.exec(runtimeText)?.[1];
assert.equal(wasmConstant, sourceWasmRecord.sha256, "runtime expected WASM SHA does not match source WASM");
assert.equal(fontConstant, sourceFontRecord.sha256, "runtime expected font SHA does not match source font");

async function allTextFiles(dist) {
  const names = await readdir(resolve(dist, "assets"));
  const paths = [];
  for (const candidate of [
    resolve(dist, "index.html"),
    resolve(dist, "office-dialog.html"),
    resolve(dist, "dialog", "index.html"),
  ]) {
    try { if ((await stat(candidate)).isFile()) paths.push(candidate); } catch {}
  }
  for (const name of names) if (/\.(?:js|css)$/u.test(name)) paths.push(resolve(dist, "assets", name));
  return paths;
}
async function references(dist, assetName) {
  const refs = [];
  for (const file of await allTextFiles(dist)) {
    const text = await readFile(file, "utf8");
    if (text.includes(assetName)) refs.push(file);
  }
  return refs;
}
for (const [dist, wasmPath, fontPath, label] of [
  [desktopDist, desktopWasm, desktopFont, "desktop"],
  [officeDist, officeWasm, officeFont, "office"],
]) {
  const wasmName = basename(wasmPath);
  const fontName = basename(fontPath);
  assert.ok((await references(dist, wasmName)).length > 0, `${label} JS/CSS/HTML does not reference current WASM asset`);
  assert.ok((await references(dist, fontName)).length > 0, `${label} JS/CSS/HTML does not reference current font asset`);
}

const tempRoot = await mkdtemp(join(tmpdir(), "VisualTeX P14 资源路径-中文-"));
const copied = resolve(tempRoot, "Office 静态资源 with spaces");
await cp(officeDist, copied, { recursive: true });
const copiedWasm = resolve(copied, "assets", basename(officeWasm));
const copiedFont = resolve(copied, "assets", basename(officeFont));
assert.equal((await record(copiedWasm)).sha256, sourceWasmRecord.sha256);
assert.equal((await record(copiedFont)).sha256, sourceFontRecord.sha256);

const mime = new Map([
  [".html", "text/html; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".wasm", "application/wasm"],
  [".otf", "font/otf"],
]);
const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url ?? "/", "http://127.0.0.1");
    const clean = decodeURIComponent(url.pathname).replace(/^\/+/, "");
    const path = resolve(copied, clean || "dialog/index.html");
    if (!path.startsWith(copied)) { res.writeHead(403); res.end(); return; }
    const bytes = await readFile(path);
    res.writeHead(200, { "content-type": mime.get(extname(path)) ?? "application/octet-stream", "content-length": String(bytes.length) });
    res.end(bytes);
  } catch {
    res.writeHead(404); res.end();
  }
});
await new Promise((resolveListen) => server.listen(0, "127.0.0.1", resolveListen));
try {
  const address = server.address();
  assert.ok(address && typeof address === "object");
  const origin = `http://127.0.0.1:${address.port}`;
  const wasmResponse = await fetch(`${origin}/assets/${encodeURIComponent(basename(officeWasm))}`);
  const fontResponse = await fetch(`${origin}/assets/${encodeURIComponent(basename(officeFont))}`);
  assert.equal(wasmResponse.status, 200);
  assert.equal(wasmResponse.headers.get("content-type"), "application/wasm");
  assert.equal(sha(Buffer.from(await wasmResponse.arrayBuffer())), sourceWasmRecord.sha256);
  assert.equal(fontResponse.status, 200);
  assert.match(fontResponse.headers.get("content-type") ?? "", /font\/otf/u);
  assert.equal(sha(Buffer.from(await fontResponse.arrayBuffer())), sourceFontRecord.sha256);
} finally {
  await new Promise((resolveClose) => server.close(resolveClose));
  await rm(tempRoot, { recursive: true, force: true });
}

console.log(JSON.stringify({
  status: "PASS",
  source: { wasm: sourceWasmRecord, font: sourceFontRecord },
  desktop: { wasm: desktopWasmRecord, font: desktopFontRecord },
  office: { wasm: officeWasmRecord, font: officeFontRecord },
  runtimeExpectedHashesMatch: true,
  desktopAndOfficeBytesMatchSource: true,
  oneHashedAssetPerResourcePerHost: true,
  hashedAssetReferencesPresent: true,
  nonAsciiSpacePathRoundTrip: true,
  mime: { wasm: "application/wasm", font: "font/otf" },
}, null, 2));

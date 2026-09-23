import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawn } from "node:child_process";
import { existsSync, readdirSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { join, resolve } from "node:path";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const root = process.cwd();
const executable = resolve(root, "src-tauri/target/debug/visualtex.exe");
const assetsDir = resolve(root, "dist-office-windows-native/assets");
const wasmName = readdirSync(assetsDir).find((name) => /^visualtex_engine_wasm_bg-.*\.wasm$/u.test(name));
const fontName = readdirSync(assetsDir).find((name) => /^latinmodern-math-.*\.otf$/u.test(name));
if (!existsSync(executable)) throw new Error(`Debug companion is missing: ${executable}`);
if (!wasmName || !fontName) throw new Error("Office production bundle is missing core resources.");

const base = "https://127.0.0.1:43127";
const sha = (bytes) => createHash("sha256").update(bytes).digest("hex");
const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

async function fetchHealth() {
  try {
    const response = await fetch(`${base}/health`, { cache: "no-store" });
    if (!response.ok) return null;
    const body = await response.json();
    return body?.ok === true ? body : null;
  } catch {
    return null;
  }
}

if (await fetchHealth()) {
  throw new Error("Refusing to run: an Office companion is already listening on 43127.");
}

const child = spawn(executable, ["--office-background"], {
  cwd: resolve(root, "src-tauri/target/debug"),
  stdio: "ignore",
  windowsHide: true,
});
let exit = null;
child.once("exit", (code, signal) => {
  exit = { code, signal, at: Date.now() };
});

try {
  const deadline = Date.now() + 20_000;
  let health = null;
  while (!(health = await fetchHealth())) {
    if (exit) throw new Error(`Test companion exited before health: ${JSON.stringify(exit)}`);
    if (Date.now() >= deadline) throw new Error("Timed out waiting for test companion health.");
    await sleep(100);
  }

  async function probe(name, expectedTypes) {
    const response = await fetch(`${base}/assets/${encodeURIComponent(name)}`, { cache: "no-store" });
    assert.equal(response.status, 200, `${name} returned HTTP ${response.status}`);
    const contentType = response.headers.get("content-type") ?? "";
    const normalizedType = contentType.split(";", 1)[0].trim().toLowerCase();
    assert.ok(expectedTypes.includes(normalizedType), `${name} has unexpected MIME: ${contentType}`);
    assert.equal(response.headers.get("x-content-type-options"), "nosniff", `${name} is missing nosniff`);
    assert.equal(response.headers.get("cross-origin-resource-policy"), "same-origin", `${name} is missing same-origin CORP`);
    const bytes = Buffer.from(await response.arrayBuffer());
    const disk = await readFile(join(assetsDir, name));
    assert.equal(sha(bytes), sha(disk), `${name} bytes differ between HTTPS companion and production bundle`);
    return {
      name,
      contentType,
      cacheControl: response.headers.get("cache-control"),
      bytes: bytes.length,
      sha256: sha(bytes),
    };
  }

  const wasm = await probe(wasmName, ["application/wasm"]);
  const font = await probe(fontName, ["application/font-sfnt", "font/otf"]);
  assert.equal(wasm.cacheControl, null, "hashed WASM asset should remain cacheable instead of inheriting no-store API policy");
  assert.equal(font.cacheControl, null, "hashed font asset should remain cacheable instead of inheriting no-store API policy");

  console.log(JSON.stringify({
    status: "PASS",
    companionPid: child.pid,
    health: {
      appVersion: health.appVersion,
      officeUiVersion: health.officeUiVersion,
      protocolVersion: health.protocolVersion,
    },
    wasm,
    font,
    correctMime: true,
    productionBytesMatchHttps: true,
    hashedAssetsCacheable: true,
    securityHeaders: true,
    wordAutomation: 0,
    officeProcessActions: 0,
  }, null, 2));
} finally {
  if (!exit) {
    child.kill();
    const stopDeadline = Date.now() + 5_000;
    while (!exit && Date.now() < stopDeadline) await sleep(50);
  }
}

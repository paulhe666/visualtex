import { spawnSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const rounds = Number(process.argv[2] ?? 30);
const pid = Number(process.argv[3] ?? 38508);
const base = "https://127.0.0.1:43127";
const root = process.cwd();
const enumScript = resolve(root, "engine/tools/p13-enum-process-windows.ps1");
const closeScript = resolve(root, "engine/tools/p13-close-owned-converter-window.ps1");
const installPath = resolve(process.env.APPDATA ?? "", "com.visualtex.studio/office/install.json");
const install = JSON.parse(await readFile(installPath, "utf8"));
if (typeof install.installToken !== "string" || install.installToken.length !== 64) {
  throw new Error("Office install token is missing or invalid.");
}
const headers = {
  "content-type": "application/json",
  "X-VisualTeX-Install-Token": install.installToken,
};
const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

async function request(path, options = {}) {
  const response = await fetch(`${base}${path}`, options);
  const text = await response.text();
  if (!response.ok) throw new Error(`${options.method ?? "GET"} ${path} -> ${response.status}: ${text}`);
  return text ? JSON.parse(text) : null;
}

async function healthy() {
  try {
    const body = await request("/health");
    return body?.ok === true;
  } catch {
    return false;
  }
}

function enumerateWindows() {
  const result = spawnSync("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", enumScript,
    "-ProcessId", String(pid),
  ], { encoding: "utf8", windowsHide: true });
  if (result.status !== 0) throw new Error(`Window enumeration failed: ${result.stderr || result.stdout}`);
  return JSON.parse(result.stdout.trim());
}

function converterWindow() {
  const inventory = enumerateWindows();
  const matches = inventory.windows.filter((window) =>
    window.class === "Tauri Window" &&
    window.title.startsWith("VisualTeX") &&
    window.title.includes("Office"),
  );
  if (matches.length !== 1) {
    throw new Error(`Expected exactly one converter window, found ${matches.length}: ${JSON.stringify(inventory)}`);
  }
  return matches[0];
}

function closeConverter(handle) {
  const result = spawnSync("powershell.exe", [
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", closeScript,
    "-ProcessId", String(pid),
    "-WindowHandle", String(handle),
  ], { encoding: "utf8", windowsHide: true });
  if (result.status !== 0) throw new Error(`Converter close failed: ${result.stderr || result.stdout}`);
  const body = JSON.parse(result.stdout.trim());
  if (body.stillExists) throw new Error(`Converter HWND ${handle} did not close.`);
}

if (!(await healthy())) throw new Error("Companion is not healthy before lifecycle stress.");
let previousHandle = null;
for (let round = 1; round <= rounds; round += 1) {
  const before = converterWindow();
  closeConverter(before.handle);
  await sleep(100);
  if (!(await healthy())) throw new Error(`Companion became unhealthy after close at round ${round}.`);

  await request("/api/v1/app/converter/prewarm", {
    method: "POST",
    headers,
    body: "{}",
  });
  await sleep(150);
  if (!(await healthy())) throw new Error(`Companion became unhealthy after recreate at round ${round}.`);
  const after = converterWindow();
  if (after.handle === before.handle) {
    throw new Error(`Round ${round} reused a closed HWND unexpectedly: ${after.handle}`);
  }
  console.log(JSON.stringify({
    round,
    beforeHandle: before.handle,
    afterHandle: after.handle,
    changedFromPrevious: previousHandle === null || after.handle !== previousHandle,
    alive: true,
  }));
  previousHandle = after.handle;
}

console.log(JSON.stringify({ status: "PASS", rounds, pid, alive: await healthy() }));

import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const root = process.cwd();
const executable = resolve(root, "apps/windows/src-tauri/target/debug/visualtex.exe");
const installPath = resolve(process.env.APPDATA ?? "", "com.visualtex.studio/office/install.json");
const base = "https://127.0.0.1:43127";
const total = Number(process.argv[2] ?? 1000);
const chunkSize = 256;
const sleep = (ms) => new Promise((resolveSleep) => setTimeout(resolveSleep, ms));

async function request(path, options = {}) {
  const response = await fetch(`${base}${path}`, options);
  const text = await response.text();
  let body = null;
  if (text) {
    try { body = JSON.parse(text); } catch { body = text; }
  }
  if (!response.ok) {
    throw new Error(`${options.method ?? "GET"} ${path} -> ${response.status}: ${typeof body === "string" ? body : JSON.stringify(body)}`);
  }
  return body;
}

async function health() {
  try {
    const value = await request("/health");
    return value?.ok === true;
  } catch {
    return false;
  }
}

if (await health()) throw new Error("Refusing to run while an existing Office companion owns port 43127.");
const install = JSON.parse(await readFile(installPath, "utf8"));
if (typeof install.installToken !== "string" || install.installToken.length !== 64) {
  throw new Error("Office install token is missing or invalid.");
}
const headers = {
  "content-type": "application/json",
  "X-VisualTeX-Install-Token": install.installToken,
};

const child = spawn(executable, ["--office-background"], {
  cwd: resolve(root, "apps/windows/src-tauri/target/debug"),
  stdio: "ignore",
  windowsHide: true,
});
let exit = null;
child.once("exit", (code, signal) => { exit = { code, signal, at: Date.now() }; });

async function assertAlive(stage) {
  if (exit || !(await health())) throw new Error(`Companion unavailable at ${stage}: ${JSON.stringify(exit)}`);
}

try {
  const startupDeadline = Date.now() + 15_000;
  while (!(await health())) {
    if (exit) throw new Error(`Companion exited during startup: ${JSON.stringify(exit)}`);
    if (Date.now() >= startupDeadline) throw new Error("Timed out waiting for Office companion health.");
    await sleep(100);
  }
  await request("/api/v1/app/converter/prewarm", { method: "POST", headers, body: "{}" });
  await sleep(750);

  const ids = [];
  const objectModes = ["nativeOle", "wordOmml", "mathTypeOle", "crossPlatformPicture"];
  for (let index = 0; index < total; index += 1) {
    const lineId = randomUUID();
    const objectMode = objectModes[index % objectModes.length];
    const created = await request("/api/v1/sessions", {
      method: "POST",
      headers,
      body: JSON.stringify({
        mode: "create",
        host: objectMode === "crossPlatformPicture" ? "powerpoint" : "word",
        formulaId: randomUUID(),
        sourceDocumentId: null,
        sourceObjectId: null,
        title: `P13 companion burst ${index}`,
        lines: [{ id: lineId, latex: index % 4 === 0 ? `x_{${index}}^2+y_{${index}}^2` : index % 4 === 1 ? `\\frac{a_${index}+b_${index}}{c_${index}+d_${index}}` : index % 4 === 2 ? `\\sqrt{x_${index}^2+y_${index}^2}` : `\\sum_{k=1}^{n_${index}}k` }],
        activeLineId: lineId,
        codeFormat: "latex",
        displayMode: index % 2 === 0 ? "inline" : "block",
        objectMode,
        numbered: false,
        mathTypeNumberPosition: "right",
        fontSizePt: 10.5,
        originalMetadata: null,
        autoCommitOnClose: false,
      }),
    });
    ids.push(created.id);
    if (index % 128 === 0) await assertAlive(`create-${index}`);
  }
  console.log(JSON.stringify({ stage: "created", total: ids.length, companionPid: child.pid }));

  const chunks = [];
  for (let index = 0; index < ids.length; index += chunkSize) chunks.push(ids.slice(index, index + chunkSize));
  for (const [index, chunk] of chunks.entries()) {
    await request("/api/v1/app/converter/convert-batch", {
      method: "POST",
      headers,
      body: JSON.stringify({ sessionIds: chunk }),
    });
    console.log(JSON.stringify({ stage: "queued", chunk: index + 1, chunks: chunks.length, size: chunk.length }));
  }

  const terminalCounts = { committing: 0, completed: 0, failed: 0, cancelled: 0 };
  const deadline = Date.now() + 180_000;
  for (let index = 0; index < ids.length; index += 1) {
    const id = ids[index];
    while (true) {
      if (Date.now() >= deadline) throw new Error(`Timed out waiting for session ${index + 1}/${ids.length}.`);
      await assertAlive(`session-${index + 1}`);
      const session = await request(`/api/v1/sessions/${encodeURIComponent(id)}`, { headers });
      if (["committing", "completed", "failed", "cancelled"].includes(session.status)) {
        terminalCounts[session.status] += 1;
        if (session.status === "failed") throw new Error(`Converter session ${session.id} failed: ${session.error ?? "unknown error"}`);
        break;
      }
      await sleep(16);
    }
    if ((index + 1) % 100 === 0 || index + 1 === ids.length) {
      console.log(JSON.stringify({ stage: "terminal", completed: index + 1, total: ids.length, terminalCounts }));
    }
  }

  await sleep(5000);
  await assertAlive("settle");
  for (const chunk of chunks) {
    await request("/api/v1/sessions/batch-delete", {
      method: "POST",
      headers,
      body: JSON.stringify({ sessionIds: chunk }),
    });
  }
  await sleep(2000);
  await assertAlive("post-delete");
  console.log(JSON.stringify({ status: "PASS", total, chunks: chunks.map((chunk) => chunk.length), terminalCounts, companionPid: child.pid }));
} finally {
  if (!exit) {
    try { child.kill(); } catch {}
    await Promise.race([
      new Promise((resolveExit) => child.once("exit", resolveExit)),
      sleep(2000),
    ]);
  }
}

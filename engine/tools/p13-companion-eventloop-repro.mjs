import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";

const root = process.cwd();
const executable = resolve(root, "apps/windows/src-tauri/target/debug/visualtex.exe");
const installPath = resolve(process.env.APPDATA ?? "", "com.visualtex.studio/office/install.json");
const base = "https://127.0.0.1:43127";
const rounds = Number(process.argv[2] ?? 12);
const batchSize = Number(process.argv[3] ?? 32);
const reuseExisting = process.env.P13_REUSE_EXISTING_COMPANION === "1";

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
    const result = await request("/health");
    return result?.ok === true;
  } catch {
    return false;
  }
}

const initiallyHealthy = await health();
if (initiallyHealthy && !reuseExisting) {
  throw new Error("Refusing to run: an Office companion is already listening on 43127.");
}
if (!initiallyHealthy && reuseExisting) {
  throw new Error("P13_REUSE_EXISTING_COMPANION=1 requires an already healthy Office companion.");
}

const install = JSON.parse(await readFile(installPath, "utf8"));
if (typeof install.installToken !== "string" || install.installToken.length !== 64) {
  throw new Error("Office install token is missing or invalid.");
}
const headers = {
  "content-type": "application/json",
  "X-VisualTeX-Install-Token": install.installToken,
};

const child = reuseExisting ? null : spawn(executable, ["--office-background"], {
  cwd: resolve(root, "apps/windows/src-tauri/target/debug"),
  stdio: "ignore",
  windowsHide: true,
});
let exit = null;
child?.once("exit", (code, signal) => { exit = { code, signal, at: Date.now() }; });

try {
  const startDeadline = Date.now() + 15_000;
  while (!(await health())) {
    if (exit) throw new Error(`Companion exited before health became ready: ${JSON.stringify(exit)}`);
    if (Date.now() >= startDeadline) throw new Error("Timed out waiting for Office companion health.");
    await sleep(100);
  }

  await request("/api/v1/app/converter/prewarm", {
    method: "POST",
    headers,
    body: "{}",
  });
  await sleep(750);

  let created = 0;
  let completed = 0;
  for (let round = 0; round < rounds; round += 1) {
    const ids = [];
    for (let index = 0; index < batchSize; index += 1) {
      const sessionId = randomUUID();
      const formulaId = randomUUID();
      const lineId = randomUUID();
      const objectModes = ["nativeOle", "wordOmml", "mathTypeOle", "crossPlatformPicture"];
      const objectMode = objectModes[(round * batchSize + index) % objectModes.length];
      const displayMode = index % 2 === 0 ? "inline" : "block";
      const body = await request("/api/v1/sessions", {
        method: "POST",
        headers,
        body: JSON.stringify({
          mode: "create",
          host: objectMode === "crossPlatformPicture" ? "powerpoint" : "word",
          formulaId,
          sourceDocumentId: null,
          sourceObjectId: null,
          title: `P13 companion repro ${round}-${index}`,
          lines: [{ id: lineId, latex: index % 3 === 0 ? "x^2+y^2" : index % 3 === 1 ? "\\frac{a+b}{c+d}" : "\\sqrt{x_1^2+x_2^2}" }],
          activeLineId: lineId,
          codeFormat: "latex",
          displayMode,
          objectMode,
          numbered: false,
          mathTypeNumberPosition: "right",
          fontSizePt: 10.5,
          originalMetadata: null,
          autoCommitOnClose: false,
        }),
      });
      ids.push(body.id ?? sessionId);
      created += 1;
    }

    await request("/api/v1/app/converter/convert-batch", {
      method: "POST",
      headers,
      body: JSON.stringify({ sessionIds: ids }),
    });

    const pending = new Set(ids);
    const deadline = Date.now() + 45_000;
    while (pending.size) {
      if (exit || !(await health())) {
        throw new Error(`Companion crashed while batch ${round + 1}/${rounds} was active: ${JSON.stringify(exit)}`);
      }
      if (Date.now() >= deadline) {
        throw new Error(`Timed out waiting for batch ${round + 1}; pending=${pending.size}`);
      }
      for (const id of [...pending]) {
        const session = await request(`/api/v1/sessions/${encodeURIComponent(id)}`, { headers });
        if (["committing", "completed", "failed", "cancelled"].includes(session.status)) {
          if (session.status === "failed") {
            throw new Error(`Converter session ${id} failed: ${session.error ?? "unknown error"}`);
          }
          pending.delete(id);
          completed += 1;
        }
      }
      if (pending.size) await sleep(30);
    }

    await request("/api/v1/sessions/batch-delete", {
      method: "POST",
      headers,
      body: JSON.stringify({ sessionIds: ids }),
    });
    await sleep(250);
    if (exit || !(await health())) {
      throw new Error(`Companion crashed after deleting batch ${round + 1}/${rounds}: ${JSON.stringify(exit)}`);
    }

    console.log(JSON.stringify({
      round: round + 1,
      rounds,
      batchSize,
      created,
      completed,
      companionPid: child?.pid ?? "existing",
      alive: !exit && await health(),
    }));
  }

  await sleep(Number(process.argv[4] ?? 0));
  if (exit || !(await health())) {
    throw new Error(`Companion crashed during the post-batch settle window: ${JSON.stringify(exit)}`);
  }

  console.log(JSON.stringify({
    status: "PASS",
    rounds,
    batchSize,
    sessions: created,
    completed,
    companionPid: child?.pid ?? "existing",
    crashed: Boolean(exit),
  }));
} finally {
  if (!reuseExisting && child && !exit) {
    try { child.kill(); } catch {}
    await Promise.race([
      new Promise((resolveExit) => child.once("exit", resolveExit)),
      sleep(2000),
    ]);
  }
}

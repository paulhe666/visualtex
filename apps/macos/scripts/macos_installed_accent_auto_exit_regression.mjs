import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const installedApp = "/Applications/VisualTeX.app";
const candidateApp = join(
  process.cwd(),
  "src-tauri/target/release/bundle/macos/VisualTeX.app",
);
const backupApp = join(tmpdir(), `VisualTeX-accent-backup-${process.pid}.app`);
const axProbeSource = join(process.cwd(), "scripts/macos_ax_ime_probe.swift");
const axProbeBinary = join(tmpdir(), `visualtex-accent-ax-${process.pid}`);
const commandContentCase = process.argv.includes("--command-content");
const nestedCase = process.argv.includes("--nested");

function run(program, args, options = {}) {
  return execFileSync(program, args, {
    encoding: "utf8",
    timeout: options.timeout ?? 30_000,
    maxBuffer: 16 * 1024 * 1024,
    ...options,
  }).trim();
}

function bestEffort(program, args, options = {}) {
  try {
    return run(program, args, options);
  } catch {
    return "";
  }
}

function jxa(source) {
  return run("/usr/bin/osascript", ["-l", "JavaScript", "-e", source]);
}

function appPids() {
  const output = bestEffort("/usr/bin/pgrep", [
    "-f",
    "/Applications/VisualTeX.app/Contents/MacOS/visualtex",
  ]);
  return output
    .split(/\s+/)
    .map((value) => Number.parseInt(value, 10))
    .filter((value) => Number.isInteger(value) && value > 0);
}

async function stopInstalledApp() {
  bestEffort("/usr/bin/osascript", [
    "-e",
    'tell application id "com.visualtex.studio" to quit',
  ], { timeout: 10_000 });
  const deadline = Date.now() + 5_000;
  while (Date.now() < deadline && appPids().length) await sleep(100);
  for (const pid of appPids()) {
    try { process.kill(pid, "SIGTERM"); } catch {}
  }
  await sleep(500);
  for (const pid of appPids()) {
    try { process.kill(pid, "SIGKILL"); } catch {}
  }
}

async function waitForMainWindow(timeoutMs = 15_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const snapshot = appSnapshot();
      if (
        snapshot.pid &&
        (snapshot.fields.length || snapshot.areas.length || snapshot.buttons.length)
      ) {
        return snapshot;
      }
    } catch {}
    await sleep(150);
  }
  const ps = bestEffort("/bin/ps", ["-axo", "pid=,command="])
    .split("\n")
    .filter((line) => line.includes("visualtex"))
    .join("\n");
  const ax = bestEffort("/usr/bin/osascript", [
    "-l", "JavaScript", "-e",
    `(() => { const se = Application("System Events"); return JSON.stringify(se.applicationProcesses().map((p) => { let name="", bundle="", pid=null, windows=0; try{name=String(p.name()||"")}catch{}; try{bundle=String(p.bundleIdentifier()||"")}catch{}; try{pid=p.unixId()}catch{}; try{windows=p.windows().length}catch{}; return {name,bundle,pid,windows}; }).filter((x) => x.name.toLowerCase().includes("visualtex") || x.bundle.includes("visualtex"))); })()`
  ]);
  throw new Error(`Timed out waiting for VisualTeX main WKWebView; ps=${JSON.stringify(ps)}; ax=${ax}`);
}

async function waitForFormulaFields(timeoutMs = 10_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    try {
      const snapshot = appSnapshot();
      if (formulaFields(snapshot).length) return snapshot;
    } catch {}
    await sleep(150);
  }
  throw new Error("Timed out waiting for VisualTeX formula fields");
}

function appSnapshot() {
  const pid = appPids()[0];
  if (!Number.isInteger(pid)) throw new Error("VisualTeX process missing");
  const payload = run(axProbeBinary, ["snapshot", String(pid)], { timeout: 10_000 });
  const raw = JSON.parse(payload);
  return {
    pid,
    fields: (raw.fields ?? []).map((item) => ({
      ...item,
      name: item.title ?? "",
      desc: item.description ?? "",
      pos: item.position ?? null,
      size: item.size ?? null,
    })),
    areas: (raw.areas ?? []).map((item) => ({
      ...item,
      name: item.title ?? "",
      desc: item.description ?? "",
      pos: item.position ?? null,
      size: item.size ?? null,
    })),
    buttons: (raw.buttons ?? []).map((item) => ({
      ...item,
      name: item.title ?? "",
      desc: item.description ?? "",
      pos: item.position ?? null,
      size: item.size ?? null,
    })),
  };
}

function formulaFields(snapshot) {
  return snapshot.fields.filter((field) =>
    field.desc !== "公式文档标题" &&
    field.desc !== "Formula document title" &&
    field.name !== "公式文档标题" &&
    field.name !== "Formula document title"
  );
}

function sourceText(snapshot) {
  return snapshot.areas.find((area) => area.value.includes("$$"))?.value ?? "";
}

function pressButtonByName(names) {
  return jxa(`(() => {
    const se = Application("System Events");
    const process = se.applicationProcesses.whose({ bundleIdentifier: { _equals: "com.visualtex.studio" } })()[0];
    const wanted = ${JSON.stringify(names)};
    const walk = (element, depth = 0) => {
      if (depth > 18) return false;
      let role = "", name = "", desc = "";
      try { role = String(element.role() || ""); } catch {}
      try { name = String(element.name() || ""); } catch {}
      try { desc = String(element.description() || ""); } catch {}
      if (role === "AXButton" && wanted.some((value) => value === name || value === desc)) {
        try { element.actions.byName("AXPress").perform(); return true; } catch {}
        try { element.click(); return true; } catch {}
      }
      let children = [];
      try { children = element.uiElements(); } catch {}
      for (const child of children) if (walk(child, depth + 1)) return true;
      return false;
    };
    for (const window of process.windows()) if (walk(window, 0)) return "true";
    return "false";
  })()`);
}

function clickAxItem(field, pid) {
  assert.ok(Array.isArray(field.pos) && Array.isArray(field.size), JSON.stringify(field));
  const x = Math.round(field.pos[0] + Math.max(8, Math.min(field.size[0] / 2, 28)));
  const y = Math.round(field.pos[1] + Math.max(6, field.size[1] / 2));
  run("/usr/bin/osascript", [
    "-e", 'tell application "System Events"',
    "-e", `set frontmost of first application process whose unix id is ${pid} to true`,
    "-e", `click at {${x}, ${y}}`,
    "-e", "end tell",
  ]);
}

function clickField(field, pid) {
  clickAxItem(field, pid);
}

function focusLastFormula(pid) {
  const result = JSON.parse(run(axProbeBinary, ["focus-last", String(pid)], { timeout: 10_000 }));
  assert.equal(result.result, 0, `AX focus-last failed: ${JSON.stringify(result)}`);
  return result;
}

function keys(pid, lines) {
  const args = ["-e", 'tell application "System Events"', "-e", `set frontmost of first application process whose unix id is ${pid} to true`, "-e", "delay 0.08"];
  for (const line of lines) args.push("-e", line);
  args.push("-e", "end tell");
  run("/usr/bin/osascript", args);
}

function inputSourceKind() {
  const value = bestEffort("/usr/bin/defaults", ["read", "com.apple.HIToolbox", "AppleSelectedInputSources"]);
  if (value.includes("com.apple.inputmethod.SCIM.ITABC")) return "chinese";
  if (value.includes("KeyboardLayout Name") && value.includes("ABC")) return "abc";
  return "unknown";
}

async function ensureInputSource(target) {
  for (let attempt = 0; attempt < 4; attempt += 1) {
    if (inputSourceKind() === target) return;
    run("/usr/bin/osascript", [
      "-e", 'tell application "System Events" to key code 49 using {control down}',
    ]);
    await sleep(400);
  }
  throw new Error(`Could not switch input source to ${target}; current=${inputSourceKind()}`);
}

async function ensureSourcePanel() {
  let snapshot = appSnapshot();
  if (sourceText(snapshot)) return snapshot;
  const button = snapshot.buttons.find((item) =>
    ["LaTeX 源码", "LaTeX source"].includes(item.name) ||
    ["LaTeX 源码", "LaTeX source"].includes(item.desc)
  );
  if (!button) throw new Error(`LaTeX source button was not found: ${JSON.stringify(snapshot.buttons.slice(-20))}`);
  clickAxItem(button, snapshot.pid);
  await sleep(400);
  snapshot = appSnapshot();
  if (!sourceText(snapshot)) throw new Error(`LaTeX source panel did not open: ${JSON.stringify(snapshot.buttons.slice(-20))}`);
  return snapshot;
}

async function createScratchLine(snapshot) {
  const fields = formulaFields(snapshot);
  const last = fields.at(-1);
  assert.ok(last, "No formula field found");
  focusLastFormula(snapshot.pid);
  await sleep(120);
  keys(snapshot.pid, ['key code 124 using {command down}', 'key code 36']);
  await sleep(350);
  const next = appSnapshot();
  assert.equal(formulaFields(next).length, fields.length + 1, `Scratch line was not created: ${JSON.stringify(next.fields)}`);
  const scratch = formulaFields(next).at(-1);
  assert.ok(scratch, "Scratch field missing");
  focusLastFormula(next.pid);
  await sleep(160);
  return next;
}

async function removeScratchLine(expectedCount) {
  let snapshot = appSnapshot();
  const scratch = formulaFields(snapshot).at(-1);
  if (scratch) {
    focusLastFormula(snapshot.pid);
    await sleep(80);
    keys(snapshot.pid, ['keystroke "a" using {command down}', 'key code 51']);
    await sleep(220);
    snapshot = appSnapshot();
    if (formulaFields(snapshot).length > expectedCount) {
      const empty = formulaFields(snapshot).at(-1);
      if (empty) {
        focusLastFormula(snapshot.pid);
        keys(snapshot.pid, ['key code 51']);
        await sleep(250);
      }
    }
  }
  snapshot = appSnapshot();
  assert.equal(formulaFields(snapshot).length, expectedCount, "Scratch field cleanup changed formula count");
}

assert.ok(existsSync(installedApp), `Installed app missing: ${installedApp}`);
assert.ok(existsSync(candidateApp), `Candidate app missing: ${candidateApp}`);
assert.ok(existsSync(axProbeSource), `AX probe missing: ${axProbeSource}`);
run("/usr/bin/swiftc", ["-O", "-o", axProbeBinary, axProbeSource], { timeout: 120_000 });
rmSync(backupApp, { recursive: true, force: true });
const originalInput = inputSourceKind();
let baselineCount = 0;
let candidateWasInstalled = false;
const originalClipboard = bestEffort("/usr/bin/pbpaste", []);

try {
  run("/usr/bin/ditto", [installedApp, backupApp], { timeout: 60_000 });
  await stopInstalledApp();
  rmSync(installedApp, { recursive: true, force: true });
  run("/usr/bin/ditto", [candidateApp, installedApp], { timeout: 60_000 });
  candidateWasInstalled = true;
  // The Office LaunchAgent can relaunch --office-background while the app is
  // being replaced. Kill once more after the candidate bytes are complete so
  // the main-window acceptance cannot attach to a background-only resident.
  await stopInstalledApp();
  await sleep(500);
  run("/usr/bin/open", ["-na", installedApp]);
  bestEffort("/usr/bin/osascript", [
    "-e", 'tell application id "com.visualtex.studio" to activate',
  ]);
  await sleep(500);
  await waitForMainWindow();
  await sleep(500);

  let startupSnapshot = appSnapshot();
  if (!formulaFields(startupSnapshot).length) {
    const dismiss = startupSnapshot.buttons.find((item) =>
      ["知道了", "Got it"].includes(item.name) ||
      ["知道了", "Got it"].includes(item.desc)
    );
    if (dismiss) clickAxItem(dismiss, startupSnapshot.pid);
    else keys(startupSnapshot.pid, ['key code 53']);
    await sleep(350);
  }
  await waitForFormulaFields();

  let snapshot = appSnapshot();
  baselineCount = formulaFields(snapshot).length;
  assert.ok(baselineCount >= 1, "Baseline formula field is missing");

  snapshot = await createScratchLine(snapshot);
  await ensureInputSource("abc");
  const pid = snapshot.pid;
  if (nestedCase) {
    keys(pid, ['keystroke "("', 'keystroke "a+"', 'keystroke "("']);
    await sleep(250);
  }
  keys(pid, ['key code 42', 'keystroke "vec"', 'key code 49']);
  await sleep(300);
  if (commandContentCase) {
    keys(pid, ['key code 42', 'keystroke "nabla"', 'key code 49']);
    await sleep(350);
  } else {
    keys(pid, ['keystroke "B"']);
    await sleep(350);
  }
  keys(pid, ['keystroke "q"']);
  if (nestedCase) {
    await sleep(180);
    keys(pid, ['keystroke ")"', 'keystroke ")"']);
  }
  await sleep(450);

  const after = appSnapshot();
  const scratch = formulaFields(after).at(-1);
  assert.ok(scratch, "Scratch field disappeared before clipboard verification");
  focusLastFormula(pid);
  await sleep(100);
  keys(pid, ['keystroke "a" using {command down}', 'keystroke "c" using {command down}']);
  await sleep(250);
  const copied = run("/usr/bin/pbpaste", []);
  if (commandContentCase) {
    assert.match(
      copied,
      /\\vec\{\\nabla\}q/,
      `WKWebView command-content accent auto-exit failed: ${JSON.stringify({ copied, fields: formulaFields(after) })}`,
    );
    assert.doesNotMatch(
      copied,
      /\\vec\{\\nabla q\}/,
      `Input after \\nabla remained inside the accent: ${JSON.stringify(copied)}`,
    );
  } else {
    assert.match(
      copied,
      /\\vec\{B\}q/,
      `WKWebView physical-key accent auto-exit failed: ${JSON.stringify({ copied, fields: formulaFields(after) })}`,
    );
    assert.doesNotMatch(
      copied,
      /\\vec\{Bq\}/,
      `Second character remained inside the accent: ${JSON.stringify(copied)}`,
    );
  }
  console.log(JSON.stringify({ status: "PASS", case: `${nestedCase ? "nested-" : ""}${commandContentCase ? "command-content" : "single-character"}`, copied, pid }, null, 2));

  await removeScratchLine(baselineCount);
} finally {
  if (candidateWasInstalled) await stopInstalledApp();
  if (existsSync(backupApp)) {
    rmSync(installedApp, { recursive: true, force: true });
    run("/usr/bin/ditto", [backupApp, installedApp], { timeout: 60_000 });
    rmSync(backupApp, { recursive: true, force: true });
    bestEffort("/usr/bin/open", [installedApp]);
  }
  if ((originalInput === "abc" || originalInput === "chinese") && inputSourceKind() !== originalInput) {
    await ensureInputSource(originalInput);
  }
  bestEffort("/usr/bin/pbcopy", [], { input: originalClipboard });
  rmSync(axProbeBinary, { force: true });
}

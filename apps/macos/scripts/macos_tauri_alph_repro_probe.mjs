import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { rmSync } from "node:fs";

const pidArg = process.argv.find((arg) => arg.startsWith("--pid="));
const appArg = process.argv.find((arg) => arg.startsWith("--app="));
const pid = Number(pidArg?.slice(6));
const appPath = appArg?.slice(6) ?? "";
if (!Number.isInteger(pid) || pid <= 0 || !appPath) throw new Error("Usage: node scripts/macos_tauri_alph_repro_probe.mjs --pid=<visualtex-pid> --app=<app-path>");
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function run(command, args, options = {}) {
  const result = spawnSync(command, args, { encoding: "utf8", ...options });
  if (result.status !== 0) throw new Error(result.stderr || result.stdout || `${command} failed`);
  return result.stdout;
}

const keyHelper = `/tmp/visualtex-repro-key-${process.pid}`;
const tisHelper = `/tmp/visualtex-repro-tis-${process.pid}`;

function compile(output, source) {
  const result = spawnSync("swiftc", ["-O", "-o", output, "-"], { encoding: "utf8", input: source });
  if (result.status !== 0) throw new Error(result.stderr || result.stdout || `swiftc failed for ${output}`);
}

function buildHelpers() {
  compile(keyHelper, String.raw`
import CoreGraphics
import Foundation
let pid = pid_t(Int(CommandLine.arguments[1])!)
for token in CommandLine.arguments.dropFirst(2) {
  let parts = token.split(separator: ":", omittingEmptySubsequences: false)
  guard let raw = UInt16(parts[0]) else { continue }
  let flags: CGEventFlags = parts.count > 1 && parts[1] == "cmd" ? [.maskCommand] : []
  let down = CGEvent(keyboardEventSource: nil, virtualKey: CGKeyCode(raw), keyDown: true)!
  down.flags = flags
  down.postToPid(pid)
  usleep(35000)
  let up = CGEvent(keyboardEventSource: nil, virtualKey: CGKeyCode(raw), keyDown: false)!
  up.flags = flags
  up.postToPid(pid)
  usleep(35000)
}
`);
  compile(tisHelper, String.raw`
import Carbon
import Foundation
let target = CommandLine.arguments[1]
let filter = [kTISPropertyInputSourceID as String: target] as CFDictionary
let list = TISCreateInputSourceList(filter, false).takeRetainedValue() as! [TISInputSource]
if list.isEmpty { fputs("missing input source\\n", stderr); exit(2) }
let status = TISSelectInputSource(list[0])
if status != noErr { fputs("TISSelectInputSource failed: \\(status)\\n", stderr); exit(3) }
`);
}

const keys = (...tokens) => run(keyHelper, [String(pid), ...tokens.map(String)]);
const selectSource = (id) => run(tisHelper, [id]);

function currentSourceKind() {
  const out = run("defaults", ["read", "com.apple.HIToolbox", "AppleSelectedInputSources"]);
  if (out.includes("com.apple.inputmethod.SCIM.ITABC")) return "pinyin";
  if (out.includes("KeyboardLayout Name") && out.includes("ABC")) return "abc";
  return "other";
}
function restoreSource(kind) {
  if (kind === "pinyin") selectSource("com.apple.inputmethod.SCIM.ITABC");
  if (kind === "abc") selectSource("com.apple.keylayout.ABC");
}
function jxa(script) { return run("osascript", ["-l", "JavaScript", "-e", script]); }

function snapshot() {
  return JSON.parse(jxa(`(() => {
    const se = Application("System Events");
    const p = se.applicationProcesses.whose({ unixId: { _equals: ${pid} } })()[0];
    if (!p) throw new Error("probe process missing");
    const result = { source: "", fields: [], diagnostic: "" };
    const walk = (e, depth = 0) => {
      if (depth > 12) return;
      let role = ""; try { role = String(e.role() || ""); } catch {}
      if (role === "AXTextField") {
        let name = "", focused = false;
        try { name = String(e.name() || ""); } catch {}
        try { focused = Boolean(e.focused()); } catch {}
        if (name !== "公式文档标题" && name !== "Formula document title") result.fields.push({ name, focused });
      }
      if (role === "AXTextArea") {
        let name = "", value = "";
        try { name = String(e.name() || ""); } catch {}
        try { value = String(e.value() || ""); } catch {}
        if (name === "VisualTeX IME Diagnostic Trace") result.diagnostic = value;
        else if (value.includes("$$")) result.source = value;
      }
      let children = []; try { children = e.uiElements(); } catch {}
      for (const child of children) walk(child, depth + 1);
    };
    for (const w of p.windows()) { let name = ""; try { name = String(w.name() || ""); } catch {}; if (name === "VisualTeX") { walk(w); break; } }
    return JSON.stringify(result);
  })()`));
}

function focusLastFormulaField() {
  return JSON.parse(jxa(`(() => {
    const se = Application("System Events");
    const p = se.applicationProcesses.whose({ unixId: { _equals: ${pid} } })()[0];
    if (!p) throw new Error("probe process missing");
    const fields = [];
    const walk = (e, depth = 0) => {
      if (depth > 12) return;
      let role = ""; try { role = String(e.role() || ""); } catch {}
      if (role === "AXTextField") { let name = ""; try { name = String(e.name() || ""); } catch {}; if (name !== "公式文档标题" && name !== "Formula document title") fields.push(e); }
      let children = []; try { children = e.uiElements(); } catch {}
      for (const child of children) walk(child, depth + 1);
    };
    for (const w of p.windows()) { let name = ""; try { name = String(w.name() || ""); } catch {}; if (name === "VisualTeX") { walk(w); break; } }
    if (!fields.length) throw new Error("no formula fields");
    fields[fields.length - 1].focused = true;
    return JSON.stringify({ count: fields.length, focused: Boolean(fields[fields.length - 1].focused()) });
  })()`));
}

function lastBlock(source) {
  const matches = [...source.matchAll(/\$\$\n([\s\S]*?)\n\$\$/g)];
  return matches.at(-1)?.[1] ?? "";
}

async function createScratch() {
  const beforeCount = snapshot().fields.length;
  focusLastFormulaField();
  keys("124:cmd", "36");
  for (let i = 0; i < 20; i += 1) {
    await sleep(80);
    const after = snapshot();
    if (after.fields.length === beforeCount + 1) {
      assert.equal(lastBlock(after.source), "", `scratch row is not empty: ${JSON.stringify(lastBlock(after.source))}`);
      focusLastFormulaField();
      return;
    }
  }
  throw new Error(`failed to create scratch row: ${JSON.stringify(snapshot())}`);
}

async function runCase(timing) {
  await createScratch();
  selectSource("com.apple.inputmethod.SCIM.ITABC");
  await sleep(60);
  focusLastFormulaField();
  keys("0");
  if (timing.aToBackspaceMs) await sleep(timing.aToBackspaceMs);
  keys("51");
  if (timing.deleteToSwitchMs) await sleep(timing.deleteToSwitchMs);
  assert.equal(lastBlock(snapshot().source), "", "Pinyin a/Backspace changed the formula");
  selectSource("com.apple.keylayout.ABC");
  if (timing.switchToCommandMs) await sleep(timing.switchToCommandMs);
  focusLastFormulaField();
  if (timing.commandKeyGapMs === 0) keys("42", "0", "37", "35", "4");
  else for (const key of ["42", "0", "37", "35", "4"]) { keys(key); await sleep(timing.commandKeyGapMs); }
  await sleep(120);
  const beforeSpace = snapshot();
  keys("49");
  const samples = [];
  let elapsed = 0;
  for (const delay of [80, 220, 500, 1000]) {
    await sleep(delay); elapsed += delay;
    const block = lastBlock(snapshot().source);
    samples.push({ elapsedMs: elapsed, block, alphaCount: block.split("\\alpha").length - 1 });
  }
  return { timing, beforeSpaceBlock: lastBlock(beforeSpace.source), samples, diagnosticTail: beforeSpace.diagnostic.split("\n").slice(-40) };
}

const originalSource = currentSourceKind();
buildHelpers();
run("open", [appPath]);
await sleep(900);
const timings = [
  { aToBackspaceMs: 0, deleteToSwitchMs: 0, switchToCommandMs: 0, commandKeyGapMs: 0 },
  { aToBackspaceMs: 20, deleteToSwitchMs: 0, switchToCommandMs: 0, commandKeyGapMs: 0 },
  { aToBackspaceMs: 50, deleteToSwitchMs: 20, switchToCommandMs: 20, commandKeyGapMs: 0 },
  { aToBackspaceMs: 100, deleteToSwitchMs: 50, switchToCommandMs: 50, commandKeyGapMs: 0 },
  { aToBackspaceMs: 180, deleteToSwitchMs: 80, switchToCommandMs: 80, commandKeyGapMs: 15 },
];
try {
  const caseArg = process.argv.find((arg) => arg.startsWith("--case="));
  const caseIndex = caseArg ? Number(caseArg.slice(7)) : 0;
  const timing = timings[caseIndex];
  if (!timing) throw new Error(`Unknown --case=${caseIndex}`);
  console.log(`RUN_CASE ${caseIndex} ${JSON.stringify(timing)}`);
  const result = await runCase(timing);
  console.log(JSON.stringify({ caseIndex, result }, null, 2));
  if (result.samples.some((sample) => sample.alphaCount >= 2)) console.log("DUPLICATE_REPRODUCED");
  else { console.log("NO_DUPLICATE_REPRODUCED"); process.exitCode = 2; }
} finally {
  restoreSource(originalSource);
  rmSync(keyHelper, { force: true });
  rmSync(tisHelper, { force: true });
}

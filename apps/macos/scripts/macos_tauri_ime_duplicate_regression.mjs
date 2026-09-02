import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { rmSync } from "node:fs";

if (
  !process.argv.includes("--cleanup-last-scratch") &&
  !process.argv.includes("--allow-legacy-probe-automation")
) {
  throw new Error(
    "Legacy scratch/probe automation is disabled: it cannot prove the main-bundle physical IME mechanism. Use macos_main_ime_overlap_regression.mjs for fail-closed evidence.",
  );
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function run(command, args) {
  const result = spawnSync(command, args, { encoding: "utf8" });
  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || `${command} failed`);
  }
  return result.stdout;
}

function osascript(lines) {
  const args = [];
  for (const line of lines) args.push("-e", line);
  return run("osascript", args);
}

function jxa(script) {
  return run("osascript", ["-l", "JavaScript", "-e", script]);
}

const tisHelper = `/tmp/visualtex-tis-select-${process.pid}`;
const hidHelper = `/tmp/visualtex-hid-key-${process.pid}`;
let tisHelperReady = false;
let hidHelperReady = false;
function ensureTisHelper() {
  if (tisHelperReady) return;
  const source = String.raw`
import Carbon
import Foundation
let target = CommandLine.arguments[1]
func currentSourceID() -> String {
    let source = TISCopyCurrentKeyboardInputSource().takeRetainedValue()
    guard let rawID = TISGetInputSourceProperty(source, kTISPropertyInputSourceID) else {
        fputs("current input source has no ID\\n", stderr)
        exit(4)
    }
    return Unmanaged<CFString>.fromOpaque(rawID).takeUnretainedValue() as String
}
if target == "--current" {
    print(currentSourceID())
    exit(0)
}
let filter = [kTISPropertyInputSourceID as String: target] as CFDictionary
let list = TISCreateInputSourceList(filter, false).takeRetainedValue() as! [TISInputSource]
if list.isEmpty { fputs("missing input source\\n", stderr); exit(2) }
let status = TISSelectInputSource(list[0])
if status != noErr { fputs("TISSelectInputSource failed: \\(status)\\n", stderr); exit(3) }
for _ in 0..<25 {
    if currentSourceID() == target { exit(0) }
    usleep(40_000)
}
fputs("input source selection was not applied: expected \\(target), actual \\(currentSourceID())\\n", stderr)
exit(5)
`;
  const result = spawnSync("swiftc", ["-O", "-o", tisHelper, "-"], {
    encoding: "utf8",
    input: source,
  });
  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || "swiftc TIS helper failed");
  }
  tisHelperReady = true;
}
function selectInputSource(id) {
  ensureTisHelper();
  run(tisHelper, [id]);
}

function currentInputSourceID() {
  ensureTisHelper();
  return run(tisHelper, ["--current"]).trim();
}

function ensureHidHelper() {
  if (hidHelperReady) return;
  const source = String.raw`
import CoreGraphics
import Foundation
let mode = CommandLine.arguments[1]
let key: CGKeyCode
let flags: CGEventFlags
switch mode {
case "ctrl-space": key = 49; flags = [.maskControl]
case "caps-lock": key = 57; flags = []
default: fputs("unsupported HID mode\\n", stderr); exit(2)
}
let down = CGEvent(keyboardEventSource: nil, virtualKey: key, keyDown: true)!
down.flags = flags
down.post(tap: .cghidEventTap)
usleep(80000)
let up = CGEvent(keyboardEventSource: nil, virtualKey: key, keyDown: false)!
up.flags = flags
up.post(tap: .cghidEventTap)
usleep(180000)
`;
  const result = spawnSync("swiftc", ["-O", "-o", hidHelper, "-"], {
    encoding: "utf8",
    input: source,
  });
  if (result.status !== 0) throw new Error(result.stderr || result.stdout || "swiftc HID helper failed");
  hidHelperReady = true;
}

function physicalSystemSwitch(mode) {
  ensureHidHelper();
  run(hidHelper, [mode]);
}

function selectedInputSourceDescription() {
  return run("defaults", ["read", "com.apple.HIToolbox", "AppleSelectedInputSources"]);
}

function currentSourceKind() {
  const id = currentInputSourceID();
  if (id === "com.apple.inputmethod.SCIM.ITABC") return "pinyin";
  if (id === "com.apple.keylayout.ABC") return "abc";
  return "other";
}

function restoreInputSource(kind) {
  if (kind === "pinyin") selectInputSource("com.apple.inputmethod.SCIM.ITABC");
  if (kind === "abc") selectInputSource("com.apple.keylayout.ABC");
}

function appSnapshot() {
  const output = jxa(`(() => {
    const se = Application("System Events");
    const processes = se.applicationProcesses.whose({ name: { _equals: "visualtex" } })();
    const walk = (element, depth, result) => {
      if (depth > 15) return;
      let role = "";
      try { role = String(element.role() ?? ""); } catch {}
      if (role === "AXTextArea") {
        try { result.sources.push(String(element.value() ?? "")); } catch {}
      }
      if (role === "AXTextField") {
        let name = null, desc = null, pos = null, size = null;
        try { name = element.name() == null ? null : String(element.name()); } catch {}
        try { desc = element.description() == null ? null : String(element.description()); } catch {}
        try { pos = element.position(); } catch {}
        try { size = element.size(); } catch {}
        result.fields.push({ name, desc, pos, size });
      }
      if (role === "AXButton") {
        let name = null;
        try { name = element.name() == null ? null : String(element.name()); } catch {}
        if (name && name.startsWith("\\\\")) result.commandButtons.push(name);
      }
      let children = [];
      try { children = element.uiElements(); } catch {}
      for (const child of children) walk(child, depth + 1, result);
    };
    for (const process of processes) {
      let windows = [];
      try { windows = process.windows(); } catch {}
      for (const window of windows) {
        const result = { sources: [], fields: [], commandButtons: [], pid: null };
        try { result.pid = process.unixId(); } catch {}
        walk(window, 0, result);
        const hasSourceEditor = result.sources.some((value) => value.includes("$$"));
        const hasFormulaField = result.fields.some((field) =>
          field.desc !== "公式文档标题" && field.desc !== "Formula document title"
        );
        if (hasSourceEditor && hasFormulaField) return JSON.stringify(result);
      }
    }
    throw new Error("VisualTeX main window is not visible");
  })()`);
  return JSON.parse(output);
}

function sourceFrom(snapshot) {
  const source = snapshot.sources.find((value) => value.includes("$$"));
  if (typeof source !== "string") {
    throw new Error(`LaTeX source panel is not visible: ${JSON.stringify(snapshot)}`);
  }
  return source;
}

function formulaFields(snapshot) {
  return snapshot.fields.filter((field) => field.desc !== "公式文档标题" && field.desc !== "Formula document title");
}

function clickField(field, pid) {
  assert.ok(Array.isArray(field.pos) && Array.isArray(field.size), `Field has no AX geometry: ${JSON.stringify(field)}`);
  const x = Math.round(field.pos[0] + Math.max(3, Math.min(field.size[0] / 2, 20)));
  const y = Math.round(field.pos[1] + Math.max(3, field.size[1] / 2));
  osascript([
    'tell application "System Events"',
    `set frontmost of first application process whose unix id is ${pid} to true`,
    `click at {${x}, ${y}}`,
    'end tell',
  ]);
}

function keys(pid, lines) {
  osascript([
    'tell application "System Events"',
    `set frontmost of first application process whose unix id is ${pid} to true`,
    'delay 0.05',
    ...lines,
    'end tell',
  ]);
}

async function waitForFieldCount(count, timeout = 2500) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    const snapshot = appSnapshot();
    if (formulaFields(snapshot).length === count) return snapshot;
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${count} formula fields; got ${JSON.stringify(appSnapshot())}`);
}

async function restoreBaseline(baselineSource, baselineFieldCount, pid) {
  let snapshot = appSnapshot();
  let source = sourceFrom(snapshot);
  let fields = formulaFields(snapshot);
  if (source === baselineSource && fields.length === baselineFieldCount) return;

  if (fields.length === baselineFieldCount + 1) {
    let scratch = fields.at(-1);
    assert.ok(scratch, "Scratch field missing during cleanup");
    clickField(scratch, pid);
    await sleep(70);
    keys(pid, ['key code 53', 'keystroke "a" using {command down}', 'key code 51']);
    await sleep(180);

    snapshot = appSnapshot();
    source = sourceFrom(snapshot);
    fields = formulaFields(snapshot);
    if (source === baselineSource && fields.length === baselineFieldCount) return;

    if (fields.length === baselineFieldCount + 1) {
      scratch = fields.at(-1);
      assert.ok(scratch, "Empty scratch field missing during cleanup");
      clickField(scratch, pid);
      await sleep(70);
      keys(pid, ['key code 51']);
      await sleep(220);
    }
  }

  snapshot = appSnapshot();
  source = sourceFrom(snapshot);
  fields = formulaFields(snapshot);
  assert.equal(fields.length, baselineFieldCount, `Scratch cleanup changed the wrong number of formula rows: ${JSON.stringify(fields)}`);
  assert.equal(source, baselineSource, `Scratch cleanup did not restore the exact baseline: expected=${JSON.stringify(baselineSource)} actual=${JSON.stringify(source)}`);
}

async function createScratchLine(baselineFieldCount, pid) {
  let snapshot = appSnapshot();
  const fields = formulaFields(snapshot);
  assert.equal(fields.length, baselineFieldCount);
  const last = fields.at(-1);
  assert.ok(last, "No formula field found");
  clickField(last, pid);
  await sleep(120);
  keys(pid, ['key code 124 using {command down}', 'key code 36']);
  snapshot = await waitForFieldCount(baselineFieldCount + 1);
  const scratch = formulaFields(snapshot).at(-1);
  assert.ok(scratch, "Scratch formula field missing");
  clickField(scratch, pid);
  await sleep(160);
  return scratch;
}

async function typePinyinAThenDelete(pid, afterDeleteMs = 280, afterAms = 180) {
  selectInputSource("com.apple.inputmethod.SCIM.ITABC");
  await sleep(120);
  if (afterAms <= 0) {
    keys(pid, ['keystroke "a"', 'key code 51']);
  } else {
    keys(pid, ['keystroke "a"']);
    await sleep(afterAms);
    keys(pid, ['key code 51']);
  }
  if (afterDeleteMs > 0) await sleep(afterDeleteMs);
}

async function chooseCommandAndSpace(pid, prefix, expectedCommand) {
  keys(pid, ['key code 42', `keystroke ${JSON.stringify(prefix)}`]);
  await sleep(350);
  let snapshot = appSnapshot();
  const expectedPrefix = `${expectedCommand} `;
  for (let index = 0; index < 12; index += 1) {
    const first = snapshot.commandButtons[0] ?? "";
    if (first.startsWith(expectedPrefix) || first === expectedCommand) break;
    keys(pid, ['key code 125']);
    await sleep(110);
    snapshot = appSnapshot();
  }
  const selectedCandidates = snapshot.commandButtons;
  if (!selectedCandidates.some((value) => value.startsWith(expectedCommand))) {
    throw new Error(`Expected native candidate ${expectedCommand} is not visible: ${JSON.stringify(selectedCandidates)}`);
  }
  keys(pid, ['key code 49']);
}

async function snapshotsAfterSpace(expectedCommand) {
  const samples = [];
  let elapsed = 0;
  const delays = process.argv.includes("--hid-alpha")
    ? [250]
    : process.argv.includes("--hid-only")
      ? [250, 750]
      : [120, 350, 650, 1000, 1200];
  for (const delay of delays) {
    await sleep(delay);
    elapsed += delay;
    const snapshot = appSnapshot();
    const source = sourceFrom(snapshot);
    const count = source.split(expectedCommand).length - 1;
    samples.push({ elapsedMs: elapsed, count, sourceTail: source.slice(-180), fields: formulaFields(snapshot).map((field) => field.name) });
  }
  return samples;
}

async function runCase({ baselineSource, baselineFieldCount, pid, command, prefix, switchMode, afterDeleteMs = 280, afterSwitchMs = 300, afterAms = 180 }) {
  await createScratchLine(baselineFieldCount, pid);
  await typePinyinAThenDelete(pid, afterDeleteMs, afterAms);
  if (switchMode === "abc-source") {
    selectInputSource("com.apple.keylayout.ABC");
    if (afterSwitchMs > 0) await sleep(afterSwitchMs);
  } else if (switchMode === "hid-ctrl-space") {
    physicalSystemSwitch("ctrl-space");
    if (afterSwitchMs > 0) await sleep(afterSwitchMs);
    assert.equal(currentSourceKind(), "abc", `HID Control+Space did not switch to ABC: ${selectedInputSourceDescription()}`);
  } else if (switchMode === "ctrl-space-then-abc") {
    // Preserve the exact physical Control+Space transaction seen by WKWebView,
    // then force the system source to ABC because synthetic HID events are not
    // always honored by macOS's global input-source shortcut service.
    physicalSystemSwitch("ctrl-space");
    if (afterSwitchMs > 0) await sleep(afterSwitchMs);
    selectInputSource("com.apple.keylayout.ABC");
  } else if (switchMode === "caps-lock") {
    // Keep Pinyin selected, but switch its keyboard state to Latin using Caps Lock.
    physicalSystemSwitch("caps-lock");
    await sleep(300);
  } else {
    throw new Error(`Unknown switch mode ${switchMode}`);
  }
  const inputSourceIdBeforeBackslash = currentInputSourceID();
  assert.equal(
    inputSourceIdBeforeBackslash,
    "com.apple.keylayout.ABC",
    `Refusing to type Backslash before the exact ABC source is active: ${inputSourceIdBeforeBackslash}; ${selectedInputSourceDescription()}`,
  );
  await chooseCommandAndSpace(pid, prefix, command);
  const samples = await snapshotsAfterSpace(command);
  if (switchMode === "caps-lock") {
    physicalSystemSwitch("caps-lock");
    await sleep(180);
  }
  await restoreBaseline(baselineSource, baselineFieldCount, pid);
  return samples;
}

if (process.argv.includes("--cleanup-last-scratch")) {
  const snapshot = appSnapshot();
  const pid = snapshot.pid;
  const source = sourceFrom(snapshot);
  const fields = formulaFields(snapshot);
  assert.ok(/\$\$\n(?:\\(?:alpha|int)|\\text\{、\})\n\$\$\n$/.test(source), `Refusing cleanup: last source block is not one of this script's scratch lines: ${JSON.stringify(source.slice(-120))}`);
  assert.ok(fields.length >= 2, `Refusing cleanup: not enough formula fields: ${fields.length}`);
  const targetCount = fields.length - 1;
  for (let attempt = 0; attempt < 12; attempt += 1) {
    const current = appSnapshot();
    if (formulaFields(current).length === targetCount) {
      console.log("Removed the interrupted Tauri scratch line without touching the prior baseline");
      process.exit(0);
    }
    keys(pid, ['keystroke "z" using {command down}']);
    await sleep(180);
  }
  throw new Error(`Could not remove only the interrupted scratch line; current=${JSON.stringify(appSnapshot())}`);
}

const originalInputSource = currentSourceKind();
let baselineSource = "";
let baselineFieldCount = 0;
let pid = null;

try {
  const baseline = appSnapshot();
  pid = baseline.pid;
  assert.ok(Number.isInteger(pid), `VisualTeX PID missing: ${JSON.stringify(baseline)}`);
  baselineSource = sourceFrom(baseline);
  baselineFieldCount = formulaFields(baseline).length;
  assert.ok(baselineFieldCount >= 1, "No editable formula field found");

  const allCases = [
    { command: "\\alpha", prefix: "al", switchMode: "abc-source", afterDeleteMs: 280, afterSwitchMs: 300 },
    { command: "\\int", prefix: "in", switchMode: "abc-source", afterDeleteMs: 280, afterSwitchMs: 300 },
    { command: "\\alpha", prefix: "al", switchMode: "caps-lock" },
    { command: "\\int", prefix: "in", switchMode: "caps-lock" },
  ];
  const fastCases = [
    { command: "\\alpha", prefix: "al", switchMode: "abc-source", afterDeleteMs: 0, afterSwitchMs: 0 },
    { command: "\\int", prefix: "in", switchMode: "abc-source", afterDeleteMs: 0, afterSwitchMs: 0 },
    { command: "\\alpha", prefix: "al", switchMode: "abc-source", afterDeleteMs: 20, afterSwitchMs: 20 },
    { command: "\\int", prefix: "in", switchMode: "abc-source", afterDeleteMs: 20, afterSwitchMs: 20 },
    { command: "\\alpha", prefix: "al", switchMode: "abc-source", afterDeleteMs: 100, afterSwitchMs: 50 },
  ];
  const hidCases = [
    { command: "\\alpha", prefix: "al", switchMode: "hid-ctrl-space", afterDeleteMs: 0, afterSwitchMs: 0, afterAms: 0 },
    { command: "\\int", prefix: "in", switchMode: "hid-ctrl-space", afterDeleteMs: 0, afterSwitchMs: 0, afterAms: 0 },
  ];
  const shortcutThenAbcCases = [
    { command: "\\alpha", prefix: "al", switchMode: "ctrl-space-then-abc", afterDeleteMs: 0, afterSwitchMs: 0, afterAms: 0 },
    { command: "\\int", prefix: "in", switchMode: "ctrl-space-then-abc", afterDeleteMs: 0, afterSwitchMs: 0, afterAms: 0 },
  ];
  if (process.argv.includes("--hid-only") || process.argv.includes("--hid-alpha")) ensureHidHelper();
  const cases = process.argv.includes("--shortcut-alpha")
    ? shortcutThenAbcCases.slice(0, 1)
    : process.argv.includes("--shortcut-int")
      ? shortcutThenAbcCases.slice(1, 2)
      : process.argv.includes("--fast-alpha")
    ? fastCases.slice(0, 1)
    : process.argv.includes("--fast-int")
      ? fastCases.slice(1, 2)
      : process.argv.includes("--caps-alpha")
        ? allCases.filter((testCase) => testCase.switchMode === "caps-lock" && testCase.command === "\\alpha")
        : process.argv.includes("--caps-int")
          ? allCases.filter((testCase) => testCase.switchMode === "caps-lock" && testCase.command === "\\int")
          : process.argv.includes("--hid-alpha")
            ? hidCases.slice(0, 1)
            : process.argv.includes("--hid-only")
              ? hidCases
              : process.argv.includes("--fast-only")
                ? fastCases
                : process.argv.includes("--caps-only")
                  ? allCases.filter((testCase) => testCase.switchMode === "caps-lock")
                  : allCases;

  for (const testCase of cases) {
    const samples = await runCase({ baselineSource, baselineFieldCount, pid, ...testCase });
    console.log(JSON.stringify({ case: testCase, samples }, null, 2));
    const baselineCount = baselineSource.split(testCase.command).length - 1;
    for (const sample of samples) {
      assert.equal(
        sample.count,
        baselineCount + 1,
        `${testCase.switchMode}/${testCase.command} duplicated or lost command at ${sample.elapsedMs}ms: ${JSON.stringify(sample)}`,
      );
    }
  }
  console.log("Tauri/WKWebView physical IME duplicate regression passed");
} finally {
  if (pid && baselineSource) {
    try { await restoreBaseline(baselineSource, baselineFieldCount, pid); } catch (error) { console.error(error); }
  }
  restoreInputSource(originalInputSource);
}

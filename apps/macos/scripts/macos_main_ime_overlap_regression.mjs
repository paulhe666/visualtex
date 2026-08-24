import assert from "node:assert/strict";
import { execFileSync, spawnSync } from "node:child_process";
import {
  existsSync,
  mkdirSync,
  mkdtempSync,
  rmSync,
  writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import process from "node:process";

// Real-main-bundle functional regression with exact TIS source gates. This
// intentionally does not claim to emulate a physical Control+Space/Caps Lock
// handoff. The top-level verdict stays false unless WKWebView itself emits the
// original same-timestamp Unidentified/Backslash -> \\/Backslash mechanism.

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function option(name, fallback = "") {
  const prefix = `--${name}=`;
  return process.argv.find((value) => value.startsWith(prefix))?.slice(prefix.length) ?? fallback;
}

const pid = Number(option("pid"));
const maxTrials = Number(option("trials", "1"));
const requiredPassingTrials = Number(option("required-hits", "1"));
const commandKind = option("command", "alpha");
const outputDir = path.resolve(option("output-dir", `test-results/ime/main-automated-overlap-${Date.now()}-pid${pid}`));
const captureOnly = process.argv.includes("--capture-only");

if (!Number.isInteger(pid) || pid <= 0) {
  throw new Error("Usage: node scripts/macos_main_ime_overlap_regression.mjs --pid=<VisualTeX PID> [--command=alpha|int] [--trials=1] [--required-hits=1] [--output-dir=<new directory>]");
}
if (!Number.isInteger(maxTrials) || maxTrials < 1) throw new Error(`Invalid --trials=${maxTrials}`);
if (!captureOnly && maxTrials !== 1) {
  throw new Error("Refusing repeated trials in one main-app lifecycle: each exact repro must start from an independently verified pristine empty field");
}
if (!Number.isInteger(requiredPassingTrials) || requiredPassingTrials < 1) {
  throw new Error(`Invalid --required-hits=${requiredPassingTrials}`);
}
if (!captureOnly && requiredPassingTrials !== 1) {
  throw new Error("A single pristine main-app trial can require exactly one mechanism hit");
}
if (commandKind !== "alpha" && commandKind !== "int") {
  throw new Error(`Invalid --command=${commandKind}; expected alpha or int`);
}
if (existsSync(outputDir)) throw new Error(`Refusing to overwrite existing evidence directory: ${outputDir}`);
mkdirSync(outputDir, { recursive: true });

function run(command, args, options = {}) {
  return execFileSync(command, args, { encoding: "utf8", ...options });
}

const executablePath = run("ps", ["-p", String(pid), "-o", "command="]).trim();
const appBundleMarker = ".app/Contents/MacOS/";
const appBundleIndex = executablePath.indexOf(appBundleMarker);
if (appBundleIndex < 0) throw new Error(`PID ${pid} is not running from a macOS app bundle: ${executablePath}`);
const appPath = `${executablePath.slice(0, appBundleIndex)}.app`;
const bundleId = run("/usr/libexec/PlistBuddy", [
  "-c",
  "Print :CFBundleIdentifier",
  path.join(appPath, "Contents", "Info.plist"),
]).trim();
assert.equal(
  bundleId,
  "com.visualtex.studio",
  `Refusing to test a probe or non-main bundle: PID ${pid}, bundle ${bundleId}, path ${appPath}`,
);
const axProbePath = path.join(process.cwd(), "scripts", "macos_ax_ime_probe.swift");
if (!existsSync(axProbePath)) throw new Error(`Missing AX probe: ${axProbePath}`);

function reopenMainWindow() {
  run("open", [appPath]);
}

function jxa(source) {
  return run("osascript", ["-l", "JavaScript", "-e", source]);
}

function appSnapshot() {
  const raw = JSON.parse(run(axHelper, ["snapshot", String(pid)], { maxBuffer: 4 * 1024 * 1024 }));
  const sourceArea = raw.areas.find((area) => typeof area.value === "string" && area.value.includes("$$"));
  const diagnosticArea = raw.areas.find((area) =>
    area.title === "VisualTeX IME Diagnostic Trace" ||
    area.description === "VisualTeX IME Diagnostic Trace"
  );
  const formulaFields = raw.fields.filter((field) =>
    field.description !== "公式文档标题" &&
    field.description !== "Formula document title" &&
    field.title !== "公式文档标题" &&
    field.title !== "Formula document title"
  );
  return {
    pid,
    source: sourceArea?.value ?? "",
    diagnostic: diagnosticArea?.value ?? "",
    formulaFields,
  };
}

function focusOnlyFormulaField() {
  const result = JSON.parse(run(axHelper, ["focus-last", String(pid)]));
  assert.equal(result.fieldCount, 1, `Expected exactly one formula field, got ${result.fieldCount}`);
  assert.equal(result.result, 0, `AX focus failed with result ${result.result}`);
  return result;
}

function clearFormula() {
  focusOnlyFormulaField();
  run(helper, [String(pid), "", "86000", "before"], {
    env: { ...process.env, VISUALTEX_RESET_ONLY: "1" },
  });
}

function parseTrace(value) {
  return value
    .split("\n")
    .filter(Boolean)
    .map((line) => {
      try { return JSON.parse(line); } catch { return null; }
    })
    .filter(Boolean);
}

function lastTransaction(entries) {
  const startIndex = entries.findLastIndex((entry) => entry.stage === "field.compositionstart");
  return startIndex >= 0 ? entries.slice(startIndex) : entries;
}

function traceSince(snapshot, traceFloorSeq) {
  const entries = parseTrace(snapshot.diagnostic);
  return traceFloorSeq > 0
    ? entries.filter((entry) => Number(entry.seq) > traceFloorSeq)
    : lastTransaction(entries);
}

function analyzePinyinPrelude(snapshot, traceFloorSeq, traceCeilingSeq = 0) {
  const sinceFloor = traceSince(snapshot, traceFloorSeq);
  const trace = traceCeilingSeq > 0
    ? sinceFloor.filter((entry) => Number(entry.seq) <= traceCeilingSeq)
    : sinceFloor;
  const starts = trace.filter((entry) => entry.stage === "field.compositionstart");
  const insertions = trace.filter((entry) =>
    entry.stage === "field.beforeinput" &&
    entry.inputType === "insertCompositionText" &&
    entry.data === "a"
  );
  const deletes = trace.filter((entry) =>
    entry.stage === "field.beforeinput" && entry.inputType === "deleteCompositionText"
  );
  const ends = trace.filter((entry) => entry.stage === "field.compositionend.end");
  const plainAInsertions = trace.filter((entry) =>
    entry.stage === "field.beforeinput" &&
    entry.inputType === "insertText" &&
    entry.data === "a"
  );
  const ordered = Boolean(
    starts.length === 1 &&
    insertions.length === 1 &&
    deletes.length >= 1 &&
    ends.length === 1 &&
    starts[0].seq < insertions[0].seq &&
    insertions[0].seq < deletes[0].seq &&
    deletes.at(-1).seq < ends[0].seq
  );
  const verdict = {
    traceFloorSeq,
    traceEntryCount: trace.length,
    traceStartSeq: trace[0]?.seq ?? null,
    traceEndSeq: trace.at(-1)?.seq ?? null,
    singleFormulaLine: snapshot.formulaFields.length === 1,
    exactEmptySource: normalizedSource(snapshot.source) === "$$\n\n$$",
    compositionStartCount: starts.length,
    compositionInsertACount: insertions.length,
    compositionDeleteCount: deletes.length,
    compositionEndCount: ends.length,
    plainAInsertionCount: plainAInsertions.length,
    ordered,
    compositionPath: false,
    pass: false,
  };
  verdict.compositionPath = Boolean(verdict.plainAInsertionCount === 0 && ordered);
  verdict.pass = Boolean(
    verdict.singleFormulaLine &&
    verdict.exactEmptySource &&
    verdict.compositionPath
  );
  return { trace, verdict };
}

function commandCount(source, command) {
  return source.split(command).length - 1;
}

function normalizedSource(source) {
  return source.endsWith("\n") ? source.slice(0, -1) : source;
}

function analyze(
  snapshot,
  expectedCommand,
  expectedRaw,
  traceFloorSeq = 0,
  preludeTraceEndSeq = 0,
) {
  const trace = traceSince(snapshot, traceFloorSeq);
  const prelude = analyzePinyinPrelude(
    snapshot,
    traceFloorSeq,
    preludeTraceEndSeq,
  ).verdict;
  const commandTrace = preludeTraceEndSeq > 0
    ? trace.filter((entry) => Number(entry.seq) > preludeTraceEndSeq)
    : trace;
  const keydowns = commandTrace.filter((entry) => entry.stage === "window.capture.keydown");
  const backslashes = keydowns.filter((entry) => entry.code === "Backslash");
  const spaces = keydowns.filter((entry) => entry.code === "Space" && !entry.ctrlKey && !entry.metaKey && !entry.altKey);
  const selected = commandTrace.filter((entry) => entry.stage === "commitNativeSuggestion.selected");
  const insertBefore = commandTrace.filter((entry) => entry.stage === "commitNativeSuggestion.insert.before");
  const insertAfter = commandTrace.filter((entry) => entry.stage === "commitNativeSuggestion.insert.after");
  const first = backslashes[0] ?? null;
  const second = backslashes[1] ?? null;
  const firstKeydownIndex = first ? keydowns.indexOf(first) : -1;
  const exactMechanism = Boolean(
    backslashes.length === 2 &&
    first &&
    second &&
    first.key === "Unidentified" &&
    first.code === "Backslash" &&
    second.key === "\\" &&
    second.code === "Backslash" &&
    first.timeStamp === second.timeStamp &&
    firstKeydownIndex >= 0 &&
    keydowns[firstKeydownIndex + 1] === second,
  );
  const sourceBody = normalizedSource(snapshot.source).match(/^\$\$\n([\s\S]*)\n\$\$$/)?.[1] ?? null;
  const verdict = {
    singleFormulaLine: snapshot.formulaFields.length === 1,
    sourceBody,
    expectedSource: sourceBody === expectedCommand,
    commandCount: commandCount(snapshot.source, expectedCommand),
    exactMechanism,
    backslashes: backslashes.map(({ seq, eventId, key, code, keyCode, timeStamp, value, raw, mode }) => ({
      seq, eventId, key, code, keyCode, timeStamp, value, raw, mode,
    })),
    spaceCount: spaces.length,
    commitCount: selected.length,
    insertBeforeCount: insertBefore.length,
    insertAfterCount: insertAfter.length,
    selected: selected.at(-1) ?? null,
    insertBefore: insertBefore.at(-1) ?? null,
    insertAfter: insertAfter.at(-1) ?? null,
    traceEntryCount: trace.length,
    commandTraceEntryCount: commandTrace.length,
    traceFloorSeq,
    preludeTraceEndSeq,
    traceStartSeq: trace[0]?.seq ?? null,
    traceEndSeq: trace.at(-1)?.seq ?? null,
    pinyinCompositionPath: prelude.compositionPath,
    transactionReached: false,
    functionalPass: false,
    pass: false,
  };
  verdict.transactionReached = Boolean(
    verdict.pinyinCompositionPath &&
    verdict.spaceCount === 1 &&
    verdict.commitCount === 1 &&
    verdict.insertBeforeCount === 1 &&
    verdict.insertAfterCount === 1 &&
    verdict.selected?.rawInput === expectedRaw &&
    verdict.selected?.selectedCommand === expectedCommand
  );
  verdict.functionalPass = Boolean(
    verdict.singleFormulaLine &&
    verdict.expectedSource &&
    verdict.commandCount === 1 &&
    verdict.spaceCount === 1 &&
    verdict.commitCount === 1 &&
    verdict.insertBeforeCount === 1 &&
    verdict.insertAfterCount === 1 &&
    verdict.transactionReached &&
    verdict.insertBefore?.hasAnchor === true &&
    verdict.insertBefore?.value === "" &&
    verdict.insertBefore?.raw === "" &&
    verdict.insertAfter?.value === expectedCommand
  );
  verdict.pass = Boolean(verdict.functionalPass && verdict.exactMechanism);
  return { trace, verdict };
}

function currentInputSourceId() {
  const result = JSON.parse(run(axHelper, ["current-source"]));
  return result.id;
}

function currentInputSource() {
  const id = currentInputSourceId();
  if (id === "com.apple.inputmethod.SCIM.ITABC") return "pinyin";
  if (id === "com.apple.keylayout.ABC") return "abc";
  return "other";
}

async function waitForInputSource(expected, timeout = 1800) {
  const started = Date.now();
  while (Date.now() - started < timeout) {
    const actual = currentInputSource();
    if (actual === expected) return actual;
    await sleep(40);
  }
  throw new Error(`Input-source gate failed: expected ${expected}, actual ${currentInputSource()}`);
}

const workDir = mkdtempSync(path.join(tmpdir(), "visualtex-main-ime-overlap-"));
const helper = path.join(workDir, "physical-ime-sequence");
const helperSource = String.raw`
import AppKit
import ApplicationServices
import Carbon
import CoreGraphics
import Foundation

let pid = pid_t(Int(CommandLine.arguments[1])!)
let typed = CommandLine.arguments[2]
let capsToBackslashUs = useconds_t(Int(CommandLine.arguments[3])!)
let releaseStage = CommandLine.arguments[4]

func currentSourceID() -> String {
    let source = TISCopyCurrentKeyboardInputSource().takeRetainedValue()
    guard let rawID = TISGetInputSourceProperty(source, kTISPropertyInputSourceID) else {
        fputs("current input source has no ID\n", stderr)
        exit(24)
    }
    return Unmanaged<CFString>.fromOpaque(rawID).takeUnretainedValue() as String
}

func selectSource(_ id: String) {
    let filter = [kTISPropertyInputSourceID as String: id] as CFDictionary
    let list = TISCreateInputSourceList(filter, false).takeRetainedValue() as! [TISInputSource]
    guard let source = list.first else { fputs("missing input source\n", stderr); exit(20) }
    if TISSelectInputSource(source) != noErr { fputs("input source selection failed\n", stderr); exit(21) }
    for _ in 0..<25 {
        if currentSourceID() == id { return }
        usleep(40_000)
    }
    fputs("input source selection was not applied: expected \(id), actual \(currentSourceID())\n", stderr)
    exit(25)
}

if let selectOnly = ProcessInfo.processInfo.environment["VISUALTEX_SELECT_ONLY"] {
    selectSource(selectOnly)
    exit(0)
}

func event(_ code: CGKeyCode, down: Bool, flags: CGEventFlags = []) -> CGEvent {
    let value = CGEvent(keyboardEventSource: nil, virtualKey: code, keyDown: down)!
    value.flags = flags
    return value
}

func post(_ code: CGKeyCode, down: Bool, flags: CGEventFlags = []) {
    event(code, down: down, flags: flags).post(tap: .cghidEventTap)
}

func postToPid(_ code: CGKeyCode, down: Bool, flags: CGEventFlags = []) {
    event(code, down: down, flags: flags).postToPid(pid)
}

func tap(_ code: CGKeyCode, holdUs: useconds_t = 85_000, gapUs: useconds_t = 85_000, flags: CGEventFlags = []) {
    post(code, down: true, flags: flags)
    usleep(holdUs)
    post(code, down: false, flags: flags)
    usleep(gapUs)
}

func tapToPid(_ code: CGKeyCode, holdUs: useconds_t = 85_000, gapUs: useconds_t = 85_000, flags: CGEventFlags = []) {
    postToPid(code, down: true, flags: flags)
    usleep(holdUs)
    postToPid(code, down: false, flags: flags)
    usleep(gapUs)
}

func code(_ char: Character) -> CGKeyCode {
    switch char {
    case "a": return 0
    case "h": return 4
    case "i": return 34
    case "l": return 37
    case "n": return 45
    case "p": return 35
    case "t": return 17
    default: fputs("unsupported character\n", stderr); exit(22)
    }
}

func activateTarget() {
    guard let app = NSRunningApplication(processIdentifier: pid) else { exit(23) }
    _ = app.unhide()
    _ = app.activate(options: [])
    for _ in 0..<25 {
        if NSWorkspace.shared.frontmostApplication?.processIdentifier == pid { return }
        usleep(40_000)
    }
    fputs("VisualTeX PID \(pid) did not become frontmost\n", stderr)
    exit(27)
}

func requireABCBeforeBackslash() {
    let id = currentSourceID()
    guard id == "com.apple.keylayout.ABC" else {
        fputs("refusing Backslash outside exact ABC source: \(id)\n", stderr)
        exit(26)
    }
}

if ProcessInfo.processInfo.environment["VISUALTEX_RESET_ONLY"] == "1" {
    activateTarget()
    usleep(180_000)
    tapToPid(0, holdUs: 65_000, gapUs: 80_000, flags: [.maskCommand])
    tapToPid(51, holdUs: 70_000, gapUs: 180_000)
    exit(0)
}

if ProcessInfo.processInfo.environment["VISUALTEX_PINYIN_CONTEXT_ONLY"] == "1" {
    guard let app = NSRunningApplication(processIdentifier: pid) else { exit(23) }
    _ = app.hide()
    usleep(160_000)
    // Force a real source transition so a stale Pinyin-in-Latin submode from a
    // previous Caps Lock experiment cannot masquerade as Chinese composition.
    selectSource("com.apple.keylayout.ABC")
    usleep(120_000)
    selectSource("com.apple.inputmethod.SCIM.ITABC")
    usleep(180_000)
    exit(0)
}

if ProcessInfo.processInfo.environment["VISUALTEX_A_DELETE_ONLY"] == "1" {
    activateTarget()
    let sourceBeforeA = currentSourceID()
    guard sourceBeforeA == "com.apple.inputmethod.SCIM.ITABC" else {
        fputs("refusing Pinyin prelude outside exact Pinyin source: \(sourceBeforeA)\n", stderr)
        exit(28)
    }
    tap(0, holdUs: 83_000, gapUs: 108_000)
    tap(51, holdUs: 104_000, gapUs: 214_000)
    exit(0)
}

if ProcessInfo.processInfo.environment["VISUALTEX_COMMAND_ONLY"] == "1" {
    activateTarget()
    usleep(150_000)
    requireABCBeforeBackslash()
    tap(42, holdUs: 92_000, gapUs: 175_000)
    for char in typed {
        tap(code(char), holdUs: 82_000, gapUs: 72_000)
    }
    tap(49, holdUs: 70_000, gapUs: 500_000)
    exit(0)
}

activateTarget()
usleep(140_000)
selectSource("com.apple.inputmethod.SCIM.ITABC")
usleep(180_000)

// The measured path starts here: Pinyin -> a -> Backspace.
tap(0, holdUs: 83_000, gapUs: 108_000)
tap(51, holdUs: 104_000, gapUs: 214_000)

// The original failing HID trace had Caps Lock down, then Backslash down before
// a distinct Caps Lock up appeared. Varying this overlap deterministically
// exercises that real input-source handoff rather than selecting ABC via API.
post(57, down: true)
if releaseStage == "before" {
    usleep(80_000)
    post(57, down: false)
    usleep(capsToBackslashUs)
    requireABCBeforeBackslash()
    post(42, down: true)
} else {
    usleep(capsToBackslashUs)
    requireABCBeforeBackslash()
    post(42, down: true)
    if releaseStage == "during" {
        usleep(18_000)
        post(57, down: false)
    }
}
usleep(92_000)
post(42, down: false)
if releaseStage == "after" {
    usleep(18_000)
    post(57, down: false)
}
usleep(175_000)

for char in typed {
    tap(code(char), holdUs: 82_000, gapUs: 72_000)
}
tap(49, holdUs: 70_000, gapUs: 500_000)

if releaseStage == "late" { post(57, down: false) }
`;

const compile = spawnSync("swiftc", ["-O", "-o", helper, "-"], {
  encoding: "utf8",
  input: helperSource,
});
if (compile.status !== 0) throw new Error(compile.stderr || compile.stdout || "Failed to compile physical IME helper");
const axHelper = path.join(workDir, "ax-ime-probe");
const axCompile = spawnSync("swiftc", ["-O", "-o", axHelper, axProbePath], { encoding: "utf8" });
if (axCompile.status !== 0) throw new Error(axCompile.stderr || axCompile.stdout || "Failed to compile AX probe");

function writeEvidence(name, snapshot, extra = {}) {
  const { trace, verdict } = analyze(
    snapshot,
    extra.expectedCommand ?? "\\alpha",
    extra.expectedRaw ?? "\\alph",
    extra.traceFloorSeq ?? 0,
    extra.preludeTraceEndSeq ?? 0,
  );
  writeFileSync(path.join(outputDir, `${name}-trace.jsonl`), `${trace.map((entry) => JSON.stringify(entry)).join("\n")}\n`);
  writeFileSync(path.join(outputDir, `${name}.json`), `${JSON.stringify({ ...extra, snapshot: { ...snapshot, diagnostic: undefined }, verdict }, null, 2)}\n`);
  return verdict;
}

function writePreludeEvidence(name, snapshot, traceFloorSeq, extra = {}) {
  const { trace, verdict } = analyzePinyinPrelude(snapshot, traceFloorSeq);
  writeFileSync(
    path.join(outputDir, `${name}-trace.jsonl`),
    `${trace.map((entry) => JSON.stringify(entry)).join("\n")}\n`,
  );
  writeFileSync(
    path.join(outputDir, `${name}.json`),
    `${JSON.stringify({ ...extra, snapshot: { ...snapshot, diagnostic: undefined }, verdict }, null, 2)}\n`,
  );
  return verdict;
}

function captureScreenshot(name) {
  try {
    const geometry = JSON.parse(jxa(`(() => {
      const se = Application("System Events");
      const app = se.applicationProcesses.whose({ unixId: { _equals: ${pid} } })()[0];
      const window = app.windows.whose({ name: { _equals: "VisualTeX" } })()[0];
      return JSON.stringify(window ? { position: window.position(), size: window.size() } : {});
    })()`));
    const position = geometry.position;
    const size = geometry.size;
    if (!Array.isArray(position) || !Array.isArray(size)) throw new Error("VisualTeX window geometry is unavailable");
    const region = `${Math.round(position[0])},${Math.round(position[1])},${Math.round(size[0])},${Math.round(size[1])}`;
    run("screencapture", ["-x", `-R${region}`, path.join(outputDir, `${name}.png`)]);
  } catch (error) {
    writeFileSync(path.join(outputDir, `${name}-screenshot-error.txt`), `${error.stack ?? error}\n`);
  }
}

async function requirePristineEmptyBaseline() {
  const snapshot = appSnapshot();
  assert.equal(snapshot.formulaFields.length, 1, `Expected exactly one formula line: ${JSON.stringify(snapshot.formulaFields)}`);
  assert.equal(
    normalizedSource(snapshot.source),
    "$$\n\n$$",
    `Refusing to modify a non-empty main-app document: ${JSON.stringify(snapshot.source)}`,
  );
  focusOnlyFormulaField();
  await sleep(160);
  return snapshot;
}

async function restoreCapturedTestOutput(capturedSource, expectedCommand, expectedRaw) {
  if (!expectedCommand) return;
  const current = appSnapshot();
  assert.equal(current.formulaFields.length, 1, "Cleanup observed an unexpected formula line count");
  if (normalizedSource(current.source) === "$$\n\n$$") return;
  assert.ok(
    capturedSource,
    "Refusing cleanup of non-empty content without a successfully captured trial source",
  );
  assert.equal(
    current.source,
    capturedSource,
    "Refusing cleanup because the main-app document changed after evidence capture",
  );
  const body = normalizedSource(current.source).match(/^\$\$\n([\s\S]*)\n\$\$$/)?.[1] ?? null;
  const allowedBodies = new Set([
    expectedCommand,
    expectedRaw,
    `${expectedRaw}${expectedCommand}`,
    `${expectedCommand}${expectedCommand}`,
  ]);
  assert.ok(allowedBodies.has(body), `Refusing cleanup of unrecognized content: ${JSON.stringify(body)}`);
  clearFormula();
  await sleep(450);
  const restored = appSnapshot();
  assert.equal(restored.formulaFields.length, 1, "Cleanup changed the formula line count");
  assert.equal(normalizedSource(restored.source), "$$\n\n$$", `Cleanup did not restore the pristine empty source: ${JSON.stringify(restored.source)}`);
}

function writeFinalizationError(name, error) {
  try {
    writeFileSync(path.join(outputDir, name), `${error.stack ?? error}\n`);
  } catch (writeError) {
    console.error(`Could not write ${name}: ${writeError.stack ?? writeError}`);
  }
}

const originalInputSourceId = currentInputSourceId();
const results = [];
let capturedTrialSource = "";
let activeExpectedCommand = "";
let activeExpectedRaw = "";
let pendingSummary = null;
let cleanupVerified = false;
let inputSourceRestored = false;
try {
  reopenMainWindow();
  await sleep(350);
  const initial = appSnapshot();
  writeFileSync(path.join(outputDir, "initial-snapshot.json"), `${JSON.stringify({ ...initial, diagnostic: undefined }, null, 2)}\n`);
  if (captureOnly) {
    const verdict = writeEvidence("capture-current", initial, {
      capturedAt: new Date().toISOString(),
      kind: "manual-current-state",
      expectedCommand: "\\alpha",
      expectedRaw: "\\alph",
    });
    captureScreenshot("capture-current");
    const summary = {
      captureOnly: true,
      regressionEligible: false,
      pass: false,
      pid,
      observedVerdict: verdict,
    };
    writeFileSync(path.join(outputDir, "summary.json"), `${JSON.stringify(summary, null, 2)}\n`);
    console.log(JSON.stringify({ outputDir, summary }, null, 2));
  } else {
    for (let trial = 1; trial <= maxTrials; trial += 1) {
      const expectedCommand = commandKind === "int" ? "\\int" : "\\alpha";
      const typed = expectedCommand === "\\int" ? "int" : "alph";
      const expectedRaw = `\\${typed}`;
      const schedule = { switchMode: "TIS Pinyin to verified ABC" };
      const baseline = await requirePristineEmptyBaseline();
      const traceFloorSeq = parseTrace(baseline.diagnostic).at(-1)?.seq ?? 0;
      activeExpectedCommand = expectedCommand;
      activeExpectedRaw = expectedRaw;
      run(helper, [String(pid), typed, "0", "before"], {
        env: { ...process.env, VISUALTEX_PINYIN_CONTEXT_ONLY: "1" },
      });
      await waitForInputSource("pinyin");
      reopenMainWindow();
      await sleep(350);
      focusOnlyFormulaField();
      await sleep(180);
      run(helper, [String(pid), typed, "0", "before"], {
        env: { ...process.env, VISUALTEX_A_DELETE_ONLY: "1" },
      });
      const afterDelete = appSnapshot();
      const inputSourceIdAfterDelete = currentInputSourceId();
      const preludeVerdict = writePreludeEvidence(
        `trial-${String(trial).padStart(3, "0")}-pinyin-prelude`,
        afterDelete,
        traceFloorSeq,
        {
          trial,
          capturedAt: new Date().toISOString(),
          inputSourceIdAfterDelete,
        },
      );
      assert.equal(
        inputSourceIdAfterDelete,
        "com.apple.inputmethod.SCIM.ITABC",
        `Pinyin source changed before the composition prelude was verified: ${inputSourceIdAfterDelete}`,
      );
      assert.equal(
        preludeVerdict.pass,
        true,
        `Refusing to switch or type a command without a fresh Pinyin a/composition-delete path: ${JSON.stringify(preludeVerdict)}`,
      );
      run(helper, [String(pid), "", "0", "before"], {
        env: { ...process.env, VISUALTEX_SELECT_ONLY: "com.apple.keylayout.ABC" },
      });
      await waitForInputSource("abc");
      const afterSwitch = appSnapshot();
      assert.equal(
        normalizedSource(afterSwitch.source),
        "$$\n\n$$",
        `Pinyin a/Backspace or input-source switch changed the empty formula: ${JSON.stringify(afterSwitch.source)}`,
      );
      const inputSourceIdBeforeBackslash = currentInputSourceId();
      assert.equal(
        inputSourceIdBeforeBackslash,
        "com.apple.keylayout.ABC",
        `Refusing to type Backslash before the exact ABC source is active: ${inputSourceIdBeforeBackslash}`,
      );
      run(helper, [String(pid), typed, "0", "before"], {
        env: { ...process.env, VISUALTEX_COMMAND_ONLY: "1" },
      });
      await sleep(280);
      const after = appSnapshot();
      capturedTrialSource = after.source;
      const verdict = writeEvidence(`trial-${String(trial).padStart(3, "0")}`, after, {
        trial,
        capturedAt: new Date().toISOString(),
        baselineSource: baseline.source,
        inputSourceIdAfterDelete,
        preludeVerdict,
        inputSourceIdBeforeBackslash,
        sourceAfterVerifiedSwitch: currentInputSource(),
        inputSourceAfter: currentInputSource(),
        schedule,
        expectedCommand,
        expectedRaw,
        traceFloorSeq,
        preludeTraceEndSeq: preludeVerdict.traceEndSeq,
      });
      const row = { trial, schedule, expectedCommand, inputSourceIdBeforeBackslash, ...verdict };
      results.push(row);
      console.log(JSON.stringify({
        trial,
        expectedCommand,
        exactMechanism: verdict.exactMechanism,
        functionalPass: verdict.functionalPass,
        exactPass: verdict.pass,
        sourceBody: verdict.sourceBody,
        hasAnchor: verdict.insertBefore?.hasAnchor ?? null,
        backslashes: verdict.backslashes.map(({ key, code, timeStamp }) => ({ key, code, timeStamp })),
      }));
      if (!verdict.functionalPass) {
        captureScreenshot(`trial-${String(trial).padStart(3, "0")}-failure`);
        throw new Error(`Automated main-app regression failed at trial ${trial}: ${JSON.stringify(row)}`);
      }
      const passing = results.filter((result) => result.pass);
      if (passing.length >= requiredPassingTrials) break;
    }

    const validResults = results.filter((result) => result.transactionReached);
    const mechanismHits = validResults.filter((result) => result.exactMechanism);
    const summary = {
      capturedAt: new Date().toISOString(),
      pid,
      commandKind,
      maxTrials,
      requiredPassingTrials,
      trialsRun: results.length,
      validTrials: validResults.length,
      invalidTimingTrials: results.length - validResults.length,
      functionalPassed: validResults.filter((result) => result.functionalPass).length,
      passed: validResults.filter((result) => result.pass).length,
      mechanismHitCount: mechanismHits.length,
      alphaMechanismHitCount: mechanismHits.filter((result) => result.expectedCommand === "\\alpha").length,
      intMechanismHitCount: mechanismHits.filter((result) => result.expectedCommand === "\\int").length,
      results: results.map(({ selected, insertBefore, insertAfter, ...result }) => result),
    };
    summary.pass = mechanismHits.length >= requiredPassingTrials &&
      mechanismHits.every((result) => result.pass) &&
      mechanismHits.every((result) => result.expectedCommand === activeExpectedCommand);
    pendingSummary = summary;
    captureScreenshot("final");
  }
} finally {
  let finalizationError = null;
  if (!captureOnly) {
    try {
      await restoreCapturedTestOutput(capturedTrialSource, activeExpectedCommand, activeExpectedRaw);
      cleanupVerified = true;
    } catch (error) {
      writeFinalizationError("cleanup-error.txt", error);
      finalizationError = error;
    }
  }
  if (!captureOnly) {
    try {
      if (currentInputSourceId() !== originalInputSourceId) {
        run(helper, [String(pid), "", "86000", "before"], {
          env: { ...process.env, VISUALTEX_SELECT_ONLY: originalInputSourceId },
        });
      }
      assert.equal(
        currentInputSourceId(),
        originalInputSourceId,
        "Input source did not remain restored after cleanup",
      );
      inputSourceRestored = true;
    } catch (error) {
      writeFinalizationError("input-source-restore-error.txt", error);
      finalizationError ??= error;
    }
  }
  try { rmSync(workDir, { recursive: true }); } catch {}
  if (finalizationError) throw finalizationError;
}

if (!captureOnly && pendingSummary) {
  pendingSummary.finalization = {
    cleanupVerified,
    inputSourceRestored,
    sourceRestoredTo: originalInputSourceId,
  };
  pendingSummary.pass = Boolean(
    pendingSummary.pass && cleanupVerified && inputSourceRestored,
  );
  writeFileSync(
    path.join(outputDir, "summary.json"),
    `${JSON.stringify(pendingSummary, null, 2)}\n`,
  );
  assert.equal(
    pendingSummary.pass,
    true,
    `Did not hit the original WKWebView mechanism often enough: ${JSON.stringify(pendingSummary)}`,
  );
  console.log(JSON.stringify({ outputDir, summary: pendingSummary }, null, 2));
}

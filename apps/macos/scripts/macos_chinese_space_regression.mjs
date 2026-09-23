import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import process from "node:process";

const sleep = (milliseconds) =>
  new Promise((resolve) => setTimeout(resolve, milliseconds));

const pidArgument = process.argv.find((argument) => argument.startsWith("--pid="));
const pid = Number(pidArgument?.slice("--pid=".length));
if (!Number.isInteger(pid) || pid <= 0) {
  throw new Error("Usage: node scripts/macos_chinese_space_regression.mjs --pid=<VisualTeX PID>");
}

function run(command, args, options = {}) {
  return execFileSync(command, args, {
    encoding: "utf8",
    maxBuffer: 8 * 1024 * 1024,
    ...options,
  });
}

const executablePath = run("ps", ["-p", String(pid), "-o", "command="]).trim();
const appMarker = ".app/Contents/MacOS/";
const appMarkerIndex = executablePath.indexOf(appMarker);
if (appMarkerIndex < 0) {
  throw new Error(`PID ${pid} is not a bundled macOS application: ${executablePath}`);
}
const appPath = `${executablePath.slice(0, appMarkerIndex)}.app`;
const bundleIdentifier = run("/usr/libexec/PlistBuddy", [
  "-c",
  "Print :CFBundleIdentifier",
  path.join(appPath, "Contents", "Info.plist"),
]).trim();
assert.equal(bundleIdentifier, "com.visualtex.studio");

const axProbeSource = path.join(process.cwd(), "scripts", "macos_ax_ime_probe.swift");
if (!existsSync(axProbeSource)) throw new Error(`Missing AX probe: ${axProbeSource}`);
const helperDirectory = run("mktemp", ["-d", `${tmpdir()}/visualtex-chinese-space.XXXXXX`]).trim();
const axProbe = path.join(helperDirectory, "ax-probe");
run("swiftc", ["-O", "-o", axProbe, axProbeSource]);

function snapshot() {
  const raw = JSON.parse(run(axProbe, ["snapshot", String(pid)]));
  const sourceArea = raw.areas.find(
    (area) => typeof area.value === "string" && area.value.includes("$$"),
  );
  const formulaFields = raw.fields.filter((field) => {
    const label = `${field.title ?? ""} ${field.description ?? ""}`;
    return !label.includes("公式文档标题") && !label.includes("Formula document title");
  });
  return {
    source: sourceArea?.value ?? "",
    formulaFieldCount: formulaFields.length,
  };
}

function currentInputSource() {
  return JSON.parse(run(axProbe, ["current-source"])).id;
}

function selectInputSource(identifier) {
  run(axProbe, ["source", identifier]);
}

function activateRunningApplication() {
  const script = `tell application "System Events" to set frontmost of first application process whose unix id is ${pid} to true`;
  run("osascript", ["-e", script]);
}

function focusLastFormula() {
  const result = JSON.parse(run(axProbe, ["focus-last", String(pid)]));
  assert.equal(result.result, 0, `Unable to focus the last formula field: ${JSON.stringify(result)}`);
  return result.fieldCount;
}

function key(keyCode, ...flags) {
  run(axProbe, ["key", String(pid), String(keyCode), ...flags]);
}

async function waitForSnapshot(predicate, description, timeout = 3000) {
  const startedAt = Date.now();
  let latest = snapshot();
  while (Date.now() - startedAt < timeout) {
    latest = snapshot();
    if (predicate(latest)) return latest;
    await sleep(80);
  }
  throw new Error(`Timed out waiting for ${description}: ${JSON.stringify(latest)}`);
}

function displayFormulaBodies(source) {
  return [...source.matchAll(/\$\$\n([\s\S]*?)\n\$\$/g)].map((match) =>
    match[1].trim(),
  );
}

const pinyinKeyCodes = {
  z: 6,
  h: 4,
  o: 31,
  n: 45,
  g: 5,
  w: 13,
  e: 14,
};

const originalInputSource = currentInputSource();
const before = snapshot();
let scratch = null;
let observed = null;
let cleanupVerified = false;

try {
  // The target PID is already running. Re-opening its bundle exercises the
  // single-instance Office fast-open path and can display an unrelated native
  // error dialog, stealing IME focus from the temporary formula row.
  activateRunningApplication();
  await sleep(300);
  assert.ok(before.source, "VisualTeX source panel is unavailable");
  const focusedFieldCount = focusLastFormula();
  assert.equal(focusedFieldCount, before.formulaFieldCount);

  key(124, "meta");
  key(36);
  scratch = await waitForSnapshot(
    (state) => state.formulaFieldCount === before.formulaFieldCount + 1,
    "one temporary formula row",
  );
  // Empty formula rows are intentionally omitted from the formatted source;
  // the extra AX mathfield is the authoritative creation check.
  assert.equal(scratch.source, before.source);

  selectInputSource("com.apple.inputmethod.SCIM.ITABC");
  assert.equal(currentInputSource(), "com.apple.inputmethod.SCIM.ITABC");
  focusLastFormula();
  for (const character of "zhongwen") {
    key(pinyinKeyCodes[character]);
  }
  key(49);
  await sleep(500);

  observed = snapshot();
  assert.equal(observed.formulaFieldCount, scratch.formulaFieldCount);
  const beforeBodies = displayFormulaBodies(scratch.source);
  const afterBodies = displayFormulaBodies(observed.source);
  const insertedBody = afterBodies.at(-1) ?? "";
  const previousLastBody = beforeBodies.at(-1) ?? null;
  const verdict = {
    insertedBody,
    previousLastBody,
    exactChineseText: insertedBody === "\\text{中文}",
    hasTrailingExplicitSpace: /\\text\{中文\}\\\s/.test(insertedBody),
    hasTrailingTextSpace: /\\text\{中文\s+\}/.test(insertedBody),
    hasTrailingPlainSpace: /中文\s+$/u.test(insertedBody),
  };
  console.log(JSON.stringify({ phase: "observed", verdict }, null, 2));

  focusLastFormula();
  for (let attempt = 0; attempt < 8; attempt += 1) {
    if (snapshot().source === scratch.source) break;
    key(51);
    await sleep(100);
  }
  const clearedScratch = snapshot();
  assert.equal(
    clearedScratch.source,
    scratch.source,
    "Refusing row removal because the temporary Chinese input did not clear exactly",
  );
  key(51);
  const restored = await waitForSnapshot(
    (state) =>
      state.formulaFieldCount === before.formulaFieldCount &&
      state.source === before.source,
    "the exact pre-test document",
  );
  assert.equal(restored.source, before.source);
  cleanupVerified = true;

  if (!verdict.exactChineseText) {
    throw new Error(`Real macOS Pinyin input appended unwanted content: ${JSON.stringify(verdict)}`);
  }
  console.log(
    JSON.stringify(
      {
        pass: true,
        pid,
        appPath,
        inputMethod: "com.apple.inputmethod.SCIM.ITABC",
        typedKeys: "zhongwen + Space",
        verdict,
        cleanupVerified,
      },
      null,
      2,
    ),
  );
} finally {
  if (currentInputSource() !== originalInputSource) {
    selectInputSource(originalInputSource);
  }
  if (!cleanupVerified && scratch && observed) {
    console.error("Native IME probe stopped before it could verify exact document restoration.");
  }
}

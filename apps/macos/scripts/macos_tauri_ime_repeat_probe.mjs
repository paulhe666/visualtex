import { execFileSync, spawnSync } from "node:child_process";
import process from "node:process";

const appPath = `${process.cwd()}/src-tauri/target/release/bundle/macos/VisualTeX IME Probe 2.app`;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const trials = Number(process.argv[2] || 10);
const typed = process.argv[3] || "alph";

function run(command, args, options = {}) {
  return execFileSync(command, args, { encoding: "utf8", ...options });
}

run("open", [appPath]);
await sleep(700);

function probePid() {
  const ps = run("ps", ["-axo", "pid=,command="]);
  const line = ps.split("\n").find((value) => value.includes("/VisualTeX IME Probe 2.app/Contents/MacOS/visualtex"));
  if (!line) throw new Error("IME Probe 2 process is not running");
  return Number(line.trim().split(/\s+/, 1)[0]);
}

const pid = probePid();

const helperPath = `/tmp/visualtex-ime-repeat-helper-${process.pid}`;
const helperSource = String.raw`
import AppKit
import ApplicationServices
import Carbon
import CoreGraphics
import Foundation

let pid = pid_t(Int(CommandLine.arguments[1])!)
let typed = CommandLine.arguments[2]

func selectSource(_ id: String) {
    let filter = [kTISPropertyInputSourceID as String: id] as CFDictionary
    let list = TISCreateInputSourceList(filter, false).takeRetainedValue() as! [TISInputSource]
    guard let source = list.first else { fputs("missing input source\\n", stderr); exit(20) }
    if TISSelectInputSource(source) != noErr { fputs("select input source failed\\n", stderr); exit(21) }
}

func keyCode(for char: Character) -> CGKeyCode? {
    switch char {
    case "a": return 0
    case "b": return 11
    case "c": return 8
    case "d": return 2
    case "e": return 14
    case "f": return 3
    case "g": return 5
    case "h": return 4
    case "i": return 34
    case "j": return 38
    case "k": return 40
    case "l": return 37
    case "m": return 46
    case "n": return 45
    case "o": return 31
    case "p": return 35
    case "q": return 12
    case "r": return 15
    case "s": return 1
    case "t": return 17
    case "u": return 32
    case "v": return 9
    case "w": return 13
    case "x": return 7
    case "y": return 16
    case "z": return 6
    default: return nil
    }
}

func postToProbe(_ code: CGKeyCode, _ flags: CGEventFlags = []) {
    let down = CGEvent(keyboardEventSource: nil, virtualKey: code, keyDown: true)!
    down.flags = flags
    down.postToPid(pid)
    usleep(45_000)
    let up = CGEvent(keyboardEventSource: nil, virtualKey: code, keyDown: false)!
    up.flags = flags
    up.postToPid(pid)
    usleep(90_000)
}

func physicalControlSpace() {
    let controlDown = CGEvent(keyboardEventSource: nil, virtualKey: 59, keyDown: true)!
    controlDown.flags = [.maskControl]
    controlDown.post(tap: .cghidEventTap)
    usleep(55_000)
    let spaceDown = CGEvent(keyboardEventSource: nil, virtualKey: 49, keyDown: true)!
    spaceDown.flags = [.maskControl]
    spaceDown.post(tap: .cghidEventTap)
    usleep(70_000)
    let spaceUp = CGEvent(keyboardEventSource: nil, virtualKey: 49, keyDown: false)!
    spaceUp.flags = [.maskControl]
    spaceUp.post(tap: .cghidEventTap)
    usleep(55_000)
    let controlUp = CGEvent(keyboardEventSource: nil, virtualKey: 59, keyDown: false)!
    controlUp.post(tap: .cghidEventTap)
    usleep(260_000)
}

func mainWindow() -> AXUIElement {
    let app = AXUIElementCreateApplication(pid)
    var value: CFTypeRef?
    guard AXUIElementCopyAttributeValue(app, kAXWindowsAttribute as CFString, &value) == .success,
          let windows = value as? [AXUIElement] else { exit(30) }
    for window in windows {
        var titleValue: CFTypeRef?
        _ = AXUIElementCopyAttributeValue(window, kAXTitleAttribute as CFString, &titleValue)
        if (titleValue as? String) == "VisualTeX" { return window }
    }
    exit(31)
}

func formulaFields(_ root: AXUIElement) -> [AXUIElement] {
    var result: [AXUIElement] = []
    func walk(_ element: AXUIElement, _ depth: Int) {
        if depth > 16 { return }
        var roleValue: CFTypeRef?
        _ = AXUIElementCopyAttributeValue(element, kAXRoleAttribute as CFString, &roleValue)
        if (roleValue as? String) == (kAXTextFieldRole as String) {
            var descriptionValue: CFTypeRef?
            _ = AXUIElementCopyAttributeValue(element, kAXDescriptionAttribute as CFString, &descriptionValue)
            let description = (descriptionValue as? String) ?? ""
            if description != "公式文档标题" && description != "Formula document title" {
                result.append(element)
            }
        }
        var childrenValue: CFTypeRef?
        if AXUIElementCopyAttributeValue(element, kAXChildrenAttribute as CFString, &childrenValue) == .success,
           let children = childrenValue as? [AXUIElement] {
            for child in children { walk(child, depth + 1) }
        }
    }
    walk(root, 0)
    return result
}

_ = NSRunningApplication(processIdentifier: pid)?.activate(options: [])
usleep(180_000)
let root = mainWindow()
let fields = formulaFields(root)
guard let field = fields.last else { exit(32) }
_ = AXUIElementSetAttributeValue(field, kAXFocusedAttribute as CFString, kCFBooleanTrue)
usleep(100_000)

// Clear only the probe scratch field.
postToProbe(0, [.maskCommand])
postToProbe(51)
usleep(180_000)

selectSource("com.apple.inputmethod.SCIM.ITABC")
usleep(180_000)
postToProbe(0)       // Pinyin composition: a
usleep(130_000)
postToProbe(51)      // cancel with Backspace
usleep(180_000)
physicalControlSpace()

// Keep the same field focused across the real input-source shortcut.
for char in typed {
    if let code = keyCode(for: char) { postToProbe(code) }
}
postToProbe(49)      // Space confirms MathLive native candidate
usleep(500_000)
`;

const compile = spawnSync("swiftc", ["-O", "-o", helperPath, "-"], {
  input: helperSource,
  encoding: "utf8",
});
if (compile.status !== 0) throw new Error(compile.stderr || compile.stdout);

function readTrace() {
  const script = `(() => {
    const se = Application("System Events");
    const processes = se.applicationProcesses.whose({ unixId: { _equals: ${pid} } })();
    if (!processes.length) throw new Error("probe process unavailable");
    const walk = (element, depth = 0) => {
      if (depth > 16) return null;
      try {
        if (String(element.role()) === "AXTextArea" && String(element.name() || "") === "VisualTeX IME Diagnostic Trace") {
          return String(element.value() || "");
        }
      } catch {}
      try {
        for (const child of element.uiElements()) {
          const value = walk(child, depth + 1);
          if (value !== null) return value;
        }
      } catch {}
      return null;
    };
    for (const window of processes[0].windows()) {
      const value = walk(window);
      if (value !== null) return value;
    }
    throw new Error("diagnostic trace unavailable");
  })()`;
  return run("osascript", ["-l", "JavaScript", "-e", script]);
}

function currentSourceKind() {
  const text = run("defaults", ["read", "com.apple.HIToolbox", "AppleSelectedInputSources"]);
  if (text.includes("KeyboardLayout Name") && text.includes("ABC")) return "abc";
  if (text.includes("com.apple.inputmethod.SCIM.ITABC")) return "pinyin";
  return "other";
}

const results = [];
for (let trial = 1; trial <= trials; trial += 1) {
  run(helperPath, [String(pid), typed]);
  const trace = readTrace();
  const entries = trace
    .split("\n")
    .filter(Boolean)
    .map((line) => {
      try { return JSON.parse(line); } catch { return null; }
    })
    .filter(Boolean);
  const insertBefore = [...entries].reverse().find((entry) => entry.stage === "commitNativeSuggestion.insert.before");
  const insertAfter = [...entries].reverse().find((entry) => entry.stage === "commitNativeSuggestion.insert.after");
  const commitSelected = [...entries].reverse().find((entry) => entry.stage === "commitNativeSuggestion.selected");
  const space = [...entries].reverse().find((entry) => entry.stage === "window.capture.keydown" && entry.code === "Space" && !entry.ctrlKey);
  const compositionEnd = [...entries].reverse().find((entry) => entry.stage === "field.compositionend.end");
  const row = {
    trial,
    sourceAfter: currentSourceKind(),
    rawInput: commitSelected?.rawInput ?? null,
    selectedCommand: commitSelected?.selectedCommand ?? null,
    hasAnchor: insertBefore?.hasAnchor ?? null,
    valueBeforeInsert: insertBefore?.value ?? null,
    valueAfterInsert: insertAfter?.value ?? null,
    compositionCancelled: compositionEnd?.data === "",
    duplicated: insertAfter?.value === "\\alpha\\alpha",
    spaceSeq: space?.seq ?? null,
  };
  results.push(row);
  console.log(JSON.stringify(row));
  await sleep(250);
}

const duplicated = results.filter((row) => row.duplicated).length;
console.log(JSON.stringify({ trials, typed, duplicated, rate: `${duplicated}/${trials}` }, null, 2));

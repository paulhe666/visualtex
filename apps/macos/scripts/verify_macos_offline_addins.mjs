import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") process.exit(0);

const appRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const rootArgument = process.argv.indexOf("--offline-root");
if (rootArgument >= 0 && !process.argv[rootArgument + 1]) {
  throw new Error("--offline-root requires the packaged Office resource directory");
}
const offlineRoot = rootArgument >= 0
  ? resolve(process.argv[rootArgument + 1])
  : join(appRoot, "office", "macos-offline");
const resourcesRoot = join(offlineRoot, "resources");
const packageVersion = JSON.parse(readFileSync(join(appRoot, "package.json"), "utf8")).version;
const manifest = JSON.parse(readFileSync(join(resourcesRoot, "addins.json"), "utf8"));

// A VBE-compiled template can contain every macro while still lacking the
// Ribbon package. Validate the final OOXML, including its relationships, before
// accepting hashes or promoting it into a DMG.
execFileSync("python3", [join(appRoot, "scripts", "verify_macos_office_ribbon.py"), offlineRoot], {
  stdio: "inherit",
});

function sha256(path) {
  return createHash("sha256").update(readFileSync(path)).digest("hex");
}

function containsMarker(buffer, marker) {
  return (
    buffer.includes(Buffer.from(marker, "utf8")) ||
    buffer.includes(Buffer.from(marker, "utf16le"))
  );
}

function inspectAddin(name, vbaEntry, requiredMarkers) {
  const path = join(resourcesRoot, name);
  execFileSync("/usr/bin/unzip", ["-t", path], { stdio: "ignore" });
  const project = execFileSync("/usr/bin/unzip", ["-p", path, vbaEntry], {
    encoding: null,
    maxBuffer: 32 * 1024 * 1024,
  });
  const missing = requiredMarkers.filter((marker) => !containsMarker(project, marker));
  if (missing.length > 0) {
    throw new Error(
      `${name} is stale and is missing reviewed VBA marker(s): ${missing.join(", ")}. Recompile the reviewed VBA sources in Office for Mac before building the DMG.`,
    );
  }
  const expectedHash = manifest.files?.[name]?.sha256;
  if (typeof expectedHash !== "string" || expectedHash.toLowerCase() !== sha256(path)) {
    throw new Error(`${name} does not match office/macos-offline/resources/addins.json`);
  }
}

if (manifest.pluginVersion !== packageVersion) {
  throw new Error(
    `The Office add-in manifest version ${manifest.pluginVersion ?? "missing"} does not match VisualTeX ${packageVersion}`,
  );
}

inspectAddin("VisualTeX.dotm", "word/vbaProject.bin", [
  packageVersion,
  "VTWordAdapter",
  "VTWordEvents",
  "AutoExec",
  "App_WindowBeforeDoubleClick",
  "App_WindowSelectionChange",
  "VisualTeX_StabilizeImageEquationNumberSelection",
  "word-office-performance-20260801-r87",
  "VTTraceWordDoubleClick",
  "VTWordRibbonApplyImageFontSizePreset",
  "VisualTeX_EditImageField",
  "VisualTeX_EditSelectedImageFromNativeMonitor",
  "VTEnsureVisualTeXImageMacroButton",
  "VTAppendText",
  "VTWriteAndLaunchSession",
  "VTPrewarmApplication",
]);
inspectAddin("VisualTeX.ppam", "ppt/vbaProject.bin", [
  packageVersion,
  "VTPowerPointAdapter",
  "VTPowerPointEvents",
  "Auto_Open",
  "App_WindowBeforeDoubleClick",
  "App_WindowSelectionChange",
  "VTPowerPointRibbonApplyFormulaFontSizePreset",
  "powerpoint-office-performance-20260801-r4",
]);

process.stdout.write("VisualTeX compiled macOS Office add-ins: PASS\n");

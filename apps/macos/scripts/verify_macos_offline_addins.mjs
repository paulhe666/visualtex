import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "darwin") process.exit(0);

const appRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const resourcesRoot = join(appRoot, "office", "macos-offline", "resources");
const packageVersion = JSON.parse(readFileSync(join(appRoot, "package.json"), "utf8")).version;
const manifest = JSON.parse(readFileSync(join(resourcesRoot, "addins.json"), "utf8"));

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
  "word-image-number-deterministic-assertion-20260723-r39",
]);
inspectAddin("VisualTeX.ppam", "ppt/vbaProject.bin", [
  packageVersion,
  "VTPowerPointAdapter",
  "VTPowerPointEvents",
  "Auto_Open",
  "App_WindowBeforeDoubleClick",
]);

process.stdout.write("VisualTeX compiled macOS Office add-ins: PASS\n");

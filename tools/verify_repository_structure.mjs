import { existsSync, lstatSync, readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const failures = [];

function requirePath(relativePath) {
  if (!existsSync(resolve(root, relativePath))) {
    failures.push(`Missing required path: ${relativePath}`);
  }
}

function forbidPath(relativePath) {
  if (existsSync(resolve(root, relativePath))) {
    failures.push(`Retired path must not exist: ${relativePath}`);
  }
}

for (const path of [
  "apps/macos/package.json",
  "apps/macos/src",
  "apps/macos/src-tauri",
  "apps/macos/office/macos-offline",
  "apps/macos/office/macos-offline/resources/VisualTeX.dotm",
  "apps/macos/office/macos-offline/resources/VisualTeX.ppam",
  "apps/windows/package.json",
  "apps/windows/src",
  "apps/windows/src-tauri",
  "apps/windows/src-windows",
  "apps/windows/office/windows/ole/manifests/visualtex-word.template.xml",
  "apps/windows/office/windows/ole/manifests/visualtex-powerpoint.template.xml",
  "README.md",
  "docs/ARCHITECTURE.md",
  "docs/images/visualtex-light-mode.png",
  "docs/images/visualtex-dark-mode.png",
  "docs/images/visualtex-macos-word-ribbon.png",
  "docs/images/visualtex-windows-word-ribbon.png",
]) {
  requirePath(path);
}

for (const path of [
  "src",
  "src-tauri",
  "src-windows",
  "office",
  "ocr",
  "scripts",
  "apps/macos/src-windows",
  "apps/macos/office/vendor",
  "apps/macos/office/manifests",
  "apps/macos/office/macos",
  "apps/macos/office/windows",
  "apps/macos/office-bridge.html",
  "apps/macos/office-dialog.html",
  "apps/macos/office-macos-bridge.html",
  "apps/macos/office-windows-ole-bridge.html",
  "apps/windows/office/vendor",
  "apps/windows/office/manifests",
  "apps/windows/office/macos",
  "apps/windows/office-bridge.html",
  "apps/windows/office-macos-bridge.html",
  "apps/windows/src/office/macos",
]) {
  forbidPath(path);
}

const macPackage = JSON.parse(
  readFileSync(resolve(root, "apps/macos/package.json"), "utf8"),
);
const windowsPackage = JSON.parse(
  readFileSync(resolve(root, "apps/windows/package.json"), "utf8"),
);

if (macPackage.name !== "visualtex-macos") {
  failures.push(`Unexpected macOS package name: ${macPackage.name}`);
}
if (windowsPackage.name !== "visualtex-windows") {
  failures.push(`Unexpected Windows package name: ${windowsPackage.name}`);
}
if (macPackage.version !== windowsPackage.version) {
  failures.push(
    `Platform versions differ: macOS ${macPackage.version}, Windows ${windowsPackage.version}`,
  );
}
if (macPackage.dependencies?.["@microsoft/office-js"] || macPackage.devDependencies?.["@types/office-js"]) {
  failures.push("The macOS application must not depend on Office.js");
}
if (!windowsPackage.devDependencies?.["@microsoft/office-js"] || !windowsPackage.devDependencies?.["@types/office-js"]) {
  failures.push("The Windows compatibility build must retain its explicit Office.js dependencies");
}

for (const path of ["apps/macos", "apps/windows"]) {
  if (lstatSync(resolve(root, path)).isSymbolicLink()) {
    failures.push(`${path} must be a real directory, not a symbolic link`);
  }
}

const trackedFiles = execFileSync("git", ["ls-files", "apps/macos", "apps/windows"], {
  cwd: root,
  encoding: "utf8",
})
  .split("\n")
  .filter(Boolean);

for (const relativePath of trackedFiles) {
  if (/^apps\/macos\/office\/vendor\/office-js\//.test(relativePath)) {
    failures.push(`Committed Office.js vendor file remains in macOS: ${relativePath}`);
  }
  if (/^apps\/windows\/office\/vendor\/office-js\//.test(relativePath)) {
    failures.push(`Generic committed Office.js vendor file remains in Windows: ${relativePath}`);
  }
  if (!/\.(?:ts|tsx|js|mjs|cjs|rs|cs|cpp|c|h|hpp|ps1|sh|json|toml|xml|md)$/.test(relativePath)) {
    continue;
  }
  const content = readFileSync(resolve(root, relativePath), "utf8");
  if (relativePath.startsWith("apps/macos/") && content.includes("apps/windows/")) {
    failures.push(`macOS source references the Windows application: ${relativePath}`);
  }
  if (relativePath.startsWith("apps/windows/") && content.includes("apps/macos/")) {
    failures.push(`Windows source references the macOS application: ${relativePath}`);
  }
}

if (failures.length) {
  console.error("VisualTeX repository structure verification failed:\n");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log("VisualTeX repository structure and legacy Office cleanup are complete.");

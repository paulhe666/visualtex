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
    failures.push(`Legacy root path must not exist: ${relativePath}`);
  }
}

for (const path of [
  "apps/macos/package.json",
  "apps/macos/src",
  "apps/macos/src-tauri",
  "apps/macos/office/macos-offline",
  "apps/windows/package.json",
  "apps/windows/src",
  "apps/windows/src-tauri",
  "apps/windows/src-windows",
  "README.md",
  "docs/ARCHITECTURE.md",
  "docs/images/visualtex-macos-editor.png",
  "docs/images/visualtex-windows-editor.png",
]) {
  requirePath(path);
}

for (const path of ["src", "src-tauri", "src-windows", "office", "ocr", "scripts"]) {
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

console.log("VisualTeX repository structure is platform-separated and complete.");

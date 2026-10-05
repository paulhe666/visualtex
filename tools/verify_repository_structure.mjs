import { existsSync, lstatSync, readFileSync } from "node:fs";
import { spawn } from "node:child_process";
import { dirname, relative, resolve } from "node:path";

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
  "docs/images/application.png",
  "docs/images/wechat-pay.jpg",
  "docs/images/alipay.jpg",
  "docs/images/qq-group.png",
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
for (const [platform, packageInfo] of [["macos", macPackage], ["windows", windowsPackage]]) {
  const config = JSON.parse(readFileSync(resolve(root, `apps/${platform}/src-tauri/tauri.conf.json`), "utf8"));
  const cargo = readFileSync(resolve(root, `apps/${platform}/src-tauri/Cargo.toml`), "utf8");
  const cargoVersion = cargo.match(/^version\s*=\s*"([^"]+)"/m)?.[1];
  if (config.version !== packageInfo.version || cargoVersion !== packageInfo.version) {
    failures.push(`${platform} package, Tauri and Cargo versions must agree`);
  }
}
if (macPackage.dependencies?.["@microsoft/office-js"] || macPackage.devDependencies?.["@types/office-js"]) {
  failures.push("The macOS application must not depend on Office.js");
}
if (windowsPackage.dependencies?.["@microsoft/office-js"] || windowsPackage.devDependencies?.["@microsoft/office-js"] || windowsPackage.devDependencies?.["@types/office-js"]) {
  failures.push("The Windows native VSTO build must not depend on retired Office.js packages");
}

for (const path of ["apps/macos", "apps/windows"]) {
  if (lstatSync(resolve(root, path)).isSymbolicLink()) {
    failures.push(`${path} must be a real directory, not a symbolic link`);
  }
}

// Stream NUL-delimited names: Unicode paths and large repositories are ordinary input.
const git = spawn("git", ["ls-files", "-z", "apps/macos", "apps/windows"], { cwd: root });
git.stdout.setEncoding("utf8");
let gitError = "";
git.stderr.setEncoding("utf8");
git.stderr.on("data", chunk => { gitError += chunk; });
const finished = new Promise((resolve, reject) => {
  git.on("error", reject);
  git.on("close", code => code === 0 ? resolve() : reject(new Error(gitError || `git exited ${code}`)));
});
function checkTrackedPath(relativePath) {
  if (/^apps\/(macos|windows)\/office\/vendor\/office-js\//.test(relativePath)) {
    failures.push(`Retired committed Office.js vendor file: ${relativePath}`);
  }
  // Check declared source dependencies, not prose, archived evidence or test launchers.
  if (!/^apps\/(?:macos|windows)\/(?:src\/|src-tauri\/src\/|src-windows\/)/.test(relativePath)
      || /\/(?:bin|obj|artifacts)\//.test(relativePath)
      || !/\.(?:ts|tsx|rs|csproj|projitems)$/.test(relativePath)) return;
  const content = readFileSync(resolve(root, relativePath), "utf8");
  const platform = relativePath.split("/")[1];
  const other = platform === "windows" ? "macos" : "windows";
  const references = content.matchAll(/(?:from\s*|import\s*\(\s*|(?:Include|Project|path)\s*=\s*)["']([^"']+)["']/g);
  for (const [, specifier] of references) {
    if (!specifier.startsWith(".")) continue;
    const target = relative(root, resolve(dirname(resolve(root, relativePath)), specifier.replaceAll("\\", "/"))).replaceAll("\\", "/");
    if (target.startsWith(`apps/${other}/`)) failures.push(`${relativePath} imports the other platform: ${specifier}`);
  }
}
let pending = "";
for await (const chunk of git.stdout) {
  const paths = (pending + chunk).split("\0");
  pending = paths.pop();
  for (const path of paths) checkTrackedPath(path);
}
await finished;
if (pending) throw new Error("Git file listing did not end at a path boundary");

if (failures.length) {
  console.error("VisualTeX repository structure verification failed:\n");
  for (const failure of failures) console.error(`- ${failure}`);
  process.exit(1);
}

console.log("VisualTeX repository structure and legacy Office cleanup are complete.");

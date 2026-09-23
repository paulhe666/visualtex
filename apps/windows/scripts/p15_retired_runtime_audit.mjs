import assert from "node:assert/strict";
import { existsSync } from "node:fs";
import { readFile, readdir } from "node:fs/promises";
import { dirname, extname, join, normalize, resolve } from "node:path";
import process from "node:process";

const root = process.cwd();
const srcRoot = resolve(root, "src");
const roots = [
  resolve(srcRoot, "main.tsx"),
  resolve(srcRoot, "desktop/main.tsx"),
  resolve(srcRoot, "office/dialog/main.tsx"),
];
const sourceExtensions = [".ts", ".tsx", ".js", ".jsx", ".mts", ".mjs"];
const assetExtensions = [".css"];
const bundledFontManifest = JSON.parse(await readFile(resolve(root, "src-tauri/resources/formula-fonts/manifest.json"), "utf8"));
const bundledFontFiles = bundledFontManifest.fonts.map((font) => font.file);
const importPattern = /(?:import|export)\s+(?:type\s+)?(?:[^"']*?\s+from\s+)?["']([^"']+)["']|import\s*\(\s*["']([^"']+)["']\s*\)/gu;
const forbiddenPackages = /^(?:mathlive(?:\/|$)|mathjax-full(?:\/|$))/u;
const forbiddenReachableBasenames = new Set([
  "MathEditor.tsx",
  "mathLiveIntegralCompatibility.ts",
  "mathLiveOptionCompatibility.ts",
  "mathLiveCompatibilityMacros.ts",
  "formulaFontRuntime.ts",
  "greekLetterHotkeyRuntime.ts",
]);

async function resolveRelative(from, specifier) {
  if (!specifier.startsWith(".")) return null;
  const clean = specifier.split("?", 1)[0];
  const base = resolve(dirname(from), clean);
  if (existsSync(base) && extname(base)) return normalize(base);
  for (const extension of [...sourceExtensions, ...assetExtensions]) {
    const candidate = `${base}${extension}`;
    if (existsSync(candidate)) return normalize(candidate);
  }
  for (const extension of sourceExtensions) {
    const candidate = join(base, `index${extension}`);
    if (existsSync(candidate)) return normalize(candidate);
  }
  return null;
}

const visited = new Set();
const packageImports = [];
async function visit(file) {
  file = normalize(file);
  if (visited.has(file)) return;
  visited.add(file);
  const text = await readFile(file, "utf8");
  if (file.endsWith(".css")) return;
  for (const match of text.matchAll(importPattern)) {
    const specifier = match[1] ?? match[2];
    if (!specifier) continue;
    if (forbiddenPackages.test(specifier)) packageImports.push({ file, specifier });
    const resolved = await resolveRelative(file, specifier);
    if (resolved) await visit(resolved);
  }
}
for (const entry of roots) await visit(entry);
assert.deepEqual(packageImports, [], `Production graph imports retired packages: ${JSON.stringify(packageImports)}`);
const forbiddenReachable = [...visited].filter((file) => forbiddenReachableBasenames.has(file.split(/[\\/]/u).at(-1)));
assert.deepEqual(forbiddenReachable, [], `Production graph reaches legacy editor/runtime files: ${JSON.stringify(forbiddenReachable)}`);

const packageJson = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));
assert.equal(packageJson.dependencies?.mathlive, undefined, "mathlive must not be a production dependency");
assert.equal(packageJson.dependencies?.["mathjax-full"], undefined, "mathjax-full must not be a production dependency");
assert.equal(typeof packageJson.devDependencies?.mathlive, "string", true, "legacy MathLive oracle must be explicitly dev-only");
assert.equal(typeof packageJson.devDependencies?.["mathjax-full"], "string", true, "legacy MathJax oracle must be explicitly dev-only");

for (const config of ["vite.config.ts", "vite.office.windows-native.config.ts"]) {
  const text = await readFile(resolve(root, config), "utf8");
  assert.doesNotMatch(text, /mathlive|mathjax/iu, `${config} still configures a retired runtime`);
}

async function walkFiles(dir) {
  const result = [];
  for (const entry of await readdir(dir, { withFileTypes: true })) {
    const path = join(dir, entry.name);
    if (entry.isDirectory()) result.push(...await walkFiles(path));
    else result.push(path);
  }
  return result;
}
const cssFiles = (await walkFiles(srcRoot)).filter((file) => file.endsWith(".css"));
for (const file of cssFiles) {
  const text = await readFile(file, "utf8");
  assert.doesNotMatch(text, /\.ML__|#mathlive-suggestion-popover|#visualtex-native-input-suggestion-popover/iu, `${file} still contains legacy MathLive CSS`);
}

async function auditDist(relative) {
  const dir = resolve(root, relative);
  assert.equal(existsSync(dir), true, `${relative} must exist before P15 audit`);
  const files = await walkFiles(dir);
  const names = files.map((file) => file.replaceAll("\\", "/"));
  const wasm = names.filter((name) => /visualtex_engine_wasm_bg-[^/]+\.wasm$/u.test(name));
  const mathFonts = names.filter((name) => /latinmodern-math-[^/]+\.otf$/u.test(name));
  // KaTeX font files are now explicit VisualTeX app-private data and are not
  // evidence that the retired MathLive editor/runtime returned. Verify the
  // complete licensed private-font manifest separately and reject only actual
  // MathLive/MathJax runtime assets.
  const retiredAssets = names.filter((name) => /mathlive|mathjax/iu.test(name));
  const bundledFonts = bundledFontFiles.map((fileName) => {
    const dot = fileName.lastIndexOf(".");
    const stem = dot >= 0 ? fileName.slice(0, dot) : fileName;
    const extension = dot >= 0 ? fileName.slice(dot) : "";
    const matches = names.filter((name) => name.split("/").at(-1)?.startsWith(`${stem}-`) && name.endsWith(extension));
    assert.equal(matches.length, 1, `${relative} must contain exactly one bundled font asset for ${fileName}`);
    return matches[0].split("/").at(-1);
  });
  assert.equal(wasm.length, 1, `${relative} must contain exactly one core WASM asset`);
  assert.equal(mathFonts.length, 1, `${relative} must contain exactly one verified math font asset`);
  assert.deepEqual(retiredAssets, [], `${relative} contains retired runtime assets: ${JSON.stringify(retiredAssets)}`);
  const textFiles = files.filter((file) => /\.(?:js|css|html)$/u.test(file));
  for (const file of textFiles) {
    const text = await readFile(file, "utf8");
    assert.doesNotMatch(text, /mathjax-full|mathlive\/static\.css|node_modules\/mathlive|MathfieldElement|<math-field/iu, `${relative} emitted a retired runtime reference`);
  }
  return { fileCount: files.length, wasm: wasm[0].split("/").at(-1), mathFont: mathFonts[0].split("/").at(-1), bundledFormulaFonts: bundledFonts.length };
}

const desktop = await auditDist("dist");
const office = await auditDist("dist-office-windows-native");
console.log(JSON.stringify({
  status: "PASS",
  productionEntrypoints: roots.map((entry) => entry.slice(root.length + 1).replaceAll("\\", "/")),
  reachableProductionFiles: visited.size,
  retiredPackageImports: 0,
  retiredLegacyReachableFiles: 0,
  retiredCssSelectors: 0,
  productionDependencies: { mathlive: false, mathjaxFull: false },
  devOracleDependencies: { mathlive: true, mathjaxFull: true },
  desktop,
  office,
}, null, 2));

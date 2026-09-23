import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { dirname, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const windowsRoot = resolve(here, "..");
const repoRoot = resolve(windowsRoot, "../..");
const engineRoot = resolve(repoRoot, "engine");
const stamp = new Date().toISOString().replaceAll(":", "-").replaceAll(".", "-");
const output = resolve(engineRoot, "target/p16-formula-font-truth-review", stamp);
mkdirSync(output, { recursive: true });

if (process.platform !== "win32") {
  throw new Error("P16 formula-font truth review is Windows-only.");
}

const sha256 = (bytes) => createHash("sha256").update(bytes).digest("hex");
const shaFile = (path) => sha256(readFileSync(path));
const rel = (path) => relative(repoRoot, path).replaceAll("\\", "/");
const windowsFonts = resolve(process.env.WINDIR ?? "C:/Windows", "Fonts");

const families = [
  {
    id: "times",
    label: "Times New Roman",
    files: ["times.ttf", "timesi.ttf", "timesbd.ttf", "timesbi.ttf"],
  },
  {
    id: "palatino",
    label: "Palatino Linotype",
    files: ["pala.ttf", "palai.ttf", "palab.ttf", "palabi.ttf"],
  },
  {
    id: "helvetica",
    label: "Arial",
    files: ["arial.ttf", "ariali.ttf", "arialbd.ttf", "arialbi.ttf"],
  },
];

const faceRecords = families.map((family) => ({
  ...family,
  faces: family.files.map((name, index) => {
    const path = resolve(windowsFonts, name);
    assert.ok(existsSync(path), `${family.label} face is missing: ${path}`);
    const stat = statSync(path);
    assert.ok(stat.size > 0, `${family.label} face is empty: ${path}`);
    return {
      role: ["regular", "italic", "bold", "boldItalic"][index],
      name,
      bytes: stat.size,
      sha256: shaFile(path),
    };
  }),
}));

function runCargoTest(name, testFile = "styles") {
  const result = spawnSync(
    "cargo",
    [
      "+1.96.0",
      "test",
      "--manifest-path",
      "engine/Cargo.toml",
      "-p",
      "visualtex-engine-core",
      "--test",
      testFile,
      name,
      "--locked",
      "--",
      "--exact",
    ],
    { cwd: repoRoot, encoding: "utf8", timeout: 120_000, maxBuffer: 16 * 1024 * 1024 },
  );
  const log = `${result.stdout ?? ""}${result.stderr ?? ""}`;
  const logPath = resolve(output, `${name}.log`);
  writeFileSync(logPath, log);
  assert.equal(result.error, undefined, `${name} could not run: ${result.error?.message}`);
  assert.equal(result.status, 0, `${name} failed; see ${rel(logPath)}`);
  assert.match(log, /test result: ok/u, `${name} did not report an OK test result`);
  return { name, status: "PASS", log: rel(logPath), logSha256: shaFile(logPath) };
}

const tests = [
  runCargoTest("product_font_faces_own_normal_symbols_and_common_math_styles"),
  runCargoTest("every_installed_product_family_uses_its_own_four_faces_for_common_math"),
  runCargoTest("mathlive_katex_role_policy_is_reproduced_without_mathlive_runtime", "bundled_tex_fonts_review"),
  runCargoTest("all_shipped_faces_fit_one_bounded_registry_and_use_no_machine_fonts", "bundled_tex_fonts_review"),
];

const runtimeSource = readFileSync(resolve(windowsRoot, "src/engine/coreEngineRuntime.ts"), "utf8");
for (const field of [
  "alphabetUprightFont",
  "alphabetItalicFont",
  "alphabetBoldFont",
  "alphabetBoldItalicFont",
]) {
  assert.ok(runtimeSource.includes(field), `Runtime layout is missing ${field}`);
}
assert.ok(runtimeSource.includes("loadBundledTexFontSet"), "Default KaTeX is not sourced from the VisualTeX private font bundle");
assert.ok(runtimeSource.includes('loadBundledFormulaFont(context.runtime, "STIXTwoMath-Regular")'), "STIX is not available as an app-private mathematical fallback");
assert.ok(runtimeSource.includes('bundled: "termes"'), "Times missing-system fallback is not explicitly TeX Gyre Termes");
assert.ok(runtimeSource.includes('bundled: "pagella"'), "Palatino missing-system fallback is not explicitly TeX Gyre Pagella");
assert.ok(runtimeSource.includes('bundled: "heros"'), "Helvetica/Arial missing-system fallback is not explicitly TeX Gyre Heros");

const privateFontManifest = JSON.parse(readFileSync(resolve(windowsRoot, "src-tauri/resources/formula-fonts/manifest.json"), "utf8"));
assert.equal(privateFontManifest.fonts.length, 37, "VisualTeX private font manifest does not contain the accepted 37 faces");
for (const required of ["KaTeX_Main-Regular.ttf", "KaTeX_Math-Italic.ttf", "KaTeX_Size4-Regular.ttf", "STIXTwoMath-Regular.otf", "NotoSansCJKsc-Regular.otf"]) {
  assert.ok(privateFontManifest.fonts.some((font) => font.file === required), `Private font bundle is missing ${required}`);
}

const settingsSource = readFileSync(resolve(windowsRoot, "src/components/SettingsDialog.tsx"), "utf8");
assert.ok(settingsSource.includes("actualLetterFamily"), "Settings UI does not disclose the actual resolved font family");
assert.ok(settingsSource.includes("软件内置"), "Settings UI does not identify app-private bundled font use");
assert.ok(settingsSource.includes("所选系统字体不可用，已明确使用内置替代字体"), "Settings UI does not disclose explicit fallback identity");

const report = {
  schema: "visualtex-p16-formula-font-truth-v2",
  status: "PASS",
  date: new Date().toISOString(),
  scope: "Windows product-facing formula font identity. VisualTeX ships the complete KaTeX role set and explicit open-font fallbacks app-private; true Windows families use their real Regular/Italic/Bold/BoldItalic faces when present. Font selection changes glyph/layout resources only, never editor syntax or behavior. Stretch capability may fall through to another bundled MATH construction without reintroducing MathLive runtime.",
  families: faceRecords,
  tests,
  privateFontManifest: { fontCount: privateFontManifest.fonts.length, sha256: shaFile(resolve(windowsRoot, "src-tauri/resources/formula-fonts/manifest.json")) },
  defaultKaTeXIsBundled: true,
  bundledStixAvailableOffline: true,
  systemFamiliesPreferredWhenPresent: true,
  explicitNamedFallbackWhenSystemFamilyMissing: true,
  settingsResolvedFamilyDisclosure: true,
  legacyMathLiveRuntimeReintroduced: false,
};
const reportPath = resolve(output, "report.json");
writeFileSync(reportPath, JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify({
  status: report.status,
  report: rel(reportPath),
  receiptSha256: shaFile(reportPath),
  families: report.families.map((family) => family.label),
  tests: report.tests.map((test) => test.name),
}, null, 2));

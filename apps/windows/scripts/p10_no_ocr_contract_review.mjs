import assert from "node:assert/strict";
import { readFileSync } from "node:fs";

const json = (path) => JSON.parse(readFileSync(path, "utf8"));
const base = json("src-tauri/tauri.conf.json");
const noOcr = json("src-tauri/tauri.no-ocr.conf.json");
const fullOcr = json("src-tauri/tauri.ocr-bundle.conf.json");
const hooks = readFileSync("src-tauri/windows/hooks.nsh", "utf8");
const coreSources = [
  "src/editor/CoreMathEditor.tsx",
  "src/engine/coreEngineRuntime.ts",
  "src/engine/coreStaticPreview.ts",
  "../../engine/web/host/editor-host.ts",
].map((path) => [path, readFileSync(path, "utf8")]);

const baseResources = base.bundle?.resources ?? {};
assert.equal(Object.entries(baseResources).some(([key, value]) =>
  /(^|[\\/])ocr([\\/-]|$)/iu.test(String(key)) || /(^|[\\/])ocr([\\/-]|$)/iu.test(String(value))), false,
  "Base application configuration embeds OCR resources");
assert.equal(noOcr.bundle?.resources, undefined, "No-OCR overlay must not add resources");
assert.equal(noOcr.bundle?.windows?.nsis?.template, "./target-no-ocr/nsis-template/visualtex-installer.nsi");
for (const source of ["ocr/worker.py", "resources/ocr-python/windows-x64", "resources/ocr-models/windows-x64"]) {
  assert.ok(Object.hasOwn(fullOcr.bundle?.resources ?? {}, source), `Full OCR overlay is missing ${source}`);
}
assert.ok(hooks.includes("!ifdef VISUALTEX_NO_OCR_BUNDLE"));
assert.ok(hooks.includes('StrCpy $VisualTeXOcrChoice "none"'));
const forbiddenCoreImports = /ocrService|ocr_offline|ocr-model|ocr-python|downloadOcr|recognizeFormulaImage|prepare_offline/iu;
const offenders = coreSources.filter(([, source]) => forbiddenCoreImports.test(source)).map(([path]) => path);
assert.deepEqual(offenders, [], "Core editor/runtime imports local OCR infrastructure");
console.log(JSON.stringify({
  status: "PASS",
  baseHasOcrResources: false,
  noOcrAddsResources: false,
  noOcrTemplate: noOcr.bundle.windows.nsis.template,
  coreImportsLocalOcrRuntime: false,
  installerLevelVerification: "not-run-no-installer-produced",
}, null, 2));

// Stable-source P10 browser/main-app acceptance. This is deliberately not a
// P10-complete receipt: real Tauri close/relaunch and installer-level no-OCR
// remain explicit external gates.
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, readdirSync, mkdirSync, writeFileSync } from 'node:fs';
import { resolve, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';

const engineRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = resolve(engineRoot, '..');
const windowsRoot = resolve(repoRoot, 'apps/windows');
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const output = resolve(engineRoot, 'target/p10-verification', stamp);
mkdirSync(output, { recursive: true });
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const norm = (path) => relative(repoRoot, path).replaceAll('\\', '/');

function inputs() {
  const paths = [];
  const addFile = (path) => paths.push(resolve(repoRoot, path));
  const walk = (path) => {
    const absolute = resolve(repoRoot, path);
    for (const entry of readdirSync(absolute, { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
      const child = resolve(absolute, entry.name);
      if (entry.isDirectory()) walk(norm(child));
      else if (entry.isFile()) paths.push(child);
    }
  };
  for (const folder of ['engine/crates', 'engine/web/host', 'engine/web/sandbox', 'engine/specs', 'engine/tools', 'engine/fixtures']) walk(folder);
  for (const path of ['engine/Cargo.toml', 'engine/Cargo.lock', 'engine/rust-toolchain.toml', 'engine/package.json']) addFile(path);
  for (const path of [
    'apps/windows/package.json',
    'apps/windows/vite.config.ts',
    'apps/windows/src/App.tsx',
    'apps/windows/src/components/MathPreview.tsx',
    'apps/windows/src/editor/CoreMathEditor.tsx',
    'apps/windows/src/editor/CoreMathEditor.disabled.tsx',
    'apps/windows/src/editor/MathEditor.tsx',
    'apps/windows/src/editor/formulaFontPreferences.ts',
    'apps/windows/src/engine/coreCommandAdapter.ts',
    'apps/windows/src/engine/coreEngineRuntime.ts',
    'apps/windows/src/engine/coreStaticPreview.ts',
    'apps/windows/src/engine/coreStaticPreview.disabled.ts',
    'apps/windows/src/history/historyTypes.ts',
    'apps/windows/src/history/documentHistory.ts',
    'apps/windows/src/history/HistoryManager.ts',
    'apps/windows/src/source-editor/LatexSourceEditor.tsx',
    'apps/windows/src/stores/editorStore.ts',
    'apps/windows/src/workspace/EditorWorkspace.tsx',
    'apps/windows/src/workspace/workspaceTypes.ts',
    'apps/windows/src/autocomplete/commandRegistry.ts',
    'apps/windows/src/autocomplete/runtimeCommandRegistry.ts',
    'apps/windows/src/math/customSymbolRegistry.ts',
    'apps/windows/src/math/customSymbolTypes.ts',
    'apps/windows/src-tauri/Cargo.toml',
    'apps/windows/src-tauri/src/lib.rs',
    'apps/windows/src-tauri/src/system_math_glyphs.rs',
    'apps/windows/src-tauri/tauri.conf.json',
    'apps/windows/src-tauri/tauri.no-ocr.conf.json',
    'apps/windows/src-tauri/tauri.ocr-bundle.conf.json',
    'apps/windows/src-tauri/windows/hooks.nsh',
    'apps/windows/scripts/p10_core_editor_smoke.mjs',
    'apps/windows/scripts/p10_build_isolation_review.mjs',
    'apps/windows/scripts/p10_no_ocr_contract_review.mjs',
    'apps/windows/src-tauri/src/bin/system_font_asset_acceptance.rs',
  ]) addFile(path);
  const files = [...new Set(paths.map((path) => resolve(path)))].sort().map((path) => ({ path: norm(path), sha256: sha(readFileSync(path)) }));
  return { files, sha256: sha(JSON.stringify(files)) };
}

const before = inputs();
const steps = [];
const p09ReceiptArgument = process.env.P10_P09_RECEIPT?.trim() || null;
const node = process.execPath;
const npm = process.platform === 'win32' ? (process.env.ComSpec || 'cmd.exe') : 'npm';
const npmPrefix = process.platform === 'win32' ? ['/d', '/s', '/c', 'npm'] : [];
const cargo = process.platform === 'win32' ? 'cargo.exe' : 'cargo';
const commands = [
  ...(p09ReceiptArgument ? [] : [['p09-regression', node, ['tools/p09-verify.mjs'], engineRoot, 300000]]),
  ['engine-build', node, ['tools/build.mjs'], engineRoot, 300000],
  ['selection-token-browser', node, ['tools/p10-selection-token-browser-review.mjs'], engineRoot, 180000],
  ['windows-build-isolation', node, ['scripts/p10_build_isolation_review.mjs'], windowsRoot, 180000],
  ['windows-core-product-browser', node, ['scripts/p10_core_editor_smoke.mjs'], windowsRoot, 180000],
  ['windows-system-font-assets', cargo, ['run', '--quiet', '--manifest-path', 'src-tauri/Cargo.toml', '--features', 'system-font-asset-acceptance', '--bin', 'system_font_asset_acceptance', '--', 'Cambria Math', 'Times New Roman'], windowsRoot, 180000],
  ['windows-system-chinese-font-assets', cargo, ['run', '--quiet', '--manifest-path', 'src-tauri/Cargo.toml', '--features', 'system-font-asset-acceptance', '--bin', 'system_font_asset_acceptance', '--', 'Microsoft YaHei', 'Microsoft JhengHei', 'Noto Sans CJK SC', 'PingFang SC'], windowsRoot, 180000],
  ['no-ocr-source-contract', node, ['scripts/p10_no_ocr_contract_review.mjs'], windowsRoot, 60000],
  ['legacy-editor-regression', npm, [...npmPrefix, 'run', 'test:editor'], windowsRoot, 180000],
  ['legacy-input-behavior', npm, [...npmPrefix, 'run', 'test:input-behavior'], windowsRoot, 180000],
  ['legacy-source-editor', npm, [...npmPrefix, 'run', 'test:source-editor-ux'], windowsRoot, 180000],
  ['legacy-formula-hotkeys', npm, [...npmPrefix, 'run', 'test:formula-hotkeys'], windowsRoot, 180000],
  ['history-manager', npm, [...npmPrefix, 'run', 'test:history'], windowsRoot, 60000],
  ['application-data', npm, [...npmPrefix, 'run', 'test:application-data'], windowsRoot, 60000],
  ['application-configuration', npm, [...npmPrefix, 'run', 'test:application-configuration'], windowsRoot, 60000],
];
let failed = false;
if (p09ReceiptArgument) {
  const startedAt = new Date().toISOString();
  const start = performance.now();
  let log = `START ${startedAt}\nP09 receipt audit ${p09ReceiptArgument}\n`;
  try {
    const receiptPath = resolve(engineRoot, p09ReceiptArgument);
    const receipt = JSON.parse(readFileSync(receiptPath, 'utf8'));
    if (receipt.status !== 'PASS' || receipt.sourceUnchanged !== true || !Array.isArray(receipt.inputs)) {
      throw new Error('Referenced P09 receipt is not a stable PASS receipt');
    }
    for (const file of receipt.inputs) {
      const current = sha(readFileSync(resolve(engineRoot, file.path)));
      if (current !== file.sha256) throw new Error(`P09 source changed after receipt: ${file.path}`);
    }
    log += `PASS sourceSha256=${receipt.sourceSha256} files=${receipt.inputs.length}\n`;
    const step = { name: 'p09-current-receipt', status: 'PASS', durationMs: Math.round(performance.now() - start), exitCode: 0, error: null, log: 'p09-current-receipt.log', logSha256: sha(log), receipt: p09ReceiptArgument };
    steps.push(step); writeFileSync(resolve(output, step.log), log); console.log(JSON.stringify(step));
  } catch (error) {
    failed = true; log += `FAILED ${error instanceof Error ? error.message : String(error)}\n`;
    const step = { name: 'p09-current-receipt', status: 'FAILED', durationMs: Math.round(performance.now() - start), exitCode: 1, error: error instanceof Error ? error.message : String(error), log: 'p09-current-receipt.log', logSha256: sha(log), receipt: p09ReceiptArgument };
    steps.push(step); writeFileSync(resolve(output, step.log), log); console.error(JSON.stringify(step));
  }
}
for (const [name, executable, args, cwd, timeout] of failed ? [] : commands) {
  const startedAt = new Date().toISOString();
  const start = performance.now();
  const result = spawnSync(executable, args, { cwd, encoding: 'utf8', timeout, maxBuffer: 64 * 1024 * 1024 });
  const log = `START ${startedAt}\n${executable} ${args.join(' ')}\n${result.stdout ?? ''}\nSTDERR\n${result.stderr ?? ''}\nERROR ${result.error?.message ?? ''}\nEXIT ${result.status}\n`;
  writeFileSync(resolve(output, `${name}.log`), log);
  const step = {
    name,
    status: result.status === 0 && !result.error ? 'PASS' : 'FAILED',
    durationMs: Math.round(performance.now() - start),
    exitCode: result.status,
    error: result.error?.message ?? null,
    log: `${name}.log`,
    logSha256: sha(log),
  };
  steps.push(step);
  console.log(JSON.stringify(step));
  if (step.status !== 'PASS') {
    failed = true;
    console.error(log.slice(-16000));
    break;
  }
}
const after = inputs();
if (before.sha256 !== after.sha256) failed = true;
const report = {
  schema: 'visualtex-p10-browser-verification-v1',
  status: failed ? 'FAILED' : 'PASS',
  completion: 'browserStable-notP10Complete',
  scope: 'P10 desktop main-app core integration, explicit session preflight/locking, typed core selection ownership, single core history with CodeMirror/header arbitration, natural structural input behaviors, immutable custom symbols, verified Windows font bytes, cached noninteractive core static preview, save/open document generation boundaries, default legacy resource isolation, no-OCR source/config isolation, and passing P09/legacy regressions.',
  sourceSha256: before.sha256,
  sourceUnchanged: before.sha256 === after.sha256,
  inputs: before.files,
  steps,
  notCovered: [
    'Real Tauri process close and second launch: not run because an installed user VisualTeX process was already active and must not be closed or stolen by single-instance dev acceptance.',
    'Installer-level no-OCR size/install verification: no no-OCR installer was produced by this checkpoint.',
    'Office interaction and Office document algorithms: intentionally outside P10.',
  ],
  officeActions: 0,
  wordActions: 0,
  gitMutations: 0,
};
writeFileSync(resolve(output, 'report.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify({ status: report.status, completion: report.completion, receipt: relative(engineRoot, resolve(output, 'report.json')).replaceAll('\\', '/'), sourceSha256: report.sourceSha256, sourceUnchanged: report.sourceUnchanged, steps: steps.length, notCovered: report.notCovered }, null, 2));
if (failed) process.exitCode = 1;

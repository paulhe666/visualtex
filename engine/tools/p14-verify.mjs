import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const engineRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = resolve(engineRoot, '..');
const windowsRoot = resolve(repoRoot, 'apps/windows');
const tauriRoot = resolve(windowsRoot, 'src-tauri');
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const output = resolve(engineRoot, 'target/p14-verification', stamp);
mkdirSync(output, { recursive: true });

const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const norm = path => relative(repoRoot, path).replaceAll('\\', '/');
const inputPaths = [
  'engine/specs/P14-cross-host-resource-performance-hardening.md',
  'engine/specs/P14-windows-host-acceptance.json',
  'engine/tools/p14-verify.mjs',
  'engine/tools/p14-browser-performance-review.mjs',
  'engine/tools/editor-browser-review.mjs',
  'apps/windows/src/engine/coreEngineRuntime.ts',
  'apps/windows/src/editor/CoreMathEditor.tsx',
  'apps/windows/src/office/compat/coreOfficeBatchConversion.ts',
  'apps/windows/scripts/p14_core_resource_consistency.mjs',
  'apps/windows/scripts/p14_core_resource_recovery_browser.mjs',
  'apps/windows/scripts/p14_core_runtime_recovery_browser.mjs',
  'apps/windows/scripts/p14_office_core_race_browser.mjs',
  'apps/windows/scripts/p14_companion_static_resource_probe.mjs',
  'apps/windows/scripts/p14_webview2_live_probe.mjs',
].map(path => resolve(repoRoot, path));
const inputs = () => inputPaths.map(path => ({ path: norm(path), sha256: sha(readFileSync(path)) }));
const digest = files => sha(Buffer.from(JSON.stringify(files)));
const before = inputs();

const node = process.execPath;
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const cargo = process.platform === 'win32' ? 'cargo.exe' : 'cargo';
const tsx = resolve(windowsRoot, 'node_modules/.bin', process.platform === 'win32' ? 'tsx.cmd' : 'tsx');
const tsc = resolve(windowsRoot, 'node_modules/typescript/bin/tsc');

function spawn(executable, args, cwd, timeout = 120_000) {
  const isCmd = process.platform === 'win32' && executable.toLowerCase().endsWith('.cmd');
  const command = isCmd ? (process.env.ComSpec || 'cmd.exe') : executable;
  const commandArgs = isCmd ? ['/d', '/s', '/c', executable, ...args] : args;
  return spawnSync(command, commandArgs, {
    cwd,
    encoding: 'utf8',
    timeout,
    maxBuffer: 128 * 1024 * 1024,
    env: { ...process.env },
  });
}
function run(name, executable, args, cwd, timeout = 120_000) {
  const start = performance.now();
  const result = spawn(executable, args, cwd, timeout);
  const log = `COMMAND ${executable} ${args.join(' ')}\nSTDOUT\n${result.stdout ?? ''}\nSTDERR\n${result.stderr ?? ''}\nERROR ${result.error?.message ?? ''}\nEXIT ${result.status}\n`;
  const logName = `${name}.log`;
  writeFileSync(resolve(output, logName), log);
  const step = {
    name,
    status: result.status === 0 && !result.error ? 'PASS' : 'FAILED',
    durationMs: Math.round(performance.now() - start),
    exitCode: result.status,
    error: result.error?.message ?? null,
    log: logName,
    logSha256: sha(Buffer.from(log)),
  };
  console.log(JSON.stringify(step));
  if (step.status !== 'PASS') throw new Error(`${name} failed; see ${norm(resolve(output, logName))}`);
  return step;
}

const frozenPaths = [
  'apps/windows/src-windows/VisualTeX.WordVsto',
  'apps/windows/src-windows/VisualTeX.PowerPointVsto',
  'apps/windows/src-windows/VisualTeX.WindowsOffice.Contracts',
  'apps/windows/src-windows/VisualTeX.FormulaOleServer',
  'apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared',
  'apps/windows/src-windows/VisualTeX.WindowsOleBridge',
  'apps/windows/src-tauri/src/office/sessions.rs',
  'apps/windows/src-tauri/src/office/server.rs',
  'apps/windows/src-tauri/src/office/lifecycle.rs',
  'apps/windows/src-tauri/src/office/state.rs',
  'apps/windows/src-tauri/src/office/windows_backend.rs',
  'apps/windows/src-tauri/src/office/windows_pipe.rs',
  'apps/windows/src/office/shared',
  'apps/windows/src/office/api',
  'apps/windows/src/office/dialog/useOfficeSession.ts',
];
function frozenOfficeDiff() {
  const result = spawn('git', ['diff', '--name-only', '--', ...frozenPaths], repoRoot, 30_000);
  if (result.status !== 0 || result.error) throw new Error(`Unable to audit frozen Office paths: ${result.stderr || result.error?.message}`);
  return (result.stdout || '').split(/\r?\n/u).map(value => value.trim()).filter(Boolean);
}
function readJson(path) { return JSON.parse(readFileSync(path, 'utf8')); }
function requirePassReceipt(relativePath, label) {
  const path = resolve(repoRoot, relativePath);
  if (!existsSync(path)) throw new Error(`${label} receipt is missing: ${relativePath}`);
  const value = readJson(path);
  if (value.status !== 'PASS') throw new Error(`${label} receipt is not PASS: ${relativePath}`);
  return { path: relativePath, sha256: sha(readFileSync(path)) };
}
function realAcceptanceRecord() {
  const path = resolve(engineRoot, 'specs/P14-windows-host-acceptance.json');
  const data = readJson(path);
  const checks = [
    ['status', data.status === 'PASS'],
    ['resources.desktopAndOfficeBytesMatchSource', data.resources?.desktopAndOfficeBytesMatchSource === true],
    ['resources.runtimeExpectedHashesMatch', data.resources?.runtimeExpectedHashesMatch === true],
    ['resourceFailureRecovery.badWasmFailsClosed', data.resourceFailureRecovery?.badWasmFailsClosed === true],
    ['resourceFailureRecovery.badFontFailsClosed', data.resourceFailureRecovery?.badFontFailsClosed === true],
    ['resourceFailureRecovery.retryWithoutPageReloadOrConfigurationDeletion', data.resourceFailureRecovery?.retryWithoutPageReloadOrConfigurationDeletion === true],
    ['runtimeRecovery.fatalDispatchErrorInvalidatesRuntime', data.runtimeRecovery?.fatalDispatchErrorInvalidatesRuntime === true],
    ['runtimeRecovery.newRuntimeRecreatedFromCheckpoint', data.runtimeRecovery?.newRuntimeRecreatedFromCheckpoint === true],
    ['races.multiWindowSessionIsolation', data.races?.multiWindowSessionIsolation === true],
    ['races.doubleApplyProducesSingleCommit', data.races?.doubleApplyProducesSingleCommit === true],
    ['races.staleWriteAfterClose', data.races?.staleWriteAfterClose === false],
    ['inputAndAccessibility.actualWindowsIme', data.inputAndAccessibility?.actualWindowsIme === true],
    ['inputAndAccessibility.realBrowserClipboard', data.inputAndAccessibility?.realBrowserClipboard === true],
    ['virtualization.formulaCount', data.virtualization?.formulaCount === 1000],
    ['virtualization.disposeActiveSessions', data.virtualization?.disposeActiveSessions === 0],
    ['performance.localEditEventToFrameP95Ms', Number(data.performance?.localEditEventToFrameP95Ms) <= 32],
    ['performance.cursorMoveLayoutDelta', data.performance?.cursorMoveLayoutDelta === 0],
    ['performance.cursorMoveRevisionDelta', data.performance?.cursorMoveRevisionDelta === 0],
    ['performance.headlessMixed100.failures', data.performance?.headlessMixed100?.failures === 0],
    ['performance.headlessOmml1000.failures', data.performance?.headlessOmml1000?.failures === 0],
    ['performance.headlessOmml1000.mathMlOnly', data.performance?.headlessOmml1000?.mathMlOnly === true],
    ['realWindowsWebView2.status', data.realWindowsWebView2?.status === 'PASS'],
    ['realWindowsWebView2.productionOfficeBundleLoaded', data.realWindowsWebView2?.productionOfficeBundleLoaded === true],
    ['realWindowsWebView2.coreEditors', data.realWindowsWebView2?.coreEditors === 1],
    ['realWindowsWebView2.coreInputs', data.realWindowsWebView2?.coreInputs === 1],
    ['realWindowsWebView2.legacyMathFields', data.realWindowsWebView2?.legacyMathFields === 0],
    ['realWindowsWebView2.systemFontProbe.hasMathTable', data.realWindowsWebView2?.systemFontProbe?.hasMathTable === true],
    ['realWindowsWebView2.create.visualTeXOleCount', data.realWindowsWebView2?.create?.visualTeXOleCount === 1],
    ['realWindowsWebView2.create.source', data.realWindowsWebView2?.create?.source === 'abc'],
    ['realWindowsWebView2.edit.sourceAfter', data.realWindowsWebView2?.edit?.sourceAfter === 'abcd'],
    ['realWindowsWebView2.edit.visualTeXOleCountAfter', data.realWindowsWebView2?.edit?.visualTeXOleCountAfter === 1],
    ['realWindowsWebView2.edit.duplicateObjectInserted', data.realWindowsWebView2?.edit?.duplicateObjectInserted === false],
    ['protection.macosOfficeModified', data.protection?.macosOfficeModified === false],
    ['protection.officeServerModified', data.protection?.officeServerModified === false],
    ['protection.wordVstoModified', data.protection?.wordVstoModified === false],
    ['protection.powerPointVstoModified', data.protection?.powerPointVstoModified === false],
    ['protection.installedOfficeUiOverwritten', data.protection?.installedOfficeUiOverwritten === false],
    ['protection.userWordDocumentTouched', data.protection?.userWordDocumentTouched === false],
  ];
  const failed = checks.filter(([, ok]) => !ok).map(([name]) => name);
  if (failed.length) throw new Error(`P14 Windows host acceptance is incomplete: ${failed.join(', ')}`);

  const receipts = [
    requirePassReceipt(data.ocrAsync.browserReceipt, 'P14 OCR browser'),
    requirePassReceipt(data.ocrAsync.parityReceipt, 'P14 OCR native/WASM parity'),
    requirePassReceipt(data.inputAndAccessibility.actualWindowsImeReceipt, 'P14 Windows native IME'),
    requirePassReceipt(data.virtualization.receipt, 'P14 virtual viewport'),
    requirePassReceipt(data.performance.browserReceipt, 'P14 browser performance'),
  ];
  const currentWasm = resolve(windowsRoot, 'dist-office-windows-native/assets', data.resources.wasmAsset);
  const currentFont = resolve(windowsRoot, 'dist-office-windows-native/assets', data.resources.fontAsset);
  if (existsSync(currentWasm) && sha(readFileSync(currentWasm)) !== data.resources.wasmSha256) {
    throw new Error('P14 Office WASM bytes no longer match the frozen real-host acceptance.');
  }
  if (existsSync(currentFont) && sha(readFileSync(currentFont)) !== data.resources.fontSha256) {
    throw new Error('P14 Office font bytes no longer match the frozen real-host acceptance.');
  }
  return { status: 'PASS', path: norm(path), sha256: sha(readFileSync(path)), receipts };
}

const steps = [];
let failure = null;
let realAcceptance = null;
let frozenBefore = [];
let frozenAfter = [];
try {
  realAcceptance = realAcceptanceRecord();
  frozenBefore = frozenOfficeDiff();
  if (frozenBefore.length) throw new Error(`Protected Office tracked paths are dirty before P14 verification: ${frozenBefore.join(', ')}`);
  steps.push(run('windows-typescript', node, [tsc, '--noEmit'], windowsRoot));
  steps.push(run('office-production-bundle', npm, ['run', 'build:office:windows-native'], windowsRoot, 180_000));
  steps.push(run('desktop-core-build-isolation', node, ['scripts/p10_build_isolation_review.mjs'], windowsRoot, 180_000));
  steps.push(run('p14-resource-consistency', node, ['scripts/p14_core_resource_consistency.mjs'], windowsRoot));
  steps.push(run('desktop-core-product-smoke', node, ['scripts/p10_core_editor_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('office-interactive-core-guard', node, ['scripts/p13_office_interactive_core_guard.mjs'], windowsRoot, 120_000));
  steps.push(run('office-headless-converter-browser', node, ['scripts/p13_office_core_converter_browser_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('p14-resource-recovery', node, ['scripts/p14_core_resource_recovery_browser.mjs'], windowsRoot, 120_000));
  steps.push(run('p14-runtime-recovery', node, ['scripts/p14_core_runtime_recovery_browser.mjs'], windowsRoot, 120_000));
  steps.push(run('p14-office-races', node, ['scripts/p14_office_core_race_browser.mjs'], windowsRoot, 120_000));
  steps.push(run('debug-companion-build', cargo, ['build', '--manifest-path', 'src-tauri/Cargo.toml'], windowsRoot, 240_000));
  steps.push(run('p14-companion-static-resources', node, ['scripts/p14_companion_static_resource_probe.mjs'], windowsRoot, 120_000));
  steps.push(run('core-resource-and-fault-tests', cargo, ['test', '-p', 'visualtex-engine-core', '--test', 'font_resources', '--test', 'ocr_target_review', '--test', 'parser', '--test', 'session_contract_review', '--test', 'structural_clipboard_review'], engineRoot, 180_000));
  steps.push(run('ocr-target-browser', node, ['tools/ocr-target-browser-review.mjs'], engineRoot, 120_000));
  steps.push(run('ocr-target-native-wasm-parity', node, ['tools/ocr-target-parity-review.mjs'], engineRoot, 120_000));
  steps.push(run('editor-browser-input-accessibility', node, ['tools/editor-browser-review.mjs'], engineRoot, 120_000));
  steps.push(run('structured-clipboard-zoom-browser', node, ['tools/grid-rectangle-browser-review.mjs'], engineRoot, 120_000));
  steps.push(run('virtual-viewport-1000', node, ['tools/virtual-viewport-browser-review.mjs'], engineRoot, 180_000));
  steps.push(run('windows-native-ime', node, ['tools/owned-ime-review.mjs'], engineRoot, 180_000));
  steps.push(run('p14-browser-performance', node, ['tools/p14-browser-performance-review.mjs'], engineRoot, 120_000));
  steps.push(run('p13-headless-performance', node, ['scripts/p13_office_core_converter_performance.mjs'], windowsRoot, 120_000));
  steps.push(run('existing-office-formula-document', tsx, ['scripts/office_formula_document_regression.mts'], windowsRoot));
  steps.push(run('existing-office-session-safety', tsx, ['scripts/office_session_payload_safety_regression.mts'], windowsRoot));
  steps.push(run('existing-office-apply-shortcut', tsx, ['scripts/office_apply_shortcut_regression.mts'], windowsRoot));
} catch (error) {
  failure = error instanceof Error ? error.message : String(error);
}
try {
  frozenAfter = frozenOfficeDiff();
  if (!failure && frozenAfter.length) failure = `Protected Office tracked paths are dirty after P14 verification: ${frozenAfter.join(', ')}`;
} catch (error) {
  if (!failure) failure = error instanceof Error ? error.message : String(error);
}
const frozenOfficeWorkingTreeUnchanged = JSON.stringify(frozenBefore) === JSON.stringify(frozenAfter);
if (!failure && !frozenOfficeWorkingTreeUnchanged) failure = 'Protected Office working-tree snapshot changed during P14 verification.';
const after = inputs();
const sourceUnchanged = digest(before) === digest(after);
if (!failure && !sourceUnchanged) failure = 'P14 verifier source inputs changed while the verification was running.';
const branch = (spawn('git', ['branch', '--show-current'], repoRoot, 30_000).stdout || '').trim();
const head = (spawn('git', ['rev-parse', 'HEAD'], repoRoot, 30_000).stdout || '').trim();
const report = {
  schema: 'visualtex-p14-cross-host-hardening-verification-v1',
  status: failure ? 'FAILED' : 'PASS',
  completion: failure ? 'incomplete' : 'complete',
  p15Permitted: !failure,
  branch,
  head,
  sourceSha256: digest(before),
  sourceUnchanged,
  inputs: before,
  steps,
  realWindowsHostAcceptance: realAcceptance,
  frozenOfficeTrackedDiffAtStart: frozenBefore,
  frozenOfficeTrackedDiffAtEnd: frozenAfter,
  frozenOfficeWorkingTreeUnchanged,
  scope: 'P14 Windows/native-WASM/browser/WebView2 resource identity, failure recovery, race, input, virtualization and performance hardening. macOS/WKWebView remains untested and unchanged.',
  notClaimed: [
    'macOS/WKWebView P14 acceptance',
    'installed NSIS/no-OCR candidate acceptance',
    'P15 legacy-engine dependency removal',
    'P16 final package acceptance',
  ],
  failure,
};
writeFileSync(resolve(output, 'report.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify({
  status: report.status,
  completion: report.completion,
  receipt: norm(resolve(output, 'report.json')),
  sourceSha256: report.sourceSha256,
  sourceUnchanged,
  steps: steps.length,
  realWindowsHostAcceptance: realAcceptance?.status ?? null,
  frozenOfficeWorkingTreeUnchanged,
  p15Permitted: report.p15Permitted,
  failure,
}, null, 2));
if (report.status !== 'PASS') process.exitCode = 1;

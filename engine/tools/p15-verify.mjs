import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const engineRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = resolve(engineRoot, '..');
const windowsRoot = resolve(repoRoot, 'apps/windows');
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const output = resolve(engineRoot, 'target/p15-verification', stamp);
mkdirSync(output, { recursive: true });

const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const norm = path => relative(repoRoot, path).replaceAll('\\', '/');
const inputPaths = [
  'engine/specs/P15-retired-runtime-removal.md',
  'engine/specs/P15-windows-acceptance.json',
  'engine/tools/p15-verify.mjs',
  'apps/windows/package.json',
  'apps/windows/package-lock.json',
  'apps/windows/vite.config.ts',
  'apps/windows/vite.office.windows-native.config.ts',
  'apps/windows/src/workspace/EditorWorkspace.tsx',
  'apps/windows/src/editor/CoreMathEditor.tsx',
  'apps/windows/src/components/MathPreview.tsx',
  'apps/windows/src/clipboard/LatexCopyService.ts',
  'apps/windows/src/export/runtime.ts',
  'apps/windows/src/math/latexCompatibility.ts',
  'apps/windows/src/math/packageMacroCompatibility.ts',
  'apps/windows/scripts/p15_retired_runtime_audit.mjs',
  'apps/windows/scripts/p15_core_export_browser_smoke.mjs',
  'apps/windows/scripts/p15_core_multiline_product_smoke.mjs',
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
function requireAcceptance() {
  const path = resolve(engineRoot, 'specs/P15-windows-acceptance.json');
  const data = readJson(path);
  const checks = [
    ['status', data.status === 'PASS'],
    ['branch', data.branch === 'reference-3d207d7'],
    ['head', data.head === 'd1f0aff8db9275c90dc18c7318dc2a4d7bbbd760'],
    ['productionDependencyAudit.mathlivePackageImports', data.productionDependencyAudit?.mathlivePackageImports === 0],
    ['productionDependencyAudit.mathjaxPackageImports', data.productionDependencyAudit?.mathjaxPackageImports === 0],
    ['productionDependencyAudit.legacyEditorOrRuntimeReachableFiles', data.productionDependencyAudit?.legacyEditorOrRuntimeReachableFiles === 0],
    ['productionDependencyAudit.retiredCssSelectors', data.productionDependencyAudit?.retiredCssSelectors === 0],
    ['resources.retiredRuntimeAssetsInDesktop', data.resources?.retiredRuntimeAssetsInDesktop === 0],
    ['resources.retiredRuntimeAssetsInOffice', data.resources?.retiredRuntimeAssetsInOffice === 0],
    ['resources.desktopAndOfficeWasmIdentical', data.resources?.desktopAndOfficeWasmIdentical === true],
    ['resources.desktopAndOfficeFontIdentical', data.resources?.desktopAndOfficeFontIdentical === true],
    ['coreOnlyProduct.legacyMathFieldsInProductSmoke', data.coreOnlyProduct?.legacyMathFieldsInProductSmoke === 0],
    ['coreOnlyProduct.unsupportedSourceRemainsCoreAcrossReload', data.coreOnlyProduct?.unsupportedSourceRemainsCoreAcrossReload === true],
    ['coreOnlyProduct.outerMultiFormulaSmoke', data.coreOnlyProduct?.outerMultiFormulaSmoke === 'PASS'],
    ['exportRuntime.status', data.exportRuntime?.status === 'PASS'],
    ['exportRuntime.retiredRuntimeResourceRequests', data.exportRuntime?.retiredRuntimeResourceRequests === 0],
    ['exportRuntime.placeholderFailsClosed', data.exportRuntime?.placeholderFailsClosed === true],
    ['exportRuntime.unknownCommandFailsClosed', data.exportRuntime?.unknownCommandFailsClosed === true],
    ['officeHeadless.status', data.officeHeadless?.status === 'PASS'],
    ['officeHeadless.mathFields', data.officeHeadless?.mathFields === 0],
    ['performance.localEditEventToFrameP95Ms', Number(data.performance?.localEditEventToFrameP95Ms) <= 32],
    ['performance.cursorMoveLayoutDelta', data.performance?.cursorMoveLayoutDelta === 0],
    ['performance.cursorMoveRevisionDelta', data.performance?.cursorMoveRevisionDelta === 0],
    ['performance.headlessMixed100.failures', data.performance?.headlessMixed100?.failures === 0],
    ['performance.headlessOmml1000.failures', data.performance?.headlessOmml1000?.failures === 0],
    ['hardeningRevalidation.physicalWindowsIme', data.hardeningRevalidation?.physicalWindowsIme === 'PASS'],
    ['noOcr.frontendContract', data.noOcr?.frontendContract === 'PASS'],
    ['licenseAndData.latinModernMathLicense', data.licenseAndData?.latinModernMathLicense === 'GUST Font License'],
    ['licenseAndData.rareIntegralDistributionRegression', data.licenseAndData?.rareIntegralDistributionRegression === 'PASS'],
    ['realWindowsWebView2.status', data.realWindowsWebView2?.status === 'PASS'],
    ['realWindowsWebView2.officeScript', data.realWindowsWebView2?.officeScript === data.resources?.officeJs],
    ['realWindowsWebView2.coreEditors', data.realWindowsWebView2?.coreEditors === 1],
    ['realWindowsWebView2.coreInputs', data.realWindowsWebView2?.coreInputs === 1],
    ['realWindowsWebView2.legacyMathFields', data.realWindowsWebView2?.legacyMathFields === 0],
    ['realWindowsWebView2.create.source', data.realWindowsWebView2?.create?.source === 'abc'],
    ['realWindowsWebView2.create.visualTeXOleCount', data.realWindowsWebView2?.create?.visualTeXOleCount === 1],
    ['realWindowsWebView2.edit.sourceAfter', data.realWindowsWebView2?.edit?.sourceAfter === 'abcd'],
    ['realWindowsWebView2.edit.visualTeXOleCountAfter', data.realWindowsWebView2?.edit?.visualTeXOleCountAfter === 1],
    ['realWindowsWebView2.edit.duplicateObjectInserted', data.realWindowsWebView2?.edit?.duplicateObjectInserted === false],
    ['realWindowsWebView2.ownedWordClosedWithoutSave', data.realWindowsWebView2?.ownedWordClosedWithoutSave === true],
    ['realWindowsWebView2.ownedCompanionStopped', data.realWindowsWebView2?.ownedCompanionStopped === true],
    ['protection.userWordDocumentTouched', data.protection?.userWordDocumentTouched === false],
    ['protection.installedOfficeUiOverwritten', data.protection?.installedOfficeUiOverwritten === false],
    ['protection.installerBuilt', data.protection?.installerBuilt === false],
    ['protection.installerInstalled', data.protection?.installerInstalled === false],
    ['protection.releasePublished', data.protection?.releasePublished === false],
    ['p16Permitted', data.p16Permitted === true],
    ['p16InstallerMutationPermittedWithoutNewUserApproval', data.p16InstallerMutationPermittedWithoutNewUserApproval === false],
  ];
  const failed = checks.filter(([, ok]) => !ok).map(([name]) => name);
  if (failed.length) throw new Error(`P15 acceptance is incomplete: ${failed.join(', ')}`);

  const assets = [
    [resolve(windowsRoot, 'dist/assets', data.resources.wasmAsset), data.resources.wasmSha256, 'desktop WASM'],
    [resolve(windowsRoot, 'dist/assets', data.resources.fontAsset), data.resources.fontSha256, 'desktop font'],
    [resolve(windowsRoot, 'dist/assets', data.resources.desktopJs), data.resources.desktopJsSha256, 'desktop JS'],
    [resolve(windowsRoot, 'dist-office-windows-native/assets', data.resources.wasmAsset), data.resources.wasmSha256, 'Office WASM'],
    [resolve(windowsRoot, 'dist-office-windows-native/assets', data.resources.fontAsset), data.resources.fontSha256, 'Office font'],
    [resolve(windowsRoot, 'dist-office-windows-native/assets', data.resources.officeJs), data.resources.officeJsSha256, 'Office JS'],
  ];
  for (const [asset, expected, label] of assets) {
    if (!existsSync(asset)) throw new Error(`${label} is missing: ${norm(asset)}`);
    if (sha(readFileSync(asset)) !== expected) throw new Error(`${label} hash differs from final P15 acceptance.`);
  }
  for (const required of [data.licenseAndData.latinModernLicenseFile, data.licenseAndData.latinModernReadme, data.licenseAndData.stixLicenseFile]) {
    if (!existsSync(resolve(repoRoot, required))) throw new Error(`Required license/provenance file missing: ${required}`);
  }
  return { path: norm(path), sha256: sha(readFileSync(path)) };
}

const steps = [];
let failure = null;
let acceptance = null;
let frozenBefore = [];
let frozenAfter = [];
try {
  acceptance = requireAcceptance();
  frozenBefore = frozenOfficeDiff();
  if (frozenBefore.length) throw new Error(`Protected Office tracked paths are dirty before P15 verification: ${frozenBefore.join(', ')}`);
  steps.push(run('windows-typescript', node, [tsc, '--noEmit'], windowsRoot));
  steps.push(run('production-build-all', npm, ['run', 'build:all'], windowsRoot, 180_000));
  steps.push(run('p15-retired-runtime-audit', node, ['scripts/p15_retired_runtime_audit.mjs'], windowsRoot));
  steps.push(run('p15-core-export-browser', node, ['scripts/p15_core_export_browser_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('desktop-core-product-smoke', node, ['scripts/p10_core_editor_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('p15-outer-multiformula-product', node, ['scripts/p15_core_multiline_product_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('renderer-independent-math-variant', tsx, ['scripts/math_variant_compatibility_regression.mts'], windowsRoot));
  steps.push(run('latex-format', tsx, ['scripts/latex_format_smoke.mjs'], windowsRoot));
  steps.push(run('source-structural', tsx, ['scripts/source_structural_draft_regression.mts'], windowsRoot));
  steps.push(run('core-custom-symbol-native', cargo, ['test', '-p', 'visualtex-engine-core', '--test', 'custom_symbols_review'], engineRoot, 120_000));
  steps.push(run('core-custom-symbol-parity', node, ['tools/custom-symbol-parity-review.mjs'], engineRoot, 120_000));
  steps.push(run('core-custom-symbol-browser', node, ['tools/custom-symbol-browser-review.mjs'], engineRoot, 120_000));
  steps.push(run('office-interactive-core-guard', node, ['scripts/p13_office_interactive_core_guard.mjs'], windowsRoot, 120_000));
  steps.push(run('office-headless-converter-browser', node, ['scripts/p13_office_core_converter_browser_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('p14-resource-recovery', node, ['scripts/p14_core_resource_recovery_browser.mjs'], windowsRoot, 120_000));
  steps.push(run('p14-runtime-recovery', node, ['scripts/p14_core_runtime_recovery_browser.mjs'], windowsRoot, 120_000));
  steps.push(run('p14-office-races', node, ['scripts/p14_office_core_race_browser.mjs'], windowsRoot, 120_000));
  steps.push(run('p13-headless-performance', node, ['scripts/p13_office_core_converter_performance.mjs'], windowsRoot, 120_000));
  steps.push(run('office-formula-document', tsx, ['scripts/office_formula_document_regression.mts'], windowsRoot));
  steps.push(run('office-session-safety', tsx, ['scripts/office_session_payload_safety_regression.mts'], windowsRoot));
  steps.push(run('office-apply-shortcut', tsx, ['scripts/office_apply_shortcut_regression.mts'], windowsRoot));
  steps.push(run('document-import-parser', tsx, ['scripts/document_import_parser_coverage.mts'], windowsRoot));
  steps.push(run('no-ocr-contract', node, ['scripts/p10_no_ocr_contract_review.mjs'], windowsRoot));
  steps.push(run('integral-license-distribution', tsx, ['scripts/integral_distribution_regression.mts'], windowsRoot));
  steps.push(run('ocr-target-browser', node, ['tools/ocr-target-browser-review.mjs'], engineRoot, 120_000));
  steps.push(run('ocr-target-native-wasm-parity', node, ['tools/ocr-target-parity-review.mjs'], engineRoot, 120_000));
  steps.push(run('structured-clipboard-zoom-browser', node, ['tools/grid-rectangle-browser-review.mjs'], engineRoot, 120_000));
  steps.push(run('virtual-viewport-1000', node, ['tools/virtual-viewport-browser-review.mjs'], engineRoot, 180_000));
  steps.push(run('windows-native-ime', node, ['tools/owned-ime-review.mjs'], engineRoot, 180_000));
  steps.push(run('browser-performance', node, ['tools/p14-browser-performance-review.mjs'], engineRoot, 120_000));
} catch (error) {
  failure = error instanceof Error ? error.message : String(error);
}
try {
  frozenAfter = frozenOfficeDiff();
  if (!failure && frozenAfter.length) failure = `Protected Office tracked paths are dirty after P15 verification: ${frozenAfter.join(', ')}`;
} catch (error) {
  if (!failure) failure = error instanceof Error ? error.message : String(error);
}
const frozenOfficeWorkingTreeUnchanged = JSON.stringify(frozenBefore) === JSON.stringify(frozenAfter);
if (!failure && !frozenOfficeWorkingTreeUnchanged) failure = 'Protected Office working-tree snapshot changed during P15 verification.';
const after = inputs();
const sourceUnchanged = digest(before) === digest(after);
if (!failure && !sourceUnchanged) failure = 'P15 verifier source inputs changed while verification was running.';
const branch = (spawn('git', ['branch', '--show-current'], repoRoot, 30_000).stdout || '').trim();
const head = (spawn('git', ['rev-parse', 'HEAD'], repoRoot, 30_000).stdout || '').trim();
if (!failure && branch !== 'reference-3d207d7') failure = `Unexpected branch: ${branch}`;
if (!failure && head !== 'd1f0aff8db9275c90dc18c7318dc2a4d7bbbd760') failure = `Unexpected HEAD: ${head}`;
const report = {
  schema: 'visualtex-p15-retired-runtime-verification-v1',
  status: failure ? 'FAILED' : 'PASS',
  generatedAt: new Date().toISOString(),
  branch,
  head,
  acceptance,
  sourceSha256: digest(before),
  sourceUnchanged,
  inputs: before,
  frozenOfficeBefore: frozenBefore,
  frozenOfficeAfter: frozenAfter,
  frozenOfficeWorkingTreeUnchanged,
  steps,
  p16Permitted: !failure,
  installerMutationPermittedWithoutNewUserApproval: false,
  failure,
};
const reportPath = resolve(output, 'report.json');
writeFileSync(reportPath, `${JSON.stringify(report, null, 2)}\n`);
const receiptSha256 = sha(readFileSync(reportPath));
console.log(JSON.stringify({ status: report.status, report: norm(reportPath), receiptSha256, sourceSha256: report.sourceSha256, sourceUnchanged, steps: `${steps.filter(step => step.status === 'PASS').length}/${steps.length}`, frozenOfficeWorkingTreeUnchanged, p16Permitted: report.p16Permitted, failure }, null, 2));
if (failure) process.exitCode = 1;

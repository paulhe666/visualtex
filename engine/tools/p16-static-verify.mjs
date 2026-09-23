import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const engineRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = resolve(engineRoot, '..');
const windowsRoot = resolve(repoRoot, 'apps/windows');
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const output = resolve(engineRoot, 'target/p16-static-verification', stamp);
mkdirSync(output, { recursive: true });

const acceptancePath = resolve(engineRoot, 'specs/P16-windows-package-acceptance.json');
const acceptance = JSON.parse(readFileSync(acceptancePath, 'utf8'));
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const shaFile = path => sha(readFileSync(path));
const norm = path => relative(repoRoot, path).replaceAll('\\', '/');
const node = process.execPath;
const powershell = 'C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe';

function must(condition, message) {
  if (!condition) throw new Error(message);
}
function pathFor(relativePath) { return resolve(repoRoot, relativePath); }
function assertArtifact(record, label) {
  const path = pathFor(record.path);
  must(existsSync(path), `${label} missing: ${record.path}`);
  const stat = statSync(path);
  if (record.bytes !== undefined) must(stat.size === record.bytes, `${label} size mismatch: ${stat.size} != ${record.bytes}`);
  must(shaFile(path).toUpperCase() === record.sha256.toUpperCase(), `${label} SHA mismatch`);
  return { path: record.path, bytes: stat.size, sha256: shaFile(path) };
}
function run(name, executable, args, cwd, timeout = 180_000) {
  const start = performance.now();
  const result = spawnSync(executable, args, { cwd, encoding: 'utf8', timeout, maxBuffer: 128 * 1024 * 1024, env: { ...process.env } });
  const log = `COMMAND ${executable} ${args.join(' ')}\nSTDOUT\n${result.stdout ?? ''}\nSTDERR\n${result.stderr ?? ''}\nERROR ${result.error?.message ?? ''}\nEXIT ${result.status}\n`;
  const logName = `${name}.log`;
  writeFileSync(resolve(output, logName), log);
  const step = { name, status: result.status === 0 && !result.error ? 'PASS' : 'FAILED', durationMs: Math.round(performance.now() - start), exitCode: result.status, error: result.error?.message ?? null, log: logName, logSha256: sha(Buffer.from(log)) };
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
  const result = spawnSync('git', ['diff', '--name-only', '--', ...frozenPaths], { cwd: repoRoot, encoding: 'utf8', timeout: 30_000 });
  if (result.status !== 0 || result.error) throw new Error(`Unable to audit frozen Office paths: ${result.stderr || result.error?.message}`);
  return (result.stdout || '').split(/\r?\n/u).map(value => value.trim()).filter(Boolean);
}

const sourceInputs = [
  'engine/specs/P16-windows-package-acceptance.json',
  'engine/tools/p16-static-verify.mjs',
  'apps/windows/scripts/build_platform_bundle.mjs',
  'apps/windows/scripts/tauri_build.mjs',
  'apps/windows/scripts/verify_fixed_windows_office_inputs.ps1',
  'apps/windows/scripts/verify_windows_release_artifacts.ps1',
  'apps/windows/scripts/verify_windows_no_ocr_installer.ps1',
  'apps/windows/scripts/p16_core_visual_regression_browser_review.mjs',
  'apps/windows/scripts/p16_formula_font_truth_review.mjs',
  'apps/windows/scripts/p16_private_font_product_review.mjs',
  'apps/windows/scripts/p15_retired_runtime_audit.mjs',
  'apps/windows/scripts/verify_embedded_frontend_assets.mjs',
  'apps/windows/src-tauri/resources/formula-fonts/manifest.json',
  'apps/windows/src-tauri/tauri.conf.json',
].map(pathFor);
const hashInputs = () => sourceInputs.map(path => ({ path: norm(path), sha256: shaFile(path) }));
const before = hashInputs();
const frozenBefore = frozenOfficeDiff();
let failure = null;
const steps = [];
const evidence = {};

try {
  must(acceptance.status === 'STATIC_PASS', 'P16 package acceptance is not STATIC_PASS.');
  must(acceptance.overallComplete === false, 'Static package acceptance must not claim overall completion.');
  must(acceptance.installedAcceptance?.status === 'NOT_RUN', 'Static verifier expects installed acceptance to remain NOT_RUN.');

  evidence.fullInstaller = assertArtifact(acceptance.packages.full, 'Full installer');
  evidence.noOcrInstaller = assertArtifact(acceptance.packages.noOcr, 'No-OCR installer');
  must(evidence.fullInstaller.bytes - evidence.noOcrInstaller.bytes === acceptance.packages.noOcr.bytesSavedVsFull, 'No-OCR size reduction changed.');

  const isolatedNoOcr = pathFor(acceptance.packages.noOcr.isolatedBuiltPath);
  must(existsSync(isolatedNoOcr), 'Isolated No-OCR build artifact is missing.');
  must(shaFile(isolatedNoOcr) === evidence.noOcrInstaller.sha256, 'Copied No-OCR delivery installer differs from isolated build artifact.');

  const fullNsi = pathFor(acceptance.packages.full.installerNsi);
  const noOcrNsi = pathFor(acceptance.packages.noOcr.installerNsi);
  must(shaFile(fullNsi).toUpperCase() === acceptance.packages.full.installerNsiSha256.toUpperCase(), 'Full generated installer.nsi SHA mismatch.');
  must(shaFile(noOcrNsi).toUpperCase() === acceptance.packages.noOcr.installerNsiSha256.toUpperCase(), 'No-OCR generated installer.nsi SHA mismatch.');
  const fullText = readFileSync(fullNsi, 'utf8');
  const noOcrText = readFileSync(noOcrNsi, 'utf8');
  const normalizeNsisPathText = value => value.replace(/\\+/gu, '/');
  const normalizedFullText = normalizeNsisPathText(fullText);
  const normalizedNoOcrText = normalizeNsisPathText(noOcrText);
  const commonTokens = [
    acceptance.coreResources.wasm.asset,
    acceptance.coreResources.mathFont.asset,
    'VisualTeX-WindowsOffice-VSTO-x64.msi',
    'VisualTeX-WindowsOffice-VSTO-x86.msi',
    'visualtex-windows-office-bridge',
    'mathtype-runtime\\visualtex-runtime-manifest.json',
  ];
  for (const token of commonTokens) {
    const normalizedToken = normalizeNsisPathText(token);
    must(normalizedFullText.includes(normalizedToken), `Full generated NSIS is missing ${normalizedToken}`);
    must(normalizedNoOcrText.includes(normalizedToken), `No-OCR generated NSIS is missing ${normalizedToken}`);
  }
  for (const token of ['ocr\\worker.py', 'ocr-python\\windows-x64', 'ocr-models\\windows-x64\\catalog.json']) {
    const normalizedToken = normalizeNsisPathText(token);
    must(normalizedFullText.includes(normalizedToken), `Full generated NSIS is missing ${normalizedToken}`);
    must(!normalizedNoOcrText.includes(normalizedToken), `No-OCR generated NSIS unexpectedly contains ${normalizedToken}`);
  }
  must(noOcrText.includes('VISUALTEX_NO_OCR_BUNDLE'), 'No-OCR generated NSIS lacks compile-time exclusion marker.');
  must(!fullText.includes('!define VISUALTEX_NO_OCR_BUNDLE'), 'Full generated NSIS unexpectedly uses no-OCR marker.');
  const formulaFontLicenses = acceptance.bundledFormulaFonts?.licenses ?? [];
  for (const record of formulaFontLicenses) {
    assertArtifact(record, `Formula font license ${record.installedPath}`);
    const installedToken = normalizeNsisPathText(record.installedPath);
    must(normalizedFullText.includes(installedToken), `Full generated NSIS is missing formula font license ${installedToken}`);
    must(normalizedNoOcrText.includes(installedToken), `No-OCR generated NSIS is missing formula font license ${installedToken}`);
  }
  evidence.generatedNsi = { fullSha256: shaFile(fullNsi), noOcrSha256: shaFile(noOcrNsi), commonResourcesPresent: true, noOcrResourceExclusion: true, formulaFontLicenses: formulaFontLicenses.length };

  const coreWasm = pathFor(`apps/windows/dist/assets/${acceptance.coreResources.wasm.asset}`);
  const coreFont = pathFor(`apps/windows/dist/assets/${acceptance.coreResources.mathFont.asset}`);
  must(shaFile(coreWasm).toUpperCase() === acceptance.coreResources.wasm.sha256.toUpperCase(), 'Core WASM SHA changed.');
  must(shaFile(coreFont).toUpperCase() === acceptance.coreResources.mathFont.sha256.toUpperCase(), 'Core math font SHA changed.');

  const bundledFontManifestRecord = acceptance.bundledFormulaFonts?.manifest;
  must(bundledFontManifestRecord?.fontCount === 37, 'P16 acceptance must pin exactly 37 app-private formula fonts.');
  const bundledFontManifestArtifact = assertArtifact(bundledFontManifestRecord, 'Bundled formula font manifest');
  const bundledFontManifest = JSON.parse(readFileSync(pathFor(bundledFontManifestRecord.path), 'utf8'));
  must(Array.isArray(bundledFontManifest.fonts) && bundledFontManifest.fonts.length === bundledFontManifestRecord.fontCount, 'Bundled formula font manifest count mismatch.');
  const desktopFontAssetDir = pathFor('apps/windows/dist/assets');
  const officeFontAssetDir = pathFor('apps/windows/dist-office-windows-native/assets');
  const findBuiltFont = (dir, file) => {
    const dot = file.lastIndexOf('.');
    const stem = dot >= 0 ? file.slice(0, dot) : file;
    const extension = dot >= 0 ? file.slice(dot) : '';
    return readdirSync(dir).filter(name => name.startsWith(`${stem}-`) && name.endsWith(extension));
  };
  const desktopFonts = readdirSync(desktopFontAssetDir).filter(name => /\.(?:ttf|otf)$/iu.test(name));
  const officeFonts = readdirSync(officeFontAssetDir).filter(name => /\.(?:ttf|otf)$/iu.test(name));
  must(desktopFonts.length === acceptance.bundledFormulaFonts.desktopEmbeddedFontCount, `Desktop bundled font count mismatch: ${desktopFonts.length}`);
  must(officeFonts.length === acceptance.bundledFormulaFonts.officeCompanionFontCount, `Office companion bundled font count mismatch: ${officeFonts.length}`);
  for (const font of bundledFontManifest.fonts) {
    const sourcePath = pathFor(`apps/windows/src-tauri/resources/formula-fonts/${font.file}`);
    must(existsSync(sourcePath), `Bundled formula font source missing: ${font.file}`);
    must(statSync(sourcePath).size === font.bytes, `Bundled formula font source size mismatch: ${font.file}`);
    must(shaFile(sourcePath) === font.sha256, `Bundled formula font source SHA mismatch: ${font.file}`);
    for (const [label, dir] of [['desktop', desktopFontAssetDir], ['office', officeFontAssetDir]]) {
      const matches = findBuiltFont(dir, font.file);
      must(matches.length === 1, `${label} build must contain exactly one emitted ${font.file}; found ${matches.length}`);
      must(shaFile(resolve(dir, matches[0])) === font.sha256, `${label} emitted formula font SHA mismatch: ${font.file}`);
    }
  }
  evidence.bundledFormulaFonts = {
    manifest: bundledFontManifestArtifact,
    count: bundledFontManifest.fonts.length,
    desktopCount: desktopFonts.length,
    officeCount: officeFonts.length,
    licenseCount: formulaFontLicenses.length,
    defaultFont: acceptance.bundledFormulaFonts.defaultFont,
    productionMathLiveRuntime: acceptance.bundledFormulaFonts.productionMathLiveRuntime,
  };

  const p15ReceiptPath = pathFor(acceptance.p15Receipt.path);
  must(existsSync(p15ReceiptPath), 'P15 baseline receipt missing.');
  must(shaFile(p15ReceiptPath) === acceptance.p15Receipt.sha256, 'P15 baseline receipt SHA changed.');
  must(JSON.parse(readFileSync(p15ReceiptPath, 'utf8')).status === 'PASS', 'P15 baseline receipt is not PASS.');
  must(acceptance.postP15RegressionFixes?.sourceChangedAfterP15 === true, 'Current P16 candidate must disclose post-P15 regression fixes.');
  must(acceptance.postP15RegressionFixes?.mathLiveRuntimeReintroduced === false, 'P16 regression fixes must not reintroduce MathLive production runtime.');
  for (const [label, record] of Object.entries({
    visualRegression: acceptance.postP15RegressionFixes?.visualRegression,
    formulaFontTruth: acceptance.postP15RegressionFixes?.formulaFontTruth,
    privateFontProductUi: acceptance.postP15RegressionFixes?.privateFontProductUi,
  })) {
    must(record?.status === 'PASS', `${label} acceptance is not PASS.`);
    const evidencePath = pathFor(record.path);
    must(existsSync(evidencePath), `${label} receipt missing.`);
    must(shaFile(evidencePath).toUpperCase() === record.sha256.toUpperCase(), `${label} receipt SHA changed.`);
  }

  steps.push(run('font-truth', node, ['scripts/p16_formula_font_truth_review.mjs'], windowsRoot, 180_000));
  steps.push(run('product-font-ui', node, ['scripts/p16_private_font_product_review.mjs'], windowsRoot, 240_000));
  steps.push(run('core-visual-regression', node, ['scripts/p16_core_visual_regression_browser_review.mjs'], windowsRoot, 180_000));
  steps.push(run('runtime-isolation', node, ['scripts/p15_retired_runtime_audit.mjs'], windowsRoot, 120_000));
  steps.push(run('fixed-office-inputs', powershell, ['-NoProfile','-ExecutionPolicy','Bypass','-File','scripts/verify_fixed_windows_office_inputs.ps1'], windowsRoot, 120_000));
  steps.push(run('full-embedded-fonts', node, ['scripts/verify_embedded_frontend_assets.mjs','--exe','src-tauri/target/release/visualtex.exe','--dist','dist'], windowsRoot, 120_000));
  steps.push(run('no-ocr-embedded-fonts', node, ['scripts/verify_embedded_frontend_assets.mjs','--exe','src-tauri/target-no-ocr/release/visualtex.exe','--dist','dist'], windowsRoot, 120_000));
  steps.push(run('full-release-static', powershell, ['-NoProfile','-ExecutionPolicy','Bypass','-File','scripts/verify_windows_release_artifacts.ps1'], windowsRoot, 300_000));
  steps.push(run('no-ocr-static', powershell, ['-NoProfile','-ExecutionPolicy','Bypass','-File','scripts/verify_windows_no_ocr_installer.ps1','-InstallerPath', pathFor(acceptance.packages.noOcr.isolatedBuiltPath)], windowsRoot, 120_000));
} catch (error) {
  failure = error instanceof Error ? error.message : String(error);
}

const frozenAfter = frozenOfficeDiff();
const frozenOfficeWorkingTreeUnchanged = JSON.stringify(frozenBefore) === JSON.stringify(frozenAfter);
if (!failure && frozenBefore.length) failure = `Protected Office source paths are dirty: ${frozenBefore.join(', ')}`;
if (!failure && !frozenOfficeWorkingTreeUnchanged) failure = 'Protected Office source snapshot changed during P16 static verification.';
const after = hashInputs();
const sourceUnchanged = JSON.stringify(before) === JSON.stringify(after);
if (!failure && !sourceUnchanged) failure = 'P16 static verifier inputs changed while verification was running.';

const branchResult = spawnSync('git', ['branch', '--show-current'], { cwd: repoRoot, encoding: 'utf8' });
const headResult = spawnSync('git', ['rev-parse', 'HEAD'], { cwd: repoRoot, encoding: 'utf8' });
const report = {
  schema: 'visualtex-p16-static-package-verification-v1',
  status: failure ? 'FAILED' : 'STATIC_PASS',
  completion: 'awaiting-installed-office-acceptance',
  branch: (branchResult.stdout || '').trim(),
  head: (headResult.stdout || '').trim(),
  sourceUnchanged,
  frozenOfficeBefore: frozenBefore,
  frozenOfficeAfter: frozenAfter,
  frozenOfficeWorkingTreeUnchanged,
  acceptance: { path: norm(acceptancePath), sha256: shaFile(acceptancePath) },
  evidence,
  steps,
  installedAcceptance: { status: 'NOT_RUN', installerInstalled: false, realInstalledOfficeUi: false },
  packageArtifactsReadyForInstallAcceptance: !failure,
  overallComplete: false,
  failure,
  verifiedAt: new Date().toISOString(),
};
const reportPath = resolve(output, 'report.json');
writeFileSync(reportPath, JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify({ status: report.status, report: norm(reportPath), receiptSha256: shaFile(reportPath), sourceUnchanged, steps: `${steps.filter(step => step.status === 'PASS').length}/${steps.length}`, frozenOfficeWorkingTreeUnchanged, packageArtifactsReadyForInstallAcceptance: report.packageArtifactsReadyForInstallAcceptance, installedAcceptance: 'NOT_RUN', overallComplete: false, failure }, null, 2));
if (failure) process.exitCode = 1;

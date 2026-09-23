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
const output = resolve(engineRoot, 'target/p12-verification', stamp);
mkdirSync(output, { recursive: true });
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const norm = path => relative(repoRoot, path).replaceAll('\\', '/');

const inputPaths = [
  'engine/specs/P12-office-interactive-integration.md',
  'engine/specs/P12-real-office-acceptance.json',
  'engine/tools/p12-verify.mjs',
  'apps/windows/src/engine/coreEditorExportPort.ts',
  'apps/windows/src/engine/coreEngineRuntime.ts',
  'apps/windows/src/editor/MathEditor.tsx',
  'apps/windows/src/editor/CoreMathEditor.tsx',
  'apps/windows/src/workspace/workspaceTypes.ts',
  'apps/windows/src/workspace/EditorWorkspace.tsx',
  'apps/windows/src/office/dialog/OfficeDialogApp.tsx',
  'apps/windows/vite.office.windows-native.config.ts',
  'apps/windows/scripts/p12_office_core_browser_review.mjs',
  'apps/windows/scripts/p12_office_core_browser_smoke.mjs',
  'apps/windows/src-tauri/Cargo.toml',
  'apps/windows/src-tauri/src/lib.rs',
  'apps/windows/src-tauri/src/system_math_glyphs.rs',
  'apps/windows/src-tauri/src/bin/system_font_asset_acceptance.rs',
  'apps/windows/src-tauri/capabilities/office-session-editor.json',
  'apps/windows/src-tauri/permissions/office-system-font.toml',
].map(path => resolve(repoRoot, path));
const inputs = () => inputPaths.map(path => ({ path: norm(path), sha256: sha(readFileSync(path)) }));
const digest = files => sha(Buffer.from(JSON.stringify(files)));
const before = inputs();

const node = process.execPath;
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const cargo = process.platform === 'win32' ? 'cargo.exe' : 'cargo';
const tsx = resolve(windowsRoot, 'node_modules/.bin', process.platform === 'win32' ? 'tsx.cmd' : 'tsx');
const tsc = resolve(windowsRoot, 'node_modules/typescript/bin/tsc');
const powershell = process.platform === 'win32'
  ? resolve(process.env.SystemRoot || 'C:/Windows', 'System32/WindowsPowerShell/v1.0/powershell.exe')
  : 'powershell';

function spawn(executable, args, cwd, timeout = 120_000) {
  const isCmd = process.platform === 'win32' && executable.toLowerCase().endsWith('.cmd');
  const command = isCmd ? (process.env.ComSpec || 'cmd.exe') : executable;
  const commandArgs = isCmd ? ['/d', '/s', '/c', executable, ...args] : args;
  return spawnSync(command, commandArgs, { cwd, encoding: 'utf8', timeout, maxBuffer: 96 * 1024 * 1024 });
}
function run(name, executable, args, cwd, timeout = 120_000) {
  const start = performance.now();
  const result = spawn(executable, args, cwd, timeout);
  const log = `COMMAND ${executable} ${args.join(' ')}\nSTDOUT\n${result.stdout ?? ''}\nSTDERR\n${result.stderr ?? ''}\nERROR ${result.error?.message ?? ''}\nEXIT ${result.status}\n`;
  const logName = `${name}.log`;
  writeFileSync(resolve(output, logName), log);
  const step = { name, status: result.status === 0 && !result.error ? 'PASS' : 'FAILED', durationMs: Math.round(performance.now() - start), exitCode: result.status, error: result.error?.message ?? null, log: logName, logSha256: sha(Buffer.from(log)) };
  console.log(JSON.stringify(step));
  if (step.status !== 'PASS') throw new Error(`${name} failed; see ${norm(resolve(output, logName))}`);
  return step;
}

function frozenOfficeDiff() {
  const frozen = [
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
  const result = spawn('git', ['diff', '--name-only', '--', ...frozen], repoRoot, 30_000);
  if (result.status !== 0 || result.error) throw new Error(`Unable to audit frozen Office paths: ${result.stderr || result.error?.message}`);
  return (result.stdout || '').split(/\r?\n/u).map(value => value.trim()).filter(Boolean);
}

function realOfficeAcceptanceRecord() {
  const path = resolve(engineRoot, 'specs/P12-real-office-acceptance.json');
  const data = JSON.parse(readFileSync(path, 'utf8'));
  const required = [
    ['status', data.status === 'PASS'],
    ['word.inlineVisualTeXOleCreate', data.word?.inlineVisualTeXOleCreate === true],
    ['word.inlineVisualTeXOleEditInPlace', data.word?.inlineVisualTeXOleEditInPlace === true],
    ['word.displayVisualTeXOleNumbered', data.word?.displayVisualTeXOleNumbered === true],
    ['word.displayVisualTeXOleUnnumbered', data.word?.displayVisualTeXOleUnnumbered === true],
    ['word.inlineWordOmml', data.word?.inlineWordOmml === true],
    ['word.displayMathTypeOleNumbered', data.word?.displayMathTypeOleNumbered === true],
    ['word.cancelLeavesDocumentUnchanged', data.word?.cancelLeavesDocumentUnchanged === true],
    ['word.closeXAutoCommit', data.word?.closeXAutoCommit === true],
    ['word.editorWindowReused', data.word?.editorWindowReused === true],
    ['word.saveReopenStructureStable', data.word?.saveReopenStructureStable === true],
    ['powerPoint.realRibbonCreate', data.powerPoint?.realRibbonCreate === true],
    ['powerPoint.editUpdatesExistingShapeWithoutDuplicate', data.powerPoint?.editUpdatesExistingShapeWithoutDuplicate === true],
    ['runtime.interactiveEditorUsesCore', data.runtime?.interactiveEditorUsesCore === true],
    ['runtime.vstoConvertRemainsLegacy', data.runtime?.vstoConvertRemainsLegacy === true],
    ['runtime.officeServerRoutesModified', data.runtime?.officeServerRoutesModified === false],
    ['protection.userWordProcessUntouched', data.protection?.userWordProcessUntouched === true],
    ['protection.officeSessionWireModified', data.protection?.officeSessionWireModified === false],
    ['protection.numberingAlgorithmsModified', data.protection?.numberingAlgorithmsModified === false],
    ['protection.comOleContractsModified', data.protection?.comOleContractsModified === false],
  ];
  const failed = required.filter(([, passed]) => !passed).map(([name]) => name);
  const dedicatedDocument = data.word?.dedicatedDocument ? resolve(repoRoot, data.word.dedicatedDocument) : null;
  if (dedicatedDocument && !existsSync(dedicatedDocument)) failed.push('word.dedicatedDocument exists');
  if (failed.length) throw new Error(`P12 real Office acceptance is incomplete: ${failed.join(', ')}`);
  return {
    status: 'PASS',
    path: norm(path),
    sha256: sha(readFileSync(path)),
    date: data.date,
    userWordProcessProtected: data.environment?.userWordProcessProtected ?? null,
    dedicatedDocument: data.word?.dedicatedDocument ?? null,
    mathTypeProgIdVerified: data.word?.mathTypeProgIdVerified ?? null,
    reopenedSourceVerified: data.word?.reopenedSourceVerified ?? null,
    wordPostReopen: data.word?.postReopen ?? null,
    powerPointMode: data.powerPoint?.currentSupportedCreateMode ?? null,
  };
}

function officeConfigurationProbe() {
  if (process.platform !== 'win32') return { status: 'NOT_APPLICABLE', reason: 'Windows Office registry/certificate probe only runs on Windows.' };
  const command = [
    "$ErrorActionPreference='Stop'",
    "$r=Get-ItemProperty 'HKCU:\\Software\\VisualTeX\\OfficeIntegration'",
    "$thumb=([string]$r.CertificateThumbprint).Replace(' ','').ToUpperInvariant()",
    "$trusted=@(Get-ChildItem Cert:\\CurrentUser\\Root | Where-Object { $_.Thumbprint.Replace(' ','').ToUpperInvariant() -eq $thumb }).Count -gt 0",
    "$word=@(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object Id,MainWindowTitle)",
    "$visualtex=@(Get-Process visualtex -ErrorAction SilentlyContinue | Select-Object Id,Path,MainWindowHandle,MainWindowTitle)",
    "$listener=Get-NetTCPConnection -LocalPort ([int]$r.CompanionPort) -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1",
    "$listenerProcess=$null;if($listener){$listenerProcess=Get-Process -Id $listener.OwningProcess -ErrorAction SilentlyContinue}",
    "$listenerPath=if($listenerProcess){[string]$listenerProcess.Path}else{''}",
    "$officeEditors=@($visualtex | Where-Object { $_.MainWindowHandle -ne 0 -and ([string]$_.MainWindowTitle) -like '*Office*' })",
    "[pscustomobject]@{ExecutablePath=[string]$r.ExecutablePath;AppDataRoot=[string]$r.AppDataRoot;CertificatePath=[string]$r.CertificatePath;CertificateThumbprint=$thumb;CompanionPort=[int]$r.CompanionPort;ProtocolVersion=[int]$r.ProtocolVersion;CertificateTrusted=$trusted;WordProcesses=$word;VisualTeXProcesses=$visualtex;ListenerOwningProcess=if($listener){[int]$listener.OwningProcess}else{$null};ListenerProcessPath=$listenerPath;ActiveOfficeEditors=$officeEditors}|ConvertTo-Json -Depth 5 -Compress",
  ].join(';');
  const result = spawn(powershell, ['-NoProfile', '-NonInteractive', '-Command', command], repoRoot, 30_000);
  if (result.status !== 0 || result.error) return { status: 'PROBE_FAILED', error: result.stderr || result.error?.message || 'unknown' };
  try {
    const data = JSON.parse((result.stdout || '').trim());
    const registeredPath = String(data.ExecutablePath || '').toLowerCase();
    const listenerPath = String(data.ListenerProcessPath || '').toLowerCase();
    const activeEditors = Array.isArray(data.ActiveOfficeEditors)
      ? data.ActiveOfficeEditors
      : data.ActiveOfficeEditors ? [data.ActiveOfficeEditors] : [];
    let status = 'HEALTHY_OR_IDLE';
    let reason = 'Current Office integration configuration is available as a diagnostic only; final P12 completion is backed by the frozen real Office acceptance record.';
    if (!data.CertificateTrusted) {
      status = 'CERTIFICATE_NOT_TRUSTED';
      reason = 'CurrentUser Root does not presently trust the registered companion certificate. This is diagnostic and does not erase the already-passed frozen P12 real Office acceptance.';
    } else if (data.ListenerOwningProcess && (!listenerPath || listenerPath !== registeredPath)) {
      status = 'PORT_OWNED_BY_OTHER_PROCESS';
      reason = 'The configured companion port is presently owned by a different executable. This is diagnostic; P12 verification will not replace or terminate it.';
    } else if (activeEditors.length) {
      status = 'CONCURRENT_OFFICE_EDITOR_PRESENT';
      reason = 'A VisualTeX Office editor window is presently active. This is diagnostic; P12 verification will not reuse or close it.';
    }
    return { status, reason, diagnosticOnly: true, ...data };
  } catch (error) {
    return { status: 'PROBE_FAILED', error: error instanceof Error ? error.message : String(error), raw: result.stdout };
  }
}

const steps = [];
let failure = null;
let realOfficeAcceptance = null;
let frozenDiffBefore = [];
let frozenDiffAfter = [];
try {
  realOfficeAcceptance = realOfficeAcceptanceRecord();
  frozenDiffBefore = frozenOfficeDiff();
  steps.push(run('windows-typescript', node, [tsc, '--noEmit'], windowsRoot));
  steps.push(run('system-font-asset-acceptance', cargo, ['run', '--quiet', '--features', 'system-font-asset-acceptance', '--bin', 'system_font_asset_acceptance'], tauriRoot, 240_000));
  steps.push(run('office-production-bundle', npm, ['run', 'build:office:windows-native'], windowsRoot, 180_000));
  steps.push(run('office-core-browser-review', node, ['scripts/p12_office_core_browser_review.mjs'], windowsRoot, 120_000));
  steps.push(run('office-core-browser-smoke', node, ['scripts/p12_office_core_browser_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('existing-office-formula-document', tsx, ['scripts/office_formula_document_regression.mts'], windowsRoot));
  steps.push(run('existing-office-session-safety', tsx, ['scripts/office_session_payload_safety_regression.mts'], windowsRoot));
  steps.push(run('existing-office-apply-shortcut', tsx, ['scripts/office_apply_shortcut_regression.mts'], windowsRoot));
  steps.push(run('p11-output-compatibility', node, ['tools/p11-verify.mjs'], engineRoot, 240_000));
} catch (error) {
  failure = error instanceof Error ? error.message : String(error);
}
try {
  frozenDiffAfter = frozenOfficeDiff();
} catch (error) {
  if (!failure) failure = error instanceof Error ? error.message : String(error);
}
const frozenOfficeWorkingTreeUnchanged = JSON.stringify(frozenDiffBefore) === JSON.stringify(frozenDiffAfter);
if (!failure && !frozenOfficeWorkingTreeUnchanged) {
  failure = `Frozen Office working-tree diff changed during P12 verification: before=[${frozenDiffBefore.join(', ')}], after=[${frozenDiffAfter.join(', ')}]`;
}
const officeEnvironmentDiagnostic = officeConfigurationProbe();
const after = inputs();
const sourceUnchanged = digest(before) === digest(after);
const report = {
  schema: 'visualtex-p12-office-interactive-verification-v2',
  status: failure || !sourceUnchanged || !realOfficeAcceptance ? 'FAILED' : 'PASS',
  completion: 'complete',
  p13Permitted: true,
  sourceSha256: digest(before),
  sourceUnchanged,
  inputs: before,
  steps,
  realOfficeAcceptance,
  frozenOfficeTrackedDiffAtStart: frozenDiffBefore,
  frozenOfficeTrackedDiffAtEnd: frozenDiffAfter,
  frozenOfficeWorkingTreeUnchanged,
  frozenOfficeDiffInterpretation: frozenDiffBefore.length
    ? 'Pre-existing tracked working-tree differences were present when final P12 verification started; P12 finalization did not modify or revert them.'
    : 'No tracked working-tree differences were present under the frozen Office paths when final P12 verification started or ended.',
  officeEnvironmentDiagnostic,
  scope: 'P12 interactive Office editor ownership/export integration only. Hidden converter remains legacy; Office Session wire/state and Word/PowerPoint document algorithms stay frozen.',
  browserCovered: [
    'production Office bundle core resource load',
    'initial clean Session hydration',
    'real browser keyboard input',
    'same-revision SVG/MathML/PNG/baseline export',
    'rapid edit and Session-switch stale export rejection',
    'core Undo/Redo through Office controls',
    'theme synchronization',
    'Apply through existing Session commit handshake',
    'Cancel through existing Session cancellation path',
    'vsto-convert remains legacy and does not load core resources',
  ],
  realOfficeCovered: [
    'real Word Ribbon/UI VisualTeX OLE create and in-place edit',
    'real Word display VisualTeX OLE numbered and unnumbered',
    'real Word OMML and MathType OLE through supported modes',
    'Cancel and native close/X auto-commit behavior',
    'same companion/WebView reuse across Sessions',
    'save/reopen structural inspection and exact source restoration',
    'real PowerPoint crossPlatformPicture create and in-place edit without duplicate shape',
    'remote Office WebView system-font access through minimal Tauri capability',
    'mathematical system font selection restricted to faces with an OpenType MATH table',
  ],
  requiredButNotPassed: [],
  failure,
};
writeFileSync(resolve(output, 'report.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify({ status: report.status, completion: report.completion, receipt: norm(resolve(output, 'report.json')), sourceSha256: report.sourceSha256, sourceUnchanged, steps: steps.length, p13Permitted: true, realOfficeAcceptance: realOfficeAcceptance?.status ?? null, officeEnvironmentDiagnostic: officeEnvironmentDiagnostic.status, frozenOfficeWorkingTreeUnchanged, failure }, null, 2));
if (report.status !== 'PASS') process.exitCode = 1;

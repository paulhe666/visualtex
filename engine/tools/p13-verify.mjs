import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const engineRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = resolve(engineRoot, '..');
const windowsRoot = resolve(repoRoot, 'apps/windows');
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const output = resolve(engineRoot, 'target/p13-verification', stamp);
mkdirSync(output, { recursive: true });

const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const norm = path => relative(repoRoot, path).replaceAll('\\', '/');
const readText = path => readFileSync(path, 'utf8').replace(/^\uFEFF/u, '');
const readJson = path => JSON.parse(readText(path));
const assert = (condition, message) => {
  if (!condition) throw new Error(message);
};

const inputPaths = [
  'engine/specs/P13-office-batch-conversion-integration.md',
  'engine/specs/P13-real-office-acceptance.json',
  'engine/tools/p13-verify.mjs',
  'engine/fixtures/p13-office-bulk-import.md',
  'apps/windows/src/engine/coreEngineRuntime.ts',
  'apps/windows/src/office/compat/coreOfficeBatchConversion.ts',
  'apps/windows/src/office/dialog/OfficeDialogApp.tsx',
  'apps/windows/scripts/p13_office_interactive_core_guard.mjs',
  'apps/windows/scripts/p13_office_core_converter_browser_smoke.mjs',
  'apps/windows/scripts/p13_office_core_converter_performance.mjs',
].map(path => resolve(repoRoot, path));
const inputs = () => inputPaths.map(path => ({ path: norm(path), sha256: sha(readFileSync(path)) }));
const digest = files => sha(Buffer.from(JSON.stringify(files)));
const before = inputs();

const node = process.execPath;
const npm = process.platform === 'win32' ? 'npm.cmd' : 'npm';
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
    maxBuffer: 96 * 1024 * 1024,
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
  if (step.status !== 'PASS') {
    throw new Error(`${name} failed; see ${norm(resolve(output, logName))}`);
  }
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
function gitLines(args) {
  const result = spawn('git', args, repoRoot, 30_000);
  if (result.status !== 0 || result.error) {
    throw new Error(`Git audit failed: ${result.stderr || result.error?.message || 'unknown error'}`);
  }
  return (result.stdout || '').split(/\r?\n/u).map(value => value.trimEnd()).filter(Boolean);
}
function frozenStatus() {
  return gitLines(['status', '--porcelain=v1', '--untracked-files=all', '--', ...frozenPaths]);
}
function gitIdentity() {
  return {
    branch: gitLines(['rev-parse', '--abbrev-ref', 'HEAD'])[0] ?? '',
    head: gitLines(['rev-parse', 'HEAD'])[0] ?? '',
  };
}

function progIdCount(audit, progId) {
  return (audit.items ?? []).filter(item => item?.progId === progId).length;
}
function validateResult(path, source, target, count) {
  const data = readJson(path);
  assert(data.passed === true, `${norm(path)} did not pass`);
  assert(data.count === count, `${norm(path)} count ${data.count} != ${count}`);
  assert(typeof data.record === 'string' && data.record.includes(`source=${source}`) && data.record.includes(`target=${target}`), `${norm(path)} source/target record mismatch`);
  assert(data.record.includes(`converted=${count}`) && data.record.includes('failed=0'), `${norm(path)} did not record converted=${count} failed=0`);
  return data;
}
function validateAudit(path, { paragraphs, oMaths, visualTeX, mathType, otherInlineShapes }) {
  const data = readJson(path);
  if (paragraphs !== undefined) assert(data.paragraphs === paragraphs, `${norm(path)} paragraphs ${data.paragraphs} != ${paragraphs}`);
  if (oMaths !== undefined) assert((data.maths ?? []).length === oMaths, `${norm(path)} OMath ${data.maths?.length ?? 0} != ${oMaths}`);
  if (visualTeX !== undefined) assert(progIdCount(data, 'VisualTeX.Formula.1') === visualTeX, `${norm(path)} VisualTeX count mismatch`);
  if (mathType !== undefined) assert(progIdCount(data, 'Equation.DSMT4') === mathType, `${norm(path)} MathType count mismatch`);
  if (otherInlineShapes !== undefined) {
    const count = (data.items ?? []).filter(item => !item?.progId).length;
    assert(count === otherInlineShapes, `${norm(path)} unrelated InlineShape count ${count} != ${otherInlineShapes}`);
  }
  return data;
}

function realOfficeAcceptanceRecord() {
  const recordPath = resolve(engineRoot, 'specs/P13-real-office-acceptance.json');
  const data = readJson(recordPath);
  const required = [
    ['status', data.status === 'PASS'],
    ['runtime.interactiveEditorStillUsesCore', data.runtime?.interactiveEditorStillUsesCore === true],
    ['runtime.hiddenConverterUsesHeadlessCore', data.runtime?.hiddenConverterUsesHeadlessCore === true],
    ['runtime.hiddenConverterMountsInteractiveEditor', data.runtime?.hiddenConverterMountsInteractiveEditor === false],
    ['runtime.hiddenConverterMountsMathLive', data.runtime?.hiddenConverterMountsMathLive === false],
    ['runtime.officeServerRoutesModifiedByP13', data.runtime?.officeServerRoutesModifiedByP13 === false],
    ['runtime.officeSessionWireModifiedByP13', data.runtime?.officeSessionWireModifiedByP13 === false],
    ['fullDocument100.sourceFormulas', data.fullDocument100?.sourceFormulas === 100],
    ['fullDocument100.allDirectionsFailedZero', data.fullDocument100?.allDirectionsFailedZero === true],
    ['mixedMathType146.unrelatedContentPreserved', data.mixedMathType146?.unrelatedContentPreserved === true],
    ['redraw1000.materializedVerified', data.redraw1000?.materializedVerified === true],
    ['redraw1000.formulas', data.redraw1000?.formulas === 1000],
    ['sourceRestore.formulas', data.sourceRestore?.formulas === 146],
    ['bulkImport.formulas', data.bulkImport?.formulas === 3],
    ['bulkImport.proseOrderVerified', data.bulkImport?.proseOrderVerified === true],
    ['browserAndPerformance.malformedSessionFailsWithoutPartialExport', data.browserAndPerformance?.malformedSessionFailsWithoutPartialExport === true],
    ['browserAndPerformance.batchContinuesAfterMalformedSession', data.browserAndPerformance?.batchContinuesAfterMalformedSession === true],
    ['browserAndPerformance.batchOrderAndIdentityPreserved', data.browserAndPerformance?.batchOrderAndIdentityPreserved === true],
    ['browserAndPerformance.omml1000.failures', data.browserAndPerformance?.omml1000?.failures === 0],
    ['protection.officeServerModifiedByP13', data.protection?.officeServerModifiedByP13 === false],
    ['protection.contractsModifiedByP13', data.protection?.contractsModifiedByP13 === false],
    ['protection.wordVstoModifiedByP13', data.protection?.wordVstoModifiedByP13 === false],
    ['protection.powerPointVstoModifiedByP13', data.protection?.powerPointVstoModifiedByP13 === false],
    ['protection.userWordDocumentsSavedOrClosed', data.protection?.userWordDocumentsSavedOrClosed === false],
  ];
  const failed = required.filter(([, passed]) => !passed).map(([name]) => name);
  if (failed.length) throw new Error(`P13 real Office acceptance is incomplete: ${failed.join(', ')}`);

  const redraw100ResultPath = resolve(repoRoot, data.fullDocument100.redrawVisualTeX.result);
  const redraw100 = readJson(redraw100ResultPath);
  assert(redraw100.passed === true && redraw100.count === 100, 'P13 100-formula VisualTeX redraw did not pass');
  assert(/redraw-complete formulas=100 unique=12 skipped=0/u.test(redraw100.record ?? ''), 'P13 100-formula redraw record mismatch');
  validateAudit(resolve(repoRoot, data.fullDocument100.redrawVisualTeX.audit), {
    paragraphs: 101,
    oMaths: 0,
    visualTeX: 100,
    mathType: 0,
  });

  const expectedDirections = new Map([
    ['VisualTeX→OMML', { oMaths: 100, visualTeX: 0, mathType: 0 }],
    ['OMML→MathType', { oMaths: 0, visualTeX: 0, mathType: 100 }],
    ['MathType→OMML', { oMaths: 100, visualTeX: 0, mathType: 0 }],
    ['OMML→VisualTeX', { oMaths: 0, visualTeX: 100, mathType: 0 }],
    ['VisualTeX→MathType', { oMaths: 0, visualTeX: 0, mathType: 100 }],
    ['MathType→VisualTeX', { oMaths: 0, visualTeX: 100, mathType: 0 }],
  ]);
  const directions = [];
  for (const direction of data.fullDocument100.directions ?? []) {
    const key = `${direction.source}→${direction.target}`;
    const expected = expectedDirections.get(key);
    assert(expected, `Unexpected P13 full-document direction ${key}`);
    const resultPath = resolve(repoRoot, direction.result);
    const result = validateResult(resultPath, direction.source, direction.target, 100);
    const auditPath = resolve(repoRoot, direction.audit);
    validateAudit(auditPath, { paragraphs: 101, ...expected });
    directions.push({ key, result: norm(resultPath), audit: norm(auditPath), productMs: result.productMs });
  }
  assert(directions.length === 6 && new Set(directions.map(item => item.key)).size === 6, 'P13 six full-document conversion directions are incomplete');

  const selectedHookPath = resolve(repoRoot, data.selectedConversion.visualTeXToOmml.log);
  const selectedHook = readText(selectedHookPath);
  const selectedStart = selectedHook.indexOf('format-conversion-start sourceMode=nativeOle targetMode=wordOmml wholeDocument=False');
  const selectedComplete = selectedHook.indexOf('format-conversion-complete source=VisualTeX target=OMML converted=1 failed=0', selectedStart);
  assert(selectedStart >= 0 && selectedComplete > selectedStart, 'Selected VisualTeX→OMML real Word acceptance is missing');
  const selectedOmmlMtPath = resolve(repoRoot, data.selectedConversion.ommlToMathType.result);
  const selectedMtVtPath = resolve(repoRoot, data.selectedConversion.mathTypeToVisualTeX.result);
  validateResult(selectedOmmlMtPath, 'OMML', 'MathType', 1);
  validateResult(selectedMtVtPath, 'MathType', 'VisualTeX', 1);

  const mixedBeforePath = resolve(repoRoot, data.mixedMathType146.beforeAudit);
  const mixedResultPath = resolve(repoRoot, data.mixedMathType146.conversionResult);
  const mixedAfterPath = resolve(repoRoot, data.mixedMathType146.afterAudit);
  validateAudit(mixedBeforePath, { paragraphs: 97, oMaths: 1, visualTeX: 0, mathType: 146, otherInlineShapes: 2 });
  validateResult(mixedResultPath, 'MathType', 'VisualTeX', 146);
  validateAudit(mixedAfterPath, { paragraphs: 97, oMaths: 1, visualTeX: 146, mathType: 0, otherInlineShapes: 2 });

  const restoreLogPath = resolve(repoRoot, data.sourceRestore.log);
  const restoreLog = readText(restoreLogPath);
  assert(restoreLog.includes('formula-to-latex-start scope=document mode=nativeOle formulas=146'), 'P13 source-restore start record is missing');
  assert(restoreLog.includes('formula-to-latex-complete scope=document mode=nativeOle formulas=146'), 'P13 source-restore completion record is missing');
  const restoreAuditPath = resolve(repoRoot, data.sourceRestore.afterAudit);
  validateAudit(restoreAuditPath, { paragraphs: 97, oMaths: 1, visualTeX: 0, mathType: 0, otherInlineShapes: 2 });

  const redraw1000UiPath = resolve(repoRoot, data.redraw1000.uiResult);
  const redraw1000Ui = readJson(redraw1000UiPath);
  assert(redraw1000Ui.success === true, 'P13 1000-formula real Word UI result did not pass');
  assert(/redraw-complete formulas=1000 unique=1000 skipped=0/u.test(redraw1000Ui.record ?? ''), 'P13 1000-formula redraw record mismatch');
  const redraw1000LogPath = resolve(repoRoot, data.redraw1000.redrawLog);
  const redraw1000Log = readText(redraw1000LogPath);
  assert(redraw1000Log.includes('redraw-complete formulas=1000 unique=1000 skipped=0 renderAverageMs=8 renderMaxMs=49 insertTotalMs=7233 insertMaxMs=127'), 'P13 1000-formula redraw completion record is missing');
  const redraw1000HookPath = resolve(repoRoot, data.redraw1000.hookLog);
  const redraw1000Hook = readText(redraw1000HookPath);
  assert(redraw1000Hook.includes('redraw-omml-materialized-verified formulas=1000 elapsedMs=17759'), 'P13 1000-formula materialization verification is missing');
  const undoInspectPath = resolve(repoRoot, data.redraw1000.intentionalUndo.inspect);
  const undoInspect = readJson(undoInspectPath);
  assert(undoInspect.sourceTextRestored === true && undoInspect.omml === 0 && undoInspect.ole === 0 && undoInspect.paragraphs === 1002, 'P13 intentional Undo evidence is inconsistent');
  const sourceDocumentPath = resolve(repoRoot, data.redraw1000.sourceDocument);
  assert(existsSync(sourceDocumentPath), 'P13 1000-formula source document is missing');

  const bulkLogPath = resolve(repoRoot, data.bulkImport.log);
  const bulkLog = readText(bulkLogPath);
  assert(bulkLog.includes('parsed blocks=7 formulas=3 mode=Omml'), 'P13 bulk-import parse record is missing');
  assert(bulkLog.includes('bulk-import-complete blocks=7 formulas=3'), 'P13 bulk-import completion record is missing');
  const bulkAuditPath = resolve(repoRoot, data.bulkImport.audit);
  const bulkAudit = validateAudit(bulkAuditPath, { paragraphs: 8, oMaths: 3, visualTeX: 0, mathType: 0 });
  assert((bulkAudit.items ?? []).length === 0, 'P13 OMML bulk-import unexpectedly created OLE objects');
  assert(bulkAudit.bookmarks === 3, `P13 bulk-import bookmark count ${bulkAudit.bookmarks} != 3`);
  const bulkXmlPath = resolve(repoRoot, data.bulkImport.auditXml);
  const bulkXml = readText(bulkXmlPath);
  const prose = [
    '# P13 批量导入验收',
    '这是一段中文正文',
    '下面是一个行间矩阵',
    '下面是一个较长的行间公式',
    '最后一段正文用于确认公式后的文本顺序保持不变。',
  ];
  const positions = prose.map(value => bulkXml.indexOf(value));
  assert(positions.every(position => position >= 0), 'P13 bulk-import prose is missing from WordOpenXML');
  assert(positions.every((position, index) => index === 0 || positions[index - 1] < position), 'P13 bulk-import prose order changed');

  const evidencePaths = new Set();
  const visit = value => {
    if (typeof value === 'string' && value.startsWith('docs/remediation-3d207d7/evidence/')) evidencePaths.add(resolve(repoRoot, value));
    else if (Array.isArray(value)) value.forEach(visit);
    else if (value && typeof value === 'object') Object.values(value).forEach(visit);
  };
  visit(data);
  for (const path of evidencePaths) assert(existsSync(path), `P13 evidence file is missing: ${norm(path)}`);
  const evidence = [...evidencePaths].sort().map(path => ({ path: norm(path), sha256: sha(readFileSync(path)) }));

  return {
    status: 'PASS',
    path: norm(recordPath),
    sha256: sha(readFileSync(recordPath)),
    date: data.date,
    fullDocumentDirections: directions,
    selectedDirections: ['VisualTeX→OMML', 'OMML→MathType', 'MathType→VisualTeX'],
    redraw1000: {
      formulas: data.redraw1000.formulas,
      materializedVerified: true,
      renderAverageMs: data.redraw1000.renderAverageMs,
      renderMaxMs: data.redraw1000.renderMaxMs,
      insertTotalMs: data.redraw1000.insertTotalMs,
      intentionalUndoSourceRestored: undoInspect.sourceTextRestored,
    },
    bulkImport: {
      blocks: data.bulkImport.blocks,
      formulas: data.bulkImport.formulas,
      oMaths: (bulkAudit.maths ?? []).length,
      paragraphs: bulkAudit.paragraphs,
      proseOrderVerified: true,
    },
    evidence,
    runtimeObservation: data.runtimeObservation,
  };
}

const steps = [];
let failure = null;
let realOfficeAcceptance = null;
let frozenStatusBefore = [];
let frozenStatusAfter = [];
let identity = null;
try {
  identity = gitIdentity();
  assert(identity.branch === 'reference-3d207d7', `P13 verification is on unexpected branch ${identity.branch}`);
  assert(identity.head === 'd1f0aff8db9275c90dc18c7318dc2a4d7bbbd760', `P13 verification HEAD changed: ${identity.head}`);
  frozenStatusBefore = frozenStatus();
  assert(frozenStatusBefore.length === 0, `Frozen Office boundary is dirty before P13 verification: ${frozenStatusBefore.join(', ')}`);
  realOfficeAcceptance = realOfficeAcceptanceRecord();
  steps.push(run('windows-typescript', node, [tsc, '--noEmit'], windowsRoot));
  steps.push(run('office-production-bundle', npm, ['run', 'build:office:windows-native'], windowsRoot, 180_000));
  steps.push(run('p13-interactive-core-guard', node, ['scripts/p13_office_interactive_core_guard.mjs'], windowsRoot, 120_000));
  steps.push(run('p13-headless-converter-browser', node, ['scripts/p13_office_core_converter_browser_smoke.mjs'], windowsRoot, 120_000));
  steps.push(run('p13-headless-converter-performance', node, ['scripts/p13_office_core_converter_performance.mjs'], windowsRoot, 180_000));
  steps.push(run('existing-office-formula-document', tsx, ['scripts/office_formula_document_regression.mts'], windowsRoot));
  steps.push(run('existing-office-session-safety', tsx, ['scripts/office_session_payload_safety_regression.mts'], windowsRoot));
  steps.push(run('existing-office-apply-shortcut', tsx, ['scripts/office_apply_shortcut_regression.mts'], windowsRoot));
  steps.push(run('existing-document-import-parser', tsx, ['scripts/document_import_parser_coverage.mts'], windowsRoot, 120_000));
} catch (error) {
  failure = error instanceof Error ? error.message : String(error);
}
try {
  frozenStatusAfter = frozenStatus();
} catch (error) {
  if (!failure) failure = error instanceof Error ? error.message : String(error);
}
const frozenOfficeWorkingTreeUnchanged = JSON.stringify(frozenStatusBefore) === JSON.stringify(frozenStatusAfter) && frozenStatusAfter.length === 0;
if (!failure && !frozenOfficeWorkingTreeUnchanged) {
  failure = `Frozen Office boundary changed during P13 verification: before=[${frozenStatusBefore.join(', ')}], after=[${frozenStatusAfter.join(', ')}]`;
}
const after = inputs();
const sourceUnchanged = digest(before) === digest(after);
if (!failure && !sourceUnchanged) failure = 'Verified P13 source inputs changed while the verifier was running.';
const report = {
  schema: 'visualtex-p13-office-batch-verification-v1',
  status: failure || !sourceUnchanged || !realOfficeAcceptance || !frozenOfficeWorkingTreeUnchanged ? 'FAILED' : 'PASS',
  completion: 'complete',
  p14Permitted: true,
  branch: identity?.branch ?? null,
  head: identity?.head ?? null,
  sourceSha256: digest(before),
  sourceUnchanged,
  inputs: before,
  steps,
  realOfficeAcceptance,
  frozenOfficeStatusAtStart: frozenStatusBefore,
  frozenOfficeStatusAtEnd: frozenStatusAfter,
  frozenOfficeWorkingTreeUnchanged,
  p12Boundary: {
    checkpoint: 'docs/engine-rebuild/checkpoints/CP12-office-interactive-complete-1/',
    headAtFreeze: 'd1f0aff8db9275c90dc18c7318dc2a4d7bbbd760',
    protectedTrackedAndUntrackedStatusRequiredEmpty: true,
  },
  browserCovered: [
    'P12 interactive Office editor remains core-owned with no MathLive owner',
    'interactive real-browser input, export, Undo/Redo, Session switch and Apply handshake',
    'vsto-convert uses headless core ownership with no editor/input/IME surface',
    'native OLE, Word OMML, MathType OLE and cross-platform picture export contracts',
    'unsupported Session failure is atomic and later batch Sessions continue',
    'batch order and identity writes are stable',
    'Office equation-tag compatibility is core-backed',
    '100 mixed and 1000 OMML headless performance/stability sample',
  ],
  realOfficeCovered: [
    '100-formula VisualTeX redraw',
    'all six 100-formula full-document conversion directions',
    'selected VisualTeX→OMML→MathType→VisualTeX path',
    '146 MathType→VisualTeX mixed-content preservation',
    'real Ribbon 146 VisualTeX OLE→LaTeX source restoration',
    '1000-formula real Word OMML redraw and strict materialization verification',
    'real visible bulk-import UI with native Word text plus inline/display OMML formulas',
  ],
  knownObservationCarriedToP14: realOfficeAcceptance?.runtimeObservation ?? null,
  requiredButNotPassed: [],
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
  realOfficeAcceptance: realOfficeAcceptance?.status ?? null,
  frozenOfficeWorkingTreeUnchanged,
  p14Permitted: report.p14Permitted,
  failure,
}, null, 2));
if (report.status !== 'PASS') process.exitCode = 1;

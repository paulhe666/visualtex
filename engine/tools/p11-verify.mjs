import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, relative, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const engineRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = resolve(engineRoot, '..');
const windowsRoot = resolve(repoRoot, 'apps/windows');
const stamp = new Date().toISOString().replaceAll(':', '-').replaceAll('.', '-');
const output = resolve(engineRoot, 'target/p11-verification', stamp);
mkdirSync(output, { recursive: true });
const sha = (bytes) => createHash('sha256').update(bytes).digest('hex');
const norm = (path) => relative(repoRoot, path).replaceAll('\\', '/');

const inputPaths = [
  'engine/crates/core/src/lib.rs',
  'engine/crates/core/src/protocol.rs',
  'engine/crates/core/src/presentation_mathml.rs',
  'engine/crates/core/tests/presentation_mathml.rs',
  'engine/web/host/protocol.ts',
  'engine/web/host/editor-session.ts',
  'engine/tools/p11-mathml-parity-review.mjs',
  'engine/tools/p11-verify.mjs',
  'engine/specs/P11-office-compatibility.md',
  'apps/windows/src/office/compat/coreOfficeCompatibility.ts',
  'apps/windows/scripts/p11_core_office_compat_review.mts',
  'apps/windows/scripts/p11_emit_office_consumer_fixture.mts',
  'apps/windows/scripts/p11_office_consumer_acceptance.ps1',
].map((path) => resolve(repoRoot, path));
const inputs = () => inputPaths.map((path) => ({ path: norm(path), sha256: sha(readFileSync(path)) }));
const digest = (files) => sha(Buffer.from(JSON.stringify(files)));
const before = inputs();

const node = process.execPath;
const cargo = process.platform === 'win32' ? 'cargo.exe' : 'cargo';
const dotnet = process.platform === 'win32' ? 'dotnet.exe' : 'dotnet';
const tsx = resolve(windowsRoot, 'node_modules/.bin', process.platform === 'win32' ? 'tsx.cmd' : 'tsx');
const tsc = resolve(windowsRoot, 'node_modules/typescript/bin/tsc');
const powershell = process.platform === 'win32'
  ? resolve(process.env.SystemRoot || 'C:/Windows', 'System32/WindowsPowerShell/v1.0/powershell.exe')
  : 'powershell';

function run(name, executable, args, cwd, timeout = 120_000) {
  const isCmd = process.platform === 'win32' && executable.toLowerCase().endsWith('.cmd');
  const command = isCmd ? (process.env.ComSpec || 'cmd.exe') : executable;
  const commandArgs = isCmd ? ['/d', '/s', '/c', executable, ...args] : args;
  const start = performance.now();
  const result = spawnSync(command, commandArgs, { cwd, encoding: 'utf8', timeout, maxBuffer: 64 * 1024 * 1024 });
  const log = `COMMAND ${executable} ${args.join(' ')}\nSTDOUT\n${result.stdout ?? ''}\nSTDERR\n${result.stderr ?? ''}\nERROR ${result.error?.message ?? ''}\nEXIT ${result.status}\n`;
  const logName = `${name}.log`;
  writeFileSync(resolve(output, logName), log);
  const step = { name, status: result.status === 0 && !result.error ? 'PASS' : 'FAILED', durationMs: Math.round(performance.now() - start), exitCode: result.status, error: result.error?.message ?? null, log: logName, logSha256: sha(Buffer.from(log)) };
  console.log(JSON.stringify(step));
  if (step.status !== 'PASS') throw new Error(`${name} failed; see ${norm(resolve(output, logName))}`);
  return step;
}

const steps = [];
let failure = null;
try {
  steps.push(run('core-presentation-mathml', cargo, ['test', '--manifest-path', 'Cargo.toml', '-p', 'visualtex-engine-core', '--test', 'presentation_mathml'], engineRoot));
  steps.push(run('native-wasm-mathml-parity', node, ['tools/p11-mathml-parity-review.mjs'], engineRoot));
  steps.push(run('windows-typescript', node, [tsc, '--noEmit'], windowsRoot));
  steps.push(run('wasm-office-adapter', tsx, ['scripts/p11_core_office_compat_review.mts'], windowsRoot));
  steps.push(run('existing-mathtype-alignment', tsx, ['scripts/math_type_alignment_geometry_acceptance.mts'], windowsRoot));
  steps.push(run('existing-office-metadata', tsx, ['scripts/office_metadata_smoke.mjs'], windowsRoot));
  steps.push(run('existing-office-formula-document', tsx, ['scripts/office_formula_document_regression.mts'], windowsRoot));
  steps.push(run('existing-office-session-safety', tsx, ['scripts/office_session_payload_safety_regression.mts'], windowsRoot));
  steps.push(run('existing-office-apply-shortcut', tsx, ['scripts/office_apply_shortcut_regression.mts'], windowsRoot));
  steps.push(run('current-vsto-consumer-acceptance', powershell, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', 'scripts/p11_office_consumer_acceptance.ps1'], windowsRoot));
  steps.push(run('dotnet-mathml-mathtype-consumers', dotnet, ['test', 'src-windows/VisualTeX.WindowsOffice.Tests/VisualTeX.WindowsOffice.Tests.csproj', '-c', 'Release', '--no-restore', '--filter', 'FullyQualifiedName~MathTypeFontSizeTests|FullyQualifiedName~MathTypeFenceRoundTripTests|FullyQualifiedName~MathMlToLatexConverterTests'], windowsRoot, 180_000));
} catch (error) {
  failure = error instanceof Error ? error.message : String(error);
}
const after = inputs();
const sourceUnchanged = digest(before) === digest(after);
const report = {
  schema: 'visualtex-p11-office-compat-verification-v1',
  status: failure || !sourceUnchanged ? 'FAILED' : 'PASS',
  completion: 'P11-complete',
  sourceSha256: digest(before), sourceUnchanged, inputs: before, steps,
  scope: 'Pure core Presentation MathML and offline Office compatibility adapter. Existing Session/VSTO/Word document algorithms are consumers only and were not modified or automated.',
  officeSessionCalls: 0, wordAutomationCalls: 0, ribbonActions: 0, documentMutations: 0,
  notCovered: ['Interactive Office Session integration (P12)', 'Hidden converter/batch integration (P13)', 'Real Ribbon/UI acceptance (P12/P13/P16)'],
  failure,
};
writeFileSync(resolve(output, 'report.json'), JSON.stringify(report, null, 2));
console.log(JSON.stringify({ status: report.status, completion: report.completion, receipt: norm(resolve(output, 'report.json')), sourceSha256: report.sourceSha256, sourceUnchanged, steps: steps.length, failure }, null, 2));
if (report.status !== 'PASS') process.exitCode = 1;

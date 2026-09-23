// Read-only provenance/guard tool. Intentionally never writes, rebuilds, or resets files.
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs';
import { dirname, resolve, relative, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const git = (...args) => execFileSync('git', args, { cwd: root, encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }).trimEnd();
const sha = (data) => createHash('sha256').update(data).digest('hex');
const slash = (s) => s.replaceAll('\\', '/');
const digestFiles = (paths) => {
  const sorted = [...new Set(paths)].sort();
  const entries = sorted.map(p => [p, sha(readFileSync(resolve(root, p)))]);
  return { files: entries.length, sha256: sha(JSON.stringify(entries)) };
};
const walk = (dir) => {
  if (!existsSync(resolve(root, dir))) return [];
  return readdirSync(resolve(root, dir), { withFileTypes: true }).flatMap(e => {
    const p = slash(join(dir, e.name));
    if (e.isSymbolicLink()) throw new Error(`Symlink not allowed in frozen input: ${p}`);
    return e.isDirectory() ? walk(p) : e.isFile() ? [p] : [];
  });
};
const excluded = /\/(?:node_modules|target|bin|obj|TestResults|\.vs|build|dist|dist-office|test-results)\//i;
const known = git('ls-files', '-z', '--cached', '--others', '--exclude-standard').split('\0').filter(Boolean);
const sourceFiles = known.filter(p => p.startsWith('apps/windows/') && !excluded.test(p)
  && !p.startsWith('apps/windows/src-tauri/resources/') && existsSync(resolve(root, p)) && statSync(resolve(root, p)).isFile());
const groups = {
  windowsSource: digestFiles(sourceFiles),
  wordVsto: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src-windows/VisualTeX.WordVsto/'))),
  powerPointVsto: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src-windows/VisualTeX.PowerPointVsto/'))),
  contracts: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src-windows/VisualTeX.WindowsOffice.Contracts/'))),
  vstoShared: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/'))),
  ole: digestFiles(sourceFiles.filter(p => /\/VisualTeX\.(?:FormulaOleServer|WindowsOleBridge)\//.test(p))),
  rustOffice: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src-tauri/src/office/'))),
  tsOffice: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src/office/'))),
  export: digestFiles(sourceFiles.filter(p => p.startsWith('apps/windows/src/export/'))),
  officeResources: digestFiles(walk('apps/windows/src-tauri/resources/windows-office')),
  frontendDist: digestFiles(walk('apps/windows/dist')),
  companionDist: digestFiles(walk('apps/windows/dist-office')),
};
const artifactPaths = [
  ...walk('apps/windows/src-tauri/target/release/bundle/nsis').filter(p => p.endsWith('.exe')),
  ...walk('apps/windows/src-tauri/resources/windows-office').filter(p => /\.(?:msi|json)$/.test(p)),
  'apps/windows/src-tauri/target/release/visualtex.exe',
];
const installed = [
  `${process.env.LOCALAPPDATA}/VisualTeX/visualtex.exe`,
  `${process.env.ProgramFiles}/VisualTeX/WindowsOffice/VSTO/VisualTeX.WordVsto.dll`,
  `${process.env.ProgramFiles}/VisualTeX/WindowsOffice/VSTO/VisualTeX.FormulaOleServer.exe`,
  `${process.env.ProgramFiles}/Microsoft Office/Root/Office16/WINWORD.EXE`,
].filter(p => !p.startsWith('undefined'));
const artifacts = [...artifactPaths, ...installed].filter(p => existsSync(resolve(root, p))).map(p => {
  const abs = resolve(root, p); const s = statSync(abs);
  return { path: slash(p), bytes: s.size, mtime: s.mtime.toISOString(), sha256: sha(readFileSync(abs)) };
});
const status = git('status', '--porcelain=v1', '-z', '--untracked-files=no');
const dirty = status.split('\0').filter(Boolean).map(s => {
  const path = s.slice(3); const abs = resolve(root, path);
  return { status: s.slice(0, 2), path, sha256: existsSync(abs) && statSync(abs).isFile() ? sha(readFileSync(abs)) : null };
});
const gitCached = new Set(git('ls-files', '-z', '--cached').split('\0').filter(Boolean));
const historicalUntracked = known.filter(p => !p.startsWith('engine/') && !p.startsWith('docs/engine-rebuild/'))
  .filter(p => !gitCached.has(p)).sort();
const manifest = {
  schema: 'visualtex-engine-provenance-v1',
  capturedAt: new Date().toISOString(),
  root: slash(root), branch: git('branch', '--show-current'), head: git('rev-parse', 'HEAD'),
  indexSha256: sha(git('ls-files', '--stage', '-z')),
  trackedDirtyCount: dirty.length,
  historicalUntracked: { count: historicalUntracked.length, pathsSha256: sha(JSON.stringify(historicalUntracked)) },
  groups, artifacts, dirty,
};
if (process.argv[2] === '--check') {
  if (!process.argv[3]) throw new Error('Usage: --check <manifest.json>');
  const baseline = JSON.parse(readFileSync(resolve(root, process.argv[3]), 'utf8'));
  const failures = [];
  const transitionIndex = process.argv.indexOf('--git-transition');
  let expectedGit = baseline;
  let acknowledgedGitTransition = null;
  if (transitionIndex >= 0) {
    const transitionPath = process.argv[transitionIndex + 1];
    if (!transitionPath) throw new Error('Missing explicit Git transition manifest');
    const transition = JSON.parse(readFileSync(resolve(root, transitionPath), 'utf8'));
    if (transition.schema !== 'visualtex-engine-observed-git-transition-v1'
      || transition.fromHead !== baseline.head || transition.fromIndexSha256 !== baseline.indexSha256
      || git('rev-parse', `${transition.head}^`) !== baseline.head) {
      throw new Error('Git transition does not match the frozen baseline and direct child commit');
    }
    const promoted = git('diff', '--diff-filter=A', '--name-only', baseline.head, transition.head).split('\n').filter(Boolean);
    if (promoted.length !== transition.promotedUntrackedCount
      || promoted.some(p => p.startsWith('engine/') || p.startsWith('docs/engine-rebuild/'))
      || baseline.historicalUntracked.count - promoted.length !== transition.historicalUntracked.count) {
      throw new Error('Git transition file promotion set is inconsistent');
    }
    expectedGit = { ...baseline, head: transition.head, indexSha256: transition.indexSha256,
      historicalUntracked: transition.historicalUntracked };
    acknowledgedGitTransition = transition.head;
  }
  for (const key of ['branch', 'head', 'indexSha256']) {
    if (expectedGit[key] !== manifest[key]) failures.push(key);
  }
  for (const previous of baseline.dirty ?? []) {
    const path = resolve(root, previous.path);
    if (!existsSync(path) || sha(readFileSync(path)) !== previous.sha256) {
      failures.push(`originalWorkingBytes:${previous.path}`);
    }
  }
  for (const [key, value] of Object.entries(baseline.groups)) {
    if (JSON.stringify(value) !== JSON.stringify(groups[key])) failures.push(`groups.${key}`);
  }
  for (const b of baseline.artifacts) {
    const a = artifacts.find(a => a.path === b.path);
    if (!a || a.sha256 !== b.sha256 || a.bytes !== b.bytes) failures.push(`artifact:${b.path}`);
  }
  if (JSON.stringify(expectedGit.historicalUntracked) !== JSON.stringify(manifest.historicalUntracked)) {
    failures.push('historicalUntracked paths');
  }
  console.log(JSON.stringify({ status: failures.length ? 'FAILED' : 'PASS', checkedAt: manifest.capturedAt,
    groupCount: Object.keys(groups).length, artifactCount: baseline.artifacts.length,
    originalDirtyFilesChecked: baseline.dirty?.length ?? 0, acknowledgedGitTransition, failures }, null, 2));
  if (failures.length) process.exitCode = 1;
} else {
  console.log(JSON.stringify(manifest, null, 2));
}

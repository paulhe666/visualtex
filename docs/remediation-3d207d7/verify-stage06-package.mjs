// Read-only comparison of extracted build artifacts; does not install or test Word.
import { readFileSync, writeFileSync, readdirSync, statSync } from 'node:fs';
import { resolve, join, dirname, relative } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';

const docs = dirname(fileURLToPath(import.meta.url));
const repo = resolve(docs, '../..');
const app = join(repo, 'apps/windows');
const evidence = join(docs, 'evidence');
const payload = join(evidence, 'stage06-package-payload');
const checks = [];
const hash = bytes => createHash('sha256').update(bytes).digest('hex').toUpperCase();
const record = path => ({ path: relative(repo, path).replaceAll('\\', '/'), size: statSync(path).size, sha256: hash(readFileSync(path)) });
function compare(label, actual, expected) {
  const a = record(actual), b = record(expected);
  checks.push({ label, passed: a.sha256 === b.sha256 && a.size === b.size, packaged: a, built: b });
}
function walk(root) {
  return readdirSync(root, { withFileTypes: true }).flatMap(entry => {
    const path = join(root, entry.name);
    return entry.isDirectory() ? walk(path) : [path];
  });
}

const installer = record(join(app, 'src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe'));
const rawMainPath = join(app, 'src-tauri/target/release/visualtex.exe');
const packedMainPath = join(payload, 'visualtex.exe');
const rawMain = readFileSync(rawMainPath), packedMain = readFileSync(packedMainPath);
const marker = Buffer.from('__TAURI_BUNDLE_TYPE_VAR_UNK');
const markerOffset = rawMain.indexOf(marker);
const expectedPackedMain = Buffer.from(rawMain);
const markerUnique = markerOffset >= 0 && rawMain.indexOf(marker, markerOffset + 1) === -1;
if (markerUnique) Buffer.from('__TAURI_BUNDLE_TYPE_VAR_NSS').copy(expectedPackedMain, markerOffset);
const differences = [];
for (let i = 0; i < Math.max(rawMain.length, packedMain.length); i++) {
  if (rawMain[i] !== packedMain[i]) differences.push(i);
}
checks.push({ label: 'main executable: exact build bytes plus Tauri NSIS bundle marker',
  passed: markerUnique && expectedPackedMain.equals(packedMain),
  packaged: record(packedMainPath), built: record(rawMainPath),
  expectedBundleMarkerChange: 'UNK -> NSS', differentByteOffsets: differences });

compare('Office bridge', join(payload, 'visualtex-windows-office-bridge.exe'),
  join(app, 'src-tauri/binaries/visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe'));
const officeRoot = join(app, 'dist-office-windows-native');
for (const path of walk(officeRoot)) {
  compare('Office frontend ' + relative(officeRoot, path), join(payload, 'office', relative(officeRoot, path)), path);
}
for (const arch of ['x64', 'x86']) {
  const msiName = `VisualTeX-WindowsOffice-VSTO-${arch}.msi`;
  compare(`${arch} MSI resource`, join(payload, 'windows-office', msiName), join(app, 'src-tauri/resources/windows-office', msiName));
  compare(`${arch} MSI build`, join(payload, 'windows-office', msiName), join(app, 'src-windows/VisualTeX.WindowsOffice.Installer/bin', arch, 'Release', msiName));
  const vsto = name => join(app, 'src-windows', name, 'bin', arch, 'Release/net472', name + '.dll');
  compare(`${arch} Word VSTO from packaged MSI`, join(payload, arch, 'WordVstoAssembly'), vsto('VisualTeX.WordVsto'));
  compare(`${arch} PowerPoint VSTO from packaged MSI`, join(payload, arch, 'PowerPointVstoAssembly'), vsto('VisualTeX.PowerPointVsto'));
  compare(`${arch} contracts from packaged MSI`, join(payload, arch, 'ContractsAssembly'), join(app, 'src-windows/VisualTeX.WordVsto/bin', arch, 'Release/net472/VisualTeX.WindowsOffice.Contracts.dll'));
  compare(`${arch} OLE from packaged MSI`, join(payload, arch, 'FormulaOleServerExecutable'),
    join(app, 'src-windows/artifacts/formula-ole-server', arch === 'x64' ? 'x64' : 'Win32', 'Release/VisualTeX.FormulaOleServer.exe'));
}
const sourcePaths = execFileSync('git', ['ls-files', '-z', '--cached', '--others', '--exclude-standard', '--',
  'apps/windows/src', 'apps/windows/src-windows', 'apps/windows/src-tauri/src', 'apps/windows/scripts',
  'apps/windows/package.json', 'apps/windows/src-tauri/Cargo.toml', 'apps/windows/src-tauri/tauri.conf.json',
  'apps/windows/src-tauri/tauri.windows.conf.json'], { cwd: repo, windowsHide: true, maxBuffer: 4 * 1024 * 1024 })
  .toString('utf8').split('\0').filter(Boolean);
const sourceFiles = [...new Set(sourcePaths)].sort().map(path => record(join(repo, path)));
const result = { time: new Date().toISOString(), scope: 'Build/package identity only; no installation or Word acceptance',
  head: execFileSync('git', ['rev-parse', 'HEAD'], { cwd: repo, windowsHide: true }).toString().trim(),
  uncommittedWorkingTree: true, installer, checks, sourceFiles,
  allChecksPassed: checks.every(check => check.passed) };
writeFileSync(join(evidence, 'stage06-package-verification.json'), JSON.stringify(result, null, 2) + '\n');
console.log(JSON.stringify({ checks: checks.length, allChecksPassed: result.allChecksPassed, installer }, null, 2));
if (!result.allChecksPassed) process.exitCode = 1;

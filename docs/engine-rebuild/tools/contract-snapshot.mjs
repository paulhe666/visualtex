// Read-only snapshot of actual TS/Rust/C# contracts, not a replacement protocol.
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, resolve } from 'node:path';
import { createRequire } from 'node:module';
import assert from 'node:assert/strict';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const require = createRequire(resolve(root, 'apps/windows/package.json'));
const ts = require('typescript');
const read = p => readFileSync(resolve(root, 'apps/windows', p), 'utf8');
const hash = data => createHash('sha256').update(data).digest('hex');
const paths = [
  'src/office/shared/sessionClient.ts', 'src/office/shared/sessionPayloadValidation.ts',
  'src/office/shared/formulaMetadata.ts', 'src/office/shared/mathTypeAlignmentGeometry.ts',
  'src/office/dialog/useOfficeSession.ts', 'src/export/exportTypes.ts',
  'src-tauri/src/office/sessions.rs', 'src-tauri/src/office/server.rs',
  'src-windows/VisualTeX.WindowsOffice.Contracts/FormulaOleContract.cs',
  'src-windows/VisualTeX.WindowsOffice.Contracts/OfficeFormulaSizing.cs',
  'src-windows/VisualTeX.WindowsOffice.Contracts/FormulaFontSize.cs',
];
export function makeSnapshot(overrides = {}) {
  const contents = Object.fromEntries(paths.map(p => [p, overrides[p] ?? read(p)]));
  const interfaces = {};
  for (const p of paths.filter(p => p.endsWith('.ts'))) {
    const source = ts.createSourceFile(p, contents[p], ts.ScriptTarget.Latest, true);
    for (const node of source.statements) {
      if (ts.isInterfaceDeclaration(node) && ['OfficeExportResult', 'OfficeFormulaSession', 'VisualTeXFormulaMetadata'].includes(node.name.text)) {
        interfaces[node.name.text] = node.members.map(member => ({
          name: member.name?.getText(source), optional: !!member.questionToken,
          type: member.type?.getText(source).replace(/\s+/g, ' '),
        }));
      }
    }
  }
  const client = contents[paths[0]];
  const routes = [...new Set(client.split(/\r?\n/).filter(s => s.includes('/api/v1/')).map(s => s.trim().replace(/,$/, '')))].sort();
  const csharp = contents['src-windows/VisualTeX.WindowsOffice.Contracts/FormulaOleContract.cs'];
  const ole = Object.fromEntries([...csharp.matchAll(/public const (?:int|string) (\w+) = ([^;]+);/g)].map(m => [m[1], JSON.parse(m[2])]));
  return { schema: 'visualtex-office-contract-snapshot-v1',
    sourceHashes: Object.fromEntries(paths.map(p => [p, hash(Buffer.from(contents[p]))])),
    interfaces, clientRoutes: routes, ole,
    dimensions: { naturalUnit: 'css-px-96dpi', officeUnit: 'pt-72dpi', pointsPerPixel: 0.75, baselineOrigin: 'svg-top', mtefUnitsPerPixel: 24 },
  };
}
const snapshot = makeSnapshot();
if (process.argv.includes('--check')) {
  const expected = JSON.parse(readFileSync(resolve(root, 'docs/engine-rebuild/checkpoints/20260913-P01-contracts.json'), 'utf8'));
  assert.deepEqual(snapshot, expected);
  for (const [path, oldText, newText] of [
    [paths[0], 'mathMl?: string;', 'mathML?: string;'],
    ['src-windows/VisualTeX.WindowsOffice.Contracts/OfficeFormulaSizing.cs', '0.75f', '1f'],
    [paths[0], 'credentials: "same-origin"', 'credentials: "omit"'],
  ]) {
    assert.ok(read(path).includes(oldText));
    assert.throws(() => assert.deepEqual(makeSnapshot({ [path]: read(path).replace(oldText, newText) }), expected));
  }
  console.log(JSON.stringify({ status: 'PASS', sources: paths.length, interfaces: Object.keys(snapshot.interfaces).length,
    routes: snapshot.clientRoutes.length, negativeMutationTests: 3, writes: 0 }));
} else console.log(JSON.stringify(snapshot, null, 2));

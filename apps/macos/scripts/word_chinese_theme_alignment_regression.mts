import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { homedir } from 'node:os';
import { join } from 'node:path';
// Exercise an actual production-inserted document, including its theme fonts,
// mixed text/display styles and wrapping. This never invokes a positioning macro.
const args = process.argv.slice(2);
const mode = args.find(v => v.startsWith('--mode='))?.slice(7) ?? 'bottom';
const count = Number(args.find(v => v.startsWith('--count='))?.slice(8) ?? 5);
const output = args.find(v => v.startsWith('--output='))?.slice(9);
const alignments = { top: ['top', 0], center: ['center', 1], baseline: ['baseline', 2], bottom: ['east asian50', 3], auto: ['auto', 4] } as const;
assert.ok(mode in alignments);
function word(lines: string[]) {
  const r = spawnSync('/usr/bin/osascript', ['-e', 'tell application "Microsoft Word"', ...lines.flatMap(l => ['-e', l]), '-e', 'end tell'], { encoding: 'utf8', timeout: 30000 });
  assert.equal(r.status, 0, r.stderr || r.error?.message);
  return r.stdout.trim();
}
assert.equal(word(['return read only of active document']), 'false', 'Acceptance must use a writable Word document');
const [alignment, code] = alignments[mode as keyof typeof alignments];
word([`set base line alignment of paragraph format of text object of active document to baseline align ${alignment}`]);
const path = join(homedir(), 'Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests/word-image-typography-probe.tsv');
const deadline = Date.now() + 20000;
let header: string[] = [], rows: string[][] = [];
do {
  await new Promise(resolve => setTimeout(resolve, 1000));
  word(['run VB macro macro name "VisualTeX_DumpWordImageTypographyForRegression"']);
  const lines = readFileSync(path, 'utf8').trim().split('\n').map(l => l.split('\t'));
  header = lines.shift()!; rows = lines;
  if (rows.length === count && rows.every(r => Number(r[2]) === code && r[6] === r[10] && r[11] === '0:')) break;
} while (Date.now() < deadline);
assert.equal(header[0], 'HOST');
assert.equal(header[2], 'False', 'Word mutation guard leaked');
assert.equal(header[3], 'False', 'Word document became read-only');
assert.equal(header[5], '0:', 'Paragraph fingerprint failed');
assert.equal(rows.length, count);
for (const [i, row] of rows.entries()) {
  assert.equal(row[1], 'True', `Image ${i + 1} lost its identity`);
  assert.equal(Number(row[2]), code, `Image ${i + 1} changed the user's mode`);
  assert.equal(row[11], '0:', `Image ${i + 1} geometry query failed`);
  assert.equal(row[6], row[10], `Image ${i + 1} idle adjustment did not apply`);
  assert.ok(!row[4].startsWith('+'), 'Word theme token was passed as an installed font');
  assert.equal(row[9].split('|').length, 6, 'Native metrics unavailable');
}
const report = { result: 'PASS', mode, count, header, rows };
if (output) writeFileSync(output, JSON.stringify(report, null, 2));
console.log(JSON.stringify(report, null, 2));

import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const ext = process.platform === 'win32' ? '.exe' : '';
const nativePath = resolve(root, `target/debug/visualtex-engine-test-cli${ext}`);
const wasmPath = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const cases = [];
let source = '', revision = 0;
let lastSnapshot, previousIds = new Set(), maxId = 1;
const scopes = { none: 0, subtree: 0, formula: 0 };
const add = (command, check) => cases.push({ input: JSON.stringify({ protocol: 1, command }), check });
function children(n) {
  if (n.kind === 'row') return n.entries.map(e => e.node);
  if (n.kind === 'fraction') return [n.numerator, n.denominator];
  if (n.kind === 'scripts') return [n.base, ...n.slots.map(s => s.argument)];
  return [];
}
function nodeIds(tree) {
  const found = new Set();
  function visit(n) { assert.ok(!found.has(n.id)); found.add(n.id); children(n).forEach(visit); }
  visit(tree); return found;
}
function normalized(value) {
  if (Array.isArray(value)) return value.map(normalized);
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value)
    .filter(([k]) => k !== 'id' && k !== 'sourceId').map(([k, v]) => [k, normalized(v)]));
  return value;
}
function checkSnapshot(s, wantedSource, wantedRevision) {
  assert.equal(s.source, wantedSource); assert.equal(s.revision, wantedRevision);
  const bytes = Buffer.from(s.source);
  let byte = 0, utf16 = 0;
  for (const token of s.analysis.tokens) {
    const p = token.span;
    assert.equal(p.sourceId, s.sourceId); assert.equal(p.byteStart, byte); assert.equal(p.utf16Start, utf16);
    assert.equal(bytes.subarray(p.byteStart, p.byteEnd).toString('utf8'), s.source.slice(p.utf16Start, p.utf16End));
    byte = p.byteEnd; utf16 = p.utf16End;
  }
  assert.equal(byte, bytes.length); assert.equal(utf16, s.source.length);
  const ids = nodeIds(s.analysis.tree);
  assert.equal(ids.size, s.analysis.facts.nodes);
  return ids;
}
function fullOracle() {
  const input = source;
  add({ op: 'analyzeSource', source: input }, response => {
    assert.equal(response.status, 'ok'); assert.equal(response.result.lossless, input);
    const actual = normalized(response.result.analysis), expected = normalized(lastSnapshot.analysis);
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
      function shape(n) {
        return { kind: n.kind, text: n.value ?? n.text, grouping: n.grouping,
          start: n.origin?.byteStart, end: n.origin?.byteEnd, children: children(n).map(shape) };
      }
      console.error(JSON.stringify({ source: input, snapshotSource: lastSnapshot.source,
        actualFacts: actual.facts, expectedFacts: expected.facts,
        actualTree: shape(actual.tree), expectedTree: shape(expected.tree) }));
    }
    assert.deepEqual(actual, expected, `Full-parse oracle mismatch for ${JSON.stringify(input)}`);
  });
}
function replace(input) {
  const oldRevision = revision; if (input !== source) revision++;
  source = input; const resultRevision = revision;
  add({ op: 'replaceSource', handle: 1, expectedRevision: oldRevision, source: input }, response => {
    assert.equal(response.status, 'ok');
    lastSnapshot = response.result.snapshot;
    previousIds = checkSnapshot(lastSnapshot, input, resultRevision);
    maxId = Math.max(maxId, ...previousIds);
  });
  fullOracle();
}
function patch(startUtf16, endUtf16, replacement, expectedScope) {
  const oldSource = source, oldRevision = revision;
  source = oldSource.slice(0, startUtf16) + replacement + oldSource.slice(endUtf16);
  if (source !== oldSource) revision++;
  const resultSource = source, resultRevision = revision;
  add({ op: 'patchSource', handle: 1, expectedRevision: oldRevision, edit: { startUtf16, endUtf16, replacement } }, response => {
    assert.equal(response.status, 'ok'); assert.equal(response.result.kind, 'sourcePatched');
    const s = response.result.snapshot, map = response.result.mapping;
    const nextIds = checkSnapshot(s, resultSource, resultRevision);
    assert.deepEqual(map.retainedIds, [...previousIds].filter(id => nextIds.has(id)).sort((a,b) => a-b));
    assert.deepEqual(map.removedIds, [...previousIds].filter(id => !nextIds.has(id)).sort((a,b) => a-b));
    assert.deepEqual(map.insertedIds, [...nextIds].filter(id => !previousIds.has(id)).sort((a,b) => a-b));
    for (const id of map.insertedIds) assert.ok(id > maxId);
    assert.equal(map.oldRange.utf16Start, startUtf16); assert.equal(map.oldRange.utf16End, endUtf16);
    assert.equal(map.oldRange.byteStart, Buffer.byteLength(oldSource.slice(0, startUtf16)));
    assert.equal(map.oldRange.byteEnd, Buffer.byteLength(oldSource.slice(0, endUtf16)));
    assert.equal(map.newRange.utf16End, startUtf16 + replacement.length);
    if (expectedScope) assert.equal(map.parseScope.kind, expectedScope);
    scopes[map.parseScope.kind]++;
    if (map.parseScope.kind === 'subtree') assert.ok(map.parseScope.parsedNodes < s.analysis.facts.nodes);
    maxId = Math.max(maxId, ...nextIds); previousIds = nextIds; lastSnapshot = s;
  });
  fullOracle();
}
function failure(command, code) {
  add(command, r => { assert.equal(r.status, 'error'); assert.equal(r.error.code, code); });
  add({ op: 'inspectDocument', handle: 1, expectedRevision: revision }, r => assert.deepEqual(r.result.snapshot, lastSnapshot));
}
add({ op: 'create', documentId: 'owned-source-patch-parity' }, r => assert.equal(r.result.handle, 1));
replace(String.raw`\frac{x+x}{y+y}`);
patch(6, 7, 'q', 'subtree'); patch(6, 7, 'q', 'none');
patch(7, 7, '%', 'formula');
replace(String.raw`\frac{}{y}`); patch(6, 6, 'x+x', 'subtree');
replace(String.raw`x+x+x`); patch(2, 2, 'y+', 'formula'); patch(0, 1, '', 'formula');
replace(String.raw`\frac{\text{速度😀é}}{y}`);
const emoji = source.indexOf('😀'); patch(emoji, emoji+2, '🌍', 'subtree');
failure({ op: 'patchSource', handle: 1, expectedRevision: revision, edit: { startUtf16: emoji+1, endUtf16: emoji+1, replacement: 'x' } }, 'invalidPosition');
failure({ op: 'patchSource', handle: 1, expectedRevision: revision-1, edit: { startUtf16: 0, endUtf16: 0, replacement: '' } }, 'staleRevision');
failure({ op: 'patchSource', handle: 1, expectedRevision: revision, edit: { startUtf16: 6, endUtf16: 6, replacement: 'x'.repeat(9000) } }, 'resourceLimit');
replace(String.raw`\future*[opt]{x}+\frac{x}{y}+\alpha`);
let x = source.indexOf(String.raw`\frac{x}`) + 6; patch(x, x+1, 'zz', 'subtree');
replace(String.raw`\frac{x}{`); patch(source.length, source.length, 'y}', 'formula');
replace(String.raw`\frac{x_{1}+x_{1}}{y}+\text{速度}`);
let currentNumber = '1';
for (let i=2; i<=201; i++) {
  const start = source.indexOf(`_{${currentNumber}}`) + 2;
  patch(start, start+currentNumber.length, String(i), 'subtree');
  currentNumber = String(i);
}
add({ op: 'destroy', handle: 1, expectedRevision: revision }, r => assert.equal(r.result.kind, 'destroyed'));
const native = spawnSync(nativePath, [], { input: cases.map(c => c.input).join('\n')+'\n', encoding: 'utf8', maxBuffer: 128*1024*1024, timeout: 30000 });
assert.equal(native.error, undefined); assert.equal(native.status, 0, native.stderr);
const rows = native.stdout.trim().split(/\r?\n/).map(line => JSON.parse(line));
assert.equal(rows.length, cases.length);
const wasm = new EngineRuntime();
try {
  for (let i=0; i<cases.length; i++) {
    const actual = JSON.parse(wasm.dispatch(cases[i].input));
    assert.deepEqual(actual, rows[i], `Native/WASM source patch mismatch at ${i}`);
    cases[i].check(actual);
  }
  assert.equal(wasm.activeSessions, 0);
} finally { wasm.free(); }
const sha = p => createHash('sha256').update(readFileSync(p)).digest('hex');
console.log(JSON.stringify({ status: 'PASS', actualWasmExecution: true, requestCases: cases.length, finalRevision: revision,
  scopes, nativeSha256: sha(nativePath), wasmSha256: sha(wasmPath),
  scopeDisclosure: 'Whole-formula tokenization and validation remain enabled; closed mathematical subtrees are parsed locally.' }, null, 2));

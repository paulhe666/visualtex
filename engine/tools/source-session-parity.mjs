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
let revision = 0;
let source = '';
function add(command, check) { cases.push({ input: JSON.stringify({ protocol: 1, command }), check }); }
function failure(command, code) {
  add(command, response => { assert.equal(response.status, 'error'); assert.equal(response.error.code, code); });
}
function children(node) {
  switch (node.kind) {
    case 'row': return node.entries.map(e => e.node);
    case 'fraction': return [node.numerator, node.denominator];
    case 'scripts': return [node.base, ...node.slots.map(s => s.argument)];
    default: return [];
  }
}
function validateSnapshot(snapshot) {
  const bytes = Buffer.from(snapshot.source);
  let byte = 0, utf16 = 0;
  for (const token of snapshot.analysis.tokens) {
    const span = token.span;
    assert.equal(span.sourceId, snapshot.sourceId);
    assert.equal(span.byteStart, byte); assert.equal(span.utf16Start, utf16);
    assert.equal(bytes.subarray(span.byteStart, span.byteEnd).toString('utf8'), snapshot.source.slice(span.utf16Start, span.utf16End));
    byte = span.byteEnd; utf16 = span.utf16End;
  }
  assert.equal(byte, bytes.length); assert.equal(utf16, snapshot.source.length);
  const ids = new Set();
  function visit(node) {
    assert.ok(Number.isSafeInteger(node.id) && node.id > 0);
    assert.ok(!ids.has(node.id), `Shared identity ${node.id}`);
    ids.add(node.id);
    if (node.origin) {
      const span = node.origin;
      assert.equal(span.sourceId, snapshot.sourceId);
      assert.equal(bytes.subarray(span.byteStart, span.byteEnd).toString('utf8'), snapshot.source.slice(span.utf16Start, span.utf16End));
    }
    for (const child of children(node)) visit(child);
  }
  visit(snapshot.analysis.tree);
  assert.equal(snapshot.analysis.tree.id, 1);
  assert.equal(ids.size, snapshot.analysis.facts.nodes);
}
let lastCompleteTree = null;
let lastSnapshot = null;
function replace(nextSource, expected = {}) {
  const oldRevision = revision;
  if (nextSource !== source) revision++;
  source = nextSource;
  const resultingRevision = revision;
  add({ op: 'replaceSource', handle: 1, expectedRevision: oldRevision, source: nextSource }, response => {
    assert.equal(response.status, 'ok'); assert.equal(response.result.kind, 'document');
    const s = response.result.snapshot;
    assert.equal(s.revision, resultingRevision); assert.equal(s.source, nextSource);
    assert.equal(s.sourceId, resultingRevision + 1);
    validateSnapshot(s);
    for (const key of ['complete', 'supported', 'exportable']) {
      if (key in expected) assert.equal(s.analysis[key], expected[key], `${key}: ${nextSource}`);
    }
    if ('canonical' in expected) assert.equal(s.canonicalLatex, expected.canonical);
    if (expected.facts) for (const [key, value] of Object.entries(expected.facts)) assert.equal(s.analysis.facts[key], value);
    if (s.analysis.complete) {
      assert.equal(s.retainedTree, null);
      lastCompleteTree = structuredClone(s.analysis.tree);
    } else {
      assert.deepEqual(s.retainedTree, lastCompleteTree);
      assert.equal(s.canonicalLatex, null);
    }
    lastSnapshot = structuredClone(s);
  });
}
function inspectUnchanged() {
  const expectedRevision = revision;
  add({ op: 'inspectDocument', handle: 1, expectedRevision }, response => {
    assert.equal(response.status, 'ok');
    assert.deepEqual(response.result.snapshot, lastSnapshot);
  });
}
add({ op: 'create', documentId: 'owned-source-session-parity' }, response => assert.equal(response.result.handle, 1));
replace(String.raw`\frac{x_1+\text{速度😀é}}{y^2}`, {
  complete: true, supported: true, exportable: true,
  facts: { fractions: 1, scripts: 2, textRuns: 1 },
  canonical: String.raw`\frac{x_{1}+\text{速度😀é}}{y^{2}}`,
});
inspectUnchanged();
replace(source, { complete: true, supported: true, exportable: true });
failure({ op: 'replaceSource', handle: 1, expectedRevision: 0, source: 'must not apply' }, 'staleRevision');
inspectUnchanged();
replace(String.raw`\frac{x}{`, { complete: false, exportable: false, canonical: null });
replace(String.raw`\frac{x}{y^`, { complete: false, exportable: false, canonical: null });
replace(String.raw`\future*[x]{a{b}}{c}+z`, { complete: true, supported: false, exportable: false, canonical: null });
replace(String.raw`\frac{x}{2}`, { complete: true, supported: true, exportable: true });
failure({ op: 'replaceSource', handle: 1, expectedRevision: revision, source: 'x'.repeat(9000) }, 'resourceLimit');
inspectUnchanged();
failure({ op: 'replaceSource', handle: 1, expectedRevision: revision, source: ' '.repeat(262145) }, 'resourceLimit');
inspectUnchanged();
replace('', { complete: true, supported: true, exportable: false, canonical: '' });
replace(String.raw`\placeholder{分子}`, { complete: true, supported: true, exportable: false, canonical: String.raw`\placeholder{分子}` });
replace(String.raw`x^2^3`, { complete: false, supported: true, exportable: false, canonical: null });
replace(String.raw`\alpha x^2`, { complete: true, supported: true, exportable: true, canonical: String.raw`\alpha x^{2}` });
for (let index = 0; index < 200; index++) {
  replace(String.raw`\frac{x_{${index}}+\text{速度😀}}{y^{${index + 1}}}`, {
    complete: true, supported: true, exportable: true, facts: { fractions: 1, scripts: 2, textRuns: 1 },
  });
}
inspectUnchanged();
failure({ op: 'replaceSource', handle: 999, expectedRevision: 0, source: 'x' }, 'invalidHandle');
add({ op: 'destroy', handle: 1, expectedRevision: revision }, response => assert.equal(response.result.kind, 'destroyed'));
failure({ op: 'inspectDocument', handle: 1 }, 'invalidHandle');
const native = spawnSync(nativePath, [], { input: cases.map(c => c.input).join('\n') + '\n', encoding: 'utf8',
  maxBuffer: 64 * 1024 * 1024, timeout: 30000 });
assert.equal(native.error, undefined); assert.equal(native.status, 0, native.stderr);
const responses = native.stdout.trim().split(/\r?\n/).map(line => JSON.parse(line));
assert.equal(responses.length, cases.length);
const wasm = new EngineRuntime();
try {
  for (let index = 0; index < cases.length; index++) {
    const response = JSON.parse(wasm.dispatch(cases[index].input));
    assert.deepEqual(response, responses[index], `Stateful Native/WASM mismatch at ${index}`);
    cases[index].check(response);
  }
  assert.equal(wasm.activeSessions, 0);
} finally { wasm.free(); }
const sha = p => createHash('sha256').update(readFileSync(p)).digest('hex');
console.log(JSON.stringify({ status: 'PASS', actualWasmExecution: true, statefulCases: cases.length,
  finalRevision: revision, nativeSha256: sha(nativePath), wasmSha256: sha(wasmPath),
  coverage: ['source ownership', 'draft retention', 'opaque macros', 'Unicode spans', 'stable no-op', 'atomic failure', 'actual Rust syntax trees'] }, null, 2));

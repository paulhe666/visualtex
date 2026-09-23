import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { performance } from 'node:perf_hooks';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const cli = resolve(root, 'target/debug/visualtex-engine-test-cli.exe');
const wasmFile = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const corpus = JSON.parse(readFileSync(resolve(root, 'fixtures/k0-review-corpus.json'), 'utf8'));
const cases = corpus.cases.map(c => ({ ...c, command: { op: 'analyzeSource', source: c.source } }));
let seed = 0x1312ab56;
const next = () => { seed ^= seed << 13; seed ^= seed >>> 17; seed ^= seed << 5; return seed >>> 0; };
function expression(depth) {
  if (!depth) return ['x', '12', '\\alpha x', '\\text{速度 é}', '\\%', '(a+b)'][next() % 6];
  switch (next() % 5) {
    case 0: return `\\frac{${expression(depth - 1)}}{${expression(depth - 1)}}`;
    case 1: return `{${expression(depth - 1)}}_{${expression(depth - 1)}}^{${expression(depth - 1)}}`;
    case 2: return `{${expression(depth - 1)}}`;
    case 3: return `${expression(depth - 1)} % retained\r\n+ ${expression(depth - 1)}`;
    default: return expression(0);
  }
}
for (let i = 0; i < 256; i++) {
  const source = expression(3);
  cases.push({ id: `seed-1312ab56-${i}`, source, expect: 'supported', command: { op: 'analyzeSource', source } });
}
for (const [id, source] of [
  ['depth-budget', '{'.repeat(80) + 'x' + '}'.repeat(80)],
  ['node-budget', 'x'.repeat(8194)],
  ['diagnostic-budget', '\\unknown '.repeat(200)],
  ['source-byte-budget', '中'.repeat(90_000)],
]) cases.push({ id, source, error: 'resourceLimit', command: { op: 'analyzeSource', source } });
for (const command of [
  { op: 'analyzeSource', source: 'x', unknown: true },
  { op: 'analyzeSource', source: null },
  { op: 'analyzeSource', source: 123 },
]) cases.push({ id: 'strict-source-schema', error: 'invalidInput', command });
cases.push({ id: 'recovery-after-errors', source: '\\frac{x_1}{y^2}', expect: 'supported',
  command: { op: 'analyzeSource', source: '\\frac{x_1}{y^2}' } });

const request = command => JSON.stringify({ protocol: 1, command });
const native = spawnSync(cli, [], {
  input: cases.map(c => request(c.command)).join('\n') + '\n', encoding: 'utf8',
  maxBuffer: 64 * 1024 * 1024, timeout: 60_000,
});
assert.equal(native.error, undefined, 'native process launch failed');
assert.equal(native.status, 0, native.stderr);
const rows = native.stdout.trim().split(/\r?\n/).map(row => JSON.parse(row));
assert.equal(rows.length, cases.length);
const wasm = new EngineRuntime();
const timings = [];
const call = command => JSON.parse(wasm.dispatch(request(command)));
function verifyRanges(analysis, source, id) {
  const bytes = Buffer.from(source, 'utf8');
  let offset = 0;
  let joined = '';
  for (const token of analysis.tokens) {
    const s = token.span;
    assert.equal(s.byteStart, offset, `${id}: token gap`);
    assert.ok(s.byteEnd > s.byteStart);
    const piece = bytes.subarray(s.byteStart, s.byteEnd).toString('utf8');
    assert.equal(s.utf16Start, bytes.subarray(0, s.byteStart).toString('utf8').length);
    assert.equal(s.utf16End, bytes.subarray(0, s.byteEnd).toString('utf8').length);
    joined += piece; offset = s.byteEnd;
  }
  assert.equal(offset, bytes.length);
  assert.equal(joined, source);
  const ids = new Set();
  function walk(node) {
    assert.ok(Number.isInteger(node.id) && node.id > 0 && !ids.has(node.id), `${id}: invalid/duplicate node id`);
    ids.add(node.id);
    if (node.origin) {
      assert.equal(node.origin.sourceId, 1);
      assert.ok(node.origin.byteStart >= 0 && node.origin.byteEnd <= bytes.length);
    }
    if (node.kind === 'row') node.entries.forEach(entry => walk(entry.node));
    if (node.kind === 'fraction') { walk(node.numerator); walk(node.denominator); }
    if (node.kind === 'scripts') { walk(node.base); node.slots.forEach(slot => walk(slot.argument)); }
  }
  walk(analysis.tree);
  assert.equal(ids.size, analysis.facts.nodes);
}
try {
  // Analysis must not create/mutate any document, even on resource failure.
  const created = call({ op: 'create', documentId: 'review-preserved-document' });
  assert.equal(created.status, 'ok');
  const handle = created.result.handle;
  const before = call({ op: 'inspect', handle });
  for (let i = 0; i < cases.length; i++) {
    const c = cases[i];
    const start = performance.now();
    const result = call(c.command);
    timings.push(performance.now() - start);
    assert.deepEqual(result, rows[i], `Native/WASM mismatch: ${c.id}`);
    if (c.error) {
      assert.equal(result.status, 'error', c.id);
      assert.equal(result.error.code, c.error, c.id);
      continue;
    }
    assert.equal(result.status, 'ok', c.id);
    assert.equal(result.result.kind, 'sourceAnalysis');
    assert.equal(result.result.lossless, c.source, `${c.id}: source changed`);
    const a = result.result.analysis;
    verifyRanges(a, c.source, c.id);
    if (c.expect === 'supported') {
      assert.ok(a.supported && a.complete, c.id);
      assert.equal(typeof result.result.canonical, 'string');
      if (c.canonical !== undefined) assert.equal(result.result.canonical, c.canonical, c.id);
      if (c.fractions !== undefined) assert.equal(a.facts.fractions, c.fractions, c.id);
      if (c.scripts !== undefined) assert.equal(a.facts.scripts, c.scripts, c.id);
      const again = call({ op: 'analyzeSource', source: result.result.canonical });
      assert.equal(again.status, 'ok');
      assert.ok(again.result.analysis.complete && again.result.analysis.supported);
      assert.equal(again.result.canonical, result.result.canonical, `${c.id}: non-idempotent normalization`);
    } else {
      assert.equal(a.exportable, false, c.id);
      assert.equal(result.result.canonical, null, `${c.id}: invalid input offered as a normalized replacement`);
      if (c.expect === 'unsupported') assert.equal(a.supported, false, c.id);
      else assert.equal(a.complete, false, c.id);
    }
  }
  assert.deepEqual(call({ op: 'inspect', handle }), before, 'analysis mutated an active document');
  assert.equal(wasm.activeSessions, 1);
  assert.equal(call({ op: 'destroy', handle, expectedRevision: before.result.snapshot.revision }).status, 'ok');
  assert.equal(wasm.activeSessions, 0);
} finally { wasm.free(); }
const digest = path => createHash('sha256').update(readFileSync(path)).digest('hex');
const sorted = [...timings].sort((a, b) => a - b);
console.log(JSON.stringify({ status: 'PASS', actualWasmExecution: true, comparison: 'full parsed JSON equality, including every node, span, diagnostic and canonical source',
  cases: cases.length, independentCorpus: corpus.cases.length, generatedCases: 256, seed: '0x1312ab56',
  syntaxOnlyTimingMs: { median: sorted[Math.floor(sorted.length / 2)], p95: sorted[Math.floor(sorted.length * 0.95)] },
  nativeSha256: digest(cli), wasmSha256: digest(wasmFile), documentMutation: false,
  notCovered: ['P04 transactions', 'font/layout', 'browser events', 'system IME', 'Office UI'] }, null, 2));

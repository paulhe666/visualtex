// Real Native/WASM source-patch replay with independent source-string and
// full-parse oracles. No UI event injection or old formula engine is involved.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const cli = resolve(root, 'target/debug/visualtex-engine-test-cli.exe');
const wasmFile = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const cases = [];
let source = '';
let revision = 0;
const append = (label, command, extra = {}) => cases.push({ label, command, source, revision, ...extra });
append('open source editor', { op: 'create', documentId: 'source-edit-review' });
append('open unrelated document', { op: 'create', documentId: 'untouched' });
function patch(startUtf16, endUtf16, replacement, label = 'patch', extra = {}) {
  const expectedRevision = revision;
  const next = source.slice(0, startUtf16) + replacement + source.slice(endUtf16);
  const changed = next !== source;
  if (changed) revision++;
  source = next;
  append(label, { op: 'patchSource', handle: 1, expectedRevision, edit: { startUtf16, endUtf16, replacement } }, { changed, ...extra });
  append(`${label}: independent full parse`, { op: 'analyzeSource', source }, { fullOracle: true });
}
const original = '\\frac{x_1+\\text{速度}}{y^2}+\\frac{x}{y}+z';
patch(0, 0, original, 'initial source');
let at = source.indexOf('x_1');
patch(at, at + 1, 'a+2', 'fraction numerator edit', { requireSubtree: true });
at = source.indexOf('速度');
patch(at, at + 2, '速度😀', 'text astral edit', { requireSubtree: true });
at = source.indexOf('😀');
append('surrogate split rejected', { op: 'patchSource', handle: 1, expectedRevision: revision,
  edit: { startUtf16: at + 1, endUtf16: at + 1, replacement: 'x' } }, { error: 'invalidPosition' });
append('stale patch rejected before parsing', { op: 'patchSource', handle: 1, expectedRevision: revision - 1,
  edit: { startUtf16: 0, endUtf16: 0, replacement: 'x'.repeat(262145) } }, { error: 'staleRevision' });
append('node budget rejection is atomic', { op: 'patchSource', handle: 1, expectedRevision: revision,
  edit: { startUtf16: 0, endUtf16: source.length, replacement: 'x'.repeat(9000) } }, { error: 'resourceLimit' });
append('reversed patch rejected', { op: 'patchSource', handle: 1, expectedRevision: revision,
  edit: { startUtf16: 5, endUtf16: 2, replacement: 'x' } }, { error: 'invalidPosition' });
append('nested patch DTO rejects unknown fields', { op: 'patchSource', handle: 1, expectedRevision: revision,
  edit: { startUtf16: 0, endUtf16: 0, replacement: 'x', ignoredTypo: true } }, { error: 'invalidInput' });
patch(0, source.length, source, 'exact noop');
patch(source.length, source.length, '\\frac{x}{', 'incomplete source retained');
patch(source.length, source.length, '2}', 'incomplete source completed');
patch(0, 0, '\\unknown[a]{b}+', 'unsupported prefix retained');
patch(0, '\\unknown[a]{b}+'.length, '', 'unsupported prefix removed');

// Deterministic sequential edits, not independent resets: this exercises source
// identity rebasing and monotonically allocated identities over many revisions.
let random = 0x54203193;
const next = () => { random ^= random << 13; random ^= random >>> 17; random ^= random << 5; return random >>> 0; };
const replacements = ['', 'x', '中', '😀', 'é', '\\', '{', '}', '%', '\r\n', '^', '_', '\\frac{a}{b}', '\\text{速度}'];
for (let index = 0; index < 320; index++) {
  const boundaries = [0];
  let utf16 = 0;
  for (const ch of source) { utf16 += ch.length; boundaries.push(utf16); }
  const a = boundaries[next() % boundaries.length];
  const b = boundaries[next() % boundaries.length];
  patch(Math.min(a, b), Math.max(a, b), replacements[next() % replacements.length], `seed-54203193-${index}`);
}
append('final source read', { op: 'inspectDocument', handle: 1, expectedRevision: revision });
append('unrelated document unchanged', { op: 'inspectDocument', handle: 2, expectedRevision: 0 }, { unrelated: true });
append('close edited document', { op: 'destroy', handle: 1, expectedRevision: revision });
append('close unrelated document', { op: 'destroy', handle: 2, expectedRevision: 0 });
append('no leaked documents', { op: 'stats' });

const wire = command => JSON.stringify({ protocol: 1, command });
const processResult = spawnSync(cli, [], { input: cases.map(c => wire(c.command)).join('\n') + '\n',
  encoding: 'utf8', maxBuffer: 96 * 1024 * 1024, timeout: 60_000 });
assert.equal(processResult.error, undefined);
assert.equal(processResult.status, 0, processResult.stderr);
const native = processResult.stdout.trim().split(/\r?\n/).map(line => JSON.parse(line));
assert.equal(native.length, cases.length);
function normalize(analysis) {
  const copy = structuredClone(analysis);
  const walk = object => {
    if (Array.isArray(object)) { object.forEach(walk); return; }
    if (object && typeof object === 'object') {
      delete object.id; delete object.sourceId;
      Object.values(object).forEach(walk);
    }
  };
  walk(copy); return copy;
}
function identities(node) {
  const ids = new Set();
  function walk(n) {
    assert.ok(Number.isInteger(n.id) && n.id > 0 && !ids.has(n.id)); ids.add(n.id);
    if (n.kind === 'row') n.entries.forEach(e => walk(e.node));
    if (n.kind === 'fraction') { walk(n.numerator); walk(n.denominator); }
    if (n.kind === 'scripts') { walk(n.base); n.slots.forEach(s => walk(s.argument)); }
  }
  walk(node); return ids;
}
const sortIds = ids => [...ids].sort((a, b) => a - b);
let previousSnapshot;
let lastAnalysis;
let subtreePatches = 0;
let fullPatches = 0;
let noops = 0;
const everIssued = new Set([1]);
const wasm = new EngineRuntime();
try {
  for (let i = 0; i < cases.length; i++) {
    const c = cases[i];
    const web = JSON.parse(wasm.dispatch(wire(c.command)));
    assert.deepEqual(web, native[i], `Native/WASM mismatch at ${c.label}`);
    if (c.error) {
      assert.equal(web.status, 'error', c.label);
      assert.equal(web.error.code, c.error, c.label);
      const after = JSON.parse(wasm.dispatch(wire({ op: 'inspectDocument', handle: 1 })));
      assert.deepEqual(after.result.snapshot, previousSnapshot, `failed patch changed document: ${c.label}`);
      continue;
    }
    assert.equal(web.status, 'ok', `${c.label}: ${JSON.stringify(web)}`);
    if (c.fullOracle) {
      assert.equal(web.result.kind, 'sourceAnalysis');
      assert.equal(web.result.lossless, c.source);
      assert.deepEqual(normalize(lastAnalysis), normalize(web.result.analysis), `incremental/full mismatch: ${c.label}`);
      continue;
    }
    if (c.command.op === 'patchSource') {
      assert.equal(web.result.kind, 'sourcePatched');
      const { snapshot, mapping } = web.result;
      assert.equal(snapshot.source, c.source, c.label);
      assert.equal(snapshot.revision, c.revision, c.label);
      const currentIds = identities(snapshot.analysis.tree);
      const oldIds = previousSnapshot ? identities(previousSnapshot.analysis.tree) : new Set([1]);
      assert.deepEqual(sortIds(mapping.retainedIds), sortIds([...oldIds].filter(id => currentIds.has(id))));
      assert.deepEqual(sortIds(mapping.removedIds), sortIds([...oldIds].filter(id => !currentIds.has(id))));
      assert.deepEqual(sortIds(mapping.insertedIds), sortIds([...currentIds].filter(id => !oldIds.has(id))));
      for (const id of mapping.insertedIds) {
        assert.equal(everIssued.has(id), false, `reused a removed node identity: ${c.label}`); everIssued.add(id);
      }
      for (const id of mapping.changedIds) assert.ok(currentIds.has(id));
      if (!c.changed) {
        assert.equal(mapping.parseScope.kind, 'none');
        assert.equal(mapping.lexedTokens, 0);
        assert.deepEqual(snapshot, previousSnapshot);
        noops++;
      } else {
        assert.ok(snapshot.sourceId > (previousSnapshot?.sourceId ?? 1));
        assert.equal(mapping.oldRange.sourceId, previousSnapshot?.sourceId ?? 1);
        assert.equal(mapping.newRange.sourceId, snapshot.sourceId);
        if (mapping.parseScope.kind === 'subtree') subtreePatches++; else fullPatches++;
      }
      if (c.requireSubtree) assert.equal(mapping.parseScope.kind, 'subtree', c.label);
      previousSnapshot = snapshot;
      lastAnalysis = snapshot.analysis;
    } else if (c.unrelated) {
      assert.equal(web.result.snapshot.documentId, 'untouched');
      assert.equal(web.result.snapshot.source, '');
      assert.equal(web.result.snapshot.revision, 0);
    }
  }
  assert.equal(wasm.activeSessions, 0);
} finally { wasm.free(); }
const sha = path => createHash('sha256').update(readFileSync(path)).digest('hex');
console.log(JSON.stringify({ status: 'PASS', actualWasmExecution: true, replayRequests: cases.length,
  sequentialGeneratedPatches: 320, seed: '0x54203193', subtreePatches, fullPatches, noops,
  finalRevision: revision, historicalNodeIdentities: everIssued.size,
  comparison: 'full Native/WASM equality + independent source splice + full-parse oracle + identity partition',
  nativeSha256: sha(cli), wasmSha256: sha(wasmFile),
  notCovered: ['structural transaction undo/redo', 'pointer/keyboard input', 'IME', 'math layout', 'Office UI'] }, null, 2));

// A reproducible P03/P04 source-path baseline, NOT user-visible typing or Office
// latency. The debug WASM path includes JSON encode/decode and full snapshots.
import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { createRequire } from 'node:module';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { cpus, platform, arch } from 'node:os';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const wasmPath = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const engine = new EngineRuntime();
function call(command) {
  const request = JSON.stringify({ protocol: 1, command });
  const start = performance.now();
  const raw = engine.dispatch(request);
  const dispatched = performance.now();
  const result = JSON.parse(raw);
  const decoded = performance.now();
  assert.equal(result.status, 'ok', result.error?.message);
  return { result: result.result, dispatchMs: dispatched - start, decodeMs: decoded - dispatched,
    totalMs: decoded - start, responseBytes: Buffer.byteLength(raw) };
}
const stats = values => {
  const sorted = [...values].sort((a, b) => a - b);
  return { min: sorted[0], median: sorted[Math.floor(sorted.length / 2)], p95: sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * .95))], max: sorted.at(-1) };
};
const rows = [];
try {
  for (const terms of [32, 256, 2048]) {
    const source = '\\frac{a+b}{c}' + '+x'.repeat(terms);
    const { result: opened } = call({ op: 'create', documentId: `source-perf-${terms}` });
    const handle = opened.handle;
    const seeded = call({ op: 'replaceSource', handle, expectedRevision: 0, source });
    let revision = 1;
    const at = source.indexOf('a');
    // indexOf('a') would match the command's spelling, so choose the argument.
    const argumentAt = source.indexOf('{a') + 1;
    assert.ok(argumentAt > at);
    const timings = [];
    const scopes = [];
    let highWaterNodeId = 0;
    for (let iteration = 0; iteration < 21; iteration++) {
      const replacement = iteration % 2 ? 'a' : 'z';
      const before = revision;
      const run = call({ op: 'patchSource', handle, expectedRevision: before,
        edit: { startUtf16: argumentAt, endUtf16: argumentAt + 1, replacement } });
      revision++;
      assert.equal(run.result.snapshot.revision, revision);
      assert.equal(run.result.mapping.parseScope.kind, 'subtree');
      assert.equal(run.result.snapshot.source, source.slice(0, argumentAt) + replacement + source.slice(argumentAt + 1));
      highWaterNodeId = Math.max(highWaterNodeId, ...run.result.mapping.insertedIds);
      if (iteration >= 6) { timings.push(run); scopes.push({ parsedNodes: run.result.mapping.parseScope.parsedNodes,
        lexedTokens: run.result.mapping.lexedTokens }); }
    }
    const noop = call({ op: 'patchSource', handle, expectedRevision: revision,
      edit: { startUtf16: argumentAt, endUtf16: argumentAt + 1, replacement: 'z' } });
    assert.equal(noop.result.mapping.parseScope.kind, 'none');
    assert.equal(noop.result.snapshot.revision, revision);
    call({ op: 'destroy', handle, expectedRevision: revision });
    rows.push({ repeatedOuterTerms: terms, sourceBytes: Buffer.byteLength(source), nodes: seeded.result.snapshot.analysis.facts.nodes,
      parseScope: 'subtree', parsedNodes: scopes[0].parsedNodes, lexedTokens: scopes[0].lexedTokens,
      warmupEdits: 6, measuredEdits: timings.length, coldSourceOpenMs: seeded.totalMs,
      patchDispatchMs: stats(timings.map(r => r.dispatchMs)), jsonDecodeMs: stats(timings.map(r => r.decodeMs)),
      sourcePatchRoundTripMs: stats(timings.map(r => r.totalMs)), fullSnapshotBytes: timings.at(-1).responseBytes,
      noopRoundTripMs: noop.totalMs, highWaterNodeId });
  }
  assert.equal(engine.activeSessions, 0);
} finally { engine.free(); }
console.log(JSON.stringify({ status: 'PASS', baselineOnly: true, actualWasmExecution: true, buildProfile: 'debug',
  wasmSha256: createHash('sha256').update(readFileSync(wasmPath)).digest('hex'), environment: { node: process.version,
    platform: platform(), arch: arch(), cpu: cpus()[0]?.model }, rows,
  limitations: ['Not a browser event-to-paint or Office benchmark', 'Global lexing, rebasing, validation and full snapshot serialization remain present',
    'No claim of constant-time source edits', 'Structural selection/layout/history performance requires separate measurement'] }, null, 2));

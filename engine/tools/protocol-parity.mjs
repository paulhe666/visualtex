import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { readFileSync, existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const cli = resolve(root, 'target/debug/visualtex-engine-test-cli.exe');
const wasmPath = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
for (const path of [cli, wasmPath]) assert.ok(existsSync(path), `Build the native CLI and WASM bindings first: ${path}`);
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const cases = [];
const ok = result => ({ status: 'ok', result });
const snapshot = (handle, documentId) => ok({ kind: 'session', handle,
  snapshot: { documentId, revision: 0, rootId: 1, source: '' } });
const add = (label, command, expected) => cases.push({ label, input: JSON.stringify({ protocol: 1, command }), expected });
const addRaw = (label, input, code) => cases.push({ label, input, expected: { status: 'error', code } });
const error = code => ({ status: 'error', code });
add('initial stats', { op: 'stats' }, ok({ kind: 'stats', protocol: 1, activeSessions: 0 }));
add('Unicode opaque document ID', { op: 'create', documentId: '文档-A' }, snapshot(1, '文档-A'));
add('duplicate document', { op: 'create', documentId: '文档-A' }, error('conflict'));
add('stale read', { op: 'inspect', handle: 1, expectedRevision: 1 }, error('staleRevision'));
add('stale destructive request', { op: 'destroy', handle: 1, expectedRevision: 1 }, error('staleRevision'));
add('failed destruction preserved document', { op: 'inspect', handle: 1, expectedRevision: 0 }, snapshot(1, '文档-A'));
add('destroy', { op: 'destroy', handle: 1, expectedRevision: 0 }, ok({ kind: 'destroyed', handle: 1 }));
add('destroyed handle rejected', { op: 'inspect', handle: 1 }, error('invalidHandle'));
add('recreated document uses fresh handle', { op: 'create', documentId: '文档-A' }, snapshot(2, '文档-A'));
add('park retained session', { op: 'parkSession', handle: 2, expectedRevision: 0 }, ok({ kind: 'sessionParked', handle: 2,
  snapshot: { documentId: '文档-A', revision: 0, rootId: 1, source: '' } }));
add('parked handle is not an active editor', { op: 'inspect', handle: 2 }, error('invalidHandle'));
add('parked document identity remains reserved', { op: 'create', documentId: '文档-A' }, error('conflict'));
add('active stats exclude parked sessions', { op: 'stats' }, ok({ kind: 'stats', protocol: 1, activeSessions: 0 }));
add('resume retained session with same handle', { op: 'resumeSession', handle: 2, expectedRevision: 0 }, ok({ kind: 'sessionResumed', handle: 2,
  snapshot: { documentId: '文档-A', revision: 0, rootId: 1, source: '' } }));
addRaw('protocol independent of Office', JSON.stringify({ protocol: 2, command: { op: 'create', documentId: 'bad' } }), 'unsupportedProtocol');
for (const input of ['{}', 'null', '[]', 'not JSON',
  '{"protocol":1,"command":{"op":"inspect","handle":-1}}',
  '{"protocol":1,"command":{"op":"inspect","handle":1.2}}',
  '{"protocol":1,"command":{"op":"create","documentId":"bad","typo":0}}',
  '{"protocol":1,"command":{"op":"stats"},"unknown":false}',
  '{"protocol":1,"command":{"op":"destroy","handle":2}}']) {
  addRaw('reject invalid wire schema', input, 'invalidInput');
}
add('invalid identifier', { op: 'create', documentId: '\u0000' }, error('invalidInput'));
addRaw('bounded request', ' '.repeat(1_048_577), 'resourceLimit');
add('state intact after invalid requests', { op: 'stats' }, ok({ kind: 'stats', protocol: 1, activeSessions: 1 }));
for (let handle = 3; handle <= 65; handle++) {
  add('fill capacity', { op: 'create', documentId: `capacity-${handle}` }, snapshot(handle, `capacity-${handle}`));
}
add('bounded active sessions', { op: 'create', documentId: 'overflow' }, error('resourceLimit'));
for (let handle = 2; handle <= 65; handle++) {
  add('release capacity', { op: 'destroy', handle, expectedRevision: 0 }, ok({ kind: 'destroyed', handle }));
}
for (let handle = 66; handle < 1066; handle++) {
  add('1000 lifecycle create', { op: 'create', documentId: 'repeated' }, snapshot(handle, 'repeated'));
  add('1000 lifecycle dispose', { op: 'destroy', handle, expectedRevision: 0 }, ok({ kind: 'destroyed', handle }));
}
add('all sessions released', { op: 'stats' }, ok({ kind: 'stats', protocol: 1, activeSessions: 0 }));
const native = spawnSync(cli, [], { input: cases.map(c => c.input).join('\n') + '\n', encoding: 'utf8', maxBuffer: 16 * 1024 * 1024, timeout: 30_000 });
assert.equal(native.error, undefined);
assert.equal(native.status, 0, native.stderr);
const nativeRows = native.stdout.trim().split(/\r?\n/).map(line => JSON.parse(line));
assert.equal(nativeRows.length, cases.length);
const wasm = new EngineRuntime();
try {
  for (let index = 0; index < cases.length; index++) {
    const c = cases[index];
    const web = JSON.parse(wasm.dispatch(c.input));
    assert.deepEqual(web, nativeRows[index], `Native/WASM mismatch at ${index}: ${c.label}`);
    if (c.expected.status === 'error') {
      assert.equal(web.status, 'error', c.label);
      assert.equal(web.error.code, c.expected.code, c.label);
      assert.ok(web.error.message.length > 0 && web.error.message.length < 300);
    } else {
      assert.deepEqual(web, c.expected, `Independent oracle ${index}: ${c.label}`);
    }
  }
  assert.equal(wasm.activeSessions, 0);
  assert.equal(wasm.parkedSessions, 0);
  assert.equal(wasm.parkedWeight, 0);
} finally { wasm.free(); }
for (let index = 0; index < 100; index++) {
  const fresh = new EngineRuntime();
  assert.equal(fresh.activeSessions, 0);
  fresh.free();
}
const sha = path => createHash('sha256').update(readFileSync(path)).digest('hex');
console.log(JSON.stringify({ status: 'PASS', actualWasmExecution: true, requestCases: cases.length,
  repeatedDocumentLifecycles: 1000, repeatedRuntimeLifecycles: 100,
  nativeSha256: sha(cli), wasmSha256: sha(wasmPath), dependencies: 'independent core; no production app import' }, null, 2));

// Host-unit tests. Synthetic byte assets test loading policy only, not font
// parsing or real system IME. Compiled sources live solely under engine/target.
import test from 'node:test';
import assert from 'node:assert/strict';
import { createHash, webcrypto } from 'node:crypto';
import { createRequire } from 'node:module';
import { setImmediate as tick } from 'node:timers/promises';
const require = createRequire(import.meta.url);
const { ViewportMapping, requireCurrentGeometry } = require('../target/host-review/coordinates.js');
const { VerifiedAssetLoader, ResourceError } = require('../target/host-review/resources.js');
const hash = bytes => createHash('sha256').update(bytes).digest('hex');
const asset = (id, bytes, extra = {}) => ({ id, kind: 'font', url: `${id}.otf`, sha256: hash(bytes), byteLength: bytes.length, ...extra });
const manifest = assets => ({ schemaVersion: 1, engineProtocol: 1, buildId: 'review-build', assets });
const options = extra => ({ baseUrl: 'https://engine.invalid/assets/', expectedBuildId: 'review-build', expectedEngineProtocol: 1,
  subtle: webcrypto.subtle, timeoutMs: 1000, ...extra });
const response = (bytes, extra = {}) => new Response(bytes, { status: 200, headers: { 'content-type': 'font/otf', ...extra } });
const isError = code => error => error instanceof ResourceError && error.code === code;
const near = (actual, expected) => assert.ok(Math.abs(actual - expected) < 1e-8, `${actual} != ${expected}`);

// Geometry uses CSS coordinates regardless of devicePixelRatio.
test('viewport transforms invert across scroll, negative origins and zoom', () => {
  for (const zoom of [0.1, 0.25, 0.5, 1, 1.25, 1.5, 2, 4, 16]) {
    for (const scrollLeft of [0, 21.375, 2000]) {
      const view = new ViewportMapping({ left: -53.5, top: 72.25, scrollLeft, scrollTop: 213.125, zoom });
      for (const point of [{ x: 0, y: 0 }, { x: 3.1415926, y: -271.25 }, { x: 20000, y: 4321 }]) {
        const back = view.clientToLayout(view.layoutToClient(point));
        near(back.x, point.x); near(back.y, point.y);
      }
      const rect = { x: 22.5, y: 3, width: 124.25, height: 88.75 };
      const back = view.clientRectToLayout(view.layoutRectToClient(rect));
      for (const key of Object.keys(rect)) near(back[key], rect[key]);
    }
  }
});
test('viewport copies input state rather than observing mutable globals', () => {
  const spec = { left: 100, top: 200, scrollLeft: 10, scrollTop: 20, zoom: 2 };
  const view = new ViewportMapping(spec);
  spec.zoom = 100; spec.scrollTop = 500;
  assert.deepEqual(view.clientToLayout({ x: 110, y: 220 }), { x: 10, y: 20 });
  assert.deepEqual(view.layoutToClient({ x: 10, y: 20 }), { x: 110, y: 220 });
});
test('invalid geometry never reaches hit-testing', () => {
  for (const zoom of [0, -1, NaN, Infinity]) {
    assert.throws(() => new ViewportMapping({ left: 0, top: 0, scrollLeft: 0, scrollTop: 0, zoom }), e => e.code === 'invalidCoordinate');
  }
  const view = new ViewportMapping({ left: 0, top: 0, scrollLeft: 0, scrollTop: 0, zoom: 2 });
  assert.throws(() => view.clientToLayout({ x: NaN, y: 0 }));
  assert.throws(() => view.layoutToClient({ x: Number.MAX_VALUE, y: 0 }));
  assert.throws(() => view.layoutRectToClient({ x: 0, y: 0, width: -1, height: 1 }));
  assert.throws(() => view.clientRectToLayout({ x: 0, y: 0, width: 1, height: Infinity }));
});
test('font/layout generation guards stale geometry even when source revision matches', () => {
  const current = { documentId: 'doc-a', revision: 7, layoutGeneration: 3 };
  requireCurrentGeometry(current, { ...current });
  for (const changed of [{ ...current, documentId: 'doc-b' }, { ...current, revision: 6 }, { ...current, layoutGeneration: 2 }]) {
    assert.throws(() => requireCurrentGeometry(current, changed), e => e.code === 'staleGeometry');
  }
  assert.throws(() => requireCurrentGeometry(current, { ...current, revision: 1.5 }), e => e.code === 'invalidCoordinate');
});

test('loader verifies hash and returns caller-owned bytes, never mutable cached storage', async () => {
  const bytes = Uint8Array.from([1, 2, 3, 4]);
  const descriptor = asset('math', bytes);
  let calls = 0;
  const loader = new VerifiedAssetLoader(manifest([descriptor]), options({ fetch: async (url, init) => {
    calls++;
    assert.equal(url, 'https://engine.invalid/assets/math.otf');
    assert.equal(init.credentials, 'omit'); assert.equal(init.redirect, 'error'); assert.equal(init.mode, 'same-origin');
    return response(bytes);
  }}));
  try {
    descriptor.sha256 = '0'.repeat(64); // manifest is copied by the loader
    const first = await loader.load('math');
    assert.equal(first.sha256, hash(bytes));
    first.bytes[0] = 99;
    const second = await loader.load('math');
    assert.deepEqual(second.bytes, bytes);
    assert.notEqual(first.bytes.buffer, second.bytes.buffer);
    assert.equal(calls, 1);
    assert.deepEqual(loader.stats, { cachedBytes: 4, pendingBytes: 0, activeLoads: 0, cachedAssets: 1 });
  } finally { loader.dispose(); }
  assert.equal(loader.stats.cachedBytes, 0);
});
test('manifest/build mismatch fails before any request', () => {
  let calls = 0;
  const opts = options({ fetch: async () => { calls++; return response(new Uint8Array()); } });
  assert.throws(() => new VerifiedAssetLoader({ ...manifest([]), buildId: 'different' }, opts), isError('versionMismatch'));
  assert.throws(() => new VerifiedAssetLoader({ ...manifest([]), engineProtocol: 2 }, opts), isError('versionMismatch'));
  assert.throws(() => new VerifiedAssetLoader({ ...manifest([]), schemaVersion: 2 }, opts), isError('invalidManifest'));
  assert.equal(calls, 0);
});
test('manifest rejects duplicate ids, invalid hashes and invalid budgets', () => {
  const a = asset('a', Uint8Array.of(1));
  for (const assets of [[a, a], [{ ...a, sha256: 'not a hash' }], [{ ...a, id: ' ' }], [{ ...a, byteLength: -1 }],
    [{ ...a, byteLength: 0 }], [{ ...a, byteLength: Infinity }], [{ ...a, kind: 'script' }]]) {
    assert.throws(() => new VerifiedAssetLoader(manifest(assets), options()), isError('invalidManifest'));
  }
  for (const limits of [{ maxAssetBytes: 10, maxResidentBytes: 9 }, { maxConcurrentLoads: 0 }, { timeoutMs: -1 }, { timeoutMs: 3e9 }]) {
    assert.throws(() => new VerifiedAssetLoader(manifest([a]), options(limits)), isError('invalidManifest'));
  }
});
test('manifest URL confinement rejects remote origins, traversal and API escape', () => {
  const bytes = Uint8Array.of(1);
  for (const url of ['https://other.invalid/a.otf', '../private.otf', '%2e%2e/private.otf', 'folder/%2f..%2fprivate.otf',
    '/api/v1/sessions', 'data:font/otf;base64,QQ==', '//other.invalid/a', 'a.otf?token=secret', 'a.otf#frag',
    'a\\b.otf', 'a%00.otf', 'bad%ZZ.otf']) {
    assert.throws(() => new VerifiedAssetLoader(manifest([asset('a', bytes, { url })]), options()), isError('unsafeUrl'), url);
  }
  for (const baseUrl of ['file:///C:/secret/', 'about:blank', 'https://user:pass@engine.invalid/assets/', 'https://engine.invalid/assets']) {
    assert.throws(() => new VerifiedAssetLoader(manifest([]), options({ baseUrl })), isError('unsafeUrl'));
  }
});
test('encoded spaces and non-ASCII asset names remain loadable', async () => {
  const bytes = Uint8Array.of(1, 2);
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes, { url: '字体%20资源/公式%20数学.otf' })]), options({ fetch: async url => {
    assert.equal(new URL(url).origin, 'https://engine.invalid');
    assert.ok(decodeURIComponent(new URL(url).pathname).endsWith('字体 资源/公式 数学.otf'));
    return response(bytes);
  }}));
  try { assert.deepEqual((await loader.load('a')).bytes, bytes); } finally { loader.dispose(); }
});
test('hash failure does not poison retries or cache partially verified assets', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  let calls = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: async () => response(++calls === 1 ? Uint8Array.of(9, 9, 9) : bytes) }));
  try {
    await assert.rejects(loader.load('a'), isError('hashMismatch'));
    assert.deepEqual(loader.stats, { cachedBytes: 0, pendingBytes: 0, activeLoads: 0, cachedAssets: 0 });
    assert.deepEqual((await loader.load('a')).bytes, bytes);
    assert.equal(calls, 2);
  } finally { loader.dispose(); }
});
test('decoded streamed length is checked independently of Content-Length', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  for (const body of [Uint8Array.of(1), Uint8Array.of(1, 2, 3, 4)]) {
    const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: async () => response(body) }));
    try { await assert.rejects(loader.load('a'), isError('lengthMismatch')); } finally { loader.dispose(); }
  }
});
test('compressed transfer length is not mistaken for uncompressed font length', async () => {
  const bytes = Uint8Array.of(1, 2, 3, 4);
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: async () => response(bytes,
    { 'content-length': '22', 'content-encoding': 'gzip' }) }));
  try { assert.deepEqual((await loader.load('a')).bytes, bytes); } finally { loader.dispose(); }
});
test('bad status, MIME and length cancel response body and leave cache empty', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  for (const [status, headers, expected] of [[404, { 'content-type': 'font/otf' }, 'networkFailure'],
    [206, { 'content-type': 'font/otf' }, 'networkFailure'],
    [200, { 'content-type': 'text/html' }, 'badMime'],
    [200, { 'content-type': 'font/otf', 'content-length': '1000' }, 'lengthMismatch']]) {
    let cancelled = 0;
    const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: async () => new Response(new ReadableStream({
      start(controller) { controller.enqueue(bytes); }, cancel() { cancelled++; },
    }), { status, headers }) }));
    try {
      await assert.rejects(loader.load('a'), isError(expected));
      assert.equal(cancelled, 1);
      assert.equal(loader.stats.cachedBytes, 0);
    } finally { loader.dispose(); }
  }
});
test('WASM loading requires application/wasm rather than relying on a successful fetch alone', async () => {
  const bytes = Uint8Array.of(0, 97, 115, 109, 1, 0, 0, 0); // empty valid module
  for (const mime of ['application/wasm', 'text/html']) {
    const loader = new VerifiedAssetLoader(manifest([asset('core', bytes, { kind: 'wasm', url: 'core.wasm' })]), options({ fetch: async () => response(bytes, { 'content-type': mime }) }));
    try {
      if (mime === 'application/wasm') {
        const verified = await loader.load('core');
        assert.ok(await WebAssembly.compile(verified.bytes) instanceof WebAssembly.Module);
      } else await assert.rejects(loader.load('core'), isError('badMime'));
    } finally { loader.dispose(); }
  }
});
test('shared load has one fetch and cancelling one waiter preserves the other', async () => {
  const bytes = Uint8Array.of(4, 5, 6);
  let resolveFetch;
  let fetchSignal;
  let calls = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: (_url, init) => {
    calls++; fetchSignal = init.signal;
    return new Promise(resolve => { resolveFetch = resolve; });
  }}));
  try {
    const controller = new AbortController();
    const a = loader.load('a', controller.signal);
    const b = loader.load('a');
    const cancelled = assert.rejects(a, isError('cancelled'));
    controller.abort(); await cancelled;
    assert.equal(fetchSignal.aborted, false);
    resolveFetch(response(bytes));
    assert.deepEqual((await b).bytes, bytes);
    assert.equal(calls, 1);
  } finally { loader.dispose(); }
});
test('cancelling every waiter aborts underlying fetch and releases reservation', async () => {
  const bytes = Uint8Array.of(4, 5, 6);
  let aborted = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: (_url, init) => new Promise((_resolve, reject) => {
    init.signal.addEventListener('abort', () => { aborted++; reject(new DOMException('cancelled', 'AbortError')); }, { once: true });
  }) }));
  try {
    const controller = new AbortController();
    const pending = loader.load('a', controller.signal);
    const failure = assert.rejects(pending, isError('cancelled'));
    controller.abort(); await failure; await tick();
    assert.equal(aborted, 1);
    assert.deepEqual(loader.stats, { cachedBytes: 0, pendingBytes: 0, activeLoads: 0, cachedAssets: 0 });
  } finally { loader.dispose(); }
});
test('dispose cancels a stalled body without publishing late results', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  let cancelled = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: async () => new Response(new ReadableStream({
    start(controller) { controller.enqueue(Uint8Array.of(1)); }, cancel() { cancelled++; },
  }), { headers: { 'content-type': 'font/otf' } }) }));
  const pending = loader.load('a');
  const failure = assert.rejects(pending, e => ['cancelled', 'disposed'].includes(e.code));
  await tick(); loader.dispose(); await failure;
  assert.equal(cancelled, 1);
  assert.equal(loader.stats.cachedBytes, 0);
  assert.equal(loader.stats.pendingBytes, 0);
  await assert.rejects(loader.load('a'), isError('disposed'));
  loader.dispose();
});
test('time budget cancels a stalled body and returns a specific timeout', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  let cancelled = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ timeoutMs: 20,
    fetch: async () => new Response(new ReadableStream({ cancel() { cancelled++; } }), { headers: { 'content-type': 'font/otf' } }) }));
  try {
    await assert.rejects(loader.load('a'), isError('timeout'));
    assert.equal(cancelled, 1);
    assert.equal(loader.stats.activeLoads, 0);
  } finally { loader.dispose(); }
});
test('cache LRU and in-flight byte budget are bounded independently of callers', async () => {
  const content = { a: Uint8Array.of(1, 1, 1, 1), b: Uint8Array.of(2, 2, 2, 2), c: Uint8Array.of(3, 3, 3, 3) };
  const calls = { a: 0, b: 0, c: 0 };
  const loader = new VerifiedAssetLoader(manifest(Object.entries(content).map(([id, bytes]) => asset(id, bytes))), options({
    maxAssetBytes: 8, maxResidentBytes: 8,
    fetch: async url => { const id = new URL(url).pathname.split('/').at(-1)[0]; calls[id]++; return response(content[id]); },
  }));
  try {
    await loader.load('a'); await loader.load('b'); await loader.load('a'); await loader.load('c');
    assert.equal(loader.stats.cachedBytes, 8);
    assert.deepEqual(calls, { a: 1, b: 1, c: 1 });
    await loader.load('a'); assert.equal(calls.a, 1);
    await loader.load('b'); assert.equal(calls.b, 2);
    assert.ok(loader.stats.cachedBytes <= 8);
  } finally { loader.dispose(); }
});
test('concurrency limit rejects excess requests without cancelling accepted ones', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  let resolveFetch;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes), asset('b', bytes)]), options({ maxConcurrentLoads: 1,
    fetch: () => new Promise(resolve => { resolveFetch = resolve; }) }));
  try {
    const first = loader.load('a');
    await assert.rejects(loader.load('b'), isError('resourceLimit'));
    resolveFetch(response(bytes));
    assert.deepEqual((await first).bytes, bytes);
    assert.equal(loader.stats.activeLoads, 0);
  } finally { loader.dispose(); }
});
test('immediate retry after the last waiter cancels starts a fresh request', async () => {
  const bytes = Uint8Array.of(1, 2, 3);
  let calls = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: (_url, init) => {
    calls++;
    if (calls > 1) return Promise.resolve(response(bytes));
    return new Promise((_resolve, reject) => {
      init.signal.addEventListener('abort', () => reject(new DOMException('cancelled', 'AbortError')), { once: true });
    });
  }}));
  try {
    const controller = new AbortController();
    const first = loader.load('a', controller.signal);
    const cancelled = assert.rejects(first, isError('cancelled'));
    controller.abort();
    const retry = loader.load('a');
    await Promise.all([cancelled, retry.then(result => assert.deepEqual(result.bytes, bytes))]);
    assert.equal(calls, 2);
    assert.equal(loader.stats.pendingBytes, 0);
  } finally { loader.dispose(); }
});
test('double-encoded traversal is rejected instead of trusting server decode count', () => {
  const bytes = Uint8Array.of(1);
  for (const url of ['%252e%252e/secret.otf', 'sub/%252f..%252fsecret.otf', 'a%255cb.otf']) {
    assert.throws(() => new VerifiedAssetLoader(manifest([asset('a', bytes, { url })]), options()), isError('unsafeUrl'));
  }
});
test('unknown and already-cancelled requests do not issue any fetch', async () => {
  const bytes = Uint8Array.of(1);
  let calls = 0;
  const loader = new VerifiedAssetLoader(manifest([asset('a', bytes)]), options({ fetch: async () => { calls++; return response(bytes); } }));
  try {
    await assert.rejects(loader.load('unknown'), isError('unknownAsset'));
    const controller = new AbortController(); controller.abort();
    await assert.rejects(loader.load('a', controller.signal), isError('cancelled'));
    assert.equal(calls, 0);
  } finally { loader.dispose(); }
});

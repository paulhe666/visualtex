// Real, isolated Chromium/Edge host check. It deliberately does NOT connect to
// an existing browser/WebView/Word process and does not claim system-IME or
// visual-math-editor acceptance. Its own temporary browser is closed on exit.
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { createHash } from 'node:crypto';
import { existsSync, readFileSync, mkdtempSync, rmSync } from 'node:fs';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const browserPath = [process.env.VISUALTEX_ENGINE_REVIEW_BROWSER,
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Microsoft/Edge/Application/msedge.exe',
  'C:/Program Files/Google/Chrome/Application/chrome.exe'].find(path => path && existsSync(path));
assert.ok(browserPath, 'No supported browser executable was found; no existing browser will be attached');
const wasm = readFileSync(resolve(root, 'web/pkg/visualtex_engine_wasm_bg.wasm'));
const font = readFileSync(resolve(root, '../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf'));
assert.equal(sha(font), '6075562b771f8b82f0c179e363389684f2dd09de30038269e2628e504bd7be0f', 'read-only font fixture changed');
const buildId = `review-${sha(wasm).slice(0, 16)}`;
const resources = new Map([
  ['/modules/resources.js', { type: 'text/javascript', bytes: readFileSync(resolve(root, 'target/host-browser/modules/resources.js')) }],
  ['/modules/coordinates.js', { type: 'text/javascript', bytes: readFileSync(resolve(root, 'target/host-browser/modules/coordinates.js')) }],
  ['/assets/core.js', { type: 'text/javascript', bytes: readFileSync(resolve(root, 'web/pkg/visualtex_engine_wasm.js')) }],
  ['/assets/core.wasm', { type: 'application/wasm', bytes: wasm }],
  ['/assets/math.otf', { type: 'font/otf', bytes: font }],
  ['/assets/bad.wasm', { type: 'text/html', bytes: wasm }],
]);
const manifest = { schemaVersion: 1, engineProtocol: 1, buildId, assets: [
  { id: 'core', kind: 'wasm', url: 'core.wasm', sha256: sha(wasm), byteLength: wasm.length },
  { id: 'math', kind: 'font', url: 'math.otf', sha256: sha(font), byteLength: font.length },
  { id: 'bad', kind: 'wasm', url: 'bad.wasm', sha256: sha(wasm), byteLength: wasm.length },
] };
const requests = new Map();
const server = createServer((req, res) => {
  const path = new URL(req.url, 'http://127.0.0.1').pathname;
  requests.set(path, (requests.get(path) ?? 0) + 1);
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  if (path === '/') {
    res.setHeader('Content-Type', 'text/html; charset=utf-8');
    res.setHeader('Content-Security-Policy', "default-src 'none'; script-src 'self' 'wasm-unsafe-eval'; connect-src 'self'");
    res.end('<!doctype html><meta charset="utf-8"><title>VisualTeX engine host verification</title><p>Isolated engine host regression.</p>');
  } else if (path === '/manifest.json') {
    res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(manifest));
  } else if (resources.has(path)) {
    const resource = resources.get(path);
    res.setHeader('Content-Type', resource.type); res.setHeader('Content-Length', resource.bytes.length); res.end(resource.bytes);
  } else { res.statusCode = 404; res.end('Not found'); }
});
await new Promise((resolveListen, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolveListen); });
const baseUrl = `http://127.0.0.1:${server.address().port}`;
const profile = mkdtempSync(join(tmpdir(), 'visualtex-engine-review-browser-'));
let child;
let client;
let stderr = '';

class CdpClient {
  constructor(url) { this.url = url; this.nextId = 1; this.pending = new Map(); this.events = new Set(); }
  async connect() {
    this.socket = new WebSocket(this.url);
    await new Promise((ok, bad) => {
      const timer = setTimeout(() => bad(new Error('Owned browser CDP connection timed out')), 15_000);
      this.socket.addEventListener('open', () => { clearTimeout(timer); ok(); }, { once: true });
      this.socket.addEventListener('error', () => { clearTimeout(timer); bad(new Error('Owned browser CDP connection failed')); }, { once: true });
    });
    this.socket.addEventListener('message', event => {
      const message = JSON.parse(event.data);
      if (message.id) {
        const waiting = this.pending.get(message.id);
        if (!waiting) return;
        this.pending.delete(message.id); clearTimeout(waiting.timer);
        if (message.error) waiting.reject(new Error(message.error.message)); else waiting.resolve(message.result);
      } else {
        for (const waiting of this.events) {
          if (waiting.method === message.method && waiting.sessionId === message.sessionId) {
            this.events.delete(waiting); clearTimeout(waiting.timer); waiting.resolve(message.params);
          }
        }
      }
    });
    this.socket.addEventListener('close', () => this.rejectAll(new Error('Owned browser CDP closed')));
  }
  rejectAll(error) {
    for (const waiting of [...this.pending.values(), ...this.events]) { clearTimeout(waiting.timer); waiting.reject(error); }
    this.pending.clear(); this.events.clear();
  }
  send(method, params = {}, sessionId) {
    const id = this.nextId++;
    return new Promise((resolveRequest, reject) => {
      const timer = setTimeout(() => { this.pending.delete(id); reject(new Error(`CDP timeout: ${method}`)); }, 20_000);
      this.pending.set(id, { resolve: resolveRequest, reject, timer });
      this.socket.send(JSON.stringify({ id, method, params, sessionId }));
    });
  }
  wait(method, sessionId) {
    return new Promise((resolveEvent, reject) => {
      const waiting = { method, sessionId, resolve: resolveEvent, reject };
      waiting.timer = setTimeout(() => { this.events.delete(waiting); reject(new Error(`CDP event timeout: ${method}`)); }, 15_000);
      this.events.add(waiting);
    });
  }
  close() { this.rejectAll(new Error('Owned test connection disposed')); this.socket?.close(); }
}

try {
  child = spawn(browserPath, ['--headless=new', '--remote-debugging-port=0', '--remote-debugging-address=127.0.0.1',
    '--no-first-run', '--no-default-browser-check', '--disable-background-networking', `--user-data-dir=${profile}`, 'about:blank'],
  { stdio: ['ignore', 'ignore', 'pipe'], windowsHide: true });
  const endpoint = await new Promise((ok, bad) => {
    const timer = setTimeout(() => bad(new Error(`Owned browser startup timed out: ${stderr.slice(-1500)}`)), 20_000);
    child.once('error', error => { clearTimeout(timer); bad(error); });
    child.once('exit', code => { clearTimeout(timer); bad(new Error(`Owned browser exited ${code}: ${stderr.slice(-1500)}`)); });
    child.stderr.on('data', chunk => {
      stderr = (stderr + chunk.toString()).slice(-16_000);
      const match = stderr.match(/DevTools listening on (ws:\/\/127\.0\.0\.1:\d+\/devtools\/browser\/[^\s]+)/u);
      if (match) { clearTimeout(timer); ok(match[1]); }
    });
  });
  client = new CdpClient(endpoint); await client.connect();
  const browserVersion = await client.send('Browser.getVersion');
  const { targetId } = await client.send('Target.createTarget', { url: 'about:blank' });
  const { sessionId } = await client.send('Target.attachToTarget', { targetId, flatten: true });
  await client.send('Page.enable', {}, sessionId);
  const loaded = client.wait('Page.loadEventFired', sessionId);
  await client.send('Page.navigate', { url: baseUrl }, sessionId); await loaded;
  const evaluate = async expression => {
    const result = await client.send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true }, sessionId);
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description ?? result.exceptionDetails.text);
    return result.result.value;
  };
  const result = await evaluate(`(async () => {
    const { VerifiedAssetLoader } = await import('/modules/resources.js');
    const manifest = await (await fetch('/manifest.json')).json();
    const loader = new VerifiedAssetLoader(manifest, { baseUrl: new URL('/assets/', location.href).href,
      expectedBuildId: ${JSON.stringify(buildId)}, expectedEngineProtocol: 1 });
    let engine;
    try {
      const [first, second, font] = await Promise.all([loader.load('core'), loader.load('core'), loader.load('math')]);
      const { default: init, EngineRuntime } = await import('/assets/core.js');
      await init({ module_or_path: first.bytes });
      if (first.bytes.buffer === second.bytes.buffer) throw new Error('Resource buffers were shared with callers');
      first.bytes[0] = 99;
      if ((await loader.load('core')).bytes[0] !== 0) throw new Error('Caller mutated cached WASM bytes');
      engine = new EngineRuntime();
      const call = command => JSON.parse(engine.dispatch(JSON.stringify({ protocol: 1, command })));
      const created = call({ op: 'create', documentId: 'browser-host-review' });
      if (created.status !== 'ok') throw new Error('Creation failed');
      const handle = created.result.handle;
      const source = ${JSON.stringify('\\frac{x_1+\\text{速度}}{y^2}')};
      const opened = call({ op: 'replaceSource', handle, expectedRevision: 0, source });
      if (opened.status !== 'ok') throw new Error(JSON.stringify(opened));
      const denominatorId = opened.result.snapshot.analysis.tree.entries[0].node.denominator.id;
      const at = source.indexOf('速度');
      const edited = call({ op: 'patchSource', handle, expectedRevision: 1,
        edit: { startUtf16: at, endUtf16: at + 2, replacement: '速率😀' } });
      if (edited.status !== 'ok' || edited.result.mapping.parseScope.kind !== 'subtree') throw new Error(JSON.stringify(edited));
      if (edited.result.snapshot.analysis.tree.entries[0].node.denominator.id !== denominatorId) throw new Error('Unrelated denominator identity changed');
      const stale = call({ op: 'patchSource', handle, expectedRevision: 1,
        edit: { startUtf16: 0, endUtf16: 0, replacement: 'bad' } });
      if (stale.status !== 'error' || stale.error.code !== 'staleRevision') throw new Error('Stale edit accepted');
      const bad = await loader.load('bad').then(() => 'accepted', error => error.code);
      if (bad !== 'badMime') throw new Error('Bad resource MIME accepted');
      if (call({ op: 'destroy', handle, expectedRevision: 2 }).status !== 'ok' || engine.activeSessions !== 0) throw new Error('Session leak');
      return { secureContext: isSecureContext, actualWasmExecution: true, source: edited.result.snapshot.source,
        scope: edited.result.mapping.parseScope, denominatorIdentityPreserved: true, realFontBytes: font.bytes.byteLength,
        fontSha256: font.sha256, badMimeRejected: true, cache: loader.stats, crossOriginIsolated };
    } finally { engine?.free(); loader.dispose(); }
  })()`);
  assert.equal(result.secureContext, true);
  assert.equal(result.source, '\\frac{x_1+\\text{速率😀}}{y^2}');
  const dpiResults = [];
  for (const deviceScaleFactor of [1, 1.25, 1.5, 2]) {
    await client.send('Emulation.setDeviceMetricsOverride', { width: 900, height: 600, deviceScaleFactor, mobile: false }, sessionId);
    const dpi = await evaluate(`(async () => {
      const { ViewportMapping } = await import('/modules/coordinates.js');
      const mapping = new ViewportMapping({ left: 100, top: 200, scrollLeft: 10, scrollTop: 20, zoom: 2 });
      return { devicePixelRatio, point: mapping.clientToLayout({ x: 110, y: 220 }) };
    })()`);
    assert.deepEqual(dpi.point, { x: 10, y: 20 }); assert.equal(dpi.devicePixelRatio, deviceScaleFactor); dpiResults.push(dpi);
  }
  // These are browser-protocol composition traces, NOT physical Windows IME.
  // Keep the distinction even though Chromium labels these DOM events trusted.
  await evaluate(`(() => {
    const sink = document.createElement('textarea'); sink.id = 'owned-ime-probe';
    sink.spellcheck = false; sink.autocomplete = 'off'; document.body.append(sink);
    window.__reviewImeEvents = [];
    for (const type of ['keydown', 'compositionstart', 'compositionupdate', 'beforeinput', 'input', 'compositionend', 'keyup']) {
      sink.addEventListener(type, event => window.__reviewImeEvents.push({ type: event.type,
        data: event.data ?? null, inputType: event.inputType ?? null, composing: event.isComposing ?? null,
        trusted: event.isTrusted, cancelable: event.cancelable, value: sink.value,
        selectionStart: sink.selectionStart, selectionEnd: sink.selectionEnd }));
    }
    sink.focus(); return true;
  })()`);
  await client.send('Input.imeSetComposition', { text: 'su', selectionStart: 2, selectionEnd: 2 }, sessionId);
  await client.send('Input.imeSetComposition', { text: '速度', selectionStart: 2, selectionEnd: 2 }, sessionId);
  await client.send('Input.insertText', { text: '速度' }, sessionId);
  await client.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'x', code: 'KeyX', text: 'x',
    unmodifiedText: 'x', windowsVirtualKeyCode: 88, nativeVirtualKeyCode: 88 }, sessionId);
  await client.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'x', code: 'KeyX',
    windowsVirtualKeyCode: 88, nativeVirtualKeyCode: 88 }, sessionId);
  const committedTrace = await evaluate(`({ value: document.getElementById('owned-ime-probe').value, events: window.__reviewImeEvents })`);
  assert.equal(committedTrace.value, '速度x', JSON.stringify(committedTrace));
  assert.equal(committedTrace.events.filter(event => event.type === 'compositionstart').length, 1);
  assert.equal(committedTrace.events.filter(event => event.type === 'compositionend').length, 1);
  await evaluate(`(() => { const sink = document.getElementById('owned-ime-probe'); sink.value = '';
    sink.setSelectionRange(0, 0); sink.focus(); window.__reviewImeEvents = []; return true; })()`);
  await client.send('Input.imeSetComposition', { text: 'temporary', selectionStart: 9, selectionEnd: 9 }, sessionId);
  await client.send('Input.imeSetComposition', { text: '', selectionStart: 0, selectionEnd: 0 }, sessionId);
  const cancelledTrace = await evaluate(`({ value: document.getElementById('owned-ime-probe').value, events: window.__reviewImeEvents })`);
  assert.equal(cancelledTrace.value, '', JSON.stringify(cancelledTrace));
  const protocolComposition = { mechanism: 'CDP Input.imeSetComposition / Input.insertText, not OS IME',
    committed: committedTrace, cancelled: cancelledTrace };
  assert.equal(requests.get('/assets/core.wasm'), 1);
  assert.equal(requests.get('/assets/math.otf'), 1);
  console.log(JSON.stringify({ status: 'PASS', browser: browserVersion.product, userAgent: browserVersion.userAgent,
    isolatedOwnedProcess: true, existingProcessesAttached: false, wasmSha256: sha(wasm), result, dpiResults, protocolComposition,
    networkRequests: Object.fromEntries(requests),
    notAccepted: ['visual mathematical editing', 'font shaping/layout', 'real system IME', 'WebView2 embedding', 'Office UI'] }, null, 2));
} finally {
  if (client) { try { await client.send('Browser.close'); } catch { /* own process may close before ACK */ } client.close(); }
  if (child && child.exitCode === null) {
    const closed = await Promise.race([new Promise(ok => child.once('exit', () => ok(true))),
      new Promise(ok => setTimeout(() => ok(false), 3000))]);
    if (!closed) child.kill(); // only the process spawned above, never process-name kill
  }
  server.closeAllConnections(); await new Promise(ok => server.close(ok));
  // This path was created by mkdtemp in this invocation, not user data.
  try { rmSync(profile, { recursive: true, force: true, maxRetries: 3, retryDelay: 100 }); } catch { /* preserve owned evidence if a browser handle lingers */ }
}

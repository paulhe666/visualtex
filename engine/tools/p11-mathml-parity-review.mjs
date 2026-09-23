import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const require = createRequire(import.meta.url);
const { EngineRuntime } = require(resolve(root, 'web/pkg-node/visualtex_engine_wasm.js'));
const nativePath = resolve(root, `target/debug/visualtex-engine-test-cli${process.platform === 'win32' ? '.exe' : ''}`);
const wasmPath = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const request = (command) => JSON.stringify({ protocol: 1, command });

const customSymbol = {
  id: 'p11-symbol', command: 'pcompatible', name: 'P11 compatible', role: 'ordinary', limitsBehavior: 'auto',
  metrics: { widthEm: 0.8, ascentEm: 0.7, descentEm: 0.2 },
  artwork: { shapes: [{ kind: 'path', d: 'M0 0 L500 0 L500 500 Z' }] },
  ommlFallback: '\\alpha', createdAt: 1, updatedAt: 1,
};

const corpus = [
  String.raw`x+y=1`,
  String.raw`\frac{x_1+\text{速度}}{y^2}`,
  String.raw`\sqrt[3]{x^2+y^2}`,
  String.raw`\left(\frac{a}{b}\right)^2`,
  String.raw`\sum\limits_{i=0}^{n} i^2`,
  String.raw`\int\nolimits_0^1 x\,dx`,
  String.raw`\hat{\vec{x}}+\overline{AB}`,
  String.raw`\mathbf{x}+\mathbb{R}+\text{中文}`,
  String.raw`\begin{matrix}a&b\\c&d\end{matrix}`,
  String.raw`\begin{pmatrix}a&b\\c&d\end{pmatrix}`,
  String.raw`\begin{cases}x&x>0\\-x&x\le0\end{cases}`,
  String.raw`\begin{aligned}a&=b&&+c\\xx&=y&&+zz\end{aligned}`,
  String.raw`\begin{array}{r|l}a&b\\cc&d\end{array}`,
  String.raw`\dv{x^2}{x}+\pdv{f}{y}+\braket{a}{b}`,
  String.raw`\color{red}{x}+\boxed{y}`,
];

const calls = [];
let handle = 0;
for (const source of corpus) {
  const h = ++handle;
  calls.push(request({ op: 'create', documentId: `p11-mathml-${h}`, source }));
  calls.push(request({ op: 'exportPresentationMathMl', handle: h, expectedRevision: 0, displayMode: true }));
  calls.push(request({ op: 'exportPresentationMathMl', handle: h, expectedRevision: 0, displayMode: false }));
  calls.push(request({ op: 'destroy', handle: h, expectedRevision: 0 }));
}
{
  const h = ++handle;
  calls.push(request({ op: 'createWithSymbols', documentId: `p11-mathml-${h}`, source: String.raw`\pcompatible+x`, symbols: [customSymbol] }));
  calls.push(request({ op: 'exportPresentationMathMl', handle: h, expectedRevision: 0, displayMode: true }));
  calls.push(request({ op: 'destroy', handle: h, expectedRevision: 0 }));
}
for (const source of [String.raw`\frac{x}{`, String.raw`\future{x}`]) {
  const h = ++handle;
  calls.push(request({ op: 'create', documentId: `p11-negative-${h}`, source }));
  calls.push(request({ op: 'exportPresentationMathMl', handle: h, expectedRevision: 0, displayMode: true }));
  calls.push(request({ op: 'destroy', handle: h, expectedRevision: 0 }));
}

const native = spawnSync(nativePath, [], {
  input: calls.join('\n') + '\n', encoding: 'utf8', timeout: 60_000, maxBuffer: 64 * 1024 * 1024,
});
assert.equal(native.error, undefined);
assert.equal(native.status, 0, native.stderr);
const nativeRows = native.stdout.trim().split(/\r?\n/u).map((line) => JSON.parse(line));
assert.equal(nativeRows.length, calls.length);

const runtime = new EngineRuntime();
const transcript = createHash('sha256');
let exports = 0;
let errors = 0;
try {
  for (let index = 0; index < calls.length; index += 1) {
    const wasm = JSON.parse(runtime.dispatch(calls[index]));
    assert.deepEqual(wasm, nativeRows[index], `Native/WASM divergence at request ${index}`);
    transcript.update(JSON.stringify(wasm));
    if (wasm.status === 'ok' && wasm.result?.kind === 'presentationMathMl') {
      exports += 1;
      const value = wasm.result.result.mathMl;
      assert.match(value, /^<math xmlns="http:\/\/www\.w3\.org\/1998\/Math\/MathML" display="(?:block|inline)">/u);
      assert.ok(!value.includes('data-visualtex-grid-id='), 'private core grid identity must not escape the protocol');
      assert.ok(!value.includes('<script') && !value.includes('<foreignObject'));
    }
    if (wasm.status === 'error') {
      errors += 1;
      assert.ok(['incompleteSource', 'unsupportedSyntax'].includes(wasm.error.code));
    }
  }
  assert.equal(runtime.activeSessions, 0);
} finally {
  runtime.free();
}
assert.equal(exports, corpus.length * 2 + 1);
assert.equal(errors, 2);

console.log(JSON.stringify({
  status: 'PASS', actualNativeExecution: true, actualWasmExecution: true,
  corpus: corpus.length, presentationMathMlExports: exports, explicitFailures: errors,
  nativeSha256: sha256(readFileSync(nativePath)), wasmSha256: sha256(readFileSync(wasmPath)),
  transcriptSha256: transcript.digest('hex'),
  scope: 'P11 retained-tree Presentation MathML parity; no MathJax, Office Session, Word or UI calls.',
}, null, 2));

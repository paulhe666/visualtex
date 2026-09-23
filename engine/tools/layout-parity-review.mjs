// Execute the same font/layout requests in the native host and real WASM.
// The font stays local; neither this runner nor source archives distribute it.
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
const sha = bytes => createHash('sha256').update(bytes).digest('hex');
const fontPath = process.env.VISUALTEX_TEST_MATH_FONT || resolve(root, '../apps/windows/src-windows/artifacts/office-fonts/latinmodern-math.otf');
const font = readFileSync(fontPath);
const fontId = `${sha(font)}:0`;
const nativePath = resolve(root, `target/debug/visualtex-engine-test-cli${process.platform==='win32'?'.exe':''}`);
const wasmPath = resolve(root, 'web/pkg-node/visualtex_engine_wasm_bg.wasm');
const calls = [];
let handle = 0;
const request = command => JSON.stringify({ protocol: 1, command });
const add = (command, check = () => {}) => calls.push({input:request(command),check});
const options = (style='display', fontSizePx=16) => ({ mathFont:fontId, textFont:null, style, fontSizePx, paddingEm:0.125 });
let scenes = 0;
function compareNumbers(a,b,path='reply') {
  if(typeof a==='number' && typeof b==='number') {
    assert.ok(Number.isFinite(a)&&Number.isFinite(b),path);
    assert.ok(Math.abs(a-b)<=1e-9*Math.max(1,Math.abs(a),Math.abs(b)),`${path}: ${a} != ${b}`); return;
  }
  if(Array.isArray(a)||Array.isArray(b)) {
    assert.ok(Array.isArray(a)&&Array.isArray(b),path); assert.equal(a.length,b.length,path);
    a.forEach((v,i)=>compareNumbers(v,b[i],`${path}[${i}]`)); return;
  }
  if(a && b && typeof a==='object' && typeof b==='object') {
    assert.deepEqual(Object.keys(a).sort(),Object.keys(b).sort(),path);
    for(const k of Object.keys(a)) compareNumbers(a[k],b[k],`${path}.${k}`); return;
  }
  assert.equal(a,b,path);
}
const corpus = ['x+y=1',String.raw`\frac{x_1+1}{y^2}`,String.raw`\frac{\dfrac{a}{b}}{\tfrac{c}{d}}`,
  String.raw`x^{y^{z^2}}`,String.raw`\frac{f^{g_y}}{j_2}`,String.raw`\alpha+\beta=\gamma`,
  String.raw`\text{office AV}+x`, '\\text{Cafe\u0301}+1',String.raw`\frac{x}{2}+\frac{x}{2}`];
let seed=0x50505226;
for(let i=0;i<36;i++) {
  seed=(Math.imul(seed,1664525)+1013904223)>>>0;
  corpus.push(seed%2?String.raw`\frac{x_{${i+1}}+y}{z^{${i+2}}}`:String.raw`{\frac{a+b}{c}}^{${i+1}}+\text{AV}`);
}
for(const source of corpus) {
  const h=++handle;
  add({op:'create',documentId:`layout-review-${h}`,source},r=>assert.equal(r.result.handle,h));
  for(const style of ['display','text','script','scriptScript']) {
    let previous;
    for(const size of [16,32]) {
      add({op:'layout',handle:h,expectedRevision:0,options:options(style,size),includeSvg:true},r=>{
        assert.equal(r.status,'ok',JSON.stringify(r));
        const s=r.result.snapshot; scenes++;
        assert.equal(s.documentId,`layout-review-${h}`); assert.equal(s.revision,0);
        assert.ok(s.exportReady); assert.ok(s.width>0&&s.height>0&&s.baseline>=0&&s.baseline<=s.height);
        assert.ok(r.result.svg.startsWith('<svg '));
        for(const bad of ['NaN','Infinity','<text','<image','<script','foreignObject','href=']) assert.ok(!r.result.svg.includes(bad),bad);
        if(previous) {
          assert.equal(s.rebuiltFragments,0);
          compareNumbers(s.width,2*previous.width); compareNumbers(s.height,2*previous.height);
          compareNumbers(s.baseline,2*previous.baseline);
          compareNumbers(s.glyphs,previous.glyphs); compareNumbers(s.carets,previous.carets);
        }
        previous=s;
      });
    }
  }
  add({op:'layout',handle:h,expectedRevision:1,options:options(),includeSvg:true},r=>assert.equal(r.error.code,'staleRevision'));
  add({op:'destroy',handle:h,expectedRevision:0},r=>assert.equal(r.result.kind,'destroyed'));
}
for(const source of ['', '{}', String.raw`\frac{{}}{x}`,String.raw`\text{}`,String.raw`\frac{x}{`,String.raw`\future{x}`]) {
  const h=++handle; add({op:'create',documentId:`negative-${h}`,source});
  add({op:'layout',handle:h,expectedRevision:0,options:options(),includeSvg:true},r=>assert.equal(r.status,'error',source));
  add({op:'inspect',handle:h},r=>assert.equal(r.result.snapshot.source,source));
  add({op:'destroy',handle:h,expectedRevision:0});
}
const native = spawnSync(nativePath,['--font',fontPath],{input:calls.map(c=>c.input).join('\n')+'\n',encoding:'utf8',timeout:60000,maxBuffer:128*1024*1024});
assert.equal(native.error,undefined); assert.equal(native.status,0,native.stderr);
const rows=native.stdout.trim().split(/\r?\n/).map(r=>JSON.parse(r)); assert.equal(rows.length,calls.length);
const runtime = new EngineRuntime();
const digest=createHash('sha256');
try {
  assert.equal(JSON.parse(runtime.registerFont(font,0)).result.font.id,fontId);
  for(let i=0;i<calls.length;i++) {
    const result=JSON.parse(runtime.dispatch(calls[i].input));
    compareNumbers(result,rows[i],`request ${i}`); calls[i].check(result);
    digest.update(JSON.stringify(result));
  }
  assert.equal(runtime.activeSessions,0);
} finally {runtime.free();}
console.log(JSON.stringify({status:'PASS',actualWasmExecution:true,requests:calls.length,validScenes:scenes,
  fontSha256:sha(font),nativeSha256:sha(readFileSync(nativePath)),wasmSha256:sha(readFileSync(wasmPath)),
  tolerance:'relative 1e-9 for numeric geometry; exact SVG and identities',transcriptSha256:digest.digest('hex'),
  scope:'Native/WASM K0 Latin/math geometry. Chinese font, browser events, OS IME and Office UI are not covered by this runner.'},null,2));

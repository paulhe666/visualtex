import { EngineEditorHost } from '../host/editor-host.js';
import { VerifiedAssetLoader } from '../host/resources.js';
import type { AssetManifest } from '../host/resources';
import type { EditorSnapshot } from '../host/editing';
import type { LayoutSnapshot } from '../host/layout';
import type { InputAudit } from '../host/input-bridge';

interface WasmRuntime {
  dispatch(request: string): string;
  registerFont(bytes: Uint8Array, faceIndex: number): string;
  readonly activeSessions: number;
  free(): void;
}
interface SandboxInspection {
  ready: Promise<void>;
  inspect(): EditorSnapshot | null;
  layout(): LayoutSnapshot | null;
  events(): InputAudit[];
  stats(): { sceneBuilds: number; overlayBuilds: number; zoom: number } | null;
  diagnostics(): string[];
  resourceBuild: string;
}
declare global { interface Window { __engineSandbox: SandboxInspection } }

function required<T extends HTMLElement>(selector: string): T {
  const value = document.querySelector<T>(selector);
  if (!value) throw new Error(`Missing sandbox control: ${selector}`);
  return value;
}
const source = required<HTMLTextAreaElement>('#source');
const errorPanel = required<HTMLDivElement>('#error');
const mount = required<HTMLDivElement>('#editor');
const phase = required<HTMLSpanElement>('#phase');
const revision = required<HTMLSpanElement>('#revision');
const documentStatus = required<HTMLSpanElement>('#document-status');
const exportButton = required<HTMLButtonElement>('#export-svg');
const exportResult = required<HTMLPreElement>('#export-result');
const buildId = required<HTMLMetaElement>('meta[name="engine-build-id"]').content;
let host: EngineEditorHost | null = null;
let runtime: WasmRuntime | null = null;
let loader: VerifiedAssetLoader | null = null;
let disposed = false;
let sourceComposing = false;
let exportRevision: number | null = null;
let frame = 0;
const abort = new AbortController();
const diagnostics: string[] = [];
const expectedTitleToken = new URL(location.href).searchParams.get('run');
if (expectedTitleToken && /^[a-zA-Z0-9-]{8,80}$/u.test(expectedTitleToken)) {
  document.title = `VisualTeX engine sandbox ${expectedTitleToken}`;
}
document.body.dataset.ready = 'false';

function report(reason: unknown): void {
  const message = reason instanceof Error ? reason.message : String(reason);
  diagnostics.push(message);
  if (diagnostics.length > 128) diagnostics.shift();
  errorPanel.textContent = message;
  errorPanel.hidden = false;
}
function invalidateExport(): void {
  if (exportRevision === null) return;
  exportRevision = null;
  exportResult.textContent = '内容或数学样式已变化，请重新检查 SVG。';
}
function update(snapshot: EditorSnapshot): void {
  if (exportRevision !== null && exportRevision !== snapshot.document.revision) invalidateExport();
  if (source.value !== snapshot.document.source) {
    const start = source.selectionStart;
    const end = source.selectionEnd;
    source.value = snapshot.document.source;
    if (document.activeElement === source) source.setSelectionRange(Math.min(start,source.value.length),Math.min(end,source.value.length));
  }
  revision.textContent = `revision ${snapshot.document.revision} · undo ${snapshot.undoDepth} · redo ${snapshot.redoDepth}`;
  const analysis = snapshot.document.analysis;
  const available = analysis.complete && analysis.supported;
  const composing = snapshot.composition !== null;
  documentStatus.textContent = composing ? '输入法候选中' : analysis.exportable ? '结构完整' : available ? '待填写' : '保留源码草稿';
  exportButton.disabled = !analysis.exportable || composing;
  for (const button of document.querySelectorAll<HTMLButtonElement>('[data-structure],[data-symbol],[data-fence],[data-accent],[data-annotation],[data-style],[data-space],[data-phantom],#insert-boxed,[data-color],[data-operator],[data-general-operator],[data-limits],#math-style,#exit-structure,#focus-editor')) button.disabled = !available || composing;
  if (available) { errorPanel.hidden = true; errorPanel.textContent = ''; }
}
function safe(operation: () => void): void {
  try { operation(); }
  catch (error) {
    // Core may have accepted a source draft before rendering was refused. The
    // current source projection must remain visible and editable for recovery.
    if (host) update(host.inspect());
    report(error);
  }
}
function updateSource(): void {
  if (!host || sourceComposing) return;
  safe(() => host!.replaceSource(source.value));
}
source.addEventListener('compositionstart', () => { sourceComposing = true; }, { signal: abort.signal });
source.addEventListener('compositionend', () => { sourceComposing = false; updateSource(); }, { signal: abort.signal });
source.addEventListener('input', event => { if (!(event as InputEvent).isComposing) updateSource(); }, { signal: abort.signal });
source.addEventListener('keydown', event => {
  if (event.isComposing || sourceComposing || event.keyCode === 229 || !(event.ctrlKey || event.metaKey) || event.altKey) return;
  const key = event.key.toLowerCase();
  if (key === 'z' || key === 'y') {
    event.preventDefault();
    safe(() => host?.history(key === 'y' || event.shiftKey));
  }
}, { signal: abort.signal });
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-structure]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    const structure = button.dataset.structure;
    if (!['fraction','binomial','skewedFraction','superscript','subscript','text','squareRoot','indexedRoot'].includes(structure ?? '')) throw new Error('Unknown structure button');
    host!.applyStructure(structure as 'fraction'|'binomial'|'skewedFraction'|'superscript'|'subscript'|'text'|'squareRoot'|'indexedRoot');
  }), { signal: abort.signal });
}
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-space],[data-phantom],#insert-boxed')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    if (button.dataset.space) host!.applySpace(`\\${button.dataset.space}`);
    else if (button.dataset.phantom) host!.applyPhantom(`\\${button.dataset.phantom}`);
    else host!.applyBoxed();
  }), { signal: abort.signal });
}
const colors: Record<string, import('../host/layout').RgbColor> = {
  red: {r:255,g:0,b:0}, blue: {r:0,g:0,b:255}, black: {r:0,g:0,b:0},
};
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-color]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    const color=colors[button.dataset.color ?? ''];if(!color)throw new Error('Unknown color control');host!.applyColor(color);
  }), { signal: abort.signal });
}
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-symbol]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => host!.applySymbol(`\\${button.dataset.symbol}`)), { signal: abort.signal });
}
const fencePairs: Record<string, readonly [string,string]> = {
  parentheses: ['(',')'], brackets: ['[',']'], braces: ['\\{','\\}'],
  absolute: ['|','|'], leftOnly: ['\\{','.'], rightOnly: ['.',')'],
};
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-fence]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    const pair = fencePairs[button.dataset.fence ?? ''];
    if (!pair) throw new Error('Unknown fence button');
    host!.applyFence(pair[0],pair[1]);
  }), { signal: abort.signal });
}
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-accent],[data-annotation]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    if (button.dataset.accent) host!.applyAccent(`\\${button.dataset.accent}`);
    else if (button.dataset.annotation) host!.applyAnnotation(`\\${button.dataset.annotation}`);
  }), { signal: abort.signal });
}
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-style]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => host!.applyStyle(`\\${button.dataset.style}`)), { signal: abort.signal });
}
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-general-operator]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    const bounds = button.dataset.generalOperator === 'bounds';
    host!.applyGeneralOperator(bounds,bounds);
  }), { signal: abort.signal });
}
for (const button of document.querySelectorAll<HTMLButtonElement>('[data-operator],[data-limits]')) {
  button.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
  button.addEventListener('click', () => safe(() => {
    if (button.dataset.operator) host!.applyOperator(`\\${button.dataset.operator}`,true,button.dataset.upper !== 'false');
    else host!.setOperatorLimits(button.dataset.limits as import('../host/syntax').LimitMode);
  }), { signal: abort.signal });
}
required<HTMLSelectElement>('#math-style').addEventListener('change', event => safe(() => {
  host!.setMathStyle((event.target as HTMLSelectElement).value as import('../host/layout').MathStyle);
  invalidateExport();
}), { signal: abort.signal });
const exitButton = required<HTMLButtonElement>('#exit-structure');
exitButton.addEventListener('pointerdown', event => { if (event.button === 0) event.preventDefault(); }, { signal: abort.signal });
exitButton.addEventListener('click', () => safe(() => host!.exitStructure()), { signal: abort.signal });
required<HTMLButtonElement>('#focus-editor').addEventListener('click', () => host?.focus(), { signal: abort.signal });
required<HTMLSelectElement>('#zoom').addEventListener('change', event => safe(() => {
  host?.setZoom(Number((event.target as HTMLSelectElement).value));
}), { signal: abort.signal });
exportButton.addEventListener('click', () => safe(() => {
  const svg = host!.exportSvg();
  const layout = host!.layoutSnapshot;
  exportResult.textContent = `${layout?.width.toFixed(3)} × ${layout?.height.toFixed(3)} CSS px；基线 ${layout?.baseline.toFixed(3)} px\n${svg}`;
  exportRevision = layout?.revision ?? null;
}), { signal: abort.signal });

async function boot(): Promise<void> {
  const manifestResponse = await fetch('/manifest.json', { cache: 'no-store', signal: abort.signal });
  if (!manifestResponse.ok) throw new Error('引擎资源清单读取失败');
  const manifest = await manifestResponse.json() as AssetManifest;
  loader = new VerifiedAssetLoader(manifest, {
    baseUrl: new URL('/assets/', location.href).href,
    expectedBuildId: buildId,
    expectedEngineProtocol: 1,
  });
  const [wasm, math, text] = await Promise.all(['core','math','text'].map(id => loader!.load(id)));
  const mathFallback = manifest.assets.some(asset => asset.id === 'mathFallback') ? await loader.load('mathFallback') : null;
  if (disposed) throw new Error('Sandbox disposed during resource loading');
  const bindingUrl = '/assets/core.js';
  const bindings = await import(bindingUrl) as {
    default(options: { module_or_path: Uint8Array }): Promise<unknown>;
    EngineRuntime: new () => WasmRuntime;
  };
  await bindings.default({ module_or_path: wasm.bytes });
  runtime = new bindings.EngineRuntime();
  const fontIds = [math,text,...(mathFallback ? [mathFallback] : [])].map(font => {
    const response = JSON.parse(runtime!.registerFont(font.bytes,0));
    if (response.status !== 'ok' || response.result?.kind !== 'fontRegistered' || response.result.font.sha256 !== font.sha256) {
      throw new Error(`字体注册失败：${JSON.stringify(response)}`);
    }
    return response.result.font.id as string;
  });
  host = new EngineEditorHost(runtime,mount,{
    documentId: `sandbox-${crypto.randomUUID()}`,
    initialSource: '',
    layout: {mathFont:fontIds[0],textFont:fontIds[1],mathFallbacks:fontIds.slice(2),style:'display',fontSizePx:28,paddingEm:0.125},
    onChange:update,
    onError:report,
  });
  source.disabled = false;
  required<HTMLSelectElement>('#zoom').disabled = false;
  phase.textContent = 'K0 · 独立 WASM';
  required<HTMLSpanElement>('#resource-status').textContent = `build ${buildId} · 数学/CJK 字体已核验`;
  document.body.dataset.ready = 'true';
  update(host.inspect());
  host.focus();
  // Blink changes only overlay visibility. It never requests parse or layout.
  let previous = -1;
  const blink = (time: number) => {
    if (disposed) return;
    const tick = Math.floor(time/530);
    if (tick !== previous) {
      previous = tick;
      host?.view.setCaretVisible(document.activeElement === host?.input.textarea && tick%2 === 0);
    }
    frame = requestAnimationFrame(blink);
  };
  frame = requestAnimationFrame(blink);
}
function dispose(): void {
  if (disposed) return;
  disposed = true;
  abort.abort();
  cancelAnimationFrame(frame);
  try { host?.dispose(); } finally { runtime?.free(); loader?.dispose(); }
  host = null; runtime = null; loader = null;
}
window.addEventListener('pagehide',dispose,{once:true});
const ready = boot().catch(error => {
  phase.textContent = '加载失败 · 未切换到旧引擎';
  report(error);
  throw error;
});
window.__engineSandbox = {
  ready,
  inspect: () => host?.inspect() ?? null,
  layout: () => host?.layoutSnapshot ?? null,
  events: () => host?.input.events ?? [],
  stats: () => host?.view.stats ?? null,
  diagnostics: () => diagnostics.slice(),
  resourceBuild: buildId,
};
// A visible error is already presented; keep unhandled rejection logging from
// replacing it while tests still await the original rejected ready promise.
void ready.catch(() => undefined);

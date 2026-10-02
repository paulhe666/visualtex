// Builds src/landing/art.generated.ts: the hero brush lettering, every foreground string as brush
// lettering, and the background notebook of formulas typeset by MathJax.
// Run after editing copy.mjs or notebook.mjs:  npm run build:landing-art
import { readFileSync, writeFileSync } from "node:fs";
import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { loadFonts, pathD } from "./fonts.mjs";
import { copy } from "./copy.mjs";
import { notes, formulas } from "./notebook.mjs";

const VERSION = readFileSync(new URL("../../src/landing/LandingPage.tsx", import.meta.url), "utf8")
  .match(/const VERSION = "([^"]+)"/)[1];
const fonts = await loadFonts();
const esc = (s) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");

// ---------------------------------------------------------------- MathJax
const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);
const svgJax = new SVG({ fontCache: "global" });
const mj = mathjax.document("", { InputJax: new TeX({ packages: AllPackages }), OutputJax: svgJax });
function tex(src) {
  const out = adaptor.innerHTML(mj.convert(src, { display: true }));
  if (out.includes("merror")) throw new Error("MathJax error in: " + src);
  return out;
}

// ---------------------------------------------------------------- brush text
const isHan = (c) => /[⺀-鿿　-〿＀-￯]/.test(c);
const SIZE = { han: 100, body: 104, display: 112 };
const symbols = new Map(); // one <symbol> per distinct glyph
function symbolFor(fontKey, glyph) {
  const key = `${fontKey}-${glyph.index}`;
  if (!symbols.has(key)) {
    symbols.set(key, { id: `g${symbols.size.toString(36)}`, d: pathD(glyph.getPath(0, 0, SIZE[fontKey])) });
  }
  return symbols.get(key).id;
}
function glyphFor(ch, display) {
  const order = isHan(ch) ? ["han", "body"] : display ? ["display", "han"] : ["body", "han"];
  for (const key of order) {
    const glyph = fonts[key].charToGlyph(ch);
    if (glyph && glyph.index !== 0) return { key, glyph };
  }
  return null;
}
const BOX_TOP = -112, BOX_H = 152;
function brush(text, display = false) {
  let x = 0, prev = null;
  const uses = [];
  for (const ch of text) {
    if (ch === " ") { x += display ? 30 : 26; prev = null; continue; }
    const hit = glyphFor(ch, display);
    if (!hit) { x += 30; prev = null; continue; }
    const font = fonts[hit.key], scale = SIZE[hit.key] / font.unitsPerEm;
    if (prev && prev.key === hit.key) x += font.getKerningValue(prev.glyph, hit.glyph) * scale;
    uses.push(`<use href="#${symbolFor(hit.key, hit.glyph)}" x="${+x.toFixed(1)}"/>`);
    x += hit.glyph.advanceWidth * scale + (hit.key === "body" ? 1.5 : 0);
    prev = hit;
  }
  return `<svg class="ink" viewBox="-10 ${BOX_TOP} ${(x + 20).toFixed(1)} ${BOX_H}" role="img" aria-label="${esc(text)}"><g class="ink-g" filter="url(#vt-ink-small)">${uses.join("")}</g></svg>`;
}

const text = {};
for (const [key, entry] of Object.entries(copy)) {
  const fill = (s) => s.replaceAll("{VERSION}", VERSION);
  text[key] = {
    zh: brush(fill(entry.zh), entry.display),
    en: entry.enTex ? `<span class="ink-tex">${tex(entry.enTex)}</span>` : brush(fill(entry.en), entry.display),
  };
}
const marks = {
  eq: [1, 2, 3, 4].map((n) => tex(String.raw`(${n})`)),
  roman: ["i", "ii", "iii", "iv", "v"].map((r) => tex(String.raw`\text{(${r})}`)),
};

// ---------------------------------------------------------------- hero "VisualTeX"
const HERO = 300;
const word = fonts.word;
let cursor = 0, prevGlyph = null;
const placed = [..."VisualTeX"].map((ch) => {
  const glyph = word.charToGlyph(ch);
  if (prevGlyph) cursor += word.getKerningValue(prevGlyph, glyph) * (HERO / word.unitsPerEm);
  const item = { glyph, x: cursor };
  cursor += glyph.advanceWidth * (HERO / word.unitsPerEm);
  prevGlyph = glyph;
  return item;
});
const wordWidth = cursor;
// small hand variation per letter: [rotation deg, scale, dy]
const jitter = [[-2, 1.08, 6], [1, 0.98, -2], [-1, 1.0, 3], [2, 0.97, -3], [-1, 1.02, 2], [1, 1.0, -2], [-3, 1.06, 8], [1, 0.97, -4], [-2, 1.05, 4]];
const ink = { x1: Infinity, y1: Infinity, x2: -Infinity, y2: -Infinity };
const grow = (x, y) => { ink.x1 = Math.min(ink.x1, x); ink.y1 = Math.min(ink.y1, y); ink.x2 = Math.max(ink.x2, x); ink.y2 = Math.max(ink.y2, y); };
const letters = placed.map(({ glyph, x }, i) => {
  const [r, s, dy] = jitter[i];
  const path = glyph.getPath(x, dy, HERO);
  const bb = path.getBoundingBox();
  const cx = (bb.x1 + bb.x2) / 2, cy = (bb.y1 + bb.y2) / 2, a = (r * Math.PI) / 180;
  for (const [px, py] of [[bb.x1, bb.y1], [bb.x2, bb.y1], [bb.x1, bb.y2], [bb.x2, bb.y2]]) {
    const dx = (px - cx) * s, dyy = (py - cy) * s;
    grow(cx + dx * Math.cos(a) - dyy * Math.sin(a), cy + dx * Math.sin(a) + dyy * Math.cos(a));
  }
  return `<path d="${pathD(path, 2)}" transform="rotate(${r} ${cx} ${cy}) translate(${cx} ${cy}) scale(${s}) translate(${-cx} ${-cy})"/>`;
}).join("");
// one long sweeping stroke, thick in the middle and tapering into a dry-brush tail
function sweep(p0, p1, p2, p3, w, n = 140) {
  const pt = (t) => [0, 1].map((k) => (1 - t) ** 3 * p0[k] + 3 * (1 - t) ** 2 * t * p1[k] + 3 * (1 - t) * t ** 2 * p2[k] + t ** 3 * p3[k]);
  const left = [], right = [];
  for (let i = 0; i <= n; i++) {
    const t = i / n, [x, y] = pt(t), [x2, y2] = pt(Math.min(1, t + 1e-3)), [x1, y1] = pt(Math.max(0, t - 1e-3));
    const dx = x2 - x1, dy = y2 - y1, len = Math.hypot(dx, dy) || 1, nx = -dy / len, ny = dx / len;
    const half = (w * (Math.pow(Math.sin(Math.PI * Math.min(1, t * 1.15)), 0.55) * (1 - 0.55 * t) + 0.02)) / 2;
    const L = [x + nx * half, y + ny * half], Rr = [x - nx * half, y - ny * half];
    grow(...L); grow(...Rr);
    left.push(`${L[0].toFixed(1)},${L[1].toFixed(1)}`);
    right.unshift(`${Rr[0].toFixed(1)},${Rr[1].toFixed(1)}`);
  }
  return `M${left.join("L")}L${right.join("L")}Z`;
}
const tail = [wordWidth + 380, -260];
const stroke = sweep([-60, 150], [500, 260], [1300, 40], tail, 34);
const M = 24;
const viewBox = [ink.x1 - M, ink.y1 - M, ink.x2 - ink.x1 + 2 * M, ink.y2 - ink.y1 + 2 * M].map((v) => v.toFixed(1)).join(" ");
const hero = `<svg viewBox="${viewBox}" role="img" aria-label="VisualTeX">
<g class="hero-letters" filter="url(#vt-ink)">${letters}</g>
<g class="hero-sweep"><path d="${stroke}" mask="url(#vt-solid)" filter="url(#vt-ink)"/><path d="${stroke}" mask="url(#vt-dry)" filter="url(#vt-dry-brush)"/></g>
</svg>`;

// ---------------------------------------------------------------- tiny brush rule under card headings
function brushRule() {
  const n = 80, top = [], bot = [];
  for (let i = 0; i <= n; i++) {
    const t = i / n, x = t * 1000;
    const h = 4.2 * Math.pow(Math.sin(Math.PI * Math.min(1, t * 1.05 + 0.02)), 0.45) * (1 - 0.35 * t) + 0.4;
    const y = 12 + 2.2 * Math.sin(t * Math.PI * 0.9);
    top.push(`${x.toFixed(1)} ${(y - h).toFixed(1)}`);
    bot.unshift(`${x.toFixed(1)} ${(y + h).toFixed(1)}`);
  }
  return `<svg class="rule" viewBox="0 0 1000 24" preserveAspectRatio="none" aria-hidden="true"><path d="M${top.join("L")}L${bot.join("L")}Z" filter="url(#vt-ink-small)"/></svg>`;
}

// ---------------------------------------------------------------- background notebook
const field = [
  ...notes.map((n) => ({ html: tex(n.src), size: n.size, once: true, ...(n.pin ? { pin: n.pin } : {}), ...(n.box ? { box: true } : {}) })),
  ...formulas.map((src) => ({ html: tex(src), size: 0 })),
];

// ---------------------------------------------------------------- shared defs
const filters = `
<filter id="vt-ink" x="-5%" y="-20%" width="110%" height="140%">
  <feTurbulence type="fractalNoise" baseFrequency="0.06" numOctaves="3" seed="7" result="edge"/>
  <feDisplacementMap in="SourceGraphic" in2="edge" scale="5" xChannelSelector="R" yChannelSelector="G" result="rough"/>
  <feTurbulence type="fractalNoise" baseFrequency="0.9" numOctaves="1" seed="5" result="grain"/>
  <feColorMatrix in="grain" type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 -6 5.5" result="pits"/>
  <feComposite in="rough" in2="pits" operator="in" result="pitted"/>
  <feTurbulence type="fractalNoise" baseFrequency="0.005" numOctaves="2" seed="11" result="tone"/>
  <feColorMatrix in="tone" type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 0.7 0.62" result="toneA"/>
  <feComposite in="pitted" in2="toneA" operator="in"/>
</filter>
<filter id="vt-dry-brush" x="-5%" y="-30%" width="110%" height="160%">
  <feTurbulence type="fractalNoise" baseFrequency="0.05" numOctaves="2" seed="9" result="edge"/>
  <feDisplacementMap in="SourceGraphic" in2="edge" scale="6" xChannelSelector="R" yChannelSelector="G" result="rough"/>
  <feTurbulence type="fractalNoise" baseFrequency="0.003 0.16" numOctaves="2" seed="3" result="streak"/>
  <feColorMatrix in="streak" type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 -7 3.9" result="hair"/>
  <feComposite in="rough" in2="hair" operator="in"/>
</filter>
<filter id="vt-ink-small" x="-2%" y="-10%" width="104%" height="120%">
  <feTurbulence type="fractalNoise" baseFrequency="0.08" numOctaves="2" seed="4" result="edge"/>
  <feDisplacementMap in="SourceGraphic" in2="edge" scale="2.2" xChannelSelector="R" yChannelSelector="G" result="rough"/>
  <feTurbulence type="fractalNoise" baseFrequency="0.7" numOctaves="1" seed="8" result="grain"/>
  <feColorMatrix in="grain" type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 -7 6.6" result="pits"/>
  <feComposite in="rough" in2="pits" operator="in" result="pitted"/>
  <feTurbulence type="fractalNoise" baseFrequency="0.011 0.02" numOctaves="2" seed="21" result="tone"/>
  <feColorMatrix in="tone" type="matrix" values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 0  0 0 0 1.1 0.3" result="toneA"/>
  <feComposite in="pitted" in2="toneA" operator="in"/>
</filter>
<filter id="vt-ink-solid" x="-2%" y="-10%" width="104%" height="120%">
  <feTurbulence type="fractalNoise" baseFrequency="0.08" numOctaves="2" seed="4" result="edge"/>
  <feDisplacementMap in="SourceGraphic" in2="edge" scale="2.2" xChannelSelector="R" yChannelSelector="G"/>
</filter>
<linearGradient id="vt-tail" gradientUnits="userSpaceOnUse" x1="-60" y1="150" x2="${tail[0]}" y2="${tail[1]}"><stop offset="0" stop-color="#fff"/><stop offset=".5" stop-color="#fff"/><stop offset=".72" stop-color="#000"/></linearGradient>
<linearGradient id="vt-tail-inv" gradientUnits="userSpaceOnUse" x1="-60" y1="150" x2="${tail[0]}" y2="${tail[1]}"><stop offset="0" stop-color="#000"/><stop offset=".5" stop-color="#000"/><stop offset=".72" stop-color="#fff"/></linearGradient>
<mask id="vt-solid" maskUnits="userSpaceOnUse" x="-400" y="-900" width="3000" height="1600"><rect x="-400" y="-900" width="3000" height="1600" fill="url(#vt-tail)"/></mask>
<mask id="vt-dry" maskUnits="userSpaceOnUse" x="-400" y="-900" width="3000" height="1600"><rect x="-400" y="-900" width="3000" height="1600" fill="url(#vt-tail-inv)"/></mask>`;
const glyphSymbols = [...symbols.values()].map((s) => `<symbol id="${s.id}" overflow="visible"><path d="${s.d}"/></symbol>`).join("");
const defs = `${adaptor.outerHTML(svgJax.fontCache.getCache()).replace(/^<defs>|<\/defs>$/g, "")}${filters}${glyphSymbols}`;

const out = `// Generated by scripts/landing/build_landing_art.mjs — do not edit by hand.
// Sources: scripts/landing/copy.mjs (foreground text) and scripts/landing/notebook.mjs (background).
export type InkText = { zh: string; en: string };
export type FieldBlock = { html: string; size: number; once?: boolean; pin?: [number, number]; box?: boolean };

export const ART_DEFS = ${JSON.stringify(`<defs>${defs}</defs>`)};
export const ART_HERO = ${JSON.stringify(hero)};
export const ART_RULE = ${JSON.stringify(brushRule())};
export const ART_TEXT = ${JSON.stringify(text)} as const satisfies Record<string, InkText>;
export type InkKey = keyof typeof ART_TEXT;
export const ART_MARKS = ${JSON.stringify(marks)};
export const ART_FIELD: readonly FieldBlock[] = ${JSON.stringify(field)};
`;
writeFileSync(new URL("../../src/landing/art.generated.ts", import.meta.url), out);
console.log(`art.generated.ts: ${(out.length / 1024).toFixed(0)} KB, ${symbols.size} brush glyphs, ${field.length} field blocks, v${VERSION}`);

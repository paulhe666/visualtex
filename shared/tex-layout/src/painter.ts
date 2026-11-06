// Editor markup for a laid-out formula.
//
// Glyphs are absolutely positioned HTML text in the converted Computer Modern
// fonts (addressed through the Private Use Area); each glyph's baseline is set
// by a strut of known height, so it does not depend on font ascent metrics.
// (SVG text is not used: Chromium keeps a stale font scale for SVG text when a
// CSS transform on an HTML ancestor changes after layout, as in fitted
// previews.) Rules are drawn in one SVG. Every formula atom also gets an
// invisible element carrying data-atom-id over its TeX box, which is what the
// kernel's caret, selection and hit-testing code measures. The caret is an
// in-flow inline-block shifted onto its baseline, so the kernel's caret
// styling and caret-slot code keep working unchanged.
import type { BoxNode, Tag } from "./nodes.ts";
import { isNativeFont } from "./native-font.ts";
import { isPathFont } from "./path-font.ts";
import type { PlacedGlyph, ShipResult } from "./shipout.ts";
import type { TeXFont } from "./tfm.ts";

const TEN_PT = 10 * 65536;
const PUA = 0xe000;

export interface Region {
  tag: Tag;
  x0: number;
  y0: number; // top (sp, y grows downward, formula baseline = 0)
  x1: number;
  y1: number; // bottom
  baseline: number;
  scale: number; // font size / 10pt of the atom's first glyph (1 for display/text)
  hasGlyph: boolean;
}

function grow(map: Map<Tag, Region>, tag: Tag, x0: number, y0: number, x1: number, y1: number, baseline: number, scale: number, glyph: boolean) {
  if (tag === undefined || tag === null) return;
  const r = map.get(tag);
  if (!r) {
    map.set(tag, { tag, x0, y0, x1, y1, baseline, scale, hasGlyph: glyph });
    return;
  }
  r.x0 = Math.min(r.x0, x0);
  r.y0 = Math.min(r.y0, y0);
  r.x1 = Math.max(r.x1, x1);
  r.y1 = Math.max(r.y1, y1);
  if (glyph && !r.hasGlyph) {
    r.baseline = baseline;
    r.scale = scale;
    r.hasGlyph = true;
  }
}

/**
 * The TeX box of every tagged atom (union of its glyphs, rules and boxes).
 * Its baseline is that of the atom's outermost box (which sits on the line
 * that contains the atom), or of its first glyph when it has no box.
 */
export function computeRegions(ship: ShipResult, strut = { height: 7.5 * 65536, depth: 2.5 * 65536 }): Map<Tag, Region> {
  const map = new Map<Tag, Region>();
  const fromBox = new Set<Tag>();
  for (const b of ship.boxes) {
    const box = b.box;
    if (box.tag === undefined || box.tag === null) continue;
    grow(map, box.tag, b.x, b.y - box.height, b.x + box.width, b.y + box.depth, b.y, 1, false);
    fromBox.add(box.tag);
  }
  for (const g of ship.glyphs) {
    const f = g.font;
    const i = f.charInfo(g.char);
    const scale = f.size / TEN_PT;
    const r = map.get(g.tag);
    if (r && fromBox.has(g.tag)) {
      // keep the box baseline; take the font scale of the first glyph
      if (!r.hasGlyph) {
        r.scale = scale;
        r.hasGlyph = true;
      }
    }
    grow(map, g.tag, g.x, g.y - f.charHeight(i), g.x + f.charWidth(i), g.y + f.charDepth(i), g.y, scale, !fromBox.has(g.tag));
  }
  for (const r of ship.rules) grow(map, r.tag, r.x, r.y, r.x + r.width, r.y + r.height, r.y + r.height, 1, false);
  // a character's italic correction: the caret after a slanted letter stands
  // past its overhang
  for (const sp of ship.spans) {
    const r = sp.italic ? map.get(sp.tag) : undefined;
    if (r) r.x1 = Math.max(r.x1, sp.x + sp.width);
  }
  // glue and kerns (spacing atoms): the height of the \Mathstrut@ parenthesis (`strut`)
  for (const sp of ship.spans) {
    if (map.has(sp.tag)) continue;
    grow(map, sp.tag, sp.x, sp.y - strut.height, sp.x + sp.width, sp.y + strut.depth, sp.y, 1, false);
  }
  return map;
}

/** The caret to draw: editor geometry (sp) and the kernel's caret mode. */
export interface CaretBar {
  x: number;
  top: number;
  bottom: number;
  scale: number;
  side: "left" | "center";
  mode: "math" | "text" | "latex";
}

/** An empty-slot marker to draw (sp, formula coordinates). */
export interface PaintMarker {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  selected: boolean;
}

/**
 * An element over part of the formula (sp, formula coordinates). Atom
 * elements carry data-atom-id and nest as the atoms do, so structure queries
 * (closest(".ML__mfrac"), ...) work as on MathLive's own markup.
 */
export interface FormulaElement {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  id?: string | null;
  classes?: string;
  /** font size relative to the formula (a placeholder marker scales with its script) */
  scale?: number;
  /** hidden text content (raw LaTeX input, read back by the editor) */
  text?: string;
  /** markup inside the element, before its children (e.g. a caret) */
  inner?: string;
  children?: FormulaElement[];
  /** CSS length added to the element's left (alignment marker compensation) */
  shift?: string | null;
  /** CSS length added to the element's width */
  extraWidth?: string | null;
  /** the atom the element stands for (not rendered) */
  tag?: Tag;
}

/**
 * Explicit alignment markers (\class{visualtex-align-marker}{\kern0pt}): the
 * editor aligns formula rows by giving marker N extra space through the CSS
 * variables --visualtex-align-marker-N-left/right on the field. Everything
 * after a marker moves by the space of all markers up to it; a marker itself
 * moves by the space before it. `markers` are their x positions (sp).
 */
export function alignShift(markers: number[], x: number, markerIndex?: number): string | null {
  const space = (j: number) => `var(--visualtex-align-marker-${j}-left, 0px) + var(--visualtex-align-marker-${j}-right, 0px)`;
  const terms: string[] = [];
  if (markerIndex !== undefined) {
    for (let j = 0; j < markerIndex; j++) terms.push(space(j));
    terms.push(`var(--visualtex-align-marker-${markerIndex}-left, 0px)`);
  } else {
    for (let j = 0; j < markers.length && markers[j] <= x; j++) terms.push(space(j));
  }
  return terms.length ? terms.join(" + ") : null;
}

const plus = (base: string, extra: string | null | undefined) => (extra ? `calc(${base} + ${extra})` : base);

export interface PaintOptions {
  /** CSS font-family for a TeX font (e.g. "VTX Roman 10") */
  fontFamily(f: TeXFont): string;
  /** extra classes for an atom's glyphs and rules (e.g. ML__selected) */
  atomClasses?(tag: Tag): string;
  /** extra classes for one glyph (a bracket the editor still waits for is faded) */
  glyphClasses?(g: PlacedGlyph): string;
  /** CSS colour of an atom's glyphs (undefined = inherit) */
  atomColor?(tag: Tag): string | undefined;
  /** glyphs not to draw (placeholders: the editor paints a marker instead) */
  hideGlyph?(tag: Tag): boolean;
  /** elements over the formula, positioned from its top-left corner */
  elements?: FormulaElement[];
  /** a caret not inside any element (drawn in the formula's own box) */
  caret?: CaretBar | null;
  /** empty-slot markers */
  markers?: PaintMarker[];
  /**
   * Height and depth of the drawn box (sp) when larger than the TeX box, and
   * how far carets reach right of it: the editor's markers and caret bars
   * must not be clipped by the field.
   */
  extent?: { height: number; depth: number; right?: number };
  /** class of the outer element */
  rootClass?: string;
  /** x positions of explicit alignment markers (see alignShift) */
  alignMarkers?: number[];
  /** contour ovals drawn over integral signs (box in sp, stroke thickness in sp) */
  ovals?: { x0: number; y0: number; x1: number; y1: number; thickness: number; tag: Tag }[];
}

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
const em = (sp: number) => +(sp / TEN_PT).toFixed(4);
const fix = (v: number) => +v.toFixed(4);
/** SVG user units per em (rules) */
const UNITS_PER_EM = 100;
const u = (sp: number) => +((sp / TEN_PT) * UNITS_PER_EM).toFixed(2);
/** height of the strut that puts a glyph's baseline at a known offset (em of the glyph) */
const STRUT = 2;

/**
 * The caret bar, absolutely placed in the element whose top-left corner is
 * (ox, oy) (sp, formula coordinates): the box of its slot, as on MathLive's
 * markup. Its font size is the slot's scale, so the bar's width scales too.
 * The bar is centred on the boundary, or stands left of it against an empty
 * slot's marker.
 */
export function caretBarHtml(c: CaretBar, ox: number, oy: number, shift: string | null = null): string {
  const cls = (c.mode === "text" ? "ML__text-caret" : c.mode === "latex" ? "ML__latex-caret" : "ML__caret") + (c.side === "left" ? " ML__vtx-caret--before" : "");
  const s = c.scale;
  return (
    `<span class="${cls}" style="position:absolute;width:0;font-size:${fix(s)}em;` +
    `left:${plus(`${fix(em(c.x - ox) / s)}em`, shift)};top:${fix(em(c.top - oy) / s)}em;height:${fix(em(c.bottom - c.top) / s)}em"></span>`
  );
}

/** An empty-slot marker (the editor's box in an empty slot), placed like a rule. */
function markerHtml(m: PaintMarker, H: number, shift: string | null): string {
  return (
    `<span class="ML__vtx-marker${m.selected ? " ML__vtx-marker--selected" : ""}" style="position:absolute;pointer-events:none;` +
    `left:${plus(`${fix(em(m.x0))}em`, shift)};top:${fix(em(m.y0) + H)}em;width:${fix(em(m.x1 - m.x0))}em;height:${fix(em(m.y1 - m.y0))}em"></span>`
  );
}

function elementHtml(e: FormulaElement, ox: number, oy: number): string {
  const s = e.scale ?? 1;
  const font = s !== 1 || e.text !== undefined ? `font-size:${fix(s)}em;` : "";
  const hidden = e.text !== undefined ? "color:transparent;overflow:hidden;white-space:pre;" : "";
  const style =
    `position:absolute!important;padding:0;${font}${hidden}left:${plus(`${fix(em(e.x0 - ox) / s)}em`, e.shift)};top:${fix(em(e.y0 - oy) / s)}em;` +
    `width:${plus(`${fix(em(e.x1 - e.x0) / s)}em`, e.extraWidth)};height:${fix(em(e.y1 - e.y0) / s)}em`;
  const id = e.id ? ` data-atom-id="${esc(e.id)}"` : "";
  const cls = e.classes ? ` class="${esc(e.classes)}"` : "";
  const kids = (e.children ?? []).map((c) => elementHtml(c, e.x0, e.y0)).join("");
  return `<span${id}${cls} style="${style}">${e.inner ?? ""}${e.text !== undefined ? esc(e.text) : ""}${kids}</span>`;
}

/**
 * How far the drawn ink reaches outside the drawn box (sp): left of 0, right
 * of its width, above its height, below its depth. TeX lets ink stand out of
 * the boxes (an accent skewed over a slanted letter ends past the letter, a
 * descender starts left of its origin), but the editor clips the formula's
 * box, so the outer element gets this much margin.
 */
export function inkOverflow(ship: ShipResult, width: number, height: number, depth: number, hideGlyph?: (tag: Tag) => boolean): [number, number, number, number] {
  let x0 = 0, x1 = width, y0 = -height, y1 = depth;
  for (const g of ship.glyphs) {
    if (hideGlyph?.(g.tag)) continue;
    const i = g.font.charInfo(g.char);
    const o = g.font.inkOverhang(g.char) ?? [0, 0, 0, 0];
    x0 = Math.min(x0, g.x - o[0]);
    x1 = Math.max(x1, g.x + g.font.charWidth(i) + o[1]);
    y0 = Math.min(y0, g.y - g.font.charHeight(i) - o[2]);
    y1 = Math.max(y1, g.y + g.font.charDepth(i) + o[3]);
  }
  for (const r of ship.rules) {
    x0 = Math.min(x0, r.x);
    x1 = Math.max(x1, r.x + r.width);
    y0 = Math.min(y0, r.y);
    y1 = Math.max(y1, r.y + r.height);
  }
  return [-x0, x1 - width, -height - y0, y1 - depth];
}

export function formulaHtml(box: BoxNode, ship: ShipResult, opts: PaintOptions): string {
  const W = em(box.width);
  const empty = box.width === 0 && box.height === 0 && box.depth === 0;
  const H = Math.max(empty ? 0.75 : em(box.height), opts.extent ? em(opts.extent.height) : 0);
  const D = Math.max(empty ? 0.25 : em(box.depth), opts.extent ? em(opts.extent.depth) : 0);
  const overflow = inkOverflow(ship, box.width, H * TEN_PT, D * TEN_PT, opts.hideGlyph);
  overflow[1] = Math.max(overflow[1], opts.extent?.right ?? 0);
  const [inkLeft, inkRight, inkTop, inkBottom] = overflow.map((v) => Math.max(0, Math.ceil(em(v) * 1000) / 1000));
  const margin = inkLeft || inkRight || inkTop || inkBottom ? `margin:${inkTop}em ${inkRight}em ${inkBottom}em ${inkLeft}em;` : "";
  const markers = opts.alignMarkers ?? [];
  const total = markers.length ? alignShift(markers, Infinity) : null;
  // glyphs
  const glyphs: string[] = [];
  // outline glyphs and ovals go into the SVG of their alignment segment
  const segments = new Map<string, string[]>();
  const svgPart = (x: number, markup: string) => {
    const shift = (markers.length ? alignShift(markers, x) : null) ?? "";
    if (!segments.has(shift)) segments.set(shift, []);
    segments.get(shift)!.push(markup);
  };
  const svgPaint = (color: string | undefined) =>
    !color ? "" : color.startsWith("var(") ? ` style="fill:${esc(color)}"` : ` fill="${esc(color)}"`;
  for (const g of ship.glyphs) {
    if (opts.hideGlyph?.(g.tag)) continue;
    const cls = [opts.atomClasses?.(g.tag), opts.glyphClasses?.(g)].filter(Boolean).join(" ");
    const color = opts.atomColor?.(g.tag);
    if (isPathFont(g.font)) {
      const k = u(g.font.size) / g.font.unitsPerEm;
      svgPart(
        g.x,
        `<path transform="translate(${u(g.x)} ${u(g.y)}) scale(${fix(k)} ${fix(-k)})" d="${esc(g.font.outlines[g.char] ?? "")}"` +
          (cls ? ` class="${cls}"` : "") + svgPaint(color) + `/>`,
      );
      continue;
    }
    const font = g.font;
    const native = isNativeFont(font);
    const family = native ? font.family : `'${esc(opts.fontFamily(font))}'`;
    const style = native && font.style ? font.style + " " : "";
    const size = g.font.size / TEN_PT;
    const shift = markers.length ? alignShift(markers, g.x) : null;
    glyphs.push(
      `<span` + (cls ? ` class="${cls}"` : "") +
        ` style="position:absolute;line-height:0;pointer-events:none;left:${plus(`${fix(em(g.x) / size)}em`, shift)};top:${fix((em(g.y) + H) / size - STRUT)}em;` +
        `font:${style}${fix(size)}em ${esc(family)}` + (color ? `;color:${esc(color)}` : "") + `">` +
        `<span style="display:inline-block;height:${STRUT}em"></span>&#x${(native ? g.char : PUA + g.char).toString(16)};</span>`,
    );
  }
  // rules: boxes the browser snaps to device pixels (crisp), at least 1px
  // thick, centred on the TeX rule (fraction bars, radical roofs, frames)
  const rules: string[] = [];
  for (const r of ship.rules) {
    const color = r.color ?? opts.atomColor?.(r.tag);
    const cls = opts.atomClasses?.(r.tag) ?? "";
    const shift = markers.length ? alignShift(markers, r.x) : null;
    const w = em(r.width), h = em(r.height);
    const cx = fix(em(r.x) + w / 2), cy = fix(em(r.y) + H + h / 2);
    rules.push(
      `<span` + (cls ? ` class="${cls}"` : "") +
        ` style="position:absolute;pointer-events:none;` +
        `left:calc(${cx}em - max(${w}em, 1px) / 2${shift ? " + " + shift : ""});width:max(${w}em, 1px);` +
        `top:calc(${cy}em - max(${h}em, 1px) / 2);height:max(${h}em, 1px);background:${esc(color ?? "currentColor")}"></span>`,
    );
  }
  // contour ovals
  for (const o of opts.ovals ?? []) {
    const color = opts.atomColor?.(o.tag);
    const stroke = !color ? "currentColor" : color;
    svgPart(
      o.x0,
      `<ellipse cx="${u((o.x0 + o.x1) / 2)}" cy="${u((o.y0 + o.y1) / 2)}" rx="${u((o.x1 - o.x0) / 2)}" ry="${u((o.y1 - o.y0) / 2)}" ` +
        `fill="none" stroke="${esc(stroke)}" stroke-width="${u(o.thickness)}"/>`,
    );
  }
  const svg = [...segments]
    .map(
      ([shift, rects]) =>
        `<svg class="ML__vtx-ink" xmlns="http://www.w3.org/2000/svg" viewBox="0 ${fix(-H * UNITS_PER_EM)} ${fix(W * UNITS_PER_EM)} ${fix((H + D) * UNITS_PER_EM)}" ` +
        `style="position:absolute;left:${plus("0em", shift || null)};top:0;width:${W}em;height:${fix(H + D)}em;overflow:visible;pointer-events:none" fill="currentColor" aria-hidden="true">` +
        rects.join("") +
        `</svg>`,
    )
    .join("");
  // elements: positioned from the top-left corner, whose baseline is at H
  const top = -Math.round(H * TEN_PT);
  const elements = (opts.elements ?? []).map((e) => elementHtml(e, 0, top)).join("");
  return (
    `<span class="${opts.rootClass ?? "ML__latex ML__vtx"}" translate="no" aria-hidden="true" ` +
    `style="display:inline-block;position:relative;line-height:0;${margin}width:${plus(`${W}em`, total)};height:${fix(H + D)}em">` +
    `<span class="ML__strut" style="display:inline-block;width:0;min-height:0;height:${H}em"></span>` +
    (D !== 0 ? `<span class="ML__strut--bottom" style="display:inline-block;width:0;min-height:0;height:${fix(H + D)}em;vertical-align:${-D}em"></span>` : "") +
    (opts.markers ?? []).map((m) => markerHtml(m, H, markers.length ? alignShift(markers, m.x0) : null)).join("") +
    rules.join("") +
    svg +
    glyphs.join("") +
    elements +
    (opts.caret ? caretBarHtml(opts.caret, 0, top, markers.length ? alignShift(markers, opts.caret.x) : null) : "") +
    `</span>`
  );
}

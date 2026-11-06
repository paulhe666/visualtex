// Paint a shipped-out formula as SVG (glyph outlines, rules and, for the
// characters the TeX fonts lack, text in their native font): the images the
// app exports, copies and places in Office documents. One TeX "em" (10pt at
// the formula's size) maps to `pxPerEm` pixels; every font is scaled from its
// TeX size, so the image is the editor's layout exactly.
import type { BoxNode, Tag } from "./nodes.ts";
import { isNativeFont } from "./native-font.ts";
import { inkOverflow } from "./painter.ts";
import { isPathFont } from "./path-font.ts";
import type { ShipResult } from "./shipout.ts";

/** Outlines of one TeX font: SVG path data per character code, font units, y up. */
export interface GlyphTable {
  tfm: string;
  unitsPerEm: number;
  glyphs: Record<string, string | { d: string }>;
}

/** Something drawn over an atom's box (sp, formula coordinates), e.g. a custom symbol's artwork. */
export interface SvgOverlay {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
  color?: string;
  /** markup in a box `width` x `height` px whose top-left corner is (0, 0) */
  markup(width: number, height: number): string;
}

export interface SvgOptions {
  /** pixels per TeX em (10pt at the formula's size) */
  pxPerEm: number;
  /** outlines of a TeX font by TFM name; undefined when not loaded (the font is listed in `missing`) */
  glyphs(tfm: string): GlyphTable | undefined;
  /** colour of glyphs and rules without one of their own */
  color?: string;
  atomColor?(tag: Tag): string | undefined;
  /** glyphs not to draw */
  hideGlyph?(tag: Tag): boolean;
  /** contour ovals over integral signs (box and stroke thickness in sp) */
  ovals?: { x0: number; y0: number; x1: number; y1: number; thickness: number; tag: Tag }[];
  overlays?: SvgOverlay[];
  /** prefix of the ids of glyph definitions (unique when several images share a page) */
  idPrefix?: string;
}

export interface SvgFormula {
  /** <defs> and drawing, in px, the formula's reference point (left end of its baseline) at (0, 0) */
  body: string;
  /** how far the drawing reaches from the reference point, px: left (>= 0), right, above and below the baseline */
  left: number;
  right: number;
  ascent: number;
  descent: number;
  /** TeX fonts whose outlines were not available (their glyphs are missing) */
  missing: string[];
}

const TEN_PT = 10 * 65536;

const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const fmt = (v: number) => {
  const r = Math.round(v * 1000) / 1000;
  return (Object.is(r, -0) ? 0 : r).toString();
};

function pathOf(entry: string | { d: string } | undefined): string | undefined {
  return typeof entry === "string" ? entry : entry?.d;
}

export function formulaSvgBody(box: BoxNode, ship: ShipResult, opts: SvgOptions): SvgFormula {
  const k = opts.pxPerEm / TEN_PT; // px per sp
  const prefix = opts.idPrefix ?? "vtx-";
  const fill = (color: string | undefined) => (color ? ` fill="${esc(color)}"` : "");
  const defs = new Map<string, string>();
  const missing = new Set<string>();
  const parts: string[] = [];
  // rules first: a \colorbox's background is one
  for (const r of ship.rules) {
    const color = r.color ?? opts.atomColor?.(r.tag);
    parts.push(
      `<rect x="${fmt(r.x * k)}" y="${fmt(r.y * k)}" width="${fmt(r.width * k)}" height="${fmt(r.height * k)}"${fill(color)}/>`,
    );
  }
  for (const g of ship.glyphs) {
    if (opts.hideGlyph?.(g.tag)) continue;
    const color = opts.atomColor?.(g.tag);
    const x = g.x * k, y = g.y * k;
    if (isPathFont(g.font)) {
      const d = g.font.outlines[g.char];
      if (!d) continue;
      const s = (g.font.size * k) / g.font.unitsPerEm;
      parts.push(`<path transform="translate(${fmt(x)} ${fmt(y)}) scale(${fmt(s)} ${fmt(-s)})" d="${esc(d)}"${fill(color)}/>`);
      continue;
    }
    if (isNativeFont(g.font)) {
      const style = g.font.style.split(/\s+/);
      parts.push(
        `<text x="${fmt(x)}" y="${fmt(y)}" font-family="${esc(g.font.family)}" font-size="${fmt(g.font.size * k)}"` +
          (style.includes("italic") ? ` font-style="italic"` : "") +
          (style.includes("bold") ? ` font-weight="bold"` : "") +
          `${fill(color)}>${esc(String.fromCodePoint(g.char))}</text>`,
      );
      continue;
    }
    const table = opts.glyphs(g.font.name);
    if (!table) {
      missing.add(g.font.name);
      continue;
    }
    const d = pathOf(table.glyphs[String(g.char)]);
    if (!d) continue;
    const id = `${prefix}${table.tfm.replace(/[^A-Za-z0-9_-]/g, "_")}-${g.char}`;
    if (!defs.has(id)) defs.set(id, d);
    const s = (g.font.size * k) / table.unitsPerEm;
    parts.push(
      `<use xlink:href="#${id}" href="#${id}" transform="translate(${fmt(x)} ${fmt(y)}) scale(${fmt(s)} ${fmt(-s)})"${fill(color)}/>`,
    );
  }
  for (const o of opts.ovals ?? []) {
    const color = opts.atomColor?.(o.tag) ?? opts.color ?? "#000";
    parts.push(
      `<ellipse cx="${fmt(((o.x0 + o.x1) / 2) * k)}" cy="${fmt(((o.y0 + o.y1) / 2) * k)}" rx="${fmt(((o.x1 - o.x0) / 2) * k)}" ` +
        `ry="${fmt(((o.y1 - o.y0) / 2) * k)}" fill="none" stroke="${esc(color)}" stroke-width="${fmt(o.thickness * k)}"/>`,
    );
  }
  for (const o of opts.overlays ?? []) {
    const w = (o.x1 - o.x0) * k, h = (o.y1 - o.y0) * k;
    if (!(w > 0) || !(h > 0)) continue;
    parts.push(`<g transform="translate(${fmt(o.x0 * k)} ${fmt(o.y0 * k)})"${fill(o.color)}>${o.markup(w, h)}</g>`);
  }
  const [inkLeft, inkRight, inkTop, inkBottom] = inkOverflow(ship, box.width, box.height, box.depth, opts.hideGlyph);
  const defsXml = defs.size ? `<defs>${[...defs].map(([id, d]) => `<path id="${id}" d="${esc(d)}"/>`).join("")}</defs>` : "";
  return {
    body: `${defsXml}<g fill="${esc(opts.color ?? "#000")}">${parts.join("")}</g>`,
    left: Math.max(0, inkLeft) * k,
    right: (box.width + Math.max(0, inkRight)) * k,
    ascent: (box.height + Math.max(0, inkTop)) * k,
    descent: (box.depth + Math.max(0, inkBottom)) * k,
    missing: [...missing],
  };
}

/** A standalone SVG of the formula with `padding` px around its drawing. */
export function formulaSvg(box: BoxNode, ship: ShipResult, opts: SvgOptions & { padding?: number }): string {
  const f = formulaSvgBody(box, ship, opts);
  const pad = opts.padding ?? 0;
  const w = f.left + f.right + 2 * pad;
  const h = f.ascent + f.descent + 2 * pad;
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="${fmt(w)}" height="${fmt(h)}" ` +
    `viewBox="${fmt(-f.left - pad)} ${fmt(-f.ascent - pad)} ${fmt(w)} ${fmt(h)}">${f.body}</svg>`
  );
}

// Pointer (docs/editing-core/DESIGN.md §5, §6.9): the caret stop a point of
// the formula means. Clicks, drags, Shift+click and the app's row clicks all
// ask hitTest; a double click selects the item under the point.
import type { EditorGeometry } from "../../tex-layout/src/index.ts";
import { allowedStops, type CoreModel } from "./stops.ts";
import { slotItems, slotOrder, type CoreAtom } from "./structures.ts";

/** A point of the formula: sp, x from its left edge, y down from its baseline. */
export interface FormulaPoint {
  x: number;
  y: number;
}

interface Box {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/** A click this close to an empty-slot marker (in em of the slot's size) enters its slot. */
const MARKER_REACH = 0.15;
/** How much a vertical distance counts against a horizontal one, off the glyphs. */
const VERTICAL_WEIGHT = 0.5;
/** The marker's height in em, the caret bar's (editor-geometry.ts): converts its size to the slot's em. */
const MARKER_EM = 0.76;

const inside = (b: Box, p: FormulaPoint, pad = 0) => p.x >= b.x0 - pad && p.x <= b.x1 + pad && p.y >= b.y0 - pad && p.y <= b.y1 + pad;
const isLeaf = (atom: CoreAtom) => slotOrder(atom).length === 0;

/**
 * An atom's box for hit testing: it ends where its scripts (a subsup right
 * after it) begin, also where TeX boxed them with it (an accented letter).
 */
function regionOf(geometry: EditorGeometry, atom: CoreAtom): Box | undefined {
  const r: Box | undefined = geometry.regions.get(atom as never);
  const next = atom.rightSibling;
  const scripts: Box | undefined = next?.type === "subsup" ? geometry.regions.get(next as never) : undefined;
  return r && scripts && scripts.x0 < r.x1 ? { ...r, x1: Math.max(r.x0, scripts.x0) } : r;
}

/** The innermost atom whose box holds the point (the root when none does). */
export function atomAt(model: CoreModel, geometry: EditorGeometry, p: FormulaPoint): CoreAtom {
  let atom = model.root;
  for (;;) {
    let best: CoreAtom | null = null;
    let bestArea = Infinity;
    for (const key of slotOrder(atom)) {
      for (const item of slotItems({ owner: atom, key })) {
        const r = regionOf(geometry, item);
        if (!r || !inside(r, p)) continue;
        const area = (r.x1 - r.x0) * (r.y1 - r.y0);
        if (area < bestArea) {
          best = item;
          bestArea = area;
        }
      }
    }
    if (!best) return atom;
    atom = best;
  }
}

/** The stop of the empty slot whose marker is under the point. */
function markerStop(model: CoreModel, geometry: EditorGeometry, p: FormulaPoint, allowed: Set<number>): number | null {
  let best: number | null = null;
  let bestDistance = Infinity;
  for (const [placeholder, m] of geometry.markers) {
    if (!inside(m, p, (MARKER_REACH * (m.y1 - m.y0)) / MARKER_EM)) continue;
    const before = (placeholder as CoreAtom).leftSibling;
    const stop = before ? model.offsetOf(before) : -1;
    if (!allowed.has(stop)) continue;
    const distance = Math.hypot(p.x - (m.x0 + m.x1) / 2, p.y - (m.y0 + m.y1) / 2);
    if (distance < bestDistance) {
      best = stop;
      bestDistance = distance;
    }
  }
  return best;
}

/**
 * Beside the formula, as beside a line of text: its first stop left of it,
 * its last stop right of it (the formula's carets may reach past its box:
 * \cfrac ends with a negative kern).
 */
function beside(geometry: EditorGeometry, stops: number[], p: FormulaPoint): number | null {
  const first = geometry.stops[stops[0]];
  const last = geometry.stops[stops[stops.length - 1]];
  if (p.x < Math.min(0, first?.x ?? 0)) return stops[0];
  if (p.x > Math.max(geometry.width, last?.x ?? 0)) return stops[stops.length - 1];
  return null;
}

/** The allowed stops inside a structure (its descendants come right before it in offset order). */
function stopsWithin(model: CoreModel, atom: CoreAtom, allowed: Set<number>): number[] {
  const end = model.offsetOf(atom);
  const out: number[] = [];
  for (let stop = end - atom.children.length; stop < end; stop++) if (allowed.has(stop)) out.push(stop);
  return out;
}

/** How many structures hold the atom before a stop. */
function depthOf(model: CoreModel, stop: number): number {
  let depth = 0;
  for (let atom = model.at(stop); atom?.parent; atom = atom.parent) depth++;
  return depth;
}

/**
 * The candidate whose caret is nearest the point (an empty slot's caret covers
 * its marker). Carets that coincide (the end of a radicand and the end of the
 * radical) go to the outer stop; the inner one is reached from the item
 * beside it.
 */
function nearest(model: CoreModel, candidates: number[], geometry: EditorGeometry, p: FormulaPoint): number {
  let best = candidates[0];
  let bestDistance = Infinity;
  let bestDepth = Infinity;
  for (const stop of candidates) {
    const c = geometry.stops[stop];
    if (!c) continue;
    const m = c.empty ? geometry.markers.get(c.empty as never) : undefined;
    const dx = m ? Math.max(0, m.x0 - p.x, p.x - m.x1) : Math.abs(p.x - c.x);
    const dy = (p.y - (c.top + c.bottom) / 2) * VERTICAL_WEIGHT;
    const distance = dx * dx + dy * dy;
    if (distance > bestDistance) continue;
    const depth = depthOf(model, stop);
    if (distance < bestDistance || depth < bestDepth) {
      best = stop;
      bestDistance = distance;
      bestDepth = depth;
    }
  }
  return best;
}

/**
 * The stop a point means: beside the formula, its start or end; on an
 * empty-slot marker, that slot; on a glyph, the side of it the point is on;
 * on a structure's own ink (a fraction bar, a radical sign, a delimiter) or in
 * a gap inside it, the structure's nearest stop; elsewhere the nearest stop.
 */
export function hitTest(model: CoreModel, geometry: EditorGeometry, p: FormulaPoint): number | null {
  const stops = allowedStops(model);
  if (!stops.length) return null;
  const side = beside(geometry, stops, p);
  if (side !== null) return side;
  const allowed = new Set(stops);
  const marker = markerStop(model, geometry, p, allowed);
  if (marker !== null) return marker;
  for (let atom: CoreAtom | undefined = atomAt(model, geometry, p); atom && atom !== model.root; atom = atom.parent) {
    const before = atom.leftSibling ? model.offsetOf(atom.leftSibling) : -1;
    const after = model.offsetOf(atom);
    if (isLeaf(atom)) {
      const b = allowed.has(before);
      const a = allowed.has(after);
      const r = regionOf(geometry, atom);
      if (a && b && r) return p.x < (r.x0 + r.x1) / 2 ? before : after;
      if (a || b) return a ? after : before;
    } else {
      const candidates = stopsWithin(model, atom, allowed);
      if (allowed.has(before)) candidates.push(before);
      if (allowed.has(after)) candidates.push(after);
      if (candidates.length) return nearest(model, candidates, geometry, p);
    }
    // nothing to stand at here (inside an accent or a macro's expansion): the enclosing structure
  }
  return nearest(model, stops, geometry, p);
}

/** The item a double click selects: the glyph or structure under the point (null: none). */
export function itemAt(model: CoreModel, geometry: EditorGeometry, p: FormulaPoint): CoreAtom | null {
  const stops = allowedStops(model);
  if (!stops.length || beside(geometry, stops, p) !== null) return null;
  const allowed = new Set(stops);
  if (markerStop(model, geometry, p, allowed) !== null) return null;
  for (let atom: CoreAtom | undefined = atomAt(model, geometry, p); atom && atom !== model.root; atom = atom.parent) {
    const before = atom.leftSibling ? model.offsetOf(atom.leftSibling) : -1;
    if (allowed.has(before) && allowed.has(model.offsetOf(atom))) return atom;
  }
  return null;
}

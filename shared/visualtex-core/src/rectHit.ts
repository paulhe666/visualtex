// Click -> caret (docs/editing-core/DESIGN.md §5 hitTest, §6.9), AxMath's rule
// (research: claude-work/axmath-mouse-reverse, selection-caret.md §2): no
// nearest-point search and no glyph outlines, but rectangles. Each atom is a
// rectangle in device pixels: across, its advance, from the caret before it to
// the caret after it (AxMath's boxes are layout boxes, so the space around an
// operator is the operator's; ink boxes would leave gaps between atoms), grown
// 1 px right and 1 px left (3 px at the start of a slot, so the caret before
// the first item is easy to reach); up and down, its own box grown by a
// fraction of the font size, at the top level of the formula the whole row.
// The point is tested along a slot's atoms (where two rectangles meet, the
// first one's: a base takes the pixel or two where its scripts start); in an
// atom, its slots are tried first, in their order; when no slot holds the
// point, the middle of the atom's rectangle decides: the caret after the atom
// (right half) or before it (left half). The start of a slot (an empty slot:
// its marker) holds the caret at that start. Nothing in the row: the
// formula's start or end, whichever side of its start the point is. Pixels
// are truncated as Windows does and rectangles hold their left and top edges
// but not their right and bottom ones.
import type { EditorGeometry } from "../../tex-layout/src/index.ts";
import { regionOf } from "./pointer.ts";
import { allowedStops, type CoreModel } from "./stops.ts";
import { isPlaceholder, slotAtoms, slotOrder, type CoreAtom } from "./structures.ts";

/**
 * How far an atom's rectangle reaches above and below its box, in em of the
 * formula's font. AxMath reads it from its style profile (+0x14 times the font
 * size); the value is not in the program, so this one is a choice.
 */
export const HIT_VERTICAL_MARGIN_EM = 0.2;

interface FormulaBox {
  x0: number;
  x1: number;
  y0: number;
  y1: number;
}

/** A page rectangle in CSS pixels. */
interface PageBox {
  left: number;
  right: number;
  top: number;
  bottom: number;
}

/** What the rectangle hit test reads of a rendered formula. */
export interface RectHitView {
  model: CoreModel;
  geometry: EditorGeometry;
  /** A formula-unit box in page (CSS) pixels. */
  toPage(box: FormulaBox): PageBox | null;
  /** The row the formula stands in, in page pixels: its top level spans it. */
  band: { top: number; bottom: number };
  /** Device pixels per CSS pixel: the pixels AxMath counts in. */
  scale: number;
  /** The formula's font size in CSS pixels. */
  fontPx: number;
}

/** Device pixels: left and top included, right and bottom excluded (PtInRect). */
interface PixelRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

const contains = (r: PixelRect, x: number, y: number) => x >= r.left && x < r.right && y >= r.top && y < r.bottom;

const topLevel = (model: CoreModel, atom: CoreAtom) => atom.parent === model.root;

/** The atoms of the slot an atom is in, with the slot's start ("first") at the head. */
function siblings(atom: CoreAtom): CoreAtom[] {
  if (!atom.parent || atom.parentBranch === undefined) return [];
  return slotAtoms({ owner: atom.parent, key: atom.parentBranch });
}

/** The box of a slot's start: the empty slot's marker, else the caret's band at that start. */
function startBox(view: RectHitView, first: CoreAtom): FormulaBox | null {
  const items = siblings(first).filter((a) => a.type !== "first");
  if (items.length === 1 && isPlaceholder(items[0])) {
    const marker = view.geometry.markers.get(items[0] as never);
    if (marker) return marker;
  }
  const stop = view.geometry.stops[view.model.offsetOf(first)];
  return stop ? { x0: stop.x, x1: stop.x, y0: stop.top, y1: stop.bottom } : null;
}

/** An atom's own box (a placeholder among other items: its marker). */
function boxOf(view: RectHitView, atom: CoreAtom): FormulaBox | null {
  return regionOf(view.geometry, atom) ?? (isPlaceholder(atom) ? view.geometry.markers.get(atom as never) ?? null : null);
}

/**
 * An atom's box across its advance, as tall as its own box: from the caret
 * before it (or its ink, where that begins further left: a script after a
 * structure) to the caret after it, but not past where the next item's ink
 * begins (the caret after a structure is drawn a little outside it, over a
 * script that follows; an accent's arrow may reach over its script).
 */
function advanceBox(view: RectHitView, atom: CoreAtom): FormulaBox | null {
  const region = boxOf(view, atom);
  if (!region) return null;
  const model = view.model;
  const before = atom.leftSibling ? view.geometry.stops[model.offsetOf(atom.leftSibling)] : null;
  const after = view.geometry.stops[model.offsetOf(atom)];
  const next = atom.rightSibling ? boxOf(view, atom.rightSibling) : null;
  const x0 = before ? Math.min(region.x0, before.x) : region.x0;
  let x1 = after ? after.x : region.x1;
  if (next && next.x0 > x0 && next.x0 < x1) x1 = next.x0;
  return { x0, x1, y0: region.y0, y1: region.y1 };
}

function hitRect(view: RectHitView, atom: CoreAtom): PixelRect | null {
  const start = atom.type === "first";
  const box = start ? startBox(view, atom) : advanceBox(view, atom);
  const page = box && view.toPage(box);
  if (!page) return null;
  const s = view.scale;
  const margin = Math.trunc(HIT_VERTICAL_MARGIN_EM * view.fontPx * s);
  const vertical = topLevel(view.model, atom) ? view.band : page;
  return {
    left: Math.trunc(page.left * s) - Math.trunc(start ? 3 * s : s),
    right: Math.trunc(page.right * s) + Math.trunc(s),
    top: Math.trunc(vertical.top * s) - margin,
    bottom: Math.trunc(vertical.bottom * s) + margin,
  };
}

/**
 * The atom the caret goes after (a slot's "first": the slot's start) along a
 * slot's atoms, or null when no atom's rectangle holds the point.
 */
function hitSlot(view: RectHitView, atoms: CoreAtom[], x: number, y: number): CoreAtom | null {
  const empty = atoms.length === 2 && isPlaceholder(atoms[1]);
  for (let i = 0; i < atoms.length; i++) {
    const atom = atoms[i];
    // an empty slot is its start, whose rectangle is the marker
    if (empty && i === 1) continue;
    const rect = hitRect(view, atom);
    if (!rect || !contains(rect, x, y)) continue;
    if (atom.type === "first") return atom;
    const inSlots = hitSlots(view, atom, x, y);
    if (inSlots) return inSlots;
    const middle = Math.trunc((rect.left + rect.right) / 2);
    return x >= middle ? atom : (atoms[i - 1] ?? atom);
  }
  return null;
}

/** The first of an atom's slots, in their order, that holds the point (a macro's expansion too: the stop is checked after). */
function hitSlots(view: RectHitView, atom: CoreAtom, x: number, y: number): CoreAtom | null {
  for (const key of slotOrder(atom)) {
    const hit = hitSlot(view, slotAtoms({ owner: atom, key }), x, y);
    if (hit) return hit;
  }
  return null;
}

/**
 * The caret stop a page point means in a rendered formula (AxMath's rule;
 * null: nothing to stand at, which only a formula without stops gives).
 */
export function rectHitTest(view: RectHitView, clientX: number, clientY: number): number | null {
  const model = view.model;
  const stops = allowedStops(model);
  if (!stops.length) return null;
  const allowed = new Set(stops);
  const s = view.scale;
  const x = Math.trunc(clientX * s);
  const y = Math.trunc(clientY * s);
  const root = model.root;
  const body = root?.branch?.("body") ?? [];
  const hit = hitSlot(view, body, x, y);
  if (!hit) {
    // beside every atom: the start or the end of the formula
    const start = view.geometry.stops[stops[0]];
    const startPage = start ? view.toPage({ x0: start.x, x1: start.x, y0: start.top, y1: start.bottom }) : null;
    const startX = startPage ? Math.trunc(startPage.left * s) : 0;
    return x < startX ? stops[0] : stops[stops.length - 1];
  }
  const stop = model.offsetOf(hit);
  if (allowed.has(stop)) return stop;
  // inside something the caret does not enter (an accent's base, a macro's
  // expansion): before or after the nearest enclosing atom it may stand beside
  for (let atom: CoreAtom | undefined = hit; atom && atom !== root; atom = atom.parent) {
    const after = model.offsetOf(atom);
    const before = atom.leftSibling ? model.offsetOf(atom.leftSibling) : -1;
    const a = allowed.has(after);
    const b = allowed.has(before);
    if (!a && !b) continue;
    if (a && b) {
      const rect = hitRect(view, atom);
      if (rect) return x >= Math.trunc((rect.left + rect.right) / 2) ? after : before;
    }
    return a ? after : before;
  }
  return stops[stops.length - 1];
}

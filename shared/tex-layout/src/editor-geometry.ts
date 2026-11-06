// Editor geometry (docs/editing-core/DESIGN.md §5): where every caret stop,
// empty-slot marker and atom of an edited formula is, in formula units (sp;
// x from the formula's left edge, y down from its baseline). The editing core
// draws, hit-tests and anchors popovers with these numbers; nothing measures
// the page.
import type { Tag } from "./nodes.ts";
import type { Region } from "./painter.ts";
import type { ShipResult } from "./shipout.ts";

const EM = 10 * 65536;

/**
 * The caret bar, in every slot: from `up` above the baseline the caret stands
 * on to `down` below it, scaled with the slot's size (a script's bar is a
 * script's size). It never stretches to the slot's content, so the bar beside
 * a fraction or inside a nested structure is the size of the text it is in.
 * An empty slot's marker spans the same band, `markerWidth` wide, so the caret
 * keeps its size and place when the first character is typed into the slot.
 * `strut`: the height and depth of an empty formula (amsmath's \Mathstrut@).
 * All in sp at 10pt; they come from the profile's fonts (caretMetrics).
 */
export interface CaretMetrics {
  up: number;
  down: number;
  markerWidth: number;
  strut: { height: number; depth: number };
}

/** Computer Modern: 0.71em up (the height of a "d" and 0.0156em), 0.05em down, 0.4em markers. */
export const CM_CARET: CaretMetrics = { up: 0.71 * EM, down: 0.05 * EM, markerWidth: 0.4 * EM, strut: { height: 0.75 * EM, depth: 0.25 * EM } };

/**
 * A profile's caret, in the proportions of Computer Modern's: the bar's top
 * as far above the roman "d" as there, its foot and the strut from the
 * \Mathstrut@ parenthesis, the marker as wide relative to a math italic "x".
 * (Computer Modern's own fonts give CM_CARET exactly.)
 */
export function caretMetrics(fonts: { dHeight: number; paren: { height: number; depth: number }; xWidth: number }): CaretMetrics {
  const CM_D = 455111; // height of "d" in cmr10, sp at 10pt
  const CM_X = 374556; // width of "x" in cmmi10
  return {
    up: CM_CARET.up + (fonts.dHeight - CM_D),
    down: (CM_CARET.down * fonts.paren.depth) / CM_CARET.strut.depth,
    markerWidth: (CM_CARET.markerWidth * fonts.xWidth) / CM_X,
    strut: { ...fonts.paren },
  };
}

/** carets of neighbouring stops closer than this look like one (about 0.6px) */
const SAME_PLACE = 0.02 * EM;
/** how far apart such stops are drawn, in em of the slot's size (a little more than the bar's 0.08em width) */
const STOP_GAP = 0.1 * EM;
/** half the caret bar's width (the CSS bar is 0.08em of the slot's size) */
const BAR_HALF = 0.04 * EM;

export interface Rect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/** The part of a kernel atom the geometry reads. */
export interface GeometryAtom {
  type?: string;
  parent?: GeometryAtom;
  parentBranch?: unknown;
  branch(name: unknown): GeometryAtom[] | undefined;
  branches?: string[];
  rows?: unknown[][][];
}

export interface CaretGeometry {
  /** the boundary the caret stands at */
  x: number;
  top: number;
  bottom: number;
  baseline: number;
  /** font scale of the slot (1 text, 0.7 script, 0.5 scriptscript) */
  scale: number;
  /**
   * where the bar stands: centred on the boundary (letters touch in TeX, so
   * the bar shares the overlap evenly), or left of it, touching an empty
   * slot's marker
   */
  side: "left" | "center";
  /** the stop is an empty slot: the marker the caret touches */
  empty?: GeometryAtom;
}

export interface EditorGeometry {
  width: number;
  /** height and depth of the drawn box: the TeX box grown to cover every marker and caret position */
  height: number;
  depth: number;
  /** how far the carets reach right of the TeX box (a caret set off after a structure at the end) */
  right: number;
  /** caret geometry of every kernel offset (index = offset) */
  stops: (CaretGeometry | null)[];
  /** empty-slot markers, by placeholder atom */
  markers: Map<Tag, Rect>;
  /** each atom's region: its own material and its descendants' */
  regions: Map<Tag, Region>;
  /** x of the explicit alignment markers of the main row (sp) */
  alignMarkers: number[];
}

export interface GeometryModel {
  at(offset: number): GeometryAtom | undefined;
  lastOffset: number;
}

const isPlaceholder = (a: GeometryAtom | undefined) => a?.type === "placeholder" || a?.type === "prompt";

function siblingsOf(a: GeometryAtom): GeometryAtom[] {
  return a.parent?.branch(a.parentBranch) ?? [];
}

export function buildEditorGeometry(
  model: GeometryModel,
  regions: Map<Tag, Region>,
  ship: ShipResult,
  box: { width: number; height: number; depth: number },
  alignMarkers: number[],
  scales?: Map<Tag, number>,
  caret: CaretMetrics = CM_CARET,
): EditorGeometry {
  const { up: BAR_UP, down: BAR_DOWN, markerWidth: MARKER_WIDTH } = caret;
  const markers = new Map<Tag, Rect>();
  const lineCache = new Map<GeometryAtom, number>();
  const scaleCache = new Map<GeometryAtom, Map<string, number>>();
  // the box TeX made of each slot (a field), by the slot's start atom
  const slotBoxes = new Map<Tag, { x: number; y: number }>();
  for (const b of ship.boxes) {
    const first = b.box.firstTag;
    if (first !== undefined && !slotBoxes.has(first)) slotBoxes.set(first, b);
  }

  // the line an atom's own material sits on (a big operator centred on the
  // axis stays on its line; an empty-nucleus subsup belongs to its nucleus)
  const baselineOf = (a: GeometryAtom): number => {
    const known = lineCache.get(a);
    if (known !== undefined) return known;
    let v: number | undefined;
    if (a.type === "subsup") {
      const sibs = siblingsOf(a);
      const prev = sibs[sibs.indexOf(a) - 1];
      if (prev && prev.type !== "first") v = baselineOf(prev);
    }
    if (v === undefined) v = ship.lineOf.get(a);
    if (v === undefined) v = regions.get(a)?.baseline ?? 0;
    lineCache.set(a, v);
    return v;
  };

  // the font scale an atom is set at: that of the style TeX set it in (the
  // typesetter records it for editor layouts), else that of its own character
  const ownScale = (a: GeometryAtom): number | undefined => {
    const recorded = scales?.get(a);
    if (recorded !== undefined) return recorded;
    if (a.branches?.length || a.rows) return undefined;
    const r = regions.get(a);
    return r?.hasGlyph ? r.scale : undefined;
  };

  // the font scale of a slot: that of its first item, else its parent slot's
  const slotScale = (owner: GeometryAtom | undefined, branch: unknown): number => {
    if (!owner || owner.type === "root") return 1;
    const key = String(branch);
    let byBranch = scaleCache.get(owner);
    if (!byBranch) scaleCache.set(owner, (byBranch = new Map()));
    const cached = byBranch.get(key);
    if (cached !== undefined) return cached;
    let s: number | undefined;
    for (const item of owner.branch(branch) ?? []) {
      if (item.type === "first") continue;
      s = ownScale(item);
      if (s !== undefined) break;
    }
    if (s === undefined) s = slotScale(owner.parent, owner.parentBranch);
    byBranch.set(key, s);
    return s;
  };

  // the caret beside an atom has the atom's scale, else its slot's
  const scaleOf = (a: GeometryAtom): number => ownScale(a) ?? slotScale(a.parent, a.parentBranch);

  for (const [tag, r] of regions) {
    if (!isPlaceholder(tag as GeometryAtom)) continue;
    const s = r.scale;
    const cx = (r.x0 + r.x1) / 2;
    const baseline = baselineOf(tag as GeometryAtom);
    markers.set(tag, {
      x0: Math.round(cx - (MARKER_WIDTH * s) / 2),
      x1: Math.round(cx + (MARKER_WIDTH * s) / 2),
      y0: Math.round(baseline - BAR_UP * s),
      y1: Math.round(baseline + BAR_DOWN * s),
    });
  }

  const band = (baseline: number, s: number) => ({ top: baseline - BAR_UP * s, bottom: baseline + BAR_DOWN * s });

  const stops: (CaretGeometry | null)[] = [];
  for (let offset = 0; offset <= model.lastOffset; offset++) {
    const a = model.at(offset);
    if (!a) {
      stops.push(null);
      continue;
    }
    const owner = a.parent;
    const branch = a.parentBranch;
    const items = (owner?.branch(branch) ?? []).filter((x) => x.type !== "first");
    if (a.type === "first") {
      const only = items.length === 1 ? items[0] : undefined;
      const marker = only && isPlaceholder(only) ? markers.get(only) : undefined;
      if (marker) {
        stops.push({ x: marker.x0, top: marker.y0, bottom: marker.y1, baseline: baselineOf(only!), scale: scaleOf(only!), side: "left", empty: only });
        continue;
      }
      const lead = items.find((x) => regions.has(x));
      if (lead) {
        const baseline = baselineOf(lead);
        const s = scaleOf(lead);
        stops.push({ x: regions.get(lead)!.x0, ...band(baseline, s), baseline, scale: s, side: "center" });
        continue;
      }
      // an empty slot without a marker (a script just opened): at the empty
      // box TeX set for it, at its size
      const slotBox = slotBoxes.get(a);
      const o = owner && owner.type !== "root" ? regions.get(owner) : undefined;
      const baseline = slotBox ? slotBox.y : o ? o.baseline : 0;
      const s = ownScale(a) ?? slotScale(owner, branch);
      stops.push({ x: slotBox ? slotBox.x : o ? o.x0 : 0, ...band(baseline, s), baseline, scale: s, side: "center" });
      continue;
    }
    const r = regions.get(a);
    if (!r) {
      // an atom without material (an empty group): where the previous stop is
      stops.push(stops[offset - 1] ? { ...stops[offset - 1]!, side: "center", empty: undefined } : null);
      continue;
    }
    const baseline = baselineOf(a);
    let x = r.x1;
    const sibs = siblingsOf(a);
    const next = sibs[sibs.indexOf(a) + 1];
    const nr = next ? regions.get(next) : undefined;
    if (next?.type === "subsup") {
      // before its scripts, also where TeX boxed them with it (an accented letter)
      if (nr && nr.x0 < x) x = nr.x0;
    } else if (a.type === "mbin" || a.type === "mrel") {
      // after a binary operator or a relation the caret stands past the
      // operator's space, at the next item
      if (nr && nr.x0 > x && baselineOf(next!) === baseline) x = nr.x0;
    }
    const s = scaleOf(a);
    stops.push({ x, ...band(baseline, s), baseline, scale: s, side: "center" });
  }

  separateCoincidingStops(model, stops);

  // the drawn box covers every marker and every caret position, so nothing is
  // clipped and the line keeps its height wherever the caret goes
  const empty = box.width === 0 && box.height === 0 && box.depth === 0;
  let height = empty ? caret.strut.height : box.height;
  let depth = empty ? caret.strut.depth : box.depth;
  let right = 0;
  for (const m of markers.values()) {
    height = Math.max(height, -m.y0);
    depth = Math.max(depth, m.y1);
  }
  for (const c of stops) {
    if (!c) continue;
    height = Math.max(height, Math.ceil(-c.top));
    depth = Math.max(depth, Math.ceil(c.bottom));
    right = Math.max(right, Math.ceil(c.x + BAR_HALF * c.scale - box.width));
  }
  return { width: box.width, height, depth, right, stops, markers, regions, alignMarkers };
}

/**
 * Stops next to each other in key order whose carets would stand on one place
 * (the end of a radicand and the place after the radical, the place before an
 * overline and the start of its body): a key press between them would show
 * nothing moving. The stop outside the structure stands a little outside it;
 * at the start of the formula, where nothing is left of it, the stop inside
 * stands a little inside instead. Stops at one depth (an atom without width)
 * stay where they are.
 */
function separateCoincidingStops(model: GeometryModel, stops: (CaretGeometry | null)[]): void {
  const depthOf = (offset: number) => {
    let depth = 0;
    for (let a = model.at(offset); a?.parent; a = a.parent) depth++;
    return depth;
  };
  const together = (a: CaretGeometry | null, b: CaretGeometry | null) =>
    !!a && !!b && !a.empty && !b.empty && Math.abs(a.x - b.x) <= SAME_PLACE && Math.abs(a.top - b.top) <= SAME_PLACE;
  for (let k = 0; k < stops.length; ) {
    let end = k + 1;
    while (end < stops.length && together(stops[k], stops[end])) end++;
    const run = end - k;
    if (run > 1) {
      const depths = Array.from({ length: run }, (_, i) => depthOf(k + i));
      if (depths.some((d) => d !== depths[0])) {
        const first = stops[k]!;
        const gap = STOP_GAP * first.scale;
        // into a structure (deeper along the run): the stops before it move
        // left of the innermost one, unless that leaves the formula
        const inward = depths[run - 1] > depths[0];
        const shift = inward && first.x - (run - 1) * gap >= 0 ? -(run - 1) * gap : 0;
        for (let i = 0; i < run; i++) stops[k + i] = { ...stops[k + i]!, x: Math.round(first.x + shift + i * gap) };
      }
    }
    k = end;
  }
}

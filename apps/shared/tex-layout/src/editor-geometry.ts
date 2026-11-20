// Editor geometry (docs/editing-core/DESIGN.md §5): where every caret stop,
// empty-slot marker and atom of an edited formula is, in formula units (sp;
// x from the formula's left edge, y down from its baseline). The editing core
// draws, hit-tests and anchors popovers with these numbers; nothing measures
// the page.
import type { Tag } from "./nodes.ts";
import { isDifferentialSymbol, isIntegralSymbol, operandScope, operandTokenOfOperator, type OperandToken } from "./operand-scope.ts";
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
/** room below a region's box for the L-shaped caret's line (the CSS line is 0.05em thick, at least 1px) */
const LINE_ROOM = 0.06 * EM;

export interface Rect {
  x0: number;
  y0: number;
  x1: number;
  y1: number;
}

/** The part of a kernel atom the geometry reads. */
export interface GeometryAtom {
  type?: string;
  /** the atom's symbol (a letter, an operator character, a big operator's sign) */
  value?: string;
  parent?: GeometryAtom;
  parentBranch?: unknown;
  branch(name: unknown): GeometryAtom[] | undefined;
  branches?: string[];
  rows?: unknown[][][];
}

/**
 * The region a caret stop is in: the box of its slot's content, which AxMath's L-shaped caret is drawn with (the bar is
 * as tall as the region, the underline runs along its foot). The items' boxes (nested structures, accents and scripts
 * included) and a line's strut at the slot's size; x from the first item's left edge to the last item's right edge.
 */
export interface CaretRegion {
  x0: number;
  x1: number;
  top: number;
  bottom: number;
}

export interface CaretGeometry {
  /** the boundary the caret stands at */
  x: number;
  top: number;
  bottom: number;
  /** the slot's content box (not for the stop of an empty slot, whose marker is its box) */
  region?: CaretRegion;
  /** the caret stands after an item of its slot: the underline under the whole region is drawn (not at the slot's start) */
  underline?: boolean;
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
  /**
   * The ways the caret can stand at this one place, innermost first (a stop is a kernel offset, and the end of a big
   * operator's operand is also the end of the structures it stands in): [0] inside the operand at its end (the underline
   * under the operand only), the last outside every operand that ends here (under the whole row). Key presses step
   * through them (the editing core's level). Only at the end of an operand; the stop itself is the first of them.
   */
  levels?: CaretGeometry[];
}

/**
 * How far the whole caret reaches vertically (sp, y down from the baseline): the bar as tall as its region and, with the
 * underline, the line below the region's box. Where popovers and the IME's window stand below the caret.
 */
export function caretReach(c: CaretGeometry): { top: number; bottom: number } {
  if (!c.region) return { top: c.top, bottom: c.bottom };
  return { top: Math.min(c.top, c.region.top), bottom: Math.max(c.bottom, c.region.bottom + (c.underline ? LINE_ROOM * c.scale : 0)) };
}

/** The caret geometry at an offset in one of the ways the caret can stand there (CaretGeometry.levels; 0: the first). */
export function stopAt(stops: readonly (CaretGeometry | null)[], offset: number, level = 0): CaretGeometry | null {
  const stop = stops[offset] ?? null;
  if (!stop?.levels) return stop;
  return stop.levels[Math.min(Math.max(level, 0), stop.levels.length - 1)];
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
  /** the formula's pieces and their places, when it is set on several lines (LayoutOptions.lineLayout) */
  lines?: import("./lines.ts").FormulaLines;
}

export interface GeometryModel {
  at(offset: number): GeometryAtom | undefined;
  lastOffset: number;
}

const isPlaceholder = (a: GeometryAtom | undefined) => a?.type === "placeholder" || a?.type === "prompt";

function siblingsOf(a: GeometryAtom): GeometryAtom[] {
  return a.parent?.branch(a.parentBranch) ?? [];
}

/** What an item of a row is to the operand rule (operand-scope.ts): a fence, relation, separator, sign, differential d, spacing, or other. */
function operandTokenOf(a: GeometryAtom, next: GeometryAtom | undefined): OperandToken {
  if (a.type === "space" || a.type === "spacing") return "space";
  if (a.branches?.length || a.rows) return "other";
  const text = typeof a.value === "string" ? a.value : "";
  if (isDifferentialSymbol(text)) return next && next.type === "mord" ? "differential" : "other";
  if (a.type === "mopen" || a.type === "mclose" || a.type === "mrel" || a.type === "mbin" || a.type === "mpunct" || a.type === "mord") return operandTokenOfOperator(text);
  return "other";
}

/**
 * A big operator's operand (∫ ∑ ∏ ...): the items after the operator that it applies to, by the one rule of operand-scope.ts.
 * AxMath's operator has a slot for it, and the caret in it has its own underline under the operand only; the formula's
 * LaTeX has none (the items follow the operator in its row), so the operand is found here, from the row.
 */
interface OperandGroup {
  op: GeometryAtom;
  /** the operand's items, in order (at least one has material) */
  items: GeometryAtom[];
  owner: GeometryAtom;
  branch: unknown;
}

function operandGroups(model: GeometryModel, hasMaterial: (a: GeometryAtom) => boolean): OperandGroup[] {
  const out: OperandGroup[] = [];
  for (let offset = 0; offset <= model.lastOffset; offset++) {
    const first = model.at(offset);
    if (first?.type !== "first" || !first.parent) continue;
    const owner = first.parent;
    const items = (owner.branch(first.parentBranch) ?? []).filter((x) => x.type !== "first");
    if (!items.some((x) => x.type === "extensible-symbol")) continue;
    const tokens = items.map((x, i) => operandTokenOf(x, items[i + 1]));
    // [from, to): the operand's items in the row, by index
    const found: { op: GeometryAtom; from: number; to: number }[] = [];
    items.forEach((op, index) => {
      if (op.type === "extensible-symbol") found.push({ op, from: index + 1, to: operandScope(tokens, index + 1, isIntegralSymbol(op.value ?? "")).end });
    });
    // an operator inside another's operand cannot reach past its end (∑ ∫ f + g dx: the sum ends at the plus)
    for (const outer of found) {
      for (const inner of found) {
        if (inner !== outer && inner.from > outer.from && inner.from <= outer.to && inner.to > outer.to) inner.to = Math.max(inner.from, outer.to);
      }
    }
    for (const g of found) {
      const operand = items.slice(g.from, g.to);
      if (operand.some(hasMaterial)) out.push({ op: g.op, items: operand, owner, branch: first.parentBranch });
    }
  }
  return out;
}

export function buildEditorGeometry(
  model: GeometryModel,
  regions: Map<Tag, Region>,
  ship: ShipResult,
  box: { width: number; height: number; depth: number },
  alignMarkers: number[],
  scales?: Map<Tag, number>,
  caret: CaretMetrics = CM_CARET,
  /** the caret stops carry their regions (the L-shaped caret); not for a formula set on several lines */
  withRegions = true,
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

  // the box of a slot's content: its items' regions and a line's strut at the slot's size (AxMath's chain box: a plain
  // line is as tall as a line, whatever its letters), by owner and branch
  const regionCache = new Map<GeometryAtom, Map<string, CaretRegion | null>>();
  const regionOf = (owner: GeometryAtom | undefined, branch: unknown): CaretRegion | null => {
    if (!owner || !withRegions) return null;
    const key = String(branch);
    let byBranch = regionCache.get(owner);
    if (!byBranch) regionCache.set(owner, (byBranch = new Map()));
    const cached = byBranch.get(key);
    if (cached !== undefined) return cached;
    let x0 = Infinity, x1 = -Infinity, y0 = Infinity, y1 = -Infinity;
    let line: number | undefined;
    for (const item of owner.branch(branch) ?? []) {
      if (item.type === "first" || isPlaceholder(item)) continue;
      const r = regions.get(item);
      if (!r) continue;
      x0 = Math.min(x0, r.x0);
      x1 = Math.max(x1, r.x1);
      y0 = Math.min(y0, r.y0);
      y1 = Math.max(y1, r.y1);
      line ??= baselineOf(item);
    }
    let found: CaretRegion | null = null;
    if (line !== undefined && x1 >= x0) {
      const s = slotScale(owner, branch);
      found = { x0, x1, top: Math.min(y0, line - caret.strut.height * s), bottom: Math.max(y1, line + caret.strut.depth * s) };
    }
    byBranch.set(key, found);
    return found;
  };

  // the box of a big operator's operand, as of a slot's content
  const groupRegions = new Map<OperandGroup, CaretRegion | null>();
  const regionOfGroup = (g: OperandGroup): CaretRegion | null => {
    if (groupRegions.has(g)) return groupRegions.get(g)!;
    let x0 = Infinity, x1 = -Infinity, y0 = Infinity, y1 = -Infinity;
    let line: number | undefined;
    for (const item of g.items) {
      if (isPlaceholder(item)) continue;
      const r = regions.get(item);
      if (!r) continue;
      x0 = Math.min(x0, r.x0);
      x1 = Math.max(x1, r.x1);
      y0 = Math.min(y0, r.y0);
      y1 = Math.max(y1, r.y1);
      line ??= baselineOf(item);
    }
    let found: CaretRegion | null = null;
    if (line !== undefined && x1 >= x0) {
      const s = slotScale(g.owner, g.branch);
      found = { x0, x1, top: Math.min(y0, line - caret.strut.height * s), bottom: Math.max(y1, line + caret.strut.depth * s) };
    }
    groupRegions.set(g, found);
    return found;
  };

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
        const region = regionOf(owner, branch);
        stops.push({ x: regions.get(lead)!.x0, ...band(baseline, s), baseline, scale: s, side: "center", ...(region ? { region } : {}) });
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
      stops.push(stops[offset - 1] ? { ...stops[offset - 1]!, side: "center", empty: undefined, ...(stops[offset - 1]!.region ? { underline: true } : {}) } : null);
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
    const region = regionOf(owner, branch);
    stops.push({ x, ...band(baseline, s), baseline, scale: s, side: "center", ...(region ? { region, underline: true } : {}) });
  }

  separateCoincidingStops(model, stops);

  // A big operator's operand has its own region (AxMath: the underline under only the operand while the caret is in it):
  // the stops inside it (and the one right after the operator, its start) are in the operand; at its end the caret
  // can stand inside the operand or outside it (and outside every operand that ends there), at one place: one more
  // way to stand for each, innermost first, a little further right each (the stops' separation at a structure's end).
  if (withRegions) {
    const groups = operandGroups(model, (a) => !isPlaceholder(a) && regions.has(a));
    const holders = new Map<GeometryAtom, OperandGroup[]>();
    for (const g of groups) for (const a of [g.op, ...g.items]) (holders.get(a) ?? holders.set(a, []).get(a)!).push(g);
    for (let offset = 0; offset <= model.lastOffset; offset++) {
      const a = model.at(offset);
      const base = stops[offset];
      const held = a ? holders.get(a) : undefined;
      if (!a || !base || !held) continue;
      // innermost first (the operator that stands latest in the row holds the others' items)
      const chain = [...held].reverse().filter((g) => regionOfGroup(g));
      if (!chain.length) continue;
      const ending = chain.filter((g) => g.items[g.items.length - 1] === a).length;
      const states: CaretGeometry[] = [];
      for (let j = 0; j <= ending; j++) {
        const g = chain[j];
        const gap = STOP_GAP * base.scale;
        const state: CaretGeometry = { ...base, x: Math.round(base.x + j * gap) };
        if (g) {
          state.region = regionOfGroup(g)!;
          state.underline = a !== g.op;
        }
        states.push(state);
      }
      stops[offset] = states.length > 1 ? { ...states[0], levels: states } : states[0];
    }
  }

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
  for (const stop of stops) {
    if (!stop) continue;
    for (const c of stop.levels ?? [stop]) {
      // (the L-shaped caret reaches the foot of its region and the line below it)
      const reach = caretReach(c);
      height = Math.max(height, Math.ceil(-reach.top));
      depth = Math.max(depth, Math.ceil(reach.bottom));
      right = Math.max(right, Math.ceil(c.x + BAR_HALF * c.scale - box.width));
    }
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

// Selection (docs/editing-core/DESIGN.md §6.9). A selection is a run of
// siblings in one slot: extending it past the slot's edge selects the whole
// structure (a script with its base, a macro as a whole). Shift+arrows,
// Shift+click, mouse drags and double/triple clicks all go through select().
import { goTo } from "./navigation.ts";
import { allowedStops, type CoreModel } from "./stops.ts";
import {
  enclosingMacro,
  isMacroInternalSlot,
  isPlaceholder,
  sameKey,
  slotAtoms,
  slotOf,
  type CoreAtom,
  type Slot,
} from "./structures.ts";

type Direction = -1 | 1;

/** A kernel selection: setSelection(anchor, focus). */
export interface SelectionRange {
  anchor: number;
  focus: number;
}

interface Level {
  atom: CoreAtom;
  slot: Slot;
}

/** An atom and its ancestors, each with the slot it is an item of (innermost first). */
function levels(atom: CoreAtom): Level[] {
  const out: Level[] = [];
  for (let a: CoreAtom | undefined = atom; a; a = a.parent) {
    const slot = slotOf(a);
    if (!slot) break;
    out.push({ atom: a, slot });
  }
  return out;
}

/**
 * The selection within one slot between the anchor side and the focus side:
 * a side is a stop of the slot itself (direct) or the item holding the stop.
 * Boundaries are indexes into the slot's atoms (k = after atom k, 0 = the
 * slot's start), so an item k spans k-1..k.
 */
function inSlot(model: CoreModel, slot: Slot, itemA: CoreAtom, directA: boolean, itemF: CoreAtom, directF: boolean, forward: boolean): SelectionRange {
  // a macro's expansion is not a slot of its own: the macro is the item
  if (isMacroInternalSlot(slot)) {
    const macro = slot.owner.type === "macro" ? slot.owner : enclosingMacro(slot.owner);
    if (macro) return whole(model, macro, forward);
  }
  const atoms = slotAtoms(slot);
  const ia = atoms.indexOf(itemA);
  const jf = atoms.indexOf(itemF);
  const lo = Math.min(directA ? ia : ia - 1, directF ? jf : jf - 1);
  let hi = Math.max(ia, jf);
  // a base takes its scripts along (the scripts alone can be selected)
  if (hi > lo && atoms[hi + 1]?.type === "subsup") hi += 1;
  const start = model.offsetOf(atoms[lo]);
  const end = model.offsetOf(atoms[hi]);
  return forward ? { anchor: start, focus: end } : { anchor: end, focus: start };
}

/** The selection of one whole item. */
function whole(model: CoreModel, item: CoreAtom, forward: boolean): SelectionRange {
  const slot = slotOf(item);
  if (!slot) return { anchor: 0, focus: 0 };
  return inSlot(model, slot, item, false, item, false, forward);
}

/** A run of matrix cells: each end stays in its cell, at the edge of the item holding it. */
function acrossCells(model: CoreModel, a: Level, directA: boolean, anchor: number, f: Level, directF: boolean, focus: number): SelectionRange {
  const forward = anchor < focus;
  const edge = (level: Level, direct: boolean, stop: number, left: boolean) => {
    if (direct) return stop;
    return left && level.atom.leftSibling ? model.offsetOf(level.atom.leftSibling) : model.offsetOf(level.atom);
  };
  return { anchor: edge(a, directA, anchor, forward), focus: edge(f, directF, focus, !forward) };
}

/** The selection between two stops: the run of siblings of their innermost common slot that holds both. */
export function selectionRange(model: CoreModel, anchor: number, focus: number): SelectionRange {
  const a = model.at(anchor);
  const f = model.at(focus);
  if (!a || !f || anchor === focus) return { anchor: focus, focus };
  const la = levels(a);
  const lf = levels(f);
  const forward = anchor < focus;
  for (let i = 0; i < la.length; i++) {
    for (let j = 0; j < lf.length; j++) {
      const owner = la[i].slot.owner;
      if (owner !== lf[j].slot.owner) continue;
      const keyA = la[i].slot.key;
      const keyF = lf[j].slot.key;
      if (sameKey(keyA, keyF)) return inSlot(model, la[i].slot, la[i].atom, i === 0, lf[j].atom, j === 0, forward);
      if (owner.rows && Array.isArray(keyA) && Array.isArray(keyF)) return acrossCells(model, la[i], i === 0, anchor, lf[j], j === 0, focus);
      // two slots of one structure (numerator and denominator, ...): the structure
      return whole(model, owner, forward);
    }
  }
  return { anchor, focus };
}

// The stop a selection was started from, kept while the selection is the one
// set here: a selection that grew to a whole structure shrinks back into the
// slot it started in.
const started = new WeakMap<CoreModel, { root: CoreAtom; anchor: number; lo: number; hi: number }>();

function currentRange(model: CoreModel): [number, number] {
  const r = model.selection.ranges[0] ?? [model.position, model.position];
  return [Math.min(r[0], r[1]), Math.max(r[0], r[1])];
}

/** The stop the current selection grows from. */
export function selectionAnchor(model: CoreModel): number {
  if (model.selectionIsPlaceholder) return Math.max(model.position, model.anchor) - 1;
  if (model.selectionIsCollapsed) return model.position;
  const memo = started.get(model);
  const [lo, hi] = currentRange(model);
  if (memo && memo.root === model.root && memo.lo === lo && memo.hi === hi && memo.anchor <= model.lastOffset) return memo.anchor;
  return model.anchor;
}

/** Select from one stop to another, widened to a run of siblings. */
export function select(model: CoreModel, anchor: number, focus: number): boolean {
  const range = selectionRange(model, anchor, focus);
  if (range.anchor === range.focus) {
    started.delete(model);
    return goTo(model, range.focus);
  }
  const previous = model.position;
  model.setSelection(range.anchor, range.focus);
  const [lo, hi] = currentRange(model);
  started.set(model, { root: model.root, anchor, lo, hi });
  model.mathfield.stopCoalescingUndo();
  model.announce("move", previous);
  return true;
}

/** Shift+Left/Right: the focus moves to the next stop that changes the selection. */
export function extend(model: CoreModel, direction: Direction): boolean {
  const anchor = selectionAnchor(model);
  const ranged = !model.selectionIsCollapsed && !model.selectionIsPlaceholder;
  const focus = ranged ? model.position : anchor;
  const [lo, hi] = ranged ? currentRange(model) : [anchor, anchor];
  const stops = allowedStops(model);
  const ahead = direction > 0 ? stops.filter((s) => s > focus) : stops.filter((s) => s < focus).reverse();
  for (const stop of ahead) {
    const range = selectionRange(model, anchor, stop);
    if (Math.min(range.anchor, range.focus) !== lo || Math.max(range.anchor, range.focus) !== hi) return select(model, anchor, stop);
  }
  model.announce("plonk");
  return false;
}

/** Extend the selection from its anchor to a stop (Shift+click, drags, Shift+Home/End, word steps). */
export function extendTo(model: CoreModel, focus: number): boolean {
  return select(model, selectionAnchor(model), focus);
}

const WORD_CHARACTER = /^[0-9\p{Letter}]$/u;

/**
 * Select an item: a run of text letters and digits, a number, or the item
 * itself (a math letter is a variable, not part of a word; a structure is
 * selected whole).
 */
export function selectItem(model: CoreModel, item: CoreAtom): boolean {
  const before = item.leftSibling;
  if (!before) return false;
  if (isPlaceholder(item)) return goTo(model, model.offsetOf(before));
  const inRun = (x: CoreAtom | undefined): x is CoreAtom => {
    if (!x || x.mode !== item.mode || x._branches || x.rows) return false;
    return item.mode === "text" ? WORD_CHARACTER.test(x.value ?? "") : !!x.isDigit?.();
  };
  let first = item;
  let last = item;
  if (inRun(item)) {
    while (inRun(first.leftSibling)) first = first.leftSibling;
    while (inRun(last.rightSibling)) last = last.rightSibling;
  }
  return select(model, model.offsetOf(first.leftSibling ?? before), model.offsetOf(last));
}

/** Select the whole slot a stop is in (an empty slot: its stop). */
export function selectSlot(model: CoreModel, stop: number): boolean {
  const atom = model.at(stop);
  const slot = atom ? slotOf(atom) : null;
  if (!slot) return false;
  const atoms = slotAtoms(slot);
  const start = model.offsetOf(atoms[0]);
  const items = atoms.slice(1);
  if (items.length === 0 || (items.length === 1 && isPlaceholder(items[0]))) return goTo(model, start);
  return select(model, start, model.offsetOf(atoms[atoms.length - 1]));
}

/** The selectGroup command: the item at the caret. */
export function selectGroup(model: CoreModel): boolean {
  let atom = model.at(model.selectionIsCollapsed ? model.position : Math.min(model.position, model.anchor) + 1);
  if (atom?.type === "first") atom = atom.rightSibling;
  if (!atom) {
    model.announce("plonk");
    return false;
  }
  return selectItem(model, atom);
}

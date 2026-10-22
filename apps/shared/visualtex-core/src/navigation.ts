// Caret movement (docs/editing-core/DESIGN.md §6.1-6.4): Left/Right, Home/End,
// Tab/Shift+Tab and Up/Down, all over the allowed stops of stops.ts and the
// slot order of structures.ts.
import { extendTo } from "./selection.ts";
import { besideTextRun } from "./text.ts";
import { allowedStops, isAllowedStop, type CoreModel } from "./stops.ts";
import {
  isEmptySlot,
  isPlaceholder,
  logicalSlots,
  logicalStructure,
  sameKey,
  slotAtoms,
  slotItems,
  slotOf,
  type CoreAtom,
  type FractionOrder,
  type Slot,
} from "./structures.ts";

export interface EditingHost {
  /** Up/Down found nothing to move to inside the formula: the app may move to the previous/next row. */
  verticalExit?(direction: -1 | 1): boolean;
  /**
   * The app's key commands (rows, hotkeys, history, ...), in order. The field's key entry offers each keydown to
   * them before the kernel handles it; a command returns true when it handled the key and ended its propagation.
   */
  keyCommands?: readonly ((event: KeyboardEvent) => boolean)[];
}

export interface NavigationContext {
  fractionOrder: FractionOrder;
  host: EditingHost | undefined;
}

/** Kernel operations a Tab may need (the kernel owns the array model). */
export interface TabKernel {
  /** Add a row after the caret's row of the array the caret is in; true if added. */
  addRowAfter(): boolean;
}

type Direction = -1 | 1;

/** The stop the caret moves from: the slot start in an empty slot, the moving edge of a range. */
function originStop(model: CoreModel, direction: Direction): number {
  if (model.selectionIsCollapsed) return model.position;
  if (model.selectionIsPlaceholder) {
    const placeholder = Math.max(model.position, model.anchor);
    return placeholder - 1;
  }
  return direction > 0 ? Math.max(model.position, model.anchor) : Math.min(model.position, model.anchor);
}

/** Put the caret at a stop; a placeholder right after it is selected (the caret in its empty slot). */
export function goTo(model: CoreModel, offset: number): boolean {
  const previous = model.position;
  const next = model.at(offset)?.rightSibling;
  if (isPlaceholder(next)) model.setSelection(offset, offset + 1);
  else model.setSelection(offset, offset);
  model.mathfield.stopCoalescingUndo();
  model.announce("move", previous);
  return true;
}

/**
 * Enter a kernel offset as the caret: the offset of a placeholder means the
 * placeholder's empty slot (one stop), which is shown selected.
 */
export function goToStop(model: CoreModel, offset: number): boolean {
  const atom = model.at(offset);
  if (isPlaceholder(atom) && offset > 0) return goTo(model, offset - 1);
  return goTo(model, offset);
}

/** Where the caret goes after inserting `atoms`: their first empty slot, in keyboard order. */
export function firstEmptySlotStop(model: CoreModel, atoms: readonly CoreAtom[]): number | undefined {
  return emptySlotStops(model, [...atoms])[0];
}

function plonk(model: CoreModel): boolean {
  model.announce("plonk");
  return true;
}

// ------------------------------------------------------------- Left/Right --

export function moveHorizontal(model: CoreModel, direction: Direction): boolean {
  // A range collapses to its edge first (an empty slot is not a range).
  if (!model.selectionIsCollapsed && !model.selectionIsPlaceholder) {
    const edge = originStop(model, direction);
    if (isAllowedStop(model, edge)) return goTo(model, edge);
  }
  const stops = allowedStops(model);
  const from = originStop(model, direction);
  const next = direction > 0 ? stops.find((s) => s > from) : [...stops].reverse().find((s) => s < from);
  return next === undefined ? plonk(model) : goTo(model, next);
}

// --------------------------------------------------------------- Home/End --

/** Start (End: end) of the caret's slot; at that edge already, of the whole formula. */
export function moveToSlotEdge(model: CoreModel, direction: Direction, thenFormula: boolean): boolean {
  const from = originStop(model, direction);
  const atom = model.at(from);
  const slot = atom ? slotOf(atom) : null;
  const stops = allowedStops(model);
  if (slot) {
    const atoms = slotAtoms(slot);
    const edgeAtom = direction < 0 ? atoms[0] : atoms[atoms.length - 1];
    const edge = edgeAtom ? model.offsetOf(edgeAtom) : -1;
    if (edge >= 0 && edge !== from && isAllowedStop(model, edge)) return goTo(model, edge);
  }
  if (!thenFormula) return plonk(model);
  const target = direction < 0 ? stops[0] : stops[stops.length - 1];
  return target === undefined || target === from ? plonk(model) : goTo(model, target);
}

// -------------------------------------------------------------------- Tab --

/** The stops at which Tab stops in an empty slot within `atoms` (in keyboard order). */
function emptySlotStops(model: CoreModel, atoms: CoreAtom[]): number[] {
  const out: number[] = [];
  const visit = (a: CoreAtom) => {
    for (const d of [...a.children, a]) {
      if (isPlaceholder(d)) {
        const before = d.leftSibling ?? null;
        if (before) {
          const offset = model.offsetOf(before);
          if (isAllowedStop(model, offset)) out.push(offset);
        }
      } else if (d.type === "first" && d.isLastSibling && d.parent?.parent) {
        // a branch holding nothing at all
        const offset = model.offsetOf(d);
        if (isAllowedStop(model, offset)) out.push(offset);
      }
    }
  };
  for (const a of atoms) visit(a);
  return [...new Set(out)];
}

/** Enter a slot: at its first (Shift+Tab: last) empty slot, else at its end. */
export function enterSlot(model: CoreModel, slot: Slot, direction: Direction): boolean {
  const items = slotItems(slot);
  const empties = emptySlotStops(model, items);
  if (isEmptySlot(slot)) {
    const first = slotAtoms(slot)[0];
    return first ? goTo(model, model.offsetOf(first)) : plonk(model);
  }
  const target = direction > 0 ? empties[0] : empties[empties.length - 1];
  if (target !== undefined) return goTo(model, target);
  const last = items[items.length - 1];
  return goTo(model, model.offsetOf(last));
}

export function tab(model: CoreModel, direction: Direction, context: NavigationContext, kernel: TabKernel): boolean {
  if (model.mode === "latex") return false;
  const from = originStop(model, direction);
  const atom = model.at(from);
  if (!atom) return plonk(model);
  // Text inside a formula is a run of siblings, not a structure: first go to
  // the end (Shift+Tab: start) of the run, then continue as math.
  if (atom.mode === "text") {
    let edge: CoreAtom = atom;
    if (direction > 0) while (edge.rightSibling && edge.rightSibling.mode === "text") edge = edge.rightSibling;
    else while (edge.mode === "text" && edge.leftSibling) edge = edge.leftSibling;
    if (edge !== atom || !model.selectionIsCollapsed) {
      goTo(model, model.offsetOf(edge));
      model.mathfield.switchMode("math");
      return true;
    }
    model.mathfield.switchMode("math");
  }
  const slot = slotOf(atom);
  if (!slot) return plonk(model);
  const structure = logicalStructure(slot);
  // 1. An empty slot still ahead (behind) in the current slot.
  const items = slotItems(slot);
  const index = atom.type === "first" ? -1 : items.indexOf(atom);
  const ahead = direction > 0 ? items.slice(index + 1) : items.slice(0, index + 1);
  const empties = emptySlotStops(model, ahead).filter((s) => s !== from);
  const near = direction > 0 ? empties[0] : empties[empties.length - 1];
  if (near !== undefined) return goTo(model, near);
  // 2. The next (previous) slot of the same structure.
  const slots = structure.type === "root" ? [] : logicalSlots(structure, context.fractionOrder);
  const current = slots.findIndex((s) => s.owner === slot.owner && sameKey(s.key, slot.key));
  const next = current < 0 ? undefined : slots[current + direction];
  const inLastCell = !!structure.rows && Array.isArray(slot.key) && next === undefined && direction > 0;
  if (next) return enterSlot(model, next, direction);
  // 3. The last cell of a matrix, cases or aligned block: a new row (the
  // kernel puts the caret in it).
  if (inLastCell && structure.environmentName !== "lines") {
    if (kernel.addRowAfter()) return true;
  }
  // 4. Out of the structure, one level. The top level has nowhere to go.
  if (!structure.parent || structure.type === "root") return plonk(model);
  // (after it; a placeholder right after it is selected by goTo)
  if (direction > 0) return goTo(model, model.offsetOf(structure));
  const before = structure.leftSibling;
  return before ? goTo(model, model.offsetOf(before)) : plonk(model);
}

// ------------------------------------------------------------------ Space --

/** Kernel operations Space may need. */
export interface SpaceKernel {
  /** Insert a visible space (\ ) at the caret. */
  insertVisibleSpace(): boolean;
  /** Insert a text-mode space at the caret (it joins the text run beside it). */
  insertTextSpace(): boolean;
}

/**
 * Space in math (after the input conversion had its turn): beside a text run
 * it is a word space of that text; inside a structure it leaves the structure
 * (a macro such as \ket as a whole); at the top level it inserts a visible
 * space.
 */
export function space(model: CoreModel, kernel: SpaceKernel): boolean {
  if (model.selectionIsCollapsed && besideTextRun(model, model.position)) return kernel.insertTextSpace();
  const atom = model.at(originStop(model, 1));
  const slot = atom ? slotOf(atom) : null;
  const structure = slot ? logicalStructure(slot) : null;
  if (!structure || structure.type === "root" || !structure.parent) return kernel.insertVisibleSpace();
  return goTo(model, model.offsetOf(structure));
}

// ---------------------------------------------------------------- Up/Down --

function stopsInBranch(model: CoreModel, owner: CoreAtom, key: Slot["key"]): number[] {
  const atoms = slotAtoms({ owner, key });
  if (!atoms.length) return [];
  const first = model.offsetOf(atoms[0]);
  const lastAtom = atoms[atoms.length - 1];
  // the last atom's offset comes after all its descendants (post-order)
  const last = model.offsetOf(lastAtom);
  const out: number[] = [];
  for (let offset = first; offset <= last; offset++) {
    // keep offsets inside this branch (not after it)
    let a: CoreAtom | undefined = model.at(offset);
    while (a && a.parent !== owner) a = a.parent;
    if (!a || !sameKey(a.parentBranch, key)) continue;
    if (isAllowedStop(model, offset)) out.push(offset);
  }
  return out;
}

/** The stop of a branch nearest the caret horizontally (the editor's geometry). */
function nearestStop(model: CoreModel, candidates: number[], fromX: number | null): number | null {
  if (!candidates.length) return null;
  const geometry = model.mathfield.visualTexGeometry;
  if (fromX === null || !geometry) return candidates[0];
  let best = candidates[0];
  let bestDistance = Infinity;
  for (const offset of candidates) {
    const c = geometry.stops[offset];
    if (!c) continue;
    const distance = Math.abs(c.x - fromX);
    if (distance < bestDistance) {
      best = offset;
      bestDistance = distance;
    }
  }
  return best;
}

export function vertical(model: CoreModel, direction: Direction, context: NavigationContext, extend = false): boolean {
  const from = extend ? model.position : originStop(model, direction);
  const source = model.at(from);
  if (!source) return plonk(model);
  const fromX = model.mathfield.visualTexGeometry?.stops[from]?.x ?? null;
  const go = (offset: number) => (extend ? extendTo(model, offset) : goTo(model, offset));
  const tryBranch = (owner: CoreAtom | undefined, key: Slot["key"] | undefined): boolean => {
    if (!owner || key === undefined) return false;
    const target = nearestStop(model, stopsInBranch(model, owner, key), fromX);
    if (target === null || target === from) return false;
    return go(target);
  };
  // From the innermost structure outward: a script inside a numerator visits
  // its sibling script before the denominator.
  for (let child: CoreAtom | undefined = source; child?.parent; child = child.parent) {
    const owner = child.parent;
    const branch = child.parentBranch;
    if (owner.rows && Array.isArray(branch)) {
      const row = branch[0] + direction;
      if (row >= 0 && row < owner.rows.length && tryBranch(owner, [row, branch[1]])) return true;
    } else if (branch === "superscript" || branch === "subscript") {
      if (branch === "superscript" && direction > 0 && owner.branch("subscript") && tryBranch(owner, "subscript")) return true;
      if (branch === "subscript" && direction < 0 && owner.branch("superscript") && tryBranch(owner, "superscript")) return true;
      // braces and other annotated over/under atoms: the label and the base
      if (owner.type === "overunder" && ((branch === "superscript" && direction > 0) || (branch === "subscript" && direction < 0)) && tryBranch(owner, "body")) return true;
      // a lone script returns to its baseline (it never creates the other script)
      if (owner.type === "subsup" && owner.leftSibling && ((branch === "superscript" && direction > 0) || (branch === "subscript" && direction < 0))) {
        const baseline = model.offsetOf(owner.leftSibling);
        if (baseline >= 0 && baseline !== from && isAllowedStop(model, baseline)) return go(baseline);
      }
    } else if (branch === "above" && direction > 0) {
      if ((owner.type === "overunder" || owner.type === "surd") && tryBranch(owner, "body")) return true;
      if (owner.branch("below") && tryBranch(owner, "below")) return true;
    } else if (branch === "below" && direction < 0) {
      if (owner.type === "overunder" && owner.branch("body") && tryBranch(owner, "body")) return true;
      if (owner.branch("above") && tryBranch(owner, "above")) return true;
    } else if (branch === "body") {
      const attached = child.rightSibling;
      if (attached?.type === "subsup" && tryBranch(attached, direction < 0 ? "superscript" : "subscript")) return true;
      if (owner.type === "surd" && direction < 0 && tryBranch(owner, "above")) return true;
      if (owner.type === "overunder" && (tryBranch(owner, direction < 0 ? "above" : "below") || tryBranch(owner, direction < 0 ? "superscript" : "subscript"))) return true;
    }
  }
  // Nothing above (below) inside the formula: the app may move to another row.
  if (!extend && context.host?.verticalExit?.(direction)) return true;
  return plonk(model);
}

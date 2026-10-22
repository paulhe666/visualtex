// Caret stops (docs/editing-core/DESIGN.md §3). A stop is a kernel offset: the
// position after the atom at that offset (a slot's start is the offset of its
// "first" atom). This module is the one rule for which offsets the caret may
// stand at; keys, Tab, extension and (later) clicks only ever produce these.
import { inMacroInternals, isPlaceholder, slotItems, slotOf, type CoreAtom } from "./structures.ts";

/** The part of the kernel model the core reads and drives. */
export interface CoreModel {
  readonly lastOffset: number;
  position: number;
  readonly anchor: number;
  readonly selectionIsCollapsed: boolean;
  readonly selectionIsPlaceholder: boolean;
  readonly selection: { readonly ranges: readonly (readonly [number, number])[] };
  readonly root: CoreAtom;
  readonly mode: string;
  at(offset: number): CoreAtom | undefined;
  offsetOf(atom: CoreAtom): number;
  setSelection(anchor: number, focus: number): boolean | void;
  collapseSelection(direction: "forward" | "backward"): boolean;
  announce(command: string, previousPosition?: number): void;
  readonly mathfield: {
    readonly hasEditablePrompts?: boolean;
    readonly host?: Element | null;
    readonly visualTexGeometry?: import("../../tex-layout/src/index.ts").EditorGeometry;
    stopCoalescingUndo(): void;
    switchMode(mode: string): void;
  };
}

/** The start of an empty slot (its only item, if any, is a placeholder). */
function isEmptySlotStart(atom: CoreAtom): boolean {
  if (atom.type !== "first") return false;
  const slot = slotOf(atom);
  if (!slot) return false;
  const items = slotItems(slot);
  return items.length === 0 || (items.length === 1 && isPlaceholder(items[0]));
}

export function isAllowedStop(model: CoreModel, offset: number): boolean {
  const atom = model.at(offset);
  if (!atom) return false;
  // a placeholder is reached from the stop before it: an empty slot has one stop
  if (isPlaceholder(atom) || atom.type === "prompt") return false;
  // a macro (\ket, \braket, ...) is one structure: only its arguments have stops
  if (atom.type === "macro-argument" || inMacroInternals(atom)) return false;
  // structures that capture the selection (accents, boxes, ...) have no inner
  // stops, except an empty slot inside accents, which keys reach as clicks do
  const capturing: CoreAtom[] = [];
  for (let p = atom.parent; p; p = p.parent) if (p.captureSelection) capturing.push(p);
  if (capturing.length > 0 && !(capturing.every((p) => p.type === "accent") && isEmptySlotStart(atom))) return false;
  // structures that skip their inner boundaries (accents, ...) still have
  // the one stop of an empty slot
  if (atom.parent?.skipBoundary && !isEmptySlotStart(atom)) {
    if (!atom.isFirstSibling && atom.isLastSibling) return false;
    if (atom.type === "first") return false;
  }
  if (model.mathfield.hasEditablePrompts && !atom.parentPrompt) return false;
  return true;
}

/** Every allowed stop, in keyboard order (the kernel's offset order follows the core's slot order). */
export function allowedStops(model: CoreModel): number[] {
  const out: number[] = [];
  for (let offset = 0; offset <= model.lastOffset; offset++) if (isAllowedStop(model, offset)) out.push(offset);
  return out;
}

// Structures and slots (docs/editing-core/DESIGN.md §3-4). The keyboard order
// of every structure's slots is defined here and nowhere else: the kernel's
// atom order (and so arrows, Tab, Backspace and offsets) is derived from it.

export type SlotKey = string | [number, number];

/** The part of a kernel atom the core reads. */
export interface CoreAtom {
  type?: string;
  mode?: string;
  value?: string;
  command?: string;
  parent?: CoreAtom;
  parentBranch?: SlotKey;
  _branches?: Record<string, CoreAtom[] | undefined>;
  rows?: (CoreAtom[] | undefined)[][];
  environmentName?: string;
  captureSelection?: boolean;
  skipBoundary?: boolean;
  parentPrompt?: unknown;
  children: readonly CoreAtom[];
  leftSibling?: CoreAtom;
  rightSibling?: CoreAtom;
  isFirstSibling?: boolean;
  isLastSibling?: boolean;
  branch(name: SlotKey): CoreAtom[] | undefined;
  getCell?(row: number, column: number): CoreAtom[] | undefined;
  isDigit?(): boolean;
}

/** A slot: an editable list in a structure (or the formula's top level). */
export interface Slot {
  owner: CoreAtom;
  key: SlotKey;
}

export type FractionOrder = "numerator-denominator" | "denominator-numerator";

const GENERIC_ORDER = ["body", "above", "below", "subscript", "superscript"] as const;

/**
 * The keyboard order of an atom's slots: reading order, a lower script or
 * limit before the upper one, a root's index before its radicand, matrix cells
 * row by row, and fractions in the user's fraction order.
 */
export function slotOrder(atom: CoreAtom, fractionOrder: FractionOrder = "numerator-denominator"): SlotKey[] {
  const has = (name: string) => atom._branches?.[name] !== undefined;
  if (atom.rows) {
    // cells row by row, then any scripts attached to the matrix
    const keys: SlotKey[] = [];
    atom.rows.forEach((row, r) => row.forEach((cell, c) => cell && keys.push([r, c])));
    return [...keys, ...GENERIC_ORDER.filter(has)];
  }
  if (atom.type === "genfrac")
    return (fractionOrder === "denominator-numerator" ? ["below", "above"] : ["above", "below"]).filter(has);
  if (atom.type === "surd") return ["above", "body"].filter(has);
  // \overset{A}{B} and its kin: the arguments in the order they are written
  const argumentOrder = ARGUMENT_ORDER[atom.command ?? ""];
  if (atom.type === "overunder" && argumentOrder) return [...argumentOrder, ...GENERIC_ORDER.filter((b) => !argumentOrder.includes(b))].filter(has);
  return GENERIC_ORDER.filter(has);
}

/** Over/under commands whose slots are navigated in argument order. */
const ARGUMENT_ORDER: Record<string, string[]> = {
  "\\overset": ["above", "body"],
  "\\stackrel": ["above", "body"],
  "\\stackbin": ["above", "body"],
  "\\underset": ["below", "body"],
  "\\overunderset": ["above", "below", "body"],
};

export function sameKey(a: SlotKey | undefined, b: SlotKey | undefined): boolean {
  if (Array.isArray(a) && Array.isArray(b)) return a[0] === b[0] && a[1] === b[1];
  return a === b;
}

export function slotAtoms(slot: Slot): CoreAtom[] {
  return (Array.isArray(slot.key) ? slot.owner.getCell?.(slot.key[0], slot.key[1]) : slot.owner.branch(slot.key)) ?? [];
}

/** The items of a slot (without the kernel's leading "first" atom). */
export function slotItems(slot: Slot): CoreAtom[] {
  return slotAtoms(slot).filter((a) => a.type !== "first");
}

/** The slot an atom sits in. */
export function slotOf(atom: CoreAtom): Slot | null {
  return atom.parent && atom.parentBranch !== undefined ? { owner: atom.parent, key: atom.parentBranch } : null;
}

export const isPlaceholder = (a: CoreAtom | undefined): boolean => a?.type === "placeholder";

/** An empty slot: no items, or only a placeholder (its persistence form). */
export function isEmptySlot(slot: Slot): boolean {
  const items = slotItems(slot);
  return items.length === 0 || (items.length === 1 && isPlaceholder(items[0]));
}

/** The macro whose expansion contains `atom`, if any. */
export function enclosingMacro(atom: CoreAtom | undefined): CoreAtom | null {
  for (let p = atom?.parent; p; p = p.parent) if (p.type === "macro") return p;
  return null;
}

/** Whether `atom` lies in a macro's expansion outside every argument. */
export function inMacroInternals(atom: CoreAtom): boolean {
  if (!enclosingMacro(atom)) return false;
  for (let a: CoreAtom | undefined = atom; a && a.type !== "macro"; a = a.parent) if (a.parent?.type === "macro-argument") return false;
  return true;
}

/** Whether a slot is a macro's own expansion (not one of its arguments): the macro is one item there. */
export function isMacroInternalSlot(slot: Slot): boolean {
  return slot.owner.type === "macro" || (slot.owner.type !== "macro-argument" && inMacroInternals(slot.owner));
}

/**
 * The structure a slot belongs to for the user. A macro such as \ket is one
 * structure whose slots are its arguments; the rest of its expansion
 * (delimiters, wrappers) is not navigable.
 */
export function logicalStructure(slot: Slot): CoreAtom {
  if (slot.owner.type === "macro-argument") return enclosingMacro(slot.owner) ?? slot.owner;
  return slot.owner;
}

/** The slots of a structure, in keyboard order. */
export function logicalSlots(structure: CoreAtom, fractionOrder: FractionOrder): Slot[] {
  if (structure.type === "macro") {
    const slots: Slot[] = [];
    for (const a of structure.children) if (a.type === "macro-argument" && enclosingMacro(a) === structure) slots.push({ owner: a, key: "body" });
    return slots;
  }
  return slotOrder(structure, fractionOrder).map((key) => ({ owner: structure, key }));
}

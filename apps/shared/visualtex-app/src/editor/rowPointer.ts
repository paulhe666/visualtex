// The formula rows under the pointer (docs/editing-core/DESIGN.md §6.9,
// §6.12), AxMath's rules (research: claude-work/axmath-mouse-reverse,
// selection-caret.md §2.1, §4.2): the row a point means is the first one whose
// bottom is at or below it (the space above a row is that row's; below the
// last row there is none, and a drag there changes nothing); in the row, the
// formula's own hit test gives the stop. A selection across rows takes whole
// top-level items: an end inside a structure takes the whole structure (the
// first row from before it, the last row up to after it).
import type { MathfieldElement } from "mathlive";
import { editingModelOf, visualTexEditingCore } from "./visualTexLayoutRuntime";

/**
 * AxMath's row rule: the index of the first row whose bottom is at or below
 * the point (rows stack top to bottom, so it is a binary search over their
 * bottoms); null below every row.
 */
export function rowIndexAt(count: number, bottomOf: (index: number) => number, clientY: number): number | null {
  let low = 0;
  let high = count;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (clientY <= bottomOf(middle)) high = middle;
    else low = middle + 1;
  }
  return low < count ? low : null;
}

type KernelField = Parameters<typeof visualTexEditingCore.hitTest>[0];

/** The stop a page point means in a row's formula (beside the formula: its start or end). */
export function stopAtPoint(field: MathfieldElement, clientX: number, clientY: number): number | null {
  const mathfield = (field as unknown as { _mathfield?: KernelField })._mathfield;
  return mathfield ? visualTexEditingCore.hitTest(mathfield, clientX, clientY) : null;
}

/**
 * An end of a selection across rows at the top level of its formula: a stop
 * inside a structure becomes the stop before the structure (the selection's
 * start) or after it (its end), so the selection holds the whole structure.
 */
export function topLevelStop(field: MathfieldElement, offset: number, side: "start" | "end"): number {
  const model = editingModelOf(field);
  let atom = model?.at(offset);
  if (!model || !atom || atom.parent === model.root) return offset;
  while (atom.parent && atom.parent !== model.root) atom = atom.parent;
  if (side === "end") return model.offsetOf(atom);
  return atom.leftSibling ? model.offsetOf(atom.leftSibling) : 0;
}

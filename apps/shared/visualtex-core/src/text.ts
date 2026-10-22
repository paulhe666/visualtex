// Text runs (docs/editing-core/DESIGN.md §6.6). \text{...} is a run of
// text-mode siblings. The caret is in the text only between two of its
// characters; at either edge it is in math. This is the one rule for the
// input mode after a caret move, a setValue, a paste and an IME composition,
// and for Space beside text.
import type { CoreModel } from "./stops.ts";

/** Whether the caret at a stop stands inside a text run: between two of its characters. */
export function inTextRun(model: CoreModel, offset: number): boolean {
  const left = model.at(offset);
  return left?.mode === "text" && left.rightSibling?.mode === "text";
}

/** The input mode at a stop: text inside a text run, else the mode of the material there (math, or latex in the LaTeX box). */
export function caretMode(model: CoreModel, offset: number): string {
  const left = model.at(offset);
  let mode = left?.mode;
  for (let p = left?.parent; !mode && p; p = p.parent) mode = p.mode;
  if (mode === "text") return inTextRun(model, offset) ? "text" : "math";
  return mode ?? "math";
}

/** Whether the caret at a stop touches a text run (at its start, inside or at its end). */
export function besideTextRun(model: CoreModel, offset: number): boolean {
  const left = model.at(offset);
  return left?.mode === "text" || left?.rightSibling?.mode === "text";
}

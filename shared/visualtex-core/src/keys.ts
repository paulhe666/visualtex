// Enter and Escape (docs/editing-core/DESIGN.md §6.11): what these keys do
// at the caret, decided in one place. The core performs what it owns
// (collapsing a selection); the editor performs row commands (rows are its
// model) and the kernel completes the LaTeX box.
import { goTo } from "./navigation.ts";
import type { CoreModel } from "./stops.ts";
import type { CoreAtom } from "./structures.ts";

export interface KeyModifiers {
  shiftKey: boolean;
  altKey: boolean;
  ctrlKey: boolean;
  metaKey: boolean;
  isComposing?: boolean;
}

/** What Enter does. */
export type EnterAction =
  /** Shift+Enter: a line break inside the formula (it becomes aligned or gathered) */
  | "row-break"
  /** in the LaTeX box: the kernel completes the typed command */
  | "complete"
  /** inside cases: a new case row */
  | "case-row"
  /** elsewhere: split the formula row at the caret (a new row after it at its end or inside a structure) */
  | "split-row";

/** What & does. */
export type AmpersandAction =
  /** in aligned/align: an alignment point (the kernel's alignment rules) */
  | "alignment-point"
  /** at the top level of a row whose rows are written as align: the row's alignment marker */
  | "alignment-marker"
  /** in a matrix, cases or array: the cell on the right, or a new column at the end of the row */
  | "next-column"
  /** elsewhere: the character */
  | "character";

/** What Escape does. */
export type EscapeAction =
  /** a selection collapses to its end (the kernel would turn it into a raw LaTeX group whose text the value drops) */
  | "collapse"
  /** the kernel opens the LaTeX box, or completes it when open */
  | "latex-box";

/** The environment of the array holding the caret (matrix, cases, aligned, ...), if any. */
function environmentAt(model: CoreModel): string | undefined {
  for (let a: CoreAtom | undefined = model.at(model.position)?.parent; a; a = a.parent) if (a.type === "array") return a.environmentName;
  return undefined;
}

const plain = (keys: KeyModifiers) => !keys.altKey && !keys.ctrlKey && !keys.metaKey;

/** Enter (rawLatex: the LaTeX box holds typed text). The editor's row hotkeys are matched before this. */
export function enterAction(model: CoreModel, keys: KeyModifiers, rawLatex: boolean): EnterAction {
  if (keys.shiftKey && plain(keys) && !keys.isComposing && !rawLatex) return "row-break";
  if (rawLatex || model.mode === "latex") return "complete";
  if (environmentAt(model) === "cases") return "case-row";
  return "split-row";
}

/** How deep the caret's atom is (0: in the formula's top level). */
function caretDepth(model: CoreModel): number {
  const atom = model.at(model.position) ?? model.at(Math.max(0, model.position - 1));
  let depth = -1;
  for (let a = atom?.parent; a; a = a.parent) depth += 1;
  return Math.max(0, depth);
}

/**
 * & (rawLatex: the LaTeX box holds typed text; alignRows: multi-line formulas are written as align, so a row's
 * top level takes an alignment marker).
 */
export function ampersandAction(model: CoreModel, keys: KeyModifiers, rawLatex: boolean, alignRows: boolean): AmpersandAction {
  if (keys.isComposing || !plain(keys) || rawLatex) return "character";
  const environment = environmentAt(model);
  if (environment === "aligned" || environment === "align" || environment === "align*") return "alignment-point";
  if (alignRows && environment === undefined && caretDepth(model) === 0) return "alignment-marker";
  return environment !== undefined ? "next-column" : "character";
}

/** Escape (rawLatex: the LaTeX box holds typed text). */
export function escapeAction(model: CoreModel, keys: KeyModifiers, rawLatex: boolean): EscapeAction {
  const ranged = model.selection.ranges.some(([start, end]) => start !== end);
  if (plain(keys) && !keys.shiftKey && !rawLatex && model.mode !== "latex" && ranged) return "collapse";
  return "latex-box";
}

/** Collapse the selection to its end. */
export function collapseToEnd(model: CoreModel): boolean {
  return goTo(model, Math.max(...model.selection.ranges.flat()));
}

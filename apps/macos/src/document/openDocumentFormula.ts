// The open document formula's field, for the formula toolbar and tiles: they insert into it exactly as into a
// formula row (prepare/attemptLatexCommandInsertion), recorded as one document edit.
import type { MathfieldElement } from "mathlive";

export interface OpenDocumentFormula {
  field: MathfieldElement;
  /** a display formula (its own line) rather than an inline one */
  display: boolean;
  /** Run a kernel change as one formula edit; true when it changed the formula. */
  mutate(change: () => boolean): boolean;
}

let open: OpenDocumentFormula | null = null;

export function setOpenDocumentFormula(next: OpenDocumentFormula | null, field?: MathfieldElement) {
  if (next) open = next;
  else if (!field || open?.field === field) open = null;
}

export function openDocumentFormula(): OpenDocumentFormula | null {
  return open?.field.isConnected ? open : null;
}

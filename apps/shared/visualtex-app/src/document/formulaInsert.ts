// Formula commands of the app (toolbar, tiles, the "\" menu in the text) into the document's open formula, exactly
// as into a formula row (prepare/attemptLatexCommandInsertion), recorded as one document edit.
import { attemptLatexCommandInsertion, prepareLatexCommandInsertion } from "../editor/MathEditor";
import { useEditorStore } from "../stores/editorStore";
import type { LatexCommand } from "../types/command";
import { openDocumentFormula, type OpenDocumentFormula } from "./openDocumentFormula";

const nextFrame = () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));

/** The open formula's editor; one that was just opened mounts it within a frame or two. */
export async function waitForOpenFormula(): Promise<OpenDocumentFormula | null> {
  let open = openDocumentFormula();
  for (let i = 0; i < 6 && !open; i++) {
    await nextFrame();
    open = openDocumentFormula();
  }
  return open;
}

export function insertCommandIntoFormula(open: OpenDocumentFormula, command: LatexCommand) {
  const { field, mutate } = open;
  field.focus();
  const inputBehavior = useEditorStore.getState().inputBehavior;
  const prepared = prepareLatexCommandInsertion(field, command);
  if (attemptLatexCommandInsertion(field, command, prepared, inputBehavior, mutate)) return;
  requestAnimationFrame(() => {
    if (!field.isConnected) return;
    field.focus();
    attemptLatexCommandInsertion(field, command, prepared, inputBehavior, mutate);
  });
}

export function insertLatexIntoFormula(open: OpenDocumentFormula, latex: string) {
  const { field, mutate } = open;
  field.focus();
  mutate(() =>
    field.insert(latex, {
      mode: "math",
      format: "latex",
      insertionMode: "replaceSelection",
      selectionMode: latex.includes("\\placeholder{}") ? "placeholder" : "after",
      focus: true,
      scrollIntoView: false,
    }),
  );
}

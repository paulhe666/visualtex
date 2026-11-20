// The engine and the editing core the app registers at startup. This module
// imports nothing, so every kernel module can reach them without changing the
// order in which the kernel's modules are evaluated.

// VisualTeX layout engine: LaTeX-exact layout and drawing of formulas.
// The app registers it at startup; the kernel keeps its model, editing,
// caret and selection, and asks the engine for the formula markup.
export let visualTexLayoutEngine: any = null;
export function setVisualTexLayoutEngine(engine: any): void {
  visualTexLayoutEngine = engine;
}

// VisualTeX editing core (apps/shared/visualtex-core): owns where the caret is
// shown and where atoms are on the page. The kernel asks it instead of
// measuring the drawn formula.
export let visualTexEditingCore: any = null;
export function setVisualTexEditingCore(core: any): void {
  visualTexEditingCore = core;
}

/**
 * Kernel functions that modules evaluated early call, but that live in
 * modules evaluated later (the keyboard input). An import from an early module
 * would move those modules ahead of the command registry they register with,
 * so they put their functions here when they are evaluated.
 */
export const visualTexLateFunctions: {
  visualTexMaybeAutoCommitSourceWrapper?: (mathfield: any) => void;
  visualTexCancelSourceWrapper?: (mathfield: any) => boolean;
  visualTexSourceWrapperKind?: (command: string) => string | null;
  visualTexCanonicalVariantInputCommand?: (command: string) => string;
  visualTexStartSourceWrapper?: (model: any, latex: string, structuralSlot: any) => boolean;
  visualTexSyncSourceWrapper?: (mathfield: any) => void;
  /** LaTeX the user typed or pasted as the formulas store it (the app's exam-zh spelling setting); identity if unset */
  visualTexUserLatexIn?: (latex: string) => string;
} = {};

/** The app's reading of LaTeX the user writes (typed commands, pasted text, previews of what typing inserts). */
export function setVisualTexUserLatexIn(read: ((latex: string) => string) | null): void {
  visualTexLateFunctions.visualTexUserLatexIn = read ?? undefined;
}

/** LaTeX the user typed or pasted, as the formulas store it. */
export function visualTexUserLatex(latex: string): string {
  return visualTexLateFunctions.visualTexUserLatexIn ? visualTexLateFunctions.visualTexUserLatexIn(latex) : latex;
}

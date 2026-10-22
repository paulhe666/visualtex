// The host module each app provides: tables generated from the app's
// registries, and the editor's input judgment.
declare module 'virtual:visualtex-kernel-host' {
  export const visualTexDefaultMacros: Record<string, any>;
  /** One entry of the editor's command catalog: a completion candidate */
  export type VisualTexCompletionRecord = {
    /** false for a shared variant: found by its command, never claims an exactly typed one */
    exact: boolean;
    id: string;
    command: string;
    /** what is inserted: a template with \placeholder{} slots */
    latex: string;
    preview: string;
    label: string;
    /** English label, aliases and keywords */
    words: string[];
    priority?: number;
  };
  export const visualTexCompletionCatalogData: VisualTexCompletionRecord[];
  export const visualTexCommandFrequencyData: Record<string, number>;
  export const visualTexRareIntegralSymbols: Record<string, string>;
  export const visualTexSourceWrapperKindsData: Record<string, string>;
  export const visualTexSourceWrapperCanonicalTargetsData: Record<string, string>;
  /** Sample formulas the completion list draws for bare kernel commands (\\left -> \\left( x \\right)) */
  export const visualTexNativePreviewsData: Record<string, string>;
  export const VISUALTEX_SYMBOL_SEQUENCES: Record<string, any>;
  export function applyInputJudgment(before: string, judgment: any): any;
}

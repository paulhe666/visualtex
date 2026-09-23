import { forwardRef } from "react";
import type { MathEditorHandle } from "./mathEditorTypes";

/**
 * Default production build stub. vite.config.ts resolves the experimental
 * CoreMathEditor import here unless VITE_VISUALTEX_CORE_EDITOR=1 is explicitly
 * supplied at build time. Keeping the stub resource-free prevents the verified
 * WASM/font candidate from changing ordinary legacy bundles before P10 approval.
 */
export async function probeCoreEditorSession() {
  return {
    eligible: false,
    reason: "The VisualTeX core editor is not included in this build.",
  } as const;
}

export const CoreMathEditor = forwardRef<MathEditorHandle, Record<string, unknown>>(
  function CoreMathEditorDisabled() {
    return null;
  },
);

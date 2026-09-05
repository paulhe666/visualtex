import { CUSTOM_SYMBOL_PROTOTYPE_DEFINITION } from "./customSymbolRegistry.ts";

/**
 * Compatibility facade kept for the phase-1 regression files. Runtime
 * rendering now comes entirely from CustomSymbolRegistry/Rendering.
 */
export const CUSTOM_SYMBOL_PROTOTYPE_COMMAND =
  CUSTOM_SYMBOL_PROTOTYPE_DEFINITION.command;
export const CUSTOM_SYMBOL_PROTOTYPE_LATEX =
  `\\${CUSTOM_SYMBOL_PROTOTYPE_COMMAND}`;

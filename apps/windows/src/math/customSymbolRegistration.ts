import {
  registerCustomSymbol,
  updateCustomSymbol,
} from "./customSymbolRegistry.ts";

/**
 * P15: command ownership is defined by VisualTeX's own registry/core reserved
 * control-word tables. The previous MathLive render probe was both redundant
 * and a production dependency on the retired editor runtime.
 */
export function registerCustomSymbolSafely(value: unknown) {
  return registerCustomSymbol(value);
}

export function updateCustomSymbolSafely(id: string, patch: unknown) {
  return updateCustomSymbol(id, patch);
}

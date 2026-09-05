import { convertLatexToMarkup } from "mathlive";
import {
  registerCustomSymbol,
  setCustomSymbolCommandAvailabilityValidator,
  updateCustomSymbol,
} from "./customSymbolRegistry.ts";

function commandFromValue(value: unknown) {
  if (!value || typeof value !== "object") return "";
  const candidate = (value as Record<string, unknown>).command;
  return typeof candidate === "string" ? candidate.trim().replace(/^\\/, "") : "";
}

function mathLiveCustomSymbolCommandIsAvailable(command: string) {
  const normalized = command.trim().replace(/^\\/, "");
  if (!/^[A-Za-z]+$/.test(normalized)) return false;
  const markup = convertLatexToMarkup(`\\${normalized}`, {
    defaultMode: "math",
  });
  return markup.includes("ML__error");
}

setCustomSymbolCommandAvailabilityValidator(
  mathLiveCustomSymbolCommandIsAvailable,
);

export function assertMathLiveCustomSymbolCommandAvailable(command: string) {
  const normalized = command.trim().replace(/^\\/, "");
  if (!/^[A-Za-z]+$/.test(normalized)) return;
  if (!mathLiveCustomSymbolCommandIsAvailable(normalized)) {
    throw new Error(`\\${normalized} is already defined by MathLive/LaTeX.`);
  }
}

export function registerCustomSymbolSafely(value: unknown) {
  const command = commandFromValue(value);
  assertMathLiveCustomSymbolCommandAvailable(command);
  return registerCustomSymbol(value);
}

export function updateCustomSymbolSafely(id: string, patch: unknown) {
  const command = commandFromValue(patch);
  if (command) assertMathLiveCustomSymbolCommandAvailable(command);
  return updateCustomSymbol(id, patch);
}

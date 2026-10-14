import { readLocalStorage, writeLocalStorage } from "../../runtime/safeStorage";

export type WordFormulaOutputKind = "image" | "omml";
const legacyKey = "visualtex.office.word.create.font-size-pt";
const keyFor = (kind: WordFormulaOutputKind) => `${legacyKey}.${kind}`;

export function readWordFormulaFontSize(kind: WordFormulaOutputKind): number | null {
  // Existing installations used one preference for both formula types.
  const stored = readLocalStorage(keyFor(kind)) ?? readLocalStorage(legacyKey);
  if (stored === null) return null;
  const value = Number(stored);
  return Number.isFinite(value) && value >= 5 && value <= 200
    ? Math.round(value * 2) / 2
    : null;
}

export function writeWordFormulaFontSize(kind: WordFormulaOutputKind, value: number) {
  if (!Number.isFinite(value)) return;
  writeLocalStorage(keyFor(kind), String(Math.round(Math.min(200, Math.max(5, value)) * 2) / 2));
}

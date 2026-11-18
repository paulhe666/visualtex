import {
  DEFAULT_FORMULA_CHINESE_FONT,
  DEFAULT_FORMULA_LETTER_FONT,
  FORMULA_FACE_ROLES,
  normalizeFormulaChineseFont,
  normalizeFormulaLetterFont,
  readPersistedFormulaFaces,
  readPersistedFormulaFontPreferences,
  type FormulaChineseFont,
  type FormulaFaces,
  type FormulaLetterFont,
} from "../../editor/formulaFontPreferences";
import type { VisualTeXFormulaMetadata } from "./formulaMetadata";

// Update Equation Format's font step: a VisualTeX formula drawn in other fonts than the ones the app draws in now is
// drawn again from its LaTeX. The rule is the Windows add-ins' (FormulaFontSettings.DifferFrom) and the macOS app's
// (src-tauri office/formula_fonts.rs), which the Office add-ins ask.

/** The fonts the app draws formulas in: the math font (KaTeX or newtxmath, stored as "times"), the Chinese font and
 * the fonts for letters, Greek letters and digits ("" follows the math font). */
export interface FormulaFontSettings {
  letterFont: FormulaLetterFont;
  chineseFont: FormulaChineseFont;
  faces: FormulaFaces;
}

/** What a formula records of the fonts it was drawn in. */
export type FormulaFontRecord = Pick<
  VisualTeXFormulaMetadata,
  "latex" | "formulaLetterFont" | "formulaChineseFont" | "formulaFaces"
>;

/** The fonts of the Settings, as stored (every window of the app reads the same storage). */
export function currentFormulaFontSettings(): FormulaFontSettings {
  const stored = readPersistedFormulaFontPreferences();
  return {
    letterFont: normalizeFormulaLetterFont(stored.formulaLetterFont ?? DEFAULT_FORMULA_LETTER_FONT),
    chineseFont: normalizeFormulaChineseFont(stored.formulaChineseFont ?? DEFAULT_FORMULA_CHINESE_FONT),
    faces: readPersistedFormulaFaces(),
  };
}

// CJK ideographs, compatibility ideographs, CJK punctuation and full-width forms: what the Chinese font draws
const CHINESE_TEXT = /[\u3400-\u9fff\uf900-\ufaff\u3000-\u303f\uff00-\uffef]/;

function sameName(recorded: string | null | undefined, current: string | null | undefined) {
  return (recorded ?? "").trim().toLowerCase() === (current ?? "").trim().toLowerCase();
}

/**
 * Whether a formula was drawn in other fonts than `current`. A record the formula lacks counts as what VisualTeX drew
 * without it: KaTeX, the system Chinese font, no faces. The Chinese font counts only for a formula with Chinese
 * text, which is all it changes.
 */
export function formulaFontsDiffer(metadata: FormulaFontRecord, current: FormulaFontSettings): boolean {
  const letter = metadata.formulaLetterFont?.trim() || DEFAULT_FORMULA_LETTER_FONT;
  if (!sameName(letter, current.letterFont)) return true;
  const faces = metadata.formulaFaces;
  if (FORMULA_FACE_ROLES.some((role) => !sameName(faces?.[role], current.faces[role]))) return true;
  const chinese = metadata.formulaChineseFont?.trim() || DEFAULT_FORMULA_CHINESE_FONT;
  return !sameName(chinese, current.chineseFont) && CHINESE_TEXT.test(metadata.latex ?? "");
}

/** The fonts in a few words, for confirmations: "KaTeX; letters, Greek letters, digits: …" (as on Windows). */
export function describeFormulaFonts(current: FormulaFontSettings, english: boolean): string {
  const math = current.letterFont === "times" ? "newtxmath" : "KaTeX";
  const names: Record<keyof FormulaFaces, [string, string]> = {
    letter: ["字母", "letters"],
    greek: ["希腊字母", "Greek letters"],
    digit: ["数字", "digits"],
  };
  const parts = FORMULA_FACE_ROLES.filter((role) => current.faces[role].trim()).map((role) => {
    const [zh, en] = names[role];
    return english ? `${en}: ${current.faces[role].trim()}` : `${zh}：${current.faces[role].trim()}`;
  });
  if (!parts.length) return math;
  return english ? `${math}; ${parts.join(", ")}` : `${math}；${parts.join("，")}`;
}

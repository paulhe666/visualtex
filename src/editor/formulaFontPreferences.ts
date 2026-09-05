import { safeStorage } from "../runtime/safeStorage.ts";

export type FormulaLetterFont =
  | "katex"
  | "times"
  | "cambria"
  | "stix"
  | "palatino"
  | "helvetica";

export type FormulaChineseFont =
  | "system"
  | "pingfang"
  | "songti"
  | "kaiti"
  | "heiti";

export const DEFAULT_FORMULA_LETTER_FONT: FormulaLetterFont = "katex";
export const DEFAULT_FORMULA_CHINESE_FONT: FormulaChineseFont = "system";

const FORMULA_LETTER_FONT_STORAGE_KEY = "visualtex.formula-letter-font";
const FORMULA_CHINESE_FONT_STORAGE_KEY = "visualtex.formula-chinese-font";
export const FORMULA_FONT_PREFERENCES_CHANGED_EVENT =
  "visualtex-formula-font-preferences-changed";

function notifyFormulaFontPreferencesChanged() {
  if (typeof window === "undefined") return;
  window.dispatchEvent(new CustomEvent(FORMULA_FONT_PREFERENCES_CHANGED_EVENT));
}

export const FORMULA_LETTER_FONT_OPTIONS: ReadonlyArray<{
  id: FormulaLetterFont;
  label: string;
}> = [
  { id: "katex", label: "KaTeX / Computer Modern" },
  { id: "times", label: "Times New Roman" },
  { id: "cambria", label: "Cambria Math" },
  { id: "stix", label: "STIX" },
  { id: "palatino", label: "Palatino" },
  { id: "helvetica", label: "Helvetica / Arial" },
];

export const FORMULA_CHINESE_FONT_OPTIONS: ReadonlyArray<{
  id: FormulaChineseFont;
  labelZh: string;
  labelEn: string;
}> = [
  { id: "system", labelZh: "系统默认", labelEn: "System default" },
  { id: "pingfang", labelZh: "苹方（兼容）", labelEn: "PingFang compatible" },
  { id: "songti", labelZh: "宋体", labelEn: "SimSun / Songti" },
  { id: "kaiti", labelZh: "楷体", labelEn: "KaiTi" },
  { id: "heiti", labelZh: "黑体", labelEn: "SimHei" },
];

const LETTER_FONT_FAMILIES: Record<
  FormulaLetterFont,
  { upright: string; italic: string }
> = {
  katex: {
    upright: "KaTeX_Main, serif",
    italic: "KaTeX_Math, KaTeX_Main, serif",
  },
  times: {
    upright: '"Times New Roman", Times, serif',
    italic: '"Times New Roman", Times, serif',
  },
  cambria: {
    upright: '"Cambria Math", Cambria, "Times New Roman", Times, serif',
    italic: '"Cambria Math", Cambria, "Times New Roman", Times, serif',
  },
  stix: {
    upright: '"STIX Two Math", "STIX Two Text", STIXGeneral, "Times New Roman", Times, serif',
    italic: '"STIX Two Math", "STIX Two Text", STIXGeneral, "Times New Roman", Times, serif',
  },
  palatino: {
    upright: 'Palatino, "Palatino Linotype", "Book Antiqua", serif',
    italic: 'Palatino, "Palatino Linotype", "Book Antiqua", serif',
  },
  helvetica: {
    upright: '"Helvetica Neue", Helvetica, Arial, sans-serif',
    italic: '"Helvetica Neue", Helvetica, Arial, sans-serif',
  },
};

const LETTER_PRIMARY_FONT_NAMES: Record<FormulaLetterFont, string> = {
  katex: "KaTeX_Math",
  times: "Times New Roman",
  cambria: "Cambria Math",
  stix: "STIX Two Math",
  palatino: "Palatino Linotype",
  helvetica: "Arial",
};

const CHINESE_PRIMARY_FONT_NAMES: Record<FormulaChineseFont, string> = {
  system: "Microsoft YaHei",
  pingfang: "Microsoft YaHei",
  songti: "SimSun",
  kaiti: "KaiTi",
  heiti: "SimHei",
};

const CHINESE_FONT_FAMILIES: Record<FormulaChineseFont, string> = {
  system:
    '"Microsoft YaHei", "Microsoft JhengHei", "Noto Sans CJK SC", "PingFang SC", sans-serif',
  pingfang:
    '"PingFang SC", "Microsoft YaHei", "Noto Sans CJK SC", sans-serif',
  songti: 'SimSun, NSimSun, "Songti SC", STSong, serif',
  kaiti: 'KaiTi, "KaiTi_GB2312", "Kaiti SC", STKaiti, serif',
  heiti: 'SimHei, "Microsoft YaHei", "Heiti SC", STHeiti, sans-serif',
};

export function normalizeFormulaLetterFont(value: unknown): FormulaLetterFont {
  return FORMULA_LETTER_FONT_OPTIONS.some((item) => item.id === value)
    ? (value as FormulaLetterFont)
    : DEFAULT_FORMULA_LETTER_FONT;
}

export function normalizeFormulaChineseFont(value: unknown): FormulaChineseFont {
  return FORMULA_CHINESE_FONT_OPTIONS.some((item) => item.id === value)
    ? (value as FormulaChineseFont)
    : DEFAULT_FORMULA_CHINESE_FONT;
}

export function readPersistedFormulaFontPreferences(): {
  formulaLetterFont: FormulaLetterFont | null;
  formulaChineseFont: FormulaChineseFont | null;
} {
  const letter = safeStorage.getItem(FORMULA_LETTER_FONT_STORAGE_KEY);
  const chinese = safeStorage.getItem(FORMULA_CHINESE_FONT_STORAGE_KEY);
  return {
    formulaLetterFont: FORMULA_LETTER_FONT_OPTIONS.some((item) => item.id === letter)
      ? (letter as FormulaLetterFont)
      : null,
    formulaChineseFont: FORMULA_CHINESE_FONT_OPTIONS.some((item) => item.id === chinese)
      ? (chinese as FormulaChineseFont)
      : null,
  };
}

export function persistFormulaLetterFontPreference(value: FormulaLetterFont) {
  safeStorage.setItem(
    FORMULA_LETTER_FONT_STORAGE_KEY,
    normalizeFormulaLetterFont(value),
  );
  notifyFormulaFontPreferencesChanged();
}

export function persistFormulaChineseFontPreference(value: FormulaChineseFont) {
  safeStorage.setItem(
    FORMULA_CHINESE_FONT_STORAGE_KEY,
    normalizeFormulaChineseFont(value),
  );
  notifyFormulaFontPreferencesChanged();
}

export function persistFormulaFontPreferences(
  formulaLetterFont: FormulaLetterFont,
  formulaChineseFont: FormulaChineseFont,
) {
  persistFormulaLetterFontPreference(formulaLetterFont);
  persistFormulaChineseFontPreference(formulaChineseFont);
}

export function formulaLetterFontFamilies(value: FormulaLetterFont) {
  return LETTER_FONT_FAMILIES[normalizeFormulaLetterFont(value)];
}

export function formulaChineseFontFamily(value: FormulaChineseFont) {
  return CHINESE_FONT_FAMILIES[normalizeFormulaChineseFont(value)];
}

export function formulaLetterPrimaryFontName(value: FormulaLetterFont) {
  return LETTER_PRIMARY_FONT_NAMES[normalizeFormulaLetterFont(value)];
}

export function formulaChinesePrimaryFontName(value: FormulaChineseFont) {
  return CHINESE_PRIMARY_FONT_NAMES[normalizeFormulaChineseFont(value)];
}

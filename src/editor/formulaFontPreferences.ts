import { safeStorage } from "../runtime/safeStorage";

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

export const VISUALTEX_FORMULA_LETTER_GLYPH_CLASS =
  "visualtex-formula-letter-glyph";
export const VISUALTEX_FORMULA_CHINESE_GLYPH_CLASS =
  "visualtex-chinese-glyph";

const VISUALTEX_FORMULA_FONT_GLYPH_SELECTOR =
  ".ML__cmr, .ML__mathbf, .ML__mathit, .ML__mathbfit, .ML__text";
const VISUALTEX_FORMULA_LETTER_GLYPH_PATTERN =
  /^[\p{Script=Latin}\p{Script=Greek}0-9]+$/u;
const VISUALTEX_FORMULA_CHINESE_GLYPH_PATTERN =
  /^[\p{Script=Han}，。；：！？、（）【】《》“”‘’]+$/u;

const FORMULA_LETTER_FONT_STORAGE_KEY = "visualtex.formula-letter-font";
const FORMULA_CHINESE_FONT_STORAGE_KEY = "visualtex.formula-chinese-font";

export const FORMULA_LETTER_FONT_OPTIONS: ReadonlyArray<{
  id: FormulaLetterFont;
  label: string;
}> = [
  { id: "katex", label: "KaTeX / Computer Modern" },
  { id: "times", label: "Times New Roman" },
  { id: "cambria", label: "Cambria Math" },
  { id: "stix", label: "STIX" },
  { id: "palatino", label: "Palatino" },
  { id: "helvetica", label: "Helvetica" },
];

export const FORMULA_CHINESE_FONT_OPTIONS: ReadonlyArray<{
  id: FormulaChineseFont;
  labelZh: string;
  labelEn: string;
}> = [
  { id: "system", labelZh: "系统默认", labelEn: "System default" },
  { id: "pingfang", labelZh: "苹方", labelEn: "PingFang SC" },
  { id: "songti", labelZh: "宋体", labelEn: "Songti SC" },
  { id: "kaiti", labelZh: "楷体", labelEn: "Kaiti SC" },
  { id: "heiti", labelZh: "黑体", labelEn: "Heiti SC" },
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
  palatino: "Palatino",
  helvetica: "Helvetica Neue",
};

const CHINESE_PRIMARY_FONT_NAMES: Record<FormulaChineseFont, string> = {
  system: "PingFang SC",
  pingfang: "PingFang SC",
  songti: "Songti SC",
  kaiti: "Kaiti SC",
  heiti: "Heiti SC",
};

const CHINESE_FONT_FAMILIES: Record<FormulaChineseFont, string> = {
  system:
    '-apple-system, BlinkMacSystemFont, "PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC", "Microsoft YaHei", sans-serif',
  pingfang:
    '"PingFang SC", "Hiragino Sans GB", "Noto Sans CJK SC", "Microsoft YaHei", sans-serif',
  songti: '"Songti SC", STSong, SimSun, "Songti TC", serif',
  kaiti: '"Kaiti SC", STKaiti, KaiTi, "KaiTi_GB2312", serif',
  heiti:
    '"Heiti SC", STHeiti, "PingFang SC", "Microsoft YaHei", sans-serif',
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
    formulaLetterFont: FORMULA_LETTER_FONT_OPTIONS.some(
      (item) => item.id === letter,
    )
      ? (letter as FormulaLetterFont)
      : null,
    formulaChineseFont: FORMULA_CHINESE_FONT_OPTIONS.some(
      (item) => item.id === chinese,
    )
      ? (chinese as FormulaChineseFont)
      : null,
  };
}

export function persistFormulaLetterFontPreference(value: FormulaLetterFont) {
  safeStorage.setItem(
    FORMULA_LETTER_FONT_STORAGE_KEY,
    normalizeFormulaLetterFont(value),
  );
}

export function persistFormulaChineseFontPreference(value: FormulaChineseFont) {
  safeStorage.setItem(
    FORMULA_CHINESE_FONT_STORAGE_KEY,
    normalizeFormulaChineseFont(value),
  );
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

export function markVisualTexFormulaFontGlyphs(root: ParentNode) {
  root
    .querySelectorAll<HTMLElement>(VISUALTEX_FORMULA_FONT_GLYPH_SELECTOR)
    .forEach((node) => {
      const text = (node.textContent ?? "").trim();
      const isChinese =
        Boolean(text) && VISUALTEX_FORMULA_CHINESE_GLYPH_PATTERN.test(text);
      const isLetterGlyph =
        Boolean(text) &&
        !isChinese &&
        VISUALTEX_FORMULA_LETTER_GLYPH_PATTERN.test(text);

      node.classList.toggle(
        VISUALTEX_FORMULA_CHINESE_GLYPH_CLASS,
        isChinese,
      );
      node.classList.toggle(
        VISUALTEX_FORMULA_LETTER_GLYPH_CLASS,
        isLetterGlyph,
      );
    });
}

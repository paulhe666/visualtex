import { safeStorage } from "../runtime/safeStorage.ts";

// The formula fonts of the main app: KaTeX (LaTeX's Computer Modern) and
// Times New Roman (LaTeX's newtxtext + newtxmath). Earlier versions offered
// more; formulas saved with those carry the old names (LegacyFormulaLetterFont)
// and are shown in KaTeX.
export type FormulaLetterFont = "katex" | "times";
export type LegacyFormulaLetterFont = FormulaLetterFont | "cambria" | "stix" | "palatino" | "helvetica";

/** Every letter font a saved formula may name (validated as such, shown via normalizeFormulaLetterFont). */
export const LEGACY_FORMULA_LETTER_FONTS: readonly LegacyFormulaLetterFont[] = [
  "katex",
  "times",
  "cambria",
  "stix",
  "palatino",
  "helvetica",
];

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
};

const LETTER_PRIMARY_FONT_NAMES: Record<FormulaLetterFont, string> = {
  katex: "KaTeX_Main",
  times: "Times New Roman",
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

// local() names (full and PostScript) of the faces in CHINESE_FONT_FAMILIES.
const CHINESE_LOCAL_FACES: Record<FormulaChineseFont, readonly string[]> = {
  system: [
    "Microsoft YaHei", "MicrosoftYaHei", "Microsoft JhengHei", "MicrosoftJhengHeiRegular",
    "Noto Sans CJK SC", "NotoSansCJKsc-Regular", "PingFang SC", "PingFangSC-Regular",
  ],
  pingfang: [
    "PingFang SC", "PingFangSC-Regular", "Microsoft YaHei", "MicrosoftYaHei",
    "Noto Sans CJK SC", "NotoSansCJKsc-Regular",
  ],
  songti: ["SimSun", "NSimSun", "Songti SC", "STSongti-SC-Regular", "STSong"],
  kaiti: ["KaiTi", "KaiTi_GB2312", "Kaiti SC", "STKaiti-SC-Regular", "STKaiti"],
  heiti: ["SimHei", "Microsoft YaHei", "MicrosoftYaHei", "Heiti SC", "STHeiti"],
};

// CJK ideographs plus the punctuation in VISUALTEX_FORMULA_CHINESE_GLYPH_PATTERN.
const CHINESE_UNICODE_RANGE = [
  "U+2E80-2FDF", "U+3001-3002", "U+3005-3007", "U+300A-300B", "U+3010-3011",
  "U+3021-3029", "U+3038-303B", "U+3400-4DBF", "U+4E00-9FFF", "U+F900-FAFF",
  "U+FF01", "U+FF08-FF09", "U+FF0C", "U+FF1A-FF1B", "U+FF1F", "U+2018-2019",
  "U+201C-201D", "U+20000-2FA1F",
].join(", ");

const CHINESE_FALLBACK_FONT_FACE_STYLE_ID = "visualtex-formula-chinese-fallback-faces";

function chineseFallbackFamilyName(value: FormulaChineseFont) {
  return `VisualTeX Formula Chinese ${value}`;
}

/**
 * Family that renders only CJK code points with the chosen formula Chinese
 * font. Appended after a MathLive KaTeX family it catches Chinese glyphs in
 * any font class without touching the KaTeX glyphs or the fallback for other
 * characters missing from KaTeX.
 */
export function formulaChineseFallbackFamily(value: FormulaChineseFont) {
  return `"${chineseFallbackFamilyName(normalizeFormulaChineseFont(value))}"`;
}

/** @font-face rules must live in the document; shadow roots do not register them. */
export function installFormulaChineseFallbackFontFaces() {
  if (typeof document === "undefined") return;
  if (document.getElementById(CHINESE_FALLBACK_FONT_FACE_STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = CHINESE_FALLBACK_FONT_FACE_STYLE_ID;
  style.textContent = FORMULA_CHINESE_FONT_OPTIONS.map(({ id }) => `
@font-face {
  font-family: "${chineseFallbackFamilyName(id)}";
  src: ${CHINESE_LOCAL_FACES[id].map((name) => `local("${name}")`).join(", ")};
  font-weight: 400;
  unicode-range: ${CHINESE_UNICODE_RANGE};
}`).join("\n");
  document.head.append(style);
}

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

// Formulas are laid out again on the change notification, so it is sent only
// when the stored value changes (the Office window applies the main app's
// preferences on every sync).
export function persistFormulaLetterFontPreference(value: FormulaLetterFont) {
  const next = normalizeFormulaLetterFont(value);
  if (safeStorage.getItem(FORMULA_LETTER_FONT_STORAGE_KEY) === next) return;
  safeStorage.setItem(FORMULA_LETTER_FONT_STORAGE_KEY, next);
  notifyFormulaFontPreferencesChanged();
}

export function persistFormulaChineseFontPreference(value: FormulaChineseFont) {
  const next = normalizeFormulaChineseFont(value);
  if (safeStorage.getItem(FORMULA_CHINESE_FONT_STORAGE_KEY) === next) return;
  safeStorage.setItem(FORMULA_CHINESE_FONT_STORAGE_KEY, next);
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

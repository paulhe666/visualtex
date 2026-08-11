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

export const FORMULA_LETTER_FONT_OPTIONS: ReadonlyArray<{
  id: FormulaLetterFont;
  label: string;
}> = [
  { id: "katex", label: "KaTeX Math" },
  { id: "times", label: "Times / TeX Gyre Termes" },
  { id: "cambria", label: "Cambria Math" },
  { id: "stix", label: "STIX Two Math" },
  { id: "palatino", label: "Palatino / TeX Gyre Pagella" },
  { id: "helvetica", label: "Helvetica / Arial" },
];

export const FORMULA_CHINESE_FONT_OPTIONS: ReadonlyArray<{
  id: FormulaChineseFont;
  labelZh: string;
  labelEn: string;
}> = [
  { id: "system", labelZh: "系统默认", labelEn: "System default" },
  { id: "pingfang", labelZh: "苹方", labelEn: "PingFang SC" },
  { id: "songti", labelZh: "宋体", labelEn: "Songti / SimSun" },
  { id: "kaiti", labelZh: "楷体", labelEn: "Kaiti / KaiTi" },
  { id: "heiti", labelZh: "黑体", labelEn: "Heiti / SimHei" },
];

const FORMULA_FONT_STORAGE_KEY = "visualtex.formula-fonts.v1";
export const FORMULA_FONT_PREFERENCES_CHANGED_EVENT =
  "visualtex-formula-font-preferences-changed";
const FORMULA_FONT_CHANNEL = "visualtex-formula-fonts";

const letterFonts = new Set<FormulaLetterFont>(
  FORMULA_LETTER_FONT_OPTIONS.map((option) => option.id),
);
const chineseFonts = new Set<FormulaChineseFont>(
  FORMULA_CHINESE_FONT_OPTIONS.map((option) => option.id),
);

export function normalizeFormulaLetterFont(
  value: unknown,
): FormulaLetterFont {
  return typeof value === "string" && letterFonts.has(value as FormulaLetterFont)
    ? (value as FormulaLetterFont)
    : DEFAULT_FORMULA_LETTER_FONT;
}

export function normalizeFormulaChineseFont(
  value: unknown,
): FormulaChineseFont {
  return typeof value === "string" && chineseFonts.has(value as FormulaChineseFont)
    ? (value as FormulaChineseFont)
    : DEFAULT_FORMULA_CHINESE_FONT;
}

export function formulaLetterFontFamilies(font: FormulaLetterFont) {
  switch (normalizeFormulaLetterFont(font)) {
    case "times":
      return {
        upright: '"TeX Gyre Termes", "Times New Roman", Times, serif',
        italic: '"TeX Gyre Termes Math", "Times New Roman", Times, serif',
      };
    case "cambria":
      return {
        upright: '"Cambria Math", Cambria, serif',
        italic: '"Cambria Math", Cambria, serif',
      };
    case "stix":
      return {
        upright: '"STIX Two Math", "STIX Two Text", "STIXGeneral", serif',
        italic: '"STIX Two Math", "STIX Two Text", "STIXGeneral", serif',
      };
    case "palatino":
      return {
        upright: '"TeX Gyre Pagella", Palatino, "Palatino Linotype", serif',
        italic: '"TeX Gyre Pagella Math", Palatino, "Palatino Linotype", serif',
      };
    case "helvetica":
      return {
        upright: 'Helvetica, Arial, "Noto Sans", sans-serif',
        italic: 'Helvetica, Arial, "Noto Sans", sans-serif',
      };
    default:
      return {
        upright: "KaTeX_Main, serif",
        italic: "KaTeX_Math, KaTeX_Main, serif",
      };
  }
}

export function formulaChineseFontFamily(font: FormulaChineseFont) {
  switch (normalizeFormulaChineseFont(font)) {
    case "pingfang":
      return '"PingFang SC", "PingFang TC", "Microsoft YaHei", sans-serif';
    case "songti":
      return '"Songti SC", "STSong", "SimSun", serif';
    case "kaiti":
      return '"Kaiti SC", "STKaiti", "KaiTi", serif';
    case "heiti":
      return '"Heiti SC", "STHeiti", "SimHei", sans-serif';
    default:
      return 'system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif';
  }
}

export interface FormulaFontPreferences {
  formulaLetterFont: FormulaLetterFont;
  formulaChineseFont: FormulaChineseFont;
}

export function readFormulaFontPreferences(): FormulaFontPreferences {
  try {
    const raw = safeStorage.getItem(FORMULA_FONT_STORAGE_KEY);
    if (!raw) {
      return {
        formulaLetterFont: DEFAULT_FORMULA_LETTER_FONT,
        formulaChineseFont: DEFAULT_FORMULA_CHINESE_FONT,
      };
    }
    const parsed = JSON.parse(raw) as Partial<FormulaFontPreferences>;
    return {
      formulaLetterFont: normalizeFormulaLetterFont(parsed.formulaLetterFont),
      formulaChineseFont: normalizeFormulaChineseFont(parsed.formulaChineseFont),
    };
  } catch {
    return {
      formulaLetterFont: DEFAULT_FORMULA_LETTER_FONT,
      formulaChineseFont: DEFAULT_FORMULA_CHINESE_FONT,
    };
  }
}

export function writeFormulaFontPreferences(
  preferences: FormulaFontPreferences,
) {
  const normalized: FormulaFontPreferences = {
    formulaLetterFont: normalizeFormulaLetterFont(preferences.formulaLetterFont),
    formulaChineseFont: normalizeFormulaChineseFont(preferences.formulaChineseFont),
  };
  safeStorage.setItem(FORMULA_FONT_STORAGE_KEY, JSON.stringify(normalized));
  if (typeof window !== "undefined") {
    window.dispatchEvent(
      new CustomEvent(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, {
        detail: normalized,
      }),
    );
    if (typeof BroadcastChannel !== "undefined") {
      const channel = new BroadcastChannel(FORMULA_FONT_CHANNEL);
      channel.postMessage(normalized);
      channel.close();
    }
  }
  return normalized;
}

export function subscribeFormulaFontPreferences(
  listener: (preferences: FormulaFontPreferences) => void,
) {
  if (typeof window === "undefined") return () => undefined;
  const local = (event: Event) => {
    const detail = (event as CustomEvent<FormulaFontPreferences>).detail;
    listener(detail ?? readFormulaFontPreferences());
  };
  const storage = (event: StorageEvent) => {
    if (event.key === FORMULA_FONT_STORAGE_KEY) {
      listener(readFormulaFontPreferences());
    }
  };
  window.addEventListener(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, local);
  window.addEventListener("storage", storage);
  const channel =
    typeof BroadcastChannel === "undefined"
      ? null
      : new BroadcastChannel(FORMULA_FONT_CHANNEL);
  if (channel) channel.onmessage = () => listener(readFormulaFontPreferences());
  return () => {
    window.removeEventListener(FORMULA_FONT_PREFERENCES_CHANGED_EVENT, local);
    window.removeEventListener("storage", storage);
    channel?.close();
  };
}

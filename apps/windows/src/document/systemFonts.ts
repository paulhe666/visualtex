// The fonts installed on the computer, for the document's body font menus (docs/document-editor/DESIGN.md §8.1):
// the desktop app reads the system's font files (list_system_font_families); elsewhere (the document lab, a
// browser) a list of the usual Windows and macOS fonts stands in for it.
import { invoke, isTauri } from "@tauri-apps/api/core";

export interface SystemFontFamily {
  /** the name the font is asked for by (CSS, LaTeX) */
  family: string;
  /** its Chinese name, when it has one (宋体 for SimSun) */
  localized: string | null;
  /** the font has Chinese characters */
  cjk: boolean;
}

/** Fonts most Chinese Windows and macOS systems have (the list when the system's cannot be read). */
const COMMON_FONTS: SystemFontFamily[] = [
  { family: "SimSun", localized: "宋体", cjk: true },
  { family: "NSimSun", localized: "新宋体", cjk: true },
  { family: "SimHei", localized: "黑体", cjk: true },
  { family: "KaiTi", localized: "楷体", cjk: true },
  { family: "FangSong", localized: "仿宋", cjk: true },
  { family: "Microsoft YaHei", localized: "微软雅黑", cjk: true },
  { family: "DengXian", localized: "等线", cjk: true },
  { family: "Songti SC", localized: "宋体-简", cjk: true },
  { family: "PingFang SC", localized: "苹方-简", cjk: true },
  { family: "Source Han Serif SC", localized: "思源宋体", cjk: true },
  { family: "Source Han Sans SC", localized: "思源黑体", cjk: true },
  { family: "Times New Roman", localized: null, cjk: false },
  { family: "Arial", localized: null, cjk: false },
  { family: "Calibri", localized: null, cjk: false },
  { family: "Cambria", localized: null, cjk: false },
  { family: "Georgia", localized: null, cjk: false },
  { family: "Garamond", localized: null, cjk: false },
  { family: "Palatino Linotype", localized: null, cjk: false },
  { family: "Book Antiqua", localized: null, cjk: false },
  { family: "Segoe UI", localized: null, cjk: false },
  { family: "Verdana", localized: null, cjk: false },
  { family: "Consolas", localized: null, cjk: false },
  { family: "Courier New", localized: null, cjk: false },
];

/** Fonts a Chinese document most often takes (first in the menus when the system has them). */
export const PREFERRED_CJK = ["SimSun", "SimHei", "KaiTi", "FangSong", "Microsoft YaHei", "DengXian", "Source Han Serif SC", "Songti SC"];
export const PREFERRED_LATIN = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond", "Palatino Linotype", "Latin Modern Roman"];

let cached: Promise<SystemFontFamily[]> | null = null;

function isFamily(value: unknown): value is SystemFontFamily {
  if (!value || typeof value !== "object") return false;
  const item = value as Record<string, unknown>;
  return (
    typeof item.family === "string" &&
    item.family.length > 0 &&
    item.family.length <= 128 &&
    (item.localized === null || item.localized === undefined || typeof item.localized === "string") &&
    typeof item.cjk === "boolean"
  );
}

async function readSystemFonts(): Promise<SystemFontFamily[]> {
  if (isTauri()) {
    try {
      const result = await invoke<unknown>("list_system_font_families");
      if (Array.isArray(result)) {
        const fonts = result.filter(isFamily).map((item) => ({ ...item, localized: item.localized ?? null }));
        if (fonts.length) return fonts;
      }
    } catch {
      // an older app or a platform without the list: the usual fonts
    }
  }
  return COMMON_FONTS;
}

/** The system's font families (read once). */
export function systemFontFamilies(): Promise<SystemFontFamily[]> {
  cached ??= readSystemFonts();
  return cached;
}

/** A family's name in the menus: 宋体 (SimSun) in Chinese, SimSun (宋体) in English. */
export function fontLabel(font: SystemFontFamily | { family: string; localized?: string | null }, isEn: boolean): string {
  if (!font.localized) return font.family;
  return isEn ? `${font.family} (${font.localized})` : font.localized;
}

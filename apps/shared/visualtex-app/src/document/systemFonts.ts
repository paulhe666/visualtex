// The fonts installed on the computer, for the document's body font menus and the formula faces
// (docs/document-editor/DESIGN.md §8.1): the desktop app reads the system's font files (list_system_font_families);
// a browser lists the fonts the user's system lets it use (the Local Font Access API of Chrome and Edge, once the user
// allows it: asked when a font menu opens), each font's file read for its Chinese characters and Chinese name; a
// browser without it gets the usual fonts of Windows, macOS and Linux that it finds installed.
import { invoke, isTauri } from "@tauri-apps/api/core";
import { readLocalStorage, writeLocalStorage } from "../runtime/safeStorage";

export interface SystemFontFamily {
  /** the name the font is asked for by (CSS, LaTeX) */
  family: string;
  /** its Chinese name, when it has one (宋体 for SimSun) */
  localized: string | null;
  /** the font has Chinese characters */
  cjk: boolean;
}

const cjkFont = (family: string, localized: string | null = null): SystemFontFamily => ({ family, localized, cjk: true });
const latinFont = (family: string): SystemFontFamily => ({ family, localized: null, cjk: false });

/** Fonts most Chinese Windows and macOS systems have (the list when nothing else can be read). */
const COMMON_FONTS: SystemFontFamily[] = [
  cjkFont("SimSun", "宋体"),
  cjkFont("NSimSun", "新宋体"),
  cjkFont("SimHei", "黑体"),
  cjkFont("KaiTi", "楷体"),
  cjkFont("FangSong", "仿宋"),
  cjkFont("Microsoft YaHei", "微软雅黑"),
  cjkFont("DengXian", "等线"),
  cjkFont("Songti SC", "宋体-简"),
  cjkFont("PingFang SC", "苹方-简"),
  cjkFont("Source Han Serif SC", "思源宋体"),
  cjkFont("Source Han Sans SC", "思源黑体"),
  latinFont("Times New Roman"),
  latinFont("Arial"),
  latinFont("Calibri"),
  latinFont("Cambria"),
  latinFont("Georgia"),
  latinFont("Garamond"),
  latinFont("Palatino Linotype"),
  latinFont("Book Antiqua"),
  latinFont("Segoe UI"),
  latinFont("Verdana"),
  latinFont("Consolas"),
  latinFont("Courier New"),
];

/** The usual fonts of Windows, macOS, Linux and Office, and the common free ones: looked for in a browser that cannot
 *  list its fonts. */
const KNOWN_FONTS: SystemFontFamily[] = [
  ...COMMON_FONTS,
  cjkFont("Microsoft JhengHei", "微軟正黑體"),
  cjkFont("MingLiU", "細明體"),
  cjkFont("PMingLiU", "新細明體"),
  cjkFont("DFKai-SB", "標楷體"),
  cjkFont("YouYuan", "幼圆"),
  cjkFont("LiSu", "隶书"),
  cjkFont("STSong", "华文宋体"),
  cjkFont("STZhongsong", "华文中宋"),
  cjkFont("STKaiti", "华文楷体"),
  cjkFont("STFangsong", "华文仿宋"),
  cjkFont("STHeiti", "华文黑体"),
  cjkFont("STXihei", "华文细黑"),
  cjkFont("STXinwei", "华文新魏"),
  cjkFont("STLiti", "华文隶书"),
  cjkFont("STXingkai", "华文行楷"),
  cjkFont("STCaiyun", "华文彩云"),
  cjkFont("STHupo", "华文琥珀"),
  cjkFont("FZShuTi", "方正舒体"),
  cjkFont("FZYaoti", "方正姚体"),
  cjkFont("PingFang TC", "蘋方-繁"),
  cjkFont("Heiti SC", "黑体-简"),
  cjkFont("Kaiti SC", "楷体-简"),
  cjkFont("Hiragino Sans GB", "冬青黑体简体中文"),
  cjkFont("Lantinghei SC", "兰亭黑-简"),
  cjkFont("Yuanti SC", "圆体-简"),
  cjkFont("Baoli SC", "报隶-简"),
  cjkFont("Libian SC", "隶变-简"),
  cjkFont("Weibei SC", "魏碑-简"),
  cjkFont("Xingkai SC", "行楷-简"),
  cjkFont("Hannotate SC", "手札体-简"),
  cjkFont("HanziPen SC", "翩翩体-简"),
  cjkFont("Wawati SC", "娃娃体-简"),
  cjkFont("Noto Sans CJK SC", "思源黑体"),
  cjkFont("Noto Serif CJK SC", "思源宋体"),
  cjkFont("Noto Sans SC"),
  cjkFont("Noto Serif SC"),
  cjkFont("WenQuanYi Micro Hei", "文泉驿微米黑"),
  cjkFont("WenQuanYi Zen Hei", "文泉驿正黑"),
  cjkFont("AR PL UMing CN"),
  cjkFont("AR PL UKai CN"),
  cjkFont("LXGW WenKai", "霞鹜文楷"),
  cjkFont("HarmonyOS Sans SC"),
  cjkFont("MiSans"),
  cjkFont("Alibaba PuHuiTi", "阿里巴巴普惠体"),
  cjkFont("Yu Gothic"),
  cjkFont("Meiryo"),
  cjkFont("MS Mincho"),
  cjkFont("MS Gothic"),
  cjkFont("Malgun Gothic"),
  ...[
    "Candara", "Constantia", "Corbel", "Tahoma", "Trebuchet MS", "Lucida Console", "Lucida Sans Unicode",
    "Franklin Gothic Medium", "Gill Sans MT", "Century Gothic", "Century Schoolbook", "Bookman Old Style",
    "Arial Narrow", "Arial Black", "Impact", "Comic Sans MS", "Sylfaen", "Bahnschrift", "Cascadia Code",
    "Cascadia Mono", "Aptos", "Segoe Print", "Segoe Script", "Baskerville Old Face", "Bodoni MT", "Goudy Old Style",
    "Rockwell", "Perpetua", "Helvetica", "Helvetica Neue", "Avenir", "Avenir Next", "Futura", "Gill Sans", "Optima",
    "Palatino", "Baskerville", "Didot", "Hoefler Text", "American Typewriter", "Menlo", "Monaco", "Charter",
    "Iowan Old Style", "Seravek", "Big Caslon", "Bodoni 72", "Cochin", "Copperplate", "Courier", "Times", "Geneva",
    "Lucida Grande", "Skia", "Athelas", "DejaVu Sans", "DejaVu Serif", "DejaVu Sans Mono", "Liberation Serif",
    "Liberation Sans", "Liberation Mono", "Noto Sans", "Noto Serif", "Ubuntu", "Cantarell", "Fira Sans", "Fira Code",
    "Roboto", "Open Sans", "Source Sans Pro", "Source Serif Pro", "Source Code Pro", "Latin Modern Roman",
    "TeX Gyre Termes", "TeX Gyre Pagella", "TeX Gyre Heros", "Linux Libertine O", "Libertinus Serif", "EB Garamond",
    "STIX Two Text", "Inter", "JetBrains Mono", "Lato", "Merriweather", "Crimson Text",
  ].map(latinFont),
];

/** Fonts a Chinese document most often takes (first in the menus when the system has them). */
export const PREFERRED_CJK = ["SimSun", "SimHei", "KaiTi", "FangSong", "Microsoft YaHei", "DengXian", "Source Han Serif SC", "Songti SC"];
export const PREFERRED_LATIN = ["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond", "Palatino Linotype", "Latin Modern Roman"];

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

const byName = (a: SystemFontFamily, b: SystemFontFamily) => a.family.toLowerCase().localeCompare(b.family.toLowerCase());

// --------------------------------------------------------------- the desktop app --

let desktopFonts: Promise<SystemFontFamily[]> | null = null;

async function readDesktopFonts(): Promise<SystemFontFamily[]> {
  try {
    const result = await invoke<unknown>("list_system_font_families");
    if (Array.isArray(result)) {
      const fonts = result.filter(isFamily).map((item) => ({ ...item, localized: item.localized ?? null }));
      if (fonts.length) return fonts;
    }
  } catch {
    // an older app or a platform without the list: the usual fonts
  }
  return COMMON_FONTS;
}

// ------------------------------------------------------------------- a browser --

/** A face the Local Font Access API lists. */
interface LocalFontData {
  family: string;
  fullName: string;
  postscriptName: string;
  style: string;
  blob(): Promise<Blob>;
}

type QueryLocalFonts = () => Promise<LocalFontData[]>;

/** What a font file says (cached by its PostScript name: a font's file is read once). */
interface FontFileFacts {
  cjk: boolean;
  localized: string | null;
}

const FACTS_STORAGE_KEY = "visualtex.local-font-facts.v1";
const CJK_NAME = /[㐀-鿿豈-﫿]/;
const REGULAR_STYLE = /^(regular|normal|book|roman|plain|w3|w4)$/i;

/** Whether a font file maps 中 and 文 to glyphs (its cmap), and its family's Chinese name (its name table). */
async function readFontFile(blob: Blob): Promise<FontFileFacts> {
  const read = async (offset: number, length: number) => new DataView(await blob.slice(offset, offset + length).arrayBuffer());
  const facts: FontFileFacts = { cjk: false, localized: null };
  const header = await read(0, 16);
  if (header.byteLength < 12) return facts;
  // a collection (ttcf): its first font
  const directory = header.getUint32(0) === 0x74746366 && header.byteLength >= 16 ? header.getUint32(12) : 0;
  const count = (await read(directory + 4, 2)).getUint16(0);
  const records = await read(directory + 12, count * 16);
  const tables = new Map<string, { offset: number; length: number }>();
  for (let i = 0; i + 16 <= records.byteLength; i += 16) {
    const tag = String.fromCharCode(records.getUint8(i), records.getUint8(i + 1), records.getUint8(i + 2), records.getUint8(i + 3));
    tables.set(tag, { offset: records.getUint32(i + 8), length: records.getUint32(i + 12) });
  }
  const cmap = tables.get("cmap");
  if (cmap) facts.cjk = await mapsCharacters(read, cmap.offset, [0x4e2d, 0x6587]);
  const name = tables.get("name");
  if (facts.cjk && name) facts.localized = await chineseFamilyName(await read(name.offset, Math.min(name.length, 1 << 20)));
  return facts;
}

/** Whether a cmap table (format 4 or 12, Unicode) maps every one of `codes` to a glyph. */
async function mapsCharacters(read: (offset: number, length: number) => Promise<DataView>, offset: number, codes: number[]): Promise<boolean> {
  const head = await read(offset, 4);
  const count = head.getUint16(2);
  const records = await read(offset + 4, count * 8);
  const subtables: Array<{ rank: number; at: number }> = [];
  for (let i = 0; i + 8 <= records.byteLength; i += 8) {
    const platform = records.getUint16(i);
    const encoding = records.getUint16(i + 2);
    const rank = platform === 3 && encoding === 10 ? 0 : platform === 0 && (encoding === 4 || encoding === 6) ? 1 : platform === 3 && encoding === 1 ? 2 : platform === 0 ? 3 : -1;
    if (rank >= 0) subtables.push({ rank, at: offset + records.getUint32(i + 4) });
  }
  subtables.sort((a, b) => a.rank - b.rank);
  for (const { at } of subtables) {
    const format = (await read(at, 2)).getUint16(0);
    if (format === 12) {
      const groups = (await read(at + 12, 4)).getUint32(0);
      const data = await read(at + 16, Math.min(groups, 1 << 20) * 12);
      return codes.every((code) => {
        let low = 0;
        let high = data.byteLength / 12 - 1;
        while (low <= high) {
          const middle = (low + high) >> 1;
          const start = data.getUint32(middle * 12);
          const end = data.getUint32(middle * 12 + 4);
          if (code < start) high = middle - 1;
          else if (code > end) low = middle + 1;
          else return data.getUint32(middle * 12 + 8) + (code - start) !== 0;
        }
        return false;
      });
    }
    if (format === 4) {
      const length = (await read(at + 2, 2)).getUint16(0);
      const table = await read(at, length);
      const segments = table.getUint16(6) / 2;
      const ends = 14;
      const starts = ends + segments * 2 + 2;
      const deltas = starts + segments * 2;
      const ranges = deltas + segments * 2;
      return codes.every((code) => {
        for (let s = 0; s < segments; s++) {
          if (table.getUint16(ends + s * 2) < code) continue;
          const start = table.getUint16(starts + s * 2);
          if (start > code) return false;
          const delta = table.getUint16(deltas + s * 2);
          const range = table.getUint16(ranges + s * 2);
          if (range === 0) return ((code + delta) & 0xffff) !== 0;
          const at = ranges + s * 2 + range + (code - start) * 2;
          if (at + 2 > table.byteLength) return false;
          const glyph = table.getUint16(at);
          return glyph !== 0 && ((glyph + delta) & 0xffff) !== 0;
        }
        return false;
      });
    }
  }
  return false;
}

/** The family's Chinese name in a name table (typographic family first; Simplified Chinese, then the others). */
function chineseFamilyName(table: DataView): string | null {
  const count = table.getUint16(2);
  const strings = table.getUint16(4);
  const RANK: Record<number, number> = { 0x0804: 0, 0x1004: 1, 0x0404: 2, 0x0c04: 3, 0x1404: 4 };
  let best: { rank: number; value: string } | null = null;
  for (let i = 0; i < count; i++) {
    const at = 6 + i * 12;
    if (at + 12 > table.byteLength) break;
    const platform = table.getUint16(at);
    const language = table.getUint16(at + 4);
    const id = table.getUint16(at + 6);
    if (platform !== 3 || (id !== 16 && id !== 1) || !(language in RANK)) continue;
    const rank = RANK[language] * 2 + (id === 16 ? 0 : 1);
    if (best && best.rank <= rank) continue;
    const length = table.getUint16(at + 8);
    const offset = strings + table.getUint16(at + 10);
    if (offset + length > table.byteLength) continue;
    let value = "";
    for (let k = 0; k + 1 < length; k += 2) value += String.fromCharCode(table.getUint16(offset + k));
    value = value.trim();
    if (value && value.length <= 128) best = { rank, value };
  }
  return best?.value ?? null;
}

function readFactsCache(): Record<string, [number, string | null]> {
  try {
    const value = JSON.parse(readLocalStorage(FACTS_STORAGE_KEY) ?? "{}");
    return value && typeof value === "object" ? value : {};
  } catch {
    return {};
  }
}

async function localFontsPermission(): Promise<PermissionState | null> {
  try {
    return (await navigator.permissions.query({ name: "local-fonts" as PermissionName })).state;
  } catch {
    return null;
  }
}

/** The fonts the browser may use (Local Font Access), or null (no API, not allowed, not asked). */
async function readBrowserFonts(ask: boolean): Promise<SystemFontFamily[] | null> {
  const query = (window as unknown as { queryLocalFonts?: QueryLocalFonts }).queryLocalFonts;
  if (typeof query !== "function") return null;
  const permission = await localFontsPermission();
  if (permission === "denied" || (permission !== "granted" && !ask)) return null;
  let faces: LocalFontData[];
  try {
    faces = await query.call(window);
  } catch {
    return null;
  }
  // one face of each family: its regular one when it has one
  const families = new Map<string, LocalFontData>();
  for (const face of faces) {
    const family = face.family?.trim();
    if (!family || family.startsWith(".") || family.length > 128) continue;
    const kept = families.get(family);
    if (!kept || (!REGULAR_STYLE.test(kept.style) && REGULAR_STYLE.test(face.style))) families.set(family, face);
  }
  const cache = readFactsCache();
  const next: Record<string, [number, string | null]> = {};
  const list = [...families.entries()];
  const fonts: SystemFontFamily[] = [];
  // the files read a few at a time (a Chinese font's file is large)
  for (let i = 0; i < list.length; i += 4) {
    await Promise.all(
      list.slice(i, i + 4).map(async ([family, face]) => {
        const key = face.postscriptName || face.fullName || family;
        let facts: FontFileFacts | null = cache[key] ? { cjk: cache[key][0] === 1, localized: cache[key][1] } : null;
        if (!facts) {
          facts = CJK_NAME.test(family) ? { cjk: true, localized: null } : await face.blob().then(readFontFile).catch(() => ({ cjk: false, localized: null }));
        }
        next[key] = [facts.cjk ? 1 : 0, facts.localized];
        const localized = facts.localized && facts.localized !== family ? facts.localized : null;
        fonts.push({ family, localized, cjk: facts.cjk });
      }),
    );
  }
  writeLocalStorage(FACTS_STORAGE_KEY, JSON.stringify(next));
  return fonts.length ? fonts.sort(byName) : null;
}

/** The known fonts the browser draws with their own faces (a text measured in each against the generic families). */
function detectKnownFonts(): SystemFontFamily[] {
  const context = typeof document !== "undefined" ? document.createElement("canvas").getContext("2d") : null;
  if (!context) return COMMON_FONTS;
  const sample = "mmmmmmmmmmlliWW0OQ@#中文永字";
  const generics = ["monospace", "serif", "sans-serif"];
  const widths = generics.map((generic) => {
    context.font = `32px ${generic}`;
    return context.measureText(sample).width;
  });
  const found = KNOWN_FONTS.filter((font) =>
    generics.some((generic, i) => {
      context.font = `32px "${font.family}", ${generic}`;
      return Math.abs(context.measureText(sample).width - widths[i]) > 0.01;
    }),
  );
  const unique = new Map(found.map((font) => [font.family.toLowerCase(), font]));
  return unique.size ? [...unique.values()].sort(byName) : COMMON_FONTS;
}

let browserFonts: SystemFontFamily[] | null = null;
let detected: SystemFontFamily[] | null = null;
/** The reading in progress (one at a time: the permission is asked once, each file read once). */
let reading: { ask: boolean; fonts: Promise<SystemFontFamily[] | null> } | null = null;

async function readFontsInBrowser(ask: boolean): Promise<SystemFontFamily[]> {
  if (browserFonts) return browserFonts;
  if (!reading || (ask && !reading.ask)) {
    // a reading that may ask after one that may not: it goes on from that one's result
    const before = reading?.fonts ?? null;
    const fonts = (async () => (before && (await before)) || readBrowserFonts(ask))();
    const current = { ask, fonts };
    reading = current;
    void fonts.finally(() => {
      if (reading === current) reading = null;
    });
  }
  const fonts = await reading.fonts;
  if (fonts) return (browserFonts = fonts);
  return (detected ??= detectKnownFonts());
}

/**
 * The installed font families: the system's in the desktop app (read once); in a browser the ones it may use — `ask`
 * (a font menu opened) lets it ask the user once for the Local Font Access permission; until then, or without it, the
 * known fonts found installed.
 */
export function systemFontFamilies(options: { ask?: boolean } = {}): Promise<SystemFontFamily[]> {
  if (isTauri()) {
    desktopFonts ??= readDesktopFonts();
    return desktopFonts;
  }
  return readFontsInBrowser(options.ask === true);
}

/** A family's name in the menus: 宋体 (SimSun) in Chinese, SimSun (宋体) in English. */
export function fontLabel(font: SystemFontFamily | { family: string; localized?: string | null }, isEn: boolean): string {
  if (!font.localized) return font.family;
  return isEn ? `${font.family} (${font.localized})` : font.localized;
}

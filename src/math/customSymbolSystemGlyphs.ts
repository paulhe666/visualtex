import type { CustomSymbolGlyphAsset } from "./customSymbolDesignerTypes";
import type { CustomSymbolTextShape } from "./customSymbolTypes";
import {
  compileNativeSystemMathGlyphAsset,
  probeNativeSystemMathFonts,
  systemFontFamilyList,
} from "./systemMathGlyphLibrary";

export type SystemMathFontId =
  | "cambria-math"
  | "stix-two-math"
  | "latin-modern-math"
  | "xits-math"
  | "system-serif";

export interface SystemMathFontPreset {
  id: SystemMathFontId;
  labelZh: string;
  labelEn: string;
  primaryFamily: string;
  family: string;
}

export const SYSTEM_MATH_FONT_PRESETS: readonly SystemMathFontPreset[] = [
  {
    id: "cambria-math",
    labelZh: "Cambria Math（系统）",
    labelEn: "Cambria Math (system)",
    primaryFamily: "Cambria Math",
    family: "Cambria Math, STIX Two Math, Times New Roman, serif",
  },
  {
    id: "stix-two-math",
    labelZh: "STIX Two Math（系统）",
    labelEn: "STIX Two Math (system)",
    primaryFamily: "STIX Two Math",
    family: "STIX Two Math, Cambria Math, Times New Roman, serif",
  },
  {
    id: "latin-modern-math",
    labelZh: "Latin Modern Math（系统）",
    labelEn: "Latin Modern Math (system)",
    primaryFamily: "Latin Modern Math",
    family: "Latin Modern Math, STIX Two Math, Cambria Math, serif",
  },
  {
    id: "xits-math",
    labelZh: "XITS Math（系统）",
    labelEn: "XITS Math (system)",
    primaryFamily: "XITS Math",
    family: "XITS Math, STIX Two Math, Cambria Math, serif",
  },
  {
    id: "system-serif",
    labelZh: "系统数学衬线回退",
    labelEn: "System math serif fallback",
    primaryFamily: "Times New Roman",
    family: "Times New Roman, Times, serif",
  },
] as const;

export type SystemGlyphCategory =
  | "basic-italic"
  | "math-alphanumeric"
  | "greek"
  | "operators"
  | "relations"
  | "arrows"
  | "letterlike"
  | "geometry";

export const SYSTEM_GLYPH_CATEGORY_LABELS: Record<
  SystemGlyphCategory,
  { zh: string; en: string }
> = {
  "basic-italic": { zh: "字母与数字", en: "Letters & digits" },
  "math-alphanumeric": { zh: "数学字母扩展", en: "Math alphanumerics" },
  greek: { zh: "希腊字母", en: "Greek" },
  operators: { zh: "运算符", en: "Operators" },
  relations: { zh: "关系与逻辑", en: "Relations & logic" },
  arrows: { zh: "箭头", en: "Arrows" },
  letterlike: { zh: "字母式符号", en: "Letterlike" },
  geometry: { zh: "几何与技术符号", en: "Geometry & technical" },
};

export const SYSTEM_GLYPH_CATEGORIES = Object.keys(
  SYSTEM_GLYPH_CATEGORY_LABELS,
) as SystemGlyphCategory[];

export interface SystemGlyphDefinition {
  character: string;
  codePoint: number;
  category: SystemGlyphCategory;
  label: string;
}

function codePointLabel(codePoint: number) {
  return `U+${codePoint.toString(16).toUpperCase().padStart(4, "0")}`;
}

function isSupportedDisplayCharacter(character: string) {
  return !/[\p{Cc}\p{Cs}\p{Cn}]/u.test(character);
}

function definitionsFromCharacters(
  characters: Iterable<string>,
  category: SystemGlyphCategory,
) {
  const result: SystemGlyphDefinition[] = [];
  const seen = new Set<number>();
  for (const character of characters) {
    const codePoint = character.codePointAt(0);
    if (codePoint === undefined || seen.has(codePoint)) continue;
    if (!isSupportedDisplayCharacter(character)) continue;
    seen.add(codePoint);
    result.push({
      character,
      codePoint,
      category,
      label: codePointLabel(codePoint),
    });
  }
  return result;
}

function definitionsFromRanges(
  ranges: readonly (readonly [number, number])[],
  category: SystemGlyphCategory,
) {
  function* characters() {
    for (const [start, end] of ranges) {
      for (let codePoint = start; codePoint <= end; codePoint += 1) {
        yield String.fromCodePoint(codePoint);
      }
    }
  }
  return definitionsFromCharacters(characters(), category);
}

const basicCharacters =
  "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
const greekCharacters =
  "ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩαβγδεζηθικλμνξοπρστυφχψωϑϕϖϱϵϰ";

const categoryDefinitions: Record<
  SystemGlyphCategory,
  readonly SystemGlyphDefinition[]
> = {
  "basic-italic": definitionsFromCharacters(basicCharacters, "basic-italic"),
  "math-alphanumeric": definitionsFromRanges(
    [
      [0x1d400, 0x1d7ff],
    ],
    "math-alphanumeric",
  ),
  greek: definitionsFromCharacters(greekCharacters, "greek"),
  operators: definitionsFromRanges(
    [
      [0x2200, 0x223f],
      [0x2290, 0x22ff],
      [0x2a00, 0x2a6f],
    ],
    "operators",
  ),
  relations: definitionsFromRanges(
    [
      [0x2240, 0x228f],
      [0x2a70, 0x2aff],
    ],
    "relations",
  ),
  arrows: definitionsFromRanges(
    [
      [0x2190, 0x21ff],
      [0x27f0, 0x27ff],
      [0x2900, 0x297f],
    ],
    "arrows",
  ),
  letterlike: definitionsFromRanges(
    [
      [0x2100, 0x214f],
    ],
    "letterlike",
  ),
  geometry: definitionsFromRanges(
    [
      [0x2300, 0x23ff],
      [0x25a0, 0x25ff],
      [0x27c0, 0x27ef],
    ],
    "geometry",
  ),
};

export function systemGlyphsForCategory(category: SystemGlyphCategory) {
  return categoryDefinitions[category];
}

export function searchSystemGlyphs(
  category: SystemGlyphCategory,
  query: string,
  limit = 2_048,
) {
  const normalized = query.trim().toLocaleUpperCase().replace(/^U\+/, "");
  const source = systemGlyphsForCategory(category);
  if (!normalized) return source.slice(0, limit);
  return source
    .filter((glyph) => {
      const hex = glyph.codePoint.toString(16).toUpperCase();
      const decimal = String(glyph.codePoint);
      return (
        glyph.character === query.trim() ||
        hex.includes(normalized) ||
        decimal.includes(normalized) ||
        glyph.label.includes(normalized)
      );
    })
    .slice(0, limit);
}

function quoteCanvasFamily(family: string) {
  return family
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean)
    .map((entry) =>
      /\s/.test(entry) && !/^['"].*['"]$/.test(entry)
        ? `"${entry.replaceAll('"', "")}"`
        : entry,
    )
    .join(", ");
}

export async function detectSystemMathFontAvailability() {
  const result = {} as Record<SystemMathFontId, boolean>;
  const nativeProbes = await probeNativeSystemMathFonts(
    SYSTEM_MATH_FONT_PRESETS.map((preset) => preset.primaryFamily),
  ).catch(() => null);
  if (nativeProbes) {
    const availability = new Map(
      nativeProbes.map((probe) => [
        probe.requestedFamily.toLocaleLowerCase(),
        probe.available,
      ]),
    );
    for (const preset of SYSTEM_MATH_FONT_PRESETS) {
      result[preset.id] =
        availability.get(preset.primaryFamily.toLocaleLowerCase()) ?? false;
    }
    return result;
  }

  if (typeof document === "undefined" || !document.fonts) {
    for (const preset of SYSTEM_MATH_FONT_PRESETS) {
      result[preset.id] = preset.id === "system-serif";
    }
    return result;
  }
  await document.fonts.ready.catch(() => undefined);
  for (const preset of SYSTEM_MATH_FONT_PRESETS) {
    result[preset.id] =
      preset.id === "system-serif" ||
      document.fonts.check(`32px "${preset.primaryFamily}"`, "∫𝑥α");
  }
  return result;
}

function characterEncodesMathematicalStyle(character: string) {
  const codePoint = character.codePointAt(0) ?? 0;
  return (
    (codePoint >= 0x1d400 && codePoint <= 0x1d7ff) ||
    [0x210e, 0x210f, 0x2113, 0x2118].includes(codePoint)
  );
}

export interface CreateSystemFontGlyphAssetOptions {
  character: string;
  font: SystemMathFontPreset;
  italic: boolean;
  fontWeight?: number;
}

export function createSystemFontGlyphAsset({
  character,
  font,
  italic,
  fontWeight = 400,
}: CreateSystemFontGlyphAssetOptions): CustomSymbolGlyphAsset {
  if (typeof document === "undefined") {
    throw new Error("System-font glyph creation requires the VisualTeX browser runtime.");
  }
  const normalized = character.normalize("NFC");
  const codePoints = Array.from(normalized);
  if (!normalized || codePoints.length !== 1 || /[\p{Cc}\p{Cs}\p{Cn}]/u.test(normalized)) {
    throw new Error("Choose one valid Unicode mathematical character.");
  }
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d");
  if (!context) throw new Error("VisualTeX could not create a font measurement canvas.");

  const fontSize = 1000;
  const effectiveItalic = italic && !characterEncodesMathematicalStyle(normalized);
  context.font = `${effectiveItalic ? "italic " : ""}${fontWeight} ${fontSize}px ${quoteCanvasFamily(font.family)}`;
  context.textBaseline = "alphabetic";
  const measurement = context.measureText(normalized);
  const left = Math.max(0, measurement.actualBoundingBoxLeft || 0);
  const right = Math.max(0, measurement.actualBoundingBoxRight || measurement.width);
  const ascent = Math.max(20, measurement.actualBoundingBoxAscent || fontSize * 0.78);
  const descent = Math.max(0, measurement.actualBoundingBoxDescent || fontSize * 0.22);
  const inkWidth = Math.max(20, left + right);
  const naturalWidth = Math.max(inkWidth, measurement.width || inkWidth);
  const padding = 36;
  const horizontalSlack = Math.max(0, naturalWidth - inkWidth);
  const x = padding + left + horizontalSlack / 2;
  const y = padding + ascent;
  const width = naturalWidth + padding * 2;

  const shape: CustomSymbolTextShape = {
    kind: "text",
    text: normalized,
    x,
    y,
    fontFamily: font.family,
    fontSize,
    fontStyle: effectiveItalic ? "italic" : "normal",
    fontWeight,
    fill: true,
  };

  return {
    sourceLatex: normalized,
    metrics: {
      widthEm: Number((width / 1000).toFixed(6)),
      ascentEm: Number((y / 1000).toFixed(6)),
      descentEm: Number(((descent + padding) / 1000).toFixed(6)),
    },
    shapes: [shape],
  };
}

export interface CreatedSystemFontGlyphAsset {
  asset: CustomSymbolGlyphAsset;
  requestedFamily: string;
  resolvedFamily: string;
  fallbackUsed: boolean;
  vectorOutline: boolean;
  warning?: string;
}

function applySyntheticMathItalic(asset: CustomSymbolGlyphAsset) {
  const originX = (asset.metrics.widthEm * 1000) / 2;
  const originY =
    ((asset.metrics.ascentEm + asset.metrics.descentEm) * 1000) / 2;
  return {
    ...asset,
    shapes: asset.shapes.map((shape) => ({
      ...shape,
      transform: {
        ...shape.transform,
        skewXDeg: (shape.transform?.skewXDeg ?? 0) - 12,
        originX: shape.transform?.originX ?? originX,
        originY: shape.transform?.originY ?? originY,
      },
    })),
  } as CustomSymbolGlyphAsset;
}

export async function createSystemFontGlyphAssetAsync(
  options: CreateSystemFontGlyphAssetOptions,
): Promise<CreatedSystemFontGlyphAsset> {
  const requestedFamily = options.font.primaryFamily;
  const fontFamilies = systemFontFamilyList(options.font.family);
  try {
    const native = await compileNativeSystemMathGlyphAsset(
      options.character,
      fontFamilies,
    );
    if (native) {
      const asset =
        options.italic && !characterEncodesMathematicalStyle(options.character)
          ? applySyntheticMathItalic(native.asset)
          : native.asset;
      return {
        asset,
        requestedFamily,
        resolvedFamily: native.resolvedFamily,
        fallbackUsed: native.fallbackUsed,
        vectorOutline: true,
      };
    }
  } catch (error) {
    const asset = createSystemFontGlyphAsset(options);
    return {
      asset,
      requestedFamily,
      resolvedFamily: requestedFamily,
      fallbackUsed: false,
      vectorOutline: false,
      warning: error instanceof Error ? error.message : String(error),
    };
  }

  return {
    asset: createSystemFontGlyphAsset(options),
    requestedFamily,
    resolvedFamily: requestedFamily,
    fallbackUsed: false,
    vectorOutline: false,
  };
}

export function systemFontPresetById(id: SystemMathFontId) {
  return (
    SYSTEM_MATH_FONT_PRESETS.find((preset) => preset.id === id) ??
    SYSTEM_MATH_FONT_PRESETS[0]
  );
}

import type { CustomSymbolGlyphAsset } from "./customSymbolDesignerTypes";

export interface NativeSystemMathFontProbe {
  requestedFamily: string;
  resolvedFamily: string;
  available: boolean;
}

export interface NativeSystemMathGlyphAsset {
  asset: CustomSymbolGlyphAsset;
  requestedFamily: string;
  resolvedFamily: string;
  fallbackUsed: boolean;
  glyphId: number;
}

export function systemFontFamilyList(value: string) {
  const result: string[] = [];
  for (const candidate of value.split(",")) {
    const family = candidate.trim().replace(/^['"]|['"]$/g, "");
    if (!family || family.toLocaleLowerCase() === "serif") continue;
    if (!result.some((item) => item.toLocaleLowerCase() === family.toLocaleLowerCase())) {
      result.push(family);
    }
  }
  return result;
}

/**
 * The browser editor deliberately has no native font bridge. Returning null
 * lets customSymbolSystemGlyphs use document.fonts for availability checks.
 */
export async function probeNativeSystemMathFonts(
  _fontFamilies: readonly string[],
): Promise<NativeSystemMathFontProbe[] | null> {
  return null;
}

/**
 * Native glyph outlining is a desktop-only capability. The browser caller
 * falls back to a Canvas-measured text glyph, so no local font data or Tauri
 * command crosses the web boundary.
 */
export async function compileNativeSystemMathGlyphAsset(
  _character: string,
  _fontFamilies: readonly string[],
): Promise<NativeSystemMathGlyphAsset | null> {
  return null;
}

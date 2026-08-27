import { invoke, isTauri } from "@tauri-apps/api/core";
import type { CustomSymbolGlyphAsset } from "./customSymbolDesignerTypes";

export interface NativeSystemMathFontProbe {
  requestedFamily: string;
  resolvedFamily: string;
  available: boolean;
}

interface NativeSystemMathGlyphOutline {
  character: string;
  requestedFamily: string;
  resolvedFamily: string;
  fallbackUsed: boolean;
  glyphId: number;
  path: string;
  metrics: {
    widthEm: number;
    ascentEm: number;
    descentEm: number;
  };
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

export async function probeNativeSystemMathFonts(
  fontFamilies: readonly string[],
): Promise<NativeSystemMathFontProbe[] | null> {
  if (!isTauri()) return null;
  return invoke<NativeSystemMathFontProbe[]>("probe_system_math_fonts", {
    fontFamilies: [...fontFamilies],
  });
}

export async function compileNativeSystemMathGlyphAsset(
  character: string,
  fontFamilies: readonly string[],
): Promise<NativeSystemMathGlyphAsset | null> {
  if (!isTauri()) return null;
  const outline = await invoke<NativeSystemMathGlyphOutline>(
    "extract_system_math_glyph",
    {
      fontFamilies: [...fontFamilies],
      character,
    },
  );
  return {
    asset: {
      sourceLatex: outline.character,
      metrics: outline.metrics,
      shapes: [
        {
          kind: "path",
          d: outline.path,
          fill: true,
        },
      ],
    },
    requestedFamily: outline.requestedFamily,
    resolvedFamily: outline.resolvedFamily,
    fallbackUsed: outline.fallbackUsed,
    glyphId: outline.glyphId,
  };
}

import { gunzipSync, strFromU8 } from "fflate";
import { ESINT_GLYPHS_GZIP_BASE64 } from "./esintGlyphs.generatedData.ts";

export interface EsintGlyphBounds {
  xMin: number;
  xMax: number;
  yMin: number;
  yMax: number;
}

export interface EsintGlyphVariant {
  glyphName: string;
  slot: number;
  path: string;
  mathJaxPath: string;
  advanceWidth: number;
  leftSideBearing: number;
  italicCorrection: number;
  height: number;
  depth: number;
  tfmHeight: number;
  tfmDepth: number;
  bounds: EsintGlyphBounds;
}

export interface EsintGlyphDefinition {
  command: string;
  aliases: string[];
  character: string;
  small: EsintGlyphVariant;
  large: EsintGlyphVariant;
}

export interface EsintGlyphPayload {
  source: {
    family: string;
    package: string;
    version: string;
    pfbSha256: string;
    tfmSha256: string;
    license: string;
  };
  unitsPerEm: number;
  glyphs: EsintGlyphDefinition[];
}

function decodeBase64(value: string) {
  const binary = globalThis.atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

export const ESINT_GLYPH_PAYLOAD = Object.freeze(
  JSON.parse(
    strFromU8(gunzipSync(decodeBase64(ESINT_GLYPHS_GZIP_BASE64))),
  ) as EsintGlyphPayload,
);

export const ESINT_INTEGRAL_GLYPH_UNITS_PER_EM =
  ESINT_GLYPH_PAYLOAD.unitsPerEm;

const ESINT_COMPATIBILITY_ALIASES: Readonly<Record<string, readonly string[]>> =
  Object.freeze({
    // MathLive and UnicodeMath use \\intclockwise for the same clockwise
    // contour operator that esint names \\ointclockwise.
    ointclockwise: ["intclockwise"],
  });

export const ESINT_INTEGRAL_GLYPHS = Object.freeze(
  ESINT_GLYPH_PAYLOAD.glyphs.map((glyph) => ({
    ...glyph,
    aliases: Array.from(
      new Set([
        ...glyph.aliases,
        ...(ESINT_COMPATIBILITY_ALIASES[glyph.command] ?? []),
      ]),
    ),
  })),
);

export const ESINT_INTEGRAL_GLYPHS_BY_COMMAND: Readonly<
  Record<string, EsintGlyphDefinition>
> = Object.freeze(
  Object.fromEntries(
    ESINT_INTEGRAL_GLYPHS.flatMap((glyph) =>
      [glyph.command, ...glyph.aliases].map((command) => [command, glyph]),
    ),
  ),
);

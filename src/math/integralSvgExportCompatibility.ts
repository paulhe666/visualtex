import { gunzipSync, strFromU8 } from "fflate";
import {
  OIINT_SIZE1_OVAL_PATH,
  OIINT_SIZE2_OVAL_PATH,
  OIIINT_SIZE1_OVAL_PATH,
  OIIINT_SIZE2_OVAL_PATH,
} from "./integralGlyphs.ts";
import { RARE_INTEGRAL_GLYPHS_GZIP_BASE64 } from "./rareIntegralGlyphs.generatedData.ts";
import { ESINT_INTEGRAL_GLYPHS } from "./esintGlyphs.ts";

interface RareIntegralGlyphVariant {
  path: string;
}

interface RareIntegralGlyphDefinition {
  command: string;
  aliases: string[];
  small: RareIntegralGlyphVariant;
  large: RareIntegralGlyphVariant;
}

interface RareIntegralGlyphPayload {
  glyphs: RareIntegralGlyphDefinition[];
}

let rareIntegralGlyphs: Map<string, RareIntegralGlyphDefinition> | null = null;

function decodeBase64(value: string) {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}

function getRareIntegralGlyphs() {
  if (rareIntegralGlyphs) return rareIntegralGlyphs;
  const payload = JSON.parse(
    strFromU8(gunzipSync(decodeBase64(RARE_INTEGRAL_GLYPHS_GZIP_BASE64))),
  ) as RareIntegralGlyphPayload;
  const result = new Map<string, RareIntegralGlyphDefinition>();
  for (const definition of payload.glyphs) {
    result.set(definition.command, definition);
    for (const alias of definition.aliases) result.set(alias, definition);
  }
  // Official esint10 outlines override similarly named Unicode/STIX glyphs.
  for (const definition of ESINT_INTEGRAL_GLYPHS) {
    result.set(definition.command, definition);
    for (const alias of definition.aliases) result.set(alias, definition);
  }
  rareIntegralGlyphs = result;
  return result;
}

const markedOperator =
  /<g\b([^>]*\bclass="[^"]*\bvisualtex-integral-export-([A-Za-z]+)\b[^"]*"[^>]*)>(<use\b[^>]*><\/use>)<\/g>/g;

function contourOverlay(command: "oiint" | "oiiint", large: boolean) {
  const path = command === "oiint"
    ? large
      ? OIINT_SIZE2_OVAL_PATH
      : OIINT_SIZE1_OVAL_PATH
    : large
      ? OIIINT_SIZE2_OVAL_PATH
      : OIIINT_SIZE1_OVAL_PATH;
  // The MathJax SVG root flips the mathematical y-axis. MathLive positions the
  // display-style oval 0.08 em lower on screen, hence -80 in math coordinates.
  const transform = large ? ' transform="translate(0 -80)"' : "";
  return `<path data-visualtex-integral="${command}"${transform} d="${path}"></path>`;
}

/**
 * Replace MathJax's marked placeholder operators with the same vector geometry
 * used by the patched MathLive editor. The placeholder remains responsible for
 * operator spacing and limit placement; only its painted glyph is changed.
 */
export function applyVisualTexIntegralSvgGlyphs(svg: string, displayMode: boolean) {
  return svg.replace(
    markedOperator,
    (whole, attributes: string, command: string, placeholder: string) => {
      const large = /-LO-/.test(placeholder) || displayMode;
      if (command === "oiint" || command === "oiiint") {
        return `<g${attributes}>${placeholder}${contourOverlay(command, large)}</g>`;
      }

      const definition = getRareIntegralGlyphs().get(command);
      if (!definition) return whole;
      const variant = large ? definition.large : definition.small;
      return (
        `<g${attributes}>` +
        `<path data-visualtex-integral="${definition.command}" d="${variant.path}"></path>` +
        "</g>"
      );
    },
  );
}

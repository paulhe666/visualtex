import { createHash } from "node:crypto";
import { gunzipSync } from "node:zlib";
import { fileURLToPath } from "node:url";
import { normalizePath, type Plugin } from "vite";
import { patchVisualTexMathLiveRuntimeSafety } from "./vite.mathliveRuntimeSafety";
import {
  RARE_INTEGRAL_GLYPHS_GZIP_BASE64,
  RARE_INTEGRAL_GLYPHS_JSON_SHA256,
} from "./src/math/rareIntegralGlyphs.generatedData";

interface RareIntegralGlyphVariant {
  path: string;
  advanceWidth: number;
  italicCorrection: number;
  height: number;
  depth: number;
}

interface RareIntegralGlyphDefinition {
  command: string;
  aliases: string[];
  character: string;
  small: RareIntegralGlyphVariant;
  large: RareIntegralGlyphVariant;
}

interface RareIntegralGlyphPayload {
  unitsPerEm: number;
  axisHeight: number;
  source: {
    family: string;
    version: string;
    tag: string;
    sha256: string;
    url: string;
    license: string;
  };
  glyphs: RareIntegralGlyphDefinition[];
}

function loadRareIntegralGlyphPayload(): RareIntegralGlyphPayload {
  const json = gunzipSync(
    Buffer.from(RARE_INTEGRAL_GLYPHS_GZIP_BASE64, "base64"),
  );
  const digest = createHash("sha256").update(json).digest("hex");
  if (digest !== RARE_INTEGRAL_GLYPHS_JSON_SHA256) {
    throw new Error(
      "VisualTeX rare-integral glyph data checksum mismatch: " +
        `${digest} != ${RARE_INTEGRAL_GLYPHS_JSON_SHA256}`,
    );
  }
  const payload = JSON.parse(json.toString("utf8")) as RareIntegralGlyphPayload;
  if (
    payload.unitsPerEm !== 1000 ||
    payload.axisHeight !== 250 ||
    payload.source.family !== "STIX Two Math" ||
    payload.source.version !== "2.13 b171" ||
    payload.source.sha256 !==
      "3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c" ||
    payload.source.license !== "SIL Open Font License 1.1" ||
    payload.glyphs.length !== 21
  ) {
    throw new Error("VisualTeX rare-integral glyph metadata is invalid.");
  }
  return payload;
}

const rareIntegralPayload = loadRareIntegralGlyphPayload();

/**
 * The MathLive browser entry must be pinned explicitly so both the desktop
 * editor and the native Windows Office dialog receive the same build-time
 * symbol registration patch.
 */
export const mathLiveBrowserEntry = fileURLToPath(
  new URL("./node_modules/mathlive/mathlive.mjs", import.meta.url),
);
const normalizedMathLiveBrowserEntry = normalizePath(mathLiveBrowserEntry);

/**
 * Register the uncommon integral glyphs used by the macOS implementation in
 * MathLive 0.109. This patch is intentionally limited to MathLive rendering;
 * it does not import any macOS Office, VBA, document-import, or lifecycle code.
 */
export function visualTexMathLiveContourIntegralCompatibility(): Plugin {
  const mathLiveVariant = (variant: RareIntegralGlyphVariant) => ({
    path: variant.path,
    advanceWidth: variant.advanceWidth,
    italicCorrection: variant.italicCorrection,
    height: variant.height,
    depth: variant.depth,
  });
  const rareIntegralGlyphs = Object.fromEntries(
    rareIntegralPayload.glyphs.flatMap((definition) =>
      [definition.command, ...definition.aliases].map((command) => [
        command,
        {
          character: definition.character,
          small: mathLiveVariant(definition.small),
          large: mathLiveVariant(definition.large),
        },
      ]),
    ),
  );
  const rareIntegralGlyphSource = JSON.stringify(rareIntegralGlyphs);
  const rareIntegralCharacters = Object.fromEntries(
    Object.entries(rareIntegralGlyphs).map(([command, definition]) => [
      command,
      definition.character,
    ]),
  );
  const rareIntegralCharacterEntries = Object.entries(rareIntegralCharacters)
    .filter(
      ([command]) =>
        ![
          "intclockwise",
          "varointclockwise",
          "ointctrclockwise",
          "intctrclockwise",
        ].includes(command),
    )
    .map(
      ([command, character]) =>
        `  ${JSON.stringify(command)}: ${JSON.stringify(character)},`,
    )
    .join("\n");

  const svgBodyAnchor = "function svgBodyToMarkup(svgBodyName) {";
  const baseAnchor = [
    '    const large = context.isDisplayStyle && this.value !== "\\\\smallint";',
    "    const base = new Box(this.value, {",
  ].join("\n");
  const classAnchor =
    '      classes: "ML__op-symbol " + (large ? "ML__large-op" : "ML__small-op"),';
  const metricsAnchor = [
    "    });",
    "    if (!base) return null;",
    "    base.right = base.italic;",
  ].join("\n");
  const symbolsAnchor = "var EXTENSIBLE_SYMBOLS = {";

  return {
    name: "visualtex-mathlive-integral-compatibility",
    enforce: "pre",
    transform(source, id) {
      if (normalizePath(id.split("?", 1)[0]) !== normalizedMathLiveBrowserEntry) {
        return null;
      }

      const svgBodyMatches = source.split(svgBodyAnchor).length - 1;
      const baseMatches = source.split(baseAnchor).length - 1;
      const classMatches = source.split(classAnchor).length - 1;
      const metricsMatches = source.split(metricsAnchor).length - 1;
      const symbolsMatches = source.split(symbolsAnchor).length - 1;
      if (
        svgBodyMatches !== 1 ||
        baseMatches !== 1 ||
        classMatches !== 1 ||
        metricsMatches !== 1 ||
        symbolsMatches !== 1
      ) {
        throw new Error(
          "MathLive integral patch anchors changed " +
            `(${svgBodyMatches}/${baseMatches}/${classMatches}/${metricsMatches}/${symbolsMatches}).`,
        );
      }

      const patched = patchVisualTexMathLiveRuntimeSafety(
        source
        .replace(
          svgBodyAnchor,
          [
            `var VISUALTEX_RARE_INTEGRAL_GLYPHS = ${rareIntegralGlyphSource};`,
            `var VISUALTEX_RARE_INTEGRAL_UNITS_PER_EM = ${rareIntegralPayload.unitsPerEm};`,
            "function visualTexRareIntegralSvgBodyToMarkup(svgBodyName) {",
            '  if (!svgBodyName.startsWith("visualtex-integral:")) return null;',
            '  const parts = svgBodyName.split(":");',
            "  const definition = VISUALTEX_RARE_INTEGRAL_GLYPHS[parts[1]];",
            '  const glyph = definition == null ? void 0 : definition[parts[2] === "large" ? "large" : "small"];',
            "  if (!glyph) return null;",
            "  const units = VISUALTEX_RARE_INTEGRAL_UNITS_PER_EM;",
            "  const width = glyph.advanceWidth / units;",
            "  const height = glyph.height / units;",
            "  const depth = glyph.depth / units;",
            "  const totalHeight = height + depth;",
            '  return `<span class="visualtex-integral-svg-body" style="display:inline-block;width:${width}em;height:${totalHeight}em;vertical-align:${-depth}em"><svg aria-hidden="true" focusable="false" width="${width}em" height="${totalHeight}em" viewBox="0 ${-glyph.height} ${glyph.advanceWidth} ${glyph.height + glyph.depth}" preserveAspectRatio="xMinYMin meet" style="display:block;overflow:visible"><path fill="currentColor" transform="scale(1,-1)" d="${glyph.path}"></path></svg></span>`;',
            "}",
            svgBodyAnchor,
            "  const visualTexIntegralMarkup = visualTexRareIntegralSvgBodyToMarkup(svgBodyName);",
            "  if (visualTexIntegralMarkup) return visualTexIntegralMarkup;",
          ].join("\n"),
        )
        .replace(
          baseAnchor,
          [
            '    const large = context.isDisplayStyle && this.value !== "\\\\smallint";',
            '    const visualTexIntegralCommand = typeof this.command === "string" ? this.command.slice(1) : "";',
            "    const visualTexIntegralDefinition = VISUALTEX_RARE_INTEGRAL_GLYPHS[visualTexIntegralCommand];",
            '    const visualTexIntegralGlyph = visualTexIntegralDefinition == null ? void 0 : visualTexIntegralDefinition[large ? "large" : "small"];',
            '    const visualTexContourValue = this.value === "\\u222F" ? "\\u222C" : this.value === "\\u2230" ? "\\u222D" : this.value;',
            '    const visualTexContourClass = visualTexIntegralGlyph ? " visualtex-integral-svg" : this.value === "\\u222F" ? " visualtex-oiint" : this.value === "\\u2230" ? " visualtex-oiiint" : "";',
            '    const base = new Box(visualTexIntegralGlyph ? "" : visualTexContourValue, {',
          ].join("\n"),
        )
        .replace(
          classAnchor,
          '      classes: "ML__op-symbol " + (large ? "ML__large-op" : "ML__small-op") + visualTexContourClass,',
        )
        .replace(
          metricsAnchor,
          [
            "    });",
            "    if (!base) return null;",
            "    if (visualTexIntegralGlyph) {",
            "      const units = VISUALTEX_RARE_INTEGRAL_UNITS_PER_EM;",
            "      base.width = visualTexIntegralGlyph.advanceWidth / units;",
            "      base.height = visualTexIntegralGlyph.height / units;",
            "      base.depth = visualTexIntegralGlyph.depth / units;",
            "      base.italic = visualTexIntegralGlyph.italicCorrection / units;",
            "      base.skew = 0;",
            '      base.svgBody = `visualtex-integral:${visualTexIntegralCommand}:${large ? "large" : "small"}`;',
            '      base.setStyle("display", "inline-block");',
            "    }",
            "    base.right = base.italic;",
          ].join("\n"),
        )
        .replace(
          symbolsAnchor,
          [symbolsAnchor, rareIntegralCharacterEntries].join("\n"),
        ),
      );

      return { code: patched, map: null };
    },
  };
}

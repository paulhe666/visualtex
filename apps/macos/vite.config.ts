import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import { fileURLToPath } from "node:url";
import {
  RARE_INTEGRAL_GLYPHS,
  RARE_INTEGRAL_GLYPH_UNITS_PER_EM,
} from "./src/math/rareIntegralGlyphs.generated.ts";
import {
  ESINT_INTEGRAL_GLYPHS,
  ESINT_INTEGRAL_GLYPH_UNITS_PER_EM,
} from "./src/math/esintGlyphs.ts";
import { patchVisualTexMathLiveRuntimeSafety } from "./vite.mathliveRuntimeSafety";

const host = process.env.TAURI_DEV_HOST;
const mathLiveBrowserEntry = fileURLToPath(
  new URL("./node_modules/mathlive/mathlive.mjs", import.meta.url),
);

function visualTexMathLiveIntegralCompatibility() {
  const mathLiveVariant = (
    variant: {
      path: string;
      advanceWidth: number;
      italicCorrection: number;
      height: number;
      depth: number;
    },
  ) => ({
    path: variant.path,
    advanceWidth: variant.advanceWidth,
    italicCorrection: variant.italicCorrection,
    height: variant.height,
    depth: variant.depth,
  });
  if (ESINT_INTEGRAL_GLYPH_UNITS_PER_EM !== RARE_INTEGRAL_GLYPH_UNITS_PER_EM) {
    throw new Error("VisualTeX integral glyph registries use incompatible units.");
  }
  // esint definitions intentionally come last: where a package command also
  // has a similarly named Unicode/STIX glyph (notably \\fint), the official
  // esint10 outline is the source of truth.
  const integralGlyphDefinitions = [
    ...RARE_INTEGRAL_GLYPHS,
    ...ESINT_INTEGRAL_GLYPHS,
  ];
  const rareIntegralGlyphs = Object.fromEntries(
    integralGlyphDefinitions.flatMap((definition) =>
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
    enforce: "pre" as const,
    transform(source: string, id: string) {
      if (id.split("?", 1)[0] !== mathLiveBrowserEntry) return null;

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
            `var VISUALTEX_RARE_INTEGRAL_UNITS_PER_EM = ${RARE_INTEGRAL_GLYPH_UNITS_PER_EM};`,
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

export default defineConfig({
  plugins: [visualTexMathLiveIntegralCompatibility(), react()],
  resolve: {
    alias: [
      {
        find: /^mathlive$/,
        replacement: mathLiveBrowserEntry,
      },
    ],
  },
  optimizeDeps: {
    exclude: ["mathlive"],
  },
  clearScreen: false,
  server: {
    port: 1420,
    strictPort: true,
    host: host || false,
    hmr: host
      ? { protocol: "ws", host, port: 1421 }
      : undefined,
  },
  envPrefix: ["VITE_", "TAURI_ENV_*"],
  build: {
    target: process.env.TAURI_ENV_PLATFORM === "windows" ? "chrome105" : "safari13",
    minify: process.env.TAURI_ENV_DEBUG ? false : "esbuild",
    sourcemap: !!process.env.TAURI_ENV_DEBUG,
    rollupOptions: {
      input: {
        main: "index.html",
        officeNativeDialog: "office-native-dialog.html",
      },
    },
  },
});

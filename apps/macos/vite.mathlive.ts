import { createHash } from "node:crypto";
import { readFileSync, readdirSync } from "node:fs";
import { gunzipSync } from "node:zlib";
import { fileURLToPath } from "node:url";
import { type Plugin } from "vite";
import { commandRegistry } from "./src/autocomplete/commandRegistry";
import { VISUALTEX_CORE_MATHLIVE_MACROS } from "./src/math/coreLatexAliases";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "./src/math/physicsKernelMacros";
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

export const mathLiveBrowserEntry = "virtual:visualtex-mathlive-kernel";
const resolvedMathLiveBrowserEntry = "\0visualtex-mathlive-kernel";
const mathLiveKernelPartsDirectory = fileURLToPath(
  new URL("./vendor/mathlive/kernel-parts/", import.meta.url),
);

export function loadVisualTexMathLiveKernel() {
  const parts = readdirSync(mathLiveKernelPartsDirectory)
    .filter((name) => /^part-\d{3}\.mjsfrag$/.test(name))
    .sort();
  if (parts.length === 0 || parts.some((name, index) => name !== `part-${String(index).padStart(3, "0")}.mjsfrag`)) {
    throw new Error("VisualTeX MathLive kernel source fragments are missing or non-contiguous.");
  }
  return parts
    .map((name) =>
      readFileSync(
        new URL("./vendor/mathlive/kernel-parts/" + name, import.meta.url),
        "utf8",
      ),
    )
    .join("");
}


// The kernel owns behavior. These are shared product data, also consumed by
// the toolbar, source validator and export pipeline; keep a single definition.
export function mathLiveKernelData() {
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

  const macros = Object.fromEntries(
    [...VISUALTEX_CORE_MATHLIVE_MACROS, ...VISUALTEX_PHYSICS_KERNEL_MACROS]
      .map(({ name, def, args }) => [name, { def, args, expand: false, captureSelection: false }]),
  );
  const completionCatalog = commandRegistry.map(command => ({
    id: command.id, command: command.command, latex: command.insertTemplate,
    preview: command.previewLatex, label: "",
    words: [command.labelEn, ...command.aliases, ...command.keywords],
    priority: command.defaultPriority,
  }));
  const integralCharacters = Object.fromEntries(Object.entries(rareIntegralGlyphs)
    .filter(([name]) => !["intclockwise", "varointclockwise", "ointctrclockwise", "intctrclockwise"].includes(name))
    .map(([name, glyph]) => [name, glyph.character]));
  return { macros, completionCatalog, integralCharacters, integralGlyphs: rareIntegralGlyphs,
    integralUnitsPerEm: rareIntegralPayload.unitsPerEm };
}

export function visualTexMathLiveKernel(): Plugin {
  const dataId = "virtual:visualtex-mathlive-data";
  return {
    name: "visualtex-mathlive-kernel",
    resolveId(id) {
      if (id === mathLiveBrowserEntry) return resolvedMathLiveBrowserEntry;
      if (id === dataId) return "\0" + dataId;
      return null;
    },
    load(id) {
      if (id === "\0" + dataId) {
        return Object.entries(mathLiveKernelData())
          .map(([name, value]) => `export const ${name} = ${JSON.stringify(value)};`).join("\n");
      }
      if (id !== resolvedMathLiveBrowserEntry) return null;
      for (const name of readdirSync(mathLiveKernelPartsDirectory)) {
        this.addWatchFile(fileURLToPath(new URL("./vendor/mathlive/kernel-parts/" + name, import.meta.url)));
      }
      return loadVisualTexMathLiveKernel();
    },
  };
}

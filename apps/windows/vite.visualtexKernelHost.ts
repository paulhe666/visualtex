import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { gunzipSync } from "node:zlib";
import type { Plugin } from "vite";
import { commandRegistry } from "./src/autocomplete/commandRegistry";
import { latestSharedCommands } from "./src/autocomplete/latestSharedCommands";
import { COMMAND_FREQUENCY } from "./src/autocomplete/commandFrequency.generated";
import { sourceWrapperCanonicalTargets, sourceWrapperKinds } from "./src/editor/sourceWrapperRegistry";
import { VISUALTEX_CORE_MATHLIVE_MACROS } from "./src/math/coreLatexAliases";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "./src/math/physicsKernelMacros";
import {
  RARE_INTEGRAL_GLYPHS_GZIP_BASE64,
  RARE_INTEGRAL_GLYPHS_JSON_SHA256,
} from "./src/math/rareIntegralGlyphs.generatedData";

/**
 * The host module of the editing kernel (apps/shared/visualtex-core/src/kernel):
 * `virtual:visualtex-kernel-host` gives the kernel the tables generated from
 * this app's registries and the app's input judgment, so the kernel never
 * copies them by hand.
 */
export const kernelHostModuleId = "virtual:visualtex-kernel-host";
const resolvedKernelHostModuleId = "\0visualtex-kernel-host";
const inputJudgmentModule = fileURLToPath(new URL("./src/editor/inputJudgment.ts", import.meta.url)).replace(/\\/g, "/");

interface RareIntegralGlyphDefinition {
  command: string;
  aliases: string[];
  character: string;
}

function rareIntegralGlyphs(): RareIntegralGlyphDefinition[] {
  const json = gunzipSync(Buffer.from(RARE_INTEGRAL_GLYPHS_GZIP_BASE64, "base64"));
  const digest = createHash("sha256").update(json).digest("hex");
  if (digest !== RARE_INTEGRAL_GLYPHS_JSON_SHA256) {
    throw new Error(`VisualTeX rare-integral glyph data checksum mismatch: ${digest} != ${RARE_INTEGRAL_GLYPHS_JSON_SHA256}`);
  }
  const payload = JSON.parse(json.toString("utf8"));
  if (
    payload.unitsPerEm !== 1000 ||
    payload.axisHeight !== 250 ||
    payload.source.family !== "STIX Two Math" ||
    payload.source.version !== "2.13 b171" ||
    payload.source.sha256 !== "3a5f3f26f40d5698b3c62dd085d48d6663696a3f80825aab8b553d5097518e8c" ||
    payload.source.license !== "SIL Open Font License 1.1" ||
    payload.glyphs.length !== 21
  ) {
    throw new Error("VisualTeX rare-integral glyph metadata is invalid.");
  }
  return payload.glyphs;
}

/** The uncommon integrals of the macOS implementation as large operators of the kernel's symbol table. */
function rareIntegralSymbols() {
  // esint spellings with no STIX glyph of their own map to the matching STIX operator
  const editorGlyphAliases: Record<string, string[]> = { ointctrclockwise: ["varointctrclockwise"] };
  const characters = Object.fromEntries(
    rareIntegralGlyphs().flatMap((definition) =>
      [definition.command, ...definition.aliases, ...(editorGlyphAliases[definition.command] ?? [])].map(
        (command) => [command, definition.character],
      ),
    ),
  );
  return Object.entries(characters)
    .filter(([command]) => !["intclockwise", "varointclockwise", "ointctrclockwise", "intctrclockwise"].includes(command))
    .map(([command, character]) => `  ${JSON.stringify(command)}: ${JSON.stringify(character)},`)
    .join("\n");
}

/** VisualTeX's common LaTeX compatibility in the parser's default macro dictionary. */
function defaultMacros() {
  return [...VISUALTEX_CORE_MATHLIVE_MACROS, ...VISUALTEX_PHYSICS_KERNEL_MACROS]
    .map(
      (macro) =>
        `  ${JSON.stringify(macro.name)}: { def: ${JSON.stringify(macro.def)}, args: ${macro.args}, expand: false, captureSelection: false },`,
    )
    .join("\n");
}

/** The command catalog of the kernel's completion list. */
function completionCatalog() {
  const nativeCompletionTemplateOverrides = new Map<string, string>([
    ["\\begin{cases}", "\\begin{cases}\\placeholder{} & \\placeholder{}\\end{cases}"],
  ]);
  // Shared variants (\bigcap with limits, three-row cases, ...) are found by
  // their command but never claim an exactly typed command: \bigcap + Space
  // still inserts the bare operator, the variant follows it.
  const variantIds = new Set(latestSharedCommands.map((c) => c.id));
  const records = commandRegistry.map((c) => ({
    exact: !variantIds.has(c.id),
    id: c.id,
    command: c.command,
    latex: nativeCompletionTemplateOverrides.get(c.command) ?? c.insertTemplate,
    preview: c.previewLatex,
    // Search can still use translated keywords, but the native list only
    // presents the command and its mathematical preview.
    label: "",
    words: [c.labelEn, ...c.aliases, ...c.keywords],
    priority: c.defaultPriority,
  }));
  return JSON.stringify(records);
}

/** The host module's source. */
export function kernelHostModuleSource(): string {
  return [
    `export { VISUALTEX_SYMBOL_SEQUENCES, applyInputJudgment } from ${JSON.stringify(inputJudgmentModule)};`,
    // the parser's default macros start with these (a later built-in of the same name wins, as before)
    `export const visualTexDefaultMacros = {\n${defaultMacros()}\n};`,
    `export const visualTexCompletionCatalogData = ${completionCatalog()};`,
    `export const visualTexCommandFrequencyData = ${JSON.stringify(COMMAND_FREQUENCY)};`,
    `export const visualTexRareIntegralSymbols = {\n${rareIntegralSymbols()}\n};`,
    `export const visualTexSourceWrapperKindsData = ${JSON.stringify(Object.fromEntries(sourceWrapperKinds))};`,
    `export const visualTexSourceWrapperCanonicalTargetsData = ${JSON.stringify(Object.fromEntries(sourceWrapperCanonicalTargets))};`,
    "",
  ].join("\n");
}

/** Vite plugin serving the host module. */
export function visualTexKernelHost(): Plugin {
  return {
    name: "visualtex-kernel-host",
    resolveId(source) {
      return source === kernelHostModuleId ? resolvedKernelHostModuleId : null;
    },
    load(id) {
      return id === resolvedKernelHostModuleId ? kernelHostModuleSource() : null;
    },
  };
}

/** The editing kernel's entry, imported by the app as `mathlive`. */
export const visualTexKernelEntry = fileURLToPath(
  new URL("../shared/visualtex-core/src/kernel/mathlive.ts", import.meta.url),
);

/** The kernel's DOM-free entry (parsing, validation, markup), imported as `mathlive/ssr`. */
export const visualTexKernelSsrEntry = fileURLToPath(
  new URL("../shared/visualtex-core/src/kernel/public/mathlive-ssr.ts", import.meta.url),
);

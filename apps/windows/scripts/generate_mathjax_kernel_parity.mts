// Writes src/math/kernelParityMacros.generated.ts: every command the editing
// kernel accepts that MathJax (the MathML behind OMML, MathType and the
// MathJax SVG) does not define, expanded to the kernel's own meaning:
//   - symbols: the kernel's code point with its TeX class,
//   - named functions (\tg, \arcsec, ...): \operatorname,
//   - macros (texvc, amsmath, braket, VisualTeX aliases): their definition.
// Commands with arguments and their own layout (\overarc, ...) are written by
// hand in latexCompatibility.ts, which takes precedence over this table.
//   tsx --import ./scripts/kernel-node.mts scripts/generate_mathjax_kernel_parity.mts [--check]
import fs from "node:fs";
import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { STATE } from "mathjax-full/js/core/MathItem.js";
import { SerializedMmlVisitor } from "mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js";
import { EXPLICIT_MATH_FONT_PACKAGE } from "../src/export/explicit-math-font.ts";
import { VISUALTEX_HANDWRITTEN_MATHML_MACROS } from "../src/math/latexCompatibility.ts";
import { VISUALTEX_CORE_MATHLIVE_MACROS } from "../src/math/coreLatexAliases.ts";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "../src/math/physicsKernelMacros.ts";
import { VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS } from "../src/math/mathLiveCompatibilityMacros.ts";
import { kernelDictionary as kernel } from "./kernel-dictionary.mts";

type Macro = string | readonly [string, number];
const output = new URL("../src/math/kernelParityMacros.generated.ts", import.meta.url);
const check = process.argv.includes("--check");

RegisterHTMLHandler(liteAdaptor());
function compiler(macros: Record<string, Macro>) {
  const document = mathjax.document("", {
    InputJax: new TeX({
      packages: [...AllPackages.filter((name) => name !== "noundefined" && name !== "noerrors"), EXPLICIT_MATH_FONT_PACKAGE],
      macros: macros as Record<string, string | [string, number]>,
      formatError: (_jax: unknown, error: unknown) => { throw error; },
    }),
    // as in runtime.ts: some packages (bussproofs) need an output jax to compile
    OutputJax: new SVG({ fontCache: "local", internalSpeechTitles: false }),
  });
  const serializer = new SerializedMmlVisitor(document.mmlFactory);
  return (latex: string) => {
    try {
      // as latexToMathMlSync: compile only; errors throw or leave an merror node
      const root = document.convert(latex, { display: true, end: STATE.COMPILED });
      const mathMl = serializer.visitTree(root as any);
      return /<merror\b/.test(mathMl) ? `merror: ${mathMl.replace(/<[^>]+>/g, " ").trim()}` : null;
    } catch (error) {
      return String((error as Error)?.message ?? error);
    }
  };
}

const usage = (name: string, args: number) => `\\${name}` + ["{x}", "{y}", "{z}", "{w}", "{v}"].slice(0, args).join("");
const undefinedHere = (error: string | null, name: string) =>
  error !== null && error.includes(`Undefined control sequence \\${name}`);
const base = compiler(VISUALTEX_HANDWRITTEN_MATHML_MACROS);
const needed = (name: string, args: number) => undefinedHere(base(usage(name, args)), name);

const table = new Map<string, { macro: Macro; source: string; fallback?: Macro }>();
const add = (name: string, macro: Macro, source: string, fallback?: Macro) => {
  if (!/^[A-Za-z]+$/.test(name) || table.has(name) || name in VISUALTEX_HANDWRITTEN_MATHML_MACROS) return;
  table.set(name, { macro, source, fallback });
};

// symbols: code point and TeX class
const TEX_SPECIAL = new Set([...'#$%&~_^\\{}']);
const CLASS_WRAPPER: Record<string, string> = {
  mrel: "\\mathrel", mbin: "\\mathbin", mop: "\\mathop", mpunct: "\\mathpunct", minner: "\\mathinner",
};
for (const [command, definition] of Object.entries(kernel.MATH_SYMBOLS as Record<string, any>)) {
  const name = command.slice(1);
  if (!command.startsWith("\\") || !definition?.codepoint || !needed(name, 0)) continue;
  const character = String.fromCodePoint(definition.codepoint);
  const atom = TEX_SPECIAL.has(character) ? `\\unicode{${definition.codepoint}}` : character;
  // opening/closing symbols stay bare so \left and \right accept them
  const wrapper = CLASS_WRAPPER[definition.type as string];
  // a character MathJax has no math glyph for (the euro sign) is set as text
  add(name, wrapper ? `${wrapper}{${atom}}` : atom, `symbol ${definition.type} U+${definition.codepoint.toString(16).toUpperCase()}`,
    TEX_SPECIAL.has(character) ? undefined : wrapper ? `${wrapper}{\\text{${character}}}` : `\\text{${character}}`);
}

// named functions without arguments
for (const [command, definition] of Object.entries(kernel.LATEX_COMMANDS as Record<string, any>)) {
  const name = command.slice(1);
  if (!definition?.isFunction || (definition.params ?? []).length || !needed(name, 0)) continue;
  add(name, `\\operatorname{${name}}`, "function");
}

// macro tables
const argumentCount = (def: string, declared: number) =>
  Math.max(declared, ...[...def.matchAll(/#(\d)/g)].map((match) => Number(match[1])), 0);
const macroSources: [string, Record<string, unknown>][] = [
  ["texvc", kernel.TEXVC_MACROS],
  ["amsmath", kernel.AMSMATH_MACROS],
  ["braket", kernel.BRAKET_MACROS],
  ["visualtex", Object.fromEntries(VISUALTEX_CORE_MATHLIVE_MACROS.map((m) => [m.name, { def: m.def, args: m.args }]))],
  ["physics", Object.fromEntries(VISUALTEX_PHYSICS_KERNEL_MACROS.map((m: any) => [m.name, { def: m.def, args: m.args }]))],
  ["mathlive-compatibility", VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS as Record<string, unknown>],
];
for (const [source, dictionary] of macroSources) {
  for (const [name, value] of Object.entries(dictionary)) {
    const def = typeof value === "string" ? value : String((value as any)?.def ?? "");
    if (!def) continue;
    const args = argumentCount(def, typeof value === "string" ? 0 : Number((value as any)?.args ?? 0));
    if (!needed(name, args)) continue;
    add(name, args ? [def, args] as const : def, source);
  }
}

// keep only expansions MathJax can compile (a definition may itself use a
// kernel-only command, which another entry of this table can provide)
let entries = [...table.entries()];
const dropped: string[] = [];
for (let pass = 0; pass < 4; pass++) {
  const compile = compiler({ ...Object.fromEntries(entries.map(([n, e]) => [n, e.macro])), ...VISUALTEX_HANDWRITTEN_MATHML_MACROS });
  const failing = entries.filter(([name, entry]) =>
    compile(usage(name, Array.isArray(entry.macro) ? entry.macro[1] : 0)) !== null);
  if (!failing.length) break;
  for (const [name, entry] of failing) {
    if (entry.fallback !== undefined) {
      entry.macro = entry.fallback;
      entry.fallback = undefined;
      entry.source += " as text";
      continue;
    }
    dropped.push(`${name} (${entry.source}): ${compile(usage(name, Array.isArray(entry.macro) ? entry.macro[1] : 0))}`);
  }
  const droppedNames = new Set(dropped.map((line) => line.split(" ")[0]));
  entries = entries.filter(([name]) => !droppedNames.has(name));
}
entries.sort(([a], [b]) => a.localeCompare(b));

const literal = (value: string) => JSON.stringify(value);
const body = entries.map(([name, { macro, source }]) =>
  `  ${/^[A-Za-z_]\w*$/.test(name) ? name : literal(name)}: ${Array.isArray(macro) ? `[${literal(macro[0])}, ${macro[1]}]` : literal(macro as string)}, // ${source}`,
).join("\n");
const text = `// Generated by scripts/generate_mathjax_kernel_parity.mts. Do not edit.
// Commands the editing kernel accepts that MathJax does not define, expanded to
// the kernel's own meaning so OMML, MathType and MathJax SVG export accept them.
import type { VisualTexCompatibilityMacro } from "./latexCompatibility.ts";

export const KERNEL_PARITY_MACROS: Record<string, VisualTexCompatibilityMacro> = {
${body}
};
`;
if (check) {
  const current = fs.existsSync(output) ? fs.readFileSync(output, "utf8").replace(/\r\n/g, "\n") : "";
  if (current !== text) {
    console.error("src/math/kernelParityMacros.generated.ts is out of date: run scripts/generate_mathjax_kernel_parity.mts");
    process.exit(1);
  }
} else {
  // the working tree checks text out with CRLF on Windows (core.autocrlf)
  fs.writeFileSync(output, process.platform === "win32" ? text.replace(/\n/g, "\r\n") : text, "utf8");
}
console.log(JSON.stringify({ entries: entries.length, dropped }, null, 1));

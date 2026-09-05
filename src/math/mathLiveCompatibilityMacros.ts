import type { MacroDictionary } from "mathlive";
import { VISUALTEX_MATHLIVE_PACKAGE_MACROS } from "./packageMacroCompatibility";

const macro = (def: string): MacroDictionary[string] => ({
  def,
  args: 1,
  expand: false,
  captureSelection: false,
});

const inputAliasMacro = (def: string): MacroDictionary[string] => ({
  def,
  args: 1,
  expand: true,
  captureSelection: false,
});

/**
 * MathLive rendering/editing compatibility shared by formula fields and static
 * previews. These are rendering aliases only: `expand: false` keeps supported
 * source spellings such as `\\bm{...}` and `\\symbfit{...}` intact.
 */
export const VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS: MacroDictionary = {
  ...VISUALTEX_MATHLIVE_PACKAGE_MACROS,

  // MathLive accepts \nicefrac as source but does not expose its two
  // arguments as stable editable atoms in every Windows runtime. Define the
  // visual expansion explicitly while preserving the canonical source token.
  nicefrac: {
    def: "{}^{#1}\\!/\\!{}_{#2}",
    args: 2,
    expand: false,
    captureSelection: false,
  },

  // The default MathLive bold wrapper is upright. Force mathematical bold
  // aliases through mathbfit so Latin variables retain their italic shape.
  boldsymbol: macro("\\mathbfit{#1}"),
  bm: macro("\\mathbfit{#1}"),

  // unicode-math alphabet aliases. MathLive 0.109 natively exposes mathbfit but
  // does not cover the full unicode-math `\\sym...` command family.
  symup: macro("\\mathrm{#1}"),
  symit: macro("\\mathit{#1}"),
  symbf: macro("\\mathbf{#1}"),
  symbfup: macro("\\mathbf{#1}"),
  symbfit: macro("\\mathbfit{#1}"),
  simbfit: inputAliasMacro("\\symbfit{#1}"),
  symbb: macro("\\mathbb{#1}"),
  symcal: macro("\\mathcal{#1}"),
  symbfcal: macro("\\mathbf{\\mathcal{#1}}"),
  symscr: macro("\\mathscr{#1}"),
  symbfscr: macro("\\mathbf{\\mathscr{#1}}"),
  symfrak: macro("\\mathfrak{#1}"),
  symbffrak: macro("\\mathbf{\\mathfrak{#1}}"),
  symsfup: macro("\\mathsf{#1}"),
  symsfit: macro("\\mathsf{\\mathit{#1}}"),
  symbfsfup: macro("\\mathbf{\\mathsf{#1}}"),
  symbfsfit: macro("\\mathbfit{#1}"),
  symtt: macro("\\mathtt{#1}"),

  // Input/declaration aliases participate in MathLive's native command
  // completion, but expand to canonical scoped commands when parsed directly.
  boldmath: inputAliasMacro("\\mathbfit{#1}"),
  bold: inputAliasMacro("\\mathbfit{#1}"),
  pmb: inputAliasMacro("\\mathbfit{#1}"),
  bf: inputAliasMacro("\\mathbf{#1}"),
  bfseries: inputAliasMacro("\\mathbf{#1}"),
  it: inputAliasMacro("\\mathit{#1}"),
  rm: inputAliasMacro("\\mathrm{#1}"),
  sf: inputAliasMacro("\\mathsf{#1}"),
  tt: inputAliasMacro("\\mathtt{#1}"),
  cal: inputAliasMacro("\\mathcal{#1}"),
  Bbb: inputAliasMacro("\\mathbb{#1}"),
  frak: inputAliasMacro("\\mathfrak{#1}"),
};

type ExpandableMacroDefinition = {
  def: string;
  args: number;
};

type MacroArgument = {
  content: string;
  end: number;
};

function compatibilityMacroDefinition(
  name: string,
  excludedCommands: ReadonlySet<string>,
): ExpandableMacroDefinition | null {
  if (excludedCommands.has(name)) return null;
  const candidate = VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS[name] as unknown;
  if (!candidate || typeof candidate !== "object") return null;
  const definition = candidate as Partial<ExpandableMacroDefinition>;
  return typeof definition.def === "string" &&
    typeof definition.args === "number" &&
    Number.isInteger(definition.args) &&
    definition.args > 0
    ? { def: definition.def, args: definition.args }
    : null;
}

function isEscapedCharacter(source: string, index: number) {
  let backslashes = 0;
  for (let cursor = index - 1; cursor >= 0 && source[cursor] === "\\"; cursor -= 1) {
    backslashes += 1;
  }
  return backslashes % 2 === 1;
}

function skipWhitespace(source: string, start: number) {
  let index = start;
  while (/\s/.test(source[index] ?? "")) index += 1;
  return index;
}

function readBalancedMacroArgument(
  source: string,
  start: number,
): MacroArgument | null {
  if (source[start] !== "{") return null;
  let depth = 0;
  for (let index = start; index < source.length; index += 1) {
    const character = source[index];
    if (isEscapedCharacter(source, index)) continue;
    if (character === "{") depth += 1;
    else if (character === "}") {
      depth -= 1;
      if (depth === 0) {
        return {
          content: source.slice(start + 1, index),
          end: index + 1,
        };
      }
    }
  }
  return null;
}

function readCompatibilityMacroArgument(
  source: string,
  start: number,
): MacroArgument | null {
  const position = skipWhitespace(source, start);
  if (position >= source.length) return null;
  if (source[position] === "{") {
    return readBalancedMacroArgument(source, position);
  }
  if (source[position] === "\\") {
    const command = source.slice(position).match(/^\\(?:[A-Za-z@]+|.)/);
    return command
      ? { content: command[0], end: position + command[0].length }
      : null;
  }
  return { content: source[position], end: position + 1 };
}

function expandCompatibilityMacros(
  source: string,
  excludedCommands: ReadonlySet<string>,
  depth: number,
): string {
  if (depth >= 16 || !source.includes("\\")) return source;
  let output = "";
  let index = 0;
  while (index < source.length) {
    if (
      source[index] === "%" &&
      !isEscapedCharacter(source, index)
    ) {
      const lineEnd = source.indexOf("\n", index);
      if (lineEnd < 0) return output + source.slice(index);
      output += source.slice(index, lineEnd + 1);
      index = lineEnd + 1;
      continue;
    }
    if (
      source[index] !== "\\" ||
      isEscapedCharacter(source, index)
    ) {
      output += source[index];
      index += 1;
      continue;
    }

    const commandMatch = source.slice(index).match(/^\\([A-Za-z@]+)/);
    if (!commandMatch) {
      output += source[index];
      index += 1;
      continue;
    }
    const commandName = commandMatch[1];
    const definition = compatibilityMacroDefinition(
      commandName,
      excludedCommands,
    );
    if (!definition) {
      output += commandMatch[0];
      index += commandMatch[0].length;
      continue;
    }

    let cursor = index + commandMatch[0].length;
    const argumentsFound: string[] = [];
    for (let argument = 0; argument < definition.args; argument += 1) {
      const parsed = readCompatibilityMacroArgument(source, cursor);
      if (!parsed) break;
      argumentsFound.push(
        expandCompatibilityMacros(
          parsed.content,
          excludedCommands,
          depth + 1,
        ),
      );
      cursor = parsed.end;
    }
    if (argumentsFound.length !== definition.args) {
      output += commandMatch[0];
      index += commandMatch[0].length;
      continue;
    }

    const replacement = definition.def.replace(
      /#([1-9])/g,
      (whole, rawIndex: string) =>
        argumentsFound[Number.parseInt(rawIndex, 10) - 1] ?? whole,
    );
    output += expandCompatibilityMacros(
      replacement,
      excludedCommands,
      depth + 1,
    );
    index = cursor;
  }
  return output;
}

/**
 * Expand only VisualTeX's compatibility aliases before static MathLive
 * rendering. Passing the compatibility dictionary through `options.macros`
 * replaces MathLive's private default macro set in 0.109.x, which would break
 * built-ins such as `\\strut` and `\\thetasym` in suggestion previews.
 */
export function expandVisualTexMathLiveCompatibilityMacros(
  source: string,
  excludedCommands: ReadonlySet<string> = new Set(),
) {
  return expandCompatibilityMacros(source, excludedCommands, 0);
}

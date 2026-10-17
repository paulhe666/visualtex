import { normalizeExtendedIntegralLatexCommands } from "../math/extendedIntegralCompatibility.ts";
import { INPUT_SHORTCUT_GROUPS, type InputShortcutGroup } from "./inputDictionary.ts";
import {
  VISUALTEX_SYMBOL_SEQUENCES,
  applyInputJudgment,
  type InputJudgmentOptions,
} from "./inputJudgment.ts";

const chineseChar = /[\u3400-\u9fff\uf900-\ufaff，。；：！？、（）【】《》“”‘’]/;

const CANONICAL_UPRIGHT_COMMANDS: ReadonlyArray<
  readonly [command: string, standardLatex: string]
> = [
  ["capitalDifferentialD", "\\mathrm{D}"],
  ["differentialD", "\\mathrm{d}"],
  ["exponentialE", "\\mathrm{e}"],
  ["imaginaryI", "\\mathrm{i}"],
  ["imaginaryJ", "\\mathrm{j}"],
];

export type VisualTexInlineShortcutDefinition =
  | string
  | { value: string; after?: string };

export type VisualTexInlineShortcutDefinitions = Record<
  string,
  VisualTexInlineShortcutDefinition
>;

export interface VisualTexAutoEscapeShortcutGroup
  extends InputShortcutGroup {}

// The settings page lists the one dictionary of the input judgment.
export const visualTexAutoEscapeShortcutGroups: readonly VisualTexAutoEscapeShortcutGroup[] =
  INPUT_SHORTCUT_GROUPS;

export const visualTexAutoEscapeInlineShortcuts: VisualTexInlineShortcutDefinitions =
  Object.assign({}, ...INPUT_SHORTCUT_GROUPS.map((group) => group.shortcuts));

/**
 * The kernel's as-you-type shortcuts: symbol sequences only (>=, ->).
 * Typed words (alpha, sin, km) are decided by the input judgment once
 * finished, so they never fire inside spin, point or for.
 */
export function resolveVisualTexInlineShortcuts(
  _mathLiveDefaults: Readonly<VisualTexInlineShortcutDefinitions>,
  autoEscapeShortcuts: boolean,
): VisualTexInlineShortcutDefinitions {
  return autoEscapeShortcuts ? { ...VISUALTEX_SYMBOL_SEQUENCES } : {};
}

export function normalizeCanonicalUprightCommands(
  source: string,
): string {
  let normalized = source;
  for (const [command, standardLatex] of CANONICAL_UPRIGHT_COMMANDS) {
    normalized = normalized.replace(
      new RegExp(`\\\\${command}(?![A-Za-z])[ \\t]*`, "g"),
      standardLatex,
    );
  }
  return normalized.replace(
    /\\(?:mathrm|textrm)\{d([A-Za-z])\}/g,
    "\\mathrm{d}$1",
  );
}

// Compatibility name retained for the restored MathLive/MathJax product path.
// The implementation is renderer-independent; this alias does not reintroduce
// any MathLive runtime behavior into normalization itself.
export const normalizeMathLiveCanonicalUprightCommands =
  normalizeCanonicalUprightCommands;

/**
 * Upright/italic for stored, pasted and imported LaTeX: the input judgment
 * without quick escapes (those apply only to words typed in the editor).
 */
export function normalizeContextualUprightSymbols(
  source: string,
  options: Pick<InputJudgmentOptions, "deferLetterRuns" | "upright"> = {},
): string {
  return applyInputJudgment(source, options);
}

function readBracedCommand(source: string, start: number): number {
  const openingBrace = source.indexOf("{", start);
  if (openingBrace < 0) return start;
  let depth = 0;
  for (let index = openingBrace; index < source.length; index += 1) {
    if (source[index] === "{") depth += 1;
    if (source[index] === "}") {
      depth -= 1;
      if (depth === 0) return index + 1;
    }
  }
  return source.length;
}

export function normalizeChineseLatex(source: string): string {
  const normalizedTextCommands = normalizeContextualUprightSymbols(
    normalizeCanonicalUprightCommands(
      normalizeExtendedIntegralLatexCommands(source),
    ),
  ).replace(
    /\\(?:mathrm|textrm)\{([\u3400-\u9fff\uf900-\ufaff，。；：！？、（）【】《》“”‘’\s]+)\}/g,
    "\\text{$1}",
  );

  let result = "";
  let index = 0;

  while (index < normalizedTextCommands.length) {
    if (normalizedTextCommands.startsWith("\\text{", index)) {
      const end = readBracedCommand(normalizedTextCommands, index);
      result += normalizedTextCommands.slice(index, end);
      index = end;
      continue;
    }

    if (chineseChar.test(normalizedTextCommands[index])) {
      let end = index + 1;
      while (
        end < normalizedTextCommands.length &&
        (chineseChar.test(normalizedTextCommands[end]) ||
          (normalizedTextCommands[end] === " " &&
            end + 1 < normalizedTextCommands.length &&
            chineseChar.test(normalizedTextCommands[end + 1])))
      ) {
        end += 1;
      }
      result += "\\text{" + normalizedTextCommands.slice(index, end) + "}";
      index = end;
      continue;
    }

    result += normalizedTextCommands[index];
    index += 1;
  }

  return result;
}

export function normalizeMultilineLatex(source: string): string {
  return source
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .map(normalizeChineseLatex)
    .join("\n");
}

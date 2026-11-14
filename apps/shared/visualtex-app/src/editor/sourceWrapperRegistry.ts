import {
  compatibilitySourceWrapperKinds,
  compatibilityWrapperCanonicalTargets,
} from "../autocomplete/compatibilityCommands";

export type VisualTexSourceWrapperKind = "variant" | "accent";

/**
 * Native MathLive/LaTeX variant commands that use VisualTeX's source-first
 * editor shell. Compatibility aliases are merged below from the command
 * registry so the kernel never needs a second handwritten alias list.
 */
export const primaryVariantSourceWrapperCommands = Object.freeze([
  "\\mathbf",
  "\\mathbb",
  "\\mathit",
  "\\mathrm",
  "\\mathsf",
  "\\mathtt",
  "\\mathcal",
  "\\mathscr",
  "\\mathfrak",
  "\\boldsymbol",
  "\\bm",
  "\\mathbfit",
  "\\mathnormal",
] as const);

/**
 * Accent-like commands intentionally edited as raw source first. This is also
 * the single source of truth for MathEditor accent wrapper templates.
 */
export const accentSourceWrapperCommands = Object.freeze([
  "\\acute",
  "\\grave",
  "\\dot",
  "\\ddot",
  "\\dddot",
  "\\ddddot",
  "\\tilde",
  "\\bar",
  "\\breve",
  "\\check",
  "\\hat",
  "\\vec",
  "\\widehat",
  "\\widetilde",
  "\\overline",
  "\\overrightarrow",
  "\\overleftarrow",
  "\\overleftrightarrow",
  "\\mathring",
  "\\c",
] as const);

export const sourceWrapperCanonicalTargets = new Map<string, string>(
  compatibilityWrapperCanonicalTargets,
);

export const sourceWrapperKinds = new Map<string, VisualTexSourceWrapperKind>([
  ...primaryVariantSourceWrapperCommands.map(
    (command) => [command, "variant"] as const,
  ),
  ...accentSourceWrapperCommands.map(
    (command) => [command, "accent"] as const,
  ),
  ...compatibilitySourceWrapperKinds,
]);

export function canonicalSourceWrapperCommand(command: string) {
  return sourceWrapperCanonicalTargets.get(command) ?? command;
}

export function sourceWrapperKind(
  command: string,
): VisualTexSourceWrapperKind | null {
  const canonical = canonicalSourceWrapperCommand(command);
  return sourceWrapperKinds.get(command) ?? sourceWrapperKinds.get(canonical) ?? null;
}

export const accentSourceWrapperTemplates = new Map<string, string>(
  accentSourceWrapperCommands.map((command) => [
    command,
    `${command}{\\placeholder{}}`,
  ]),
);

export * from './public/mathlive';

import {
  LATEX_COMMANDS,
  MATH_SYMBOLS,
  getMacros,
} from './latex-commands/definitions-utils';

/**
 * Current version: `0.109.2`
 *
 * The version string of the SDK using the [semver](https://semver.org/) convention:
 *
 * `MAJOR`.`MINOR`.`PATCH`
 *
 * * **`MAJOR`** is incremented for incompatible API changes
 * * **`MINOR`** is incremented for new features
 * * **`PATCH`** is incremented for bug fixes
 */

export const version = {
  mathlive: '0.109.2',
};

export { setVisualTexLayoutEngine, setVisualTexEditingCore } from './editor-mathfield/visualtex-hooks';

/** Every command the parser knows, with its number of required and optional arguments (for the LaTeX source editor's completion). */
export function visualTexLatexCommandSignatures(): [string, number, number][] {
  const hidden = new Set(['\\emph', '\\error', '\\the']);
  const signatures: [string, number, number][] = [];
  for (const [name, definition] of Object.entries(LATEX_COMMANDS)) {
    if (definition.infix || hidden.has(name)) continue;
    const params = definition.params || [];
    signatures.push([name, params.filter((p) => !p.isOptional).length, params.filter((p) => p.isOptional).length]);
  }
  for (const name of Object.keys(MATH_SYMBOLS)) signatures.push([name, 0, 0]);
  for (const [name, definition] of Object.entries(getMacros())) signatures.push(['\\' + name, definition?.args || 0, 0]);
  return signatures;
}

/** @internal */
// export const debug = {
//   FUNCTIONS: MathLiveDebug.FUNCTIONS,
//   MATH_SYMBOLS: MathLiveDebug.MATH_SYMBOLS,
//   ENVIRONMENTS: MathLiveDebug.ENVIRONMENTS,
//   DEFAULT_KEYBINDINGS: MathLiveDebug.DEFAULT_KEYBINDINGS,
//   getKeybindingMarkup: MathLiveDebug.getKeybindingMarkup,
// };

/**
 * Server-side rendering exports.
 *
 * These functions do not require a DOM environment and can
 * be used from a server-side environment.
 *
 */

import '../core/math-environment';

import { Atom } from '../core/atom-class';

import '../latex-commands/definitions';

import { visualTexLayoutEngine } from '../editor-mathfield/visualtex-hooks';
import { Context } from '../core/context';
import { parseLatex } from '../core/parser';
import { validateLatex as validateLatexInternal } from '../core/parser';

import type { LatexSyntaxError, ParseMode } from './core-types';

import '../core/modes';
import { getDefaultContext } from '../core/context-utils';
import { LayoutOptions } from './options';
import { ContextInterface } from '../core/types';
import {
  getMacroDefinition,
  normalizeMacroDictionary,
} from '../latex-commands/definitions';

/**
 * Convert a LaTeX string to a string of HTML markup.
 *
 * :::info[Note]
 *
 * This function does not interact with the DOM. It does not load fonts or
 * inject stylesheets in the document. It can safely be used on the server side.
 * :::
 *
 * To get the output of this function to correctly display
 * in a document, use the mathlive static style sheet by adding the following
 * to the `<head>` of the document:
 *
 * ```html
 * <link
 *  rel="stylesheet"
 *  href="https://cdn.jsdelivr.net/npm/mathlive/mathlive-static.css"
 * />
 * ```
 *
 * or
 *
 * ```html
 * <link
 *  rel="stylesheet"
 *  href="https://unpkg.com/mathlive/mathlive-static.css"
 * />
 * ```
 *
 *
 *
 * @param text A string of valid LaTeX. It does not have to start
 * with a mode token such as `$$` or `\(`.
 *
 * @param options.defaultMode If `"displaystyle"` the "display" mode of TeX
 * is used to typeset the formula, which is most appropriate for formulas that are
 * displayed in a standalone block.
 *
 * If `"textstyle"` is used, the "text" mode of TeX is used, which is most
 * appropriate when displaying math "inline" with other text (on the same line).
 *
 * @category Conversion
 * @keywords convert, latex, markup
 */
export function convertLatexToMarkup(
  text: string,
  options?: Partial<LayoutOptions>
): string {
  const { atoms, display } = parseForLayout(text, options);
  //
  // 2. The layout engine lays out and draws the formula
  //
  return visualTexLayoutEngine.renderAtoms(atoms, { display });
}

/**
 * The formula as SVG, laid out exactly as the editor lays it out (exported,
 * copied and Office images). `svgOptions` go to the layout engine's
 * renderSvg (pixels per em, glyph outlines, colour, custom artwork); the
 * result is its drawing and extent.
 *
 * @category Conversion
 */
export function convertLatexToSvg(
  text: string,
  options: Partial<LayoutOptions> | undefined,
  svgOptions: Record<string, unknown>
): any {
  const { atoms, display } = parseForLayout(text, options);
  return visualTexLayoutEngine.renderSvg(atoms, { ...svgOptions, display });
}

function parseForLayout(
  text: string,
  options?: Partial<LayoutOptions>
): { atoms: readonly Atom[]; display: boolean } {
  const from: ContextInterface = { ...getDefaultContext() };
  if (options?.letterShapeStyle && options?.letterShapeStyle !== 'auto')
    from.letterShapeStyle = options.letterShapeStyle;

  if (options?.macros) {
    const macros = normalizeMacroDictionary(options?.macros);
    from.getMacro = (token) => getMacroDefinition(token, macros);
  }
  if (options?.registers)
    from.registers = { ...from.registers, ...options.registers };

  const defaultMode = options?.defaultMode ?? 'math';
  let parseMode: ParseMode = 'math';
  let mathstyle: 'displaystyle' | 'textstyle' = 'displaystyle';
  if (defaultMode === 'inline-math') mathstyle = 'textstyle';
  else if (defaultMode === 'math') mathstyle = 'displaystyle';
  else if (defaultMode === 'text') {
    mathstyle = 'textstyle';
    parseMode = 'text';
  }
  const effectiveContext = new Context({ from });

  //
  // 1. Parse the formula and return a tree of atoms, e.g. 'genfrac'.
  //
  const root = new Atom({
    type: 'root',
    mode: parseMode,
    body: parseLatex(text, { context: effectiveContext, parseMode, mathstyle }),
  });
  return { atoms: root.body ?? [], display: mathstyle === 'displaystyle' };
}

/**
 * Check if a string of LaTeX is valid and return an array of syntax errors.
 *
 * @category Conversion
 */
export function validateLatex(s: string): LatexSyntaxError[] {
  return validateLatexInternal(s, { context: getDefaultContext() });
}


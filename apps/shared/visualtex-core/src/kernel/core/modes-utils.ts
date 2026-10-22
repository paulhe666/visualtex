import { Atom } from './atom-class';
import type { ParseMode, Style } from '../public/core-types';
import { getDefinition } from '../latex-commands/definitions-utils';
import { joinLatex, latexCommand } from './tokenizer';
import type { TokenDefinition } from '../latex-commands/types';
import type { ToLatexOptions } from './types';

export abstract class Mode {
  static _registry: Record<string, Mode> = {};
  constructor(name: string) {
    Mode._registry[name] = this;
  }

  static createAtom(
    mode: ParseMode,
    command: string,
    style?: Style
  ): Atom | null {
    return Mode._registry[mode].createAtom(
      command,
      getDefinition(command, mode),
      style
    );
  }

  static serialize(
    atoms: readonly Atom[] | undefined,
    options: ToLatexOptions
  ): string {
    if (!atoms || atoms.length === 0) return '';

    if (options.skipStyles ?? false) {
      const body: string[] = [];
      for (const run of getModeRuns(atoms)) {
        const mode = Mode._registry[run[0].mode];
        body.push(...mode.serialize(run, options));
      }
      return joinLatex(body);
    }

    return joinLatex(emitFontSizeRun(atoms, options));
  }

  abstract createAtom(
    command: string,
    info: TokenDefinition | null,
    style?: Style
  ): Atom | null;

  abstract serialize(run: readonly Atom[], options: ToLatexOptions): string[];
}

/*
 * Return an array of runs with the same mode
 */
export function getModeRuns(atoms: readonly Atom[]): (readonly Atom[])[] {
  const result: Atom[][] = [];
  let run: Atom[] = [];
  let currentMode = 'NONE';
  for (const atom of atoms) {
    if (atom.type !== 'first') {
      if (atom.mode !== currentMode) {
        if (run.length > 0) result.push(run);
        run = [atom];
        currentMode = atom.mode;
      } else run.push(atom);
    }
  }
  // Push whatever is left
  if (run.length > 0) result.push(run);
  return result;
}

/** This "weight" is only for math mode. Text mode uses fontSeries. */
export function weightString(atom: Atom): string {
  if (atom?.mode !== 'math') return '';
  const { style } = atom;
  if (!style) return '';
  if (!style.variantStyle) return '';
  if (style.variantStyle === 'bold' || style.variantStyle === 'bolditalic')
    return 'bold';

  return '';
}

/** Combine variant and variantStyle in one string */
export function variantString(atom: Atom): string {
  if (!atom) return '';
  const { style } = atom;
  if (!style) return '';

  let result = style.variant;
  if (result === undefined) return 'normal';

  if (
    ![
      'calligraphic',
      'fraktur',
      'double-struck',
      'script',
      'monospace',
      'sans-serif',
    ].includes(result) &&
    style.variantStyle
  )
    result += '-' + style.variantStyle;

  return result;
}

/*
 * Return an array of runs (array of atoms with the same value
 *   for the specified property)
 */
export function getPropertyRuns(
  atoms: readonly Atom[],
  property: keyof Style | 'bold' | 'visualTexBoldScope'
): (readonly Atom[])[] {
  const result: Atom[][] = [];
  let run: Atom[] = [];
  let currentValue: string | number | undefined = undefined;
  for (const atom of atoms) {
    if (atom.type === 'first') continue;
    // The 'variant' property combines the variant and variantStyle
    let value;
    if (property === 'variant') value = variantString(atom);
    else if (property === 'bold') value = weightString(atom);
    else if (property === 'visualTexBoldScope')
      value = weightString(atom) + ':' + (atom.style.variant === 'normal' && atom.style.variantStyle === 'bold' ? 'basic-font' : 'other');
    else value = atom.style[property];

    if (value === currentValue) {
      // Same value, add it to the current run
      run.push(atom);
    } else {
      // The value of property for this atom is different from the
      // current value, start a new run
      if (run.length > 0) result.push(run);
      run = [atom];
      currentValue = value;
    }
  }

  // Push whatever is left
  if (run.length > 0) result.push(run);

  return result;
}

function emitColorRun(run: readonly Atom[], options: ToLatexOptions): string[] {
  const { parent } = run[0];
  const parentColor = parent?.style.color;

  const result: string[] = [];
  // Since `\textcolor{}` applies to both text and math mode, wrap mode first, then
  // textcolor
  for (const modeRun of getModeRuns(run)) {
    const mode = options.defaultMode;

    for (const colorRun of getPropertyRuns(modeRun, 'color')) {
      const style = colorRun[0].style;
      const body = Mode._registry[colorRun[0].mode].serialize(colorRun, {
        ...options,
        defaultMode: mode === 'text' ? 'text' : 'math',
      });
      if (
        !options.skipStyles &&
        style.color &&
        style.color !== 'none' &&
        (!parent || parentColor !== style.color)
      ) {
        result.push(
          latexCommand(
            '\\textcolor',
            style.verbatimColor ?? style.color,
            joinLatex(body)
          )
        );
      } else result.push(joinLatex(body));
    }
  }

  return result;
}

function emitBackgroundColorRun(
  run: readonly Atom[],
  options: ToLatexOptions
): string[] {
  const { parent } = run[0];
  const parentColor = parent?.style.backgroundColor;
  return getPropertyRuns(run, 'backgroundColor').map((x) => {
    if (x.length > 0 || x[0].type !== 'box') {
      const style = x[0].style;
      if (
        style.backgroundColor &&
        style.backgroundColor !== 'none' &&
        (!parent || parentColor !== style.backgroundColor)
      ) {
        return latexCommand(
          '\\colorbox',
          style.verbatimBackgroundColor ?? style.backgroundColor,
          joinLatex(emitColorRun(x, { ...options, defaultMode: 'text' }))
        );
      }
    }
    return joinLatex(emitColorRun(x, options));
  });
}

function emitFontSizeRun(
  run: readonly Atom[],
  options: ToLatexOptions
): string[] {
  if (run.length === 0) return [];
  const { parent } = run[0];
  const contextFontsize = parent?.style.fontSize;
  const result: string[] = [];
  for (const sizeRun of getPropertyRuns(run, 'fontSize')) {
    const fontsize = sizeRun[0].style.fontSize;
    const body = emitBackgroundColorRun(sizeRun, options);
    if (body) {
      if (
        fontsize &&
        fontsize !== 'auto' &&
        (!parent || contextFontsize !== fontsize)
      ) {
        result.push(
          [
            '',
            '\\tiny',
            '\\scriptsize',
            '\\footnotesize',
            '\\small',
            '\\normalsize',
            '\\large',
            '\\Large',
            '\\LARGE',
            '\\huge',
            '\\Huge',
          ][fontsize],
          ...body
        );
      } else result.push(...body);
    }
  }

  return result;
}

import { Atom } from './atom';
import { joinLatex, latexCommand } from './tokenizer';
import {
  getPropertyRuns,
  Mode,
  variantString,
  weightString,
} from './modes-utils';
import type { Style } from '../public/core-types';
import type { TokenDefinition } from '../latex-commands/types';
import type { ToLatexOptions } from './types';
import { getDefinition } from '../latex-commands/definitions';

const GREEK_LOWERCASE = /^[\u03B1-\u03C9]|\u03D1|\u03D5|\u03D6|\u03F1|\u03F5]$/;
const GREEK_UPPERCASE =
  /^[\u0393|\u0394\u0398\u039B\u039E\u03A0\u03A3\u03A5\u03A6\u03A8\u03A9]$/;

export const LETTER_SHAPE_RANGES = [
  /^[a-z]$/, // Lowercase latin
  /^[A-Z]$/, // Uppercase latin
  GREEK_LOWERCASE,
  GREEK_UPPERCASE,
];

// The letterShapeStyle property indicates which characters should be
// automatically italicized (see LETTER_SHAPE_RANGES)
export const LETTER_SHAPE_MODIFIER = {
  iso: ['it', 'it', 'it', 'it'],
  tex: ['it', 'it', 'it', 'up'],
  french: ['it', 'up', 'up', 'up'],
  upright: ['up', 'up', 'up', 'up'],
};

// See http://ctan.math.illinois.edu/macros/latex/base/fntguide.pdf

export class MathMode extends Mode {
  constructor() {
    super('math');
  }

  createAtom(command: string, info: TokenDefinition, style?: Style): Atom {
    if (info === null) {
      return new Atom({
        type: 'mord',
        mode: 'math',
        command,
        value: command,
        style,
      });
    }

    let isFunction: boolean;
    try {
      isFunction =
        globalThis.MathfieldElement?.isFunction(info.command ?? command) ??
        false;
    } catch (e) {
      isFunction = false;
    }

    if (info.definitionType === 'symbol') {
      const result = new Atom({
        type: info.type ?? 'mord',
        mode: 'math',
        command: info.command ?? command,
        value: String.fromCodePoint(info.codepoint),
        isFunction,
        style,
      });

      if (command.startsWith('\\')) result.verbatimLatex = command;
      return result;
    }
    const result = new Atom({
      type: 'mord',
      mode: 'math',
      command: info.command ?? command,
      value: command,
      isFunction,
      style,
    });
    if (command.startsWith('\\')) result.verbatimLatex = command;

    return result;
  }

  serialize(run: Atom[], options: ToLatexOptions): string[] {
    const result = emitBoldRun(run, { ...options, defaultMode: 'math' });
    if (result.length === 0 || options.defaultMode !== 'text') return result;
    return ['$ ', ...result, ' $'];
  }
}

function emitBoldRun(run: Atom[], options: ToLatexOptions): string[] {
  // VisualTeX: runs split by weight and by whether the bold is the basic
  // font's (\mathbf: an upright bold variant) or symbol-wide (\bm)
  return getPropertyRuns(run, 'visualTexBoldScope').map((x) => {
    const weight = weightString(x[0]);
    const visualTexVariant = variantString(x[0]);
    // bolditalic is already serialized by emitVariantRun as \mathbfit.
    // Wrapping it in an additional \mathbf is redundant and changes
    // round-trip style semantics.
    if (weight !== 'bold' || visualTexVariant.endsWith('-bolditalic'))
      return joinLatex(emitVariantRun(x, options));

    // If the parent is already bold the same way, don't emit the bold command
    if (weightString(x[0].parent!) === 'bold' &&
        (x[0].style.variant === 'normal' && x[0].style.variantStyle === 'bold') ===
        (x[0].parent!.style.variant === 'normal' && x[0].parent!.style.variantStyle === 'bold'))
      return joinLatex(emitVariantRun(x, options));

    if (x[0].style.variant === 'normal' && x[0].style.variantStyle === 'bold')
      return latexCommand('\\mathbf', joinLatex(emitVariantRun(x, options)));

    // Explicit basic font scope was handled above; preserve symbol-wide bold.
    return latexCommand('\\bm', joinLatex(emitVariantRun(x, options)));
  });
}

function emitVariantRun(
  run: readonly Atom[],
  options: ToLatexOptions
): string[] {
  const { parent } = run[0];
  const contextVariant = variantString(parent!);
  return getPropertyRuns(run, 'variant').map((x) => {
    const variant = variantString(x[0]);
    let command = '';
    if (variant && variant !== contextVariant) {
      // Note that bold is handled separately, so we ignore it here.
      command = {
        'calligraphic': '\\mathcal',
        'calligraphic-uo': '\\mathcal',
        'fraktur': '\\mathfrak',
        'fraktur-uo': '\\mathfrak',
        'double-struck': '\\mathbb',
        'double-struck-uo': '\\mathbb',
        'script': '\\mathscr',
        'script-uo': '\\mathscr',
        'monospace': '\\mathtt',
        'monospace-uo': '\\mathtt',
        'sans-serif': '\\mathsf',
        'sans-serif-uo': '\\mathsf',
        'normal': '',
        'normal-up': '\\mathrm',
        'normal-italic': '\\mathnormal',
        'normal-bold': '',
        'normal-bolditalic': '\\mathbfit',
        'ams': '',
        'ams-up': '\\mathrm',
        'ams-italic': '\\mathit',
        'ams-bold': '',
        'ams-bolditalic': '\\mathbfit',
        'main': '',
        'main-up': '\\mathrm',
        'main-italic': '\\mathit',
        'main-bold': '',
        'main-bolditalic': '\\mathbfit',
        // There are a few rare font families possible, which
        // are not supported:
        // mathbbm, mathbbmss, mathbbmtt, mathds, swab, goth
        // In addition, the 'main' and 'math' font technically
        // map to \mathnormal{}
      }[variant!]!;
      console.assert(command !== undefined);
    }

    // When we have a variant command, superscripts/subscripts should be
    // applied AFTER the variant command, not inside it.
    // For example, \mathbb{R}^0 not \mathbb{R^0}
    // See issue #2867
    let trailingSupsub = '';
    const lastAtom = x[x.length - 1];
    const hasTrailingSupsub =
      command &&
      x.length > 0 &&
      (lastAtom.branch('subscript') !== undefined ||
        lastAtom.branch('superscript') !== undefined);

    if (hasTrailingSupsub) trailingSupsub = lastAtom.supsubToLatex(options);

    const arg = joinLatex(
      x.map((atom, index) => {
        // For the last atom with a superscript/subscript in a variant command,
        // serialize without the supsub (we'll add it after the command)
        if (hasTrailingSupsub && index === x.length - 1) {
          // Get the definition to check if it has a custom serializer
          const def = getDefinition(atom.command, atom.mode);
          if (def?.serialize) return def.serialize(atom, options);

          // Serialize the atom without superscript/subscript
          if (atom.body && atom.command)
            return latexCommand(atom.command, atom.bodyToLatex(options));

          if (atom.body) return atom.bodyToLatex(options);

          if (!atom.value || atom.value === '\u200B') return '';
          return atom.command || atom.value;
        }
        return atom._serialize(options);
      })
    );

    if (!command) return arg;
    return latexCommand(command, arg) + trailingSupsub;
  });
}

// Singleton class
new MathMode();

import type { MathfieldOptions } from '../public/options';
import type { InsertStyleHook } from '../public/mathfield';

import { isArray } from '../common/types';

import { l10n } from '../core/l10n';
import { defaultBackgroundColorMap, defaultColorMap } from '../core/color';

import { normalizeMacroDictionary } from '../latex-commands/definitions-utils';

import { defaultExportHook } from './mode-editor';

import { INLINE_SHORTCUTS } from '../editor/shortcuts-definitions';
import { DEFAULT_KEYBINDINGS } from '../editor/keybindings-definitions';
import { defaultInsertStyleHook } from './styling';

/** @internal */
export type _MathfieldOptions = MathfieldOptions & {
  value: string;
};

/** An options object read by name */
type AnyRecord = Record<string, any>;

export function update(
  updates: Partial<_MathfieldOptions>
): Partial<_MathfieldOptions> {
  // Options are copied by name
  const result: Partial<_MathfieldOptions> & Record<string, any> = {};
  for (const key of Object.keys(updates)) {
    switch (key) {
      case 'scriptDepth':
        const scriptDepth = updates.scriptDepth;
        if (isArray<number>(scriptDepth))
          result.scriptDepth = [scriptDepth[0], scriptDepth[1]];
        else if (typeof scriptDepth === 'number')
          result.scriptDepth = [scriptDepth, scriptDepth];
        else if (typeof scriptDepth === 'string') {
          const [from, to] = (scriptDepth as string)
            .split(',')
            .map((x) => parseInt(x.trim()));
          result.scriptDepth = [from, to];
        } else throw new TypeError('Unexpected value for scriptDepth');

        break;

      case 'letterShapeStyle':
        if (updates.letterShapeStyle === 'auto') {
          // Letter shape style (locale dependent)
          if (l10n.locale.startsWith('fr')) result.letterShapeStyle = 'french';
          else result.letterShapeStyle = 'tex';
        } else result.letterShapeStyle = updates.letterShapeStyle!;

        break;

      case 'defaultMode':
        if (
          !['text', 'math', 'inline-math'].includes(
            updates.defaultMode as string
          )
        ) {
          console.error(
            `MathLive 0.109.2:  valid values for defaultMode are "text", "math" or "inline-math"`
          );
          result.defaultMode = 'math';
        } else result.defaultMode = updates.defaultMode!;
        break;

      case 'macros':
        result.macros = normalizeMacroDictionary(updates.macros!);
        break;

      default:
        if (isArray((updates as AnyRecord)[key]))
          result[key] = [...(updates as AnyRecord)[key]];
        else if (
          typeof (updates as AnyRecord)[key] === 'object' &&
          !((updates as AnyRecord)[key] instanceof Element)
        )
          result[key] = { ...(updates as AnyRecord)[key] };
        else result[key] = (updates as AnyRecord)[key];
    }
  }

  return result;
}

export function get(
  config: Required<_MathfieldOptions>,
  keys?: keyof _MathfieldOptions | string[]
): any | Partial<_MathfieldOptions> {
  let resolvedKeys: string[];
  if (typeof keys === 'string') resolvedKeys = [keys];
  else if (keys === undefined) resolvedKeys = Object.keys(config);
  else resolvedKeys = keys;

  const result: Partial<_MathfieldOptions> & Record<string, any> = {};
  for (const x of resolvedKeys) {
    if ((config as AnyRecord)[x] === null) result[x] = null;
    else if (isArray((config as AnyRecord)[x]))
      result[x] = [...(config as AnyRecord)[x]];
    else if (
      typeof (config as AnyRecord)[x] === 'object' &&
      !((config as AnyRecord)[x] instanceof Element)
    ) {
      // Some object literal, make a copy (for keypressSound, macros, etc...)
      result[x] = { ...(config as AnyRecord)[x] };
    } else result[x] = (config as AnyRecord)[x];
  }
  // If requested a single key, return its value
  if (typeof keys === 'string') return result[keys];

  return result;
}

export function getDefault(): Required<_MathfieldOptions> {
  return {
    readOnly: false,

    defaultMode: 'math',
    macros: {},
    registers: {},
    colorMap: defaultColorMap,
    backgroundColorMap: defaultBackgroundColorMap,
    letterShapeStyle: l10n.locale.startsWith('fr') ? 'french' : 'tex',
    minFontScale: 0,
    maxMatrixCols: 10,

    smartMode: false,
    smartFence: true,
    smartSuperscript: true,
    scriptDepth: [Infinity, Infinity],
    removeExtraneousParentheses: true,
    isImplicitFunction: (x) =>
      [
        '\\sin',
        '\\cos',
        '\\tan',

        '\\arcsin',
        '\\arccos',
        '\\arctan',
        '\\arcsec',
        '\\arccsc',

        '\\arsinh',
        '\\arcosh',
        '\\artanh',
        '\\arcsech',
        '\\arccsch',
        '\\arg',
        '\\ch',
        '\\cosec',
        '\\cosh',
        '\\cot',
        '\\cotg',
        '\\coth',
        '\\csc',
        '\\ctg',
        '\\cth',
        '\\sec',
        '\\sinh',
        '\\sh',
        '\\tanh',
        '\\tg',
        '\\th',

        '\\lg',
        '\\lb',
        '\\log',
        '\\ln',
      ].includes(x),

    mathModeSpace: '',
    placeholderSymbol: '▢',
    contentPlaceholder: '',
    popoverPolicy: 'auto',

    keybindings: DEFAULT_KEYBINDINGS,

    inlineShortcuts: INLINE_SHORTCUTS,
    inlineShortcutTimeout: 0,

    onInsertStyle: defaultInsertStyleHook as InsertStyleHook,
    onInlineShortcut: () => '',
    onScrollIntoView: null,
    onExport: defaultExportHook,
    value: '',
  };
}

export function effectiveMode(options: MathfieldOptions): 'math' | 'text' {
  if (options.defaultMode === 'inline-math') return 'math';
  return options.defaultMode;
}

import type { Style } from '../public/core-types';

import { Atom } from '../core/atom-class';
import { latexCommand } from '../core/tokenizer';
import { getDefinition } from '../latex-commands/definitions-utils';
import type { AtomJson, ToLatexOptions } from '../core/types';


export type EncloseAtomOptions = {
  shadow?: string;
  strokeWidth?: string;
  strokeStyle?: string;
  svgStrokeStyle?: string;
  strokeColor?: string;
  borderStyle?: string;
  padding?: string;
  backgroundcolor?: string;
  style: Style;
};
export type Notations = {
  downdiagonalstrike?: boolean;
  updiagonalstrike?: boolean;
  verticalstrike?: boolean;
  horizontalstrike?: boolean;
  updiagonalarrow?: boolean;
  right?: boolean;
  bottom?: boolean;
  left?: boolean;
  top?: boolean;
  circle?: boolean;
  roundedbox?: boolean;
  madruwb?: boolean;
  actuarial?: boolean;
  box?: boolean;
  phasorangle?: boolean;
  // radical?: boolean;
  longdiv?: boolean;
};

export class EncloseAtom extends Atom {
  backgroundcolor?: string;
  border!: string;

  readonly notation: Notations;
  private readonly shadow?: string;
  private readonly strokeWidth?: string;
  private readonly strokeStyle?: string;
  private readonly svgStrokeStyle?: string;
  private readonly strokeColor?: string;
  private readonly borderStyle?: string;
  private readonly padding?: string;

  constructor(
    command: string,
    body: readonly Atom[],
    notation: Notations,
    options: EncloseAtomOptions
  ) {
    super({ type: 'enclose', command, style: options.style });
    this.body = body;
    this.backgroundcolor = options.backgroundcolor;
    if (notation.updiagonalarrow) notation.updiagonalstrike = false;

    if (notation.box) {
      notation.left = false;
      notation.right = false;
      notation.bottom = false;
      notation.top = false;
    }

    this.notation = notation;
    this.shadow = options.shadow ?? 'none';
    this.strokeWidth = options.strokeWidth ?? '0.06em';
    if (!this.strokeWidth) this.strokeWidth = '0.06em';
    this.strokeStyle = options.strokeStyle;
    this.svgStrokeStyle = options.svgStrokeStyle;
    this.strokeColor = options.strokeColor;
    this.borderStyle = options.borderStyle;
    this.padding = options.padding;

    this.captureSelection = false;
  }

  static fromJson(json: AtomJson): EncloseAtom {
    return new EncloseAtom(
      json.command,
      json.body,
      json.notation,
      json as any as EncloseAtomOptions
    );
  }

  toJson(): AtomJson {
    return {
      ...super.toJson(),
      notation: this.notation,
      shadow: this.shadow,
      strokeWidth: this.strokeWidth,
      strokeStyle: this.strokeStyle,
      svgStrokeStyle: this.svgStrokeStyle,
      strokeColor: this.strokeColor,
      borderStyle: this.borderStyle,
      padding: this.padding,
    };
  }

  _serialize(options: ToLatexOptions): string {
    if (
      !(
        options.expandMacro ||
        options.skipStyles ||
        options.skipPlaceholders
      ) &&
      typeof this.verbatimLatex === 'string'
    )
      return this.verbatimLatex;
    const def = getDefinition(this.command, this.mode);
    if (def?.serialize) return def.serialize(this, options);

    let command = this.command ?? '';
    if (this.command === '\\enclose') {
      command += '{' + Object.keys(this.notation).join(' ') + '}';

      // \enclose can have optional parameters...
      let style = '';
      let sep = '';
      if (this.backgroundcolor && this.backgroundcolor !== 'transparent') {
        style += sep + 'mathbackground="' + this.backgroundcolor + '"';
        sep = ',';
      }

      if (this.shadow && this.shadow !== 'auto') {
        style += sep + 'shadow="' + this.shadow + '"';
        sep = ',';
      }

      if (this.strokeWidth || this.strokeStyle !== 'solid') {
        style += sep + this.borderStyle;
        sep = ',';
      } else if (this.strokeColor && this.strokeColor !== 'currentColor') {
        style += sep + 'mathcolor="' + this.strokeColor + '"';
        sep = ',';
      }

      if (style) command += `[${style}]`;
    }

    return latexCommand(command, this.bodyToLatex(options));
  }


}





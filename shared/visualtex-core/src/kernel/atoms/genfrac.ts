import type { MathstyleName, Style } from '../public/core-types';

import { Atom } from '../core/atom-class';
import type { AtomJson } from '../core/types';

export type GenfracOptions = {
  continuousFraction?: boolean;
  align?: 'left' | 'right' | 'center';
  numerPrefix?: string;
  denomPrefix?: string;
  leftDelim?: string;
  rightDelim?: string;
  hasBarLine?: boolean;
  mathstyleName?: MathstyleName;
  style?: Style;
};

/**
 * Genfrac -- Generalized Fraction
 *
 * Decompose fractions, binomials, and in general anything made
 * of a numerator and denominator, optionally separated by a fraction bar,
 * and optionally surrounded by delimiters (parentheses, brackets, etc...).
 *
 * Depending on the type of fraction the mathstyle is either
 * displaystyle or textstyle. This value can also be set to 'auto',
 * to indicate it should use the current mathstyle
 */
export class GenfracAtom extends Atom {
  readonly hasBarLine: boolean;
  readonly leftDelim?: string;
  readonly rightDelim?: string;
  private readonly continuousFraction: boolean;
  private readonly align: 'left' | 'right' | 'center';
  private readonly numerPrefix?: string;
  private readonly denomPrefix?: string;
  private readonly mathstyleName?: MathstyleName;

  constructor(
    above: readonly Atom[],
    below: readonly Atom[],
    options: GenfracOptions
  ) {
    super({
      ...options,
      type: 'genfrac',
      displayContainsHighlight: true,
    });
    this.above = above;
    this.below = below;
    this.hasBarLine = options?.hasBarLine ?? true;
    this.continuousFraction = options?.continuousFraction ?? false;
    this.align = options?.align ?? 'center';
    this.numerPrefix = options?.numerPrefix;
    this.denomPrefix = options?.denomPrefix;
    this.mathstyleName = options?.mathstyleName;
    this.leftDelim = options?.leftDelim;
    this.rightDelim = options?.rightDelim;
  }

  static fromJson(json: AtomJson): GenfracAtom {
    return new GenfracAtom(
      json.above,
      json.below,
      json as any as GenfracOptions
    );
  }

  toJson(): AtomJson {
    const options: GenfracOptions = {};
    if (this.continuousFraction) options.continuousFraction = true;
    if (this.align !== 'center') options.align = this.align;
    if (this.numerPrefix) options.numerPrefix = this.numerPrefix;
    if (this.denomPrefix) options.denomPrefix = this.denomPrefix;
    if (this.leftDelim) options.leftDelim = this.leftDelim;
    if (this.rightDelim) options.rightDelim = this.rightDelim;
    if (!this.hasBarLine) options.hasBarLine = false;
    if (this.mathstyleName) options.mathstyleName = this.mathstyleName;
    return { ...super.toJson(), ...options };
  }
}

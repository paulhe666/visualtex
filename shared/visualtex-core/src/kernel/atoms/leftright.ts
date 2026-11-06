import type { Style } from '../public/core-types';

import { Atom } from '../core/atom-class';
import { RIGHT_DELIM } from '../core/delimiters';
import { joinLatex } from '../core/tokenizer';
import type { AtomJson, ToLatexOptions } from '../core/types';

/**
 *  \left....\right
 *
 * Note that we can encounter malformed \left...\right, for example
 * a \left without a matching \right or vice versa. In that case, the
 * leftDelim (resp. rightDelim) will be undefined. We still need to handle
 * those cases.
 *
 */
export class LeftRightAtom extends Atom {
  leftDelim?: string;
  rightDelim?: string;

  // Indicate which command should the delimiters be serialized to:
  // regular delimiter, `\left...\right` or `\mleft...mright`.
  private readonly variant: '' | 'left...right' | 'mleft...mright';

  constructor(
    variant: '' | 'left...right' | 'mleft...mright',
    body: readonly Atom[],
    options: {
      leftDelim: string;
      rightDelim: string;
      style?: Style;
    }
  ) {
    super({
      type: 'leftright',
      style: options.style,
      displayContainsHighlight: true,
    });
    this.variant = variant;
    this.body = body;
    this.leftDelim = options.leftDelim;
    this.rightDelim = options.rightDelim;
  }

  static fromJson(json: AtomJson): LeftRightAtom {
    return new LeftRightAtom(json.variant ?? '', json.body, json as any);
  }

  toJson(): AtomJson {
    const result = super.toJson();
    if (this.variant) result.variant = this.variant;
    if (this.leftDelim) result.leftDelim = this.leftDelim;
    if (this.rightDelim) result.rightDelim = this.rightDelim;
    return result;
  }

  _serialize(options: ToLatexOptions): string {
    const rightDelim = this.matchingRightDelim();

    if (this.variant === 'left...right') {
      return joinLatex([
        '\\left',
        this.leftDelim ?? '.',
        this.bodyToLatex(options),
        '\\right',
        rightDelim,
      ]);
    }

    if (this.variant === 'mleft...mright') {
      return joinLatex([
        '\\mleft',
        this.leftDelim ?? '.',
        this.bodyToLatex(options),
        '\\mright',
        rightDelim,
      ]);
    }

    return joinLatex([
      !this.leftDelim || this.leftDelim === '.' ? '' : this.leftDelim,
      this.bodyToLatex(options),
      rightDelim,
    ]);
  }

  matchingRightDelim(): string {
    if (this.rightDelim && this.rightDelim !== '?') return this.rightDelim;
    const leftDelim = this.leftDelim ?? '.';
    return RIGHT_DELIM[leftDelim] ?? leftDelim;
  }

}


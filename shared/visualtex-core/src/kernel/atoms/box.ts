import type { LatexValue } from '../public/core-types';

import { Atom } from '../core/atom-class';
import { joinLatex } from '../core/tokenizer';
import type {
  CreateAtomOptions,
  AtomJson,
  ToLatexOptions,
} from '../core/types';

/**
 * A BoxAtom is an atom that renders a box around its content.
 * Not to be confused with the Box class, which represents the geometric
 * rendering of an atom.
 */
export class BoxAtom extends Atom {
  readonly framecolor?: LatexValue;
  readonly backgroundcolor?: LatexValue;
  readonly padding?: LatexValue;
  readonly offset?: LatexValue; // Vertical offset (used by \raisebox)
  readonly border?: string;

  constructor(
    options: CreateAtomOptions & {
      body: readonly Atom[];
      framecolor?: LatexValue;
      backgroundcolor?: LatexValue;
      padding?: LatexValue;
      offset?: LatexValue;
      border?: string;
    }
  ) {
    super({
      mode: options.mode,
      command: options.command,
      style: options.style,
      body: options.body,
      type: 'box',
    });

    this.framecolor = options.framecolor;
    this.backgroundcolor = options.backgroundcolor;
    this.padding = options.padding;
    this.offset = options.offset;
    this.border = options.border;
  }

  static fromJson(json: { [key: string]: any }): BoxAtom {
    return new BoxAtom(json as any);
  }

  toJson(): AtomJson {
    return {
      ...super.toJson(),
      framecolor: this.framecolor,
      backgroundcolor: this.backgroundcolor,
      padding: this.padding,
      offset: this.offset,
      border: this.border,
    };
  }


  _serialize(options: ToLatexOptions): string {
    if (!options.skipStyles) return super._serialize(options);

    return joinLatex([this.bodyToLatex(options), this.supsubToLatex(options)]);
  }
}

// function lineWidth(s: string): string {
//   const m = s.match(/[\d]+(\.[\d]+)?[a-zA-Z]+/);
//   if (m) return m[0];
//   return '';
// }

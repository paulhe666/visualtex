import { Atom } from '../core/atom-class';
import type { AtomJson, ToLatexOptions } from '../core/types';

/**
 * Atom for raw latex character, while in LaTeX editing mode
 */
export class LatexAtom extends Atom {
  isSuggestion: boolean; // Display suggestions with reduced opacity
  isError: boolean; // Display errors with wavy red line

  constructor(value: string, options?: { isSuggestion: boolean }) {
    super({ type: 'latex', value, mode: 'latex' });
    this.isSuggestion = options?.isSuggestion ?? false;
    this.isError = false;
  }

  static fromJson(json: AtomJson): LatexAtom {
    const result = new LatexAtom(json.value);
    if (json.isSuggestion) result.isSuggestion = true;
    if (json.isError) result.isError = true;
    return result;
  }

  toJson(): AtomJson {
    const options: { [key: string]: any } = {};
    if (this.isSuggestion) options.isSuggestion = true;
    if (this.isError) options.isError = true;
    return { type: 'latex', value: this.value, ...options };
  }

}

/**
 * A group that represents a raw LaTeX editing zone.
 * There is only one LatexGroupAtom at a time in an expression.
 * All the children of a LatexGroupAtom are LatexAtom.
 */
export class LatexGroupAtom extends Atom {
  /** VisualTeX: set while this group is a source wrapper's command (keyboard-input.ts) */
  visualTexSourceWrapper?: any;
  constructor(latex = '') {
    super({ type: 'latexgroup', mode: 'latex' });
    this.body = [...latex].map((c) => new LatexAtom(c));
    this.skipBoundary = true;
  }

  static fromJson(_json: AtomJson): LatexGroupAtom {
    return new LatexGroupAtom();
  }

  toJson(): AtomJson {
    return super.toJson();
  }


  _serialize(_options: ToLatexOptions): string {
    return this.body?.map((x) => x.value).join('') ?? '';
  }
}

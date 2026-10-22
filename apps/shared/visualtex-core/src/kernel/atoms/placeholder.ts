import type { ParseMode, Style } from '../public/core-types';

import { Atom } from '../core/atom-class';
import type { AtomJson, ToLatexOptions } from '../core/types';

export class PlaceholderAtom extends Atom {
  constructor(options?: { mode?: ParseMode; style?: Style }) {
    super({
      type: 'placeholder',
      command: '\\placeholder',
      mode: options?.mode ?? 'math',
      style: options?.style,
    });
    this.captureSelection = true;
  }

  static fromJson(json: AtomJson): PlaceholderAtom {
    return new PlaceholderAtom(json as any);
  }

  toJson(): AtomJson {
    return super.toJson();
  }


  _serialize(options: ToLatexOptions): string {
    if (options.skipPlaceholders) return '';
    return '\\placeholder{}';
  }
}

import type { ParseMode } from '../public/core-types';

import { Atom } from '../core/atom-class';
import type { AtomJson, ToLatexOptions } from '../core/types';

export class CompositionAtom extends Atom {
  constructor(value: string, options?: { mode: ParseMode }) {
    super({ type: 'composition', mode: options?.mode ?? 'math', value });
  }

  static fromJson(json: { [key: string]: any }): CompositionAtom {
    return new CompositionAtom(json.value, json as any);
  }

  toJson(): AtomJson {
    return super.toJson();
  }


  _serialize(_options: ToLatexOptions): string {
    return '';
  }
}

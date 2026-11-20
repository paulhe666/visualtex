import type { Style } from '../public/core-types';

import { Atom, NAMED_BRANCHES } from '../core/atom-class';
import type { ToLatexOptions } from '../core/types';

export class SubsupAtom extends Atom {
  constructor(options?: { style?: Style }) {
    super({ type: 'subsup', style: options?.style });
    this.subsupPlacement = 'auto';
  }


  static fromJson(json: { [key: string]: any }): SubsupAtom {
    const result = new SubsupAtom(json as any);
    for (const branch of NAMED_BRANCHES)
      if (json[branch]) result.setChildren(json[branch], branch);

    return result;
  }


  _serialize(options: ToLatexOptions): string {
    return this.supsubToLatex(options);
  }
}

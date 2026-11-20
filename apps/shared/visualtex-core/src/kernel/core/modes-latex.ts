import { Mode } from './modes-utils';
import { Atom } from './atom';

import { LatexAtom, LatexGroupAtom } from '../atoms/latex';
import type { ToLatexOptions } from './types';

export class LatexMode extends Mode {
  constructor() {
    super('latex');
  }

  createAtom(command: string): Atom | null {
    return new LatexAtom(command);
  }

  serialize(run: Atom[], _options: ToLatexOptions): string[] {
    // a VisualTeX source wrapper (a command group typed in the LaTeX box) is kept whole
    return run
      .filter((x) => (x instanceof LatexAtom && !x.isSuggestion) || (x instanceof LatexGroupAtom && x.visualTexSourceWrapper))
      .map((x) => (x instanceof LatexGroupAtom ? x._serialize(_options) : x.value));
  }
}

new LatexMode();

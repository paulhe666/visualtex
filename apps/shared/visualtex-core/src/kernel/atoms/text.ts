import type { Style } from '../public/core-types';

import { Atom } from '../core/atom-class';
import { charToLatex } from '../latex-commands/definitions-utils';
import type { AtomJson, ToLatexOptions } from '../core/types';

export class TextAtom extends Atom {
  constructor(command: string, value: string, style: Style) {
    super({
      type: 'text',
      command,
      mode: 'text',
      displayContainsHighlight: true,
    });
    this.value = value;
    this.verbatimLatex = value;
    this.applyStyle(style);
  }

  static fromJson(json: AtomJson): TextAtom {
    return new TextAtom(json.command, json.value, json.style);
  }

  _serialize(_options: ToLatexOptions): string {
    return this.verbatimLatex ?? charToLatex('text', this.value.codePointAt(0));
  }
}

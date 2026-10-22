import type { ParseMode } from '../public/core-types';

import { Atom } from '../core/atom-class';
import type { AtomJson, ToLatexOptions } from '../core/types';
import { getDefinition } from '../latex-commands/definitions-utils';

export class GroupAtom extends Atom {
  constructor(arg: readonly Atom[], mode: ParseMode) {
    super({ type: 'group', mode });
    this.body = arg;

    this.skipBoundary = true;
    this.displayContainsHighlight = false;

    // French decimal point, i.e. `{,}`
    if (arg?.length === 1 && arg[0].command === ',')
      this.captureSelection = true;
  }

  static fromJson(json: AtomJson): GroupAtom {
    return new GroupAtom(json.body, json.mode);
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

    return `{${this.bodyToLatex(options)}}`;
  }
}

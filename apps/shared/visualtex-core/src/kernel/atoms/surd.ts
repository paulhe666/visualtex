import type { ParseMode } from '../public/core-types';

import { Atom } from '../core/atom-class';

import { latexCommand } from '../core/tokenizer';
import { getDefinition } from '../latex-commands/definitions-utils';
import type {
  CreateAtomOptions,
  AtomJson,
  ToLatexOptions,
} from '../core/types';

export class SurdAtom extends Atom {
  constructor(
    options: CreateAtomOptions & {
      mode?: ParseMode;
      body: readonly Atom[];
      index: undefined | readonly Atom[];
    }
  ) {
    super({
      ...options,
      type: 'surd',
      mode: options.mode ?? 'math',
      style: options.style,
      displayContainsHighlight: true,
      body: options.body,
    });
    this.above = options.index;
  }

  static fromJson(json: AtomJson): SurdAtom {
    return new SurdAtom({
      ...(json as any),
      index: json.above,
    });
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

    const command = this.command;
    const body = this.bodyToLatex(options);
    if (this.above && !this.hasEmptyBranch('above'))
      return latexCommand(`${command}[${this.aboveToLatex(options)}]`, body);

    if (/^[0-9]$/.test(body)) return `${command}${body}`;

    return latexCommand(command, body);
  }


}

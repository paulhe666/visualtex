import { Atom } from '../core/atom-class';
import { latexCommand } from '../core/tokenizer';
import { getDefinition } from '../latex-commands/definitions-utils';
import type {
  CreateAtomOptions,
  AtomJson,
  ToLatexOptions,
} from '../core/types';

export class MiddleDelimAtom extends Atom {
  size: 1 | 2 | 3 | 4;
  constructor(
    options: CreateAtomOptions & {
      delim: string;
      size: 1 | 2 | 3 | 4;
    }
  ) {
    super({ ...options, type: 'delim' });
    this.value = options.delim;
    this.size = options.size;
  }

  static fromJson(json: AtomJson): MiddleDelimAtom {
    return new MiddleDelimAtom(json as any);
  }

  toJson(): AtomJson {
    return { ...super.toJson(), delim: this.value, size: this.size };
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

    return latexCommand(this.command, this.value);
  }
}

/** The math class a sized delimiter takes (\bigl opens, \bigr closes, \bigm relates) */
type SizedDelimClass = 'mopen' | 'mclose' | 'mrel' | 'mord';

export class SizedDelimAtom extends Atom {
  protected delimType: SizedDelimClass;
  private readonly size: 1 | 2 | 3 | 4;
  constructor(
    options: CreateAtomOptions & {
      delim: string;
      delimType: SizedDelimClass;
      size: 1 | 2 | 3 | 4;
    }
  ) {
    super({ ...options, type: 'sizeddelim', value: options.delim });
    this.delimType = options.delimType;
    this.size = options.size;
  }

  static fromJson(json: AtomJson): SizedDelimAtom {
    return new SizedDelimAtom(json as any);
  }

  toJson(): AtomJson {
    return {
      ...super.toJson(),
      delim: this.value,
      size: this.size,
      delimType: this.delimType,
    };
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
    return latexCommand(this.command, this.value);
  }
}

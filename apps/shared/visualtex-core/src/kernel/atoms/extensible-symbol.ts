import type { Variant } from '../public/core-types';

import { Atom } from '../core/atom-class';
import { joinLatex } from '../core/tokenizer';
import { getDefinition } from '../latex-commands/definitions-utils';
import type {
  CreateAtomOptions,
  AtomJson,
  ToLatexOptions,
} from '../core/types';

/**
 * Operators are handled in the TeXbook pg. 443-444, rule 13(a).
 */
export class ExtensibleSymbolAtom extends Atom {
  private readonly variant?: Variant;

  constructor(
    symbol: string,
    options: CreateAtomOptions & {
      isFunction?: boolean;
      // Unlike `style`, `variant` and `variantStyle` are applied to the
      // content of this atom, but not propagated to the next atom
      variant?: Variant;
      limits?: 'auto' | 'over-under' | 'adjacent';
    }
  ) {
    super({
      ...options,
      type: 'extensible-symbol',
      isFunction: options?.isFunction,
    });
    this.value = symbol;

    this.variant = options?.variant;
    this.subsupPlacement = options?.limits;
  }

  static fromJson(json: AtomJson): ExtensibleSymbolAtom {
    return new ExtensibleSymbolAtom(json.symbol, json as any);
  }

  toJson(): AtomJson {
    const result = super.toJson();

    if (this.variant) result.variant = this.variant;
    if (this.subsupPlacement) result.limits = this.subsupPlacement;
    if (this.value) result.symbol = this.value;
    return result;
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

    const result: string[] = [];

    result.push(this.command!);

    if (this.explicitSubsupPlacement) {
      if (this.subsupPlacement === 'over-under') result.push('\\limits');
      if (this.subsupPlacement === 'adjacent') result.push('\\nolimits');
      if (this.subsupPlacement === 'auto') result.push('\\displaylimits');
    }

    result.push(this.supsubToLatex(options));
    return joinLatex(result);
  }
}

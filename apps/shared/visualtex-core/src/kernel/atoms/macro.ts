import { Atom } from '../core/atom-class';
import type { ParseMode, Style } from '../public/core-types';
import type { AtomJson, ToLatexOptions } from '../core/types';

export class MacroAtom extends Atom {
  readonly macroArgs: null | string;
  // VisualTeX: the number of arguments of an unexpanded macro whose expansion
  // marks them (\visualtexmacroarg), so they can be edited as its slots.
  readonly argumentCount: number;
  // If false, even if `expandMacro` is true, do not expand.
  private readonly expand: boolean;

  constructor(
    macro: string,
    options: {
      expand?: boolean;
      args: null | string;
      argumentCount?: number;
      body: readonly Atom[];
      captureSelection?: boolean;
      style: Style;
    }
  ) {
    super({ type: 'macro', command: macro, style: options.style });
    this.body = options.body;
    // Set the `captureSelection` attribute to true so that the atom is handled
    // as an unbreakable unit
    if (options.captureSelection === undefined) {
      if (options.args) this.captureSelection = false;
      else this.captureSelection = true;
    } else this.captureSelection = options.captureSelection;

    // Don't use verbatimLatex to save the macro, as it can get wiped when
    // the atom is modified (adding super/subscript, for example).
    this.macroArgs = options.args;
    this.argumentCount = options.argumentCount ?? 0;

    this.expand = options.expand ?? false;
  }

  static fromJson(json: AtomJson): MacroAtom {
    return new MacroAtom(json.command, json as any);
  }

  toJson(): AtomJson {
    const options = super.toJson();
    if (this.expand) options.expand = true;
    if (this.captureSelection !== undefined)
      options.captureSelection = this.captureSelection;
    if (this.macroArgs) options.args = this.macroArgs;
    if (this.argumentCount) options.argumentCount = this.argumentCount;
    return options;
  }

  _serialize(options: ToLatexOptions): string {
    if (options.expandMacro && this.expand) return this.bodyToLatex(options);
    // An argument edited in place is saved from its slot, not from the
    // argument text the macro was parsed with.
    if (this.argumentCount > 0) {
      const argumentsByIndex = new Map<number, MacroArgumentAtom>();
      for (const child of this.children) {
        if (
          child.type === 'macro-argument' &&
          Number.isInteger((child as MacroArgumentAtom).argumentIndex) &&
          !argumentsByIndex.has((child as MacroArgumentAtom).argumentIndex)
        ) {
          argumentsByIndex.set((child as MacroArgumentAtom).argumentIndex, child as MacroArgumentAtom);
        }
      }
      if (
        Array.from({ length: this.argumentCount }, (_, index) => index + 1)
          .every((index) => argumentsByIndex.has(index))
      ) {
        return this.command + Array.from(
          { length: this.argumentCount },
          (_, index) => {
            const argument = argumentsByIndex.get(index + 1)!;
            return '{' + argument.bodyToLatex(options) + '}';
          }
        ).join('');
      }
    }
    return this.command + (this.macroArgs ?? '');
  }

  applyStyle(style: Style, options?: { unstyledOnly: boolean }): void {
    // For macros, we only allow color styling. The macro itself has control
    // over the other style attributes
    const allowedStyle: Style = {};
    if (style.color) allowedStyle.color = style.color;
    if (style.backgroundColor)
      allowedStyle.backgroundColor = style.backgroundColor;

    super.applyStyle(allowedStyle, options);
  }
}

/** One argument of an unexpanded macro, inside its expansion: an editable slot. */
export class MacroArgumentAtom extends Atom {
  argumentIndex: number;

  constructor(
    argumentIndex: number = 0,
    body: readonly Atom[] = [],
    options: { mode?: ParseMode; style?: Style } = {}
  ) {
    super({
      type: 'macro-argument',
      mode: options.mode ?? 'math',
      style: options.style,
      body,
    });
    this.argumentIndex = Number(argumentIndex) || 0;
  }

  static fromJson(json: AtomJson): MacroArgumentAtom {
    return new MacroArgumentAtom(
      json.argumentIndex ?? 0,
      json.body ?? [],
      json as any
    );
  }

  toJson(): AtomJson {
    const options = super.toJson();
    if (this.argumentIndex) options.argumentIndex = this.argumentIndex;
    return options;
  }

  _serialize(options: ToLatexOptions): string {
    return this.bodyToLatex(options);
  }
}

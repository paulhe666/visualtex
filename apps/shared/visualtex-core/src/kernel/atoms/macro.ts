import { Atom } from '../core/atom-class';
import type { ParseMode, Style } from '../public/core-types';
import type { AtomJson, ToLatexOptions } from '../core/types';

/**
 * VisualTeX: how an edited argument slot of a package command is written back
 * into its source, by argument kind (a siunitx number, a matrix body...).
 * The package commands register their writers; an unknown kind is written as
 * the slot's LaTeX.
 */
export const macroArgumentWriters: Record<
  string,
  (latex: string, kind: string) => string
> = {};

function writeBackArgument(kind: string, latex: string): string {
  const writer = macroArgumentWriters[kind.replace(/:.*/, '')];
  return writer ? writer(latex, kind) : latex;
}

/** A command name and what follows it: a space ends the name before a letter. */
function joinCommand(command: string, rest: string): string {
  return /\\[a-zA-Z]+$/.test(command) && /^[a-zA-Z0-9]/.test(rest) ? `${command} ${rest}` : command + rest;
}

export class MacroAtom extends Atom {
  readonly macroArgs: null | string;
  // VisualTeX: the number of arguments of an unexpanded macro whose expansion
  // marks them (\visualtexmacroarg), so they can be edited as its slots.
  readonly argumentCount: number;
  // If false, even if `expandMacro` is true, do not expand.
  private readonly expand: boolean;
  // VisualTeX: a package command with xparse-style arguments (physics,
  // siunitx, braket: latex-commands/visualtex-packages.ts). Its source is the
  // template: literal text (star, brackets, delimiters) and argument numbers;
  // an argument not edited since parsing is saved as written (its source),
  // an edited one from its slot (`argumentKinds` says how to write it back).
  readonly argumentTemplate?: readonly (string | number)[];
  readonly argumentSources?: readonly string[];
  readonly argumentKinds?: readonly string[];
  argumentDisplays?: string[];

  constructor(
    macro: string,
    options: {
      expand?: boolean;
      args: null | string;
      argumentCount?: number;
      argumentTemplate?: readonly (string | number)[];
      argumentSources?: readonly string[];
      argumentKinds?: readonly string[];
      argumentDisplays?: string[];
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

    if (options.argumentTemplate) {
      this.argumentTemplate = options.argumentTemplate;
      this.argumentSources = options.argumentSources ?? [];
      this.argumentKinds = options.argumentKinds ?? [];
      // what each slot holds as parsed: a slot still holding it is unedited
      this.argumentDisplays =
        options.argumentDisplays ??
        this.argumentSources.map((_, index) => {
          const slot = this.argumentSlots().get(index + 1);
          return slot ? slot.bodyToLatex({ defaultMode: 'math' }) : '';
        });
    } else if (this.argumentCount > 0)
      // VisualTeX: likewise for a plain macro: unedited, it is saved with its
      // arguments as written (\mod 3, \sideset{_a}{_b}\sum), not braced
      this.argumentDisplays = options.argumentDisplays ?? this.slotDisplays();
  }

  /** VisualTeX: what each of this macro's slots holds now. */
  private slotDisplays(): string[] {
    const slots = this.argumentSlots();
    return Array.from({ length: this.argumentCount }, (_, index) =>
      slots.get(index + 1)?.bodyToLatex({ defaultMode: 'math' }) ?? ''
    );
  }

  /**
   * VisualTeX: this macro's own argument slots by number (the first of each),
   * not those of a macro nested in one of its arguments.
   */
  argumentSlots(): Map<number, MacroArgumentAtom> {
    const slots = new Map<number, MacroArgumentAtom>();
    for (const child of this.children) {
      if (child.type !== 'macro-argument') continue;
      const index = (child as MacroArgumentAtom).argumentIndex;
      if (!Number.isInteger(index) || slots.has(index)) continue;
      let owner = child.parent;
      while (owner && owner.type !== 'macro') owner = owner.parent;
      if (owner === this) slots.set(index, child as MacroArgumentAtom);
    }
    return slots;
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
    if (this.argumentTemplate) {
      options.argumentTemplate = this.argumentTemplate;
      options.argumentSources = this.argumentSources;
      options.argumentKinds = this.argumentKinds;
    }
    if (this.argumentDisplays) options.argumentDisplays = this.argumentDisplays;
    return options;
  }

  _serialize(options: ToLatexOptions): string {
    if (options.expandMacro && this.expand) return this.bodyToLatex(options);
    if (this.argumentTemplate) return this.packageSource(options);
    // An argument edited in place is saved from its slot, not from the
    // argument text the macro was parsed with.
    if (this.argumentCount > 0) {
      // VisualTeX: only this macro's slots (a nested macro's come first
      // among the descendants: \comm{\comm{A}{B}}{C})
      const argumentsByIndex = this.argumentSlots();
      if (
        Array.from({ length: this.argumentCount }, (_, index) => index + 1)
          .every((index) => argumentsByIndex.has(index))
      ) {
        const displays = this.slotDisplays();
        if (
          this.macroArgs &&
          this.argumentDisplays?.every((latex, index) => latex === displays[index])
        )
          return joinCommand(this.command, this.macroArgs);
        return this.command + Array.from(
          { length: this.argumentCount },
          (_, index) => {
            const argument = argumentsByIndex.get(index + 1)!;
            return '{' + argument.bodyToLatex(options) + '}';
          }
        ).join('');
      }
    }
    return joinCommand(this.command, this.macroArgs ?? '');
  }

  /** VisualTeX: a package command's source from its template. */
  private packageSource(options: ToLatexOptions): string {
    const slots = this.argumentSlots();
    let result = this.command;
    for (const part of this.argumentTemplate!) {
      let text: string;
      if (typeof part === 'string') text = part;
      else {
        const source = this.argumentSources![part - 1] ?? '';
        const kind = this.argumentKinds![part - 1] ?? '';
        const slot = slots.get(part);
        const latex = slot?.bodyToLatex({ defaultMode: 'math' });
        if (!slot || latex === this.argumentDisplays?.[part - 1]) text = source;
        else {
          text = writeBackArgument(kind, slot.bodyToLatex(options));
          // an unbraced argument edited into more than one token
          if (kind === 'bare' && !/^(?:\\[a-zA-Z]+|\\.|[^\\{}\s])$/.test(text))
            text = `{${text}}`;
        }
      }
      // a command name is ended by a space before a letter
      if (/\\[a-zA-Z]+$/.test(result) && /^[a-zA-Z]/.test(text)) result += ' ';
      result += text;
    }
    return result;
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

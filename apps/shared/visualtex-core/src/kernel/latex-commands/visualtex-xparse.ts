// VisualTeX: xparse-style argument scanning for the package commands of
// visualtex-packages.ts (physics, siunitx, braket). The arguments are read
// from the parser's tokens as text; the command's expansion is LaTeX in which
// each argument the user can edit is a slot (\visualtexmacroarg), and the
// command's source is kept as a template (MacroAtom.argumentTemplate).
import { tokensToString } from '../core/tokenizer';

/**
 * The parser's tokens and the position of the next one. The tokenizer makes
 * the stars after a command name part of it (`\abs*`): `stars` are those of
 * the command being read, taken by its `s` arguments first.
 */
export type Cursor = { readonly tokens: readonly string[]; index: number; stars: number };

const SPACE = '<space>';

/** The index of the next token that is not a space (xparse skips spaces before optional arguments). */
function nextIndex(c: Cursor): number {
  let i = c.index;
  while (c.tokens[i] === SPACE) i += 1;
  return i;
}

export function tokenText(tokens: readonly string[]): string {
  return tokensToString(tokens as string[]);
}

/** The next token, spaces skipped (the cursor does not move). */
export function peekToken(c: Cursor): string | undefined {
  return c.tokens[nextIndex(c)];
}

/** xparse `t`/`s`: the token, if it comes next. */
export function scanToken(c: Cursor, token: string): boolean {
  const i = nextIndex(c);
  if (c.tokens[i] !== token) return false;
  c.index = i + 1;
  return true;
}

/** xparse `g`: a brace group's text, if a brace group comes next. */
export function scanGroup(c: Cursor): string | undefined {
  const i = nextIndex(c);
  if (c.tokens[i] !== '<{>') return undefined;
  let depth = 0;
  for (let j = i; j < c.tokens.length; j += 1) {
    const t = c.tokens[j];
    if (t === '<{>') depth += 1;
    else if (t === '<}>') {
      depth -= 1;
      if (depth === 0) {
        c.index = j + 1;
        return tokenText(c.tokens.slice(i + 1, j));
      }
    }
  }
  return undefined;
}

/**
 * xparse `o` / `d<open><close>`: the text between the delimiters, if they come
 * next. Braces are balanced, and so are nested pairs of different delimiters
 * (`\qty(a(b))`).
 */
export function scanDelimited(c: Cursor, open: string, close: string): string | undefined {
  const i = nextIndex(c);
  if (c.tokens[i] !== open) return undefined;
  let braces = 0;
  let depth = 0;
  for (let j = i + 1; j < c.tokens.length; j += 1) {
    const t = c.tokens[j];
    if (t === '<{>') braces += 1;
    else if (t === '<}>') {
      if (braces === 0) return undefined;
      braces -= 1;
    } else if (braces === 0) {
      if (t === close && depth === 0) {
        c.index = j + 1;
        return tokenText(c.tokens.slice(i + 1, j));
      }
      if (open !== close && t === open) depth += 1;
      else if (open !== close && t === close) depth -= 1;
    }
  }
  return undefined;
}

const NOT_AN_ARGUMENT = new Set(['<}>', '&', '\\\\', '\\cr', '\\end', '^', '_', '<$>', '<$$>', '\\right', '\\middle']);

/** xparse `m`: a brace group, else the next token. */
export function scanMandatory(c: Cursor): { text: string; braced: boolean } | undefined {
  const group = scanGroup(c);
  if (group !== undefined) return { text: group, braced: true };
  const i = nextIndex(c);
  const t = c.tokens[i];
  if (t === undefined || NOT_AN_ARGUMENT.has(t)) return undefined;
  c.index = i + 1;
  return { text: t, braced: false };
}

/**
 * The arguments of one command as they are read: its source template and the
 * argument texts, and the slots of its expansion.
 */
export class Args {
  readonly template: (string | number)[] = [];
  readonly sources: string[] = [];
  readonly kinds: string[] = [];
  private readonly shown = new Set<number>();

  literal(text: string): void {
    if (text) this.template.push(text);
  }

  /** An argument written between `open` and `close` in the source: its number. */
  add(open: string, source: string, close: string, kind = ''): number {
    this.sources.push(source);
    this.kinds.push(kind);
    const n = this.sources.length;
    this.literal(open);
    this.template.push(n);
    this.literal(close);
    return n;
  }

  source(n: number): string {
    return this.sources[n - 1] ?? '';
  }

  /**
   * An argument in the expansion: its first appearance is its editable slot
   * (drawn as `display`), a repeated one a copy.
   */
  slot(n: number, display = this.source(n)): string {
    if (this.shown.has(n)) return `{${display}}`;
    this.shown.add(n);
    return `\\visualtexmacroarg{${n}}{${display}}`;
  }

  /** An argument where it is not seen (a phantom): never a slot. */
  copy(n: number | undefined, display = n === undefined ? '' : this.source(n)): string {
    return `{${display}}`;
  }

  // ------------------------------------------------ xparse argument types --
  /** `s`: a star. */
  s(c: Cursor): boolean {
    if (c.stars > 0) {
      c.stars -= 1;
      return true;
    }
    if (!scanToken(c, '*')) return false;
    this.literal('*');
    return true;
  }

  /** `t<token>`. */
  t(c: Cursor, token: string): boolean {
    if (!scanToken(c, token)) return false;
    this.literal(token);
    return true;
  }

  /** `m`: missing (a command typed without its argument), an empty one. */
  m(c: Cursor, kind = ''): number {
    const arg = scanMandatory(c);
    if (!arg) return this.add('', '', '', 'bare');
    return arg.braced ? this.add('{', arg.text, '}', kind) : this.add('', arg.text, '', kind || 'bare');
  }

  /** `g`. */
  g(c: Cursor, kind = ''): number | undefined {
    const text = scanGroup(c);
    return text === undefined ? undefined : this.add('{', text, '}', kind);
  }

  /** `o`. */
  o(c: Cursor, kind = ''): number | undefined {
    return this.d(c, '[', ']', kind);
  }

  /** `d<open><close>`. */
  d(c: Cursor, open: string, close: string, kind = ''): number | undefined {
    const text = scanDelimited(c, open, close);
    return text === undefined ? undefined : this.add(open, text, close, kind);
  }
}

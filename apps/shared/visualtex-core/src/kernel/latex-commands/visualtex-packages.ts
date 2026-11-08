// VisualTeX: commands of physics.sty (1.3), braket.sty and siunitx (v3) whose
// arguments are optional, delimited or variable (xparse), read as the
// packages read them. Each becomes a MacroAtom: its expansion follows the
// package's definition (the .sty code is quoted at each command), its
// arguments are the slots the user edits, and its source is saved as written.
//
// The parser asks here first for every command (Parser.scanMacro): a command
// of the table below, in math mode, is read here; a form that is not the
// package's (`\cos x`, a plain LaTeX meaning) is left to the parser.
import { MacroAtom, macroArgumentWriters } from '../atoms/macro';
import type { Atom } from '../core/atom-class';
import type { Style } from '../public/core-types';
import { Args, type Cursor, peekToken, scanDelimited, scanGroup, scanToken, tokenText } from './visualtex-xparse';
import { SIUNITX_COMMANDS, isSiunitxQuantity } from './visualtex-siunitx';

/** A command's reading: its expansion, or null when the input is not its form. */
type PackageCommand = (c: Cursor, a: Args, name: string) => string | null;

const BIG = ['\\big', '\\Big', '\\bigg', '\\Bigg'];

/** physics `l` before \braces: `*` (no resize) or a manual size. */
function sizeArgument(c: Cursor, a: Args): string {
  if (a.s(c)) return '*';
  for (const size of BIG) if (a.t(c, size)) return size;
  return '';
}

/**
 * physics \braces: `{\left<open>{#}\right<close>}`. The braces (physics
 * `{\ifnum\z@=`}\fi ... \ifnum\z@=`{\fi}`) make it an ordinary atom.
 */
function braces(size: string, open: string, close: string, body: string, phantom: string): string {
  if (size === '*') return `{\\left${open}\\smash{${body}}\\right${close}\\vphantom{${phantom}}}`;
  if (size) return `{${size}l${open}{${body}}${size}r${close}}`;
  return `{\\left${open}{${body}}\\right${close}}`;
}

/**
 * physics \fbraces: the function's argument after it,
 * \argopen = \mathopen{}\mathclose\bgroup\left, \argclose = \aftergroup\egroup\right.
 */
function fbraces(size: string, open: string, close: string, fn: string, body: string, phantom: string): string {
  if (size === '*') return `${fn}\\mathopen{}\\mathclose{\\left${open}\\smash{${body}}\\right${close}}\\vphantom{${phantom}}`;
  if (size) return `${fn}\\mathopen{}\\mathclose{${size}l${open}{${body}}}${size}r${close}`;
  return `${fn}\\mathopen{}\\mathclose{\\left${open}{${body}}\\right${close}}`;
}

/** `\argopen(#\argclose)` */
const argParens = (body: string) => `\\mathopen{}\\mathclose{\\left(${body}\\right)}`;

/** physics `l m`: \absolutevalue, \norm, \pqty... */
function bracesCommand(open: string, close: string): PackageCommand {
  return (c, a) => {
    const size = sizeArgument(c, a);
    const n = a.m(c);
    return braces(size, open, close, a.slot(n), a.copy(n));
  };
}

/** physics `l m m`: \commutator, \anticommutator, \poissonbracket. */
function bracketCommand(open: string, close: string): PackageCommand {
  return (c, a) => {
    const size = sizeArgument(c, a);
    const x = a.m(c);
    const y = a.m(c);
    return braces(size, open, close, `${a.slot(x)},${a.slot(y)}`, `${a.copy(x)},${a.copy(y)}`);
  };
}

/** The optional arguments physics prints as written after the used one (`\qty{a}[b]`). */
function trailing(a: Args, parts: [number | undefined, string, string][]): string {
  return parts.map(([n, open, close]) => (n === undefined ? '' : `${open}${a.slot(n)}${close}`)).join('');
}

/**
 * physics \quantity: `t\big t\Big t\bigg t\Bigg g o d() d||`, in a group.
 *   {\ltag\lbrace#5\rtag\rbrace}, \ltag[#6\rtag], \ltag(#7\rtag), \ltag\lvert{#8}\rtag\rvert, else ()
 */
const quantity: PackageCommand = (c, a) => {
  let size = '';
  for (const big of BIG) if (!size && a.t(c, big)) size = big;
  const [l, r] = size ? [`${size}l`, `${size}r`] : ['\\left', '\\right'];
  const brace = a.g(c);
  const bracket = a.o(c);
  const paren = a.d(c, '(', ')');
  const bar = a.d(c, '|', '|');
  let body: string;
  if (brace !== undefined)
    body = `${l}\\lbrace ${a.slot(brace)}${r}\\rbrace ` + trailing(a, [[bracket, '[', ']'], [paren, '(', ')'], [bar, '|', '|']]);
  else if (bracket !== undefined)
    body = `${l}[${a.slot(bracket)}${r}] ` + trailing(a, [[paren, '(', ')'], [bar, '|', '|']]);
  else if (paren !== undefined) body = `${l}(${a.slot(paren)}${r}) ` + trailing(a, [[bar, '|', '|']]);
  else if (bar !== undefined) body = `${l}\\lvert{${a.slot(bar)}}${r}\\rvert `;
  else body = '()';
  return `{${body}}`;
};

// -------------------------------------------------------------- matrices --
/**
 * physics' matrix generators inside a matrix's content, written out as TeX
 * expands them (\identitymatrix, \paulimatrix, \xmatrix, \zeromatrix).
 */
function expandMatrixGenerators(content: string): string {
  const generators: [RegExp, (...m: string[]) => string | null][] = [
    [/\\(?:imat|identitymatrix)\s*\{\s*(\d+)\s*\}/g, (_, n) => matrixOf(+n, +n, (i, j) => (i === j ? '1' : '0'))],
    [
      /\\(?:pmat|paulimatrix)\s*\{\s*([0-3xyz])\s*\}/g,
      (_, k) =>
        ({
          '0': '1&0\\\\0&1',
          '1': '0&1\\\\1&0',
          x: '0&1\\\\1&0',
          '2': '0&-i\\\\i&0',
          y: '0&-i\\\\i&0',
          '3': '1&0\\\\0&-1',
          z: '1&0\\\\0&-1',
        })[k] ?? null,
    ],
    [
      /\\(?:xmat|xmatrix)(\*?)\s*\{([^{}]*)\}\s*\{\s*(\d+)\s*\}\s*\{\s*(\d+)\s*\}/g,
      (_, star, x, r, cols) =>
        matrixOf(+r, +cols, (i, j) => (star ? `${x}_{${+r > 1 ? i + 1 : ''}${+cols > 1 ? j + 1 : ''}}` : x)),
    ],
    [/\\(?:zmat|zeromatrix)\s*\{\s*(\d+)\s*\}(?:\s*\{\s*(\d+)\s*\})?/g, (_, r, cols) => matrixOf(+r, +(cols ?? r), () => '0')],
  ];
  for (const [pattern, expand] of generators) content = content.replace(pattern, (...m) => expand(...(m as string[])) ?? m[0]);
  return content;
}

function matrixOf(rows: number, cols: number, cell: (i: number, j: number) => string): string | null {
  if (!(rows > 0 && cols > 0 && rows <= 20 && cols <= 20)) return null;
  return Array.from({ length: rows }, (_, i) => Array.from({ length: cols }, (_, j) => cell(i, j)).join('&')).join('\\\\');
}

/** A matrix environment whose content is the argument's slot. */
function matrixSlot(a: Args, n: number, environment: string): string {
  a.kinds[n - 1] = `matrix:${environment}`;
  return a.slot(n, `\\begin{${environment}}${expandMatrixGenerators(a.source(n))}\\end{${environment}}`);
}

/**
 * physics \matrixquantity / \smallmatrixquantity: `s g o d() d||`, an ord:
 * {matrix}, \bmqty, \pmqty (\Pmqty starred), \vmqty, else ().
 */
function matrixQuantity(small: boolean): PackageCommand {
  return (c, a) => {
    const star = a.s(c);
    const brace = a.g(c);
    const bracket = a.o(c);
    const paren = a.d(c, '(', ')');
    const bar = a.d(c, '|', '|');
    const env = small ? 'smallmatrix' : 'matrix';
    const smallIn = (open: string, close: string, n: number) => `{\\left${open}{${matrixSlot(a, n, 'smallmatrix')}}\\right${close}}`;
    let body: string;
    if (brace !== undefined)
      body = matrixSlot(a, brace, env) + trailing(a, [[bracket, '[', ']'], [paren, '(', ')'], [bar, '|', '|']]);
    else if (bracket !== undefined)
      body = (small ? smallIn('[', ']', bracket) : matrixSlot(a, bracket, 'bmatrix')) + trailing(a, [[paren, '(', ')'], [bar, '|', '|']]);
    else if (paren !== undefined) {
      const matrix = star
        ? `\\left\\lgroup ${matrixSlot(a, paren, env)}\\right\\rgroup `
        : small
          ? smallIn('(', ')', paren)
          : matrixSlot(a, paren, 'pmatrix');
      body = matrix + trailing(a, [[bar, '|', '|']]);
    } else if (bar !== undefined) body = small ? smallIn('\\lvert', '\\rvert', bar) : matrixSlot(a, bar, 'vmatrix');
    else body = '()';
    return `\\mathord{${body}}`;
  };
}

/**
 * An edited matrix slot: the environment's content; anything typed beside
 * the matrix in its slot joins its first or last cell.
 */
macroArgumentWriters.matrix = (latex, kind) => {
  const environment = kind.slice('matrix:'.length);
  const open = latex.indexOf(`\\begin{${environment}}`);
  const close = latex.lastIndexOf(`\\end{${environment}}`);
  if (open < 0 || close < open) return latex;
  const inner = latex.slice(open + `\\begin{${environment}}`.length, close);
  return latex.slice(0, open).trim() + inner + latex.slice(close + `\\end{${environment}}`.length).trim();
};

// ------------------------------------------------------------ derivatives --
const DIFF = '\\mathrm{d}';

/** physics' \fractype: \frac, or \flatfrac (`\left.#1\middle\slash#2\right.`) when starred. */
function fraction(flat: boolean, num: string, den: string): string {
  return flat ? `\\left.${num}\\middle/${den}\\right.` : `\\frac{${num}}{${den}}`;
}

/**
 * physics \derivative / \functionalderivative: `s o m g d()`.
 *   no f: \fractype{\diffd^{n}}{\diffd #3^{n}} [\argopen(#5\argclose)]
 *   f:    \fractype{\diffd^{n} #3}{\diffd #4^{n}}
 */
function derivative(d: string): PackageCommand {
  return (c, a) => {
    const flat = a.s(c);
    const n = a.o(c);
    const x = a.m(c);
    const f = a.g(c);
    const paren = f === undefined ? a.d(c, '(', ')') : undefined;
    const pow = n === undefined ? '' : `^{${a.slot(n)}}`;
    if (f !== undefined) return fraction(flat, `${d}${pow} ${a.slot(x)}`, `${d}${a.slot(f)}${pow}`);
    const op = fraction(flat, `${d}${pow}`, `${d}${a.slot(x)}${pow}`);
    return paren === undefined ? op : op + argParens(a.slot(paren));
  };
}

/**
 * physics \partialderivative: `s o m g g d()`; with two variables the mixed
 * \fractype{\partial^2 #3}{\partial #4 \partial #5}.
 */
const partialDerivative: PackageCommand = (c, a) => {
  const flat = a.s(c);
  const n = a.o(c);
  const x = a.m(c);
  const f = a.g(c);
  const y = f === undefined ? undefined : a.g(c);
  const paren = f === undefined ? a.d(c, '(', ')') : undefined;
  const pow = n === undefined ? '' : `^{${a.slot(n)}}`;
  if (y !== undefined) return fraction(flat, `\\partial^2 ${a.slot(x)}`, `\\partial ${a.slot(f!)} \\partial ${a.slot(y)}`);
  if (f !== undefined) return fraction(flat, `\\partial${pow} ${a.slot(x)}`, `\\partial ${a.slot(f)}${pow}`);
  const op = fraction(flat, `\\partial${pow}`, `\\partial ${a.slot(x)}${pow}`);
  return paren === undefined ? op : op + argParens(a.slot(paren));
};

/**
 * physics \differential / \variation: `o g d()`.
 *   \diffd^{n};  \mathinner{\diffd^{n}\argopen(#3\argclose)};  \mathinner{\diffd^{n}#2} (#3)
 */
function differential(d: string): PackageCommand {
  return (c, a) => {
    const n = a.o(c);
    const x = a.g(c);
    const paren = a.d(c, '(', ')');
    const pow = n === undefined ? '' : `^{${a.slot(n)}}`;
    if (x !== undefined) return `\\mathinner{${d}${pow}${a.slot(x)}}` + trailing(a, [[paren, '(', ')']]);
    if (paren !== undefined) return `\\mathinner{${d}${pow}${argParens(a.slot(paren))}}`;
    return `${d}${pow}`;
  };
}

// ---------------------------------------------------------- vector calculus --
const VNABLA = '\\boldsymbol{\\nabla}';

/**
 * physics \gradient, \curl, \laplacian: `g o d()`.
 *   \grad: \vnabla #1 | \fbraces{[}{]}{\vnabla}{#2} | \fbraces{(}{)}{\vnabla}{#3} | \vnabla
 *   \curl: \vnabla\cross #1 | \vnabla\cross\quantity[#2] | \vnabla\cross\quantity(#3) | \vnabla\cross
 */
function vectorOperator(op: string, fbraced: boolean): PackageCommand {
  return (c, a) => {
    const x = a.g(c);
    const bracket = a.o(c);
    const paren = a.d(c, '(', ')');
    if (x !== undefined) return `${op} ${a.slot(x)}` + trailing(a, [[bracket, '[', ']'], [paren, '(', ')']]);
    if (bracket !== undefined) {
      const body = fbraced ? fbraces('', '[', ']', op, a.slot(bracket), '') : `${op}{\\left[${a.slot(bracket)}\\right]}`;
      return body + trailing(a, [[paren, '(', ')']]);
    }
    if (paren !== undefined) return fbraced ? fbraces('', '(', ')', op, a.slot(paren), '') : `${op}{\\left(${a.slot(paren)}\\right)}`;
    return op;
  };
}

/**
 * physics \opbraces (\tr, \Tr, \Res with the trig option): `g o d()`,
 * \fbraces{\lbrace}{\rbrace}, [ ], ( ) after the operator.
 */
function opBraces(op: string): PackageCommand {
  return (c, a) => {
    const brace = a.g(c);
    const bracket = a.o(c);
    const paren = a.d(c, '(', ')');
    if (brace !== undefined)
      return fbraces('', '\\lbrace', '\\rbrace', op, a.slot(brace), '') + trailing(a, [[bracket, '[', ']'], [paren, '(', ')']]);
    if (bracket !== undefined) return fbraces('', '[', ']', op, a.slot(bracket), '') + trailing(a, [[paren, '(', ')']]);
    if (paren !== undefined) return fbraces('', '(', ')', op, a.slot(paren), '');
    return op;
  };
}

/**
 * physics \trigbraces: `o d()`: #1^{#2}\argopen(#3\argclose), else #1[#2].
 * For a function plain LaTeX defines (\sin, \log, ...) only the physics form
 * with a power, `\cos[2](x)`, is read here: \sin(x) keeps its LaTeX meaning.
 */
function trigBraces(op: string, latexDefines: boolean): PackageCommand {
  return (c, a) => {
    if (latexDefines) {
      // [power] followed by (argument)
      const probe = { ...c };
      if (scanDelimited(probe, '[', ']') === undefined || peekToken(probe) !== '(') return null;
    }
    const n = a.o(c);
    const paren = a.d(c, '(', ')');
    if (paren === undefined) return n === undefined ? op : `${op}[${a.slot(n)}]`;
    return `${op}${n === undefined ? '' : `^{${a.slot(n)}}`}${argParens(a.slot(paren))}`;
  };
}

// ---------------------------------------------------------------- Dirac --
/**
 * physics \innerproduct (\braket, \ip): `s m g`.
 *   \left\langle{#2}\middle\vert{#3}\right\rangle (#3 = #2 when absent);
 *   starred, no resize: \vphantom{#2#3}\left\langle\smash{#2}\middle\vert\smash{#3}\right\rangle
 */
function innerProduct(a: Args, star: boolean, x: number, y: number | undefined): string {
  const ket = y ?? x;
  if (star)
    return `\\vphantom{${a.copy(x)}${a.copy(ket)}}\\left\\langle\\smash{${a.slot(x)}}\\middle\\vert\\smash{${a.slot(ket)}}\\right\\rangle `;
  return `\\left\\langle{${a.slot(x)}}\\middle\\vert{${a.slot(ket)}}\\right\\rangle `;
}

/** Whether braket.sty's argument has a | of its own (not inside braces). */
function hasTopLevelBar(text: string): boolean {
  let depth = 0;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === '\\') {
      i += 1;
      continue;
    }
    if (ch === '{') depth += 1;
    else if (ch === '}') depth -= 1;
    else if (ch === '|' && depth === 0) return true;
  }
  return false;
}

/**
 * \braket: physics' \innerproduct, except braket.sty's one argument with
 * its own | (\braket{\phi|\psi}: \mathinner{\langle{#1}\rangle}).
 */
const braket: PackageCommand = (c, a, name) => {
  const star = a.s(c);
  const x = a.m(c);
  const y = a.g(c);
  if (name === 'braket' && !star && y === undefined && hasTopLevelBar(a.source(x)))
    return `\\mathinner{\\langle{${a.slot(x)}}\\rangle}`;
  return innerProduct(a, star, x, y);
};

/**
 * physics \bra: `s m t\ket s g`; followed by \ket it is their contraction,
 * \braket{#2}{#5} (\braket* if either is starred).
 */
const bra: PackageCommand = (c, a) => {
  const star = a.s(c);
  const x = a.m(c);
  // \ket, \ket* (the tokenizer joins a command's stars to its name)
  const next = peekToken(c);
  if (next !== undefined && /^\\ket\**$/.test(next)) {
    scanToken(c, next);
    a.literal(next);
    const ketStar = next.endsWith('*');
    const y = a.g(c);
    return innerProduct(a, star || ketStar, x, y ?? a.add('', '', ''));
  }
  if (star) return `\\vphantom{${a.copy(x)}}\\left\\langle\\smash{${a.slot(x)}}\\right\\rvert `;
  return `\\left\\langle{${a.slot(x)}}\\right\\rvert `;
};

/** physics \ket: `s m`. */
const ket: PackageCommand = (c, a) => {
  const star = a.s(c);
  const x = a.m(c);
  if (star) return `\\vphantom{${a.copy(x)}}\\left\\lvert\\smash{${a.slot(x)}}\\right\\rangle `;
  return `\\left\\lvert{${a.slot(x)}}\\right\\rangle `;
};

/** physics \outerproduct (\dyad, \ketbra, \op): `s m g`. */
const outerProduct: PackageCommand = (c, a) => {
  const star = a.s(c);
  const x = a.m(c);
  const y = a.g(c) ?? x;
  if (star)
    return `\\vphantom{${a.copy(x)}${a.copy(y)}}\\left\\lvert\\smash{${a.slot(x)}}\\middle\\rangle\\!\\middle\\langle\\smash{${a.slot(y)}}\\right\\rvert `;
  return `\\left\\lvert{${a.slot(x)}}\\middle\\rangle\\!\\middle\\langle{${a.slot(y)}}\\right\\rvert `;
};

/**
 * physics \expectationvalue (\expval, \ev): `s s m g`.
 *   ⟨#3⟩;  with a state: \vphantom{#3}\left\langle{#4}\middle\vert\smash{#3}\middle\vert{#4}\right\rangle
 *   (* : nothing resizes, ** : everything does)
 */
const expectationValue: PackageCommand = (c, a) => {
  const star = a.s(c);
  const double = star && a.s(c);
  const op = a.m(c);
  const state = a.g(c);
  if (state === undefined) {
    if (star) return `\\vphantom{${a.copy(op)}}\\left\\langle\\smash{${a.slot(op)}}\\right\\rangle `;
    return `\\left\\langle{${a.slot(op)}}\\right\\rangle `;
  }
  if (double)
    return `\\left\\langle{${a.slot(state)}}\\middle\\vert{${a.slot(op)}}\\middle\\vert{${a.slot(state)}}\\right\\rangle `;
  if (star)
    return `\\vphantom{${a.copy(op)}${a.copy(state)}}\\left\\langle\\smash{${a.slot(state)}}\\middle\\vert\\smash{${a.slot(op)}}\\middle\\vert\\smash{${a.slot(state)}}\\right\\rangle `;
  return `\\vphantom{${a.copy(op)}}\\left\\langle{${a.slot(state)}}\\middle\\vert\\smash{${a.slot(op)}}\\middle\\vert{${a.slot(state)}}\\right\\rangle `;
};

/** physics \matrixelement (\mel, \matrixel): `s s m m m`. */
const matrixElement: PackageCommand = (c, a) => {
  const star = a.s(c);
  const double = star && a.s(c);
  const x = a.m(c);
  const op = a.m(c);
  const y = a.m(c);
  if (double) return `\\left\\langle{${a.slot(x)}}\\middle\\vert{${a.slot(op)}}\\middle\\vert{${a.slot(y)}}\\right\\rangle `;
  if (star)
    return `\\vphantom{${a.copy(x)}${a.copy(op)}${a.copy(y)}}\\left\\langle\\smash{${a.slot(x)}}\\middle\\vert\\smash{${a.slot(op)}}\\middle\\vert\\smash{${a.slot(y)}}\\right\\rangle `;
  return `\\vphantom{${a.copy(op)}}\\left\\langle{${a.slot(x)}}\\middle\\vert\\smash{${a.slot(op)}}\\middle\\vert{${a.slot(y)}}\\right\\rangle `;
};

/** physics \vev: \expectationvalue{#1}{0}. */
const vacuumExpectationValue: PackageCommand = (c, a) => {
  const op = a.m(c);
  return `\\vphantom{${a.copy(op)}}\\left\\langle{0}\\middle\\vert\\smash{${a.slot(op)}}\\middle\\vert{0}\\right\\rangle `;
};

/**
 * physics \evaluated (\eval): `s g d[| d(|`, the bar as tall as a \int at
 * least: \left.#2\vphantom{\int}\right\rvert (\left[ or \left( with d[| d(|)
 */
const evaluated: PackageCommand = (c, a) => {
  const star = a.s(c);
  const x = a.g(c);
  const bracket = x === undefined ? a.d(c, '[', '|') : undefined;
  const paren = x === undefined && bracket === undefined ? a.d(c, '(', '|') : undefined;
  const open = x !== undefined ? '.' : bracket !== undefined ? '[' : paren !== undefined ? '(' : null;
  const n = x ?? bracket ?? paren;
  if (open === null || n === undefined) return `\\mathopen{}\\mathclose{\\left.\\vphantom{\\int}\\right\\rvert}`;
  if (star) return `\\vphantom{${a.copy(n)}}\\left${open}\\smash{${a.slot(n)}}\\vphantom{\\int}\\right\\rvert `;
  return `\\left${open}${a.slot(n)}\\vphantom{\\int}\\right\\rvert `;
};

// ------------------------------------------------------------- braket.sty --
/** braket.sty's argument split at its own | (|| is one double bar), as slots. */
function barParts(c: Cursor, a: Args): { slots: number[]; bars: string[] } | null {
  const text = scanGroup(c);
  if (text === undefined) return null;
  const tokens = tokenizeBars(text);
  a.literal('{');
  const slots: number[] = [];
  const bars: string[] = [];
  tokens.forEach((part, i) => {
    if (i % 2 === 1) {
      bars.push(part);
      a.literal(part);
    } else slots.push(a.add('', part.trim(), ''));
  });
  a.literal('}');
  return { slots, bars };
}

/** Split at the | not inside braces: [part, bar, part, bar, part...]; || and \| are double bars. */
function tokenizeBars(text: string): string[] {
  const out: string[] = [];
  let depth = 0;
  let start = 0;
  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch === '\\') {
      if (text[i + 1] === '|' && depth === 0) {
        out.push(text.slice(start, i), '\\|');
        start = i + 2;
      }
      i += 1;
      continue;
    }
    if (ch === '{') depth += 1;
    else if (ch === '}') depth -= 1;
    else if (ch === '|' && depth === 0) {
      const double = text[i + 1] === '|';
      out.push(text.slice(start, i), double ? '||' : '|');
      if (double) i += 1;
      start = i + 1;
    }
  }
  out.push(text.slice(start));
  return out;
}

/**
 * braket.sty \Braket: \left\langle{#1}\right\rangle, each | of #1 is
 * `}\,\middle|\,{` (\BraVert), || and \| `}\,\middle\|\,{` (\BraDoubleVert).
 */
const bigBraket: PackageCommand = (c, a) => {
  const parts = barParts(c, a);
  if (!parts) return null;
  const middle = (bar: string) => (bar === '|' ? '\\middle|' : '\\middle\\|');
  return (
    `\\left\\langle{${a.slot(parts.slots[0])}}` +
    parts.bars.map((bar, i) => `\\,${middle(bar)}\\,{${a.slot(parts.slots[i + 1])}}`).join('') +
    `\\right\\rangle `
  );
};

/** braket.sty \Set: \left\{\:{#1}\:\right\}, each | `}\;\middle|\;{` (\SetVert). */
const bigSet: PackageCommand = (c, a) => {
  const parts = barParts(c, a);
  if (!parts) return null;
  const middle = (bar: string) => (bar === '|' ? '\\middle|' : '\\middle\\|');
  return (
    `\\left\\{\\:{${a.slot(parts.slots[0])}}` +
    parts.bars.map((bar, i) => `\\;${middle(bar)}\\;{${a.slot(parts.slots[i + 1])}}`).join('') +
    `\\:\\right\\} `
  );
};

/** braket.sty \set: \mathinner{\lbrace\,{#1}\,\rbrace}, each | `}\mid{` (\midvert). */
const smallSet: PackageCommand = (c, a) => {
  const parts = barParts(c, a);
  if (!parts) return null;
  return (
    `\\mathinner{\\lbrace\\,{${a.slot(parts.slots[0])}}` +
    parts.bars.map((bar, i) => `${bar === '|' ? '\\mid' : '\\|'}{${a.slot(parts.slots[i + 1])}}`).join('') +
    `\\,\\rbrace}`
  );
};

// ------------------------------------------------------------------ table --
const vectorStyle = (plain: (x: string) => string, starred: (x: string) => string): PackageCommand => (c, a) => {
  const star = a.s(c);
  const x = a.slot(a.m(c));
  return star ? starred(x) : plain(x);
};

const PHYSICS_ONLY_FUNCTIONS = ['arccsc', 'arcsec', 'arccot', 'asin', 'acos', 'atan', 'acsc', 'asec', 'acot', 'csch', 'sech'];
const LATEX_FUNCTIONS = [
  'sin', 'cos', 'tan', 'csc', 'sec', 'cot', 'arcsin', 'arccos', 'arctan',
  'sinh', 'cosh', 'tanh', 'coth', 'log', 'ln',
];

const PACKAGE_COMMANDS: Record<string, PackageCommand> = {
  // siunitx (its \qty: below)
  ...SIUNITX_COMMANDS,
  // physics: brackets and braces
  quantity,
  qty: (c, a, name) => (isSiunitxQuantity(c) ? SIUNITX_COMMANDS.qty(c, a, name) : quantity(c, a, name)),
  pqty: bracesCommand('(', ')'),
  bqty: bracesCommand('[', ']'),
  Bqty: bracesCommand('\\lbrace', '\\rbrace'),
  vqty: bracesCommand('\\lvert', '\\rvert'),
  absolutevalue: bracesCommand('\\lvert', '\\rvert'),
  abs: bracesCommand('\\lvert', '\\rvert'),
  norm: bracesCommand('\\lVert', '\\rVert'),
  order: (c, a) => {
    const size = sizeArgument(c, a);
    const n = a.m(c);
    return fbraces(size, '(', ')', '\\mathcal{O}', a.slot(n), a.copy(n));
  },
  evaluated,
  eval: evaluated,
  commutator: bracketCommand('[', ']'),
  comm: bracketCommand('[', ']'),
  anticommutator: bracketCommand('\\lbrace', '\\rbrace'),
  acommutator: bracketCommand('\\lbrace', '\\rbrace'),
  acomm: bracketCommand('\\lbrace', '\\rbrace'),
  poissonbracket: bracketCommand('\\lbrace', '\\rbrace'),
  pb: bracketCommand('\\lbrace', '\\rbrace'),
  // physics: matrices
  matrixquantity: matrixQuantity(false),
  mqty: matrixQuantity(false),
  smallmatrixquantity: matrixQuantity(true),
  smqty: matrixQuantity(true),
  // physics: vectors
  vectorbold: vectorStyle((x) => `\\mathbf{${x}}`, (x) => `\\boldsymbol{${x}}`),
  vb: vectorStyle((x) => `\\mathbf{${x}}`, (x) => `\\boldsymbol{${x}}`),
  vectorarrow: vectorStyle((x) => `\\vec{\\mathbf{${x}}}`, (x) => `\\vec{\\boldsymbol{${x}}}`),
  va: vectorStyle((x) => `\\vec{\\mathbf{${x}}}`, (x) => `\\vec{\\boldsymbol{${x}}}`),
  vectorunit: vectorStyle((x) => `\\mathbf{\\hat{${x}}}`, (x) => `\\boldsymbol{\\hat{${x}}}`),
  vu: vectorStyle((x) => `\\mathbf{\\hat{${x}}}`, (x) => `\\boldsymbol{\\hat{${x}}}`),
  gradient: vectorOperator(VNABLA, true),
  grad: vectorOperator(VNABLA, true),
  curl: vectorOperator(`${VNABLA}\\boldsymbol{\\times}`, false),
  laplacian: vectorOperator('\\nabla^2', true),
  // physics: operators
  tr: opBraces('\\operatorname{tr}'),
  Tr: opBraces('\\operatorname{Tr}'),
  Res: opBraces('\\operatorname{Res}'),
  principalvalue: (c, a) => {
    const x = a.g(c);
    return x === undefined ? '\\mathcal{P}' : `\\mathcal{P}\\mathord{${a.slot(x)}}`;
  },
  pv: (c, a) => {
    const x = a.g(c);
    return x === undefined ? '\\mathcal{P}' : `\\mathcal{P}\\mathord{${a.slot(x)}}`;
  },
  PV: (c, a) => {
    const x = a.g(c);
    return x === undefined ? '\\mathrm{P.V.}' : `\\mathrm{P.V.}\\mathord{${a.slot(x)}}`;
  },
  ...Object.fromEntries(PHYSICS_ONLY_FUNCTIONS.map((name) => [name, trigBraces(`\\operatorname{${name}}`, false)])),
  ...Object.fromEntries(LATEX_FUNCTIONS.map((name) => [name, trigBraces(`\\${name}`, true)])),
  // physics: derivatives
  differential: differential(DIFF),
  dd: differential(DIFF),
  variation: differential('\\delta'),
  var: differential('\\delta'),
  derivative: derivative(DIFF),
  dv: derivative(DIFF),
  partialderivative: partialDerivative,
  pderivative: partialDerivative,
  pdv: partialDerivative,
  functionalderivative: derivative('\\delta'),
  fderivative: derivative('\\delta'),
  fdv: derivative('\\delta'),
  // physics: Dirac notation (\bra, \ket and \braket as physics reads them;
  // \braket{\phi|\psi} as braket.sty)
  bra,
  ket,
  innerproduct: (c, a) => innerProduct(a, a.s(c), a.m(c), a.g(c)),
  ip: (c, a) => innerProduct(a, a.s(c), a.m(c), a.g(c)),
  braket,
  outerproduct: outerProduct,
  dyad: outerProduct,
  ketbra: outerProduct,
  op: outerProduct,
  expectationvalue: expectationValue,
  expval: expectationValue,
  ev: expectationValue,
  vev: vacuumExpectationValue,
  matrixelement: matrixElement,
  matrixel: matrixElement,
  mel: matrixElement,
  // braket.sty
  Braket: bigBraket,
  Set: bigSet,
  set: smallSet,
};

/** The commands read here (the app leaves them to the kernel). */
export const VISUALTEX_PACKAGE_COMMANDS: readonly string[] = Object.freeze(Object.keys(PACKAGE_COMMANDS));

/**
 * A package command at the cursor (just after its name), as a MacroAtom; null
 * when the command is not one of them or its input is not the package's form
 * (the cursor is then unchanged).
 */
export function scanVisualTexPackageCommand(
  command: string,
  tokens: readonly string[],
  index: number,
  parse: (latex: string) => Atom[],
  style: Style
): { atom: MacroAtom; index: number } | null {
  const name = command.slice(1).replace(/\*+$/, '');
  const definition = Object.prototype.hasOwnProperty.call(PACKAGE_COMMANDS, name) ? PACKAGE_COMMANDS[name] : undefined;
  if (!definition) return null;
  const cursor: Cursor = { tokens, index, stars: command.length - 1 - name.length };
  const args = new Args();
  const body = definition(cursor, args, name);
  // a star the command does not take is not its form
  if (body === null || cursor.stars > 0) return null;
  const atom = new MacroAtom(command, {
    args: cursor.index === index ? null : tokenText(tokens.slice(index, cursor.index)),
    argumentCount: args.sources.length,
    argumentTemplate: args.template,
    argumentSources: args.sources,
    argumentKinds: args.kinds,
    body: parse(body),
    style,
    expand: false,
    captureSelection: args.sources.length === 0,
  });
  return { atom, index: cursor.index };
}

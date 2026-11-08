// VisualTeX: commands of amsmath, mathtools, cancel, stmaryrd, upgreek and
// tensor that MathLive does not define. The layout engine sets each one as
// the package's TeX code does (shared/tex-layout latex-adapter.ts, by
// command); here they parse, keep their source and keep the user's content in
// editable branches.

import { Atom } from '../core/atom-class';
import { GenfracAtom } from '../atoms/genfrac';
import { SpacingAtom } from '../atoms/spacing';
import { SubsupAtom } from '../atoms/subsup';
import type { Parser } from '../core/parser';
import { joinLatex, latexCommand } from '../core/tokenizer';
import { serializeLatexValue } from '../core/registers-utils';
import type { LatexValue, MathstyleName } from '../public/core-types';
import type { ToLatexOptions } from '../core/types';
import { argAtoms, defineFunction, defineSymbols } from './definitions-utils';
import type { Argument } from './types';

// upgreek.sty (Euler option): upright Greek letters
defineSymbols(
  [
    ['\\upalpha', 0x03b1],
    ['\\upbeta', 0x03b2],
    ['\\upgamma', 0x03b3],
    ['\\updelta', 0x03b4],
    ['\\upepsilon', 0x03f5],
    ['\\upvarepsilon', 0x03b5],
    ['\\upzeta', 0x03b6],
    ['\\upeta', 0x03b7],
    ['\\uptheta', 0x03b8],
    ['\\upvartheta', 0x03d1],
    ['\\upiota', 0x03b9],
    ['\\upkappa', 0x03ba],
    ['\\uplambda', 0x03bb],
    ['\\upmu', 0x03bc],
    ['\\upnu', 0x03bd],
    ['\\upxi', 0x03be],
    ['\\uppi', 0x03c0],
    ['\\upvarpi', 0x03d6],
    ['\\uprho', 0x03c1],
    ['\\upvarrho', 0x03f1],
    ['\\upsigma', 0x03c3],
    ['\\upvarsigma', 0x03c2],
    ['\\uptau', 0x03c4],
    ['\\upupsilon', 0x03c5],
    ['\\upphi', 0x03d5],
    ['\\upvarphi', 0x03c6],
    ['\\upchi', 0x03c7],
    ['\\uppsi', 0x03c8],
    ['\\upomega', 0x03c9],
    ['\\Upgamma', 0x0393],
    ['\\Updelta', 0x0394],
    ['\\Uptheta', 0x0398],
    ['\\Uplambda', 0x039b],
    ['\\Upxi', 0x039e],
    ['\\Uppi', 0x03a0],
    ['\\Upsigma', 0x03a3],
    ['\\Upupsilon', 0x03a5],
    ['\\Upphi', 0x03a6],
    ['\\Uppsi', 0x03a8],
    ['\\Upomega', 0x03a9],
  ],
  'mord'
);

// stmaryrd.sty symbols the kernel did not have (the others are in symbols.ts)
defineSymbols([
  ['\\llparenthesis', 0x2987, 'mopen'],
  ['\\rrparenthesis', 0x2988, 'mclose'],
  ['\\varcurlyvee', 0x22ce, 'mbin'],
  ['\\varcurlywedge', 0x22cf, 'mbin'],
  ['\\shortleftarrow', 0x2190, 'mrel'],
  ['\\shortrightarrow', 0x2192, 'mrel'],
  ['\\shortuparrow', 0x2191, 'mrel'],
  ['\\shortdownarrow', 0x2193, 'mrel'],
  ['\\varoplus', 0x2295, 'mbin'],
  ['\\varominus', 0x2296, 'mbin'],
  ['\\varotimes', 0x2297, 'mbin'],
  ['\\varoslash', 0x2298, 'mbin'],
  ['\\varodot', 0x2299, 'mbin'],
  ['\\varocircle', 0x229a, 'mbin'],
  ['\\varobslash', 0x29b8, 'mbin'],
  ['\\obslash', 0x29b8, 'mbin'],
  ['\\obar', 0x233d, 'mbin'],
  ['\\olessthan', 0x29c0, 'mbin'],
  ['\\ogreaterthan', 0x29c1, 'mbin'],
  ['\\boxempty', 0x25a1, 'mbin'],
  ['\\Mapsto', 0x2907, 'mrel'],
  ['\\llfloor', 0x230a, 'mopen'],
  ['\\rrfloor', 0x230b, 'mclose'],
  ['\\llceil', 0x2308, 'mopen'],
  ['\\rrceil', 0x2309, 'mclose'],
  ['\\bigsqcap', 0x2a05, 'mop'],
]);

/** A brace-delimited argument as written (\genfrac's delimiters, thickness and style). */
function rawArgument(parser: Parser): string {
  parser.skipFiller();
  if (!parser.match('<{>')) return parser.end() ? '' : parser.get();
  const s = parser.scanBalancedString();
  parser.match('<}>');
  return s.trim();
}

/** A TeX dimension written as text ("1pt", ".5ex"), or undefined when empty. */
function dimensionOf(s: string): { dimension: number; unit: any } | undefined {
  const m = /^\s*([-+]?(?:\d+\.?\d*|\.\d+))\s*(pt|mm|cm|ex|em|bp|dd|pc|in|mu|px|sp)\s*$/.exec(s);
  if (!m) return undefined;
  return { dimension: Number.parseFloat(m[1]), unit: m[2] as any };
}

// amsmath \genfrac{left}{right}{thickness}{style}{num}{den}
defineFunction('genfrac', '', {
  ifMode: 'math',
  parse: (parser: Parser): Argument[] => {
    const raw = [0, 1, 2, 3].map(() => rawArgument(parser));
    return [...raw, parser.scanArgument('expression'), parser.scanArgument('expression')] as Argument[];
  },
  createAtom: (options) => {
    const args = options.args! as (string | Argument | null)[];
    const [left, right, thickness, style] = args.slice(0, 4) as string[];
    const styles: Record<string, MathstyleName> = { '0': 'displaystyle', '1': 'textstyle', '2': 'scriptstyle', '3': 'scriptscriptstyle' };
    const barThickness = dimensionOf(thickness ?? '');
    return new GenfracAtom(argAtoms(args[4] as Argument), argAtoms(args[5] as Argument), {
      ...options,
      leftDelim: left || undefined,
      rightDelim: right || undefined,
      hasBarLine: !(barThickness && barThickness.dimension === 0),
      barThickness,
      mathstyleName: styles[(style ?? '').trim()],
    } as any);
  },
  serialize: (atom: GenfracAtom, options) => {
    const args = (atom.args ?? []) as unknown as string[];
    return `\\genfrac{${args[0] ?? ''}}{${args[1] ?? ''}}{${args[2] ?? ''}}{${args[3] ?? ''}}{${atom.aboveToLatex(options)}}{${atom.belowToLatex(options)}}`;
  },
});

// TeX's \above <dimen> and \abovewithdelims <delim><delim><dimen>
defineFunction('above', '{numer:auto}{denom:auto}{thickness:value}', {
  infix: true,
  createAtom: (options) => {
    const args = options.args! as Argument[];
    return new GenfracAtom(argAtoms(args[0]), argAtoms(args[1]), {
      ...options,
      hasBarLine: true,
      barThickness: (args[2] as LatexValue) ?? { dimension: 0 },
    } as any);
  },
  serialize: (atom: GenfracAtom, options) =>
    `${atom.aboveToLatex(options)} \\above ${serializeLatexValue(atom.barThickness ?? { dimension: 0 })} ${atom.belowToLatex(options)}`,
});

defineFunction('abovewithdelims', '{numer:auto}{denom:auto}{left-delim:delim}{right-delim:delim}{thickness:value}', {
  infix: true,
  createAtom: (options) => {
    const args = options.args! as Argument[];
    return new GenfracAtom(argAtoms(args[0]), argAtoms(args[1]), {
      ...options,
      leftDelim: (args[2] as string) ?? '.',
      rightDelim: (args[3] as string) ?? '.',
      hasBarLine: true,
      barThickness: (args[4] as LatexValue) ?? { dimension: 0 },
    } as any);
  },
  serialize: (atom: GenfracAtom, options) =>
    `${atom.aboveToLatex(options)} \\abovewithdelims${atom.leftDelim}${atom.rightDelim} ${serializeLatexValue(
      atom.barThickness ?? { dimension: 0 }
    )} ${atom.belowToLatex(options)}`,
});

/** `\cmd[a][b]{body}` and the scripts the atom carries (optional arguments kept as written). */
function withOptions(atom: Atom, options: ToLatexOptions, defaultMode: 'math' | 'text' = 'math'): string {
  const opts = ((atom.args ?? []) as (LatexValue | string | null)[])
    .slice(0, 2)
    .map((v) => (v === null || v === undefined ? null : typeof v === 'string' ? v : serializeLatexValue(v) ?? ''));
  let result = atom.command;
  if (opts[0] !== null && opts[0] !== undefined) result += `[${opts[0]}]`;
  if (opts[1] !== null && opts[1] !== undefined) result += `[${opts[1]}]`;
  return joinLatex([
    latexCommand(result, atom.bodyToLatex({ ...options, defaultMode })),
    atom.supsubToLatex(options),
  ]);
}

// mathtools \overbracket[rule thickness][bracket height]{arg}, \underbracket
defineFunction(['overbracket', 'underbracket'], '[:value][:value]{:auto}', {
  ifMode: 'math',
  createAtom: (options) =>
    new Atom({
      ...options,
      type: 'mord',
      body: argAtoms(options.args![2] as Argument),
      limits: 'over-under',
    }),
  serialize: (atom, options) => withOptions(atom, options),
});

// amsmath \leftroot{n} and \uproot{n}, inside the index of \sqrt[...]{...}
defineFunction(['leftroot', 'uproot'], '{:value}', {
  ifMode: 'math',
  createAtom: (options) => new SpacingAtom({ ...options }),
  serialize: (atom) => `${atom.command}{${serializeLatexValue(atom.args?.[0] as LatexValue) ?? '0'}}`,
});

// \nobreak (\penalty10000): nothing visible
defineFunction('nobreak', '', {
  createAtom: (options) => new SpacingAtom({ ...options }),
});

// cancel.sty \cancelto{value}{expression}
defineFunction('cancelto', '{:auto}{:auto}', {
  ifMode: 'math',
  createAtom: (options) => {
    const atom = new Atom({ ...options, type: 'mord', body: argAtoms(options.args![1] as Argument) });
    atom.above = argAtoms(options.args![0] as Argument);
    return atom;
  },
  serialize: (atom, options) =>
    joinLatex([latexCommand('\\cancelto', atom.aboveToLatex(options), atom.bodyToLatex(options)), atom.supsubToLatex(options)]),
});

// LaTeX's \fbox{...} and \hbox{...} (text-mode content, an \hbox)
defineFunction(['fbox', 'hbox'], '{:text}', {
  ifMode: 'math',
  createAtom: (options) => new Atom({ ...options, type: 'mord', body: argAtoms(options.args![0] as Argument), mode: 'math' }),
  serialize: (atom, options) => latexCommand(atom.command, atom.bodyToLatex({ ...options, defaultMode: 'text' })),
});

// \framebox[width][position]{...}
defineFunction('framebox', '[:value][:string]{:text}', {
  ifMode: 'math',
  createAtom: (options) => new Atom({ ...options, type: 'mord', body: argAtoms(options.args![2] as Argument), mode: 'math' }),
  serialize: (atom, options) => withOptions(atom, options, 'text'),
});

/**
 * tensor.sty's index string `{^a_b^c}`: one subsup atom per index, in the
 * order written (the order sets the staggering).
 */
function scanIndices(parser: Parser): Atom[] | null {
  parser.skipFiller();
  if (!parser.match('<{>')) return null;
  const items: Atom[] = [];
  while (!parser.end() && !parser.match('<}>')) {
    parser.skipFiller();
    if (parser.match('<}>')) break;
    const branch = parser.match('^') ? 'superscript' : parser.match('_') ? 'subscript' : null;
    if (!branch) {
      // tensor.sty: misordered index tokens are dropped
      parser.get();
      continue;
    }
    const item = new SubsupAtom({ style: parser.style });
    item.setChildren(argAtoms(parser.scanArgument('math')), branch);
    items.push(item);
  }
  return items;
}

// tensor.sty \indices{^a_b} (scripts of the atom before it), \tensor[^a_b]{T}{^c_d}
defineFunction(['indices', 'indices*'], '', {
  ifMode: 'math',
  parse: (parser: Parser): Argument[] => [scanIndices(parser) ?? []],
  createAtom: (options) => new Atom({ ...options, type: 'mord', body: options.args![0] as Atom[] }),
  serialize: (atom, options) => `${atom.command}{${atom.bodyToLatex(options)}}`,
});

defineFunction(['tensor', 'tensor*'], '', {
  ifMode: 'math',
  parse: (parser: Parser): Argument[] => {
    parser.skipFiller();
    let pre: Atom[] | null = null;
    if (parser.match('[')) {
      // the prescripts: an index string up to the closing bracket
      const items: Atom[] = [];
      while (!parser.end() && !parser.match(']')) {
        parser.skipFiller();
        if (parser.match(']')) break;
        const branch = parser.match('^') ? 'superscript' : parser.match('_') ? 'subscript' : null;
        if (!branch) {
          parser.get();
          continue;
        }
        const item = new SubsupAtom({ style: parser.style });
        item.setChildren(argAtoms(parser.scanArgument('math')), branch);
        items.push(item);
      }
      pre = items;
    }
    const nucleus = parser.scanArgument('math');
    return [pre, nucleus, scanIndices(parser) ?? []] as Argument[];
  },
  createAtom: (options) => {
    const args = options.args! as (Atom[] | Argument | null)[];
    const atom = new Atom({ ...options, type: 'mord', body: argAtoms(args[1] as Argument) });
    if (args[0]) atom.above = args[0] as Atom[];
    atom.below = (args[2] as Atom[]) ?? [];
    return atom;
  },
  serialize: (atom, options) => {
    const pre = atom.branch('above') !== undefined ? `[${atom.aboveToLatex(options)}]` : '';
    return `${atom.command}${pre}{${atom.bodyToLatex(options)}}{${atom.belowToLatex(options)}}`;
  },
});

// Internal to \ce (latex-commands/mhchem.ts writes them; a chem atom keeps its
// own source): mhchem.sty v4's arrows (\__mhchem_arrow_output:nnn), the
// prescripts' \llap in the current style and \mhchem@minispace.
defineFunction('vtxcearrow', '{:string}{:math}{:math}', {
  ifMode: 'math',
  createAtom: (options) => {
    const atom = new Atom({ ...options, type: 'mrel' });
    atom.above = argAtoms(options.args![1] as Argument);
    atom.below = argAtoms(options.args![2] as Argument);
    return atom;
  },
  serialize: (atom, options) =>
    `\\vtxcearrow{${atom.args?.[0] ?? ''}}{${atom.aboveToLatex(options)}}{${atom.belowToLatex(options)}}`,
});

defineFunction('vtxcellap', '{:math}', {
  ifMode: 'math',
  createAtom: (options) => new Atom({ ...options, type: 'mord', body: argAtoms(options.args![0] as Argument) }),
});

defineFunction('vtxceminispace', '', {
  ifMode: 'math',
  createAtom: (options) => new SpacingAtom({ ...options }),
});

import type { Dimension, Environment } from '../public/core-types';

import { Atom } from '../core/atom-class';
import { ArrayAtom, ColumnFormat } from '../atoms/array';

import {
  argAtoms,
  defineFunction,
  defineRootEnvironment,
  defineTabularEnvironment,
} from './definitions-utils';
import type { Argument } from './types';
import { Parser, parseLatex } from '../core/parser';
import { latexCommand } from '../core/tokenizer';
import { MacroAtom, MacroArgumentAtom } from '../atoms/macro';
import { VISUALTEX_PHYSICS_MATRIX_WRAPPERS } from './visualtex-physics-matrices';

/*
 The star at the end of the name of an environment indicates that the
 equations will not be numbered and the counter for the equations
 will not be incremented.

\notag will also turn off the numbering (but the counter will still be incremented).
\shoveright and \shoveleft will force alignment of a line

The only difference between `align` and `equation` is the spacing of the formulas.
You should attempt to use `equation` when possible, and `align` when you have multi-line formulas.
`equation` will have space before/after < 1em if line before/after is short enough.

Also: `equation` throws an error when you have an & inside the environment,
so look out for that when converting between the two.

Whereas `align` produces a structure whose width is the full line width
(it's a root environment), `aligned` gives a width that is the actual width
of the contents, thus it can be used as a component in a containing
expression, e.g. for putting the entire alignment in a parenthesis
*/

/** `math` and `displaymath` are older environments, they
 * are equivalent to `\(...\)` and `\[...\]`
 */
defineRootEnvironment(['math', 'displaymath'], makeEnvironment);

/** The `center` environment is a text environment, but
 * we extend it to support math content as well.
 */
defineRootEnvironment('center', makeEnvironment);

/** Note spelling: MULTLINE, not multiline.
 * The `multline` environment is used for long formulas that don't fit on a single line.
 * The first line is left-aligned, the last line is right-aligned, and all lines in between
 * are centered. The `multline*` variant does not number the equation.
 */
defineTabularEnvironment(['multline', 'multline*'], '', makeEnvironment);

/** This is not exactly an environment, but a function that behaves like one: its content is interpreted as multiple rows/lines.
 * The environment `lines` is identical using `\displaylines`.
 */
defineFunction('displaylines', '', {
  parse: (parser: Parser) => {
    parser.skipWhitespace();
    if (!parser.match('<{>')) return [];

    const lines: Atom[][][] = [];
    let line: Atom[] = [];

    // Collect any elements left of the command
    if (parser.mathlist.length > 0) lines.push([parser.mathlist]);

    parser.beginContext({ tabular: true, root: true });
    do {
      if (parser.end() || parser.match('<}>')) break;
      if (parser.matchColumnSeparator() || parser.matchRowSeparator()) {
        lines.push([line]);
        line = [];
      } else {
        line.push(
          ...parser.scan((token) =>
            ['<}>', '&', '\\cr', '\\\\', '\\tabularnewline'].includes(token)
          )
        );
      }
    } while (true);
    parser.endContext();
    lines.push([line]);

    // Append any elements right of the command
    const rhs = parser.scan();
    if (rhs.length > 0) lines.push([rhs]);

    return lines as Argument[];
  },
  createAtom: (options) =>
    new ArrayAtom('lines', options.args as (readonly Atom[])[][], [], {
      // arraystretch: 1.2,
      leftDelim: '.',
      rightDelim: '.',
      columns: [{ align: 'l' }],
      classes: ['ML__multiline_environment'],
      isRoot: true,
      minColumns: 1,
      maxColumns: 1,
      minRows: 1,
    }),
});

/** The `split` environment is used to split a long formula over multiple lines.
 * It is used inside another environment, typically `equation` or `align`.
 * The first column is right-aligned, the second column is left-aligned.
 * The entire construct is numbered as a single equation.
 * The `split*` variant does not number the equation.
```latex
\begin{equation}
\begin{split}
    f(x) &= a + b + c + d \\
         &\quad + e + f + g
\end{split}
\end{equation}
```
 */
defineTabularEnvironment('split', '', makeEnvironment);

// An AMS-Math environment
// %    The \env{gathered} environment is for several lines that are
// %    centered independently.
// From amstex.sty
// \newenvironment{gathered}[1][c]{%
//   \relax\ifmmode\else\nonmatherr@{\begin{gathered}}\fi
//   \null\,%
//   \if #1t\vtop \else \if#1b\vbox \else \vcenter \fi\fi
//   \bgroup\Let@\restore@math@cr
//   \ifinany@\else\openup\jot\fi\ialign
//   \bgroup\hfil\strut@$\m@th\displaystyle##$\hfil\crcr
/** This is the equivalent of the `center` environment, but for math content. `gather` and `gather*` are root environments, `gathered` is a sub-environment */
defineTabularEnvironment(
  ['gather', 'gather*', 'gathered'],
  '',
  makeEnvironment
);

defineTabularEnvironment(
  ['equation', 'equation*', 'subequations'],
  '',
  makeEnvironment
);

// An AMS-Math environment
// See amsmath.dtx:3565
// Note that some versions of AMS-Math have a gap on the left.
// More recent version suppresses that gap, but have an option to turn it
// back on for backward compatibility.
// Note that technically, 'eqnarray' behaves (slightly) differently. However,
// is is generally recommended to avoid using eqnarray and use align instead.
// https://texblog.net/latex-archive/maths/eqnarray-align-environment/
defineRootEnvironment(['align', 'align*', 'eqnarray'], makeEnvironment, {
  tabular: true,
});

defineTabularEnvironment(['aligned'], '', makeEnvironment);

// VisualTeX: amsmath's other display alignments. alignat's argument is its
// number of column pairs (kept for saving); flalign spreads its pairs over the
// line; eqnarray* is eqnarray unnumbered.
defineRootEnvironment(['eqnarray*', 'flalign', 'flalign*'], makeEnvironment, {
  tabular: true,
});
defineRootEnvironment(['alignat', 'alignat*'], makeEnvironment, {
  tabular: true,
  params: '{:string}',
});

// VisualTeX: aligned / gathered / alignedat take amsmath's vertical position
// [t] / [b] / [c] (the parser reads only those: any other bracket is content,
// as amsmath's \ams@start@box puts it back)
defineTabularEnvironment(['aligned', 'gathered'], '[:string]', makeEnvironment);
defineTabularEnvironment('alignedat', '[:string]{:string}', makeEnvironment);

// DefineEnvironment('alignat', '', function(name, args) {
//     return {

//     };
// });

// defineEnvironment('flalign', '', function(name, args) {
//     return {

//     };
// });

// DefineEnvironment('cardinality', '',  function() {
//     const result = {};

//     result.mathstyle = 'textstyle';
//     result.lFence = '|';
//     result.rFence = '|';

//     return result;
// });

defineTabularEnvironment(
  'array',
  '{columns:colspec}',
  (name, array, rowGaps, args) => {
    return new ArrayAtom(name, array, rowGaps, {
      columns: args[0] as ColumnFormat[],
      mathstyleName: 'textstyle',
    });
  }
);

defineTabularEnvironment(
  [
    'matrix',
    'pmatrix',
    'bmatrix',
    'Bmatrix',
    'vmatrix',
    'Vmatrix',
    'matrix*',
    'pmatrix*',
    'bmatrix*',
    'Bmatrix*',
    'vmatrix*',
    'Vmatrix*',
  ],
  '[columns:colspec]',
  makeEnvironment
);

defineTabularEnvironment(
  ['smallmatrix', 'smallmatrix*'],
  '[columns:colspec]',
  makeEnvironment
);

// physics.sty matrix shorthands need a real tabular parse. A regular MathLive
// macro expands too late: its argument ampersands become literal \& atoms and
// source serialization loses the matrix column separators.
defineFunction(Object.keys(VISUALTEX_PHYSICS_MATRIX_WRAPPERS), '', {
  ifMode: 'math',
  parse: (parser) => {
    parser.skipWhitespace();
    if (parser.peek() !== '<{>') {
      parser.onError({ code: 'missing-argument' });
      return [''];
    }
    return [parser.scanLiteralGroup()];
  },
  createAtom: (options) => {
    const [open, close] = VISUALTEX_PHYSICS_MATRIX_WRAPPERS[options.command!.slice(1)];
    const argument = (options.args![0] as string) ?? '';
    const atom = new MacroAtom(options.command!, {
      args: `{${argument}}`,
      body: parseLatex(open + argument + close),
      style: options.style!,
      expand: false,
      captureSelection: false,
    });
    atom._serialize = (serializeOptions) => {
      const expanded = atom.bodyToLatex(serializeOptions);
      if (!expanded.startsWith(open) || !expanded.endsWith(close)) return expanded;
      return options.command + '{' + expanded.slice(open.length, -close.length) + '}';
    };
    return atom;
  },
});

// A tracked source argument must contain the real array, not a flat A\\B
// string, so vertical navigation can traverse \\substack rows.
defineFunction('substack', '', {
  ifMode: 'math',
  parse: (parser) => {
    parser.skipWhitespace();
    if (parser.peek() !== '<{>') {
      parser.onError({ code: 'missing-argument' });
      return [''];
    }
    return [parser.scanLiteralGroup()];
  },
  createAtom: (options) => {
    const argument = (options.args![0] as string) ?? '';
    const open = '\\begin{array}{c}';
    const close = '\\end{array}';
    const tracked = new MacroArgumentAtom(
      1,
      parseLatex(open + argument + close),
      { style: options.style }
    );
    const atom = new MacroAtom(options.command!, {
      args: `{${argument}}`,
      argumentCount: 1,
      body: [tracked],
      style: options.style!,
      expand: false,
      captureSelection: false,
    });
    atom._serialize = (serializeOptions) => {
      const expanded = tracked.bodyToLatex(serializeOptions);
      if (!expanded.startsWith(open) || !expanded.endsWith(close)) return expanded;
      return options.command + '{' + expanded.slice(open.length, -close.length) + '}';
    };
    return atom;
  },
});

// \qq and \qqtext have text arguments. A generic macro reparses their text
// wrapper into the editable argument and rewrites \qq{if} as \qq{\text{if}}.
defineFunction(['qq', 'qqtext'], '', {
  ifMode: 'math',
  parse: (parser) => {
    parser.skipWhitespace();
    if (parser.peek() !== '<{>') {
      parser.onError({ code: 'missing-argument' });
      return [''];
    }
    return [parser.scanLiteralGroup()];
  },
  createAtom: (options) => {
    const argument = (options.args![0] as string) ?? '';
    const open = '\\quad\\text{';
    const close = '}\\quad';
    const atom = new MacroAtom(options.command!, {
      args: `{${argument}}`,
      body: parseLatex(open + argument + close),
      style: options.style!,
      expand: false,
      captureSelection: false,
    });
    atom._serialize = (serializeOptions) => {
      const expanded = atom.bodyToLatex(serializeOptions);
      if (!expanded.startsWith(open) || !expanded.endsWith(close)) return expanded;
      return options.command + '{' + expanded.slice(open.length, -close.length) + '}';
    };
    return atom;
  },
});

// \cases is standard LaTeX
// \dcases is from the mathtools package
// \rcases is from the mathtools package
// From amstex.sty:
// \def\cases{\left\{\def\arraystretch{1.2}\hskip-\arraycolsep
//   \array{l@{\quad}l}}
// \def\endcases{\endarray\hskip-\arraycolsep\right.}
// From amsmath.dtx
// \def\env@cases{%
//   \let\@ifnextchar\new@ifnextchar
//   \left\lbrace
//   \def\arraystretch{1.2}%
//   \array{@{}l@{\quad}l@{}}%
defineTabularEnvironment(['cases', 'dcases', 'rcases'], '', makeEnvironment);

// VisualTeX: mathtools' other cases (the starred ones set their second column
// as text) and amsmath's subarray (\substack's environment, {l} or {c})
defineTabularEnvironment(
  ['cases*', 'dcases*', 'rcases*', 'drcases', 'drcases*'],
  '',
  makeEnvironment
);
defineTabularEnvironment('subarray', '{columns:colspec}', makeEnvironment);

// VisualTeX: mathtools' starred matrices take one alignment (l, c or r) for
// every column, not a column specification
defineTabularEnvironment(
  ['matrix*', 'pmatrix*', 'bmatrix*', 'Bmatrix*', 'vmatrix*', 'Vmatrix*'],
  '[:string]',
  makeEnvironment
);

// This is a text mode environment
/*
\begin{theorem}
Let $f$ be a function whose derivative exists in every point, then $f$
is a continuous function.
\end{theorem}
*/
// defineEnvironment('theorem', '', function () {
//     return {};
// });

export function makeEnvironment(
  name: Environment,
  content: (readonly Atom[])[][] = [[[]]],
  rowGaps: readonly Dimension[] = [],
  args: readonly (null | Argument)[] = [],
  maxMatrixCols?: number
): ArrayAtom {
  const atom = makeEnvironmentAtom(name, content, rowGaps, args, maxMatrixCols);
  // VisualTeX: the arguments read after \begin{name}, as written
  const written = environmentArguments(name, args);
  if (written) atom.environmentArguments = written;
  return atom;
}

/** VisualTeX: the LaTeX of the arguments the parser read for these environments. */
function environmentArguments(
  name: Environment,
  args: readonly (null | Argument)[]
): string {
  const text = (x: null | Argument | undefined) => (typeof x === 'string' ? x.trim() : '');
  switch (name) {
    case 'aligned':
    case 'gathered':
    case 'matrix*':
    case 'pmatrix*':
    case 'bmatrix*':
    case 'Bmatrix*':
    case 'vmatrix*':
    case 'Vmatrix*':
      return text(args[0]) ? `[${text(args[0])}]` : '';
    case 'alignedat':
      return (text(args[0]) ? `[${text(args[0])}]` : '') + `{${text(args[1])}}`;
    case 'alignat':
    case 'alignat*':
      return `{${text(args[0])}}`;
  }
  return '';
}

function makeEnvironmentAtom(
  name: Environment,
  content: (readonly Atom[])[][],
  rowGaps: readonly Dimension[],
  args: readonly (null | Argument)[],
  maxMatrixCols?: number
): ArrayAtom {
  switch (name) {
    case 'math':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        isRoot: true,
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
        maxRows: 1,
      });

    case 'displaymath':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        isRoot: true,
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
        maxRows: 1,
      });

    case 'center':
      return new ArrayAtom(name, content, rowGaps, {
        columns: [{ align: 'c' }],
        classes: ['ML__center_environment'],
        isRoot: true,
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
        maxRows: 1,
      });

    case 'multline':
    case 'multline*':
      return new ArrayAtom(name, content, rowGaps, {
        columns: [{ align: 'm' }],
        leftDelim: '.',
        rightDelim: '.',
        isRoot: true,
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
      });

    case 'split':
      return new ArrayAtom(name, content, rowGaps, {
        columns: [{ align: 'r' }, { align: 'l' }],
        minColumns: 2,
        minRows: 1,
        isRoot: false,
      });

    case 'gather':
    case 'gather*':
      return new ArrayAtom(name, content, rowGaps, {
        columns: [{ gap: 0.25 }, { align: 'c' }, { gap: 0 }],
        // colSeparationType: 'gather',
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
        isRoot: true,
        classes: ['ML__gather_environment'],
      });

    case 'gathered':
      return new ArrayAtom(name, content, rowGaps, {
        columns: [{ gap: 0.25 }, { align: 'c' }, { gap: 0 }],
        // colSeparationType: 'gather',
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
      });

    case 'equation':
    case 'equation*':
    case 'subequations':
      return new ArrayAtom(name, content, rowGaps, {
        columns: [{ align: 'c' }],
        isRoot: true,
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
      });

    case 'subarray':
      return new ArrayAtom(name, content, rowGaps, {
        columns: (args[0] as ColumnFormat[] | null) ?? [{ align: 'c' }],
        mathstyleName: 'scriptstyle',
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
      });

    // VisualTeX: amsmath's alignment environments keep every column pair
    // (`x&=1&y&=2`), as aligned does
    case 'align':
    case 'align*':
    case 'flalign':
    case 'flalign*':
    case 'alignat':
    case 'alignat*':
    case 'alignedat':
    case 'aligned': {
      let colCount = 0;
      for (const row of content) colCount = Math.max(colCount, row.length);

      const columns: ColumnFormat[] = [
        { gap: 0 },
        { align: 'r' },
        { gap: 0.25 },
        { align: 'l' },
      ];
      let i = 2;
      while (i < colCount) {
        columns.push({ gap: 1 }, { align: 'r' }, { gap: 0.25 }, { align: 'l' });
        i += 2;
      }

      columns.push({ gap: 0 });

      return new ArrayAtom(name, content, rowGaps, {
        arraycolsep: 0,
        columns,
        // colSeparationType: 'align',
        minColumns: 2,
        minRows: 1,
        isRoot: name !== 'aligned' && name !== 'alignedat',
        classes: name.startsWith('align') && name !== 'aligned' && name !== 'alignedat' ? ['ML__align_environment'] : [],
      });
    }

    case 'eqnarray':
    case 'eqnarray*':
      return new ArrayAtom(name, content, rowGaps, {
        arraycolsep: 0,
        columns: [{ gap: 0 }, { align: 'r' }, { align: 'c' }, { align: 'l' }],
        minColumns: 3,
        maxColumns: 3,
        minRows: 1,
        isRoot: true,
        classes: ['ML__eqnarray_environment'],
      });

    case 'pmatrix':
    case 'pmatrix*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        leftDelim: '(',
        rightDelim: ')',
        columns: defaultColumns(args[0], maxMatrixCols),
      });

    case 'bmatrix':
    case 'bmatrix*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        leftDelim: '[',
        rightDelim: ']',
        columns: defaultColumns(args[0], maxMatrixCols),
      });

    case 'Bmatrix':
    case 'Bmatrix*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        leftDelim: '\\lbrace',
        rightDelim: '\\rbrace',
        columns: defaultColumns(args[0], maxMatrixCols),
      });

    case 'vmatrix':
    case 'vmatrix*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        leftDelim: '\\vert',
        rightDelim: '\\vert',
        columns: defaultColumns(args[0], maxMatrixCols),
      });

    case 'Vmatrix':
    case 'Vmatrix*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        leftDelim: '\\Vert',
        rightDelim: '\\Vert',
        columns: defaultColumns(args[0], maxMatrixCols),
      });

    case 'matrix':
    case 'matrix*':
      // Specifying a fence, even a null fence,
      // will prevent the insertion of an initial and final gap
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'textstyle',
        leftDelim: '.',
        rightDelim: '.',
        columns: defaultColumns(args?.[0], maxMatrixCols),
      });

    case 'smallmatrix':
    case 'smallmatrix*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: 'scriptstyle',
        columns: defaultColumns(args?.[0], maxMatrixCols),
        colSeparationType: 'small',
        arraystretch: 0.5,
      });

    case 'cases':
    case 'dcases':
    case 'cases*':
    case 'dcases*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: name.startsWith('dcases') ? 'displaystyle' : 'textstyle',
        arraystretch: 1.2,
        leftDelim: '\\lbrace',
        rightDelim: '.',
        columns: casesColumns(),
      });

    case 'rcases':
    case 'rcases*':
    case 'drcases':
    case 'drcases*':
      return new ArrayAtom(name, content, rowGaps, {
        mathstyleName: name.startsWith('d') ? 'displaystyle' : 'textstyle',
        arraystretch: 1.2,
        leftDelim: '.',
        rightDelim: '\\rbrace',
        columns: casesColumns(),
      });

    case 'lines':
      return new ArrayAtom(name, content, rowGaps, {
        // arraystretch: 1.2,
        leftDelim: '.',
        rightDelim: '.',
        columns: [{ align: 'l' }],
        isRoot: true,
        minColumns: 1,
        maxColumns: 1,
        minRows: 1,
      });
  }

  // 'math'
  return new ArrayAtom(name, content, rowGaps, {
    mathstyleName: 'textstyle',
  });
}

function defaultColumns(
  args: null | Argument,
  maxMatrixCols: number = 10
): ColumnFormat[] {
  // VisualTeX: mathtools' matrix*[r]: one alignment for every column
  if (typeof args === 'string') {
    const align = args.trim();
    return Array(maxMatrixCols).fill({ align: align === 'l' || align === 'r' ? align : 'c' });
  }
  return (args as ColumnFormat[]) ?? Array(maxMatrixCols).fill({ align: 'c' });
}

function casesColumns(maxCasesColumns: number = 10): ColumnFormat[] {
  const columns: ColumnFormat[] = [];
  for (let i = 0; i < maxCasesColumns; i++) {
    if (i > 0) columns.push({ gap: 1 });
    columns.push({ align: 'l' });
  }
  return columns;
}

// VisualTeX: amsmath's equation tags and labels. \tag{x} draws "(x)" and
// \tag*{x} draws "x" at the right end of its line (the layout engine places
// them); its argument is text. \label, \notag and \nonumber draw nothing and
// are kept for saving.
defineFunction(['tag', 'tag*'], '{:text}', {
  ifMode: 'math',
  createAtom: (options) =>
    new Atom({
      ...options,
      type: 'mord',
      body: argAtoms(options.args![0]),
      mode: 'math',
    }),
  serialize: (atom, options) =>
    latexCommand(atom.command, atom.bodyToLatex({ ...options, defaultMode: 'text' })),
});

defineFunction('label', '{:string}', {
  ifMode: 'math',
  createAtom: (options) =>
    new Atom({ ...options, type: 'mord', mode: 'math', captureSelection: true }),
  serialize: (atom) => latexCommand('\\label', (atom.args?.[0] as string) ?? ''),
});

defineFunction(['notag', 'nonumber'], '', {
  ifMode: 'math',
  createAtom: (options) =>
    new Atom({ ...options, type: 'mord', mode: 'math', captureSelection: true }),
  serialize: (atom) => atom.command,
});

// VisualTeX: amsmath's \intertext and mathtools' \shortintertext: a line of
// text between the rows of an alignment (the layout engine sets it)
defineFunction(['intertext', 'shortintertext'], '{:text}', {
  ifMode: 'math',
  createAtom: (options) =>
    new Atom({
      ...options,
      type: 'mord',
      body: argAtoms(options.args![0]),
      mode: 'math',
    }),
  serialize: (atom, options) =>
    latexCommand(atom.command, atom.bodyToLatex({ ...options, defaultMode: 'text' })),
});

// VisualTeX: LaTeX's tabular, as in \text{\begin{tabular}{|l|r|}...}: its
// cells are text (the parser reads them in text mode)
defineTabularEnvironment(
  'tabular',
  '{columns:colspec}',
  (name, array, rowGaps, args) =>
    new ArrayAtom(name, array, rowGaps, {
      columns: args[0] as ColumnFormat[],
      mode: 'text',
    })
);

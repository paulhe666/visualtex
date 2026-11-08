import type {
  MathstyleName,
  Dimension,
  Environment,
  ParseMode,
} from '../public/core-types';

import { Atom, isCellBranch, isNamedBranch } from '../core/atom-class';
import { joinLatex } from '../core/tokenizer';

import { PlaceholderAtom } from './placeholder';
import type { AtomJson, Branch, ToLatexOptions } from '../core/types';

export type ColumnFormat =
  | {
      // A regular content column, with the specified alignment.
      // 'm' is a special alignment for multline: left on first row, right on last
      // row, centered otherwise
      align: 'l' | 'c' | 'r' | 'm';
    }
  | {
      // The width of a gap between columns, or a LaTeX expression between columns
      gap: number | readonly Atom[];
    }
  | {
      // A rule (line) separating columns
      separator: 'solid' | 'dashed';
    };

export type ColSeparationType =
  // | 'align'
  // | 'alignat'
  // | 'gather'
  | 'small'
  // | 'CD'
  | undefined;

export type ArrayAtomConstructorOptions = {
  isRoot?: boolean;

  mathstyleName?: MathstyleName;

  colSeparationType?: ColSeparationType;
  leftDelim?: string;
  rightDelim?: string;
  // A multiplication factor applied to the spacing between rows and columns
  arraystretch?: number;
  // The spacing between columns
  arraycolsep?: number;

  columns?: ColumnFormat[];
  minColumns?: number;
  maxColumns?: number;
  minRows?: number;
  maxRows?: number;

  displayEquationNumber?: boolean;

  // VisualTeX: the \hline rules after the last row (`\\ \hline\end{array}`),
  // as JSON of the starting commit had them (now in rowRules)
  bottomRules?: number;

  // VisualTeX: the environment's arguments as written after \begin{name}
  // (`{2}` of alignat, `[t]` of aligned, `[r]` of pmatrix*), kept for saving
  environmentArguments?: string;

  // VisualTeX: the rules before each row and, last, after the last row
  // (\hline, \cline{2-3}, \toprule, ...), as written
  rowRules?: readonly (readonly string[])[];

  // VisualTeX: the column specification as written (`|p{2cm}|c|`)
  colSpec?: string;

  // VisualTeX: the \multicolumn, \multirow and \hdotsfor around cells, by
  // [row, column] (JSON; the parser gives TableMarkAtoms in the cells)
  cellWrappers?: readonly [number, number, readonly CellWrapper[]][];

  mode?: ParseMode;

  classes?: string[];
};

/**
 * VisualTeX: what a cell's content is wrapped in, outermost first, as written:
 * \multicolumn{columns}{spec}{…}, multirow.sty's
 * \multirow[vpos]{rows}[bigstruts]{width}[vmove]{…} (its content is text) and
 * amsmath's \hdotsfor[spacing]{columns} (dots over the columns; no content).
 */
export type CellWrapper =
  | { command: '\\multicolumn'; columns: number; spec: string }
  | {
      command: '\\multirow';
      rows: string;
      width: string;
      vpos?: string;
      bigstruts?: string;
      vmove?: string;
    }
  | { command: '\\hdotsfor'; columns: number; spacing?: string };

/** VisualTeX: a wrapper's LaTeX before its content argument */
export function cellWrapperLatex(w: CellWrapper): string {
  const opt = (x: string | undefined) => (x === undefined ? '' : `[${x}]`);
  if (w.command === '\\multicolumn')
    return `\\multicolumn{${w.columns}}{${w.spec}}`;
  if (w.command === '\\hdotsfor')
    return `\\hdotsfor${opt(w.spacing)}{${w.columns}}`;
  return `\\multirow${opt(w.vpos)}{${w.rows}}${opt(w.bigstruts)}{${w.width}}${opt(w.vmove)}`;
}

/** VisualTeX: the columns a cell spans (\multicolumn, \hdotsfor) */
export function cellWrapperSpan(wrappers: readonly CellWrapper[] | undefined): number {
  let span = 1;
  for (const w of wrappers ?? [])
    if (w.command !== '\\multirow') span = Math.max(1, w.columns);
  return span;
}

/**
 * VisualTeX: a rule at a row start (\hline, \cline{2-3}, \toprule, ...) or a
 * wrapper at a cell start (\multicolumn{3}{c}, ...), as the parser meets it.
 * The ArrayAtom takes these out of its cells: rules go to rowRules, wrappers
 * to the cell (CELL_WRAPPERS), so the caret never stands beside them.
 */
export class TableMarkAtom extends Atom {
  readonly rule?: string;
  readonly wrapper?: CellWrapper;
  constructor(mark: { rule: string } | { wrapper: CellWrapper }) {
    const latex = 'rule' in mark ? mark.rule : cellWrapperLatex(mark.wrapper);
    super({
      type: 'error',
      command: /^\\[a-zA-Z]+/.exec(latex)?.[0] ?? latex,
      verbatimLatex: latex,
    });
    if ('rule' in mark) this.rule = mark.rule;
    else this.wrapper = mark.wrapper;
  }
}

/** VisualTeX: the wrappers of a cell, by the cell's `first` atom */
const CELL_WRAPPERS = new WeakMap<Atom, readonly CellWrapper[]>();

function wrappersOf(cell: readonly Atom[] | undefined): readonly CellWrapper[] | undefined {
  return cell?.[0] ? CELL_WRAPPERS.get(cell[0]) : undefined;
}

function rowWidth(row: readonly (readonly Atom[] | undefined)[]): number {
  let width = 0;
  for (const cell of row) if (cell) width += cellWrapperSpan(wrappersOf(cell));
  return width;
}

/**
 * Normalize cells:
 * - ensure the array of cells is dense (not sparse)
 * - fold rows that overflow (longer than maximum number of columns)
 * - ensure each cell begins with a `first` atom
 * - ensure the minimum/maximum number of cells is respected
 * - remove last row if empty (TeX behavior)
 * - take the \hline rules that follow the last row out of it (no new row)
 * - VisualTeX: take the rules and cell wrappers out of the cells; a spanned
 *   cell's covered columns are `undefined`
 */

function normalizeCells(
  atom: ArrayAtom,
  input: (readonly Atom[])[][],
  options: {
    columns: readonly ColumnFormat[];
    minColumns: number;
    minRows: number;
    maxRows: number;
    cellWrappers?: readonly [number, number, readonly CellWrapper[]][];
  }
): { rows: (readonly Atom[] | undefined)[][]; rowRules: string[][] } {
  //
  // 0/ VisualTeX: rules at row starts and wrappers at cell starts
  // (TableMarkAtom) leave the cells; each cell gets its `first` atom here,
  // which carries the cell's wrappers (and, for now, its row's rules).
  // Cells a span covers (JSON) are dropped: step 3 adds them back.
  //
  const rulesOf = new Map<Atom, string[]>();
  const covered = new Set<string>();
  const jsonWrappers = new Map<string, readonly CellWrapper[]>();
  for (const [r, c, wrappers] of options.cellWrappers ?? []) {
    jsonWrappers.set(`${r},${c}`, wrappers);
    for (let k = 1; k < cellWrapperSpan(wrappers); k++) covered.add(`${r},${c + k}`);
  }
  const cells: (readonly Atom[])[][] = input.map((row, r) => {
    const result: (readonly Atom[])[] = [];
    row.forEach((cell, c) => {
      if (covered.has(`${r},${c}`) && (cell ?? []).every((x) => x.type === 'first'))
        return;
      const atoms = (cell ?? []).filter((x) => x.type !== 'first');
      const rules: string[] = [];
      const wrappers: CellWrapper[] = [...(jsonWrappers.get(`${r},${c}`) ?? [])];
      let k = 0;
      for (; k < atoms.length; k++) {
        const x = atoms[k];
        if (x instanceof TableMarkAtom && x.rule !== undefined && wrappers.length === 0)
          rules.push(x.rule);
        else if (x instanceof TableMarkAtom && x.wrapper) wrappers.push(x.wrapper);
        // a starting commit's JSON: \hline as an unknown command
        else if (x.type === 'error' && x.command === '\\hline' && c === 0 && wrappers.length === 0)
          rules.push('\\hline');
        else break;
      }
      const first = cell?.[0]?.type === 'first' ? cell[0] : new Atom({ type: 'first', mode: atom.mode });
      if (rules.length > 0) rulesOf.set(first, rules);
      if (wrappers.length > 0) CELL_WRAPPERS.set(first, wrappers);
      result.push([first, ...atoms.slice(k)]);
    });
    return result;
  });

  //
  // 1/
  // - Fold the array so that there are no more columns of content than
  // there are columns prescribed by the column format.
  // - Fill rows that have fewer cells than expected with empty cells
  // - Ensure that all the cells have a `first` atom.
  //

  // The number of column is determined by the colFormat
  let maxColCount = 0;
  for (const colSpec of options.columns)
    if ('align' in colSpec) maxColCount += 1;
  maxColCount = Math.max(maxColCount, options.minColumns);

  // Actual number of columns (at most `maxColCount`)
  let colCount = 0;
  const rows: (readonly Atom[])[][] = [];

  for (const row of cells) {
    colCount = Math.max(colCount, Math.min(row.length, maxColCount));
    let colIndex = 0;
    while (colIndex < row.length) {
      const newRow: (readonly Atom[])[] = [];
      const lastCol = Math.min(row.length, colIndex + maxColCount);
      while (colIndex < lastCol) {
        newRow.push(normalizeCell(row[colIndex], atom.columnMode(colIndex % maxColCount)));
        colIndex += 1;
      }

      if (newRow.length < options.minColumns) {
        while (newRow.length < options.minColumns)
          newRow.push([new Atom({ type: 'first', mode: atom.mode })]);
      }

      rows.push(newRow);
    }
  }

  //
  // 2/ If the last row is empty, ignore it (TeX behavior)
  // (unless there's only one row).
  // A last row of nothing but \hline is the rules under the row before it
  // (`\\ \hline\end{array}`): it is no row either, its rules are kept.
  //
  let bottomRules: string[] = [];
  if (!atom.isMultiline && rows.length > 0 && rows[rows.length - 1].length === 1) {
    const last = rows[rows.length - 1][0];
    const rules = rulesOf.get(last[0]) ?? [];
    // (a cell with a wrapper, \hdotsfor{3}, is no empty row)
    if (isEmptyCell(last) && !CELL_WRAPPERS.has(last[0]) && (rules.length === 0 || rows.length > 1)) {
      bottomRules = rules;
      rows.pop();
    }
  }

  // VisualTeX: a span counts the columns it covers
  for (const row of rows) colCount = Math.max(colCount, Math.min(rowWidth(row), maxColCount));

  //
  // 3/ Fill out any missing cells
  //
  const result: (readonly Atom[] | undefined)[][] = [];
  for (const row of rows) {
    if (rowWidth(row) !== colCount) {
      for (let i = rowWidth(row); i < colCount; i++) {
        if (atom.isMultiline)
          row.push([new Atom({ type: 'first', mode: atom.mode })]);
        else {
          row.push([
            new Atom({ type: 'first', mode: atom.mode }),
            new PlaceholderAtom(),
          ]);
        }
      }
    }
    // VisualTeX: the columns a span covers are no cells (`undefined`)
    const spanned: (readonly Atom[] | undefined)[] = [];
    for (const cell of row) {
      spanned.push(cell);
      const span = Math.min(cellWrapperSpan(wrappersOf(cell)), colCount - spanned.length + 1);
      for (let k = 1; k < span; k++) spanned.push(undefined);
    }
    result.push(spanned);
  }

  //
  // 4/ Set the `parent` and `treeBranch` for each cell
  //
  let rowIndex = 0;
  let colIndex = 0;
  for (const row of result) {
    colIndex = 0;
    for (const cell of row) {
      for (const element of cell ?? []) {
        element.parent = atom;
        element.parentBranch = [rowIndex, colIndex];
      }
      colIndex += 1;
    }
    rowIndex += 1;
  }

  atom.isDirty = true;

  const rowRules = result.map((row) => rulesOf.get(row[0]?.[0] as Atom) ?? []);
  rowRules.push(bottomRules);
  return { rows: result, rowRules };
}

function normalizeCell(
  cell: readonly Atom[],
  mode: ParseMode
): readonly Atom[] {
  if (cell.length === 0) return [new Atom({ type: 'first', mode })];
  if (cell[0].type !== 'first')
    return [new Atom({ type: 'first', mode }), ...cell];
  console.assert(!cell.slice(1).some((x) => x.type === 'first'));
  return cell;
}

function isEmptyCell(cell: readonly Atom[]): boolean {
  return cell.length === 1 && cell[0].type === 'first';
}

// See http://ctan.math.utah.edu/ctan/tex-archive/macros/latex/base/lttab.dtx
export class ArrayAtom extends Atom {
  environmentName: Environment;

  /** True if the environment is multiline. In a multiline environment,
   * pressing return automatically creates a new row. In non-multiline
   * environments, pressing alt+return creates a new row.
   */
  get isMultiline(): boolean {
    const env = this.environmentName;
    return [
      'lines',
      'multline',
      'multline*',
      'align',
      'split',
      'gather',
      'gathered',
      // VisualTeX: every amsmath display environment is a list of lines
      'align*',
      'gather*',
      'flalign',
      'flalign*',
      'alignat',
      'alignat*',
      'eqnarray',
      'eqnarray*',
    ].includes(env);
  }

  // The array is a 2D array of cells, each cell being an array of atoms
  private _rows: (undefined | readonly Atom[])[][];

  /**
   * VisualTeX: the rules before each row (\hline, \cline{2-3}, \toprule,
   * ...) as written (a space after an argument kept, see scanTableMark);
   * the entry after the last row's is the rules under it
   */
  rowRules: string[][];

  /** VisualTeX: the number of rules under the last row, whichever row is last */
  get bottomRules(): number {
    return this.rowRules[this._rows.length]?.length ?? 0;
  }

  /** VisualTeX: the column specification as written (array, tabular) */
  colSpec?: string;

  /** VisualTeX: the arguments after \begin{name} (`{2}`, `[t]`, `[r]`), as written */
  environmentArguments?: string;

  /** VisualTeX: mathtools' cases* family sets its second column in text mode */
  columnMode(col: number): ParseMode {
    return col === 1 && /^d?r?cases\*$/.test(this.environmentName) ? 'text' : this.mode;
  }

  rowGaps: readonly Dimension[];
  arraystretch?: number;
  arraycolsep?: number;
  colSeparationType?: ColSeparationType;
  leftDelim?: string;
  rightDelim?: string;
  mathstyleName?: MathstyleName;

  colFormat: readonly ColumnFormat[];
  /** VisualTeX: aligned rows still typing into their second cell (see visualTexInsertAlignmentPoint) */
  visualTexProvisionalAlignRows?: Set<number>;
  minColumns: number;
  minRows: number;
  maxRows: number;

  // Additional classes when rendering the array
  classes: string[];

  constructor(
    envName: Environment,
    array: (readonly Atom[])[][],
    rowGaps: readonly Dimension[],
    options: Readonly<ArrayAtomConstructorOptions> = {}
  ) {
    super({ type: 'array', isRoot: options.isRoot, mode: options.mode });
    this.environmentName = envName;
    if (options.environmentArguments)
      this.environmentArguments = options.environmentArguments;

    if (options.columns) {
      if (options.columns.length === 0) this.colFormat = [{ align: 'l' }];
      else this.colFormat = [...options.columns];
    } else {
      if (options.minColumns) {
        // Repeat {align: 'l'} as many times as minColumns
        const columns: ColumnFormat[] = [];
        for (let i = 0; i < options.minColumns; i++)
          columns.push({ align: 'l' });
        this.colFormat = columns;
      } else {
        // In TeX definition, if no columns are specified, a maximum of
        // 10 left-aligned columns is assumed

        this.colFormat = [
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
          { align: 'l' },
        ];
      }
    }

    this.minColumns = options.minColumns ?? 1;
    this.minRows = options.minRows ?? 1;
    this.maxRows = options.maxRows ?? Infinity;

    const { rows, rowRules } = normalizeCells(this, array, {
      columns: this.colFormat,
      minColumns: this.minColumns,
      minRows: this.minRows,
      maxRows: this.maxRows,
      cellWrappers: options.cellWrappers,
    });
    this._rows = rows;
    // JSON gives the rules apart from the cells
    if (options.rowRules)
      this.rowRules = rows.map((_row, r) => [...(options.rowRules![r] ?? [])]);
    else this.rowRules = rowRules.slice(0, -1);
    const bottom = options.rowRules?.[rows.length] ?? rowRules[rowRules.length - 1];
    this.rowRules.push([
      ...bottom,
      ...(options.rowRules ? [] : Array<string>(options.bottomRules ?? 0).fill('\\hline')),
    ]);
    const colSpec = options.colSpec ?? (options.columns as { source?: string } | undefined)?.source;
    if (colSpec !== undefined) this.colSpec = colSpec;

    this.rowGaps = rowGaps;

    if (options.arraycolsep !== undefined)
      this.arraycolsep = options.arraycolsep;
    this.colSeparationType = options.colSeparationType;
    // Default \arraystretch from lttab.dtx
    if (options.arraystretch !== undefined)
      this.arraystretch = options.arraystretch;

    if (options.mathstyleName) this.mathstyleName = options.mathstyleName;
    if (options.leftDelim) this.leftDelim = options.leftDelim;
    if (options.rightDelim) this.rightDelim = options.rightDelim;

    this.classes = options.classes ?? [];
  }

  static fromJson(json: AtomJson): ArrayAtom {
    return new ArrayAtom(
      json.environmentName,
      json.array,
      json.rowGaps,
      json as any
    );
  }

  toJson(): AtomJson {
    const result: AtomJson = {
      ...super.toJson(),
      environmentName: this.environmentName,
      array: this._rows.map((row) =>
        row.map((col) => (col ?? []).map((x) => x.toJson()))
      ),
      rowGaps: this.rowGaps,
      // VisualTeX: the atoms of @{...} and !{...} as JSON (core/atom.ts fromJson)
      columns: this.colFormat.map((c) =>
        'gap' in c && Array.isArray(c.gap) ? { gap: c.gap.map((x) => x.toJson()) } : c
      ),
      colSeparationType: this.colSeparationType,
      minColumns: this.minColumns,
      minRows: this.minRows,
      maxRows: this.maxRows,
    };

    if (this.arraystretch !== undefined)
      result.arraystretch = this.arraystretch;
    if (this.arraycolsep !== undefined) result.arraycolsep = this.arraycolsep;
    if (this.leftDelim) result.leftDelim = this.leftDelim;
    if (this.rightDelim) result.rightDelim = this.rightDelim;
    if (this.isRoot) result.isRoot = true;
    result.minColumns = this.minColumns;
    result.minRows = this.minRows;
    result.maxRows = this.maxRows;
    if (this.mathstyleName) result.mathstyleName = this.mathstyleName;
    if (this.bottomRules > 0) result.bottomRules = this.bottomRules;
    if (this.environmentArguments)
      result.environmentArguments = this.environmentArguments;
    if (this.rowRules.some((rules) => rules.length > 0))
      result.rowRules = this.rowRules.map((rules) => [...rules]);
    if (this.colSpec !== undefined) result.colSpec = this.colSpec;
    const wrappers: [number, number, readonly CellWrapper[]][] = [];
    this._rows.forEach((row, r) =>
      row.forEach((cell, c) => {
        const w = wrappersOf(cell);
        if (w) wrappers.push([r, c, w]);
      })
    );
    if (wrappers.length > 0) result.cellWrappers = wrappers;
    if (this.classes.length > 0) result.classes = this.classes;

    return result;
  }

  branch(cell: Branch): readonly Atom[] | undefined {
    if (!isCellBranch(cell)) return undefined;
    return this._rows[cell[0]][cell[1]] ?? undefined;
  }

  createBranch(cell: Branch): Atom[] {
    if (!isCellBranch(cell)) return [];
    this.isDirty = true;
    return (this.branch(cell) as Atom[]) ?? [];
  }

  get rowCount(): number {
    return this._rows.length;
  }

  get colCount(): number {
    return this._rows[0].length;
  }

  /** VisualTeX: the \multicolumn, \multirow or \hdotsfor around a cell, outermost first */
  getCellWrappers(row: number, col: number): readonly CellWrapper[] | undefined {
    return wrappersOf(this._rows[row]?.[col]);
  }

  get maxColumns(): number {
    return this.colFormat.filter((col) => Boolean((col as { align?: string })['align'])).length;
  }

  removeBranch(name: Branch): readonly Atom[] {
    if (isNamedBranch(name)) return super.removeBranch(name);

    const [_first, ...children] = this.branch(name)!;
    // Drop the 'first' element
    console.assert(_first.type === 'first');

    this._rows[name[0]][name[1]] = undefined;
    children.forEach((x) => {
      x.parent = undefined;
      x.parentBranch = undefined;
    });

    this.isDirty = true;
    return children;
  }

  get hasChildren(): boolean {
    return this.children.length > 0;
  }

  _serialize(options: ToLatexOptions): string {
    const result: string[] = [];

    // Only wrap in \displaylines{} if there are multiple rows
    if (this.environmentName === 'lines' && this._rows.length > 1)
      result.push(`\\displaylines{`);
    else if (this.environmentName !== 'lines')
      result.push(`\\begin{${this.environmentName}}`);

    if (this.environmentArguments) result.push(this.environmentArguments);

    if (
      this.environmentName === 'array' ||
      this.environmentName === 'subarray' ||
      this.environmentName === 'tabular'
    ) {
      result.push('{');
      // VisualTeX: the column specification as written
      if (this.colSpec !== undefined) result.push(this.colSpec);
      else if (this.colFormat !== undefined) {
        for (const format of this.colFormat) {
          if ('align' in format && typeof format.align === 'string')
            result.push(format.align);
          else if ('separator' in format && format.separator === 'solid')
            result.push('|');
          else if ('separator' in format && format.separator === 'dashed')
            result.push(':');
        }
      }

      result.push('}');
    }

    for (let row = 0; row < this._rows.length; row++) {
      // VisualTeX: the rules before the row; no `&` for the columns a span
      // covers; the cell's wrappers around its content
      if (this.rowRules[row]?.length) result.push(this.rowRules[row].join('') + ' ');
      let firstCell = true;
      for (let col = 0; col < this._rows[row].length; col++) {
        const cell = this._rows[row][col];
        if (!cell) continue;
        if (!firstCell) result.push(' & ');
        firstCell = false;
        result.push(
          serializeCell(
            cell,
            options,
            this.environmentName === 'tabular' || this.columnMode(col) === 'text'
          )
        );
      }

      // Adds a separator between rows (but not after the last row, unless
      // rules follow it)
      if (row < this._rows.length - 1 || this.bottomRules > 0) {
        const gap = this.rowGaps[row];
        if (gap?.dimension)
          result.push(`\\\\[${gap.dimension} ${gap.unit ?? 'pt'}] `);
        else result.push('\\\\ ');
      }
    }
    if (this.bottomRules > 0) result.push(this.rowRules[this._rows.length].join('') + ' ');

    // Only close \displaylines{} if there are multiple rows
    if (this.environmentName === 'lines' && this._rows.length > 1)
      result.push(`}`);
    else if (this.environmentName !== 'lines')
      result.push(`\\end{${this.environmentName}}`);

    return joinLatex(result);
  }

  forEachCell(
    callback: (cell: readonly Atom[], row: number, col: number) => void
  ): void {
    for (let i = 0; i < this.rowCount; i++)
      for (let j = 0; j < this.colCount; j++) {
        // VisualTeX: a column a span covers is no cell
        const cell = this._rows[i][j];
        if (cell) callback(cell, i, j);
      }
  }

  getCell(row: number, col: number): readonly Atom[] | undefined {
    return this._rows[row]?.[col];
  }

  setCell(row: number, column: number, value: readonly Atom[]): void {
    console.assert(
      this.type === 'array' &&
        Array.isArray(this._rows) &&
        this._rows[row][column] !== undefined
    );
    for (const atom of this._rows[row][column]!) {
      atom.parent = undefined;
      atom.parentBranch = undefined;
    }

    const atoms = [
      new Atom({ type: 'first', mode: this.columnMode(column) }),
      ...value.filter((x) => x.type !== 'first'),
    ];
    // VisualTeX: the cell keeps its wrappers
    const wrappers = wrappersOf(this._rows[row][column]);
    if (wrappers) CELL_WRAPPERS.set(atoms[0], wrappers);

    this._rows[row][column] = atoms;

    adjustBranches(this);

    this.isDirty = true;
  }

  addRowBefore(row: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));

    this._rows.splice(
      row,
      0,
      Array.from({ length: this.colCount }, (_, col) =>
        makeEmptyCell(this, !this.isMultiline, this.columnMode(col))
      )
    );
    // VisualTeX: the rules stay with the rows they are before
    this.rowRules.splice(row, 0, []);
    adjustBranches(this);
    this.isDirty = true;
  }

  addRowAfter(row: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));

    this._rows.splice(
      row + 1,
      0,
      Array.from({ length: this.colCount }, (_, col) =>
        makeEmptyCell(this, !this.isMultiline, this.columnMode(col))
      )
    );
    this.rowRules.splice(row + 1, 0, []);

    adjustBranches(this);
    this.isDirty = true;
  }

  removeRow(row: number): void {
    console.assert(
      this.type === 'array' && Array.isArray(this._rows) && this.rowCount > row
    );

    const deleted = this._rows.splice(row, 1);
    this.rowRules.splice(row, 1);
    for (const column of deleted) {
      for (const cell of column) {
        if (cell) {
          for (const child of cell) {
            child.parent = undefined;
            child.parentBranch = undefined;
          }
        }
      }
    }

    adjustBranches(this);
    this.isDirty = true;
  }

  addColumnBefore(col: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));
    for (const row of this._rows) row.splice(col, 0, makeEmptyCell(this));

    fixSpans(this);
    adjustBranches(this);
    this.isDirty = true;
  }

  addColumnAfter(col: number): void {
    console.assert(this.type === 'array' && Array.isArray(this._rows));
    for (const row of this._rows) row.splice(col + 1, 0, makeEmptyCell(this));

    fixSpans(this);
    adjustBranches(this);
    this.isDirty = true;
  }

  addColumn(): void {
    this.addColumnAfter(this.colCount - 1);
  }

  removeColumn(col: number): void {
    console.assert(
      this.type === 'array' && Array.isArray(this._rows) && this.colCount > col
    );
    for (const row of this._rows) {
      const deleted = row.splice(col, 1);
      for (const cell of deleted) {
        if (cell) {
          for (const child of cell) {
            child.parent = undefined;
            child.parentBranch = undefined;
          }
        }
      }
    }
    fixSpans(this);
    adjustBranches(this);
    this.isDirty = true;
  }

  get cells(): (readonly Atom[])[] {
    const result: (readonly Atom[])[] = [];
    for (const row of this._rows) {
      for (const cell of row)
        if (cell) result.push(cell.filter((x) => x.type !== 'first'));
    }
    return result;
  }

  get rows(): (undefined | readonly Atom[])[][] {
    return this._rows!;
  }
}

/**
 * Create an empty cell
 */
function makeEmptyCell(
  parent: ArrayAtom,
  withPlaceholder = false,
  mode: ParseMode = parent.mode
): readonly Atom[] {
  const first = new Atom({ type: 'first', mode });
  first.parent = parent;
  const result = [first];
  if (withPlaceholder) {
    const placeholder = new PlaceholderAtom();
    placeholder.parent = parent;
    result.push(placeholder);
  }
  return result;
}

/** VisualTeX: a cell's LaTeX inside its wrappers (a tabular's cells and multirow's content are text) */
function serializeCell(cell: readonly Atom[], options: ToLatexOptions, tabular: boolean): string {
  const wrappers = wrappersOf(cell) ?? [];
  const text = tabular || wrappers.some((w) => w.command === '\\multirow');
  let latex = Atom.serialize(cell, text ? { ...options, defaultMode: 'text' } : options);
  for (let i = wrappers.length - 1; i >= 0; i--) {
    const head = cellWrapperLatex(wrappers[i]);
    if (wrappers[i].command === '\\hdotsfor') latex = latex.trim() ? `${head} ${latex}` : head;
    else latex = `${head}{${latex}}`;
  }
  return latex;
}

/**
 * VisualTeX: after a column was added or removed, the columns each span
 * covers are `undefined` again, and every other column is a cell.
 */
function fixSpans(array: ArrayAtom): void {
  const rows = array.rows;
  for (const row of rows) {
    let covered = 0;
    for (let c = 0; c < row.length; c++) {
      const cell = row[c];
      if (covered > 0 && (!cell || cell.every((x) => x.type === 'first'))) {
        row[c] = undefined;
        covered -= 1;
        continue;
      }
      covered = 0;
      if (!cell) row[c] = makeEmptyCell(array);
      else covered = cellWrapperSpan(wrappersOf(cell)) - 1;
    }
  }
}

function adjustBranches(array: ArrayAtom): void {
  for (let i = 0; i < array.rowCount; i++) {
    for (let j = 0; j < array.colCount; j++) {
      const atoms = array.getCell(i, j);
      if (atoms) {
        for (const atom of atoms) {
          if (atom) {
            atom.parent = array;
            atom.parentBranch = [i, j];
          }
        }
      }
    }
  }
}
